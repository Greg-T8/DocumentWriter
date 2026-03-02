<#
.SYNOPSIS
Add AI-generated commentary to a Word document using GitHub Models or Azure OpenAI.

.DESCRIPTION
Reads a Word (.docx) file, extracts its section outline and embedded
screenshots via a Python helper, sends each section (with images) to a
chat completions endpoint for analysis, and writes the commentary back
into the document as styled annotation paragraphs.

Supports two providers:
  - GitHub  : Uses GitHub Models (models.inference.ai.azure.com) with a
              GitHub CLI token. Free tier with rate limits.
  - Azure   : Uses an Azure OpenAI deployment (via AI Foundry) with an
              Entra ID bearer token from the Azure CLI. Production-grade
              throughput and token limits.

.CONTEXT
DocumentWriter project - AI-powered document review.

.AUTHOR
Greg Tate

.NOTES
Program: Invoke-DocumentCommentary.ps1
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$DocumentPath,

    [ValidateSet('GitHub', 'Azure')]
    [string]$Provider = 'GitHub',

    [string]$Model = 'openai/gpt-4o',

    [string]$AzureEndpoint,

    [string]$AzureDeployment,

    [string]$OutputPath,

    [string]$SystemPromptFile = 'prompts/SystemPrompt.md',

    [ValidateRange(0, 120)]
    [int]$RequestDelaySeconds = 30,

    [string]$PythonExe = 'python'
)

# Configuration
$ProjectRoot          = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$GitHubModelsEndpoint = 'https://models.inference.ai.azure.com/chat/completions'
$ProcessorScript      = Join-Path $ProjectRoot 'src/python/docx_processor.py'
$TempDir              = Join-Path $ProjectRoot '.docwriter_temp'
$ResolvedToken        = $null
$ResolvedEndpoint     = $null
$ResolvedModel        = $null
$ResolvedSystemPrompt = $null

$Main = {
    . $Helpers

    Confirm-Prerequisite
    Initialize-Authentication
    Initialize-SystemPrompt
    $extracted = Export-DocumentSection
    $commentary = Get-SectionCommentary -Sections $extracted.sections
    Import-Commentary -Commentary $commentary
    Remove-TempArtifact
}

$Helpers = {

    #region PREREQUISITE VALIDATION
    function Confirm-Prerequisite {
        # Verify Python and the processor script are available

        # Check that the Python helper script exists
        if (-not (Test-Path $ProcessorScript)) {
            throw "Python processor not found at: $ProcessorScript"
        }

        # Test Python availability
        try {
            $null = & $PythonExe --version 2>&1
        }
        catch {
            throw "Python executable not found: $PythonExe"
        }

        # Ensure python-docx is installed
        $pipCheck = & $PythonExe -c "import docx" 2>&1
        if ($LASTEXITCODE -ne 0) {
            Write-Host "Installing required Python packages..." -ForegroundColor Yellow

            $reqFile = Join-Path $ProjectRoot 'requirements.txt'
            if (Test-Path $reqFile) {
                & $PythonExe -m pip install -r $reqFile --quiet
            }
            else {
                & $PythonExe -m pip install python-docx Pillow requests --quiet
            }
        }

        # Create temp directory
        if (-not (Test-Path $TempDir)) {
            New-Item -ItemType Directory -Path $TempDir -Force | Out-Null
        }

        Write-Verbose "Prerequisites validated."
    }

    function Initialize-Authentication {
        # Route authentication to the selected provider

        if ($Provider -eq 'Azure') {
            Initialize-AzureAuthentication
        }
        else {
            Initialize-GitHubAuthentication
        }
    }

    function Initialize-GitHubAuthentication {
        # Authenticate using GitHub CLI and capture a token for API calls

        # Verify GitHub CLI is installed
        try {
            $null = & gh --version 2>&1
        }
        catch {
            throw "GitHub CLI ('gh') was not found. Install it from: https://cli.github.com/"
        }

        # Attempt to retrieve an auth token from existing CLI login
        $token = & gh auth token 2>$null

        # Launch interactive login if no token is available
        if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($token)) {
            Write-Host "GitHub CLI is not authenticated. Starting interactive login..." -ForegroundColor Yellow
            & gh auth login

            if ($LASTEXITCODE -ne 0) {
                throw "GitHub CLI authentication failed. Run 'gh auth login' and try again."
            }

            $token = & gh auth token 2>$null
        }

        # Validate token retrieval after authentication
        if ([string]::IsNullOrWhiteSpace($token)) {
            throw "Unable to retrieve a GitHub token from CLI authentication."
        }

        $script:ResolvedToken = $token.Trim()

        # Resolve model name for GitHub Models endpoint compatibility
        $resolvedModel = $Model
        if ($GitHubModelsEndpoint -like 'https://models.inference.ai.azure.com*' -and $resolvedModel -match '/') {
            $resolvedModel = $resolvedModel.Split('/')[-1]
        }

        $script:ResolvedEndpoint = $GitHubModelsEndpoint
        $script:ResolvedModel    = $resolvedModel
        Write-Verbose "GitHub CLI authentication completed."
    }

    function Initialize-AzureAuthentication {
        # Authenticate using Azure CLI and obtain an Entra ID bearer token

        # Validate required Azure parameters
        if ([string]::IsNullOrWhiteSpace($AzureEndpoint)) {
            throw "The -AzureEndpoint parameter is required when using the Azure provider."
        }
        if ([string]::IsNullOrWhiteSpace($AzureDeployment)) {
            throw "The -AzureDeployment parameter is required when using the Azure provider."
        }

        # Verify Azure CLI is installed
        try {
            $null = & az --version 2>&1
        }
        catch {
            throw "Azure CLI ('az') was not found. Install it from: https://learn.microsoft.com/cli/azure/install-azure-cli"
        }

        # Verify Azure CLI login status
        $accountJson = & az account show 2>$null
        if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($accountJson)) {
            Write-Host "Azure CLI is not authenticated. Starting interactive login..." -ForegroundColor Yellow
            & az login

            if ($LASTEXITCODE -ne 0) {
                throw "Azure CLI authentication failed. Run 'az login' and try again."
            }
        }

        # Obtain a bearer token scoped to Cognitive Services
        $tokenJson = & az account get-access-token --resource https://cognitiveservices.azure.com --query accessToken -o tsv 2>$null

        if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($tokenJson)) {
            throw "Unable to retrieve an Azure access token for Cognitive Services."
        }

        $script:ResolvedToken = $tokenJson.Trim()

        # Build the Azure OpenAI chat completions endpoint URL
        $baseUrl = $AzureEndpoint.TrimEnd('/')
        $script:ResolvedEndpoint = "${baseUrl}/openai/deployments/${AzureDeployment}/chat/completions?api-version=2024-12-01-preview"
        $script:ResolvedModel    = $AzureDeployment

        Write-Verbose "Azure CLI authentication completed. Endpoint: $script:ResolvedEndpoint"
    }

    function Initialize-SystemPrompt {
        # Load the system prompt text from a markdown file

        # Resolve relative prompt file paths against the script root
        if ([System.IO.Path]::IsPathRooted($SystemPromptFile)) {
            $promptPath = $SystemPromptFile
        }
        else {
            $promptPath = Join-Path $ProjectRoot $SystemPromptFile
        }

        # Ensure the prompt file exists
        if (-not (Test-Path $promptPath -PathType Leaf)) {
            throw "System prompt file not found: $promptPath"
        }

        # Read and validate the prompt contents
        $promptText = Get-Content -Path $promptPath -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($promptText)) {
            throw "System prompt file is empty: $promptPath"
        }

        $script:ResolvedSystemPrompt = $promptText.Trim()
        Write-Verbose "Loaded system prompt file: $promptPath"
    }
    #endregion

    #region DOCUMENT EXTRACTION
    function Export-DocumentSection {
        # Call Python to extract section outlines and images from the docx

        $jsonOut = Join-Path $TempDir 'extracted_sections.json'
        $resolvedDoc = (Resolve-Path $DocumentPath).Path

        Write-Host "Extracting document sections..." -ForegroundColor Cyan

        # Run the Python extraction command
        $result = & $PythonExe $ProcessorScript extract $resolvedDoc $jsonOut 2>&1
        if ($LASTEXITCODE -ne 0) {
            throw "Document extraction failed: $result"
        }

        Write-Verbose $result

        # Load and return the extracted JSON data
        $extracted = Get-Content $jsonOut -Raw | ConvertFrom-Json
        $sectionCount = $extracted.sections.Count
        Write-Host "Found $sectionCount sections in document." -ForegroundColor Green

        return $extracted
    }
    #endregion

    #region GITHUB MODELS API
    function Get-SectionCommentary {
        # Send each section to GitHub Models and collect AI commentary
        param(
            [Parameter(Mandatory)]
            [array]$Sections
        )

        $commentary = @()
        $totalSections = $Sections.Count

        foreach ($section in $Sections) {
            $index = $commentary.Count + 1
            $heading = $section.heading

            # Pace requests to reduce per-minute model rate-limit failures
            if ($index -gt 1 -and $RequestDelaySeconds -gt 0) {
                Start-Sleep -Seconds $RequestDelaySeconds
            }

            Write-Host "[$index/$totalSections] Analyzing: $heading" -ForegroundColor Cyan

            # Build the prompt content for this section
            $messages = Build-ApiMessage -Section $section

            # Call the chat completions API
            $response = Invoke-ModelApi -Messages $messages

            # Extract the commentary text from the response
            $commentaryText = $response.choices[0].message.content

            $commentary += @{
                heading    = $heading
                commentary = $commentaryText
            }

            Write-Verbose "Commentary generated for: $heading"
        }

        return $commentary
    }

    function Build-ApiMessage {
        # Construct the messages array for the chat completions API
        param(
            [Parameter(Mandatory)]
            [object]$Section
        )

        # Ensure a system prompt is loaded before constructing messages
        if ([string]::IsNullOrWhiteSpace($script:ResolvedSystemPrompt)) {
            throw "System prompt is unavailable. Ensure 'Initialize-SystemPrompt' runs before API calls."
        }

        # Build the user content parts (text + optional images)
        $contentParts = @()

        # Compose the section text prompt
        $sectionText = "## Section: $($Section.heading)`n`n"
        if ($Section.text_content -and $Section.text_content.Count -gt 0) {
            $bodyText = $Section.text_content -join "`n"
            $sectionText += "Content:`n$bodyText"
        }
        else {
            $sectionText += "Content: (no body text)"
        }

        $contentParts += @{
            type = 'text'
            text = $sectionText
        }

        # Add any images from the section as base64-encoded parts
        if ($Section.images -and $Section.images.Count -gt 0) {
            $imgIndex = 0
            foreach ($img in $Section.images) {
                $imgIndex++

                $contentParts += @{
                    type = 'text'
                    text = "[Screenshot $imgIndex in this section]"
                }

                $contentParts += @{
                    type      = 'image_url'
                    image_url = @{
                        url = "data:$($img.mime_type);base64,$($img.base64)"
                    }
                }
            }
        }

        # Assemble the full messages array
        $messages = @(
            @{
                role    = 'system'
                content = $script:ResolvedSystemPrompt
            },
            @{
                role    = 'user'
                content = $contentParts
            }
        )

        return $messages
    }

    function Get-RetryDelayFromHeaders {
        # Parse server-provided retry headers and return delay + source metadata
        param(
            [Parameter(Mandatory)]
            [object]$ErrorRecord
        )

        $response = $null
        if ($ErrorRecord.Exception -and $ErrorRecord.Exception.Response) {
            $response = $ErrorRecord.Exception.Response
        }

        if (-not $response -or -not $response.Headers) {
            return $null
        }

        $headerMap = @{}

        # Normalize response headers into a case-insensitive hashtable
        foreach ($name in $response.Headers.Keys) {
            $rawValue = $response.Headers[$name]
            if ($rawValue -is [System.Array] -or $rawValue -is [System.Collections.IEnumerable]) {
                $headerMap[$name.ToLowerInvariant()] = ($rawValue | ForEach-Object { $_.ToString() }) -join ','
            }
            else {
                $headerMap[$name.ToLowerInvariant()] = [string]$rawValue
            }
        }

        $candidateDelays = @()

        # Interpret Retry-After as either seconds or an HTTP date
        if ($headerMap.ContainsKey('retry-after')) {
            $retryAfter = $headerMap['retry-after']

            $retryAfterSeconds = 0
            if ([int]::TryParse($retryAfter, [ref]$retryAfterSeconds)) {
                if ($retryAfterSeconds -gt 0) {
                    $candidateDelays += [pscustomobject]@{
                        DelaySeconds = $retryAfterSeconds
                        Source       = 'retry-after-seconds'
                    }
                }
            }
            else {
                $retryAfterDate = [DateTimeOffset]::MinValue
                if ([DateTimeOffset]::TryParse($retryAfter, [ref]$retryAfterDate)) {
                    $dateDelay = [int][Math]::Ceiling(($retryAfterDate - [DateTimeOffset]::UtcNow).TotalSeconds)
                    if ($dateDelay -gt 0) {
                        $candidateDelays += [pscustomobject]@{
                            DelaySeconds = $dateDelay
                            Source       = 'retry-after-date'
                        }
                    }
                }
            }
        }

        # Interpret x-ms-retry-after-ms in milliseconds
        if ($headerMap.ContainsKey('x-ms-retry-after-ms')) {
            $retryAfterMs = 0
            if ([int]::TryParse($headerMap['x-ms-retry-after-ms'], [ref]$retryAfterMs)) {
                $msDelay = [int][Math]::Ceiling($retryAfterMs / 1000.0)
                if ($msDelay -gt 0) {
                    $candidateDelays += [pscustomobject]@{
                        DelaySeconds = $msDelay
                        Source       = 'x-ms-retry-after-ms'
                    }
                }
            }
        }

        # Interpret x-ratelimit-reset headers as either epoch time or relative seconds
        foreach ($headerName in @('x-ratelimit-reset', 'x-ratelimit-reset-requests', 'x-ratelimit-reset-tokens')) {
            if (-not $headerMap.ContainsKey($headerName)) {
                continue
            }

            $resetValue = 0L
            if (-not [long]::TryParse($headerMap[$headerName], [ref]$resetValue)) {
                continue
            }

            $relativeDelay = [int]$resetValue

            if ($resetValue -gt 1000000000) {
                $epochSeconds = $resetValue
                if ($resetValue -gt 20000000000) {
                    $epochSeconds = [long][Math]::Floor($resetValue / 1000.0)
                }

                $resetAt = [DateTimeOffset]::FromUnixTimeSeconds($epochSeconds)
                $relativeDelay = [int][Math]::Ceiling(($resetAt - [DateTimeOffset]::UtcNow).TotalSeconds)
            }

            if ($relativeDelay -gt 0) {
                $candidateDelays += [pscustomobject]@{
                    DelaySeconds = $relativeDelay
                    Source       = $headerName
                }
            }
        }

        if ($candidateDelays.Count -eq 0) {
            return $null
        }

        $selected = $candidateDelays | Sort-Object DelaySeconds -Descending | Select-Object -First 1
        return $selected
    }

    function Invoke-ModelApi {
        # Send a request to the configured chat completions endpoint
        param(
            [Parameter(Mandatory)]
            [array]$Messages
        )

        # Ensure authentication has already produced a token
        if ([string]::IsNullOrWhiteSpace($script:ResolvedToken)) {
            throw "API token is unavailable. Ensure 'Initialize-Authentication' runs before API calls."
        }

        $headers = @{
            'Authorization' = "Bearer $script:ResolvedToken"
            'Content-Type'  = 'application/json'
        }

        $body = @{
            model       = $script:ResolvedModel
            messages    = $Messages
            temperature = 0.4
            max_tokens  = 500
        } | ConvertTo-Json -Depth 20

        # Retry logic for transient API failures
        $maxRetries = 6
        $retryDelay = 5

        for ($attempt = 1; $attempt -le $maxRetries; $attempt++) {
            try {
                $response = Invoke-RestMethod `
                    -Uri $script:ResolvedEndpoint `
                    -Method Post `
                    -Headers $headers `
                    -Body $body `
                    -ErrorAction Stop

                return $response
            }
            catch {
                $statusCode = $null
                if ($_.Exception.Response) {
                    $statusCode = $_.Exception.Response.StatusCode.value__
                }

                $errorText = $_.ToString()
                $serverSuggestedDelay = $null
                $delaySource = 'exponential-backoff'

                # Prefer explicit server retry headers when available
                $headerRetryInfo = Get-RetryDelayFromHeaders -ErrorRecord $_
                if ($headerRetryInfo) {
                    $serverSuggestedDelay = [int]$headerRetryInfo.DelaySeconds
                    $delaySource = [string]$headerRetryInfo.Source
                }

                # Fall back to parsing a delay hint from the response text
                if ($errorText -match 'Please wait\s+(\d+)\s+seconds') {
                    $messageSuggestedDelay = [int]$Matches[1]
                    if (-not $serverSuggestedDelay -or $messageSuggestedDelay -gt $serverSuggestedDelay) {
                        $serverSuggestedDelay = $messageSuggestedDelay
                        $delaySource = 'response-message'
                    }
                }

                if ($serverSuggestedDelay) {
                    $boundedDelay = [Math]::Min(180, [Math]::Max(3, $serverSuggestedDelay + 1))
                    $retryDelay = [Math]::Max($retryDelay, $boundedDelay)
                }

                Write-Verbose "Retry delay source: $delaySource; next wait: ${retryDelay}s"

                # Retry on rate-limit (429) or server errors (5xx)
                if ($attempt -lt $maxRetries -and ($statusCode -eq 429 -or $statusCode -ge 500)) {
                    Write-Warning "API request failed (HTTP $statusCode). Retrying in ${retryDelay}s... (attempt $attempt/$maxRetries)"
                    Start-Sleep -Seconds $retryDelay
                    $retryDelay *= 2
                }
                else {
                    throw "$Provider API error: $_"
                }
            }
        }
    }
    #endregion

    #region COMMENTARY INSERTION
    function Import-Commentary {
        # Write commentary JSON and call Python to insert into document
        param(
            [Parameter(Mandatory)]
            [array]$Commentary
        )

        # Determine output file path
        $resolvedDoc = (Resolve-Path $DocumentPath).Path
        $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
        $defaultOutputDirectory = Join-Path $ProjectRoot 'Output'

        # Create the default output directory when it does not exist
        if (-not (Test-Path -Path $defaultOutputDirectory -PathType Container)) {
            New-Item -ItemType Directory -Path $defaultOutputDirectory -Force | Out-Null
        }

        if ($OutputPath) {
            $outputDirectory = [System.IO.Path]::GetDirectoryName($OutputPath)
            $outputFileName = [System.IO.Path]::GetFileNameWithoutExtension($OutputPath)
            $outputExtension = [System.IO.Path]::GetExtension($OutputPath)

            # Default to the project Output directory when only a filename is provided
            if ([string]::IsNullOrWhiteSpace($outputDirectory)) {
                $outputDirectory = $defaultOutputDirectory
            }

            # Ensure output has a .docx extension when omitted
            if ([string]::IsNullOrWhiteSpace($outputExtension)) {
                $outputExtension = '.docx'
            }

            # Fall back to source document name if output filename is omitted
            if ([string]::IsNullOrWhiteSpace($outputFileName)) {
                $outputFileName = [System.IO.Path]::GetFileNameWithoutExtension($resolvedDoc)
            }

            $outFile = Join-Path `
                $outputDirectory `
                "${outputFileName}_${timestamp}${outputExtension}"
        }
        else {
            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($resolvedDoc)
            $directory = $defaultOutputDirectory
            $outFile = Join-Path $directory "${baseName}_commented_${timestamp}.docx"
        }

        # Ensure the final output directory exists
        $finalOutputDirectory = [System.IO.Path]::GetDirectoryName($outFile)
        if (-not (Test-Path -Path $finalOutputDirectory -PathType Container)) {
            New-Item -ItemType Directory -Path $finalOutputDirectory -Force | Out-Null
        }

        # Ensure the output path never points to the original source document
        $sourceFullPath = [System.IO.Path]::GetFullPath($resolvedDoc)
        $outputFullPath = [System.IO.Path]::GetFullPath($outFile)

        if (
            [System.StringComparer]::OrdinalIgnoreCase.Equals(
                $sourceFullPath,
                $outputFullPath
            )
        ) {
            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($resolvedDoc)
            $directory = $defaultOutputDirectory
            $fallbackTimestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
            $outputFullPath = Join-Path $directory "${baseName}_commented_${fallbackTimestamp}.docx"

            Write-Warning (
                "Requested output matched the source document. " +
                "Using a new file instead: $outputFullPath"
            )
        }

        $outFile = $outputFullPath

        # Write commentary to a temp JSON file
        $commentaryJson = Join-Path $TempDir 'commentary.json'
        $Commentary | ConvertTo-Json -Depth 10 | Set-Content -Path $commentaryJson -Encoding UTF8

        Write-Host "Inserting commentary into document..." -ForegroundColor Cyan

        # Run the Python insert command
        $result = & $PythonExe $ProcessorScript insert $resolvedDoc $commentaryJson $outFile 2>&1
        if ($LASTEXITCODE -ne 0) {
            throw "Commentary insertion failed: $result"
        }

        Write-Verbose $result
        Write-Host "Annotated document saved to: $outFile" -ForegroundColor Green
    }
    #endregion

    #region CLEANUP
    function Remove-TempArtifact {
        # Clean up temporary files created during processing

        if (Test-Path $TempDir) {
            Remove-Item -Path $TempDir -Recurse -Force
            Write-Verbose "Cleaned up temp directory: $TempDir"
        }
    }
    #endregion
}

try {
    Push-Location -Path $PSScriptRoot
    & $Main
}
finally {
    Pop-Location
}
