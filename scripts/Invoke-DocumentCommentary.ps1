<#
.SYNOPSIS
Revise a Word document using AI analysis from GitHub Models or Azure OpenAI.

.DESCRIPTION
Reads a Word (.docx) file, extracts its section outline and embedded
text via a Python helper, sends each section to a chat completions
endpoint for analysis, and revises the document based on AI-generated
suggestions.

Supports two providers:
  - GitHub  : Uses GitHub Models (models.inference.ai.azure.com) with a
              GitHub CLI token. Free tier with rate limits.
  - Azure   : Uses an Azure OpenAI deployment (via AI Foundry) with an
              Entra ID bearer token from the Azure CLI. Production-grade
              throughput and token limits.

Supports multiple Azure subscriptions via the -Subscription parameter:
  - msdn : MSDN subscription (spending limit, restricted model catalog)
  - payg : Pay-As-You-Go subscription (full model catalog access)

Each subscription maps to a Terraform workspace for isolated state
management while sharing the same infrastructure code.

.CONTEXT
DocumentWriter project - AI-powered document revision.

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
    [string]$Provider = 'Azure',

    [ValidateSet('msdn', 'payg')]
    [string]$Subscription = 'msdn',

    [string]$Model = 'gpt-5.1-chat',

    [string]$AzureEndpoint,

    [string]$AzureDeployment,

    [string]$OutputPath,

    [string]$SystemPromptFile = 'prompts/SystemPrompt.md',

    [ValidateRange(0, 120)]
    [int]$RequestDelaySeconds = 5,

    [string]$PythonExe = 'python'
)

# Configuration
$ProjectRoot          = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$GitHubModelsEndpoint = 'https://models.inference.ai.azure.com/chat/completions'
$ProcessorScript      = Join-Path $ProjectRoot 'src/python/docx_processor.py'
$TerraformDirectory   = Join-Path $ProjectRoot 'terraform'
$TempDir              = Join-Path $ProjectRoot '.docwriter_temp'
$ResolvedToken        = $null
$ResolvedEndpoint     = $null
$ResolvedModel        = $null
$ResolvedSystemPrompt = $null

# Subscription-to-workspace mapping
$SubscriptionMap = @{
    'msdn' = @{
        SubscriptionId     = 'e091f6e7-031a-4924-97bb-8c983ca5d21a'
        TerraformWorkspace = 'msdn'
        TfVarsFile         = 'msdn.tfvars'
    }
    'payg' = @{
        SubscriptionId     = 'e6ad7655-b3ba-4324-8361-fcfdc59973a5'
        TerraformWorkspace = 'payg'
        TfVarsFile         = 'payg.tfvars'
    }
}

$Main = {
    . $Helpers

    Confirm-Prerequisite
    Initialize-Authentication
    Initialize-SystemPrompt
    $extracted = Export-DocumentSection
    $revision = Get-SectionRevision -Sections $extracted.sections
    Import-Revision -Revision $revision
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

        # Look up the selected subscription details from the mapping table
        $subInfo = $SubscriptionMap[$Subscription]
        $targetSubscriptionId = $subInfo.SubscriptionId

        $resolvedAzureEndpoint = $AzureEndpoint
        $resolvedAzureDeployment = $AzureDeployment

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

        # Set the Azure CLI context to the target subscription
        Write-Verbose "Setting Azure subscription to: $Subscription ($targetSubscriptionId)"
        & az account set --subscription $targetSubscriptionId 2>$null

        if ($LASTEXITCODE -ne 0) {
            throw "Failed to set Azure subscription to '$Subscription' ($targetSubscriptionId). Verify access with 'az account list'."
        }

        Write-Host "Using Azure subscription: $Subscription ($targetSubscriptionId)" -ForegroundColor Cyan

        # Resolve missing Azure settings from Terraform outputs when available
        if (
            [string]::IsNullOrWhiteSpace($resolvedAzureEndpoint) -or
            [string]::IsNullOrWhiteSpace($resolvedAzureDeployment)
        ) {
            $terraformSettings = Get-AzureSettingsFromTerraform

            if (
                [string]::IsNullOrWhiteSpace($resolvedAzureEndpoint) -and
                -not [string]::IsNullOrWhiteSpace($terraformSettings.AzureEndpoint)
            ) {
                $resolvedAzureEndpoint = $terraformSettings.AzureEndpoint
                Write-Verbose "Resolved Azure endpoint from Terraform output."
            }

            if (
                [string]::IsNullOrWhiteSpace($resolvedAzureDeployment) -and
                -not [string]::IsNullOrWhiteSpace($terraformSettings.AzureDeployment)
            ) {
                $resolvedAzureDeployment = $terraformSettings.AzureDeployment
                Write-Verbose "Resolved Azure deployment from Terraform output."
            }
        }

        # Validate required Azure parameters
        if ([string]::IsNullOrWhiteSpace($resolvedAzureEndpoint)) {
            throw "The -AzureEndpoint parameter is required when using the Azure provider."
        }
        if ([string]::IsNullOrWhiteSpace($resolvedAzureDeployment)) {
            throw "The -AzureDeployment parameter is required when using the Azure provider."
        }

        # Obtain a bearer token scoped to Cognitive Services
        $tokenJson = & az account get-access-token --resource https://cognitiveservices.azure.com --query accessToken -o tsv 2>$null

        if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($tokenJson)) {
            throw "Unable to retrieve an Azure access token for Cognitive Services."
        }

        $script:ResolvedToken = $tokenJson.Trim()

        # Build the Azure OpenAI chat completions endpoint URL
        $baseUrl = $resolvedAzureEndpoint.TrimEnd('/')
        $script:ResolvedEndpoint = "${baseUrl}/openai/deployments/${resolvedAzureDeployment}/chat/completions?api-version=2024-12-01-preview"
        $script:ResolvedModel    = $resolvedAzureDeployment

        Write-Verbose "Azure CLI authentication completed. Endpoint: $script:ResolvedEndpoint"
    }

    function Get-AzureSettingsFromTerraform {
        # Read Azure endpoint and deployment values from the Terraform workspace

        $azureSettings = [pscustomobject]@{
            AzureEndpoint   = $null
            AzureDeployment = $null
        }

        # Skip Terraform resolution when the terraform directory is unavailable
        if (-not (Test-Path -Path $TerraformDirectory -PathType Container)) {
            return $azureSettings
        }

        # Skip Terraform resolution when terraform CLI is not installed
        try {
            $null = & terraform version 2>$null
        }
        catch {
            return $azureSettings
        }

        # Resolve the Terraform workspace for the selected subscription
        $subInfo = $SubscriptionMap[$Subscription]
        $targetWorkspace = $subInfo.TerraformWorkspace

        try {
            Push-Location -Path $TerraformDirectory

            # Select the workspace that matches the target subscription
            $currentWorkspace = (& terraform workspace show 2>$null)
            if ($currentWorkspace -and $currentWorkspace.Trim() -ne $targetWorkspace) {
                Write-Verbose "Switching Terraform workspace to: $targetWorkspace"
                $null = & terraform workspace select $targetWorkspace 2>$null

                if ($LASTEXITCODE -ne 0) {
                    Write-Warning "Terraform workspace '$targetWorkspace' not found. Run 'terraform workspace new $targetWorkspace' to create it."
                    return $azureSettings
                }
            }

            # Read endpoint and deployment from known output names
            $endpointOutput = & terraform output -raw azure_openai_endpoint 2>$null
            if ($LASTEXITCODE -eq 0 -and -not [string]::IsNullOrWhiteSpace($endpointOutput)) {
                $azureSettings.AzureEndpoint = $endpointOutput.Trim()
            }

            $deploymentOutput = & terraform output -raw model_deployment_name 2>$null
            if ($LASTEXITCODE -eq 0 -and -not [string]::IsNullOrWhiteSpace($deploymentOutput)) {
                $azureSettings.AzureDeployment = $deploymentOutput.Trim()
            }
        }
        catch {
            return $azureSettings
        }
        finally {
            Pop-Location
        }

        return $azureSettings
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

    function Show-VerboseIfNotEmpty {
        # Write a verbose message only when text content is present
        param(
            [AllowNull()]
            [AllowEmptyString()]
            [string]$Message
        )

        if (-not [string]::IsNullOrWhiteSpace($Message)) {
            Write-Verbose $Message
        }
    }
    #endregion

    #region DOCUMENT EXTRACTION
    function Export-DocumentSection {
        # Call Python to extract section outlines and text from the docx

        $jsonOut = Join-Path $TempDir 'extracted_sections.json'
        $resolvedDoc = (Resolve-Path $DocumentPath).Path

        Write-Host "Extracting document sections..." -ForegroundColor Cyan

        # Run the Python extraction command
        $result = & $PythonExe $ProcessorScript extract $resolvedDoc $jsonOut 2>&1
        if ($LASTEXITCODE -ne 0) {
            throw "Document extraction failed: $result"
        }
        Show-VerboseIfNotEmpty -Message ([string]$result)

        # Load and return the extracted JSON data
        $extracted = Get-Content $jsonOut -Raw | ConvertFrom-Json
        $sectionCount = $extracted.sections.Count
        Write-Host "Found $sectionCount sections in document." -ForegroundColor Green

        return $extracted
    }
    #endregion

    #region GITHUB MODELS API
    function Get-SectionRevision {
        # Send each section to the model and collect rewritten prose
        param(
            [Parameter(Mandatory)]
            [array]$Sections
        )

        $revision = @()
        $totalSections = $Sections.Count

        foreach ($section in $Sections) {
            $index = $revision.Count + 1
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

            # Extract the revised text from the response
            $revisedText = $response.choices[0].message.content

            $revision += @{
                heading      = $heading
                revised_text = $revisedText
            }

            Write-Verbose "Revision generated for: $heading"
        }

        return $revision
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

        # Build the user content parts as text-only input
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

        $requestBody = @{
            model       = $script:ResolvedModel
            messages    = $Messages
        }

        # Set a custom temperature only for providers that support it
        if ($Provider -ne 'Azure') {
            $requestBody.temperature = 0.4
        }

        # Use provider-specific completion token parameter name
        if ($Provider -eq 'Azure') {
            $requestBody.max_completion_tokens = 500
        }
        else {
            $requestBody.max_tokens = 500
        }

        $body = $requestBody | ConvertTo-Json -Depth 20

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

    #region DOCUMENT REVISION
    function Import-Revision {
        # Write section revisions to JSON and call Python to rewrite the document
        param(
            [Parameter(Mandatory)]
            [array]$Revision
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
            $outFile = Join-Path $directory "${baseName}_revised_${timestamp}.docx"
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
            $outputFullPath = Join-Path $directory "${baseName}_revised_${fallbackTimestamp}.docx"

            Write-Warning (
                "Requested output matched the source document. " +
                "Using a new file instead: $outputFullPath"
            )
        }

        $outFile = $outputFullPath

        # Write section revisions to a temp JSON file
        $revisionJson = Join-Path $TempDir 'revision.json'
        $Revision | ConvertTo-Json -Depth 10 | Set-Content -Path $revisionJson -Encoding UTF8

        Write-Host "Rewriting document prose..." -ForegroundColor Cyan

        # Run the Python revise command
        $result = & $PythonExe $ProcessorScript revise $resolvedDoc $revisionJson $outFile 2>&1
        if ($LASTEXITCODE -ne 0) {
            throw "Document revision failed: $result"
        }
        Show-VerboseIfNotEmpty -Message ([string]$result)
        Write-Host "Revised document saved to: $outFile" -ForegroundColor Green
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
