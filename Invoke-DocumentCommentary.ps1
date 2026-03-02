<#
.SYNOPSIS
Add AI-generated commentary to a Word document using GitHub Models.

.DESCRIPTION
Reads a Word (.docx) file, extracts its section outline and embedded
screenshots via a Python helper, sends each section (with images) to a
GitHub Models endpoint for analysis, and writes the commentary back into
the document as styled annotation paragraphs.

.CONTEXT
DocumentWriter project - GitHub Models integration for document review.

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

    [string]$Model = 'openai/gpt-4.1',

    [string]$OutputPath,

    [string]$SystemPrompt = @"
You are a technical document reviewer. For each section of the document you
receive, provide concise, constructive commentary that:
1. Summarizes what the section covers.
2. Notes any issues, ambiguities, or missing information.
3. If screenshots are provided, describes what they show and whether they
   support the section text.
Keep commentary to 2-4 sentences per section. Be professional and helpful.
"@,

    [string]$PythonExe = 'python'
)

# Configuration
$GitHubModelsEndpoint = 'https://models.inference.ai.azure.com/chat/completions'
$ProcessorScript      = Join-Path $PSScriptRoot 'docx_processor.py'
$TempDir              = Join-Path $PSScriptRoot '.docwriter_temp'
$ResolvedGitHubToken  = $null

$Main = {
    . $Helpers

    Confirm-Prerequisite
    Initialize-GitHubAuthentication
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

            $reqFile = Join-Path $PSScriptRoot 'requirements.txt'
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

        $script:ResolvedGitHubToken = $token.Trim()
        Write-Verbose "GitHub CLI authentication completed."
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
            Write-Host "[$index/$totalSections] Analyzing: $heading" -ForegroundColor Cyan

            # Build the prompt content for this section
            $messages = Build-ApiMessage -Section $section

            # Call the GitHub Models API
            $response = Invoke-GitHubModel -Messages $messages

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
                content = $SystemPrompt
            },
            @{
                role    = 'user'
                content = $contentParts
            }
        )

        return $messages
    }

    function Invoke-GitHubModel {
        # Send a request to the GitHub Models chat completions endpoint
        param(
            [Parameter(Mandatory)]
            [array]$Messages
        )

        # Ensure authentication has already produced a token
        if ([string]::IsNullOrWhiteSpace($script:ResolvedGitHubToken)) {
            throw "GitHub token is unavailable. Ensure 'Initialize-GitHubAuthentication' runs before API calls."
        }

        $headers = @{
            'Authorization' = "Bearer $script:ResolvedGitHubToken"
            'Content-Type'  = 'application/json'
        }

        $body = @{
            model       = $Model
            messages    = $Messages
            temperature = 0.4
            max_tokens  = 500
        } | ConvertTo-Json -Depth 20

        # Retry logic for transient API failures
        $maxRetries = 3
        $retryDelay = 5

        for ($attempt = 1; $attempt -le $maxRetries; $attempt++) {
            try {
                $response = Invoke-RestMethod `
                    -Uri $GitHubModelsEndpoint `
                    -Method Post `
                    -Headers $headers `
                    -Body $body `
                    -ErrorAction Stop

                return $response
            }
            catch {
                $statusCode = $_.Exception.Response.StatusCode.value__

                # Retry on rate-limit (429) or server errors (5xx)
                if ($attempt -lt $maxRetries -and ($statusCode -eq 429 -or $statusCode -ge 500)) {
                    Write-Warning "API request failed (HTTP $statusCode). Retrying in ${retryDelay}s... (attempt $attempt/$maxRetries)"
                    Start-Sleep -Seconds $retryDelay
                    $retryDelay *= 2
                }
                else {
                    throw "GitHub Models API error: $_"
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
        if ($OutputPath) {
            $outFile = $OutputPath
        }
        else {
            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($resolvedDoc)
            $directory = [System.IO.Path]::GetDirectoryName($resolvedDoc)
            $outFile = Join-Path $directory "${baseName}_commented.docx"
        }

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
