<#
.SYNOPSIS
Invoke document commentary from the repository root entry point.

.DESCRIPTION
Provides a backward-compatible entry point and forwards parameters to the
implementation script in the scripts folder. Supports both GitHub Models
and Azure OpenAI providers for text-only document revision.

.CONTEXT
DocumentWriter project - root launcher for document commentary.

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

    [string]$SystemPromptFile = 'Input/SystemPrompt.md',

    [ValidateRange(0, 120)]
    [int]$RequestDelaySeconds = 30,

    [string]$PythonExe = 'python'
)

# Resolve the implementation script location inside the scripts folder.
$ImplementationScript = Join-Path $PSScriptRoot 'scripts/Invoke-DocumentCommentary.ps1'

# Ensure the implementation script exists before forwarding execution.
if (-not (Test-Path -Path $ImplementationScript -PathType Leaf)) {
    throw "Implementation script not found at: $ImplementationScript"
}

try {
    Push-Location -Path $PSScriptRoot
    & $ImplementationScript @PSBoundParameters
}
finally {
    Pop-Location
}
