<#
.SYNOPSIS
Lists models from GitHub Models and Azure Foundry in a single sorted view.

.DESCRIPTION
Retrieves model catalogs from the GitHub Models endpoint and from Azure
Cognitive Services account models. Azure account context is resolved from
Terraform outputs in the project terraform folder.

.CONTEXT
DocumentWriter project - model inventory across GitHub and Azure Foundry.

.AUTHOR
Greg Tate

.NOTES
Program: Get-ModelCatalog.ps1
#>

[CmdletBinding()]
param(
    [string]$TerraformDirectory = (Join-Path (Resolve-Path (Join-Path $PSScriptRoot '..')).Path 'terraform')
)

# Configuration
$GitHubModelsListEndpoint = 'https://models.inference.ai.azure.com/models'

$Main = {
    . $Helpers

    Confirm-Prerequisite
    $terraformContext = Get-TerraformFoundryContext
    $githubApiModels = Get-GitHubModelCatalog
    $githubCliModels = Get-GitHubCliModelCatalog
    $azureModels = Get-AzureFoundryModelCatalog -TerraformContext $terraformContext
    $combinedCatalog = New-CombinedModelCatalog -GitHubApiModels $githubApiModels -GitHubCliModels $githubCliModels -AzureModels $azureModels
    Show-ModelCatalog -Catalog $combinedCatalog
}

$Helpers = {

    #region PREREQUISITE VALIDATION
    # Functions for validating required CLIs and input directories.
    function Confirm-Prerequisite {
        # Validate that required CLIs and Terraform directory are available.

        if (-not (Test-Path -Path $TerraformDirectory -PathType Container)) {
            throw "Terraform directory not found: $TerraformDirectory"
        }

        try {
            $null = & terraform version 2>$null
        }
        catch {
            throw "Terraform CLI was not found. Install Terraform and try again."
        }

        try {
            $null = & az --version 2>$null
        }
        catch {
            throw "Azure CLI ('az') was not found. Install Azure CLI and try again."
        }

        try {
            $null = & gh --version 2>$null
        }
        catch {
            throw "GitHub CLI ('gh') was not found. Install GitHub CLI and try again."
        }
    }
    #endregion

    #region TERRAFORM RESOLUTION
    # Functions for resolving Azure Foundry context from Terraform configuration.
    function Get-TerraformFoundryContext {
        # Read Azure resource group, account name, and subscription from Terraform artifacts.

        $context = [pscustomobject]@{
            ResourceGroup  = $null
            AccountName    = $null
            SubscriptionId = $null
        }

        try {
            Push-Location -Path $TerraformDirectory

            $resourceGroup = & terraform output -raw resource_group_name 2>$null
            if ($LASTEXITCODE -eq 0 -and -not [string]::IsNullOrWhiteSpace($resourceGroup)) {
                $context.ResourceGroup = $resourceGroup.Trim()
            }

            $accountName = & terraform output -raw ai_foundry_name 2>$null
            if ($LASTEXITCODE -eq 0 -and -not [string]::IsNullOrWhiteSpace($accountName)) {
                $context.AccountName = $accountName.Trim()
            }
        }
        finally {
            Pop-Location
        }

        $tfVarsPath = Join-Path $TerraformDirectory 'terraform.tfvars'
        if (Test-Path -Path $tfVarsPath -PathType Leaf) {
            $tfVarsContent = Get-Content -Path $tfVarsPath -Raw -Encoding UTF8
            $subscriptionMatch = [regex]::Match($tfVarsContent, 'subscription_id\s*=\s*"([^"]+)"')

            if ($subscriptionMatch.Success) {
                $context.SubscriptionId = $subscriptionMatch.Groups[1].Value.Trim()
            }
        }

        if ([string]::IsNullOrWhiteSpace($context.ResourceGroup) -or [string]::IsNullOrWhiteSpace($context.AccountName)) {
            throw 'Unable to resolve Azure Foundry account from Terraform outputs. Run terraform apply in the terraform folder first.'
        }

        return $context
    }
    #endregion

    #region CATALOG RETRIEVAL
    # Functions for collecting model catalogs from GitHub and Azure Foundry.
    function Get-GitHubAccessToken {
        # Resolve a GitHub token from environment or GitHub CLI login.

        if (-not [string]::IsNullOrWhiteSpace($env:GITHUB_TOKEN)) {
            return $env:GITHUB_TOKEN.Trim()
        }

        if (-not [string]::IsNullOrWhiteSpace($env:GH_TOKEN)) {
            return $env:GH_TOKEN.Trim()
        }

        $token = & gh auth token 2>$null
        if ($LASTEXITCODE -eq 0 -and -not [string]::IsNullOrWhiteSpace($token)) {
            return $token.Trim()
        }

        throw "No GitHub token available. Run 'gh auth login' or set GITHUB_TOKEN."
    }

    function Get-GitHubModelCatalog {
        # Retrieve and normalize the GitHub Models catalog.

        $token = Get-GitHubAccessToken

        $headers = @{
            Authorization = "Bearer $token"
            Accept        = 'application/json'
        }

        $response = Invoke-RestMethod -Method Get -Uri $GitHubModelsListEndpoint -Headers $headers

        if ($response -is [array]) {
            $sourceItems = $response
        }
        elseif ($null -ne $response.data -and $response.data -is [array]) {
            $sourceItems = $response.data
        }
        else {
            $sourceItems = @($response)
        }

        return $sourceItems |
            Where-Object { $_ -ne $null } |
            ForEach-Object {
                [pscustomobject]@{
                    Model    = if (-not [string]::IsNullOrWhiteSpace($_.name)) { $_.name } elseif (-not [string]::IsNullOrWhiteSpace($_.id)) { $_.id } else { $null }
                    Version  = if (-not [string]::IsNullOrWhiteSpace($_.version)) { $_.version } else { 'n/a' }
                    Provider = if (-not [string]::IsNullOrWhiteSpace($_.publisherName)) { $_.publisherName } elseif (-not [string]::IsNullOrWhiteSpace($_.publisher)) { $_.publisher } elseif (-not [string]::IsNullOrWhiteSpace($_.provider)) { $_.provider } else { 'GitHub' }
                    Source   = 'GitHub Endpoint'
                }
            } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_.Model) }
    }

    function Get-GitHubCliModelCatalog {
        # Retrieve and normalize models returned by gh models list.

        $rows = & gh models list 2>$null
        if ($LASTEXITCODE -ne 0 -or $null -eq $rows) {
            return @()
        }

        $modelIds = $rows |
            ForEach-Object {
                $match = [regex]::Match($_, '^\s*([a-z0-9][a-z0-9\-]*/[a-z0-9][a-z0-9\-._]*)\b', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                if ($match.Success) {
                    $match.Groups[1].Value.Trim()
                }
            } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Sort-Object -Unique

        return $modelIds |
            ForEach-Object {
                $provider = if ($_ -match '^([^/]+)/') { $Matches[1] } else { 'GitHub' }

                [pscustomobject]@{
                    Model    = $_
                    Version  = 'n/a'
                    Provider = $provider
                    Source   = 'GitHub CLI'
                }
            } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_.Model) }
    }

    function Get-AzureFoundryModelCatalog {
        # Retrieve and normalize the Azure Foundry model catalog.
        param(
            [Parameter(Mandatory)]
            [pscustomobject]$TerraformContext
        )

        if (-not [string]::IsNullOrWhiteSpace($TerraformContext.SubscriptionId)) {
            $null = & az account set --subscription $TerraformContext.SubscriptionId 2>$null
        }

        $modelsJson = & az cognitiveservices account list-models --resource-group $TerraformContext.ResourceGroup --name $TerraformContext.AccountName -o json
        if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($modelsJson)) {
            throw "Failed to list models for account '$($TerraformContext.AccountName)' in resource group '$($TerraformContext.ResourceGroup)'."
        }

        $items = $modelsJson | ConvertFrom-Json

        return $items |
            Where-Object { $_ -ne $null } |
            ForEach-Object {
                [pscustomobject]@{
                    Model    = if (-not [string]::IsNullOrWhiteSpace($_.name)) { $_.name } else { $null }
                    Version  = if (-not [string]::IsNullOrWhiteSpace($_.version)) { $_.version } else { 'n/a' }
                    Provider = if (-not [string]::IsNullOrWhiteSpace($_.publisherName)) { $_.publisherName } elseif (-not [string]::IsNullOrWhiteSpace($_.provider)) { $_.provider } else { 'Azure' }
                    Source   = 'Azure Foundry'
                }
            } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_.Model) }
    }
    #endregion

    #region OUTPUT
    # Functions for combining and presenting model inventory output.
    function New-CombinedModelCatalog {
        # Merge provider catalogs and sort by model, version, and provider.
        param(
            [Parameter(Mandatory)]
            [array]$GitHubApiModels,

            [Parameter(Mandatory)]
            [array]$GitHubCliModels,

            [Parameter(Mandatory)]
            [array]$AzureModels
        )

        return @($GitHubApiModels + $GitHubCliModels + $AzureModels) |
            Sort-Object -Property Model, Version, Provider, Source -Unique
    }

    function Show-ModelCatalog {
        # Render the merged model catalog to the console.
        param(
            [Parameter(Mandatory)]
            [array]$Catalog
        )

        if ($Catalog.Count -eq 0) {
            Write-Warning 'No models were returned from GitHub or Azure Foundry.'
            return
        }

        $Catalog |
            Select-Object Model, Version, Provider, Source |
            Format-Table -AutoSize
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
