# DocumentWriter

AI-powered Word document revision tool that analyzes `.docx` files section by section and rewrites content into polished, production-ready documentation.

## Overview

DocumentWriter reads a Word document, extracts its section outline and body text via a Python helper, and sends each section to a chat completions model. The AI applies a senior technical writer persona — following Microsoft Writing Style Guide conventions — and returns revised, client-facing prose. The script then rewrites each section body in the document with the revised prose.

Two inference providers are supported:

| Provider | Endpoint | Auth | Use Case |
|----------|----------|------|----------|
| **GitHub** | GitHub Models (`models.inference.ai.azure.com`) | GitHub CLI token | Free tier, rate-limited |
| **Azure** | Azure OpenAI via AI Foundry | Azure CLI Entra ID token | Production-grade throughput |

## Project Structure

```
DocumentWriter/
├── scripts/
│   └── Invoke-DocumentCommentary.ps1   # Main PowerShell orchestrator
├── src/python/
│   └── docx_processor.py               # Python helper: extract & revise .docx
├── prompts/
│   └── SystemPrompt.md                 # System prompt for the AI model
├── terraform/                          # Azure AI Foundry infrastructure (optional)
│   ├── main.tf
│   ├── variables.tf
│   ├── outputs.tf
│   └── providers.tf
│   ├── msdn.tfvars
│   ├── payg.tfvars
│   └── Invoke-TerraformWorkspace.ps1   # Workspace + subscription helper
├── input/                              # Place source .docx files here
├── output/                             # Revised documents are written here
└── requirements.txt                    # Python dependencies
```

## Prerequisites

- **PowerShell 7+**
- **Python 3.10+** with `pip`
- For the **GitHub** provider: [GitHub CLI](https://cli.github.com/) (`gh`), authenticated via `gh auth login`
- For the **Azure** provider: [Azure CLI](https://learn.microsoft.com/cli/azure/install-azure-cli), authenticated via `az login`

## Setup

### 1. Clone the repository

```powershell
git clone <repo-url>
cd DocumentWriter
```

### 2. Create and activate a Python virtual environment

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

### 3. Install Python dependencies

```powershell
pip install -r requirements.txt
```

Dependencies:

| Package | Minimum Version |
|---------|----------------|
| `python-docx` | 1.1.0 |
| `requests` | 2.31.0 |

> The script will attempt to install dependencies automatically on first run if they are not found.

## Usage

### GitHub Provider (free tier)

```powershell
.\scripts\Invoke-DocumentCommentary.ps1 `
    -DocumentPath "input\MyDocument.docx" `
    -Provider GitHub
```

### Azure OpenAI Provider (production)

```powershell
$endpoint   = terraform -chdir=terraform output -raw azure_openai_endpoint
$deployment = terraform -chdir=terraform output -raw model_deployment_name

.\scripts\Invoke-DocumentCommentary.ps1 `
    -DocumentPath "input\MyDocument.docx" `
    -Provider Azure `
    -Subscription payg `
    -AzureEndpoint $endpoint `
    -AzureDeployment $deployment
```

### Parameters

| Parameter | Required | Default | Description |
|-----------|----------|---------|-------------|
| `-DocumentPath` | Yes | — | Path to the source `.docx` file |
| `-Provider` | No | `GitHub` | `GitHub` or `Azure` |
| `-Subscription` | No | `msdn` | Azure target subscription key: `msdn` or `payg` |
| `-Model` | No | `openai/gpt-4o` | Model name (GitHub provider) |
| `-AzureEndpoint` | Azure only | — | Azure OpenAI endpoint URL |
| `-AzureDeployment` | Azure only | — | Azure OpenAI deployment name |
| `-OutputPath` | No | Auto-generated | Path for the revised output `.docx` |
| `-SystemPromptFile` | No | `prompts/SystemPrompt.md` | Path to a custom system prompt |
| `-RequestDelaySeconds` | No | `30` | Delay between API calls (rate-limit management) |
| `-PythonExe` | No | `python` | Python executable path |

### Output

The revised document is written to the path specified by `-OutputPath`. If omitted, the output file is placed alongside the source document with a timestamp suffix.

## System Prompt

The AI behavior is governed by [prompts/SystemPrompt.md](prompts/SystemPrompt.md). The prompt instructs the model to act as a senior Azure technical writer, producing concise prose from the extracted document text and following Microsoft Writing Style Guide conventions.

Customize this file to change the revision style, tone, or domain focus.

## Azure Infrastructure (Optional)

The `terraform/` directory provisions an Azure AI Foundry account and supports isolated state per subscription using Terraform workspaces.

### Terraform Variables

| Variable | Default | Description |
|----------|---------|-------------|
| `location` | `eastus` | Azure region |
| `model_name` | `gpt-4o` | OpenAI model to deploy |
| `model_version` | `2024-11-20` | Model version |
| `model_capacity` | `10` | Throughput in thousands of tokens per minute |
| `ai_services_sku` | `S0` | AI Services account SKU |

Subscription IDs are managed in per-subscription variable files:

| File | Subscription | ID |
|------|--------------|----|
| `msdn.tfvars` | `sub-tatelab-msdn` | `e091f6e7-031a-4924-97bb-8c983ca5d21a` |
| `payg.tfvars` | `sub-tatelab-payg` | `e6ad7655-b3ba-4324-8361-fcfdc59973a5` |

### Deploy

```powershell
cd terraform
.\Invoke-TerraformWorkspace.ps1 -Subscription msdn -Action apply -Initialize

# Deploy to PAYG subscription workspace/state
.\Invoke-TerraformWorkspace.ps1 -Subscription payg -Action apply -Initialize
```

### Workspace Commands (manual)

```powershell
cd terraform
terraform init

# MSDN state (dedicated workspace)
terraform workspace new msdn    # run once
terraform workspace select msdn
terraform plan  -var-file="msdn.tfvars"
terraform apply -var-file="msdn.tfvars"

# PAYG state (separate workspace)
terraform workspace new payg    # run once
terraform workspace select payg
terraform plan  -var-file="payg.tfvars"
terraform apply -var-file="payg.tfvars"
```

After deployment, use the Terraform outputs to populate the Azure provider parameters:

```powershell
$endpoint   = terraform output -raw azure_openai_endpoint
$deployment = terraform output -raw model_deployment_name

.\scripts\Invoke-DocumentCommentary.ps1 `
    -DocumentPath "input\MyDocument.docx" `
    -Provider Azure `
    -AzureEndpoint $endpoint `
    -AzureDeployment $deployment
```

## How It Works

1. **Extract** — `docx_processor.py extract` walks the document paragraphs and maps heading levels to section boundaries in a text-only JSON payload.
2. **Analyze** — The PowerShell script sends each section text to the configured chat completions endpoint with the system prompt. Requests are paced by `-RequestDelaySeconds` to stay within rate limits.
3. **Revise** — `docx_processor.py revise` replaces each section body with the AI-rewritten prose while preserving section headings and document structure.
4. **Cleanup** — Temporary extraction artifacts are removed after the revised document is saved.

## Authentication

### GitHub Provider

The script calls `gh auth token` to obtain a token automatically. If no active session exists, it launches `gh auth login` interactively.

### Azure Provider

The script calls `az account get-access-token --resource https://cognitiveservices.azure.com` to obtain a short-lived Entra ID bearer token. Run `az login` beforehand if not already authenticated.

## Author

Greg Tate
