# -------------------------------------------------------------------------
# Program: outputs.tf
# Description: Output values for DocumentWriter Azure AI resources
# Context: DocumentWriter project - Azure AI Foundry for production-grade inference
# Author: Greg Tate
# Date: 2026-03-02
# -------------------------------------------------------------------------

output "resource_group_name" {
  description = "Name of the resource group"
  value       = azurerm_resource_group.main.name
}

# =========================================================================
# AI Foundry Account outputs
# =========================================================================

output "ai_foundry_name" {
  description = "AI Foundry account name"
  value       = azurerm_cognitive_account.ai_foundry.name
}

output "azure_openai_endpoint" {
  description = "Azure OpenAI endpoint for chat completions (use with -AzureEndpoint parameter)"
  value       = azurerm_cognitive_account.ai_foundry.endpoint
}

# =========================================================================
# Model Deployment outputs
# =========================================================================

output "model_deployment_name" {
  description = "Deployed model name (use with -AzureDeployment parameter)"
  value       = try(azurerm_cognitive_deployment.model[0].name, null)
}

# =========================================================================
# Script invocation example
# =========================================================================

output "script_usage" {
  description = "Example command to run Invoke-DocumentRevision with Azure provider"
  value       = var.deploy_model ? "Invoke-DocumentRevision.ps1 -DocumentPath <file.docx> -Provider Azure -AzureEndpoint '${azurerm_cognitive_account.ai_foundry.endpoint}' -AzureDeployment '${azurerm_cognitive_deployment.model[0].name}'" : "Model deployment disabled. Re-run with deploy_model=true and a quota-enabled model, then use: Invoke-DocumentRevision.ps1 -DocumentPath <file.docx> -Provider Azure -AzureEndpoint '${azurerm_cognitive_account.ai_foundry.endpoint}' -AzureDeployment '<deployment-name>'"
}
