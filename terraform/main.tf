# -------------------------------------------------------------------------
# Program: main.tf
# Description: Deploy Azure AI Foundry account with GPT-4o for
#              production-grade document revision inference
# Context: DocumentWriter project - Azure AI Foundry for production-grade inference
# Author: Greg Tate
# Date: 2026-03-02
# -------------------------------------------------------------------------

# =========================================================================
# Local values
# =========================================================================

locals {
  resource_group_name = "project-docwriter-tf"

  common_tags = {
    Environment      = "Production"
    Project          = "DocumentWriter"
    Purpose          = "AI Document Revision"
    Owner            = var.owner
    DateCreated      = var.date_created
    DeploymentMethod = "Terraform"
    Workspace        = terraform.workspace
  }
}

# Random suffix for globally unique resource names
resource "random_string" "suffix" {
  length  = 4
  upper   = false
  special = false
}

# =========================================================================
# Resource Group
# =========================================================================

resource "azurerm_resource_group" "main" {
  name     = local.resource_group_name
  location = var.location
  tags     = local.common_tags

  lifecycle {
    ignore_changes = [
      tags,
    ]
  }
}

# =========================================================================
# Azure AI Foundry Account (Cognitive Services - AIServices kind)
# =========================================================================

# AI Services account for production-grade GPT-4o inference with vision
resource "azurerm_cognitive_account" "ai_foundry" {
  name                = "cog-docwriter-${random_string.suffix.result}"
  location            = azurerm_resource_group.main.location
  resource_group_name = azurerm_resource_group.main.name

  kind                  = "AIServices"
  sku_name              = var.ai_services_sku
  custom_subdomain_name = "cog-docwriter-${random_string.suffix.result}"

  # Enable public network access for API calls
  public_network_access_enabled = true

  # System-assigned managed identity
  identity {
    type = "SystemAssigned"
  }

  tags = local.common_tags

  lifecycle {
    ignore_changes = [
      tags,
    ]
  }
}

# =========================================================================
# Model Deployment (GPT-4o for document commentary with vision support)
# =========================================================================

# Deploy GPT-4o with GlobalStandard SKU for production-grade throughput
resource "azurerm_cognitive_deployment" "model" {
  count = var.deploy_model ? 1 : 0

  name                 = var.model_name
  cognitive_account_id = azurerm_cognitive_account.ai_foundry.id

  sku {
    name     = "GlobalStandard"
    capacity = var.model_capacity
  }

  model {
    format  = "OpenAI"
    name    = var.model_name
    version = var.model_version
  }

  depends_on = [azurerm_cognitive_account.ai_foundry]
}

# =========================================================================
# RBAC: Grant deployer Cognitive Services User for inference
# =========================================================================

# Get current client config for role assignments
data "azurerm_client_config" "current" {}

# Grant Cognitive Services User on the account for model inference
resource "azurerm_role_assignment" "deployer_cognitive_user" {
  scope                = azurerm_cognitive_account.ai_foundry.id
  role_definition_name = "Cognitive Services User"
  principal_id         = data.azurerm_client_config.current.object_id
  description          = "Deployer access for Azure OpenAI chat completions"
}
