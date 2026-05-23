```bash
# Azure OpenAI Provisioning Setup Commands
# Configure these variables before running the commands:

SUBSCRIPTION_ID="YOUR_SUBSCRIPTION_ID"
RESOURCE_GROUP_NAME="YOUR_RESOURCE_GROUP_NAME"
LOCATION="YOUR_AZURE_REGION"                  # e.g., eastus
COGNITIVE_ACCOUNT_NAME="YOUR_OPENAI_RESOURCE_NAME"
DEPLOYMENT_NAME="YOUR_DEPLOYMENT_NAME"        # e.g., gpt-4-mini
MODEL_NAME="YOUR_MODEL_NAME"                  # e.g., gpt-4.1-mini
MODEL_VERSION="YOUR_MODEL_VERSION"            # e.g., 2025-04-14
MODEL_FORMAT="OpenAI"
SKU_NAME="S0"
SKU_CAPACITY="20"                             # 20k TPM

# 1. Set the active subscription
az account set --subscription "$SUBSCRIPTION_ID"

# 2. Register the Cognitive Services provider (if not already done)
az provider register --namespace Microsoft.CognitiveServices --wait

# 3. Create the Resource Group
az group create --name "$RESOURCE_GROUP_NAME" --location "$LOCATION"

# 4. Create the Cognitive Services Azure OpenAI Resource
az cognitiveservices account create \
  --name "$COGNITIVE_ACCOUNT_NAME" \
  --resource-group "$RESOURCE_GROUP_NAME" \
  --kind OpenAI \
  --sku "$SKU_NAME" \
  --location "$LOCATION" \
  --yes

# 5. Deploy the specified model
az cognitiveservices account deployment create \
  --name "$COGNITIVE_ACCOUNT_NAME" \
  --resource-group "$RESOURCE_GROUP_NAME" \
  --deployment-name "$DEPLOYMENT_NAME" \
  --model-name "$MODEL_NAME" \
  --model-version "$MODEL_VERSION" \
  --model-format "$MODEL_FORMAT" \
  --sku-capacity "$SKU_CAPACITY" \
  --sku-name Standard

# 6. Retrieve keys and endpoints
echo "Endpoint:"
az cognitiveservices account show \
  --name "$COGNITIVE_ACCOUNT_NAME" \
  --resource-group "$RESOURCE_GROUP_NAME" \
  --query properties.endpoint \
  --output tsv

echo "Key 1:"
az cognitiveservices account keys list \
  --name "$COGNITIVE_ACCOUNT_NAME" \
  --resource-group "$RESOURCE_GROUP_NAME" \
  --query key1 \
  --output tsv
```
