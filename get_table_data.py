from storage_services.dynamodb_operations import get_company
from core.config import DYNAMODB_COMPANIES_TABLE
import boto3
import json
import dotenv
import os
dotenv.load_dotenv()

AWS_REGION = os.getenv("AWS_REGION")
AWS_ACCESS_KEY_ID = os.getenv("AWS_ACCESS_KEY_ID")
AWS_SECRET_ACCESS_KEY = os.getenv("AWS_SECRET_ACCESS_KEY")

prospect_company_name = "juniper"
seller_company_name = "icicilombard"

seller_company_data = get_company(DYNAMODB_COMPANIES_TABLE, seller_company_name, "seller")
prospect_company_data = get_company(DYNAMODB_COMPANIES_TABLE, prospect_company_name, "prospect")

seller_company_context = seller_company_data['structured_company_data']
prospect_company_context = prospect_company_data['structured_company_data']

client = boto3.client("bedrock-runtime", region_name=AWS_REGION, aws_access_key_id=AWS_ACCESS_KEY_ID, aws_secret_access_key=AWS_SECRET_ACCESS_KEY)

# Read prompt from file
with open("prompts/presentation_slides_generator_prompt.txt", "r") as f:
    prompt = f.read()

prompt = prompt.replace("{{seller_company_context}}", str(seller_company_context))
prompt = prompt.replace("{{prospect_company_context}}", str(prospect_company_context))
# prompt = prompt.replace("{{slides_template}}", slides_template)

response = client.invoke_model_with_response_stream(
    # modelId="anthropic.claude-3-haiku-20240307-v1:0", 
    modelId="us.anthropic.claude-haiku-4-5-20251001-v1:0", 
    body=json.dumps({ 
        "anthropic_version": "bedrock-2023-05-31", 
        "max_tokens": 2048, 
        "messages": [
            { "role": "user", "content": prompt }
        ] 
    }) 
)  

# Collect and process output
output_content = ""
for event in response["body"]: 
    chunk = json.loads(event["chunk"]["bytes"]) 
    if chunk["type"] == "content_block_delta": 
        text = chunk["delta"].get("text", "")
        output_content += text

# Save output to file
with open("sample_content.md", "w", encoding="utf-8") as output_file:
    output_file.write(output_content)

# print(output_content)