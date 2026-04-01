import os
import asyncio
import sys
from dotenv import load_dotenv
import boto3

load_dotenv()

AWS_REGION = os.getenv("AWS_REGION")

DYNAMODB_COMPANIES_TABLE = os.getenv("DYNAMODB_COMPANIES_TABLE")
DYNAMODB_COMPANIES_ACCESS_TABLE = os.getenv("DYNAMODB_COMPANIES_ACCESS_TABLE")
DYNAMODB_SIMULATIONS_TABLE = os.getenv("DYNAMODB_SIMULATIONS_TABLE")
BEDROCK_MODEL_ID = os.getenv("BEDROCK_MODEL_ID")

# S3_COMPANIES_BUCKET_NAME = os.getenv("S3_COMPANIES_BUCKET_NAME")
# S3_SALES_REPORTS_BUCKET_NAME = os.getenv("S3_SALES_REPORTS_BUCKET_NAME", "insightsphere-sales-reports")

# s3_client = boto3.client("s3", region_name=AWS_REGION)
# bedrock_runtime = boto3.client("bedrock-runtime", region_name=AWS_REGION)
dynamodb = boto3.resource("dynamodb", region_name=AWS_REGION)
bedrock_agent = boto3.client("bedrock-runtime", region_name=AWS_REGION)
# s3vectors = boto3.client('s3vectors', region_name=AWS_REGION)

# SERPER_API_KEY = os.getenv("SERPER_API_KEY")
# SERPER_API_URL = "https://google.serper.dev/search"

# # AgentCore Runtime ARNs (set after deploying agents via scripts/deploy_agents.py)
# AGENTCORE_ARN_CONTEXT_STRUCTURING = os.getenv("AGENTCORE_ARN_CONTEXT_STRUCTURING", "")
# AGENTCORE_ARN_CASE_STUDY = os.getenv("AGENTCORE_ARN_CASE_STUDY", "")
# AGENTCORE_ARN_SALES_STRATEGY = os.getenv("AGENTCORE_ARN_SALES_STRATEGY", "")
# AGENTCORE_ARN_SALES_REPORT = os.getenv("AGENTCORE_ARN_SALES_REPORT", "")
# AGENTCORE_ARN_JUDGE = os.getenv("AGENTCORE_ARN_JUDGE", "")
# AGENTCORE_ARN_SELLER = os.getenv("AGENTCORE_ARN_SELLER", "")
# AGENTCORE_ARN_PROSPECT = os.getenv("AGENTCORE_ARN_PROSPECT", "")

# agents_source = "agentcore-runtime" # or "local"
# # agents_source = "local"