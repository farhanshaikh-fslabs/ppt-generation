"""
PPT Slide Content Generator

This script orchestrates the complete workflow:
1. Load theme suggestion (from PPT analysis)
2. Load company data (seller and prospect)
3. Run content generation prompt through Claude
4. Save generated slides as JSON
"""

import json
import os
from storage_services.bedrock_operations import invoke_model, extract_content_from_response
from storage_services.dynamodb_operations import get_company
from core.config import BEDROCK_MODEL_ID, DYNAMODB_COMPANIES_TABLE

# Ensure output directory exists
os.makedirs("outputs/generated_slides", exist_ok=True)


def load_theme_suggestion(theme_file: str) -> str:
    """Load the theme suggestion (markdown).
    
    Args:
        theme_file: Path to the theme suggestion file
    
    Returns:
        Theme suggestion content as string
    """
    try:
        with open(theme_file, 'r', encoding='utf-8') as f:
            theme_content = f.read()
        print(f"[OK] Loaded theme suggestion from {theme_file}")
        return theme_content
    except FileNotFoundError:
        print(f"[ERROR] Theme file not found: {theme_file}")
        raise


def load_generator_prompt(prompt_file: str) -> str:
    """Load the slides generator prompt template.
    
    Args:
        prompt_file: Path to the prompt template file
    
    Returns:
        Prompt template as string
    """
    try:
        with open(prompt_file, 'r', encoding='utf-8') as f:
            prompt_content = f.read()
        print(f"[OK] Loaded generator prompt from {prompt_file}")
        return prompt_content
    except FileNotFoundError:
        print(f"[ERROR] Prompt file not found: {prompt_file}")
        raise


def get_company_data(
    companies_table: str,
    seller_name: str,
    prospect_name: str,
) -> tuple:
    """Fetch seller and prospect company data from DynamoDB.
    
    Args:
        companies_table: DynamoDB table name
        seller_name: Seller company name
        prospect_name: Prospect company name
    
    Returns:
        Tuple of (seller_data, prospect_data)
    """
    try:
        seller_data = get_company(companies_table, seller_name, "seller")
        prospect_data = get_company(companies_table, prospect_name, "prospect")
        
        seller_context = seller_data.get('structured_company_data', {})
        prospect_context = prospect_data.get('structured_company_data', {})
        
        print(f"[OK] Loaded seller company data: {seller_name}")
        print(f"[OK] Loaded prospect company data: {prospect_name}")
        
        return seller_context, prospect_context
    except Exception as e:
        print(f"[ERROR] Error fetching company data: {e}")
        raise


def generate_slide_content(
    seller_context: dict,
    prospect_context: dict,
    theme_content: str,
    prompt_template: str,
    model_id: str = BEDROCK_MODEL_ID,
    temperature: float = 0.7,
) -> str:
    """Generate slide content using Claude and the presentation prompt.
    
    Args:
        seller_context: Seller company structured data
        prospect_context: Prospect company structured data
        theme_content: Theme suggestion content (markdown)
        prompt_template: Generator prompt template
        model_id: Bedrock model ID
        temperature: Model temperature for sampling
    
    Returns:
        Generated slide content (JSON)
    """
    # Prepare the prompt by replacing placeholders
    # Use default=str to handle non-serializable objects (like datetime)
    user_prompt = prompt_template.replace(
        "{{seller_company_context}}", json.dumps(seller_context, indent=2, default=str)
    ).replace(
        "{{prospect_company_context}}", json.dumps(prospect_context, indent=2, default=str)
    ).replace(
        "{{design_theme}}", theme_content
    )
    
    # Prepare messages for the model
    messages = [
        {"role": "user", "content": [{"text": user_prompt}]}
    ]
    
    print(f"\nInvoking {model_id} to generate slide content...")
    print("-" * 60)
    
    try:
        # Invoke the model
        response = invoke_model(
            model_id=model_id,
            messages=messages,
            system_prompt="""You are an expert presentation designer and B2B sales strategist. 
Your task is to generate a complete, valid JSON structure for PowerPoint slides that adheres to 
the provided design theme and schema. 

The output must be:
1. Valid JSON format
2. Strictly following the provided schema
3. Using ONLY colors from the design theme
4. Maintaining typography hierarchy from the theme
5. Including all required fields for each slide

Return ONLY the JSON object, no other text or markdown.""",
            max_tokens=10240,
            temperature=temperature,
        )
        
        # Extract the content from response
        slide_content = extract_content_from_response(response)
        
        if slide_content:
            print("[OK] Slide content generated successfully")
            return slide_content
        else:
            print("[ERROR] Failed to extract content from model response")
            return None
    except Exception as e:
        print(f"[ERROR] Error generating slide content: {e}")
        raise


def save_slide_content(
    slide_content: str,
    output_file: str,
) -> str:
    """Save generated slide content to a file.
    
    Args:
        slide_content: Generated slide JSON content
        output_file: Path to save the content
    
    Returns:
        Path to the saved file
    """
    try:
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(slide_content)
        print(f"[OK] Slide content saved to {output_file}")
        return output_file
    except Exception as e:
        print(f"[ERROR] Error saving slide content: {e}")
        raise


def generate_slides(
    seller_company_name: str = "icicilombard",
    prospect_company_name: str = "juniper",
    theme_file: str = None,
    prompt_file: str = "prompts/presentation_slides_generator_prompt.txt",
    output_file: str = None,
    companies_table: str = DYNAMODB_COMPANIES_TABLE,
    model_id: str = BEDROCK_MODEL_ID,
) -> dict:
    """Complete workflow: Generate slides from company data and theme.
    
    Args:
        seller_company_name: Name of seller company in table
        prospect_company_name: Name of prospect company in table
        theme_file: Path to theme suggestion (auto-detected if not provided)
        prompt_file: Path to generator prompt template
        output_file: Path to save generated slides (auto-generated if not provided)
        companies_table: DynamoDB table name
        model_id: Bedrock model ID to use
    
    Returns:
        Dictionary containing generation results and metadata
    """
    # Auto-detect theme file if not provided
    if theme_file is None:
        # Try prospect-specific theme first
        prospect_theme = f"outputs/theme_suggestions/theme_suggestion_{prospect_company_name}.md"
        # Fallback to template theme
        template_theme = "outputs/theme_suggestions/theme_suggestion_ppt-template.md"
        
        if os.path.exists(prospect_theme):
            theme_file = prospect_theme
        elif os.path.exists(template_theme):
            print(f"[INFO] Using template theme: {template_theme}")
            theme_file = template_theme
        else:
            print(f"[ERROR] No theme file found. Checked:")
            print(f"  - {prospect_theme}")
            print(f"  - {template_theme}")
            raise FileNotFoundError("Theme file not found")
    
    # Auto-generate output file path if not provided
    if output_file is None:
        output_file = f"outputs/generated_slides/slides_{prospect_company_name}.json"
    
    print("=" * 60)
    print("PPT Slide Content Generator")
    print("=" * 60)
    
    # Step 1: Load theme suggestion
    theme_content = load_theme_suggestion(theme_file)
    
    # Step 2: Load generator prompt template
    prompt_template = load_generator_prompt(prompt_file)
    
    # Step 3: Get company data from DynamoDB
    seller_context, prospect_context = get_company_data(
        companies_table,
        seller_company_name,
        prospect_company_name
    )
    
    # Step 4: Generate slide content
    slide_content = generate_slide_content(
        seller_context,
        prospect_context,
        theme_content,
        prompt_template,
        model_id
    )
    
    if slide_content is None:
        return {"error": "Failed to generate slide content"}
    
    # Step 5: Save slide content
    saved_file = save_slide_content(slide_content, output_file)
    
    result = {
        "status": "success",
        "seller_company": seller_company_name,
        "prospect_company": prospect_company_name,
        "theme_file": theme_file,
        "prompt_file": prompt_file,
        "output_file": saved_file,
        "model_id": model_id,
    }
    
    print("=" * 60)
    print("Summary:")
    print(f"  Seller: {seller_company_name}")
    print(f"  Prospect: {prospect_company_name}")
    print(f"  Theme: {theme_file}")
    print(f"  Output: {saved_file}")
    print("=" * 60)
    
    return result


if __name__ == "__main__":
    """
    Main execution - generates slides for specified companies
    """
    # Configuration (customize as needed)
    SELLER_COMPANY = "icicilombard"
    PROSPECT_COMPANY = "juniper"
    
    # Run generation
    result = generate_slides(
        seller_company_name=SELLER_COMPANY,
        prospect_company_name=PROSPECT_COMPANY,
    )
    
    print("\nGeneration Result:")
    print(json.dumps(result, indent=2))
    
    # Optionally parse and display the generated slides
    if result.get("status") == "success":
        try:
            with open(result["output_file"], 'r', encoding='utf-8') as f:
                slides_data = json.load(f)
            print(f"\nGenerated {slides_data['presentation_metadata']['total_slides']} slides")
            print(f"Presentation: {slides_data['presentation_metadata']['title']}")
        except Exception as e:
            print(f"\n[WARNING] Could not parse generated slides: {e}")
