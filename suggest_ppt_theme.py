import json
import os
from storage_services.bedrock_operations import invoke_model, extract_content_from_response
from core.config import BEDROCK_MODEL_ID

# Ensure output directory exists
os.makedirs("outputs/theme_suggestions", exist_ok=True)


def load_prompts():
    """Load prompt templates from the prompts folder.
    
    Returns:
        Tuple of (system_prompt, user_prompt_template)
    """
    prompt_file = "prompts/ppt_theme_suggestion_prompt.txt"
    try:
        with open(prompt_file, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Extract system prompt (between "## System Prompt" and "---")
        system_start = content.find("## System Prompt\n") + len("## System Prompt\n")
        system_end = content.find("\n---", system_start)
        system_prompt = content[system_start:system_end].strip()
        
        # Extract the simpler user prompt template (between "## User Prompt Template" and the first "---" after it)
        user_start = content.find("## User Prompt Template\n") + len("## User Prompt Template\n")
        user_end = content.find("\n---", user_start)
        user_section = content[user_start:user_end].strip()
        
        # Extract just the text lines without JSON blocks
        user_lines = []
        for line in user_section.split('\n'):
            if not line.startswith('```'):
                user_lines.append(line)
        
        user_template = '\n'.join(user_lines).strip()
        
        if system_prompt and user_template:
            return system_prompt, user_template
        else:
            raise ValueError("Could not parse prompts from file")
    
    except FileNotFoundError:
        print(f"[ERROR] Prompt file not found: {prompt_file}")
        raise
    except Exception as e:
        print(f"[ERROR] Error loading prompts: {e}")
        raise


# Load prompts when module is imported
THEME_SUGGESTION_SYSTEM_PROMPT, THEME_ANALYSIS_PROMPT_TEMPLATE = load_prompts()


def load_ppt_analysis(analysis_file: str) -> dict:
    """Load the analyzed PPT design system.
    
    Args:
        analysis_file: Path to the PPT analysis JSON file
    
    Returns:
        Dictionary containing the PPT analysis
    """
    try:
        with open(analysis_file, 'r', encoding='utf-8') as f:
            analysis = json.load(f)
        print(f"[OK] Loaded PPT analysis from {analysis_file}")
        return analysis
    except FileNotFoundError:
        print(f"[ERROR] PPT analysis file not found: {analysis_file}")
        raise
    except json.JSONDecodeError:
        print(f"[ERROR] Failed to parse JSON from {analysis_file}")
        raise


def generate_theme_suggestion(
    ppt_analysis: dict,
    model_id: str = BEDROCK_MODEL_ID,
    temperature: float = 0.7,
) -> str:
    """Generate theme suggestions based on PPT analysis.
    
    Args:
        ppt_analysis: Dictionary containing the analyzed PPT
        model_id: Bedrock model ID to use
        temperature: Temperature for model sampling
    
    Returns:
        Theme suggestion text from the model
    """
    # Format the analysis as JSON string for the prompt
    analysis_json = json.dumps(ppt_analysis, indent=2)
    
    # Create the user prompt with the analysis using replace instead of format (to avoid {} conflicts)
    user_prompt = THEME_ANALYSIS_PROMPT_TEMPLATE.replace(
        "{ppt_analysis}", analysis_json
    )
    
    # Prepare messages for the model - content must be a list of content blocks
    messages = [
        {"role": "user", "content": [{"text": user_prompt}]}
    ]
    
    print(f"Invoking {model_id} to generate theme suggestions...")
    
    try:
        # Invoke the model
        response = invoke_model(
            model_id=model_id,
            messages=messages,
            system_prompt=THEME_SUGGESTION_SYSTEM_PROMPT,
            max_tokens=4096,
            temperature=temperature,
        )
        
        # Extract the content from response
        theme_suggestion = extract_content_from_response(response)
        
        if theme_suggestion:
            print("[OK] Theme suggestion generated successfully")
            return theme_suggestion
        else:
            print("[ERROR] Failed to extract content from model response")
            return None
    except Exception as e:
        print(f"[ERROR] Error generating theme suggestion: {e}")
        raise


def save_theme_suggestion(
    theme_suggestion: str,
    output_file: str,
) -> str:
    """Save theme suggestion to a file (markdown or JSON).
    
    Args:
        theme_suggestion: The theme suggestion text (markdown or JSON)
        output_file: Path to save the suggestion
    
    Returns:
        Path to the saved file
    """
    try:
        # Save content as provided (markdown or JSON)
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(theme_suggestion)
        
        file_format = "Markdown" if output_file.endswith('.md') else "JSON"
        print(f"[OK] Theme suggestion saved as {file_format} to {output_file}")
        return output_file
    except Exception as e:
        print(f"[ERROR] Error saving theme suggestion: {e}")
        raise


def create_theme_guide(
    ppt_file: str = "templates/ppt-template.pptx",
    analysis_file: str = None,
    output_file: str = None,
    model_id: str = BEDROCK_MODEL_ID,
) -> dict:
    """Complete workflow: Analyze PPT and generate theme suggestion.
    
    Args:
        ppt_file: Path to the PPT file
        analysis_file: Path to the PPT analysis JSON (auto-generated if not provided)
        output_file: Path to save theme suggestion (auto-generated if not provided)
        model_id: Bedrock model ID to use
    
    Returns:
        Dictionary containing the analysis and suggestion
    """
    # Auto-generate analysis file path if not provided
    if analysis_file is None:
        ppt_name = ppt_file.split('/')[-1].split('.')[0]
        analysis_file = f"outputs/ppt_analysis/ppt_detailed_analysis_{ppt_name}.json"
    
    # Auto-generate output file path if not provided
    if output_file is None:
        ppt_name = ppt_file.split('/')[-1].split('.')[0]
        output_file = f"outputs/theme_suggestions/theme_suggestion_{ppt_name}.md"
    
    print("=" * 60)
    print("PPT Theme Suggestion Generator")
    print("=" * 60)
    
    # Load PPT analysis
    ppt_analysis = load_ppt_analysis(analysis_file)
    
    # Generate theme suggestion
    theme_suggestion = generate_theme_suggestion(ppt_analysis, model_id)
    
    if theme_suggestion is None:
        return {"error": "Failed to generate theme suggestion"}
    
    # Save theme suggestion
    saved_file = save_theme_suggestion(theme_suggestion, output_file)
    
    result = {
        "status": "success",
        "analysis_file": analysis_file,
        "suggestion_file": saved_file,
        "ppt_metadata": ppt_analysis.get("metadata", {}),
        "theme_suggestion_preview": theme_suggestion[:500] + "..." if len(theme_suggestion) > 500 else theme_suggestion,
    }
    
    print("=" * 60)
    print("Summary:")
    print(f"  Input PPT: {ppt_file}")
    print(f"  Analysis: {analysis_file}")
    print(f"  Theme Suggestion: {saved_file}")
    print("=" * 60)
    
    return result


if __name__ == "__main__":
    # Example usage
    result = create_theme_guide()
    print("\nResult:", json.dumps(result, indent=2, default=str))
