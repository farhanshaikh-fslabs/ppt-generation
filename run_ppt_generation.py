"""
PPT Generation Pipeline Orchestrator

Complete end-to-end workflow for generating PowerPoint presentations:
1. Analyze PPT template and extract design system
2. Generate theme suggestion from design system analysis
3. Generate slide content JSON from company data and theme
4. Create PowerPoint presentation from slide JSON

Usage:
    python run_ppt_generation.py
    
Or programmatically:
    from run_ppt_generation import run_full_pipeline
    result = run_full_pipeline(seller="icicilombard", prospect="juniper")
"""

import json
import os
from typing import Dict, Optional

# Import pipeline steps
from analyze_ppt import extract_design_system
from suggest_ppt_theme import create_theme_guide
from generate_slide_content import generate_slides
from create_presentation import generate_presentation


def validate_ppt_template(ppt_file: str) -> bool:
    """Validate that the PPT template exists.
    
    Args:
        ppt_file: Path to PPT file
    
    Returns:
        True if file exists, False otherwise
    """
    if not os.path.exists(ppt_file):
        print(f"[ERROR] PPT template not found: {ppt_file}")
        return False
    print(f"[OK] PPT template found: {ppt_file}")
    return True


def validate_company_data(
    companies_table: str,
    seller_name: str,
    prospect_name: str,
) -> bool:
    """Validate that company data can be accessed.
    
    Args:
        companies_table: DynamoDB table name
        seller_name: Seller company name
        prospect_name: Prospect company name
    
    Returns:
        True if accessible, False otherwise
    """
    try:
        from storage_services.dynamodb_operations import get_company
        
        seller_data = get_company(companies_table, seller_name, "seller")
        prospect_data = get_company(companies_table, prospect_name, "prospect")
        
        if seller_data and prospect_data:
            print(f"[OK] Company data accessible")
            print(f"     Seller: {seller_name}")
            print(f"     Prospect: {prospect_name}")
            return True
        return False
    except Exception as e:
        print(f"[ERROR] Failed to access company data: {e}")
        return False


def run_full_pipeline(
    seller_company: str = "icicilombard",
    prospect_company: str = "juniper",
    ppt_template: str = "templates/ppt-template.pptx",
    companies_table: str = None,
    skip_steps: Optional[list] = None,
) -> Dict:
    """Run the complete PPT generation pipeline.
    
    Args:
        seller_company: Seller company name
        prospect_company: Prospect company name
        ppt_template: Path to PPT template file
        companies_table: DynamoDB companies table name
        skip_steps: List of steps to skip (e.g., ['analyze', 'theme'])
                   Steps: 'analyze', 'theme', 'slides', 'presentation'
    
    Returns:
        Dictionary containing results from all pipeline steps
    """
    if skip_steps is None:
        skip_steps = []
    
    # Initialize from config
    if companies_table is None:
        try:
            from core.config import DYNAMODB_COMPANIES_TABLE
            companies_table = DYNAMODB_COMPANIES_TABLE
        except ImportError:
            companies_table = "companies"
    
    # print("=" * 70)
    # print(" " * 15 + "PPT GENERATION PIPELINE ORCHESTRATOR")
    # print("=" * 70)
    
    results = {
        "status": "in_progress",
        "seller_company": seller_company,
        "prospect_company": prospect_company,
        "steps": {},
        "errors": []
    }
    
    # Pre-flight checks
    print("\n[PHASE 0] Pre-flight Validation")
    print("-" * 70)
    
    if not validate_ppt_template(ppt_template):
        results["status"] = "failed"
        results["errors"].append(f"PPT template not found: {ppt_template}")
        return results
    
    if not validate_company_data(companies_table, seller_company, prospect_company):
        results["status"] = "failed"
        results["errors"].append("Company data validation failed")
        return results
    
    # Step 1: Analyze PPT
    if 'analyze' not in skip_steps:
        print("\n[PHASE 1] Analyzing PPT Template")
        print("-" * 70)
        try:
            analysis_result = extract_design_system(ppt_template)
            
            analysis_file = f"outputs/ppt_analysis/ppt_detailed_analysis_{ppt_template.split('/')[-1].split('.')[0]}.json"
            
            results["steps"]["analyze_ppt"] = {
                "status": "success",
                "output_file": analysis_file,
                "metadata": analysis_result.get("metadata", {}) if isinstance(analysis_result, dict) else None,
            }
            print(f"[OK] PPT analysis completed")
        except Exception as e:
            print(f"[ERROR] PPT analysis failed: {e}")
            results["steps"]["analyze_ppt"] = {"status": "failed", "error": str(e)}
            results["errors"].append(f"PPT analysis: {e}")
            return results
    else:
        print("\n[PHASE 1] Skipped: Analyzing PPT Template")
        analysis_file = f"outputs/ppt_analysis/ppt_detailed_analysis_{ppt_template.split('/')[-1].split('.')[0]}.json"
    
    # Step 2: Generate Theme Suggestion
    if 'theme' not in skip_steps:
        print("\n[PHASE 2] Generating Theme Suggestion")
        print("-" * 70)
        try:
            theme_result = create_theme_guide()
            
            results["steps"]["generate_theme"] = {
                "status": "success",
                "output_file": theme_result.get("suggestion_file"),
                "metadata": theme_result.get("ppt_metadata", {}),
            }
            print(f"[OK] Theme suggestion completed")
        except Exception as e:
            print(f"[ERROR] Theme suggestion failed: {e}")
            results["steps"]["generate_theme"] = {"status": "failed", "error": str(e)}
            results["errors"].append(f"Theme generation: {e}")
            return results
    else:
        print("\n[PHASE 2] Skipped: Generating Theme Suggestion")
    
    # Step 3: Generate Slide Content
    if 'slides' not in skip_steps:
        print("\n[PHASE 3] Generating Slide Content")
        print("-" * 70)
        try:
            slides_result = generate_slides(
                seller_company_name=seller_company,
                prospect_company_name=prospect_company,
                companies_table=companies_table,
            )
            
            results["steps"]["generate_slides"] = {
                "status": slides_result.get("status"),
                "output_file": slides_result.get("output_file"),
            }
            print(f"[OK] Slide content generation completed")
        except Exception as e:
            print(f"[ERROR] Slide content generation failed: {e}")
            results["steps"]["generate_slides"] = {"status": "failed", "error": str(e)}
            results["errors"].append(f"Slide generation: {e}")
            return results
    else:
        print("\n[PHASE 3] Skipped: Generating Slide Content")
    
    # Step 4: Create PowerPoint Presentation
    if 'presentation' not in skip_steps:
        print("\n[PHASE 4] Creating PowerPoint Presentation")
        print("-" * 70)
        try:
            ppt_result = generate_presentation(prospect_company=prospect_company)
            
            results["steps"]["create_presentation"] = {
                "status": "success",
                "output_file": ppt_result,
            }
            print(f"[OK] PowerPoint presentation created")
        except Exception as e:
            print(f"[ERROR] PowerPoint creation failed: {e}")
            results["steps"]["create_presentation"] = {"status": "failed", "error": str(e)}
            results["errors"].append(f"Presentation creation: {e}")
            return results
    else:
        print("\n[PHASE 4] Skipped: Creating PowerPoint Presentation")
    
    # Final summary
    print("\n" + "=" * 70)
    print(" " * 25 + "PIPELINE COMPLETED")
    print("=" * 70)
    
    print("\nGenerated Files:")
    print("-" * 70)
    
    for step_name, step_result in results["steps"].items():
        if step_result.get("status") == "success" and step_result.get("output_file"):
            print(f"  ✓ {step_name:25} → {step_result['output_file']}")
        elif step_result.get("status") == "failed":
            print(f"  ✗ {step_name:25} → FAILED: {step_result.get('error', 'Unknown error')}")
    
    if results["errors"]:
        print("\nErrors Encountered:")
        print("-" * 70)
        for error in results["errors"]:
            print(f"  • {error}")
        results["status"] = "completed_with_errors"
    else:
        results["status"] = "success"
    
    print("\n" + "=" * 70)
    
    return results


def run_pipeline_interactive():
    """Run pipeline with interactive prompts for configuration."""
    print("=" * 70)
    print(" " * 15 + "PPT GENERATION PIPELINE - INTERACTIVE MODE")
    print("=" * 70)
    
    print("\nConfiguration:")
    print("-" * 70)
    
    seller = input("Seller company name [icicilombard]: ").strip() or "icicilombard"
    prospect = input("Prospect company name [juniper]: ").strip() or "juniper"
    ppt_template = input("PPT template path [templates/ppt-template.pptx]: ").strip() or "templates/ppt-template.pptx"
    
    print("\nSkip any steps? (comma-separated: analyze,theme,slides,presentation)")
    skip_input = input("Skip steps [none]: ").strip()
    skip_steps = [s.strip() for s in skip_input.split(",")] if skip_input else []
    
    print("\n")
    result = run_full_pipeline(
        seller_company=seller,
        prospect_company=prospect,
        ppt_template=ppt_template,
        skip_steps=skip_steps,
    )
    
    return result


if __name__ == "__main__":
    """
    Main execution
    """
    import sys
    
    # Check for command line arguments
    if len(sys.argv) > 1:
        if sys.argv[1] == "interactive" or sys.argv[1] == "-i":
            # Interactive mode
            result = run_pipeline_interactive()
        else:
            # Show usage
            print("Usage:")
            print("  python run_ppt_generation.py              # Run with defaults")
            print("  python run_ppt_generation.py interactive  # Interactive mode")
            sys.exit(1)
    else:
        # Run with defaults
        print("Running PPT generation pipeline with default settings...")
        print("For interactive mode, run: python run_ppt_generation.py interactive\n")
        
        result = run_full_pipeline(
            seller_company="icicilombard",
            prospect_company="juniper",
        )
    
    # Output result summary
    print("\nFinal Result:")
    print(json.dumps({
        "status": result["status"],
        "seller": result["seller_company"],
        "prospect": result["prospect_company"],
        "completed_steps": [s for s, r in result["steps"].items() if r.get("status") == "success"],
        "failed_steps": [s for s, r in result["steps"].items() if r.get("status") == "failed"],
        "output_files": {s: r.get("output_file") for s, r in result["steps"].items() if r.get("output_file")},
    }, indent=2))
