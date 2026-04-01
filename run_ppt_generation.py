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
    result = run_full_pipeline(seller_company="seller-name", prospect_company="prospect-name")
"""

import json
import logging
import os
from typing import Dict, Optional

# Import pipeline steps
from analyze_ppt import extract_design_system
from suggest_ppt_theme import create_theme_guide
from generate_slide_content import generate_slides
from create_presentation import generate_presentation

logging.basicConfig(
    level=os.getenv("LOG_LEVEL", "INFO").upper(),
    format="%(asctime)s %(levelname)s %(name)s - %(message)s",
)
logger = logging.getLogger(__name__)


def validate_ppt_template(ppt_file: str) -> bool:
    """Validate that the PPT template exists.
    
    Args:
        ppt_file: Path to PPT file
    
    Returns:
        True if file exists, False otherwise
    """
    if not os.path.exists(ppt_file):
        logger.error("PPT template not found: %s", ppt_file)
        return False
    logger.info("PPT template found: %s", ppt_file)
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
            logger.info("Company data accessible")
            logger.info("Seller: %s", seller_name)
            logger.info("Prospect: %s", prospect_name)
            return True
        return False
    except Exception as e:
        logger.exception("Failed to access company data: %s", e)
        return False


def run_full_pipeline(
    seller_company: str = None,
    prospect_company: str = None,
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
    logger.info("[PHASE 0] Pre-flight validation")
    
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
        logger.info("[PHASE 1] Analyzing PPT template")
        try:
            analysis_result = extract_design_system(ppt_template)
            
            analysis_file = f"outputs/ppt_analysis/ppt_detailed_analysis_{ppt_template.split('/')[-1].split('.')[0]}.json"
            
            results["steps"]["analyze_ppt"] = {
                "status": "success",
                "output_file": analysis_file,
                "metadata": analysis_result.get("metadata", {}) if isinstance(analysis_result, dict) else None,
            }
            logger.info("PPT analysis completed")
        except Exception as e:
            logger.exception("PPT analysis failed: %s", e)
            results["steps"]["analyze_ppt"] = {"status": "failed", "error": str(e)}
            results["errors"].append(f"PPT analysis: {e}")
            return results
    else:
        logger.info("[PHASE 1] Skipped: analyzing PPT template")
        analysis_file = f"outputs/ppt_analysis/ppt_detailed_analysis_{ppt_template.split('/')[-1].split('.')[0]}.json"
    
    # Step 2: Generate Theme Suggestion
    if 'theme' not in skip_steps:
        logger.info("[PHASE 2] Generating theme suggestion")
        try:
            theme_result = create_theme_guide()
            
            results["steps"]["generate_theme"] = {
                "status": "success",
                "output_file": theme_result.get("suggestion_file"),
                "metadata": theme_result.get("ppt_metadata", {}),
            }
            logger.info("Theme suggestion completed")
        except Exception as e:
            logger.exception("Theme suggestion failed: %s", e)
            results["steps"]["generate_theme"] = {"status": "failed", "error": str(e)}
            results["errors"].append(f"Theme generation: {e}")
            return results
    else:
        logger.info("[PHASE 2] Skipped: generating theme suggestion")
    
    # Step 3: Generate Slide Content
    if 'slides' not in skip_steps:
        logger.info("[PHASE 3] Generating slide content")
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
            logger.info("Slide content generation completed")
        except Exception as e:
            logger.exception("Slide content generation failed: %s", e)
            results["steps"]["generate_slides"] = {"status": "failed", "error": str(e)}
            results["errors"].append(f"Slide generation: {e}")
            return results
    else:
        logger.info("[PHASE 3] Skipped: generating slide content")
    
    # Step 4: Create PowerPoint Presentation
    if 'presentation' not in skip_steps:
        logger.info("[PHASE 4] Creating PowerPoint presentation")
        try:
            ppt_result = generate_presentation(prospect_company=prospect_company)
            
            results["steps"]["create_presentation"] = {
                "status": "success",
                "output_file": ppt_result,
            }
            logger.info("PowerPoint presentation created")
        except Exception as e:
            logger.exception("PowerPoint creation failed: %s", e)
            results["steps"]["create_presentation"] = {"status": "failed", "error": str(e)}
            results["errors"].append(f"Presentation creation: {e}")
            return results
    else:
        logger.info("[PHASE 4] Skipped: creating PowerPoint presentation")
    
    # Final summary
    logger.info("PIPELINE COMPLETED")
    logger.info("Generated files:")
    
    for step_name, step_result in results["steps"].items():
        if step_result.get("status") == "success" and step_result.get("output_file"):
            logger.info("SUCCESS %-25s -> %s", step_name, step_result["output_file"])
        elif step_result.get("status") == "failed":
            logger.error("FAILED %-25s -> %s", step_name, step_result.get("error", "Unknown error"))
    
    if results["errors"]:
        logger.error("Errors encountered:")
        for error in results["errors"]:
            logger.error("- %s", error)
        results["status"] = "completed_with_errors"
    else:
        results["status"] = "success"
    
    return results


def run_pipeline_interactive():
    """Run pipeline with interactive prompts for configuration."""
    default_seller = os.getenv("DEFAULT_SELLER_COMPANY", "")
    default_prospect = os.getenv("DEFAULT_PROSPECT_COMPANY", "")

    logger.info("PPT generation pipeline - interactive mode")
    seller = input(f"Seller company name [{default_seller}]: ").strip() or default_seller
    prospect = input(f"Prospect company name [{default_prospect}]: ").strip() or default_prospect
    ppt_template = input("PPT template path [templates/ppt-template.pptx]: ").strip() or "templates/ppt-template.pptx"
    if not seller or not prospect:
        raise ValueError("Seller and prospect company names are required.")

    skip_input = input("Skip steps [none]: ").strip()
    skip_steps = [s.strip() for s in skip_input.split(",")] if skip_input else []

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
    import argparse

    parser = argparse.ArgumentParser(description="Run PPT generation pipeline")
    parser.add_argument("mode", nargs="?", default="", help="Use 'interactive' for prompts")
    parser.add_argument("--interactive", "-i", action="store_true", help="Run in interactive mode")
    parser.add_argument("--seller-company", default=os.getenv("DEFAULT_SELLER_COMPANY"))
    parser.add_argument("--prospect-company", default=os.getenv("DEFAULT_PROSPECT_COMPANY"))
    parser.add_argument("--ppt-template", default="templates/ppt-template.pptx")
    parser.add_argument(
        "--skip-steps",
        default="",
        help="Comma-separated list: analyze,theme,slides,presentation",
    )
    args = parser.parse_args()

    if args.interactive or args.mode == "interactive":
        result = run_pipeline_interactive()
    else:
        if not args.seller_company or not args.prospect_company:
            raise ValueError(
                "Missing seller/prospect company. Provide --seller-company and "
                "--prospect-company or set DEFAULT_SELLER_COMPANY and "
                "DEFAULT_PROSPECT_COMPANY."
            )

        result = run_full_pipeline(
            seller_company=args.seller_company,
            prospect_company=args.prospect_company,
            ppt_template=args.ppt_template,
            skip_steps=[s.strip() for s in args.skip_steps.split(",") if s.strip()],
        )
    
    logger.info("Final result:\n%s", json.dumps({
        "status": result["status"],
        "seller": result["seller_company"],
        "prospect": result["prospect_company"],
        "completed_steps": [s for s, r in result["steps"].items() if r.get("status") == "success"],
        "failed_steps": [s for s, r in result["steps"].items() if r.get("status") == "failed"],
        "output_files": {s: r.get("output_file") for s, r in result["steps"].items() if r.get("output_file")},
    }, indent=2))
