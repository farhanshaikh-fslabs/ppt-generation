"""
PPT Presentation Generator

Converts the generated JSON slide structure into a actual PowerPoint presentation
using the python-pptx library. Handles all slide types and design specifications
from the JSON schema.
"""

import json
import os
import re
from typing import Dict, List, Optional, Tuple
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor


def parse_color(hex_color: str) -> RGBColor:
    """Convert hex color code to RGBColor.
    
    Args:
        hex_color: Hex color code (e.g., '#02428E')
    
    Returns:
        RGBColor object
    """
    hex_color = hex_color.lstrip('#')
    return RGBColor(int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16))


def load_slide_json(json_file: str) -> dict:
    """Load and parse the slide JSON file.
    
    Args:
        json_file: Path to the JSON file
    
    Returns:
        Parsed JSON dictionary
    """
    try:
        with open(json_file, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Strip markdown code block wrapper if present
        if content.strip().startswith('```'):
            # Remove ```json at start
            content = re.sub(r'^```json\n', '', content.strip())
            # Remove ``` at end
            content = re.sub(r'\n```$', '', content)
        
        slides_data = json.loads(content)
        print(f"[OK] Loaded slide JSON from {json_file}")
        return slides_data
    except FileNotFoundError:
        print(f"[ERROR] Slide JSON file not found: {json_file}")
        raise
    except json.JSONDecodeError as e:
        print(f"[ERROR] Failed to parse JSON: {e}")
        raise


def create_presentation(slides_data: dict, output_file: str) -> str:
    """Create PowerPoint presentation from slide JSON structure.
    
    Args:
        slides_data: Parsed slide JSON data
        output_file: Path to save the PowerPoint file
    
    Returns:
        Path to the created presentation
    """
    # Create presentation
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)
    
    # Get metadata and design system
    metadata = slides_data.get('presentation_metadata', {})
    design_system = slides_data.get('design_system_reference', {})
    slides_list = slides_data.get('slides', [])
    
    print(f"\nCreating presentation: {metadata.get('title', 'Untitled')}")
    print(f"Total slides: {len(slides_list)}")
    
    # Create each slide
    for slide_data in slides_list:
        slide_number = slide_data.get('slide_number', 0)
        slide_type = slide_data.get('slide_type', 'content')
        
        print(f"  Creating slide {slide_number}: {slide_type}")
        
        if slide_type == 'title':
            _create_title_slide(prs, slide_data, design_system)
        elif slide_type == 'content':
            _create_content_slide(prs, slide_data, design_system)
        elif slide_type == 'two_column':
            _create_two_column_slide(prs, slide_data, design_system)
        elif slide_type == 'image_text':
            _create_image_text_slide(prs, slide_data, design_system)
        elif slide_type == 'data_chart':
            _create_data_chart_slide(prs, slide_data, design_system)
        elif slide_type == 'centered_content':
            _create_centered_slide(prs, slide_data, design_system)
        elif slide_type == 'comparison':
            _create_comparison_slide(prs, slide_data, design_system)
        elif slide_type == 'closing':
            _create_closing_slide(prs, slide_data, design_system)
        else:
            print(f"    [WARNING] Unknown slide type: {slide_type}")
    
    # Save presentation
    try:
        prs.save(output_file)
        print(f"\n[OK] Presentation saved to {output_file}")
        return output_file
    except Exception as e:
        print(f"[ERROR] Failed to save presentation: {e}")
        raise


def _create_title_slide(prs: Presentation, slide_data: dict, design_system: dict):
    """Create a title/hero slide."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout
    background = slide.background
    fill = background.fill
    
    # Set gradient background
    design_notes = slide_data.get('design_notes', {})
    gradient_colors = design_notes.get('gradient_colors', [])
    
    if gradient_colors and len(gradient_colors) >= 2:
        fill.gradient()
        fill.gradient_angle = 45.0
        fill.gradient_stops[0].color.rgb = parse_color(gradient_colors[0])
        fill.gradient_stops[1].color.rgb = parse_color(gradient_colors[1])
    else:
        fill.solid()
        fill.fore_color.rgb = parse_color(design_notes.get('background', '#02428E'))
    
    # Add title
    left = Inches(0.5)
    top = Inches(2.5)
    width = Inches(9)
    height = Inches(1.5)
    title_box = slide.shapes.add_textbox(left, top, width, height)
    title_frame = title_box.text_frame
    title_frame.word_wrap = True
    
    title_para = title_frame.paragraphs[0]
    title_para.text = slide_data.get('title', '')
    title_para.font.size = Pt(54)
    title_para.font.bold = True
    title_para.font.color.rgb = parse_color(design_notes.get('text_color', '#FFFFFF'))
    title_para.alignment = PP_ALIGN.CENTER
    
    # Add subtitle
    subtitle_box = slide.shapes.add_textbox(Inches(0.5), Inches(4.2), Inches(9), Inches(1.5))
    subtitle_frame = subtitle_box.text_frame
    subtitle_frame.word_wrap = True
    
    subtitle_para = subtitle_frame.paragraphs[0]
    subtitle_para.text = slide_data.get('subtitle', '')
    subtitle_para.font.size = Pt(24)
    subtitle_para.font.color.rgb = parse_color(design_notes.get('text_color', '#FFFFFF'))
    subtitle_para.alignment = PP_ALIGN.CENTER


def _create_content_slide(prs: Presentation, slide_data: dict, design_system: dict):
    """Create a standard content slide with bullets."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout
    background = slide.background
    fill = background.fill
    fill.solid()
    
    design_notes = slide_data.get('design_notes', {})
    fill.fore_color.rgb = parse_color(design_notes.get('background', '#FFFFFF'))
    
    # Add title with accent line
    title_top = Inches(0.35)
    title_left = Inches(0.5)
    
    title_box = slide.shapes.add_textbox(title_left, title_top, Inches(9), Inches(0.6))
    title_frame = title_box.text_frame
    title_para = title_frame.paragraphs[0]
    title_para.text = slide_data.get('title', '')
    title_para.font.size = Pt(24)
    title_para.font.bold = True
    title_para.font.color.rgb = parse_color(design_system.get('header_color', '#02428E'))
    
    # Add accent line below title
    if design_notes.get('accent_line') == 'yes':
        line = slide.shapes.add_shape(1, title_left, Inches(1.05), Inches(9), Inches(0.05))
        line.fill.solid()
        line.fill.fore_color.rgb = parse_color(design_system.get('accent_color', '#F26633'))
        line.line.color.rgb = parse_color(design_system.get('accent_color', '#F26633'))
    
    # Add content bullets
    content = slide_data.get('content', {})
    bullets = content.get('bullets', [])
    
    bullet_top = Inches(1.3)
    bullet_left = Inches(0.7)
    bullet_width = Inches(8.6)
    
    bullet_box = slide.shapes.add_textbox(bullet_left, bullet_top, bullet_width, Inches(5.5))
    text_frame = bullet_box.text_frame
    text_frame.word_wrap = True
    
    for idx, bullet in enumerate(bullets):
        if idx > 0:
            text_frame.add_paragraph()
        
        para = text_frame.paragraphs[idx]
        para.text = bullet.get('text', '')
        para.level = bullet.get('level', 1) - 1
        para.font.size = Pt(15 if para.level == 0 else 13)
        para.font.bold = bullet.get('emphasis', 'normal') == 'bold'
        
        if bullet.get('accent') == 'orange':
            para.font.color.rgb = parse_color(design_system.get('accent_color', '#F26633'))
        else:
            para.font.color.rgb = parse_color(design_system.get('text_color', '#000000'))
        
        para.space_before = Pt(6)
        para.space_after = Pt(6)
    
    # Add callout box if present
    callout = design_notes.get('callout_box')
    if callout and callout.get('text'):
        callout_left = Inches(0.7)
        callout_top = Inches(6.0)
        callout_width = Inches(8.6)
        callout_height = Inches(0.9)
        
        box = slide.shapes.add_shape(1, callout_left, callout_top, callout_width, callout_height)
        box.fill.solid()
        box.fill.fore_color.rgb = parse_color(callout.get('background_color', '#F26633'))
        box.line.color.rgb = parse_color(callout.get('background_color', '#F26633'))
        
        box_text = box.text_frame
        box_text.word_wrap = True
        box_text.vertical_anchor = MSO_ANCHOR.MIDDLE
        
        callout_para = box_text.paragraphs[0]
        callout_para.text = callout.get('text', '')
        callout_para.font.size = Pt(14)
        callout_para.font.bold = True
        callout_para.font.color.rgb = parse_color(callout.get('text_color', '#FFFFFF'))
        callout_para.alignment = PP_ALIGN.CENTER


def _create_two_column_slide(prs: Presentation, slide_data: dict, design_system: dict):
    """Create a two-column content slide."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = parse_color('#FFFFFF')
    
    # Add title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.35), Inches(9), Inches(0.6))
    title_frame = title_box.text_frame
    title_para = title_frame.paragraphs[0]
    title_para.text = slide_data.get('title', '')
    title_para.font.size = Pt(24)
    title_para.font.bold = True
    title_para.font.color.rgb = parse_color(design_system.get('header_color', '#02428E'))
    
    # Add two columns
    columns = slide_data.get('columns', [])
    col_top = Inches(1.2)
    col_height = Inches(5.8)
    
    for col_idx, column in enumerate(columns):
        col_left = Inches(0.5 + col_idx * 4.9)
        col_width = Inches(4.6)
        
        # Column header
        header_box = slide.shapes.add_textbox(col_left, col_top, col_width, Inches(0.4))
        header_frame = header_box.text_frame
        header_para = header_frame.paragraphs[0]
        header_para.text = column.get('header', '')
        header_para.font.size = Pt(18)
        header_para.font.bold = True
        
        # Column content
        content_box = slide.shapes.add_textbox(col_left, Inches(1.8), col_width, col_height)
        content_frame = content_box.text_frame
        content_frame.word_wrap = True
        
        items = column.get('content', {}).get('items', [])
        for item_idx, item in enumerate(items):
            if item_idx > 0:
                content_frame.add_paragraph()
            
            para = content_frame.paragraphs[item_idx]
            para.text = item
            para.font.size = Pt(13)
            para.space_before = Pt(4)
            para.space_after = Pt(4)


def _create_image_text_slide(prs: Presentation, slide_data: dict, design_system: dict):
    """Create a slide with image and text."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = parse_color('#FFFFFF')
    
    # Add title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.35), Inches(9), Inches(0.6))
    title_frame = title_box.text_frame
    title_para = title_frame.paragraphs[0]
    title_para.text = slide_data.get('title', '')
    title_para.font.size = Pt(24)
    title_para.font.bold = True
    title_para.font.color.rgb = parse_color(design_system.get('header_color', '#02428E'))
    
    # Placeholder for image (would need actual image file)
    image_desc_box = slide.shapes.add_textbox(Inches(0.7), Inches(1.3), Inches(4.5), Inches(5.2))
    image_frame = image_desc_box.text_frame
    image_para = image_frame.paragraphs[0]
    image_para.text = "[Image Placeholder]\n\n" + slide_data.get('image', {}).get('description', '')
    image_para.font.size = Pt(14)
    image_para.font.italic = True
    image_para.font.color.rgb = RGBColor(128, 128, 128)
    
    # Add content bullets
    content_box = slide.shapes.add_textbox(Inches(5.5), Inches(1.3), Inches(4), Inches(5.2))
    text_frame = content_box.text_frame
    text_frame.word_wrap = True
    
    bullets = slide_data.get('content', {}).get('bullets', [])
    for idx, bullet in enumerate(bullets):
        if idx > 0:
            text_frame.add_paragraph()
        
        para = text_frame.paragraphs[idx]
        para.text = bullet.get('text', '')
        para.font.size = Pt(14)
        para.font.bold = bullet.get('emphasis', 'normal') == 'bold'
        para.space_before = Pt(4)
        para.space_after = Pt(4)


def _create_data_chart_slide(prs: Presentation, slide_data: dict, design_system: dict):
    """Create a data/chart slide."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = parse_color('#FFFFFF')
    
    # Add title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.35), Inches(9), Inches(0.6))
    title_frame = title_box.text_frame
    title_para = title_frame.paragraphs[0]
    title_para.text = slide_data.get('title', '')
    title_para.font.size = Pt(24)
    title_para.font.bold = True
    title_para.font.color.rgb = parse_color(design_system.get('header_color', '#02428E'))
    
    # Chart placeholder
    chart_desc_box = slide.shapes.add_textbox(Inches(1), Inches(1.3), Inches(8), Inches(3.5))
    chart_frame = chart_desc_box.text_frame
    chart_para = chart_frame.paragraphs[0]
    
    chart_data = slide_data.get('chart', {})
    chart_para.text = f"[{chart_data.get('chart_type', 'Chart')} Chart]\n\n{chart_data.get('description', '')}"
    chart_para.font.size = Pt(16)
    chart_para.font.italic = True
    chart_para.font.color.rgb = RGBColor(128, 128, 128)
    chart_para.alignment = PP_ALIGN.CENTER
    
    # Add key insight callout
    insight = slide_data.get('key_insight', {})
    if insight and insight.get('text'):
        insight_left = Inches(1)
        insight_top = Inches(5.3)
        insight_width = Inches(8)
        insight_height = Inches(0.9)
        
        box = slide.shapes.add_shape(1, insight_left, insight_top, insight_width, insight_height)
        box.fill.solid()
        box.fill.fore_color.rgb = parse_color(insight.get('background_color', '#F26633'))
        
        box_text = box.text_frame
        box_text.word_wrap = True
        box_text.vertical_anchor = MSO_ANCHOR.MIDDLE
        
        insight_para = box_text.paragraphs[0]
        insight_para.text = insight.get('text', '')
        insight_para.font.size = Pt(16)
        insight_para.font.bold = True
        insight_para.font.color.rgb = parse_color(insight.get('text_color', '#FFFFFF'))
        insight_para.alignment = PP_ALIGN.CENTER


def _create_centered_slide(prs: Presentation, slide_data: dict, design_system: dict):
    """Create a centered content slide."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = parse_color('#FFFFFF')
    
    # Add centered title
    title_box = slide.shapes.add_textbox(Inches(1), Inches(2.5), Inches(8), Inches(1.5))
    title_frame = title_box.text_frame
    title_frame.word_wrap = True
    
    title_para = title_frame.paragraphs[0]
    title_para.text = slide_data.get('title', '')
    title_para.font.size = Pt(28)
    title_para.font.bold = True
    title_para.font.color.rgb = parse_color(design_system.get('header_color', '#02428E'))
    title_para.alignment = PP_ALIGN.CENTER
    
    # Add subtitle
    subtitle_box = slide.shapes.add_textbox(Inches(1), Inches(4.2), Inches(8), Inches(1))
    subtitle_frame = subtitle_box.text_frame
    subtitle_frame.word_wrap = True
    
    subtitle_para = subtitle_frame.paragraphs[0]
    subtitle_para.text = slide_data.get('subtitle', '')
    subtitle_para.font.size = Pt(20)
    subtitle_para.font.color.rgb = parse_color(design_system.get('text_color', '#000000'))
    subtitle_para.alignment = PP_ALIGN.CENTER


def _create_comparison_slide(prs: Presentation, slide_data: dict, design_system: dict):
    """Create a before/after comparison slide."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = parse_color('#FFFFFF')
    
    # Add title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.35), Inches(9), Inches(0.6))
    title_frame = title_box.text_frame
    title_para = title_frame.paragraphs[0]
    title_para.text = slide_data.get('title', '')
    title_para.font.size = Pt(24)
    title_para.font.bold = True
    title_para.font.color.rgb = parse_color(design_system.get('header_color', '#02428E'))
    
    # Create two comparison boxes
    columns = slide_data.get('columns', [])
    col_width = Inches(4.3)
    col_height = Inches(5.2)
    col_top = Inches(1.3)
    
    for col_idx, column in enumerate(columns):
        col_left = Inches(0.5 + col_idx * 4.8)
        
        # Background box for column
        box = slide.shapes.add_shape(1, col_left, col_top, col_width, col_height)
        box.fill.solid()
        box.fill.fore_color.rgb = parse_color(column.get('background_color', '#02428E'))
        box.line.color.rgb = parse_color(column.get('background_color', '#02428E'))
        
        # Column label
        label_box = slide.shapes.add_textbox(col_left + Inches(0.1), col_top + Inches(0.1), col_width - Inches(0.2), Inches(0.4))
        label_frame = label_box.text_frame
        label_para = label_frame.paragraphs[0]
        label_para.text = column.get('column_label', '')
        label_para.font.size = Pt(16)
        label_para.font.bold = True
        label_para.font.color.rgb = parse_color(column.get('text_color', '#FFFFFF'))
        label_para.alignment = PP_ALIGN.CENTER
        
        # Items
        items_box = slide.shapes.add_textbox(col_left + Inches(0.15), col_top + Inches(0.6), col_width - Inches(0.3), col_height - Inches(0.8))
        items_frame = items_box.text_frame
        items_frame.word_wrap = True
        
        items = column.get('items', [])
        for item_idx, item in enumerate(items):
            if item_idx > 0:
                items_frame.add_paragraph()
            
            para = items_frame.paragraphs[item_idx]
            para.text = f"• {item}"
            para.font.size = Pt(12)
            para.font.color.rgb = parse_color(column.get('text_color', '#FFFFFF'))
            para.space_before = Pt(3)
            para.space_after = Pt(3)
            para.level = 0


def _create_closing_slide(prs: Presentation, slide_data: dict, design_system: dict):
    """Create a closing/CTA slide."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    background = slide.background
    fill = background.fill
    
    # Set gradient background
    design_notes = slide_data.get('design_notes', {})
    fill.gradient()
    fill.gradient_angle = 45.0
    fill.gradient_stops[0].color.rgb = parse_color('#02428E')
    fill.gradient_stops[1].color.rgb = parse_color('#00498F')
    
    # Add title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(2), Inches(9), Inches(1))
    title_frame = title_box.text_frame
    title_frame.word_wrap = True
    
    title_para = title_frame.paragraphs[0]
    title_para.text = slide_data.get('title', '')
    title_para.font.size = Pt(54)
    title_para.font.bold = True
    title_para.font.color.rgb = parse_color('#FFFFFF')
    title_para.alignment = PP_ALIGN.CENTER
    
    # Add subtitle
    subtitle_box = slide.shapes.add_textbox(Inches(0.5), Inches(3.2), Inches(9), Inches(0.8))
    subtitle_frame = subtitle_box.text_frame
    subtitle_frame.word_wrap = True
    
    subtitle_para = subtitle_frame.paragraphs[0]
    subtitle_para.text = slide_data.get('subtitle', '')
    subtitle_para.font.size = Pt(18)
    subtitle_para.font.color.rgb = parse_color('#FFFFFF')
    subtitle_para.alignment = PP_ALIGN.CENTER
    
    # Add contact info
    contact_info = slide_data.get('contact_info', {})
    contact_box = slide.shapes.add_textbox(Inches(1), Inches(4.5), Inches(8), Inches(1.5))
    contact_frame = contact_box.text_frame
    contact_frame.word_wrap = True
    
    contact_items = [
        contact_info.get('name', ''),
        contact_info.get('email', ''),
        contact_info.get('phone', ''),
    ]
    
    for idx, contact_item in enumerate(contact_items):
        if contact_item:
            if idx > 0:
                contact_frame.add_paragraph()
            
            para = contact_frame.paragraphs[idx]
            para.text = contact_item
            para.font.size = Pt(14)
            para.font.color.rgb = parse_color('#FFFFFF')
            para.alignment = PP_ALIGN.CENTER
    
    # Add CTA button (as a shape)
    cta = slide_data.get('cta_button', {})
    if cta and cta.get('text'):
        btn_left = Inches(3)
        btn_top = Inches(6.2)
        btn_width = Inches(4)
        btn_height = Inches(0.6)
        
        btn = slide.shapes.add_shape(1, btn_left, btn_top, btn_width, btn_height)
        btn.fill.solid()
        btn.fill.fore_color.rgb = parse_color(cta.get('background_color', '#F26633'))
        btn.line.color.rgb = parse_color(cta.get('background_color', '#F26633'))
        
        btn_text = btn.text_frame
        btn_text.word_wrap = True
        btn_text.vertical_anchor = MSO_ANCHOR.MIDDLE
        
        btn_para = btn_text.paragraphs[0]
        btn_para.text = cta.get('text', '')
        btn_para.font.size = Pt(16)
        btn_para.font.bold = True
        btn_para.font.color.rgb = parse_color(cta.get('text_color', '#FFFFFF'))
        btn_para.alignment = PP_ALIGN.CENTER


def generate_presentation(
    json_file: str = None,
    output_file: str = None,
    prospect_company: str = "juniper",
) -> str:
    """Complete workflow: Generate PowerPoint from slide JSON.
    
    Args:
        json_file: Path to slide JSON file
        output_file: Path to save PowerPoint
        prospect_company: Prospect company name for auto-detection
    
    Returns:
        Path to the created PowerPoint file
    """
    # Auto-detect JSON file
    if json_file is None:
        json_file = f"outputs/generated_slides/slides_{prospect_company}.json"
    
    # Auto-generate output file
    if output_file is None:
        prospect_clean = prospect_company.lower().replace(' ', '_')
        output_file = f"outputs/presentations/presentation_{prospect_clean}.pptx"
    
    # Ensure output directory exists
    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    
    print("=" * 60)
    print("PPT Presentation Generator")
    print("=" * 60)
    
    # Load slide JSON
    slides_data = load_slide_json(json_file)
    
    # Create presentation
    created_file = create_presentation(slides_data, output_file)
    
    print("\n" + "=" * 60)
    print("Summary:")
    print(f"  Input JSON: {json_file}")
    print(f"  Output PPT: {created_file}")
    print("=" * 60)
    
    return created_file


if __name__ == "__main__":
    """
    Main execution - generates PowerPoint from slide JSON
    """
    PROSPECT_COMPANY = "juniper"
    
    result = generate_presentation(prospect_company=PROSPECT_COMPANY)
    print(f"\nPresentation created successfully: {result}")
