import zipfile
import xml.etree.ElementTree as ET
import json
import logging
import os
import re

logging.basicConfig(
    level=os.getenv("LOG_LEVEL", "INFO").upper(),
    format="%(asctime)s %(levelname)s %(name)s - %(message)s",
)
logger = logging.getLogger(__name__)

os.makedirs("outputs/ppt_analysis", exist_ok=True)

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
}

EMU_TO_PX = 914400 / 96


def emu_to_px(val):
    return round(int(val) / EMU_TO_PX, 2)


def parse_theme(root):
    theme = {"colors": {}, "fonts": {}}

    for clr in root.findall(".//a:clrScheme/*", NS):
        name = clr.tag.split("}")[-1]
        srgb = clr.find(".//a:srgbClr", NS)
        if srgb is not None:
            theme["colors"][name] = "#" + srgb.attrib["val"]

    latin = root.find(".//a:fontScheme//a:latin", NS)
    if latin is not None:
        theme["fonts"]["primary"] = latin.attrib.get("typeface")

    return theme


def extract_colors(root, theme_colors):
    used = set()

    for srgb in root.findall(".//a:srgbClr", NS):
        used.add("#" + srgb.attrib["val"])

    for scheme in root.findall(".//a:schemeClr", NS):
        key = scheme.attrib.get("val")
        if key in theme_colors:
            used.add(theme_colors[key])

    return list(used)


def extract_typography(root):
    fonts = set()
    sizes = set()

    for r in root.findall(".//a:r", NS):
        rpr = r.find("a:rPr", NS)
        if rpr is None:
            continue

        size = rpr.attrib.get("sz")
        if size:
            sizes.add(int(size) / 100)

        font = rpr.find("a:latin", NS)
        if font is not None:
            fonts.add(font.attrib.get("typeface"))

    return list(fonts), list(sizes)


def extract_gradients(root):
    gradients = []

    for grad in root.findall(".//a:gradFill", NS):
        stops = []

        for gs in grad.findall(".//a:gs", NS):
            color = gs.find(".//a:srgbClr", NS)
            if color is not None:
                stops.append("#" + color.attrib["val"])

        if stops:
            gradients.append(stops)

    return gradients


def extract_layout(root):
    layouts = []

    for xfrm in root.findall(".//a:xfrm", NS):
        off = xfrm.find("a:off", NS)
        ext = xfrm.find("a:ext", NS)

        if off is not None and ext is not None:
            layouts.append({
                "x": emu_to_px(off.attrib.get("x", 0)),
                "y": emu_to_px(off.attrib.get("y", 0)),
                "width": emu_to_px(ext.attrib.get("cx", 0)),
                "height": emu_to_px(ext.attrib.get("cy", 0)),
            })

    return layouts


def extract_text_formatting(text_run):
    """Extract detailed formatting from a text run"""
    rpr = text_run.find("a:rPr", NS)
    text_content = text_run.findtext("a:t", namespaces=NS) or ""
    formatting = {
        "text": text_content,
        "bold": False,
        "italic": False,
        "underline": None,
        "fontSize": None,
        "fontName": None,
        "color": None,
    }

    if rpr is not None:
        formatting["bold"] = rpr.attrib.get("b") == "1"
        formatting["italic"] = rpr.attrib.get("i") == "1"
        formatting["underline"] = rpr.attrib.get("u")

        size = rpr.attrib.get("sz")
        if size:
            formatting["fontSize"] = int(size) / 100

        # Font name
        latin = rpr.find("a:latin", NS)
        if latin is not None:
            formatting["fontName"] = latin.attrib.get("typeface")

        # Color - look for srgbClr or schemeClr
        srgb = rpr.find(".//a:srgbClr", NS)
        if srgb is not None:
            formatting["color"] = "#" + srgb.attrib.get("val", "000000")
        else:
            scheme = rpr.find(".//a:schemeClr", NS)
            if scheme is not None:
                formatting["color"] = scheme.attrib.get("val")

    return formatting


def extract_paragraph_info(paragraph):
    """Extract formatting from a paragraph"""
    ppr = paragraph.find("a:pPr", NS)
    para_info = {
        "text": "",
        "alignment": None,
        "textRuns": [],
    }

    if ppr is not None:
        para_info["alignment"] = ppr.attrib.get("algn", "left")

    # Extract text runs
    for text_run in paragraph.findall("a:r", NS):
        text_formatting = extract_text_formatting(text_run)
        para_info["textRuns"].append(text_formatting)
        para_info["text"] += text_formatting["text"]

    return para_info


def extract_shape_info(shape):
    """Extract detailed information from a shape"""
    shape_info = {
        "type": "shape",
        "position": {},
        "text": "",
        "paragraphs": [],
        "fill": None,
        "line": None,
    }

    # Position and size
    xfrm = shape.find(".//a:xfrm", NS)
    if xfrm is not None:
        off = xfrm.find("a:off", NS)
        ext = xfrm.find("a:ext", NS)
        if off is not None and ext is not None:
            shape_info["position"] = {
                "x": emu_to_px(off.attrib.get("x", 0)),
                "y": emu_to_px(off.attrib.get("y", 0)),
                "width": emu_to_px(ext.attrib.get("cx", 0)),
                "height": emu_to_px(ext.attrib.get("cy", 0)),
            }

    # Text content and formatting
    text_frame = shape.find(".//p:txBody", NS)
    if text_frame is not None:
        for paragraph in text_frame.findall("a:p", NS):
            para_info = extract_paragraph_info(paragraph)
            shape_info["paragraphs"].append(para_info)
            shape_info["text"] += para_info["text"]

    # Fill color
    solid_fill = shape.find(".//a:solidFill", NS)
    if solid_fill is not None:
        srgb = solid_fill.find(".//a:srgbClr", NS)
        if srgb is not None:
            shape_info["fill"] = "#" + srgb.attrib.get("val", "000000")

    # Line/stroke
    line = shape.find(".//a:ln", NS)
    if line is not None:
        solid_fill = line.find(".//a:solidFill", NS)
        if solid_fill is not None:
            srgb = solid_fill.find(".//a:srgbClr", NS)
            if srgb is not None:
                shape_info["line"] = "#" + srgb.attrib.get("val", "000000")

    return shape_info


def extract_slide_info(slide_xml, slide_number):
    """Extract detailed information from a single slide"""
    root = ET.fromstring(slide_xml)

    slide_info = {
        "slideNumber": slide_number,
        "shapes": [],
        "textFrames": [],
        "images": [],
        "allText": "",
    }

    # Extract all shapes (including text boxes)
    for shape in root.findall(".//p:sp", NS):
        shape_info = extract_shape_info(shape)
        slide_info["shapes"].append(shape_info)
        slide_info["allText"] += shape_info["text"]

    # Extract images
    for image in root.findall(".//p:pic", NS):
        pic_info = {"type": "image"}
        # Position
        xfrm = image.find(".//a:xfrm", NS)
        if xfrm is not None:
            off = xfrm.find("a:off", NS)
            ext = xfrm.find("a:ext", NS)
            if off is not None and ext is not None:
                pic_info["position"] = {
                    "x": emu_to_px(off.attrib.get("x", 0)),
                    "y": emu_to_px(off.attrib.get("y", 0)),
                    "width": emu_to_px(ext.attrib.get("cx", 0)),
                    "height": emu_to_px(ext.attrib.get("cy", 0)),
                }
        slide_info["images"].append(pic_info)

    return slide_info


def extract_design_system(pptx_path):
    system = {
        "metadata": {
            "filePath": pptx_path,
            "totalSlides": 0,
        },
        "designSystem": {
            "colors": set(),
            "fonts": set(),
            "fontSizes": set(),
            "gradients": [],
            "layout": []
        },
        "slides": []
    }

    theme_colors = {}

    with zipfile.ZipFile(pptx_path, 'r') as z:
        slide_files = []

        for file in z.namelist():

            if file == "ppt/theme/theme1.xml":
                root = ET.fromstring(z.read(file))
                theme = parse_theme(root)
                theme_colors = theme["colors"]

            # Match ppt/slides/slide#.xml exactly
            if re.match(r'^ppt/slides/slide\d+\.xml$', file):
                slide_files.append((file, z.read(file)))

        # Sort slide files numerically by extracting the slide number
        slide_files.sort(key=lambda x: int(re.search(r'slide(\d+)\.xml', x[0]).group(1)))

        # Process each slide
        for slide_idx, (file, slide_content) in enumerate(slide_files, 1):
            root = ET.fromstring(slide_content)

            # Extract design system info
            system["designSystem"]["colors"].update(extract_colors(root, theme_colors))

            fonts, sizes = extract_typography(root)
            system["designSystem"]["fonts"].update(fonts)
            system["designSystem"]["fontSizes"].update(sizes)

            system["designSystem"]["gradients"].extend(extract_gradients(root))
            system["designSystem"]["layout"].extend(extract_layout(root))

            # Extract detailed slide information
            slide_info = extract_slide_info(slide_content, slide_idx)
            system["slides"].append(slide_info)

        system["metadata"]["totalSlides"] = len(slide_files)

    # Convert sets → lists
    system["designSystem"]["colors"] = list(system["designSystem"]["colors"])
    system["designSystem"]["fonts"] = list(system["designSystem"]["fonts"])
    system["designSystem"]["fontSizes"] = sorted(list(system["designSystem"]["fontSizes"]))

    return system

if __name__ == "__main__":
    pptx_file = os.getenv("PPT_TEMPLATE_PATH", "templates/ppt-template.pptx")
    output_file = f"outputs/ppt_analysis/ppt_analysis_{pptx_file.split('/')[-1].split('.')[0]}.json"  # output file

    analysis = extract_design_system(pptx_file)

    # Save to JSON file
    with open(output_file, "w", encoding="utf-8") as f:
        json.dump(analysis, f, indent=2, sort_keys=True)

    logger.info("Detailed PPT analysis saved to %s", output_file)
    logger.info("Total slides analyzed: %s", analysis["metadata"]["totalSlides"])