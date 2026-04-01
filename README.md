# PPT Generation Pipeline

A complete, end-to-end automation system for generating tailored PowerPoint presentations. The pipeline analyzes a template PPT, extracts its design system, generates theme guidelines, creates prospect-specific slide content using AI, and produces a fully formatted PowerPoint file ready for presentation.

## 🎯 What It Does

This project automates the entire presentation creation workflow:

1. **Analyzes PPT Templates** - Extracts design system (colors, fonts, layouts, typography)
2. **Generates Theme Guides** - Uses Claude AI to create comprehensive theme documentation
3. **Creates Slide Content** - Generates prospect-specific sales presentations from company data
4. **Builds PowerPoint** - Creates fully formatted, design-compliant presentations

## 🚀 Quick Start

### Prerequisites

- Python 3.8+
- Virtual environment (venv)
- AWS credentials (for Bedrock and DynamoDB)
- `.env` file with AWS configuration

### Installation

```bash
# Clone or navigate to project
cd ppt-generation-6

# Create virtual environment
python -m venv venv

# Activate virtual environment
# Windows:
venv\Scripts\activate
# macOS/Linux:
source venv/bin/activate

# Install dependencies
pip install -r requirements.txt
```

### Environment Setup

Create a `.env` file in the project root:

```env
AWS_REGION=us-east-1
AWS_ACCESS_KEY_ID=your_access_key
AWS_SECRET_ACCESS_KEY=your_secret_key
BEDROCK_MODEL_ID=us.anthropic.claude-haiku-4-5-20251001-v1:0
DYNAMODB_COMPANIES_TABLE=companies
```

## 📋 Usage

### Option 1: Run Complete Pipeline (Recommended)

```bash
# Automated mode (uses defaults: seller=icicilombard, prospect=juniper)
python run_ppt_generation.py

# Interactive mode (prompts for configuration)
python run_ppt_generation.py interactive
```

**Output:** `outputs/presentations/presentation_juniper.pptx`

### Option 2: Run Individual Steps

```bash
# Step 1: Analyze PPT template
python analyze_ppt.py
# Output: outputs/ppt_analysis/ppt_detailed_analysis_ppt-template.json

# Step 2: Generate theme guide
python suggest_ppt_theme.py
# Output: outputs/theme_suggestions/theme_suggestion_ppt-template.md

# Step 3: Generate slide content
python generate_slide_content.py
# Output: outputs/generated_slides/slides_juniper.json

# Step 4: Create PowerPoint
python create_presentation.py
# Output: outputs/presentations/presentation_juniper.pptx
```

### Option 3: Programmatic Usage

```python
from run_ppt_generation import run_full_pipeline

# Run complete pipeline
result = run_full_pipeline(
    seller_company="icicilombard",
    prospect_company="juniper",
    ppt_template="templates/ppt-template.pptx"
)

# Skip certain steps
result = run_full_pipeline(
    seller_company="icicilombard",
    prospect_company="juniper",
    skip_steps=['analyze', 'theme']  # Already done, skip these
)

print(result)  # See summary of all outputs
```

## 📁 Project Structure

```
ppt-generation-6/
├── analyze_ppt.py                          # PPT analysis & design system extraction
├── suggest_ppt_theme.py                    # Theme guide generation
├── generate_slide_content.py               # Slide content generation
├── create_presentation.py                  # PowerPoint creation
├── run_ppt_generation.py                   # Pipeline orchestrator (MAIN ENTRY POINT)
│
├── core/
│   ├── config.py                           # Configuration & AWS setup
│   └── __pycache__/
│
├── storage_services/
│   ├── bedrock_operations.py               # Bedrock model invocation
│   ├── dynamodb_operations.py              # DynamoDB operations
│   └── __pycache__/
│
├── prompts/
│   ├── ppt_theme_suggestion_prompt.txt     # Theme analysis prompts
│   └── presentation_slides_generator_prompt.txt  # Slide generation prompts
│
├── templates/
│   └── ppt-template.pptx                   # PPT template to analyze
│
├── outputs/
│   ├── ppt_analysis/                       # Design system analysis
│   ├── theme_suggestions/                  # Theme guides (markdown)
│   ├── generated_slides/                   # Slide content (JSON)
│   └── presentations/                      # Final PowerPoint files
│
├── requirements.txt                        # Python dependencies
├── .env                                    # Environment variables (create this)
├── .gitignore                              # Git ignore rules
└── README.md                               # This file
```

## 🔄 Pipeline Stages Explained

### Stage 1: Analyze PPT (`analyze_ppt.py`)

**What it does:**
- Reads the PPT template (templates/ppt-template.pptx)
- Extracts colors, fonts, font sizes, gradients
- Analyzes slide layouts, positioning, spacing
- Extracts text content and formatting from each slide

**Output:** JSON file with complete design system analysis
```json
{
  "metadata": {...},
  "designSystem": {
    "colors": ["#02428E", "#F26633", ...],
    "fonts": ["Arial"],
    "fontSizes": [15, 18, 32, 40, 44],
    "gradients": [...]
  },
  "slides": [...]
}
```

### Stage 2: Generate Theme (`suggest_ppt_theme.py`)

**What it does:**
- Loads the PPT analysis JSON
- Sends design system to Claude via AWS Bedrock
- Claude analyzes and generates comprehensive theme documentation
- Creates markdown guide with design principles, color strategy, typography rules, etc.

**Output:** Markdown file with theme guidelines
```markdown
# InsightSphere Professional Theme

## Color Palette
- Primary: #02428E (Navy Blue)
- Accent: #F26633 (Orange)
- ...

## Typography
- Headers: 40-44pt, bold, Arial
- Body: 15pt, regular, Arial
- ...
```

### Stage 3: Generate Slides (`generate_slide_content.py`)

**What it does:**
- Loads company data (seller and prospect) from DynamoDB
- Loads theme guide (markdown)
- Sends everything to Claude with slide generation prompt
- Claude generates prospect-specific presentation content
- Content includes slide structure, copy, design specifications

**Output:** JSON file with structured slide data
```json
{
  "presentation_metadata": {...},
  "design_system_reference": {...},
  "slides": [
    {
      "slide_number": 1,
      "slide_type": "title",
      "title": "...",
      "design_notes": {...}
    },
    ...
  ]
}
```

### Stage 4: Create PowerPoint (`create_presentation.py`)

**What it does:**
- Loads the slide JSON (handles markdown code block wrappers)
- Iterates through each slide data
- Creates PowerPoint slides using python-pptx
- Applies colors, fonts, layouts from design system
- Adds text, shapes, gradients, callout boxes
- Saves as .pptx file

**Output:** PowerPoint presentation file ready for use

## 📊 Data Flow

```
PPT Template
    ↓
[analyze_ppt.py]
    ↓
Design System JSON
    ↓
[suggest_ppt_theme.py] + AWS Bedrock (Claude)
    ↓
Theme Guide (Markdown)
    ↓
[generate_slide_content.py] + AWS Bedrock (Claude) + DynamoDB
    ↓
Slide Content JSON
    ↓
[create_presentation.py]
    ↓
PowerPoint Presentation (.pptx)
```

## 🧪 Example: Generate Presentation for Juniper Networks

```bash
# Run the complete pipeline
python run_ppt_generation.py

# Or interactively
python run_ppt_generation.py interactive
# Then enter: seller=icicilombard, prospect=juniper
```

**Output files created:**
- `outputs/ppt_analysis/ppt_detailed_analysis_ppt-template.json`
- `outputs/theme_suggestions/theme_suggestion_ppt-template.md`
- `outputs/generated_slides/slides_juniper.json`
- `outputs/presentations/presentation_juniper.pptx` ✓ **READY TO USE**

## 🔧 Configuration

### Core Config (`core/config.py`)

Set these for your AWS environment:
- `BEDROCK_MODEL_ID` - Claude model to use (default: Haiku 4.5)
- `DYNAMODB_COMPANIES_TABLE` - DynamoDB table name
- AWS credentials (from `.env`)

### Bedrock Operations (`storage_services/bedrock_operations.py`)

Customize model parameters:
- `max_tokens` - Max response length (default: 4096)
- `temperature` - Creativity vs. consistency (default: 0.7)

### Prompts

**Theme Suggestion** (`prompts/ppt_theme_suggestion_prompt.txt`):
- System prompt: Instructions for Claude to analyze design systems
- User prompt template: Template for analyzing specific PPT

**Slide Generation** (`prompts/presentation_slides_generator_prompt.txt`):
- System prompt: Instructions for Claude to create sales presentations
- User prompt template: Template with JSON schema for slide output
- Includes complete JSON schema for slide structure

## 📦 Dependencies

Key packages:
- `python-pptx` - PowerPoint generation
- `boto3` - AWS Bedrock and DynamoDB access
- `python-dotenv` - Environment variable management
- `requests` - HTTP client (if needed)

Full list: See `requirements.txt`

## 🎨 Supported Slide Types

The system can generate:
- **Title Slides** - With gradients and hero messaging
- **Content Slides** - Bullets with accent lines and callouts
- **Two-Column Slides** - Side-by-side layouts
- **Comparison Slides** - Before/After layouts
- **Data/Chart Slides** - For analytics and metrics
- **Image + Text Slides** - Mixed content layouts
- **Centered Slides** - Key message focus
- **Closing Slides** - CTA and contact information

## 🎯 Design System Features

Automatically maintained across all generated presentations:
- ✓ Color palette consistency
- ✓ Typography hierarchy (font sizes, weights, colors)
- ✓ Margin and spacing standards
- ✓ Layout patterns (margins, gutters, alignment)
- ✓ Accent elements (orange highlights, lines, callout boxes)
- ✓ Gradient applications
- ✓ Accessibility (color contrast)

## 📝 Notes

- The PPT template (`templates/ppt-template.pptx`) is analyzed once; results are cached
- Theme guide is typically generated once per template
- Slide content is generated fresh for each prospect
- Company data is fetched from DynamoDB (ensure table and records exist)
- Claude APIs are called during steps 2 and 3 (may incur costs)

## 🔐 Security

- AWS credentials stored in `.env` (never committed to git)
- `.gitignore` excludes sensitive files
- See `.gitignore` for what's excluded from version control

## 🐛 Troubleshooting

### "Theme file not found"
- Run `python analyze_ppt.py` and `python suggest_ppt_theme.py` first
- Or use `run_ppt_generation.py` which handles all steps

### "Company data not found"
- Ensure DynamoDB table exists and has seller/prospect records
- Check AWS credentials in `.env`
- Verify table name in `core/config.py`

### "Invalid JSON in slide file"
- Check if JSON starts with ` ```json ` - the parser handles this automatically
- Ensure JSON is valid when created by `generate_slide_content.py`

### AWS Connection Errors
- Verify AWS credentials in `.env`
- Check AWS region matches your setup
- Ensure Bedrock model ID is available in your region

## 📚 Advanced Usage

### Skip Steps You've Already Done

```python
from run_ppt_generation import run_full_pipeline

# Skip theme generation if already done
result = run_full_pipeline(
    skip_steps=['analyze', 'theme']
)
```

### Use Different PPT Template

```python
result = run_full_pipeline(
    ppt_template="templates/my-custom-template.pptx"
)
```

### Use Different Companies

```python
result = run_full_pipeline(
    seller_company="acme-corp",
    prospect_company="tech-startup"
)
```

## 📞 Support

For issues or questions:
1. Check the Troubleshooting section
2. Review `.env` configuration
3. Verify AWS credentials and permissions
4. Check AWS Bedrock and DynamoDB services are accessible

## 📄 License

[Add your license information here]

## 🎉 Happy Presenting!

Your AI-powered presentation generation system is ready. Generate tailored, design-compliant presentations in minutes!
