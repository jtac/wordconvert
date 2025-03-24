# WordToPPT CLI Tool

A command-line tool that converts Word documents to PowerPoint presentations using AI summarization.

## Features

- Extract content and structure from Word documents
- Use OpenAI (GPT-4o) to summarize and structure content for presentations
- Apply PowerPoint templates for consistent styling
- Add speaker notes to slides
- Simple CLI interface for easy use

## Installation

```bash
# Clone the repository
git clone https://github.com/yourusername/wordtoppt.git
cd wordtoppt

# Create and activate a virtual environment
python -m venv .venv
source .venv/bin/activate  # On Windows: .venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
```

## Usage

```bash
# Basic usage (output file will be derived from the input file, e.g., input.docx -> input.pptx)
python wordtoppt.py convert input.docx

# Use a custom PowerPoint template
python wordtoppt.py convert input.docx --template custom_template.pptx

# Enable verbose output for debugging
python wordtoppt.py convert --verbose input.docx
```

## Configuration

Create a `.env` file in the project root with the following content:

```bash
# OpenAI API key
OPENAI_API_KEY="your-openai-api-key"

# AI generation parameters
MAX_TOKENS=4000
TEMPERATURE=0.7
```

## Components

- **DocxParser**: Extracts content from Word documents
- **AIGenerator**: Generates presentation content using OpenAI (GPT-4o)
- **TemplateManager**: Analyzes PowerPoint templates
- **PPTXCreator**: Creates PowerPoint presentations

## Development

```bash
# Install development dependencies
pip install -r requirements-dev.txt

# Run tests
python -m pytest

# Run linting
pylint wordtoppt/*.py
```

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## How It Works

The tool follows a four-stage process:

1. **Document Analysis**: Extracts text, headings, and structure from the Word document
2. **Content Generation**: Uses OpenAI (GPT-4o) to generate a structured presentation outline in JSON format
3. **Template Analysis**: Identifies slide layouts and placeholders from the PowerPoint template
4. **Presentation Creation**: Builds the final PowerPoint by populating slides with the AI-generated content and applying consistent styling

### AI Providers

The tool exclusively uses OpenAI (GPT-4o) for generating presentation content.

## Architecture

1. **Input Processing**
   - Parses command-line arguments
   - Extracts text and structure from DOCX files
2. **Content Generation**
   - Sends document content to the OpenAI API
   - Receives a structured JSON response detailing the presentation outline
3. **Template Management**
   - Analyzes the provided PPTX template to determine slide layouts
4. **Presentation Assembly**
   - Creates a new presentation based on the template and populates it with the AI-generated content
   - Applies a consistent visual style across slides

## Technology Stack

- Python 3.9+
- Libraries:
  - python-docx
  - python-pptx
  - openai (GPT-4o)
  - python-dotenv
  - typer

## Project Structure

```
wordtoppt/
├── wordtoppt.py            # Main CLI entry point
├── requirements.txt        # Project dependencies
├── .env                    # Environment variables (API keys)
├── .env.example            # Example environment file
├── README.md               # Project documentation
├── wordtoppt/              # Source code
│   ├── __init__.py
│   ├── docx_parser.py      # DOCX content extraction
│   ├── ai_generator.py     # AI integration
│   ├── pptx_creator.py     # PowerPoint creation
│   ├── template_manager.py # Template handling
│   └── utils.py            # Utility functions
└── tests/                  # Unit tests
    ├── __init__.py
    ├── test_docx_parser.py
    ├── test_ai_generator.py
    ├── test_pptx_creator.py
    └── test_template_manager.py
```

## Development

To contribute to the project:

1. Fork the repository
2. Create a feature branch: `git checkout -b new-feature`
3. Make your changes
4. Run tests: `python -m unittest discover tests`
5. Submit a pull request

## License

MIT 