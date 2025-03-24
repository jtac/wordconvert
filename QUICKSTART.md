# WordToPPT Quickstart Guide

This guide will help you get started with WordToPPT, a tool that converts Word documents to PowerPoint presentations using AI.

## Prerequisites

Before you begin, ensure you have:

1. Python 3.9 or higher installed
2. A valid Word document (.docx) that you want to convert
3. (Optional) A PowerPoint template (.pptx) for styling
4. An API key for OpenAI (GPT-4o)

## Installation

1. Clone or download the repository:

```bash
git clone https://github.com/yourusername/wordtoppt.git
cd wordtoppt
```

2. Create and activate a virtual environment:

```bash
# On Windows
python -m venv venv
venv\Scripts\activate

# On macOS/Linux
python -m venv venv
source venv/bin/activate
```

3. Install the required dependencies:

```bash
pip install -r requirements.txt
```

4. Set up your API key:

```bash
# Copy the example environment file
cp .env.example .env

# Edit the .env file with your OpenAI API key
```

## Basic Usage

The output presentation will be saved with the same filename as the input, but with a .pptx extension. For example, converting `document.docx` will produce `document.pptx`.

```bash
# Convert a Word document (basic usage)
python wordtoppt.py convert document.docx
```

If you have a custom PowerPoint template, use the `--template` option:

```bash
python wordtoppt.py convert document.docx --template custom_template.pptx
```

For verbose output to help with debugging:

```bash
python wordtoppt.py convert --verbose document.docx
```

## Example

Let's convert a sample document:

```python
from docx import Document

# Create a new Word document
doc = Document()
doc.add_heading('My Document', 0)
doc.add_paragraph('This is a sample paragraph.')
doc.add_heading('Section 1', level=1)
doc.add_paragraph('Content for section 1.')
doc.save('sample.docx')
```

Then convert it by running:

```bash
python wordtoppt.py convert --verbose sample.docx
```

## Troubleshooting

- **API Key Errors**: Verify that your API key is correctly set in the `.env` file.
- **Module Not Found**: Ensure you have installed all dependencies (`pip install -r requirements.txt`).
- **Invalid Document**: Make sure your Word document is a valid .docx file.
- **Missing Content**: If the generated presentation lacks content, ensure your document contains headings and paragraphs for the AI to parse.

## Next Steps

- Experiment with different documents and templates.
- Review the AI-generated presentation structure for further customizations.
- Adjust AI parameters in the `.env` file if needed. 