#!/usr/bin/env python3
"""
Main module for WordToPPT package.
"""

import logging
import os
import traceback
from typing import Optional

import typer
from dotenv import load_dotenv

from .ai_generator import AIGenerator
from .docx_parser import DocxParser
from .pptx_creator import PPTXCreator
from .template_manager import TemplateManager
from .utils import validate_file_path, get_api_key

# Load environment variables
load_dotenv()

# Create Typer app
app = typer.Typer()


def extract_document_content(input_file: str) -> str:
    """
    Extract content from Word document.

    Args:
        input_file: Path to input Word document

    Returns:
        Extracted document content
    """
    parser = DocxParser(file_path=input_file)
    return parser.extract_content()


def generate_presentation_content(
    content: dict,
    api_key: str,
) -> dict:
    """
    Generate presentation content using AI.

    Args:
        content: Document content dictionary
        api_key: API key for OpenAI

    Returns:
        Generated presentation content
    """
    generator = AIGenerator(provider="openai", api_key=api_key)
    return generator.generate_presentation(content)


def create_presentation(
    content: dict, output_file: str, template_file: Optional[str] = None
) -> None:
    """
    Create PowerPoint presentation.

    Args:
        content: Presentation content
        output_file: Output file path
        template_file: Optional template file path
    """
    template_manager = TemplateManager(template_path=template_file)
    layouts = template_manager.analyze_template() if template_file else None

    creator = PPTXCreator(template_path=template_file, output_path=output_file)
    creator.create_presentation(content, layouts)


@app.command()
def convert(
    input_file: str = typer.Argument(..., help="Input Word document to convert"),
    template_file: str = typer.Option("template.pptx", "--template", "-t", help="PowerPoint template to use (defaults to template.pptx)"),
    verbose: bool = typer.Option(False, "--verbose", "-v", help="Enable verbose output")
) -> None:
    """
    Convert a Word document to a PowerPoint presentation.

    The output presentation will be saved with the same name as the input file but with .pptx extension.
    """
    # Set up logging
    log_level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=log_level,
        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    )
    logger = logging.getLogger(__name__)

    try:
        # Validate input file
        if not validate_file_path(input_file, must_exist=True):
            logger.error("Input file does not exist: %s", input_file)
            raise typer.Exit(1)

        # Validate template file
        if not validate_file_path(template_file, must_exist=True):
            logger.error("Template file does not exist: %s", template_file)
            raise typer.Exit(1)

        # Set output file path
        output_file = os.path.splitext(input_file)[0] + ".pptx"

        # Validate output file path
        if not validate_file_path(output_file, must_exist=False):
            logger.error("Cannot write to output file: %s", output_file)
            raise typer.Exit(1)

        # Get API key
        api_key = get_api_key()
        if not api_key:
            logger.error("No API key found. Set OPENAI_API_KEY environment variable.")
            raise typer.Exit(1)

        # Extract content from Word document
        logger.info("Extracting content from Word document...")
        document_content = extract_document_content(input_file)

        # Generate presentation content
        logger.info("Generating presentation content...")
        presentation_content = generate_presentation_content(
            document_content,
            api_key=api_key
        )

        # Create PowerPoint presentation
        logger.info("Creating PowerPoint presentation...")
        create_presentation(presentation_content, output_file, template_file)

        logger.info("Presentation created successfully: %s", output_file)

    except Exception as e:
        logger.error("Error: %s", str(e))
        if verbose:
            logger.debug(traceback.format_exc())
        raise typer.Exit(1)


if __name__ == "__main__":
    app()
