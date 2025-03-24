"""
Utility functions for WordToPPT.
"""

import logging
import os
from pathlib import Path
from typing import Optional


def setup_logging(verbose: bool = False) -> None:
    """
    Set up logging configuration.

    Args:
        verbose: Whether to enable debug logging
    """
    log_level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=log_level,
        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    )


def validate_file_path(file_path: str, must_exist: bool = True) -> bool:
    """
    Validate a file path.

    Args:
        file_path: Path to validate
        must_exist: Whether the file must already exist

    Returns:
        True if valid, False otherwise
    """
    if must_exist and not os.path.exists(file_path):
        return False

    # Check if directory is writable if we're creating a file
    if not must_exist:
        dir_path = os.path.dirname(file_path)
        if dir_path and not os.path.exists(dir_path):
            try:
                os.makedirs(dir_path)
            except Exception:
                return False

        if not os.access(dir_path or ".", os.W_OK):
            return False

    return True


def get_api_key() -> Optional[str]:
    """
    Get the OpenAI API key from environment variables.

    Returns:
        API key if found, None otherwise
    """
    return os.getenv("OPENAI_API_KEY")


def format_slide_notes(notes: str) -> str:
    """
    Format slide notes for better readability.

    Args:
        notes: Raw notes text

    Returns:
        Formatted notes text
    """
    if not notes:
        return ""

    # Split into lines and strip whitespace
    lines = [line.strip() for line in notes.split("\n")]

    # Remove empty lines and join with newlines
    return "\n".join(line for line in lines if line)
