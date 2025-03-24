"""
AI content generation for WordToPPT.
"""

import json
import logging
import os

import openai

logger = logging.getLogger(__name__)


class AIGenerator:
    """
    Generate presentation content using AI.
    """

    def __init__(
        self,
        provider: str = "openai",
        max_tokens: int = 4000,
        temperature: float = 0.7,
        api_key: str = None,
    ):
        """
        Initialize the AI generator.

        Args:
            provider: AI provider to use (openai)
            max_tokens: Maximum tokens for generation
            temperature: Temperature for generation
            api_key: API key (optional, will use environment variable if not provided)
        """
        self.provider = provider
        self.max_tokens = max_tokens
        self.temperature = temperature
        self.api_key = api_key or os.getenv("OPENAI_API_KEY")
        self.logger = logging.getLogger(__name__)

        if self.provider != "openai":
            raise ValueError("Invalid AI provider. Must be 'openai'")

    def generate_presentation(self, content: dict) -> dict:
        """
        Generate presentation content from document content.

        Args:
            content: Document content dictionary

        Returns:
            Generated presentation structure
        """
        prompt = self._create_prompt(content)
        return self._generate_with_openai(prompt)

    def _create_prompt(self, content: dict) -> str:
        """
        Create the prompt for AI content generation.

        Args:
            content: Document content dictionary from DocxParser

        Returns:
            Generated prompt
        """
        # Extract sections for better prompting
        sections_text = ""
        for section in content.get("sections", []):
            heading = section["heading"]["text"]
            section_content = "\n".join(section["content"])
            sections_text += f"\nSection: {heading}\nContent:\n{section_content}\n"

        return (
            "Create a complete presentation outline from the following document content. "
            "For each section, create a section slide followed by content slides summarizing the section's content as bullet points. "
            "Do not limit the number of slides; generate as many slides as necessary based on the document content. "
            "Format the response as a JSON object with these keys:\n"
            "- presentation_title: The document title\n"
            "- presentation_subtitle: Optional subtitle\n"
            "- slides: Array of slide objects containing:\n"
            "  - slide_type: One of 'title', 'section', or 'content'\n"
            "  - title: Slide title\n"
            "  - subtitle: Optional subtitle (for title slides)\n"
            "  - bullets: Array of bullet points summarizing the section content (for content slides)\n"
            "  - notes: Optional speaker notes\n\n"
            f"Document Title: {content.get('title', '')}\n"
            "Document Content:"
            f"{sections_text}"
        )

    def _generate_with_openai(self, prompt: str) -> dict:
        """
        Generate content using OpenAI.

        Args:
            prompt: Prompt for generation

        Returns:
            Generated content as dictionary
        """
        client = openai.OpenAI(api_key=self.api_key)
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {
                    "role": "developer",
                    "content": "You are a helpful presentation creation assistant.",
                },
                {"role": "user", "content": prompt},
            ],
            temperature=self.temperature,
            max_tokens=self.max_tokens,
        )

        try:
            content = response.choices[0].message.content.strip()
            # If the response is wrapped in markdown code block, remove the markers
            if content.startswith("```json"):
                content = content[7:].strip()
                if content.endswith("```"):
                    content = content[:-3].strip()
            elif content.startswith("```"):
                content = content[3:].strip()
                if content.endswith("```"):
                    content = content[:-3].strip()

            if not content:
                self.logger.error("OpenAI returned empty content.")
                raise ValueError("Empty content returned from OpenAI")

            return json.loads(content)
        except json.JSONDecodeError as e:
            self.logger.error("Failed to parse OpenAI response as JSON: %s", str(e))
            raise ValueError("Invalid JSON response from OpenAI") from e
