#!/usr/bin/env python3

from setuptools import setup, find_packages

# Read version from package
with open("wordtoppt/__init__.py", "r") as f:
    for line in f:
        if line.startswith("__version__"):
            version = line.split("=")[1].strip().strip('"').strip("'")
            break
    else:
        version = "0.1.0"

# Read long description from README
with open("README.md", "r") as f:
    long_description = f.read()

setup(
    name="wordtoppt",
    version=version,
    description="Convert Word documents to PowerPoint presentations using AI",
    long_description=long_description,
    long_description_content_type="text/markdown",
    author="Your Name",
    author_email="your.email@example.com",
    url="https://github.com/yourusername/wordtoppt",
    packages=find_packages(),
    install_requires=[
        "python-docx>=0.8.11",
        "python-pptx>=0.6.21",
        "openai>=1.0.0",
        "python-dotenv>=1.0.0",
        "typer[all]>=0.9.0",
    ],
    entry_points={
        "console_scripts": [
            "wordtoppt=wordtoppt.wordtoppt:app",
        ],
    },
    classifiers=[
        "Development Status :: 3 - Alpha",
        "Intended Audience :: End Users/Desktop",
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Topic :: Office/Business",
        "Topic :: Utilities",
    ],
    python_requires=">=3.9",
)
