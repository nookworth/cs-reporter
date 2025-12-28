"""
Setup script for CS Reporter.
Allows installation via pip and creates the 'reporter' command.
"""

from setuptools import setup, find_packages
from pathlib import Path

# Read the README file
readme_file = Path(__file__).parent / "README.md"
long_description = readme_file.read_text() if readme_file.exists() else ""

setup(
    name="cs-reporter",
    version="0.1.0",
    description="Excel to PowerPoint report generator",
    long_description=long_description,
    long_description_content_type="text/markdown",
    author="Christopher Morrison",
    author_email="chrismorrison1987@gmail.com",
    url="https://github.com/nookworth/cs-reporter",
    packages=find_packages(),
    include_package_data=True,
    install_requires=[
        "pandas>=2.0.0",
        "python-pptx>=0.6.21",
        "PyYAML>=6.0",
        "openpyxl>=3.1.0",
    ],
    entry_points={
        "console_scripts": [
            "reporter=src.main:cli_entry_point",
        ],
    },
    python_requires=">=3.8",
    classifiers=[
        "Development Status :: 3 - Alpha",
        "Intended Audience :: End Users/Desktop",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
    ],
)
