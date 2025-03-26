#!/usr/bin/env python3
from setuptools import setup, find_packages

setup(
    name="x2md",
    version="0.1.0",
    description="Universal File to Markdown Converter",
    author="Your Name",
    author_email="your.email@example.com",
    url="https://github.com/yourusername/x2md",
    packages=find_packages(),
    py_modules=["x2md", "txt2md", "csv2md", "xlsx2md", "docx2md", "pdf2md", "msg2md", "eml2md"],
    entry_points={
        "console_scripts": [
            "x2md=x2md:main",
            "txt2md=txt2md:main",
            "csv2md=csv2md:main",
            "xlsx2md=xlsx2md:main",
            "docx2md=docx2md:main",
            "pdf2md=pdf2md:main",
            "msg2md=msg2md:main",
            "eml2md=eml2md:main",
        ],
    },
    install_requires=[
        "pandas>=2.1.0",
        "tabulate>=0.9.0",
        "openpyxl>=3.1.0",
        "python-docx>=1.1.0",
        "PyPDF2>=3.0.0",
        "pdfplumber>=0.10.0",
        "extract-msg>=0.45.0",
    ],
    classifiers=[
        "Development Status :: 3 - Alpha",
        "Intended Audience :: End Users/Desktop",
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Topic :: Text Processing :: Markup :: Markdown",
    ],
    python_requires=">=3.8",
)
