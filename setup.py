from setuptools import setup, find_packages

setup(
    name="omnipdf",
    version="1.0.0",
    package_dir={"": "src"},
    packages=find_packages(where="src"),
    install_requires=[
        "pymupdf>=1.24.0",
        "python-docx>=1.1.0",
        "pdf2docx>=0.5.8",
        "pillow>=10.0.0",
        "fastapi>=0.110.0",
        "uvicorn>=0.28.0",
        "python-multipart>=0.0.9",
        "jinja2>=3.1.0",
        "typer>=0.9.0",
        "rich>=13.0.0",
        "lxml>=5.0.0",
    ],
    entry_points={
        "console_scripts": [
            "omnipdf=omnipdf.cli.main:main",
        ],
    },
)
