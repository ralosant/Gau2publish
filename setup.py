from setuptools import setup, find_packages
from pathlib import Path

README = Path(__file__).with_name("README.md")
long_description = README.read_text(encoding="utf-8") if README.exists() else ""

setup(
    name="gau2publish",
    version="1.0.0",
    description=(
        "Convert Gaussian (.log) outputs into .geom/.png/.cdxml and assemble a DOCX "
        "plus an optional ALL.cdxml mosaic"
    ),
    long_description=long_description,
    long_description_content_type="text/markdown",
    author="Raúl Losantos Cabello",
    license="MIT",
    python_requires=">=3.9",
    packages=find_packages(where="src"),
    package_dir={"": "src"},
    install_requires=[
        "numpy>=1.23",
        "pandas>=1.5",
        "cclib>=1.8",
        "periodictable>=1.6",
        "rdkit>=2022.09",
        "python-docx>=0.8.11"
    ],
    extras_require={
        "render": ["xyzrender"],  # optional
    },
    entry_points={
        "console_scripts": [
            "gau2publish=gau2publish.cli:main",
        ]
    },
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
        "Topic :: Scientific/Engineering :: Chemistry",
    ],
    keywords=["Gaussian", "chemistry", "cclib", "rdkit", "docx", "cdxml"],
    include_package_data=True,
    zip_safe=False,
)