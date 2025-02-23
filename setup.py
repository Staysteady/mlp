"""
Market Maker System Setup

This module configures the package installation for the market maker system.
It specifies:
- Package metadata
- Python version requirements
- Required dependencies
"""
from setuptools import setup, find_packages

setup(
    name="market_maker",
    version="0.1.0",
    description="A market making system for processing Excel data",
    author="John Stedman",
    packages=find_packages(),
    python_requires=">=3.10.11",
    install_requires=[
        "pandas",
        "sqlalchemy",
        "xlwings;platform_system=='Darwin'",  # Mac only
        "pywin32;platform_system=='Windows'",  # Windows only
        "pytest",
        "freezegun",
        "schedule",
        "openpyxl",
        "click"
    ],
    entry_points={
        'console_scripts': [
            'market-maker=market_maker.main:main',
            'market-maker-cli=market_maker.cli:cli'
        ]
    },
    classifiers=[
        "Development Status :: 3 - Alpha",
        "Intended Audience :: Financial and Insurance Industry",
        "Programming Language :: Python :: 3.10",
        "Topic :: Office/Business :: Financial :: Investment"
    ]
)
