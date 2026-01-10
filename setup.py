from setuptools import setup, find_packages

setup(
    name="visio_restyle",
    version="0.1.0",
    description="Convert Visio diagrams to different visual styles using LLM-powered shape mapping",
    author="",
    packages=find_packages(),
    install_requires=[
        "vsdx>=0.5.0",
        "openai>=1.0.0",
        "pydantic>=2.0.0",
        "pyyaml>=6.0",
        "python-dotenv>=1.0.0",
    ],
    entry_points={
        "console_scripts": [
            "visio-restyle=visio_restyle.main:main",
        ],
    },
    python_requires=">=3.8",
)
