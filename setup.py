import setuptools

from analysis_engine import __version__

with open("README.md", "r") as fh:
    long_description = fh.read()

setuptools.setup(
    name="analysis_engine",
    version=__version__,
    author="Will Grant",
    author_email="will.granty@yahoo.co.uk",
    description="Analysis for the UK Department for Transport's major projects portfolio",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/banillie/analysis_engine",
    packages=['analysis_engine'],
    entry_points={"console_scripts": ["analysis = analysis_engine.main:main"]},
    setup_requires=["wheel"],
    install_requires=[
        "datamaps",
        "python-docx==0.8.10",
        "openpyxl",
        "matplotlib==3.4.1",
        "pdf2image",
    ],
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires=">=3.7",
)
