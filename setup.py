from setuptools import setup

with open("README.md", "r") as fh:
    long_description = fh.read()

setup(
    name="sappy",
    version="0.0.1",
    author="Nicolas Passos",
    license="MIT License",
    description="Manipulate SAP GUI with some lines of code",
    long_description=long_description,
    long_description_content_type="text/markdown",
    author_email="nicolasduart21@gmail.com",
    packages=["sappy"],
    keywords="sap",
    python_requires='>=3.8',
    install_requires=[]
)