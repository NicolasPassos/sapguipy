from setuptools import setup
from os import path

here = path.abspath(path.dirname(__file__))

with open("README.md", "r") as arq:
    long_description = arq.read()

with open(path.join(here, "requirements.txt"), "r", encoding="utf-8") as arq:
    requirement = arq.readlines()
requirements_list = [requirement.strip() for requirement in requirement]


setup(
    name="sapguipy",
    version="0.0.5",
    author="Nicolas Passos",
    license="MIT License",
    description="Manipulate SAP GUI with some lines of code",
    long_description=long_description,
    long_description_content_type="text/markdown",
    author_email="nicolasduart21@gmail.com",
    packages=["sapguipy"],
    keywords="sap",
    python_requires='>=3.8',
    install_requires=requirements_list,
)