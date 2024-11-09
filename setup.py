from setuptools import setup

with open("README.md", "r") as arq:
    long_description = arq.read()

with open("requirements.txt", "r", encoding="utf-8") as arq:
    requirement = arq.readlines()
requirements_list = [requirement.strip() for requirement in requirement]


setup(
    name="sapguipy",
    version="0.0.2",
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