
# sappy

Sappy is designed to manipulate SAP GUI with some lines of code. It can be used to facilitate complexes RPA developments, realize tests or even to execute simple macros on SAP GUI.

### Key Features

1. **SAP Integration:** Uses SAP Script, that was especifully builded for manipulate the SAP GUI. If you never heard about, see de docs of [SAP Script](https://help.sap.com/docs/sap_gui_for_windows/b47d018c3b9b45e897faf66a6c0885a8/babdf65f4d0a4bd8b40f5ff132cb12fa.html).
2. **Flexible Configuration:** Configure to suit your needs.

## SAP Configuration
Before use this package, you must configure your SAP to accept scripts without notify.
Go to SAP configuration -> Accessbility and Scripting -> Scripting -> Uncheck the option "Notify when a script attaches to SAP GUI".

![SAP GUI Config](https://i.sstatic.net/lATNJ.jpg)

Save and you can go ahead to use this package.

## Usage:

To install this package, ensure you have Python installed.

### install with pip:

```bash
pip install sappy
```
### Import it
```python
from sappy import SapGui
```

### Run it
Fill the parameters with your SAP info and credentials:

```python
sap = SapGui(sid='PRD',user='USR',pwd='AnPassword',mandante='900',language='PT')
sap.login()
sap.start_transaction('SU01D')
```
You can use with context management too (i recommend using instead of the previous example):
```python
with SapGui(sid='PRD',user='USR',pwd='AnPassword',mandante='900',language='PT') as sap:
	sap.open_transaction('SU01D')
```

### The sky is the limit (or in this case, SAP Script)
You can do what do you want to do with SAP session generated.
If the methods of this package don't do what you need, you can build for your own using SAP Script docs.
```python
with SapGui(sid='PRD',user='USR',pwd='AnPassword',mandante='900',language='PT') as sap:
	height_windows = sap.session.findById('wnd[0]').Height
	with_window = sap.session.findById('wnd[0]').Width
	status_bar_text = sap.session.findById("wnd[0]/sbar/pane[0]").text
```
### How it works
This package abstracts and padronize the most utilized methods of SAP Script. So, you can worry only about   your business rules.
### How to Contribute
If you find some other functionality that is util or relevant to implement in this package, feel free to contribute.
We welcome contributions from the community and are pleased to have you join us.

### Prerequisites

Before contributing, please ensure you have the following:
- A basic understanding of Python and SAP Script.

### Setting Up Your Development Environment

1. **Fork the Repository**: Start by forking the repository to your GitHub account.
2. **Clone the Repository**: Clone your fork to your local machine.
   ```bash
   git clone https://github.com/NicolasPassos/sappy
   cd repository-name
   ```
3. **Install Dependencies**: Install the required dependencies.
   ```bash
   pip install -r requirements.txt
   ```

### Making Changes

1. **Create a New Branch**: Create a new branch for your feature or bug fix.
   ```bash
   git checkout -b awesome-feature
   ```
2. **Make Your Changes**: Implement your feature or fix a bug. Be sure to adhere to the coding standards and include comments where necessary.

### Submitting a Pull Request

1. **Commit Your Changes**: Once your tests pass, commit your changes.
   ```bash
   git commit -m 'Add some feature'
   ```
2. **Push to GitHub**: Push your changes to your fork on GitHub.
   ```bash
   git push origin awesome-feature
   ```
3. **Open a Pull Request**: Go to the original repository and click the *Compare & pull request* button. Then submit your pull request with a clear title and description.

### Code Review

Once your pull request is opened, it will be reviewed by the maintainers. Some changes may be requested. Please be patient and responsive. Once the pull request has been approved, it will be merged into the master branch.

Thank you for contributing!
