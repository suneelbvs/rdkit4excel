# rdkit4excel
# Pyxll, RDKit, and Jupyter-Pyxll Integration Guide

Existing Excel add-ons for RDKit typically utilize VBA and Macros. For example [rdkit4excel](https://github.com/janholstjensen/rdkit4excel)

If you're looking for VBA & Macro's free add-on?. This is for you!

This repository provides a step-by-step guide on how to integrate `pyxll`, `rdkit`, and `jupyter-pyxll` for a seamless Python and Excel experience, especially for cheminformatics tasks.

## About Pyxll:
‘PyXLL add-in’ lets you integrate Python into Excel and use Python instead of VBA. Please note that, PyXLL is commercial and you can still access PyXLL for 30 days. Refer [PyXLL website](https://www.pyxll.com/) for more information.

## Table of Contents

- [Requirements](#Requirements)
- [Setting Up the Environment](#setting-up-the-environment)
- [Creating Enviorment from YML file](#Creating-Enviorment-from-YML-file)
- [Configuration](#Configuration)
- [License](#license) : Needs PyXLL Commercial license. Reach out to [PyXLL](https://www.pyxll.com/).

## Requirements:

Python >=3.10

PyXLL >= 5.1.0

Jupyter >= 1.0.0

rdkit >= 2022.2

notebook >= 6.0.0

PySide2, or PySide6 for Python >= 3.10

Optional

jupyterlab >= 4.0.0

## Setting Up the Environment

### 1. Install Anaconda

If you haven't already, download and install [Anaconda](https://www.anaconda.com/products/distribution) for your operating system.

### 2. Create a New Anaconda Environment

Open the Anaconda prompt or terminal and run the following command:

```bash
conda create --name rdkit_env python=3.10

or alternatively you can also create an environment using rdkit_env.yml file

### Step 1:
To set up ‘PyXLL Excel add-in (pip install pyxll) and then use the PyXLL command line tool to install the Excel add-in: 

> pip install pyxll

> pyxll install 

Follow the instructions to complete the installation process.

When the installer asks “Have you already downloaded the PyXLL add-in?”,

> enter “n” and the installer will download everything you need automatically.

If you prefer, you can download the PyXLL add-in from the [download page](https://www.pyxll.com/download.html),

> but please be sure to select the correct Python and Excel options to get the right version of the PyXLL add-in.

To activate the free 30 day trial when the installer asks *“Do you have a PyXLL license key?”*

> enter “n”. This will install PyXLL without a license key, activating the 30 day free trial automatically.

Step 2:
Once you have the PyXLL Excel add-in installed the next step is to install the pyxll-jupyter package. This package provides the glue between PyXLL and Jupyter so that we can use our Jupyter notebooks inside of Excel.

The pyxll-jupyter package is installed using pip:

> pip install pyxll-jupyter
Once both the PyXLL Excel add-in and the PyXLL-Jupyter package are installed start Excel and you will see a new “Jupyter” button in the PyXLL tab.

## 3. Creating Enviorment from YML file:

- [Anaconda](https://www.anaconda.com/products/distribution) or [Miniconda](https://docs.conda.io/en/latest/miniconda.html) installed on your system.

- conda env create -f rdkit_env.yml

- conda activate rdkit_env

## 4. Configuration

- The pyxll.cfg configuration file. You can locate this typically in *C:\Users\username\AppData\Local\Programs\PyXLL*

- This is the PyXLL configuration file and you will need to modify this to load your own Python modules

- Right click on pyxll.cfg
                    > locate modules > add 'rdkit_functions'
                    > locate pythonpath > add ./src

Now, download src folder from this repo and paste into your pyxll folder. Now restart excel to see Pyxll example tab

Please reach out to me for any queries.
