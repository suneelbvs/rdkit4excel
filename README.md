# rdkit4excel
# Pyxll, RDKit, and Jupyter-Pyxll Integration Guide

This repository provides a step-by-step guide on how to integrate `pyxll`, `rdkit`, and `jupyter-pyxll` for a seamless Python and Excel experience, especially for cheminformatics tasks.

# About Pyxll:
‘PyXLL add-in’ lets you integrate Python into Excel and use Python instead of VBA. Please note that, PyXLL is commercial and you can still access PyXLL for 30 days. Refer [PyXLL website](https://www.pyxll.com/) for more information.

## Table of Contents

- [Setting Up the Environment](#setting-up-the-environment)
- [Creating the RDKit Functions File](#creating-the-rdkit-functions-file)
- [Examples](#examples)
- [Contributing](#contributing)
- [License](#license)

## Setting Up the Environment

### 1. Install Anaconda

If you haven't already, download and install [Anaconda](https://www.anaconda.com/products/distribution) for your operating system.

### 2. Create a New Anaconda Environment

Open the Anaconda prompt or terminal and run the following command:

```bash
conda create --name rdkit_env python=3.10

or alternatively you can also create an environment using rdkit_env.yml file

To set up ‘PyXLL Excel add-in (pip install pyxll) and then use the PyXLL command line tool to install the Excel add-in: 

> pip install pyxll

> pyxll install 
