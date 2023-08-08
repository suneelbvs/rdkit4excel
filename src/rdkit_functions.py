from pyxll import xl_func, xl_app
from rdkit import Chem
from rdkit.Chem import Descriptors, Lipinski, AllChem
from rdkit import DataStructs
from rdkit.Chem import PandasTools as pt
from rdkit.Chem import Draw
from PIL import Image
import io
import tempfile
import pandas as pd
import os

@xl_func(volatile=True)
def NumHAcceptors(smiles):
    """Calculates drug-like properties for a given SMILES string."""
    mol = Chem.MolFromSmiles(smiles)
    if not mol:
        return "Invalid SMILES"
    
    properties = {
        "HBA": Descriptors.NumHAcceptors(mol)
    }
    
    return NumHAcceptors
    
@xl_func(volatile=True)
def NumHDonors(smiles):
    """Calculates drug-like properties for a given SMILES string."""
    mol = Chem.MolFromSmiles(smiles)
    if not mol:
        return "Invalid SMILES"
    
    properties = {
        "HBD": Descriptors.NumHDonors(mol)
    }
    
    return NumHDonors
    
@xl_func(volatile=True)
def smiles_to_mol(smiles):
    """Converts a SMILES string to a molecule representation."""
    mol = Chem.MolFromSmiles(smiles)
    if not mol:
        return "Invalid SMILES"
    return Chem.MolToMolBlock(mol)

@xl_func("string molecule, string query: string", volatile=True)
def substructure_search(molecule, query):
    mol = Chem.MolFromSmiles(molecule)
    q = Chem.MolFromSmiles(query)
    if not mol or not q:
        return "Invalid SMILES"
    return "Yes" if mol.HasSubstructMatch(q) else "No"

@xl_func("string smiles1, string smiles2: float", volatile=True)
def molecule_similarity(smiles1, smiles2):
    m1 = Chem.MolFromSmiles(smiles1)
    m2 = Chem.MolFromSmiles(smiles2)
    if not m1 or not m2:
        return None
    fp1 = AllChem.GetMorganFingerprint(m1, 2)
    fp2 = AllChem.GetMorganFingerprint(m2, 2)
    return DataStructs.TanimotoSimilarity(fp1, fp2)
    
@xl_func("string smiles: float", volatile=True)
def MolecularWeight(smiles):
    mol = Chem.MolFromSmiles(smiles)
    if not mol:
        return "Invalid SMILES", None, None, None
    mw = Descriptors.MolWt(mol)
    return mw
    
@xl_func("string smiles: float", volatile=True)
def logp(smiles):
    mol = Chem.MolFromSmiles(smiles)
    if not mol:
        return "Invalid SMILES", None, None, None
    logp = Descriptors.MolLogP(mol)
    return logp
    
@xl_func("string smiles: float", volatile=True)
def rot_bonds(smiles):
    mol = Chem.MolFromSmiles(smiles)
    if not mol:
        return "Invalid SMILES", None, None, None
    rot_bonds = Descriptors.NumRotatableBonds(mol)
    return rot_bonds