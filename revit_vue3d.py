# -*- coding: utf-8 -*-
# Ouvre la vue 3D du projet Revit (la crée si elle n'existe pas).
# À lancer DANS Revit: pyRevit (bouton ou "Run script") ou RevitPythonShell.
# Compatible IronPython 2.7 et CPython 3.
from Autodesk.Revit.DB import (FilteredElementCollector, View3D, ViewFamilyType, ViewFamily, Transaction)

uidoc = __revit__.ActiveUIDocument   # noqa: F821  (variable fournie par pyRevit / RevitPythonShell)
doc = uidoc.Document

NOM_DEFAUT = "{3D}"                  # nom de la vue 3D par défaut de Revit

def find_3d():
    """Vue 3D par défaut "{3D}", sinon la première vue 3D (hors gabarits) en perspective ou non."""
    vues = [v for v in FilteredElementCollector(doc).OfClass(View3D) if not v.IsTemplate]
    for v in vues:
        if v.Name == NOM_DEFAUT:
            return v
    return vues[0] if vues else None

def create_3d():
    vft = next((t for t in FilteredElementCollector(doc).OfClass(ViewFamilyType)
                if t.ViewFamily == ViewFamily.ThreeDimensional), None)
    if vft is None:
        raise Exception("Aucun type de vue 3D dans ce projet.")
    t = Transaction(doc, "Créer vue 3D")
    t.Start()
    try:
        v = View3D.CreateIsometric(doc, vft.Id)
        try:
            v.Name = NOM_DEFAUT
        except Exception:                # nom déjà pris: on garde le nom proposé par Revit
            pass
        t.Commit()
    except Exception:
        t.RollBack()
        raise
    return v

def main():
    if doc.IsFamilyDocument:
        print("Ouvrez un projet (pas une famille).")
        return
    v = find_3d() or create_3d()
    uidoc.ActiveView = v
    print("Vue active : %s" % v.Name)

main()
