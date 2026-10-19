# -*- coding: utf-8 -*-
# Liste des feuilles Revit -> CSV (à ouvrir dans Excel, à comparer avec l'inventaire PDF)
# À lancer DANS Revit: pyRevit (bouton ou "Run script") ou RevitPythonShell.
# Compatible IronPython 2.7 et CPython 3 (pas de f-string, pas de dépendance externe).
import os, io, re, datetime
from Autodesk.Revit.DB import FilteredElementCollector, ViewSheet, BuiltInParameter, BuiltInCategory

doc = __revit__.ActiveUIDocument.Document   # noqa: F821  (variable fournie par pyRevit / RevitPythonShell)

# Paramètres d'occurrence lus sur chaque feuille: (colonne, BuiltInParameter)
BUILTIN = [
    ("Indice", BuiltInParameter.SHEET_CURRENT_REVISION),
    ("Date indice", BuiltInParameter.SHEET_CURRENT_REVISION_DATE),
    ("Description indice", BuiltInParameter.SHEET_CURRENT_REVISION_DESCRIPTION),
    ("Date d'émission", BuiltInParameter.SHEET_ISSUE_DATE),
    ("Dessiné par", BuiltInParameter.SHEET_DRAWN_BY),
    ("Vérifié par", BuiltInParameter.SHEET_CHECKED_BY),
]
# Paramètres partagés / de projet à ajouter (nom exact tel qu'affiché dans Revit), ex: ["Lot", "Phase"]
EXTRA = []

HEADERS = (["N° feuille", "Nom feuille"] + [h for h, _ in BUILTIN] + EXTRA +
           ["Cartouche", "Nb vues", "Nom PDF attendu", "Id"])

def txt(x):
    return u"" if x is None else (x if isinstance(x, type(u"")) else u"%s" % x)

def param(el, p):
    """Valeur d'un paramètre (BuiltInParameter ou nom) sous forme de texte, "" si absent."""
    pr = el.get_Parameter(p) if not isinstance(p, (str, type(u""))) else el.LookupParameter(p)
    if pr is None or not pr.HasValue:
        return u""
    return txt(pr.AsString() if pr.AsString() is not None else pr.AsValueString())

def safe_name(s):
    """Nom de fichier tel que Revit le produit à l'export PDF (caractères interdits remplacés)."""
    return re.sub(r'[\\/:*?"<>|]', "-", s).strip()

def title_block(sheet):
    tb = (FilteredElementCollector(doc, sheet.Id).OfCategory(BuiltInCategory.OST_TitleBlocks)
          .WhereElementIsNotElementType().FirstElement())
    if tb is None:
        return u""
    sym = doc.GetElement(tb.GetTypeId())
    return txt(sym.FamilyName + u" : " + sym.Name) if sym is not None else u""

def rows():
    sheets = [s for s in FilteredElementCollector(doc).OfClass(ViewSheet) if not s.IsPlaceholder]
    sheets.sort(key=lambda s: s.SheetNumber)
    for s in sheets:
        num, name = txt(s.SheetNumber), txt(s.Name)
        yield ([num, name] + [param(s, b) for _, b in BUILTIN] + [param(s, e) for e in EXTRA] +
               [title_block(s), len(list(s.GetAllPlacedViews())),
                safe_name(u"Feuille - %s - %s.pdf" % (num, name)), txt(s.Id)])

def csv_line(values):
    """Séparateur ';' et BOM UTF-8: s'ouvre directement dans Excel français avec les accents."""
    out = []
    for v in values:
        v = txt(v)
        if any(c in v for c in u';"\n\r'):
            v = u'"' + v.replace(u'"', u'""') + u'"'
        out.append(v)
    return u";".join(out) + u"\r\n"

def target_path():
    title = re.sub(r'[\\/:*?"<>|]', "_", doc.Title or "Projet")
    default = u"Feuilles_%s_%s.csv" % (title, datetime.datetime.now().strftime("%Y%m%d_%H%M"))
    try:                                     # boîte "Enregistrer sous" de pyRevit si disponible
        from pyrevit import forms
        return forms.save_file(file_ext="csv", default_name=default)
    except ImportError:
        pass
    home = os.path.expanduser("~")
    for d in [os.path.join(home, "Desktop"), os.path.join(home, "OneDrive", "Desktop"),
              os.path.join(home, "Documents"), home]:
        if os.path.isdir(d):
            return os.path.join(d, default)
    return default

def main():
    data = list(rows())
    if not data:
        print(u"Aucune feuille dans ce projet.")
        return
    path = target_path()
    if not path:                             # annulé par l'utilisateur
        return
    with io.open(path, "w", encoding="utf-8-sig", newline="") as f:
        f.write(csv_line(HEADERS))
        for r in data:
            f.write(csv_line(r))
    print(u"%d feuilles -> %s" % (len(data), path))
    try:
        os.startfile(path)
    except Exception:
        pass

main()
