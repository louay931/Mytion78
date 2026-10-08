"""Corrige l'erreur 1004 « Impossible d'accéder à 'fusion_certificats.xlsx' » de la macro GenererCertificats :
fichier temporaire au nom unique à chaque génération (jamais bloqué par un Word resté ouvert),
anciens fichiers temporaires nettoyés, et gestion d'erreur couvrant aussi la création de ce fichier.
Usage : python corrige_fusion_temp.py <entrée.xlsm> <sortie.xlsm>"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from patch_vba import patcher

REMPLACEMENTS = [
("""    Application.CalculateFull
    tmp = Environ$("TEMP") & "\\fusion_certificats.xlsx"
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("Fiches clients").Copy
    Set wbTmp = ActiveWorkbook
""",
"""    Application.CalculateFull
    ' anciens fichiers temporaires (ceux encore bloqués par un Word resté ouvert sont ignorés)
    On Error Resume Next
    Kill Environ$("TEMP") & "\\fusion_certificats*.xlsx"
    On Error GoTo Erreur
    ' nom unique à chaque génération : jamais bloqué par une génération précédente
    Randomize
    tmp = Environ$("TEMP") & "\\fusion_certificats_" & Format(Now, "yyyymmdd_hhnnss") & "_" & CStr(Int(Rnd * 100000)) & ".xlsx"
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("Fiches clients").Copy
    Set wbTmp = ActiveWorkbook
"""),
("""    wbTmp.SaveAs Filename:=tmp, FileFormat:=51
    wbTmp.Close SaveChanges:=False
    Application.DisplayAlerts = True
""",
"""    wbTmp.SaveAs Filename:=tmp, FileFormat:=51
    wbTmp.Close SaveChanges:=False
    Set wbTmp = Nothing
    Application.DisplayAlerts = True
"""),
("""    MsgBox "Erreur pendant la génération : " & Err.Description, vbCritical, "Générer les certificats"
    On Error Resume Next
""",
"""    MsgBox "Erreur pendant la génération : " & Err.Description, vbCritical, "Générer les certificats"
    On Error Resume Next
    If Not wbTmp Is Nothing Then wbTmp.Close SaveChanges:=False
"""),
]

patcher(sys.argv[1], sys.argv[2], 'GenererCertificats', REMPLACEMENTS)
print('ok', sys.argv[2])
