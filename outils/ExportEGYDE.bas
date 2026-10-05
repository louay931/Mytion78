Attribute VB_Name = "ExportEGYDE"
Option Explicit

' Crée un fichier Excel séparé contenant uniquement l'onglet « Demande EGYDE »,
' en valeurs (sans formules ni lien vers ce classeur), prêt à être envoyé.
Sub ExporterEGYDE()
    Dim src As Worksheet, wb As Workbook, ws As Worksheet
    Dim nom As String, chemin As Variant, c As Variant, i As Long

    Set src = ThisWorkbook.Worksheets("Demande EGYDE")

    ' Nom de fichier proposé : DEMANDE_EGYDE_<nom du site client>
    nom = Trim(CStr(src.Range("B2").Value))
    If nom = "" Then nom = "SITE"
    For Each c In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        nom = Replace(nom, c, "_")
    Next c

    chemin = Application.GetSaveAsFilename( _
        InitialFileName:="DEMANDE_EGYDE_" & nom & ".xlsx", _
        FileFilter:="Classeur Excel (*.xlsx), *.xlsx", _
        Title:="Enregistrer la demande EGYDE")
    If chemin = False Then Exit Sub

    Application.ScreenUpdating = False
    src.Copy                       ' nouveau classeur avec seulement cet onglet
    Set wb = ActiveWorkbook
    Set ws = wb.Worksheets(1)
    ws.Cells.Copy
    ws.Cells.PasteSpecial Paste:=xlPasteValues   ' formules -> valeurs
    Application.CutCopyMode = False
    ws.Range("A1").Select
    ws.Name = "Demande EGYDE"

    ' le bouton ne sert à rien dans le fichier envoyé
    On Error Resume Next
    ws.Shapes("BoutonExportEGYDE").Delete
    On Error GoTo 0

    ' supprime les noms hérités qui pointent vers d'autres fichiers
    On Error Resume Next
    For i = wb.Names.Count To 1 Step -1
        wb.Names(i).Delete
    Next i
    On Error GoTo 0

    Application.DisplayAlerts = False
    wb.SaveAs Filename:=chemin, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Demande EGYDE enregistrée :" & vbCrLf & chemin & vbCrLf & vbCrLf & _
           "Le fichier est ouvert, tu peux l'envoyer.", vbInformation, "Export EGYDE"
End Sub
