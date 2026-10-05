"""Transforme Fiches_clients.xlsx en Fiches_clients.xlsm : ajoute la macro ExportEGYDE.bas
et un bouton « Exporter EGYDE » sur l'onglet Demande EGYDE.
Usage : python construire_xlsm.py Fiches_clients.xlsx ExportEGYDE.bas Fiches_clients.xlsm
Nécessite : pip install openpyxl ms_ovba ms_cfb"""
import os, re, sys, tempfile, uuid, zipfile
import openpyxl
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import ovba_compression
from ms_ovba_compression.ms_ovba import MsOvba
MsOvba.compress = lambda self, data: ovba_compression.compress(data)   # la compression du paquet est boguée
from ms_ovba.vbaProject import VbaProject
from ms_ovba.Models.Entities.doc_module import DocModule
from ms_ovba.Models.Entities.std_module import StdModule
from ms_ovba.Models.Entities.reference import Reference
from ms_ovba.Models.Entities.reference_registered import ReferenceRegistered
from ms_ovba.Models.Fields.libid_reference import LibidReference
from ms_ovba.Views.project_ole_file import ProjectOleFile

src, bas, out = (os.path.abspath(a) for a in sys.argv[1:4])
FEUILLE, BOUTON, MACRO = 'Demande EGYDE', 'BoutonExportEGYDE', 'ExporterEGYDE'

# 1. noms de code VBA des feuilles (doivent correspondre aux modules du projet)
wb = openpyxl.load_workbook(src)
wb.code_name = 'ThisWorkbook'
codes = {}
for i, ws in enumerate(wb.worksheets, 1):
    ws.sheet_properties.codeName = codes[ws.title] = f'Sheet{i}'
tmp = tempfile.mkdtemp()
xlsx = os.path.join(tmp, 'base.xlsx'); wb.save(xlsx)

# 2. vbaProject.bin (source seule, sans cache compilé : Excel compile à l'ouverture)
os.chdir(tmp)
projet = VbaProject(); projet.project_id = '{' + str(uuid.uuid4()).upper() + '}'
def module_document(nom, guid):
    with open(nom + '.cls', 'w') as f:
        f.write('VERSION 1.0 CLASS\nBEGIN\n  MultiUse = -1  \'True\nEND\n'
                f'Attribute VB_Name = "{nom}"\nAttribute VB_GlobalNameSpace = False\n'
                'Attribute VB_Creatable = False\nAttribute VB_PredeclaredId = True\nAttribute VB_Exposed = True\n')
    m = DocModule(nom); m.add_file(nom + '.cls'); m.add_guid(uuid.UUID(guid)); m.normalize_file()
    projet.add_module(m)
module_document('ThisWorkbook', '0002081900000000C000000000000046')
for nom in codes.values():
    module_document(nom, '0002082000000000C000000000000046')
m = StdModule('ExportEGYDE'); m.add_file('ExportEGYDE')
with open(bas, 'rb') as f, open('ExportEGYDE.new', 'wb') as g:   # source en cp1252 / CRLF, telle quelle
    g.write(f.read())
projet.add_module(m)
for guid, ver, chemin, nom, alias in [
        ('0002043000000000C000000000000046', '2.0', r'C:\Windows\System32\stdole2.tlb', 'OLE Automation', 'stdole'),
        ('2DF8D04C5BFA101BBDE500AA0044DE52', '2.0',
         r'C:\Program Files\Common Files\Microsoft Shared\OFFICE16\MSO.DLL', 'Microsoft Office 16.0 Object Library', 'Office')]:
    projet.add_reference(Reference(ReferenceRegistered(LibidReference(uuid.UUID(guid), ver, '0', chemin, nom)), alias))
ProjectOleFile.write_file(projet)
vba = open(os.path.join(tmp, 'vbaProject.bin'), 'rb').read()

# 3. bouton (forme avec macro affectée, non imprimée) dans le dessin de l'onglet EGYDE
FORME = (f'<xdr:twoCellAnchor xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
         'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
         f'<xdr:from><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row>'
         f'<xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>10</xdr:col><xdr:colOff>0</xdr:colOff>'
         f'<xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'
         f'<xdr:sp macro="[0]!{MACRO}" textlink=""><xdr:nvSpPr><xdr:cNvPr id="100" name="{BOUTON}"/><xdr:cNvSpPr/></xdr:nvSpPr>'
         '<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>'
         '<a:solidFill><a:srgbClr val="1F4D78"/></a:solidFill><a:ln><a:noFill/></a:ln></xdr:spPr>'
         '<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip" rtlCol="0" anchor="ctr"/><a:lstStyle/>'
         '<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="fr-FR" sz="1200" b="1"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
         '</a:rPr><a:t>Exporter EGYDE</a:t></a:r></a:p></xdr:txBody></xdr:sp><xdr:clientData fPrintsWithSheet="0"/></xdr:twoCellAnchor>')

zin = zipfile.ZipFile(xlsx)
wbxml = zin.read('xl/workbook.xml').decode()
rels = zin.read('xl/_rels/workbook.xml.rels').decode()
rid = re.search(r'<sheet [^>]*name="' + re.escape(FEUILLE) + r'"[^>]*r:id="([^"]+)"', wbxml).group(1)
cible = re.search(r'Id="' + rid + r'"[^>]*Target="([^"]+)"|Target="([^"]+)"[^>]*Id="' + rid + '"', rels)
feuille = 'xl/' + (cible.group(1) or cible.group(2)).lstrip('/').replace('xl/', '')
frels = zin.read(feuille.replace('worksheets/', 'worksheets/_rels/') + '.rels').decode()
dessin = 'xl/' + re.search(r'Target="(?:\.\./|/xl/)([^"]*drawing[^"]*)"', frels).group(1)

zout = zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED)
for item in zin.infolist():
    data = zin.read(item.filename)
    if item.filename == '[Content_Types].xml':
        t = data.decode().replace('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
                                  'application/vnd.ms-excel.sheet.macroEnabled.main+xml')
        t = t.replace('</Types>', '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>')
        data = t.encode()
    elif item.filename == 'xl/_rels/workbook.xml.rels':
        data = data.decode().replace('</Relationships>',
            '<Relationship Id="rIdVBA" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" '
            'Target="vbaProject.bin"/></Relationships>').encode()
    elif item.filename == dessin:
        t = data.decode()
        fin = t.rindex('</')   # </wsDr> ou </xdr:wsDr> selon le préfixe utilisé
        data = (t[:fin] + FORME + t[fin:]).encode()
    zout.writestr(item, data)
zout.writestr('xl/vbaProject.bin', vba)
zout.close()
assert FORME.split('>')[0] and BOUTON in zipfile.ZipFile(out).read(dessin).decode()
print(out, '| bouton dans', dessin)
