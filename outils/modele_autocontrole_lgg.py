"""Crée Modele_Autocontrole_LGG.docx à partir du formulaire FR.50.07.Form.005 (centrale LIQ/GAZ/GAZ) :
les valeurs d'exemple (SMR BONVERT…) de l'en-tête sont remplacées par des champs de fusion
lus dans l'onglet « Fiches clients » (mêmes noms que le modèle autocontrôle GGG).
Usage : python modele_autocontrole_lgg.py <Form.005 LGG.docx> <sortie.docx>"""
import re, sys, zipfile

src, out = sys.argv[1], sys.argv[2]

# texte d'exemple -> (préfixe conservé, champ de fusion)
REMPLACEMENTS = [
    ('SMR BONVERT', '', 'Client'),
    ('84009437', '', 'NumOrdre'),
    ('ROUTE DE BONVERT', '', 'Adresse'),
    ('42300 MABLY', '', 'CodePostalVille'),
    ('LIQ - 3 000L', 'LIQ – ', 'Source1'),            # source principale (liquide)
    ('GAZ – 1*CV18 (190m3)', 'GAZ – ', 'Source2'),    # source attente
    ('GAZ – 1*CV18 (190m3)', 'GAZ – ', 'Source3'),    # source de secours
]

with zipfile.ZipFile(src) as z:
    xml = z.read('word/document.xml').decode('utf8')

for exemple, prefixe, champ in REMPLACEMENTS:
    # run d'exemple : <w:r ...><w:rPr>...</w:rPr><w:t>exemple</w:t></w:r>  (1re occurrence restante)
    motif = re.compile(r'<w:r(?: [^>]*)?>(<w:rPr>(?:(?!</w:rPr>).)*</w:rPr>)<w:t(?: [^>]*)?>'
                       + re.escape(exemple) + r'</w:t></w:r>', re.S)
    m = motif.search(xml)
    assert m, f'texte introuvable : {exemple}'
    rpr = m.group(1)
    neuf = ''
    if prefixe:
        neuf += f'<w:r>{rpr}<w:t xml:space="preserve">{prefixe}</w:t></w:r>'
    neuf += (f'<w:fldSimple w:instr=" MERGEFIELD {champ} "><w:r>{rpr}'
             f'<w:t>«{champ}»</w:t></w:r></w:fldSimple>')
    xml = xml[:m.start()] + neuf + xml[m.end():]

with zipfile.ZipFile(src) as zin, zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data = xml.encode('utf8') if item.filename == 'word/document.xml' else zin.read(item.filename)
        if item.filename == 'docProps/core.xml':   # titre du document
            data = re.sub(rb'<dc:title>[^<]*</dc:title>', '<dc:title>Autocontrôle centrale LIQ/GAZ/GAZ</dc:title>'.encode(), data)
        zout.writestr(item, data)
print('ok', out)
