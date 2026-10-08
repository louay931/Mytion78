"""Met à jour Fiches_clients (.xlsm) pour les autocontrôles LLG et LIQ, sans réenregistrer le classeur
(boutons, images et mise en forme copiés octet pour octet) :
  - macro GenererCertificats : LIQ-LIQ-GAZ -> Modele_Autocontrole_LLG.docx, LIQ-LIQ-LIQ -> Modele_Autocontrole_LIQ.docx ;
    le cache compilé (p-code) est retiré pour qu'Excel recompile la macro depuis le source modifié ;
  - onglet « Fiches clients » : colonne R_Pression = pression de service du réservoir (Saisie, section remplissage) ;
  - onglet « Saisie » : libellé de la section AUTOCONTRÔLES.
Usage : python maj_xlsm_llg_liq.py <entrée.xlsm> <sortie.xlsm>"""
import io, os, re, struct, sys, tempfile, zipfile
import olefile
from oletools.olevba import decompress_stream
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from ovba_compression import compress
from ms_cfb.ole_file import OleFile
from ms_cfb.Models.Directories.root_directory import RootDirectory
from ms_cfb.Models.Directories.storage_directory import StorageDirectory
from ms_cfb.Models.Directories.stream_directory import StreamDirectory
from ms_ovba.vbaProject import VbaProject

src, out = os.path.abspath(sys.argv[1]), os.path.abspath(sys.argv[2])
zin = zipfile.ZipFile(src)

# ---------------------------------------------------------------- macro
ANCIEN = '''    If ValeurChamp("T_LIQLIQGAZ") = "X" Then absents = absents & vbCrLf & "  - LIQ-LIQ-GAZ"
    If ValeurChamp("T_LIQLIQLIQ") = "X" Then absents = absents & vbCrLf & "  - LIQ-LIQ-LIQ"
    If absents <> "" Then
        MsgBox "Autocontrôle pas encore disponible pour :" & absents, vbInformation, "Générer les certificats"
    ElseIf n = 0 Then
        MsgBox "Coche (X) la typologie de l'installation finale dans la Saisie" & vbCrLf & _
               "(GAZ-GAZ-GAZ, LIQ-GAZ-GAZ, MESS'AIR ou MESS-VAC).", vbExclamation, "Générer les certificats"
    End If'''.replace('\n', '\r\n')
NOUVEAU = '''    If ValeurChamp("T_LIQLIQGAZ") = "X" Then Generer "Modele_Autocontrole_LLG.docx", "AUTOCONTROLE_LLG": n = n + 1
    If ValeurChamp("T_LIQLIQLIQ") = "X" Then Generer "Modele_Autocontrole_LIQ.docx", "AUTOCONTROLE_LIQ": n = n + 1
    If n = 0 Then
        MsgBox "Coche (X) la typologie de l'installation finale dans la Saisie" & vbCrLf & _
               "(LIQ-LIQ-LIQ, LIQ-LIQ-GAZ, LIQ-GAZ-GAZ, GAZ-GAZ-GAZ, MESS'AIR ou MESS-VAC).", vbExclamation, "Générer les certificats"
    End If'''.replace('\n', '\r\n')

ole = olefile.OleFileIO(io.BytesIO(zin.read('xl/vbaProject.bin')))
from oletools.olevba import VBA_Parser
DEJA_FAIT = 'Modele_Autocontrole_LLG.docx' in ''.join(c for (_, _, _, c) in VBA_Parser(src).extract_macros())
dir_data = bytearray(decompress_stream(bytearray(ole.openstream('VBA/dir').read())))

# parcours du flux dir : modules (nom de flux + offset du source), et remise à zéro des offsets
modules, i, nom = {}, 0, None
while i < len(dir_data):
    rid, size = struct.unpack_from('<HI', dir_data, i)
    if rid == 0x0009:          # PROJECTVERSION : taille réservée, 6 octets de données
        size = 6
    body = i + 6
    if rid == 0x001A:          # MODULESTREAMNAME
        nom = bytes(dir_data[body:body + size]).decode('cp1252')
    elif rid == 0x0031:        # MODULEOFFSET : début du source compressé dans le flux du module
        modules[nom] = struct.unpack_from('<I', dir_data, body)[0]
        struct.pack_into('<I', dir_data, body, 0)
    i = body + size

flux = {}
for nom, offset in ({} if DEJA_FAIT else modules).items():
    data = ole.openstream('VBA/' + nom).read()
    source = bytes(decompress_stream(bytearray(data[offset:])))
    if nom == 'GenererCertificats':
        texte = source.decode('cp1252')
        assert texte.count(ANCIEN) == 1, 'bloc à remplacer introuvable dans la macro'
        texte = texte.replace(ANCIEN, NOUVEAU).replace(
            '    Dim n As Integer, absents As String\r\n', '    Dim n As Integer\r\n')
        source = texte.encode('cp1252')
    flux['VBA/' + nom] = compress(source)           # source seul, sans p-code
if DEJA_FAIT:   # macro déjà modifiée (et compilée par Excel) : projet VBA conservé tel quel
    vba_bin = zin.read('xl/vbaProject.bin')
else:
  flux['VBA/dir'] = compress(bytes(dir_data))
  flux['VBA/_VBA_PROJECT'] = struct.pack('<HHBH', 0x61CC, 0xFFFF, 0x00, 0x0003)  # pas de cache : recompilation
  for n in ('PROJECT', 'PROJECTwm', 'PROJECTlk'):
      if ole.exists(n):
          flux[n] = ole.openstream(n).read()
  autres = ['/'.join(e) for e in ole.listdir() if '/'.join(e) not in flux and not e[-1].startswith('__SRP_')]
  assert not autres, f'flux non gérés : {autres}'   # (__SRP_* = caches, volontairement supprimés)

  tmp = tempfile.mkdtemp(); os.chdir(tmp)
  date = VbaProject().default_date
  racine = RootDirectory(); racine.set_modified(date)
  vba = StorageDirectory('VBA'); vba.set_created(date); vba.set_modified(date)
  for chemin, data in flux.items():
      fichier = chemin.replace('/', '_') + '.bin'
      open(fichier, 'wb').write(data)
      (vba if chemin.startswith('VBA/') else racine).add_directory(StreamDirectory(chemin.split('/')[-1], fichier))
  racine.add_directory(vba)
  o = OleFile(); o.root_directory = racine; o.create_file('vbaProject.bin')
  vba_bin = open('vbaProject.bin', 'rb').read()

# ---------------------------------------------------------------- feuilles
ss = zin.read('xl/sharedStrings.xml').decode('utf8')
uniques = len(re.findall(r'<si>', ss))
def remplace_chaine(ancien, nouveau):
    global ss
    assert ss.count(f'<t>{ancien}</t>') == 1, ancien
    ss = ss.replace(f'<t>{ancien}</t>', f'<t>{nouveau}</t>')
remplace_chaine("AUTOCONTRÔLES  (MESS'AIR / MESS'VAC ; GGG et LGG utilisent les sources ci-dessus)",
                "AUTOCONTRÔLES  (MESS'AIR / MESS'VAC ; GGG, LGG, LLG et LIQ utilisent les sources ci-dessus)")
remplace_chaine('Ex : 10 bar', 'Ex : 10 bar (repris aussi dans l’autocontrôle LIQ)')
idx_entete = uniques
ss = ss.replace('</sst>', '<si><t>R_Pression</t></si></sst>')
ss = re.sub(r'uniqueCount="\d+"', f'uniqueCount="{uniques + 1}"', ss, count=1)
ss = re.sub(r' count="(\d+)"', lambda m: f' count="{int(m.group(1)) + 1}"', ss, count=1)

wbxml = zin.read('xl/workbook.xml').decode('utf8'); wbrels = zin.read('xl/_rels/workbook.xml.rels').decode('utf8')
def feuille(nom):
    m = re.search(r'<sheet name="' + re.escape(nom) + r'" sheetId="(\d+)" r:id="([^"]+)"', wbxml)
    cible = re.search(r'<Relationship Id="' + m.group(2) + r'"[^>]*Target="([^"]+)"', wbrels).group(1)
    return m.group(1), 'xl/' + cible.lstrip('/').replace('xl/', '')
id_fc, chemin_fc = feuille('Fiches clients')
_, chemin_saisie = feuille('Saisie')
fc = zin.read(chemin_fc).decode('utf8')
assert '<c r="BS1"' not in fc and '<dimension ref="A1:BR2"/>' in fc
ligne1 = re.search(r'<c r="BR1"[^>]*>.*?</c>', fc, re.S).group(0)
style = re.search(r' s="(\d+)"', ligne1).group(1)
fc = fc.replace(ligne1, ligne1 + f'<c r="BS1" s="{style}" t="s"><v>{idx_entete}</v></c>')
ligne2 = re.search(r'<c r="BR2"[^>]*>.*?</c>', fc, re.S).group(0)
fc = fc.replace(ligne2, ligne2 + '<c r="BS2" t="str"><f>Saisie!$B$131&amp;""</f><v></v></c>')
fc = fc.replace('<dimension ref="A1:BR2"/>', '<dimension ref="A1:BS2"/>').replace('spans="1:70"', 'spans="1:71"')

saisie = zin.read(chemin_saisie).decode('utf8')
assert re.search(r'<c r="A131"[^>]*><v>(\d+)</v>', saisie).group(1) and 'Pression de service du réservoir (bar)' in ss

cc = zin.read('xl/calcChain.xml').decode('utf8')
cc = cc.replace('</calcChain>', f'<c r="BS2" i="{id_fc}"/></calcChain>')

# ---------------------------------------------------------------- écriture
nouveaux = {'xl/vbaProject.bin': vba_bin, 'xl/sharedStrings.xml': ss.encode('utf8'),
            chemin_fc: fc.encode('utf8'), 'xl/calcChain.xml': cc.encode('utf8')}
with zipfile.ZipFile(out, 'w') as zout:
    for item in zin.infolist():
        zout.writestr(item, nouveaux.get(item.filename, zin.read(item.filename)), compress_type=zipfile.ZIP_DEFLATED)
print('ok', out)
