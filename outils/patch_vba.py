"""Modifie le code source d'un module VBA dans un .xlsm sans réenregistrer le classeur :
les autres parties du fichier sont recopiées octet pour octet, le cache compilé (p-code, __SRP_*)
est retiré pour qu'Excel recompile depuis le source modifié à l'ouverture."""
import io, os, struct, sys, tempfile, zipfile
import olefile
from oletools.olevba import decompress_stream
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from ovba_compression import compress
from ms_cfb.ole_file import OleFile
from ms_cfb.Models.Directories.root_directory import RootDirectory
from ms_cfb.Models.Directories.storage_directory import StorageDirectory
from ms_cfb.Models.Directories.stream_directory import StreamDirectory
from ms_ovba.vbaProject import VbaProject


def sources_vba(vba_bin):
    """{nom du module: source (str)} et le flux dir décompressé (offsets remis à zéro)."""
    ole = olefile.OleFileIO(io.BytesIO(vba_bin))
    d = bytearray(decompress_stream(bytearray(ole.openstream('VBA/dir').read())))
    offsets, i, nom = {}, 0, None
    while i < len(d):
        rid, size = struct.unpack_from('<HI', d, i)
        if rid == 0x0009:                  # PROJECTVERSION : 6 octets de données
            size = 6
        if rid == 0x001A:                  # MODULESTREAMNAME
            nom = bytes(d[i + 6:i + 6 + size]).decode('cp1252')
        elif rid == 0x0031:                # MODULEOFFSET -> 0 (plus de p-code avant le source)
            offsets[nom] = struct.unpack_from('<I', d, i + 6)[0]
            struct.pack_into('<I', d, i + 6, 0)
        i += 6 + size
    sources = {n: bytes(decompress_stream(bytearray(ole.openstream('VBA/' + n).read()[o:]))).decode('cp1252')
               for n, o in offsets.items()}
    return ole, sources, bytes(d)


def reconstruire(ole, sources, dir_data):
    flux = {'VBA/' + n: compress(s.encode('cp1252')) for n, s in sources.items()}
    flux['VBA/dir'] = compress(dir_data)
    flux['VBA/_VBA_PROJECT'] = struct.pack('<HHBH', 0x61CC, 0xFFFF, 0x00, 0x0003)
    for n in ('PROJECT', 'PROJECTwm', 'PROJECTlk'):
        if ole.exists(n):
            flux[n] = ole.openstream(n).read()
    autres = ['/'.join(e) for e in ole.listdir() if '/'.join(e) not in flux and not e[-1].startswith('__SRP_')]
    assert not autres, f'flux non gérés : {autres}'
    ici = os.getcwd(); tmp = tempfile.mkdtemp(); os.chdir(tmp)
    try:
        date = VbaProject().default_date
        racine = RootDirectory(); racine.set_modified(date)
        vba = StorageDirectory('VBA'); vba.set_created(date); vba.set_modified(date)
        for chemin, data in flux.items():
            fichier = chemin.replace('/', '_') + '.bin'
            open(fichier, 'wb').write(data)
            (vba if chemin.startswith('VBA/') else racine).add_directory(StreamDirectory(chemin.split('/')[-1], fichier))
        racine.add_directory(vba)
        o = OleFile(); o.root_directory = racine; o.create_file('vbaProject.bin')
        return open('vbaProject.bin', 'rb').read()
    finally:
        os.chdir(ici)


def patcher(entree, sortie, module, remplacements):
    """remplacements : liste de (ancien, nouveau) ; chaque « ancien » doit apparaître exactement une fois."""
    zin = zipfile.ZipFile(entree)
    ole, sources, dir_data = sources_vba(zin.read('xl/vbaProject.bin'))
    texte = sources[module]
    for ancien, nouveau in remplacements:
        ancien, nouveau = ancien.replace('\r\n', '\n').replace('\n', '\r\n'), nouveau.replace('\r\n', '\n').replace('\n', '\r\n')
        assert texte.count(ancien) == 1, f'texte à remplacer trouvé {texte.count(ancien)} fois :\n{ancien}'
        texte = texte.replace(ancien, nouveau)
    sources[module] = texte
    vba_bin = reconstruire(ole, sources, dir_data)
    with zipfile.ZipFile(sortie, 'w') as zout:
        for item in zin.infolist():
            data = vba_bin if item.filename == 'xl/vbaProject.bin' else zin.read(item.filename)
            zout.writestr(item, data, compress_type=zipfile.ZIP_DEFLATED)
