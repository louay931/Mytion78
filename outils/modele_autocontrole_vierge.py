"""Crée un modèle de publipostage à partir d'un formulaire d'autocontrôle VIERGE (Form.008 LLG, Form.009 LIQ) :
un champ de fusion est inséré juste après le « : » de chaque ligne d'en-tête (Client, ordre, adresse…).
Usage : python modele_autocontrole_vierge.py LLG|LIQ <formulaire.docx> <sortie.docx>"""
import re, sys, zipfile

LIGNES = {   # début du texte de la ligne -> (préfixe affiché avant la valeur, champ de « Fiches clients »)
    'LLG': [('Client', '', 'Client'), ('Numéro de Chantier', '', 'NumOrdre'), ('Adresse', '', 'Adresse'),
            ('Code postal', '', 'CodePostalVille'),
            ('Capacité de la source Principale', 'LIQ – ', 'Source1'),
            ('Capacité de la source Attente', 'LIQ – ', 'Source2'),
            ('Capacité de la source de secours', 'GAZ – ', 'Source3')],
    'LIQ': [('Client', '', 'Client'), ('Numéro de Chantier', '', 'NumOrdre'), ('Adresse', '', 'Adresse'),
            ('Code postal', '', 'CodePostalVille'),
            ('Capacité de la source Principale', '', 'Source1'),
            ('PRESSION DE SERVICE', '', 'R_Pression')],
}
TITRES = {'LLG': 'Autocontrôle centrale LIQ/LIQ/GAZ', 'LIQ': 'Autocontrôle une source liquide'}

typ, src, out = sys.argv[1], sys.argv[2], sys.argv[3]
with zipfile.ZipFile(src) as z:
    xml = z.read('word/document.xml').decode('utf8')

def texte(fragment):
    return ''.join(re.findall(r'<w:t(?: [^>]*)?>([^<]*)</w:t>', fragment))

def bleu(rpr):   # même mise en forme que le libellé, en bleu (comme les autres modèles)
    return re.sub(r'(<w:sz )', r'<w:color w:val="0070C0"/>\1', rpr, count=1) if '<w:sz ' in rpr \
        else rpr.replace('</w:rPr>', '<w:color w:val="0070C0"/></w:rPr>')

def run(rpr, t):
    return f'<w:r>{rpr}<w:t xml:space="preserve">{t}</w:t></w:r>' if t else ''

paras = list(re.finditer(r'<w:p[ >].*?</w:p>', xml, flags=re.S))
morceaux, pos, faits = [], 0, {}
for m in paras:
    p = m.group(0); t = texte(p)
    cible = next(((d, pre, ch) for d, pre, ch in LIGNES[typ] if t.startswith(d) and '…' in t), None)
    if not cible:
        continue
    debut, prefixe, champ = cible
    # 1er run contenant le « : » : on le coupe juste après les deux-points (et l'espace qui suit)
    for r in re.finditer(r'<w:r(?: [^>]*)?>(<w:rPr>.*?</w:rPr>)?<w:t(?: [^>]*)?>([^<]*)</w:t></w:r>', p, flags=re.S):
        if ':' in r.group(2):
            rpr = r.group(1) or '<w:rPr></w:rPr>'
            txt = r.group(2); k = txt.index(':') + 1
            while k < len(txt) and txt[k] in '  ':
                k += 1
            if k == len(txt) or txt[k - 1] == ':':
                txt = txt[:k] + ' ' + txt[k:]; k += 1   # espace entre « : » et la valeur
            neuf = (run(rpr, txt[:k]) + run(bleu(rpr), prefixe)
                    + f'<w:fldSimple w:instr=" MERGEFIELD {champ} "><w:r>{bleu(rpr)}<w:t>«{champ}»</w:t></w:r></w:fldSimple>'
                    + run(rpr, txt[k:]))
            p = p[:r.start()] + neuf + p[r.end():]
            break
    else:
        raise SystemExit(f'pas de « : » dans la ligne {debut}')
    morceaux += [xml[pos:m.start()], p]; pos = m.end()
    faits[champ] = faits.get(champ, 0) + 1
morceaux.append(xml[pos:])
xml = ''.join(morceaux)
manque = [ch for _, _, ch in LIGNES[typ] if ch not in faits]
assert not manque, f'lignes non trouvées : {manque}'

with zipfile.ZipFile(src) as zin, zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data = xml.encode('utf8') if item.filename == 'word/document.xml' else zin.read(item.filename)
        if item.filename == 'docProps/core.xml':
            data = re.sub(rb'<dc:title>[^<]*</dc:title>|<dc:title/>', f'<dc:title>{TITRES[typ]}</dc:title>'.encode(), data)
        zout.writestr(item, data)
print('ok', out, faits)
