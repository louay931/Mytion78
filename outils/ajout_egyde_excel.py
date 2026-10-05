"""Ajoute la section DEMANDE EGYDE à Fiches_clients.xlsx (onglet Saisie + colonnes Fiches clients)."""
import copy, sys, openpyxl
from openpyxl.worksheet.datavalidation import DataValidation
src, out = sys.argv[1], sys.argv[2]
wb = openpyxl.load_workbook(src)
ws, fc, lm = wb['Saisie'], wb['Fiches clients'], wb['Lisez-moi']

def style(dst, model):
    for a in ('font', 'fill', 'border', 'alignment', 'number_format', 'protection'):
        setattr(dst, a, copy.copy(getattr(model, a)))

TITRE, LIB, VAL, AIDE = ws['A21'], ws['A22'], ws['B22'], ws['C22']
# (libellé, nom du champ Word, aide, type)  type: t=texte, d=date, on=liste oui/non
champs = [('Date de la demande', 'E_DateDemande', 'Format jj/mm/aaaa', 'd'),
          ('Date de M.E.S. souhaitée', 'E_DateMES', '3 jours ouvrés minimum si type existant', 'd'),
          ('Nom du site (création)', 'E_NomSite', 'Ex : CH DE PITHIVIERS / PITHIVIERS (45)', 't'),
          ('Type de site', 'E_TypeSite', 'Ex : LGG', 't'),
          ('N° de téléphone du site', 'E_TelSite', '', 't'),
          ('N° identification', 'E_NumIdentification', '', 't'),
          ('N° esclave', 'E_NumEsclave', '', 't'),
          ('N° astreinte', 'E_Astreinte', '', 't'),
          ('Source principale', 'E_SourcePrincipale', 'Ex : 3000L', 't'),
          ('Source attente', 'E_SourceAttente', 'Ex : 2XCV18', 't'),
          ('Source secours', 'E_SourceSecours', 'Ex : 3B50', 't'),
          ('Commercial responsable', 'E_Commercial', '', 't')]
contact = [('Nom', 'Nom', '', 't'), ('Prénom', 'Prenom', '', 't'), ('Adresse mail', 'Mail', '', 't'),
           ('N° de téléphone', 'Tel', '', 't'), ('Service', 'Service', '', 't'),
           ('Identifiant', 'Identifiant', '', 't'), ('Mot de passe', 'MotDePasse', '', 't'),
           ('Type d’accès client', 'TypeAcces', 'Ex : etendu', 't'), ('Ajout alarme', 'AjoutAlarme', 'oui / non', 'on')]
blocs = [('DEMANDE EGYDE – SITE  (modèle Demande EGYDE uniquement ; client, adresse, CP et ville = INFOS CLIENT)', champs)]
for n in (1, 2, 3):
    blocs.append((f'DEMANDE EGYDE – CONTACT CLIENT {n}  (laisser vide si pas de contact {n})',
                  [(l, f'E_C{n}_{k}', a, t) for l, k, a, t in contact]))

assert ws.max_row <= 40, 'la section EGYDE semble déjà présente'
dv_on = DataValidation(type='list', formula1='"oui,non"', allow_blank=True)
ws.add_data_validation(dv_on)
row, col = 41, fc.max_column + 1
for titre, lignes in blocs:
    ws.cell(row, 1, titre); style(ws.cell(row, 1), TITRE)
    for c in (2, 3): style(ws.cell(row, c), ws.cell(21, c))
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
    row += 1
    for lib, nom, aide, typ in lignes:
        a, b, c = ws.cell(row, 1, lib), ws.cell(row, 2), ws.cell(row, 3, aide or None)
        style(a, LIB); style(b, VAL); style(c, AIDE)
        b.number_format = 'DD/MM/YYYY' if typ == 'd' else '@'
        if typ == 'on': dv_on.add(b.coordinate)
        ref = f'Saisie!$B${row}'
        fc.cell(1, col, nom); style(fc.cell(1, col), fc.cell(1, 1))
        # dates écrites jj/mm/aaaa sans TEXT(..,"jj/mm/aaaa") qui dépend de la langue d'Excel
        fc.cell(2, col, f'=IF({ref}="","",TEXT(DAY({ref}),"00")&"/"&TEXT(MONTH({ref}),"00")&"/"&YEAR({ref}))'
                if typ == 'd' else f'={ref}&""')
        style(fc.cell(2, col), fc.cell(2, 1))
        fc.column_dimensions[fc.cell(1, col).column_letter].width = 18
        col += 1; row += 1
    row += 1

# Lisez-moi : point 6 inséré après le point 5
for r in range(1, lm.max_row + 1):
    if str(lm.cell(r, 1).value or '').startswith('5.'):
        lm.insert_rows(r + 1)
        for im in lm._images:  # les captures ne suivent pas insert_rows
            if im.anchor._from.row >= r:
                im.anchor._from.row += 1; im.anchor.to.row += 1
        lm.cell(r + 1, 1, '6. Section « DEMANDE EGYDE » (en bas de Saisie) : ne sert qu’au modèle Demande EGYDE. '
                          'Le nom du site client, l’adresse, le CP et la ville sont repris des INFOS CLIENT.')
        style(lm.cell(r + 1, 1), lm.cell(r, 1))
        break
wb.calculation.fullCalcOnLoad = True
wb.save(out)
print('ok', fc.max_column, 'colonnes')
