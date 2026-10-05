import copy, re, sys, openpyxl
from docx import Document
from docx.oxml.ns import qn
from docxcompose.composer import Composer
U='/root/.claude/uploads/405fc51f-0221-5003-9df1-3315ba71f4e8/'
ws=openpyxl.load_workbook(U+'3a4f6c54-Fiches_clients.xlsx',data_only=True)['Fiches clients']
data={h.value:('' if v.value is None else str(v.value).strip()) for h,v in zip(ws[1],ws[2])}
order=['7f1e4e8a-Modele_Certificat_reception_PROVISOIRE.docx','41800d00-Modele_Certificat_tests_alarmes.docx',
       '9b442162-Modele_Certificat_reception_pharmaceutique.docx','289d14eb-Modele_Form004_autocontrole_centrale.docx']
def fill(el):
    for f in list(el.iter(qn('w:fldSimple'))):
        m=re.search(r'MERGEFIELD\s+"?([^\s"]+)',f.get(qn('w:instr')))
        if not m: continue
        runs=f.findall(qn('w:r'))
        r=copy.deepcopy(runs[0]) if runs else f.makeelement(qn('w:r'),{})
        for c in list(r):
            if c.tag!=qn('w:rPr'): r.remove(c)
        t=r.makeelement(qn('w:t'),{}); t.text=data[m.group(1)]; t.set('{http://www.w3.org/XML/1998/namespace}space','preserve'); r.append(t)
        f.addprevious(r); f.getparent().remove(f)
docs=[]
for n in order:
    d=Document(U+n); fill(d.element.body)
    for s in d.sections:
        for hf in (s.header,s.footer,s.first_page_header,s.first_page_footer,s.even_page_header,s.even_page_footer):
            if not hf.is_linked_to_previous: fill(hf._element)
    docs.append(d)
def close_section(d):
    # move the document's final section properties onto its last paragraph so it stays a separate section
    body=d.element.body; sp=body.find(qn('w:sectPr'))
    last=[e for e in body if e.tag!=qn('w:sectPr')][-1]
    if last.tag!=qn('w:p'):
        last=body.makeelement(qn('w:p'),{}); sp.addprevious(last)
    ppr=last.find(qn('w:pPr'))
    if ppr is None:
        ppr=last.makeelement(qn('w:pPr'),{}); last.insert(0,ppr)
    ppr.append(copy.deepcopy(sp))
def sect_refs(d):
    # per section: list of (tag, type, source part) for its own header/footer references
    out=[]
    for sp in d.element.body.iter(qn('w:sectPr')):
        out.append([(e.tag,e.get(qn('w:type')),d.part.related_parts[e.get(qn('r:id'))])
                    for e in sp if e.tag in (qn('w:headerReference'),qn('w:footerReference'))])
    return out
refs=[r for d in docs for r in sect_refs(d)]
for d in docs[:-1]: close_section(d)
final=copy.deepcopy(docs[-1].element.body.find(qn('w:sectPr')))
c=Composer(docs[0])
for d in docs[1:]:
    c.append(d)
# last section = last template's layout; it has no own header refs so it inherits the previous section's
body=c.doc.element.body; old=body.find(qn('w:sectPr'))
for e in final.findall(qn('w:headerReference'))+final.findall(qn('w:footerReference')): final.remove(e)
old.addprevious(final); body.remove(old)
out='/home/user/Mytion78/Certificats_'+re.sub(r'\W+','_',data['Client']).strip('_')+'.docx'
# docxcompose drops the appended templates' headers/footers: re-attach each template's own ones
from io import BytesIO
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.parts.hdrftr import HeaderPart, FooterPart
R='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
pkg=c.doc.part.package
sects=list(body.iter(qn('w:sectPr')))
assert len(sects)==len(refs),(len(sects),len(refs))
for sp,own in zip(sects,refs):
    for e in sp.findall(qn('w:headerReference'))+sp.findall(qn('w:footerReference')): sp.remove(e)
    for i,(tag,typ,src) in enumerate(own):
        new=(HeaderPart if tag==qn('w:headerReference') else FooterPart).new(pkg)
        el=copy.deepcopy(src.element)
        for a in el.iter():
            for k,v in list(a.attrib.items()):
                if k.startswith(R) and v in src.rels:
                    rel=src.rels[v]
                    if rel.is_external: a.set(k,new.relate_to(rel.target_ref,rel.reltype,is_external=True))
                    elif rel.reltype==RT.IMAGE: a.set(k,new.relate_to(pkg.get_or_add_image_part(BytesIO(rel.target_part.blob)),RT.IMAGE))
        new._element=el
        rid=c.doc.part.relate_to(new,RT.HEADER if tag==qn('w:headerReference') else RT.FOOTER)
        ref=sp.makeelement(tag,{qn('w:type'):typ,qn('r:id'):rid}); sp.insert(i,ref)
c.save(out); print(out)
