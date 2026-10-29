# xemnguon.py (Quy tac 39): boc HYPERLINK vao o Don gia goc VL/NC (MEP, bang TV). Dung: python3 -I xemnguon.py SRC.xlsx DST.xlsx
import sys,zipfile,re
from lxml import etree
M='http://schemas.openxmlformats.org/spreadsheetml/2006/main'; ns={'m':M}
def q(t): return '{%s}%s'%(M,t)
SRC,DST=sys.argv[1],sys.argv[2]
SHEET='xl/worksheets/sheet2.xml'
z=zipfile.ZipFile(SRC)
S=[''.join(si.itertext()) for si in etree.fromstring(z.read('xl/sharedStrings.xml'))]
def cval(c):
    v=c.find('m:v',ns)
    if v is None: return None
    return S[int(v.text)] if c.get('t')=='s' else v.text
sh=etree.fromstring(z.read(SHEET)); sd=sh.find('m:sheetData',ns)
hdr={}
for r in sd:
    if r.get('r')=='4':
        for c in r:
            i=c.find('m:is',ns); hdr[cval(c) if i is None else ''.join(i.itertext())]=re.sub(r'\d','',c.get('r'))
cAX,cBB,cBC,cBD=[hdr[k] for k in ['Nhóm VT','Loại dòng','Mã tra cứu','Tên tra cứu']]
assert 'Xem nguồn VL' not in hdr
def wrap(orig,n,tvcol):
    ax,bc,bd=f'${cAX}{n}',f'${cBC}{n}',f'${cBD}{n}'
    I=(f'MATCH(1,INDEX(((({ax}="")+({ax}="Chưa xác định")+(TV[Loại vật liệu]={ax}))>0)'
       f'*((({bc}="")+(TV[Mã hiệu]={bc}))>0)*((({bd}="")+(TV[Tên công việc]={bd}))>0),0),0)')
    link=(f'HYPERLINK("#\'2 Thu vien gia\'!"&ADDRESS(ROW(INDEX(TV[{tvcol}],{I})),MIN(COLUMN(TV[{tvcol}])),4),{orig})')
    return f'IF(AND({bc}="",{bd}=""),{orig},IFERROR({link},{orig}))'
st={'S':0,'T':0,'skip_shared':0,'skip_other':0}
for r in sd:
    n=int(r.get('r'))
    if n<6: continue
    cells={re.sub(r'\d','',c.get('r')):c for c in r}
    if cBB not in cells or cval(cells[cBB]) not in ('Dòng','Con'): continue
    for col,key,tvcol in (('S','VL','Giá VL thư viện (thô)'),('T','NC','Giá NC thư viện (thô)')):
        c=cells.get(col); f=c.find('m:f',ns) if c is not None else None
        if f is None or not f.text or f'TV[Giá {key} gốc]' not in f.text:
            if f is not None and f.get('t')=='shared': st['skip_shared']+=1
            continue
        if f.get('t') in ('shared','array'): st['skip_shared']+=1; continue
        f.text=wrap(f.text,n,tvcol); st[col]+=1
print(st)
wb=z.read('xl/workbook.xml').decode('utf8')
if 'fullCalcOnLoad' not in wb:
    wb=re.sub(r'<calcPr([^>]*?)(/?)>',lambda m:'<calcPr'+m.group(1)+' fullCalcOnLoad="1"'+m.group(2)+'>',wb,count=1)
out=zipfile.ZipFile(DST,'w',zipfile.ZIP_DEFLATED)
for it in z.infolist():
    d=z.read(it.filename)
    if it.filename==SHEET: d=etree.tostring(sh,xml_declaration=True,encoding='UTF-8',standalone=True)
    if it.filename=='xl/workbook.xml': d=wb.encode('utf8')
    out.writestr(it,d)
out.close()
