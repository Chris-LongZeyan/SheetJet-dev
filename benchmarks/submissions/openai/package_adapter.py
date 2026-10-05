"""Read selected XLSX evidence and preserve packages around artifact-authored values."""
import json, posixpath, re, sys, zipfile
import xml.etree.ElementTree as ET

NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
RID = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'

def sheets(z):
    rels = {x.attrib['Id']: x.attrib['Target'] for x in ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
    return {x.attrib['name']: posixpath.normpath(posixpath.join('xl', rels[x.attrib[RID]])) if not rels[x.attrib[RID]].startswith('/') else rels[x.attrib[RID]].lstrip('/') for x in ET.fromstring(z.read('xl/workbook.xml')).find('s:sheets', NS)}

def strings(z):
    if 'xl/sharedStrings.xml' not in z.namelist(): return []
    return [''.join(t.text or '' for t in x.iter('{'+NS['s']+'}t')) for x in ET.fromstring(z.read('xl/sharedStrings.xml'))]

def value(c, ss):
    t=c.attrib.get('t')
    if t=='inlineStr': return ''.join(x.text or '' for x in c.iter('{'+NS['s']+'}t'))
    v=c.find('s:v',NS)
    if v is None or v.text is None: return None
    if t=='s': return ss[int(v.text)]
    if t in ('str','e'): return v.text
    return float(v.text)

def rows(z, name, ss):
    with z.open(name) as f:
        for event,e in ET.iterparse(f,events=('end',)):
            if e.tag=='{'+NS['s']+'}row':
                yield int(e.attrib['r']), {re.sub(r'\d+$','',c.attrib['r']):value(c,ss) for c in e.findall('s:c',NS)}
                e.clear()

def read(task, inp):
    with zipfile.ZipFile(inp) as z:
        sm=sheets(z); ss=strings(z)
        if task=='edit':
            row=next(r for n,r in rows(z,sm['Assumptions'],ss) if n==1)
            return {'sheet':'Assumptions','updates':{'B1':0.12},'result':{'changed_cell':'Assumptions!B1','previous_value':row['B'],'new_value':0.12,'recalculation':{'performed':False,'needed':True,'reason':'Existing formulas and caches preserved; calculate in Excel when desired.'}}}
        if task=='aggregate':
            labels=[r['A'] for n,r in rows(z,sm['Summary'],ss) if 2<=n<=5]
            totals={label:0 for label in labels}
            for name in ['Period01','Period03']:
                it=rows(z,sm[name],ss); _,header=next(it)
                region=next(c for c,v in header.items() if v=='Region'); amount=next(c for c,v in header.items() if v=='Revenue')
                for _,r in it:
                    if r.get(region) in totals and isinstance(r.get(amount),(int,float)): totals[r[region]]+=r[amount]
            return {'sheet':'Summary','updates':{f'B{i+2}':totals[label] for i,label in enumerate(labels)},'result':{'totals':totals,'selected_sheets':['Period01','Period03']}}
        if task=='wide':
            it=rows(z,sm['Wide'],ss); _,header=next(it)
            region=next(c for c,v in header.items() if v=='Region'); amount=next(c for c,v in header.items() if v=='Amount')
            total=0
            for n,r in it:
                if n>41: break
                if n>=2 and r.get(region)=='APAC' and isinstance(r.get(amount),(int,float)): total+=r[amount]
            return {'sheet':'Summary','updates':{'B2':total},'result':{'total':total,'source_coordinates':{'sheet':'Wide','region':f'{region}2:{region}41','amount':f'{amount}2:{amount}41'}}}
    raise ValueError(task)

def patch(inp, authored, output, plan):
    """Artifact tool owns numeric authoring; this adapter preserves unsupported ZIP parts."""
    with zipfile.ZipFile(inp) as src, zipfile.ZipFile(authored) as donor:
        target=sheets(src)[plan['sheet']]; donor_sheet=sheets(donor)[plan['sheet']]
        cells={c.attrib['r']:c for c in ET.fromstring(donor.read(donor_sheet)).findall('.//s:c',NS)}
        xml=src.read(target).decode('utf-8')
        for address, intended in plan['updates'].items():
            c=cells[address]; v=c.find('s:v',NS)
            if v is None or float(v.text)!=float(intended): raise ValueError('Artifact-authored value mismatch')
            pattern=rf'(<c\b[^>]*\br="{re.escape(address)}"[^>]*>)(.*?)(</c>)'
            def replace(m):
                start=re.sub(r'\s+t="[^"]*"','',m.group(1))
                body=re.sub(r'<(?:f|v|is)\b[^>]*>.*?</(?:f|v|is)>|<(?:f|v|is)\b[^>]*/>','',m.group(2),flags=re.S)
                return start+body+'<v>'+v.text+'</v>'+m.group(3)
            xml,n=re.subn(pattern,replace,xml,count=1,flags=re.S)
            if n!=1: raise ValueError('Expected existing target cell: '+address)
        with zipfile.ZipFile(output,'w') as out:
            out.comment=src.comment
            for info in src.infolist(): out.writestr(info,xml.encode('utf-8') if info.filename==target else src.read(info.filename))
    with zipfile.ZipFile(inp) as a, zipfile.ZipFile(output) as b:
        assert a.namelist()==b.namelist()
        assert all(a.read(n)==b.read(n) for n in a.namelist() if n!=target)

if __name__=='__main__':
    op=sys.argv[1]
    if op=='read':
        data=read(sys.argv[2],sys.argv[3]); print(json.dumps(data,separators=(',',':')))
    elif op=='patch':
        patch(sys.argv[2],sys.argv[3],sys.argv[4],json.loads(sys.argv[5]))
