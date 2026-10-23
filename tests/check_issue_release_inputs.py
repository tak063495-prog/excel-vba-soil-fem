"""Read XLSM XML without Excel; compare all populated canonical input cells."""
import argparse, json, zipfile, hashlib
from decimal import Decimal
from pathlib import Path
from xml.etree import ElementTree as ET

N={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
R='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
SHEETS=('節点データ','形状入力','接合','要素データ','載荷','水条件','材料データ','ステージ','設定')

def inputs(path):
    with zipfile.ZipFile(path) as z:
        strings=[]
        if 'xl/sharedStrings.xml' in z.namelist():
            strings=[''.join(t.text or '' for t in si.findall('.//m:t',N)) for si in ET.fromstring(z.read('xl/sharedStrings.xml')).findall('m:si',N)]
        relations={r.attrib['Id']:r.attrib['Target'] for r in ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
        sheets={s.attrib['name']:relations[s.attrib[R]] for s in ET.fromstring(z.read('xl/workbook.xml')).findall('m:sheets/m:sheet',N)}
        result={}
        for name in SHEETS:
            target=sheets[name];target=target.lstrip('/') if target.startswith('/') else 'xl/'+target
            cells={}
            for c in ET.fromstring(z.read(target)).findall('.//m:sheetData/m:row/m:c',N):
                f=c.find('m:f',N);v=c.find('m:v',N);kind=c.attrib.get('t','n')
                if f is not None: value=('formula',f.text or '',tuple(sorted(f.attrib.items())))
                elif kind=='inlineStr': value=('text',''.join(t.text or '' for t in c.findall('.//m:t',N)))
                elif v is None or v.text is None: continue
                elif kind=='s': value=('text',strings[int(v.text)])
                elif kind=='n': value=('number',str(Decimal(v.text).normalize()))
                else: value=(kind,v.text)
                cells[c.attrib['r']]=value
            result[name]=cells
        return result

if __name__=='__main__':
    ap=argparse.ArgumentParser()
    for name in ('reference','candidate','out'): ap.add_argument('--'+name,type=Path,required=True)
    a=ap.parse_args();old,new=inputs(a.reference),inputs(a.candidate)
    differences={name:[cell for cell in sorted(set(old[name])|set(new[name])) if old[name].get(cell)!=new[name].get(cell)] for name in SHEETS}
    assert not any(differences.values()),differences
    result={'status':'PASS','scope':'All populated cells in nine canonical input/settings sheets, including complete hydraulic/time tables; styles and formula cached values excluded',
            'reference':str(a.reference),'candidate':str(a.candidate),
            'reference_sha256':hashlib.sha256(a.reference.read_bytes()).hexdigest(),
            'candidate_sha256':hashlib.sha256(a.candidate.read_bytes()).hexdigest(),
            'cells':{name:len(new[name]) for name in SHEETS}}
    a.out.write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8');print('PASS: complete canonical input cells',sum(result['cells'].values()))
