"""Read saved OpenXML independently of Excel and compare every fixture input cell."""
import argparse, json, math, posixpath, xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZipFile
NS={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
REL='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'

def workbook_tables(path):
    with ZipFile(path) as z:
        strings=[]
        if 'xl/sharedStrings.xml' in z.namelist():
            strings=[''.join(t.text or '' for t in si.iter('{'+NS['s']+'}t')) for si in ET.fromstring(z.read('xl/sharedStrings.xml'))]
        relationships={r.attrib['Id']:r.attrib['Target'] for r in ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
        tables={}
        for sheet in ET.fromstring(z.read('xl/workbook.xml')).find('s:sheets',NS):
            target=relationships[sheet.attrib[REL]]
            target=target.lstrip('/') if target.startswith('/') else posixpath.normpath('xl/'+target)
            cells={}
            for cell in ET.fromstring(z.read(target)).findall('.//s:sheetData/s:row/s:c',NS):
                kind=cell.attrib.get('t','n');v=cell.find('s:v',NS)
                if kind=='inlineStr':value=''.join(t.text or '' for t in cell.findall('.//s:t',NS))
                elif v is None:value=None
                elif kind=='s':value=strings[int(v.text)]
                elif kind in ['str','e']:value=v.text
                else:value=float(v.text)
                cells[cell.attrib['r']]=value
            tables[sheet.attrib['name']]=cells
        return tables

def equal(expected,actual):
    if expected is None:return actual is None or actual==''
    if isinstance(expected,str):return isinstance(actual,str) and expected==actual
    return isinstance(actual,(float,int)) and math.isclose(expected,actual,rel_tol=1e-10,abs_tol=1e-10)

def column(i):
    return chr(65+i) # fixtures use A:M only

def main():
    parser=argparse.ArgumentParser()
    parser.add_argument('--fixtures',type=Path,default=Path(__file__).with_name('fixtures.json'))
    parser.add_argument('--output',type=Path,required=True)
    args=parser.parse_args();fixtures={f['name']:f for f in json.loads(args.fixtures.read_text(encoding='utf-8-sig'))['fixtures']}
    checks=[]
    def check(case,name,ok,detail):checks.append(dict(case=case,check=name,ok=bool(ok),detail=detail))
    for path in sorted((args.output/'results').glob('*.json')):
        result=json.loads(path.read_text(encoding='utf-8-sig'))
        if 'metrics' not in result:continue
        case=result['case'];fixture=fixtures[result['fixture']];tables=workbook_tables(args.output/'cases'/f'{case}.xlsm')
        for key,sheet,ncols in [('materials','材料データ',13),('nodes','節点データ',9),('elements','要素データ',10),('stages','ステージ',6),('loads','載荷',10),('joints','接合',5)]:
            rows=fixture.get(key,[]);cells=tables[sheet];mismatches=[];inert_default=False
            if key=='loads' and not rows:
                # FEMEnsureLoadSheetLayout inserts this zero DISP example in an
                # empty table. These fixtures have no LOAD stage, and DISP=0 is
                # also ignored by P3ReadLoadingData. Verify every value rather
                # than overlooking extra rows generally.
                default=[1,'TOP','Y','DISP',0,None,None,None,None,None]
                if not any(s[1]=='LOAD' for s in fixture['stages']) and all(equal(v,cells.get(f'{column(i)}2')) for i,v in enumerate(default)):
                    rows=[default];inert_default=True
            for r,row in enumerate(rows,2):
                for c,value in enumerate(row):
                    addr=f'{column(c)}{r}';actual=cells.get(addr)
                    if not equal(value,actual):mismatches.append(dict(cell=addr,expected=value,actual=actual))
            extra=[a for a,v in cells.items() if a.startswith('A') and a[1:].isdigit() and int(a[1:])>len(rows)+1 and v not in (None,'')]
            check(case,key,not mismatches and not extra,dict(compared_cells=len(rows)*ncols,inert_default_row=inert_default,mismatches=mismatches[:5],extra_id_cells=extra[:5]))
        settings=tables['設定'];matches=[a[1:] for a,v in settings.items() if a.startswith('E') and v=='FLOW_POLICY']
        actual=[settings.get('C'+row) for row in matches]
        check(case,'flow_policy',actual==[result['flow_policy']],dict(expected=result['flow_policy'],actual=actual))
    report=dict(ok=all(c['ok'] for c in checks),check_count=len(checks),failure_count=sum(not c['ok'] for c in checks),checks=checks)
    dest=args.output/'verification';dest.mkdir(exist_ok=True)
    (dest/'verify_inputs.json').write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    print('Input checks:',report['check_count'],'failures:',report['failure_count'])
    if not report['ok']:
        print(json.dumps([c for c in checks if not c['ok']],ensure_ascii=False,indent=2));raise SystemExit(1)
if __name__=='__main__':main()
