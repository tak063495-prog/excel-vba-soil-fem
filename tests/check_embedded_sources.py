"""Check exported embedded VBA without ignoring case inside string literals/comments."""
from pathlib import Path
import argparse,json,hashlib

def normalize(source):
    lines=[]
    for line in source.splitlines():
        if line.startswith('Attribute VB_'):continue
        out=[];quoted=False;i=0
        while i<len(line):
            c=line[i]
            if c=='"':
                out.append(c)
                if quoted and i+1<len(line) and line[i+1]=='"':out.append('"');i+=2;continue
                quoted=not quoted
            elif c=="'" and not quoted:out.append(line[i:]);break
            else:out.append(c if quoted else c.lower())
            i+=1
        lines.append(''.join(out).rstrip())
    return '\n'.join(lines).rstrip()

if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('dump');p.add_argument('report');a=p.parse_args()
    root=Path(__file__).resolve().parents[1]
    embedded=json.loads(Path(a.dump).read_text(encoding='utf-8-sig'));hashes={}
    for name,code in embedded.items():
        assert name.isidentifier(),name
        source=root/'src/vba'/(name+('.cls' if name=='ThisWorkbook' else '.bas'))
        assert normalize(source.read_text(encoding='utf-8-sig'))==normalize(code),name
        hashes[name]=hashlib.sha256(source.read_bytes()).hexdigest()
    assert len(hashes)==33,len(hashes)
    report=dict(status='PASS',modules=hashes,normalization='VBE identifier/keyword case outside string literals and apostrophe comments; line endings/trailing whitespace/Attribute VB_. String literals and comment text preserved.')
    Path(a.report).write_bytes((json.dumps(report,indent=2)+'\n').encode());print('PASS: 33 embedded VBA modules; string literals and comments preserved.')
