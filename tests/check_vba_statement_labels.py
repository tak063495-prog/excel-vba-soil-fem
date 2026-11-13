"""Reject parameterless procedure calls accidentally parsed as VBA labels."""
from pathlib import Path
import re

source_dir=Path(__file__).resolve().parents[1]/'src/vba'
sources={p:p.read_text(encoding='utf-8-sig') for p in source_dir.iterdir() if p.suffix in ('.bas','.cls')}
declaration=re.compile(r'^\s*(?:(?:Public|Private|Friend|Static)\s+)*(?:Sub|Function)\s+(\w+)\s*\(\s*\)',re.I|re.M)
names={name.lower() for text in sources.values() for name in declaration.findall(text)}
bad=[]
for path,text in sources.items():
 for row,line in enumerate(text.splitlines(),1):
  match=re.match(r'^\s*(\w+)\s*:(?!=)',line)
  if match and match[1].lower() in names:bad.append((path.name,row,line.strip()))
assert not bad,('VBA interprets these as labels, not procedure calls',bad)
print('PASS:',len(sources),'VBA modules have no parameterless procedure labels')
