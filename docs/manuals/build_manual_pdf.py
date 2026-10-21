import markdown,sys
md=open('2DSoilFEM_Manual.md',encoding='utf8').read()
body=markdown.markdown(md,extensions=['tables'])
css="""@page{size:A4;margin:16mm 15mm}body{font-family:'Noto Sans CJK JP','Noto Sans JP',sans-serif;font-size:10.5pt;line-height:1.7;color:#222}
h1{font-size:22pt;border-bottom:3px solid #2f5d50}h2{font-size:15pt;background:#2f5d50;color:#fff;padding:4px 8px;break-before:page}
h3{font-size:12pt;border-left:5px solid #2f5d50;padding-left:6px}table{border-collapse:collapse;width:100%;font-size:9pt;margin:6px 0}th,td{border:1px solid #aaa;padding:3px 5px;vertical-align:top}th{background:#e6efe9}
img{max-width:100%;border:1px solid #999;display:block;margin:8px auto 2px;break-inside:avoid}code{background:#f3f3f3;padding:0 3px}"""
open('manual.html','w',encoding='utf8').write(f"<html><head><meta charset=utf-8><style>{css}</style></head><body>{body}</body></html>")
