"""Tanam scripts/data.json ke index.html (ganti blok `const D=...`)."""
t=open('index.html',encoding='utf-8').read();d=open('scripts/data.json',encoding='utf-8').read()
a=t.index('const D=')+len('const D=');b=t.index(',N=Object.fromEntries')
open('index.html','w',encoding='utf-8').write(t[:a]+d+t[b:]);print('index.html diperbarui')
