from pathlib import Path
p = Path('c:/Users/Usuario/Desktop/py/base/proyecto proveedores/app.py')
text = p.read_text(encoding='utf-8')
idx = text.find('HERRAMIENTA')
print('idx', idx)
if idx != -1:
    print(repr(text[idx-80:idx+220]))
