"""Build a small vector toolbar symbol and raster previews from the same paths."""
from pathlib import Path
from PIL import Image, ImageDraw
import math

folder = Path(__file__).resolve().parent / 'icons'
folder.mkdir(exist_ok=True)
paths = [
    ('#ffffff', 1.7, [(7, 17), (7, 5), (16, 5), (20, 9), (20, 17)]),
    ('#ffffff', 1.5, [(16, 5), (16, 9), (20, 9)]),
    ('#ffffff', 2.5, [(13, 13), (26, 13)]),
    ('#ffffff', 2.5, [(23, 10), (26, 13), (23, 16)]),
    ('#ffffff', 1.6, [(6, 26), (6, 21), (9, 21), (11, 22.5), (11, 24.5), (9, 26), (6, 26)]),
    ('#ffffff', 1.6, [(14, 21), (18, 26)]),
    ('#ffffff', 1.6, [(18, 21), (14, 26)]),
    ('#ffffff', 1.6, [(22, 26), (22, 21), (27, 21)]),
    ('#ffffff', 1.6, [(22, 23.5), (26, 23.5)]),
]
svg = ['<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32" width="32" height="32">',
       '<rect x="1" y="1" width="30" height="30" rx="6" fill="#087bd5" stroke="#bde7ff" stroke-width="1"/>']
for color, width, points in paths:
    value = ' '.join(f'{x},{y}' for x, y in points)
    svg.append(f'<polyline points="{value}" fill="none" stroke="{color}" stroke-width="{width}" stroke-linecap="round" stroke-linejoin="round"/>')
svg.append('</svg>')
(folder / 'CorelDXF.svg').write_text('\n'.join(svg), encoding='ascii')
for size in (24, 32, 64):
    scale = size / 32 * 8
    bitmap = Image.new('RGBA', (size * 8, size * 8), (0, 0, 0, 0))
    draw = ImageDraw.Draw(bitmap)
    draw.rounded_rectangle((scale, scale, 31 * scale, 31 * scale), radius=6 * scale, fill='#087bd5', outline='#bde7ff', width=round(scale))
    for color, width, points in paths:
        draw.line([(x * scale, y * scale) for x, y in points], fill=color, width=round(width * scale), joint='curve')
        radius = width * scale / 2
        for x, y in (points[0], points[-1]):
            draw.ellipse((x * scale-radius, y * scale-radius, x * scale+radius, y * scale+radius), fill=color)
    bitmap = bitmap.resize((size, size), Image.Resampling.LANCZOS)
    bitmap.save(folder / f'CorelDXF_{size}.png')
    if size == 64:
        bitmap.save(folder / 'CorelDXF.ico', bitmap_format='bmp', sizes=[(16,16), (24,24), (32,32), (48,48), (64,64)])
    background = Image.new('RGB', bitmap.size, '#f0f0f0')
    background.paste(bitmap, mask=bitmap.getchannel('A'))
    background.save(folder / f'CorelDXF_{size}.bmp')
print(folder)

# A drawn eight-tooth gear, not a font glyph (identical on every workstation).
scale = 8
bitmap = Image.new('RGB', (20 * scale, 20 * scale), '#f0f0f0')
draw = ImageDraw.Draw(bitmap)
points = []
for tooth in range(8):
    for offset, radius in ((0, 7), (.18, 9), (.52, 9), (.70, 7)):
        angle = (tooth + offset) * math.tau / 8
        points.append(((10 + math.cos(angle) * radius) * scale,
                       (10 + math.sin(angle) * radius) * scale))
draw.polygon(points, fill='#164754')
draw.ellipse((6*scale, 6*scale, 14*scale, 14*scale), fill='#f0f0f0')
draw.ellipse((8*scale, 8*scale, 12*scale, 12*scale), fill='#164754')
bitmap.resize((20, 20), Image.Resampling.LANCZOS).save(folder / 'Settings_20.bmp')

# Distinct inspection and connection symbols, readable on dark toolbars.
for name,bg in [('VectorCheck','#dd3c42'),('VectorJoin','#15976a')]:
    factor=16
    pic=Image.new('RGBA',(512,512),(0,0,0,0));d=ImageDraw.Draw(pic)
    d.rounded_rectangle((factor,factor,31*factor,31*factor),radius=6*factor,fill=bg,outline='white',width=factor)
    if name=='VectorCheck':
        d.ellipse((6*factor,5*factor,22*factor,21*factor),outline='white',width=2*factor)
        d.line((20*factor,20*factor,27*factor,27*factor),fill='white',width=3*factor)
        d.line((10*factor,9*factor,18*factor,17*factor),fill='white',width=2*factor)
        d.line((18*factor,9*factor,10*factor,17*factor),fill='white',width=2*factor)
    else:
        d.rounded_rectangle((5*factor,10*factor,18*factor,22*factor),radius=5*factor,outline='white',width=2*factor)
        d.rounded_rectangle((14*factor,10*factor,27*factor,22*factor),radius=5*factor,outline='white',width=2*factor)
    pic.resize((64,64),Image.Resampling.LANCZOS).save(folder/(name+'.ico'),bitmap_format='bmp',sizes=[(16,16),(24,24),(32,32),(48,48),(64,64)])

# Layer stack and question badge distinguish help from export even without text.
for size in (24, 32, 64):
    factor = size / 32 * 8
    bitmap = Image.new('RGBA', (size * 8, size * 8), (0, 0, 0, 0))
    draw = ImageDraw.Draw(bitmap)
    draw.rounded_rectangle((factor, factor, 31*factor, 31*factor), radius=6*factor,
                           fill='#ffc247', outline='#fff0bc', width=round(factor))
    for y in (22, 17, 12):
        points = [(5, y), (13, y-4), (21, y), (13, y+4), (5, y)]
        draw.line([(x*factor, v*factor) for x,v in points], fill='#263342', width=round(2*factor), joint='curve')
    draw.ellipse((17*factor, 3*factor, 30*factor, 17*factor), fill='#263342')
    points = [(21, 7), (22, 6), (25, 6), (26, 8), (23.5, 10), (23.5, 11)]
    draw.line([(x*factor, y*factor) for x,y in points], fill='white', width=round(1.5*factor), joint='curve')
    draw.ellipse((22.8*factor, 13*factor, 24.2*factor, 14.4*factor), fill='white')
    bitmap = bitmap.resize((size,size), Image.Resampling.LANCZOS)
    bitmap.save(folder / f'LayerHelp_{size}.png')
    background = Image.new('RGB', bitmap.size, '#f0f0f0')
    background.paste(bitmap, mask=bitmap.getchannel('A'))
    background.save(folder / f'LayerHelp_{size}.bmp')
    if size == 64:
        bitmap.save(folder / 'LayerHelp.ico', bitmap_format='bmp', sizes=[(16,16), (24,24), (32,32), (48,48), (64,64)])
