"""Extract original concept-image panels from a blue-divider quotation screenshot.

The agent reads names visually; this script does not invent OCR results.
Requires Pillow and numpy. --boxes accepts JSON pixel boxes for other layouts.
"""
import argparse
import json
from pathlib import Path
import numpy as np
from PIL import Image, ImageOps, ImageDraw


def extract(source, output, expected=None, boxes=None):
    output = Path(output)
    output.mkdir(parents=True, exist_ok=True)
    im = ImageOps.exif_transpose(Image.open(source)).convert('RGB')
    w, h = im.size
    if boxes is None:
        a = np.asarray(im.resize((800, round(h * 800 / w)))).astype(np.int16)
        blue = (a[:, :, 2] - a[:, :, 0] > 45) & (a[:, :, 2] - a[:, :, 1] > 20)
        rows = np.flatnonzero(blue[:, 25:-25].mean(axis=1) > .7)
        groups = np.split(rows, np.flatnonzero(np.diff(rows) > 1) + 1)
        lines = [round(float(g[-1]) * w / 800) for g in groups
                 if len(g) and len(g) < 8 and g[0] > 40]
        if not lines or (expected is not None and len(lines) != expected):
            raise ValueError(f'Detected {len(lines)} panels, expected {expected}; inspect source and use --boxes.')
        side = round(w * .377)
        boxes = [[round(w*.035), y+round(w*.012), round(w*.035)+side,
                  min(h, y+round(w*.012)+side)] for y in lines]
    if expected is not None and len(boxes) != expected:
        raise ValueError('Box count does not match expected item count')
    records = []
    contact = Image.new('RGB', (900, ((len(boxes)+2)//3)*340), 'white')
    draw = ImageDraw.Draw(contact)
    for n, box in enumerate(boxes, 1):
        x1,y1,x2,y2 = map(int, box)
        if not (0 <= x1 < x2 <= w and 0 <= y1 < y2 <= h):
            raise ValueError(f'Invalid box {box}')
        panel = im.crop((x1,y1,x2,y2))
        path = output / f'item-{n:02}.png'
        panel.save(path)
        thumb = ImageOps.contain(panel, (280,300))
        col,row = (n-1)%3, (n-1)//3
        contact.paste(thumb, (col*300+(300-thumb.width)//2,row*340+25))
        draw.text((col*300+12,row*340+7), str(n), fill='black')
        records.append({'index':n,'box':[x1,y1,x2,y2],'image':str(path.resolve())})
    contact.save(output/'contact-sheet.png')
    payload = {'source_size':[w,h], 'items':records}
    (output/'panels.json').write_text(json.dumps(payload,ensure_ascii=False,indent=2),encoding='utf-8')
    return payload


if __name__ == '__main__':
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('source'); p.add_argument('output'); p.add_argument('--expected',type=int)
    p.add_argument('--boxes',help='JSON file containing pixel boxes [[left,top,right,bottom],...]')
    a=p.parse_args()
    boxes=json.loads(Path(a.boxes).read_text(encoding='utf-8')) if a.boxes else None
    print(json.dumps(extract(a.source,a.output,a.expected,boxes),ensure_ascii=False))
