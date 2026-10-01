"""Seeded PDF page/object edits with independent post-save assertions."""
import pathlib, random, sys, hashlib
import fitz
sys.path.insert(0,str(pathlib.Path(__file__).resolve().parents[1]/'tools'))
from pdf_objects import apply_ops
from safe_output import staged_output
ROOT=pathlib.Path(__file__).resolve().parents[1]
OUT=ROOT/'out/randomized/pdf'; OUT.mkdir(parents=True,exist_ok=True)
rng=random.Random(421001)
image=ROOT/'out/randomized/chart.png'
for i in range(20):
    src,dst=OUT/f'case-{i:02}.pdf',OUT/f'case-{i:02}.edited.pdf'
    doc=fitz.open()
    for j in range(rng.randrange(3,6)):
        p=doc.new_page(width=rng.choice([595,612]),height=rng.choice([792,842]))
        if j: p.insert_text((50,60),f'CONTROL PAGE {j}',fontsize=12)
    doc.save(src); doc.close()
    original=src.read_bytes()
    with fitz.open(src) as doc:
        last_pixels=doc[-1].get_pixmap(matrix=fitz.Matrix(.5,.5)).samples
        angle=rng.choice([90,180,270])
        rows=[['Метрика','Значение']]+[[f'Строка {k}',str(rng.randrange(100,999))] for k in range(rng.randrange(2,6))]
        ops=[{'op':'insert_text','page':1,'rect':[40,25,550,100],'text':f'Случайный PDF {i}'},
             {'op':'insert_image','page':1,'rect':[50,120,300,250],'file':str(image)},
             {'op':'insert_table','page':1,'rect':[40,300,550,620],'rows':rows},
             {'op':'duplicate_page','page':2},
             {'op':'rotate_page','page':3,'angle':angle},
             {'op':'delete_page','page':2}]
        count=len(doc)
        apply_ops(doc,ops,OUT)
        with staged_output(src,dst) as temp:
            doc.save(temp,garbage=3,deflate=True)
            with fitz.open(temp) as check:
                assert len(check)==count
                assert check[1].rotation==angle
                assert 'CONTROL PAGE 1' in check[1].get_text()
                assert f'Случайный PDF {i}' in check[0].get_text()
                assert len(check[0].get_images())==1
                for row in rows:
                    for value in row: assert value in check[0].get_text()
                assert check[-1].get_pixmap(matrix=fitz.Matrix(.5,.5)).samples==last_pixels
    assert src.read_bytes()==original
    print(f'PASS PDF {i}: {count} pages, {len(rows)} table rows, rotation {angle}')
print('ALL GREEN: 20 randomized PDF documents')
