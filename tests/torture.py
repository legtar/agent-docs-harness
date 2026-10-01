"""Adversarial behavioral tests; fixtures saved under out/hardcore."""
import contextlib, copy, io, json, pathlib, random, subprocess, sys, unittest, zipfile
from lxml import etree
from docx import Document
from docx.oxml import OxmlElement
from docx.shared import Cm
import fitz
ROOT=pathlib.Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'tools'))
import docx_edit as de
from docx_edit import q, Editor, OpError
from pdf_objects import apply_ops
from qa import Report, check_pdf
from render import render
OUT=ROOT/'out/hardcore'; OUT.mkdir(parents=True,exist_ok=True)


def xml(s): return etree.fromstring(f'<w:p xmlns:w="{de.W}">{s}</w:p>')
def trees(doc): return {'word/document.xml':etree.fromstring(doc._element.xml.encode())}
def replace(p,old,new,ed=None):
    with contextlib.redirect_stdout(io.StringIO()): de.do_replace({'find':old,'replace':new}, {'word/document.xml':p},ed or Editor(),set())
def visible(p):
    return ''.join(t.text or '' for t in p.iter(q('t')) if not any(a.tag in (q('del'),q('moveFrom')) for a in t.iterancestors()))
def cli(tool,*args):
    return subprocess.run([sys.executable,str(ROOT/'tools'/tool),*map(str,args)],cwd=ROOT,capture_output=True,text=True,encoding='utf-8',timeout=180)


class Torture(unittest.TestCase):
    def test_cross_run_property_1000_plain_and_tracked(self):
        rng=random.Random(191026)
        alphabet='абвXYZ 0123-«»'
        for i in range(1000):
            old=''.join(rng.choice(alphabet) for _ in range(rng.randrange(1,28)))
            new=''.join(rng.choice(alphabet) for _ in range(rng.randrange(0,28)))
            p=xml('')
            for char in old:
                r=etree.SubElement(p,q('r')); prop=etree.SubElement(r,q('rPr'))
                if rng.choice([True,False]): etree.SubElement(prop,q('b'))
                etree.SubElement(r,q('t')).text=char
            replace(p,old,new,Editor('Auditor') if i%2 else Editor())
            self.assertEqual(visible(p),new,f'case {i}: {old!r} -> {new!r}')
            etree.fromstring(etree.tostring(p))

    def test_simple_field_result_not_replaced(self):
        p=xml('<w:r><w:t>Caption </w:t></w:r><w:fldSimple w:instr="DATE"><w:r><w:t>2026</w:t></w:r></w:fldSimple>')
        with self.assertRaises(OpError): replace(p,'2026','2030')

    def test_complex_field_result_not_replaced(self):
        p=xml('<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>PAGE</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>123</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>')
        with self.assertRaises(OpError): replace(p,'123','456')

    def test_field_across_paragraphs_not_replaced(self):
        d=Document(); p=d.add_paragraph(); r=p.add_run()._r; b=OxmlElement('w:fldChar');b.set(q('fldCharType'),'begin');r.append(b)
        d.add_paragraph('Generated result')
        r=d.add_paragraph().add_run()._r; e=OxmlElement('w:fldChar');e.set(q('fldCharType'),'end');r.append(e)
        with self.assertRaises(OpError): replace(trees(d)['word/document.xml'],'Generated result','wrong')

    def test_field_not_corrupted_by_set_text(self):
        p=xml('<w:r><w:t>Old </w:t></w:r><w:fldSimple w:instr="PAGE"><w:r><w:t>77</w:t></w:r></w:fldSimple>')
        field=etree.tostring(p.find(q('fldSimple')))
        de.set_para_text(p,'New',Editor())
        self.assertEqual(etree.tostring(p.find(q('fldSimple'))),field)
        self.assertIn('New',visible(p))

    def test_extra_cell_paragraph_drawing_is_preserved(self):
        d=Document(); t=d.add_table(rows=1,cols=1); c=t.cell(0,0);c.text='Old'
        p=c.add_paragraph(); p.add_run()._r.append(OxmlElement('w:drawing'))
        de.set_cell(c._tc,'New',Editor())
        self.assertEqual(len(list(c._tc.iter(q('drawing')))),1)

    def test_nested_table_survives_cell_text_change(self):
        d=Document(); cell=d.add_table(rows=1,cols=1).cell(0,0);cell.text='Outer'
        inner=cell.add_table(rows=1,cols=1);inner.cell(0,0).text='Nested'
        before=etree.tostring(inner._tbl)
        de.set_cell(cell._tc,'Changed outer',Editor())
        self.assertEqual(etree.tostring(inner._tbl),before)

    def test_tracked_edit_existing_insert_refuses_nested_revision(self):
        p=xml('<w:ins w:id="1" w:author="Original"><w:r><w:t>old</w:t></w:r></w:ins>')
        with self.assertRaises(OpError): replace(p,'old','new',Editor('Second'))

    def test_vml_row_clone_refuses_duplicate_objects(self):
        d=Document(); t=d.add_table(rows=1,cols=1);t.cell(0,0).text='text'
        t.cell(0,0).paragraphs[0].add_run()._r.append(OxmlElement('w:pict'))
        with self.assertRaises(OpError): de.add_row(t._tbl,0,['new'],Editor())

    def test_row_clone_refuses_duplicate_footnote(self):
        d=Document();t=d.add_table(rows=1,cols=1);t.cell(0,0).text='text'
        note=OxmlElement('w:footnoteReference');note.set(q('id'),'1');t.cell(0,0).paragraphs[0].add_run()._r.append(note)
        with self.assertRaises(OpError): de.add_row(t._tbl,0,['new'],Editor())

    def test_insert_into_vertical_merge_refuses_breaking_chain(self):
        d=Document();t=d.add_table(rows=3,cols=2);t.cell(0,0).merge(t.cell(2,0)).text='Merged'
        with self.assertRaises(OpError): de.add_row(t._tbl,0,['Inserted','x'],Editor())

    def test_delete_merge_start_preserves_remaining_merge_text(self):
        d=Document();t=d.add_table(rows=3,cols=2);t.cell(0,0).merge(t.cell(2,0)).text='Merged text'
        de.del_row(t._tbl,0,Editor())
        self.assertEqual(t.cell(0,0).text,'Merged text')
        marker=t._tbl.findall(q('tr'))[0].findall(q('tc'))[0].find(f'{q("tcPr")}/{q("vMerge")}')
        self.assertIsNotNone(marker); self.assertEqual(marker.get(q('val')),'restart')

    def test_resize_image_inside_cell_refuses_cell_overflow(self):
        image=OUT/'marker.png';pix=fitz.Pixmap(fitz.csRGB,fitz.IRect(0,0,100,50),False);pix.clear_with(160);pix.save(image)
        d=Document(); t=d.add_table(rows=1,cols=2);t.autofit=False;t.columns[0].width=Cm(2);t.cell(0,0).width=Cm(2)
        t.cell(0,0).paragraphs[0].add_run().add_picture(str(image),width=Cm(1))
        with self.assertRaises(OpError): de.resize_image(trees(d)['word/document.xml'],{'image':0,'width_cm':5},Editor())

    def test_delete_section_boundary_refuses_silent_reformat(self):
        d=Document();d.add_paragraph('Before');d.add_section();d.add_paragraph('After')
        tree=trees(d);body=tree['word/document.xml'].find(q('body'))
        target=next(i for i,b in enumerate(de.blocks(body)) if b.find(f'{q("pPr")}/{q("sectPr")}') is not None)
        with self.assertRaises(OpError): de.run_ops([{'op':'delete','block':f'#{target}'}],tree,Editor(),{},set())

    def test_atomic_batch_after_twenty_successful_ops(self):
        d=Document();d.add_paragraph('anchor');src=OUT/'batch.docx';dst=OUT/'batch.out.docx';d.save(src);dst.write_bytes(b'previous')
        ops=[{'op':'insert','after':'anchor','text':f'line {i}'} for i in range(20)]+[{'op':'replace','find':'does not exist','replace':'x'}]
        plan=OUT/'batch.json';plan.write_text(json.dumps(ops),encoding='utf-8')
        result=cli('docx_edit.py',src,dst,'--ops',plan)
        self.assertNotEqual(result.returncode,0);self.assertEqual(dst.read_bytes(),b'previous')
        self.assertEqual(Document(src).paragraphs[0].text,'anchor')

    def test_qa_rotated_page_does_not_claim_valid_text_is_outside(self):
        d=fitz.open();p=d.new_page(width=595,height=842)
        p.insert_font(fontname='embedded',fontbuffer=fitz.Font('helv').buffer)
        p.insert_text((50,750),'Valid bottom caption',fontname='embedded',fontsize=12);p.set_rotation(90)
        src=OUT/'rotated-valid.pdf';d.save(src);d.close();rep=Report();check_pdf(src,rep,set(),None)
        self.assertFalse([e for e in rep.errors if 'outside page' in e],rep.errors)

    def test_qa_detects_vertical_vector_overflow(self):
        d=fitz.open();p=d.new_page(width=300,height=400);p.draw_rect(fitz.Rect(50,380,200,450),color=(0,0,0))
        src=OUT/'vector-overflow.pdf';d.save(src);d.close();rep=Report();check_pdf(src,rep,set(),None)
        self.assertTrue(any('drawing' in e for e in rep.errors),rep.errors)

    def test_pdf_insertion_refuses_annotation_overlap(self):
        d=fitz.open();p=d.new_page();p.add_rect_annot(fitz.Rect(50,50,300,100)).update()
        with self.assertRaises(ValueError): apply_ops(d,[{'op':'insert_text','rect':[50,50,300,100],'text':'Over annotation'}],OUT)
        d.close()

    def test_pdf_insertion_refuses_form_field_overlap(self):
        d=fitz.open();p=d.new_page();w=fitz.Widget();w.field_name='Client';w.field_type=fitz.PDF_WIDGET_TYPE_TEXT;w.rect=fitz.Rect(50,50,300,100);p.add_widget(w)
        with self.assertRaises(ValueError): apply_ops(d,[{'op':'insert_text','rect':[50,50,300,100],'text':'Over form'}],OUT)
        d.close()

    def test_pdf_replacement_preserves_links_and_vector_art(self):
        d=fitz.open();p=d.new_page();font=fitz.Font('helv');p.insert_font(fontname='embedded',fontbuffer=font.buffer)
        p.insert_text((50,80),'Alpha',fontname='embedded',fontsize=12)
        p.draw_line((50,84),(120,84),color=(0,0,1));p.insert_link({'kind':fitz.LINK_URI,'from':fitz.Rect(48,65,120,85),'uri':'https://example.com/'})
        src,dst=OUT/'linked.pdf',OUT/'linked.edited.pdf';d.save(src);d.close()
        r=cli('pdf_edit.py',src,dst,'--replace','Alpha','Alphx')
        self.assertEqual(r.returncode,0,r.stdout+r.stderr)
        with fitz.open(dst) as chk:
            self.assertEqual(len(chk[0].get_links()),1,'redaction must not silently remove a hyperlink')
            self.assertEqual(chk[0].get_links()[0]['uri'],'https://example.com/')
            self.assertEqual(len(chk[0].get_drawings()),1)

    def test_all_pages_of_long_pdf_have_contact_sheets(self):
        d=fitz.open()
        for i in range(31): d.new_page(width=100,height=140).insert_text((10,30),f'PAGE {i+1}',fontsize=7)
        src=OUT/'long.pdf';d.save(src);d.close();folder=OUT/'long-preview';render(src,folder,dpi=24)
        self.assertEqual(len(list(folder.glob('page-*.png'))),31)
        self.assertTrue((folder/'sheet-002.png').exists(),'pages after 24 must not disappear from preview')

    def test_fractional_or_boolean_row_index_refused(self):
        d=Document();t=d.add_table(rows=2,cols=2)
        for value in (.9, True, None, '0.5'):
            with self.subTest(value=value):
                with self.assertRaises(OpError): de.get_row(t._tbl,value)

    def test_merged_continuation_cell_edits_visible_restart(self):
        d=Document();t=d.add_table(rows=3,cols=2);t.cell(0,0).merge(t.cell(2,0)).text='Merged'
        ts=trees(d)
        de.run_ops([{'op':'cell','table':0,'row':2,'col':0,'text':'Changed merged'}],ts,Editor(),{},set())
        tbl=next(ts['word/document.xml'].iter(q('tbl')))
        self.assertEqual(de.cell_text(tbl.findall(q('tr'))[0].findall(q('tc'))[0]),'Changed merged')

    def test_delete_multiple_vertical_starts_keeps_both(self):
        d=Document();t=d.add_table(rows=4,cols=3)
        t.cell(0,0).merge(t.cell(3,0)).text='First group'
        t.cell(0,1).merge(t.cell(2,1)).text='Second group'
        de.del_row(t._tbl,0,Editor())
        self.assertEqual(t.cell(0,0).text,'First group')
        self.assertEqual(t.cell(0,1).text,'Second group')
        self.assertEqual(len(t.rows),3)

    def test_hyperlink_docx_replace_keeps_relationship(self):
        p=xml('<w:r><w:t>See </w:t></w:r><w:hyperlink xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId42"><w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>customer portal</w:t></w:r></w:hyperlink>')
        replace(p,'customer portal','partner portal')
        link=p.find(q('hyperlink'))
        self.assertIsNotNone(link)
        self.assertEqual(link.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'),'rId42')
        self.assertEqual(visible(p),'See partner portal')

    def test_repeated_replacements_in_headers_and_footnotes(self):
        d=Document();d.add_paragraph('Number 30');d.sections[0].header.paragraphs[0].text='Number 30'
        ts=trees(d);ts['word/header1.xml']=etree.fromstring(d.sections[0].header._element.xml.encode())
        ts['word/footnotes.xml']=etree.fromstring(f'<w:footnotes xmlns:w="{de.W}"><w:footnote w:id="1"><w:p><w:r><w:t>Number 30</w:t></w:r></w:p></w:footnote></w:footnotes>')
        changed=set()
        with contextlib.redirect_stdout(io.StringIO()):de.do_replace({'find':'Number 30','replace':'Number 45','count':3},ts,Editor(),changed)
        self.assertEqual(len(changed),3)
        for tree in ts.values():self.assertIn('Number 45',visible(tree))

    def test_new_table_refuses_ragged_rows_without_mutating_document(self):
        from docx_kit import add_table,new_document
        d=new_document()
        with self.assertRaises(ValueError): add_table(d,[['H1','H2'],['A','B','LOST']],caption='caption')
        self.assertEqual(len(d.tables),0)

    def test_new_table_refuses_invalid_column_weights(self):
        from docx_kit import add_table,new_document
        for widths in ([1], [1,-1], [0,1], [float('nan'),1]):
            with self.subTest(widths=widths):
                with self.assertRaises(ValueError): add_table(new_document(),[['A','B'],['1','2']],widths=widths)

    def test_polish_merged_cell_width_equals_sum_of_columns(self):
        from docx_kit import polish_table
        d=Document();t=d.add_table(rows=2,cols=3);t.cell(0,0).merge(t.cell(0,1)).text='Merged header'
        polish_table(t,9000,widths=[1,1,1])
        self.assertEqual(t.cell(0,0).width.twips,6000)

    def test_insert_table_uses_anchor_section_not_final_section(self):
        from docx.enum.section import WD_ORIENT
        d=Document();d.add_paragraph('Portrait anchor');first=d.sections[0];first.page_width=Cm(21);first.left_margin=Cm(2);first.right_margin=Cm(2)
        second=d.add_section();second.orientation=WD_ORIENT.LANDSCAPE;second.page_width=Cm(29.7);second.page_height=Cm(21)
        ts=trees(d)
        de.run_ops([{'op':'insert_table','after':'Portrait anchor','rows':[['A','B'],['1','2']]}],ts,Editor(),{},set())
        table=next(ts['word/document.xml'].iter(q('tbl')))
        width=int(table.find(f'{q("tblPr")}/{q("tblW")}').get(q('w')))
        self.assertAlmostEqual(width,Cm(17).twips,delta=2)

    def test_complex_field_is_barrier_to_cross_run_replace(self):
        p=xml('<w:r><w:t>left</w:t></w:r><w:fldSimple w:instr="PAGE"><w:r><w:t>X</w:t></w:r></w:fldSimple><w:r><w:t>right</w:t></w:r>')
        with self.assertRaises(OpError): replace(p,'leftXright','changed')


    def test_wrapped_table_header_and_cells_are_not_orphaned(self):
        from qa import check_table_splits
        d=fitz.open()
        for n in range(2):
            p=d.new_page()
            p.insert_text((50,70),'Identifier');p.insert_text((180,70),'Description')
            p.insert_text((50,82),'continued')
            for i in range(4):
                y=110+i*35
                p.insert_text((50,y),str(i));p.insert_text((180,y),'Row body')
                p.insert_text((180,y+12),'wrapped description')
        rep=Report();check_table_splits(d,rep)
        self.assertFalse(rep.errors,rep.errors)
        d.close()

    def test_real_table_orphan_still_detected_with_wrapped_header(self):
        from qa import check_table_splits
        d=fitz.open()
        for n in range(2):
            p=d.new_page();y=690 if n==0 else 70
            p.insert_text((50,y),'Identifier');p.insert_text((180,y),'Description')
            p.insert_text((50,y+12),'continued')
            p.insert_text((50,y+35),'1');p.insert_text((180,y+35),'Row body')
        rep=Report();check_table_splits(d,rep)
        self.assertTrue(any('table starts' in e for e in rep.errors),rep.errors)
        d.close()

    def test_pdf_duplicate_preserves_external_and_internal_links(self):
        d=fitz.open();d.new_page();d.new_page()
        d[0].insert_link({'kind':fitz.LINK_URI,'from':fitz.Rect(20,20,80,40),'uri':'https://example.com'})
        d[0].insert_link({'kind':fitz.LINK_GOTO,'from':fitz.Rect(20,60,80,80),'page':1,'to':fitz.Point(0,0)})
        data=d.tobytes();d.close();d=fitz.open(stream=data,filetype='pdf')
        apply_ops(d,[{'op':'duplicate_page','page':1}],OUT)
        data=d.tobytes();d.close();d=fitz.open(stream=data,filetype='pdf')
        for p in (d[0],d[1]):
            links=p.get_links();self.assertEqual(len(links),2)
            self.assertTrue(any(l.get('uri')=='https://example.com' for l in links))
            self.assertTrue(any(l.get('page')==2 for l in links))
        d.close()


    def test_same_batch_tracked_insert_replace_move_has_no_nested_revisions(self):
        d=Document();d.add_paragraph('Start anchor');d.add_paragraph('End anchor')
        ts=trees(d);ed=Editor('Auditor')
        de.run_ops([{'op':'insert','after':'Start anchor','text':'New alpha'},
                    {'op':'replace','find':'New alpha','replace':'New beta'},
                    {'op':'move','block':'New beta','after':'End anchor'}],ts,ed,{},set())
        root=ts['word/document.xml']
        self.assertFalse(root.xpath('//w:ins//w:ins | //w:ins//w:del',namespaces={'w':de.W}))
        self.assertEqual(visible(root).count('New beta'),1)
        self.assertTrue(visible(root).endswith('New beta'))

    def test_field_cache_does_not_leak_between_batches_or_after_error(self):
        d=Document();d.add_paragraph('Anchor')
        ts=trees(d)
        with self.assertRaises(OpError):
            de.run_ops([{'op':'replace','find':'Missing','replace':'Fail'}],ts,Editor(),{},set())
        self.assertIsNone(de._FIELD_CACHE)
        p=xml('<w:fldSimple w:instr="PAGE"><w:r><w:t>77</w:t></w:r></w:fldSimple>')
        with self.assertRaises(OpError):replace(p,'77','99')


    def test_contact_sheet_keeps_bottom_text_at_every_rotation(self):
        from render import sheet
        from PIL import Image
        for angle in (0,90,180,270):
            with self.subTest(angle=angle):
                d=fitz.open();p=d.new_page(width=595,height=842)
                p.insert_text((50,750),'Visible bottom caption',fontsize=20);p.set_rotation(angle)
                path=OUT/f'rotation-{angle}.png';sheet(d,path);d.close()
                im=Image.open(path).convert('RGB');im=im.crop((9,9,im.width-9,im.height-9))
                self.assertGreater(sum(1 for pixel in im.get_flattened_data() if max(pixel)<100),30)


    def test_malformed_input_errors_are_clean_and_preserve_output(self):
        src=OUT/'corrupt.docx';src.write_bytes(b'not a ZIP document')
        dst=OUT/'corrupt.out.docx';dst.write_bytes(b'previous result')
        result=cli('docx_edit.py',src,dst,'--replace','Old','New')
        self.assertNotEqual(result.returncode,0)
        self.assertIn('ERROR',result.stdout);self.assertNotIn('Traceback',result.stdout+result.stderr)
        self.assertEqual(dst.read_bytes(),b'previous result')


if __name__=='__main__': unittest.main(verbosity=2)
