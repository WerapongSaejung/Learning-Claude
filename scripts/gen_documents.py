# -*- coding: utf-8 -*-
"""สร้างเอกสารข้อมูลสมมติ (docx/pdf) สำหรับ workshop "ใช้ AI ทำงานจริงด้วย Claude"

ข้อมูลทั้งหมดแต่งขึ้นเพื่อการฝึกอบรม ไม่ใช่เอกสารจริงของบริษัท

รันแล้วได้ 4 ไฟล์ใน dataset/
  TP-04-exec-meeting-minutes-Q2.docx  รายงานประชุมผู้บริหารไตรมาส 2 (เอกสารยาว)
  TP-05-regulatory-notice.pdf         ประกาศหลักเกณฑ์ฉบับสมมติ
  TP-06-export-customer-emails.docx   อีเมลลูกค้าต่างประเทศ 8 ฉบับ
  TP-07-sop-complaint-handling.pdf    SOP การจัดการข้อร้องเรียน

ต้องมี LibreOffice (soffice) สำหรับแปลง docx เป็น pdf
"""
import os
import subprocess
import tempfile
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, 'dataset')

# Tahoma มีอยู่บน Windows/Mac และรองรับไทยครบ — ใช้กับไฟล์ที่ส่งมอบเป็น .docx
# Loma ติดตั้งอยู่บนเครื่องที่ใช้แปลง PDF — ใช้กับไฟล์ที่จะแปลงเป็น .pdf เพื่อให้ฝังฟอนต์ถูกต้อง
FONT_DOCX = 'Tahoma'
FONT_PDF = 'Loma'

DISCLAIMER = ('เอกสารสมมติเพื่อการฝึกอบรม — ไม่ใช่เอกสารจริงของบริษัท '
              'ตัวเลข ชื่อบุคคล ชื่อลูกค้า และเหตุการณ์ทั้งหมดแต่งขึ้น ห้ามนำไปใช้อ้างอิงในงานจริง')


def new_doc(font):
    doc = Document()
    st = doc.styles['Normal']
    st.font.name = font
    st.font.size = Pt(10.5)
    st.paragraph_format.space_after = Pt(4)
    st.paragraph_format.line_spacing = 1.15
    for s in doc.sections:
        s.top_margin = s.bottom_margin = Cm(2.0)
        s.left_margin = s.right_margin = Cm(2.2)
    return doc


def banner(doc, font):
    p = doc.add_paragraph()
    r = p.add_run(DISCLAIMER)
    r.font.size = Pt(8.5)
    r.font.name = font
    r.font.color.rgb = RGBColor(0xB0, 0x30, 0x30)
    r.italic = True
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER


def title(doc, text, font, size=16):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.bold = True
    r.font.size = Pt(size)
    r.font.name = font
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    return p


def heading(doc, text, font, size=12):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    r = p.add_run(text)
    r.bold = True
    r.font.size = Pt(size)
    r.font.name = font
    return p


def body(doc, text, font, indent=0.0):
    p = doc.add_paragraph()
    if indent:
        p.paragraph_format.left_indent = Cm(indent)
    r = p.add_run(text)
    r.font.name = font
    r.font.size = Pt(10.5)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    return p


def bullets(doc, items, font, indent=0.8):
    for it in items:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(indent)
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run('•  ' + it)
        r.font.name = font
        r.font.size = Pt(10.5)


def table(doc, headers, rows, font, widths=None):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = 'Table Grid'
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(headers):
        cell = t.rows[0].cells[i]
        cell.text = ''
        r = cell.paragraphs[0].add_run(h)
        r.bold = True
        r.font.name = font
        r.font.size = Pt(9.5)
    for row in rows:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = ''
            r = cells[i].paragraphs[0].add_run(str(v))
            r.font.name = font
            r.font.size = Pt(9.5)
    if widths:
        for row in t.rows:
            for i, w in enumerate(widths):
                row.cells[i].width = Cm(w)
    return t


def to_pdf(docx_path, pdf_path):
    outdir = tempfile.mkdtemp()
    subprocess.run(['soffice', '--headless', '--norestore', '--convert-to', 'pdf',
                    '--outdir', outdir, docx_path],
                   check=True, capture_output=True, timeout=420)
    produced = os.path.join(outdir, os.path.splitext(os.path.basename(docx_path))[0] + '.pdf')
    if not os.path.exists(produced):
        raise RuntimeError('แปลง PDF ไม่สำเร็จ: %s' % docx_path)
    os.replace(produced, pdf_path)


# ============================================ TP-04 รายงานประชุมผู้บริหาร Q2
# ใช้ "ตำแหน่ง" แทนชื่อบุคคล เพื่อไม่ให้เอกสารสมมติไปพัวพันกับพนักงานจริงของบริษัท

MINUTES = [
    ('วาระที่ 1  รับรองรายงานการประชุมครั้งที่ 1/2568', [
        ('p', 'ประธานที่ประชุมเสนอให้ที่ประชุมพิจารณารับรองรายงานการประชุมคณะผู้บริหาร ครั้งที่ 1/2568 '
              'ซึ่งประชุมเมื่อวันที่ 18 มกราคม 2568 โดยฝ่ายเลขานุการได้จัดส่งร่างรายงานให้ผู้เข้าประชุมทุกท่าน '
              'ล่วงหน้าไม่น้อยกว่าเจ็ดวันทำการ'),
        ('p', 'ผู้อำนวยการฝ่ายบัญชีและการเงินขอแก้ไขข้อความในวาระที่ 6 หน้า 9 จากเดิม '
              '"งบลงทุนปรับปรุงระบบน้ำบริสุทธิ์วงเงิน 12.0 ล้านบาท" เป็น "วงเงินไม่เกิน 12.8 ล้านบาท '
              'ตามที่ได้รับอนุมัติในหลักการ" เนื่องจากตัวเลขเดิมคลาดเคลื่อนจากเอกสารประกอบ'),
        ('p', 'มติ: ที่ประชุมรับรองรายงานการประชุมครั้งที่ 1/2568 พร้อมการแก้ไขตามที่เสนอ'),
    ]),
    ('วาระที่ 2  รายงานผลการดำเนินงานฝ่ายขายและการตลาด', [
        ('p', 'ผู้อำนวยการฝ่ายขายและการตลาดรายงานว่ายอดขายรวมไตรมาสที่ 2 ปี 2568 อยู่ที่ประมาณ 52.4 ล้านบาท '
              'เพิ่มขึ้นจากไตรมาสที่ 1 ซึ่งอยู่ที่ประมาณ 49.0 ล้านบาท คิดเป็นการเติบโตร้อยละ 6.9 '
              'โดยการเติบโตมาจากช่องทางส่งออกเป็นหลัก และจากคำสั่งซื้อของโรงพยาบาลเอกชนที่ฟื้นตัวในเดือนมิถุนายน'),
        ('p', 'เมื่อพิจารณาแยกตามช่องทางจำหน่าย ช่องทางโรงพยาบาลรัฐยังคงเป็นสัดส่วนใหญ่ที่สุดประมาณร้อยละ 41 '
              'ของยอดขายรวม รองลงมาคือช่องทางร้านขายยาประมาณร้อยละ 22 ช่องทางส่งออกประมาณร้อยละ 19 '
              'และช่องทางโรงพยาบาลเอกชนประมาณร้อยละ 18 ทั้งนี้สัดส่วนของช่องทางส่งออกเพิ่มขึ้นต่อเนื่องทุกเดือน '
              'ตั้งแต่เดือนเมษายนเป็นต้นมา'),
        ('h', 'ประเด็นที่ต้องเฝ้าระวัง'),
        ('b', [
            'ผลิตภัณฑ์ Ceftriaxone T P ยังเป็นสินค้าที่ทำรายได้สูงสุด โดยพึ่งพาช่องทางโรงพยาบาลรัฐมากกว่าร้อยละ 40 '
            'ของยอดขายผลิตภัณฑ์นี้ ความเสี่ยงจากการประมูลจึงกระจุกตัวสูง',
            'ราคาขายเฉลี่ยต่อหน่วยของช่องทางส่งออกต่ำกว่าราคาอ้างอิงประมาณร้อยละ 22 '
            'ยอดขายที่โตจากช่องทางนี้จึงไม่ได้ทำให้กำไรขั้นต้นโตในอัตราเดียวกัน',
            'ปริมาณการสั่งซื้อของตัวแทนจำหน่ายในประเทศเมียนมาและกัมพูชาเพิ่มขึ้นเร็วกว่าที่ประมาณการไว้ '
            'ฝ่ายวางแผนการผลิตแจ้งว่ากำลังการบรรจุสายที่ 2 เริ่มตึงตัวในเดือนมิถุนายน',
        ]),
        ('p', 'ที่ประชุมตั้งข้อสังเกตว่าการรายงานยอดขายในรูปยอดรวมรายเดือนอย่างเดียว '
              'ทำให้มองไม่เห็นว่าการเติบโตของช่องทางหนึ่งกำลังกลบการหดตัวของอีกช่องทางหนึ่ง '
              'จึงขอให้ฝ่ายขายปรับรูปแบบรายงานประจำเดือนให้แสดงยอดแยกตามช่องทางและตามผลิตภัณฑ์หลักควบคู่กันไป '
              'เริ่มตั้งแต่รายงานเดือนกรกฎาคม 2568'),
        ('p', 'มติ: รับทราบรายงาน และให้ฝ่ายขายและการตลาดปรับรูปแบบรายงานประจำเดือนตามข้อสังเกตของที่ประชุม'),
    ]),
    ('วาระที่ 3  รายงานตลาดต่างประเทศและการส่งออก', [
        ('p', 'ผู้จัดการฝ่ายธุรกิจต่างประเทศรายงานว่าปัจจุบันบริษัทมีการจำหน่ายในตลาดต่างประเทศรวม 19 ประเทศ '
              'โดยตลาดที่มีอัตราการเติบโตสูงสุดในไตรมาสนี้คือเมียนมา กัมพูชา และฟิลิปปินส์ ตามลำดับ '
              'ส่วนตลาดที่ยอดขายทรงตัวคือเวียดนามและศรีลังกา'),
        ('h', 'สถานะการขึ้นทะเบียนในตลาดต่างประเทศ'),
        ('t', (['ประเทศ', 'ผลิตภัณฑ์', 'สถานะการขึ้นทะเบียน', 'กำหนดเดิม', 'สถานะล่าสุด'],
               [['เมียนมา', 'Ceftriaxone T P', 'ยื่นเอกสารเพิ่มเติมรอบที่ 2', 'มี.ค. 2568', 'ล่าช้า 3 เดือน'],
                ['กัมพูชา', 'Clatacef', 'รอผลพิจารณา', 'พ.ค. 2568', 'ตามกำหนด'],
                ['ฟิลิปปินส์', 'Paramol T P', 'รอเอกสาร CoA ย้อนหลัง 3 รุ่นผลิต', 'เม.ย. 2568', 'ล่าช้า 2 เดือน'],
                ['เวียดนาม', 'Tramadol T P', 'เตรียมยื่นใหม่', 'ก.ค. 2568', 'ตามกำหนด'],
                ['ศรีลังกา', 'Furosemide Injection T P', 'ต่ออายุทะเบียน', 'ส.ค. 2568', 'ตามกำหนด'],
                ['ไนจีเรีย', 'Vitamin C Injection T P', 'ยื่นครั้งแรก', 'ก.ย. 2568', 'รอเอกสารโรงงาน'],
                ],
               [2.6, 5.4, 5.4, 2.6, 3.0])),
        ('p', 'ผู้จัดการฝ่ายธุรกิจต่างประเทศรายงานเพิ่มเติมว่าความล่าช้าของการขึ้นทะเบียนในเมียนมาและฟิลิปปินส์ '
              'ทำให้ตัวแทนจำหน่ายทั้งสองรายส่งหนังสือสอบถามเข้ามาหลายฉบับในรอบสองเดือนที่ผ่านมา '
              'และเริ่มมีการสอบถามถึงความแน่นอนของแผนการส่งมอบในครึ่งปีหลัง'),
        ('p', 'สาเหตุหลักของความล่าช้าเกิดจากปริมาณงานจัดทำเอกสารของฝ่ายทะเบียนยาที่เพิ่มขึ้นพร้อมกัน '
              'ทั้งงานขึ้นทะเบียนใหม่ งานต่ออายุ และงานตอบข้อซักถามของหน่วยงานกำกับดูแลในหลายประเทศ '
              'ขณะที่จำนวนบุคลากรยังเท่าเดิม'),
        ('p', 'มติ: ให้ฝ่ายทะเบียนยาจัดทำแผนงานเอกสารรายตลาดพร้อมระบุลำดับความสำคัญ '
              'และให้ฝ่ายธุรกิจต่างประเทศจัดทำหนังสือชี้แจงความคืบหน้าส่งตัวแทนจำหน่ายทุกรายภายในเดือนกรกฎาคม 2568'),
    ]),
    ('วาระที่ 4  รายงานฝ่ายผลิตและวิศวกรรม', [
        ('p', 'ผู้อำนวยการฝ่ายผลิตรายงานผลการผลิตไตรมาสที่ 2 โดยภาพรวมเป็นไปตามแผน '
              'อัตราการใช้กำลังการผลิตของสายผลิตยาฉีดปราศจากเชื้ออยู่ที่ประมาณร้อยละ 78 '
              'และสายผลิตยาเม็ดอยู่ที่ประมาณร้อยละ 64'),
        ('h', 'ประเด็นด้านกำลังการผลิต'),
        ('b', [
            'สายบรรจุที่ 2 มีชั่วโมงเดินเครื่องสูงกว่าแผนร้อยละ 11 ในเดือนมิถุนายน จากคำสั่งซื้อส่งออกที่เพิ่มขึ้น',
            'เครื่องตรวจสอบอนุภาคด้วยสายตาของสายบรรจุที่ 2 หยุดเดินเครื่องเพื่อซ่อมบำรุงนอกแผนรวม 3 ครั้ง ในเดือนพฤษภาคม',
            'ฝ่ายวิศวกรรมเสนอให้เร่งแผนเปลี่ยนชุดหัวบรรจุของสายที่ 2 จากไตรมาสที่ 4 มาเป็นไตรมาสที่ 3',
        ]),
        ('p', 'ที่ประชุมสอบถามถึงความเชื่อมโยงระหว่างการหยุดเดินเครื่องของเครื่องตรวจสอบอนุภาคในเดือนพฤษภาคม '
              'กับข้อร้องเรียนเรื่องอนุภาคแขวนลอยที่ฝ่ายประกันคุณภาพจะรายงานในวาระถัดไป '
              'ผู้อำนวยการฝ่ายผลิตชี้แจงว่าอยู่ระหว่างการสอบสวนร่วมกับฝ่ายประกันคุณภาพ ยังไม่สรุปสาเหตุราก'),
        ('p', 'มติ: อนุมัติในหลักการให้เร่งแผนเปลี่ยนชุดหัวบรรจุสายที่ 2 มาเป็นไตรมาสที่ 3 '
              'โดยให้ฝ่ายวิศวกรรมเสนอรายละเอียดวงเงินและผลกระทบต่อแผนการผลิตในการประชุมครั้งถัดไป'),
    ]),
    ('วาระที่ 5  รายงานฝ่ายประกันคุณภาพและควบคุมคุณภาพ', [
        ('p', 'ผู้อำนวยการฝ่ายประกันคุณภาพรายงานว่าในไตรมาสที่ 2 บริษัทได้รับข้อร้องเรียนผลิตภัณฑ์รวม 21 เรื่อง '
              'เพิ่มขึ้นจากไตรมาสที่ 1 ซึ่งมี 14 เรื่อง โดยการเพิ่มขึ้นเกือบทั้งหมดมาจากกลุ่มข้อร้องเรียนเรื่องเดียว'),
        ('h', 'ข้อร้องเรียนเรื่องอนุภาคแขวนลอยในผลิตภัณฑ์ Gentamicin T P'),
        ('p', 'ตั้งแต่วันที่ 11 มิถุนายน 2568 บริษัทได้รับข้อร้องเรียนเรื่องพบอนุภาคแขวนลอยในหลอดบรรจุ '
              'ของผลิตภัณฑ์ Gentamicin T P จากผู้แจ้งหลายรายต่อเนื่องกัน โดยข้อร้องเรียนกระจุกตัวใน '
              'สามรุ่นผลิตติดกันคือ GTM-2568-0416, GTM-2568-0417 และ GTM-2568-0418 '
              'ซึ่งผลิตในช่วงปลายเดือนเมษายนถึงต้นเดือนพฤษภาคม 2568 บนสายบรรจุที่ 2'),
        ('p', 'ฝ่ายควบคุมคุณภาพได้สุ่มตัวอย่างจากตัวอย่างเก็บอ้างอิงของทั้งสามรุ่นผลิตมาตรวจซ้ำ '
              'ผลการตรวจสอบพบอนุภาคเกินเกณฑ์ที่กำหนดในสองรุ่นผลิต และอยู่ในเกณฑ์แต่ค่าใกล้ขอบบนในอีกหนึ่งรุ่นผลิต '
              'ฝ่ายประกันคุณภาพจึงเปิดการสอบสวนและออกเลขที่ CAPA-2568-014'),
        ('h', 'มาตรการที่ดำเนินการแล้ว'),
        ('b', [
            'ระงับการจำหน่ายผลิตภัณฑ์ Gentamicin T P ทั้งสามรุ่นผลิตทันทีเมื่อวันที่ 18 กรกฎาคม 2568',
            'แจ้งผู้แจ้งข้อร้องเรียนทุกรายเป็นหนังสือ พร้อมดำเนินการรับคืนและเปลี่ยนสินค้า',
            'ตรวจสอบย้อนกลับรุ่นผลิตอื่นบนสายบรรจุที่ 2 ในช่วงเวลาเดียวกัน ไม่พบความผิดปกติเพิ่มเติม',
            'หยุดจำหน่าย Gentamicin T P ชั่วคราวตั้งแต่เดือนสิงหาคม 2568 จนกว่าการสอบสวนสาเหตุรากจะได้ข้อสรุป '
            'และมาตรการแก้ไขได้รับการยืนยันประสิทธิผล',
            'รายงานหน่วยงานกำกับดูแลตามกรอบเวลาที่กำหนด',
        ]),
        ('p', 'ที่ประชุมสอบถามถึงผลกระทบต่อยอดขาย ผู้อำนวยการฝ่ายขายและการตลาดประมาณการว่าการหยุดจำหน่าย '
              'Gentamicin T P จะทำให้ยอดขายหายไปประมาณ 0.7 ถึง 0.8 ล้านบาทต่อเดือน '
              'และหากหยุดจำหน่ายเต็มเดือนจะกระทบยอดขายรวมของบริษัทประมาณร้อยละ 4 ถึง 5 ในเดือนนั้น'),
        ('p', 'ที่ประชุมยังตั้งข้อสังเกตว่ากระบวนการจัดการข้อร้องเรียนตามขั้นตอนปฏิบัติฉบับปัจจุบัน '
              'ใช้เวลาจากวันรับแจ้งจนถึงวันเปิดการสอบสวนนานเกินไปในกรณีที่มีข้อร้องเรียนซ้ำในผลิตภัณฑ์เดียวกัน '
              'จึงควรมีกลไกยกระดับเรื่องอัตโนมัติเมื่อพบข้อร้องเรียนลักษณะเดียวกันเกินจำนวนที่กำหนด'),
        ('p', 'มติ: รับทราบรายงานและมาตรการที่ดำเนินการ และให้ฝ่ายประกันคุณภาพทบทวนขั้นตอนปฏิบัติ '
              'เรื่องการจัดการข้อร้องเรียนผลิตภัณฑ์ โดยเพิ่มเกณฑ์การยกระดับเรื่องกรณีข้อร้องเรียนซ้ำ '
              'และเสนอฉบับแก้ไขภายในไตรมาสที่ 3 ปี 2568'),
    ]),
    ('วาระที่ 6  รายงานฝ่ายทะเบียนยาและการกำกับดูแล', [
        ('p', 'ผู้จัดการฝ่ายทะเบียนยารายงานว่าในไตรมาสที่ 2 มีงานยื่นเอกสารต่อหน่วยงานกำกับดูแลรวม 34 รายการ '
              'แบ่งเป็นงานขึ้นทะเบียนใหม่ 6 รายการ งานต่ออายุทะเบียน 11 รายการ '
              'งานแก้ไขเปลี่ยนแปลงทะเบียน 9 รายการ และงานตอบข้อซักถาม 8 รายการ'),
        ('p', 'ฝ่ายทะเบียนยาแจ้งว่าอยู่ระหว่างศึกษาร่างประกาศหลักเกณฑ์เรื่องการรายงานข้อร้องเรียน '
              'และการเรียกคืนผลิตภัณฑ์ยาฉบับใหม่ ซึ่งจะมีผลบังคับใช้ในปี 2568 '
              'โดยสาระสำคัญที่กระทบต่อบริษัทคือกรอบเวลาการรายงานที่สั้นลง '
              'และการกำหนดให้ต้องมีระบบทบทวนแนวโน้มข้อร้องเรียนเป็นลายลักษณ์อักษร'),
        ('p', 'ที่ประชุมขอให้ฝ่ายทะเบียนยาจัดทำสรุปสาระสำคัญของประกาศฉบับดังกล่าว '
              'เปรียบเทียบกับแนวปฏิบัติปัจจุบันของบริษัท ระบุช่องว่างที่ต้องปิด และประมาณการทรัพยากรที่ต้องใช้ '
              'เสนอที่ประชุมครั้งถัดไป'),
        ('p', 'มติ: รับทราบ และให้ฝ่ายทะเบียนยาดำเนินการตามที่ที่ประชุมมอบหมาย'),
    ]),
    ('วาระที่ 7  รายงานฝ่ายบัญชีและการเงิน และฝ่ายจัดซื้อ', [
        ('p', 'ผู้อำนวยการฝ่ายบัญชีและการเงินรายงานว่าอัตรากำไรขั้นต้นรวมของไตรมาสที่ 2 อยู่ที่ประมาณร้อยละ 37 '
              'ลดลงจากไตรมาสที่ 1 ซึ่งอยู่ที่ประมาณร้อยละ 39 ทั้งที่ยอดขายเติบโต '
              'สาเหตุหลักมาจากสัดส่วนของยอดขายช่องทางส่งออกที่เพิ่มขึ้น ซึ่งมีราคาต่อหน่วยต่ำกว่าช่องทางอื่น '
              'ประกอบกับต้นทุนวัตถุดิบตั้งต้นกลุ่ม cephalosporin ที่ปรับขึ้น'),
        ('h', 'ต้นทุนวัตถุดิบที่ปรับขึ้นในไตรมาส'),
        ('t', (['กลุ่มวัตถุดิบ', 'การเปลี่ยนแปลงราคา', 'ผลกระทบต่อต้นทุน', 'การรับมือ'],
               [['วัตถุดิบตั้งต้น cephalosporin', 'เพิ่มขึ้นร้อยละ 8', 'ประมาณ 0.9 ล้านบาทต่อไตรมาส',
                 'เจรจาสัญญาระยะยาว 12 เดือน'],
                ['หลอดแก้วบรรจุยาฉีด', 'เพิ่มขึ้นร้อยละ 5', 'ประมาณ 0.3 ล้านบาทต่อไตรมาส',
                 'เพิ่มผู้ขายรายที่สอง'],
                ['ค่าขนส่งทางเรือ', 'ลดลงร้อยละ 3', 'ลดลงประมาณ 0.1 ล้านบาทต่อไตรมาส', 'คงแผนเดิม'],
                ],
               [4.6, 3.6, 4.4, 4.4])),
        ('p', 'ที่ประชุมตั้งข้อสังเกตว่าการติดตามผลการดำเนินงานด้วยตัวเลขยอดขายเพียงอย่างเดียวไม่เพียงพอ '
              'เนื่องจากยอดขายที่เติบโตอาจมาพร้อมกับอัตรากำไรขั้นต้นที่ลดลง '
              'จึงควรรายงานยอดขายและกำไรขั้นต้นแยกตามช่องทางควบคู่กันทุกเดือน'),
        ('p', 'มติ: ให้ฝ่ายบัญชีและการเงินร่วมกับฝ่ายขายจัดทำรายงานกำไรขั้นต้นแยกตามช่องทางจำหน่าย '
              'เริ่มใช้กับรายงานเดือนกรกฎาคม 2568'),
    ]),
    ('วาระที่ 8  รายงานฝ่ายทรัพยากรบุคคลและเทคโนโลยีสารสนเทศ', [
        ('p', 'ผู้จัดการฝ่ายทรัพยากรบุคคลรายงานว่าอัตราการลาออกของพนักงานในไตรมาสที่ 2 อยู่ที่ร้อยละ 3.1 '
              'ต่ำกว่าไตรมาสก่อน ตำแหน่งที่ยังสรรหาไม่ได้ตามแผนคือเจ้าหน้าที่ฝ่ายทะเบียนยา 2 อัตรา '
              'และเจ้าหน้าที่ประกันคุณภาพ 1 อัตรา'),
        ('p', 'ผู้จัดการฝ่ายเทคโนโลยีสารสนเทศรายงานความคืบหน้าการปรับปรุงระบบจัดเก็บเอกสารอิเล็กทรอนิกส์ '
              'ปัจจุบันอยู่ในขั้นย้ายข้อมูลเอกสารระบบคุณภาพ คาดว่าจะแล้วเสร็จภายในไตรมาสที่ 3'),
        ('p', 'ฝ่ายเทคโนโลยีสารสนเทศรายงานเพิ่มเติมว่าพบการใช้เครื่องมือปัญญาประดิษฐ์แบบสนทนา '
              'ผ่านบัญชีส่วนตัวของพนักงานในหลายฝ่าย โดยบริษัทยังไม่มีนโยบายกำกับการใช้งาน '
              'จึงยังไม่สามารถระบุได้ว่ามีการนำข้อมูลประเภทใดออกไปใช้กับเครื่องมือดังกล่าวบ้าง'),
        ('p', 'ที่ประชุมเห็นว่าเรื่องนี้มีความสำคัญทั้งในมิติความปลอดภัยของข้อมูลและมิติการกำกับดูแล '
              'จึงควรกำหนดแนวปฏิบัติให้ชัดเจนก่อนที่การใช้งานจะกระจายไปมากกว่านี้'),
        ('p', 'มติ: ให้ฝ่ายเทคโนโลยีสารสนเทศร่วมกับฝ่ายประกันคุณภาพจัดทำร่างแนวปฏิบัติการใช้เครื่องมือ '
              'ปัญญาประดิษฐ์ในการปฏิบัติงาน โดยกำหนดประเภทข้อมูลที่ห้ามนำออกนอกองค์กร '
              'และเงื่อนไขการใช้กับเอกสารในระบบคุณภาพ เสนอที่ประชุมครั้งถัดไป'),
    ]),
    ('วาระที่ 9  เรื่องพิจารณา: การประมูลจัดซื้อยาของโรงพยาบาลรัฐรอบเดือนกรกฎาคม 2568', [
        ('p', 'ผู้อำนวยการฝ่ายขายและการตลาดรายงานว่าการประมูลจัดซื้อยาของกลุ่มโรงพยาบาลรัฐรอบเดือนกรกฎาคม 2568 '
              'ครอบคลุมรายการยาที่บริษัทเป็นผู้ส่งมอบรายหลักหลายรายการ โดยรายการที่มีมูลค่าสูงสุดคือ '
              'Ceftriaxone T P'),
        ('p', 'จากการสำรวจราคาตลาด พบว่ามีผู้ผลิตรายใหม่สองรายเสนอราคาต่ำกว่าราคาที่บริษัทเสนอในรอบก่อน '
              'ประมาณร้อยละ 10 ถึง 12 หากบริษัทคงราคาเดิม มีความเป็นไปได้สูงที่จะไม่ได้รับงานในรายการนี้ '
              'ซึ่งจะกระทบยอดขายช่องทางโรงพยาบาลรัฐอย่างมีนัยสำคัญตั้งแต่เดือนกรกฎาคมเป็นต้นไป'),
        ('h', 'ทางเลือกที่เสนอต่อที่ประชุม'),
        ('b', [
            'ทางเลือกที่ 1 ลดราคาลงร้อยละ 12 เพื่อรักษาส่วนแบ่ง ซึ่งจะทำให้อัตรากำไรขั้นต้นของรายการนี้ '
            'ลดลงต่ำกว่าเกณฑ์ขั้นต่ำที่บริษัทกำหนด',
            'ทางเลือกที่ 2 คงราคาเดิมและยอมรับความเสี่ยงที่จะเสียงาน แล้วผลักปริมาณไปยังช่องทางส่งออก '
            'และช่องทางโรงพยาบาลเอกชนแทน',
            'ทางเลือกที่ 3 เสนอราคาลดลงร้อยละ 6 พร้อมเงื่อนไขปริมาณขั้นต่ำและกำหนดส่งมอบที่ยืดหยุ่นกว่า',
        ]),
        ('p', 'ที่ประชุมอภิปรายอย่างกว้างขวาง โดยฝ่ายบัญชีและการเงินไม่เห็นด้วยกับทางเลือกที่ 1 '
              'เนื่องจากจะสร้างฐานราคาใหม่ที่ต่ำจนกลับขึ้นได้ยากในรอบถัดไป '
              'ขณะที่ฝ่ายผลิตตั้งข้อสังเกตว่าหากเลือกทางเลือกที่ 2 กำลังการบรรจุของสายที่ 2 '
              'อาจไม่เพียงพอรองรับปริมาณส่งออกที่เพิ่มขึ้นในไตรมาสที่ 3'),
        ('p', 'มติ: เห็นชอบทางเลือกที่ 3 เป็นแนวทางหลัก และมอบหมายให้ฝ่ายขายและการตลาดร่วมกับ '
              'ฝ่ายบัญชีและการเงินจัดทำโครงสร้างราคาและเงื่อนไขเสนอประธานพิจารณาก่อนวันยื่นเสนอราคา '
              'พร้อมจัดทำแผนรองรับกรณีไม่ได้รับงาน'),
    ]),
    ('วาระที่ 10  เรื่องพิจารณา: งบลงทุนปรับปรุงสายบรรจุที่ 2', [
        ('p', 'ฝ่ายวิศวกรรมเสนอขออนุมัติงบลงทุนเปลี่ยนชุดหัวบรรจุและปรับปรุงระบบตรวจสอบอนุภาค '
              'ของสายบรรจุที่ 2 วงเงินไม่เกิน 18.5 ล้านบาท โดยให้เหตุผลว่าจะช่วยลดความเสี่ยง '
              'ด้านคุณภาพที่พบในไตรมาสนี้ และเพิ่มกำลังการบรรจุได้ประมาณร้อยละ 15'),
        ('p', 'ฝ่ายบัญชีและการเงินตั้งข้อสังเกตว่าวงเงินที่เสนอสูงกว่ากรอบงบลงทุนประจำปีที่เหลืออยู่ '
              'จึงต้องพิจารณาจัดลำดับความสำคัญกับโครงการอื่นที่ได้รับอนุมัติในหลักการไว้แล้ว'),
        ('p', 'มติ: อนุมัติในหลักการวงเงินไม่เกิน 18.5 ล้านบาท '
              'โดยให้ฝ่ายวิศวกรรมและฝ่ายบัญชีและการเงินจัดทำแผนการเบิกจ่ายแยกเป็นสองระยะ '
              'และเสนอรายละเอียดพร้อมผลตอบแทนที่คาดหวังในการประชุมครั้งถัดไป'),
    ]),
    ('วาระที่ 11  เรื่องอื่น ๆ', [
        ('p', 'ประธานที่ประชุมแจ้งกำหนดการตรวจประเมินระบบคุณภาพจากหน่วยงานกำกับดูแลในไตรมาสที่ 4 ปี 2568 '
              'และขอให้ทุกฝ่ายเตรียมความพร้อมด้านเอกสารตั้งแต่ไตรมาสที่ 3'),
        ('p', 'ฝ่ายเลขานุการแจ้งกำหนดการประชุมคณะผู้บริหารครั้งที่ 3/2568 ในเดือนตุลาคม 2568'),
        ('p', 'ปิดประชุมเวลา 16:20 นาฬิกา'),
    ]),
]

ACTIONS = [
    ['1', 'ปรับรูปแบบรายงานยอดขายประจำเดือนให้แยกตามช่องทางและผลิตภัณฑ์หลัก', 'ฝ่ายขายและการตลาด', 'ก.ค. 2568'],
    ['2', 'จัดทำแผนงานเอกสารรายตลาดพร้อมลำดับความสำคัญ', 'ฝ่ายทะเบียนยา', 'ก.ค. 2568'],
    ['3', 'จัดทำหนังสือชี้แจงความคืบหน้าส่งตัวแทนจำหน่ายต่างประเทศทุกราย', 'ฝ่ายธุรกิจต่างประเทศ', 'ก.ค. 2568'],
    ['4', 'เสนอรายละเอียดวงเงินและผลกระทบของการเร่งแผนเปลี่ยนชุดหัวบรรจุสายที่ 2', 'ฝ่ายวิศวกรรม', 'ประชุมครั้งถัดไป'],
    ['5', 'สรุปสาเหตุรากข้อร้องเรียนอนุภาคแขวนลอย Gentamicin T P และยืนยันประสิทธิผลมาตรการแก้ไข',
     'ฝ่ายประกันคุณภาพ', 'ก.ย. 2568'],
    ['6', 'ทบทวนขั้นตอนปฏิบัติการจัดการข้อร้องเรียน เพิ่มเกณฑ์ยกระดับเรื่องกรณีร้องเรียนซ้ำ',
     'ฝ่ายประกันคุณภาพ', 'ไตรมาส 3/2568'],
    ['7', 'สรุปสาระสำคัญประกาศหลักเกณฑ์ฉบับใหม่ เทียบกับแนวปฏิบัติปัจจุบัน และระบุช่องว่าง',
     'ฝ่ายทะเบียนยา', 'ประชุมครั้งถัดไป'],
    ['8', 'จัดทำรายงานกำไรขั้นต้นแยกตามช่องทางจำหน่าย', 'ฝ่ายบัญชีและการเงิน / ฝ่ายขาย', 'ก.ค. 2568'],
    ['9', 'จัดทำร่างแนวปฏิบัติการใช้เครื่องมือปัญญาประดิษฐ์ในการปฏิบัติงาน',
     'ฝ่ายเทคโนโลยีสารสนเทศ / ฝ่ายประกันคุณภาพ', 'ประชุมครั้งถัดไป'],
    ['10', 'จัดทำโครงสร้างราคาและเงื่อนไขเสนอราคาประมูล พร้อมแผนรองรับกรณีไม่ได้รับงาน',
     'ฝ่ายขายและการตลาด / ฝ่ายบัญชีและการเงิน', 'ก่อนวันยื่นเสนอราคา'],
    ['11', 'จัดทำแผนการเบิกจ่ายงบลงทุนสายบรรจุที่ 2 แยกสองระยะ',
     'ฝ่ายวิศวกรรม / ฝ่ายบัญชีและการเงิน', 'ประชุมครั้งถัดไป'],
    ['12', 'เตรียมความพร้อมด้านเอกสารรองรับการตรวจประเมินระบบคุณภาพ', 'ทุกฝ่าย', 'ไตรมาส 3/2568'],
]


APPENDIX_SALES = (
    ['เดือน', 'โรงพยาบาลรัฐ', 'โรงพยาบาลเอกชน', 'ร้านขายยา', 'ส่งออก', 'รวม'],
    [['เมษายน 2568', '6.72', '2.98', '3.66', '2.69', '16.56'],
     ['พฤษภาคม 2568', '6.98', '3.13', '3.85', '2.94', '17.45'],
     ['มิถุนายน 2568', '7.22', '3.31', '4.03', '3.27', '18.34'],
     ['รวมไตรมาส 2', '20.92', '9.42', '11.54', '8.90', '52.35'],
     ['รวมไตรมาส 1 (เทียบ)', '19.95', '8.81', '10.87', '7.35', '48.97'],
     ['เปลี่ยนแปลง', '+4.9%', '+6.9%', '+6.2%', '+21.1%', '+6.9%'],
     ],
    [4.0, 3.4, 3.6, 3.0, 2.8, 2.6])

APPENDIX_COMPLAINTS = (
    ['ผลิตภัณฑ์', 'จำนวนข้อร้องเรียน Q2', 'จำนวน Q1', 'ประเภทข้อบกพร่องหลัก', 'สถานะ'],
    [['Gentamicin T P', '7', '0', 'อนุภาคแขวนลอยในหลอดบรรจุ', 'ปิดเรื่องแล้วทุกเรื่อง / CAPA-2568-014 ยังไม่ปิด'],
     ['Ceftriaxone T P', '3', '4', 'ฉลากพิมพ์เลื่อน', 'ปิดเรื่องแล้ว'],
     ['Paramol T P', '3', '2', 'จำนวนในกล่องไม่ครบ', 'ปิดเรื่องแล้ว'],
     ['Furosemide Injection T P', '2', '3', 'กล่องบุบระหว่างขนส่ง', 'ปิดเรื่องแล้ว'],
     ['Vitamin C Injection T P', '2', '1', 'สีของสารละลายเปลี่ยน', 'อยู่ระหว่างสอบสวน 1 เรื่อง'],
     ['Transic', '2', '2', 'ซีลฝาขวดไม่แน่น', 'ปิดเรื่องแล้ว'],
     ['ผลิตภัณฑ์อื่น รวม', '2', '2', 'เอกสารกำกับและการจัดส่ง', 'ปิดเรื่องแล้ว'],
     ['รวม', '21', '14', '-', '-'],
     ],
    [4.6, 3.4, 2.2, 4.6, 5.2])

APPENDIX_CAPA = (
    ['เลขที่ CAPA', 'เรื่อง', 'ผู้รับผิดชอบ', 'เปิดเมื่อ', 'กำหนดปิด', 'สถานะ'],
    [['CAPA-2568-006', 'ปรับปรุงการควบคุมการพิมพ์ฉลากสายบรรจุที่ 1', 'ฝ่ายผลิต', 'ก.พ. 2568', 'ก.ค. 2568',
      'ดำเนินการแล้ว รอยืนยันประสิทธิผล'],
     ['CAPA-2568-009', 'ทบทวนวิธีนับจำนวนเม็ดก่อนบรรจุกล่อง', 'ฝ่ายผลิต', 'มี.ค. 2568', 'ส.ค. 2568',
      'อยู่ระหว่างดำเนินการ'],
     ['CAPA-2568-011', 'ปรับปรุงบรรจุภัณฑ์ขนส่งเส้นทางภาคเหนือ', 'ฝ่ายคลังสินค้า', 'เม.ย. 2568', 'ก.ย. 2568',
      'อยู่ระหว่างดำเนินการ'],
     ['CAPA-2568-014', 'สอบสวนสาเหตุรากอนุภาคแขวนลอย สายบรรจุที่ 2', 'ฝ่ายประกันคุณภาพ', 'ก.ค. 2568',
      'ก.ย. 2568', 'อยู่ระหว่างสอบสวน — ลำดับความสำคัญสูงสุด'],
     ['CAPA-2567-038', 'ปรับปรุงระบบเฝ้าระวังอุณหภูมิระหว่างขนส่งส่งออก', 'ฝ่ายคลังสินค้า', 'พ.ย. 2567',
      'มี.ค. 2568', 'เกินกำหนด 4 เดือน'],
     ['CAPA-2567-041', 'ทบทวนขั้นตอนการสอบเทียบเครื่องตรวจสอบอนุภาค', 'ฝ่ายวิศวกรรม', 'ธ.ค. 2567',
      'เม.ย. 2568', 'เกินกำหนด 3 เดือน'],
     ],
    [3.2, 5.4, 3.2, 2.4, 2.4, 4.6])


def _appendices(doc, f):
    doc.add_page_break()
    heading(doc, 'ภาคผนวก ก  สรุปยอดขายไตรมาสที่ 2 ปี 2568 แยกตามช่องทางจำหน่าย (หน่วย: ล้านบาท)', f, 12)
    table(doc, APPENDIX_SALES[0], APPENDIX_SALES[1], f, APPENDIX_SALES[2])
    doc.add_paragraph()
    body(doc, 'ข้อสังเกตของฝ่ายเลขานุการ: ช่องทางส่งออกเติบโตสูงที่สุดในไตรมาสนี้ '
              'ขณะที่ช่องทางโรงพยาบาลรัฐเติบโตต่ำที่สุด ทั้งที่ยังเป็นช่องทางที่มีสัดส่วนยอดขายสูงสุด '
              'เมื่อพิจารณาประกอบกับผลการประมูลรอบเดือนกรกฎาคมตามวาระที่ 9 '
              'ที่ประชุมควรเตรียมแผนรองรับกรณียอดช่องทางโรงพยาบาลรัฐลดลงในครึ่งปีหลัง', f)

    doc.add_page_break()
    heading(doc, 'ภาคผนวก ข  สรุปข้อร้องเรียนผลิตภัณฑ์ไตรมาสที่ 2 ปี 2568', f, 12)
    table(doc, APPENDIX_COMPLAINTS[0], APPENDIX_COMPLAINTS[1], f, APPENDIX_COMPLAINTS[2])
    doc.add_paragraph()
    body(doc, 'ข้อสังเกตของฝ่ายเลขานุการ: การเพิ่มขึ้นของจำนวนข้อร้องเรียนจาก 14 เรื่องในไตรมาสที่ 1 '
              'เป็น 21 เรื่องในไตรมาสที่ 2 มาจากกลุ่มข้อร้องเรียนของผลิตภัณฑ์ Gentamicin T P เพียงเรื่องเดียว '
              'จำนวน 7 เรื่อง หากไม่นับกลุ่มนี้ จำนวนข้อร้องเรียนของผลิตภัณฑ์อื่นอยู่ในระดับเดียวกับไตรมาสก่อน', f)

    doc.add_page_break()
    heading(doc, 'ภาคผนวก ค  สถานะมาตรการแก้ไขและป้องกันที่ยังไม่ปิด ณ วันที่ประชุม', f, 12)
    table(doc, APPENDIX_CAPA[0], APPENDIX_CAPA[1], f, APPENDIX_CAPA[2])
    doc.add_paragraph()
    body(doc, 'ข้อสังเกตของฝ่ายเลขานุการ: มีมาตรการแก้ไขและป้องกันที่เกินกำหนดปิดจำนวน 2 เรื่อง '
              'ซึ่งทั้งสองเรื่องเกี่ยวข้องกับการเฝ้าระวังคุณภาพ และเรื่องหนึ่งเกี่ยวข้องโดยตรง '
              'กับการสอบเทียบเครื่องตรวจสอบอนุภาค อันเป็นประเด็นเดียวกับที่รายงานในวาระที่ 4 และวาระที่ 5', f)

def build_minutes():
    f = FONT_DOCX
    doc = new_doc(f)
    banner(doc, f)
    title(doc, 'บริษัท ที.พี. ดรัก แลบบอราทอรี่ (1969) จำกัด', f, 14)
    title(doc, 'รายงานการประชุมคณะผู้บริหาร ครั้งที่ 2/2568', f, 16)
    p = doc.add_paragraph()
    r = p.add_run('วันพฤหัสบดีที่ 17 กรกฎาคม 2568  เวลา 13:00 – 16:20 นาฬิกา\n'
                  'ณ ห้องประชุมใหญ่ สำนักงานสุขุมวิท 62')
    r.font.name = f
    r.font.size = Pt(10.5)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    heading(doc, 'ผู้เข้าประชุม', f)
    bullets(doc, [
        'ประธานที่ประชุม (กรรมการผู้จัดการ)',
        'ผู้อำนวยการฝ่ายขายและการตลาด',
        'ผู้จัดการฝ่ายธุรกิจต่างประเทศ',
        'ผู้อำนวยการฝ่ายผลิต',
        'ผู้จัดการฝ่ายวิศวกรรม',
        'ผู้อำนวยการฝ่ายประกันคุณภาพ',
        'ผู้จัดการฝ่ายควบคุมคุณภาพ',
        'ผู้จัดการฝ่ายทะเบียนยา',
        'ผู้อำนวยการฝ่ายบัญชีและการเงิน',
        'ผู้จัดการฝ่ายจัดซื้อ',
        'ผู้จัดการฝ่ายทรัพยากรบุคคล',
        'ผู้จัดการฝ่ายเทคโนโลยีสารสนเทศ',
        'เลขานุการที่ประชุม (ผู้จัดการสำนักกรรมการผู้จัดการ)',
    ], f)
    heading(doc, 'ผู้ไม่เข้าประชุม', f)
    bullets(doc, ['ผู้จัดการฝ่ายคลังสินค้าและกระจายสินค้า (ลาป่วย)'], f)
    body(doc, 'หมายเหตุ: เอกสารนี้อ้างถึงผู้เข้าประชุมด้วยตำแหน่งเท่านั้น '
              'ไม่มีการระบุชื่อบุคคล เนื่องจากเป็นเอกสารสมมติสำหรับการฝึกอบรม', f)

    for head, blocks in MINUTES:
        heading(doc, head, f, 12.5)
        for kind, payload in blocks:
            if kind == 'p':
                body(doc, payload, f)
            elif kind == 'h':
                heading(doc, payload, f, 11)
            elif kind == 'b':
                bullets(doc, payload, f)
            elif kind == 't':
                headers, rows, widths = payload
                table(doc, headers, rows, f, widths)
                doc.add_paragraph()

    doc.add_page_break()
    heading(doc, 'สรุปมติและงานที่ต้องติดตาม', f, 13)
    table(doc, ['ลำดับ', 'งานที่ต้องดำเนินการ', 'ผู้รับผิดชอบ', 'กำหนดเสร็จ'], ACTIONS, f,
          [1.6, 8.6, 5.0, 3.8])
    doc.add_paragraph()
    body(doc, 'ผู้จดรายงาน: เลขานุการที่ประชุม        ผู้ตรวจรายงาน: ประธานที่ประชุม', f)
    _appendices(doc, f)
    banner(doc, f)

    path = os.path.join(OUT, 'TP-04-exec-meeting-minutes-Q2.docx')
    doc.save(path)
    return path


# ================================== TP-06 อีเมลลูกค้าต่างประเทศ (ภาษาอังกฤษ)

EMAILS = [
    dict(no=1, frm='Sales Manager, Golden Bough Pharma Trading Co., Ltd. (Yangon, Myanmar)',
         to='Export Sales, T.P. Drug Laboratories (1969) Co., Ltd.',
         date='9 July 2025', subj='THIRD FOLLOW-UP: Ceftriaxone registration dossier - no response since April',
         body=[
             'Dear Export Sales Team,',
             'This is my third message on the same subject. I wrote on 14 April and again on 26 May, and I have '
             'not received a substantive reply to either.',
             'Our registration file for Ceftriaxone T P was submitted to the authority here in November 2024. '
             'In February the authority requested a second round of supporting documents from your side: the '
             'updated stability summary, the finished-product specification with the revised impurity limits, '
             'and a current GMP certificate with an English translation. We forwarded that request to you the '
             'same week.',
             'We have now missed the original March target by more than three months. Two hospital tenders that '
             'we had planned to enter this year require a valid registration number, and we could not '
             'participate in either. Our commercial team is asking me whether we should look for an alternative '
             'supplier for this molecule, and I am running out of arguments.',
             'I am not asking for the documents today. I am asking for one thing only: a realistic date. If the '
             'stability summary cannot be ready before September, tell me that and I will manage expectations '
             'here. What I cannot manage is silence.',
             'Please reply by Friday this week.',
             'Best regards,',
             'Sales Manager, Golden Bough Pharma Trading Co., Ltd.',
         ]),
    dict(no=2, frm='Regulatory Officer, Mekong Medical Supply Co., Ltd. (Phnom Penh, Cambodia)',
         to='Regulatory Affairs, T.P. Drug Laboratories (1969) Co., Ltd.',
         date='2 July 2025', subj='Clatacef - status enquiry and CPP request',
         body=[
             'Dear Regulatory Affairs,',
             'Thank you for the documents sent on 12 June. The authority has confirmed our submission is under '
             'review and we expect an outcome within eight weeks.',
             'One item is still outstanding. The reviewer has asked for a Certificate of a Pharmaceutical '
             'Product issued within the last twelve months. The copy in our file is dated March 2024 and will '
             'not be accepted.',
             'Could you please advise how long a fresh CPP normally takes to obtain, so that we can inform the '
             'reviewer of a realistic date?',
             'Kind regards,',
             'Regulatory Officer, Mekong Medical Supply Co., Ltd.',
         ]),
    dict(no=3, frm='Quality Assurance Head, Isla Health Distribution Inc. (Manila, Philippines)',
         to='Quality Assurance, T.P. Drug Laboratories (1969) Co., Ltd.',
         date='24 June 2025', subj='Request: Certificates of Analysis for Paramol T P - three previous batches',
         body=[
             'Dear Quality Assurance,',
             'As part of our local registration variation, the authority has requested Certificates of Analysis '
             'for three consecutive commercial batches of Paramol T P, together with the corresponding batch '
             'release statements.',
             'Please send the CoA for the three most recent batches shipped to us, in PDF with authorised '
             'signature. If any of those batches were released under a different specification version, please '
             'indicate which version applies to each.',
             'Our submission window closes on 31 July, so we would appreciate receiving these by 15 July.',
             'Thank you,',
             'Quality Assurance Head, Isla Health Distribution Inc.',
         ]),
    dict(no=4, frm='Procurement Director, Truong Phat Pharma JSC (Hanoi, Vietnam)',
         to='Export Sales, T.P. Drug Laboratories (1969) Co., Ltd.',
         date='19 June 2025', subj='Request for price review - 2026 supply agreement',
         body=[
             'Dear Export Sales,',
             'We are preparing our 2026 budget and would like to open the annual price discussion early.',
             'Two points from our side. First, the local currency has moved against the US dollar by roughly '
             'four per cent since our last agreement, which has compressed our margin on your products. Second, '
             'two competing generics entered our market in the first quarter at prices approximately eight per '
             'cent below our current landed cost for Tramadol T P.',
             'We are not asking for a unilateral discount. We would rather discuss a volume-linked structure: a '
             'firm annual commitment from us in exchange for a tiered price. If that is workable, we can '
             'provide indicative volumes within two weeks.',
             'Please let me know whether your team is open to this approach.',
             'Best regards,',
             'Procurement Director, Truong Phat Pharma JSC',
         ]),
    dict(no=5, frm='Business Development Manager, Sahel Medipharm Ltd. (Lagos, Nigeria)',
         to='Export Sales, T.P. Drug Laboratories (1969) Co., Ltd.',
         date='16 June 2025', subj='New market registration - plant documentation package required',
         body=[
             'Dear Sir or Madam,',
             'Following our meeting at the trade exhibition in March, we would like to proceed with the first '
             'registration of Vitamin C Injection T P in our market.',
             'To open the file we need the following from your side: a valid GMP certificate, the site master '
             'file, the product dossier in CTD format, and a letter of authorisation naming us as the '
             'registration holder. The authority will also require the certificate to be legalised.',
             'Please confirm which of these you can provide and the expected timeline for each. We would also '
             'appreciate an indication of your minimum order quantity for the first shipment.',
             'Yours faithfully,',
             'Business Development Manager, Sahel Medipharm Ltd.',
         ]),
    dict(no=6, frm='Registration Manager, Serendib Pharma (Pvt) Ltd. (Colombo, Sri Lanka)',
         to='Regulatory Affairs, T.P. Drug Laboratories (1969) Co., Ltd.',
         date='11 June 2025', subj='Furosemide Injection T P - renewal due August 2025',
         body=[
             'Dear Regulatory Affairs,',
             'A reminder that the registration for Furosemide Injection T P in our market expires in August '
             'this year. The renewal application must be lodged no later than sixty days before expiry, which '
             'means we need your documents in hand by early July.',
             'The renewal set is the same as last time. If nothing has changed in the specification or the '
             'manufacturing site, a signed declaration to that effect will be sufficient.',
             'Please confirm that this is on your schedule.',
             'Regards,',
             'Registration Manager, Serendib Pharma (Pvt) Ltd.',
         ]),
    dict(no=7, frm='Logistics Supervisor, Lanexang Medical Import Co., Ltd. (Vientiane, Lao PDR)',
         to='Customer Service, T.P. Drug Laboratories (1969) Co., Ltd.',
         date='5 June 2025', subj='Complaint: temperature excursion on shipment INV-2568-0431',
         body=[
             'Dear Customer Service,',
             'Shipment INV-2568-0431 arrived at our warehouse on 3 June. The data logger inside the second '
             'pallet recorded 34 degrees Celsius for approximately eleven hours during the road leg, which is '
             'above the labelled storage condition.',
             'We have quarantined the affected cartons and have not released them for sale. Please advise '
             'whether the goods remain fit for distribution, and if not, how you wish to proceed with '
             'replacement or credit.',
             'We would also like to understand what will be done to prevent this on future shipments, as this '
             'is the second excursion on this route in twelve months.',
             'Regards,',
             'Logistics Supervisor, Lanexang Medical Import Co., Ltd.',
         ]),
    dict(no=8, frm='Managing Director, Al-Noor Medical Trading (Aden, Yemen)',
         to='Export Sales, T.P. Drug Laboratories (1969) Co., Ltd.',
         date='28 May 2025', subj='Increased forecast for H2 2025 - supply assurance needed',
         body=[
             'Dear Export Sales,',
             'Demand in our market has grown faster than we projected. Our forecast for the second half of this '
             'year is approximately forty per cent above the first half, concentrated in the injectable '
             'antibiotic range.',
             'Before we commit to our customers, we need written assurance on two points: that you can supply '
             'the forecast volumes on a monthly schedule, and that lead time will not exceed six weeks from '
             'order confirmation.',
             'If your capacity cannot cover the full forecast, please tell us now and we will phase our '
             'commitments accordingly. A clear partial commitment is more useful to us than an optimistic full '
             'one.',
             'Best regards,',
             'Managing Director, Al-Noor Medical Trading',
         ]),
]


def build_emails():
    f = FONT_DOCX
    doc = new_doc(f)
    banner(doc, f)
    title(doc, 'T.P. Drug Laboratories (1969) Co., Ltd.', f, 14)
    title(doc, 'Export Customer Correspondence — Collected Inbox Extract', f, 15)
    body(doc, 'Period covered: 28 May 2025 – 9 July 2025.  Eight messages, most recent first order of arrival '
              'preserved as received.  Compiled for internal review.', f)
    body(doc, 'หมายเหตุ: อีเมลทั้งหมดในเอกสารนี้แต่งขึ้นเพื่อการฝึกอบรม '
              'ชื่อบริษัทลูกค้า ชื่อผู้ส่ง และเหตุการณ์ไม่มีอยู่จริง', f)

    for e in EMAILS:
        doc.add_paragraph()
        heading(doc, 'Message %d of 8' % e['no'], f, 11.5)
        for label, value in (('From:', e['frm']), ('To:', e['to']),
                             ('Date:', e['date']), ('Subject:', e['subj'])):
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(1)
            r1 = p.add_run(label + '  ')
            r1.bold = True
            r1.font.name = f
            r1.font.size = Pt(10)
            r2 = p.add_run(value)
            r2.font.name = f
            r2.font.size = Pt(10)
        doc.add_paragraph()
        for para in e['body']:
            body(doc, para, f)

    banner(doc, f)
    path = os.path.join(OUT, 'TP-06-export-customer-emails.docx')
    doc.save(path)
    return path


# ============== TP-05 ประกาศหลักเกณฑ์ฉบับสมมติ  /  TP-07 SOP การจัดการข้อร้องเรียน
# ออกแบบให้ SOP (ฉบับปี 2566) ล้าสมัยกว่าประกาศ (ปี 2568) โดยเจตนา ใน 3 จุด:
#   1. กรอบเวลารายงานข้อร้องเรียนวิกฤต  ประกาศ 3 วันทำการ  /  SOP 7 วันทำการ
#   2. ความถี่การทบทวนแนวโน้ม           ประกาศ ทุกไตรมาส   /  SOP ปีละครั้ง
#   3. เกณฑ์ยกระดับเรื่องกรณีร้องเรียนซ้ำ ประกาศ กำหนดไว้   /  SOP ไม่มีเลย
# ทำให้โจทย์ "ประกาศฉบับนี้ทำให้เราต้องแก้ SOP ตรงไหน" มีคำตอบที่ตรวจสอบได้จริง

NOTICE = [
    ('หมวด 1  บททั่วไป', [
        ('p', 'ข้อ 1  ประกาศนี้เรียกว่า "ประกาศเรื่อง หลักเกณฑ์ วิธีการ และเงื่อนไขการรายงานข้อร้องเรียน '
              'ผลิตภัณฑ์ยาและการเรียกคืนผลิตภัณฑ์ยา พ.ศ. 2568"'),
        ('p', 'ข้อ 2  ประกาศนี้ให้ใช้บังคับเมื่อพ้นกำหนดหนึ่งร้อยแปดสิบวันนับแต่วันประกาศเป็นต้นไป'),
        ('p', 'ข้อ 3  ให้ยกเลิกประกาศเรื่องหลักเกณฑ์การรายงานข้อร้องเรียนผลิตภัณฑ์ยา พ.ศ. 2562 '
              'และให้ใช้ประกาศนี้แทน'),
        ('p', 'ข้อ 4  ในประกาศนี้'),
        ('b', [
            '"ผู้รับอนุญาต" หมายความว่า ผู้รับอนุญาตผลิตยาแผนปัจจุบัน',
            '"ข้อร้องเรียน" หมายความว่า การแจ้งข้อบกพร่องหรือข้อสงสัยในคุณภาพ ความปลอดภัย '
            'หรือประสิทธิผลของผลิตภัณฑ์ยา ที่ได้รับจากบุคคลภายนอกหรือจากภายในองค์กร',
            '"ข้อร้องเรียนวิกฤต" หมายความว่า ข้อร้องเรียนที่อาจก่อให้เกิดอันตรายถึงชีวิต '
            'หรืออาจทำให้เกิดความเสียหายต่อสุขภาพอย่างร้ายแรง รวมถึงกรณีผลิตภัณฑ์ปลอมปน '
            'การระบุฉลากผิดที่มีผลต่อขนาดยา และการปนเปื้อนที่ตรวจพบในผลิตภัณฑ์ปราศจากเชื้อ',
            '"ข้อร้องเรียนสำคัญ" หมายความว่า ข้อร้องเรียนที่ผลิตภัณฑ์ไม่เป็นไปตามข้อกำหนด '
            'แต่ไม่ถึงระดับที่จะก่อให้เกิดอันตรายร้ายแรง',
            '"ข้อร้องเรียนเล็กน้อย" หมายความว่า ข้อร้องเรียนที่ไม่กระทบคุณภาพ ความปลอดภัย '
            'หรือประสิทธิผลของผลิตภัณฑ์',
            '"ข้อร้องเรียนซ้ำ" หมายความว่า ข้อร้องเรียนที่มีลักษณะของข้อบกพร่องอย่างเดียวกัน '
            'ในผลิตภัณฑ์เดียวกัน ไม่ว่าจะเป็นรุ่นผลิตเดียวกันหรือต่างรุ่นผลิต',
        ]),
        ('p', 'ข้อ 5  ผู้รับอนุญาตต้องจัดให้มีระบบการจัดการข้อร้องเรียนเป็นลายลักษณ์อักษร '
              'และต้องกำหนดผู้รับผิดชอบซึ่งเป็นเภสัชกรผู้มีหน้าที่ปฏิบัติการหรือผู้ที่ได้รับมอบหมายไว้อย่างชัดเจน'),
    ]),
    ('หมวด 2  การรับและบันทึกข้อร้องเรียน', [
        ('p', 'ข้อ 6  ผู้รับอนุญาตต้องบันทึกข้อร้องเรียนทุกเรื่องที่ได้รับ ไม่ว่าจะได้รับด้วยวิธีใด '
              'ภายในหนึ่งวันทำการนับแต่วันที่ได้รับแจ้ง'),
        ('p', 'ข้อ 7  บันทึกข้อร้องเรียนต้องมีรายการอย่างน้อยดังต่อไปนี้'),
        ('b', [
            'เลขที่อ้างอิงข้อร้องเรียนซึ่งไม่ซ้ำกัน',
            'วันที่ได้รับแจ้ง และช่องทางที่ได้รับแจ้ง',
            'ชื่อผลิตภัณฑ์ ความแรง รูปแบบยา และเลขที่รุ่นผลิต',
            'รายละเอียดของข้อบกพร่องที่ได้รับแจ้ง',
            'การจำแนกระดับความรุนแรงตามข้อ 12 และเหตุผลของการจำแนก',
            'ชื่อผู้รับผิดชอบการสอบสวน และวันที่เปิดการสอบสวน',
            'ผลการสอบสวน สาเหตุราก และมาตรการแก้ไขและป้องกัน',
            'วันที่ปิดเรื่อง และเลขที่อ้างอิงมาตรการแก้ไขและป้องกัน',
        ]),
        ('p', 'ข้อ 8  ผู้รับอนุญาตต้องจำแนกระดับความรุนแรงของข้อร้องเรียนภายในสองวันทำการ '
              'นับแต่วันที่บันทึกตามข้อ 6'),
        ('p', 'ข้อ 9  ในกรณีที่ยังไม่สามารถจำแนกระดับความรุนแรงได้อย่างแน่ชัด '
              'ให้จำแนกในระดับที่สูงกว่าไว้ก่อน และทบทวนเมื่อได้ข้อมูลเพิ่มเติม'),
        ('p', 'ข้อ 10  ผู้รับอนุญาตต้องเปิดการสอบสวนข้อร้องเรียนวิกฤตภายในสองวันทำการ '
              'ข้อร้องเรียนสำคัญภายในห้าวันทำการ และข้อร้องเรียนเล็กน้อยภายในสิบวันทำการ '
              'นับแต่วันที่จำแนกระดับความรุนแรง'),
        ('p', 'ข้อ 11  ผู้รับอนุญาตต้องแจ้งผลการพิจารณาให้ผู้แจ้งข้อร้องเรียนทราบเป็นลายลักษณ์อักษร '
              'ภายในสามสิบวันนับแต่วันที่ปิดเรื่อง'),
    ]),
    ('หมวด 3  การจำแนกระดับความรุนแรงและกรอบเวลาการรายงาน', [
        ('p', 'ข้อ 12  ให้จำแนกข้อร้องเรียนเป็นสามระดับ คือ ข้อร้องเรียนวิกฤต ข้อร้องเรียนสำคัญ '
              'และข้อร้องเรียนเล็กน้อย ตามนิยามในข้อ 4'),
        ('p', 'ข้อ 13  ผู้รับอนุญาตต้องรายงานข้อร้องเรียนต่อพนักงานเจ้าหน้าที่ตามกรอบเวลาดังต่อไปนี้'),
        ('t', (['ระดับความรุนแรง', 'กรอบเวลารายงานครั้งแรก', 'กรอบเวลารายงานผลการสอบสวน', 'ช่องทางรายงาน'],
               [['ข้อร้องเรียนวิกฤต', 'ภายใน 3 วันทำการนับแต่วันที่จำแนกระดับ',
                 'ภายใน 30 วันนับแต่วันรายงานครั้งแรก', 'ระบบรายงานอิเล็กทรอนิกส์ และหนังสือ'],
                ['ข้อร้องเรียนสำคัญ', 'ภายใน 15 วันทำการนับแต่วันที่จำแนกระดับ',
                 'ภายใน 60 วันนับแต่วันรายงานครั้งแรก', 'ระบบรายงานอิเล็กทรอนิกส์'],
                ['ข้อร้องเรียนเล็กน้อย', 'รวมในรายงานสรุปรายไตรมาส',
                 'รวมในรายงานสรุปรายไตรมาส', 'รายงานสรุปรายไตรมาส'],
                ],
               [4.0, 5.2, 5.0, 4.4])),
        ('p', 'ข้อ 14  ในกรณีที่การสอบสวนไม่แล้วเสร็จภายในกรอบเวลาตามข้อ 13 '
              'ผู้รับอนุญาตต้องรายงานความคืบหน้าพร้อมเหตุผลของความล่าช้าและกำหนดเวลาที่คาดว่าจะแล้วเสร็จ'),
        ('p', 'ข้อ 15  ผู้รับอนุญาตต้องกำหนดเกณฑ์การยกระดับเรื่องไว้เป็นลายลักษณ์อักษร '
              'โดยอย่างน้อยต้องยกระดับเรื่องเป็นการสอบสวนเชิงระบบเมื่อได้รับข้อร้องเรียนซ้ำ '
              'ตั้งแต่สามเรื่องขึ้นไปในผลิตภัณฑ์เดียวกันภายในระยะเวลาเก้าสิบวัน '
              'ไม่ว่าข้อร้องเรียนแต่ละเรื่องจะจำแนกอยู่ในระดับใด'),
        ('p', 'ข้อ 16  การยกระดับเรื่องตามข้อ 15 ต้องมีการประเมินความจำเป็นในการระงับการจำหน่าย '
              'หรือการเรียกคืนผลิตภัณฑ์ และต้องบันทึกผลการประเมินไว้เป็นหลักฐาน'),
        ('p', 'ข้อ 17  ผู้รับอนุญาตต้องแจ้งพนักงานเจ้าหน้าที่ทราบเมื่อมีการยกระดับเรื่องตามข้อ 15 '
              'ภายในห้าวันทำการ'),
        ('p', 'ข้อ 18  การจำแนกระดับความรุนแรงและการยกระดับเรื่องต้องได้รับการทบทวนและลงนามรับรอง '
              'โดยผู้รับผิดชอบตามข้อ 5'),
    ]),
    ('หมวด 4  การทบทวนแนวโน้มข้อร้องเรียน', [
        ('p', 'ข้อ 19  ผู้รับอนุญาตต้องจัดให้มีการทบทวนแนวโน้มข้อร้องเรียนเป็นลายลักษณ์อักษร '
              'อย่างน้อยทุกไตรมาส'),
        ('p', 'ข้อ 20  การทบทวนตามข้อ 19 ต้องครอบคลุมอย่างน้อยดังต่อไปนี้'),
        ('b', [
            'จำนวนข้อร้องเรียนแยกตามผลิตภัณฑ์ ระดับความรุนแรง และประเภทข้อบกพร่อง',
            'การเปรียบเทียบกับไตรมาสก่อนหน้าและกับช่วงเดียวกันของปีก่อน',
            'การระบุข้อร้องเรียนซ้ำและผลการดำเนินการตามข้อ 15',
            'สถานะของมาตรการแก้ไขและป้องกันที่ยังไม่ปิด',
            'ระยะเวลาเฉลี่ยจากวันรับแจ้งถึงวันปิดเรื่อง',
        ]),
        ('p', 'ข้อ 21  ผลการทบทวนตามข้อ 19 ต้องเสนอต่อผู้บริหารระดับสูงของผู้รับอนุญาต '
              'และต้องบันทึกการพิจารณาของผู้บริหารไว้เป็นหลักฐาน'),
        ('p', 'ข้อ 22  ผู้รับอนุญาตต้องนำผลการทบทวนตามข้อ 19 ไปประกอบการทบทวนคุณภาพผลิตภัณฑ์ประจำปี'),
    ]),
    ('หมวด 5  การเรียกคืนผลิตภัณฑ์ยา', [
        ('p', 'ข้อ 23  ผู้รับอนุญาตต้องจัดให้มีขั้นตอนการเรียกคืนผลิตภัณฑ์เป็นลายลักษณ์อักษร '
              'และต้องทดสอบความพร้อมของระบบเรียกคืนอย่างน้อยปีละหนึ่งครั้ง'),
        ('p', 'ข้อ 24  ให้จำแนกระดับการเรียกคืนเป็นสามระดับตามความรุนแรงของความเสี่ยงต่อผู้ใช้ยา'),
        ('p', 'ข้อ 25  ในกรณีเรียกคืนระดับที่หนึ่ง ผู้รับอนุญาตต้องแจ้งพนักงานเจ้าหน้าที่ '
              'ภายในยี่สิบสี่ชั่วโมงนับแต่มีการตัดสินใจเรียกคืน และต้องเริ่มดำเนินการเรียกคืน '
              'ภายในสี่สิบแปดชั่วโมง'),
        ('p', 'ข้อ 26  ในกรณีเรียกคืนระดับที่สอง ให้แจ้งภายในสามวันทำการ '
              'และในกรณีเรียกคืนระดับที่สาม ให้แจ้งภายในเจ็ดวันทำการ'),
        ('p', 'ข้อ 27  ผู้รับอนุญาตต้องจัดทำรายงานผลการเรียกคืน โดยระบุปริมาณที่กระจายออกไป '
              'ปริมาณที่ได้รับคืน อัตราการได้รับคืนเป็นร้อยละ และการจัดการผลิตภัณฑ์ที่ได้รับคืน'),
        ('p', 'ข้อ 28  ในกรณีที่ผลิตภัณฑ์ที่เรียกคืนได้ถูกส่งออกไปจำหน่ายในต่างประเทศ '
              'ผู้รับอนุญาตต้องแจ้งผู้นำเข้าในต่างประเทศทุกรายภายในกรอบเวลาเดียวกับที่กำหนดในข้อ 25 และข้อ 26'),
        ('p', 'ข้อ 29  การระงับการจำหน่ายผลิตภัณฑ์เป็นการชั่วคราวเพื่อรอผลการสอบสวน '
              'ไม่ถือเป็นการเรียกคืน แต่ต้องบันทึกเหตุผล ขอบเขต และระยะเวลาไว้เป็นหลักฐาน '
              'และต้องแจ้งพนักงานเจ้าหน้าที่ภายในห้าวันทำการ'),
        ('p', 'ข้อ 30  ผู้รับอนุญาตต้องทบทวนความเหมาะสมของการตัดสินใจตามข้อ 29 '
              'อย่างน้อยทุกสามสิบวันจนกว่าจะยุติการระงับการจำหน่าย'),
    ]),
    ('หมวด 6  การเก็บรักษาบันทึกและบทเฉพาะกาล', [
        ('p', 'ข้อ 31  บันทึกข้อร้องเรียน บันทึกการสอบสวน และบันทึกการเรียกคืน '
              'ต้องเก็บรักษาไว้ไม่น้อยกว่าห้าปีนับแต่วันสิ้นอายุของผลิตภัณฑ์รุ่นผลิตที่เกี่ยวข้อง '
              'หรือไม่น้อยกว่าหนึ่งปีนับแต่วันปิดเรื่อง แล้วแต่ระยะเวลาใดจะยาวกว่า'),
        ('p', 'ข้อ 32  บันทึกที่จัดเก็บในรูปอิเล็กทรอนิกส์ต้องมีการควบคุมการเข้าถึง '
              'และต้องสามารถแสดงร่องรอยการแก้ไขเปลี่ยนแปลงได้'),
        ('p', 'ข้อ 33  ผู้รับอนุญาตต้องจัดให้บันทึกตามข้อ 31 พร้อมให้พนักงานเจ้าหน้าที่ตรวจสอบได้ '
              'ณ สถานที่ผลิตในเวลาทำการ'),
        ('p', 'ข้อ 34  ผู้รับอนุญาตที่ดำเนินการอยู่ก่อนวันที่ประกาศนี้ใช้บังคับ '
              'ต้องปรับปรุงเอกสารระบบคุณภาพให้สอดคล้องกับประกาศนี้ '
              'ให้เสร็จสิ้นภายในหนึ่งร้อยแปดสิบวันนับแต่วันที่ประกาศนี้ใช้บังคับ'),
        ('p', 'ข้อ 35  ข้อร้องเรียนที่อยู่ระหว่างการสอบสวนในวันที่ประกาศนี้ใช้บังคับ '
              'ให้ดำเนินการต่อไปตามหลักเกณฑ์เดิมจนแล้วเสร็จ '
              'แต่การทบทวนแนวโน้มตามหมวด 4 ให้เริ่มใช้กับไตรมาสถัดไปทันที'),
        ('p', 'ข้อ 36  ให้ผู้รับอนุญาตประเมินความสอดคล้องของเอกสารระบบคุณภาพที่ใช้อยู่กับประกาศนี้ '
              'และจัดทำบันทึกช่องว่างที่ต้องปรับปรุงพร้อมกำหนดเวลาแล้วเสร็จ '
              'ให้พนักงานเจ้าหน้าที่ตรวจสอบได้'),
    ]),
]

SOP = [
    ('1.  วัตถุประสงค์', [
        ('p', 'เพื่อกำหนดวิธีปฏิบัติในการรับ บันทึก จำแนก สอบสวน และปิดเรื่องข้อร้องเรียนผลิตภัณฑ์ '
              'ให้เป็นไปในแนวทางเดียวกัน และเพื่อให้มั่นใจว่าข้อร้องเรียนทุกเรื่องได้รับการพิจารณา '
              'อย่างเหมาะสมและมีหลักฐานการดำเนินการครบถ้วน'),
    ]),
    ('2.  ขอบเขต', [
        ('p', 'ใช้กับข้อร้องเรียนผลิตภัณฑ์ทุกรายการที่บริษัทผลิตและจำหน่าย '
              'ทั้งที่จำหน่ายในประเทศและส่งออก ครอบคลุมข้อร้องเรียนที่ได้รับจากลูกค้า ตัวแทนจำหน่าย '
              'สถานพยาบาล ผู้บริโภค และที่ตรวจพบภายในองค์กร'),
    ]),
    ('3.  คำจำกัดความ', [
        ('b', [
            'ข้อร้องเรียนผลิตภัณฑ์ — การแจ้งข้อบกพร่องหรือข้อสงสัยในคุณภาพ ความปลอดภัย '
            'หรือประสิทธิผลของผลิตภัณฑ์ที่บริษัทผลิต',
            'ระดับความรุนแรง — การจำแนกข้อร้องเรียนออกเป็น 4 ระดับ คือ วิกฤต สูง ปานกลาง และต่ำ',
            'CAPA — มาตรการแก้ไขและป้องกันการเกิดซ้ำ',
            'การระงับการจำหน่าย — การหยุดการจ่ายผลิตภัณฑ์ออกจากคลังเป็นการชั่วคราวเพื่อรอผลการสอบสวน',
        ]),
    ]),
    ('4.  หน้าที่และความรับผิดชอบ', [
        ('t', (['ผู้รับผิดชอบ', 'หน้าที่'],
               [['ฝ่ายขาย / ฝ่ายลูกค้าสัมพันธ์', 'รับแจ้งข้อร้องเรียนและส่งต่อให้ฝ่ายประกันคุณภาพ '
                                              'ภายใน 2 วันทำการ'],
                ['ฝ่ายประกันคุณภาพ', 'บันทึกข้อร้องเรียน จำแนกระดับความรุนแรง มอบหมายผู้สอบสวน '
                                  'ติดตามความคืบหน้า และปิดเรื่อง'],
                ['ฝ่ายควบคุมคุณภาพ', 'ตรวจสอบตัวอย่างเก็บอ้างอิงและรายงานผลการตรวจวิเคราะห์'],
                ['ฝ่ายผลิต', 'ตรวจสอบบันทึกการผลิตของรุ่นผลิตที่เกี่ยวข้องและร่วมสอบหาสาเหตุราก'],
                ['เภสัชกรผู้มีหน้าที่ปฏิบัติการ', 'พิจารณาอนุมัติการระงับการจำหน่ายและการเรียกคืน '
                                              'และลงนามปิดเรื่อง'],
                ],
               [5.6, 11.4])),
    ]),
    ('5.  ขั้นตอนปฏิบัติ', [
        ('p', '5.1  การรับและบันทึกข้อร้องเรียน'),
        ('b', [
            'ผู้รับแจ้งกรอกแบบฟอร์ม FM-QA-012-01 และส่งให้ฝ่ายประกันคุณภาพภายใน 2 วันทำการ',
            'ฝ่ายประกันคุณภาพบันทึกลงทะเบียนข้อร้องเรียนและออกเลขที่อ้างอิงภายใน 2 วันทำการ '
            'นับแต่วันที่ได้รับแบบฟอร์ม',
            'กรณีข้อมูลไม่ครบถ้วน ให้ประสานผู้แจ้งเพื่อขอข้อมูลเพิ่มเติมภายใน 5 วันทำการ',
        ]),
        ('p', '5.2  การจำแนกระดับความรุนแรง'),
        ('b', [
            'ฝ่ายประกันคุณภาพจำแนกระดับความรุนแรงเป็น 4 ระดับ ภายใน 3 วันทำการนับแต่วันที่บันทึก',
            'กรณีจำแนกเป็นระดับวิกฤตหรือสูง ให้รายงานเภสัชกรผู้มีหน้าที่ปฏิบัติการทันที',
        ]),
        ('p', '5.3  การสอบสวน'),
        ('b', [
            'เปิดการสอบสวนภายใน 5 วันทำการนับแต่วันที่จำแนกระดับความรุนแรง',
            'ตรวจสอบตัวอย่างเก็บอ้างอิง บันทึกการผลิต บันทึกการควบคุมสภาวะแวดล้อม '
            'และบันทึกการสอบเทียบเครื่องมือที่เกี่ยวข้อง',
            'สรุปสาเหตุรากและกำหนดมาตรการแก้ไขและป้องกันให้แล้วเสร็จภายใน 30 วันทำการ',
        ]),
        ('p', '5.4  การรายงานหน่วยงานกำกับดูแล'),
        ('b', [
            'กรณีข้อร้องเรียนระดับวิกฤต ให้รายงานหน่วยงานกำกับดูแลภายใน 7 วันทำการ '
            'นับแต่วันที่จำแนกระดับความรุนแรง',
            'กรณีข้อร้องเรียนระดับสูง ให้รายงานภายใน 20 วันทำการ',
            'กรณีข้อร้องเรียนระดับปานกลางและต่ำ ให้รวมในรายงานสรุปประจำปี',
        ]),
        ('p', '5.5  การระงับการจำหน่ายและการเรียกคืน'),
        ('b', [
            'เภสัชกรผู้มีหน้าที่ปฏิบัติการพิจารณาความจำเป็นในการระงับการจำหน่ายหรือเรียกคืน '
            'โดยอ้างอิงผลการสอบสวนเบื้องต้น',
            'กรณีระงับการจำหน่าย ให้แจ้งฝ่ายคลังสินค้าและฝ่ายขายทันที และบันทึกเหตุผลไว้ในทะเบียน',
            'ขั้นตอนการเรียกคืนให้ปฏิบัติตาม SOP-QA-015 เรื่องการเรียกคืนผลิตภัณฑ์',
        ]),
        ('p', '5.6  การปิดเรื่องและการแจ้งผู้แจ้ง'),
        ('b', [
            'ปิดเรื่องเมื่อมาตรการแก้ไขและป้องกันได้รับการดำเนินการและยืนยันประสิทธิผลแล้ว',
            'แจ้งผลการพิจารณาให้ผู้แจ้งทราบเป็นลายลักษณ์อักษรภายใน 45 วันนับแต่วันปิดเรื่อง',
        ]),
        ('p', '5.7  การทบทวนแนวโน้ม'),
        ('b', [
            'ฝ่ายประกันคุณภาพจัดทำรายงานทบทวนแนวโน้มข้อร้องเรียนปีละหนึ่งครั้ง '
            'เพื่อประกอบการทบทวนคุณภาพผลิตภัณฑ์ประจำปี',
            'รายงานต้องแสดงจำนวนข้อร้องเรียนแยกตามผลิตภัณฑ์และประเภทข้อบกพร่อง',
        ]),
    ]),
    ('6.  บันทึกที่เกี่ยวข้อง', [
        ('b', [
            'FM-QA-012-01  แบบฟอร์มรับแจ้งข้อร้องเรียนผลิตภัณฑ์',
            'FM-QA-012-02  ทะเบียนข้อร้องเรียนผลิตภัณฑ์',
            'FM-QA-012-03  แบบฟอร์มรายงานผลการสอบสวน',
            'FM-QA-012-04  แบบฟอร์มบันทึกการระงับการจำหน่าย',
        ]),
    ]),
    ('7.  เอกสารอ้างอิง', [
        ('b', [
            'ประกาศเรื่องหลักเกณฑ์การรายงานข้อร้องเรียนผลิตภัณฑ์ยา พ.ศ. 2562',
            'SOP-QA-015  การเรียกคืนผลิตภัณฑ์',
            'SOP-QA-021  การจัดการมาตรการแก้ไขและป้องกัน',
        ]),
    ]),
    ('8.  ประวัติการแก้ไข', [
        ('t', (['ฉบับที่', 'วันที่มีผลบังคับใช้', 'สาระสำคัญของการแก้ไข'],
               [['1', '1 มีนาคม 2562', 'จัดทำครั้งแรก'],
                ['2', '1 กรกฎาคม 2564', 'เพิ่มขั้นตอนการระงับการจำหน่าย'],
                ['3', '1 กุมภาพันธ์ 2566', 'ปรับกรอบเวลาการรายงานหน่วยงานกำกับดูแล '
                                        'และเพิ่มแบบฟอร์ม FM-QA-012-04'],
                ],
               [2.2, 4.4, 10.4])),
    ]),
]


def _render_sections(doc, font, sections):
    for head, blocks in sections:
        heading(doc, head, font, 12)
        for kind, payload in blocks:
            if kind == 'p':
                body(doc, payload, font)
            elif kind == 'b':
                bullets(doc, payload, font)
            elif kind == 't':
                headers, rows, widths = payload
                table(doc, headers, rows, font, widths)
                doc.add_paragraph()


def build_notice():
    f = FONT_PDF
    doc = new_doc(f)
    banner(doc, f)
    title(doc, 'ประกาศ (ฉบับสมมติเพื่อการฝึกอบรม)', f, 13)
    title(doc, 'เรื่อง หลักเกณฑ์ วิธีการ และเงื่อนไขการรายงานข้อร้องเรียนผลิตภัณฑ์ยา', f, 15)
    title(doc, 'และการเรียกคืนผลิตภัณฑ์ยา พ.ศ. 2568', f, 15)
    body(doc, 'คำเตือน: ประกาศฉบับนี้ไม่มีอยู่จริง จัดทำขึ้นเพื่อใช้เป็นเอกสารฝึกอบรมเท่านั้น '
              'ห้ามนำไปอ้างอิงเป็นข้อกำหนดทางกฎหมายในทุกกรณี', f)
    _render_sections(doc, f, NOTICE)
    banner(doc, f)

    tmp = os.path.join(OUT, '_tmp-notice.docx')
    doc.save(tmp)
    pdf = os.path.join(OUT, 'TP-05-regulatory-notice.pdf')
    to_pdf(tmp, pdf)
    os.remove(tmp)
    return pdf


def build_sop():
    f = FONT_PDF
    doc = new_doc(f)
    banner(doc, f)
    title(doc, 'บริษัท ที.พี. ดรัก แลบบอราทอรี่ (1969) จำกัด', f, 13)
    title(doc, 'ขั้นตอนปฏิบัติมาตรฐาน  SOP-QA-012', f, 15)
    title(doc, 'เรื่อง การจัดการข้อร้องเรียนผลิตภัณฑ์', f, 14)
    table(doc, ['รหัสเอกสาร', 'ฉบับที่', 'วันที่มีผลบังคับใช้', 'ผู้จัดทำ', 'ผู้อนุมัติ'],
          [['SOP-QA-012', '3', '1 กุมภาพันธ์ 2566', 'ฝ่ายประกันคุณภาพ',
            'เภสัชกรผู้มีหน้าที่ปฏิบัติการ']], f, [3.2, 1.8, 4.0, 3.6, 4.4])
    doc.add_paragraph()
    _render_sections(doc, f, SOP)
    banner(doc, f)

    tmp = os.path.join(OUT, '_tmp-sop.docx')
    doc.save(tmp)
    pdf = os.path.join(OUT, 'TP-07-sop-complaint-handling.pdf')
    to_pdf(tmp, pdf)
    os.remove(tmp)
    return pdf


if __name__ == '__main__':
    os.makedirs(OUT, exist_ok=True)
    for fn in (build_minutes, build_emails, build_notice, build_sop):
        path = fn()
        print('%-46s %8.1f KB' % (os.path.basename(path), os.path.getsize(path) / 1024.0))
