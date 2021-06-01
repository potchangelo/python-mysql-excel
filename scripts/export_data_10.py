# Export เหมือนไฟล์ "export_data_09"
# แต่คราวนี้จะต้องตกแต่งเนื้อหาในไฟล์ Excel ให้ดูดีมีระดับ

import mysql.connector
from openpyxl import Workbook
from openpyxl.styles import Font, NamedStyle, Alignment
from openpyxl.styles.fills import PatternFill, FILL_SOLID
from openpyxl.styles.borders import Border, Side, BORDER_THIN

def run():
    # Database
    # - เชื่อมต่อ Database (เปลี่ยนค่า Connection เป็นของเครื่องตัวเองเน่อ)
    db = mysql.connector.connect(
        host="localhost",
        port=3306,
        user="root",
        password="password1234",
        database='golf_want_to_buy'
    )
    cursor = db.cursor()

    # - โหลดข้อมูลสินค้าที่รวมประเภทสินค้ามาด้วย
    sql = '''
        SELECT p.id AS id, p.title AS title, p.price AS price, c.title AS category
        FROM products AS p
        LEFT JOIN categories AS c
        ON p.category_id = c.id
    '''
    cursor.execute(sql)
    products = cursor.fetchall()

    # - โหลดข้อมูลประเภทสินค้า พร้อมนับจำนวนสินค้าด้วย
    sql = '''
        SELECT c.id AS id, c.title AS title, COUNT(p.id) AS products_count
        FROM categories AS c
        LEFT JOIN products AS p
        ON c.id = p.category_id
        GROUP BY c.id
    '''
    cursor.execute(sql)
    categories = cursor.fetchall()

    # Excel
    # - การตกแต่งทั้งหมด
    border = Border(
        top=Side(border_style=BORDER_THIN, color='333333'),
        right=Side(border_style=BORDER_THIN, color='333333'),
        bottom=Side(border_style=BORDER_THIN, color='333333'),
        left=Side(border_style=BORDER_THIN, color='333333')
    )

    header_style = NamedStyle(name='header')
    header_style.fill = PatternFill(fill_type=FILL_SOLID, fgColor='EEEEEE')
    header_style.border = border
    header_style.font = Font(name='Helvetica Neue', bold=True, size=16)
    header_style.alignment = Alignment(vertical='center')

    row_style = NamedStyle(name='row')
    row_style.border = border
    row_style.font = Font(name='Helvetica Neue', size=16)
    row_style.alignment = Alignment(vertical='center')

    # - สร้างไฟล์ใหม่
    workbook = Workbook()

    # - สร้างชีทสำหรับสินค้า
    products_sheet = workbook.active
    products_sheet.title = 'สินค้า'
    products_sheet.append(['ID', 'ชื่อสินค้า', 'ราคา', 'ประเภทสินค้า'])

    # - ตกแต่งโดยรวม
    products_sheet.row_dimensions[1].height = 32
    products_sheet.column_dimensions['A'].width = 10
    products_sheet.column_dimensions['B'].width = 30
    products_sheet.column_dimensions['C'].width = 18
    products_sheet.column_dimensions['D'].width = 30

    # - ตกแต่งแถวแรก (Header)
    product_header_row = products_sheet[1]
    for cell in product_header_row:
        cell.style = header_style

    # - ใส่ข้อมูลสินค้าทีละอัน เพิ่มลงไปทีละแถว
    for p in products:
        print(p)
        products_sheet.append(p)

    # - ตกแต่งแถวข้อมูล (Row)
    for (number, row) in enumerate(products_sheet.iter_rows(min_row=2), 2):
        products_sheet.row_dimensions[number].height = 32
        for cell in row:
            cell.style = row_style

    # - สร้างชีทสำหรับประเภทสินค้า
    categories_sheet = workbook.create_sheet('ประเภทสินค้า', 1)
    categories_sheet.append(['ID', 'ประเภทสินค้า', 'จำนวนสินค้า'])

    # - ตกแต่งโดยรวม
    categories_sheet.row_dimensions[1].height = 32
    categories_sheet.column_dimensions['A'].width = 10
    categories_sheet.column_dimensions['B'].width = 30
    categories_sheet.column_dimensions['C'].width = 20

    # - ตกแต่งแถวแรก (Header)
    categories_header_row = categories_sheet[1]
    for cell in categories_header_row:
        cell.style = header_style

    # - ใส่ข้อมูลประเภทสินค้าทีละอัน เพิ่มลงไปทีละแถว
    for c in categories:
        print(c)
        categories_sheet.append(c)

    # - ตกแต่งแถวข้อมูล (Row)
    for (number, row) in enumerate(categories_sheet.iter_rows(min_row=2), 2):
        categories_sheet.row_dimensions[number].height = 32
        for cell in row:
            cell.style = row_style

    # - Export ไฟล์ Excel
    workbook.save(filename="./files/exported_10.xlsx")

    # ปิดการเชื่อมต่อ Database
    cursor.close()
    db.close()
