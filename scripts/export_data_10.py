# Export ข้อมูลจาก Database MySQL ออกมาเป็นไฟล์ Excel (.xlsx)
# เป็นการ Export ข้อมูลสินค้าทุกแถว รวมประเภทสินค้าด้วย มาอยู่ในชีทแรก
# และ Export ข้อมูลประเภทสินค้าทุกแถว พร้อมนับจำนวนสินค้า มาอยู่ในชีทที่สอง

import mysql.connector
from openpyxl import Workbook

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
    sql_select_products = '''
        SELECT p.id AS id, p.title AS title, p.price AS price, c.title AS category
        FROM products AS p
        LEFT JOIN categories AS c
        ON p.category_id = c.id;
    '''
    cursor.execute(sql_select_products)
    products = cursor.fetchall()

    # - โหลดข้อมูลประเภทสินค้า พร้อมนับจำนวนสินค้าด้วย
    sql_select_categories = '''
        SELECT c.id AS id, c.title AS title, COUNT(p.id) AS products_count
        FROM categories AS c
        LEFT JOIN products AS p
        ON c.id = p.category_id
        GROUP BY c.id;
    '''
    cursor.execute(sql_select_categories)
    categories = cursor.fetchall()

    # Excel
    # - สร้างไฟล์ใหม่
    workbook = Workbook()

    # - สร้างชีทสำหรับสินค้า
    products_sheet = workbook.active
    products_sheet.title = 'สินค้า'
    products_sheet.append(['ID', 'ชื่อสินค้า', 'ราคา', 'ประเภทสินค้า'])

    # - ใส่ข้อมูลสินค้าทีละอัน เพิ่มลงไปทีละแถว
    for p in products:
        print(p)
        products_sheet.append(p)

    # - สร้างชีทสำหรับประเภทสินค้า
    categories_sheet = workbook.create_sheet('ประเภทสินค้า', 1)
    categories_sheet.append(['ID', 'ประเภทสินค้า', 'จำนวนสินค้า'])

    # - ใส่ข้อมูลประเภทสินค้าทีละอัน เพิ่มลงไปทีละแถว
    for c in categories:
        print(c)
        categories_sheet.append(c)

    # - Export ไฟล์ Excel
    workbook.save(filename="./files/exported_10.xlsx")

    # ปิดการเชื่อมต่อ Database
    cursor.close()
    db.close()
