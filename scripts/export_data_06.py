# Export ข้อมูลจาก Database MySQL ออกมาเป็นไฟล์ Excel (.xlsx)
# เป็นการ Export ข้อมูลสินค้าทุกแถว
# และรวมประเภทสินค้าของแต่ละอัน มาแสดงด้วย

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

    # - โหลดข้อมูลสินค้าที่รวมแฮชแท็กทั้งหมดมาด้วย
    cursor = db.cursor()
    sql = '''
        SELECT p.id AS id, p.title AS title, p.price AS price, c.title AS category
        FROM products AS p
        LEFT JOIN categories AS c
        ON p.category_id = c.id;
    '''
    cursor.execute(sql)
    products = cursor.fetchall()

    # Excel
    # - สร้างไฟล์ใหม่ สร้างชีท และใส่แถวสำหรับเป็นหัวข้อตาราง
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(['ID', 'ชื่อสินค้า', 'ราคา', 'ประเภทสินค้า'])

    # - ใส่ข้อมูลทีละอัน เพิ่มลงไปทีละแถว
    for p in products:
        print(p)
        sheet.append(p)

    # - Export ไฟล์ Excel
    workbook.save(filename="./files/exported_06.xlsx")

    # ปิดการเชื่อมต่อ Database
    cursor.close()
    db.close()
