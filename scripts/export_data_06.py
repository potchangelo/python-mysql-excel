# Export ข้อมูลจาก Database MySQL ออกมาเป็นไฟล์ Excel (.xlsx)
# เป็นการ Export ข้อมูลสินค้าทุกแถว
# และรวมแฮชแท็กทั้งหมดของสินค้าแต่ละอัน มาแสดงด้วย

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

    # - ส่งคำสั่ง SQL ไปให้ MySQL ทำการโหลดข้อมูล
    # - Python จะรับข้อมูลทั้งหมดมาเป็น List ผ่านคำสั่ง fetchall()
    cursor = db.cursor()
    sql = '''
        SELECT p.id AS id, p.title AS title, p.price AS price, GROUP_CONCAT(ph.hashtag SEPARATOR ' ') AS hashtags
        FROM products AS p
        LEFT JOIN (
            SELECT ph1.product_id AS product_id, h1.hashtag AS hashtag
            FROM products_hashtags AS ph1
            LEFT JOIN hashtags AS h1
            ON ph1.hashtag_id = h1.id
        ) AS ph
        ON p.id = ph.product_id
        GROUP BY p.id;
    '''
    cursor.execute(sql)
    products = cursor.fetchall()

    # Excel
    # - สร้างไฟล์ใหม่ สร้างชีท และใส่แถวสำหรับเป็นหัวข้อตาราง
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(['ID', 'ชื่อสินค้า', 'ราคา', 'แฮชแท็ก'])

    # - ใส่ข้อมูลทีละอัน เพิ่มลงไปทีละแถว
    for p in products:
        print(p)
        sheet.append(p)

    # - Export ไฟล์ Excel
    workbook.save(filename="./files/exported_06.xlsx")

    # ปิดการเชื่อมต่อ Database
    cursor.close()
    db.close()
