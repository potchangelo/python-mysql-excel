# Export ข้อมูลทุกแถวจาก Database MySQL ออกมาเป็นไฟล์ Excel (.xlsx)

import mysql.connector
from openpyxl import Workbook

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

# - โหลดข้อมูลสินค้า พร้อมประเภทสินค้า
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
sheet.append(['ID', 'ชื่อสินค้า', 'ราคาสินค้า', 'ประเภทสินค้า'])

# - ใส่ข้อมูลทีละอัน เพิ่มลงไปทีละแถว
for p in products:
    print(p)
    sheet.append(p)

# - Export ไฟล์ Excel
workbook.save(filename="exported.xlsx")

# ปิดการเชื่้อมต่อ Database
cursor.close()
db.close()