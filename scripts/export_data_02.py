# Export ข้อมูลจาก Database MySQL ออกมาเป็นไฟล์ Excel (.xlsx)
# เป็นการ Export ข้อมูลทุกแถว แต่เอาเฉพาะข้อมูลชื่อสินค้า, ราคาสินค้า

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
        SELECT title, price 
        FROM products;
    '''
    cursor.execute(sql)
    products = cursor.fetchall()

    # Excel
    # - สร้างไฟล์ใหม่ สร้างชีท และใส่แถวสำหรับเป็นหัวข้อตาราง
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(['ชื่อสินค้า', 'ราคา'])

    # - ใส่ข้อมูลทีละอัน เพิ่มลงไปทีละแถว
    for p in products:
        print(p)
        sheet.append(p)

    # - Export ไฟล์ Excel
    workbook.save(filename="./files/exported_02.xlsx")