# Export ข้อมูลจาก Database MySQL ออกมาเป็นไฟล์ Excel (.xlsx)
# เป็นการ Export ข้อมูลทุกแถว
# เอาข้อมูลชื่อสินค้า, ราคาสินค้า, วันที่บันทึก มาแสดง
# แสดงวันที่บันทึก ในรูปแบบ "25 November 2021"

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
        SELECT title, price, created_at 
        FROM products;
    '''
    cursor.execute(sql)
    products = cursor.fetchall()

    # Excel
    # - สร้างไฟล์ใหม่ สร้างชีท และใส่แถวสำหรับเป็นหัวข้อตาราง
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(['ชื่อสินค้า', 'ราคา', 'วันที่บันทึก'])

    # - ใส่ข้อมูลทีละอัน เพิ่มลงไปทีละแถว
    for p in products:
        p_real = list(p)
        p_real[2] = p_real[2].strftime('%d %B %Y')
        print(p_real)
        sheet.append(p_real)

    # - Export ไฟล์ Excel
    workbook.save(filename="./files/exported_03.xlsx")

    # ปิดการเชื่อมต่อ Database
    cursor.close()
    db.close()
