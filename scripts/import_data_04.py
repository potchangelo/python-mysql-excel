# Import ข้อมูลจากไฟล์ Excel (.xlsx) เข้าสู่ Database MySQL
# เป็นการ Import ข้อมูลทุกแถวมาใส่ใน products
# แต่จะเอาข้อมูลประเภทสินค้า แยกไปใส่ใน categories (ถ้ายังไม่มี)
# และที่สำคัญ ต้องเชื่อม products และ categories เข้ากันให้เรียบร้อย

import mysql.connector
from openpyxl import load_workbook

def run():
    # Database
    # - เชื่อมต่อ Database (เปลี่ยนค่า Connection เป็นของเครื่องตัวเองเน่อ)
    db = mysql.connector.connect(
        host="localhost",
        port=3306,
        user="root",
        password="password1234",
        database='golf_want_to_buy_completed'
    )
    cursor = db.cursor()

    # Database
    # - โหลดข้อมูล categories
    sql = '''
        SELECT * 
        FROM categories;
    '''
    cursor.execute(sql)
    categories = cursor.fetchall()

    # Excel
    # - โหลดไฟล์ และโหลดชีทที่เปิดอยู่
    workbook = load_workbook('./files/imported_04.xlsx')
    sheet = workbook.active

    # - เอา Category ใหม่มาใส่ใน List
    categories_values = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        is_new = True
        category = row[3]

        for c in categories:
            if category == c[1]:
                is_new = False
                break

        if is_new:
            print((category,))
            categories_values.append((category,))

    # Database
    # - เพิ่มข้อมูลประเภทสินค้า (ถ้ามีอันใหม่)
    if len(categories_values) > 0:
        sql = '''
            INSERT INTO categories (title)
            VALUES (%s);
        '''
        cursor.executemany(sql, categories_values)
        db.commit()
        print('เพิ่มประเภทสินค้าจำนวน ' + str(cursor.rowcount) + ' แถว')
    else:
        print('ไม่มีประเภทสินค้าใหม่มาเพิ่ม')

    # Database
    # - โหลดข้อมูลประเภทสินค้า หลังบันทึกล่าสุดไปแล้ว
    sql = '''
        SELECT * 
        FROM categories;
    '''
    cursor.execute(sql)
    categories = cursor.fetchall()

    # - ใส่ category_id ให้สินค้าแต่ละอย่าง
    products_values = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        category_title = row[3]
        category_id = 'NULL'

        for c in categories:
            if category_title == c[1]:
                category_id = c[0]
                break

        product = (row[0], row[1], row[2], category_id)
        print(product)
        products_values.append(product)

    # Database
    # - เพิ่มข้อมูลสินค้า
    cursor = db.cursor()
    sql = '''
        INSERT INTO products (title, price, is_necessary, category_id)
        VALUES (%s, %s, %s, %s);
    '''
    cursor.executemany(sql, products_values)
    db.commit()
    print('เพิ่มสินค้าจำนวน ' + str(cursor.rowcount) + ' แถว')

    # ปิดการเชื่อมต่อ Database
    cursor.close()
    db.close()
