# Import ข้อมูลจากไฟล์ Excel (.xlsx) เข้าสู่ Database MySQL
# เป็นการ Import ข้อมูลทุกแถวมาใส่ใน products
# แต่จะเอาข้อมูลแฮชแท็ก แยกไปใส่ใน hashtags (ถ้ายังไม่มี)
# และที่สำคัญ ต้องเชื่อม products และ hashtags เข้ากันให้เรียบร้อย

import mysql.connector
from openpyxl import load_workbook

def run():
    # Excel
    # - โหลดไฟล์ และโหลดชีทที่เปิดอยู่
    workbook = load_workbook('./files/imported_06.xlsx')
    sheet = workbook.active

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

    # - โหลดข้อมูลแฮชแท็กทั้งหมด
    sql = '''
        SELECT * 
        FROM hashtags;
    '''
    cursor.execute(sql)
    hashtags = cursor.fetchall()

    # - เอาแฮชแท็กใหม่ มาใส่ใน List
    hashtags_values = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        hashtags_all_string = row[3]
        hashtags_list = hashtags_all_string.split(' ')

        for hashtag_excel in hashtags_list:
            is_new = True
            for hashtag_table in hashtags:
                if hashtag_excel == hashtag_table[1]:
                    is_new = False
                    break
            for hashtag_value in hashtags_values:
                if hashtag_excel == hashtag_value[0]:
                    is_new = False
                    break

            if is_new:
                print((hashtag_excel,))
                hashtags_values.append((hashtag_excel,))

    # - เพิ่มข้อมูลแฮชแท็ก (ถ้ามีอันใหม่)
    if len(hashtags_values) > 0:
        sql = '''
            INSERT INTO hashtags (title)
            VALUES (%s);
        '''
        cursor.executemany(sql, hashtags_values)
        db.commit()
        print('เพิ่มแฮชแท็กจำนวน ' + str(cursor.rowcount) + ' แถว')
    else:
        print('ไม่มีแฮชแท็กใหม่มาเพิ่ม')

    # - ใส่สินค้าทั้งหมดลง List
    products_values = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        product = (row[0], row[1], row[2])
        print(product)
        products_values.append(product)

    # - เพิ่มข้อมูลสินค้า
    sql = '''
        INSERT INTO products (title, price, is_necessary)
        VALUES (%s, %s, %s);
    '''
    cursor.executemany(sql, products_values)
    db.commit()
    print('เพิ่มสินค้าจำนวน ' + str(cursor.rowcount) + ' แถว')

    # - ดึง ID ของสินค้าแรกที่เพิ่มไปเมื่อกี๊ มาเตรียมตัวใช้
    first_product_id = cursor.lastrowid

    # - โหลดข้อมูลแฮชแท็กทั้งหมด (หลังบันทึกล่าสุดไปแล้ว)
    sql = '''
        SELECT *
        FROM hashtags;
    '''
    cursor.execute(sql)
    hashtags = cursor.fetchall()

    # - เชื่อมสินค้ากับแฮชแท็กเข้าด้วยกัน
    products_hashtags_values = []
    for (p_id, row) in enumerate(sheet.iter_rows(min_row=2, values_only=True), first_product_id):
        hashtags_all_string = row[3]
        hashtags_list = hashtags_all_string.split(' ')

        for hashtag_excel in hashtags_list:
            hashtag_id = None
            for hashtag_table in hashtags:
                if hashtag_excel == hashtag_table[1]:
                    hashtag_id = hashtag_table[0]
                    break
            print((p_id, hashtag_id))
            products_hashtags_values.append((p_id, hashtag_id))

    # - ใส่ข้อมูลลงในตารางเชื่อมโยง products_hashtags
    sql = '''
        INSERT INTO products_hashtags (product_id, hashtag_id)
        VALUES (%s, %s);
    '''
    cursor.executemany(sql, products_hashtags_values)
    db.commit()
    print('เพิ่มการเชื่อมโยงจำนวน ' + str(cursor.rowcount) + ' แถว')

    # ปิดการเชื่อมต่อ Database
    cursor.close()
    db.close()
