import mysql.connector
from openpyxl import load_workbook

# Excel
workbook = load_workbook('imported.xlsx')
sheet = workbook.active

values = []
for row in sheet.iter_rows(min_row=2, values_only=True):
    print(row)
    values.append(row)

# Database (เปลี่ยนค่า Connection เป็นของเครื่องตัวเองเน่อ)
db = mysql.connector.connect(
    host="localhost",
    port=3306,
    user="root",
    password="password1234",
    database='golf_want_to_buy'
)

cursor = db.cursor()
sql = '''
    INSERT INTO products (title, price, is_necessary)
    VALUES (%s, %s, %s);
'''
cursor.executemany(sql, values)
db.commit()
print('เพิ่ม ' + str(cursor.rowcount) + ' ข้อมูล')