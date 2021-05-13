import mysql.connector
from openpyxl import Workbook

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
    SELECT * 
    FROM products;
'''
cursor.execute(sql)
products = cursor.fetchall()

# Excel
workbook = Workbook()
sheet = workbook.active
sheet.append(['ID', 'ชื่อสินค้า', 'ราคา', 'ต้องการมากๆ', 'วันที่บันทึก'])

for p in products:
    print(p)
    sheet.append(p)

workbook.save(filename="exported.xlsx")