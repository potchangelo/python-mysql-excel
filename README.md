# Python x MySQL x Excel

ตัวอย่างโปรเจ็คอ่าน/บันทึกข้อมูลระหว่าง Database MySQL กับไฟล์ Excel จากคลิปสอน MySQL เบื้องต้น Ep.1-2 ของ Zinglecode

## YouTube videos

- [สอน MySQL เบื้องต้น #01](https://www.youtube.com/watch?v=axraNvtHjO4)
- [สอน MySQL เบื้องต้น #02](https://www.youtube.com/watch?v=xXDR9rxVfA8)

## Setup database table

![products table structure](https://raw.githubusercontent.com/potchangelo/python-mysql-excel/dev/snapshots/yt-2-db-table-products-structure.jpg "products table structure")
*products table structure*

![product_notes table structure](https://raw.githubusercontent.com/potchangelo/python-mysql-excel/dev/snapshots/yt-2-db-table-product-notes-structure.jpg "product_notes table structure")
*product_notes table structure*

![categories table structure](https://raw.githubusercontent.com/potchangelo/python-mysql-excel/dev/snapshots/yt-2-db-table-categories-structure.jpg "categories table structure")
*categories table structure*

![hashtags table structure](https://raw.githubusercontent.com/potchangelo/python-mysql-excel/dev/snapshots/yt-2-db-table-hashtags-structure.jpg "hashtags table structure")
*hashtags table structure*

![products_hashtags table structure](https://raw.githubusercontent.com/potchangelo/python-mysql-excel/dev/snapshots/yt-2-db-table-products-hashtags-structure.jpg "products_hashtags table structure")
*products_hashtags table structure*

## Install and run project by PyCharm

0. ติดตั้ง MySQL, MySQL Workbench, Python 3, Pipenv, และ PyCharm ลงเครื่องให้เรียบร้อยก่อน

1. ดาวน์โหลดโปรเจ็คนี้ลงเครื่อง

2. เปิดโฟลเดอร์โปรเจ็คใน PyCharm โดยเลือกที่เมนู File -> Open... -> และเลือกโฟลเดอร์

3. แก้ไข Warning อะไรก็ตามที่ขึ้นมาใน PyCharm ให้เรียบร้อย

4. ติดตั้ง Packages

```
pipenv install
```

5. เปิดไฟล์ main.py และคลิกขวาที่พื้นที่เขียนโค้ด แล้วเลือก Run 'main'

6. โปรแกรมจะให้ระบุชื่อไฟล์ที่ต้องการรัน ดูชื่อไฟล์ได้จากโฟลเดอร์ scripts (ใส่ไปแบบไม่ต้องเติม .py)

## Install and run project by CLI

0. ติดตั้ง MySQL, MySQL Workbench, Python 3, และ Pipenv ลงเครื่องให้เรียบร้อยก่อน

1. ดาวน์โหลดโปรเจ็คนี้ลงเครื่อง

2. เปิด Terminal หรือ Command Prompt หรือ PowerShell ที่โฟลเดอร์โปรเจ็ค

3. ติดตั้ง Packages

```
pipenv install
```

4. Activate pipenv environment

```
pipenv shell
```

5. รันไฟล์ main.py

```
python main.py
```

6. โปรแกรมจะให้ระบุชื่อไฟล์ที่ต้องการรัน ดูชื่อไฟล์ได้จากโฟลเดอร์ scripts (ใส่ไปแบบไม่ต้องเติม .py)
