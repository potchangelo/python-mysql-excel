import os
from importlib import import_module

def main():
    print('พิมพ์ชื่อไฟล์ในโฟลเดอร์ "scripts" เพื่อสั่งการทำงาน (ตัวอย่าง : "import_data_01")')
    file_name = input('ชื่อไฟล์ : ')
    files = [f.replace('.py', '') for f in os.listdir(os.curdir + '/scripts') if 'import_' in f or 'export_' in f]

    if file_name not in files:
        print('ไม่พบไฟล์ที่ระบุ')
        return

    print('--- "' + file_name + '" กำลังทำงาน ---')
    module = import_module('scripts.' + file_name)
    module.run()
    print('--- "' + file_name + '" ทำงานเสร็จสิ้น ---')

if __name__ == '__main__':
    main()
