import mysql.connector
from openpyxl import Workbook

#Database
db = mysql.connector.connect(
    host='localhost',
    port=3306,
    user='root',
    password='Title@20572',
    database='title_database'

)

cursor = db.cursor()
sql = '''
    select p.id as id , p.title as title , p.price as price , c.title as catagories 
    from products as p
    left join catagories as c
    on p.catagory_id = c.id

'''

cursor.execute(sql)
product = cursor.fetchall()

#excel

workbook = Workbook()
sheet = workbook.active
sheet.append(['id','ชื่อสินค้า','ราคาสินค้า','ประเภทสินค้า'])

for p in product:
    print(p)
    sheet.append(p)


workbook.save(filename='exex.xlsx')

cursor.close()
db.close()