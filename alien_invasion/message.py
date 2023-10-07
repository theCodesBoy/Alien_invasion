import sqlite3

# 连接到 SQLite 数据库
conn = sqlite3.connect('messages.db')

# 创建游标对象
cursor = conn.cursor()

# 读取 SQL 文件并执行查询
try:
    with open('D:\Py-file\自定义案例\love_comment.sql', 'r',  encoding='utf-8') as sql_file:
        sql_queries = sql_file.read().split(',')
        for query in sql_queries:
            if query.strip():
                cursor.execute(query)
except UnicodeDecodeError as e:
    print(f"UnicodeDecodeError: {e}")
    # 在此处添加适当的处理代码，或跳过错误并继续处理文件

# 提交更改
conn.commit()

# 关闭连接
conn.close()


from openpyxl import Workbook

# 创建一个新的工作簿
workbook = Workbook()

# 创建一个工作表
sheet = workbook.active

# 执行 SQL 查询以获取结果
conn = sqlite3.connect('messages.db')
cursor = conn.cursor()
cursor.execute("SELECT * FROM love_comment")

# 将查询结果写入 XLSX 文件
for row_index, row in enumerate(cursor.fetchall()):
    for col_index, value in enumerate(row):
        sheet.cell(row=row_index+1, column=col_index+1, value=value)

# 保存 XLSX 文件
workbook.save('messages.xlsx')

# 关闭连接
conn.close()

