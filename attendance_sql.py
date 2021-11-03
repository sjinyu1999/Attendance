import sqlite3
import os
path=os.getcwd()
path1="\\attendance_sql.db"
fname=path+path1
conn = sqlite3.connect(fname)
cursor = conn.cursor()

# 添加信息
#sql_insert = '''insert into attendance_sheet
#                (name,date,sign_in,sign_out,mark,number,department)
#                values
#                ('"dd" ', '"2021-11-03" ', '"8:56" ', '"18:00" ', '" " ', '"555" ', '"芯片生态" ')'''

# 执行语句
#results = cursor.execute(sql_insert)
#print("插入数据成功！！！")
#conn.commit()

'''# 删除信息
sql_delete = 'delete from attendance_sheet where number=='"555"' '
# 执行语句
results = cursor.execute(sql_delete)
print("删除成功！！！")
conn.commit()'''

# 修改信息
#sql_update = '''update `attendance_sheet` set `name`='"ee"' where `name`=='"dd"' '''
# 执行语句
#results = cursor.execute(sql_update)
#print("修改成功！！！")
#conn.commit()

# 查询所有的信息
# sql_query = '''select * from attendance_sheet'''
''' 得到数据库中的信息'''
sql_query = '''
            select
            *
            from
            attendance_sheet
            where
            jobnumber=64
            '''
# 执行语句
results = cursor.execute(sql_query)

# 遍历打印输出
all_attendance_sheet = results.fetchall()
for name,date,sign_in,sign_out,mark,number,department in all_attendance_sheet:
    print(name+date+sign_in+sign_out+mark+number+department)
#print("查询数据成功！！！")
conn.commit()
conn.close()
