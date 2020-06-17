# 导入xlrd库
import xlrd

# 打开刚才写入的test_w.xls文件
wb = xlrd.open_workbook("../write_excel/test_xlwt_write_excel.xls")

# 获取并打印sheet数量
print("sheet 数量：", wb.nsheets)

# 获取并打印 sheet 名称
print("sheet 名称:", wb.sheet_names())

# 根据 sheet 索引获取内容
sh1 = wb.sheet_by_index(0)
# 或者
# 也可根据 sheet 名称获取内容
# sh = wb.sheet_by_name('成绩')

# 获取并打印该 sheet 行数和列数
print("sheet %s 共 %d 行 %d 列" % (sh1.name, sh1.nrows, sh1.ncols))

# 获取并打印某个单元格的值
print("第一行第二列的值为:", sh1.cell_value(0, 1))

# 获取整行或整列的值
rows = sh1.row_values(0) # 获取第一行内容
cols = sh1.col_values(1) # 获取第二列内容

# 打印获取的行列值
print("第一行的值为:", rows)
print("第二列的值为:", cols)

# 获取单元格内容的数据类型
print("第二行第一列的值类型为:", sh1.cell(1, 0).ctype)

# 遍历所有表格内容
print("遍历所有表格内容：")
for sh in wb.sheets():
    print("开始输出sheet：" + sh.name)
    for r in range(sh.nrows):
        # 输出指定行
        print(sh.row(r))

"""
执行结果：
------------------------------
sheet 数量： 2
sheet 名称: ['成绩', '汇总']
sheet 成绩 共 3 行 2 列
第一行第二列的值为: 成绩
第一行的值为: ['姓名', '成绩']
第二列的值为: ['成绩', 88.0, 99.5]
第二行第一列的值类型为: 1
遍历所有表格内容：
开始输出sheet：成绩
[text:'姓名', text:'成绩']
[text:'张三', number:88.0]
[text:'李四', number:99.5]
开始输出sheet：汇总
[text:'总分']
[number:187.5]
"""

"""
细心的朋友可能注意到，这里我们可以获取到单元格的类型，
上面我们读取类型时获取的是数字1，那1表示什么类型，又都有什么类型呢？别急下面我们通过一个表格展示下：
----------------
数值	类型	说明
0	empty	空
1	string	字符串
2	number	数字
3	date	日期
4	boolean	布尔值
5	error	错误
-----------------
通过上面表格，我们可以知道刚获取单元格类型返回的数字1对应的就是字符串类型。
"""