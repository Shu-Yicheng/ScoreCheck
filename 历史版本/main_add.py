from pandas import read_excel, DataFrame
from xlwings import App
from os import walk, path, system, _exit, _exists
from random import choice
import re

#遍历文件夹得到个人综测分xls，读取信息并输出
def walk_main():
    global file_path
    global cnt
    global df_log
    app = App(visible=False, add_book=False)  # 创建一个Excel主进程
    df_log = DataFrame(columns=["学号","姓名","总分"])
    cnt = 0
    for root, dirs, files in walk("test_data"):
        for file in files:
            if (file.endswith('.xls') or file.endswith('.xlsx')) and "学年纪实测评表" in file:
                file_path = path.join(root, file)
                get_personal_score(file_path, app)
                cnt += 1
    df_log.to_excel("个人综测分.xlsx", index=False)
    print("个人综测分已保存至 个人综测分.xlsx")
    print(f"共处理 {cnt} 个文件")
    app.quit()  # 最后关闭Excel主进程

def do_write(wb, sheet, score):
    first_namelist = ["李四","王五","赵六"]
    second_name = "张三"
    random_name = choice(first_namelist)
    title_list = ["一校", "二校", "三校", "分数"]
    sheet.range('f3').expand('table').value =[title_list,[random_name,second_name,"",score]]
    wb.save()

def get_personal_score(filepath, app):
    wb = app.books.open(filepath)
    sheet = wb.sheets[0]
    # 尝试从表格获取姓名和学号
    try:
        name_val = sheet.range("A2").value
        if name_val and isinstance(name_val, str) and len(name_val) > 3:
            name = name_val[3:]
        else:
            name = None
    except:
        name = None
    try:
        number_val = sheet.range("C2").value
        if number_val and isinstance(number_val, str) and len(number_val) > 3:
            number = number_val[3:]
        else:
            number = None
    except:
        number = None
    # 如果表格中没有找到，从文件名提取
    if not name or not number:
        filename = path.basename(filepath).replace('.xls', '').replace('.xlsx', '')
        parts = re.split(r'[-—]', filename)
        if len(parts) >= 3 and parts[0] == "学年纪实测评表":
            number = parts[1]
            name = parts[2]
        else:
            number = None
            name = None
    if not name or not number:
        print(f"未找到学号或姓名，跳过该文件: {filepath}")
        wb.close()
        return
    scorelist = sheet.range('D5:D42').value
    for i in range(len(scorelist)):
        if scorelist[i] is None:
            scorelist[i] = 0
        else:
            scorelist[i] = float(scorelist[i])
    score = sum(scorelist) #总分
    df_log.loc[len(df_log)] = [number, name, score]
    #print(df_log)
    print(f"{number} {name} {score}")
    do_write(wb, sheet, score)
    wb.close()
    #app.quit()

walk_main()
