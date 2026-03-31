from pandas import read_excel, DataFrame
from xlwings import App
from os import walk, path
from random import choice, shuffle
import re

# 遍历文件夹得到个人综测分 xls，读取信息并输出
def walk_main():
    global file_path
    global cnt
    global df_log
    app = App(visible=False, add_book=False)  # 创建一个 Excel 主进程
    df_log = DataFrame(columns=["学号", "姓名", "总分"])
    cnt = 0
    for root, dirs, files in walk("test_data"):
        for file in files:
            if (file.endswith('.xls') or file.endswith('.xlsx')) and "学年纪实测评表" in file:
                file_path = path.join(root, file)
                try:
                    get_personal_score(file_path, app)
                    cnt += 1
                except Exception as e:
                    print(f"处理文件 {file_path} 时出错: {e}")
    df_log.to_excel("个人综测分.xlsx", index=False)
    print("个人综测分已保存至 个人综测分.xlsx")
    print(f"共处理 {cnt} 个文件")
    app.quit()  # 最后关闭 Excel 主进程

# 获取个人分数并写入表格
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

    # 动态确定总分的终止行
    try:
        start_row = 5
        end_row = sheet.used_range.last_cell.row  # 默认终止行为表格最后一行
        for row in range(start_row, sheet.used_range.last_cell.row + 1):
            if sheet.range(f"A{row}").value == "其他":
                end_row = row - 1  # “其他”所在行的上一行为终止行
                break

        scorelist = sheet.range(f'E{start_row}:E{end_row}').value
        scorelist = [float(val) if val is not None else 0 for val in scorelist]
        score = sum(scorelist)
        sheet.range('I4').value = score
    except Exception as e:
        print(f"计算总分时出错: {e}")
        wb.close()
        return

    # 填写
    all_names = ["张三","李四", "王五", "赵六"]
    shuffle(all_names)  # 随机打乱顺序
    sheet.range('H4').value = all_names[cnt % len(all_names)]

    # 保存数据到汇总表
    df_log.loc[len(df_log)] = [number, name, score]
    print(f"{number} {name} {score}")

    wb.save()
    wb.close()

walk_main()
