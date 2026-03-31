from sqlite3 import connect
from pandas import read_excel, DataFrame
from xlwings import App
from os import walk, path, system, _exit, _exists, makedirs
import re

#数据库数据准备，需要另外连接 
def prepare_sql(conn):
    #try:
    #读取data文件夹下准备好的文件
    df_PE = read_excel("data/PE.xlsx")
    df_dom = read_excel("data/dom.xlsx")
    df_members = read_excel("data/members.xlsx")
    df_names = read_excel("data/names.xlsx")

    #新增PE表，包含待查人员的体育成绩
    df_PE_1 = df_PE[["学号","姓名","总分","等级"]]
    df_PE_1.columns = ["id","name","score_PE","level_PE"]
    df_PE_1.to_sql("PE", conn, if_exists="replace", index=False)

    #新增dom表，包含待查人员的寝室成绩
    df_dom_1 = df_dom[["学号","姓名","寝室名称","学院","平均分","寝室等级"]]
    df_dom_1.columns = ["id","name","dormitory","dom_name","score_dom","hygiene_level"]
    df_dom_1.to_sql("dom", conn, if_exists="replace", index=False)
    
    #新增members表，包含待查人员的班干部任职情况（此项可能需要扩充）
    df_members_1 = df_members[["学号","学生姓名","岗位名称","考核结果"]]
    df_members_1.columns = ["id","name","position","assessment_result"]
    df_members_1.to_sql("members", conn, if_exists="replace", index=False)

    # 新增 names 表，包含待检查人员名单
    df_names_1 = df_names[["学号","姓名"]]
    df_names_1.columns = ["id","name"]
    df_names_1["checked"] = 0
    df_names_1.to_sql("names", conn, if_exists="replace", index=False)

    conn.commit()
    print("数据库数据准备完成")
"""
    except Exception as e:
        print(str(e))"""
    
#读取单个文件的个人信息
def get_personal_info(app, file_path): #file_path为单个文件的路径，遍历放在外层程序
    wb = app.books.open(file_path)
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
        filename = path.basename(file_path).replace('.xls', '').replace('.xlsx', '')
        parts = re.split(r'[-—]', filename)
        if len(parts) >= 3 and parts[0] == "学年纪实测评表":
            number = parts[1]
            name = parts[2]
        else:
            number = None
            name = None
    PE_score= sheet.range("D36").value #体育成绩
    dom_score= sheet.range("D38").value #寝室分数，需要结合寝室成绩和是否寝室长
    is_level = sheet.range("D39").value #优秀寝室加分
    position_score = sheet.range("D32").value #班干部考核分数
    wb.close()
    return [number, name, PE_score, dom_score, is_level, position_score]

#存储并被调用分数计算规则
def score(PE_raw, dom_raw, is_level_raw, position_raw, position_level_raw ):
    PE_score_cal = dom_score_cal = is_level_cal = position_score_cal = 0
    #体育分数计算
    if PE_raw >= 90:
        PE_score_cal = 5
    elif PE_raw >= 80:
        PE_score_cal = 4
    elif PE_raw >= 70:
        PE_score_cal = 3
    elif PE_raw >= 60:
        PE_score_cal = 2
    else:
        PE_score_cal = 0
    #寝室分数计算
    if dom_raw >= 100:
        dom_score_cal = 5
    elif dom_raw >= 95:
        dom_score_cal = 4
    elif dom_raw >= 85:
        dom_score_cal = 3
    elif dom_raw >= 75:
        dom_score_cal = 2
    else:
        dom_score_cal = 0
    #优秀寝室加分
    if is_level_raw == "模范寝室":
        is_level_cal = 2
    elif is_level_raw == "文明寝室":
        is_level_cal = 1
    else:
        is_level_cal = 0
    #班干部考核分数
    position_data = {"班长":{"优秀":4,"合格":3},
                     "团支书":{"优秀":4,"合格":3},
                     "学习委员":{"优秀":3,"合格":2},
                     "生活委员":{"优秀":3,"合格":2},
                     "组宣委员":{"优秀":3,"合格":2},
                     "其他班级委员":{"优秀":3,"合格":2},
                     "寝室长":{"优秀":2,"合格":1}}
    if position_raw in position_data:
        position_score_cal = position_data[position_raw][position_level_raw]
    return PE_score_cal, dom_score_cal, is_level_cal, position_score_cal

#判断获取到的数据是否符合规则
def check(cur,personal_info,):
    number = personal_info[0]
    name = personal_info[1]
    PE_score = int(personal_info[2]) if personal_info[2] else 0
    dom_score = int(personal_info[3]) if personal_info[3] else 0
    is_level = int(personal_info[4]) if personal_info[4] else 0
    position_score = int(personal_info[5]) if personal_info[5] else 0
    if not number or not name:
        df_log.loc[len(df_log)] = [name,"无名氏",  "", "",file_path]
        print(f"未找到学号或姓名，跳过该文件")
        return

    # 标记 names 表，该人员已被检查到
    cur.execute("UPDATE names SET checked = 1 WHERE id = ? AND name = ?", (number, name))
    
    if PE_score == 0 and dom_score == 0 and is_level == 0 and position_score == 0:
        print(f"{name} 的待查项均为0，跳过该文件")
        return
    
    data_get = []
    #检查体育成绩
    cur.execute("SELECT * FROM PE WHERE id = ? AND name = ?", (number, name))
    result = cur.fetchone()
    if not result:
        df_log.loc[len(df_log)] = [name,"未找到体育成绩", PE_score, "", file_path]
        data_get.append("") #体育成绩
    else:
        data_get.append(result[2]) if result[2] else  data_get.append("")
    #检查寝室分数
    cur.execute("SELECT * FROM dom WHERE id = ? AND name = ?", (number, name))
    result = cur.fetchone()
    if not result:
        df_log.loc[len(df_log)] = [name,"未找到寝室分数", dom_score, "", file_path]
        data_get.append("") #寝室分数
        data_get.append("") #寝室等级
    else:
        data_get.append(result[4]) if result[4] else data_get.append("")
        data_get.append(result[5]) if result[5] else data_get.append("")

    #检查班干部考核
    cur.execute("SELECT * FROM members WHERE id = ? AND name = ?", (number, name))
    result = cur.fetchone()
    if not result:
        df_log.loc[len(df_log)] = [name,"未找到班干部考核", position_score, "", file_path]
        data_get.append("") #职位
        data_get.append("") #考核结果
    else:
        data_get.append(result[2]) if result[2] else data_get.append("")
        data_get.append(result[3]) if result[3] else data_get.append("")

    PE_score_cal, dom_score_cal, is_level_cal, position_score_cal = score(data_get[0], data_get[1], data_get[2], data_get[3], data_get[4])
    testflag = True
    if PE_score_cal != PE_score:
        df_log.loc[len(df_log)] = [name,"体育成绩", PE_score, PE_score_cal, file_path]
        testflag = False
    if dom_score_cal != dom_score:
        df_log.loc[len(df_log)] = [name,"寝室分数", dom_score, dom_score_cal, file_path]
        testflag = False   
    if is_level_cal != is_level:
        df_log.loc[len(df_log)] = [name,"寝室加分", is_level, is_level_cal, file_path]
        testflag = False
    if position_score_cal != position_score:
        df_log.loc[len(df_log)] = [name,"班干部考核", position_score, position_score_cal, file_path]
        testflag = False
    
    if testflag:
        print(f"{name} 检查通过")

#遍历文件夹得到个人综测分xls，读取信息并输出
def walk_main(cur, app):
    global file_path
    global cnt
    cnt = 0
    for root, dirs, files in walk("test_data"):
        for file in files:
            if (file.endswith('.xls')or file.endswith('.xlsx')) and "学年纪实测评表" in file:
                file_path = path.join(root, file)
                personal_info = get_personal_info(app, file_path)
                check(cur, personal_info)
                cnt += 1

# 检查data文件夹及文件是否存在
def check_data_files():
    data_dir_1 = "data" #被调用数据文件夹
    data_dir_2 = "test_data" #待测试数据文件夹
    required_files = ["PE.xlsx", "dom.xlsx", "members.xlsx", "names.xlsx"]
    flag= True
    if not path.exists(data_dir_1) and not path.exists("data.db"): #data文件夹和数据库文件都不存在
        makedirs(data_dir_1)
        print(f"已创建缺失的文件夹: {data_dir_1}")
        flag= False
    if not path.exists(data_dir_2):
        makedirs(data_dir_2)
        print(f"已创建缺失的文件夹: {data_dir_2}")
        flag= False
    print(f"请确认数据文件放入 {data_dir_1} 文件夹，待检查文件放入 {data_dir_2} 文件夹")
    for f in required_files:
        if not path.exists(path.join(data_dir_1, f)):
            print(f"缺少文件: {f}")
            flag= False
    return flag

def main():
    global df_log

    while check_data_files() == False:
        input("请将缺失的文件准备好后，按回车键继续...")

    conn = connect('data.db')
    cur = conn.cursor()
    df_log = DataFrame(columns=["姓名","不匹配项","表格中分数","数据库计算分数","文件路径"])
    """
    if path.getsize("data.db") > 0:
        pass
    else:
    """
    prepare_sql(conn)

    app = App(visible=False, add_book=False)
    try:
        walk_main(cur, app)
    finally:
        # 保证程序结束时只关闭这一主进程，避免残留 workbook 进程
        app.quit()

    # 检查 names 表是否所有目标都被扫到
    cur.execute("SELECT id,name FROM names WHERE checked = 0")
    missed = cur.fetchall()
    if not missed:
        print("所有人员均已检查到。")
    else:
        missed_names = ", ".join([f"{r[1]}({r[0]})" for r in missed])
        print(f"以下人员未被检查到：{missed_names}")
        df_log.loc[len(df_log)] = ["[未检查到]", "名单缺失", "", missed_names, ""]

    df_log.to_excel("报错汇总.xlsx", index=False)
    print("检查完成，报错汇总已保存至 报错汇总.xlsx")
    print(f"共检查 {cnt} 个文件")
    system("pause")

    cur.close() 
    conn.close()

if __name__ == "__main__":
    main()