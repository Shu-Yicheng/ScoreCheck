# 检查并安装必要依赖
#引入内置库
from json import load, JSONDecodeError, dump
from os import walk, path, system, _exit, _exists, makedirs, listdir
from random import choice
from re import split
from sqlite3 import connect
#pandas和xlwings库需要安装，需要检查并自动补齐
try:
    from pandas import read_excel, DataFrame
    from xlwings import App
except ImportError:
    print("缺少必要运行库，正在尝试安装...")
    try:
        from subprocess import check_call
        from  sys import executable, exit
        mirror_source = "https://pypi.tuna.tsinghua.edu.cn/simple"  # 请填写镜像源，如 "https://pypi.tuna.tsinghua.edu.cn/simple"
        packages = ["pandas", "xlwings"]
        for pkg in packages:
            cmd = [executable, "-m", "pip", "install", pkg]
            if mirror_source:
                cmd.extend(["-i", mirror_source])
            check_call(cmd)
        print("依赖库已安装，请重新运行程序。")
        system("pause")
        exit(0)
    except Exception as e:
        print(f"安装依赖失败: {e}")
        system("pause")
        exit(1)

# 全局配置变量
g_config = None

# 加载并确认指定类型的签字配置
def load_and_confirm_signature_config(sig_type):
    """
    根据签字类型加载对应的配置项（带重试机制）
    sig_type: '本班' 或 '其他班'
    返回: 配置字典 或 None（用户放弃）
    """
    config_file = "config.json"
    while True: 
        try:
            with open(config_file, 'r', encoding='utf-8') as f:
                config = load(f)
            print(f"\n已加载配置文件: {config_file}")
            #从配置文件中获取签字字段
            sig_cfg = config.get('signature_config')
            if not sig_cfg:
                raise ValueError("配置文件中缺少 signature_config 字段")
            #依据签字类型获取相应字段
            if sig_type == '本班':
                config_item = sig_cfg.get('current_class')
                config_name = "本班审核署名"
            elif sig_type == '其他班':
                config_item = sig_cfg.get('other_class')
                config_name = "其他班审核署名"
            else:
                raise ValueError(f"未知的签字类型: {sig_type}")
            if not config_item:
                raise ValueError(f"配置文件中缺少 {sig_type} 的配置项")
            # 打印配置供用户确认
            print(f"【{config_name}配置】")
            if sig_type == '本班':
                print(f"  • 签字人选项: {config_item.get('reviewers', [])}")
                print(f"  • 副审核人: {config_item.get('secondary_reviewer', '')}")
                print(f"  • 表头位置: {config_item.get('header_position', 'f3').upper()}")
            else:  # 其他班
                print(f"  • 签字人选项: {config_item.get('reviewers', [])}")
                print(f"  • 签名位置: {config_item.get('name_cell', 'h4').upper()}")
                print(f"  • 分数位置: {config_item.get('score_cell', 'i4').upper()}")
            # 检查配置完整性
            if not config_item.get('reviewers'):
                print("一校人员列表为空，请修改配置文件")
                input("配置修改完毕后，按回车键继续...")
                continue
            if sig_type == '本班' and not config_item.get('secondary_reviewer'):
                print("二校人员配置为空，请修改配置文件")
                input("配置修改完毕后，按回车键继续...")
                continue
            #用户确认配置
            user_input = input("\n是否使用上述配置？(Y/N): ").strip().upper()
            if user_input != 'Y':
                print("请修改配置文件后重试")
                input("配置修改完毕后，按回车键继续...")
                continue
            return config_item
        #如果config文件不存在，直接生成空配置文件并提醒用户填写    
        except FileNotFoundError:
            print(f"\n✗ 配置文件 {config_file} 不存在")
            # 复用load_config函数生成空配置文件
            load_config()
            input(f"  请完成设置后按回车键继续...")
            continue
        except JSONDecodeError as e:
            print(f"\n✗ 配置文件格式错误: {e}")
            print(f"  请检查 {config_file} 的 JSON 格式")
            input("配置文件修复完毕后，按回车键继续...")
            continue
        except Exception as e:
            print(f"\n✗ 加载配置出错: {e}")
            input("配置文件修复完毕后，按回车键继续...")
            continue

# 旧的 load_config 函数保留用于第一次生成空配置文件
def load_config():
    """
    加载 config.json 配置文件
    如果文件不存在，生成空配置文件并提示用户
    """
    config_file = "config.json"
    #配置文件模板
    empty_config = {
        "signature_config": {
            "current_class": {
                "reviewers": [],
                "secondary_reviewer": "",
                "header_position": "f3"
            },
            "other_class": {
                "reviewers": [],
                "name_cell": "h4",
                "score_cell": "i4"
            }
        }
    }
    #加载后直接生成空配置文件
    print(f"  正在生成空配置文件...")
    with open(config_file, 'w', encoding='utf-8') as f:
        dump(empty_config, f, ensure_ascii=False, indent=2)
    print(f"  已生成 {config_file}")
    print(f"  请参阅程序说明文档填写相关信息后重新运行程序")
    return None

# 安全地解析单元格数值（支持浮点数和表达式）
def parse_score_value(value):
    """
    安全地将单元格值解析为浮点数
    支持：纯数字、浮点数、形如 "1+1=2" 的表达式（取等号前部分）
    """
    if value is None:
        return 0
    # 处理形如 "1+1=2" 的情况，提取等号前面的部分
    value_str = str(value).strip()
    if "=" in value_str:
        value_str = value_str.split("=")[0].strip()
    try:
        # 尝试直接转换为浮点数
        return float(value_str)
    except ValueError:
        # 尝试计算表达式（仅允许数字和基本操作符）
        try:
            # 检查是否包含危险内容（防止注入）
            dangerous_keywords = ['import', 'os', 'sys', 'exec', 'eval', '__', 'open', 'file']
            if any(keyword in value_str.lower() for keyword in dangerous_keywords):
                return 0
            # 只允许数字、小数点、括号和基本操作符
            allowed_chars = set('0123456789.+-*/(). \t')
            if not all(c in allowed_chars for c in value_str):
                return 0
            # 安全地计算表达式
            result = eval(value_str, {"__builtins__": {}}, {})
            return float(result)
        except:
            return 0

# 计算总分（支持动态行数确定）
def calculate_total_score(sheet, score_column='D', start_row=5):
    """
    计算表格中的总分
    score_column: 分数所在列（'D' 或 'E'）
    start_row: 开始行（默认为第5行）
    returns: 计算得到的总分（浮点数）
    """
    try:
        # 确定结束行（动态确定表格末尾）
        end_row = sheet.used_range.last_cell.row
        # 获取分数列的范围
        score_range = f'{score_column}{start_row}:{score_column}{end_row}'
        scorelist = sheet.range(score_range).value
        if scorelist is None:
            return 0
        # 展平列表（处理单个值的情况）
        if not isinstance(scorelist, list):
            scorelist = [scorelist]
        # 计算总分
        total = 0
        for score_value in scorelist:
            total += parse_score_value(score_value)
        return total
    except Exception as e:
        print(f"计算总分时出错: {e}")
        return 0

# 填写签字信息
def handle_signature(sheet, total_score, sig_type='', score_column='D'):
    """
    在表格中填写签字信息
    sig_type: 签字类型（'本班' 或 '其他班'）
    """
    global g_config
    try:
        # 从配置中获取对应签字类型的配置项
        if sig_type == '本班':
            config_item = g_config['signature_config'].get('本班') or g_config['signature_config'].get('current_class')
        else:  # 其他班
            config_item = g_config['signature_config'].get('其他班') or g_config['signature_config'].get('other_class')
        
        if not config_item:
            raise ValueError(f"未找到 {sig_type} 的配置项")
        
        if sig_type == '本班':
            # 本班审核署名配置
            reviewers = config_item.get('reviewers', [])
            if not reviewers:
                raise ValueError("本班一校人员列表为空")
            secondary_reviewer = config_item.get('secondary_reviewer', '')
            if not secondary_reviewer:
                raise ValueError("本班二校人员配置为空")
            #本班框定范围起始位置获取
            header_pos = config_item.get('header_position', 'f3').upper()
            # 从位置推导范围（如 F3 -> F3:I3）
            header_col = header_pos[0]
            header_row = header_pos[1:]
            end_col = chr(ord(header_col) + 3)  # 相隔3列
            header_range = f'{header_col}{header_row}:{end_col}{header_row}'
            # 数据位置硬编码为第二行第四列
            data_row = str(int(header_row) + 1)
            data_range = f'{header_col}{data_row}:{end_col}{data_row}'
            #获取随机名字
            random_name = choice(reviewers)
            # 填写表头
            sheet.range(header_range).value = ["一校", "二校", "三校", "分数"]
            # 填写数据
            sheet.range(data_range).value = [random_name, secondary_reviewer, "", total_score]
        elif sig_type == '其他班':
            # 其他班审核署名配置
            reviewers = config_item.get('reviewers', [])
            if not reviewers:
                raise ValueError("其他班签字人员列表为空")
            #定义签名和总分格子
            name_cell = config_item.get('name_cell', 'h4').upper()
            score_cell = config_item.get('score_cell', 'i4').upper()
            #随机选人并签名
            random_name = choice(reviewers)
            sheet.range(name_cell).value = random_name
            sheet.range(score_cell).value = total_score
    except Exception as e:
        print(f"✗ 填写签字信息时出错: {e}")

#数据库数据准备，需要另外连接 
def prepare_sql(conn):
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
    #提交保存
    conn.commit()
    print("数据库数据准备完成")
    
#读取单个文件的个人信息
def get_personal_info(app, file_path, enable_signature=False, sig_type='', score_column='D'): 
    #file_path为单个文件的路径，遍历放在外层程序
    wb = app.books.open(file_path)
    sheet = wb.sheets[0]
    # 尝试从表格A2和C2获取姓名和学号
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
    # 如果表格中没有找到，从文件名提取（这里需要文件名以“学年纪实测评表-学号-姓名”的规范命名，中间的-号起到分隔识别的作用，不能省略
    if not name or not number:
        filename = path.basename(file_path).replace('.xls', '').replace('x', '')
        parts = split(r'[-—]', filename)
        if len(parts) >= 3 and parts[0] == "学年纪实测评表":
            number = parts[1]
            name = parts[2]
        else:
            number = None
            name = None
    # 根据 B 列标签文本动态查找分数对应的行
    def get_score_by_label(label_text, start_row=3):
        """根据 B 列的标签文本查找对应行，返回 D 列的值"""
        end_row = sheet.used_range.last_cell.row
        for row in range(start_row, end_row + 1):
            cell_value = sheet.range(f"B{row}").value
            if cell_value and label_text in str(cell_value):
                return sheet.range(f"D{row}").value
        return None
    #获取待查项信息
    PE_score = get_score_by_label("体测成绩") #体育成绩
    dom_score = get_score_by_label("宿舍卫生") #寝室分数，需要结合寝室成绩和是否寝室长
    is_level = get_score_by_label("优秀宿舍") #优秀寝室加分
    position_score = get_score_by_label("社会工作") #班干部考核分数
    # 计算总分（如果启用签字功能）
    total_score = 0
    if enable_signature:
        total_score = calculate_total_score(sheet, score_column=score_column)
        # 填写签字信息
        handle_signature(sheet, total_score, sig_type=sig_type, score_column=score_column)
        wb.save()  # 保存修改
    wb.close()
    return [number, name, PE_score, dom_score, is_level, position_score, total_score]

#存储并调用分数计算规则
def score(PE_raw, dom_raw, is_level_raw, position_raw, position_level_raw, check_PE=True, check_dom=True, check_position=True):
    PE_score_cal = dom_score_cal = is_level_cal = position_score_cal = 0
    #体育分数计算
    if check_PE and PE_raw:
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
    if check_dom and dom_raw:
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
    if check_position and position_raw:
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
def check(cur, personal_info, check_PE=True, check_dom=True, check_position=True, enable_signature=False):
    global df_signature
    number = personal_info[0]
    name = personal_info[1]
    PE_score = int(personal_info[2]) if personal_info[2] else 0
    dom_score = int(personal_info[3]) if personal_info[3] else 0
    is_level = int(personal_info[4]) if personal_info[4] else 0
    position_score = int(personal_info[5]) if personal_info[5] else 0
    total_score = personal_info[6] if len(personal_info) > 6 else 0
    #纯无名氏
    if not number or not name:
        if check_PE or check_dom or check_position:
            df_log.loc[len(df_log)] = list(get_log_row([name,"无名氏", "", "", file_path], check_PE, check_dom, check_position))
        print(f"未找到学号或姓名，跳过该文件")
        return
    # 标记 names 表，该人员已被检查到
    cur.execute("UPDATE names SET checked = 1 WHERE id = ? AND name = ?", (number, name))
    # 如果仅执行签字功能（不执行任何检查功能）
    if enable_signature and not (check_PE or check_dom or check_position):
        print(f"{name} 处理完成，总分为{total_score}")
        # 添加到签字结果表格
        df_signature.loc[len(df_signature)] = [number, name, total_score]
        return
    # 如果待查项均为0，直接跳过
    if PE_score == 0 and dom_score == 0 and is_level == 0 and position_score == 0:
        print(f"{name} 的待查项均为0", end="")
        if enable_signature:
            print(f"，处理完成，总分为{total_score}")
            df_signature.loc[len(df_signature)] = [number, name, total_score]
        else:
            print("，跳过该文件")
        return
    # 追踪是否找到了各项数据
    data_get = []
    found_PE = False
    found_dom = False
    found_position = False
    #检查体育成绩
    if check_PE:
        cur.execute("SELECT * FROM PE WHERE id = ? AND name = ?", (number, name))
        result = cur.fetchone()
        if not result:
            df_log.loc[len(df_log)] = list(get_log_row([name,"未找到体育成绩", PE_score, "", file_path], check_PE, check_dom, check_position))
            data_get.append("") #体育成绩
            found_PE = False
        else:
            data_get.append(result[2]) if result[2] else  data_get.append("")
            found_PE = True
    else:
        data_get.append("")
        found_PE = True  # 未选中该项目也视为"找到"，避免后续比对
    #检查寝室分数
    if check_dom:
        cur.execute("SELECT * FROM dom WHERE id = ? AND name = ?", (number, name))
        result = cur.fetchone()
        if not result:
            df_log.loc[len(df_log)] = list(get_log_row([name,"未找到寝室分数", dom_score, "", file_path], check_PE, check_dom, check_position))
            data_get.append("") #寝室分数
            data_get.append("") #寝室等级
            found_dom = False
        else:
            data_get.append(result[4]) if result[4] else data_get.append("")
            data_get.append(result[5]) if result[5] else data_get.append("")
            found_dom = True
    else:
        data_get.append("")
        data_get.append("")
        found_dom = True
    #检查班干部考核
    if check_position:
        cur.execute("SELECT * FROM members WHERE id = ? AND name = ?", (number, name))
        result = cur.fetchone()
        if not result:
            df_log.loc[len(df_log)] = list(get_log_row([name,"未找到班干部考核", position_score, "", file_path], check_PE, check_dom, check_position))
            data_get.append("") #职位
            data_get.append("") #考核结果
            found_position = False
        else:
            data_get.append(result[2]) if result[2] else data_get.append("")
            data_get.append(result[3]) if result[3] else data_get.append("")
            found_position = True
    else:
        data_get.append("")
        data_get.append("")
        found_position = True
    PE_score_cal, dom_score_cal, is_level_cal, position_score_cal = score(data_get[0], data_get[1], data_get[2], data_get[3], data_get[4], check_PE, check_dom, check_position)
    testflag = True
    # 只在找到数据的情况下才进行分数比对，避免重复报错
    #体育成绩
    if check_PE and found_PE and PE_score_cal != PE_score:
        df_log.loc[len(df_log)] = list(get_log_row([name,"体育成绩", PE_score, PE_score_cal, file_path], check_PE, check_dom, check_position))
        testflag = False
    #寝室分数和优秀寝室
    if check_dom and found_dom:
        if dom_score_cal != dom_score:
            df_log.loc[len(df_log)] = list(get_log_row([name,"寝室分数", dom_score, dom_score_cal, file_path], check_PE, check_dom, check_position))
            testflag = False   
        if is_level_cal != is_level:
            df_log.loc[len(df_log)] = list(get_log_row([name,"寝室加分", is_level, is_level_cal, file_path], check_PE, check_dom, check_position))
            testflag = False
    #班干部考核
    if check_position and found_position and position_score_cal != position_score:
        df_log.loc[len(df_log)] = list(get_log_row([name,"班干部考核", position_score, position_score_cal, file_path], check_PE, check_dom, check_position))
        testflag = False
    # 检查是否有未找到的项
    if (check_PE and not found_PE) or (check_dom and not found_dom) or (check_position and not found_position):
        testflag = False
    #输出结果
    if testflag:
        print(f"{name} 检查通过")
    else:
        print(f"{name} 可能存在问题")
    # 如果同时启用了签字功能，也在检查后记录签字结果
    if enable_signature:
        print(f"{name} 签字信息已处理，总分为{total_score}")
        df_signature.loc[len(df_signature)] = [number, name, total_score]
        
#遍历文件夹得到个人综测分xls，读取信息并输出
def walk_main(cur, app, check_PE=True, check_dom=True, check_position=True, enable_signature=False, sig_type='', score_column='D'):
    global file_path
    global cnt
    global df_signature
    cnt = 0
    for root, dirs, files in walk("test_data"):
        for file in files:
            if (file.endswith('.xls')or file.endswith('.xlsx')) and "学年纪实测评表" in file:
                file_path = path.join(root, file)
                personal_info = get_personal_info(app, file_path, enable_signature=enable_signature, sig_type=sig_type, score_column=score_column)
                check(cur, personal_info, check_PE, check_dom, check_position, enable_signature)
                cnt += 1

# 检查data文件夹及文件是否存在（交互式版本）
def check_data_files_interactive():
    data_dir_1 = "data" #被调用数据文件夹
    data_dir_2 = "test_data" #待测试数据文件夹
    required_files = ["PE.xlsx", "dom.xlsx", "members.xlsx", "names.xlsx"]
    
    if not path.exists(data_dir_1):
        makedirs(data_dir_1)
        print(f"已创建缺失的文件夹: {data_dir_1}")
    if not path.exists(data_dir_2):
        makedirs(data_dir_2)
        print(f"已创建缺失的文件夹: {data_dir_2}")
    # 循环检查直到所有文件齐备
    while True:
        missing_files = []
        for f in required_files:
            if not path.exists(path.join(data_dir_1, f)):
                missing_files.append(f)
        if not missing_files:
            print(f"所有数据文件已准备完毕")
            return True
        #检查data文件夹的缺失文件
        print(f"\n缺少以下文件，请放入 {data_dir_1} 文件夹中：")
        for f in missing_files:
            print(f"  - {f}")
        input("文件准备完毕后，按回车键继续...")

# 检查test_data文件夹并统计内容
def check_test_data_status():
    test_data_dir = "test_data"
    if not path.exists(test_data_dir):
        print("test_data 文件夹不存在")
        return False
    # 统计根目录内的子文件夹和xls/xlsx文件
    items = []
    folders_count = 0
    excel_files_count = 0
    try:
        for item in listdir(test_data_dir):
            item_path = path.join(test_data_dir, item)
            if path.isdir(item_path):
                folders_count += 1
            elif path.isfile(item_path) and (item.endswith('.xls') or item.endswith('.xlsx')):
                excel_files_count += 1
    except Exception as e:
        print(f"读取 test_data 文件夹出错: {e}")
        return False
    #输出结果
    print(f"\ntest_data 文件夹统计:")
    print(f"  子文件夹个数: {folders_count}")
    print(f"  xls/xlsx 文件个数: {excel_files_count}")
    # 用户确认
    user_input = input("\n是否继续处理？(Y/N): ").strip().upper()
    if user_input == 'Y':
        return True
    else:
        print("已退出程序")
        return False

# 功能菜单定义和获取
def get_function_selection():
    #定义并输出菜单
    print("\n请选择要执行的功能（可同时输入多个数字，输入完成后按回车）：")
    print("  1 - 检查体育成绩")
    print("  2 - 检查寝室成绩")
    print("  3 - 检查班干部工作")
    print("  4 - 本班审核和计算签字")
    print("  5 - 其他班审核和计算签字")
    #用户选择
    user_input = input("请输入功能编号（如：123表示全选检查）: ").strip()
    check_PE = '1' in user_input
    check_dom = '2' in user_input
    check_position = '3' in user_input
    enable_signature_current_class = '4' in user_input
    enable_signature_other_class = '5' in user_input
    # 4 和 5 中只能选一个或都不选
    if enable_signature_current_class and enable_signature_other_class:
        print("警告：4 和 5 只能同时选择一个，已取消选择。")
        enable_signature_current_class = False
        enable_signature_other_class = False
    # 检查是否至少选择了一项功能
    if not (check_PE or check_dom or check_position or enable_signature_current_class or enable_signature_other_class):
        print("未识别到有效的功能编号，将退出程序")
        return None, None, None, False, ''
    #收集获取结果
    selected = []
    if check_PE:
        selected.append("体育成绩")
    if check_dom:
        selected.append("寝室成绩")
    if check_position:
        selected.append("班干部工作")
    if enable_signature_current_class:
        selected.append("本班签字")
    if enable_signature_other_class:
        selected.append("其他班签字")
    print(f"已选择功能: {', '.join(selected)}")
    # 确定签字类型
    sig_type = ''
    if enable_signature_current_class:
        sig_type = '本班'
    elif enable_signature_other_class:
        sig_type = '其他班'
    # 返回选择的功能（增加了签字功能标志和签字类型）
    return check_PE, check_dom, check_position, (enable_signature_current_class or enable_signature_other_class), sig_type

# 根据选择的功能项构建日志行记录
def get_log_row(full_row, check_PE, check_dom, check_position):
    """
    根据选择的功能项返回对应的日志行
    full_row: [姓名, 不匹配项, 表格中分数, 数据库计算分数, 文件路径]
    """
    name, mismatch_item, table_score, db_score, file_path = full_row
    return [name, mismatch_item, table_score, db_score, file_path]

# 获取报告列名
def get_report_columns(check_PE, check_dom, check_position):
    """根据选择的功能项返回报告的列名"""
    return ["姓名", "不匹配项", "表格中分数", "数据库计算分数", "文件路径"]

def main():
    global df_log
    global df_signature
    global g_config
    check_data_files_interactive() # Step 1: 检查 data 文件夹及文件（交互式）   
    if not check_test_data_status(): # Step 2: 检查 test_data 文件夹状态
        _exit(0)
    check_PE, check_dom, check_position, enable_signature, sig_type = get_function_selection()  # Step 3: 获取用户选择的功能
    if check_PE is None:
        _exit(0)
    if enable_signature:  # Step 4: 如果选择了签字功能（4 或 5），先检查配置文件，再询问列号
        # 加载并确认对应签字类型的配置（带死循环重试）
        sig_config_item = load_and_confirm_signature_config(sig_type)
        # 构建全局配置对象（键名与签字类型对应）
        g_config = {
            'signature_config': {
                sig_type: sig_config_item
            }
        }
        # 询问列号
        score_column = 'D'  # 默认列号
        print(f"\n默认使用 D 列计算总分")
        column_input = input("是否更改为 E 列？(Y/N): ").strip().upper()
        if column_input == 'Y':
            score_column = 'E'
            print(f"已更改为使用 {score_column} 列计算总分")
    else:
        score_column = 'D'
    # Step 5: 初始化数据库连接和 DataFrame
    conn = connect('data.db') 
    cur = conn.cursor()
    df_log = DataFrame(columns=get_report_columns(check_PE, check_dom, check_position))
    # 初始化签字结果表格
    df_signature = DataFrame(columns=["学号", "姓名", "总分"])
    prepare_sql(conn) # Step 6: 准备数据库
    app = App(visible=False, add_book=False) # Step 7: 创建 xlwings 应用并执行检查
    try:
        walk_main(cur, app, check_PE, check_dom, check_position, enable_signature, sig_type, score_column)
    finally:
        # 保证程序结束时只关闭这一主进程，避免残留 workbook 进程
        app.quit()
    cur.execute("SELECT id,name FROM names WHERE checked = 0") # Step 8: 检查 names 表是否所有目标都被扫到
    missed = cur.fetchall()
    if not missed:
        print("所有人员均已检查到。")
    else:
        missed_names = ", ".join([f"{r[1]}({r[0]})" for r in missed])
        print(f"以下人员未被检查到：{missed_names}")
        df_log.loc[len(df_log)] = ["[未检查到]", "名单缺失", "", "", ""]
    if len(df_log) > 0: # Step 9: 保存报告
        df_log.to_excel("报错汇总.xlsx", index=False)
        print("检查完成，报错汇总已保存至 报错汇总.xlsx")
    if enable_signature and len(df_signature) > 0: # Step 10: 保存签字结果
        df_signature.to_excel("个人综测分.xlsx", index=False)
        print("签字结果已保存至 个人综测分.xlsx")
    #报告结果并结束程序
    print(f"共检查 {cnt} 个文件")
    system("pause")
    cur.close() 
    conn.close()

if __name__ == "__main__":
    main()