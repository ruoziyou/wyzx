import pymysql
import pandas as pd
import openpyxl
import os

# 获取数据库名称和MySQL密码
db_name = input("数据库名称: ")
db_pw = input("MySQL密码: ")

# 连接MySQL并获取游标
conn = pymysql.connect(host="127.0.0.1", user="root", passwd=db_pw, db=db_name, charset="utf8")
cursor = conn.cursor()

# 获取输出表格的类型，名称，和路径。检查是否合理并报错
export_type = input("输出表格类型（输入1或2）：")
if not (export_type == "1" or export_type == "2"):
    raise ValueError("表格类型错误")

export_name = input("输出文件名称：")
if not export_name[-5:] == ".xlsx":
    raise ValueError("文件名称错误")

export_path = input("输出文件路径（如果没有路径要求不用输入）：")
if export_path == "":
    export_file = export_name
else:
    if not os.path.exists(export_path) or os.path.isfile(export_path):
        raise ValueError("文件路径错误")
    else:
        export_file = f"{export_path}/{export_name}"

# 获取全部user的id和姓名
cursor.execute("SELECT id FROM user")
all_id = cursor.fetchall()
cursor.execute("SELECT uuid FROM user")
all_uuid = cursor.fetchall()


# 定义获取数据的函数，返回含有该数据的list
def get(value_name, table_name, user_id):
    # 如果需要把Blob转换成Text：
    # cursor.execute(f"SELECT convert({value_name} using utf8) FROM {table_name} WHERE user_id = {user_id}")
    cursor.execute(f"SELECT {value_name} FROM {table_name} WHERE user_id = {user_id}")
    output_tuple = cursor.fetchall()

    # 把tuple转换为list（为了去掉tuple里每一项的括号和逗号，只保留数据，美化输出数据）
    output_list = []
    for output in output_tuple:
        output_list.append(output[0])
    return output_list


# 第1种表格类型
if export_type == "1":
    # 获取要查询的姓名及对应id，并检查该id是否在数据库中
    input_uuid = input("登记人姓名：")
    if not (input_uuid,) in all_uuid:
        raise ValueError("姓名错误")
    cursor.execute(f"SELECT id FROM user WHERE uuid= '{input_uuid}'")
    input_id = cursor.fetchone()[0]

    # 获取内网数据
    private_time = get("create_time", "private", input_id)
    private_ip = get("ip", "private", input_id)
    private_ggroup = get("ggroup", "private", input_id)
    private_ports = get("ports", "private", input_id)

    # 把内网数据整合成DataFrame并按日期生序排序
    df_private = pd.DataFrame(
        {"启用日期": private_time, "IP": private_ip, "项目组": private_ggroup, "开放端口数量": private_ports})
    df_private = df_private.sort_values(by="启用日期", axis=0, ascending=True)

    # 获取外网数据
    public_time = get("create_time", "public", input_id)
    public_ip = get("ip", "public", input_id)
    public_servers = get("servers", "public", input_id)
    public_ports = get("ports", "public", input_id)

    # 把外网数据整合成DataFrame并按日期生序排序
    df_public = pd.DataFrame(
        {"启用日期": public_time, "IP": public_ip, "服务": public_servers, "开放端口数量": public_ports})
    df_public = df_public.sort_values(by="启用日期", axis=0, ascending=True)

    # 计算资产数量
    counter = df_private.shape[0] + df_public.shape[0]

    # 创建Excel文件并输入表格的header
    wb = openpyxl.Workbook()
    sheet = wb.create_sheet("sheet1", 0)
    sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)
    sheet.merge_cells(start_row=2, start_column=1, end_row=2, end_column=4)
    sheet.merge_cells(start_row=2, start_column=5, end_row=2, end_column=8)
    sheet.cell(1, 1).value = f"{input_uuid}: {counter}"
    sheet.cell(2, 1).value = "内网资产"
    sheet.cell(2, 5).value = "外网资产"
    wb.save(export_file)

    # 输出内网数据和外网数据到Excel表格里面
    with pd.ExcelWriter(export_file, mode="a", engine="openpyxl", if_sheet_exists="overlay") as writer:
        df_private.to_excel(writer, sheet_name="sheet1", index=False, header=True, startrow=2)
        df_public.to_excel(writer, sheet_name="sheet1", index=False, header=True, startrow=2, startcol=4)

# 第二种表格类型
else:
    # 创建Excel文件并输入表格的header
    wb = openpyxl.Workbook()
    sheet = wb.create_sheet("sheet1", 0)
    sheet.merge_cells(start_row=1, start_column=3, end_row=1, end_column=5)
    sheet.merge_cells(start_row=1, start_column=6, end_row=1, end_column=8)
    sheet.cell(1, 3).value = "内网"
    sheet.cell(1, 6).value = "外网"
    wb.save(export_file)

    # 设置计数器（计数器设置为0是为了只在第一次输出DataFrame的时候打印header）
    counter = 0
    user_counter = 0

    # 获取所有登记日期
    cursor.execute("SELECT create_time FROM user")
    all_time = cursor.fetchall()

    # 遍历全部user id
    while user_counter < len(all_id):
        # 获取内网数据
        private_ip = get("ip", "private", all_id[user_counter][0])
        private_ggroup = get("ggroup", "private", all_id[user_counter][0])
        private_ports = get("ports", "private", all_id[user_counter][0])

        # 把内网数据整合成DataFrame
        df_private = pd.DataFrame({"IP": private_ip, "项目组": private_ggroup, "开放端口数量": private_ports})
        df_private = df_private.sort_values(by="IP", axis=0, ascending=True)

        # 获取外网数据
        public_ip = get("ip", "public", all_id[user_counter][0])
        public_servers = get("servers", "public", all_id[user_counter][0])
        public_ports = get("ports", "public", all_id[user_counter][0])

        # 把外网数据整合成DataFrame并按IP生序排序
        df_public = pd.DataFrame({"IP": public_ip, "服务": public_servers, "开放端口数量": public_ports})
        df_public = df_public.sort_values(by="IP", axis=0, ascending=True)

        # 当前user id的内网/外网资源占据的行数
        incre = max(df_private.shape[0], df_public.shape[0])

        # 获取登记人信息并整合成对应行数的DataFrame
        user_time = (all_time[user_counter]) * incre
        user_uuid = (all_uuid[user_counter]) * incre
        df_user = pd.DataFrame({"日期": user_time, "登记人": user_uuid})

        # 输出登记人信息，内网数据，和外网数据到Excel表格里面
        with pd.ExcelWriter(export_file, mode="a", engine="openpyxl", if_sheet_exists="overlay") as writer:
            df_user.to_excel(writer, sheet_name="sheet1", index=False, header=not counter, startrow=counter + 1)
            df_private.to_excel(writer, sheet_name="sheet1", index=False, header=not counter, startrow=counter + 1,
                                startcol=2)
            df_public.to_excel(writer, sheet_name="sheet1", index=False, header=not counter, startrow=counter + 1,
                               startcol=5)

        # 更新计数器
        if counter == 0:
            counter += 1  # 由于第一次输出要打印header，避免后面输出会覆盖第一次的最后一行数据
        counter += incre
        user_counter += 1

# 关闭游标和连接
cursor.close()
conn.close()
