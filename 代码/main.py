import os
import pandas as pd


def concatXls(file_path, type_name):
    all_fund_df = pd.DataFrame()

    for file_name in os.listdir(file_path):  # 外循环：对所有文件做循环
        # 以下只处理对应种类的数据
        if type_name not in file_name or file_name.endswith('.xlsx') is False:
            continue

        excel_file = pd.io.excel.ExcelFile(file_path + file_name)  # +为连接两个字符串
        # 从文件名字符串里提取年和月
        year = file_name[0:4]
        month = file_name[4:6]

        for sheet_name in excel_file.sheet_names:  # 内循环：对每个sheet做循环
            fund_df = pd.read_excel(excel_file, sheet_name, index_col=0)
            if fund_df.empty is False:
                # 增加'年月'列名，放在最左边
                old_columns = list(fund_df.columns)
                new_columns = ['年月'] + old_columns
                # 新增'年月'列到df
                fund_df['年月'] = pd.to_datetime(
                    [year + '-' + month] * fund_df.shape[0]).to_period('M')  # 年月是字符串year+'-'+month
                # 合并到all_fund_df中
                all_fund_df = pd.concat(
                    [all_fund_df, fund_df.loc[:, new_columns]], ignore_index=True)
                # all_fund_df.to_csv(r"./新表格/所有公积金.csv", encoding='utf_8_sig')
    if type_name != "医疗险":
        all_fund_df.rename(
            columns={'姓名': '员工姓名', '身份证号': '员工身份证号'}, inplace=True)

    return all_fund_df


if __name__ == "__main__":
    all_types = ["公积金", "四险", "医疗险"]
    file_path = "../原始表格/"
    merge_t = ["年月", "员工身份证号", "员工姓名"]
    all_sum = 0
    average = {}
    df0 = concatXls(file_path, "公积金")
    df1 = concatXls(file_path, "四险")
    df2 = concatXls(file_path, "医疗险")

    res = pd.merge(df0, df1, left_on=merge_t, right_on=merge_t, how='outer')
    res = pd.merge(res, df2, left_on=merge_t, right_on=merge_t, how='outer')
    res.drop(columns=['Unnamed: 0.1'], inplace=True)
    res = res[["年月", "员工身份证号", "员工姓名", "公积金",
               "养老险", "失业险", "工伤险", "生育险", "医疗险"]]
    res.to_csv(r"./新表格/五险一金汇总表.csv", encoding='utf_8_sig', index=False)

    # 求总和
    for i in range(3, len(res.columns)):
        all_sum += res[res.columns[i]].sum()
    print("五险一金的总缴纳金额: ", all_sum)

    # 求平均值
    for index, row in res.iterrows():
        year = row["年月"].strftime('%F')
        e_name = row["员工姓名"]
        p_fund = row["公积金"]
        if year == "2019":
            if pd.isnull(p_fund):
                continue
            if e_name in average:
                average[e_name][0] += p_fund
                average[e_name][1] += 1
            else:
                average[e_name] = [p_fund, 1]

    print("2019年月均公积金不足1000元的员工有: ")
    for key, value in average.items():
        if value[0] / value[1] < 1000:
            print(key, value[0] / value[1])
    # print(res)
