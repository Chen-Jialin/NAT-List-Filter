print("正在导入 pandas 库，请稍等……")
import pandas as pd
print("导入成功！")
print("读取待检名单，请稍等……")
should_test = pd.read_excel("待检名单.xlsx")
print("读取成功！")
print("读取近期登记名单，请稍等……")
recent = pd.read_excel("近期登记名单.xlsx")
print("读取成功！")
test_interval = int(input("请输入规定的检查时间间隔（按回车提交）："))
# test_interval = 2

print("启动筛选，请稍等……")
recent["登记时间"] = recent["登记时间"].map(lambda x: x.floor(freq="D"))
date_range_tmp = sorted(recent["登记时间"].drop_duplicates())

data = should_test.copy(deep=True)
date_range = []
this_date = date_range_tmp[0]
while (this_date <= date_range_tmp[-1]):
    date_range.append(this_date)
    this_date += pd.Timedelta(days=1)
for i in range(len(date_range)):
    col_name = date_range[i].strftime("%Y-%m-%d")
    data[col_name] = False

for i in range(len(recent)):
    data.loc[data["姓名"] == recent.loc[i, "姓名"], recent.loc[i,"登记时间"].strftime("%Y-%m-%d")] = True

data_violate = pd.DataFrame(columns=data.columns.values.tolist())
for i in range(len(data)):
    for j in range(3,data.shape[1] - test_interval):
        tested = False
        for k in range(test_interval):
            if data.iloc[i,j + k]:
                tested = True
                break
        if not tested:
            data_violate = data_violate.append(data.iloc[i])
            break

def num2letter(num):
    letter_list = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    if num <= 26:
        return letter_list[num]
    else:
        return "Z"


writer = pd.ExcelWriter("筛选结果.xlsx")
workbook = writer.book
data_violate.to_excel(writer, 'sheet')
worksheet = writer.sheets['sheet']
# colored_range = 
fmt = workbook.add_format({'bg_color': "green"})
worksheet.conditional_format("E1:{0}{1}".format(num2letter(data_violate.shape[1]),len(data_violate)), {"type": "cell", "criteria": "=", "value": True, "format": fmt})
writer.save()
# data_violate.to_excel("筛选结果.xlsx")
input('筛选完毕，筛选结果已存入"筛选结果.xlsx"，按回车就可以关闭这个窗口了~')