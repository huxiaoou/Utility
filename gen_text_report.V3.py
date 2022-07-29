import os
import sys
import pandas as pd
import datetime as dt

'''
created @ 2021-03-19
0.  derived from V2. since account 10003000010 has been given back to FICC,
    no reports about this account is required.
updated @ 2021-07-22
1.  10003000010 is added to project again, so the old code is used again.
'''

report_date = sys.argv[1]
base_date = "20211231"  # change every year
WANYUAN = 1e4

sep_line = "-" * 100
key_elements = {
    "futures": {
        "src_dir": os.path.join("/Works", "Trade", "Reports", "output"),
        "file_name_03": "03_衍生品当日成交汇总_大宗商品_{}.xlsx".format(report_date),
        "file_name_04": "04_衍生品持仓情况明细表_大宗商品_{}.xlsx".format(report_date),
        "description": "2、大宗商品策略：",
        "cost_val_adj_ratio": 2,
    },
    # "equity": {
    #     "src_dir": os.path.join("/Works", "Trade", "Reports_Equity", "output"),
    #     "file_name_03": "03_当日成交汇总_股票可转债_1001000016_{}.xlsx".format(report_date),
    #     "file_name_04": "04_持仓情况明细表_股票可转债_1001000016_{}.xlsx".format(report_date),
    #     "description": "1、国轩高科托管项目：",
    #     "cost_val_adj_ratio": 2,

    # },
    "equity2": {
        "src_dir": os.path.join("/Works", "Trade", "Reports_Equity2", "output"),
        "file_name_03": "03_当日成交汇总_股票可转债_1003000010_{}.xlsx".format(report_date),
        "file_name_04": "04_持仓情况明细表_股票可转债_1003000010_{}.xlsx".format(report_date),
        "description": "2、可转债托管项目：",
        "cost_val_adj_ratio": 2,
    },
}

# for weekly report
print(sep_line)
for report_type, report_type_data in key_elements.items():
    # shared settings
    report_dir = report_type_data.get("src_dir")
    title_description = report_type_data.get("description")
    cost_val_adj_ratio = report_type_data.get("cost_val_adj_ratio")

    # load 03 traded
    traded_file = report_type_data.get("file_name_03")
    traded_path = os.path.join(report_dir, report_date[0:4], report_date, traded_file)
    if not os.path.exists(traded_path):
        print("{} does not exist, please check again".format(traded_path))
        continue

    # load 04 position
    position_file = report_type_data.get("file_name_04")
    position_path = os.path.join(report_dir, report_date[0:4], report_date, position_file)
    if not os.path.exists(position_path):
        print("{} does not exist, please check again".format(position_path))
        continue

    # load traded and position data
    traded_df = pd.read_excel(traded_path, header=2)
    position_df = pd.read_excel(position_path, dtype={"证券名称": str, "总成本": float}, header=3)

    # print(traded_df)
    # print(position_df)

    # print section title
    if report_type == "futures":
        print("二、衍生品类：")
    elif report_type == "equity":
        print("三、协作类：")

    # print
    if report_type == "futures":
        for x in position_df["证券名称"].astype(str):
            if x.find("年初至今") > 0:
                z = x.strip().replace(":", "：").replace(",", "，")
                z = z.replace("注：年初至今，", title_description + "相较去年末，")
                print(z)
    else:
        traded_qty = traded_df["成交数量"].sum()
        cost_val = position_df["总成本"].sum() / cost_val_adj_ratio
        if traded_qty > 0:
            print("{}今日有交易，持仓成本{:.0f}万元。".format(title_description, cost_val / WANYUAN))
        else:
            print("{}今日无交易，持仓成本{:.0f}万元。".format(title_description, cost_val / WANYUAN))
print(sep_line)

# ---- for weekly report
src_file = "组合净值.xlsx"
src_path = os.path.join("/Works", "Trade", "Reports", "intermediary", src_file)
src_df = pd.read_excel(src_path, sheet_name="期货净值表").set_index("日期")
# print(src_df.tail(30))
base_realized = src_df.at[base_date[0:4] + "-" + base_date[4:6] + "-" + base_date[6:8], "累积实现盈亏"]
base_unrealized = src_df.at[base_date[0:4] + "-" + base_date[4:6] + "-" + base_date[6:8], "持仓盈亏"]
base_tot = base_realized + base_unrealized
report_realized = src_df.at[report_date[0:4] + "-" + report_date[4:6] + "-" + report_date[6:8], "累积实现盈亏"]
report_unrealized = src_df.at[report_date[0:4] + "-" + report_date[4:6] + "-" + report_date[6:8], "持仓盈亏"]
report_tot = report_realized + report_unrealized
week_report = "至{}日收盘，组合今年以来已实现盈利{:.2f}万元，加上当前持仓浮盈{:.2f}万元，累积总盈利{:.2f}万元，业务开展以来总盈利{:.2f}万元。".format(
    report_date[0:4] + "-" + report_date[4:6] + "-" + report_date[6:8],
    (report_realized - base_realized) / WANYUAN,
    (report_unrealized - base_unrealized) / WANYUAN,
    (report_tot - base_tot) / WANYUAN,
    report_tot / WANYUAN
)
print("周报：")
print(week_report)
print(sep_line)
