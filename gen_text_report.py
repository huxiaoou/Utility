import os
import pandas as pd
import datetime as dt

report_date = input("Please input the report date [format = 'YYYYMMDD'], enter the 'ENTER' key directly will use today as report date:\n") or dt.datetime.now().strftime("%Y%m%d")

key_elements = {
    "futures": {
        "src_dir": os.path.join("E:\\", "Works", "Trade", "Reports", "output"),
        "file_name": "04_衍生品持仓情况明细表_大宗商品_{}.xlsx".format(report_date),
        "description": "2、大宗商品策略：",
    },
    "equity": {
        "src_dir": os.path.join("E:\\", "Works", "Trade", "Reports_Equity", "output"),
        "file_name": "04_持仓情况明细表_股票可转债_1001000016_{}.xlsx".format(report_date),
        "description": "1、国轩高科托管项目：",
    },
    "equity2": {
        "src_dir": os.path.join("E:\\", "Works", "Trade", "Reports_Equity2", "output"),
        "file_name": "04_持仓情况明细表_股票可转债_1003000010_{}.xlsx".format(report_date),
        "description": "2、可转债托管项目：",
    },
}

for report_type, report_type_data in key_elements.items():
    report_file = report_type_data.get("file_name")
    report_dir = report_type_data.get("src_dir")
    report_descrption = report_type_data.get("description")
    report_path = os.path.join(report_dir, report_date[0:4], report_date, report_file)
    if not os.path.exists(report_path):
        print("{} does not exist, please check again".format(report_path))
        continue
    
    if report_type == "futures":
        print("二、衍生品类：")
    elif report_type == "equity":
        print("三、协作类：")
    
    report_df = pd.read_excel(report_path, dtype="str", header=3)
    for x in report_df["证券名称"].astype(str):
        if x.find("年初至今") > 0:
            z = x.strip().replace(":", "：").replace(",", "，")
            z = z.replace("注：年初至今，", report_descrption)
            print(z)
