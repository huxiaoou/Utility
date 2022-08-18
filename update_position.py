import sys
import pandas as pd
import xlwings as xw
import xlwings.utils


def get_realized_pnl_cumsum(t_report_date, t_config: dict) -> (float, float, dict):
    nav_df = pd.read_excel(t_config["nav"]["path"], sheet_name="期货净值表", header=0, index_col=0)
    nav_df["trade_date"] = nav_df.index.map(lambda z: z.strftime("%Y%m%d"))
    nav_df = nav_df.set_index("trade_date")
    realized_pnl_cumsum_all = nav_df.at[t_report_date, "累积实现盈亏"]

    # by portfolio
    nav_df = pd.read_excel(t_config["nav"]["path"], sheet_name="按策略计盈亏", header=[0, 1], index_col=0)
    nav_df["trade_date"] = nav_df.index.map(lambda z: z.strftime("%Y%m%d"))
    nav_df = nav_df.set_index("trade_date")

    strategy_name_list = ["趋势跟踪", "OU过程", "期限结构", "LTRR", "MTM", "RS"]
    strategy_realized_pnl_cumsum = {z: nav_df.at[t_report_date, (z, "累积实现盈亏")] for z in strategy_name_list}
    realized_pnl_cumsum_sub = sum(strategy_realized_pnl_cumsum.values())

    print("截至{}, 按策略计算，累积实现盈亏:{:.2f}元。".format(t_report_date, realized_pnl_cumsum_sub))
    print("截至{}, 按总体计算，累积实现盈亏:{:.2f}元。".format(t_report_date, realized_pnl_cumsum_all))
    return realized_pnl_cumsum_all, realized_pnl_cumsum_sub, strategy_realized_pnl_cumsum


def update_position_and_pnl(t_config: dict, t_pos_type_list: list, t_realized_pnl_cumsum_all: float, t_realized_pnl_cumsum_sub: float, t_pnl_by_strategy: dict):
    """

    :param t_config:
    :param t_pos_type_list: ["pos_all", "pos_sub"]
    :param t_realized_pnl_cumsum_all:
    :param t_realized_pnl_cumsum_sub:
    :param t_pnl_by_strategy:
    :return:
    """

    wb = xw.Book(t_config["monitor"]["path"])
    for pos_type in t_pos_type_list:
        pos_df = pd.read_csv(t_config[pos_type]["path"], encoding="GB18030")
        ws = wb.sheets[t_config[pos_type]["monitor_sheet_name"]]

        # delete old rows
        focus_row_id = 2
        while ws.range("A{}".format(focus_row_id)).value != "合计":
            ws.api.Rows(focus_row_id).Delete()

        # add new rows
        for pos_row_id in range(len(pos_df)):
            ws.api.Rows(focus_row_id).Insert()

            if pos_type == "pos_all":
                ws.range("A{}".format(focus_row_id)).value = pos_df.at[pos_row_id, "合约"]
                ws.range("B{}".format(focus_row_id)).value = pos_df.at[pos_row_id, "方向"]
                ws.range("C{}".format(focus_row_id)).value = pos_df.at[pos_row_id, "数量"]
                ws.range("D{}".format(focus_row_id)).value = pos_df.at[pos_row_id, "成本价"]

                ws.range("E{}".format(focus_row_id)).value = "=RTD(\"wdf.rtq\",,A{},\"LastPrice\")".format(focus_row_id)
                ws.range("F{}".format(focus_row_id)).value = "=s_info_contractmultiplier(A{})".format(focus_row_id)
                ws.range("G{}".format(focus_row_id)).value = "=D{0}*F{0}*C{0}".format(focus_row_id)
                ws.range("H{}".format(focus_row_id)).value = "=E{0}*F{0}*C{0}".format(focus_row_id)
                ws.range("I{}".format(focus_row_id)).value = "=IF(B{0}=\"多\",1,-1)*(E{0}-D{0})*C{0}*F{0}".format(focus_row_id)

            if pos_type == "pos_sub":
                ws.range("A{}".format(focus_row_id)).value = pos_df.at[pos_row_id, "组合标志"]
                ws.range("B{}".format(focus_row_id)).value = pos_df.at[pos_row_id, "策略标志"]
                ws.range("C{}".format(focus_row_id)).value = pos_df.at[pos_row_id, "合约"]
                ws.range("D{}".format(focus_row_id)).value = pos_df.at[pos_row_id, "方向"]
                ws.range("E{}".format(focus_row_id)).value = pos_df.at[pos_row_id, "数量"]
                ws.range("F{}".format(focus_row_id)).value = pos_df.at[pos_row_id, "成本价"]

                ws.range("G{}".format(focus_row_id)).value = "=RTD(\"wdf.rtq\",,C{},\"LastPrice\")".format(focus_row_id)
                ws.range("H{}".format(focus_row_id)).value = "=s_info_contractmultiplier(C{})".format(focus_row_id)
                ws.range("I{}".format(focus_row_id)).value = "=F{0}*H{0}*E{0}".format(focus_row_id)
                ws.range("J{}".format(focus_row_id)).value = "=G{0}*H{0}*E{0}".format(focus_row_id)
                ws.range("K{}".format(focus_row_id)).value = "=IF(D{0}=\"多\",1,-1)*(G{0}-F{0})*E{0}*H{0}".format(focus_row_id)

            # for next row
            focus_row_id += 1

        # update sum
        if pos_type == "pos_all":
            ws.range("C{}".format(focus_row_id)).value = "=sum(C2:C{})".format(len(pos_df) + 1)
            ws.range("G{}".format(focus_row_id)).value = "=sum(G2:G{})".format(len(pos_df) + 1)
            ws.range("H{}".format(focus_row_id)).value = "=sum(H2:H{})".format(len(pos_df) + 1)
            ws.range("I{}".format(focus_row_id)).value = "=sum(I2:I{})".format(len(pos_df) + 1)

        if pos_type == "pos_sub":
            ws.range("E{}".format(focus_row_id)).value = "=sum(E2:E{})".format(len(pos_df) + 1)
            ws.range("I{}".format(focus_row_id)).value = "=sum(I2:I{})".format(len(pos_df) + 1)
            ws.range("J{}".format(focus_row_id)).value = "=sum(J2:J{})".format(len(pos_df) + 1)
            ws.range("K{}".format(focus_row_id)).value = "=sum(K2:K{})".format(len(pos_df) + 1)

        # format adjustment
        end_column_id = "A"
        if pos_type == "pos_all":
            end_column_id = "I"
        if pos_type == "pos_sub":
            end_column_id = "K"
        ws.range("A2:{}{}".format(end_column_id, len(pos_df) + 1)).row_height = 12
        ws.range("A2:{}{}".format(end_column_id, len(pos_df) + 1)).color = (255, 255, 255)
        ws.range("A2:{}{}".format(end_column_id, len(pos_df) + 1)).api.Font.Size = 10
        ws.range("A2:{}{}".format(end_column_id, len(pos_df) + 1)).api.Font.Color = xlwings.utils.rgb_to_int((0, 0, 0))

    # ================= update summary =================
    ws = wb.sheets["summary"]
    # ----------------- update total description -----------------
    focus_row_id = 2
    while True:
        conds = (
            ws.range("A{}".format(focus_row_id - 1)).value == "总计",
            ws.range("A{}".format(focus_row_id)).value == "组合",
            ws.range("A{}".format(focus_row_id + 1)).value == "总体",
        )
        if all(conds):
            break
        focus_row_id += 1

    ws.range("F{}".format(focus_row_id)).value = t_realized_pnl_cumsum_sub
    ws.range("F{}".format(focus_row_id + 1)).value = t_realized_pnl_cumsum_all

    # ----------------- update strategy realized pnl -----------------
    focus_row_id = 2
    strategy_name = ""
    while strategy_name != "总计":
        strategy_name = ws.range("A{}".format(focus_row_id)).value
        if strategy_name in t_pnl_by_strategy:
            ws.range("F{}".format(focus_row_id)).value = t_pnl_by_strategy[strategy_name]
        focus_row_id += 1

    # save and close
    wb.save()
    wb.close()
    return 0


report_date = sys.argv[1]
config = {
    "monitor": {"path": "\Works\Monitor\monitor.V3.xlsx"},
    "pos_all": {
        "path": "\Works\Monitor\pos.copy_and_paste_to_monitor.sheet_portfolio.csv",
        "monitor_sheet_name": "Portfolio",
    },
    "pos_sub": {
        "path": "\Works\TradeClearing\data\pos.by_portfolio\{}\pos.by_portfolio.{}.csv".format(report_date[0:4], report_date),
        "monitor_sheet_name": "Portfolio_GroupBy",
    },
    "nav": {"path": "\Works\Trade\Reports\intermediary\组合净值.xlsx"},
}

pnl_all, pnl_sub, pnl_by_strategy = get_realized_pnl_cumsum(t_report_date=report_date, t_config=config)

update_position_and_pnl(
    t_config=config,
    t_pos_type_list=["pos_all", "pos_sub"],
    t_realized_pnl_cumsum_all=pnl_all,
    t_realized_pnl_cumsum_sub=pnl_sub,
    t_pnl_by_strategy=pnl_by_strategy,
)
