# -*- coding: utf-8 -*-
"""
TW Stock Topic

說明:
    1. 讀取台股題材資料
    2. 透過輸入的股票代碼，查詢相關題材資料

"""

import os
import pandas as pd

# 引用自建公用模組
from proj_util_pkg.settings import ProjEnvSettings


def read_topic_stocks(inp_stock_no):
    """ 讀取股票代號題材資訊 """

    # 設定環境參數
    TW_STOCK_TOPIC_DATA = os.environ.get('external_data_path')  # 台股題材資料路徑
    tw_topis_df = pd.read_excel(f"{TW_STOCK_TOPIC_DATA}/tw_stock_topics.xlsx", dtype=str)

    return str(tw_topis_df[tw_topis_df["stock_no"] == inp_stock_no].topic.to_list())[1:-1]


if __name__ == '__main__':
    print(f'2330相關概念股: {read_topic_stocks("2330")}')
