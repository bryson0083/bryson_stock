# -*- coding: utf-8 -*-
"""
google spreadsheet API

說明:
    1. 建立google spreadsheet API 公用模組
    2. 透過公用程式模組，提供google spreadsheet API功能

"""

import os
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# 引用自建公用模組
from proj_util_pkg.settings import ProjEnvSettings


class GspreadManager:
    def __init__(self):
        self.gc = self.gspread_init()

    def gspread_init(self):
        """ google spreadsheet API 初始化 """

        # 連線參數設定
        scope = ['https://spreadsheets.google.com/feeds',
                 'https://www.googleapis.com/auth/drive']
        CREDS_FILE = f"{os.environ.get('cfg_path')}/gspread_credentials.json"
        creds = ServiceAccountCredentials.from_json_keyfile_name(CREDS_FILE, scope)
        gc = gspread.authorize(creds)

        return gc

    def get_spreadsheet(self, spreadsheet_key):
        """ 取得google spreadsheet """
        # return self.gc.open(spreadsheet_key)
        return self.gc.open_by_key(spreadsheet_key)

    def get_worksheet(self, spreadsheet_key, worksheet_name):
        """ 取得google spreadsheet worksheet """
        return self.get_spreadsheet(spreadsheet_key).worksheet(worksheet_name)

    def get_worksheet_values(self, spreadsheet_key, worksheet_name):
        """ 取得google spreadsheet worksheet values """
        return self.get_worksheet(spreadsheet_key, worksheet_name).get_all_values()

    def update_worksheet_values(self, spreadsheet_key, worksheet_name, values):
        """ 更新google spreadsheet worksheet values """
        self.get_worksheet(spreadsheet_key, worksheet_name).update(values)

    def append_worksheet_values(self, spreadsheet_key, worksheet_name, values):
        """ 新增google spreadsheet worksheet values """
        self.get_worksheet(spreadsheet_key, worksheet_name).append_row(values)

    def clear_worksheet(self, spreadsheet_key, worksheet_name):
        """ 清除google spreadsheet worksheet """
        self.get_worksheet(spreadsheet_key, worksheet_name).clear()

    def create_worksheet(self, spreadsheet_key, worksheet_name):
        """ 建立google spreadsheet worksheet """
        self.get_spreadsheet(spreadsheet_key).add_worksheet(title=worksheet_name, rows="100", cols="20")

    def delete_worksheet(self, spreadsheet_key, worksheet_name):
        """ 刪除google spreadsheet worksheet """
        self.get_spreadsheet(spreadsheet_key).del_worksheet(self.get_worksheet(spreadsheet_key, worksheet_name))

    def get_worksheet_list(self, spreadsheet_key):
        """ 取得google spreadsheet worksheet list """
        return self.get_spreadsheet(spreadsheet_key).worksheets()

    def recreate_worksheet(self, spreadsheet_key, sheet_name):
        """ 刪除後重建工作表 """
        
        sheet_list = [sht.title for sht in self.get_worksheet_list(spreadsheet_key)]
        if sheet_name in sheet_list:
            self.delete_worksheet(spreadsheet_key, sheet_name)
        
        self.create_worksheet(spreadsheet_key, sheet_name)


if __name__ == '__main__':
    # 測試
    gspread_mgr = GspreadManager()
    print(gspread_mgr.get_worksheet_list("台股選股清單"))
    # print(gspread_mgr.get_worksheet_values("test", "工作表1"))
    # gspread_mgr.append_worksheet_values("test", "工作表1", ["測試1", "測試2", "測試3"])
    # print(gspread_mgr.get_worksheet_values("test", "工作表1"))
    # gspread_mgr.update_worksheet_values("test", "工作表1", [[1, 2, 3], [4, 5, 6], [7, 8, 9]])
    # print(gspread_mgr.get_worksheet_values("test", "工作表1"))
    # gspread_mgr.clear_worksheet("test", "工作表1")
    # print(gspread_mgr.get_worksheet_values("test", "工作表1"))
    # gspread_mgr.create_worksheet("test", "工作表2")
    # print(gspread_mgr.get_worksheet_list("test"))
    # gspread_mgr.delete_worksheet("test", "工作表2")
    # print(gspread_mgr.get_worksheet_list("test"))