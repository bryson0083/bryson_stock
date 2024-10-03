# -*- coding: utf-8 -*-
"""
Finlab API Manager

說明:
    1. 建立 API 公用模組
    2. 透過公用程式模組，封裝提供API功能

"""

import os
import finlab

# 引用自建公用模組
from proj_util_pkg.settings import ProjEnvSettings


class FinlabManager:
    def __init__(self):
        self.finlab = self.finlab_init()

    def finlab_init(self):
        """ Finlab API 初始化 """

        # 讀取api token
        FINLAB_API_TOKEN = os.environ.get('finlab_api_token')

        # finlab api 服務設定
        finlab.login(FINLAB_API_TOKEN)

        return finlab