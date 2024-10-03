# -*- coding: utf-8 -*-
"""
初始環境參數設定

"""
import os
from dotenv import load_dotenv

class ProjEnvSettings:

    def __init__(self):
        self.proj_env_settings_init()

    def proj_env_settings_init(self):
        """ Load .env file """
        # Get the path to the directory this file is in
        base_dir = os.path.abspath(os.path.dirname(__file__))

        # Get the path to the project root directory
        base_dir = os.path.abspath(os.path.join(base_dir, os.pardir))

        # Setting app project head path to system environment variable
        os.environ["app_home"] = base_dir

        # Connect the path with your '.env' file name
        load_dotenv(os.path.join(base_dir, "proj_util_pkg", "config", ".env"))

settings = ProjEnvSettings()

# print(f"[環境變數]app_home: {os.environ.get('app_home')}")
# print(f"[環境變數]cfg_path: {os.environ.get('cfg_path')}")