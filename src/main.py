#!/usr/bin/env python
# -*- coding: UTF-8 -*-
"""
@Project : officeauto 
@File    : main.py
@IDE     : vscode 
@Author  : xie.fangyu
@Date    : 2025/3/31 下午1:58 
"""

import sys
from PyQt5.QtWidgets import QApplication
from gui import WeChatSchedulerUI
from scheduler import WeChatScheduler
import logging
import logging.handlers

if __name__ == "__main__":
    try:
        # 配置日志
        handler = logging.handlers.RotatingFileHandler(
            'error.log', maxBytes=5*1024*1024, backupCount=3)
        logging.basicConfig(
            level=logging.ERROR,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[handler]
        )
        
        app = QApplication(sys.argv)
        
        # 创建业务逻辑层实例
        scheduler = WeChatScheduler()
        
        
        # 创建UI层实例并注入业务逻辑
        window = WeChatSchedulerUI(scheduler)
        window.show()
        
        sys.exit(app.exec_())
    except Exception as e:
        logging.error(f"{e} exception occurred", exc_info=True)
        sys.exit(f"Error: {e}")
