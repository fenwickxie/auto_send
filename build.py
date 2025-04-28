#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2025 4月 16 22:25
# @Author  : fenwickxie
# @PROJECT : officeauto
# @File    : build.py
# @Software: PyCharm

import PyInstaller.__main__
import platform

version = "1.1.0"  # 版本号

# 获取系统名称和架构
system_name = platform.system()  # 返回 'Windows', 'Linux', 'Darwin' 等
machine_arch = platform.machine()  # 返回 'ARM', 'AMD64' 等

# 动态生成名称
app_name: str = f"auto_send_{system_name}_{machine_arch}_v{version}"

# 打包配置
auto_send_build = [
    "./src/main.py",  # 主入口文件
    f"--name={app_name}",  # 输出文件名
    "--onefile",
    "--windowed",
    "--icon=./assets/icon.ico",
    "--hidden-import=PyQt5.QtCore",
    "--hidden-import=PyQt5.QtGui",
    "--hidden-import=PyQt5.QtWidgets",
    "--hidden-import=keyboard",
    "--hidden-import=pyautogui",
    # "--exclude-module=tkinter",
    # "--exclude-module=unittest",
    # "--exclude-module=numpy",
    "--distpath=./dist",  # 输出目录
    "--workpath=./build",  # 构建目录
    "--clean",  # 清理临时文件
    # "--log-level=WARN",  # 日志级别
]

PyInstaller.__main__.run(auto_send_build)
