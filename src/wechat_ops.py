#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
author: fenwickxie
date: 2025-07-04 17:34:41
project: auto_send
filename: wechat_ops.py
version: 1.0
"""

"""封装所有与企业微信窗口、键盘操作相关的底层实现"""
import win32gui
import keyboard
import time
from config import WECHAT_WINDOW_TITLE


def activate_wechat(window_title=WECHAT_WINDOW_TITLE, delay=1.0):
    hwnd = win32gui.FindWindow(None, window_title)
    if hwnd:
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(delay)
        return True
    return False


def search_and_select_chat(target, shortcuts, delays):
    keyboard.press_and_release(shortcuts["open_search"])
    time.sleep(delays["search_delay"])
    keyboard.write(target)
    time.sleep(delays["search_result_delay"])
    keyboard.press_and_release("enter")
    time.sleep(delays["chat_delay"])


def input_message_content(content, delays):
    lines = content.split("\n")
    for i, line in enumerate(lines):
        keyboard.write(line)
        time.sleep(0.1)
        keyboard.press_and_release("shift+enter")
        time.sleep(delays["line_delay"])


def send_message(shortcuts):
    keyboard.press_and_release(shortcuts["send_message"])
