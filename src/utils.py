#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
author: fenwickxie
date: 2025-07-04 17:35:05
project: auto_send
filename: utils.py
version: 1.0
"""

import threading
from datetime import datetime


def async_log(log_signal, message):
    threading.Thread(target=lambda: log_signal.emit(message), daemon=True).start()


def format_time(dt=None):
    if dt is None:
        dt = datetime.now()
    return dt.strftime("%H:%M:%S.%f")[:-3]
