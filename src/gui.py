#!/usr/bin/env python
# -*- coding: UTF-8 -*-
"""
@Project : officeauto
@File    : gui.py
@IDE     : PyCharm
@Author  : xie.fangyu
@Date    : 2025/3/31 下午2:00
"""

from PyQt5.QtCore import Qt, QDateTime, QTime
from PyQt5.QtGui import QIntValidator, QKeySequence
from PyQt5.QtWidgets import (
    QMainWindow,
    QVBoxLayout,
    QHBoxLayout,
    QWidget,
    QLabel,
    QLineEdit,
    QTextEdit,
    QPushButton,
    QGroupBox,
    QDateTimeEdit,
    QCheckBox,
    QComboBox,
    QSplitter,
    QPlainTextEdit,
    QKeySequenceEdit,
    QSpinBox,  # 添加 QSpinBox 用于时间偏移输入
)
from auto_send.src.scheduler import STATUS, IDLE, RUNNING


class WeChatSchedulerUI(QMainWindow):
    def __init__(self, scheduler):
        super().__init__()
        self.scheduler = scheduler
        self.setWindowTitle("企业微信定时消息发送工具")
        self.setGeometry(100, 100, 900, 600)

        # 连接业务逻辑的信号
        self.scheduler.log_signal.connect(self.append_log)
        self.scheduler.status_signal.connect(self.update_status)

        self.init_ui()

    def init_ui(self):
        main_widget = QWidget()
        main_layout = QHBoxLayout()

        # 左侧控制面板
        left_panel = QWidget()
        left_layout = QVBoxLayout()

        # 消息内容设置
        msg_group = QGroupBox("消息内容设置")
        msg_layout = QVBoxLayout()

        self.target_label = QLabel("目标对话名称(必须准确,如不唯一则为检索到的第一个):")
        self.target_input = QLineEdit()
        self.target_input.setPlaceholderText("请输入完整的企业微信对话名称")
        self.content_label = QLabel("发送内容:")
        self.content_input = QTextEdit()
        self.content_input.setPlaceholderText("输入要发送的消息内容，支持多行")

        msg_layout.addWidget(self.target_label)
        msg_layout.addWidget(self.target_input)
        msg_layout.addWidget(self.content_label)
        msg_layout.addWidget(self.content_input)
        msg_group.setLayout(msg_layout)

        # 定时设置
        time_group = QGroupBox("定时设置")
        time_layout = QVBoxLayout()

        # 一次性定时
        self.once_checkbox = QCheckBox("一次性定时")
        self.once_checkbox.setChecked(True)
        self.once_checkbox.toggled.connect(self.update_once_schedule)

        self.datetime_label = QLabel("发送时间:")
        self.oncetime_input = QDateTimeEdit()
        self.oncetime_input.setDateTime(QDateTime.currentDateTime().addSecs(60))
        self.oncetime_input.setDisplayFormat("yyyy-MM-dd HH:mm:ss.zzz")
        # self.oncetime_input.setMinimumDateTime(QDateTime.currentDateTime())

        # 循环定时
        self.repeat_checkbox = QCheckBox("循环定时")
        self.repeat_type = QComboBox()
        self.repeat_type.addItems(["每周", "每月"])
        self.repeat_type.setEnabled(False)

        self.weekly_options = QWidget()
        weekly_layout = QHBoxLayout()
        self.weekdays = []
        for day in ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]:
            cb = QCheckBox(day)
            self.weekdays.append(cb)
            weekly_layout.addWidget(cb)
        self.weekly_options.setLayout(weekly_layout)
        self.weekly_options.setVisible(False)

        self.monthly_options = QWidget()
        monthly_layout = QHBoxLayout()
        self.day_label = QLabel("每月第几天:")
        self.day_input = QLineEdit()
        self.day_input.setValidator(QIntValidator(1, 31))
        self.day_input.setPlaceholderText("1-31")
        monthly_layout.addWidget(self.day_label)
        monthly_layout.addWidget(self.day_input)
        self.monthly_options.setLayout(monthly_layout)
        self.monthly_options.setEnabled(False)
        self.monthly_options.setVisible(False)

        self.repeat_time_label = QLabel("每天发送时间:")
        self.repeat_time_input = QDateTimeEdit()
        self.repeat_time_input.setDisplayFormat("HH:mm:ss.zzz")
        self.repeat_time_input.setTime(QTime(8, 0, 0))
        self.repeat_time_input.setEnabled(False)

        self.offset_label = QLabel("时间偏移(毫秒):")
        self.offset_input = QSpinBox()
        self.offset_input.setRange(-60000, 60000)  # 偏移范围为 -60秒 到 +60秒
        self.offset_input.setSuffix(" ms")
        self.offset_input.setValue(0)

        time_layout.addWidget(self.once_checkbox)
        time_layout.addWidget(self.datetime_label)
        time_layout.addWidget(self.oncetime_input)
        time_layout.addWidget(self.repeat_checkbox)
        time_layout.addWidget(self.repeat_type)
        time_layout.addWidget(self.weekly_options)
        time_layout.addWidget(self.monthly_options)
        time_layout.addWidget(self.repeat_time_label)
        time_layout.addWidget(self.repeat_time_input)
        time_layout.addWidget(self.offset_label)
        time_layout.addWidget(self.offset_input)
        time_group.setLayout(time_layout)

        # 控制按钮
        control_layout = QHBoxLayout()
        self.start_btn = QPushButton("开始定时")
        self.start_btn.clicked.connect(self.start_scheduler)
        self.stop_btn = QPushButton("停止定时")
        self.stop_btn.clicked.connect(self.stop_scheduler)
        self.stop_btn.setEnabled(False)
        self.send_now_btn = QPushButton("立即发送")
        self.send_now_btn.clicked.connect(self.send_message_now)

        control_layout.addWidget(self.start_btn)
        control_layout.addWidget(self.stop_btn)
        control_layout.addWidget(self.send_now_btn)

        # 状态显示
        self.status_label = QLabel("状态: 未运行")
        self.status_label.setAlignment(Qt.AlignCenter)

        # 添加到左侧布局
        left_layout.addWidget(msg_group)
        left_layout.addWidget(time_group)
        left_layout.addLayout(control_layout)
        left_layout.addWidget(self.status_label)
        left_panel.setLayout(left_layout)

        # 右侧布局
        right_panel = QWidget()
        right_layout = QVBoxLayout()

        # 快捷键设置
        shortcut_group = QGroupBox("快捷键设置")
        shortcut_layout = QVBoxLayout()

        self.open_wechat_label = QLabel("打开应用快捷键:")
        self.open_wechat_input = QKeySequenceEdit(QKeySequence("Ctrl+Alt+W"))
        self.open_search_label = QLabel("打开搜索框快捷键:")
        self.open_search_input = QKeySequenceEdit(QKeySequence("Alt+F"))
        self.send_message_label = QLabel("发送消息快捷键:")
        self.send_message_input = QKeySequenceEdit(QKeySequence("Ctrl+Enter"))

        shortcut_layout.addWidget(self.open_wechat_label)
        shortcut_layout.addWidget(self.open_wechat_input)
        shortcut_layout.addWidget(self.open_search_label)
        shortcut_layout.addWidget(self.open_search_input)
        shortcut_layout.addWidget(self.send_message_label)
        shortcut_layout.addWidget(self.send_message_input)
        shortcut_group.setLayout(shortcut_layout)

        # 添加到右侧布局
        right_layout.addWidget(shortcut_group)

        self.log_label = QLabel("运行日志:")
        self.log_display = QPlainTextEdit()
        self.log_display.setReadOnly(True)
        self.log_display.setLineWrapMode(QPlainTextEdit.NoWrap)
        self.log_display.setStyleSheet("font-family:Consolas,'Courier New',monospace;")

        # 添加清除日志按钮
        self.clear_log_btn = QPushButton("清空日志")
        self.clear_log_btn.clicked.connect(self.clear_log)

        right_layout.addWidget(self.log_label)
        right_layout.addWidget(self.log_display)
        right_layout.addWidget(self.clear_log_btn)
        right_panel.setLayout(right_layout)

        # 使用QSpliter实现可调整大小的左右面板
        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(left_panel)
        splitter.addWidget(right_panel)
        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 2)

        main_layout.addWidget(splitter)
        main_widget.setLayout(main_layout)
        self.setCentralWidget(main_widget)

        # 连接信号
        self.repeat_checkbox.toggled.connect(self.update_repeat_schedule)
        self.repeat_type.currentIndexChanged.connect(self.update_repeat_options)

    def update_ui_state(self):
        """根据当前运行状态及一次性/循环定时状态更新输入框和按钮的可用性"""
        is_running = self.scheduler.is_running

        # 一次性定时相关
        self.once_checkbox.setEnabled(not is_running)
        self.oncetime_input.setEnabled(
            not is_running and self.once_checkbox.isChecked()
        )

        # 循环定时相关
        self.repeat_checkbox.setEnabled(not is_running)
        self.repeat_type.setEnabled(not is_running and self.repeat_checkbox.isChecked())
        self.weekly_options.setEnabled(
            not is_running
            and self.repeat_checkbox.isChecked()
            and self.repeat_type.currentIndex() == 0
        )
        self.monthly_options.setEnabled(
            not is_running
            and self.repeat_checkbox.isChecked()
            and self.repeat_type.currentIndex() == 1
        )
        self.repeat_time_input.setEnabled(
            not is_running and self.repeat_checkbox.isChecked()
        )

        # 控制按钮
        self.start_btn.setEnabled(not is_running)
        self.stop_btn.setEnabled(is_running)
        self.offset_input.setEnabled(not is_running)

    def update_repeat_schedule(self):
        """更新循环定时相关选项的可见性和状态"""
        if self.repeat_checkbox.isChecked():
            self.once_checkbox.setChecked(False)
            self.update_repeat_options()
        self.update_ui_state()

    def update_once_schedule(self):
        """更新一次性定时相关选项的状态"""
        if self.once_checkbox.isChecked():
            self.repeat_checkbox.setChecked(False)
        self.update_ui_state()

    def update_repeat_options(self):
        """根据循环定时类型更新选项的可见性"""
        if self.repeat_type.currentIndex() == 0:  # 每周
            self.weekly_options.setVisible(True)
            self.monthly_options.setVisible(False)
        elif self.repeat_type.currentIndex() == 1:  # 每月
            self.weekly_options.setVisible(False)
            self.monthly_options.setVisible(True)
        self.update_ui_state()

    def start_scheduler(self):
        target = self.target_input.text().strip()
        content = self.content_input.toPlainText().strip()
        offset = self.offset_input.value()

        # 更新快捷键设置
        shortcuts = {
            "open_wechat": self.open_wechat_input.keySequence().toString(),
            "open_search": self.open_search_input.keySequence().toString(),
            "send_message": self.send_message_input.keySequence().toString(),
        }
        self.scheduler.update_shortcuts(shortcuts)

        if self.once_checkbox.isChecked():
            scheduled_time = (
                self.oncetime_input.dateTime().addMSecs(offset).toPyDateTime()
            )
            self.scheduler.start_once_schedule(target, content, scheduled_time)
        elif self.repeat_checkbox.isChecked():
            if self.repeat_type.currentIndex() == 0:  # 0 corresponds to "每周"
                days = [i for i, cb in enumerate(self.weekdays) if cb.isChecked()]
            elif self.repeat_type.currentIndex() == 1:  # 1 corresponds to "每月"
                try:
                    days = [int(self.day_input.text())]
                except:
                    days = []

            send_time = self.repeat_time_input.time().addMSecs(offset).toPyTime()
            self.scheduler.start_repeating_schedule(
                target,
                content,
                days,
                send_time,
                self.repeat_type.currentIndex(),
            )
        ## 禁用时间相关输入
        # self.set_time_inputs_enabled(False)

    def stop_scheduler(self):
        self.scheduler.stop_scheduler()
        ## 恢复时间相关输入的可用状态
        # self.set_time_inputs_enabled(True)

    def set_time_inputs_enabled(self, enabled):
        """设置所有时间相关输入框的可用状态"""
        self.oncetime_input.setEnabled(enabled)
        self.once_checkbox.setEnabled(enabled)

        self.repeat_checkbox.setEnabled(enabled)
        self.repeat_type.setEnabled(enabled)
        self.monthly_options.setEnabled(enabled)
        self.day_input.setEnabled(enabled)
        self.weekly_options.setEnabled(enabled)
        for cb in self.weekdays:
            cb.setEnabled(enabled)
        self.repeat_time_input.setEnabled(enabled)

        # 根据当前选择的定时类型调整显示
        self.update_once_schedule()
        self.update_repeat_schedule()

    def send_message_now(self):
        """立即发送消息，不影响定时任务状态"""
        target = self.target_input.text().strip()
        content = self.content_input.toPlainText().strip()

        # 保存当前按钮状态
        start_btn_state = self.start_btn.isEnabled()
        stop_btn_state = self.stop_btn.isEnabled()

        # 执行发送，不改变定时任务状态
        self.scheduler.send_message_now(target, content)

        # 恢复按钮状态
        self.start_btn.setEnabled(start_btn_state)
        self.stop_btn.setEnabled(stop_btn_state)

    def append_log(self, message):
        """追加日志到显示区域"""
        self.log_display.appendPlainText(message)
        # 自动滚动到底部
        self.log_display.verticalScrollBar().setValue(
            self.log_display.verticalScrollBar().maximum()
        )

    def update_status(self, status):
        """更新状态显示并调整 UI"""
        self.status_label.setText(f"状态: {STATUS[status]}")
        self.update_ui_state()

    def clear_log(self):
        """清除日志"""
        self.log_display.clear()

    def closeEvent(self, event):
        """窗口关闭事件"""
        self.scheduler.stop_scheduler()
        event.accept()
