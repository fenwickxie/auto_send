#!/usr/bin/env python
# -*- coding: UTF-8 -*-
"""
@Project : officeauto
@File    : scheduler.py.py
@IDE     : PyCharm
@Author  : xie.fangyu
@Date    : 2025/3/31 下午1:58
"""

import threading
import time
from datetime import datetime

import keyboard
from PyQt5.QtCore import QTimer, QObject, pyqtSignal

IDLE = 0
RUNNING = 1
STATUS = {IDLE: "空闲中", RUNNING: "运行中"}


class WeChatScheduler(QObject):
    log_signal = pyqtSignal(str)  # 日志信号
    status_signal = pyqtSignal(int)  # 状态信号

    def __init__(self):
        super().__init__()
        self.scheduled_time = None
        self.timer = QTimer()
        self.timer.timeout.connect(self.message_send)
        self.is_running = False
        self.schedule_thread = None
        self.current_window = None
        self.prepare_time = 10.0  # 预估准备操作需要的时间(秒)
        self.once_schedule = False

        # 消息内容缓存
        self.target = ""
        self.content = ""
        self.shortcuts = {
            # "open_wechat": "ctrl+alt+w",
            "open_search": "alt+f",
            "send_message": "ctrl+enter",
        }
        self.delays = {
            # "window_delay": 0.8,
            "search_delay": 1.0,
            "search_result_delay": 1.5,
            "chat_delay": 1,
            "line_delay": 0.1,
        }

    def update_shortcuts(self, shortcuts):
        """更新快捷键设置"""
        self.shortcuts.update(shortcuts)
        self.log_signal.emit(f"快捷键已更新: {self.shortcuts}")

    def update_delays(self, delays):
        """更新延时设置"""
        self.delays.update(delays)
        self.log_signal.emit(f"延时设置已更新: {self.delays}")

    def start_once_schedule(self, target, content, scheduled_time):
        """启动一次性定时任务"""
        self.once_schedule = True
        if not self.validate_inputs(target, content, scheduled_time):
            return

        self.target = target
        self.content = content
        self.scheduled_time = scheduled_time

        self.is_running = True
        self.status_signal.emit(RUNNING)
        self.log_signal.emit("启动一次性定时任务")

        now = datetime.now()
        seconds_remain = (self.scheduled_time - now).total_seconds()

        if seconds_remain > self.prepare_time:

            delay_ms = int((seconds_remain - self.prepare_time) * 1000)
            # delay_time = seconds_remain - self.prepare_time
            self.log_signal.emit(f"将在 {delay_ms/1000.0:.2f} 秒后开始准备消息")
            # self.log_signal.emit(f"将在 {delay_time:.2f} 秒后开始准备消息")
            QTimer.singleShot(
                delay_ms, lambda: self.message_schedule(self.scheduled_time)
            )
            # threading.Timer(
            #     delay_time, self.message_schedule, [self.scheduled_time]
            # ).start()
        else:
            self.log_signal.emit("时间紧张，立即开始准备消息")
            self.message_schedule(self.scheduled_time)
        # self._on_once_task_complete()
        # # 确保任务完成后状态更新
        # self.timer.timeout.connect(self._on_once_task_complete)

    def start_repeating_schedule(self, target, content, days, send_time, repeat_type):
        """启动循环定时任务"""
        if self.schedule_thread and self.schedule_thread.is_alive():
            self.log_signal.emit("循环定时任务已在运行")
            return
        self.once_schedule = False  # 设置非一次性任务
        if not self.validate_repeat_inputs(target, content, days, repeat_type):
            return

        self.target = target
        self.content = content
        self.is_running = True
        self.status_signal.emit(RUNNING)
        self.log_signal.emit("启动循环定时任务")

        self.schedule_thread = threading.Thread(
            target=self.run_repeating_schedule,
            args=(days, send_time, repeat_type),
        )
        self.schedule_thread.daemon = True
        self.schedule_thread.start()

    def run_repeating_schedule(self, days, send_time, repeat_type):
        """循环定时任务线程"""
        self.log_signal.emit("循环定时任务线程已启动")
        try:
            while self.is_running:
                self.log_signal.emit("检查任务状态...")
                now = datetime.now()
                self.log_signal.emit(f"当前时间: {now}")

                # 检查是否是选定的日期或工作日
                if repeat_type == 0:  # 按日期
                    if now.day not in days:
                        # self.log_signal.emit(
                        #     f"当前日期 {now.day} 不在选定日期 {days} 中"
                        # )
                        time.sleep(3)
                        continue
                elif repeat_type == 1:  # 按工作日
                    weekday = now.weekday()
                    if weekday not in days:
                        # self.log_signal.emit(
                        #     f"当前工作日 {weekday} 不在选定工作日 {days} 中"
                        # )
                        time.sleep(3)
                        continue

                # 检查时间是否匹配
                target_datetime = datetime.combine(now.date(), send_time)
                self.log_signal.emit(f"目标时间: {target_datetime}")
                if now >= target_datetime:
                    self.log_signal.emit("今天的目标时间已过，等待下一天")
                    time.sleep(60)
                    continue

                # 计算等待时间
                wait_seconds = (target_datetime - now).total_seconds()
                self.log_signal.emit(f"距离目标时间还有 {wait_seconds:.1f} 秒")
                if wait_seconds > self.prepare_time:
                    delay_ms = int((wait_seconds - self.prepare_time) * 1000)
                    # delay_time = wait_seconds - self.prepare_time
                    self.log_signal.emit(f"等待 {delay_ms/1000:.1f} 秒后开始准备消息")
                    # self.log_signal.emit(f"等待 {delay_time:.2f} 秒后开始准备消息")
                    QTimer.singleShot(
                        delay_ms, lambda: self.message_schedule(target_datetime)
                    )
                    # threading.Timer(
                    #     delay_time, self.message_schedule, [self.scheduled_time]
                    # ).start()
                else:
                    # # 立即在主线程中执行
                    QTimer.singleShot(0, lambda: self.message_schedule(target_datetime))
                    # threading.Timer(
                    #     0, self.message_schedule, [self.scheduled_time]
                    # ).start()
                # 发送后等待一段时间再检查，避免重复发送
                for _ in range(60):
                    if not self.is_running:
                        self.log_signal.emit("任务已停止，退出循环")
                        return
                    time.sleep(1)

        except Exception as e:
            self.log_signal.emit(f"循环定时任务线程出错: {str(e)}")
        finally:
            self.log_signal.emit("循环定时任务线程已退出")

    def stop_scheduler(self):
        """停止所有定时任务"""
        self.is_running = False
        self.timer.stop()
        try:
            self.timer.timeout.disconnect()  # 尝试断开所有信号连接
        except TypeError:
            self.log_signal.emit("警告: 定时器信号未连接或已断开")

        if self.schedule_thread and self.schedule_thread.is_alive():
            self.schedule_thread.join(timeout=5)  # 等待线程退出

        self.status_signal.emit(IDLE)
        self.log_signal.emit("定时任务已停止")

        # # 尝试恢复窗口状态
        # self.activate_previous_window()

    def message_send_immed(self, target, content):
        """立即发送消息，不影响定时任务运行状态"""
        if not target or not content:
            self.log_signal.emit("错误: 目标或内容为空")
            return

        # 不改变运行状态，只执行发送流程
        original_target = self.target
        original_content = self.content

        self.target = target
        self.content = content

        self.status_signal.emit(RUNNING)
        self.log_signal.emit("开始立即发送流程...")
        if self.message_prepare():
            time.sleep(0.1)  # 确保准备完成
            self.message_send()

        # 恢复原来的消息内容
        self.target = original_target
        self.content = original_content

        # 不改变定时任务运行状态
        if self.is_running:
            self.status_signal.emit(RUNNING)
        else:
            self.status_signal.emit(IDLE)
            # 通过进程名激活企业微信窗口

    def message_prepare(self):
        """预先打开聊天窗口并输入消息内容，只差发送"""
        try:
            # # 保存当前活动窗口以便后续恢复
            # self.current_window = pyautogui.getActiveWindow()
            # self.log_signal.emit("保存当前窗口状态")

            # # 通过企微进程名（'WXWork.exe'）打开界面
            # self.activate_window("WXWork.exe")
            # time.sleep(self.delays["window_delay"])  # 等待窗口激活
            # self.log_signal.emit("激活企业微信窗口")

            # # 模拟按下快捷键打开企微
            # keyboard.press_and_release(self.shortcuts["open_wechat"])
            # time.sleep(0.8)  # 等待企微界面出现

            # 按下Alt+F打开搜索框
            self.log_signal.emit(f"模拟按下 {self.shortcuts['open_search']} 打开搜索框")
            keyboard.press_and_release(self.shortcuts["open_search"])
            time.sleep(self.delays["search_delay"])  # 等待搜索框出现

            # 输入目标对话名称
            self.log_signal.emit(f"输入目标对话: {self.target}")
            keyboard.write(self.target)
            time.sleep(self.delays["search_result_delay"])  # 等待搜索结果

            # 按Enter选择第一个结果
            self.log_signal.emit("模拟按下 Enter 选择对话")
            keyboard.press_and_release("enter")
            time.sleep(self.delays["chat_delay"])  # 等待对话打开

            # 输入消息内容并保持分段
            self.log_signal.emit("开始输入消息内容")
            for line in self.content.split("\n"):
                keyboard.write(line)
                # Shift+Enter换行
                keyboard.press_and_release("shift+enter")
                time.sleep(self.delays["line_delay"])

            # # 删除最后一个多余的换行
            # keyboard.press_and_release("backspace")
            # time.sleep(0.1)

            self.log_signal.emit("消息准备完成，等待发送时机")
            return True

        except Exception as e:
            error_msg = f"准备消息出错: {str(e)}"
            self.log_signal.emit(error_msg)
            # self.activate_previous_window()
            return False

    def message_schedule(self, target_time):
        """准备发送并启动精准定时"""
        if not self.is_running:
            return

        self.log_signal.emit("开始准备消息...")
        try:
            if self.message_prepare():
                remaining_ms = int(
                    (target_time - datetime.now()).total_seconds() * 1000
                )
                if remaining_ms > 0:
                    self.log_signal.emit(
                        f"消息准备完成，等待 {remaining_ms:.0f} 毫秒后发送"
                    )
                    self.timer.setSingleShot(True)
                    self.timer.start(remaining_ms)
                else:
                    self.log_signal.emit("立即发送消息")
                    self.message_send()
        except Exception as e:
            self.log_signal.emit(f"消息准备失败: {str(e)}")
        if self.once_schedule:
            self.is_running = False
            self.status_signal.emit(IDLE)

    def message_send(self):
        """在精确时间执行发送操作"""
        try:
            self.log_signal.emit("正在执行发送操作...")
            self.log_signal.emit(f"模拟按下 {self.shortcuts['send_message']} 发送消息")
            keyboard.press_and_release(self.shortcuts["send_message"])
            send_time = datetime.now().strftime("%H:%M:%S.%f")[:-3]
            status_msg = f"消息已于 ({send_time}) 发送"
            time.sleep(0.5)

            self.log_signal.emit(status_msg)

            # 停止定时器
            self.timer.stop()

        except Exception as e:
            error_msg = f"发送出错: {str(e)}"
            self.log_signal.emit(error_msg)

        # 如果是一次性任务，停止调度器并更新状态
        if self.once_schedule:
            self.log_signal.emit("一次性任务完成，停止定时器")
            self.is_running = False
            self.status_signal.emit(IDLE)

        # # 恢复之前的活动窗口
        # self.activate_previous_window()

    def activate_previous_window(self):
        """恢复之前的活动窗口"""
        if hasattr(self, "current_window") and self.current_window:
            try:
                self.log_signal.emit("恢复之前的窗口状态")
                self.current_window.activate()
            except Exception as e:
                self.log_signal.emit(f"恢复窗口失败: {str(e)}")

    def validate_inputs(self, target, content, scheduled_time):
        """验证一次性任务的输入"""
        if not target:
            self.log_signal.emit("错误: 请输入目标对话名称")
            self.restore_inputs()  # 恢复输入框和按钮的可用性
            return False

        if not content:
            self.log_signal.emit("错误: 请输入发送内容")
            self.restore_inputs()  # 恢复输入框和按钮的可用性
            return False

        if scheduled_time <= datetime.now():
            self.log_signal.emit("错误: 请选择未来的时间")
            self.restore_inputs()  # 恢复输入框和按钮的可用性
            return False

        return True

    def restore_inputs(self):
        """恢复输入框和按钮的可用性"""
        self.status_signal.emit(IDLE)  # 更新状态为“空闲中”
        self.log_signal.emit("恢复输入框和按钮的可用性")

    def validate_repeat_inputs(self, target, content, days, repeat_type):
        """验证循环任务的输入"""
        if not target:
            self.log_signal.emit("错误: 请输入目标对话名称")
            return False

        if not content:
            self.log_signal.emit("错误: 请输入发送内容")
            return False

        if not days:
            msg = (
                "错误: 请输入每月发送的日期"
                if repeat_type
                else "错误: 请至少选择一个工作日"
            )
            self.log_signal.emit(msg)
            return False

        return True

    def test_keyboard_operations(self, target):
        """测试键盘操作"""
        if not target:
            self.log_signal.emit("错误: 请输入目标对话名称")
            return False
        try:
            self.log_signal.emit("开始测试键盘操作...")

            # 测试搜索框打开
            self.log_signal.emit("测试打开搜索框...")
            keyboard.press_and_release(self.shortcuts["open_search"])
            time.sleep(1)

            # 测试输入
            self.log_signal.emit("测试输入...")
            keyboard.write(target)
            time.sleep(1)

            # 测试回车
            self.log_signal.emit("测试回车...")
            keyboard.press_and_release("enter")
            time.sleep(1)

            self.log_signal.emit("键盘操作测试完成")
            return True
        except Exception as e:
            self.log_signal.emit(f"键盘操作测试失败: {str(e)}")
            return False
