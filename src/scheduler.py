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
import pyautogui
from PyQt5.QtCore import QTimer, QObject, pyqtSignal

IDLE = 0
RUNNING = 1
STATUS = {IDLE: "空闲中", RUNNING: "运行中"}


class WeChatScheduler(QObject):
    log_signal = pyqtSignal(str)  # 日志信号
    status_signal = pyqtSignal(int)  # 状态信号
    start_timer_signal = pyqtSignal(int)  # 启动定时器信号（以毫秒为单位）

    def __init__(self):
        super().__init__()
        self.scheduled_time = None
        self.timer = QTimer()
        self.timer.timeout.connect(self.message_send)
        self.is_running = False
        self.schedule_thread = None
        self.current_window = None
        self.preparation_time = 5.0  # 预估准备操作需要的时间(秒)
        self.once_schedule = False

        # 连接信号到槽
        self.start_timer_signal.connect(self.start_timer)

        # 消息内容缓存
        self.target = ""
        self.content = ""
        self.shortcuts = {
            "open_wechat": "ctrl+alt+w",
            "open_search": "alt+f",
            "send_message": "ctrl+enter",
        }

    def update_shortcuts(self, shortcuts):
        """更新快捷键设置"""
        self.shortcuts.update(shortcuts)
        self.log_signal.emit(f"快捷键已更新: {self.shortcuts}")

    def start_once_schedule(self, target, content, scheduled_time):
        """启动一次性定时任务"""
        self.once_schedule = True  # 设置一次性任务标志
        if not self.validate_inputs(target, content, scheduled_time):
            return

        self.target = target
        self.content = content
        self.scheduled_time = scheduled_time

        self.is_running = True
        self.status_signal.emit(RUNNING)
        self.log_signal.emit("启动一次性定时任务")

        now = datetime.now()
        seconds_left = (self.scheduled_time - now).total_seconds()

        if seconds_left > self.preparation_time:
            # 提前准备消息
            delay_time = seconds_left - self.preparation_time
            self.log_signal.emit(f"将在 {delay_time:.1f} 秒后开始准备消息")
            threading.Timer(
                delay_time,
                lambda: self.message_schedule(self.scheduled_time),
            ).start()
        else:
            # 时间太紧，立即准备
            self.log_signal.emit("时间紧张，立即开始准备消息")
            self.message_schedule(self.scheduled_time)

        # 确保任务完成后状态更新
        self.timer.timeout.connect(self._on_once_task_complete)

    def _on_once_task_complete(self):
        """一次性任务完成后的清理操作"""
        self.log_signal.emit("一次性任务完成，清理状态")
        self.timer.stop()
        self.is_running = False
        self.status_signal.emit(IDLE)

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
                        self.log_signal.emit(f"当前日期 {now.day} 不在选定日期 {days} 中")
                        time.sleep(1)
                        continue
                elif repeat_type == 1:  # 按工作日
                    weekday = now.weekday()
                    if weekday not in days:
                        self.log_signal.emit(f"当前工作日 {weekday} 不在选定工作日 {days} 中")
                        time.sleep(1)
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
                if wait_seconds > self.preparation_time:
                    self.log_signal.emit(
                        f"等待 {wait_seconds - self.preparation_time:.1f} 秒后开始准备消息"
                    )
                    for _ in range(int(wait_seconds - self.preparation_time)):
                        if not self.is_running:
                            self.log_signal.emit("任务已停止，退出循环")
                            return
                        time.sleep(1)
                    self.message_schedule(target_datetime)
                else:
                    self.log_signal.emit("时间紧张，立即准备并发送消息")
                    self.message_schedule(target_datetime)

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

        # 尝试恢复窗口状态
        self.activate_previous_window()

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

    def message_prepare(self):
        """预先打开聊天窗口并输入消息内容，只差发送"""
        try:
            # 保存当前活动窗口以便后续恢复
            self.current_window = pyautogui.getActiveWindow()
            self.log_signal.emit("保存当前窗口状态")

            # 打开企业微信
            self.log_signal.emit(
                f"模拟按下 {self.shortcuts['open_wechat']} 打开应用"
            )
            keyboard.press_and_release(self.shortcuts["open_wechat"])
            time.sleep(0.8)  # 等待搜索框出现
            # 按下Alt+F打开搜索框
            self.log_signal.emit(f"模拟按下 {self.shortcuts['open_search']} 打开搜索框")
            keyboard.press_and_release(self.shortcuts["open_search"])
            time.sleep(0.8)  # 等待搜索框出现

            # 输入目标对话名称
            self.log_signal.emit(f"输入目标对话: {self.target}")
            keyboard.write(self.target)
            time.sleep(1)  # 等待搜索结果

            # 按Enter选择第一个结果
            self.log_signal.emit("模拟按下 Enter 选择对话")
            keyboard.press_and_release("enter")
            time.sleep(1)  # 等待对话打开

            # 输入消息内容并保持分段
            self.log_signal.emit("开始输入消息内容")
            for line in self.content.split("\n"):
                keyboard.write(line)
                # Shift+Enter换行
                keyboard.press_and_release("shift+enter")
                time.sleep(0.1)

            # 删除最后一个多余的换行
            keyboard.press_and_release("backspace")
            time.sleep(0.1)

            self.log_signal.emit("消息准备完成，等待发送时机")
            return True

        except Exception as e:
            error_msg = f"准备消息出错: {str(e)}"
            self.log_signal.emit(error_msg)
            self.activate_previous_window()
            return False

    def message_schedule(self, target_time):
        """准备发送并启动精准定时"""
        if not self.is_running:

            return

        self.log_signal.emit("开始准备消息...")
        if self.message_prepare() and self.is_running:
            # 计算剩余时间
            remaining_ms = int((target_time - datetime.now()).total_seconds() * 1000)
            if remaining_ms > 0:
                self.log_signal.emit(
                    f"消息准备完成，等待 {remaining_ms:.0f} 毫秒后发送"
                )
                self.start_timer_signal.emit(
                    remaining_ms
                )  # 发射信号，通知主线程启动定时器
            else:
                self.log_signal.emit("立即发送消息（准备时间过长）")
                self.message_send()

    def message_send(self):
        """在精确时间执行发送操作"""
        SEND_SUCCESS = True
        try:
            self.log_signal.emit("正在执行发送操作...")
            self.log_signal.emit(f"模拟按下 {self.shortcuts['send_message']} 发送消息")
            keyboard.press_and_release(self.shortcuts["send_message"])
            send_time = datetime.now().strftime("%H:%M:%S.%f")[:-3]
            status_msg = f"消息已于 ({send_time}) 发送"
            time.sleep(0.5)

            self.log_signal.emit(status_msg)

        except Exception as e:
            error_msg = f"发送出错: {str(e)}"
            self.log_signal.emit(error_msg)
            SEND_SUCCESS = False

        # 如果是一次性任务，停止调度器并更新状态
        if self.once_schedule:
            self.log_signal.emit("一次性任务完成，停止定时器")
            self.timer.stop()  # 停止 QTimer
            self.is_running = False
            self.status_signal.emit(IDLE)

        # 恢复之前的活动窗口
        self.activate_previous_window()
        return SEND_SUCCESS

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

    def start_timer(self, remaining_ms):
        """在主线程中启动定时器"""
        self.log_signal.emit(f"主线程启动定时器，等待 {remaining_ms} 毫秒后发送消息")
        self.timer.start(remaining_ms)
