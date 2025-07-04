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
import win32gui

import keyboard
from PyQt5.QtCore import QTimer, QObject, pyqtSignal, QMutex, QMutexLocker

IDLE = 0
RUNNING = 1
STATUS = {IDLE: "空闲中", RUNNING: "运行中"}


class WeChatScheduler(QObject):
    log_signal = pyqtSignal(str)  # 日志信号
    status_signal = pyqtSignal(int)  # 状态信号

    # 新增信号用于主线程定时器操作
    start_timer_signal = pyqtSignal(int)

    def __init__(self):
        super().__init__()
        self.scheduled_time = None
        self.timer = QTimer()
        self.timer.timeout.connect(self.message_send)
        self.is_running = False
        self.schedule_thread = None
        self.current_window = None
        self.once_schedule = False
        # 消息内容缓存
        self.target = ""
        self.content = ""
        self.window_title = "企业微信"
        self.shortcuts = {
            "open_search": "alt+s",
            "send_message": "ctrl+enter",
        }
        self.delays = {
            "prepare_pre_time": 10.0,  # 消息准备提前时间
            "window_active_delay": 1.0,  # 窗口激活等待延时
            "search_delay": 1.0,
            "search_result_delay": 1.5,
            "chat_delay": 1,
            "line_delay": 0.1,
        }

        # 添加线程锁
        self.mutex = QMutex()
        self.running_mutex = QMutex()  # 专门用于控制运行状态的锁

        # 信号连接到主线程的定时器启动槽
        self.start_timer_signal.connect(self._start_timer_mainthread)

    def _start_timer_mainthread(self, ms):
        self.timer.setSingleShot(True)
        self.timer.start(ms)

    def update_shortcuts(self, shortcuts):
        """更新快捷键设置"""
        self.shortcuts.update(shortcuts)
        self.log_signal.emit(f"快捷键已更新")

    def update_delays(self, delays):
        """更新延时设置"""
        self.delays.update(delays)
        self.log_signal.emit(f"延时设置已更新")


    def update_window_title(self, title):
        self.window_title = title
        self.log_signal.emit(f"窗口标题已更新为: {title}")

    def start_once_schedule(self, target, content, scheduled_time):
        """启动一次性定时任务（多线程优化）"""
        def run_once():
            with QMutexLocker(self.mutex):
                self.once_schedule = True
                if not self.validate_inputs(target, content, scheduled_time):
                    return
                self.target = target
                self.content = content
                self.scheduled_time = scheduled_time
            with QMutexLocker(self.running_mutex):
                self.is_running = True
            self.status_signal.emit(RUNNING)
            self.log_signal.emit("启动一次性定时任务")
            now = datetime.now()
            seconds_remain = (self.scheduled_time - now).total_seconds()
            pre_time = self.delays.get("prepare_pre_time", 10.0)
            if seconds_remain > pre_time:
                delay_ms = int((seconds_remain - pre_time) * 1000)
                self.log_signal.emit(
                    f"现在是 {datetime.now()}, 将在 {delay_ms/1000.0:.2f} 秒后开始准备消息"
                )
                timer = threading.Timer((seconds_remain - pre_time), lambda: self.message_schedule(self.scheduled_time))
                timer.daemon = True
                timer.start()
            else:
                self.log_signal.emit("时间紧张，立即开始准备消息")
                self.message_schedule(self.scheduled_time)
        t = threading.Thread(target=run_once)
        t.daemon = True
        t.start()

    def start_repeating_schedule(self, target, content, days, send_time, repeat_type):
        """启动循环定时任务"""
        if self.schedule_thread and self.schedule_thread.is_alive():
            self.log_signal.emit("循环定时任务已在运行")
            return
        if not self.validate_repeat_inputs(target, content, days, repeat_type):
            return

        with QMutexLocker(self.mutex):
            self.target = target
            self.content = content
            self.once_schedule = False  # 设置非一次性任务

        with QMutexLocker(self.running_mutex):
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
        """循环定时任务线程（异步优化）"""
        self._async_log("循环定时任务线程已启动")
        pre_time = self.delays.get("prepare_pre_time", 10.0)
        try:
            while True:
                with QMutexLocker(self.running_mutex):
                    if not self.is_running:
                        break
                now = datetime.now()
                check_interval = 15
                day_match = False
                if repeat_type == 0:
                    day_match = now.day in days
                elif repeat_type == 1:
                    day_match = now.weekday() in days
                if not day_match:
                    time.sleep(check_interval)
                    continue
                target_datetime = datetime.combine(now.date(), send_time)
                if now >= target_datetime:
                    self._async_log("今天的目标时间已过，等待下一天")
                    time.sleep(60)
                    continue
                wait_seconds = (target_datetime - now).total_seconds()
                if wait_seconds < 300:
                    self._async_log(f"当前时间: {now}, 目标时间: {target_datetime}")
                    self._async_log(f"距离目标时间还有 {wait_seconds:.1f} 秒")
                if wait_seconds > pre_time:
                    if wait_seconds > 60:
                        time.sleep(min(wait_seconds - pre_time - 30, check_interval))
                        continue
                    delay_ms = int((wait_seconds - pre_time) * 1000)
                    self._async_log(f"等待 {delay_ms/1000:.1f} 秒后开始准备消息")
                    # 通过信号让主线程启动定时器
                    self.start_timer_signal.emit(delay_ms)
                    # 定时器到时后执行 message_schedule
                    self.timer.timeout.disconnect()
                    self.timer.timeout.connect(lambda: self.message_schedule(target_datetime))
                else:
                    self._async_log("时间紧张，立即开始准备消息")
                    # 通过信号让主线程立即执行
                    self.start_timer_signal.emit(0)
                    self.timer.timeout.disconnect()
                    self.timer.timeout.connect(lambda: self.message_schedule(target_datetime))
                wait_count = 60
                for _ in range(wait_count):
                    with QMutexLocker(self.running_mutex):
                        if not self.is_running:
                            self._async_log("任务已停止，退出循环")
                            return
                    time.sleep(1)
        except Exception as e:
            self._async_log(f"循环定时任务线程出错: {str(e)}")
        finally:
            self._async_log("循环定时任务线程已退出")

    def stop_scheduler(self):
        """停止所有定时任务"""
        with QMutexLocker(self.running_mutex):
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
        """立即发送消息，不影响定时任务运行状态（多线程优化）"""
        def send_job():
            if not target or not content:
                self.log_signal.emit("错误: 目标或内容为空")
                return
            with QMutexLocker(self.mutex):
                original_target = self.target
                original_content = self.content
                self.target = target
                self.content = content
            self.status_signal.emit(RUNNING)
            self.log_signal.emit("开始立即发送流程...")
            if self.message_prepare():
                time.sleep(0.1)
                self.message_send()
            with QMutexLocker(self.mutex):
                self.target = original_target
                self.content = original_content
            with QMutexLocker(self.running_mutex):
                running_state = self.is_running
            if running_state:
                self.status_signal.emit(RUNNING)
            else:
                self.status_signal.emit(IDLE)
        t = threading.Thread(target=send_job)
        t.daemon = True
        t.start()

    def message_prepare(self):
        """预先打开聊天窗口并输入消息内容，只差发送（异步优化）"""
        result = {'success': False}
        def prepare_job():
            max_retries = 3
            retry_count = 0
            self._async_log(f"开始准备消息 - 目标: {self.target}, 内容长度: {len(self.content)} 字符")
            while retry_count < max_retries:
                try:
                    wechat_window = win32gui.FindWindow(None, self.window_title)
                    if wechat_window:
                        self._async_log(f"找到{self.window_title}窗口，尝试激活")
                        win32gui.SetForegroundWindow(wechat_window)
                        time.sleep(self.delays["window_active_delay"])
                        self._async_log(f"窗口已等待激活 {self.delays['window_active_delay']}秒")
                    else:
                        self._async_log(f"警告: 未找到{self.window_title}窗口")
                    self._async_log(f"[第{retry_count+1}次尝试] 模拟按下 {self.shortcuts['open_search']} 打开搜索框")
                    keyboard.press_and_release(self.shortcuts["open_search"])
                    time.sleep(self.delays["search_delay"])
                    self._async_log(f"搜索框已等待 {self.delays['search_delay']}秒")
                    self._async_log(f"输入目标对话: {self.target}")
                    keyboard.write(self.target)
                    time.sleep(self.delays["search_result_delay"])
                    self._async_log(f"搜索结果已等待 {self.delays['search_result_delay']}秒")
                    self._async_log("模拟按下 Enter 选择对话")
                    keyboard.press_and_release("enter")
                    time.sleep(self.delays["chat_delay"])
                    self._async_log(f"聊天窗口已等待 {self.delays['chat_delay']}秒")
                    chat_window = win32gui.GetForegroundWindow()
                    if not chat_window:
                        self._async_log("警告: 未能获取聊天窗口句柄")
                        time.sleep(0.5)
                    self._async_log("开始输入消息内容，共 {} 行".format(len(self.content.split("\n"))))
                    for i, line in enumerate(self.content.split("\n")):
                        try:
                            keyboard.write(line)
                            time.sleep(0.1)
                            keyboard.press_and_release("shift+enter")
                            time.sleep(self.delays["line_delay"])
                            if (i + 1) % 5 == 0:
                                self._async_log(f"已输入 {i+1}/{len(self.content.split('\\n'))} 行")
                        except Exception as line_e:
                            self._async_log(f"输入第 {i+1} 行时出错: {str(line_e)}，尝试继续")
                            time.sleep(0.5)
                    self._async_log("消息准备完成，等待发送时机")
                    result['success'] = True
                    return
                except Exception as e:
                    retry_count += 1
                    error_msg = f"准备消息出错: {str(e)}, 错误类型: {type(e).__name__}, 重试 ({retry_count}/{max_retries})"
                    self._async_log(error_msg)
                    time.sleep(1.0)
            self._async_log(f"消息准备失败，已达到最大重试次数 {max_retries}")
            result['success'] = False
        t = threading.Thread(target=prepare_job)
        t.daemon = True
        t.start()
        t.join()  # 等待异步准备完成
        return result['success']

    def _async_log(self, message):
        """日志异步写入"""
        threading.Thread(target=lambda: self.log_signal.emit(message), daemon=True).start()

    def message_schedule(self, target_time):
        """准备发送并启动精准定时"""
        start_time = datetime.now()
        self.log_signal.emit(
            f"[{start_time.strftime('%H:%M:%S.%f')[:-3]}] 消息调度开始，目标时间: {target_time.strftime('%H:%M:%S.%f')[:-3]}"
        )

        with QMutexLocker(self.running_mutex):
            if not self.is_running:
                self.log_signal.emit("任务已停止，取消消息调度")
                return

        self.log_signal.emit("开始准备消息流程...")
        try:
            # 消息准备
            prepare_start = datetime.now()
            self.log_signal.emit(
                f"[{prepare_start.strftime('%H:%M:%S.%f')[:-3]}] 开始执行消息准备"
            )
            if self.message_prepare():
                prepare_end = datetime.now()
                prepare_duration = (prepare_end - prepare_start).total_seconds()
                self.log_signal.emit(f"消息准备耗时: {prepare_duration:.3f}秒")

                # 计算剩余时间，使用毫秒级精度
                now = datetime.now()
                remaining_seconds = (target_time - now).total_seconds()
                remaining_ms = int(remaining_seconds * 1000)

                if remaining_ms > 0:
                    # 通过信号让主线程启动定时器
                    self.start_timer_signal.emit(remaining_ms)
                    self.timer.timeout.disconnect()
                    self.timer.timeout.connect(self.message_send)
                else:
                    self.log_signal.emit(
                        f"[{now.strftime('%H:%M:%S.%f')[:-3]}] 已超过目标时间，立即发送消息"
                    )
                    self.message_send()
        except Exception as e:
            error_time = datetime.now().strftime("%H:%M:%S.%f")[:-3]
            self.log_signal.emit(f"[{error_time}] 消息准备失败: {str(e)}")
            import traceback

            self.log_signal.emit(f"详细错误: {traceback.format_exc()}")

    def message_send(self):
        """在精确时间执行发送操作（异步优化）"""
        def send_job():
            max_retries = 3
            retry_count = 0
            while retry_count < max_retries:
                try:
                    self._async_log(f"[第{retry_count+1}次尝试] 正在执行发送操作...")
                    self._async_log(f"模拟按下 {self.shortcuts['send_message']} 发送消息")
                    keyboard.press_and_release(self.shortcuts["send_message"])
                    send_time = datetime.now().strftime("%H:%M:%S.%f")[:-3]
                    status_msg = f"消息已于 ({send_time}) 发送完成"
                    self._async_log(status_msg)
                    self.timer.stop()
                    break
                except Exception as e:
                    retry_count += 1
                    error_msg = f"发送出错: {str(e)}, 错误类型: {type(e).__name__}, 重试 ({retry_count}/{max_retries})"
                    self._async_log(error_msg)
                    if retry_count < max_retries:
                        self._async_log(f"等待1秒后重试...")
                        time.sleep(1.0)
                    else:
                        self._async_log("发送失败，已达到最大重试次数")
            with QMutexLocker(self.mutex):
                once_task = self.once_schedule
            if once_task:
                self._async_log("一次性任务完成，停止定时器")
                with QMutexLocker(self.running_mutex):
                    self.is_running = False
                self.status_signal.emit(IDLE)
        t = threading.Thread(target=send_job)
        t.daemon = True
        t.start()

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
