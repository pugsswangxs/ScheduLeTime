import sys
import json
import subprocess
from datetime import datetime, timedelta
from enum import Enum
from typing import Dict, List, Any, Optional

from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QHBoxLayout, QTableWidget, QTableWidgetItem,
                             QPushButton, QLabel, QLineEdit, QComboBox,
                             QTextEdit, QSpinBox, QCheckBox, QTimeEdit,
                             QSystemTrayIcon, QMenu, QDialog,
                             QFormLayout, QTabWidget, QMessageBox,
                             QHeaderView, QStyle, QAbstractItemView)
from PyQt6.QtCore import Qt, QTime, QDate, QTimer, pyqtSignal, QObject
from PyQt6.QtGui import QAction, QColor
import schedule
import threading
import time as time_module


class SignalHandler(QObject):
    refresh_tasks_signal = pyqtSignal()
    show_notification_signal = pyqtSignal(str, str)
    execute_task_signal = pyqtSignal(object)


class TaskType(Enum):
    CMD = "CMD命令"
    NOTIFICATION = "提醒任务"


class PopupType(Enum):
    SYSTEM_TRAY = "右下角弹窗"
    WINDOW_POPUP = "窗口弹窗"


class TaskStatus(Enum):
    ENABLED = "启用"
    DISABLED = "禁用"


class Task:
    def __init__(self):
        self.id = str(int(time_module.time() * 1000))
        self.name = ""
        self.description = ""
        self.task_type = TaskType.NOTIFICATION  # 修改默认值为提醒任务
        self.status = TaskStatus.ENABLED
        self.schedule_type = "interval"
        self.interval_seconds = 60
        self.daily_time = QTime.currentTime()
        self.weekly_day = 0  # 0-6, Monday to Sunday
        self.monthly_day = 1  # 1-31, day of month
        self.start_date = QDate.currentDate()
        self.end_date = QDate.currentDate().addYears(1)
        self.cmd_command = ""
        self.notification_title = ""
        self.notification_content = ""
        self.notification_timeout = 3000  # 添加弹窗显示时间属性，默认3秒
        self.popup_type = "system_tray"  # 新增弹窗类型，默认系统托盘
        self.last_execution = None
        self.execution_count = 0
        self.retry_count = 0
        self.enable_logging = True

    def to_dict(self) -> Dict[str, Any]:
        return {
            "id": self.id,
            "name": self.name,
            "description": self.description,
            "task_type": self.task_type.value,
            "status": self.status.value,
            "schedule_type": self.schedule_type,
            "interval_seconds": self.interval_seconds,
            "daily_time": self.daily_time.toString("hh:mm"),
            "weekly_day": self.weekly_day,
            "monthly_day": self.monthly_day,
            "start_date": self.start_date.toString("yyyy-MM-dd"),
            "end_date": self.end_date.toString("yyyy-MM-dd"),
            "cmd_command": self.cmd_command,
            "notification_title": self.notification_title,
            "notification_content": self.notification_content,
            "notification_timeout": self.notification_timeout,  # 保存弹窗显示时间
            "popup_type": self.popup_type,
            "last_execution": self.last_execution.isoformat() if self.last_execution else None,
            "execution_count": self.execution_count,
            "retry_count": self.retry_count,
            "enable_logging": self.enable_logging
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'Task':
        task = cls()
        task.id = data.get("id", str(int(time_module.time() * 1000)))
        task.name = data.get("name", "")
        task.description = data.get("description", "")

        # 数据迁移：将旧的窗口弹窗提醒类型转换为提醒任务
        raw_task_type = data.get("task_type", TaskType.NOTIFICATION.value)
        if raw_task_type in [TaskType.NOTIFICATION.value, "窗口弹窗提醒"]:
            task.task_type = TaskType.NOTIFICATION
        elif raw_task_type == TaskType.CMD.value:
            task.task_type = TaskType.CMD
        else:
            task.task_type = TaskType.NOTIFICATION  # 默认值

        task.status = TaskStatus(data.get("status", TaskStatus.ENABLED.value))
        task.schedule_type = data.get("schedule_type", "interval")
        task.interval_seconds = data.get("interval_seconds", 60)
        task.daily_time = QTime.fromString(data.get("daily_time", "00:00"), "hh:mm")
        task.weekly_day = data.get("weekly_day", 0)
        task.monthly_day = data.get("monthly_day", 1)  # 新增每月执行日期
        task.start_date = QDate.fromString(data.get("start_date", QDate.currentDate().toString("yyyy-MM-dd")),
                                           "yyyy-MM-dd")
        task.end_date = QDate.fromString(data.get("end_date", QDate.currentDate().addYears(1).toString("yyyy-MM-dd")),
                                         "yyyy-MM-dd")
        task.cmd_command = data.get("cmd_command", "")
        task.notification_title = data.get("notification_title", "")
        task.notification_content = data.get("notification_content", "")
        task.notification_timeout = data.get("notification_timeout", 3000)  # 加载弹窗显示时间
        task.popup_type = data.get("popup_type", "system_tray")  # 加载弹窗类型
        if data.get("last_execution"):
            task.last_execution = datetime.fromisoformat(data["last_execution"])
        task.execution_count = data.get("execution_count", 0)
        task.retry_count = data.get("retry_count", 0)
        task.enable_logging = data.get("enable_logging", True)
        return task

    def get_schedule_description(self) -> str:
        if self.schedule_type == "interval":
            if self.interval_seconds < 60:
                return f"每{self.interval_seconds}秒"
            elif self.interval_seconds < 3600:
                minutes = self.interval_seconds // 60
                seconds = self.interval_seconds % 60
                if seconds == 0:
                    return f"每{minutes}分钟"
                else:
                    return f"每{minutes}分{seconds}秒"
            else:
                hours = self.interval_seconds // 3600
                minutes = (self.interval_seconds % 3600) // 60
                seconds = self.interval_seconds % 60
                if minutes == 0 and seconds == 0:
                    return f"每{hours}小时"
                elif seconds == 0:
                    return f"每{hours}小时{minutes}分钟"
                else:
                    return f"每{hours}小时{minutes}分{seconds}秒"
        elif self.schedule_type == "daily":
            return f"每天 {self.daily_time.toString('hh:mm')}"
        elif self.schedule_type == "weekly":
            days = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]
            return f"每周{days[self.weekly_day]} {self.daily_time.toString('hh:mm')}"
        elif self.schedule_type == "monthly":
            return f"每月{self.monthly_day}日 {self.daily_time.toString('hh:mm')}"
        return "未知"

    def get_next_run_time(self) -> str:
        if self.status != TaskStatus.ENABLED:
            return "未启用"

        try:
            now = datetime.now()

            if self.schedule_type == "interval":
                if self.last_execution:
                    next_run = self.last_execution + timedelta(seconds=self.interval_seconds)
                    return next_run.strftime("%Y-%m-%d %H:%M:%S")
                else:
                    next_run = now + timedelta(seconds=self.interval_seconds)
                    return next_run.strftime("%Y-%m-%d %H:%M:%S")

            elif self.schedule_type == "daily":
                time_parts = self.daily_time.toString("HH:mm").split(":")
                hour, minute = int(time_parts[0]), int(time_parts[1])
                next_run = now.replace(hour=hour, minute=minute, second=0, microsecond=0)

                if next_run <= now:
                    next_run += timedelta(days=1)
                return next_run.strftime("%Y-%m-%d %H:%M:%S")

            elif self.schedule_type == "weekly":
                time_parts = self.daily_time.toString("HH:mm").split(":")
                hour, minute = int(time_parts[0]), int(time_parts[1])
                days = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]
                target_weekday = self.weekly_day

                days_ahead = target_weekday - now.weekday()
                if days_ahead <= 0:
                    days_ahead += 7

                next_run = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
                next_run += timedelta(days_ahead)

                if target_weekday == now.weekday() and next_run <= now:
                    next_run += timedelta(days=7)

                return next_run.strftime("%Y-%m-%d %H:%M:%S")

            elif self.schedule_type == "monthly":
                time_parts = self.daily_time.toString("HH:mm").split(":")
                hour, minute = int(time_parts[0]), int(time_parts[1])
                next_run = now.replace(day=self.monthly_day, hour=hour, minute=minute, second=0, microsecond=0)

                if next_run <= now:
                    # 如果本月日期已过，移到下个月
                    if now.month == 12:
                        next_run = next_run.replace(year=now.year + 1, month=1)
                    else:
                        next_run = next_run.replace(month=now.month + 1)
                return next_run.strftime("%Y-%m-%d %H:%M:%S")

        except Exception:
            return "计算错误"

        return "未知"


class PopupDialog(QDialog):
    def __init__(self, title: str, content: str, timeout: int = 3000):
        super().__init__()
        self.title = title
        self.content = content
        self.timeout = timeout
        self.setup_ui()

    def setup_ui(self):
        self.setWindowTitle(self.title)
        self.setModal(False)  # 非模态窗口
        self.setWindowFlags(Qt.WindowType.Tool | Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint)

        layout = QVBoxLayout()

        # 标题
        title_label = QLabel(self.title)
        title_label.setStyleSheet("font-size: 14px; font-weight: bold; color: #333;")
        layout.addWidget(title_label)

        # 内容
        content_label = QLabel(self.content)
        content_label.setWordWrap(True)
        layout.addWidget(content_label)

        # 确认按钮
        confirm_button = QPushButton("确认")
        confirm_button.clicked.connect(self.accept)
        layout.addWidget(confirm_button)

        self.setLayout(layout)

        # 设置窗口大小和位置 - 居中显示
        self.resize(300, 150)
        screen_geometry = QApplication.primaryScreen().availableGeometry()
        x = (screen_geometry.width() - self.width()) // 2
        y = (screen_geometry.height() - self.height()) // 2
        self.move(x, y)

        # 自动关闭
        if self.timeout > 0:
            QTimer.singleShot(self.timeout, self.accept)


class TaskEditDialog(QDialog):
    def __init__(self, task: Optional[Task] = None, parent=None):
        super().__init__(parent)
        self.task = task if task else Task()
        self.setup_ui()
        self.load_task_data()

    def setup_ui(self):
        self.setWindowTitle("编辑任务")
        self.setModal(True)
        self.resize(600, 500)  # 增大窗口尺寸

        # 设置全局字体大小
        font = self.font()
        font.setPointSize(10)  # 增大字体
        self.setFont(font)

        tab_widget = QTabWidget()
        layout = QVBoxLayout()
        layout.addWidget(tab_widget)

        # 基本信息
        basic_widget = QWidget()
        basic_layout = QFormLayout()
        # 设置表单布局的间距和边距
        basic_layout.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        basic_layout.setHorizontalSpacing(15)
        basic_layout.setVerticalSpacing(10)

        self.name_edit = QLineEdit()
        self.name_edit.setMinimumHeight(25)  # 增大控件高度
        self.name_edit.addAction(
            self.style().standardIcon(QStyle.StandardPixmap.SP_FileIcon),
            QLineEdit.ActionPosition.LeadingPosition
        )

        self.description_edit = QTextEdit()
        self.description_edit.setMaximumHeight(100)  # 增大高度
        self.description_edit.setMinimumHeight(80)

        self.type_combo = QComboBox()
        self.type_combo.setMinimumHeight(25)
        self.type_combo.addItems([t.value for t in TaskType])
        self.type_combo.setItemIcon(0, self.style().standardIcon(QStyle.StandardPixmap.SP_CommandLink))
        self.type_combo.setItemIcon(1, self.style().standardIcon(QStyle.StandardPixmap.SP_MessageBoxInformation))

        basic_layout.addRow("任务名称:", self.name_edit)
        basic_layout.addRow("任务描述:", self.description_edit)
        basic_layout.addRow("任务类型:", self.type_combo)

        basic_widget.setLayout(basic_layout)
        tab_widget.addTab(basic_widget, "基本信息")

        # 定时配置
        schedule_widget = QWidget()
        schedule_layout = QFormLayout()
        schedule_layout.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        schedule_layout.setHorizontalSpacing(15)
        schedule_layout.setVerticalSpacing(10)

        self.schedule_type_combo = QComboBox()
        self.schedule_type_combo.setMinimumHeight(25)
        self.schedule_type_combo.addItems(["固定间隔", "每日", "每周", "每月"])
        self.schedule_type_combo.setItemIcon(0, self.style().standardIcon(QStyle.StandardPixmap.SP_ArrowUp))
        self.schedule_type_combo.setItemIcon(1, self.style().standardIcon(QStyle.StandardPixmap.SP_DialogOkButton))
        self.schedule_type_combo.setItemIcon(2, self.style().standardIcon(QStyle.StandardPixmap.SP_DialogOpenButton))
        self.schedule_type_combo.setItemIcon(3, self.style().standardIcon(QStyle.StandardPixmap.SP_DialogOpenButton))

        self.interval_spin = QSpinBox()
        self.interval_spin.setMinimumHeight(25)
        self.interval_spin.setRange(1, 100000)

        self.interval_unit_combo = QComboBox()
        self.interval_unit_combo.setMinimumHeight(25)
        self.interval_unit_combo.addItems(["秒", "分钟", "小时"])

        self.daily_time_edit = QTimeEdit()
        self.daily_time_edit.setMinimumHeight(25)
        self.daily_time_edit.setTime(QTime.currentTime())

        self.weekly_combo = QComboBox()
        self.weekly_combo.setMinimumHeight(25)
        self.weekly_combo.addItems(["周一", "周二", "周三", "周四", "周五", "周六", "周日"])

        self.monthly_day_spin = QSpinBox()
        self.monthly_day_spin.setMinimumHeight(25)
        self.monthly_day_spin.setRange(1, 31)
        self.monthly_day_spin.setValue(1)

        schedule_layout.addRow("定时类型:", self.schedule_type_combo)
        schedule_layout.addRow("间隔时间:", self.interval_spin)
        schedule_layout.addRow("时间单位:", self.interval_unit_combo)
        schedule_layout.addRow("执行时间:", self.daily_time_edit)
        schedule_layout.addRow("星期:", self.weekly_combo)
        schedule_layout.addRow("每月日期:", self.monthly_day_spin)

        schedule_widget.setLayout(schedule_layout)
        tab_widget.addTab(schedule_widget, "定时配置")

        # 任务内容
        content_widget = QWidget()
        content_layout = QVBoxLayout()
        content_layout.setSpacing(10)

        self.cmd_text = QTextEdit()
        self.cmd_text.setMinimumHeight(100)
        self.cmd_text.setPlaceholderText("输入要执行的CMD命令...")

        self.notification_title_edit = QLineEdit()
        self.notification_title_edit.setMinimumHeight(25)
        self.notification_title_edit.setPlaceholderText("提醒标题...")

        self.notification_content_edit = QTextEdit()
        self.notification_content_edit.setMinimumHeight(100)
        self.notification_content_edit.setMaximumHeight(150)
        self.notification_content_edit.setPlaceholderText("提醒内容...")

        # 添加弹窗类型选择
        self.popup_type_combo = QComboBox()
        self.popup_type_combo.setMinimumHeight(25)
        self.popup_type_combo.addItems([PopupType.SYSTEM_TRAY.value, PopupType.WINDOW_POPUP.value])

        # 添加弹窗显示时间输入框
        self.notification_timeout_spin = QSpinBox()
        self.notification_timeout_spin.setMinimumHeight(25)
        self.notification_timeout_spin.setRange(1, 60000)  # 1毫秒到60秒
        self.notification_timeout_spin.setValue(3000)  # 默认3秒
        self.notification_timeout_spin.setSuffix(" 毫秒")

        content_layout.addWidget(QLabel("CMD命令:"))
        content_layout.addWidget(self.cmd_text)
        content_layout.addWidget(QLabel("提醒标题:"))
        content_layout.addWidget(self.notification_title_edit)
        content_layout.addWidget(QLabel("提醒内容:"))
        content_layout.addWidget(self.notification_content_edit)
        content_layout.addWidget(QLabel("弹窗类型:"))
        content_layout.addWidget(self.popup_type_combo)
        content_layout.addWidget(QLabel("弹窗显示时间:"))
        content_layout.addWidget(self.notification_timeout_spin)

        content_widget.setLayout(content_layout)
        tab_widget.addTab(content_widget, "任务内容")

        # 高级选项
        advanced_widget = QWidget()
        advanced_layout = QFormLayout()
        advanced_layout.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        advanced_layout.setHorizontalSpacing(15)
        advanced_layout.setVerticalSpacing(10)

        self.enable_check = QCheckBox("启用任务")
        self.enable_check.setChecked(True)

        self.retry_spin = QSpinBox()
        self.retry_spin.setMinimumHeight(25)
        self.retry_spin.setRange(0, 10)

        self.logging_check = QCheckBox("记录执行日志")
        self.logging_check.setChecked(True)

        advanced_layout.addRow(self.enable_check)
        advanced_layout.addRow("重试次数:", self.retry_spin)
        advanced_layout.addRow(self.logging_check)

        advanced_widget.setLayout(advanced_layout)
        tab_widget.addTab(advanced_widget, "高级选项")

        # 按钮
        button_layout = QHBoxLayout()
        self.ok_button = QPushButton("确定")
        self.ok_button.setMinimumHeight(30)  # 增大按钮
        self.ok_button.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DialogOkButton))
        self.cancel_button = QPushButton("取消")
        self.cancel_button.setMinimumHeight(30)
        self.cancel_button.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DialogCancelButton))

        button_layout.addWidget(self.ok_button)
        button_layout.addWidget(self.cancel_button)
        layout.addLayout(button_layout)

        self.setLayout(layout)

        # 连接信号
        self.schedule_type_combo.currentTextChanged.connect(self.on_schedule_type_changed)
        self.type_combo.currentTextChanged.connect(self.on_task_type_changed)
        self.ok_button.clicked.connect(self.accept)
        self.cancel_button.clicked.connect(self.reject)

        self.on_schedule_type_changed(self.schedule_type_combo.currentText())
        self.on_task_type_changed(self.type_combo.currentText())

    def on_schedule_type_changed(self, schedule_type):
        self.interval_spin.setVisible(False)
        self.interval_unit_combo.setVisible(False)
        self.daily_time_edit.setVisible(False)
        self.weekly_combo.setVisible(False)
        self.monthly_day_spin.setVisible(False)

        if schedule_type == "固定间隔":
            self.interval_spin.setVisible(True)
            self.interval_unit_combo.setVisible(True)
        elif schedule_type == "每日":
            self.daily_time_edit.setVisible(True)
        elif schedule_type == "每周":
            self.daily_time_edit.setVisible(True)
            self.weekly_combo.setVisible(True)
        elif schedule_type == "每月":  # 新增每月执行配置
            self.daily_time_edit.setVisible(True)
            self.monthly_day_spin.setVisible(True)

    def on_task_type_changed(self, task_type):
        is_notification_task = task_type == TaskType.NOTIFICATION.value
        self.cmd_text.setVisible(task_type == TaskType.CMD.value)
        self.notification_title_edit.setVisible(is_notification_task)
        self.notification_content_edit.setVisible(is_notification_task)
        self.popup_type_combo.setVisible(is_notification_task)
        self.notification_timeout_spin.setVisible(is_notification_task)

    def load_task_data(self):
        self.name_edit.setText(self.task.name)
        self.description_edit.setPlainText(self.task.description)
        self.type_combo.setCurrentText(self.task.task_type.value)
        self.enable_check.setChecked(self.task.status == TaskStatus.ENABLED)
        self.retry_spin.setValue(self.task.retry_count)
        self.logging_check.setChecked(self.task.enable_logging)

        if self.task.schedule_type == "interval":
            self.schedule_type_combo.setCurrentText("固定间隔")
            self.interval_spin.setValue(self.task.interval_seconds)
        elif self.task.schedule_type == "daily":
            self.schedule_type_combo.setCurrentText("每日")
            self.daily_time_edit.setTime(self.task.daily_time)
        elif self.task.schedule_type == "weekly":
            self.schedule_type_combo.setCurrentText("每周")
            self.daily_time_edit.setTime(self.task.daily_time)
            self.weekly_combo.setCurrentIndex(self.task.weekly_day)
        elif self.task.schedule_type == "monthly":
            self.schedule_type_combo.setCurrentText("每月")
            self.daily_time_edit.setTime(self.task.daily_time)
            self.monthly_day_spin.setValue(self.task.monthly_day)

        self.cmd_text.setPlainText(self.task.cmd_command)
        self.notification_title_edit.setText(self.task.notification_title)
        self.notification_content_edit.setPlainText(self.task.notification_content)
        self.notification_timeout_spin.setValue(self.task.notification_timeout)  # 加载弹窗显示时间
        self.popup_type_combo.setCurrentText(self.task.popup_type)  # 加载弹窗类型

    def get_task_data(self):
        self.task.name = self.name_edit.text()
        self.task.description = self.description_edit.toPlainText()
        self.task.task_type = TaskType(self.type_combo.currentText())
        self.task.status = TaskStatus.ENABLED if self.enable_check.isChecked() else TaskStatus.DISABLED
        self.task.retry_count = self.retry_spin.value()
        self.task.enable_logging = self.logging_check.isChecked()

        schedule_type = self.schedule_type_combo.currentText()
        if schedule_type == "固定间隔":
            self.task.schedule_type = "interval"
            interval = self.interval_spin.value()
            unit = self.interval_unit_combo.currentText()
            if unit == "秒":
                self.task.interval_seconds = interval
            elif unit == "分钟":
                self.task.interval_seconds = interval * 60
            else:
                self.task.interval_seconds = interval * 3600
        elif schedule_type == "每日":
            self.task.schedule_type = "daily"
            self.task.daily_time = self.daily_time_edit.time()
        elif schedule_type == "每周":
            self.task.schedule_type = "weekly"
            self.task.daily_time = self.daily_time_edit.time()
            self.task.weekly_day = self.weekly_combo.currentIndex()
        elif schedule_type == "每月":  # 新增每月执行配置
            self.task.schedule_type = "monthly"
            self.task.daily_time = self.daily_time_edit.time()
            self.task.monthly_day = self.monthly_day_spin.value()

        self.task.cmd_command = self.cmd_text.toPlainText()
        self.task.notification_title = self.notification_title_edit.text()
        self.task.notification_content = self.notification_content_edit.toPlainText()
        self.task.notification_timeout = self.notification_timeout_spin.value()  # 保存弹窗显示时间
        self.task.popup_type = self.popup_type_combo.currentText()  # 保存弹窗类型

        return self.task


class TaskManager(QMainWindow):
    def __init__(self):
        super().__init__()
        self.tasks: List[Task] = []
        self.scheduler_thread = None
        self.scheduler_running = False
        self.is_minimized_to_tray = False

        self.signal_handler = SignalHandler()
        self.signal_handler.refresh_tasks_signal.connect(self.refresh_tasks)
        self.signal_handler.show_notification_signal.connect(self.show_notification)
        self.signal_handler.execute_task_signal.connect(self.execute_task)

        self.refresh_timer = QTimer()
        self.refresh_timer.timeout.connect(self.refresh_next_run_times)
        self.refresh_timer.start(30000)

        self.setStyleSheet("""
            QMainWindow {
                background-color: #f0f0f0;
            }
            QGroupBox {
                font-weight: bold;
                border: 1px solid #cccccc;
                border-radius: 5px;
                margin-top: 1ex;
                padding-top: 10px;
            }
            QGroupBox::title {
                subline-offset: -2px;
                padding: 0 5px;
                background-color: #f0f0f0;
            }
            QPushButton {
                background-color: #4CAF50;
                border: none;
                color: white;
                padding: 8px 16px;
                text-align: center;
                text-decoration: none;
                font-size: 14px;
                border-radius: 4px;
                min-width: 80px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:pressed {
                background-color: #3d8b40;
            }
            QPushButton#delete_button {
                background-color: #f44336;
            }
            QPushButton#delete_button:hover {
                background-color: #d32f2f;
            }
            QPushButton#delete_button:pressed {
                background-color: #b71c1c;
            }
            QLineEdit, QTextEdit, QComboBox, QSpinBox, QTimeEdit {
                padding: 5px;
                border: 1px solid #cccccc;
                border-radius: 4px;
            }
            QTableWidget {
                alternate-background-color: #f9f9f9;
                background-color: white;
            }
            QHeaderView::section {
                background-color: #e0e0e0;
                padding: 4px;
                border: 1px solid #cccccc;
                font-weight: bold;
            }
        """)

        self.setup_ui()
        self.setup_tray()
        self.load_tasks()
        self.start_scheduler()

    def setup_ui(self):
        self.setWindowTitle("定时任务管理器")
        self.setGeometry(100, 100, 1000, 700)
        self.setWindowIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_TitleBarMenuButton))
        self.setWindowFlags(Qt.WindowType.WindowCloseButtonHint | Qt.WindowType.WindowMinimizeButtonHint)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # 工具栏
        toolbar_layout = QHBoxLayout()

        self.new_button = QPushButton("新建任务")
        self.new_button.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_FileIcon))

        self.edit_button = QPushButton("编辑任务")
        self.edit_button.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_FileDialogDetailedView))

        self.delete_button = QPushButton("删除任务")
        self.delete_button.setObjectName("delete_button")
        self.delete_button.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_TrashIcon))

        self.refresh_button = QPushButton("刷新")
        self.refresh_button.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_BrowserReload))

        self.enable_button = QPushButton("启用任务")
        self.enable_button.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DialogOkButton))
        self.disable_button = QPushButton("禁用任务")
        self.disable_button.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DialogCancelButton))

        self.select_all_button = QPushButton("全选")
        self.select_all_button.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_ArrowRight))

        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("搜索任务...")
        self.search_edit.addAction(
            self.style().standardIcon(QStyle.StandardPixmap.SP_FileDialogContentsView),
            QLineEdit.ActionPosition.LeadingPosition
        )

        toolbar_layout.addWidget(self.new_button)
        toolbar_layout.addWidget(self.edit_button)
        toolbar_layout.addWidget(self.delete_button)
        toolbar_layout.addWidget(self.enable_button)
        toolbar_layout.addWidget(self.disable_button)
        toolbar_layout.addWidget(self.select_all_button)
        toolbar_layout.addWidget(self.refresh_button)
        toolbar_layout.addStretch()
        toolbar_layout.addWidget(QLabel("搜索:"))
        toolbar_layout.addWidget(self.search_edit)

        layout.addLayout(toolbar_layout)

        # 任务列表 (7列)
        self.task_table = QTableWidget()
        self.task_table.setColumnCount(7)  # 从6列增加到7列
        self.task_table.setHorizontalHeaderLabels(
            ["任务名称", "类型", "弹窗类型", "定时规则", "状态", "上次执行", "下次执行"])
        self.task_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        # 启用拖拽排序
        self.task_table.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        self.task_table.setDragEnabled(True)
        self.task_table.setAcceptDrops(True)
        self.task_table.setDropIndicatorShown(True)
        # 连接拖拽事件
        self.task_table.model().rowsMoved.connect(self.on_task_order_changed)
        layout.addWidget(self.task_table)

        # 状态栏
        self.status_label = QLabel("就绪")
        self.statusBar().addWidget(self.status_label)

        # 连接信号
        self.new_button.clicked.connect(self.new_task)
        self.edit_button.clicked.connect(self.edit_task)
        self.delete_button.clicked.connect(self.delete_task)
        self.enable_button.clicked.connect(self.enable_task)
        self.disable_button.clicked.connect(self.disable_task)
        self.select_all_button.clicked.connect(self.select_all_tasks)
        self.refresh_button.clicked.connect(self.refresh_tasks)
        self.search_edit.textChanged.connect(self.filter_tasks)
        self.task_table.doubleClicked.connect(self.edit_task_on_double_click)

    def edit_task_on_double_click(self, index):
        current_row = index.row()
        if current_row >= 0:
            task_id = self.task_table.item(current_row, 0).data(Qt.ItemDataRole.UserRole)
            task = next((t for t in self.tasks if t.id == task_id), None)
            if task:
                dialog = TaskEditDialog(task)
                if dialog.exec() == QDialog.DialogCode.Accepted:
                    dialog.get_task_data()
                    self.save_tasks()
                    self.refresh_tasks()
                    self.status_label.setText("任务更新成功")

    def setup_tray(self):
        self.tray_icon = QSystemTrayIcon(self)
        self.tray_icon.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_TitleBarMenuButton))
        self.tray_icon.setToolTip("定时任务管理器")

        tray_menu = QMenu()

        show_action = QAction("显示主窗口", self)
        show_action.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_TitleBarNormalButton))
        show_action.triggered.connect(self.show)

        pause_action = QAction("暂停所有任务", self)
        pause_action.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_MediaPause))
        pause_action.triggered.connect(self.pause_all_tasks)

        resume_action = QAction("恢复所有任务", self)
        resume_action.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_MediaPlay))
        resume_action.triggered.connect(self.resume_all_tasks)

        quit_action = QAction("退出", self)
        quit_action.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DialogCloseButton))
        quit_action.triggered.connect(self.quit_application)

        tray_menu.addAction(show_action)
        tray_menu.addAction(pause_action)
        tray_menu.addAction(resume_action)
        tray_menu.addSeparator()
        tray_menu.addAction(quit_action)

        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.activated.connect(self.tray_icon_activated)
        self.tray_icon.show()

    def tray_icon_activated(self, reason):
        if reason == QSystemTrayIcon.ActivationReason.DoubleClick:
            if self.isVisible():
                self.hide()
            else:
                self.show()
                self.activateWindow()

    def new_task(self):
        dialog = TaskEditDialog()
        if dialog.exec() == QDialog.DialogCode.Accepted:
            task = dialog.get_task_data()
            self.tasks.append(task)
            self.save_tasks()
            self.refresh_tasks()
            self.on_tasks_changed()
            self.status_label.setText("任务创建成功")

    def edit_task(self):
        current_row = self.task_table.currentRow()
        if current_row >= 0:
            task_id = self.task_table.item(current_row, 0).data(Qt.ItemDataRole.UserRole)
            task = next((t for t in self.tasks if t.id == task_id), None)
            if task:
                dialog = TaskEditDialog(task)
                if dialog.exec() == QDialog.DialogCode.Accepted:
                    dialog.get_task_data()
                    self.save_tasks()
                    self.refresh_tasks()
                    self.on_tasks_changed()
                    self.status_label.setText("任务更新成功")

    def get_selected_tasks(self):
        selected_rows = set()
        for item in self.task_table.selectedItems():
            selected_rows.add(item.row())

        selected_tasks = []
        for row in selected_rows:
            task_id = self.task_table.item(row, 0).data(Qt.ItemDataRole.UserRole)
            task = next((t for t in self.tasks if t.id == task_id), None)
            if task:
                selected_tasks.append(task)

        return selected_tasks

    def enable_task(self):
        selected_tasks = self.get_selected_tasks()
        if not selected_tasks:
            self.status_label.setText("请先选择任务")
            return

        for task in selected_tasks:
            task.status = TaskStatus.ENABLED

        self.save_tasks()
        self.refresh_tasks()
        self.on_tasks_changed()
        self.status_label.setText(f"已启用 {len(selected_tasks)} 个任务")

    def disable_task(self):
        selected_tasks = self.get_selected_tasks()
        if not selected_tasks:
            self.status_label.setText("请先选择任务")
            return

        for task in selected_tasks:
            task.status = TaskStatus.DISABLED

        self.save_tasks()
        self.refresh_tasks()
        self.on_tasks_changed()
        self.status_label.setText(f"已禁用 {len(selected_tasks)} 个任务")

    def select_all_tasks(self):
        self.task_table.selectAll()
        self.status_label.setText(f"已选择 {len(self.tasks)} 个任务")

    def delete_task(self):
        selected_tasks = self.get_selected_tasks()
        if not selected_tasks:
            self.status_label.setText("请先选择任务")
            return

        task_names = [task.name for task in selected_tasks]
        reply = QMessageBox.question(
            self,
            "确认删除",
            f"确定要删除选中的 {len(selected_tasks)} 个任务吗？\n{', '.join(task_names)}",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            for task in selected_tasks:
                self.tasks.remove(task)

            self.save_tasks()
            self.refresh_tasks()
            self.on_tasks_changed()
            self.status_label.setText(f"已删除 {len(selected_tasks)} 个任务")

    def refresh_tasks(self):
        self.task_table.setRowCount(0)
        current_time = datetime.now()

        for task in self.tasks:
            # 更新任务的上次执行时间为当前时间
            task.last_execution = current_time
            task.execution_count += 1  # 增加执行次数

            row = self.task_table.rowCount()
            self.task_table.insertRow(row)

            name_item = QTableWidgetItem(task.name)
            name_item.setData(Qt.ItemDataRole.UserRole, task.id)
            self.task_table.setItem(row, 0, name_item)

            type_item = QTableWidgetItem(task.task_type.value)
            self.task_table.setItem(row, 1, type_item)

            # 新增弹窗类型列 - 只对提醒任务显示弹窗类型
            if task.task_type == TaskType.NOTIFICATION:
                popup_type_item = QTableWidgetItem(task.popup_type)
            else:
                popup_type_item = QTableWidgetItem("-")
            self.task_table.setItem(row, 2, popup_type_item)

            rule_item = QTableWidgetItem(task.get_schedule_description())
            self.task_table.setItem(row, 3, rule_item)

            status_item = QTableWidgetItem(task.status.value)
            self.task_table.setItem(row, 4, status_item)

            # 显示更新后的上次执行时间
            last_exec = task.last_execution.strftime("%Y-%m-%d %H:%M:%S") if task.last_execution else "从未执行"
            last_exec_item = QTableWidgetItem(last_exec)
            self.task_table.setItem(row, 5, last_exec_item)

            # 显示更新后的下次执行时间
            next_run_item = QTableWidgetItem(task.get_next_run_time())
            self.task_table.setItem(row, 6, next_run_item)

            color = QColor(200, 255, 200) if task.status == TaskStatus.ENABLED else QColor(255, 200, 200)
            for col in range(self.task_table.columnCount()):
                self.task_table.item(row, col).setBackground(color)

        # 保存更新后的任务数据
        self.save_tasks()
        self.status_label.setText("任务已刷新，执行时间和下次执行时间已更新")

    def refresh_next_run_times(self):
        try:
            for row in range(self.task_table.rowCount()):
                if row < len(self.tasks):
                    task_id = self.task_table.item(row, 0).data(Qt.ItemDataRole.UserRole)
                    task = next((t for t in self.tasks if t.id == task_id), None)
                    if task and self.task_table.columnCount() > 6:  # 确保有第7列
                        next_run_item = QTableWidgetItem(task.get_next_run_time())
                        self.task_table.setItem(row, 6, next_run_item)  # 更新第7列(索引为6)

                        color = QColor(200, 255, 200) if task.status == TaskStatus.ENABLED else QColor(255, 200, 200)
                        self.task_table.item(row, 6).setBackground(color)
        except Exception:
            pass

    def filter_tasks(self):
        search_text = self.search_edit.text().lower()
        for row in range(self.task_table.rowCount()):
            task_name = self.task_table.item(row, 0).text().lower()
            task_type = self.task_table.item(row, 1).text().lower()
            should_show = search_text in task_name or search_text in task_type
            self.task_table.setRowHidden(row, not should_show)

    def on_task_order_changed(self, sourceParent, sourceStart, sourceEnd, destinationParent, destinationRow):
        # 获取拖拽的行和目标位置
        if sourceStart == destinationRow or destinationRow == sourceStart + 1:
            return  # 没有实际移动

        # 重新排序任务列表
        moved_task = self.tasks.pop(sourceStart)

        # 计算实际插入位置
        if destinationRow > sourceStart:
            destinationRow -= 1  # 因为删除了一个元素，所以目标位置前移

        self.tasks.insert(destinationRow, moved_task)
        self.save_tasks()
        self.refresh_tasks()
        self.status_label.setText("任务顺序已更新")

    def load_tasks(self):
        try:
            with open("tasks.json", "r", encoding="utf-8") as f:
                data = json.load(f)
                self.tasks = [Task.from_dict(task_data) for task_data in data]
        except FileNotFoundError:
            self.tasks = []
        self.refresh_tasks()

    def save_tasks(self):
        with open("tasks.json", "w", encoding="utf-8") as f:
            data = [task.to_dict() for task in self.tasks]
            json.dump(data, f, ensure_ascii=False, indent=2)

    def start_scheduler(self):
        self.scheduler_running = True
        self.scheduler_thread = threading.Thread(target=self.scheduler_loop, daemon=True)
        self.scheduler_thread.start()

    def scheduler_loop(self):
        self.reschedule_all_tasks()

        while self.scheduler_running:
            try:
                schedule.run_pending()
            except Exception:
                pass

            time_module.sleep(0.1)

    def reschedule_all_tasks(self):
        try:
            schedule.clear()

            for task in self.tasks:
                if task.status == TaskStatus.ENABLED:
                    self.schedule_task(task)

        except Exception:
            pass

    def on_tasks_changed(self):
        if self.scheduler_running:
            self.reschedule_all_tasks()

    def schedule_task(self, task: Task):
        try:
            today = QDate.currentDate()
            if today < task.start_date or today > task.end_date:
                return

            job = None

            if task.schedule_type == "interval":
                job = schedule.every(task.interval_seconds).seconds
            elif task.schedule_type == "daily":
                time_str = task.daily_time.toString("HH:mm")
                job = schedule.every().day.at(time_str)
            elif task.schedule_type == "weekly":
                time_str = task.daily_time.toString("HH:mm")
                days = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]
                job = getattr(schedule.every(), days[task.weekly_day]).at(time_str)
            elif task.schedule_type == "monthly":  # 新增每月执行
                time_str = task.daily_time.toString("HH:mm")
                job = schedule.every().day.at(time_str)

            if job:
                def job_wrapper(task):
                    try:
                        # 检查是否是每月执行的特定日期
                        if task.schedule_type == "monthly":
                            now = datetime.now()
                            if now.day != task.monthly_day:
                                return
                        self.signal_handler.execute_task_signal.emit(task)
                    except Exception:
                        pass

                job.do(job_wrapper, task)
        except Exception:
            pass

    def execute_task(self, task: Task):
        try:
            task.last_execution = datetime.now()
            task.execution_count += 1

            if task.task_type == TaskType.CMD:
                self.execute_cmd_task(task)
            else:
                self.execute_notification_task(task)

            self.signal_handler.refresh_tasks_signal.emit()
            self.save_tasks()

            # 记录必要的执行日志到文件
            if task.enable_logging:
                try:
                    with open("execution.log", "a", encoding="utf-8") as f:
                        f.write(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - 任务执行成功: {task.name}\n")
                except:
                    pass

        except Exception as e:
            self.refresh_tasks()
            self.save_tasks()

    def execute_cmd_task(self, task: Task):
        try:
            result = subprocess.run(task.cmd_command, shell=True, capture_output=True, text=True)
            if result.returncode == 0:
                self.show_notification("CMD任务执行成功", f"任务 '{task.name}' 执行成功", 3000)
            else:
                self.show_notification("CMD任务执行失败", f"任务 '{task.name}' 执行失败: {result.stderr}", 3000)
        except Exception as e:
            self.show_notification("CMD任务执行错误", f"任务 '{task.name}' 执行错误: {str(e)}", 3000)

    def execute_notification_task(self, task: Task):
        try:
            if task.popup_type == PopupType.WINDOW_POPUP.value or task.popup_type == "window_popup":
                # 创建Windows系统原生提示框，需要用户点击确认
                QMessageBox.information(self, task.notification_title, task.notification_content)
            else:
                # 系统托盘弹窗，使用弹窗显示时间
                self.show_notification(task.notification_title, task.notification_content, task.notification_timeout)
        except Exception:
            self.show_notification("任务提醒", f"任务 '{task.name}' 已执行", task.notification_timeout)

    def show_notification(self, title: str, message: str, timeout: int = 3000):
        try:
            if len(message) > 200:
                message = message[:200] + "..."

            self.tray_icon.showMessage(title, message, QSystemTrayIcon.MessageIcon.Information, timeout)
            self.status_label.setText(f"{title}: {message[:50]}...")

        except Exception:
            try:
                QMessageBox.information(self, title, message)
            except:
                pass

    def pause_all_tasks(self):
        for task in self.tasks:
            task.status = TaskStatus.DISABLED
        self.save_tasks()
        self.refresh_tasks()
        self.status_label.setText("所有任务已暂停")

    def resume_all_tasks(self):
        for task in self.tasks:
            task.status = TaskStatus.ENABLED
        self.save_tasks()
        self.refresh_tasks()
        self.status_label.setText("所有任务已恢复")

    def quit_application(self):
        self.scheduler_running = False
        self.refresh_timer.stop()
        self.tray_icon.hide()
        QApplication.quit()

    def closeEvent(self, event):
        reply = QMessageBox.question(
            self,
            '确认退出',
            '确定要退出定时任务管理器吗？',
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            self.quit_application()
            event.accept()
        else:
            event.ignore()

    def changeEvent(self, event):
        if event.type() == event.Type.WindowStateChange:
            if self.windowState() & Qt.WindowState.WindowMinimized:
                self.hide()
                self.is_minimized_to_tray = True
        super().changeEvent(event)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)
    app.setWindowIcon(app.style().standardIcon(QStyle.StandardPixmap.SP_TitleBarMenuButton))

    manager = TaskManager()
    manager.show()

    sys.exit(app.exec())
