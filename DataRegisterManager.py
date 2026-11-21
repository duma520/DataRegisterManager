#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import os
import sqlite3
import json
import shutil
import threading
import time
import hashlib
import re
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Any, Union
from dataclasses import dataclass, asdict, field
from enum import Enum
import queue
import traceback
import zipfile
from functools import partial
import base64

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QGridLayout, QTableWidget, QTableWidgetItem, QPushButton, QLabel,
    QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox, QCheckBox, 
    QGroupBox, QTabWidget, QSplitter, QMessageBox, QFileDialog,
    QDialog, QDialogButtonBox, QFormLayout, QTextEdit, QHeaderView,
    QProgressBar, QStatusBar, QMenuBar, QMenu, QToolBar,
    QTreeWidget, QTreeWidgetItem, QDateTimeEdit, QDateEdit, QCalendarWidget,
    QScrollArea, QFrame, QStackedWidget, QInputDialog, QListWidget,
    QListWidgetItem, QAbstractItemView, QStyleFactory, QApplication as QApp,
    QTableWidgetSelectionRange
)
from PySide6.QtCore import (
    Qt, Signal, QThread, QObject, QTimer, QDateTime, QDate, 
    QSize, QCoreApplication, QEvent, QMutex, QWaitCondition
)
from PySide6.QtGui import (
    QIcon, QFont, QColor, QPalette, QAction, QActionGroup, QKeySequence,
    QStandardItemModel, QStandardItem, QPixmap, QDesktopServices,
    QTextDocument, QTextCursor, QTextCharFormat, QTextBlockFormat, QBrush,
    QUndoStack, QUndoCommand 
)


def get_app_dir() -> Path:
    """获取应用程序所在目录"""
    if getattr(sys, 'frozen', False):
        app_dir = Path(sys.executable).parent
    else:
        app_dir = Path(__file__).parent
    return app_dir


def create_default_icon() -> QIcon:
    """创建默认图标（使用SVG风格的简单图标）"""
    # 创建一个简单的图标（蓝色圆形中间有白色"登"字）
    pixmap = QPixmap(64, 64)
    pixmap.fill(Qt.transparent)
    
    from PySide6.QtGui import QPainter, QBrush, QPen, QColor, QFont
    
    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.Antialiasing)
    
    # 绘制背景圆形
    painter.setBrush(QBrush(QColor(52, 152, 219)))  # 蓝色
    painter.setPen(QPen(Qt.NoPen))
    painter.drawEllipse(4, 4, 56, 56)
    
    # 绘制内圈
    painter.setBrush(QBrush(QColor(41, 128, 185)))  # 深蓝色
    painter.drawEllipse(8, 8, 48, 48)
    
    # 绘制文字
    painter.setPen(QPen(QColor(255, 255, 255)))
    font = QFont("Microsoft YaHei", 24, QFont.Bold)
    painter.setFont(font)
    painter.drawText(pixmap.rect(), Qt.AlignCenter, "登")
    
    painter.end()
    
    return QIcon(pixmap)


def get_icon_path() -> Path:
    """获取图标文件路径"""
    # 尝试多个可能的图标位置
    icon_names = ["app_icon.ico", "app_icon.png", "icon.ico", "icon.png"]
    
    for icon_name in icon_names:
        # 在应用程序目录查找
        icon_path = APP_DIR / icon_name
        if icon_path.exists():
            return icon_path
            
        icon_path = Path(__file__).parent / icon_name
        if icon_path.exists():
            return icon_path
            
    return None


# 应用程序目录
APP_DIR = get_app_dir()


# ==================== 项目信息元数据 ====================
class ProjectInfo:
    """项目信息元数据（集中管理所有项目相关信息）"""
    VERSION = "3.0.0"
    BUILD_DATE = "2026-08-20"
    AUTHOR = "杜玛"
    LICENSE = "GNU Affero General Public License v3.0"
    COPYRIGHT = "© 永久 杜玛"
    URL = "https://github.com/duma520/DataRegisterManager"
    MAINTAINER_EMAIL = "不提供"
    NAME = "资料登记管理系统"
    DESCRIPTION = "资料登记管理系统（支持多用户、灵活的表格配置、数据搜索和自动备份）"
    
    @classmethod
    def get_full_name(cls) -> str:
        """获取完整的程序名称（带版本）"""
        return f"{cls.NAME} v{cls.VERSION}"
    
    @classmethod
    def get_full_title(cls, username: str = None) -> str:
        """获取完整的窗口标题"""
        base = f"{cls.get_full_name()} 构建:{cls.BUILD_DATE}"
        if username:
            return f"{base} - 当前用户: {username}"
        return base
    
    @classmethod
    def get_about_text(cls) -> str:
        """获取关于信息文本"""
        return f"""
        <h2>{cls.NAME}</h2>
        <p><b>版本:</b> {cls.VERSION}</p>
        <p><b>构建日期:</b> {cls.BUILD_DATE}</p>
        <p><b>作者:</b> {cls.AUTHOR}</p>
        <p><b>版权:</b> {cls.COPYRIGHT}</p>
        <p><b>许可证:</b> {cls.LICENSE}</p>
        <p><b>项目主页:</b> <a href='{cls.URL}'>{cls.URL}</a></p>
        <p><b>描述:</b> {cls.DESCRIPTION}</p>
        """


# 配置常量（使用 ProjectInfo 统一管理）
APP_NAME = ProjectInfo.NAME
APP_VERSION = ProjectInfo.VERSION
BACKUP_DIR = APP_DIR / "backups"
CONFIG_FILE = APP_DIR / "app_config.json"
DB_FILE = APP_DIR / f"{APP_NAME.replace(' ', '_')}{'.db'}"
SETTINGS_FILE = APP_DIR / "settings.json"

WAL_MODE = "WAL"  # WAL模式

# 确保备份目录存在
try:
    BACKUP_DIR.mkdir(exist_ok=True)
except Exception as e:
    print(f"  - 备份目录创建失败: {e}")


# ==================== 应用设置管理（文件存储，避免注册表泄露） ====================
class AppSettings:
    """应用程序设置管理器
    
    将所有设置保存在程序目录下的 settings.json 文件中，
    避免使用 QSettings 将数据写入 Windows 注册表，防止数据泄露。
    """
    
    def __init__(self, file_path: Path = None):
        self._file_path = file_path or SETTINGS_FILE
        self._data: Dict[str, Any] = {}
        self._load()
    
    def _load(self):
        """从文件加载设置"""
        if self._file_path.exists():
            try:
                with open(self._file_path, 'r', encoding='utf-8') as f:
                    self._data = json.load(f)
            except Exception as e:
                print(f"[设置] 加载设置文件失败: {e}，使用默认设置")
                self._data = {}
        else:
            self._data = {}
    
    def _save(self):
        """保存设置到文件"""
        try:
            with open(self._file_path, 'w', encoding='utf-8') as f:
                json.dump(self._data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"[设置] 保存设置文件失败: {e}")
    
    def get(self, key: str, default: Any = None) -> Any:
        """获取设置值"""
        return self._data.get(key, default)
    
    def value(self, key: str, default: Any = None, type=None) -> Any:
        """获取设置值（兼容 QSettings 接口）"""
        val = self._data.get(key, default)
        if type is not None and val is not None:
            try:
                return type(val)
            except (ValueError, TypeError):
                return default
        return val
    
    def set(self, key: str, value: Any):
        """设置设置值"""
        self._data[key] = value
        self._save()
    
    def setValue(self, key: str, value: Any):
        """设置设置值（兼容 QSettings 接口）"""
        self.set(key, value)
    
    def remove(self, key: str):
        """移除设置项"""
        self._data.pop(key, None)
        self._save()
    
    def clear(self):
        """清空所有设置"""
        self._data.clear()
        self._save()


# ==================== 多线程后台工作器 ====================
class BackgroundWorker(QThread):
    """后台工作线程基类
    
    用于将耗时操作放到后台线程执行，避免阻塞 UI 主线程。
    通过信号与主线程通信，确保 UI 更新的线程安全性。
    """
    
    # 信号定义
    started = Signal()                          # 任务开始
    finished = Signal(object)                   # 任务完成，携带结果
    progress = Signal(int, int)                 # 进度更新 (当前值, 最大值)
    status = Signal(str)                        # 状态消息
    error = Signal(str)                         # 错误消息
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self._target = None
        self._args = ()
        self._kwargs = {}
        self._is_cancelled = False
        self._result = None
        self._mutex = QMutex()
    
    def setup(self, target, *args, **kwargs):
        """设置要执行的目标函数和参数"""
        self._target = target
        self._args = args
        self._kwargs = kwargs
    
    def run(self):
        """在工作线程中执行任务"""
        self._is_cancelled = False
        self.started.emit()
        try:
            self._result = self._target(*self._args, **self._kwargs)
            if not self._is_cancelled:
                self.finished.emit(self._result)
        except Exception as e:
            if not self._is_cancelled:
                error_msg = f"{type(e).__name__}: {str(e)}\n{traceback.format_exc()}"
                self.error.emit(error_msg)
    
    def cancel(self):
        """取消任务"""
        self._mutex.lock()
        self._is_cancelled = True
        self._mutex.unlock()
        if self.isRunning():
            self.quit()
            self.wait(3000)
    
    def report_progress(self, current: int, maximum: int):
        """报告进度（从工作函数中调用）"""
        self.progress.emit(current, maximum)
    
    def report_status(self, msg: str):
        """报告状态消息（从工作函数中调用）"""
        self.status.emit(msg)


# 单元格配置方案
class CellScheme(Enum):
    STANDARD = "标准方案"
    CUSTOM = "自定义方案"
    ENHANCED = "增强方案"
    SIMPLE = "简化方案"
    

class TitleMode(Enum):
    """标题模式"""
    UNIFIED = "统一标题"      # 所有列使用相同标题
    INDEPENDENT = "独立标题"   # 每列独立标题


@dataclass
class ColumnConfig:
    """列配置类"""
    title: str = ""           # 列标题
    width: int = 100          # 列宽（像素）
    visible: bool = True      # 是否可见
    
    
@dataclass
class CellConfig:
    """单元格配置类"""
    row_count: int = 3
    col_count: int = 4
    require_login_time: bool = False
    calculate_days_diff: bool = False
    require_title: bool = True
    title_mode: str = TitleMode.UNIFIED.value  # 标题模式
    title_text: str = ""                        # 统一标题文本
    column_titles: Dict[int, str] = field(default_factory=dict)  # 列索引 -> 标题
    title_unique: bool = True                   # 保留兼容性，实际用title_mode
    scheme_type: str = CellScheme.STANDARD.value
    custom_config: Dict = None
    
    def __post_init__(self):
        if self.custom_config is None:
            self.custom_config = {}
        if self.column_titles is None:
            self.column_titles = {}
            
    def get_column_title(self, col_index: int) -> str:
        """获取指定列的标题"""
        if self.title_mode == TitleMode.INDEPENDENT.value:
            return self.column_titles.get(col_index, f"列{col_index + 1}")
        else:
            return self.title_text or f"列{col_index + 1}"
            
    def set_column_title(self, col_index: int, title: str):
        """设置指定列的标题"""
        self.title_mode = TitleMode.INDEPENDENT.value
        self.column_titles[col_index] = title
        
    def get_all_column_titles(self) -> List[str]:
        """获取所有列标题"""
        titles = []
        for i in range(self.col_count):
            titles.append(self.get_column_title(i))
        return titles


@dataclass
class UserConfig:
    """用户配置类"""
    user_id: int
    username: str
    created_at: str
    cell_count: int = 6          # 单元格数量
    cell_configs: Dict[int, CellConfig] = None  # key: 单元格索引
    auto_resize_columns: bool = True  # 是否自动调整列宽
    
    def __post_init__(self):
        if self.cell_configs is None:
            self.cell_configs = {}


class BackupManager:
    """备份管理器类"""
    
    def __init__(self, db_path: str):
        self.db_path = db_path
        self.backup_dir = BACKUP_DIR
        self.max_backups = 30
        
    def create_backup(self, backup_type: str = "manual") -> Tuple[bool, str, Dict]:
        """创建数据库备份"""
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_name = f"backup_{timestamp}_{backup_type}.db"
            backup_path = self.backup_dir / backup_name
            
            # 获取备份前信息
            source_size = os.path.getsize(self.db_path)
            
            # 创建备份
            shutil.copy2(self.db_path, backup_path)
            
            # 获取备份信息
            backup_size = os.path.getsize(backup_path)
            backup_info = {
                "name": backup_name,
                "path": str(backup_path),
                "type": backup_type,
                "time": timestamp,
                "size": backup_size,
                "datetime": datetime.now()
            }

            # 清理旧备份
            self._cleanup_old_backups()
            
            return True, backup_name, backup_info
            
        except Exception as e:
            print(f"[备份] 备份失败: {e}")
            return False, str(e), {}
            
    def _cleanup_old_backups(self):
        """清理旧备份"""
        backups = sorted(self.backup_dir.glob("*.db"), key=os.path.getmtime)
        while len(backups) > self.max_backups:
            backups[0].unlink()
            backups.pop(0)
            
    def get_backup_list(self) -> List[Dict]:
        """获取备份列表"""
        backups = []
        for backup_file in sorted(self.backup_dir.glob("*.db"), key=os.path.getmtime, reverse=True):
            # 解析备份信息
            name = backup_file.name
            parts = name.replace(".db", "").split("_")
            backup_type = parts[2] if len(parts) > 2 else "unknown"
            backup_time = parts[1] if len(parts) > 1 else "unknown"
            
            size_bytes = backup_file.stat().st_size
            size_mb = size_bytes / (1024 * 1024)
            size_str = f"{size_mb:.2f} MB" if size_mb >= 1 else f"{size_bytes / 1024:.2f} KB"
            
            # 获取备份的数据库信息
            db_info = self._get_backup_info(str(backup_file))
            
            backups.append({
                "name": name,
                "path": str(backup_file),
                "type": backup_type,
                "time": backup_time,
                "datetime": datetime.strptime(backup_time, "%Y%m%d_%H%M%S") if backup_time != "unknown" else None,
                "size": size_bytes,
                "size_str": size_str,
                "info": db_info
            })
        return backups
        
    def _get_backup_info(self, backup_path: str) -> Dict:
        """获取备份数据库信息"""
        try:
            conn = sqlite3.connect(backup_path)
            cursor = conn.cursor()
            
            # 获取版本信息
            cursor.execute("SELECT sqlite_version()")
            version = cursor.fetchone()[0]
            
            # 获取表信息
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
            tables = [row[0] for row in cursor.fetchall()]
            
            # 获取排班日期范围（如果有相关表）
            date_range = None
            if "data_records" in tables:
                cursor.execute("SELECT MIN(record_date), MAX(record_date) FROM data_records")
                result = cursor.fetchone()
                if result and result[0] and result[1]:
                    date_range = f"{result[0]} 至 {result[1]}"
                    
            conn.close()
            
            return {
                "version": version,
                "tables": tables,
                "date_range": date_range
            }
        except:
            return {"version": "unknown", "tables": [], "date_range": None}
            
    def restore_backup(self, backup_path: str) -> Tuple[bool, str]:
        """恢复数据库备份"""
        try:
            # 先创建当前数据库的备份（回滚点）
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            rollback_name = f"rollback_{timestamp}_before_restore.db"
            rollback_path = self.backup_dir / rollback_name
            shutil.copy2(self.db_path, rollback_path)
            
            # 恢复备份
            shutil.copy2(backup_path, self.db_path)
            
            return True, rollback_name
            
        except Exception as e:
            print(f"[恢复] 恢复失败: {e}")
            return False, str(e)
            
    def set_max_backups(self, max_count: int):
        """设置最大备份数量"""
        self.max_backups = max_count
        self._cleanup_old_backups()


class DatabaseManager:
    """数据库管理器类"""
    
    def __init__(self, db_path: str):
        self.db_path = db_path
        self.connection = None
        self.cursor = None
        self._lock = threading.Lock()  # 线程安全锁
        self._init_database()
        
    def _init_database(self):
        """初始化数据库"""
        try:
            # 启用WAL模式
            self.connection = sqlite3.connect(self.db_path, check_same_thread=False)
            
            self.connection.execute("PRAGMA journal_mode=WAL")
            wal_status = self.connection.execute("PRAGMA journal_mode").fetchone()[0]
            
            self.connection.execute("PRAGMA synchronous=NORMAL")
            self.cursor = self.connection.cursor()
            
            # 创建基础表
            self.cursor.execute("""
                CREATE TABLE IF NOT EXISTS users (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    username TEXT UNIQUE NOT NULL,
                    password TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    config TEXT
                )
            """)
            
            self.cursor.execute("""
                CREATE TABLE IF NOT EXISTS user_configs (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    user_id INTEGER,
                    config_key TEXT,
                    config_value TEXT,
                    FOREIGN KEY (user_id) REFERENCES users(id)
                )
            """)
            
            self.cursor.execute("""
                CREATE TABLE IF NOT EXISTS data_records (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    user_id INTEGER,
                    cell_index INTEGER,
                    row_index INTEGER,
                    col_index INTEGER,
                    record_date DATE,
                    login_time TIMESTAMP,
                    content TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (user_id) REFERENCES users(id)
                )
            """)
            
            self.cursor.execute("""
                CREATE TABLE IF NOT EXISTS cell_templates (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    user_id INTEGER,
                    cell_index INTEGER,
                    config TEXT,
                    FOREIGN KEY (user_id) REFERENCES users(id)
                )
            """)
            
            self.connection.commit()
            
            # 创建默认管理员用户（如果不存在）
            self._create_default_admin()
            
        except Exception as e:
            print(f"[数据库] 初始化错误: {e}")
            import traceback
            traceback.print_exc()
            
    def _create_default_admin(self):
        """创建默认管理员用户"""
        admin_exists = self.execute_query("SELECT id FROM users WHERE username = ?", ("admin",))
        if not admin_exists:
            # 默认密码: admin123
            default_password = hashlib.sha256("admin123".encode()).hexdigest()
            self.execute_update(
                "INSERT INTO users (username, password, config) VALUES (?, ?, ?)",
                ("admin", default_password, json.dumps({"cell_count": 6, "default_cell": {}, "auto_resize_columns": True}))
            )
            
    def execute_query(self, query: str, params: tuple = ()) -> List[tuple]:
        """执行查询（线程安全）"""
        with self._lock:
            try:
                self.cursor.execute(query, params)
                return self.cursor.fetchall()
            except Exception as e:
                print(f"[数据库] 查询错误: {e}")
                return []
            
    def execute_update(self, query: str, params: tuple = ()) -> bool:
        """执行更新（线程安全）"""
        with self._lock:
            try:
                self.cursor.execute(query, params)
                self.connection.commit()
                return True
            except Exception as e:
                print(f"[数据库] 更新错误: {e}")
                self.connection.rollback()
                return False
            
    def verify_user(self, username: str, password: str) -> Optional[Dict]:
        """验证用户登录（支持空密码）"""
        # 如果密码为空字符串，直接使用空字符串验证
        if password == "":
            result = self.execute_query(
                "SELECT id, username, created_at, config FROM users WHERE username = ? AND (password = '' OR password IS NULL)",
                (username,)
            )
        else:
            # 非空密码需要哈希验证
            hashed_password = hashlib.sha256(password.encode()).hexdigest()
            result = self.execute_query(
                "SELECT id, username, created_at, config FROM users WHERE username = ? AND password = ?",
                (username, hashed_password)
            )
            
        if result:
            return {
                "id": result[0][0],
                "username": result[0][1],
                "created_at": result[0][2],
                "config": json.loads(result[0][3]) if result[0][3] else {}
            }
        return None
        
    def get_user(self, username: str) -> Optional[Dict]:
        """获取用户信息"""
        result = self.execute_query(
            "SELECT id, username, created_at, config FROM users WHERE username = ?",
            (username,)
        )
        if result:
            return {
                "id": result[0][0],
                "username": result[0][1],
                "created_at": result[0][2],
                "config": json.loads(result[0][3]) if result[0][3] else {}
            }
        return None
        
    def get_user_by_id(self, user_id: int) -> Optional[Dict]:
        """根据ID获取用户信息"""
        result = self.execute_query(
            "SELECT id, username, created_at, config FROM users WHERE id = ?",
            (user_id,)
        )
        if result:
            return {
                "id": result[0][0],
                "username": result[0][1],
                "created_at": result[0][2],
                "config": json.loads(result[0][3]) if result[0][3] else {}
            }
        return None
        
    def create_user(self, username: str, password: str, config: Dict) -> bool:
        """创建用户（支持空密码）"""
        if password == "":
            # 空密码直接存储空字符串
            return self.execute_update(
                "INSERT INTO users (username, password, config) VALUES (?, ?, ?)",
                (username, "", json.dumps(config))
            )
        else:
            # 非空密码哈希后存储
            hashed_password = hashlib.sha256(password.encode()).hexdigest()
            return self.execute_update(
                "INSERT INTO users (username, password, config) VALUES (?, ?, ?)",
                (username, hashed_password, json.dumps(config))
            )
        
    def delete_user(self, user_id: int) -> bool:
        """删除用户"""
        # 删除用户相关数据
        self.execute_update("DELETE FROM user_configs WHERE user_id = ?", (user_id,))
        self.execute_update("DELETE FROM data_records WHERE user_id = ?", (user_id,))
        self.execute_update("DELETE FROM cell_templates WHERE user_id = ?", (user_id,))
        return self.execute_update("DELETE FROM users WHERE id = ?", (user_id,))
        
    def update_username(self, user_id: int, new_username: str) -> bool:
        """更新用户名"""
        return self.execute_update(
            "UPDATE users SET username = ? WHERE id = ?",
            (new_username, user_id)
        )
        
    def update_user_password(self, user_id: int, new_password: str) -> bool:
        """更新用户密码（支持空密码）"""
        if new_password == "":
            # 空密码直接存储空字符串
            return self.execute_update(
                "UPDATE users SET password = ? WHERE id = ?",
                ("", user_id)
            )
        else:
            # 非空密码哈希后存储
            hashed_password = hashlib.sha256(new_password.encode()).hexdigest()
            return self.execute_update(
                "UPDATE users SET password = ? WHERE id = ?",
                (hashed_password, user_id)
            )
        
    def update_user_config(self, user_id: int, config: Dict) -> bool:
        """更新用户配置"""
        return self.execute_update(
            "UPDATE users SET config = ? WHERE id = ?",
            (json.dumps(config), user_id)
        )
        
    def get_all_users(self) -> List[Dict]:
        """获取所有用户（不包含密码）"""
        results = self.execute_query("SELECT id, username, created_at FROM users ORDER BY id")
        users = [{"id": r[0], "username": r[1], "created_at": r[2]} for r in results]
        return users
        
    def save_cell_config(self, user_id: int, cell_index: int, config: CellConfig) -> bool:
        """保存单元格配置"""
        return self.execute_update(
            "INSERT OR REPLACE INTO cell_templates (user_id, cell_index, config) VALUES (?, ?, ?)",
            (user_id, cell_index, json.dumps(asdict(config)))
        )
        
    def get_cell_config(self, user_id: int, cell_index: int) -> Optional[CellConfig]:
        """获取单元格配置"""
        result = self.execute_query(
            "SELECT config FROM cell_templates WHERE user_id = ? AND cell_index = ?",
            (user_id, cell_index)
        )
        if result:
            config_dict = json.loads(result[0][0])
            return CellConfig(**config_dict)
        return None
        
    def save_data_record(self, user_id: int, cell_index: int, row_index: int, 
                        col_index: int, content: str, login_time: datetime = None) -> bool:
        """保存数据记录"""
        if login_time is None:
            login_time = datetime.now()
            
        record_date = datetime.now().date()
        
        # 检查是否已存在
        existing = self.execute_query(
            "SELECT id FROM data_records WHERE user_id = ? AND cell_index = ? AND row_index = ? AND col_index = ? AND record_date = ?",
            (user_id, cell_index, row_index, col_index, record_date.isoformat())
        )
        
        if existing:
            return self.execute_update(
                """UPDATE data_records 
                   SET content = ?, updated_at = CURRENT_TIMESTAMP 
                   WHERE user_id = ? AND cell_index = ? AND row_index = ? AND col_index = ? AND record_date = ?""",
                (content, user_id, cell_index, row_index, col_index, record_date.isoformat())
            )
        else:
            return self.execute_update(
                """INSERT INTO data_records (user_id, cell_index, row_index, col_index, record_date, login_time, content) 
                   VALUES (?, ?, ?, ?, ?, ?, ?)""",
                (user_id, cell_index, row_index, col_index, record_date.isoformat(), 
                 login_time.isoformat(), content)
            )
            
    def get_data_records(self, user_id: int, cell_index: int = None) -> List[Dict]:
        """获取数据记录"""
        if cell_index is not None:
            query = """SELECT id, cell_index, row_index, col_index, record_date, login_time, content 
                      FROM data_records WHERE user_id = ? AND cell_index = ? ORDER BY record_date, row_index, col_index"""
            params = (user_id, cell_index)
        else:
            query = """SELECT id, cell_index, row_index, col_index, record_date, login_time, content 
                      FROM data_records WHERE user_id = ? ORDER BY record_date, row_index, col_index"""
            params = (user_id,)
            
        results = self.execute_query(query, params)
        records = []
        for r in results:
            records.append({
                "id": r[0],
                "cell_index": r[1],
                "row_index": r[2],
                "col_index": r[3],
                "record_date": r[4],
                "login_time": r[5],
                "content": r[6]
            })
        return records
        
    def search_data(self, user_id: int, keyword: str) -> List[Dict]:
        """搜索数据"""
        keyword_lower = keyword.lower()
        records = self.get_data_records(user_id)
        
        filtered = []
        for record in records:
            if record["content"] and keyword_lower in record["content"].lower():
                filtered.append(record)
                
        return filtered
        
    def close(self):
        """关闭数据库连接"""
        if self.connection:
            self.connection.close()


class CellEditCommand(QUndoCommand):
    """单元格编辑命令"""
    
    def __init__(self, table_widget, cell_index, row, col, old_text, new_text, description=""):
        super().__init__(description)
        self.table_widget = table_widget
        self.cell_index = cell_index
        self.row = row
        self.col = col
        self.old_text = old_text
        self.new_text = new_text
        
    def undo(self):
        """撤销操作"""
        self.table_widget._set_cell_text(self.row, self.col, self.old_text)
        
    def redo(self):
        """重做操作"""
        self.table_widget._set_cell_text(self.row, self.col, self.new_text)


class BatchEditCommand(QUndoCommand):
    """批量编辑命令（用于粘贴等操作）"""
    
    def __init__(self, table_widget, changes, description=""):
        """
        changes: [(row, col, old_text, new_text), ...]
        """
        super().__init__(description)
        self.table_widget = table_widget
        self.changes = changes
        
    def undo(self):
        """撤销操作"""
        for row, col, old_text, _ in self.changes:
            self.table_widget._set_cell_text(row, col, old_text)
            
    def redo(self):
        """重做操作"""
        for row, col, _, new_text in self.changes:
            self.table_widget._set_cell_text(row, col, new_text)



class DataTableWidget(QTableWidget):
    """数据表格组件类"""
    
    data_changed = Signal(int, int, int, str)  # cell_index, row, col, content
    resize_requested = Signal()  # 请求调整列宽
    paste_extended = Signal(int, int, int)  # cell_index, new_rows, new_cols - 粘贴扩展信号
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.cell_config = None
        self.user_id = None
        self.cell_index = None
        self.auto_resize_columns = True
        self.highlight_keyword = None
        self._clipboard_data = []  # 存储剪贴板数据（用于重复粘贴）
        self._clipboard_rows = 0
        self._clipboard_cols = 0
        self._undo_stack = QUndoStack(self)  # 撤销堆栈
        self._is_programmatic_change = False  # 标记是否为程序化修改
        self._old_text_before_edit = ""  # 记录编辑前的文本
        self.setup_ui()
        self.setup_clipboard_shortcuts()
        self.setup_undo_shortcuts()
        
    def setup_ui(self):
        """设置UI"""
        self.setAlternatingRowColors(True)
        # 支持多选（Ctrl+点击多选，Shift+点击范围选择）
        self.setSelectionBehavior(QTableWidget.SelectItems)
        self.setSelectionMode(QTableWidget.ExtendedSelection)
        self.itemChanged.connect(self.on_item_changed)
        
    def setup_clipboard_shortcuts(self):
        """设置剪贴板快捷键"""
        # 复制 Ctrl+C
        copy_action = QAction("复制", self)
        copy_action.setShortcut(QKeySequence.Copy)
        copy_action.triggered.connect(self.copy_selected)
        self.addAction(copy_action)
        
        # 剪切 Ctrl+X
        cut_action = QAction("剪切", self)
        cut_action.setShortcut(QKeySequence.Cut)
        cut_action.triggered.connect(self.cut_selected)
        self.addAction(cut_action)
        
        # 粘贴 Ctrl+V
        paste_action = QAction("粘贴", self)
        paste_action.setShortcut(QKeySequence.Paste)
        paste_action.triggered.connect(self.paste_from_clipboard)
        self.addAction(paste_action)
        
        # 也添加右键菜单支持
        self.setContextMenuPolicy(Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self.show_context_menu)
        
    def setup_undo_shortcuts(self):
        """设置撤销/重做快捷键"""
        # 撤销 Ctrl+Z
        undo_action = self._undo_stack.createUndoAction(self, "撤销")
        undo_action.setShortcut(QKeySequence.Undo)
        self.addAction(undo_action)
        
        # 重做 Ctrl+Y 或 Ctrl+Shift+Z
        redo_action = self._undo_stack.createRedoAction(self, "重做")
        redo_action.setShortcuts([QKeySequence.Redo, QKeySequence("Ctrl+Y")])
        self.addAction(redo_action)
        
    def _set_cell_text(self, row, col, text):
        """程序化设置单元格文本（不触发撤销）"""
        if 0 <= row < self.rowCount() and 0 <= col < self.columnCount():
            item = self.item(row, col)
            if not item:
                item = QTableWidgetItem()
                self.setItem(row, col, item)
            
            self._is_programmatic_change = True
            item.setText(text)
            self._is_programmatic_change = False
            self.apply_highlight_to_item(item)
        
    def show_context_menu(self, pos):
        """显示右键菜单"""
        menu = QMenu(self)
        
        # 撤销/重做
        undo_action = self._undo_stack.createUndoAction(self, "撤销\tCtrl+Z")
        menu.addAction(undo_action)
        
        redo_action = self._undo_stack.createRedoAction(self, "重做\tCtrl+Y")
        menu.addAction(redo_action)
        
        menu.addSeparator()
        
        copy_action = menu.addAction("复制\tCtrl+C")
        copy_action.triggered.connect(self.copy_selected)
        copy_action.setEnabled(len(self.selectedRanges()) > 0)
        
        cut_action = menu.addAction("剪切\tCtrl+X")
        cut_action.triggered.connect(self.cut_selected)
        cut_action.setEnabled(len(self.selectedRanges()) > 0)
        
        menu.addSeparator()
        
        paste_action = menu.addAction("粘贴（覆盖）\tCtrl+V")
        paste_action.triggered.connect(self.paste_from_clipboard)
        paste_action.setEnabled(self._can_paste())
        
        paste_insert_action = menu.addAction("粘贴（插入新行列）")
        paste_insert_action.triggered.connect(self.paste_with_insert)
        paste_insert_action.setEnabled(self._can_paste())
        
        menu.addSeparator()
        
        clear_action = menu.addAction("清除选中内容\tDelete")
        clear_action.triggered.connect(self.clear_selected)
        clear_action.setEnabled(len(self.selectedRanges()) > 0)
        
        menu.addSeparator()
        
        # 添加行列操作
        insert_row_action = menu.addAction("在当前位置插入行")
        insert_row_action.triggered.connect(self.insert_row_at_current)
        
        insert_col_action = menu.addAction("在当前位置插入列")
        insert_col_action.triggered.connect(self.insert_col_at_current)
        
        delete_row_action = menu.addAction("删除当前行")
        delete_row_action.triggered.connect(self.delete_current_row)
        
        delete_col_action = menu.addAction("删除当前列")
        delete_col_action.triggered.connect(self.delete_current_col)
        
        menu.exec(self.mapToGlobal(pos))
        
    def keyPressEvent(self, event):
        """处理键盘事件，支持Delete键清除和Ctrl+Z/Y"""
        if event.key() == Qt.Key_Delete:
            self.clear_selected()
        elif event.key() == Qt.Key_Return or event.key() == Qt.Key_Enter:
            # 回车键开始编辑当前单元格
            current = self.currentItem()
            if current:
                self.editItem(current)
        else:
            super().keyPressEvent(event)
    
    def get_all_selected_ranges(self) -> List[Tuple[int, int, int, int]]:
        """获取所有选中的区域（支持多个不连续区域）
        
        Returns:
            List of (top_row, bottom_row, left_col, right_col)
        """
        ranges = []
        for sel_range in self.selectedRanges():
            top_row = sel_range.topRow()
            bottom_row = sel_range.bottomRow()
            left_col = sel_range.leftColumn()
            right_col = sel_range.rightColumn()
            ranges.append((top_row, bottom_row, left_col, right_col))
        return ranges
    
    def get_selected_data(self) -> tuple:
        """获取选中区域的数据（支持多个不连续区域）
        
        Returns:
            tuple: (all_data, min_row, min_col)
                all_data: 列表，包含 (data_matrix, start_row, start_col) 元组
                min_row: 所有选中区域的最小行号
                min_col: 所有选中区域的最小列号
        """
        ranges = self.get_all_selected_ranges()
        if not ranges:
            return [], -1, -1
        
        min_row = min(r[0] for r in ranges)
        min_col = min(r[2] for r in ranges)
        
        all_data = []
        for top_row, bottom_row, left_col, right_col in ranges:
            data = []
            for row in range(top_row, bottom_row + 1):
                row_data = []
                for col in range(left_col, right_col + 1):
                    item = self.item(row, col)
                    row_data.append(item.text() if item else "")
                data.append(row_data)
            all_data.append((data, top_row, left_col))
        
        return all_data, min_row, min_col
    
    def copy_selected(self):
        """复制选中内容到剪贴板"""
        all_data, _, _ = self.get_selected_data()
        if not all_data:
            return False
        
        # 对于多个区域，使用特殊分隔符
        text_parts = []
        for data, _, _ in all_data:
            text_parts.append(self._data_to_tsv(data))
        
        # 多个区域用双换行 + 特殊标记分隔
        if len(text_parts) > 1:
            text = "\n\n---SELECTION_SEPARATOR---\n\n".join(text_parts)
        else:
            text = text_parts[0] if text_parts else ""
        
        # 复制到剪贴板
        clipboard = QApplication.clipboard()
        clipboard.setText(text)
        
        # 同时保存内部剪贴板（用于后续粘贴操作）
        self._store_clipboard_data(all_data)
        
        return True
    
    def _store_clipboard_data(self, all_data):
        """存储剪贴板数据到内部缓存"""
        if not all_data:
            return
        
        # all_data 是 [(data_matrix, start_row, start_col), ...] 的列表
        # 先找到最小行列作为基准
        min_row = min(rd[1] for rd in all_data)  # rd[1] 是 start_row
        min_col = min(rd[2] for rd in all_data)  # rd[2] 是 start_col
        
        # 计算最大行列
        max_row = max(rd[1] + len(rd[0]) - 1 for rd in all_data)
        max_col = max(rd[2] + (len(rd[0][0]) - 1 if rd[0] and rd[0][0] else 0) for rd in all_data)
        
        rows = max_row - min_row + 1
        cols = max_col - min_col + 1
        
        # 创建空矩阵
        matrix = [[""] * cols for _ in range(rows)]
        
        # 填充数据
        for data, start_row, start_col in all_data:
            for i, row_data in enumerate(data):
                for j, value in enumerate(row_data):
                    matrix_row = start_row - min_row + i
                    matrix_col = start_col - min_col + j
                    if 0 <= matrix_row < rows and 0 <= matrix_col < cols:
                        matrix[matrix_row][matrix_col] = value
        
        self._clipboard_data = matrix
        self._clipboard_rows = rows
        self._clipboard_cols = cols
    
    def cut_selected(self):
        """剪切选中内容（支持撤销）"""
        all_data, _, _ = self.get_selected_data()
        if not all_data:
            return False
        
        # 先复制
        self.copy_selected()
        
        # 记录所有更改
        changes = []
        for data, start_row, start_col in all_data:
            for i, row_data in enumerate(data):
                for j, value in enumerate(row_data):
                    if value:  # 只清空非空的单元格
                        row = start_row + i
                        col = start_col + j
                        item = self.item(row, col)
                        old_text = item.text() if item else ""
                        changes.append((row, col, old_text, ""))
        
        if changes:
            # 创建批量编辑命令
            command = BatchEditCommand(self, changes, f"剪切 {len(changes)} 个单元格")
            self._undo_stack.push(command)
            
            # 执行清空操作
            self.blockSignals(True)
            for row, col, old_text, _ in changes:
                item = self.item(row, col)
                if item:
                    item.setText("")
                else:
                    new_item = QTableWidgetItem("")
                    self.setItem(row, col, new_item)
            self.blockSignals(False)
            
            # 触发数据变更信号
            for row, col, _, _ in changes:
                self.data_changed.emit(
                    self.cell_index if self.cell_index is not None else -1,
                    row, col, ""
                )
        
        return True
    
    def paste_from_clipboard(self, insert_mode: bool = False):
        """从剪贴板粘贴数据
        
        Args:
            insert_mode: 如果为True，则在需要时插入新行列
        """
        clipboard = QApplication.clipboard()
        text = clipboard.text()
        
        if not text:
            # 尝试使用内部剪贴板数据
            if self._clipboard_data:
                return self._paste_internal_data(insert_mode)
            return False
        
        # 检查是否有多区域分隔符
        if "---SELECTION_SEPARATOR---" in text:
            parts = text.split("\n\n---SELECTION_SEPARATOR---\n\n")
            # 只使用第一个区域进行粘贴（简化处理）
            data = self._tsv_to_data(parts[0])
        else:
            # 解析TSV数据
            data = self._tsv_to_data(text)
        
        if not data:
            return False
        
        return self._do_paste(data, insert_mode)
    
    def _paste_internal_data(self, insert_mode: bool = False) -> bool:
        """使用内部缓存的剪贴板数据进行粘贴"""
        if not self._clipboard_data:
            return False
        return self._do_paste(self._clipboard_data, insert_mode)
    
    def paste_with_insert(self):
        """粘贴数据，在需要时插入新行/列"""
        return self.paste_from_clipboard(insert_mode=True)
    
    def _do_paste(self, data: List[List[str]], insert_mode: bool = False) -> bool:
        """执行粘贴操作（支持撤销）
        
        Args:
            data: 要粘贴的二维数据
            insert_mode: 是否在超出范围时插入行/列
        """
        if not data:
            return False
        
        paste_rows = len(data)
        paste_cols = len(data[0]) if data else 0
        
        # 【关键修复】获取当前选中区域的左上角作为粘贴起始位置
        ranges = self.selectedRanges()
        if ranges:
            # 使用选中区域的左上角作为粘贴起始位置
            start_row = ranges[0].topRow()
            start_col = ranges[0].leftColumn()
        else:
            # 如果没有选中区域，尝试使用当前单元格
            current_item = self.currentItem()
            if current_item:
                start_row = current_item.row()
                start_col = current_item.column()
            else:
                # 如果都没有，从第一个单元格开始
                start_row = 0
                start_col = 0
        
        # 记录更改前的值（用于撤销）
        changes = []
        
        # 检查是否超出表格范围
        max_row = start_row + paste_rows
        max_col = start_col + paste_cols
        
        need_expand = max_row > self.rowCount() or max_col > self.columnCount()
        
        if need_expand:
            if insert_mode:
                # 插入模式：自动扩展表格
                self._expand_table(max_row, max_col)
                # 发送扩展信号
                if hasattr(self, 'paste_extended'):
                    self.paste_extended.emit(
                        self.cell_index if self.cell_index is not None else -1,
                        max(self.rowCount(), max_row),
                        max(self.columnCount(), max_col)
                    )
            else:
                # 覆盖模式：询问用户
                reply = QMessageBox.question(
                    self,
                    "粘贴范围超出",
                    f"粘贴内容将超出表格范围\n\n"
                    f"当前表格: {self.rowCount()}行 × {self.columnCount()}列\n"
                    f"粘贴起始: 第{start_row+1}行 第{start_col+1}列\n"
                    f"需要范围: {paste_rows}行 × {paste_cols}列\n\n"
                    f"是否扩展表格？（是=扩展表格，否=裁剪数据，取消=不粘贴）",
                    QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel
                )
                
                if reply == QMessageBox.Cancel:
                    return False
                elif reply == QMessageBox.Yes:
                    # 扩展表格
                    self._expand_table(max_row, max_col)
                    if hasattr(self, 'paste_extended'):
                        self.paste_extended.emit(
                            self.cell_index if self.cell_index is not None else -1,
                            max(self.rowCount(), max_row),
                            max(self.columnCount(), max_col)
                        )
                else:
                    # 裁剪数据以适应表格
                    data = self._clip_data_to_table(data, start_row, start_col)
                    if not data:
                        return False
                    paste_rows = len(data)
                    paste_cols = len(data[0]) if data else 0
        
        # 收集更改信息
        self.blockSignals(True)
        for i in range(paste_rows):
            for j in range(paste_cols):
                target_row = start_row + i
                target_col = start_col + j
                
                if target_row < self.rowCount() and target_col < self.columnCount():
                    value = data[i][j] if i < len(data) and j < len(data[i]) else ""
                    
                    # 获取旧值
                    item = self.item(target_row, target_col)
                    if not item:
                        old_text = ""
                        item = QTableWidgetItem()
                        self.setItem(target_row, target_col, item)
                    else:
                        old_text = item.text()
                    
                    # 如果值有变化，记录下来
                    if old_text != value:
                        changes.append((target_row, target_col, old_text, value))
                    
                    item.setText(value)
                    self.apply_highlight_to_item(item)
        self.blockSignals(False)
        
        # 创建撤销命令
        if changes:
            command = BatchEditCommand(self, changes, f"粘贴 {paste_rows}×{paste_cols} 数据")
            self._undo_stack.push(command)
        
        # 选择粘贴的区域
        self.clearSelection()
        self.setRangeSelected(
            QTableWidgetSelectionRange(start_row, start_col, 
                                       start_row + paste_rows - 1, 
                                       start_col + paste_cols - 1),
            True
        )
        
        # 批量发送数据变更信号
        for row, col, _, new_text in changes:
            if new_text:  # 只发送非空的变更
                self.data_changed.emit(
                    self.cell_index if self.cell_index is not None else -1,
                    row, col, new_text
                )
        
        # 如果启用自动调整列宽，请求调整
        if self.auto_resize_columns:
            for col in range(start_col, start_col + paste_cols):
                self.resizeColumnToContents(col)
        
        return True
    
    def clear_selected(self):
        """清除选中单元格的内容（支持撤销和多区域）"""
        all_data, min_row, min_col = self.get_selected_data()
        if not all_data:
            return
        
        # 计算总单元格数
        total_cells = sum(len(data) * (len(data[0]) if data else 0) 
                         for data, _, _ in all_data)
        
        reply = QMessageBox.question(
            self,
            "确认清除",
            f"确定要清除选中的 {total_cells} 个单元格的内容吗？\n"
            f"（选中了 {len(all_data)} 个区域）",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply != QMessageBox.Yes:
            return
        
        # 记录更改
        changes = []
        for data, start_row, start_col in all_data:
            for i, row_data in enumerate(data):
                for j, value in enumerate(row_data):
                    if value:  # 只清除非空的单元格
                        row = start_row + i
                        col = start_col + j
                        item = self.item(row, col)
                        old_text = item.text() if item else ""
                        changes.append((row, col, old_text, ""))
        
        if changes:
            # 创建撤销命令
            command = BatchEditCommand(self, changes, f"清除 {len(changes)} 个单元格")
            self._undo_stack.push(command)
            
            # 执行清除
            self.blockSignals(True)
            for row, col, old_text, _ in changes:
                item = self.item(row, col)
                if item:
                    item.setText("")
            self.blockSignals(False)
            
            # 触发数据变更信号
            for row, col, _, _ in changes:
                self.data_changed.emit(
                    self.cell_index if self.cell_index is not None else -1,
                    row, col, ""
                )

    def insert_row_at_current(self):
        """在当前选中位置插入一行"""
        current_item = self.currentItem()
        if not current_item:
            QMessageBox.warning(self, "提示", "请先选择一个单元格")
            return
        
        insert_row = current_item.row()
        self.insertRow(insert_row)
        
        # 为新行初始化空项目
        for col in range(self.columnCount()):
            item = QTableWidgetItem("")
            self.setItem(insert_row, col, item)
        
        # 更新行配置
        if self.cell_config:
            self.cell_config.row_count = self.rowCount()
    
    def insert_col_at_current(self):
        """在当前选中位置插入一列"""
        current_item = self.currentItem()
        if not current_item:
            QMessageBox.warning(self, "提示", "请先选择一个单元格")
            return
        
        insert_col = current_item.column()
        self.insertColumn(insert_col)
        
        # 为新列初始化空项目
        for row in range(self.rowCount()):
            item = QTableWidgetItem("")
            self.setItem(row, insert_col, item)
        
        # 更新列标题
        if self.cell_config and self.cell_config.require_title:
            title = self.cell_config.get_column_title(insert_col)
            header_item = QTableWidgetItem(title)
            self.setHorizontalHeaderItem(insert_col, header_item)
        
        # 更新列配置
        if self.cell_config:
            self.cell_config.col_count = self.columnCount()
    
    def delete_current_row(self):
        """删除当前行"""
        current_item = self.currentItem()
        if not current_item:
            QMessageBox.warning(self, "提示", "请先选择一个单元格")
            return
        
        if self.rowCount() <= 1:
            QMessageBox.warning(self, "提示", "表格至少保留一行")
            return
        
        row = current_item.row()
        reply = QMessageBox.question(
            self, "确认删除",
            f"确定要删除第 {row+1} 行吗？\n该行的所有数据将被清除！",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            self.removeRow(row)
            if self.cell_config:
                self.cell_config.row_count = self.rowCount()
    
    def delete_current_col(self):
        """删除当前列"""
        current_item = self.currentItem()
        if not current_item:
            QMessageBox.warning(self, "提示", "请先选择一个单元格")
            return
        
        if self.columnCount() <= 1:
            QMessageBox.warning(self, "提示", "表格至少保留一列")
            return
        
        col = current_item.column()
        reply = QMessageBox.question(
            self, "确认删除",
            f"确定要删除第 {col+1} 列吗？\n该列的所有数据将被清除！",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            self.removeColumn(col)
            if self.cell_config:
                self.cell_config.col_count = self.columnCount()

    def _can_paste(self) -> bool:
        """检查是否可以粘贴"""
        clipboard = QApplication.clipboard()
        text = clipboard.text()
        has_external = bool(text) and self._is_valid_tsv(text)
        has_internal = bool(self._clipboard_data)
        return has_external or has_internal
    
    def _is_valid_tsv(self, text: str) -> bool:
        """检查文本是否为有效的TSV格式"""
        if not text:
            return False
        lines = text.strip().split('\n')
        return len(lines) > 0
    
    def _data_to_tsv(self, data: List[List[str]]) -> str:
        """将二维数据转换为TSV文本"""
        lines = []
        for row in data:
            # 转义制表符和换行符
            escaped_row = [cell.replace('\t', '\\t').replace('\n', '\\n') for cell in row]
            lines.append('\t'.join(escaped_row))
        return '\n'.join(lines)
    
    def _tsv_to_data(self, text: str) -> List[List[str]]:
        """将TSV文本转换为二维数据"""
        data = []
        lines = text.strip().split('\n')
        for line in lines:
            # 分割并反转义
            cells = line.split('\t')
            unescaped_cells = [cell.replace('\\t', '\t').replace('\\n', '\n') for cell in cells]
            data.append(unescaped_cells)
        return data
    
    def _expand_table(self, needed_rows: int, needed_cols: int):
        """扩展表格以容纳更多数据"""
        current_rows = self.rowCount()
        current_cols = self.columnCount()
        
        if needed_rows > current_rows:
            self.setRowCount(needed_rows)
            # 初始化新行 - 确保所有单元格都有 QTableWidgetItem
            for row in range(current_rows, needed_rows):
                for col in range(self.columnCount()):
                    if not self.item(row, col):
                        item = QTableWidgetItem("")
                        self.setItem(row, col, item)
        
        if needed_cols > current_cols:
            self.setColumnCount(needed_cols)
            # 初始化新列
            for col in range(current_cols, needed_cols):
                # 为所有行创建空项目
                for row in range(self.rowCount()):
                    if not self.item(row, col):
                        item = QTableWidgetItem("")
                        self.setItem(row, col, item)
                
                # 为新列设置标题
                if self.cell_config and self.cell_config.require_title:
                    if self.cell_config.title_mode == TitleMode.UNIFIED.value:
                        title = self.cell_config.title_text or ""
                    else:
                        title = self.cell_config.get_column_title(col)
                    header_item = QTableWidgetItem(title)
                    self.setHorizontalHeaderItem(col, header_item)
    
    def _clip_data_to_table(self, data: List[List[str]], start_row: int, start_col: int) -> List[List[str]]:
        """裁剪数据以适应表格大小"""
        max_rows = self.rowCount() - start_row
        max_cols = self.columnCount() - start_col
        
        if max_rows <= 0 or max_cols <= 0:
            return []
        
        clipped_data = []
        for i, row in enumerate(data):
            if i >= max_rows:
                break
            clipped_row = row[:max_cols]
            clipped_data.append(clipped_row)
        
        return clipped_data

    def on_item_changed(self, item: QTableWidgetItem):
        """数据变化处理（添加撤销支持）"""
        if self._is_programmatic_change:
            return
            
        if item and self.user_id is not None and self.cell_index is not None:
            row = item.row()
            col = item.column()
            new_text = item.text()
            
            self.apply_highlight_to_item(item)
            
            # 使用之前记录的旧文本创建撤销命令
            old_text = self._old_text_before_edit
            if old_text != new_text:
                self._undo_stack.push(
                    CellEditCommand(self, self.cell_index, row, col, 
                                  old_text, new_text, f"编辑单元格 ({row+1},{col+1})")
                )
            
            self.data_changed.emit(self.cell_index, row, col, new_text)
            
            if self.auto_resize_columns:
                self.resizeColumnToContents(col)
    
    def editItem(self, item):
        """重写editItem以记录编辑前的文本"""
        if item:
            self._old_text_before_edit = item.text()
        super().editItem(item)

    
    

    def configure(self, user_id: int, cell_index: int, config: CellConfig, auto_resize: bool = True):
        """配置表格"""
        self.user_id = user_id
        self.cell_index = cell_index
        self.cell_config = config
        self.auto_resize_columns = auto_resize
        
        # 设置表格大小
        self.setRowCount(config.row_count)
        self.setColumnCount(config.col_count)
        
        # 设置标题
        self._update_headers()
        
        # 设置字体
        font = QFont()
        font.setPointSize(10)
        self.setFont(font)
        
        # 设置列宽自动调整模式
        if self.auto_resize_columns:
            self.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        else:
            self.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
            
    def _update_headers(self):
        """更新表头标题"""
        if not self.cell_config or not self.cell_config.require_title:
            # 如果不需要标题，隐藏表头
            self.horizontalHeader().setVisible(False)
            return
            
        self.horizontalHeader().setVisible(True)
        
        if self.cell_config.title_mode == TitleMode.UNIFIED.value:
            # 统一标题模式：所有列使用相同标题
            title = self.cell_config.title_text or ""
            for col in range(self.columnCount()):
                header_item = QTableWidgetItem(title)
                self.setHorizontalHeaderItem(col, header_item)
        else:
            # 独立标题模式：每列使用独立标题
            for col in range(self.columnCount()):
                title = self.cell_config.get_column_title(col)
                header_item = QTableWidgetItem(title)
                self.setHorizontalHeaderItem(col, header_item)
                
    def set_column_title(self, col: int, title: str):
        """设置指定列的标题"""
        if self.cell_config:
            self.cell_config.set_column_title(col, title)
            header_item = QTableWidgetItem(title)
            self.setHorizontalHeaderItem(col, header_item)
            
    def get_column_title(self, col: int) -> str:
        """获取指定列的标题"""
        if self.cell_config:
            return self.cell_config.get_column_title(col)
        return ""
        
    def update_column_titles(self, titles: List[str]):
        """批量更新列标题"""
        if self.cell_config and self.cell_config.require_title:
            self.cell_config.title_mode = TitleMode.INDEPENDENT.value
            for col, title in enumerate(titles):
                if col < self.columnCount():
                    self.cell_config.column_titles[col] = title
                    header_item = QTableWidgetItem(title)
                    self.setHorizontalHeaderItem(col, header_item)
                    
    def set_highlight(self, keyword: str = None):
        """设置高亮关键词"""
        self.highlight_keyword = keyword
        self.update_all_items_highlight()
        
    def clear_highlight(self):
        """清除高亮"""
        self.highlight_keyword = None
        self.update_all_items_highlight()
        
    def update_all_items_highlight(self):
        """更新所有单元格的高亮"""
        for row in range(self.rowCount()):
            for col in range(self.columnCount()):
                item = self.item(row, col)
                if item:
                    self.apply_highlight_to_item(item)
                    
    def apply_highlight_to_item(self, item: QTableWidgetItem):
        """对单个单元格应用高亮"""
        text = item.text()
        if not text:
            return
            
        if self.highlight_keyword and self.highlight_keyword.lower() in text.lower():
            # 高亮显示
            item.setBackground(QColor("#ffeb3b"))  # 黄色背景
            item.setForeground(QColor("#000000"))  # 黑色文字
            font = item.font()
            font.setBold(True)
            item.setFont(font)
        else:
            # 正常显示 - 交替行颜色
            if item.row() % 2 == 0:
                item.setBackground(QColor("#ffffff"))
            else:
                item.setBackground(QColor("#f5f5f5"))
            item.setForeground(QColor("#000000"))
            font = item.font()
            font.setBold(False)
            item.setFont(font)
                    
    def load_data(self, records: List[Dict]):
        """加载数据"""
        self.blockSignals(True)
        
        for record in records:
            if record["cell_index"] == self.cell_index:
                row = record["row_index"]
                col = record["col_index"]
                content = record["content"]
                
                if row < self.rowCount() and col < self.columnCount():
                    item = self.item(row, col)
                    if not item:
                        item = QTableWidgetItem()
                        self.setItem(row, col, item)
                    item.setText(content or "")
                    self.apply_highlight_to_item(item)
                    
        self.blockSignals(False)
        
        # 加载完成后自动调整列宽
        if self.auto_resize_columns:
            self.resize_columns()
            
    def resize_columns(self):
        """调整所有列宽以适应内容"""
        for col in range(self.columnCount()):
            self.resizeColumnToContents(col)
        
    def set_auto_resize(self, enabled: bool):
        """设置是否自动调整列宽"""
        self.auto_resize_columns = enabled
        if enabled:
            self.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
            self.resize_columns()
        else:
            self.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        
    def calculate_days_diff(self, login_time: datetime) -> Optional[int]:
        """计算与登录时间相隔的天数"""
        if self.cell_config and self.cell_config.calculate_days_diff:
            now = datetime.now()
            diff = (now - login_time).days
            return diff
        return None


class ColumnTitleDialog(QDialog):
    """列标题设置对话框"""
    
    def __init__(self, table_widget: DataTableWidget, parent=None):
        super().__init__(parent)
        self.table_widget = table_widget
        self.setup_ui()
        
    def setup_ui(self):
        """设置UI"""
        self.setWindowTitle("设置列标题")
        self.resize(400, 300)
        
        # 设置窗口图标
        self.setWindowIcon(get_app_icon())
        
        layout = QVBoxLayout()
        
        # 说明
        info_label = QLabel("为每一列设置独立的标题：")
        layout.addWidget(info_label)
        
        # 滚动区域
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout(scroll_widget)
        
        # 为每一列创建标题输入框
        self.title_edits = []
        col_count = self.table_widget.columnCount()
        
        for col in range(col_count):
            group = QGroupBox(f"第 {col + 1} 列")
            group_layout = QHBoxLayout()
            
            label = QLabel("标题:")
            edit = QLineEdit()
            edit.setText(self.table_widget.get_column_title(col))
            edit.setPlaceholderText(f"请输入第{col + 1}列的标题")
            
            group_layout.addWidget(label)
            group_layout.addWidget(edit)
            group.setLayout(group_layout)
            
            scroll_layout.addWidget(group)
            self.title_edits.append(edit)
            
        scroll_layout.addStretch()
        scroll.setWidget(scroll_widget)
        layout.addWidget(scroll)
        
        # 按钮
        button_layout = QHBoxLayout()
        
        # 统一设置按钮
        self.unified_btn = QPushButton("统一设置")
        self.unified_btn.clicked.connect(self.set_unified_titles)
        button_layout.addWidget(self.unified_btn)
        
        button_layout.addStretch()
        
        self.ok_btn = QPushButton("确定")
        self.ok_btn.clicked.connect(self.apply_titles)
        button_layout.addWidget(self.ok_btn)
        
        self.cancel_btn = QPushButton("取消")
        self.cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(self.cancel_btn)
        
        layout.addLayout(button_layout)
        
        self.setLayout(layout)
        
    def set_unified_titles(self):
        """统一设置所有列标题"""
        text, ok = QInputDialog.getText(self, "统一标题", "请输入统一的列标题:")
        if ok and text:
            for edit in self.title_edits:
                edit.setText(text)
                
    def apply_titles(self):
        """应用标题设置"""
        titles = [edit.text() for edit in self.title_edits]
        self.table_widget.update_column_titles(titles)
        self.accept()


class UserConfigDialog(QDialog):
    """用户配置对话框类"""
    
    def __init__(self, db_manager: DatabaseManager, user_id: int = None, parent=None):
        super().__init__(parent)
        self.db_manager = db_manager
        self.user_id = user_id
        self._existing_cell_configs = {}  # 已有各单元格配置（保留每个单元格独立的标题等）
        self.cell_title_edits = []        # 各单元格标题输入框
        self.setup_ui()
        
    def setup_ui(self):
        """设置UI"""
        self.setWindowTitle("用户配置")
        
        # 设置窗口图标
        self.setWindowIcon(get_app_icon())
        
        layout = QVBoxLayout()
        
        # 单元格数量配置
        cell_count_group = QGroupBox("单元格数量配置")
        cell_count_layout = QFormLayout()
        
        # 单元格数量
        self.cell_count_spin = QSpinBox()
        self.cell_count_spin.setRange(1, 100)
        self.cell_count_spin.setValue(6)
        self.cell_count_spin.setToolTip("设置需要多少个独立的单元格（每个单元格是一个独立的表格）")
        cell_count_layout.addRow("单元格数量:", self.cell_count_spin)
        
        # 添加说明标签
        hint_label = QLabel("说明：每个单元格是一个独立的表格，可通过下拉框切换编辑不同的单元格")
        hint_label.setStyleSheet("color: gray; font-size: 10px;")
        cell_count_layout.addRow(hint_label)
        
        cell_count_group.setLayout(cell_count_layout)
        layout.addWidget(cell_count_group)
        
        # 界面设置
        ui_group = QGroupBox("界面设置")
        ui_layout = QVBoxLayout()
        
        self.auto_resize_cb = QCheckBox("根据填充内容自动调整列宽")
        self.auto_resize_cb.setChecked(True)
        ui_layout.addWidget(self.auto_resize_cb)
        
        ui_group.setLayout(ui_layout)
        layout.addWidget(ui_group)
        
        # 单元格默认配置
        cell_group = QGroupBox("单元格默认配置")
        cell_layout = QFormLayout()
        
        self.cell_rows_spin = QSpinBox()
        self.cell_rows_spin.setRange(1, 1000)
        self.cell_rows_spin.setValue(3)
        cell_layout.addRow("每个单元格的行数:", self.cell_rows_spin)
        
        self.cell_cols_spin = QSpinBox()
        self.cell_cols_spin.setRange(1, 1000)
        self.cell_cols_spin.setValue(4)
        cell_layout.addRow("每个单元格的列数:", self.cell_cols_spin)
        
        self.require_login_time_cb = QCheckBox("需要登录时间")
        cell_layout.addRow(self.require_login_time_cb)
        
        self.calculate_days_cb = QCheckBox("计算相隔天数")
        cell_layout.addRow(self.calculate_days_cb)
        
        self.require_title_cb = QCheckBox("需要标题")
        self.require_title_cb.toggled.connect(self.on_require_title_changed)
        cell_layout.addRow(self.require_title_cb)
        
        # 标题模式选择
        self.title_mode_combo = QComboBox()
        self.title_mode_combo.addItem(TitleMode.UNIFIED.value, TitleMode.UNIFIED.value)
        self.title_mode_combo.addItem(TitleMode.INDEPENDENT.value, TitleMode.INDEPENDENT.value)
        self.title_mode_combo.setToolTip("统一标题：所有列使用相同标题；独立标题：每列可单独设置标题")
        cell_layout.addRow("标题模式:", self.title_mode_combo)
        
        # 统一标题文本（仅在统一标题模式下显示）
        self.title_text_edit = QLineEdit()
        self.title_text_edit.setPlaceholderText("统一标题文本")
        cell_layout.addRow("统一标题文本:", self.title_text_edit)
        
        cell_group.setLayout(cell_layout)
        layout.addWidget(cell_group)
        
        # 各单元格标题配置（不同单元格可以使用不同标题）
        cell_titles_group = QGroupBox("各单元格标题")
        cell_titles_group.setToolTip("为每个单元格设置独立的标题，不同单元格可以使用不同的标题")
        cell_titles_layout = QVBoxLayout()
        
        titles_hint = QLabel("说明：每个单元格可以设置独立的标题（如“登记表”、“考勤表”等），互不影响")
        titles_hint.setStyleSheet("color: gray; font-size: 10px;")
        cell_titles_layout.addWidget(titles_hint)
        
        self.cell_titles_scroll = QScrollArea()
        self.cell_titles_scroll.setWidgetResizable(True)
        self.cell_titles_scroll.setMaximumHeight(200)
        self.cell_titles_widget = QWidget()
        self.cell_titles_form = QFormLayout(self.cell_titles_widget)
        self.cell_titles_scroll.setWidget(self.cell_titles_widget)
        cell_titles_layout.addWidget(self.cell_titles_scroll)
        
        cell_titles_group.setLayout(cell_titles_layout)
        layout.addWidget(cell_titles_group)
        
        # 提示信息
        notice_label = QLabel("提示: 行列数设置过大会影响性能，请根据实际需求设置")
        notice_label.setStyleSheet("color: orange; font-size: 10px;")
        layout.addWidget(notice_label)
        
        # 按钮
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        
        self.setLayout(layout)
        
        # 连接信号
        self.title_mode_combo.currentTextChanged.connect(self.on_title_mode_changed)
        self.cell_count_spin.valueChanged.connect(self.rebuild_cell_titles)
        
        # 初始化各单元格标题输入框
        self.rebuild_cell_titles()
        
        # 加载现有配置
        if self.user_id:
            self.load_config()
            
        # 初始化控件状态
        self.on_require_title_changed(self.require_title_cb.isChecked())
        self.on_title_mode_changed(self.title_mode_combo.currentText())
        
    def on_require_title_changed(self, checked: bool):
        """是否需要标题变化"""
        self.title_mode_combo.setEnabled(checked)
        self.title_text_edit.setEnabled(checked and self.title_mode_combo.currentText() == TitleMode.UNIFIED.value)
        
    def on_title_mode_changed(self, mode: str):
        """标题模式变化"""
        self.title_text_edit.setEnabled(
            self.require_title_cb.isChecked() and mode == TitleMode.UNIFIED.value
        )
            
    def rebuild_cell_titles(self):
        """根据单元格数量重建各单元格标题输入框"""
        # 先记住已填写的标题，避免重建时丢失
        current_titles = []
        for edit in self.cell_title_edits:
            current_titles.append(edit.text())
        self.cell_title_edits = []
        
        # 清空原有表单行
        while self.cell_titles_form.rowCount() > 0:
            self.cell_titles_form.removeRow(0)
        
        count = self.cell_count_spin.value()
        for i in range(count):
            label = QLabel(f"单元格 {i + 1}:")
            edit = QLineEdit()
            edit.setPlaceholderText(f"请输入单元格 {i + 1} 的标题（默认为“单元格 {i + 1}”）")
            if i < len(current_titles):
                edit.setText(current_titles[i])
            self.cell_titles_form.addRow(label, edit)
            self.cell_title_edits.append(edit)
            
    def load_config(self):
        """加载现有配置"""
        # 从数据库加载用户配置
        user = self.db_manager.get_user_by_id(self.user_id)
        if user and user.get("config"):
            config = user["config"]
            # 保留已有的各单元格配置（避免编辑时丢失每个单元格独立的标题等配置）
            self._existing_cell_configs = config.get("cell_configs", {}) or {}
            self.cell_count_spin.setValue(config.get("cell_count", 6))
            self.auto_resize_cb.setChecked(config.get("auto_resize_columns", True))
            default_cell = config.get("default_cell", {})
            self.cell_rows_spin.setValue(default_cell.get("row_count", 3))
            self.cell_cols_spin.setValue(default_cell.get("col_count", 4))
            self.require_login_time_cb.setChecked(default_cell.get("require_login_time", False))
            self.calculate_days_cb.setChecked(default_cell.get("calculate_days_diff", False))
            self.require_title_cb.setChecked(default_cell.get("require_title", True))
            
            # 加载标题模式
            title_mode = default_cell.get("title_mode", TitleMode.UNIFIED.value)
            index = self.title_mode_combo.findData(title_mode)
            if index >= 0:
                self.title_mode_combo.setCurrentIndex(index)
            else:
                self.title_mode_combo.setCurrentIndex(0)
                
            self.title_text_edit.setText(default_cell.get("title_text", ""))
            
            # 填充各单元格标题（有独立配置的取原标题，否则使用默认的“单元格 N”)
            self.rebuild_cell_titles()
            for i, edit in enumerate(self.cell_title_edits):
                key = str(i)
                existing_cfg = self._existing_cell_configs.get(key, {})
                edit.setText(existing_cfg.get("title_text") or f"单元格 {i + 1}")
            
    def get_config(self) -> Dict:
        """获取配置"""
        # 默认单元格配置
        default_cell = {
            "row_count": self.cell_rows_spin.value(),
            "col_count": self.cell_cols_spin.value(),
            "require_login_time": self.require_login_time_cb.isChecked(),
            "calculate_days_diff": self.calculate_days_cb.isChecked(),
            "require_title": self.require_title_cb.isChecked(),
            "title_mode": self.title_mode_combo.currentData(),
            "title_text": self.title_text_edit.text(),
            "column_titles": {},
            "title_unique": self.title_mode_combo.currentData() == TitleMode.UNIFIED.value,
            "scheme_type": CellScheme.STANDARD.value
        }
        
        # 构建各单元格配置：保留已有的独立配置，并应用各单元格标题
        cell_configs = {}
        for i in range(self.cell_count_spin.value()):
            key = str(i)
            existing = dict(self._existing_cell_configs.get(key, {}) or {})
            new_title = ""
            if i < len(self.cell_title_edits):
                new_title = self.cell_title_edits[i].text().strip()
            
            if existing:
                # 已有独立配置：保留原有配置，仅更新标题
                cell_cfg = dict(default_cell)
                cell_cfg.update(existing)
                cell_cfg["title_text"] = new_title
                # 未设置独立列标题时，使用统一标题模式以显示标题
                if not cell_cfg.get("column_titles"):
                    cell_cfg["title_mode"] = TitleMode.UNIFIED.value
                cell_configs[key] = cell_cfg
            elif new_title:
                # 无独立配置但设置了标题：基于默认配置创建
                cell_cfg = dict(default_cell)
                cell_cfg["title_text"] = new_title
                cell_cfg["title_mode"] = TitleMode.UNIFIED.value
                cell_configs[key] = cell_cfg
            # 否则不创建独立配置，加载时使用默认标题（单元格 N）
        
        return {
            "cell_count": self.cell_count_spin.value(),
            "auto_resize_columns": self.auto_resize_cb.isChecked(),
            "default_cell": default_cell,
            "cell_configs": cell_configs
        }


class CellConfigDialog(QDialog):
    """单个单元格配置对话框"""
    
    def __init__(self, cell_config: CellConfig, cell_index: int, parent=None):
        super().__init__(parent)
        self.cell_config = cell_config
        self.cell_index = cell_index
        self.setup_ui()
        
    def setup_ui(self):
        """设置UI"""
        self.setWindowTitle(f"单元格 {self.cell_index + 1} 配置")
        
        # 设置窗口图标
        self.setWindowIcon(get_app_icon())
        
        layout = QVBoxLayout()
        
        # 表格配置
        table_group = QGroupBox("表格配置")
        table_layout = QFormLayout()
        
        self.row_spin = QSpinBox()
        self.row_spin.setRange(1, 1000)
        self.row_spin.setValue(self.cell_config.row_count)
        table_layout.addRow("行数:", self.row_spin)
        
        self.col_spin = QSpinBox()
        self.col_spin.setRange(1, 1000)
        self.col_spin.setValue(self.cell_config.col_count)
        table_layout.addRow("列数:", self.col_spin)
        
        table_group.setLayout(table_layout)
        layout.addWidget(table_group)
        
        # 标题配置
        title_group = QGroupBox("标题配置")
        title_layout = QFormLayout()
        
        self.require_title_cb = QCheckBox("需要标题")
        self.require_title_cb.setChecked(self.cell_config.require_title)
        title_layout.addRow(self.require_title_cb)
        
        self.title_mode_combo = QComboBox()
        self.title_mode_combo.addItem(TitleMode.UNIFIED.value, TitleMode.UNIFIED.value)
        self.title_mode_combo.addItem(TitleMode.INDEPENDENT.value, TitleMode.INDEPENDENT.value)
        self.title_mode_combo.setCurrentIndex(
            0 if self.cell_config.title_mode == TitleMode.UNIFIED.value else 1
        )
        title_layout.addRow("标题模式:", self.title_mode_combo)
        
        self.title_text_edit = QLineEdit()
        self.title_text_edit.setText(self.cell_config.title_text)
        self.title_text_edit.setPlaceholderText("统一标题文本")
        title_layout.addRow("统一标题:", self.title_text_edit)
        
        # 设置列标题按钮（仅在独立标题模式下显示）
        self.set_titles_btn = QPushButton("设置各列标题...")
        self.set_titles_btn.clicked.connect(self.on_set_column_titles)
        title_layout.addRow(self.set_titles_btn)
        
        title_group.setLayout(title_layout)
        layout.addWidget(title_group)
        
        # 功能配置
        func_group = QGroupBox("功能配置")
        func_layout = QVBoxLayout()
        
        self.require_login_time_cb = QCheckBox("记录登录时间")
        self.require_login_time_cb.setChecked(self.cell_config.require_login_time)
        func_layout.addWidget(self.require_login_time_cb)
        
        self.calculate_days_cb = QCheckBox("计算相隔天数")
        self.calculate_days_cb.setChecked(self.cell_config.calculate_days_diff)
        func_layout.addWidget(self.calculate_days_cb)
        
        func_group.setLayout(func_layout)
        layout.addWidget(func_group)
        
        # 按钮
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        
        self.setLayout(layout)
        
        # 连接信号
        self.require_title_cb.toggled.connect(self.on_require_title_toggled)
        self.title_mode_combo.currentTextChanged.connect(self.on_title_mode_changed)
        
        # 初始化控件状态
        self.on_require_title_toggled(self.require_title_cb.isChecked())
        self.on_title_mode_changed(self.title_mode_combo.currentText())
        
    def on_require_title_toggled(self, checked: bool):
        """是否需要标题切换"""
        self.title_mode_combo.setEnabled(checked)
        self.title_text_edit.setEnabled(checked and self.title_mode_combo.currentText() == TitleMode.UNIFIED.value)
        self.set_titles_btn.setEnabled(checked and self.title_mode_combo.currentText() == TitleMode.INDEPENDENT.value)
        
    def on_title_mode_changed(self, mode: str):
        """标题模式切换"""
        is_unified = mode == TitleMode.UNIFIED.value
        self.title_text_edit.setEnabled(self.require_title_cb.isChecked() and is_unified)
        self.set_titles_btn.setEnabled(self.require_title_cb.isChecked() and not is_unified)
        
    def on_set_column_titles(self):
        """设置各列标题"""
        QMessageBox.information(self, "提示", "请在主界面对应表格的菜单中设置各列标题")
        
    def get_config(self) -> CellConfig:
        """获取配置"""
        self.cell_config.row_count = self.row_spin.value()
        self.cell_config.col_count = self.col_spin.value()
        self.cell_config.require_login_time = self.require_login_time_cb.isChecked()
        self.cell_config.calculate_days_diff = self.calculate_days_cb.isChecked()
        self.cell_config.require_title = self.require_title_cb.isChecked()
        self.cell_config.title_mode = self.title_mode_combo.currentData()
        self.cell_config.title_text = self.title_text_edit.text()
        return self.cell_config


class LoginDialog(QDialog):
    """登录对话框"""
    
    login_success = Signal(dict)  # 登录成功信号
    
    def __init__(self, db_manager: DatabaseManager, parent=None):
        super().__init__(parent)
        self.db_manager = db_manager
        self.settings = AppSettings()
        self.setup_ui()
        self.load_last_user()
        
    def setup_ui(self):
        """设置UI"""
        self.setWindowTitle(f"用户登录 - {ProjectInfo.get_full_name()}")
        self.setFixedSize(450, 380)
        
        # 设置窗口图标
        self.setWindowIcon(get_app_icon())
        
        layout = QVBoxLayout()
        
        # 标题
        title_label = QLabel(ProjectInfo.get_full_name())
        title_label.setAlignment(Qt.AlignCenter)
        title_font = QFont()
        title_font.setPointSize(18)
        title_font.setBold(True)
        title_label.setFont(title_font)
        layout.addWidget(title_label)
        
        layout.addSpacing(20)
        
        # 表单
        form_layout = QFormLayout()
        
        # 用户下拉选择
        self.user_combo = QComboBox()
        self.user_combo.setEditable(True)
        self.user_combo.setPlaceholderText("请选择或输入用户名")
        self.user_combo.currentTextChanged.connect(self.on_user_selected)
        form_layout.addRow("用户名:", self.user_combo)
        
        self.password_edit = QLineEdit()
        self.password_edit.setEchoMode(QLineEdit.Password)
        self.password_edit.setPlaceholderText("请输入密码（可为空）")
        form_layout.addRow("密码:", self.password_edit)
        
        layout.addLayout(form_layout)
        
        layout.addSpacing(10)
        
        # 提示信息
        hint_label = QLabel("提示: 密码可以为空，直接登录")
        hint_label.setStyleSheet("color: gray; font-size: 10px;")
        hint_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(hint_label)
        
        layout.addSpacing(10)
        
        # 记住用户名复选框
        self.remember_checkbox = QCheckBox("记住用户名")
        self.remember_checkbox.setChecked(True)
        layout.addWidget(self.remember_checkbox)
        
        layout.addSpacing(10)
        
        # 按钮
        button_layout = QHBoxLayout()
        
        self.login_btn = QPushButton("登录")
        self.login_btn.clicked.connect(self.on_login)
        self.login_btn.setDefault(True)
        button_layout.addWidget(self.login_btn)
        
        self.user_mgr_btn = QPushButton("用户管理")
        self.user_mgr_btn.clicked.connect(self.on_user_management)
        button_layout.addWidget(self.user_mgr_btn)
        
        layout.addLayout(button_layout)
        
        # 状态标签
        self.status_label = QLabel()
        self.status_label.setStyleSheet("color: red;")
        layout.addWidget(self.status_label)
        
        self.setLayout(layout)
        
        # 加载用户列表
        self.load_user_list()
        
    def load_user_list(self):
        """加载用户列表"""
        self.user_combo.clear()
        users = self.db_manager.get_all_users()
        for user in users:
            self.user_combo.addItem(user['username'])
            
    def load_last_user(self):
        """加载上次登录的用户"""
        last_user = self.settings.value("last_user", "")
        if last_user:
            index = self.user_combo.findText(last_user)
            if index >= 0:
                self.user_combo.setCurrentIndex(index)
            else:
                self.user_combo.setEditText(last_user)
                
    def save_last_user(self, username):
        """保存最后登录的用户"""
        if self.remember_checkbox.isChecked():
            self.settings.setValue("last_user", username)
        else:
            self.settings.setValue("last_user", "")
            
    def on_user_selected(self, username):
        """用户选择变化时清空状态"""
        self.status_label.setText("")
        
    def on_login(self):
        """登录处理"""
        username = self.user_combo.currentText().strip()
        password = self.password_edit.text()
        
        if not username:
            self.status_label.setText("请输入用户名")
            return
            
        user_data = self.db_manager.verify_user(username, password)
        
        if user_data:
            self.save_last_user(username)
            self.login_success.emit(user_data)
            self.accept()
        else:
            self.status_label.setText("用户名或密码错误")
            
    def on_user_management(self):
        """打开用户管理"""
        dialog = UserManagementDialog(self.db_manager, self)
        if dialog.exec():
            # 刷新用户列表
            self.load_user_list()
            # 不清空表单


class UserManagementDialog(QDialog):
    """用户管理对话框"""
    
    def __init__(self, db_manager: DatabaseManager, parent=None):
        super().__init__(parent)
        self.db_manager = db_manager
        self.setup_ui()
        
    def setup_ui(self):
        """设置UI"""
        self.setWindowTitle("用户管理")
        self.resize(600, 450)
        
        # 设置窗口图标
        self.setWindowIcon(get_app_icon())
        
        layout = QVBoxLayout()
        
        # 用户列表
        self.user_list = QListWidget()
        self.user_list.itemClicked.connect(self.on_user_selected)
        layout.addWidget(self.user_list)
        
        # 按钮布局
        button_layout = QHBoxLayout()
        
        self.add_btn = QPushButton("添加用户")
        self.add_btn.clicked.connect(self.on_add_user)
        button_layout.addWidget(self.add_btn)
        
        self.edit_btn = QPushButton("编辑用户名")
        self.edit_btn.clicked.connect(self.on_edit_user)
        button_layout.addWidget(self.edit_btn)
        
        self.delete_btn = QPushButton("删除用户")
        self.delete_btn.clicked.connect(self.on_delete_user)
        button_layout.addWidget(self.delete_btn)
        
        self.change_pwd_btn = QPushButton("修改密码")
        self.change_pwd_btn.clicked.connect(self.on_change_password)
        button_layout.addWidget(self.change_pwd_btn)
        
        self.config_btn = QPushButton("编辑配置")
        self.config_btn.clicked.connect(self.on_edit_config)
        button_layout.addWidget(self.config_btn)
        
        button_layout.addStretch()
        
        self.close_btn = QPushButton("关闭")
        self.close_btn.clicked.connect(self.accept)
        button_layout.addWidget(self.close_btn)
        
        layout.addLayout(button_layout)
        
        self.setLayout(layout)
        self.load_users()
        
    def load_users(self):
        """加载用户列表"""
        self.user_list.clear()
        users = self.db_manager.get_all_users()
        for user in users:
            item = QListWidgetItem(f"{user['username']} (创建于: {user['created_at']})")
            item.setData(Qt.UserRole, user)
            self.user_list.addItem(item)
            
    def on_user_selected(self, item: QListWidgetItem):
        """用户选择处理"""
        self.current_user = item.data(Qt.UserRole)
        
    def on_add_user(self):
        """添加用户"""
        dialog = UserAddEditDialog(self.db_manager, parent=self)
        if dialog.exec():
            username = dialog.username_edit.text().strip()
            password = dialog.password_edit.text()
            confirm_password = dialog.confirm_edit.text()
            
            if not username:
                QMessageBox.warning(self, "提示", "用户名不能为空")
                return
                
            if password != confirm_password:
                QMessageBox.warning(self, "提示", "两次输入的密码不一致")
                return
                
            # 检查用户名是否已存在
            existing = self.db_manager.get_user(username)
            if existing:
                QMessageBox.warning(self, "提示", "用户名已存在")
                return
                
            # 创建用户
            config = {
                "cell_count": 6,
                "auto_resize_columns": True,
                "default_cell": {
                    "row_count": 3,
                    "col_count": 4,
                    "require_login_time": True,
                    "calculate_days_diff": True,
                    "require_title": True,
                    "title_mode": TitleMode.UNIFIED.value,
                    "title_text": "",
                    "column_titles": {},
                    "title_unique": True,
                    "scheme_type": CellScheme.STANDARD.value
                },
                "cell_configs": {}
            }
            
            if self.db_manager.create_user(username, password, config):
                QMessageBox.information(self, "成功", f"用户 {username} 创建成功")
                self.load_users()
            else:
                QMessageBox.warning(self, "错误", "创建用户失败")
                
    def on_edit_user(self):
        """编辑用户名"""
        if not hasattr(self, 'current_user') or not self.current_user:
            QMessageBox.warning(self, "提示", "请先选择要编辑的用户")
            return
            
        # 防止编辑admin用户
        if self.current_user["username"] == "admin":
            QMessageBox.warning(self, "提示", "不能编辑管理员用户")
            return
            
        new_username, ok = QInputDialog.getText(
            self, "编辑用户名", "请输入新的用户名:", text=self.current_user["username"]
        )
        
        if ok and new_username and new_username != self.current_user["username"]:
            # 检查新用户名是否已存在
            existing = self.db_manager.get_user(new_username)
            if existing:
                QMessageBox.warning(self, "提示", "用户名已存在")
                return
                
            if self.db_manager.update_username(self.current_user["id"], new_username):
                QMessageBox.information(self, "成功", "用户名更新成功")
                self.load_users()
            else:
                QMessageBox.warning(self, "错误", "更新用户名失败")
                
    def on_edit_config(self):
        """编辑用户配置"""
        if not hasattr(self, 'current_user') or not self.current_user:
            QMessageBox.warning(self, "提示", "请先选择要编辑的用户")
            return
            
        dialog = UserConfigDialog(self.db_manager, self.current_user["id"], self)
        if dialog.exec():
            config = dialog.get_config()
            if self.db_manager.update_user_config(self.current_user["id"], config):
                QMessageBox.information(self, "成功", "配置已更新")
            else:
                QMessageBox.warning(self, "错误", "保存配置失败")
                
    def on_delete_user(self):
        """删除用户"""
        if not hasattr(self, 'current_user') or not self.current_user:
            QMessageBox.warning(self, "提示", "请先选择要删除的用户")
            return
            
        # 防止删除admin用户
        if self.current_user["username"] == "admin":
            QMessageBox.warning(self, "提示", "不能删除管理员用户")
            return
            
        reply = QMessageBox.question(
            self, "确认删除", 
            f"确定要删除用户 {self.current_user['username']} 吗？\n此操作不可撤销，该用户的所有数据将被删除！",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            if self.db_manager.delete_user(self.current_user["id"]):
                QMessageBox.information(self, "成功", "用户删除成功")
                self.load_users()
                self.current_user = None
            else:
                QMessageBox.warning(self, "错误", "删除用户失败")
                
    def on_change_password(self):
        """修改用户密码"""
        if not hasattr(self, 'current_user') or not self.current_user:
            QMessageBox.warning(self, "提示", "请先选择要修改密码的用户")
            return
            
        dialog = ChangePasswordDialog(self.db_manager, self.current_user["id"], self)
        if dialog.exec():
            QMessageBox.information(self, "成功", "密码修改成功")


class UserAddEditDialog(QDialog):
    """用户添加对话框"""
    
    def __init__(self, db_manager: DatabaseManager, parent=None):
        super().__init__(parent)
        self.db_manager = db_manager
        self.setup_ui()
        
    def setup_ui(self):
        """设置UI"""
        self.setWindowTitle("添加用户")
        self.setFixedSize(350, 280)
        
        # 设置窗口图标
        self.setWindowIcon(get_app_icon())
        
        layout = QVBoxLayout()
        
        form_layout = QFormLayout()
        
        self.username_edit = QLineEdit()
        self.username_edit.setPlaceholderText("请输入用户名")
        form_layout.addRow("用户名:", self.username_edit)
        
        self.password_edit = QLineEdit()
        self.password_edit.setEchoMode(QLineEdit.Password)
        self.password_edit.setPlaceholderText("请输入密码（可为空）")
        form_layout.addRow("密码:", self.password_edit)
        
        self.confirm_edit = QLineEdit()
        self.confirm_edit.setEchoMode(QLineEdit.Password)
        self.confirm_edit.setPlaceholderText("请再次输入密码")
        form_layout.addRow("确认密码:", self.confirm_edit)
        
        layout.addLayout(form_layout)
        
        # 提示信息
        hint_label = QLabel("提示: 密码可以为空，直接登录")
        hint_label.setStyleSheet("color: gray; font-size: 10px;")
        layout.addWidget(hint_label)
        
        layout.addSpacing(20)
        
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        
        self.setLayout(layout)


class ChangePasswordDialog(QDialog):
    """修改密码对话框"""
    
    def __init__(self, db_manager: DatabaseManager, user_id: int, parent=None):
        super().__init__(parent)
        self.db_manager = db_manager
        self.user_id = user_id
        self.setup_ui()
        
    def setup_ui(self):
        """设置UI"""
        self.setWindowTitle("修改密码")
        self.setFixedSize(350, 220)
        
        # 设置窗口图标
        self.setWindowIcon(get_app_icon())
        
        layout = QVBoxLayout()
        
        form_layout = QFormLayout()
        
        self.new_password_edit = QLineEdit()
        self.new_password_edit.setEchoMode(QLineEdit.Password)
        self.new_password_edit.setPlaceholderText("请输入新密码（可为空）")
        form_layout.addRow("新密码:", self.new_password_edit)
        
        self.confirm_edit = QLineEdit()
        self.confirm_edit.setEchoMode(QLineEdit.Password)
        self.confirm_edit.setPlaceholderText("请再次输入新密码")
        form_layout.addRow("确认密码:", self.confirm_edit)
        
        layout.addLayout(form_layout)
        
        # 提示信息
        hint_label = QLabel("提示: 留空表示不设密码")
        hint_label.setStyleSheet("color: gray; font-size: 10px;")
        layout.addWidget(hint_label)
        
        layout.addSpacing(20)
        
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.on_confirm)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        
        self.setLayout(layout)
        
    def on_confirm(self):
        """确认修改"""
        new_password = self.new_password_edit.text()
        confirm = self.confirm_edit.text()
        
        if new_password != confirm:
            QMessageBox.warning(self, "提示", "两次输入的密码不一致")
            return
            
        if self.db_manager.update_user_password(self.user_id, new_password):
            self.accept()
        else:
            QMessageBox.warning(self, "错误", "修改密码失败")


class BackupRestoreDialog(QDialog):
    """备份恢复对话框类"""
    
    backup_restored = Signal()  # 备份恢复信号
    
    def __init__(self, backup_manager: BackupManager, parent=None):
        super().__init__(parent)
        self.backup_manager = backup_manager
        self.setup_ui()
        
    def setup_ui(self):
        """设置UI"""
        self.setWindowTitle("数据库备份恢复")
        self.resize(800, 600)
        
        # 设置窗口图标
        self.setWindowIcon(get_app_icon())
        
        layout = QVBoxLayout()
        
        # 筛选区域
        filter_group = QGroupBox("筛选")
        filter_layout = QHBoxLayout()
        
        filter_layout.addWidget(QLabel("备份类型:"))
        self.type_combo = QComboBox()
        self.type_combo.addItems(["全部", "manual", "auto", "rollback"])
        self.type_combo.currentTextChanged.connect(self.refresh_backup_list)
        filter_layout.addWidget(self.type_combo)
        
        filter_layout.addWidget(QLabel("开始时间:"))
        self.start_date = QDateTimeEdit()
        self.start_date.setCalendarPopup(True)
        self.start_date.setDateTime(QDateTime.currentDateTime().addDays(-30))
        filter_layout.addWidget(self.start_date)
        
        filter_layout.addWidget(QLabel("结束时间:"))
        self.end_date = QDateTimeEdit()
        self.end_date.setCalendarPopup(True)
        self.end_date.setDateTime(QDateTime.currentDateTime())
        filter_layout.addWidget(self.end_date)
        
        self.filter_btn = QPushButton("应用筛选")
        self.filter_btn.clicked.connect(self.refresh_backup_list)
        filter_layout.addWidget(self.filter_btn)
        
        filter_group.setLayout(filter_layout)
        layout.addWidget(filter_group)
        
        # 备份列表
        self.backup_tree = QTreeWidget()
        self.backup_tree.setHeaderLabels(["备份时间", "备份类型", "文件大小", "数据库版本", "排班日期范围", "操作"])
        self.backup_tree.setAlternatingRowColors(True)
        self.backup_tree.itemDoubleClicked.connect(self.on_backup_selected)
        layout.addWidget(self.backup_tree)
        
        # 详细信息预览
        detail_group = QGroupBox("详细信息")
        detail_layout = QVBoxLayout()
        
        self.detail_text = QTextEdit()
        self.detail_text.setReadOnly(True)
        detail_layout.addWidget(self.detail_text)
        
        detail_group.setLayout(detail_layout)
        layout.addWidget(detail_group)
        
        # 按钮
        button_layout = QHBoxLayout()
        
        self.restore_btn = QPushButton("恢复选中备份")
        self.restore_btn.clicked.connect(self.on_restore_backup)
        button_layout.addWidget(self.restore_btn)
        
        self.backup_btn = QPushButton("手动备份")
        self.backup_btn.clicked.connect(self.on_manual_backup)
        button_layout.addWidget(self.backup_btn)
        
        self.refresh_btn = QPushButton("刷新")
        self.refresh_btn.clicked.connect(self.refresh_backup_list)
        button_layout.addWidget(self.refresh_btn)
        
        layout.addLayout(button_layout)
        
        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)
        
        self.setLayout(layout)
        
        # 加载备份列表
        self.refresh_backup_list()
        
    def refresh_backup_list(self):
        """刷新备份列表（后台线程）"""
        self.backup_tree.clear()
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)
        
        self._refresh_worker = BackgroundWorker(self)
        self._refresh_worker.setup(self._do_refresh_backup_list)
        self._refresh_worker.finished.connect(self._on_refresh_list_finished)
        self._refresh_worker.error.connect(self._on_bg_error)
        self._refresh_worker.start()
    
    def _do_refresh_backup_list(self):
        """在后台线程中获取备份列表并筛选"""
        backups = self.backup_manager.get_backup_list()
        backup_type = self.type_combo.currentText()
        start_datetime = self.start_date.dateTime()
        end_datetime = self.end_date.dateTime()
        
        filtered = []
        for backup in backups:
            if backup_type != "全部" and backup["type"] != backup_type:
                continue
            if backup["datetime"]:
                if backup["datetime"] < start_datetime or backup["datetime"] > end_datetime:
                    continue
            filtered.append(backup)
        return filtered
    
    def _on_refresh_list_finished(self, backups):
        """备份列表加载完成（主线程回调）"""
        self.progress_bar.setVisible(False)
                    
        for backup in backups:
            # 显示类型
            type_display = {
                "manual": "手动备份",
                "auto": "自动备份",
                "rollback": "回滚备份"
            }.get(backup["type"], backup["type"])
            
            item = QTreeWidgetItem([
                backup["time"],
                type_display,
                backup["size_str"],
                backup["info"].get("version", "未知"),
                backup["info"].get("date_range", "无数据") or "无数据",
                "恢复"
            ])
            
            # 存储备份信息
            item.setData(0, Qt.UserRole, backup)
            
            # 设置操作按钮
            restore_btn = QPushButton("恢复")
            restore_btn.clicked.connect(partial(self.on_restore_specific, backup["path"]))
            self.backup_tree.setItemWidget(item, 5, restore_btn)
            
            self.backup_tree.addTopLevelItem(item)
            
        # 调整列宽
        for i in range(5):
            self.backup_tree.resizeColumnToContents(i)
            
    def on_backup_selected(self, item: QTreeWidgetItem, column: int):
        """备份选中处理"""
        backup_data = item.data(0, Qt.UserRole)
        if backup_data:
            # 显示详细信息
            detail = f"""
备份文件: {backup_data['name']}
备份类型: {backup_data['type']}
备份时间: {backup_data['time']}
文件大小: {backup_data['size_str']}
数据库版本: {backup_data['info'].get('version', '未知')}
包含的表: {', '.join(backup_data['info'].get('tables', []))}
排班日期范围: {backup_data['info'].get('date_range', '无')}
            """
            self.detail_text.setText(detail)
            
    def on_restore_backup(self):
        """恢复选中备份"""
        current_item = self.backup_tree.currentItem()
        if not current_item:
            QMessageBox.warning(self, "提示", "请先选择要恢复的备份")
            return
            
        backup_data = current_item.data(0, Qt.UserRole)
        if backup_data:
            self.restore_backup(backup_data["path"])
            
    def on_restore_specific(self, backup_path: str):
        """恢复指定备份"""
        self.restore_backup(backup_path)
        
    def restore_backup(self, backup_path: str):
        """执行备份恢复（后台线程）"""
        reply = QMessageBox.question(
            self, "确认恢复", 
            "恢复操作将替换当前数据库，并创建回滚备份。\n确定要继续吗？",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            self.progress_bar.setVisible(True)
            self.progress_bar.setRange(0, 0)  # 不确定进度
            self.restore_btn.setEnabled(False)
            self.backup_btn.setEnabled(False)
            
            self._restore_worker = BackgroundWorker(self)
            self._restore_worker.setup(self.backup_manager.restore_backup, backup_path)
            self._restore_worker.finished.connect(self._on_restore_finished)
            self._restore_worker.error.connect(self._on_bg_error)
            self._restore_worker.start()
    
    def _on_restore_finished(self, result):
        """恢复完成（主线程回调）"""
        self.progress_bar.setVisible(False)
        self.restore_btn.setEnabled(True)
        self.backup_btn.setEnabled(True)
        
        success, msg = result
        if success:
            QMessageBox.information(self, "成功", f"数据库恢复成功！\n回滚备份已创建: {msg}")
            self.backup_restored.emit()
            self.refresh_backup_list()
        else:
            QMessageBox.critical(self, "错误", f"数据库恢复失败: {msg}")
                
    def on_manual_backup(self):
        """手动备份（后台线程）"""
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)
        self.backup_btn.setEnabled(False)
        self.restore_btn.setEnabled(False)
        
        self._backup_worker = BackgroundWorker(self)
        self._backup_worker.setup(self.backup_manager.create_backup, "manual")
        self._backup_worker.finished.connect(self._on_manual_backup_finished)
        self._backup_worker.error.connect(self._on_bg_error)
        self._backup_worker.start()
    
    def _on_manual_backup_finished(self, result):
        """手动备份完成（主线程回调）"""
        self.progress_bar.setVisible(False)
        self.backup_btn.setEnabled(True)
        self.restore_btn.setEnabled(True)
        
        success, backup_name, info = result
        if success:
            QMessageBox.information(self, "成功", f"手动备份成功: {backup_name}")
            self.refresh_backup_list()
        else:
            QMessageBox.critical(self, "错误", f"备份失败: {backup_name}")
    
    def _on_bg_error(self, error_msg: str):
        """后台任务出错处理"""
        self.progress_bar.setVisible(False)
        self.restore_btn.setEnabled(True)
        self.backup_btn.setEnabled(True)
        print(f"[备份/恢复错误] {error_msg}")
        QMessageBox.critical(self, "错误", f"操作失败:\n{error_msg.split(chr(10))[0]}")


class DataEntryWindow(QMainWindow):
    """数据录入主窗口"""
    
    def __init__(self, db_manager: DatabaseManager, user_data: Dict):
        super().__init__()
        self.db_manager = db_manager
        self.user_data = user_data
        self.current_user = user_data
        self.backup_manager = BackupManager(db_manager.db_path)
        self.current_user_config = user_data.get("config", {})
        self.debug_mode = False
        self.login_time = datetime.now()
        self.search_results = []  # 存储搜索结果
        self._tables_ready = False  # 标记表格是否已创建
        self.init_settings()
        self.setup_ui()
        # 恢复上次的窗口位置和大小
        self._restore_window_geometry()
        # 窗口先显示，延迟加载配置和数据（避免启动卡顿）
        QTimer.singleShot(30, self._startup_load)
        
    def init_settings(self):
        """初始化设置"""
        self.settings = AppSettings()
        self.debug_mode = self.settings.get("debug_mode", False)
        
    def setup_ui(self):
        """设置UI"""
        self.setWindowTitle(ProjectInfo.get_full_title(self.user_data['username']))
        self.resize(1400, 800)  # 增加宽度以容纳搜索面板
        
        # 设置窗口图标
        self.setWindowIcon(get_app_icon())
        
        # 创建菜单栏
        self.create_menu_bar()
        
        # 创建工具栏
        self.create_tool_bar()
        
        # 创建主分割器
        self.main_splitter = QSplitter(Qt.Horizontal)
        
        # 左侧内容区域
        left_widget = QWidget()
        left_layout = QVBoxLayout()
        left_layout.setContentsMargins(0, 0, 0, 0)
        
        # 用户信息栏
        info_layout = QHBoxLayout()
        self.user_label = QLabel(f"当前用户: {self.user_data['username']}")
        user_font = QFont()
        user_font.setBold(True)
        self.user_label.setFont(user_font)
        info_layout.addWidget(self.user_label)
        
        info_layout.addStretch()
        
        # 登录时间
        self.login_time_label = QLabel(f"登录时间: {self.login_time.strftime('%Y-%m-%d %H:%M:%S')}")
        info_layout.addWidget(self.login_time_label)
        
        info_layout.addStretch()
        
        # 单元格选择区域
        self.cell_selector = QComboBox()
        self.cell_selector.currentIndexChanged.connect(self.on_cell_changed)
        info_layout.addWidget(QLabel("选择单元格:"))
        info_layout.addWidget(self.cell_selector)
        
        left_layout.addLayout(info_layout)
        
        # 表格区域
        self.table_tabs = QTabWidget()
        self.table_tabs.setTabsClosable(False)
        left_layout.addWidget(self.table_tabs)
        
        left_widget.setLayout(left_layout)
        self.main_splitter.addWidget(left_widget)
        
        # 右侧搜索面板
        self.search_panel = SearchResultPanel()
        self.search_panel.setVisible(False)
        self.search_panel.setMinimumWidth(350)
        self.search_panel.setMaximumWidth(500)
        self.search_panel.close_btn.clicked.connect(self.hide_search_panel)
        self.search_panel.result_selected.connect(self.on_search_result_selected)
        self.main_splitter.addWidget(self.search_panel)
        
        # 设置分割比例
        self.main_splitter.setSizes([1050, 350])
        self.main_splitter.setStretchFactor(0, 3)
        self.main_splitter.setStretchFactor(1, 1)
        
        self.setCentralWidget(self.main_splitter)
        
        # 状态栏
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        
        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.status_bar.addPermanentWidget(self.progress_bar)
        
        # 状态标签
        self.status_label = QLabel("就绪")
        self.status_bar.addWidget(self.status_label)
        
        # 启动自动备份定时器
        self.auto_backup_timer = QTimer()
        self.auto_backup_timer.timeout.connect(self.auto_backup)
        self.auto_backup_timer.start(3600000)  # 每小时备份一次
        
    def create_menu_bar(self):
        """创建菜单栏"""
        menubar = self.menuBar()
        
        # 文件菜单
        file_menu = menubar.addMenu("文件")
        
        backup_action = QAction("备份管理", self)
        backup_action.triggered.connect(self.show_backup_dialog)
        file_menu.addAction(backup_action)
        
        file_menu.addSeparator()
        
        exit_action = QAction("退出", self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)
        
        # 用户菜单
        user_menu = menubar.addMenu("用户")
        
        switch_user_action = QAction("切换用户", self)
        switch_user_action.triggered.connect(self.switch_user)
        user_menu.addAction(switch_user_action)
        
        edit_config_action = QAction("编辑配置", self)
        edit_config_action.triggered.connect(self.edit_user_config)
        user_menu.addAction(edit_config_action)
        
        change_password_action = QAction("修改密码", self)
        change_password_action.triggered.connect(self.change_password)
        user_menu.addAction(change_password_action)
        
        # 单元格菜单
        cell_menu = menubar.addMenu("单元格")
        
        edit_cell_config_action = QAction("编辑当前单元格配置", self)
        edit_cell_config_action.triggered.connect(self.edit_current_cell_config)
        cell_menu.addAction(edit_cell_config_action)
        
        set_column_titles_action = QAction("设置当前单元格列标题", self)
        set_column_titles_action.triggered.connect(self.set_current_column_titles)
        cell_menu.addAction(set_column_titles_action)
        
        # 视图菜单
        view_menu = menubar.addMenu("视图")
        
        self.debug_action = QAction("调试模式", self)
        self.debug_action.setCheckable(True)
        self.debug_action.setChecked(self.debug_mode)
        self.debug_action.triggered.connect(self.toggle_debug_mode)
        view_menu.addAction(self.debug_action)
        
        # 帮助菜单
        help_menu = menubar.addMenu("帮助")
        
        about_action = QAction("关于", self)
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)
        
    def create_tool_bar(self):
        """创建工具栏"""
        toolbar = QToolBar()
        toolbar.setObjectName("mainToolBar")
        self.addToolBar(toolbar)
        
        # 备份按钮
        backup_btn = QPushButton("备份管理")
        backup_btn.clicked.connect(self.show_backup_dialog)
        toolbar.addWidget(backup_btn)
        
        # 刷新按钮
        refresh_btn = QPushButton("刷新")
        refresh_btn.clicked.connect(self.refresh_data)
        toolbar.addWidget(refresh_btn)
        
        # 设置列标题按钮
        set_titles_btn = QPushButton("设置列标题")
        set_titles_btn.clicked.connect(self.set_current_column_titles)
        toolbar.addWidget(set_titles_btn)
        
        # 搜索框
        toolbar.addSeparator()
        toolbar.addWidget(QLabel("搜索:"))
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("支持中文、英文...（按回车搜索）")
        self.search_edit.setMinimumWidth(200)
        self.search_edit.returnPressed.connect(self.on_search_pressed)
        toolbar.addWidget(self.search_edit)
        
        # 搜索按钮
        self.search_btn = QPushButton("搜索")
        self.search_btn.clicked.connect(self.on_search_pressed)
        toolbar.addWidget(self.search_btn)
        
        # 清除搜索按钮
        self.clear_search_btn = QPushButton("清除")
        self.clear_search_btn.clicked.connect(self.clear_search)
        self.clear_search_btn.setEnabled(False)
        toolbar.addWidget(self.clear_search_btn)
        
    def load_user_config(self):
        """加载用户配置（DB查询在后台，UI创建在主线程）"""
        self._config_worker = BackgroundWorker(self)
        self._config_worker.setup(self._do_load_user_config)
        self._config_worker.finished.connect(self._on_config_loaded)
        self._config_worker.error.connect(self._on_worker_error)
        self._config_worker.start()
    
    def _do_load_user_config(self):
        """在后台线程中查询用户配置"""
        user = self.db_manager.get_user_by_id(self.user_data["id"])
        config = user.get("config", {}) if user else {}
        if not config:
            config = {
                "cell_count": 6,
                "auto_resize_columns": True,
                "cell_configs": {}
            }
        return config
    
    def _on_config_loaded(self, config):
        """配置加载完成，在主线程创建表格UI"""
        self.current_user_config = config
        auto_resize = config.get("auto_resize_columns", True)
        
        # 初始化单元格选择器
        self.cell_selector.clear()
        total_cells = config.get("cell_count", 6)
        for i in range(total_cells):
            self.cell_selector.addItem(f"单元格 {i+1}")
            
        # 初始化表格标签页
        self.table_tabs.clear()
        for i in range(total_cells):
            tab = DataTableWidget()
            
            # 获取或创建单元格配置
            cell_configs = config.get("cell_configs", {})
            if str(i) in cell_configs:
                cell_config_dict = cell_configs[str(i)]
                cell_config = CellConfig(**cell_config_dict)
            else:
                # 使用默认配置
                default_cell = config.get("default_cell", {})
                cell_config = CellConfig(
                    row_count=default_cell.get("row_count", 3),
                    col_count=default_cell.get("col_count", 4),
                    require_login_time=default_cell.get("require_login_time", True),
                    calculate_days_diff=default_cell.get("calculate_days_diff", True),
                    require_title=default_cell.get("require_title", True),
                    title_mode=default_cell.get("title_mode", TitleMode.UNIFIED.value),
                    title_text=default_cell.get("title_text") or f"单元格 {i+1}",
                    column_titles=default_cell.get("column_titles", {}),
                    title_unique=default_cell.get("title_unique", True),
                    scheme_type=default_cell.get("scheme_type", CellScheme.STANDARD.value)
                )
                
            tab.configure(self.user_data["id"], i, cell_config, auto_resize)
            tab.data_changed.connect(self.on_data_changed)
            tab.paste_extended.connect(self.on_table_extended)
            self.table_tabs.addTab(tab, f"单元格 {i+1}")
        
        self._tables_ready = True
        # 表格创建完成后，加载数据
        self.load_user_data()
    
    def _startup_load(self):
        """启动时加载配置和数据（窗口已显示后再执行）"""
        self.status_label.setText("正在加载配置...")
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)
        self.load_user_config()  # DB查询后台 → 表格创建主线程 → 自动触发 load_user_data
            
    def load_user_data(self):
        """加载用户数据（后台线程）"""
        if not self.user_data:
            return
        
        self.status_label.setText("正在加载数据...")
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)  # 不确定进度
        
        self._data_worker = BackgroundWorker(self)
        self._data_worker.setup(self._do_load_data)
        self._data_worker.finished.connect(self._on_data_loaded)
        self._data_worker.error.connect(self._on_worker_error)
        self._data_worker.start()
    
    def _do_load_data(self):
        """在后台线程中加载数据"""
        return self.db_manager.get_data_records(self.user_data["id"])
    
    def _on_data_loaded(self, records):
        """数据加载完成（主线程回调）"""
        self.progress_bar.setVisible(False)
        
        # 分发数据到各个表格
        for i in range(self.table_tabs.count()):
            tab = self.table_tabs.widget(i)
            tab.load_data(records)
        
        self.status_label.setText(f"数据加载完成，共 {len(records)} 条记录")
            
    def on_data_changed(self, cell_index: int, row: int, col: int, content: str):
        """数据变化处理"""
        if not self.user_data:
            return
            
        if self.db_manager.save_data_record(
            self.user_data["id"], cell_index, row, col, content, self.login_time
        ):
            if self.debug_mode:
                self.status_label.setText(f"数据已保存 - 单元格{cell_index+1} ({row+1},{col+1})")
                
    def on_cell_changed(self, index: int):
        """单元格切换处理"""
        if index >= 0:
            self.table_tabs.setCurrentIndex(index)
            
    def on_search_pressed(self):
        """按回车或点击按钮触发搜索"""
        keyword = self.search_edit.text().strip()
        if keyword:
            self.perform_search(keyword)
        else:
            self.status_label.setText("请输入搜索关键词")
            
    def clear_search(self):
        """清除搜索"""
        self.search_edit.clear()
        self.clear_highlight_all_tables()
        self.hide_search_panel()
        self.clear_search_btn.setEnabled(False)
        self.status_label.setText("已清除搜索结果")
        self.search_results = []
        
    def hide_search_panel(self):
        """隐藏搜索面板"""
        self.search_panel.setVisible(False)
        # 调整分割器，让左侧占据大部分空间
        total_width = self.main_splitter.width()
        self.main_splitter.setSizes([total_width - 50, 50])
        
    def clear_highlight_all_tables(self):
        """清除所有表格的高亮"""
        for i in range(self.table_tabs.count()):
            tab = self.table_tabs.widget(i)
            if hasattr(tab, 'clear_highlight'):
                tab.clear_highlight()
                
    def on_search_result_selected(self, cell_index: int, row: int, col: int):
        """选择搜索结果"""
        # 切换到对应的单元格
        if cell_index < self.cell_selector.count():
            self.cell_selector.setCurrentIndex(cell_index)
            self.table_tabs.setCurrentIndex(cell_index)
            
            # 选中对应的单元格
            tab = self.table_tabs.widget(cell_index)
            if tab:
                tab.setCurrentCell(row, col)
                # 滚动到该单元格
                item = tab.item(row, col)
                if item:
                    tab.scrollToItem(item, QTableWidget.PositionAtCenter)
                tab.setFocus()
                # 添加一个短暂的闪烁效果来突出显示选中的单元格
                tab.editItem(item)
            
    def perform_search(self, keyword: str):
        """执行搜索（后台线程）"""
        if not self.user_data:
            return
            
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)
        self.status_label.setText(f"正在搜索: {keyword}...")
        
        # 清除旧的高亮
        self.clear_highlight_all_tables()
        
        # 取消之前的搜索任务
        if hasattr(self, '_search_worker') and self._search_worker.isRunning():
            self._search_worker.cancel()
        
        self._search_worker = BackgroundWorker(self)
        self._search_worker.setup(self.db_manager.search_data, self.user_data["id"], keyword)
        self._search_worker.finished.connect(lambda results: self._on_search_finished(results, keyword))
        self._search_worker.error.connect(self._on_worker_error)
        self._search_worker.start()
    
    def _on_search_finished(self, results, keyword):
        """搜索完成（主线程回调）"""
        self.search_results = results
        self.progress_bar.setVisible(False)
        
        if results:
            # 对所有表格应用高亮
            for i in range(self.table_tabs.count()):
                tab = self.table_tabs.widget(i)
                if hasattr(tab, 'set_highlight'):
                    tab.set_highlight(keyword)
            
            # 显示搜索面板
            self.search_panel.set_results(results, keyword)
            self.search_panel.setVisible(True)
            
            # 调整分割器
            total_width = self.main_splitter.width()
            panel_width = min(400, total_width // 3)
            self.main_splitter.setSizes([total_width - panel_width, panel_width])
            
            self.clear_search_btn.setEnabled(True)
            self.status_label.setText(f"搜索完成，找到 {len(results)} 条记录")
            
            # 自动跳转到第一个结果
            if results:
                first = results[0]
                self.on_search_result_selected(
                    first['cell_index'],
                    first['row_index'],
                    first['col_index']
                )
        else:
            self.hide_search_panel()
            QMessageBox.information(self, "搜索结果", f"未找到包含 \"{keyword}\" 的记录")
            self.status_label.setText(f"未找到包含 \"{keyword}\" 的记录")
            self.clear_search_btn.setEnabled(True)
     
    def refresh_data(self):
        """刷新数据"""
        self.load_user_data()
        
    def auto_backup(self):
        """自动备份（后台线程）"""
        if self.backup_manager:
            if hasattr(self, '_backup_worker') and self._backup_worker.isRunning():
                return  # 上一次备份还在进行中，跳过
            
            self._backup_worker = BackgroundWorker(self)
            self._backup_worker.setup(self.backup_manager.create_backup, "auto")
            self._backup_worker.finished.connect(self._on_auto_backup_finished)
            self._backup_worker.error.connect(self._on_worker_error)
            self._backup_worker.start()
    
    def _on_auto_backup_finished(self, result):
        """自动备份完成（主线程回调）"""
        success, backup_name, info = result
        if success and self.debug_mode:
            self.status_label.setText(f"自动备份完成: {backup_name}")
    
    def _on_worker_error(self, error_msg: str):
        """后台任务出错处理"""
        self.progress_bar.setVisible(False)
        print(f"[后台任务错误] {error_msg}")
        self.status_label.setText("后台任务执行出错，请查看日志")
                
    def show_backup_dialog(self):
        """显示备份恢复对话框"""
        if not self.backup_manager:
            return
        dialog = BackupRestoreDialog(self.backup_manager, self)
        dialog.backup_restored.connect(self.on_backup_restored)
        dialog.exec()
        
    def on_backup_restored(self):
        """备份恢复完成处理"""
        # 重新加载数据
        self.load_user_data()
        self.status_label.setText("数据库已恢复，数据已刷新")
        QMessageBox.information(self, "成功", "数据库恢复完成，界面数据已刷新")
        
    def switch_user(self):
        """切换用户"""
        reply = QMessageBox.question(
            self, "确认切换", "确定要切换用户吗？未保存的数据将会丢失。",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            self.close()
            
    def edit_user_config(self):
        """编辑用户配置"""
        dialog = UserConfigDialog(self.db_manager, self.user_data["id"], self)
        if dialog.exec():
            config = dialog.get_config()
            if self.db_manager.update_user_config(self.user_data["id"], config):
                self.current_user_config = config  # 先更新内存中的配置
                self.load_user_config()  # 异步重新加载（完成后自动触发 load_user_data）
                QMessageBox.information(self, "成功", "配置已更新，正在重新加载...")
            else:
                QMessageBox.warning(self, "错误", "保存配置失败")
                
    def edit_current_cell_config(self):
        """编辑当前单元格配置"""
        current_tab_index = self.table_tabs.currentIndex()
        if current_tab_index < 0:
            return
            
        tab = self.table_tabs.widget(current_tab_index)
        if not hasattr(tab, 'cell_config'):
            return
            
        dialog = CellConfigDialog(tab.cell_config, current_tab_index, self)
        if dialog.exec():
            new_config = dialog.get_config()
            tab.cell_config = new_config
            
            # 重新配置表格
            tab.configure(
                tab.user_id, tab.cell_index, new_config, tab.auto_resize_columns
            )
            tab._update_headers()
            
            # 保存配置到数据库
            cell_configs = self.current_user_config.get("cell_configs", {})
            cell_configs[str(current_tab_index)] = asdict(new_config)
            self.current_user_config["cell_configs"] = cell_configs
            
            # 询问是否更新为默认配置
            update_default = QMessageBox.question(
                self, 
                "更新默认配置", 
                "是否将此单元格配置设为新建单元格的默认配置？\n"
                "（选择'是'，则后续新建单元格将使用此配置；选择'否'，则仅修改当前单元格）",
                QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel
            )
            
            if update_default == QMessageBox.Yes:
                # 更新默认配置
                self.current_user_config["default_cell"] = asdict(new_config)
                self.status_label.setText("已更新为默认配置")
            elif update_default == QMessageBox.Cancel:
                return
                
            # 保存用户配置
            self.db_manager.update_user_config(self.user_data["id"], self.current_user_config)
            
            # 重新加载数据
            self.load_user_data()
            QMessageBox.information(self, "成功", "单元格配置已更新")
            
    def set_current_column_titles(self):
        """设置当前单元格的列标题"""
        current_tab_index = self.table_tabs.currentIndex()
        if current_tab_index < 0:
            return
            
        tab = self.table_tabs.widget(current_tab_index)
        if not hasattr(tab, 'cell_config') or not tab.cell_config.require_title:
            QMessageBox.warning(self, "提示", "当前单元格未启用标题功能")
            return
            
        dialog = ColumnTitleDialog(tab, self)
        if dialog.exec():
            # 保存配置
            cell_configs = self.current_user_config.get("cell_configs", {})
            cell_configs[str(current_tab_index)] = asdict(tab.cell_config)
            self.current_user_config["cell_configs"] = cell_configs
            self.db_manager.update_user_config(self.user_data["id"], self.current_user_config)
            QMessageBox.information(self, "成功", "列标题已更新")
                
    def change_password(self):
        """修改当前用户密码"""
        dialog = ChangePasswordDialog(self.db_manager, self.user_data["id"], self)
        if dialog.exec():
            QMessageBox.information(self, "成功", "密码修改成功，请重新登录")
            self.switch_user()
                
    def toggle_debug_mode(self, checked: bool):
        """切换调试模式"""
        self.debug_mode = checked
        self.settings.setValue("debug_mode", checked)
        self.status_label.setText(f"调试模式已{'开启' if checked else '关闭'}")
        
        if self.debug_mode:
            self.status_label.setStyleSheet("color: red;")
        else:
            self.status_label.setStyleSheet("")
            
    def show_about(self):
        """显示关于对话框"""
        about_html = ProjectInfo.get_about_text()
        about_html += """
        <h3>功能特性:</h3>
        <ul>
            <li>多用户管理</li>
            <li>支持空密码登录</li>
            <li>用户下拉选择，记住上次登录</li>
            <li>灵活的表格配置（每个单元格的行列数可自定义）</li>
            <li>灵活的单元格数量设置</li>
            <li>支持统一标题和独立标题两种模式</li>
            <li>自动调整列宽（根据内容）</li>
            <li>数据库自动备份</li>
            <li>数据搜索功能（支持高亮显示）</li>
            <li>支持多选复制、剪切、粘贴（Excel风格）</li>
            <li>支持撤销/重做操作（Ctrl+Z/Ctrl+Y）</li>
            <li>支持WAL模式</li>
            <li>多线程后台处理（数据加载/搜索/备份不阻塞界面）</li>
        </ul>
        <h3>默认管理员:</h3>
        <p>账号: admin<br>密码: admin123</p>
        """
        
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("关于")
        msg_box.setIcon(QMessageBox.Information)
        msg_box.setTextFormat(Qt.RichText)
        msg_box.setText(about_html)
        msg_box.exec()

    def on_table_extended(self, cell_index: int, new_rows: int, new_cols: int):
        """表格被扩展时更新配置"""
        if cell_index >= 0:
            cell_configs = self.current_user_config.get("cell_configs", {})
            if str(cell_index) in cell_configs:
                cell_configs[str(cell_index)]["row_count"] = new_rows
                cell_configs[str(cell_index)]["col_count"] = new_cols
            else:
                # 更新默认配置
                default_cell = self.current_user_config.get("default_cell", {})
                default_cell["row_count"] = new_rows
                default_cell["col_count"] = new_cols
                self.current_user_config["default_cell"] = default_cell
            
            self.current_user_config["cell_configs"] = cell_configs
            # 自动保存配置
            self.db_manager.update_user_config(self.user_data["id"], self.current_user_config)



        
    def closeEvent(self, event):
        """关闭事件"""
        # 保存窗口位置和大小
        self._save_window_geometry()
        # 停止自动备份定时器
        if hasattr(self, 'auto_backup_timer'):
            self.auto_backup_timer.stop()
        # 取消所有运行中的后台任务
        for worker_attr in ['_data_worker', '_search_worker', '_backup_worker']:
            if hasattr(self, worker_attr):
                worker = getattr(self, worker_attr)
                if worker.isRunning():
                    worker.cancel()
        event.accept()

    def _save_window_geometry(self):
        """保存窗口位置和大小到设置文件"""
        try:
            geometry = self.saveGeometry()
            state = self.saveState()
            self.settings.set("window_geometry", geometry.toHex().data().decode())
            self.settings.set("window_state", state.toHex().data().decode())
        except Exception as e:
            print(f"[窗口] 保存窗口位置失败: {e}")

    def _restore_window_geometry(self):
        """从设置文件恢复窗口位置和大小"""
        try:
            geometry_hex = self.settings.get("window_geometry")
            state_hex = self.settings.get("window_state")
            if geometry_hex:
                from PySide6.QtCore import QByteArray
                self.restoreGeometry(QByteArray.fromHex(geometry_hex.encode()))
            if state_hex:
                from PySide6.QtCore import QByteArray
                self.restoreState(QByteArray.fromHex(state_hex.encode()))
        except Exception as e:
            print(f"[窗口] 恢复窗口位置失败: {e}")




def get_app_icon() -> QIcon:
    """获取应用程序图标"""
    # 尝试加载外部图标文件
    icon_path = get_icon_path()
    if icon_path and icon_path.exists():
        return QIcon(str(icon_path))
    else:
        # 使用默认生成的图标
        return create_default_icon()


class MainApplication(QApplication):
    """主应用程序类"""
    
    def __init__(self, argv):
        super().__init__(argv)
        self.db_manager = None
        self.main_window = None
        
        # 设置应用程序图标
        app_icon = get_app_icon()
        self.setWindowIcon(app_icon)
        
    def run(self):
        """运行应用"""
        # 使用固定路径的数据库文件
        db_path = str(DB_FILE)
        self.db_manager = DatabaseManager(db_path)
        
        # 显示登录对话框
        login_dialog = LoginDialog(self.db_manager)
        
        def on_login_success(user_data):
            login_dialog.accept()
            self.main_window = DataEntryWindow(self.db_manager, user_data)
            self.main_window.show()
            
        login_dialog.login_success.connect(on_login_success)
        
        if login_dialog.exec() != QDialog.Accepted:
            # 用户取消了登录
            return False
            
        return True



class SearchResultPanel(QWidget):
    """搜索结果面板 - 使用 QTextDocument + QTextCursor 方案"""
    
    result_selected = Signal(int, int, int)  # cell_index, row, col
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_ui()
        self.results = []
        self.keyword = ""
        self.current_index = -1
        self.result_positions = []  # 存储每个结果在文档中的位置，用于导航
        
    def setup_ui(self):
        """设置UI"""
        layout = QVBoxLayout()
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(8)
        
        # 标题栏
        title_layout = QHBoxLayout()
        self.title_label = QLabel("🔍 搜索结果")
        title_font = QFont()
        title_font.setBold(True)
        title_font.setPointSize(12)
        self.title_label.setFont(title_font)
        title_layout.addWidget(self.title_label)
        
        title_layout.addStretch()
        
        self.close_btn = QPushButton("×")
        self.close_btn.setFixedSize(28, 28)
        self.close_btn.setCursor(Qt.PointingHandCursor)
        self.close_btn.setStyleSheet("""
            QPushButton {
                border: none;
                border-radius: 14px;
                background-color: #f0f0f0;
                font-size: 18px;
                font-weight: bold;
                color: #666;
            }
            QPushButton:hover {
                background-color: #ff4444;
                color: white;
            }
        """)
        self.close_btn.setToolTip("关闭搜索面板")
        title_layout.addWidget(self.close_btn)
        
        layout.addLayout(title_layout)
        
        # 分隔线
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        line.setStyleSheet("background-color: #ddd;")
        layout.addWidget(line)
        
        # 结果统计
        self.stats_label = QLabel()
        self.stats_label.setStyleSheet("""
            color: #666; 
            padding: 5px; 
            background-color: #f8f9fa;
            border-radius: 4px;
        """)
        layout.addWidget(self.stats_label)
        
        # 使用 QTextEdit 显示搜索结果（支持富文本）
        self.result_text = QTextEdit()
        self.result_text.setReadOnly(True)
        self.result_text.setStyleSheet("""
            QTextEdit {
                border: 1px solid #ddd;
                border-radius: 4px;
                background-color: white;
                font-family: 'Microsoft YaHei', sans-serif;
                font-size: 10pt;
                line-height: 1.5;
            }
        """)
        self.result_text.setCursorWidth(0)  # 隐藏光标
        layout.addWidget(self.result_text, 1)
        
        # 导航按钮
        nav_layout = QHBoxLayout()
        nav_layout.setSpacing(10)
        
        self.prev_btn = QPushButton("◀ 上一个")
        self.prev_btn.clicked.connect(self.go_to_previous)
        self.prev_btn.setEnabled(False)
        self.prev_btn.setCursor(Qt.PointingHandCursor)
        self.prev_btn.setStyleSheet("""
            QPushButton {
                padding: 8px 16px;
                border: 1px solid #ddd;
                border-radius: 4px;
                background-color: #f8f9fa;
            }
            QPushButton:hover {
                background-color: #e9ecef;
            }
            QPushButton:disabled {
                color: #ccc;
                border-color: #eee;
            }
        """)
        nav_layout.addWidget(self.prev_btn)
        
        self.next_btn = QPushButton("下一个 ▶")
        self.next_btn.clicked.connect(self.go_to_next)
        self.next_btn.setEnabled(False)
        self.next_btn.setCursor(Qt.PointingHandCursor)
        self.next_btn.setStyleSheet("""
            QPushButton {
                padding: 8px 16px;
                border: 1px solid #ddd;
                border-radius: 4px;
                background-color: #f8f9fa;
            }
            QPushButton:hover {
                background-color: #e9ecef;
            }
            QPushButton:disabled {
                color: #ccc;
                border-color: #eee;
            }
        """)
        nav_layout.addWidget(self.next_btn)
        
        nav_layout.addStretch()
        
        # 当前位置指示
        self.position_label = QLabel("")
        self.position_label.setStyleSheet("color: #666; font-size: 11px;")
        nav_layout.addWidget(self.position_label)
        
        layout.addLayout(nav_layout)
        
        self.setLayout(layout)
        
    def set_results(self, results: List[Dict], keyword: str):
        """设置搜索结果 - 使用 QTextDocument 构建富文本"""
        self.results = results
        self.keyword = keyword
        self.current_index = -1
        self.result_positions = []
        
        if not results:
            self.title_label.setText("🔍 搜索结果 (0)")
            self.stats_label.setText("❌ 未找到匹配的记录")
            self.result_text.clear()
            self.prev_btn.setEnabled(False)
            self.next_btn.setEnabled(False)
            self.position_label.setText("")
            return
            
        self.title_label.setText(f"🔍 搜索结果 ({len(results)})")
        self.stats_label.setText(f"✅ 找到 {len(results)} 条匹配记录，关键词: \"{keyword}\"")
        
        # 创建 QTextDocument
        doc = QTextDocument()
        cursor = QTextCursor(doc)
        
        # 定义格式
        # 位置标签格式
        location_format = QTextCharFormat()
        location_format.setForeground(QColor("#1976d2"))
        location_font = QFont("Microsoft YaHei", 10)
        location_font.setBold(True)
        location_format.setFont(location_font)
        
        # 分隔线格式
        separator_format = QTextCharFormat()
        separator_format.setForeground(QColor("#cccccc"))
        
        # 内容格式（普通文本）
        normal_format = QTextCharFormat()
        normal_format.setForeground(QColor("#333333"))
        normal_font = QFont("Microsoft YaHei", 9)
        normal_format.setFont(normal_font)
        
        # 高亮格式
        highlight_format = QTextCharFormat()
        highlight_format.setBackground(QBrush(QColor("#ffeb3b")))
        highlight_format.setForeground(QColor("#000000"))
        highlight_font = QFont("Microsoft YaHei", 9)
        highlight_font.setBold(True)
        highlight_format.setFont(highlight_font)
        
        # 块格式（用于控制段落间距）
        block_format = QTextBlockFormat()
        block_format.setBottomMargin(5)
        
        keyword_lower = keyword.lower()
        
        for i, result in enumerate(results):
            # 记录当前光标位置（用于导航）
            self.result_positions.append(cursor.position())
            
            cell_idx = result['cell_index'] + 1
            row = result['row_index'] + 1
            col = result['col_index'] + 1
            content = result['content'] or ""
            
            # 添加位置标签
            cursor.insertText(f"📌 单元格 {cell_idx} - 第 {row} 行, 第 {col} 列\n", location_format)
            
            # 添加内容（带高亮）
            if keyword_lower in content.lower():
                self.insert_highlighted_text(cursor, content, keyword, 
                                             normal_format, highlight_format)
            else:
                cursor.insertText(content, normal_format)
            
            cursor.insertBlock()
            
            # 添加分隔线（最后一项不加）
            if i < len(results) - 1:
                cursor.insertText("─" * 50 + "\n\n", separator_format)
        
        # 记录最后一个结果的结束位置
        self.result_positions.append(cursor.position())
        
        # 设置文档
        self.result_text.setDocument(doc)
        
        self.prev_btn.setEnabled(True)
        self.next_btn.setEnabled(True)
        self.update_position_label()
        
        # 默认选中第一个结果
        if results:
            self.navigate_to_result(0)
        
    def insert_highlighted_text(self, cursor: QTextCursor, text: str, keyword: str,
                                 normal_format: QTextCharFormat, 
                                 highlight_format: QTextCharFormat):
        """插入带高亮的文本"""
        keyword_lower = keyword.lower()
        text_lower = text.lower()
        
        pos = 0
        while pos < len(text):
            # 查找关键词位置
            found_pos = text_lower.find(keyword_lower, pos)
            
            if found_pos == -1:
                # 没有更多关键词，插入剩余文本
                cursor.insertText(text[pos:], normal_format)
                break
            
            # 插入关键词前的普通文本
            if found_pos > pos:
                cursor.insertText(text[pos:found_pos], normal_format)
            
            # 插入高亮的关键词（保持原始大小写）
            keyword_original = text[found_pos:found_pos + len(keyword)]
            cursor.insertText(keyword_original, highlight_format)
            
            pos = found_pos + len(keyword)
    
    def navigate_to_result(self, index: int):
        """导航到指定结果"""
        if 0 <= index < len(self.results):
            result = self.results[index]
            self.current_index = index
            self.update_position_label()
            
            # 滚动到对应的结果位置
            if index < len(self.result_positions):
                cursor = self.result_text.textCursor()
                cursor.setPosition(self.result_positions[index])
                self.result_text.setTextCursor(cursor)
                self.result_text.ensureCursorVisible()
            
            # 发出信号，在主表格中选中对应的单元格
            self.result_selected.emit(
                result['cell_index'],
                result['row_index'],
                result['col_index']
            )
    
    def update_position_label(self):
        """更新位置标签"""
        if self.results:
            self.position_label.setText(f"{self.current_index + 1} / {len(self.results)}")
        else:
            self.position_label.setText("")
    
    def go_to_previous(self):
        """跳转到上一个结果"""
        if self.current_index > 0:
            self.navigate_to_result(self.current_index - 1)
    
    def go_to_next(self):
        """跳转到下一个结果"""
        if self.current_index < len(self.results) - 1:
            self.navigate_to_result(self.current_index + 1)

def main():
    """主函数"""
    app = MainApplication(sys.argv)
    
    if app.run():
        sys.exit(app.exec())
    else:
        sys.exit(0)


if __name__ == "__main__":
    main()