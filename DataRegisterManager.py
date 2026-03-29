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
    QListWidgetItem, QAbstractItemView, QStyleFactory, QApplication as QApp
)
from PySide6.QtCore import (
    Qt, Signal, QThread, QObject, QTimer, QDateTime, QDate, 
    QSize, QSettings, QCoreApplication, QEvent, QMutex, QWaitCondition
)
from PySide6.QtGui import (
    QIcon, QFont, QColor, QPalette, QAction, QActionGroup, QKeySequence,
    QStandardItemModel, QStandardItem, QPixmap, QDesktopServices
)


def get_app_dir() -> Path:
    """获取应用程序所在目录"""
    if getattr(sys, 'frozen', False):
        # 打包后的exe
        app_dir = Path(sys.executable).parent
        print(f"[路径信息] 打包模式 - 应用程序目录: {app_dir}")
    else:
        # 开发环境
        app_dir = Path(__file__).parent
        print(f"[路径信息] 开发模式 - 脚本所在目录: {app_dir}")
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
            print(f"[图标] 找到图标文件: {icon_path}")
            return icon_path
            
        # 在脚本所在目录查找
        icon_path = Path(__file__).parent / icon_name
        if icon_path.exists():
            print(f"[图标] 找到图标文件: {icon_path}")
            return icon_path
            
    print(f"[图标] 未找到外部图标文件，使用默认图标")
    return None


# 应用程序目录
APP_DIR = get_app_dir()
print(f"[路径信息] 最终应用程序目录: {APP_DIR}")

# 配置常量
APP_NAME = "资料登记管理系统"
APP_VERSION = "1.0.0"
BACKUP_DIR = APP_DIR / "backups"
CONFIG_FILE = APP_DIR / "app_config.json"
DB_FILE = APP_DIR / f"{APP_NAME.replace(' ', '_')}{'.db'}"

print(f"\n[路径信息] 主要路径配置:")
print(f"  - 应用程序目录: {APP_DIR}")
print(f"  - 数据库文件路径: {DB_FILE}")
print(f"  - 备份目录: {BACKUP_DIR}")
print(f"  - 配置文件路径: {CONFIG_FILE}")

# 检查路径是否存在或创建
print(f"\n[路径检查] 目录状态:")
print(f"  - 应用程序目录存在: {APP_DIR.exists()}")
if not BACKUP_DIR.exists():
    print(f"  - 备份目录不存在，将在首次备份时创建")
else:
    print(f"  - 备份目录存在: {BACKUP_DIR.exists()}")

WAL_MODE = "WAL"  # WAL模式


# 确保备份目录存在
try:
    BACKUP_DIR.mkdir(exist_ok=True)
    print(f"  - 备份目录已创建/确认: {BACKUP_DIR}")
except Exception as e:
    print(f"  - 备份目录创建失败: {e}")


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
        print(f"[备份管理器] 初始化 - 数据库路径: {db_path}")
        print(f"[备份管理器] 备份目录: {self.backup_dir}")
        
    def create_backup(self, backup_type: str = "manual") -> Tuple[bool, str, Dict]:
        """创建数据库备份"""
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_name = f"backup_{timestamp}_{backup_type}.db"
            backup_path = self.backup_dir / backup_name
            
            # 获取备份前信息
            source_size = os.path.getsize(self.db_path)
            print(f"[备份] 创建 {backup_type} 备份: {backup_path}")
            print(f"[备份] 源数据库大小: {source_size} 字节")
            
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
            
            print(f"[备份] 备份成功，大小: {backup_size} 字节")

            # 清理旧备份
            self._cleanup_old_backups()
            
            return True, backup_name, backup_info
            
        except Exception as e:
            print(f"[备份] 备份失败: {e}")
            return False, str(e), {}
            
    def _cleanup_old_backups(self):
        """清理旧备份"""
        backups = sorted(self.backup_dir.glob("*.db"), key=os.path.getmtime)
        print(f"[备份] 当前备份数量: {len(backups)}, 最大保留: {self.max_backups}")
        while len(backups) > self.max_backups:
            print(f"[备份] 删除旧备份: {backups[0]}")
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
            print(f"[恢复] 创建回滚备份: {rollback_path}")
            shutil.copy2(self.db_path, rollback_path)
            
            # 恢复备份
            print(f"[恢复] 从备份恢复: {backup_path} -> {self.db_path}")
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
        print(f"[数据库管理器] 初始化 - 数据库路径: {db_path}")
        print(f"[数据库管理器] 数据库文件是否存在: {os.path.exists(db_path)}")
        self._init_database()
        
    def _init_database(self):
        """初始化数据库"""
        try:
            print(f"[数据库] 开始初始化数据库...")
            # 启用WAL模式
            self.connection = sqlite3.connect(self.db_path, check_same_thread=False)
            print(f"[数据库] 数据库连接已建立")
            
            self.connection.execute("PRAGMA journal_mode=WAL")
            wal_status = self.connection.execute("PRAGMA journal_mode").fetchone()[0]
            print(f"[数据库] WAL模式状态: {wal_status}")
            
            self.connection.execute("PRAGMA synchronous=NORMAL")
            self.cursor = self.connection.cursor()
            
            # 创建基础表
            print(f"[数据库] 创建数据表...")
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
            print(f"[数据库] 数据表创建/确认完成")
            
            # 创建默认管理员用户（如果不存在）
            self._create_default_admin()
            
            # 显示数据库文件信息
            db_size = os.path.getsize(self.db_path) if os.path.exists(self.db_path) else 0
            print(f"[数据库] 数据库文件大小: {db_size} 字节 ({db_size/1024:.2f} KB)")
            
        except Exception as e:
            print(f"[数据库] 初始化错误: {e}")
            import traceback
            traceback.print_exc()
            
    def _create_default_admin(self):
        """创建默认管理员用户"""
        print(f"[数据库] 检查默认管理员用户...")
        admin_exists = self.execute_query("SELECT id FROM users WHERE username = ?", ("admin",))
        if not admin_exists:
            print(f"[数据库] 创建默认管理员用户...")
            # 默认密码: admin123
            default_password = hashlib.sha256("admin123".encode()).hexdigest()
            self.execute_update(
                "INSERT INTO users (username, password, config) VALUES (?, ?, ?)",
                ("admin", default_password, json.dumps({"cell_count": 6, "default_cell": {}, "auto_resize_columns": True}))
            )
            print(f"[数据库] 默认管理员用户创建成功 (用户名: admin, 密码: admin123)")
        else:
            print(f"[数据库] 默认管理员用户已存在")
            
    def execute_query(self, query: str, params: tuple = ()) -> List[tuple]:
        """执行查询"""
        try:
            self.cursor.execute(query, params)
            return self.cursor.fetchall()
        except Exception as e:
            print(f"[数据库] 查询错误: {e}")
            return []
            
    def execute_update(self, query: str, params: tuple = ()) -> bool:
        """执行更新"""
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
        print(f"[登录] 验证用户: {username}, 密码{'为空' if password == '' else '不为空'}")
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
            print(f"[登录] 用户 {username} 验证成功")
            return {
                "id": result[0][0],
                "username": result[0][1],
                "created_at": result[0][2],
                "config": json.loads(result[0][3]) if result[0][3] else {}
            }
        else:
            print(f"[登录] 用户 {username} 验证失败")
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
        print(f"[用户管理] 创建用户: {username}, 密码{'为空' if password == '' else '不为空'}")
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
        print(f"[用户管理] 删除用户 ID: {user_id}")
        # 删除用户相关数据
        self.execute_update("DELETE FROM user_configs WHERE user_id = ?", (user_id,))
        self.execute_update("DELETE FROM data_records WHERE user_id = ?", (user_id,))
        self.execute_update("DELETE FROM cell_templates WHERE user_id = ?", (user_id,))
        return self.execute_update("DELETE FROM users WHERE id = ?", (user_id,))
        
    def update_username(self, user_id: int, new_username: str) -> bool:
        """更新用户名"""
        print(f"[用户管理] 更新用户名 ID: {user_id} -> {new_username}")
        return self.execute_update(
            "UPDATE users SET username = ? WHERE id = ?",
            (new_username, user_id)
        )
        
    def update_user_password(self, user_id: int, new_password: str) -> bool:
        """更新用户密码（支持空密码）"""
        print(f"[用户管理] 更新密码 ID: {user_id}, 新密码{'为空' if new_password == '' else '不为空'}")
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
        print(f"[用户管理] 更新配置 ID: {user_id}")
        return self.execute_update(
            "UPDATE users SET config = ? WHERE id = ?",
            (json.dumps(config), user_id)
        )
        
    def get_all_users(self) -> List[Dict]:
        """获取所有用户（不包含密码）"""
        results = self.execute_query("SELECT id, username, created_at FROM users ORDER BY id")
        users = [{"id": r[0], "username": r[1], "created_at": r[2]} for r in results]
        print(f"[用户管理] 获取用户列表，共 {len(users)} 个用户")
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
            print(f"[数据库] 关闭数据库连接")
            self.connection.close()


class DataTableWidget(QTableWidget):
    """数据表格组件类"""
    
    data_changed = Signal(int, int, int, str)  # cell_index, row, col, content
    resize_requested = Signal()  # 请求调整列宽
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.cell_config = None
        self.user_id = None
        self.cell_index = None
        self.auto_resize_columns = True
        self.setup_ui()
        
    def setup_ui(self):
        """设置UI"""
        self.setAlternatingRowColors(True)
        self.setSelectionBehavior(QTableWidget.SelectItems)
        self.setSelectionMode(QTableWidget.SingleSelection)
        self.itemChanged.connect(self.on_item_changed)
        
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
                    
    def on_item_changed(self, item: QTableWidgetItem):
        """数据变化处理"""
        if item and self.user_id and self.cell_index is not None:
            row = item.row()
            col = item.column()
            content = item.text()
            self.data_changed.emit(self.cell_index, row, col, content)
            
            # 如果启用自动调整列宽，请求调整
            if self.auto_resize_columns:
                self.resizeColumnToContents(col)
            
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
            
    def load_config(self):
        """加载现有配置"""
        # 从数据库加载用户配置
        user = self.db_manager.get_user_by_id(self.user_id)
        if user and user.get("config"):
            config = user["config"]
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
            
    def get_config(self) -> Dict:
        """获取配置"""
        return {
            "cell_count": self.cell_count_spin.value(),
            "auto_resize_columns": self.auto_resize_cb.isChecked(),
            "default_cell": {
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
            },
            "cell_configs": {}
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
        self.settings = QSettings(APP_NAME, APP_NAME)
        self.setup_ui()
        self.load_last_user()
        
    def setup_ui(self):
        """设置UI"""
        self.setWindowTitle("用户登录")
        self.setFixedSize(450, 380)
        
        # 设置窗口图标
        self.setWindowIcon(get_app_icon())
        
        layout = QVBoxLayout()
        
        # 标题
        title_label = QLabel(APP_NAME)
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
        """刷新备份列表"""
        self.backup_tree.clear()
        backups = self.backup_manager.get_backup_list()
        
        # 应用筛选
        backup_type = self.type_combo.currentText()
        start_datetime = self.start_date.dateTime()
        end_datetime = self.end_date.dateTime()
        
        for backup in backups:
            # 类型筛选
            if backup_type != "全部" and backup["type"] != backup_type:
                continue
                
            # 时间筛选
            if backup["datetime"]:
                if backup["datetime"] < start_datetime or backup["datetime"] > end_datetime:
                    continue
                    
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
        """执行备份恢复"""
        reply = QMessageBox.question(
            self, "确认恢复", 
            "恢复操作将替换当前数据库，并创建回滚备份。\n确定要继续吗？",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            self.progress_bar.setVisible(True)
            self.progress_bar.setRange(0, 0)  # 不确定进度
            
            success, result = self.backup_manager.restore_backup(backup_path)
            
            self.progress_bar.setVisible(False)
            
            if success:
                QMessageBox.information(self, "成功", f"数据库恢复成功！\n回滚备份已创建: {result}")
                self.backup_restored.emit()
                self.refresh_backup_list()
            else:
                QMessageBox.critical(self, "错误", f"数据库恢复失败: {result}")
                
    def on_manual_backup(self):
        """手动备份"""
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)
        
        success, result, info = self.backup_manager.create_backup("manual")
        
        self.progress_bar.setVisible(False)
        
        if success:
            QMessageBox.information(self, "成功", f"手动备份成功: {result}")
            self.refresh_backup_list()
        else:
            QMessageBox.critical(self, "错误", f"备份失败: {result}")


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
        self.init_settings()
        self.setup_ui()
        
    def init_settings(self):
        """初始化设置"""
        self.settings = QSettings(APP_NAME, APP_NAME)
        self.debug_mode = self.settings.value("debug_mode", False, type=bool)
        
    def setup_ui(self):
        """设置UI"""
        self.setWindowTitle(f"{APP_NAME} - 当前用户: {self.user_data['username']}")
        self.resize(1200, 800)
        
        # 设置窗口图标
        self.setWindowIcon(get_app_icon())
        
        # 创建菜单栏
        self.create_menu_bar()
        
        # 创建工具栏
        self.create_tool_bar()
        
        # 中央部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout()
        
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
        
        main_layout.addLayout(info_layout)
        
        # 表格区域
        self.table_tabs = QTabWidget()
        self.table_tabs.setTabsClosable(False)
        main_layout.addWidget(self.table_tabs)
        
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
        
        central_widget.setLayout(main_layout)
        
        # 启动自动备份定时器
        self.auto_backup_timer = QTimer()
        self.auto_backup_timer.timeout.connect(self.auto_backup)
        self.auto_backup_timer.start(3600000)  # 每小时备份一次
        
        # 加载用户配置和数据
        self.load_user_config()
        self.load_user_data()
        
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
        self.search_edit.setPlaceholderText("支持中文、英文...")
        self.search_edit.textChanged.connect(self.on_search)
        toolbar.addWidget(self.search_edit)
        
    def load_user_config(self):
        """加载用户配置"""
        # 从数据库获取最新配置
        user = self.db_manager.get_user_by_id(self.user_data["id"])
        if user:
            self.current_user_config = user.get("config", {})
            
        # 如果没有配置，创建默认配置
        if not self.current_user_config:
            self.current_user_config = {
                "cell_count": 6,
                "auto_resize_columns": True,
                "cell_configs": {}
            }
            
        # 获取自动调整列宽设置
        auto_resize = self.current_user_config.get("auto_resize_columns", True)
            
        # 初始化单元格选择器
        self.cell_selector.clear()
        total_cells = self.current_user_config.get("cell_count", 6)
        for i in range(total_cells):
            self.cell_selector.addItem(f"单元格 {i+1}")
            
        # 初始化表格标签页
        self.table_tabs.clear()
        for i in range(total_cells):
            tab = DataTableWidget()
            
            # 获取或创建单元格配置
            cell_config = None
            cell_configs = self.current_user_config.get("cell_configs", {})
            if str(i) in cell_configs:
                cell_config_dict = cell_configs[str(i)]
                cell_config = CellConfig(**cell_config_dict)
            else:
                # 使用默认配置
                default_cell = self.current_user_config.get("default_cell", {})
                cell_config = CellConfig(
                    row_count=default_cell.get("row_count", 3),
                    col_count=default_cell.get("col_count", 4),
                    require_login_time=default_cell.get("require_login_time", True),
                    calculate_days_diff=default_cell.get("calculate_days_diff", True),
                    require_title=default_cell.get("require_title", True),
                    title_mode=default_cell.get("title_mode", TitleMode.UNIFIED.value),
                    title_text=default_cell.get("title_text", f"单元格 {i+1}"),
                    column_titles=default_cell.get("column_titles", {}),
                    title_unique=default_cell.get("title_unique", True),
                    scheme_type=default_cell.get("scheme_type", CellScheme.STANDARD.value)
                )
                
            tab.configure(self.user_data["id"], i, cell_config, auto_resize)
            tab.data_changed.connect(self.on_data_changed)
            self.table_tabs.addTab(tab, f"单元格 {i+1}")
            
    def load_user_data(self):
        """加载用户数据"""
        if not self.user_data:
            return
            
        records = self.db_manager.get_data_records(self.user_data["id"])
        
        # 分发数据到各个表格
        for i in range(self.table_tabs.count()):
            tab = self.table_tabs.widget(i)
            tab.load_data(records)
            
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
            
    def on_search(self, keyword: str):
        """搜索处理"""
        if not keyword or not self.user_data:
            return
            
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)
            
        results = self.db_manager.search_data(self.user_data["id"], keyword)
        
        self.progress_bar.setVisible(False)
        
        if results:
            result_text = f"找到 {len(results)} 条记录:\n\n"
            for r in results:
                result_text += f"单元格 {r['cell_index']+1}, 位置({r['row_index']+1},{r['col_index']+1}): {r['content']}\n"
            QMessageBox.information(self, "搜索结果", result_text)
            self.status_label.setText(f"搜索完成，找到 {len(results)} 条记录")
        else:
            QMessageBox.information(self, "搜索结果", "未找到匹配的记录")
            self.status_label.setText("未找到匹配的记录")
            
    def refresh_data(self):
        """刷新数据"""
        self.load_user_data()
        self.status_label.setText("数据已刷新")
        
    def auto_backup(self):
        """自动备份"""
        if self.backup_manager:
            success, result, info = self.backup_manager.create_backup("auto")
            if success and self.debug_mode:
                self.status_label.setText(f"自动备份完成: {result}")
                
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
                self.load_user_config()
                self.load_user_data()
                QMessageBox.information(self, "成功", "配置已更新")
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
        QMessageBox.about(
            self, "关于",
            f"{APP_NAME} v{APP_VERSION}\n\n"
            "资料登记管理系统\n"
            "功能特性:\n"
            "- 多用户管理\n"
            "- 支持空密码登录\n"
            "- 用户下拉选择，记住上次登录\n"
            "- 灵活的表格配置（每个单元格的行列数可自定义）\n"
            "- 灵活的单元格数量设置\n"
            "- 支持统一标题和独立标题两种模式\n"
            "- 自动调整列宽（根据内容）\n"
            "- 数据库自动备份\n"
            "- 数据搜索功能\n"
            "- 支持WAL模式\n\n"
            "默认管理员账号: admin\n"
            "默认管理员密码: admin123\n\n"
            "© 2024 版权所有"
        )
        
    def closeEvent(self, event):
        """关闭事件"""
        # 停止自动备份定时器
        if hasattr(self, 'auto_backup_timer'):
            self.auto_backup_timer.stop()
        event.accept()


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


def main():
    """主函数"""
    print(f"\n{'='*60}")
    print(f"启动 {APP_NAME} v{APP_VERSION}")
    print(f"{'='*60}")
    print(f"当前工作目录: {os.getcwd()}")
    print(f"应用程序目录: {APP_DIR}")
    print(f"数据库文件: {DB_FILE}")
    print(f"备份目录: {BACKUP_DIR}")
    print(f"配置文件: {CONFIG_FILE}")
    print(f"{'='*60}\n")
    
    app = MainApplication(sys.argv)
    
    if app.run():
        print(f"\n{APP_NAME} 正常退出")
        sys.exit(app.exec())
    else:
        print(f"\n{APP_NAME} 退出（用户取消登录）")
        sys.exit(0)


if __name__ == "__main__":
    main()