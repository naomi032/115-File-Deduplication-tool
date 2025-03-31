import os
import configparser
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog, Menu, Toplevel
import threading
import time
from datetime import datetime, timezone # Added timezone
from collections import defaultdict, Counter
import json
import traceback
import collections
import math # For size conversion
import sys # To get base path for PyInstaller

# --- NEW: Import Pillow for background image ---
try:
    from PIL import Image, ImageTk
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False
    print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
    print("WARNING: Pillow library not found.")
    print("The background image feature will be disabled.")
    print("To enable it, install Pillow using: pip install Pillow")
    print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")


# --- Dependency Check & Handling: matplotlib ---
# ... (Matplotlib check code remains the same) ...
try:
    import matplotlib
    matplotlib.use('TkAgg') # Use Tkinter backend for embedding
    from matplotlib.figure import Figure
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
    MATPLOTLIB_AVAILABLE = True
    try:
        preferred_fonts = ['SimHei', 'Microsoft YaHei', 'MS Gothic', 'Malgun Gothic', 'Arial Unicode MS', 'sans-serif']
        matplotlib.rcParams['font.sans-serif'] = preferred_fonts
        matplotlib.rcParams['axes.unicode_minus'] = False
        print(f"Attempting to set Matplotlib font preference: {preferred_fonts}")
    except Exception as font_error:
        print(f"WARNING: Could not set preferred CJK font for Matplotlib - {font_error}")
except ImportError:
    MATPLOTLIB_AVAILABLE = False
    print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
    print("WARNING: matplotlib library not found.")
    print("The 'Show Cloud File Types' chart feature will be disabled.")
    print("To enable it, install matplotlib using: pip install matplotlib")
    print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
# --- End Dependency Check ---

# Make sure clouddrive library is installed
try:
    from clouddrive import CloudDriveClient, CloudDriveFileSystem
except ImportError:
    print("ERROR: The 'clouddrive' library is not installed. Please install it using: pip install clouddrive")
    try:
        root_tk = tk.Tk()
        root_tk.withdraw()
        messagebox.showerror("Missing Library", "The 'clouddrive' library is not installed.\nPlease install it using: pip install clouddrive")
    except Exception:
        pass
    exit()


# --- NEW: Function to get resource path for PyInstaller ---
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        # If not running as a PyInstaller bundle, use the script's directory
        base_path = os.path.abspath(os.path.dirname(__file__)) # Use __file__ instead of .

    return os.path.join(base_path, relative_path)

# --- Constants ---
VIDEO_EXTENSIONS = {
    ".mkv", ".iso", ".ts", ".mp4", ".avi", ".rmvb",
    ".wmv", ".m2ts", ".mpg", ".flv", ".rm", ".mov",
}
# Use resource_path for files needed by the application
CONFIG_FILE = resource_path("config.ini")
LANG_PREF_FILE = resource_path("lang_pref.json")
BACKGROUND_IMAGE_FILE = resource_path("background.png") # NEW
ICON_FILE = resource_path("app_icon.ico") # NEW

DATE_FORMAT = "%Y-%m-%d %H:%M:%S" # For display
DEFAULT_LANG = "en"
# Rule constants
RULE_KEEP_SHORTEST = "shortest"
RULE_KEEP_LONGEST = "longest"
RULE_KEEP_OLDEST = "oldest"
RULE_KEEP_NEWEST = "newest"
RULE_KEEP_SUFFIX = "suffix"

# --- Translations ---
# ... (Translations dictionary remains the same) ...
translations = {
    "en": {
        # Window & Config
        "window_title": "CloudDrive2 Duplicate Video Finder & Deleter",
        "config_title": "Configuration",
        "address_label": "CloudDrive2 Address:",
        "account_label": "CloudDrive2 Account:",
        "password_label": "CloudDrive2 Password:",
        "scan_path_label": "Root Path to Scan:",
        "mount_point_label": "CloudDrive2 Mount Point:",
        "load_config_button": "Load Config",
        "save_config_button": "Save Config",
        "test_connection_button": "Test Connection",

        # --- NEW/MODIFIED Main Actions ---
        "find_button": "Find Duplicates", # MODIFIED
        "find_starting": "Starting duplicate file scan...", # NEW
        "find_complete_found": "Scan complete. Found {count} sets of duplicate video files.", # NEW
        "find_complete_none": "Scan complete. No duplicate video files found based on SHA1 hash.", # NEW
        "find_error_during": "Error during duplicate scan: {error}", # NEW
        "delete_by_rule_button": "Delete Files by Rule", # NEW
        "delete_confirm_title": "Confirm Deletion by Rule", # NEW
        "delete_confirm_msg": "This will delete files based on the selected rule: '{rule_name}'.\n\nFiles marked for deletion in the list below will be removed permanently from the cloud drive.\n\nTHIS ACTION CANNOT BE UNDONE.\n\nAre you sure you want to proceed?", # NEW
        "delete_cancelled": "Deletion operation cancelled by user.", # NEW
        "delete_starting": "Starting deletion based on rule '{rule_name}'...", # NEW
        "delete_determining": "Determining files to delete based on rule '{rule_name}'...", # NEW
        "delete_no_rule_selected": "Error: Please select a deletion rule first.", # NEW
        "delete_suffix_missing": "Error: Please enter the suffix to keep when using the 'Keep Suffix' rule.", # NEW
        "delete_rule_no_files": "No files matched the deletion criteria for the selected rule.", # NEW
        "delete_finished": "Deletion complete. Deleted {deleted_count} of {total_marked} files.", # NEW
        "delete_error_during": "Error during deletion process: {error}", # NEW
        "delete_error_file": "Error deleting {path}: {error}", # NEW

        # Deletion Rules Section
        "rules_title": "Deletion Rules (Select ONE to Keep):", # NEW
        "rule_shortest_path": "Keep Shortest Path", # NEW
        "rule_longest_path": "Keep Longest Path", # NEW
        "rule_oldest": "Keep Oldest (Modified Date)", # NEW
        "rule_newest": "Keep Newest (Modified Date)", # NEW
        "rule_keep_suffix": "Keep files ending with:", # NEW (Changed wording)
        "rule_suffix_entry_label": "Suffix:", # NEW

        # Chart & Save Report
        "show_chart_button": "Show Cloud File Types",
        "show_chart_button_disabled": "Chart (Needs Connection & Matplotlib)",
        "save_list_button": "Save Found Duplicates Report...", # MODIFIED BUTTON TEXT
        "save_report_saved": "Duplicates report saved to {file}", # MODIFIED
        "save_report_error": "Error saving duplicates report to {file}: {error}", # MODIFIED
        "save_report_no_data": "No duplicate sets found to save.", # MODIFIED
        "save_report_header": "Duplicate Video File Sets Found (Based on SHA1):", # MODIFIED
        "save_report_set_header": "Set {index} (SHA1: {sha1}) - {count} files", # MODIFIED
        "save_report_file_label": "  - File:", # NEW
        "save_report_details_label": "    (Path: {path}, Modified: {modified}, Size: {size_mb:.2f} MB)", # NEW

        # Log & Status
        "log_title": "Log",
        "status_loading_config": "Attempting to load configuration from {file}...",
        "status_config_loaded": "Configuration loaded successfully.",
        "status_config_not_found": "Warning: Config file '{file}' not found. Using defaults/empty.",
        "status_config_section_missing": "Warning: '[config]' section not found in config file.",
        "status_saving_config": "Saving configuration to {file}...",
        "status_config_saved": "Configuration saved successfully.",
        "status_connecting": "Attempting connection...",
        "status_connect_success": "Connection successful.",
        "status_scan_progress": "Scanned {count} total items... Found {video_count} videos so far.", # NEW/Refined
        "status_scan_finished": "Scan finished. Checked {count} files.", # Kept for chart scan
        "status_populating_tree": "Populating results list with {count} duplicate sets...", # NEW
        "status_tree_populated": "Results list populated.", # NEW
        "status_applying_rule": "Applying rule '{rule_name}' to {count} sets...", # NEW
        "status_rule_applied": "Rule applied. Identified {delete_count} files to delete.", # NEW
        "status_clearing_tree": "Clearing results list...", # NEW

        # Treeview Columns
        "tree_rule_action_col": "Action", # NEW (Keep/Delete)
        "tree_path_col": "File Path",
        "tree_modified_col": "Date Modified",
        "tree_size_col": "Size (MB)",
        "tree_set_col": "Duplicate Set #", # Changed Text
        "tree_set_col_value": "{index}", # Changed Format
        "tree_action_keep": "Keep", # NEW
        "tree_action_delete": "Delete", # NEW

        # Errors & Warnings
        "error_input_missing": "Error: Please fill in Address, Account, Root Path, and Mount Point.",
        "error_input_missing_conn": "Error: Please fill in Address, Account, and Mount Point for connection test.",
        "error_input_missing_chart": "Error: Please fill in Root Path to Scan and Mount Point for chart.",
        "error_input_title": "Input Error",
        "error_config_read": "Error reading config file:\n{error}",
        "error_config_title": "Config Error",
        "error_config_save": "Could not write to config file:\n{error}",
        "error_config_save_title": "Config Save Error",
        "error_unexpected": "An unexpected error occurred:\n{error}",
        "error_title": "Error",
        "error_connect": "Error connecting to CloudDrive2: {error}",
        "error_scan_path": "Critical Error walking path '{path}': {error}",
        "error_get_attrs": "Error getting attributes for {path}: {error}",
        "error_no_duplicates_found": "No duplicates were found or displayed. Cannot apply deletion rule.", # NEW
        "error_parse_date": "Error parsing date for {path}: {error}", # NEW
        "error_not_connected": "Error: Not connected to CloudDrive2. Cannot perform action.",
        "error_path_calc_failed": "Error: Could not determine a valid cloud path from Root Path and Mount Point. Check inputs.",
        "warning_path_mismatch": "Warning: Could not determine a valid cloud path based on 'Root Path to Scan' ('{scan}') and 'Mount Point' ('{mount}'). Please check inputs.",
        "path_warning_title": "Path Input Warning",
        "path_warning_suspicious_chars": "Suspicious character(s) detected in input paths!\nThis often happens from copy-pasting.\nPlease DELETE and MANUALLY RETYPE the paths in the GUI.",
        "error_img_load": "Error loading background image '{path}': {error}", # NEW
        "error_icon_load": "Error loading icon '{path}': {error}", # NEW


        # Menu & Connection Test
        "menu_language": "Language",
        "menu_english": "English",
        "menu_chinese": "中文",
        "conn_test_success_title": "Connection Test Successful",
        "conn_test_success_msg": "Successfully connected to CloudDrive2.",
        "conn_test_fail_title": "Connection Test Failed",
        "conn_test_fail_msg": "Failed to connect. Check log for details.",

        # Charting (Mostly unchanged)
        "chart_status_scanning_cloud": "Scanning cloud path '{path}' for file types...",
        "chart_error_title": "Chart Error",
        "chart_info_title": "Chart Info",
        "chart_error_no_matplotlib": "Matplotlib library not found. Please install it (`pip install matplotlib`) to use this feature.",
        "chart_error_no_connection": "Cannot generate chart: Not connected to CloudDrive2. Please test connection first.",
        "chart_error_cloud_scan": "Error scanning cloud path '{path}': {error}",
        "chart_status_no_files_found": "No files found in the cloud path '{path}'.",
        "chart_status_generating": "Generating pie chart for {count} file types ({total} total files)...",
        "chart_window_title": "File Types in Cloud Path '{path}'",
        "chart_legend_title": "File Extensions",
        "chart_label_no_extension": "[No Extension]",
        "chart_label_others": "Others",
    },
    "zh": {
        # 窗口与配置
        "window_title": "CloudDrive2 重复视频查找与删除工具",
        "config_title": "配置",
        "address_label": "CloudDrive2 地址:",
        "account_label": "CloudDrive2 账号:",
        "password_label": "CloudDrive2 密码:",
        "scan_path_label": "要扫描的根路径:",
        "mount_point_label": "CloudDrive2 挂载点:",
        "load_config_button": "加载配置",
        "save_config_button": "保存配置",
        "test_connection_button": "测试连接",

        # --- 新增/修改的主要操作 ---
        "find_button": "查找重复项", # 修改
        "find_starting": "开始扫描重复文件...", # 新增
        "find_complete_found": "扫描完成。找到 {count} 组重复的视频文件。", # 新增
        "find_complete_none": "扫描完成。未根据 SHA1 哈希找到重复的视频文件。", # 新增
        "find_error_during": "扫描重复项期间出错: {error}", # 新增
        "delete_by_rule_button": "按规则删除文件", # 新增
        "delete_confirm_title": "确认按规则删除", # 新增
        "delete_confirm_msg": "此操作将根据选定规则删除文件: '{rule_name}'。\n\n下面列表中标记为删除的文件将从云盘中永久移除。\n\n此操作无法撤销。\n\n您确定要继续吗？", # 新增
        "delete_cancelled": "用户取消了删除操作。", # 新增
        "delete_starting": "开始根据规则 '{rule_name}' 删除文件...", # 新增
        "delete_determining": "正在根据规则 '{rule_name}' 确定要删除的文件...", # 新增
        "delete_no_rule_selected": "错误：请先选择一个删除规则。", # 新增
        "delete_suffix_missing": "错误：使用“保留后缀”规则时，请输入要保留的后缀名。", # 新增
        "delete_rule_no_files": "没有文件符合所选规则的删除条件。", # 新增
        "delete_finished": "删除完成。共删除了 {total_marked} 个标记文件中的 {deleted_count} 个。", # 新增 (修正: marked_count -> total_marked)
        "delete_error_during": "删除过程中出错: {error}", # 新增
        "delete_error_file": "删除 {path} 时出错: {error}", # 新增

        # 删除规则部分
        "rules_title": "删除规则 (选择一项保留):", # 新增
        "rule_shortest_path": "保留最短路径", # 新增
        "rule_longest_path": "保留最长路径", # 新增
        "rule_oldest": "保留最旧 (修改日期)", # 新增
        "rule_newest": "保留最新 (修改日期)", # 新增
        "rule_keep_suffix": "保留以此结尾的文件:", # 新增 (修改措辞)
        "rule_suffix_entry_label": "后缀:", # 新增

        # 图表和保存报告
        "show_chart_button": "显示云盘文件类型",
        "show_chart_button_disabled": "图表 (需连接 & Matplotlib)",
        "save_list_button": "保存找到的重复项报告...", # 修改按钮文字
        "save_report_saved": "重复项报告已保存至 {file}", # 修改
        "save_report_error": "保存重复项报告至 {file} 时出错: {error}", # 修改
        "save_report_no_data": "未找到可保存的重复文件集。", # 修改
        "save_report_header": "找到的重复视频文件集 (基于 SHA1):", # 修改
        "save_report_set_header": "集合 {index} (SHA1: {sha1}) - {count} 个文件", # 修改
        "save_report_file_label": "  - 文件:", # 新增
        "save_report_details_label": "    (路径: {path}, 修改日期: {modified}, 大小: {size_mb:.2f} MB)", # 新增

        # 日志与状态
        "log_title": "日志",
        "status_loading_config": "尝试从 {file} 加载配置...",
        "status_config_loaded": "配置加载成功。",
        "status_config_not_found": "警告: 配置文件 '{file}' 未找到。将使用默认/空值。",
        "status_config_section_missing": "警告: 配置文件中未找到 '[config]' 部分。",
        "status_saving_config": "正在保存配置到 {file}...",
        "status_config_saved": "配置保存成功。",
        "status_connecting": "尝试连接中...",
        "status_connect_success": "连接成功。",
        "status_scan_progress": "已扫描 {count} 个项目... 目前找到 {video_count} 个视频。", # 新增/优化
        "status_scan_finished": "扫描完成。共检查 {count} 个文件。", # 保留用于图表扫描
        "status_populating_tree": "正在使用 {count} 个重复文件集填充结果列表...", # 新增
        "status_tree_populated": "结果列表已填充。", # 新增
        "status_applying_rule": "正在对 {count} 个集合应用规则 '{rule_name}'...", # 新增
        "status_rule_applied": "规则已应用。识别出 {delete_count} 个待删除文件。", # 新增
        "status_clearing_tree": "正在清空结果列表...", # 新增

        # Treeview 列
        "tree_rule_action_col": "操作", # 新增 (保留/删除)
        "tree_path_col": "文件路径",
        "tree_modified_col": "修改日期",
        "tree_size_col": "大小 (MB)",
        "tree_set_col": "重复集合 #", # 修改文字
        "tree_set_col_value": "{index}", # 修改格式
        "tree_action_keep": "保留", # 新增
        "tree_action_delete": "删除", # 新增

        # 错误与警告
        "error_input_missing": "错误: 请填写地址、账号、扫描根路径和挂载点。",
        "error_input_missing_conn": "错误: 请填写地址、账号和挂载点以进行连接测试。",
        "error_input_missing_chart": "错误: 请填写扫描根路径和挂载点以生成图表。",
        "error_input_title": "输入错误",
        "error_config_read": "读取配置文件时出错:\n{error}",
        "error_config_title": "配置错误",
        "error_config_save": "无法写入配置文件:\n{error}",
        "error_config_save_title": "配置保存错误",
        "error_unexpected": "发生意外错误:\n{error}",
        "error_title": "错误",
        "error_connect": "连接 CloudDrive2 时出错: {error}",
        "error_scan_path": "遍历路径 '{path}' 时发生严重错误: {error}",
        "error_get_attrs": "获取 {path} 的属性时出错: {error}",
        "error_no_duplicates_found": "未找到或显示重复项。无法应用删除规则。", # 新增
        "error_parse_date": "解析 {path} 的日期时出错: {error}", # 新增
        "error_not_connected": "错误：未连接到 CloudDrive2。无法执行操作。",
        "error_path_calc_failed": "错误：无法根据扫描根路径和挂载点确定有效的云端路径。请检查输入。",
        "warning_path_mismatch": "警告：无法根据“要扫描的根路径” ('{scan}') 和“挂载点” ('{mount}') 确定有效的云端路径。请检查输入。",
        "path_warning_title": "路径输入警告",
        "path_warning_suspicious_chars": "在输入路径中检测到可疑字符！\n这通常是复制粘贴造成的。\n请在图形界面中删除并手动重新输入路径。",
        "error_img_load": "加载背景图片 '{path}' 时出错: {error}", # 新增
        "error_icon_load": "加载图标 '{path}' 时出错: {error}", # 新增

        # 菜单与连接测试
        "menu_language": "语言",
        "menu_english": "English",
        "menu_chinese": "中文",
        "conn_test_success_title": "连接测试成功",
        "conn_test_success_msg": "已成功连接到 CloudDrive2。",
        "conn_test_fail_title": "连接测试失败",
        "conn_test_fail_msg": "连接失败。请检查日志获取详细信息。",

        # 图表 (基本不变)
        "chart_status_scanning_cloud": "正在扫描云端路径 '{path}' 以统计文件类型...",
        "chart_error_title": "图表错误",
        "chart_info_title": "图表信息",
        "chart_error_no_matplotlib": "未找到 Matplotlib 库。请安装它 (`pip install matplotlib`) 以使用此功能。",
        "chart_error_no_connection": "无法生成图表：未连接到 CloudDrive2。请先测试连接。",
        "chart_error_cloud_scan": "扫描云端路径 '{path}' 时出错: {error}",
        "chart_status_no_files_found": "在云端路径 '{path}' 中未找到任何文件。",
        "chart_status_generating": "正在为 {count} 种文件类型（共 {total} 个文件）生成饼状图...",
        "chart_window_title": "云端路径 '{path}' 中的文件类型",
        "chart_legend_title": "文件扩展名",
        "chart_label_no_extension": "[无扩展名]",
        "chart_label_others": "其它",
    }
}


# --- Helper Functions (_validate_path_chars, _build_full_path, _parse_datetime) ---
# ... (Helper functions remain the same) ...
def _validate_path_chars(path_str):
    """Checks a single path string for suspicious characters."""
    suspicious_codes = []
    if not isinstance(path_str, str): return suspicious_codes
    KNOWN_INVISIBLE_CODES = {0x200B, 0x200E, 0x200F, 0xFEFF}
    for i, char in enumerate(path_str):
        char_code = ord(char)
        is_suspicious = False; reason = ""
        if 0 <= char_code <= 31 or char_code == 127: is_suspicious = True; reason = "Control Char (C0/DEL)"
        elif 128 <= char_code <= 159: is_suspicious = True; reason = "Control Char (C1)"
        elif char_code in KNOWN_INVISIBLE_CODES: is_suspicious = True; reason = "Known Invisible Char"
        if is_suspicious: suspicious_codes.append(f"U+{char_code:04X} ({reason}) at pos {i}")
    return suspicious_codes

def _build_full_path(parent_path, item_name):
    """Helper to correctly join cloud paths using forward slashes."""
    parent_path_norm = parent_path.replace('\\', '/').rstrip('/')
    item_name_norm = item_name.replace('\\', '/').lstrip('/')
    if not parent_path_norm or parent_path_norm == '/':
        return '/' + item_name_norm
    else:
        return parent_path_norm + '/' + item_name_norm

def _parse_datetime(date_string):
    """Parses common datetime string formats into timezone-aware datetime objects."""
    if not date_string: return None
    try:
        dt = datetime.fromisoformat(date_string.replace('Z', '+00:00'))
        return dt if dt.tzinfo else dt.replace(tzinfo=timezone.utc)
    except ValueError: pass
    try:
        dt = datetime.strptime(date_string, '%Y-%m-%d %H:%M:%S')
        return dt.replace(tzinfo=timezone.utc)
    except ValueError: pass
    try: # Handle format like '2023-10-27T08:15:30.123+00:00'
        # Python 3.7+ needed for %z with colon
        if sys.version_info >= (3, 7):
             # Attempt parsing with potential colon in timezone offset
             try:
                 dt = datetime.strptime(date_string, '%Y-%m-%dT%H:%M:%S.%f%z')
                 return dt
             except ValueError:
                 # Try removing colon if present for compatibility (e.g., +0800)
                 if len(date_string) > 6 and date_string[-3] == ':':
                      date_string_no_colon = date_string[:-3] + date_string[-2:]
                      try:
                           dt = datetime.strptime(date_string_no_colon, '%Y-%m-%dT%H:%M:%S.%f%z')
                           return dt
                      except ValueError: pass # Still failed
        else: # Older Python versions might not handle %z with colon well
             pass # Skip this format on older Python
    except ValueError: pass
    except IndexError: pass # Handle potential index errors if date string is malformed

    print(f"Warning: Could not parse date string: {date_string} with available formats.")
    return None


# --- DuplicateFileFinder Class ---
# ... (DuplicateFileFinder class remains the same) ...
class DuplicateFileFinder:
    def __init__(self):
        self.clouddrvie2_address = ""
        self.clouddrive2_account = ""
        self.clouddrive2_passwd = ""
        self._raw_scan_path = ""
        self._raw_mount_point = ""
        self.fs = None
        self.progress_callback = None
        self._ = lambda key, **kwargs: kwargs.get('default', key)

    def set_translator(self, translator_func):
        self._ = translator_func

    def log(self, message):
        if self.progress_callback:
            try:
                message_str = str(message)
                self.progress_callback(message_str)
            except Exception as e:
                print(f"Error in progress callback: {e}")
                print(f"Original message: {message}")
        else:
            print(str(message))

    def set_config(
            self,
            clouddrvie2_address,
            clouddrive2_account,
            clouddrive2_passwd,
            raw_scan_path,
            raw_mount_point,
            progress_callback=None,
    ):
        self.clouddrvie2_address = clouddrvie2_address
        self.clouddrive2_account = clouddrive2_account
        self.clouddrive2_passwd = clouddrive2_passwd
        self._raw_scan_path = raw_scan_path
        self._raw_mount_point = raw_mount_point
        self.progress_callback = progress_callback
        self.fs = None
        try:
            self.log(self._("status_connecting"))
            client = CloudDriveClient(
                self.clouddrvie2_address, self.clouddrive2_account, self.clouddrive2_passwd
            )
            self.fs = CloudDriveFileSystem(client)
            self.log("Testing connection by attempting to list root directory ('/')...")
            self.fs.ls('/') # Test connectivity
            self.log(self._("status_connect_success"))
            return True
        except Exception as e:
            error_msg = self._("error_connect", error=e)
            self.log(error_msg)
            self.log(f"Connection Error Details: {traceback.format_exc()}")
            self.fs = None
            return False

    def calculate_fs_path(self, scan_path_raw, mount_point_raw):
        scan_path_issues = _validate_path_chars(scan_path_raw)
        mount_point_issues = _validate_path_chars(mount_point_raw)
        if scan_path_issues or mount_point_issues:
            all_issues = []
            if scan_path_issues: all_issues.append(f"Scan Path: {', '.join(scan_path_issues)}")
            if mount_point_issues: all_issues.append(f"Mount Point: {', '.join(mount_point_issues)}")
            self.log(f"ERROR: Invalid characters found in path inputs: {'; '.join(all_issues)}")
            return None

        scan_path_norm = scan_path_raw.replace('\\', '/').strip().rstrip('/')
        mount_point_norm = mount_point_raw.replace('\\', '/').strip().rstrip('/')
        fs_dir_path = None

        # Logic from original script to determine cloud path relative to mount point
        # Case 1: Mount point is a Windows drive letter (e.g., "D:")
        if len(mount_point_norm) == 2 and mount_point_norm[1] == ':' and mount_point_norm[0].isalpha():
            mount_point_drive_prefix = mount_point_norm.lower() + '/'
            scan_path_lower = scan_path_norm.lower()
            if scan_path_lower.startswith(mount_point_drive_prefix):
                relative_part = scan_path_norm[len(mount_point_norm):]
                fs_dir_path = '/' + relative_part.lstrip('/')
            elif scan_path_lower == mount_point_norm.lower():
                 fs_dir_path = '/' # Scanning the root represented by the drive
        # Case 2: Mount point is empty or root "/"
        elif not mount_point_norm or mount_point_norm == '/':
            fs_dir_path = '/' + scan_path_norm.lstrip('/') # Scan path is relative to cloud root
        # Case 3: Mount point is an absolute path (e.g., "/CloudMount")
        elif mount_point_norm.startswith('/'):
             mount_point_prefix = mount_point_norm + '/'
             if scan_path_norm.startswith(mount_point_prefix):
                 relative_part = scan_path_norm[len(mount_point_norm):]
                 fs_dir_path = '/' + relative_part.lstrip('/') # Scan path is inside the mount path
             elif scan_path_norm == mount_point_norm:
                 fs_dir_path = '/' # Scanning the root represented by the mount path
        # Case 4: Mount point is a relative path (e.g., "MyFolder") - Interpret relative to cloud root
        elif '/' not in mount_point_norm and ':' not in mount_point_norm:
            mount_point_prefix = mount_point_norm + '/'
            # Check if scan path starts with mount point name (e.g., mount='Movies', scan='Movies/Action')
            if scan_path_norm.startswith(mount_point_prefix):
                 relative_part = scan_path_norm[len(mount_point_norm):]
                 fs_dir_path = '/' + relative_part.lstrip('/') # Relative path from root
            # Check if scan path starts with / followed by mount point (e.g., mount='Movies', scan='/Movies/Action')
            elif scan_path_norm.startswith('/' + mount_point_prefix):
                relative_part = scan_path_norm[len('/' + mount_point_norm):]
                fs_dir_path = '/' + relative_part.lstrip('/')
            # Check if scan path *is* the mount point name
            elif scan_path_norm == mount_point_norm:
                 fs_dir_path = '/' # Scanning the root represented by the folder name
            # Fallback: If scan path is absolute, maybe use it directly? Risky. Stick to original idea: fail if ambiguous.
            # elif scan_path_norm.startswith('/'):
            #      fs_dir_path = scan_path_norm # This might ignore the mount point concept


        if fs_dir_path is None:
            warning_msg_tmpl = self._("warning_path_mismatch",
                                      default="Warning: Could not determine a valid cloud path based on 'Root Path to Scan' ('{scan}') and 'Mount Point' ('{mount}'). Please check inputs.")
            self.log(warning_msg_tmpl.format(scan=scan_path_raw, mount=mount_point_raw))
            return None

        # Normalize the final path
        if fs_dir_path:
            while '//' in fs_dir_path: fs_dir_path = fs_dir_path.replace('//', '/')
            if not fs_dir_path.startswith('/'): fs_dir_path = '/' + fs_dir_path
            # Keep trailing slash for root, remove otherwise
            if len(fs_dir_path) > 1: fs_dir_path = fs_dir_path.rstrip('/')
        # Handle case where calculation results in empty string (should become root)
        elif not fs_dir_path:
            fs_dir_path = '/'

        self.log(f"Calculated effective cloud scan path: {fs_dir_path}")
        return fs_dir_path


    def find_duplicates(self):
        """
        Scans for duplicate video files using SHA1 hash.
        Returns a dictionary: { sha1: [FileInfo, FileInfo, ...], ... }
        where FileInfo is {'path': str, 'modified': datetime, 'size': int, 'sha1': str}
        Only includes sets with 2 or more files.
        """
        if not self.fs:
            self.log(self._("error_not_connected"))
            return {}

        fs_dir_path = self.calculate_fs_path(self._raw_scan_path, self._raw_mount_point)
        if fs_dir_path is None:
            self.log(self._("error_path_calc_failed"))
            return {}

        self.log(self._("find_starting"))

        potential_duplicates = defaultdict(list)
        count = 0
        video_files_checked = 0
        errors_getting_attrs = 0
        start_time = time.time()

        try:
            self.log(f"DEBUG: Attempting fs.walk_path('{fs_dir_path}')")
            # Use detail=True to get attributes directly from walk, potentially faster
            # Note: Check if clouddrive library's walk_path supports detail=True efficiently
            # Assuming it might not, stick to individual attr calls for now.
            walk_iterator = self.fs.walk_path(fs_dir_path) # detail=False (default)

            for foldername, _, filenames in walk_iterator:
                foldername_str = str(foldername).replace('\\', '/')

                for filename_obj in filenames:
                    count += 1
                    filename_str = str(filename_obj)

                    if count % 500 == 0:
                        self.log(self._("status_scan_progress", count=count, video_count=video_files_checked))

                    file_extension = os.path.splitext(filename_str)[1].lower()

                    if file_extension in VIDEO_EXTENSIONS:
                        video_files_checked += 1
                        current_file_path = _build_full_path(foldername_str, filename_str)

                        try:
                            # Get attributes: path, modTime, size, SHA1 hash
                            attrs = self.fs.attr(current_file_path)
                            file_sha1 = None
                            # Check for SHA1 hash (key '2' in fileHashes dictionary)
                            if 'fileHashes' in attrs and isinstance(attrs['fileHashes'], dict) and '2' in attrs['fileHashes']:
                                file_sha1 = attrs['fileHashes']['2']

                            if not file_sha1:
                                # self.log(f"WARNING: SHA1 hash not found for {current_file_path}. Skipping.")
                                errors_getting_attrs += 1
                                continue # Skip files without SHA1

                            # Get and parse modification time
                            mod_time_str = attrs.get('modifiedTime')
                            mod_time_dt = _parse_datetime(mod_time_str) # Parse to datetime
                            if mod_time_dt is None:
                                 self.log(self._("error_parse_date", path=current_file_path, error="Unknown format or missing"))
                                 # Decide how to handle missing dates: skip or use fallback?
                                 # Skipping ensures rules based on date work correctly.
                                 errors_getting_attrs += 1
                                 continue # Skip files with unparseable dates


                            # Get file size
                            file_size = attrs.get('size', 0) # Default to 0 if missing

                            # Store file info
                            file_info = {
                                'path': current_file_path,
                                'modified': mod_time_dt, # Store as datetime object
                                'size': int(file_size),
                                'sha1': file_sha1
                            }
                            potential_duplicates[file_sha1].append(file_info)

                        except Exception as e:
                            # Log error getting attributes for a specific file
                            # Avoid flooding log if many files fail (e.g., permission issues)
                            # Maybe log only first few errors? For now, log all.
                            err_msg = self._("error_get_attrs", path=current_file_path, error=e)
                            self.log(err_msg)
                            errors_getting_attrs += 1
                            # Optionally add traceback here if needed for debugging specific attr errors
                            # self.log(traceback.format_exc(limit=1))


            # --- Scan finished ---
            end_time = time.time()
            duration = end_time - start_time
            self.log(f"Scan finished in {duration:.2f} seconds. Total items encountered: {count}. Video files checked: {video_files_checked}.")
            if errors_getting_attrs > 0:
                self.log(f"WARNING: Encountered {errors_getting_attrs} errors while retrieving file attributes/hashes/dates.")

            # Filter for actual duplicates (sets with more than one file)
            actual_duplicates = {sha1: files for sha1, files in potential_duplicates.items() if len(files) > 1}

            if actual_duplicates:
                self.log(self._("find_complete_found", count=len(actual_duplicates)))
            else:
                self.log(self._("find_complete_none"))

            return actual_duplicates

        except Exception as walk_e:
            # Catch errors during the walk process itself (e.g., network error, path not found)
            err_msg = self._("error_scan_path", path=fs_dir_path, error=walk_e)
            self.log(err_msg)
            self.log(traceback.format_exc()) # Log full traceback for walk errors
            # Provide a more user-friendly summary error message
            self.log(self._("find_error_during", error=walk_e))
            return {} # Return empty on major error


    def write_duplicates_report(self, duplicate_sets, output_file):
        """ Writes the dictionary of found duplicate file sets to a text file. """
        if not duplicate_sets:
             self.log(self._("save_report_no_data"))
             return False
        try:
            with open(output_file, "w", encoding='utf-8') as f:
                header = self._("save_report_header")
                f.write(f"{header}\n===================================================\n\n")
                set_count = 0

                # Sort sets by SHA1 for consistent output
                sorted_sha1s = sorted(duplicate_sets.keys())

                for sha1 in sorted_sha1s:
                    files_in_set = duplicate_sets[sha1]
                    # Basic check, though find_duplicates should ensure > 1 file
                    if not files_in_set or len(files_in_set) < 2: continue

                    set_count += 1
                    set_header = self._("save_report_set_header", index=set_count, sha1=sha1, count=len(files_in_set))
                    f.write(f"{set_header}\n")

                    # Sort files within the set by path for clarity
                    sorted_files = sorted(files_in_set, key=lambda item: item['path'])

                    for file_info in sorted_files:
                         file_label = self._("save_report_file_label")
                         # Safely format datetime, handle potential None
                         mod_time_str = file_info['modified'].strftime(DATE_FORMAT) if isinstance(file_info.get('modified'), datetime) else "N/A"
                         # Safely calculate size, handle potential None or non-numeric
                         size_mb = file_info.get('size', 0) / (1024 * 1024) if isinstance(file_info.get('size'), (int, float)) else 0.0

                         details_label = self._("save_report_details_label",
                                                path=file_info['path'],
                                                modified=mod_time_str,
                                                size_mb=size_mb)
                         f.write(f"{file_label}\n{details_label}\n")
                    f.write("\n") # Blank line between sets

            save_msg = self._("save_report_saved", file=output_file)
            self.log(save_msg)
            return True
        except Exception as e:
             error_msg = self._("save_report_error", file=output_file, error=e)
             self.log(error_msg)
             self.log(traceback.format_exc()) # Log traceback for saving errors
             return False


    def delete_files(self, files_to_delete):
        """ Deletes a list of files from the cloud drive. """
        if not self.fs:
            self.log(self._("error_not_connected"))
            return 0, 0 # deleted_count, total_attempted

        deleted_count = 0
        total_to_delete = len(files_to_delete)
        errors_deleting = []

        if total_to_delete == 0:
            # This case should ideally be handled before calling delete_files
            self.log("No files provided for deletion.")
            return 0, 0

        # Use a more specific starting message if possible, or keep generic
        self.log(f"Attempting to delete {total_to_delete} files...")

        for i, file_path in enumerate(files_to_delete):
            cloud_path = file_path.replace('\\', '/') # Ensure forward slashes
            self.log(f"Deleting [{i+1}/{total_to_delete}]: {cloud_path}")
            try:
                self.fs.remove(cloud_path)
                deleted_count += 1
                time.sleep(0.05) # Keep small delay to potentially reduce API hammering
            except Exception as e:
                # Use the specific deletion error key from translations
                error_log_msg = self._("delete_error_file", path=cloud_path, error=e)
                self.log(error_log_msg)
                errors_deleting.append(cloud_path)
                # Log traceback for the first few deletion errors for debugging?
                # if len(errors_deleting) <= 5:
                #      self.log(traceback.format_exc(limit=1))


        # Use the specific deletion finished key
        # Note: total_marked seems more accurate based on the calling context (start_delete_by_rule_thread)
        finish_msg = self._("delete_finished", deleted_count=deleted_count, total_marked=total_to_delete)
        self.log(finish_msg)

        if errors_deleting:
             # Log summary of failed deletions
             self.log(f"WARNING: Failed to delete {len(errors_deleting)} files:")
             # Log first few failed paths for quick reference
             for failed_path in errors_deleting[:10]: # Log up to 10 failed paths
                 self.log(f"  - {failed_path}")
             if len(errors_deleting) > 10:
                 self.log(f"  ... and {len(errors_deleting) - 10} more.")

        return deleted_count, total_to_delete # Return actual attempted count


# --- GUI Application Class ---
class DuplicateFinderApp:
    def __init__(self, master):
        self.master = master
        self.current_language = self.load_language_preference()
        self.finder = DuplicateFileFinder()
        self.finder.set_translator(self._)

        self.duplicate_sets = {}
        self.treeview_item_map = {}
        self.files_to_delete_cache = []

        self.widgets = {}
        self.string_vars = {}
        self.entries = {}
        self.deletion_rule_var = tk.StringVar(value="")
        self.suffix_entry_var = tk.StringVar()

        # --- NEW: Store background image reference ---
        self.background_photo = None # Keep reference

        master.title(self._("window_title"))
        master.geometry("950x750")

        # --- NEW: Set Icon ---
        try:
            if os.path.exists(ICON_FILE):
                 # Recommended way for cross-platform compatibility if using PhotoImage for icon:
                 # icon_img = tk.PhotoImage(file=ICON_FILE)
                 # master.iconphoto(True, icon_img) # True for default icon for all windows
                 # Using iconbitmap (simpler for .ico on Windows, might vary on others)
                 master.iconbitmap(ICON_FILE)
            else:
                 self.log_message(f"Warning: Icon file not found at '{ICON_FILE}'")
        except tk.TclError as e:
            # Catch errors (e.g., invalid icon format, file not found handled above)
            icon_err_msg = self._("error_icon_load", path=ICON_FILE, error=e)
            self.log_message(icon_err_msg)
            print(f"Icon load error details: {e}") # Print details for debugging
        except Exception as e: # Catch other potential errors
             icon_err_msg = self._("error_icon_load", path=ICON_FILE, error=f"Unexpected error: {e}")
             self.log_message(icon_err_msg)


        # --- NEW: Setup Background Image ---
        self.setup_background()

        # --- Menu Bar ---
        self.menu_bar = Menu(master)
        master.config(menu=self.menu_bar)
        self.create_menus()

        # --- Paned Window (Place this on top of the background) ---
        # Make PanedWindow potentially transparent IF the theme supports it (unlikely)
        # Style might be needed, but often ineffective for background transparency
        s = ttk.Style()
        s.configure('Transparent.TPanedwindow', background='') # Example, likely won't work

        self.paned_window = ttk.PanedWindow(master, orient=tk.VERTICAL) # style='Transparent.TPanedwindow'
        # Use place to put it over the background label
        self.paned_window.place(relx=0, rely=0, relwidth=1, relheight=1)

        # --- Top Frame (Config & Actions) ---
        # Need to ensure frames added to PanedWindow are also potentially transparent
        # Or set their background color explicitly if transparency fails
        self.top_frame = ttk.Frame(self.paned_window, padding=5) # style='Transparent.TFrame' (Style likely needed)
        self.paned_window.add(self.top_frame, weight=0)

        # Config Frame
        self.widgets["config_frame"] = ttk.LabelFrame(self.top_frame, text=self._("config_title"), padding="10")
        self.widgets["config_frame"].grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        self.top_frame.columnconfigure(0, weight=1)

        # ... (Config labels and entries remain the same) ...
        config_labels = {
            "address": "address_label", "account": "account_label", "password": "password_label",
            "scan_path": "scan_path_label", "mount_point": "mount_point_label"
        }
        for i, (key, label_key) in enumerate(config_labels.items()):
            label_widget = ttk.Label(self.widgets["config_frame"], text=self._(label_key))
            label_widget.grid(row=i, column=0, padx=(5,2), pady=5, sticky=tk.W)
            self.widgets[f"label_{key}"] = label_widget

            var = tk.StringVar()
            self.string_vars[key] = var
            entry_args = {"textvariable": var, "width": 60}
            if key == "password": entry_args["show"] = "*"
            entry = ttk.Entry(self.widgets["config_frame"], **entry_args)
            entry.grid(row=i, column=1, padx=(2,5), pady=5, sticky=tk.EW)
            self.entries[key] = entry
        self.widgets["config_frame"].columnconfigure(1, weight=1)


        # Action Buttons Frame (Load, Save, Test, Find)
        action_button_frame = ttk.Frame(self.top_frame, padding="5")
        action_button_frame.grid(row=1, column=0, padx=5, pady=(5,0), sticky="ew") # Reduced bottom padding

        btn_info_actions = [
            ("load_button", "load_config_button", self.load_config, 5),
            ("save_button", "save_config_button", self.save_config, 5),
            ("test_conn_button", "test_connection_button", self.start_test_connection_thread, (10, 5)),
            ("find_button", "find_button", self.start_find_duplicates_thread, (15, 5)), # Renamed key & command
        ]

        for widget_key, text_key, command, padx_val in btn_info_actions:
             # Find disabled initially until connected
             initial_state = tk.NORMAL if widget_key in ["load_button", "save_button", "test_conn_button"] else tk.DISABLED
             # Get translated text, provide default if key missing during init
             button_text = self._(text_key, default=text_key.replace('_',' ').title())
             button = ttk.Button(action_button_frame, text=button_text, command=command, state=initial_state)
             button.pack(side=tk.LEFT, padx=padx_val, pady=5)
             self.widgets[widget_key] = button


        # --- Middle Frame (Rules, TreeView) ---
        self.middle_frame = ttk.Frame(self.paned_window, padding=5)
        self.paned_window.add(self.middle_frame, weight=1) # Allow this area to expand
        self.middle_frame.rowconfigure(1, weight=1)    # Allow tree frame row (row 1) to expand
        self.middle_frame.columnconfigure(0, weight=1) # Allow content column to expand

        # Deletion Rules Frame
        rules_frame = ttk.LabelFrame(self.middle_frame, text=self._("rules_title"), padding="10")
        rules_frame.grid(row=0, column=0, padx=5, pady=(10, 5), sticky="ew")
        self.widgets["rules_frame"] = rules_frame
        rules_frame.columnconfigure(1, weight=1) # Allow space next to radio buttons

        rule_options = [
            ("rule_shortest_path", RULE_KEEP_SHORTEST),
            ("rule_longest_path", RULE_KEEP_LONGEST),
            ("rule_oldest", RULE_KEEP_OLDEST),
            ("rule_newest", RULE_KEEP_NEWEST),
            ("rule_keep_suffix", RULE_KEEP_SUFFIX),
        ]
        self.rule_radios = {}
        for i, (text_key, value) in enumerate(rule_options):
            radio_text = self._(text_key, default=text_key.replace('_',' ').title())
            radio = ttk.Radiobutton(rules_frame, text=radio_text, variable=self.deletion_rule_var,
                                    value=value, command=self._on_rule_change, state=tk.DISABLED)
            radio.grid(row=i, column=0, padx=5, pady=2, sticky="w")
            self.rule_radios[value] = radio
            self.widgets[f"radio_{value}"] = radio

        # Suffix Entry
        suffix_label_text = self._("rule_suffix_entry_label", default="Suffix:")
        suffix_label = ttk.Label(rules_frame, text=suffix_label_text, state=tk.DISABLED)
        suffix_label.grid(row=len(rule_options)-1, column=1, padx=(10, 2), pady=2, sticky="w")
        self.widgets["suffix_label"] = suffix_label

        suffix_entry = ttk.Entry(rules_frame, textvariable=self.suffix_entry_var, width=15, state=tk.DISABLED)
        suffix_entry.grid(row=len(rule_options)-1, column=2, padx=(0, 5), pady=2, sticky="w")
        self.widgets["suffix_entry"] = suffix_entry
        self.entries["suffix"] = suffix_entry


        # TreeView Frame
        tree_frame = ttk.Frame(self.middle_frame) # No specific style needed here usually
        tree_frame.grid(row=1, column=0, padx=5, pady=5, sticky="nsew") # Fill available space
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        self.columns = ("action", "path", "modified", "size_mb", "set_id")
        self.tree = ttk.Treeview(tree_frame, columns=self.columns, show="headings", selectmode="none")
        self.widgets["treeview"] = self.tree
        self.setup_treeview_headings() # Call after columns defined

        # ... (Treeview columns and scrollbars remain the same) ...
        self.tree.column("action", width=60, anchor=tk.CENTER, stretch=tk.NO)
        self.tree.column("path", width=450, anchor=tk.W)
        self.tree.column("modified", width=140, anchor=tk.W)
        self.tree.column("size_mb", width=80, anchor=tk.E)
        self.tree.column("set_id", width=80, anchor=tk.CENTER)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')


        # --- Bottom Frame (Log & Final Action Buttons) ---
        self.bottom_frame = ttk.Frame(self.paned_window, padding=5)
        self.paned_window.add(self.bottom_frame, weight=0) # Fixed height log + buttons
        self.bottom_frame.columnconfigure(0, weight=1) # Log takes bulk of width
        self.bottom_frame.rowconfigure(1, weight=1) # Log text expands vertically

        # Delete Action / Report Buttons Frame
        final_action_frame = ttk.Frame(self.bottom_frame)
        final_action_frame.grid(row=0, column=0, columnspan=2, padx=5, pady=(5,0), sticky="ew")

        style = ttk.Style()
        try: style.configure("Danger.TButton", foreground="white", background="red", font=('TkDefaultFont', 10, 'bold'))
        except tk.TclError: style.configure("Danger.TButton", foreground="red") # Fallback

        # Delete Button
        delete_btn_text = self._("delete_by_rule_button", default="Delete by Rule")
        delete_button = ttk.Button(final_action_frame, text=delete_btn_text,
                                   command=self.start_delete_by_rule_thread, state=tk.DISABLED, style="Danger.TButton")
        delete_button.pack(side=tk.LEFT, padx=5, pady=5)
        self.widgets["delete_button"] = delete_button

        # Chart Button
        chart_button_state = tk.DISABLED # Needs connection first
        chart_button_text_key = "show_chart_button_disabled" if not MATPLOTLIB_AVAILABLE else "show_chart_button"
        chart_btn_text = self._(chart_button_text_key, default="Show Chart")
        chart_button = ttk.Button(final_action_frame, text=chart_btn_text, command=self.show_cloud_file_types, state=chart_button_state)
        chart_button.pack(side=tk.LEFT, padx=(10, 5), pady=5)
        self.widgets["chart_button"] = chart_button

        # Save Report Button
        save_report_btn_text = self._("save_list_button", default="Save Report")
        save_report_button = ttk.Button(final_action_frame, text=save_report_btn_text, command=self.save_duplicates_report, state=tk.DISABLED)
        save_report_button.pack(side=tk.LEFT, padx=5, pady=5)
        self.widgets["save_list_button"] = save_report_button


        # Log Frame
        log_frame_text = self._("log_title", default="Log")
        self.widgets["log_frame"] = ttk.LabelFrame(self.bottom_frame, text=log_frame_text, padding="5")
        self.widgets["log_frame"].grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(5,0))
        self.widgets["log_frame"].rowconfigure(0, weight=1)
        self.widgets["log_frame"].columnconfigure(0, weight=1)

        self.log_text = scrolledtext.ScrolledText(self.widgets["log_frame"], wrap=tk.WORD, height=8, width=100, state='disabled', relief=tk.SUNKEN, borderwidth=1) # Reduced height slightly
        self.log_text.grid(row=0, column=0, sticky="nsew")
        self.widgets["log_text"] = self.log_text

        # --- Initial State ---
        self.load_config() # Load config first
        self.update_ui_language() # Apply language based on pref/default
        self.set_ui_state(tk.NORMAL) # Initial setup enables config/test


        # --- NEW: Bind window resize event ---
        self.master.bind('<Configure>', self.on_resize)


    # --- NEW: Background Setup Method ---
    def setup_background(self):
        """Loads and places the background image."""
        if not PIL_AVAILABLE:
             self.log_message("Pillow library not found, cannot load background image.")
             return

        try:
            if not os.path.exists(BACKGROUND_IMAGE_FILE):
                self.log_message(f"Warning: Background image not found at '{BACKGROUND_IMAGE_FILE}'")
                return

            # Load image using Pillow
            original_image = Image.open(BACKGROUND_IMAGE_FILE)

            # --- Transparency Handling ---
            # Option 1 (Recommended): Pre-process the image file (background.png)
            # in an image editor to blend it with a light grey background (e.g., #F0F0F0)
            # to achieve the desired transparency effect. Then load it directly.

            # Option 2 (Simpler code, affects whole window): Make the whole window transparent
            # self.master.attributes('-alpha', 0.9) # Value between 0.0 (invisible) and 1.0 (opaque)
            # self.master.config(bg='systemTransparent') # Might be needed on macOS

            # Option 3 (Attempt alpha blending in code - complex, might be slow/buggy):
            # Not implemented here due to complexity.

            # For now, we just load the image as is. Assume pre-processing if transparency needed.
            # Resize the image to fit the initial window size (will be updated on resize)
            win_w = self.master.winfo_width() if self.master.winfo_width() > 1 else 950 # Initial guess
            win_h = self.master.winfo_height() if self.master.winfo_height() > 1 else 750
            # Maintain aspect ratio while covering the area (or stretch)
            # Let's resize to cover, maintaining aspect ratio (Image.LANCZOS for quality)
            img_w, img_h = original_image.size
            ratio = max(win_w/img_w, win_h/img_h)
            new_size = (int(img_w * ratio), int(img_h * ratio))
            resized_image = original_image.resize(new_size, Image.Resampling.LANCZOS)

            self.background_photo = ImageTk.PhotoImage(resized_image)

            # Create a Label to hold the image
            if "background_label" not in self.widgets:
                bg_label = tk.Label(self.master, image=self.background_photo)
                bg_label.place(x=0, y=0, relwidth=1, relheight=1) # Cover the whole window
                bg_label.lower() # Send it to the bottom of the stacking order
                self.widgets["background_label"] = bg_label
            else:
                # Update existing label's image
                self.widgets["background_label"].configure(image=self.background_photo)
                self.widgets["background_label"].lower() # Ensure it stays at the back


        except FileNotFoundError:
             err_msg = self._("error_img_load", path=BACKGROUND_IMAGE_FILE, error="File not found.")
             self.log_message(err_msg)
        except Exception as e:
            err_msg = self._("error_img_load", path=BACKGROUND_IMAGE_FILE, error=e)
            self.log_message(err_msg)
            self.log_message(traceback.format_exc())

    # --- NEW: Resize Event Handler ---
    def on_resize(self, event):
        """Handle window resize events to resize the background image."""
        # Only resize if Pillow is available and image was loaded
        if not PIL_AVAILABLE or not os.path.exists(BACKGROUND_IMAGE_FILE) or "background_label" not in self.widgets:
             return

        # Get current window size from the event
        win_w = event.width
        win_h = event.height

        # Debounce or throttle this if resizing is laggy
        # For simplicity, resize directly here

        try:
            # Reload the original image
            original_image = Image.open(BACKGROUND_IMAGE_FILE)
            img_w, img_h = original_image.size

            # Calculate new size to cover window while maintaining aspect ratio
            ratio = max(win_w / img_w, win_h / img_h)
            # Prevent potential zero division or tiny sizes if window is minimized weirdly
            if ratio <= 0 or img_w <= 0 or img_h <= 0:
                return
            new_size = (max(1, int(img_w * ratio)), max(1, int(img_h * ratio))) # Ensure size > 0

            # Resize using LANCZOS for better quality
            resized_image = original_image.resize(new_size, Image.Resampling.LANCZOS)

            # Update the PhotoImage object
            self.background_photo = ImageTk.PhotoImage(resized_image)

            # Configure the background label with the new image
            self.widgets["background_label"].configure(image=self.background_photo)
            # Ensure it stays behind other widgets
            self.widgets["background_label"].lower()

        except FileNotFoundError:
            # Image might have been deleted after startup
            self.log_message(f"Error: Background image '{BACKGROUND_IMAGE_FILE}' not found during resize.")
            if "background_label" in self.widgets:
                 self.widgets["background_label"].configure(image='') # Remove image if file gone
                 self.background_photo = None
        except Exception as e:
            # Log errors during resize/reload
            self.log_message(f"Error resizing background image: {e}")
            # self.log_message(traceback.format_exc(limit=2)) # Optional traceback


    # --- Language Handling ---
    # ... (_ method remains the same) ...
    # ... (save/load language preference remain the same) ...
    # ... (create_menus remains the same) ...
    # ... (change_language remains the same) ...
    # ... (update_ui_language remains the same, check widget keys) ...
    # ... (setup_treeview_headings remains the same) ...
    def _(self, key, **kwargs):
        lang_dict = translations.get(self.current_language, translations[DEFAULT_LANG])
        default_val = kwargs.pop('default', f"<{key}?>")
        base_string = lang_dict.get(key, translations[DEFAULT_LANG].get(key, default_val))
        try:
            # Only format if placeholders exist and args are provided
            if '{' in base_string and '}' in base_string and kwargs:
                 return base_string.format(**kwargs)
            else:
                 return base_string # Return base string if no kwargs or no placeholders
        except KeyError as e:
             # Log specific missing key for easier debugging of translations
             print(f"Warning: Formatting KeyError for key '{key}' ({self.current_language}): Missing key {e}. Kwargs: {kwargs}")
             return f"{base_string} [MISSING FORMAT KEY: {e}]"
        except Exception as e: # Catch other formatting errors
             print(f"Warning: Formatting failed for key '{key}' ({self.current_language}): {e}. Kwargs: {kwargs}")
             return f"{base_string} [FORMATTING ERROR]"

    def save_language_preference(self):
        try:
            with open(LANG_PREF_FILE, 'w', encoding='utf-8') as f:
                json.dump({"language": self.current_language}, f)
        except IOError as e:
            # Log error saving language pref
            print(f"Warning: Could not save language preference to {LANG_PREF_FILE}: {e}")
            self.log_message(f"Warning: Could not save language preference: {e}") # Also log to GUI
        except Exception as e:
             print(f"Error saving language preference: {e}")
             self.log_message(f"Error saving language preference: {e}")

    def load_language_preference(self):
        try:
            if os.path.exists(LANG_PREF_FILE):
                with open(LANG_PREF_FILE, 'r', encoding='utf-8') as f:
                    prefs = json.load(f)
                    lang = prefs.get("language", DEFAULT_LANG)
                    # Validate loaded language code against available translations
                    return lang if lang in translations else DEFAULT_LANG
        except (IOError, json.JSONDecodeError) as e:
            # Log error loading language pref (e.g., file corrupted)
            print(f"Warning: Could not load language preference from {LANG_PREF_FILE}: {e}")
            # No GUI log here as logger might not be ready yet
        except Exception as e:
             print(f"Unexpected error loading language preference: {e}")
        return DEFAULT_LANG # Default language if load fails

    def create_menus(self):
        lang_menu = Menu(self.menu_bar, tearoff=0)
        self.widgets["lang_menu"] = lang_menu
        # Labels updated in update_ui_language
        lang_menu.add_command(label="English", command=lambda: self.change_language("en"))
        lang_menu.add_command(label="中文", command=lambda: self.change_language("zh"))
        self.menu_bar.add_cascade(label="Language", menu=lang_menu)


    def change_language(self, lang_code):
        if lang_code in translations and lang_code != self.current_language:
            print(f"Changing language to: {lang_code}")
            self.current_language = lang_code
            self.finder.set_translator(self._) # Update finder's translator
            self.update_ui_language() # Update all GUI elements
            self.save_language_preference() # Save the new preference
        elif lang_code not in translations:
             # Log if attempted language is not supported
             print(f"Error: Language '{lang_code}' not supported.")
             self.log_message(f"Error: Language '{lang_code}' not supported.")

    def update_ui_language(self):
        print(f"Updating UI to language: {self.current_language}")
        # Ensure master window exists
        if not self.master or not self.master.winfo_exists():
             print("Error: Master window does not exist during UI language update.")
             return

        try: # Wrap in try/except to catch errors during update
            self.master.title(self._("window_title"))

            # Update Menu
            if "lang_menu" in self.widgets and self.widgets["lang_menu"].winfo_exists():
                try: self.menu_bar.entryconfig(0, label=self._("menu_language"))
                except Exception as e: print(f"Error updating menu cascade label: {e}")
                try: self.widgets["lang_menu"].entryconfig(0, label=self._("menu_english"))
                except Exception as e: print(f"Error updating menu item 1: {e}")
                try: self.widgets["lang_menu"].entryconfig(1, label=self._("menu_chinese"))
                except Exception as e: print(f"Error updating menu item 2: {e}")

            # Update LabelFrames
            for key, title_key in [("config_frame", "config_title"),
                                ("rules_frame", "rules_title"),
                                ("log_frame", "log_title")]:
                if key in self.widgets and self.widgets[key].winfo_exists():
                    try: self.widgets[key].config(text=self._(title_key))
                    except tk.TclError: pass # Ignore if widget destroyed mid-update

            # Update Labels
            label_keys = {"label_address": "address_label", "label_account": "account_label",
                        "label_password": "password_label", "label_scan_path": "scan_path_label",
                        "label_mount_point": "mount_point_label",
                        "suffix_label": "rule_suffix_entry_label"}
            for widget_key, text_key in label_keys.items():
                if widget_key in self.widgets and self.widgets[widget_key].winfo_exists():
                    try: self.widgets[widget_key].config(text=self._(text_key))
                    except tk.TclError: pass

            # Update Buttons
            button_keys = {"load_button": "load_config_button", "save_button": "save_config_button",
                        "test_conn_button": "test_connection_button",
                        "find_button": "find_button",
                        "delete_button": "delete_by_rule_button",
                        "save_list_button": "save_list_button"}
            for widget_key, text_key in button_keys.items():
                if widget_key in self.widgets and self.widgets[widget_key].winfo_exists():
                    # Special handle chart button text based on state/matplotlib
                    if widget_key == "chart_button":
                        is_enabled = self.widgets[widget_key].cget('state') == tk.NORMAL
                        effective_text_key = "show_chart_button" if (is_enabled and MATPLOTLIB_AVAILABLE) else "show_chart_button_disabled"
                        try: self.widgets[widget_key].config(text=self._(effective_text_key))
                        except tk.TclError: pass
                    else:
                        try: self.widgets[widget_key].config(text=self._(text_key))
                        except tk.TclError: pass

            # Update Radio Buttons
            radio_keys = {RULE_KEEP_SHORTEST: "rule_shortest_path", RULE_KEEP_LONGEST: "rule_longest_path",
                        RULE_KEEP_OLDEST: "rule_oldest", RULE_KEEP_NEWEST: "rule_newest",
                        RULE_KEEP_SUFFIX: "rule_keep_suffix"}
            for value, text_key in radio_keys.items():
                # Construct the widget key used during creation
                widget_key = f"radio_{value}"
                if widget_key in self.widgets and self.widgets[widget_key].winfo_exists():
                    try: self.widgets[widget_key].config(text=self._(text_key))
                    except tk.TclError: pass

            # Update Treeview Headings
            self.setup_treeview_headings()
            # Re-apply rule highlighting/action text if tree has items
            if self.duplicate_sets:
                self._apply_rule_to_treeview() # Ensure tree text reflects new language

            print("UI Language update complete.")

        except Exception as e:
             # Catch any unexpected error during the update process
             print(f"ERROR during UI language update: {e}")
             self.log_message(f"ERROR during UI language update: {e}")
             # Optionally log traceback for debugging
             # self.log_message(traceback.format_exc(limit=3))

    def setup_treeview_headings(self):
        """Sets or updates the text of the treeview headings."""
        heading_keys = { "action": "tree_rule_action_col", "path": "tree_path_col",
                         "modified": "tree_modified_col", "size_mb": "tree_size_col",
                         "set_id": "tree_set_col" }
        if "treeview" in self.widgets and self.widgets["treeview"].winfo_exists():
            for col_id, text_key in heading_keys.items():
                try:
                    # Check if the column actually exists before configuring
                    if col_id in self.widgets["treeview"]["columns"]:
                         # Use the translation function _ to get the current language text
                         self.widgets["treeview"].heading(col_id, text=self._(text_key))
                except tk.TclError: pass # Ignore if widget is destroyed during update


    # --- GUI Logic Methods ---
    # ... (log_message, _append_log remain the same) ...
    # ... (load_config, save_config remain the same, use resource_path) ...
    # ... (_check_path_chars remains the same) ...
    # ... (set_ui_state remains the same logic) ...
    # ... (start_find_duplicates_thread, _find_duplicates_worker, _process_find_results remain the same) ...
    # ... (clear_results, populate_treeview remain the same) ...
    # ... (_on_rule_change, _apply_rule_to_treeview remain the same) ...
    # ... (_determine_files_to_delete remains the same) ...
    # ... (start_delete_by_rule_thread, _delete_worker remain the same) ...
    # ... (save_duplicates_report remains the same) ...
    # ... (start_test_connection_thread, _test_connection_worker remain the same) ...
    # ... (show_cloud_file_types, _show_cloud_file_types_worker, _create_pie_chart_window remain the same) ...
    def log_message(self, message):
        """ Safely appends a message to the log ScrolledText widget from any thread. """
        message_str = str(message) if message is not None else ""
        # Check if master and log widget exist before scheduling
        if hasattr(self, 'master') and self.master and self.master.winfo_exists() and \
           self.widgets.get("log_text") and self.widgets["log_text"].winfo_exists():
            try:
                # Use 'after' for thread safety with Tkinter
                self.master.after(0, self._append_log, message_str)
            except (tk.TclError, RuntimeError): pass # Ignore errors if window closing
            except Exception as e: print(f"Error scheduling log message: {e}\nMessage: {message_str}")

    def _append_log(self, message):
        """ Internal method to append message to log widget (MUST run in main thread). """
        log_widget = self.widgets.get("log_text")
        if not log_widget or not log_widget.winfo_exists(): return
        try:
            current_state = log_widget.cget('state')
            log_widget.configure(state='normal')
            timestamp = datetime.now().strftime("%H:%M:%S")
            log_widget.insert(tk.END, f"[{timestamp}] {message}\n")
            log_widget.see(tk.END) # Scroll to the end
            log_widget.configure(state=current_state) # Restore original state
        except tk.TclError: pass # Widget might have been destroyed
        except Exception as e: print(f"Unexpected error appending log: {e}\nMessage: {message}")

    def load_config(self):
        # Now uses resource_path to find config.ini
        config_path = CONFIG_FILE
        self.log_message(self._("status_loading_config", file=os.path.basename(config_path)))
        config = configparser.ConfigParser()
        try:
            if not os.path.exists(config_path):
                 self.log_message(self._("status_config_not_found", file=os.path.basename(config_path)))
                 return
            read_files = config.read(config_path, encoding='utf-8')
            if not read_files:
                 # File exists but couldn't be parsed or read
                 self.log_message(f"Warning: Config file '{os.path.basename(config_path)}' exists but could not be read/parsed.")
                 return

            if 'config' in config:
                cfg = config['config']
                # Populate GUI fields from config
                self.string_vars["address"].set(cfg.get("clouddrvie2_address", ""))
                self.string_vars["account"].set(cfg.get("clouddrive2_account", ""))
                self.string_vars["password"].set(cfg.get("clouddrive2_passwd", ""))
                # Use original keys 'root_path' and 'clouddrive2_root_path' for compatibility
                self.string_vars["scan_path"].set(cfg.get("root_path", ""))
                self.string_vars["mount_point"].set(cfg.get("clouddrive2_root_path", ""))
                self.log_message(self._("status_config_loaded"))
            else:
                self.log_message(self._("status_config_section_missing"))
                # Optionally set default values here if section missing

        except configparser.Error as e:
            # Error specifically during parsing
            error_msg = self._("error_config_read", error=e)
            # Show error to user and log it
            if self.master.winfo_exists(): # Check if GUI is ready
                 messagebox.showerror(self._("error_config_title"), error_msg)
            self.log_message(error_msg)
        except Exception as e:
             # Catch other potential errors (e.g., file permission)
             error_msg = f"{self._('error_unexpected', error='').rstrip(': ')} loading config: {e}"
             if self.master.winfo_exists():
                 messagebox.showerror(self._("error_title"), error_msg)
             self.log_message(error_msg)
             self.log_message(traceback.format_exc()) # Log details


    def save_config(self):
        # Now uses resource_path to find config.ini
        config_path = CONFIG_FILE
        self.log_message(self._("status_saving_config", file=os.path.basename(config_path)))
        config = configparser.ConfigParser()

        # Prepare the data to save
        config_data = {
            "clouddrvie2_address": self.string_vars["address"].get(),
            "clouddrive2_account": self.string_vars["account"].get(),
            "clouddrive2_passwd": self.string_vars["password"].get(),
            "root_path": self.string_vars["scan_path"].get(), # Use original key
            "clouddrive2_root_path": self.string_vars["mount_point"].get(), # Use original key
        }
        config['config'] = config_data

        # Attempt to preserve other sections if the file exists
        try:
            if os.path.exists(config_path):
                 config_old = configparser.ConfigParser()
                 config_old.read(config_path, encoding='utf-8')
                 for section in config_old.sections():
                     if section != 'config':
                         # Ensure the section exists before adding items
                         if not config.has_section(section):
                             config.add_section(section)
                         # Copy items from old section to new config
                         config[section] = dict(config_old.items(section))
                     else:
                         # Merge keys within 'config' section, prioritizing new values
                         # Add old keys only if they don't exist in the new data
                         for key, value in config_old.items(section):
                             if key not in config['config']:
                                 config['config'][key] = value
        except Exception as e:
            # Log warning if merging old config fails, but proceed with saving new data
            print(f"Warning: Could not merge old config sections during save: {e}")
            self.log_message(f"Warning: Could not merge old config sections during save: {e}")


        try:
            # Write the combined config to the file
            with open(config_path, 'w', encoding='utf-8') as configfile:
                config.write(configfile)
            self.log_message(self._("status_config_saved"))
        except IOError as e:
            # Specific error for file writing issues (permissions, etc.)
            error_msg = self._("error_config_save", error=e)
            if self.master.winfo_exists():
                 messagebox.showerror(self._("error_config_save_title"), error_msg)
            self.log_message(error_msg)
        except Exception as e:
             # Catch other unexpected errors during save
             error_msg = f"{self._('error_unexpected', error='').rstrip(': ')} saving config: {e}"
             if self.master.winfo_exists():
                 messagebox.showerror(self._("error_title"), error_msg)
             self.log_message(error_msg)
             self.log_message(traceback.format_exc()) # Log details

    def _check_path_chars(self, path_dict):
        # (Unchanged - uses the corrected validation logic)
        suspicious_char_found = False
        all_details = []
        # Get display names using current language
        path_display_names = {
            "address": self._("address_label").rstrip(': '),
            "scan_path": self._("scan_path_label").rstrip(': '),
            "mount_point": self._("mount_point_label").rstrip(': ')
        }
        # Add other paths if needed here
        # path_display_names["other_path"] = self._("other_path_label").rstrip(':')

        for key, path_str in path_dict.items():
            issues = _validate_path_chars(path_str)
            if issues:
                suspicious_char_found = True
                # Get the translated display name, fallback to the key name
                display_name = path_display_names.get(key, key)
                all_details.append(f"'{display_name}': {', '.join(issues)}")

        if suspicious_char_found:
            log_sep = "!" * 70
            self.log_message(log_sep)
            warning_title = self._("path_warning_title", default="Path Input Warning")
            self.log_message(f"*** {warning_title} ***")
            for detail in all_details: self.log_message(f"  -> {detail}")

            # Get the multi-line warning message
            warning_msg_template = self._("path_warning_suspicious_chars", default="Suspicious chars detected!\nCheck and retype paths.")
            warning_lines = warning_msg_template.split('\n')
            # Format for popup (first line as title/header, rest as body)
            popup_msg = f"{warning_lines[0]}\n\n" + "\n".join(warning_lines[1:]) if len(warning_lines) > 1 else warning_lines[0]
            instruction = warning_lines[1] if len(warning_lines) > 1 else "Please check and retype the paths."

            self.log_message(f"Instruction: {instruction}")
            self.log_message(log_sep)

            # Show error message box in the main thread
            if self.master.winfo_exists():
                self.master.after(0, messagebox.showerror, warning_title, popup_msg)
            return False # Indicate failure due to suspicious chars

        return True # Indicate success (no suspicious chars found)


    def set_ui_state(self, new_state):
        """Enable/disable UI elements based on state (e.g., 'normal', 'finding', 'deleting')."""
        # Define states
        STATE_NORMAL = 'normal'
        STATE_FINDING = 'finding'
        STATE_DELETING = 'deleting'
        STATE_TESTING_CONN = 'testing_connection'
        STATE_CHARTING = 'charting'

        is_normal_op = (new_state == STATE_NORMAL)
        is_busy = new_state in [STATE_FINDING, STATE_DELETING, STATE_TESTING_CONN, STATE_CHARTING]

        # Determine capabilities based on current app state (connection, results)
        has_connection = self.finder is not None and self.finder.fs is not None
        has_duplicates = bool(self.duplicate_sets)

        # --- Define enable/disable logic for each widget group ---
        can_config_interact = is_normal_op # Can load/save/edit config only when not busy
        can_test_conn = is_normal_op # Can test connection only when not busy
        can_find = is_normal_op and has_connection # Can start find only when connected and not busy
        can_apply_rules = is_normal_op and has_duplicates # Can select rules only when duplicates are found and not busy
        can_delete = is_normal_op and has_duplicates and bool(self.deletion_rule_var.get()) and bool(self.files_to_delete_cache) # Can delete only if rule selected, files marked, and not busy
        can_save_report = is_normal_op and has_duplicates # Can save report if duplicates found and not busy
        can_chart = is_normal_op and has_connection and MATPLOTLIB_AVAILABLE # Can chart if connected, matplotlib available, and not busy

        # --- Apply States to Widgets ---

        # Config Entries (Address, Account, Password, Paths)
        entry_state = tk.NORMAL if can_config_interact else tk.DISABLED
        for key, entry in self.entries.items():
            # Keep suffix entry tied to rule radio state (handled below)
            if key != "suffix" and entry.winfo_exists():
                try: entry.config(state=entry_state)
                except tk.TclError: pass

        # Config Buttons (Load, Save)
        config_btn_state = tk.NORMAL if can_config_interact else tk.DISABLED
        if self.widgets.get("load_button"): self.widgets["load_button"].config(state=config_btn_state)
        if self.widgets.get("save_button"): self.widgets["save_button"].config(state=config_btn_state)

        # Test Connection Button
        test_conn_btn_state = tk.NORMAL if can_test_conn else tk.DISABLED
        if self.widgets.get("test_conn_button"): self.widgets["test_conn_button"].config(state=test_conn_btn_state)

        # Find Duplicates Button
        find_btn_state = tk.NORMAL if can_find else tk.DISABLED
        if self.widgets.get("find_button"): self.widgets["find_button"].config(state=find_btn_state)

        # Deletion Rules Widgets (Radios, Suffix Label/Entry)
        rules_state = tk.NORMAL if can_apply_rules else tk.DISABLED
        if self.widgets.get("rules_frame"): # Check frame exists
            # Enable/disable radio buttons
            for radio in self.rule_radios.values():
                 if radio.winfo_exists():
                     try: radio.config(state=rules_state)
                     except tk.TclError: pass
            # Enable/disable suffix label/entry based on main rule state AND selected rule
            suffix_widgets_state = tk.DISABLED # Default disabled
            if rules_state == tk.NORMAL and self.deletion_rule_var.get() == RULE_KEEP_SUFFIX:
                suffix_widgets_state = tk.NORMAL # Enable only if rules enabled AND suffix rule selected
            if self.widgets.get("suffix_label"): self.widgets["suffix_label"].config(state=suffix_widgets_state)
            if self.widgets.get("suffix_entry"): self.widgets["suffix_entry"].config(state=suffix_widgets_state)

        # Delete Button
        # Check can_delete conditions: normal state, duplicates found, rule selected, files identified by rule
        delete_btn_state = tk.NORMAL if can_delete else tk.DISABLED
        if self.widgets.get("delete_button"): self.widgets["delete_button"].config(state=delete_btn_state)

        # Save Report Button
        save_report_btn_state = tk.NORMAL if can_save_report else tk.DISABLED
        if self.widgets.get("save_list_button"): self.widgets["save_list_button"].config(state=save_report_btn_state)

        # Chart Button
        chart_btn_state = tk.NORMAL if can_chart else tk.DISABLED
        if self.widgets.get("chart_button"):
            chart_text_key = "show_chart_button" if (chart_btn_state == tk.NORMAL and MATPLOTLIB_AVAILABLE) else "show_chart_button_disabled"
            self.widgets["chart_button"].config(state=chart_btn_state, text=self._(chart_text_key))

        # Treeview and Log are generally always visible, their content changes.
        # No need to disable the widgets themselves usually.

    def start_find_duplicates_thread(self):
        """ Handles the 'Find Duplicates' button click. """
        # 1. Check connection
        if not self.finder or not self.finder.fs:
             # Show warning and log if not connected
             messagebox.showwarning(self._("error_title"), self._("error_not_connected"))
             self.log_message(self._("error_not_connected"))
             return

        # 2. Check essential path inputs
        paths_to_check = {
            "scan_path": self.string_vars["scan_path"].get(),
            "mount_point": self.string_vars["mount_point"].get()
            # Add address? Usually checked by connection test, but maybe add here too?
            # "address": self.string_vars["address"].get(),
        }
        # Check for missing inputs needed for scan path calculation
        if not paths_to_check["scan_path"] or not paths_to_check["mount_point"]:
             # Use the generic missing input error message
              messagebox.showerror(self._("error_input_title"), self._("error_input_missing"))
              self.log_message(self._("error_input_missing") + " (Scan Path / Mount Point)")
              return

        # 3. Check for suspicious characters in paths
        if not self._check_path_chars(paths_to_check):
            # Error message already shown by _check_path_chars
            return

        # 4. Clear previous results and UI state before starting
        self.clear_results() # Clears tree, stored data, rule selection, delete cache
        self.log_message(self._("find_starting"))
        self.set_ui_state("finding") # Set specific state to disable UI during find

        # 5. Start the worker thread
        thread = threading.Thread(target=self._find_duplicates_worker, daemon=True)
        thread.start()

    def _find_duplicates_worker(self):
        """ Worker thread for finding duplicates. """
        # Double check connection within thread just in case
        if not self.finder or not self.finder.fs:
            self.log_message("Error: Connection lost before Find could execute.")
            # Schedule UI reset in main thread
            if self.master.winfo_exists(): self.master.after(0, self.set_ui_state, 'normal')
            return

        found_duplicates = {} # Initialize result
        try:
            start_time = time.time()
            # Call the finder method to get duplicate sets
            found_duplicates = self.finder.find_duplicates()
            end_time = time.time()
            # Log duration of the core find operation
            self.log_message(f"Find duplicates network/scan part took {end_time - start_time:.2f} seconds.")

            # --- Schedule GUI update in main thread ---
            # Pass the found duplicates to the processing function
            if self.master.winfo_exists():
                self.master.after(0, self._process_find_results, found_duplicates)

        except Exception as e:
            # Catch unexpected errors during the finder.find_duplicates() call itself
            err_msg = self._("find_error_during", error=e)
            self.log_message(err_msg)
            self.log_message(traceback.format_exc()) # Log details
            if self.master.winfo_exists():
                 # Schedule error message box and reset UI state
                 self.master.after(0, messagebox.showerror, self._("error_title"), err_msg)
                 self.master.after(0, self.set_ui_state, 'normal') # Reset state on error
        # No finally block here, _process_find_results handles the final UI state change upon success


    def _process_find_results(self, found_duplicates):
        """ Processes results from find_duplicates worker (runs in main thread). """
        if not self.master.winfo_exists(): return # Abort if window closed

        # Store the results
        self.duplicate_sets = found_duplicates if found_duplicates else {}

        if self.duplicate_sets:
            # Populate the treeview if duplicates were found
            self.populate_treeview() # This logs population start/end and applies rule
            # Log summary (already logged by find_duplicates and populate_treeview)
            # self.log_message(self._("find_complete_found", count=len(self.duplicate_sets)))
        else:
            # No duplicates found, message already logged by find_duplicates
             pass # No need to populate tree

        # Set final UI state back to normal after processing is complete
        self.set_ui_state('normal')


    def clear_results(self):
        """Clears the treeview, stored duplicates, rule selection, and delete cache."""
        self.log_message(self._("status_clearing_tree", default="Clearing results...")) # Provide default text
        self.duplicate_sets = {}
        self.treeview_item_map = {}
        self.files_to_delete_cache = []
        self.deletion_rule_var.set("") # Deselect rule radio button

        # Clear Treeview items
        tree = self.widgets.get("treeview")
        if tree and tree.winfo_exists():
            try:
                # Efficiently delete all top-level items and their children
                tree.delete(*tree.get_children())
            except tk.TclError: pass # Ignore if tree destroyed during clear

        # Update UI state to reflect cleared results (disables rules/delete buttons)
        # Call set_ui_state AFTER clearing data, otherwise conditions might be wrong
        self.set_ui_state('normal') # Reset to normal state, which will disable buttons based on empty data


    def populate_treeview(self):
        """ Populates the treeview with found duplicate sets. """
        tree = self.widgets.get("treeview")
        if not tree or not tree.winfo_exists():
            self.log_message("Error: Treeview widget not available for population.")
            return # Cannot proceed

        # Log start of population
        count = len(self.duplicate_sets)
        self.log_message(self._("status_populating_tree", count=count, default=f"Populating list with {count} sets..."))
        start_time = time.time()

        # --- Clear existing items and mapping before populating ---
        try:
            tree.delete(*tree.get_children())
        except tk.TclError:
            self.log_message("Error clearing treeview before population.")
            return # Abort if tree destroyed

        self.treeview_item_map.clear() # Clear the item-to-data mapping

        # --- Prepare data for insertion ---
        set_index = 0
        tree_items_data = [] # Store tuples: (item_id, values, file_info)

        # Sort by SHA1 for consistent set numbering
        sorted_sha1s = sorted(self.duplicate_sets.keys())

        for sha1 in sorted_sha1s:
            files_in_set = self.duplicate_sets[sha1]
            # Ensure it's actually a duplicate set (should be pre-filtered)
            if len(files_in_set) < 2: continue

            set_index += 1
            # Sort files within set by path for initial display order
            sorted_files = sorted(files_in_set, key=lambda x: x.get('path', '')) # Use get with default

            for file_info in sorted_files:
                # Safely get and format data
                path = file_info.get('path', 'N/A')
                mod_time = file_info.get('modified')
                mod_time_str = mod_time.strftime(DATE_FORMAT) if isinstance(mod_time, datetime) else "N/A"
                size = file_info.get('size', 0)
                size_mb = size / (1024 * 1024) if isinstance(size, (int, float)) else 0.0
                set_id_str = self._("tree_set_col_value", index=set_index, default=str(set_index))

                # Values tuple must match the `self.columns` order
                values = (
                    "", # Placeholder for 'Action' column
                    path,
                    mod_time_str,
                    f"{size_mb:.2f}", # Format size to 2 decimal places
                    set_id_str
                 )
                # Use the file path as the unique item ID (iid) in the tree
                # Ensure path is not empty, fallback if necessary though unlikely
                item_id = path if path != 'N/A' else f"set{set_index}_sha{sha1[:8]}_{mod_time_str}" # Fallback iid
                tree_items_data.append((item_id, values, file_info))

        # --- Insert items into the Treeview ---
        # ttk.Treeview doesn't have a true batch insert. Insert one by one.
        # This can be slow for very large numbers of items.
        items_inserted = 0
        for item_id, values, file_info in tree_items_data:
             try:
                 # Insert item with its unique ID and values
                 tree.insert("", tk.END, iid=item_id, values=values)
                 # Map the item ID back to the full file_info dictionary
                 self.treeview_item_map[item_id] = file_info
                 items_inserted += 1
             except tk.TclError as e:
                 # Handle potential error if item_id (path) causes issues
                 # or if tree is destroyed during insertion
                 self.log_message(f"Error inserting item '{item_id}' into tree: {e}")
                 # Continue attempting to insert others if possible

        end_time = time.time()
        duration = end_time - start_time
        # Log completion status
        self.log_message(self._("status_tree_populated", default="List populated.") + f" ({items_inserted} items in {duration:.2f}s)")

        # --- Apply the currently selected rule (if any) ---
        # This will update the 'Action' column and apply highlighting
        self._apply_rule_to_treeview()


    def _on_rule_change(self):
        """Called when a deletion rule radio button is selected."""
        # Update suffix entry state based on selected rule
        selected_rule = self.deletion_rule_var.get()
        is_suffix_rule = (selected_rule == RULE_KEEP_SUFFIX)
        suffix_widgets_state = tk.NORMAL if is_suffix_rule else tk.DISABLED

        if self.widgets.get("suffix_label"): self.widgets["suffix_label"].config(state=suffix_widgets_state)
        if self.widgets.get("suffix_entry"): self.widgets["suffix_entry"].config(state=suffix_widgets_state)

        # Clear suffix entry text if the rule is not 'Keep Suffix'
        if not is_suffix_rule:
            self.suffix_entry_var.set("")

        # --- Re-apply the selected rule to update the Treeview ---
        # This updates the 'Action' column and highlighting immediately
        self._apply_rule_to_treeview()


    def _apply_rule_to_treeview(self):
        """ Updates the 'Action' column and highlighting in the treeview based on the selected rule. """
        # Check if there's data and the treeview exists
        tree = self.widgets.get("treeview")
        if not self.duplicate_sets or not tree or not tree.winfo_exists():
            # Clear cache if no data/tree
            self.files_to_delete_cache = []
            # Ensure delete button is disabled if no rule applied/possible
            if self.widgets.get("delete_button"): self.widgets["delete_button"].config(state=tk.DISABLED)
            return

        selected_rule = self.deletion_rule_var.get()

        # --- Clear Action column and cache if no rule is selected ---
        if not selected_rule:
            self.files_to_delete_cache = [] # Clear deletion list
            try:
                # Reset action column and tags for all items
                for item_id in self.treeview_item_map.keys():
                    if tree.exists(item_id):
                         tree.set(item_id, "action", "")
                         tree.item(item_id, tags=()) # Remove tags
            except tk.TclError: pass # Ignore if tree destroyed
            # Disable delete button explicitly
            if self.widgets.get("delete_button"): self.widgets["delete_button"].config(state=tk.DISABLED)
            return

        # --- Apply the selected rule ---
        rule_name_display = self._(f"rule_{selected_rule}", default=selected_rule)
        self.log_message(self._("status_applying_rule", rule_name=rule_name_display, count=len(self.duplicate_sets), default=f"Applying rule '{rule_name_display}'..."))
        start_time = time.time()

        # Get translated text for Keep/Delete actions
        keep_text = self._("tree_action_keep", default="Keep")
        delete_text = self._("tree_action_delete", default="Delete")
        suffix = self.suffix_entry_var.get() if selected_rule == RULE_KEEP_SUFFIX else None

        # Define tags for highlighting (configure once, apply per item)
        try:
            tree.tag_configure('keep', foreground='green')
            tree.tag_configure('delete', foreground='red', font=('TkDefaultFont', 9, 'bold')) # Make delete bold red
        except tk.TclError: pass # Ignore if tree destroyed

        try:
            # --- Determine files to delete based on the rule ---
            # This call might raise ValueError if input is invalid (e.g., missing suffix)
            files_to_delete_list = self._determine_files_to_delete(self.duplicate_sets, selected_rule, suffix)

            # Cache the list of paths to be deleted
            self.files_to_delete_cache = files_to_delete_list
            files_to_delete_paths_set = set(files_to_delete_list) # Use set for fast lookups below
            delete_count = len(files_to_delete_paths_set)
            self.log_message(self._("status_rule_applied", delete_count=delete_count, default=f"Rule applied. {delete_count} files marked for deletion."))

            # --- Update Treeview items based on deletion list ---
            for item_id, file_info in self.treeview_item_map.items():
                 if not tree.exists(item_id): continue # Skip if item somehow removed

                 path = file_info.get('path')
                 if not path: continue # Skip if file info is malformed

                 # Determine action and tags
                 is_marked_for_delete = (path in files_to_delete_paths_set)
                 action_text = delete_text if is_marked_for_delete else keep_text
                 item_tags = ('delete',) if is_marked_for_delete else ('keep',)

                 # Update the treeview item's action column and tags
                 tree.set(item_id, "action", action_text)
                 tree.item(item_id, tags=item_tags)

            # --- Update Delete Button State ---
            # Enable delete button only if there are files marked for deletion
            delete_button_state = tk.NORMAL if delete_count > 0 else tk.DISABLED
            if self.widgets.get("delete_button"): self.widgets["delete_button"].config(state=delete_button_state)


        except ValueError as ve: # Catch specific errors from _determine_files_to_delete
             # Log and show the specific validation error (e.g., "Suffix missing")
             self.log_message(str(ve))
             messagebox.showerror(self._("error_input_title"), str(ve))
             # Clear action column and cache on rule application error
             self.files_to_delete_cache = []
             try:
                 for item_id in self.treeview_item_map.keys():
                     if tree.exists(item_id):
                         tree.set(item_id, "action", "")
                         tree.item(item_id, tags=()) # Clear tags
             except tk.TclError: pass
             # Disable delete button
             if self.widgets.get("delete_button"): self.widgets["delete_button"].config(state=tk.DISABLED)

        except tk.TclError:
             self.log_message("Error updating treeview, it might have been closed.")
             # Also clear cache if tree is gone
             self.files_to_delete_cache = []
             if self.widgets.get("delete_button"): self.widgets["delete_button"].config(state=tk.DISABLED)
        except Exception as e:
             # Catch any other unexpected error during rule application
             self.log_message(f"Unexpected error applying rule to treeview: {e}")
             self.log_message(traceback.format_exc())
             # Clear cache and disable delete on unexpected error
             self.files_to_delete_cache = []
             if self.widgets.get("delete_button"): self.widgets["delete_button"].config(state=tk.DISABLED)
        finally:
            end_time = time.time()
            self.log_message(f"Rule application / tree update took {end_time-start_time:.3f}s")


    def _determine_files_to_delete(self, duplicate_sets, rule, suffix_value):
        """
        Determines which files to delete based on the selected rule.
        Returns a list of full paths to delete.
        Raises ValueError for invalid input (e.g., missing suffix).
        """
        # --- Input Validation ---
        if not duplicate_sets: return [] # No duplicates, nothing to delete
        if not rule: raise ValueError(self._("delete_no_rule_selected", default="No deletion rule selected."))
        if rule == RULE_KEEP_SUFFIX and not suffix_value:
             raise ValueError(self._("delete_suffix_missing", default="Suffix is required for 'Keep Suffix' rule."))

        files_to_delete = [] # List to store paths of files identified for deletion

        # Iterate through each set of duplicates
        for sha1, files_in_set in duplicate_sets.items():
            # Basic sanity check, should have at least 2 files
            if len(files_in_set) < 2: continue

            keep_file = None # The FileInfo dict of the file to keep in this set
            error_in_set = None # Store any error encountered while processing this set

            try:
                # --- Apply Rule Logic ---
                if rule == RULE_KEEP_SHORTEST:
                    # Find file with the minimum path length
                    keep_file = min(files_in_set, key=lambda f: len(f.get('path', ''))) # Use get for safety
                elif rule == RULE_KEEP_LONGEST:
                    # Find file with the maximum path length
                    keep_file = max(files_in_set, key=lambda f: len(f.get('path', '')))
                elif rule == RULE_KEEP_OLDEST:
                    # Filter out files with invalid/missing dates first
                    valid_files = [f for f in files_in_set if isinstance(f.get('modified'), datetime)]
                    if not valid_files:
                        error_in_set = f"Skipping set {sha1[:8]}... for rule '{rule}' - no valid dates found."
                        continue # Skip to next set
                    # Find file with the minimum (oldest) datetime
                    keep_file = min(valid_files, key=lambda f: f['modified'])
                elif rule == RULE_KEEP_NEWEST:
                    # Filter out files with invalid/missing dates first
                    valid_files = [f for f in files_in_set if isinstance(f.get('modified'), datetime)]
                    if not valid_files:
                        error_in_set = f"Skipping set {sha1[:8]}... for rule '{rule}' - no valid dates found."
                        continue
                    # Find file with the maximum (newest) datetime
                    keep_file = max(valid_files, key=lambda f: f['modified'])
                elif rule == RULE_KEEP_SUFFIX:
                    # Find files ending with the specified suffix (case-insensitive)
                    suffix_lower = suffix_value.lower()
                    matching_files = [f for f in files_in_set if f.get('path', '').lower().endswith(suffix_lower)]

                    if len(matching_files) == 1:
                        keep_file = matching_files[0] # Keep the single match
                    elif len(matching_files) > 1:
                        # Tie-breaker: Multiple files match suffix. Keep the shortest path among them.
                        error_in_set = f"Warning: Multiple files in set {sha1[:8]}... match suffix '{suffix_value}'. Keeping shortest path among them."
                        keep_file = min(matching_files, key=lambda f: len(f.get('path', '')))
                    else: # len(matching_files) == 0
                        # Tie-breaker: No files match suffix. Keep the overall shortest path file in the set.
                        error_in_set = f"Warning: No files in set {sha1[:8]}... match suffix '{suffix_value}'. Keeping overall shortest path file."
                        keep_file = min(files_in_set, key=lambda f: len(f.get('path', '')))
                else:
                    # Should not happen if rule is validated earlier
                    error_in_set = f"Warning: Unknown rule '{rule}' encountered for set {sha1[:8]}.... Skipping deletion for this set."
                    continue # Skip to next set

                # --- Add files to delete list ---
                # If a keep_file was successfully determined for this set...
                if keep_file and keep_file.get('path'):
                    keep_path = keep_file['path']
                    # Add all other files from the set to the deletion list
                    for f_info in files_in_set:
                        path = f_info.get('path')
                        if path and path != keep_path:
                            files_to_delete.append(path)
                else:
                    # This case might happen if a rule fails to find a unique keep file
                    # (e.g., rule logic error, or all files have invalid data for the rule)
                    error_in_set = error_in_set or f"Warning: Could not determine a file to keep for set {sha1[:8]}... with rule '{rule}'. No files from this set will be deleted."

            except Exception as e:
                 # Catch potential errors during comparison (e.g., unexpected data in FileInfo)
                 error_in_set = f"Error processing rule '{rule}' for set {sha1[:8]}...: {e}. Skipping set."
                 # Log traceback for debugging rule logic errors
                 print(f"Traceback for error processing set {sha1[:8]}...:")
                 traceback.print_exc(limit=2)
                 continue # Skip to the next set on error

            finally:
                # Log any warnings or errors encountered for this specific set
                if error_in_set:
                     self.log(error_in_set) # Log warnings/errors specific to set processing

        return files_to_delete


    def start_delete_by_rule_thread(self):
        """ Handles the 'Delete Files by Rule' button click. """
        # --- Pre-checks ---
        selected_rule = self.deletion_rule_var.get()
        # Get a display name for the rule for messages
        rule_name_for_msg = self._(f"rule_{selected_rule}", default=selected_rule) if selected_rule else "N/A"

        # 1. Check if a rule is selected
        if not selected_rule:
             messagebox.showerror(self._("error_input_title"), self._("delete_no_rule_selected"))
             return
        # 2. Check if suffix is provided for the suffix rule
        if selected_rule == RULE_KEEP_SUFFIX and not self.suffix_entry_var.get():
             messagebox.showerror(self._("error_input_title"), self._("delete_suffix_missing"))
             return
        # 3. Check if the cache of files to delete (determined by rule) is populated
        if not self.files_to_delete_cache:
             # Inform user that based on the rule, no files need deletion
             messagebox.showinfo(self._("delete_by_rule_button"), self._("delete_rule_no_files"))
             self.log_message(self._("delete_rule_no_files") + f" (Rule: {rule_name_for_msg})")
             return

        # --- Confirmation Dialog ---
        # Ask user to confirm deletion, showing the rule and number of files
        num_files = len(self.files_to_delete_cache)
        confirm_msg = self._("delete_confirm_msg", rule_name=rule_name_for_msg)
        confirm_msg += f"\n\n({num_files} files will be deleted)" # Add count to confirmation

        confirm = messagebox.askyesno(
            title=self._("delete_confirm_title"),
            message=confirm_msg,
            icon='warning' # Use warning icon
        )

        if not confirm:
            self.log_message(self._("delete_cancelled"))
            return # User cancelled

        # --- Start Deletion Process ---
        self.log_message(self._("delete_starting", rule_name=rule_name_for_msg))
        self.set_ui_state("deleting") # Disable UI during deletion

        # Use the cached list of file paths determined when the rule was applied
        files_to_delete_list = list(self.files_to_delete_cache) # Create a copy to pass to thread

        # Start the worker thread for deletion
        thread = threading.Thread(target=self._delete_worker, args=(files_to_delete_list, rule_name_for_msg), daemon=True)
        thread.start()


    def _delete_worker(self, files_to_delete, rule_name_for_log):
        """ Worker thread for deleting files based on the provided list. """
        # Check connection within thread
        if not self.finder or not self.finder.fs:
            self.log_message("Error: Connection lost before Deletion could execute.")
            # Schedule UI reset
            if self.master.winfo_exists(): self.master.after(0, self.set_ui_state, 'normal')
            return

        deleted_count = 0
        total_attempted = len(files_to_delete)

        try:
            # Check if there are actually files to delete
            if not files_to_delete:
                self.log_message(self._("delete_rule_no_files") + f" (Worker check; Rule: {rule_name_for_log})")
            else:
                # Call the finder's delete method
                deleted_count, total_attempted = self.finder.delete_files(files_to_delete)
                # Log message for completion/errors is handled within finder.delete_files

            # --- Schedule GUI update after deletion attempt ---
            # Decide what to do after deletion: clear list, refresh, etc.
            # Clearing the list is simplest: assumes user wants to start fresh after deleting.
            if self.master.winfo_exists():
                # Always clear results after a delete attempt, regardless of success?
                # Or only if delete_count > 0? Let's clear always for simplicity.
                self.master.after(0, self.clear_results) # Schedule clear_results to run in main thread


        except Exception as e:
            # Catch unexpected errors during the finder.delete_files call itself
            # Error during individual file delete is handled inside finder.delete_files
            err_msg = self._("delete_error_during", error=e) # General error during the process
            self.log_message(err_msg)
            self.log_message(traceback.format_exc()) # Log details
            if self.master.winfo_exists():
                 # Show error message box
                 self.master.after(0, messagebox.showerror, self._("error_title"), err_msg)
                 # Do we clear results on error? Maybe not, allow user to retry?
                 # Let's NOT clear results here, just reset state. User can clear manually.

        finally:
            # --- ALWAYS Re-enable UI in the main thread ---
            # Ensure UI is re-enabled regardless of success or failure
            if self.master.winfo_exists():
                # Schedule state reset back to normal
                 self.master.after(0, self.set_ui_state, 'normal')


    def save_duplicates_report(self):
        """ Saves the report of FOUND duplicate file sets. """
        # Check if duplicates data exists
        if not self.duplicate_sets:
            # Use translated message key for consistency
            messagebox.showinfo(self._("save_report_no_data", default="No Duplicates Found"),
                                self._("save_report_no_data", default="No duplicate sets found to save."))
            self.log_message(self._("save_report_no_data"))
            return

        # Generate default filename with timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        initial_filename = f"duplicates_report_{timestamp}.txt"

        # Ask user for save location and filename
        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
            # Use the button's text key for the dialog title if appropriate
            title=self._("save_list_button", default="Save Found Duplicates Report As..."),
            initialfile=initial_filename
        )

        # If user provided a path (didn't cancel)
        if file_path:
            # Call the finder's report writing method
            # Pass the currently stored duplicate sets
            success = self.finder.write_duplicates_report(self.duplicate_sets, file_path)
            # Logging is handled within write_duplicates_report
            # Optionally show a success message box here if desired
            # if success and self.master.winfo_exists():
            #     messagebox.showinfo("Report Saved", self._("save_report_saved", file=file_path))
        else:
             # User cancelled the save dialog
             self.log_message("Save report operation cancelled by user.")


    def start_test_connection_thread(self):
        self.log_message(self._("status_connecting"))
        # Set UI to a busy state during connection test
        self.set_ui_state("testing_connection") # Use a specific state or just disable relevant parts
        thread = threading.Thread(target=self._test_connection_worker, daemon=True)
        thread.start()


    def _test_connection_worker(self):
        # Get necessary config values from GUI
        address = self.string_vars["address"].get()
        account = self.string_vars["account"].get()
        passwd = self.string_vars["password"].get()
        # Need scan path and mount point for finder.set_config context now
        mount_point = self.string_vars["mount_point"].get()
        scan_path = self.string_vars["scan_path"].get() # Get scan path as well
        connected = False

        try:
            # --- Input Validation ---
            # Check required fields for connection test
            if not all([address, account, mount_point]): # Password might be optional
                 error_msg = self._("error_input_missing_conn", default="Address, Account, and Mount Point are required for connection test.")
                 self.log_message(error_msg)
                 # Show error message box in main thread
                 if self.master.winfo_exists(): self.master.after(0, messagebox.showerror, (self._("error_input_title", default="Input Error"), error_msg))
                 return # Exit worker early

            # --- Path Character Validation ---
            # Check relevant paths for suspicious characters
            paths_to_check_conn = {"address": address, "mount_point": mount_point}
            # Also check scan_path as it's passed to set_config? Or only if calculating path?
            # Let's check all relevant paths used by set_config for consistency.
            paths_to_check_all = {"address": address, "mount_point": mount_point, "scan_path": scan_path}
            if not self._check_path_chars(paths_to_check_all):
                # Error already logged and shown by _check_path_chars
                return # Exit worker early

            # --- Attempt Connection ---
            # Call finder's set_config, which handles connection attempt and logging
            connected = self.finder.set_config(address, account, passwd, scan_path, mount_point, self.log_message)

            # --- Show Result Message Box ---
            # Schedule the message box to run in the main GUI thread
            if self.master.winfo_exists():
                 if connected:
                     # Success message
                     self.master.after(0, messagebox.showinfo, (self._("conn_test_success_title", default="Connection Test Successful"), self._("conn_test_success_msg", default="Successfully connected.")))
                 else:
                     # Failure message (error details already logged by set_config)
                     self.master.after(0, messagebox.showwarning, (self._("conn_test_fail_title", default="Connection Test Failed"), self._("conn_test_fail_msg", default="Failed to connect. Check log.")))

        except Exception as e:
            # Catch unexpected errors during the worker setup or validation steps
            error_msg = f'{self._("error_unexpected", default="Unexpected error")} during connection test worker: {e}'
            self.log_message(error_msg)
            self.log_message(traceback.format_exc()) # Log details
            if self.master.winfo_exists():
                self.master.after(0, messagebox.showerror, (self._("error_title", default="Error"), error_msg))
        finally:
            # --- ALWAYS Re-enable UI ---
            # Ensure UI state is reset back to normal in the main thread
            if self.master.winfo_exists():
                self.master.after(0, self.set_ui_state, 'normal')


    def show_cloud_file_types(self):
        # 1. Check Matplotlib dependency
        if not MATPLOTLIB_AVAILABLE:
            messagebox.showwarning(self._("chart_error_title"), self._("chart_error_no_matplotlib"))
            return
        # 2. Check connection status
        if not self.finder or not self.finder.fs:
            messagebox.showwarning(self._("chart_error_title"), self._("chart_error_no_connection"))
            return
        # 3. Check required inputs for chart path calculation
        scan_path_raw = self.string_vars["scan_path"].get()
        mount_point_raw = self.string_vars["mount_point"].get()
        if not scan_path_raw or not mount_point_raw:
             messagebox.showwarning(self._("error_input_title"), self._("error_input_missing_chart"))
             return
        # 4. Check path characters
        paths_to_check = {"scan_path": scan_path_raw, "mount_point": mount_point_raw}
        if not self._check_path_chars(paths_to_check): return

        # 5. Set UI state and start worker thread
        self.set_ui_state("charting") # Disable UI during chart generation
        thread = threading.Thread(target=self._show_cloud_file_types_worker,
                                  args=(scan_path_raw, mount_point_raw), daemon=True)
        thread.start()


    def _show_cloud_file_types_worker(self, scan_path_raw, mount_point_raw):
        fs_dir_path = None
        file_counts = collections.Counter()
        total_files = 0
        scan_error = None

        try:
             # --- Setup & Validation ---
             # Re-calculate path within worker to be self-contained
             # Ensure finder exists before calculating path
             if not self.finder:
                  self.log_message("Error: Finder object not initialized for charting.")
                  return
             fs_dir_path = self.finder.calculate_fs_path(scan_path_raw, mount_point_raw)
             if fs_dir_path is None:
                 # Error logged by calculate_fs_path
                 return # Exit worker

             # Double check connection within worker
             if not self.finder.fs:
                 self.log_message(self._("chart_error_no_connection"))
                 return

             # --- Scan Files ---
             self.log_message(self._("chart_status_scanning_cloud", path=fs_dir_path))
             try:
                 # Iterate through files using walk_path
                 for _, _, filenames in self.finder.fs.walk_path(fs_dir_path):
                     for filename_obj in filenames:
                         filename = str(filename_obj) # Ensure string
                         total_files += 1
                         _, ext = os.path.splitext(filename)
                         # Use translation for no extension label, fallback if needed
                         ext_label = ext.lower() if ext else self._("chart_label_no_extension", default="[No Ext]")
                         file_counts[ext_label] += 1
             except Exception as e:
                 # Catch errors during the walk (network, permissions, etc.)
                 scan_error = e
                 error_msg = self._("chart_error_cloud_scan", path=fs_dir_path, error=e)
                 self.log_message(error_msg)
                 self.log_message(traceback.format_exc()) # Log details
                 # Schedule error message box in main thread
                 if self.master.winfo_exists():
                      self.master.after(0, messagebox.showerror, (self._("chart_error_title"), error_msg))
                 # Do not proceed to chart generation if scan failed

             # --- Schedule GUI Update (Chart Creation or Info Message) ---
             # Define the function to run in the main thread
             def update_gui_after_scan():
                 if not self.master.winfo_exists(): return # Window closed
                 if scan_error: return # Error already handled, don't proceed

                 # Check if any files were found
                 if not file_counts:
                     no_files_msg = self._("chart_status_no_files_found", path=fs_dir_path)
                     self.log_message(no_files_msg)
                     # Show info message box
                     if self.master.winfo_exists(): messagebox.showinfo(self._("chart_info_title"), no_files_msg)
                     return # Nothing to chart

                 # Log generation message and create the chart window
                 self.log_message(self._("chart_status_generating", count=len(file_counts), total=total_files))
                 self._create_pie_chart_window(file_counts, fs_dir_path)

             # Schedule the update function using 'after' for thread safety
             if self.master.winfo_exists():
                 self.master.after(0, update_gui_after_scan)

        except Exception as e:
             # Catch errors during worker setup (e.g., path calculation before scan)
             err_msg = f"Error during chart worker setup for path '{scan_path_raw}': {e}"
             self.log_message(err_msg)
             self.log_message(traceback.format_exc())
             # No message box here, rely on logging. Finally block resets state.
        finally:
             # --- ALWAYS Re-enable UI in the main thread ---
             if self.master.winfo_exists():
                 self.master.after(0, self.set_ui_state, 'normal')


    def _create_pie_chart_window(self, counts, display_path):
        # (Method logic remains the same as previous version)
         if not MATPLOTLIB_AVAILABLE:
             self.log_message("Error: Matplotlib became unavailable before chart creation.")
             if self.master.winfo_exists(): # Check window exists before showing messagebox
                 messagebox.showerror(self._("chart_error_title"), self._("chart_error_no_matplotlib"))
             return
         chart_window = None # Initialize

         try:
             # Check/set Matplotlib font settings (optional, can be removed if problematic)
             try:
                 current_pref = matplotlib.rcParams['font.sans-serif']
                 matplotlib.rcParams['axes.unicode_minus'] = False # Ensure minus sign displays correctly
             except Exception as font_e:
                 self.log_message(f"Warning: Issue re-checking Matplotlib font settings: {font_e}")

             # --- Data Preparation ---
             top_n = 15 # Show top N categories + 'Others'
             total_count = sum(counts.values())
             sorted_counts = counts.most_common()

             labels = []
             sizes = []
             others_label = self._("chart_label_others", default="Others")
             others_count = 0
             others_sources = [] # Keep track of extensions grouped into 'Others'

             if len(sorted_counts) > top_n:
                 # Take top N items
                 top_items = sorted_counts[:top_n]
                 labels = [item[0] for item in top_items]
                 sizes = [item[1] for item in top_items]
                 # Group remaining items into 'Others'
                 other_items = sorted_counts[top_n:]
                 others_count = sum(item[1] for item in other_items)
                 others_sources = [item[0] for item in other_items]
                 # Add 'Others' category if it has items
                 if others_count > 0:
                     labels.append(others_label)
                     sizes.append(others_count)
             else:
                 # Fewer than top_n items, show all individually
                 labels = [item[0] for item in sorted_counts]
                 sizes = [item[1] for item in sorted_counts]

             # Log if items were grouped
             if others_count > 0:
                 self.log_message(f"Chart Note: Grouped {len(others_sources)} smaller categories ({others_count} files) into '{others_label}'.")

             # --- Create Chart Window ---
             chart_window = Toplevel(self.master)
             chart_window.title(self._("chart_window_title", path=display_path))
             chart_window.geometry("850x650") # Adjust size as needed

             # --- Create Matplotlib Figure & Axes ---
             fig = Figure(figsize=(8, 6), dpi=100) # Adjust figsize for content
             ax = fig.add_subplot(111)

             # --- Generate Pie Chart ---
             # Explode slices? (Optional visual enhancement) - e.g., explode = [0.05] * len(labels)
             wedges, texts, autotexts = ax.pie(sizes, autopct='%1.1f%%', startangle=90, pctdistance=0.85) # textprops={'fontsize': 8}
             ax.axis('equal') # Ensures pie is circular

             # --- Create Legend ---
             legend_labels = [f'{label} ({sizes[i]})' for i, label in enumerate(labels)]
             ax.legend(wedges, legend_labels,
                       title=self._("chart_legend_title", default="File Extensions"),
                       loc="center left", bbox_to_anchor=(1.05, 0, 0.5, 1), # Place outside right edge
                       fontsize='small', frameon=False) # No frame around legend

             # --- Style Percentage Labels ---
             for autotext in autotexts:
                 autotext.set_color('white')
                 autotext.set_size(8) # Adjust size as needed
                 autotext.set_weight('bold')
                 # Add background box for contrast
                 autotext.set_bbox(dict(facecolor='black', alpha=0.6, pad=1, edgecolor='none'))

             # --- Adjust Layout ---
             try:
                 # Adjust layout to prevent labels/legend overlapping plot area
                 fig.tight_layout(rect=[0, 0, 0.80, 1]) # Leave space on right (for legend)
             except Exception as layout_err:
                 print(f"Warning: tight_layout issue: {layout_err}. Legend/labels might overlap.")
                 self.log_message(f"Warning: Chart layout issue: {layout_err}")

             # --- Embed in Tkinter ---
             canvas = FigureCanvasTkAgg(fig, master=chart_window)
             canvas_widget = canvas.get_tk_widget()

             # Add Navigation Toolbar
             toolbar = NavigationToolbar2Tk(canvas, chart_window)
             toolbar.update() # Needed step

             # Pack canvas widget
             canvas_widget.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

             # Draw the canvas
             canvas.draw()

         except Exception as e:
             # Catch errors during chart creation/display
             error_msg = f"Error creating chart window: {e}"
             self.log_message(error_msg)
             self.log_message(traceback.format_exc()) # Log details
             if self.master.winfo_exists(): messagebox.showerror(self._("chart_error_title"), error_msg)
             # Destroy the Toplevel window if it exists but failed
             if chart_window and chart_window.winfo_exists():
                  chart_window.destroy()



# --- Main Execution ---
if __name__ == "__main__":
    # --- Setup Root Window ---
    root = tk.Tk()
    root.minsize(950, 750) # Adjust min size based on content

    # --- Create and Run App ---
    try:
        app = DuplicateFinderApp(root)
        root.mainloop()
    except Exception as main_e:
        # --- Fatal Error Handling ---
        print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
        print("FATAL ERROR DURING APPLICATION STARTUP / MAIN LOOP:")
        print(traceback.format_exc())
        print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
        # Attempt to show GUI error message as fallback
        try:
            # Create a temporary minimal root if mainloop failed early
            if 'root' not in locals() or not root.winfo_exists():
                 root_err = tk.Tk()
                 root_err.withdraw() # Hide the empty window
                 messagebox.showerror("Fatal Error", f"A critical error occurred:\n\n{main_e}\n\nSee console log for details.", master=root_err)
                 root_err.destroy()
            else:
                 messagebox.showerror("Fatal Error", f"A critical error occurred:\n\n{main_e}\n\nSee console log for details.", master=root) # Use existing root if possible
        except Exception as mb_err:
            # If even Tkinter fails for the message box
            print(f"Could not display fatal error in GUI: {mb_err}")