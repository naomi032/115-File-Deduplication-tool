import os
import configparser
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog, Menu, Toplevel, Canvas # Import Canvas
import threading
import time
from datetime import datetime, timezone # Added timezone
from collections import defaultdict, Counter
import json
import traceback
import collections
import math # For size conversion
import sys # To get base path for PyInstaller

# --- Pillow Check ---
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


# --- Matplotlib Check ---
try:
    import matplotlib
    matplotlib.use('TkAgg') # Use Tkinter backend for embedding
    from matplotlib.figure import Figure
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
    MATPLOTLIB_AVAILABLE = True
    try:
        # Attempt to set preferred CJK fonts
        preferred_fonts = ['SimHei', 'Microsoft YaHei', 'MS Gothic', 'Malgun Gothic', 'Arial Unicode MS', 'sans-serif']
        matplotlib.rcParams['font.sans-serif'] = preferred_fonts
        matplotlib.rcParams['axes.unicode_minus'] = False # Ensure minus sign displays correctly
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

# --- clouddrive Check ---
try:
    from clouddrive import CloudDriveClient, CloudDriveFileSystem
except ImportError:
    print("ERROR: The 'clouddrive' library is not installed. Please install it using: pip install clouddrive")
    try: # Attempt to show GUI error even if library is missing
        root_tk_err = tk.Tk()
        root_tk_err.withdraw()
        messagebox.showerror("Missing Library", "The 'clouddrive' library is not installed.\nPlease install it using: pip install clouddrive", master=root_tk_err)
        root_tk_err.destroy()
    except Exception:
        pass
    exit()


# --- Function to get resource path for PyInstaller ---
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        # If not running as a PyInstaller bundle, use the script's directory
        base_path = os.path.abspath(os.path.dirname(__file__))

    return os.path.join(base_path, relative_path)

# --- Constants ---
VIDEO_EXTENSIONS = {
    ".mkv", ".iso", ".ts", ".mp4", ".avi", ".rmvb",
    ".wmv", ".m2ts", ".mpg", ".flv", ".rm", ".mov",
}
CONFIG_FILE = resource_path("config.ini")
LANG_PREF_FILE = resource_path("lang_pref.json")
BACKGROUND_IMAGE_FILE = resource_path("background.png") # Path for background
ICON_FILE = resource_path("app_icon.ico") # Path for icon

DATE_FORMAT = "%Y-%m-%d %H:%M:%S" # For display
DEFAULT_LANG = "en"
# Rule constants
RULE_KEEP_SHORTEST = "shortest"
RULE_KEEP_LONGEST = "longest"
RULE_KEEP_OLDEST = "oldest"
RULE_KEEP_NEWEST = "newest"
RULE_KEEP_SUFFIX = "suffix"
# --- Debounce Delay (milliseconds) for background resize ---
RESIZE_DEBOUNCE_DELAY = 200 # Adjust as needed (e.g., 150-300ms)


# --- Translations ---
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

        # Main Actions
        "find_button": "Find Duplicates",
        "find_starting": "Starting duplicate file scan...",
        "find_complete_found": "Scan complete. Found {count} sets of duplicate video files.",
        "find_complete_none": "Scan complete. No duplicate video files found based on SHA1 hash.",
        "find_error_during": "Error during duplicate scan: {error}",
        "delete_by_rule_button": "Delete Files by Rule",
        "delete_confirm_title": "Confirm Deletion by Rule",
        "delete_confirm_msg": "This will delete files based on the selected rule: '{rule_name}'.\n\nFiles marked for deletion in the list below will be removed permanently from the cloud drive.\n\nTHIS ACTION CANNOT BE UNDONE.\n\nAre you sure you want to proceed?",
        "delete_cancelled": "Deletion operation cancelled by user.",
        "delete_starting": "Starting deletion based on rule '{rule_name}'...",
        "delete_determining": "Determining files to delete based on rule '{rule_name}'...",
        "delete_no_rule_selected": "Error: Please select a deletion rule first.",
        "delete_suffix_missing": "Error: Please enter the suffix to keep when using the 'Keep Suffix' rule.",
        "delete_rule_no_files": "No files matched the deletion criteria for the selected rule.",
        "delete_finished": "Deletion complete. Deleted {deleted_count} of {total_marked} files.",
        "delete_error_during": "Error during deletion process: {error}",
        "delete_error_file": "Error deleting {path}: {error}",

        # Deletion Rules Section
        "rules_title": "Deletion Rules (Select ONE to Keep):",
        "rule_shortest_path": "Keep Shortest Path",
        "rule_longest_path": "Keep Longest Path",
        "rule_oldest": "Keep Oldest (Modified Date)",
        "rule_newest": "Keep Newest (Modified Date)",
        "rule_keep_suffix": "Keep files ending with:",
        "rule_suffix_entry_label": "Suffix:",

        # Chart & Save Report
        "show_chart_button": "Show Cloud File Types",
        "show_chart_button_disabled": "Chart (Needs Connection & Matplotlib)",
        "save_list_button": "Save Found Duplicates Report...",
        "save_report_saved": "Duplicates report saved to {file}",
        "save_report_error": "Error saving duplicates report to {file}: {error}",
        "save_report_no_data": "No duplicate sets found to save.",
        "save_report_header": "Duplicate Video File Sets Found (Based on SHA1):",
        "save_report_set_header": "Set {index} (SHA1: {sha1}) - {count} files",
        "save_report_file_label": "  - File:",
        "save_report_details_label": "    (Path: {path}, Modified: {modified}, Size: {size_mb:.2f} MB)",

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
        "status_scan_progress": "Scanned {count} total items... Found {video_count} videos so far.",
        "status_scan_finished": "Scan finished. Checked {count} files.",
        "status_populating_tree": "Populating results list with {count} duplicate sets...",
        "status_tree_populated": "Results list populated.",
        "status_applying_rule": "Applying rule '{rule_name}' to {count} sets...",
        "status_rule_applied": "Rule applied. Identified {delete_count} files to delete.",
        "status_clearing_tree": "Clearing results list...",

        # Treeview Columns
        "tree_rule_action_col": "Action", # (Keep/Delete)
        "tree_path_col": "File Path",
        "tree_modified_col": "Date Modified",
        "tree_size_col": "Size (MB)",
        "tree_set_col": "Duplicate Set #",
        "tree_set_col_value": "{index}",
        "tree_action_keep": "Keep",
        "tree_action_delete": "Delete",

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
        "error_no_duplicates_found": "No duplicates were found or displayed. Cannot apply deletion rule.",
        "error_parse_date": "Error parsing date for {path}: {error}",
        "error_not_connected": "Error: Not connected to CloudDrive2. Cannot perform action.",
        "error_path_calc_failed": "Error: Could not determine a valid cloud path from Root Path and Mount Point. Check inputs.",
        "warning_path_mismatch": "Warning: Could not determine a valid cloud path based on 'Root Path to Scan' ('{scan}') and 'Mount Point' ('{mount}'). Please check inputs.",
        "path_warning_title": "Path Input Warning",
        "path_warning_suspicious_chars": "Suspicious character(s) detected in input paths!\nThis often happens from copy-pasting.\nPlease DELETE and MANUALLY RETYPE the paths in the GUI.",
        "error_img_load": "Error loading background image '{path}': {error}",
        "error_icon_load": "Error loading icon '{path}': {error}",

        # Menu & Connection Test
        "menu_language": "Language",
        "menu_english": "English",
        "menu_chinese": "中文",
        "conn_test_success_title": "Connection Test Successful",
        "conn_test_success_msg": "Successfully connected to CloudDrive2.",
        "conn_test_fail_title": "Connection Test Failed",
        "conn_test_fail_msg": "Failed to connect. Check log for details.",

        # Charting
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

        # 主要操作
        "find_button": "查找重复项",
        "find_starting": "开始扫描重复文件...",
        "find_complete_found": "扫描完成。找到 {count} 组重复的视频文件。",
        "find_complete_none": "扫描完成。未根据 SHA1 哈希找到重复的视频文件。",
        "find_error_during": "扫描重复项期间出错: {error}",
        "delete_by_rule_button": "按规则删除文件",
        "delete_confirm_title": "确认按规则删除",
        "delete_confirm_msg": "此操作将根据选定规则删除文件: '{rule_name}'。\n\n下面列表中标记为删除的文件将从云盘中永久移除。\n\n此操作无法撤销。\n\n您确定要继续吗？",
        "delete_cancelled": "用户取消了删除操作。",
        "delete_starting": "开始根据规则 '{rule_name}' 删除文件...",
        "delete_determining": "正在根据规则 '{rule_name}' 确定要删除的文件...",
        "delete_no_rule_selected": "错误：请先选择一个删除规则。",
        "delete_suffix_missing": "错误：使用“保留后缀”规则时，请输入要保留的后缀名。",
        "delete_rule_no_files": "没有文件符合所选规则的删除条件。",
        "delete_finished": "删除完成。共删除了 {total_marked} 个标记文件中的 {deleted_count} 个。",
        "delete_error_during": "删除过程中出错: {error}",
        "delete_error_file": "删除 {path} 时出错: {error}",

        # 删除规则部分
        "rules_title": "删除规则 (选择一项保留):",
        "rule_shortest_path": "保留最短路径",
        "rule_longest_path": "保留最长路径",
        "rule_oldest": "保留最旧 (修改日期)",
        "rule_newest": "保留最新 (修改日期)",
        "rule_keep_suffix": "保留以此结尾的文件:",
        "rule_suffix_entry_label": "后缀:",

        # 图表和保存报告
        "show_chart_button": "显示云盘文件类型",
        "show_chart_button_disabled": "图表 (需连接 & Matplotlib)",
        "save_list_button": "保存找到的重复项报告...",
        "save_report_saved": "重复项报告已保存至 {file}",
        "save_report_error": "保存重复项报告至 {file} 时出错: {error}",
        "save_report_no_data": "未找到可保存的重复文件集。",
        "save_report_header": "找到的重复视频文件集 (基于 SHA1):",
        "save_report_set_header": "集合 {index} (SHA1: {sha1}) - {count} 个文件",
        "save_report_file_label": "  - 文件:",
        "save_report_details_label": "    (路径: {path}, 修改日期: {modified}, 大小: {size_mb:.2f} MB)",

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
        "status_scan_progress": "已扫描 {count} 个项目... 目前找到 {video_count} 个视频。",
        "status_scan_finished": "扫描完成。共检查 {count} 个文件。",
        "status_populating_tree": "正在使用 {count} 个重复文件集填充结果列表...",
        "status_tree_populated": "结果列表已填充。",
        "status_applying_rule": "正在对 {count} 个集合应用规则 '{rule_name}'...",
        "status_rule_applied": "规则已应用。识别出 {delete_count} 个待删除文件。",
        "status_clearing_tree": "正在清空结果列表...",

        # Treeview 列
        "tree_rule_action_col": "操作", # (保留/删除)
        "tree_path_col": "文件路径",
        "tree_modified_col": "修改日期",
        "tree_size_col": "大小 (MB)",
        "tree_set_col": "重复集合 #",
        "tree_set_col_value": "{index}",
        "tree_action_keep": "保留",
        "tree_action_delete": "删除",

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
        "error_no_duplicates_found": "未找到或显示重复项。无法应用删除规则。",
        "error_parse_date": "解析 {path} 的日期时出错: {error}",
        "error_not_connected": "错误：未连接到 CloudDrive2。无法执行操作。",
        "error_path_calc_failed": "错误：无法根据扫描根路径和挂载点确定有效的云端路径。请检查输入。",
        "warning_path_mismatch": "警告：无法根据“要扫描的根路径” ('{scan}') 和“挂载点” ('{mount}') 确定有效的云端路径。请检查输入。",
        "path_warning_title": "路径输入警告",
        "path_warning_suspicious_chars": "在输入路径中检测到可疑字符！\n这通常是复制粘贴造成的。\n请在图形界面中删除并手动重新输入路径。",
        "error_img_load": "加载背景图片 '{path}' 时出错: {error}",
        "error_icon_load": "加载图标 '{path}' 时出错: {error}",

        # 菜单与连接测试
        "menu_language": "语言",
        "menu_english": "English",
        "menu_chinese": "中文",
        "conn_test_success_title": "连接测试成功",
        "conn_test_success_msg": "已成功连接到 CloudDrive2。",
        "conn_test_fail_title": "连接测试失败",
        "conn_test_fail_msg": "连接失败。请检查日志获取详细信息。",

        # 图表
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


# --- Helper Functions ---
def _validate_path_chars(path_str):
    """Checks a single path string for suspicious characters."""
    suspicious_codes = []
    if not isinstance(path_str, str): return suspicious_codes
    # Common invisible or control characters that cause issues
    KNOWN_INVISIBLE_CODES = {
        0x200B, # Zero Width Space
        0x200C, # Zero Width Non-Joiner
        0x200D, # Zero Width Joiner
        0x200E, # Left-to-Right Mark
        0x200F, # Right-to-Left Mark
        0xFEFF, # Byte Order Mark / Zero Width No-Break Space
        0x00A0, # No-Break Space (often problematic in paths)
        # Add others if needed
    }
    for i, char in enumerate(path_str):
        char_code = ord(char)
        is_suspicious = False
        reason = ""
        # C0 Controls (ASCII 0-31) and DEL (127)
        if 0 <= char_code <= 31 or char_code == 127:
            is_suspicious = True
            reason = "Control Char (C0/DEL)"
        # C1 Controls (often appear from bad encoding/copy-paste: 128-159)
        elif 128 <= char_code <= 159:
            is_suspicious = True
            reason = "Control Char (C1)"
        # Specific known problematic Unicode chars
        elif char_code in KNOWN_INVISIBLE_CODES:
            is_suspicious = True
            reason = "Known Invisible/Problematic Char"
        # Check for characters typically invalid in Windows paths (if targetting Windows primarily)
        # elif char in r'<>:"/\|?*':
        #     is_suspicious = True; reason="Invalid Windows Path Char"

        if is_suspicious:
            suspicious_codes.append(f"U+{char_code:04X} ({reason}) at pos {i}")
            # Optional: Stop after first detection for performance if needed
            # break
    return suspicious_codes

def _build_full_path(parent_path, item_name):
    """Helper to correctly join cloud paths using forward slashes."""
    # Normalize slashes and remove trailing/leading ones inappropriately
    parent_path_norm = parent_path.replace('\\', '/').rstrip('/')
    item_name_norm = item_name.replace('\\', '/').lstrip('/')
    # Handle root case
    if not parent_path_norm or parent_path_norm == '/':
        return '/' + item_name_norm
    else:
        return parent_path_norm + '/' + item_name_norm

def _parse_datetime(date_string):
    """Parses common datetime string formats into timezone-aware datetime objects."""
    if not date_string: return None

    formats_to_try = []
    original_date_string = date_string # Keep original for reprocessing if needed

    # 1. ISO 8601 format with Z (Zulu time = UTC)
    if 'Z' in date_string:
        try:
            dt = datetime.fromisoformat(date_string.replace('Z', '+00:00'))
            return dt if dt.tzinfo else dt.replace(tzinfo=timezone.utc) # Ensure TZ info
        except ValueError: pass

    # 2. ISO 8601 format with explicit offset (+HH:MM, +HHMM, -HH:MM, -HHMM)
    # Requires Python 3.7+ for handling ':' in offset with %z
    if sys.version_info >= (3, 7) and ('+' in date_string[-6:] or '-' in date_string[-6:]):
        # Try standard ISO format directly (handles optional microseconds and offset)
        try:
            dt = datetime.fromisoformat(date_string)
            return dt if dt.tzinfo else dt.replace(tzinfo=timezone.utc) # Ensure TZ info
        except ValueError: pass

        # Explicit formats if fromisoformat fails (less common these days)
        formats_with_colon = [
            '%Y-%m-%dT%H:%M:%S.%f%z', # With microseconds and colon offset
            '%Y-%m-%dT%H:%M:%S%z',    # Without microseconds and colon offset
        ]
        formats_without_colon = [
             '%Y-%m-%dT%H:%M:%S.%f%z', # With microseconds and no colon (input adjusted)
             '%Y-%m-%dT%H:%M:%S%z',    # Without microseconds and no colon (input adjusted)
        ]

        # Try formats expecting colon first
        for fmt in formats_with_colon:
            try:
                dt = datetime.strptime(date_string, fmt)
                return dt # strptime with %z should yield aware object
            except ValueError: pass

        # Try formats without colon (adjust input string if colon exists)
        date_string_no_colon = date_string
        if len(date_string) > 6 and date_string[-3] == ':':
            date_string_no_colon = date_string[:-3] + date_string[-2:]

        for fmt in formats_without_colon:
             try:
                 dt = datetime.strptime(date_string_no_colon, fmt)
                 return dt # strptime with %z should yield aware object
             except ValueError: pass

    # 3. Common non-ISO format (assume UTC if no timezone info)
    try:
        dt = datetime.strptime(date_string, '%Y-%m-%d %H:%M:%S')
        # Assume UTC if no timezone was parsed
        return dt.replace(tzinfo=timezone.utc)
    except ValueError: pass

    # Add more formats here if needed, e.g.:
    # try:
    #     dt = datetime.strptime(date_string, '%Y/%m/%d %H:%M:%S')
    #     return dt.replace(tzinfo=timezone.utc)
    # except ValueError: pass

    # --- Fallback if all parsing fails ---
    print(f"Warning: Could not parse date string: {original_date_string} with available formats.")
    return None


# --- DuplicateFileFinder Class ---
class DuplicateFileFinder:
    def __init__(self):
        self.clouddrvie2_address = ""
        self.clouddrive2_account = ""
        self.clouddrive2_passwd = ""
        self._raw_scan_path = ""
        self._raw_mount_point = ""
        self.fs = None
        self.progress_callback = None
        self._ = lambda key, **kwargs: kwargs.get('default', key) # Default translator

    def set_translator(self, translator_func):
        self._ = translator_func

    def log(self, message):
        """ Sends message to the registered progress callback (GUI logger). """
        if self.progress_callback:
            try:
                # Ensure message is a string before calling callback
                message_str = str(message) if message is not None else ""
                self.progress_callback(message_str)
            except Exception as e:
                # Fallback to print if callback fails
                print(f"Error in progress callback: {e}")
                print(f"Original message: {message}")
        else:
            # Fallback to print if no callback is registered
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
        """ Sets configuration and attempts to establish+test connection. """
        self.clouddrvie2_address = clouddrvie2_address
        self.clouddrive2_account = clouddrive2_account
        self.clouddrive2_passwd = clouddrive2_passwd
        self._raw_scan_path = raw_scan_path
        self._raw_mount_point = raw_mount_point
        self.progress_callback = progress_callback
        self.fs = None # Reset filesystem object

        try:
            self.log(self._("status_connecting", default="Attempting connection..."))
            # Initialize the client
            client = CloudDriveClient(
                self.clouddrvie2_address, self.clouddrive2_account, self.clouddrive2_passwd
            )
            # Initialize the filesystem interface
            self.fs = CloudDriveFileSystem(client)

            # Test connectivity by listing root (or another simple operation)
            self.log("Testing connection by attempting to list root directory ('/')...")
            self.fs.ls('/') # Raises exception on failure
            self.log(self._("status_connect_success", default="Connection successful."))
            return True # Connection successful

        except Exception as e:
            # Log detailed error for debugging
            error_msg = self._("error_connect", error=e, default=f"Error connecting: {e}")
            self.log(error_msg)
            self.log(f"Connection Error Details: {traceback.format_exc()}")
            self.fs = None # Ensure fs is None on error
            return False # Connection failed

    def calculate_fs_path(self, scan_path_raw, mount_point_raw):
        """ Calculates the effective cloud filesystem path based on scan root and mount point. """
        # Input validation first
        scan_path_issues = _validate_path_chars(scan_path_raw)
        mount_point_issues = _validate_path_chars(mount_point_raw)
        if scan_path_issues or mount_point_issues:
            all_issues = []
            if scan_path_issues: all_issues.append(f"Scan Path: {', '.join(scan_path_issues)}")
            if mount_point_issues: all_issues.append(f"Mount Point: {', '.join(mount_point_issues)}")
            self.log(f"ERROR: Invalid characters found in path inputs: {'; '.join(all_issues)}")
            # Optionally raise ValueError or return None
            return None

        # Normalize inputs (forward slashes, strip whitespace, remove trailing slash)
        scan_path_norm = scan_path_raw.replace('\\', '/').strip().rstrip('/')
        mount_point_norm = mount_point_raw.replace('\\', '/').strip().rstrip('/')
        fs_dir_path = None # The calculated path within the cloud drive filesystem

        # --- Logic to determine cloud path relative to mount point ---
        # Case 1: Mount point is a Windows drive letter (e.g., "D:")
        if len(mount_point_norm) == 2 and mount_point_norm[1] == ':' and mount_point_norm[0].isalpha():
            mount_point_drive_prefix = mount_point_norm.lower() + '/'
            scan_path_lower = scan_path_norm.lower()
            if scan_path_lower.startswith(mount_point_drive_prefix):
                # Scan path is inside the mounted drive
                relative_part = scan_path_norm[len(mount_point_norm):]
                fs_dir_path = '/' + relative_part.lstrip('/')
            elif scan_path_lower == mount_point_norm.lower():
                 # Scanning the root represented by the drive letter itself
                 fs_dir_path = '/'
        # Case 2: Mount point is empty or root "/" (scan path is relative to cloud root)
        elif not mount_point_norm or mount_point_norm == '/':
            fs_dir_path = '/' + scan_path_norm.lstrip('/')
        # Case 3: Mount point is an absolute path (e.g., "/CloudMount" or "C:/Users/Me/Cloud")
        # Handle both Unix-style and Windows-style absolute paths for mount point
        elif mount_point_norm.startswith('/') or (len(mount_point_norm) > 1 and mount_point_norm[1] == ':'):
             mount_point_prefix = mount_point_norm + '/'
             # Check if scan path starts with the mount point prefix
             if scan_path_norm.startswith(mount_point_prefix):
                 relative_part = scan_path_norm[len(mount_point_norm):]
                 fs_dir_path = '/' + relative_part.lstrip('/') # Path relative to cloud root
             elif scan_path_norm == mount_point_norm:
                 # Scanning the root represented by the mount path itself
                 fs_dir_path = '/'
        # Case 4: Mount point is a relative path (e.g., "MyFolder") - Treat relative to cloud root
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

        # If no logic matched, the paths are likely incompatible
        if fs_dir_path is None:
            warning_msg_tmpl = self._("warning_path_mismatch",
                                      default="Warning: Could not determine a valid cloud path based on 'Root Path to Scan' ('{scan}') and 'Mount Point' ('{mount}'). Please check inputs.")
            self.log(warning_msg_tmpl.format(scan=scan_path_raw, mount=mount_point_raw))
            return None

        # --- Final path normalization ---
        if fs_dir_path:
            # Replace double slashes
            while '//' in fs_dir_path: fs_dir_path = fs_dir_path.replace('//', '/')
            # Ensure it starts with a single slash
            if not fs_dir_path.startswith('/'): fs_dir_path = '/' + fs_dir_path
            # Keep trailing slash ONLY for root, remove otherwise
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
            self.log(self._("error_not_connected", default="Error: Not connected."))
            return {}

        fs_dir_path = self.calculate_fs_path(self._raw_scan_path, self._raw_mount_point)
        if fs_dir_path is None:
            # Error already logged by calculate_fs_path
            self.log(self._("error_path_calc_failed", default="Error: Could not calculate scan path."))
            return {}

        self.log(self._("find_starting", default="Starting duplicate file scan..."))

        potential_duplicates = defaultdict(list)
        count = 0
        video_files_checked = 0
        errors_getting_attrs = 0
        start_time = time.time()

        try:
            self.log(f"DEBUG: Attempting fs.walk_path('{fs_dir_path}')")
            # Use detail=False (default) as getting attributes individually seems more reliable
            # with potential missing hashes/dates. detail=True might be faster if all data is present.
            walk_iterator = self.fs.walk_path(fs_dir_path)

            for foldername, _, filenames in walk_iterator:
                foldername_str = str(foldername).replace('\\', '/') # Normalize folder path

                for filename_obj in filenames:
                    count += 1
                    filename_str = str(filename_obj) # Ensure filename is string

                    # Log progress periodically
                    if count % 500 == 0:
                        self.log(self._("status_scan_progress", count=count, video_count=video_files_checked, default=f"Scanned {count} items... Found {video_files_checked} videos."))

                    # Check if it's a video file based on extension
                    file_extension = os.path.splitext(filename_str)[1].lower()
                    if file_extension in VIDEO_EXTENSIONS:
                        video_files_checked += 1
                        current_file_path = _build_full_path(foldername_str, filename_str)

                        try:
                            # Get necessary attributes: path, modTime, size, SHA1 hash
                            attrs = self.fs.attr(current_file_path)

                            # Extract SHA1 hash (key '2' in fileHashes dictionary)
                            file_sha1 = None
                            # Safely access nested dictionary
                            file_hashes_dict = attrs.get('fileHashes')
                            if isinstance(file_hashes_dict, dict):
                                file_sha1 = file_hashes_dict.get('2') # '2' corresponds to SHA1 in clouddrive

                            # Skip files without SHA1 hash - they cannot be compared for duplicates
                            if not file_sha1:
                                # self.log(f"DEBUG: SHA1 hash not found for {current_file_path}. Skipping.") # Optional debug log
                                errors_getting_attrs += 1
                                continue

                            # Get and parse modification time
                            mod_time_str = attrs.get('modifiedTime')
                            mod_time_dt = _parse_datetime(mod_time_str) # Parse to aware datetime object
                            if mod_time_dt is None:
                                 self.log(self._("error_parse_date", path=current_file_path, error="Unknown format or missing", default=f"Error parsing date for {current_file_path}"))
                                 # Skip files with unparseable dates if date-based rules are important
                                 errors_getting_attrs += 1
                                 continue

                            # Get file size
                            file_size = attrs.get('size', 0) # Default to 0 if missing

                            # Store file info using SHA1 as the key
                            file_info = {
                                'path': current_file_path,
                                'modified': mod_time_dt, # Store as datetime object
                                'size': int(file_size), # Ensure integer size
                                'sha1': file_sha1       # Store SHA1 for reference within the list too
                            }
                            potential_duplicates[file_sha1].append(file_info)

                        except KeyError as ke:
                            # Handle cases where expected keys (like 'fileHashes' or 'modifiedTime') are missing
                            err_msg = self._("error_get_attrs", path=current_file_path, error=f"Missing expected attribute key: {ke}", default=f"Error getting attrs for {current_file_path}: Missing key {ke}")
                            self.log(err_msg)
                            errors_getting_attrs += 1
                        except Exception as e:
                            # Log other errors getting attributes for a specific file
                            err_msg = self._("error_get_attrs", path=current_file_path, error=e, default=f"Error getting attrs for {current_file_path}: {e}")
                            self.log(err_msg)
                            errors_getting_attrs += 1
                            # Optionally add traceback for specific file errors if needed
                            # self.log(traceback.format_exc(limit=1))


            # --- Scan finished ---
            end_time = time.time()
            duration = end_time - start_time
            self.log(f"Scan finished in {duration:.2f} seconds. Total items encountered: {count}. Video files checked: {video_files_checked}.")
            if errors_getting_attrs > 0:
                self.log(f"WARNING: Encountered {errors_getting_attrs} errors while retrieving file attributes/hashes/dates.")

            # Filter for actual duplicates (sets with more than one file sharing the same SHA1)
            actual_duplicates = {sha1: files for sha1, files in potential_duplicates.items() if len(files) > 1}

            # Log final result count
            if actual_duplicates:
                self.log(self._("find_complete_found", count=len(actual_duplicates), default=f"Found {len(actual_duplicates)} duplicate sets."))
            else:
                self.log(self._("find_complete_none", default="No duplicate video files found."))

            return actual_duplicates

        except Exception as walk_e:
            # Catch errors during the walk process itself (e.g., network error, path not found)
            err_msg = self._("error_scan_path", path=fs_dir_path, error=walk_e, default=f"Error walking path '{fs_dir_path}': {walk_e}")
            self.log(err_msg)
            self.log(traceback.format_exc()) # Log full traceback for walk errors
            # Provide a user-friendly summary error message
            self.log(self._("find_error_during", error=walk_e, default=f"Error during duplicate scan: {walk_e}"))
            return {} # Return empty on major error


    def write_duplicates_report(self, duplicate_sets, output_file):
        """ Writes the dictionary of found duplicate file sets to a text file. """
        if not duplicate_sets:
             self.log(self._("save_report_no_data", default="No duplicate sets found to save."))
             return False
        try:
            with open(output_file, "w", encoding='utf-8') as f:
                header = self._("save_report_header", default="Duplicate Video File Sets Found (Based on SHA1):")
                f.write(f"{header}\n===================================================\n\n")
                set_count = 0

                # Sort sets by SHA1 for consistent output order
                sorted_sha1s = sorted(duplicate_sets.keys())

                for sha1 in sorted_sha1s:
                    files_in_set = duplicate_sets[sha1]
                    # Basic check, though find_duplicates should ensure > 1 file
                    if not files_in_set or len(files_in_set) < 2: continue

                    set_count += 1
                    set_header = self._("save_report_set_header", index=set_count, sha1=sha1, count=len(files_in_set), default=f"Set {set_count} (SHA1: {sha1}) - {len(files_in_set)} files")
                    f.write(f"{set_header}\n")

                    # Sort files within the set by path for clarity
                    sorted_files = sorted(files_in_set, key=lambda item: item.get('path', '')) # Use get for safety

                    for file_info in sorted_files:
                         file_label = self._("save_report_file_label", default="  - File:")
                         # Safely format datetime, handle potential None
                         mod_time_obj = file_info.get('modified')
                         mod_time_str = mod_time_obj.strftime(DATE_FORMAT) if isinstance(mod_time_obj, datetime) else "N/A"
                         # Safely calculate size, handle potential None or non-numeric
                         size_bytes = file_info.get('size')
                         size_mb = size_bytes / (1024 * 1024) if isinstance(size_bytes, (int, float)) else 0.0

                         details_label = self._("save_report_details_label",
                                                path=file_info.get('path', 'N/A'),
                                                modified=mod_time_str,
                                                size_mb=size_mb,
                                                default=f"    (Path: {file_info.get('path', 'N/A')}, Modified: {mod_time_str}, Size: {size_mb:.2f} MB)")
                         f.write(f"{file_label}\n{details_label}\n")
                    f.write("\n") # Blank line between sets

            # Log success message
            save_msg = self._("save_report_saved", file=output_file, default=f"Report saved to {output_file}")
            self.log(save_msg)
            return True
        except IOError as ioe:
             # Specific file IO error
             error_msg = self._("save_report_error", file=output_file, error=ioe, default=f"Error saving report to {output_file}: {ioe}")
             self.log(error_msg)
             self.log(traceback.format_exc())
             return False
        except Exception as e:
             # Other unexpected errors
             error_msg = self._("save_report_error", file=output_file, error=e, default=f"Error saving report to {output_file}: {e}")
             self.log(error_msg)
             self.log(traceback.format_exc())
             return False


    def delete_files(self, files_to_delete):
        """ Deletes a list of files from the cloud drive. """
        if not self.fs:
            self.log(self._("error_not_connected", default="Error: Not connected."))
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
            # Ensure forward slashes for cloud path consistency
            cloud_path = file_path.replace('\\', '/')
            self.log(f"Deleting [{i+1}/{total_to_delete}]: {cloud_path}")
            try:
                self.fs.remove(cloud_path)
                deleted_count += 1
                time.sleep(0.05) # Keep small delay to potentially reduce API hammering rate limits
            except Exception as e:
                # Use the specific deletion error key from translations
                error_log_msg = self._("delete_error_file", path=cloud_path, error=e, default=f"Error deleting {cloud_path}: {e}")
                self.log(error_log_msg)
                errors_deleting.append(cloud_path)
                # Log traceback for the first few deletion errors for debugging?
                # if len(errors_deleting) <= 5:
                #      self.log(traceback.format_exc(limit=1))


        # Log final status using the specific translation key
        finish_msg = self._("delete_finished", deleted_count=deleted_count, total_marked=total_to_delete, default=f"Deletion complete. Deleted {deleted_count} of {total_to_delete} files.")
        self.log(finish_msg)

        # Log summary of failed deletions if any occurred
        if errors_deleting:
             self.log(f"WARNING: Failed to delete {len(errors_deleting)} files:")
             # Log first few failed paths for quick reference
             for failed_path in errors_deleting[:10]: # Log up to 10 failed paths
                 self.log(f"  - {failed_path}")
             if len(errors_deleting) > 10:
                 self.log(f"  ... and {len(errors_deleting) - 10} more.")

        # Return actual deleted count and the number attempted
        return deleted_count, total_to_delete


# --- GUI Application Class ---
class DuplicateFinderApp:
    def __init__(self, master):
        self.master = master
        self.current_language = self.load_language_preference()
        self.finder = DuplicateFileFinder()
        self.finder.set_translator(self._)

        # Application state
        self.duplicate_sets = {}
        self.treeview_item_map = {} # Maps tree item ID (path) to file info dict
        self.files_to_delete_cache = [] # Stores paths marked for deletion by the current rule

        # Tkinter variables
        self.widgets = {} # Dictionary to store references to widgets
        self.string_vars = {} # For Entry widgets
        self.entries = {} # References to Entry widgets
        self.deletion_rule_var = tk.StringVar(value="") # Holds the selected deletion rule
        self.suffix_entry_var = tk.StringVar() # For the suffix rule input

        # --- Canvas and Background Image Setup ---
        self.background_photo = None # Holds the PhotoImage to prevent GC
        self._original_pil_image = None # Holds the original loaded PIL Image
        self._last_canvas_width = 0
        self._last_canvas_height = 0
        self._canvas_item_bg = None # ID of the canvas image item
        self._canvas_item_main = None # ID of the canvas window item (for content)
        self._resize_job_id = None # Holds the 'after' job ID for debouncing

        # --- Window Setup ---
        master.title(self._("window_title", default="Duplicate Finder"))
        master.geometry("950x750") # Initial size
        master.minsize(800, 600) # Minimum allowed size

        # --- Create Canvas as the base layer ---
        # No border, no highlight to make it blend seamlessly
        self.background_canvas = tk.Canvas(master, borderwidth=0, highlightthickness=0)
        self.background_canvas.pack(fill=tk.BOTH, expand=True)
        self.widgets["canvas"] = self.background_canvas

        # Load original background image (if available and Pillow installed)
        self._load_original_image()
        # Create placeholder for background image item on canvas
        if PIL_AVAILABLE and self._original_pil_image:
            try:
                # Create a tiny transparent image initially - it will be replaced on first resize
                dummy_img = ImageTk.PhotoImage(Image.new('RGBA', (1, 1), (0,0,0,0)))
                self._canvas_item_bg = self.background_canvas.create_image(
                    0, 0, anchor="nw", image=dummy_img, tags=("bg_image",) # Use tuple for tags
                )
                self.background_photo = dummy_img # Keep reference! Essential.
            except Exception as e:
                print(f"Error creating initial dummy background image: {e}")
                self._canvas_item_bg = None

        # --- Set Application Icon ---
        try:
            if os.path.exists(ICON_FILE):
                master.iconbitmap(ICON_FILE)
            else:
                # Use print for early warnings as log widget might not exist yet
                print(f"Warning: Icon file not found at '{ICON_FILE}'")
        except tk.TclError as e:
            # Catch specific Tkinter errors (e.g., invalid format)
            icon_err_msg = self._("error_icon_load", path=ICON_FILE, error=e, default=f"Error loading icon '{ICON_FILE}': {e}")
            print(icon_err_msg)
        except Exception as e:
             # Catch other potential errors
             icon_err_msg = self._("error_icon_load", path=ICON_FILE, error=f"Unexpected error: {e}", default=f"Error loading icon '{ICON_FILE}': {e}")
             print(icon_err_msg)


        # --- Menu Bar ---
        self.menu_bar = Menu(master)
        master.config(menu=self.menu_bar)
        self.create_menus() # Populate the menu bar

        # --- Create Main Content Frame ---
        # This frame holds all other widgets and is placed onto the canvas
        self.main_content_frame = ttk.Frame(self.background_canvas) # Parent is canvas
        # Note: main_content_frame is NOT packed/gridded here.
        # Its size and position are controlled by the canvas's create_window item.

        # --- Create Paned Window INSIDE main_content_frame ---
        # This allows the top/middle/bottom sections to be resized vertically
        self.paned_window = ttk.PanedWindow(self.main_content_frame, orient=tk.VERTICAL)
        self.paned_window.pack(fill=tk.BOTH, expand=True) # Fills its parent (main_content_frame)

        # --- Build the rest of the UI Structure ---
        # Call a method to create and layout all the frames, buttons, labels, etc.
        # This keeps the __init__ method cleaner.
        self._build_ui_structure()

        # --- Place main_content_frame onto the canvas ---
        # This MUST be done AFTER the paned window and its children have been
        # packed/gridded inside main_content_frame.
        self._canvas_item_main = self.background_canvas.create_window(
            0, 0, # Position at top-left corner of the canvas
            anchor="nw", # Anchor the window item at its North-West corner
            window=self.main_content_frame, # The widget to embed
            tags=("main_content",) # Assign a tag for easier reference (use tuple)
        )

        # --- Initial State & Language ---
        self.load_config() # Load settings from config.ini
        self.update_ui_language() # Set initial text based on language preference
        self.set_ui_state(tk.NORMAL) # Enable/disable widgets based on initial state

        # --- Bind Resize Event to Canvas & Trigger Initial Resize ---
        # We bind to the canvas's resize event, as its size dictates background/content size
        self.background_canvas.bind('<Configure>', self.on_resize_debounced)

        # --- Trigger initial resize after event loop starts ---
        # Use after_idle to ensure the window and canvas have received their
        # initial dimensions before the first resize calculation and drawing happens.
        self.master.after_idle(lambda: self.on_resize_debounced(None))


    def _build_ui_structure(self):
        """Creates and packs/grids all the UI widgets within main_content_frame/paned_window."""

        # --- Top Frame (Config & Actions) - Parent is paned_window ---
        self.top_frame = ttk.Frame(self.paned_window, padding=5)
        self.paned_window.add(self.top_frame, weight=0) # Add to PanedWindow, fixed height
        self.top_frame.columnconfigure(0, weight=1) # Allow config frame to expand horizontally

        # Config Frame (inside top_frame)
        self.widgets["config_frame"] = ttk.LabelFrame(self.top_frame, text=self._("config_title", default="Configuration"), padding="10")
        self.widgets["config_frame"].grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        self.widgets["config_frame"].columnconfigure(1, weight=1) # Make entry column expand

        config_labels = {
            "address": "address_label", "account": "account_label", "password": "password_label",
            "scan_path": "scan_path_label", "mount_point": "mount_point_label"
        }
        for i, (key, label_key) in enumerate(config_labels.items()):
            # Labels inside config_frame
            label = ttk.Label(self.widgets["config_frame"], text=self._(label_key, default=label_key))
            label.grid(row=i, column=0, padx=(5,2), pady=5, sticky=tk.W)
            self.widgets[f"label_{key}"] = label

            # Entries inside config_frame
            var = tk.StringVar()
            self.string_vars[key] = var
            entry_args = {"textvariable": var, "width": 60}
            if key == "password": entry_args["show"] = "*"
            entry = ttk.Entry(self.widgets["config_frame"], **entry_args)
            entry.grid(row=i, column=1, padx=(2,5), pady=5, sticky=tk.EW)
            self.entries[key] = entry


        # Action Buttons Frame (inside top_frame)
        action_button_frame = ttk.Frame(self.top_frame, padding="5")
        action_button_frame.grid(row=1, column=0, padx=5, pady=(5,0), sticky="ew")

        btn_info = [
            ("load", "load_config_button", self.load_config, 5),
            ("save", "save_config_button", self.save_config, 5),
            ("test_conn", "test_connection_button", self.start_test_connection_thread, (10, 5)),
            ("find", "find_button", self.start_find_duplicates_thread, (15, 5)),
        ]

        for w_key, t_key, cmd, padx_val in btn_info:
             # Buttons inside action_button_frame
             initial_state = tk.NORMAL if w_key in ["load", "save", "test_conn"] else tk.DISABLED
             button = ttk.Button(action_button_frame, text=self._(t_key, default=t_key.replace('_',' ').title()), command=cmd, state=initial_state)
             button.pack(side=tk.LEFT, padx=padx_val, pady=5)
             self.widgets[f"{w_key}_button"] = button


        # --- Middle Frame (Rules, TreeView) - Parent is paned_window ---
        self.middle_frame = ttk.Frame(self.paned_window, padding=5)
        self.paned_window.add(self.middle_frame, weight=1) # Add to PanedWindow, allow vertical expansion
        self.middle_frame.rowconfigure(1, weight=1)    # Allow tree frame row (row 1) to expand vertically
        self.middle_frame.columnconfigure(0, weight=1) # Allow content column to expand horizontally

        # Deletion Rules Frame (inside middle_frame)
        rules_frame = ttk.LabelFrame(self.middle_frame, text=self._("rules_title", default="Deletion Rules"), padding="10")
        rules_frame.grid(row=0, column=0, padx=5, pady=(10, 5), sticky="ew")
        self.widgets["rules_frame"] = rules_frame
        # Give column 1 weight so suffix entry doesn't crowd radios if window is wide
        rules_frame.columnconfigure(1, weight=1)
        rules_frame.columnconfigure(2, weight=0) # Suffix entry has fixed width

        rule_opts = [
            ("shortest_path", RULE_KEEP_SHORTEST), ("longest_path", RULE_KEEP_LONGEST),
            ("oldest", RULE_KEEP_OLDEST), ("newest", RULE_KEEP_NEWEST),
            ("keep_suffix", RULE_KEEP_SUFFIX)
        ]
        self.rule_radios = {}
        for i, (t_key_suffix, value) in enumerate(rule_opts):
            t_key = f"rule_{t_key_suffix}"
            # Radios inside rules_frame
            radio = ttk.Radiobutton(rules_frame, text=self._(t_key, default=t_key.replace('_',' ').title()),
                                    variable=self.deletion_rule_var, value=value,
                                    command=self._on_rule_change, state=tk.DISABLED)
            radio.grid(row=i, column=0, padx=5, pady=2, sticky="w")
            self.rule_radios[value] = radio
            self.widgets[f"radio_{value}"] = radio

        # Suffix Entry Widgets (inside rules_frame, aligned with the last radio)
        suffix_row_index = len(rule_opts) - 1
        lbl = ttk.Label(rules_frame, text=self._("rule_suffix_entry_label", default="Suffix:"), state=tk.DISABLED)
        lbl.grid(row=suffix_row_index, column=1, padx=(10, 2), pady=2, sticky="w")
        self.widgets["suffix_label"] = lbl

        entry = ttk.Entry(rules_frame, textvariable=self.suffix_entry_var, width=15, state=tk.DISABLED)
        entry.grid(row=suffix_row_index, column=2, padx=(0, 5), pady=2, sticky="w")
        self.widgets["suffix_entry"] = entry
        self.entries["suffix"] = entry


        # TreeView Frame (inside middle_frame)
        tree_frame = ttk.Frame(self.middle_frame)
        tree_frame.grid(row=1, column=0, padx=5, pady=5, sticky="nsew") # Fill available space
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        # Treeview (inside tree_frame)
        self.columns = ("action", "path", "modified", "size_mb", "set_id")
        self.tree = ttk.Treeview(tree_frame, columns=self.columns, show="headings", selectmode="none")
        self.widgets["treeview"] = self.tree
        self.setup_treeview_headings() # Set initial column header text

        # Configure Treeview columns
        self.tree.column("action", width=60, anchor=tk.CENTER, stretch=tk.NO)
        self.tree.column("path", width=450, anchor=tk.W) # Allow path to stretch more if needed
        self.tree.column("modified", width=140, anchor=tk.W, stretch=tk.NO)
        self.tree.column("size_mb", width=80, anchor=tk.E, stretch=tk.NO)
        self.tree.column("set_id", width=80, anchor=tk.CENTER, stretch=tk.NO)
        # Optionally make path column stretch: self.tree.column("path", stretch=tk.YES)

        # Scrollbars (inside tree_frame)
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        # Layout tree and scrollbars using grid
        self.tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')


        # --- Bottom Frame (Log & Final Action Buttons) - Parent is paned_window ---
        self.bottom_frame = ttk.Frame(self.paned_window, padding=5)
        self.paned_window.add(self.bottom_frame, weight=0) # Add to PanedWindow, fixed height
        self.bottom_frame.columnconfigure(0, weight=1) # Allow log frame column to expand
        self.bottom_frame.rowconfigure(1, weight=1) # Allow log frame row to expand

        # Final Action Frame (inside bottom_frame)
        final_action_frame = ttk.Frame(self.bottom_frame)
        final_action_frame.grid(row=0, column=0, columnspan=2, padx=5, pady=(5,0), sticky="ew")

        # Style for the red delete button
        style = ttk.Style()
        try: style.configure("Danger.TButton", foreground="white", background="red", font=('TkDefaultFont', 10, 'bold'))
        except tk.TclError: style.configure("Danger.TButton", foreground="red") # Fallback

        # Buttons inside final_action_frame
        btn_del = ttk.Button(final_action_frame, text=self._("delete_by_rule_button", default="Delete by Rule"),
                             command=self.start_delete_by_rule_thread, state=tk.DISABLED, style="Danger.TButton")
        btn_del.pack(side=tk.LEFT, padx=5, pady=5)
        self.widgets["delete_button"] = btn_del

        chart_state = tk.DISABLED
        chart_tkey = "show_chart_button_disabled" if not MATPLOTLIB_AVAILABLE else "show_chart_button"
        btn_chart = ttk.Button(final_action_frame, text=self._(chart_tkey, default="Show Chart"),
                               command=self.show_cloud_file_types, state=chart_state)
        btn_chart.pack(side=tk.LEFT, padx=(10, 5), pady=5)
        self.widgets["chart_button"] = btn_chart

        btn_save = ttk.Button(final_action_frame, text=self._("save_list_button", default="Save Report"),
                              command=self.save_duplicates_report, state=tk.DISABLED)
        btn_save.pack(side=tk.LEFT, padx=5, pady=5)
        self.widgets["save_list_button"] = btn_save


        # Log Frame (inside bottom_frame)
        log_frame = ttk.LabelFrame(self.bottom_frame, text=self._("log_title", default="Log"), padding="5")
        log_frame.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(5,0))
        self.widgets["log_frame"] = log_frame
        log_frame.rowconfigure(0, weight=1) # Make text area expand within frame
        log_frame.columnconfigure(0, weight=1)

        # Log Text ScrolledText (inside log_frame)
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=8, width=100,
                                                  state='disabled', relief=tk.SUNKEN, borderwidth=1)
        self.log_text.grid(row=0, column=0, sticky="nsew")
        self.widgets["log_text"] = self.log_text


    def _load_original_image(self):
        """Loads the original background PIL image."""
        self._original_pil_image = None # Reset previous image
        if not PIL_AVAILABLE:
            # Log only if Pillow is expected/wanted
            # print("Pillow library not found, cannot load background image.")
            return
        if not os.path.exists(BACKGROUND_IMAGE_FILE):
            print(f"Warning: Background image file not found at '{BACKGROUND_IMAGE_FILE}'")
            return

        try:
            self._original_pil_image = Image.open(BACKGROUND_IMAGE_FILE)
            # Ensure it's in a format Tkinter can handle well, like RGBA for transparency support
            # This might slightly increase memory usage but avoids potential format issues.
            self._original_pil_image = self._original_pil_image.convert("RGBA")
            print(f"Background image loaded successfully: {BACKGROUND_IMAGE_FILE}")
        except Exception as e:
            # Log error loading image
            err_msg = self._("error_img_load", path=BACKGROUND_IMAGE_FILE, error=e, default=f"Error loading background image '{BACKGROUND_IMAGE_FILE}': {e}")
            print(err_msg) # Use print for early errors before logger might be ready
            # self.log_message(err_msg) # Can log later if needed


    # --- Debounced Resize Handler ---
    def on_resize_debounced(self, event):
        """
        Handles the <Configure> event bound to the canvas.
        Immediately resizes the main content frame on the canvas.
        Schedules the expensive background image resize using debouncing.
        """
        try:
            # Get current canvas dimensions reliably
            # Use winfo_width/height as event might be None on initial call or during rapid events
            if not self.background_canvas or not self.background_canvas.winfo_exists(): return
            new_width = self.background_canvas.winfo_width()
            new_height = self.background_canvas.winfo_height()
        except tk.TclError:
            return # Window likely closing

        # Avoid unnecessary processing for tiny/invalid sizes (e.g., during minimize)
        if new_width <= 1 or new_height <= 1:
            return

        # --- IMMEDIATE ACTION: Resize the content frame window item ---
        # This makes the UI elements (PanedWindow, frames, etc.) resize instantly with the window.
        if self._canvas_item_main is not None:
            try:
                # Configure the embedded window's size to match the canvas
                self.background_canvas.itemconfig(self._canvas_item_main, width=new_width, height=new_height)
                # Keep position at top-left (usually not necessary to change coords)
                # self.background_canvas.coords(self._canvas_item_main, 0, 0)
            except tk.TclError: pass # Ignore if canvas/window closing

        # --- DEBOUNCED ACTION: Resize background image ---
        # Optimization: Only schedule background resize if size actually changed since last time
        if new_width == self._last_canvas_width and new_height == self._last_canvas_height:
            return

        # Update the last known size *before* scheduling the job
        self._last_canvas_width = new_width
        self._last_canvas_height = new_height

        # Cancel any previously scheduled resize job (the core of debouncing)
        if self._resize_job_id:
            self.master.after_cancel(self._resize_job_id)
            self._resize_job_id = None # Clear the old job ID

        # Schedule the actual (expensive) image resize operation after the delay
        # Pass the current target width and height to the scheduled function
        self._resize_job_id = self.master.after(RESIZE_DEBOUNCE_DELAY,
                                                 self._perform_background_resize,
                                                 new_width, new_height)


    def _perform_background_resize(self, width, height):
        """
        Actually resizes the background image. This is the function called
        by the debouncer after the delay. Runs in the main GUI thread.
        """
        self._resize_job_id = None # Mark the job ID as processed

        # --- Safety Checks ---
        # Ensure Pillow is available, image was loaded, canvas item exists, and canvas is still valid
        if not PIL_AVAILABLE or not self._original_pil_image or self._canvas_item_bg is None \
           or not self.background_canvas or not self.background_canvas.winfo_exists():
            return

        # print(f"Performing background resize to {width}x{height}") # Debugging line

        try:
            # --- Resize the Original Image ---
            # Use LANCZOS for best quality, but it's slower.
            # Consider Image.Resampling.BILINEAR for better performance during rapid resize if needed.
            resampling_filter = Image.Resampling.LANCZOS
            resized_pil_image = self._original_pil_image.resize((width, height), resampling_filter)

            # --- Create New PhotoImage ---
            # Tkinter requires a *new* PhotoImage object for canvas updates. Keep a reference!
            self.background_photo = ImageTk.PhotoImage(resized_pil_image)

            # --- Update Canvas Item ---
            # Configure the existing canvas image item to use the new PhotoImage
            self.background_canvas.itemconfig(self._canvas_item_bg, image=self.background_photo)

            # --- Ensure Stacking Order ---
            # Make sure the background image item is drawn *behind* the main content item.
            if self._canvas_item_main:
                self.background_canvas.tag_lower(self._canvas_item_bg, self._canvas_item_main)

        except Exception as e:
            # Catch errors during the image processing/update
            print(f"Error during background resize execution: {e}")
            # Optionally log to the GUI log as well:
            # self.log_message(f"Error performing background resize: {e}")


    # --- Language Handling ---
    def _(self, key, **kwargs):
        """ Translation helper method. """
        lang_dict = translations.get(self.current_language, translations[DEFAULT_LANG])
        default_val = kwargs.pop('default', f"<{key}?>") # Provide a default if key is missing
        base_string = lang_dict.get(key, translations[DEFAULT_LANG].get(key, default_val))
        try:
            # Only format if placeholders exist and arguments are provided
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
        """ Saves the current language selection to a JSON file. """
        try:
            with open(LANG_PREF_FILE, 'w', encoding='utf-8') as f:
                json.dump({"language": self.current_language}, f)
        except IOError as e:
            # Log error saving language pref (e.g., permissions)
            print(f"Warning: Could not save language preference to {os.path.basename(LANG_PREF_FILE)}: {e}")
            self.log_message(f"Warning: Could not save language preference: {e}")
        except Exception as e:
             print(f"Error saving language preference: {e}")
             self.log_message(f"Error saving language preference: {e}")

    def load_language_preference(self):
        """ Loads the language preference from a JSON file. """
        # Ensure resource_path is used via the constant
        pref_path = LANG_PREF_FILE
        try:
            if os.path.exists(pref_path):
                with open(pref_path, 'r', encoding='utf-8') as f:
                    prefs = json.load(f)
                    lang = prefs.get("language", DEFAULT_LANG)
                    # Validate loaded language code against available translations
                    return lang if lang in translations else DEFAULT_LANG
        except (IOError, json.JSONDecodeError) as e:
            # Log error loading language pref (e.g., file corrupted)
            print(f"Warning: Could not load language preference from {os.path.basename(pref_path)}: {e}")
            # No GUI log here as logger might not be ready yet
        except Exception as e:
             print(f"Unexpected error loading language preference: {e}")
        return DEFAULT_LANG # Default language if load fails or file doesn't exist

    def create_menus(self):
        """ Creates the application's menu bar. """
        lang_menu = Menu(self.menu_bar, tearoff=0)
        self.widgets["lang_menu"] = lang_menu
        # Add commands - labels will be updated by update_ui_language
        lang_menu.add_command(label="English", command=lambda: self.change_language("en"))
        lang_menu.add_command(label="中文", command=lambda: self.change_language("zh"))
        # Add the cascade menu to the main menu bar - label updated by update_ui_language
        self.menu_bar.add_cascade(label="Language", menu=lang_menu)


    def change_language(self, lang_code):
        """ Changes the application language and updates the UI. """
        if lang_code in translations and lang_code != self.current_language:
            print(f"Changing language to: {lang_code}")
            self.current_language = lang_code
            self.finder.set_translator(self._) # Update finder's translator instance
            self.update_ui_language() # Update all GUI elements immediately
            self.save_language_preference() # Save the new preference
        elif lang_code not in translations:
             # Log if attempted language is not supported
             print(f"Error: Language '{lang_code}' not supported.")
             self.log_message(f"Error: Language '{lang_code}' not supported.")

    def update_ui_language(self):
        """ Updates the text of all UI elements based on the current language. """
        print(f"Updating UI language to: {self.current_language}")
        # Ensure master window exists before trying to update widgets
        if not self.master or not self.master.winfo_exists():
             print("Error: Master window does not exist during UI language update.")
             return

        try: # Wrap in try/except to catch errors during update (e.g., widget destroyed)
            self.master.title(self._("window_title", default="App"))

            # Update Menu Bar Cascade Label
            if self.menu_bar and self.menu_bar.winfo_exists():
                try: self.menu_bar.entryconfig(0, label=self._("menu_language", default="Language"))
                except tk.TclError: pass # Ignore if destroyed
                except IndexError: pass # Ignore if menu structure changed

            # Update Language Menu Items
            lang_menu = self.widgets.get("lang_menu")
            if lang_menu and lang_menu.winfo_exists():
                try: lang_menu.entryconfig(0, label=self._("menu_english", default="English"))
                except tk.TclError: pass
                except IndexError: pass
                try: lang_menu.entryconfig(1, label=self._("menu_chinese", default="中文"))
                except tk.TclError: pass
                except IndexError: pass

            # Update LabelFrames Titles
            for key, title_key in [("config_frame", "config_title"),
                                ("rules_frame", "rules_title"),
                                ("log_frame", "log_title")]:
                widget = self.widgets.get(key)
                if widget and widget.winfo_exists():
                    try: widget.config(text=self._(title_key, default=title_key.replace('_',' ').title()))
                    except tk.TclError: pass # Ignore if widget destroyed mid-update

            # Update Labels
            label_keys = {"label_address": "address_label", "label_account": "account_label",
                        "label_password": "password_label", "label_scan_path": "scan_path_label",
                        "label_mount_point": "mount_point_label",
                        "suffix_label": "rule_suffix_entry_label"}
            for widget_key, text_key in label_keys.items():
                widget = self.widgets.get(widget_key)
                if widget and widget.winfo_exists():
                    try: widget.config(text=self._(text_key, default=text_key.replace('_',' ').title()))
                    except tk.TclError: pass

            # Update Buttons
            button_keys = {"load_button": "load_config_button", "save_button": "save_config_button",
                        "test_conn_button": "test_connection_button",
                        "find_button": "find_button",
                        "delete_button": "delete_by_rule_button",
                        "save_list_button": "save_list_button"}
             # Add chart button separately as its text depends on state too
            all_button_keys = list(button_keys.items())
            all_button_keys.append(("chart_button", "show_chart_button")) # Use base key here

            for widget_key, text_key in all_button_keys:
                widget = self.widgets.get(widget_key)
                if widget and widget.winfo_exists():
                    try:
                        # Special handle chart button text based on state/matplotlib availability
                        if widget_key == "chart_button":
                            is_enabled = widget.cget('state') == tk.NORMAL
                            # Determine the correct translation key based on enabled status and library
                            effective_text_key = text_key if (is_enabled and MATPLOTLIB_AVAILABLE) else "show_chart_button_disabled"
                            widget.config(text=self._(effective_text_key, default=effective_text_key.replace('_',' ').title()))
                        else:
                            widget.config(text=self._(text_key, default=text_key.replace('_',' ').title()))
                    except tk.TclError: pass

            # Update Radio Buttons Text
            radio_keys = {RULE_KEEP_SHORTEST: "rule_shortest_path", RULE_KEEP_LONGEST: "rule_longest_path",
                        RULE_KEEP_OLDEST: "rule_oldest", RULE_KEEP_NEWEST: "rule_newest",
                        RULE_KEEP_SUFFIX: "rule_keep_suffix"}
            for value, text_key in radio_keys.items():
                # Construct the widget key used during creation
                widget_key = f"radio_{value}"
                widget = self.widgets.get(widget_key)
                if widget and widget.winfo_exists():
                    try: widget.config(text=self._(text_key, default=text_key.replace('_',' ').title()))
                    except tk.TclError: pass

            # Update Treeview Headings
            self.setup_treeview_headings()

            # Re-apply rule highlighting/action text if tree has items
            # This ensures the "Keep"/"Delete" text in the action column updates
            tree = self.widgets.get("treeview")
            if self.duplicate_sets and tree and tree.winfo_exists():
                 self._apply_rule_to_treeview()

            print("UI Language update complete.")

        except Exception as e:
             # Catch any unexpected error during the update process
             print(f"ERROR during UI language update: {e}")
             self.log_message(f"ERROR during UI language update: {e}")
             # Optionally log traceback for debugging
             self.log_message(traceback.format_exc(limit=3))

    def setup_treeview_headings(self):
        """Sets or updates the text of the treeview headings based on current language."""
        heading_keys = { "action": "tree_rule_action_col", "path": "tree_path_col",
                         "modified": "tree_modified_col", "size_mb": "tree_size_col",
                         "set_id": "tree_set_col" }
        tree = self.widgets.get("treeview")
        if tree and tree.winfo_exists():
             try:
                 # Ensure columns are actually defined in the treeview instance
                 current_columns = tree['columns']
                 for col_id, text_key in heading_keys.items():
                     if col_id in current_columns: # Check if column exists
                         # Use the translation function _ to get the current language text
                         tree.heading(col_id, text=self._(text_key, default=col_id.replace('_',' ').title()))
             except tk.TclError: pass # Ignore if widget is destroyed during update


    # --- GUI Logic Methods ---
    def log_message(self, message):
        """ Safely appends a message to the log ScrolledText widget from any thread. """
        message_str = str(message) if message is not None else ""
        log_widget = self.widgets.get("log_text")
        # Ensure master and log widget exist before scheduling update
        if hasattr(self, 'master') and self.master and self.master.winfo_exists() and \
           log_widget and log_widget.winfo_exists():
            try:
                # Use 'after' for thread safety with Tkinter GUI updates
                self.master.after(0, self._append_log, message_str)
            except (tk.TclError, RuntimeError):
                # Ignore errors if window is closing or already destroyed
                pass
            except Exception as e:
                # Log error if scheduling itself fails
                print(f"Error scheduling log message: {e}\nMessage: {message_str}")

    def _append_log(self, message):
        """ Internal method to append message to log widget (MUST run in main thread). """
        log_widget = self.widgets.get("log_text")
        # Double-check widget existence within the main thread callback
        if not log_widget or not log_widget.winfo_exists(): return
        try:
            current_state = log_widget.cget('state')
            log_widget.configure(state='normal') # Enable writing
            timestamp = datetime.now().strftime("%H:%M:%S")
            log_widget.insert(tk.END, f"[{timestamp}] {message}\n")
            log_widget.see(tk.END) # Auto-scroll to the end
            log_widget.configure(state=current_state) # Restore original state (usually 'disabled')
        except tk.TclError:
            # Widget might have been destroyed between check and update
            pass
        except Exception as e:
            # Catch other potential errors during append
            print(f"Unexpected error appending log: {e}\nMessage: {message}")

    def load_config(self):
        """ Loads configuration from the ini file. """
        # Uses resource_path via CONFIG_FILE constant
        config_path = CONFIG_FILE
        self.log_message(self._("status_loading_config", file=os.path.basename(config_path), default=f"Loading config from {os.path.basename(config_path)}..."))
        config = configparser.ConfigParser()
        try:
            if not os.path.exists(config_path):
                 self.log_message(self._("status_config_not_found", file=os.path.basename(config_path), default=f"Config file '{os.path.basename(config_path)}' not found."))
                 return # Nothing to load
            # Attempt to read the file
            read_files = config.read(config_path, encoding='utf-8')
            if not read_files:
                 # File exists but couldn't be parsed or read (e.g., empty, corrupted format)
                 self.log_message(f"Warning: Config file '{os.path.basename(config_path)}' exists but could not be read/parsed.")
                 return

            # Check if the expected section exists
            if 'config' in config:
                cfg = config['config']
                # Populate GUI fields from config, providing defaults if keys missing
                self.string_vars["address"].set(cfg.get("clouddrvie2_address", ""))
                self.string_vars["account"].set(cfg.get("clouddrive2_account", ""))
                self.string_vars["password"].set(cfg.get("clouddrive2_passwd", ""))
                # Use original keys 'root_path' and 'clouddrive2_root_path' for loading compatibility
                self.string_vars["scan_path"].set(cfg.get("root_path", ""))
                self.string_vars["mount_point"].set(cfg.get("clouddrive2_root_path", ""))
                self.log_message(self._("status_config_loaded", default="Config loaded."))
            else:
                self.log_message(self._("status_config_section_missing", default="Config section missing."))
                # Optionally set default values in GUI fields here if section missing

        except configparser.Error as e:
            # Error specifically during parsing
            error_msg = self._("error_config_read", error=e, default=f"Error reading config file:\n{e}")
            # Show error to user and log it
            if self.master.winfo_exists():
                 messagebox.showerror(self._("error_config_title", default="Config Error"), error_msg, master=self.master) # Specify master
            self.log_message(error_msg)
        except Exception as e:
             # Catch other potential errors (e.g., file permission)
             error_msg = f"{self._('error_unexpected', error='', default='Unexpected error').rstrip(': ')} loading config: {e}"
             if self.master.winfo_exists():
                 messagebox.showerror(self._("error_title", default="Error"), error_msg, master=self.master)
             self.log_message(error_msg)
             self.log_message(traceback.format_exc()) # Log details for unexpected errors


    def save_config(self):
        """ Saves current configuration from GUI fields to the ini file. """
        # Uses resource_path via CONFIG_FILE constant
        config_path = CONFIG_FILE
        self.log_message(self._("status_saving_config", file=os.path.basename(config_path), default=f"Saving config to {os.path.basename(config_path)}..."))
        config = configparser.ConfigParser()

        # Prepare the data to save from current GUI values
        config_data = {
            "clouddrvie2_address": self.string_vars["address"].get(),
            "clouddrive2_account": self.string_vars["account"].get(),
            "clouddrive2_passwd": self.string_vars["password"].get(),
            # Use original keys 'root_path' and 'clouddrive2_root_path' for writing compatibility
            "root_path": self.string_vars["scan_path"].get(),
            "clouddrive2_root_path": self.string_vars["mount_point"].get(),
        }
        config['config'] = config_data

        # Attempt to preserve other sections if the file exists
        try:
            if os.path.exists(config_path):
                 config_old = configparser.ConfigParser()
                 # Read existing config to merge sections
                 config_old.read(config_path, encoding='utf-8')
                 for section in config_old.sections():
                     # Skip the 'config' section we are overwriting
                     if section != 'config':
                         # Ensure the section exists before adding items
                         if not config.has_section(section): config.add_section(section)
                         # Copy items from old section to new config object
                         config[section] = dict(config_old.items(section))
                     else:
                         # Merge keys within 'config' section if needed (optional)
                         # Add old keys only if they don't exist in the new data (preserves unused old keys)
                         for key, value in config_old.items(section):
                             if key not in config['config']:
                                 config['config'][key] = value
        except Exception as e:
            # Log warning if merging old config fails, but proceed with saving new data
            print(f"Warning: Could not merge old config sections during save: {e}")
            self.log_message(f"Warning: Could not merge old config sections during save: {e}")

        try:
            # Write the combined config object to the file
            with open(config_path, 'w', encoding='utf-8') as configfile:
                config.write(configfile)
            self.log_message(self._("status_config_saved", default="Config saved."))
        except IOError as e:
            # Specific error for file writing issues (permissions, disk full, etc.)
            error_msg = self._("error_config_save", error=e, default=f"Could not write config:\n{e}")
            if self.master.winfo_exists():
                 messagebox.showerror(self._("error_config_save_title", default="Config Save Error"), error_msg, master=self.master)
            self.log_message(error_msg)
        except Exception as e:
             # Catch other unexpected errors during save
             error_msg = f"{self._('error_unexpected', error='', default='Unexpected error').rstrip(': ')} saving config: {e}"
             if self.master.winfo_exists():
                 messagebox.showerror(self._("error_title", default="Error"), error_msg, master=self.master)
             self.log_message(error_msg)
             self.log_message(traceback.format_exc()) # Log details

    def _check_path_chars(self, path_dict):
        """ Validates characters in specified path inputs. Shows warning if needed. """
        suspicious_char_found = False
        all_details = []
        # Define which input fields (keys in path_dict) to validate and their display names
        path_display_names = {
            "address": self._("address_label", default="Address").rstrip(': '),
            "scan_path": self._("scan_path_label", default="Scan Path").rstrip(': '),
            "mount_point": self._("mount_point_label", default="Mount Point").rstrip(': ')
            # Add other path fields here if needed, e.g., "output_dir"
        }

        for key, path_str in path_dict.items():
             # Only validate keys present in our display name mapping
             if key not in path_display_names: continue

             issues = _validate_path_chars(path_str) # Call the helper function
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

            # Get the multi-line warning message for the popup
            warning_msg_template = self._("path_warning_suspicious_chars", default="Suspicious character(s) detected!\nThis often happens from copy-pasting.\nPlease DELETE and MANUALLY RETYPE the paths in the GUI.")
            warning_lines = warning_msg_template.split('\n')
            # Format for popup (first line as title/header, rest as body)
            popup_msg = f"{warning_lines[0]}\n\n" + "\n".join(warning_lines[1:]) if len(warning_lines) > 1 else warning_lines[0]
            instruction = warning_lines[1] if len(warning_lines) > 1 else "Please check and retype the paths."

            self.log_message(f"Instruction: {instruction}")
            self.log_message(log_sep)

            # Show error message box in the main thread, associated with the main window
            if self.master.winfo_exists():
                # Use after(0, ...) to ensure it runs in the main GUI thread
                # Pass master explicitly using a dictionary for the last argument
                self.master.after(0, messagebox.showerror, warning_title, popup_msg, {"master": self.master})
            return False # Indicate failure due to suspicious chars

        return True # Indicate success (no suspicious chars found)


    def set_ui_state(self, new_state):
        """Enable/disable UI elements based on state (e.g., 'normal', 'finding', 'deleting')."""
        # Define states as strings for clarity
        STATE_NORMAL = 'normal'
        STATE_FINDING = 'finding'
        STATE_DELETING = 'deleting'
        STATE_TESTING_CONN = 'testing_connection'
        STATE_CHARTING = 'charting'

        # Determine general operational status
        is_normal_op = (new_state == STATE_NORMAL)
        is_busy = new_state in [STATE_FINDING, STATE_DELETING, STATE_TESTING_CONN, STATE_CHARTING]

        # Determine capabilities based on current application data/state
        has_connection = self.finder is not None and self.finder.fs is not None
        has_duplicates = bool(self.duplicate_sets) # Are there results loaded?
        is_rule_selected = bool(self.deletion_rule_var.get()) # Is a deletion rule radio button selected?
        # Crucially, check the CACHE for files marked by the *last applied rule*
        has_files_marked_for_deletion = bool(self.files_to_delete_cache)

        # --- Define enable/disable logic based on state and capabilities ---
        can_config_interact = is_normal_op # Can load/save/edit config only when idle
        can_test_conn = is_normal_op # Can test connection only when idle
        can_find = is_normal_op and has_connection # Can start find only when connected and idle
        can_apply_rules = is_normal_op and has_duplicates # Can select rules only when duplicates are found and idle
        # Can delete only if idle, duplicates found, a rule is selected, AND that rule marked files for deletion
        can_delete = is_normal_op and has_duplicates and is_rule_selected and has_files_marked_for_deletion
        can_save_report = is_normal_op and has_duplicates # Can save report if duplicates found and idle
        can_chart = is_normal_op and has_connection and MATPLOTLIB_AVAILABLE # Can chart if connected, lib available, and idle

        # --- Apply States to Widgets ---
        # Use .get() for widgets dictionary and widget.winfo_exists() for safety

        # Config Entries (Address, Account, Password, Paths)
        entry_state = tk.NORMAL if can_config_interact else tk.DISABLED
        for key, entry in self.entries.items():
            # Keep suffix entry tied to rule radio state (handled below)
            if key != "suffix" and entry and entry.winfo_exists():
                try: entry.config(state=entry_state)
                except tk.TclError: pass

        # Config Buttons (Load, Save)
        config_btn_state = tk.NORMAL if can_config_interact else tk.DISABLED
        load_btn = self.widgets.get("load_button")
        if load_btn and load_btn.winfo_exists(): load_btn.config(state=config_btn_state)
        save_btn = self.widgets.get("save_button")
        if save_btn and save_btn.winfo_exists(): save_btn.config(state=config_btn_state)

        # Test Connection Button
        test_conn_btn = self.widgets.get("test_conn_button")
        if test_conn_btn and test_conn_btn.winfo_exists(): test_conn_btn.config(state=tk.NORMAL if can_test_conn else tk.DISABLED)

        # Find Duplicates Button
        find_btn = self.widgets.get("find_button")
        if find_btn and find_btn.winfo_exists(): find_btn.config(state=tk.NORMAL if can_find else tk.DISABLED)

        # Deletion Rules Widgets (Radios, Suffix Label/Entry)
        rules_state = tk.NORMAL if can_apply_rules else tk.DISABLED
        rules_frame = self.widgets.get("rules_frame") # Check parent frame exists
        if rules_frame and rules_frame.winfo_exists():
            # Enable/disable radio buttons
            for radio in self.rule_radios.values():
                 if radio and radio.winfo_exists():
                     try: radio.config(state=rules_state)
                     except tk.TclError: pass
            # Enable/disable suffix label/entry based on main rule state AND selected rule
            suffix_widgets_state = tk.DISABLED # Default disabled
            # Only enable suffix controls if rules are generally active AND the suffix rule is selected
            if rules_state == tk.NORMAL and self.deletion_rule_var.get() == RULE_KEEP_SUFFIX:
                suffix_widgets_state = tk.NORMAL
            suffix_label = self.widgets.get("suffix_label")
            if suffix_label and suffix_label.winfo_exists(): suffix_label.config(state=suffix_widgets_state)
            suffix_entry = self.widgets.get("suffix_entry")
            if suffix_entry and suffix_entry.winfo_exists(): suffix_entry.config(state=suffix_widgets_state)

        # Delete Button - state depends on the calculated 'can_delete' condition
        delete_btn = self.widgets.get("delete_button")
        if delete_btn and delete_btn.winfo_exists(): delete_btn.config(state=tk.NORMAL if can_delete else tk.DISABLED)

        # Save Report Button
        save_report_btn = self.widgets.get("save_list_button")
        if save_report_btn and save_report_btn.winfo_exists(): save_report_btn.config(state=tk.NORMAL if can_save_report else tk.DISABLED)

        # Chart Button - state and text depend on 'can_chart' and MATPLOTLIB_AVAILABLE
        chart_btn = self.widgets.get("chart_button")
        if chart_btn and chart_btn.winfo_exists():
            chart_state = tk.NORMAL if can_chart else tk.DISABLED
            # Choose correct translation key based on availability and enabled status
            chart_text_key = "show_chart_button" if (chart_state == tk.NORMAL and MATPLOTLIB_AVAILABLE) else "show_chart_button_disabled"
            chart_btn.config(state=chart_state, text=self._(chart_text_key, default=chart_text_key.replace('_',' ').title()))

        # Treeview and Log are generally always visible/enabled, their content changes.

    def start_find_duplicates_thread(self):
        """ Handles the 'Find Duplicates' button click. Validates inputs, starts worker thread. """
        # 1. Check connection status
        if not self.finder or not self.finder.fs:
             messagebox.showwarning(self._("error_title", default="Error"), self._("error_not_connected", default="Not connected."), master=self.master)
             self.log_message(self._("error_not_connected", default="Error: Not connected."))
             return

        # 2. Check essential path inputs are filled
        paths_to_check = {
            "scan_path": self.string_vars["scan_path"].get(),
            "mount_point": self.string_vars["mount_point"].get()
            # Optionally add address/account if needed for path calc, but usually checked by connection
        }
        if not paths_to_check["scan_path"] or not paths_to_check["mount_point"]:
              # Use a generic missing input error first
              messagebox.showerror(self._("error_input_title", default="Input Error"), self._("error_input_missing", default="Address, Account, Root Path, and Mount Point are required."), master=self.master)
              self.log_message(self._("error_input_missing", default="Error: Missing required inputs.") + " (Scan Path / Mount Point required for finding duplicates)")
              return

        # 3. Check for suspicious characters in relevant paths
        if not self._check_path_chars(paths_to_check):
            # Error message already shown by _check_path_chars
            return

        # 4. Prepare for scan: Clear previous results and set UI to busy state
        self.clear_results() # Clears tree, stored data, rule selection, delete cache
        self.log_message(self._("find_starting", default="Starting scan..."))
        self.set_ui_state("finding") # Disable UI elements during scan

        # 5. Start the background worker thread
        thread = threading.Thread(target=self._find_duplicates_worker, daemon=True)
        thread.start()

    def _find_duplicates_worker(self):
        """ Worker thread for finding duplicates. Calls finder logic, schedules GUI update. """
        # Double check connection within thread just in case state changed rapidly
        if not self.finder or not self.finder.fs:
            self.log_message("Error: Connection lost before Find duplicates could execute.")
            # Schedule UI reset in main thread if window still exists
            if self.master.winfo_exists(): self.master.after(0, self.set_ui_state, 'normal')
            return

        found_duplicates = {} # Initialize result dictionary
        try:
            start_time = time.time()
            # Call the core logic in the finder class
            found_duplicates = self.finder.find_duplicates()
            end_time = time.time()
            # Log duration of the core find operation (network/API calls)
            self.log_message(f"Find duplicates network/scan part took {end_time - start_time:.2f} seconds.")

            # --- Schedule GUI update in main thread ---
            # Pass the found duplicates dictionary to the processing function
            if self.master.winfo_exists():
                self.master.after(0, self._process_find_results, found_duplicates)

        except Exception as e:
            # Catch unexpected errors during the finder.find_duplicates() call itself
            err_msg = self._("find_error_during", error=e, default=f"Error during scan: {e}")
            self.log_message(err_msg)
            self.log_message(traceback.format_exc()) # Log details for debugging
            if self.master.winfo_exists():
                 # Schedule error message box and reset UI state in main thread
                 self.master.after(0, messagebox.showerror, self._("error_title", default="Error"), err_msg, {"master": self.master})
                 self.master.after(0, self.set_ui_state, 'normal') # Reset state on error
        # No finally block here for UI state reset; _process_find_results handles it on success.


    def _process_find_results(self, found_duplicates):
        """ Processes results from find_duplicates worker (runs in main thread). Updates UI. """
        # Ensure window hasn't been closed while worker was running
        if not self.master.winfo_exists(): return

        # Store the results in the application instance
        self.duplicate_sets = found_duplicates if found_duplicates else {}

        if self.duplicate_sets:
            # Populate the treeview if duplicates were found
            self.populate_treeview() # This logs population start/end and applies rule
            # Log summary (already logged by find_duplicates and populate_treeview)
        else:
            # No duplicates found, message already logged by find_duplicates
             pass # No need to populate tree

        # Set final UI state back to normal *after* processing and population is complete
        self.set_ui_state('normal')


    def clear_results(self):
        """Clears the treeview, stored duplicates, rule selection, and delete cache."""
        self.log_message(self._("status_clearing_tree", default="Clearing results list..."))
        self.duplicate_sets = {}
        self.treeview_item_map = {}
        self.files_to_delete_cache = [] # IMPORTANT: Clear the delete cache
        self.deletion_rule_var.set("") # Deselect rule radio button

        # Clear Treeview items
        tree = self.widgets.get("treeview")
        if tree and tree.winfo_exists():
            try:
                # Efficiently delete all top-level items and their children
                tree.delete(*tree.get_children())
            except tk.TclError: pass # Ignore if tree destroyed during clear

        # Update UI state AFTER clearing data, so conditions (like empty cache/no rule) are correct
        # This will disable rules/delete/save buttons appropriately.
        self.set_ui_state('normal')


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

        # Sort by SHA1 for consistent set numbering across runs
        sorted_sha1s = sorted(self.duplicate_sets.keys())

        for sha1 in sorted_sha1s:
            files_in_set = self.duplicate_sets[sha1]
            # Ensure it's actually a duplicate set (should be pre-filtered, but check)
            if len(files_in_set) < 2: continue

            set_index += 1
            # Sort files within set by path for initial display order
            sorted_files = sorted(files_in_set, key=lambda x: x.get('path', '')) # Use get with default

            for file_info in sorted_files:
                # Safely get and format data, providing defaults
                path = file_info.get('path', 'N/A')
                mod_time = file_info.get('modified')
                mod_time_str = mod_time.strftime(DATE_FORMAT) if isinstance(mod_time, datetime) else "N/A"
                size = file_info.get('size')
                size_mb = size / (1024 * 1024) if isinstance(size, (int, float)) else 0.0
                set_id_str = self._("tree_set_col_value", index=set_index, default=str(set_index))

                # Values tuple must match the `self.columns` order ("action", "path", "modified", "size_mb", "set_id")
                values = (
                    "", # Placeholder for 'Action' column - filled by _apply_rule_to_treeview
                    path,
                    mod_time_str,
                    f"{size_mb:.2f}", # Format size to 2 decimal places
                    set_id_str
                 )
                # Use the file path as the unique item ID (iid) in the tree
                # Ensure path is not empty, fallback if necessary though unlikely
                item_id = path if path != 'N/A' else f"set{set_index}_sha{sha1[:8]}_{mod_time_str}" # Generate fallback iid
                tree_items_data.append((item_id, values, file_info))

        # --- Insert items into the Treeview ---
        # ttk.Treeview doesn't have a true batch insert. Insert one by one.
        # This can be slow for very large numbers of items (>10k). Consider virtual lists for extreme cases.
        items_inserted = 0
        for item_id, values, file_info in tree_items_data:
             try:
                 # Insert item with its unique ID (path) and column values
                 tree.insert("", tk.END, iid=item_id, values=values)
                 # Map the item ID back to the full file_info dictionary for later use
                 self.treeview_item_map[item_id] = file_info
                 items_inserted += 1
             except tk.TclError as e:
                 # Handle potential error if item_id (path) causes issues (e.g., contains '{' '}')
                 # or if tree is destroyed during insertion
                 self.log_message(f"Error inserting item '{item_id}' into tree: {e}")
                 # Continue attempting to insert others if possible

        end_time = time.time()
        duration = end_time - start_time
        # Log completion status
        self.log_message(self._("status_tree_populated", default="Results list populated.") + f" ({items_inserted} items in {duration:.2f}s)")

        # --- Apply the currently selected rule (if any) ---
        # This will update the 'Action' column and apply highlighting based on the rule
        self._apply_rule_to_treeview()


    def _on_rule_change(self):
        """Called when a deletion rule radio button is selected."""
        # Update suffix entry state based on selected rule
        selected_rule = self.deletion_rule_var.get()
        is_suffix_rule = (selected_rule == RULE_KEEP_SUFFIX)
        suffix_widgets_state = tk.NORMAL if is_suffix_rule else tk.DISABLED

        # Enable/disable suffix label and entry
        suffix_label = self.widgets.get("suffix_label")
        if suffix_label and suffix_label.winfo_exists(): suffix_label.config(state=suffix_widgets_state)
        suffix_entry = self.widgets.get("suffix_entry")
        if suffix_entry and suffix_entry.winfo_exists(): suffix_entry.config(state=suffix_widgets_state)

        # Clear suffix entry text if the rule is not 'Keep Suffix' for usability
        if not is_suffix_rule:
            self.suffix_entry_var.set("")

        # --- Re-apply the selected rule to update the Treeview ---
        # This updates the 'Action' column and highlighting immediately
        self._apply_rule_to_treeview()


    def _apply_rule_to_treeview(self):
        """ Updates the 'Action' column and highlighting in the treeview based on the selected rule. """
        tree = self.widgets.get("treeview")
        # If no data or tree doesn't exist, clear cache and disable delete button
        if not self.duplicate_sets or not tree or not tree.winfo_exists():
            self.files_to_delete_cache = []
            delete_btn = self.widgets.get("delete_button")
            if delete_btn and delete_btn.winfo_exists(): delete_btn.config(state=tk.DISABLED)
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
                         tree.item(item_id, tags=()) # Remove tags ('keep', 'delete')
            except tk.TclError: pass # Ignore if tree destroyed
            # Explicitly disable delete button as no rule means nothing to delete
            delete_btn = self.widgets.get("delete_button")
            if delete_btn and delete_btn.winfo_exists(): delete_btn.config(state=tk.DISABLED)
            return

        # --- Apply the selected rule ---
        # Get display name for logging
        rule_name_display = self._(f"rule_{selected_rule}", default=selected_rule)
        self.log_message(self._("status_applying_rule", rule_name=rule_name_display, count=len(self.duplicate_sets), default=f"Applying rule '{rule_name_display}'..."))
        start_time = time.time()

        # Get translated text for Keep/Delete actions for the tree column
        keep_text = self._("tree_action_keep", default="Keep")
        delete_text = self._("tree_action_delete", default="Delete")
        # Get suffix value only if the suffix rule is selected
        suffix = self.suffix_entry_var.get() if selected_rule == RULE_KEEP_SUFFIX else None

        # Define tags for highlighting (configure once, apply per item)
        try:
            tree.tag_configure('keep', foreground='green')
            tree.tag_configure('delete', foreground='red', font=('TkDefaultFont', 9, 'bold')) # Make delete bold red
        except tk.TclError: pass # Ignore if tree destroyed

        try:
            # --- Determine files to delete based on the rule ---
            # This call might raise ValueError if input is invalid (e.g., missing suffix)
            # It returns a list of full paths to be deleted.
            files_to_delete_list = self._determine_files_to_delete(self.duplicate_sets, selected_rule, suffix)

            # IMPORTANT: Update the cache *before* updating the UI or button state
            self.files_to_delete_cache = files_to_delete_list
            files_to_delete_paths_set = set(files_to_delete_list) # Use set for fast lookups below
            delete_count = len(files_to_delete_paths_set)
            self.log_message(self._("status_rule_applied", delete_count=delete_count, default=f"Rule applied. {delete_count} files marked for deletion."))

            # --- Update Treeview items based on the determined deletion list ---
            for item_id, file_info in self.treeview_item_map.items():
                 if not tree.exists(item_id): continue # Skip if item somehow removed

                 path = file_info.get('path')
                 if not path: continue # Skip if file info is malformed

                 # Determine action text and tags based on whether the path is in the delete set
                 is_marked_for_delete = (path in files_to_delete_paths_set)
                 action_text = delete_text if is_marked_for_delete else keep_text
                 item_tags = ('delete',) if is_marked_for_delete else ('keep',)

                 # Update the treeview item's action column and apply the tag
                 tree.set(item_id, "action", action_text)
                 tree.item(item_id, tags=item_tags)

            # --- Update Delete Button State ---
            # Enable delete button only if the applied rule actually marked files for deletion
            delete_button_state = tk.NORMAL if delete_count > 0 else tk.DISABLED
            delete_btn = self.widgets.get("delete_button")
            if delete_btn and delete_btn.winfo_exists(): delete_btn.config(state=delete_button_state)


        except ValueError as ve: # Catch specific errors from _determine_files_to_delete (e.g., "Suffix missing")
             # Log and show the specific validation error
             self.log_message(str(ve))
             messagebox.showerror(self._("error_input_title", default="Input Error"), str(ve), master=self.master)
             # Clear action column and cache on rule application error
             self.files_to_delete_cache = []
             try:
                 for item_id in self.treeview_item_map.keys():
                     if tree.exists(item_id):
                         tree.set(item_id, "action", "")
                         tree.item(item_id, tags=()) # Clear tags
             except tk.TclError: pass
             # Disable delete button
             delete_btn = self.widgets.get("delete_button")
             if delete_btn and delete_btn.winfo_exists(): delete_btn.config(state=tk.DISABLED)

        except tk.TclError:
             self.log_message("Error updating treeview, it might have been closed.")
             self.files_to_delete_cache = [] # Clear cache if tree is gone
             delete_btn = self.widgets.get("delete_button")
             if delete_btn and delete_btn.winfo_exists(): delete_btn.config(state=tk.DISABLED)
        except Exception as e:
             # Catch any other unexpected error during rule application
             self.log_message(f"Unexpected error applying rule to treeview: {e}")
             self.log_message(traceback.format_exc())
             # Clear cache and disable delete on unexpected error
             self.files_to_delete_cache = []
             delete_btn = self.widgets.get("delete_button")
             if delete_btn and delete_btn.winfo_exists(): delete_btn.config(state=tk.DISABLED)
        finally:
            end_time = time.time()
            # Only log duration if a rule was actually selected and applied
            if selected_rule: self.log_message(f"Rule application / tree update took {end_time-start_time:.3f}s")
            # Call set_ui_state at the end to ensure button states reflect the final cache state
            # Needed because even if no error occurs, the delete button state might change
            self.set_ui_state('normal')


    def _determine_files_to_delete(self, duplicate_sets, rule, suffix_value):
        """
        Determines which files to delete based on the selected rule.
        Returns a list of full paths to delete. Logs warnings for ambiguities.
        Raises ValueError for invalid input (e.g., missing suffix for suffix rule).
        """
        # Basic validation (caller _apply_rule_to_treeview should handle most cases)
        if not duplicate_sets: return []
        if not rule: raise ValueError(self._("delete_no_rule_selected", default="No deletion rule selected."))
        if rule == RULE_KEEP_SUFFIX and not suffix_value:
             # This specific error needs to be raised for the caller to catch
             raise ValueError(self._("delete_suffix_missing", default="Suffix required for 'Keep Suffix' rule."))

        files_to_delete = [] # List to store paths of files identified for deletion
        log_func = self.log_message # Use the GUI logger for feedback/warnings within this function

        # Iterate through each set of duplicates (grouped by SHA1)
        for sha1, files_in_set in duplicate_sets.items():
            # Basic sanity check, should have at least 2 files to be a duplicate set
            if len(files_in_set) < 2: continue

            keep_file = None # The FileInfo dict of the file to keep in this set
            error_in_set = None # Store any error/warning encountered while processing this specific set

            try:
                # --- Apply Rule Logic to find the file to KEEP ---
                if rule == RULE_KEEP_SHORTEST:
                    # Find file with the minimum path length (use get for safety)
                    keep_file = min(files_in_set, key=lambda f: len(f.get('path', '')))
                elif rule == RULE_KEEP_LONGEST:
                    # Find file with the maximum path length
                    keep_file = max(files_in_set, key=lambda f: len(f.get('path', '')))
                elif rule == RULE_KEEP_OLDEST:
                    # Filter out files with invalid/missing dates first
                    valid_files = [f for f in files_in_set if isinstance(f.get('modified'), datetime)]
                    if not valid_files:
                        error_in_set = f"Warning: Skipping set {sha1[:8]}... (Rule: {rule}) - No valid modification dates found."
                        continue # Skip to next set if no valid dates for comparison
                    # Find file with the minimum (oldest) datetime among valid files
                    keep_file = min(valid_files, key=lambda f: f['modified'])
                elif rule == RULE_KEEP_NEWEST:
                    # Filter out files with invalid/missing dates first
                    valid_files = [f for f in files_in_set if isinstance(f.get('modified'), datetime)]
                    if not valid_files:
                        error_in_set = f"Warning: Skipping set {sha1[:8]}... (Rule: {rule}) - No valid modification dates found."
                        continue
                    # Find file with the maximum (newest) datetime among valid files
                    keep_file = max(valid_files, key=lambda f: f['modified'])
                elif rule == RULE_KEEP_SUFFIX:
                    # Find files ending with the specified suffix (case-insensitive)
                    suffix_lower = suffix_value.lower()
                    matching_files = [f for f in files_in_set if f.get('path', '').lower().endswith(suffix_lower)]

                    if len(matching_files) == 1:
                        keep_file = matching_files[0] # Keep the single match
                    elif len(matching_files) > 1:
                        # Tie-breaker: Multiple files match suffix. Keep the shortest path among them.
                        error_in_set = f"Warning: Set {sha1[:8]}...: Multiple files match suffix '{suffix_value}'. Keeping shortest path among matches."
                        keep_file = min(matching_files, key=lambda f: len(f.get('path', '')))
                    else: # len(matching_files) == 0
                        # Tie-breaker: No files match suffix. Keep the overall shortest path file in the set.
                        error_in_set = f"Warning: Set {sha1[:8]}...: No files match suffix '{suffix_value}'. Keeping overall shortest path file instead."
                        keep_file = min(files_in_set, key=lambda f: len(f.get('path', '')))
                else:
                    # Should not happen if rule is validated earlier, but handle defensively
                    error_in_set = f"Internal Error: Unknown rule '{rule}' encountered for set {sha1[:8]}... Skipping deletions for this set."
                    continue # Skip to next set

                # --- Add files to delete list based on the determined keep_file ---
                if keep_file and keep_file.get('path'):
                    keep_path = keep_file['path']
                    # Add all *other* files from the set to the deletion list
                    for f_info in files_in_set:
                        path = f_info.get('path')
                        if path and path != keep_path:
                            files_to_delete.append(path)
                else: # Rule failed to select a file (e.g., invalid data)
                    error_in_set = error_in_set or f"Warning: Set {sha1[:8]}...: Could not determine a file to keep with rule '{rule}'. No files from this set will be deleted."

            except Exception as e:
                 # Catch potential errors during comparison (e.g., unexpected data in FileInfo)
                 error_in_set = f"Error processing rule '{rule}' for set {sha1[:8]}...: {e}. Skipping set."
                 # Log traceback for debugging rule logic errors
                 print(f"Traceback for error processing set {sha1[:8]}...:")
                 traceback.print_exc(limit=2)
                 continue # Skip to the next set on unexpected error
            finally:
                 # Log any warnings or errors encountered for this specific set *after* processing it
                 if error_in_set:
                     log_func(error_in_set) # Log set-specific warnings/errors

        return files_to_delete


    def start_delete_by_rule_thread(self):
        """ Handles the 'Delete Files by Rule' button click. Validates, confirms, starts worker. """
        selected_rule = self.deletion_rule_var.get()
        # Get a display name for the rule for user messages
        rule_name_for_msg = self._(f"rule_{selected_rule}", default=selected_rule) if selected_rule else "N/A"

        # --- Pre-checks ---
        # 1. Check if a rule is selected in the UI
        if not selected_rule:
             messagebox.showerror(self._("error_input_title", default="Input Error"), self._("delete_no_rule_selected", default="Please select a deletion rule."), master=self.master)
             return
        # 2. Check if suffix is provided if the suffix rule is selected
        if selected_rule == RULE_KEEP_SUFFIX and not self.suffix_entry_var.get():
             messagebox.showerror(self._("error_input_title", default="Input Error"), self._("delete_suffix_missing", default="Please enter suffix to keep."), master=self.master)
             return

        # 3. CRITICAL Check: Use the cached list determined by the *last rule application*
        #    Do NOT redetermine here, as the user sees the result of the last application.
        if not self.files_to_delete_cache:
             # Inform user that based on the *currently applied* rule, no files need deletion
             messagebox.showinfo(self._("delete_by_rule_button", default="Delete by Rule"), self._("delete_rule_no_files", default="No files matched deletion criteria."), master=self.master)
             self.log_message(self._("delete_rule_no_files", default="No files matched deletion criteria.") + f" (Rule: {rule_name_for_msg})")
             return

        # --- Confirmation Dialog ---
        num_files = len(self.files_to_delete_cache)
        # Format the confirmation message using translations
        confirm_msg = self._("delete_confirm_msg", rule_name=rule_name_for_msg, default=f"Delete files based on rule: '{rule_name_for_msg}'?\nTHIS CANNOT BE UNDONE.")
        confirm_msg += f"\n\n({num_files} files will be deleted)" # Add count for clarity

        # Ask user for confirmation
        confirm = messagebox.askyesno(
            title=self._("delete_confirm_title", default="Confirm Deletion"),
            message=confirm_msg,
            icon='warning', # Use warning icon for destructive action
            master=self.master # Associate dialog with main window
        )

        # If user cancels
        if not confirm:
            self.log_message(self._("delete_cancelled", default="Deletion cancelled."))
            return

        # --- Start Deletion Process ---
        self.log_message(self._("delete_starting", rule_name=rule_name_for_msg, default=f"Starting deletion (Rule: {rule_name_for_msg})..."))
        self.set_ui_state("deleting") # Disable UI during deletion

        # Pass a *copy* of the cached list to the thread to avoid race conditions
        # if the user changes the rule while deletion is in progress (though UI is disabled)
        files_to_delete_list = list(self.files_to_delete_cache)

        # Start the background worker thread for deletion
        thread = threading.Thread(target=self._delete_worker, args=(files_to_delete_list, rule_name_for_msg), daemon=True)
        thread.start()


    def _delete_worker(self, files_to_delete, rule_name_for_log):
        """ Worker thread for deleting files based on the provided list. """
        # Double-check connection status within the thread
        if not self.finder or not self.finder.fs:
            self.log_message("Error: Connection lost before Deletion could execute.")
            # Schedule UI reset if window exists
            if self.master.winfo_exists(): self.master.after(0, self.set_ui_state, 'normal')
            return

        deleted_count = 0
        total_attempted = len(files_to_delete)

        try:
            # Check again if there are actually files to delete (belt-and-suspenders)
            if not files_to_delete:
                self.log_message(self._("delete_rule_no_files", default="No files to delete.") + f" (Worker check; Rule: {rule_name_for_log})")
            else:
                # Call the finder's delete method (which handles logging of progress and errors)
                deleted_count, total_attempted = self.finder.delete_files(files_to_delete)
                # Final summary log ("Deletion complete...") is handled within finder.delete_files

            # --- Schedule GUI update after deletion attempt ---
            if self.master.winfo_exists():
                # Always clear results after a delete attempt (success or partial failure).
                # This reflects that the state has changed and the displayed list/actions are no longer valid.
                self.master.after(0, self.clear_results)

        except Exception as e:
            # Catch unexpected errors during the finder.delete_files call itself
            # Errors during individual file delete are handled inside finder.delete_files
            err_msg = self._("delete_error_during", error=e, default=f"Error during deletion: {e}")
            self.log_message(err_msg)
            self.log_message(traceback.format_exc()) # Log details
            if self.master.winfo_exists():
                 # Show error message box
                 self.master.after(0, messagebox.showerror, self._("error_title", default="Error"), err_msg, {"master": self.master})
                 # Do NOT clear results here on unexpected worker error; allow user to see list/retry?

        finally:
            # --- ALWAYS Re-enable UI in the main thread ---
            # Ensure UI is re-enabled regardless of success or failure of the deletion process.
            # clear_results scheduled above will handle the state update eventually,
            # but setting state back to normal here ensures UI is responsive sooner if clear takes time.
            if self.master.winfo_exists():
                # Schedule state reset back to normal
                 self.master.after(0, self.set_ui_state, 'normal')


    def save_duplicates_report(self):
        """ Saves the report of FOUND duplicate file sets to a text file. """
        # Check if duplicates data exists to save
        if not self.duplicate_sets:
            messagebox.showinfo(self._("save_report_no_data", default="No Duplicates Found"),
                                self._("save_report_no_data", default="No duplicate sets found to save."),
                                master=self.master)
            self.log_message(self._("save_report_no_data", default="No duplicate sets to save."))
            return

        # Generate default filename with timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        initial_filename = f"duplicates_report_{timestamp}.txt"

        # Ask user for save location and filename using standard dialog
        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
            title=self._("save_list_button", default="Save Found Duplicates Report As..."),
            initialfile=initial_filename,
            master=self.master # Associate dialog with main window
        )

        # If user provided a path (didn't cancel)
        if file_path:
            # Call the finder's report writing method
            # Pass the currently stored duplicate sets data
            success = self.finder.write_duplicates_report(self.duplicate_sets, file_path)
            # Logging of success/failure is handled within write_duplicates_report
            # Optionally show a success message box here if desired and successful
            # if success and self.master.winfo_exists():
            #     messagebox.showinfo("Report Saved", self._("save_report_saved", file=file_path), master=self.master)
        else:
             # User cancelled the save dialog
             self.log_message("Save report operation cancelled by user.")


    def start_test_connection_thread(self):
        """ Starts the connection test in a background thread. """
        self.log_message(self._("status_connecting", default="Attempting connection..."))
        # Set UI to a busy state during connection test
        self.set_ui_state("testing_connection") # Use a specific state or just disable relevant parts
        thread = threading.Thread(target=self._test_connection_worker, daemon=True)
        thread.start()


    def _test_connection_worker(self):
        """ Worker thread for testing the CloudDrive2 connection. """
        # Get necessary config values from GUI StringVars
        address = self.string_vars["address"].get()
        account = self.string_vars["account"].get()
        passwd = self.string_vars["password"].get()
        # Need scan path and mount point for finder.set_config context, even for test
        mount_point = self.string_vars["mount_point"].get()
        scan_path = self.string_vars["scan_path"].get()
        connected = False

        try:
            # --- Input Validation ---
            # Check required fields for connection test itself
            if not all([address, account, mount_point]): # Password might be optional depending on setup
                 error_msg = self._("error_input_missing_conn", default="Address, Account, and Mount Point are required for connection test.")
                 self.log_message(error_msg)
                 # Show error message box in main thread
                 if self.master.winfo_exists(): self.master.after(0, messagebox.showerror, self._("error_input_title", default="Input Error"), error_msg, {"master": self.master})
                 return # Exit worker early

            # --- Path Character Validation ---
            # Check relevant paths for suspicious characters before passing to set_config
            paths_to_check_conn = {"address": address, "mount_point": mount_point, "scan_path": scan_path}
            if not self._check_path_chars(paths_to_check_conn):
                # Error already logged and shown by _check_path_chars
                return # Exit worker early

            # --- Attempt Connection using finder.set_config ---
            # set_config handles the actual connection attempt and logs success/failure messages
            connected = self.finder.set_config(address, account, passwd, scan_path, mount_point, self.log_message)

            # --- Show Result Message Box ---
            # Schedule the message box to run in the main GUI thread
            if self.master.winfo_exists():
                 if connected:
                     # Success message
                     self.master.after(0, messagebox.showinfo, self._("conn_test_success_title", default="Connection Test Successful"), self._("conn_test_success_msg", default="Successfully connected."), {"master": self.master})
                 else:
                     # Failure message (error details already logged by set_config)
                     self.master.after(0, messagebox.showwarning, self._("conn_test_fail_title", default="Connection Test Failed"), self._("conn_test_fail_msg", default="Failed to connect. Check log."), {"master": self.master})

        except Exception as e:
            # Catch unexpected errors during the worker setup or validation steps
            error_msg = f'{self._("error_unexpected", default="Unexpected error", error="").rstrip(": ")} during connection test worker: {e}'
            self.log_message(error_msg)
            self.log_message(traceback.format_exc()) # Log details
            if self.master.winfo_exists():
                self.master.after(0, messagebox.showerror, self._("error_title", default="Error"), error_msg, {"master": self.master})
        finally:
            # --- ALWAYS Re-enable UI ---
            # Ensure UI state is reset back to normal in the main thread regardless of outcome
            if self.master.winfo_exists():
                self.master.after(0, self.set_ui_state, 'normal')


    def show_cloud_file_types(self):
        """ Handles the 'Show Cloud File Types' button click. Starts worker thread for chart. """
        # 1. Check Matplotlib dependency is met
        if not MATPLOTLIB_AVAILABLE:
            messagebox.showwarning(self._("chart_error_title", default="Chart Error"), self._("chart_error_no_matplotlib", default="Matplotlib library not found."), master=self.master)
            return
        # 2. Check connection status
        if not self.finder or not self.finder.fs:
            messagebox.showwarning(self._("chart_error_title", default="Chart Error"), self._("chart_error_no_connection", default="Not connected."), master=self.master)
            return
        # 3. Check required inputs for chart path calculation
        scan_path_raw = self.string_vars["scan_path"].get()
        mount_point_raw = self.string_vars["mount_point"].get()
        if not scan_path_raw or not mount_point_raw:
             messagebox.showwarning(self._("error_input_title", default="Input Error"), self._("error_input_missing_chart", default="Scan Path and Mount Point required."), master=self.master)
             return
        # 4. Check path characters for safety
        paths_to_check_chart = {"scan_path": scan_path_raw, "mount_point": mount_point_raw}
        if not self._check_path_chars(paths_to_check_chart): return

        # 5. Set UI state to busy and start worker thread
        self.set_ui_state("charting") # Disable UI during chart generation
        thread = threading.Thread(target=self._show_cloud_file_types_worker,
                                  args=(scan_path_raw, mount_point_raw), daemon=True)
        thread.start()


    def _show_cloud_file_types_worker(self, scan_path_raw, mount_point_raw):
        """ Worker thread to scan files and prepare data for the pie chart. """
        fs_dir_path = None
        file_counts = collections.Counter()
        total_files = 0
        scan_error = None

        try:
             # --- Setup & Validation inside worker ---
             # Re-calculate path within worker to be self-contained
             # Ensure finder exists before calculating path
             if not self.finder:
                  self.log_message("Error: Finder object not initialized for charting.")
                  return
             fs_dir_path = self.finder.calculate_fs_path(scan_path_raw, mount_point_raw)
             if fs_dir_path is None:
                 # Error already logged by calculate_fs_path
                 return # Exit worker

             # Double check connection within worker just before scan
             if not self.finder.fs:
                 self.log_message(self._("chart_error_no_connection", default="Cannot chart: Not connected."))
                 return

             # --- Scan Files ---
             self.log_message(self._("chart_status_scanning_cloud", path=fs_dir_path, default=f"Scanning {fs_dir_path} for file types..."))
             try:
                 # Iterate through files using walk_path to count extensions
                 for _, _, filenames in self.finder.fs.walk_path(fs_dir_path):
                     for filename_obj in filenames:
                         filename = str(filename_obj) # Ensure string
                         total_files += 1
                         _, ext = os.path.splitext(filename)
                         # Use translation for no extension label, provide fallback
                         ext_label = ext.lower() if ext else self._("chart_label_no_extension", default="[No Ext]")
                         file_counts[ext_label] += 1
             except Exception as e:
                 # Catch errors during the walk (network, permissions, etc.)
                 scan_error = e
                 error_msg = self._("chart_error_cloud_scan", path=fs_dir_path, error=e, default=f"Error scanning cloud path '{fs_dir_path}': {e}")
                 self.log_message(error_msg)
                 self.log_message(traceback.format_exc()) # Log details
                 # Schedule error message box in main thread
                 if self.master.winfo_exists():
                      self.master.after(0, messagebox.showerror, self._("chart_error_title", default="Chart Error"), error_msg, {"master": self.master})
                 # Do not proceed to chart generation if scan failed

             # --- Schedule GUI Update (Chart Creation or Info Message) ---
             # Define the function to run in the main thread using closure
             def update_gui_after_scan():
                 # Ensure window exists and no scan error occurred before proceeding
                 if not self.master.winfo_exists() or scan_error: return

                 # Check if any files were found
                 if not file_counts:
                     no_files_msg = self._("chart_status_no_files_found", path=fs_dir_path, default=f"No files found in '{fs_dir_path}'.")
                     self.log_message(no_files_msg)
                     # Show info message box
                     if self.master.winfo_exists(): messagebox.showinfo(self._("chart_info_title", default="Chart Info"), no_files_msg, master=self.master)
                     return # Nothing to chart

                 # Log generation message and create the chart window
                 self.log_message(self._("chart_status_generating", count=len(file_counts), total=total_files, default=f"Generating chart for {len(file_counts)} types ({total_files} files)..."))
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


    # This method belongs inside the DuplicateFinderApp class

    def _create_pie_chart_window(self, counts, display_path):
        """ Creates and displays the file type pie chart in a new Toplevel window. """
        # Double-check matplotlib availability right before creation
        if not MATPLOTLIB_AVAILABLE: # 8 spaces indent (inside method)
            self.log_message("Error: Matplotlib became unavailable before chart creation.")
            if self.master.winfo_exists(): # 12 spaces indent (inside if)
                messagebox.showerror(self._("chart_error_title", default="Chart Error"), self._("chart_error_no_matplotlib", default="Matplotlib library not found."), master=self.master)
            return # 12 spaces indent (inside if)

        chart_window = None # Initialize window reference (8 spaces)

        try: # 8 spaces indent
            # --- Ensure Matplotlib settings (optional, but good practice) ---
            try: # 12 spaces indent (inside outer try)
                matplotlib.rcParams['axes.unicode_minus'] = False
            except Exception as font_e: # 12 spaces indent
                self.log_message(f"Warning: Issue setting Matplotlib unicode_minus: {font_e}")

            # --- Data Preparation for Chart --- (12 spaces indent)
            top_n = 15 # Show top N categories + 'Others' for clarity
            total_count = sum(counts.values())
            sorted_counts = counts.most_common() # Sort by count descending

            labels = []
            sizes = []
            others_label = self._("chart_label_others", default="Others")
            others_count = 0
            others_sources = [] # Keep track of extensions grouped into 'Others'

            # Group small slices into 'Others' if necessary
            if len(sorted_counts) > top_n: # 12 spaces indent
                # Take top N items
                top_items = sorted_counts[:top_n] # 16 spaces indent (inside if)
                labels = [item[0] for item in top_items]
                sizes = [item[1] for item in top_items]
                # Group remaining items into 'Others'
                other_items = sorted_counts[top_n:]
                others_count = sum(item[1] for item in other_items)
                others_sources = [item[0] for item in other_items]
                # Add 'Others' category if it has items
                if others_count > 0: # 16 spaces indent
                    labels.append(others_label) # 20 spaces indent
                    sizes.append(others_count)
            else: # 12 spaces indent (else corresponding to outer if)
                # Fewer than top_n items, show all individually
                labels = [item[0] for item in sorted_counts] # 16 spaces indent
                sizes = [item[1] for item in sorted_counts]

            # Log if items were grouped for user awareness (12 spaces indent)
            if others_count > 0:
                 self.log_message(f"Chart Note: Grouped {len(others_sources)} smaller categories ({others_count} files) into '{others_label}'.")

            # --- Create Chart Window (Toplevel) --- (12 spaces indent)
            chart_window = Toplevel(self.master) # Parent is the main application window
            chart_window.title(self._("chart_window_title", path=display_path, default=f"File Types in {display_path}"))
            chart_window.geometry("850x650") # Adjust size as needed

            # --- Create Matplotlib Figure & Axes --- (12 spaces indent)
            fig = Figure(figsize=(8, 6), dpi=100) # Adjust figsize for content
            ax = fig.add_subplot(111)

            # --- Generate Pie Chart --- (12 spaces indent)
            # Customize appearance: explode slices slightly, set start angle, format percentages
            explode_value = 0.02 # Slight explosion for visual separation
            explode = [explode_value] * len(labels)
            # Make 'Others' slice less exploded if it exists
            if others_count > 0 and others_label in labels: # 12 spaces indent
                try: # Added try-except for robustness if labels list is unexpectedly empty
                    explode[labels.index(others_label)] = 0 # No explosion for 'Others' (16 spaces)
                except ValueError: # 16 spaces indent
                    pass # others_label somehow not in labels, ignore

            wedges, texts, autotexts = ax.pie(sizes, explode=explode, autopct='%1.1f%%',
                                              startangle=140, pctdistance=0.85, # Adjust pctdistance to move % labels
                                              # Optional: Add shadows, text properties
                                              # shadow=True, textprops={'fontsize': 8}
                                              ) # (continuation, 14 spaces indent relative to ax.pie start)
            ax.axis('equal') # Ensures pie is drawn as a circle. (12 spaces)

            # --- Create Legend --- (12 spaces indent)
            # Create labels with counts for the legend
            legend_labels = [f'{label} ({sizes[i]})' for i, label in enumerate(labels)]
            # Place legend outside the pie chart area
            ax.legend(wedges, legend_labels,
                      title=self._("chart_legend_title", default="File Extensions"),
                      loc="center left", bbox_to_anchor=(1.05, 0.5), # (x, y) position relative to axes
                      fontsize='small', frameon=True) # Add frame for better visibility

            # --- Style Percentage Labels --- (12 spaces indent)
            # Make percentage labels more readable, especially on dark slices
            for autotext in autotexts: # 12 spaces indent
                autotext.set_color('white') # 16 spaces indent
                autotext.set_size(8) # Adjust size as needed
                autotext.set_weight('bold')
                # Add a semi-transparent background box for contrast
                autotext.set_bbox(dict(facecolor='black', alpha=0.6, pad=1, edgecolor='none'))

            # --- Adjust Layout --- (12 spaces indent)
            try:
                # Adjust layout to prevent labels/legend overlapping plot area
                # rect=[left, bottom, right, top] defines the area for the plot elements
                fig.tight_layout(rect=[0, 0, 0.80, 1]) # Leave 20% space on right for legend
            except Exception as layout_err:
                # tight_layout can sometimes fail, log warning but proceed
                print(f"Warning: tight_layout issue: {layout_err}. Legend/labels might overlap.")
                self.log_message(f"Warning: Chart layout issue: {layout_err}")

            # --- Embed Matplotlib Figure in Tkinter Window --- (12 spaces indent)
            canvas = FigureCanvasTkAgg(fig, master=chart_window)
            canvas_widget = canvas.get_tk_widget()

            # Add standard Matplotlib navigation toolbar
            toolbar = NavigationToolbar2Tk(canvas, chart_window)
            toolbar.update() # Needed step for toolbar state

            # Pack the toolbar and canvas widget into the Toplevel window
            # Pack toolbar first so it appears above the canvas
            # toolbar.pack(side=tk.BOTTOM, fill=tk.X) # Or TOP
            canvas_widget.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

            # Draw the canvas
            canvas.draw()

        except Exception as e: # 8 spaces indent (corresponding to outer try)
            # Catch errors during chart creation/display
            error_msg = f"Error creating chart window: {e}"
            self.log_message(error_msg)
            self.log_message(traceback.format_exc()) # Log details
            if self.master.winfo_exists(): # 12 spaces indent
                messagebox.showerror(self._("chart_error_title", default="Chart Error"), error_msg, master=self.master)
            # Destroy the Toplevel window if it exists but chart creation failed
            if chart_window and chart_window.winfo_exists(): # 12 spaces indent
                 chart_window.destroy()


# --- Main Execution Block ---
if __name__ == "__main__":
    # --- Setup Root Window ---
    root = tk.Tk()
    # Minimum size is set in the App class __init__

    # --- Create and Run App ---
    try:
        # Basic check if translations seem populated (prevent hard crash if copy/paste missed)
        if not translations["en"].get("window_title") or not translations["zh"].get("window_title"):
            print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
            print("ERROR: The 'translations' dictionary appears incomplete.")
            print("Please ensure the full translation dictionary from previous examples is present in the code.")
            print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
            # Optionally show a basic Tkinter error
            try:
                root.withdraw() # Hide the potentially broken main window
                messagebox.showerror("Code Error", "Translation data is missing.\nPlease check the source code.", master=root)
            except: pass # Ignore if messagebox fails too
            sys.exit(1) # Exit if translations are missing

        app = DuplicateFinderApp(root)
        root.mainloop()
    except Exception as main_e:
        # --- Fatal Error Handling ---
        print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
        print("FATAL ERROR DURING APPLICATION STARTUP / MAIN LOOP:")
        print(traceback.format_exc())
        print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
        try:
            # Attempt to show a GUI error message as a last resort
            root_err = tk.Tk()
            root_err.withdraw() # Hide the empty error window
            messagebox.showerror("Fatal Error",
                                 f"A critical error occurred:\n\n{main_e}\n\nSee console log for technical details.",
                                 master=root_err)
            root_err.destroy()
        except Exception as mb_err:
            # If even Tkinter fails for the message box
            print(f"Could not display fatal error in GUI: {mb_err}")
