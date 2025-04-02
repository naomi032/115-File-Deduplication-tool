# -*- coding: utf-8 -*-
import os
import configparser
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog, Menu, Toplevel
import threading
import time
from datetime import datetime, timezone, timedelta # Added timezone, timedelta
from collections import defaultdict, Counter
import json
import traceback
import collections
import math # For size conversion
import sys # To get base path for PyInstaller
import re

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
        # Check if default sans-serif is already in the list, if not, prepend it for fallback
        if 'sans-serif' not in matplotlib.rcParams['font.sans-serif']:
             matplotlib.rcParams['font.sans-serif'].insert(0, 'sans-serif')
        # Try setting preferred fonts, prepending defaults if necessary
        current_sans_serif = matplotlib.rcParams['font.sans-serif']
        final_font_list = preferred_fonts + [f for f in current_sans_serif if f not in preferred_fonts]
        matplotlib.rcParams['font.sans-serif'] = final_font_list
        matplotlib.rcParams['axes.unicode_minus'] = False # Ensure minus sign displays correctly
        print(f"Attempting to set Matplotlib font preference: {final_font_list}")
    except Exception as font_error:
        print(f"WARNING: Could not set preferred CJK font for Matplotlib - {font_error}")
except ImportError:
    MATPLOTLIB_AVAILABLE = False
    # Inform user about missing optional dependency and feature impact
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
    # Required dependency missing, inform user and exit
    print("ERROR: The 'clouddrive' library is not installed. Please install it using: pip install clouddrive")
    try: # Attempt to show GUI error even if library is missing
        root_tk_err = tk.Tk()
        root_tk_err.withdraw()
        messagebox.showerror("Missing Library", "The 'clouddrive' library is not installed.\nPlease install it using: pip install clouddrive", master=root_tk_err)
        root_tk_err.destroy()
    except Exception:
        pass
    sys.exit("Required 'clouddrive' library not found.") # Exit cleanly


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
ICON_FILE = resource_path("app_icon.ico") # Path for icon

DATE_FORMAT = "%Y-%m-%d %H:%M:%S" # For display format
DEFAULT_LANG = "en"
# Rule constants for deletion logic
RULE_KEEP_SHORTEST = "shortest"
RULE_KEEP_LONGEST = "longest"
RULE_KEEP_OLDEST = "oldest"
RULE_KEEP_NEWEST = "newest"
RULE_KEEP_SUFFIX = "suffix"

# --- Translations ---
translations = {
    "en": {
        "window_title": "CloudDrive2 Duplicate Video Finder & Deleter",
        "config_title": "Configuration",
        "address_label": "API Address:",
        "account_label": "Account:",
        "password_label": "Password:",
        "scan_path_label": "Root Path to Scan:",
        "mount_point_label": "CloudDrive Mount Point:",
        "load_config_button": "Load Config",
        "save_config_button": "Save Config",
        "test_connection_button": "Test Connection",
        "find_button": "Find Duplicates",
        "rules_title": "Deletion Rules (Select one to mark files)", # Clarified purpose
        "rule_shortest_path": "Keep file with shortest path",
        "rule_longest_path": "Keep file with longest path",
        "rule_oldest": "Keep oldest file (Modified Date)",
        "rule_newest": "Keep newest file (Modified Date)",
        "rule_keep_suffix": "Keep file path ending with suffix:",
        "rule_suffix_entry_label": "Suffix:",
        "tree_rule_action_col": "Action",
        "tree_path_col": "File Path",
        "tree_modified_col": "Modified",
        "tree_size_col": "Size (MB)",
        "tree_set_col": "Set #",
        "tree_action_keep": "Keep",
        "tree_action_delete": "Delete",
        "tree_set_col_value": "{index}", # Simple index for set number
        "delete_by_rule_button": "Delete Marked Files",
        "save_list_button": "Save Report",
        "show_chart_button": "Show File Types Chart",
        "show_chart_button_disabled": "Show Chart (Install matplotlib)",
        "log_title": "Log",
        "menu_language": "Language",
        "menu_english": "English",
        "menu_chinese": "中文",
        "status_loading_config": "Loading config from {file}...",
        "status_saving_config": "Saving config to {file}...",
        "status_config_loaded": "Config loaded.",
        "status_config_saved": "Config saved.",
        "status_config_not_found": "Config file '{file}' not found. Using defaults/empty.",
        "status_config_section_missing": "Config file loaded, but '[config]' section is missing.",
        "status_connecting": "Attempting connection...",
        "status_connect_success": "Connection successful.",
        "status_scan_progress": "Scanned {count} items... Found {video_count} videos.",
        "status_populating_tree": "Populating list with {count} duplicate sets...",
        "status_tree_populated": "Results list populated.",
        "status_clearing_tree": "Clearing results list and rule selection...", # Updated log message
        "status_applying_rule": "Applying rule '{rule_name}' to {count} sets...",
        "status_rule_applied": "Rule applied. {delete_count} files marked for deletion.",
        "find_starting": "Starting duplicate file scan...",
        "find_complete_found": "Scan complete. Found {count} duplicate sets.",
        "find_complete_none": "Scan complete. No duplicate video files found based on SHA1 hash.",
        "find_error_during": "Error during duplicate scan: {error}",
        "delete_starting": "Starting deletion based on rule: {rule_name}...",
        "delete_finished": "Deletion complete. Attempted to delete {total_marked} files. Successfully deleted {deleted_count}.", # Clarified wording
        "delete_results_cleared": "Deletion process finished. Results list cleared, please re-scan if needed.", # Added completion message
        "delete_error_during": "Error during deletion process: {error}",
        "delete_error_file": "Error deleting {path}: {error}",
        "delete_cancelled": "Deletion cancelled by user.",
        "delete_confirm_title": "Confirm Deletion",
        "delete_confirm_msg": "Delete marked files based on rule: '{rule_name}'?\nTHIS ACTION CANNOT BE UNDONE.",
        "delete_no_rule_selected": "No deletion rule selected.",
        "delete_rule_no_files": "No files are currently marked for deletion based on the applied rule.", # Changed wording for clarity
        "delete_suffix_missing": "Suffix is required for the 'Keep Suffix' rule.",
        "save_report_header": "Duplicate Video File Sets Found (Based on SHA1):",
        "save_report_set_header": "Set {index} (SHA1: {sha1}) - {count} files",
        "save_report_file_label": "  - File:",
        "save_report_details_label": "    (Path: {path}, Modified: {modified}, Size: {size_mb:.2f} MB)",
        "save_report_saved": "Report saved successfully to: {file}",
        "save_report_error": "Error saving report to {file}: {error}",
        "save_report_no_data": "No duplicate sets found to save.",
        "error_title": "Error",
        "error_input_title": "Input Error",
        "error_config_title": "Config Error",
        "error_config_save_title": "Config Save Error",
        "error_rule_title": "Rule Error", # Added title for rule specific errors
        "error_connect": "Error connecting to CloudDrive2 API at '{address}': {error}",
        "error_scan_path": "Critical error walking cloud path '{path}': {error}",
        "error_get_attrs": "Error getting attributes/hash for '{path}': {error}",
        "error_parse_date": "Warning: Could not parse modification date for '{path}': {error}. Skipping date comparison for this file.",
        "error_no_duplicates_found": "No duplicates were found or displayed. Cannot apply deletion rule.",
        "error_not_connected": "Error: Not connected to CloudDrive2. Please test connection first.",
        "error_path_calc_failed": "Error: Could not determine a valid cloud scan path from Root Path and Mount Point. Check inputs.",
        "error_input_missing": "API Address, Account, Root Path to Scan, and Mount Point are required.", # Updated required fields text
        "error_input_missing_conn": "API Address, Account, and Mount Point are required for connection test.",
        "error_input_missing_chart": "Root Path to Scan and Mount Point are required to generate the chart.",
        "error_config_read": "Error reading config file: {error}",
        "error_config_save": "Could not write config file: {error}",
        "error_unexpected": "Unexpected error: {error}",
        "error_icon_load": "Error loading application icon '{path}': {error}",
        "warning_path_mismatch": "Warning: Could not determine a valid cloud path based on 'Root Path to Scan' ('{scan}') and 'Mount Point' ('{mount}'). Effective scan path might be '/'. Please verify inputs.",
        "path_warning_title": "Path Input Warning",
        "path_warning_suspicious_chars": "Suspicious character(s) detected in input paths!\nThis often happens from copy-pasting.\nPlease DELETE and MANUALLY RETYPE the paths in the GUI.",
        "conn_test_success_title": "Connection Test Successful",
        "conn_test_success_msg": "Successfully connected to CloudDrive2.",
        "conn_test_fail_title": "Connection Test Failed",
        "conn_test_fail_msg": "Failed to connect. Please check API address, credentials, and ensure CloudDrive2 is running. See log for details.",
        "chart_error_title": "Chart Error",
        "chart_info_title": "Chart Info",
        "chart_error_no_matplotlib": "The 'matplotlib' library is required for charting but not found. Please install it (`pip install matplotlib`).",
        "chart_error_no_connection": "Cannot generate chart: Not connected to CloudDrive2.",
        "chart_error_cloud_scan": "Error scanning cloud path '{path}' for chart data: {error}",
        "chart_status_scanning_cloud": "Scanning cloud path '{path}' for file types (this may take a while)...",
        "chart_status_generating": "Scan complete. Generating chart for {count} file types ({total} files)...",
        "chart_status_no_files_found": "Scan complete. No files found in '{path}'. Cannot generate chart.",
        "chart_window_title": "File Types in: {path}",
        "chart_legend_title": "File Extensions",
        "chart_label_others": "Others",
        "chart_label_no_extension": "[No Ext]",
        "tie_break_log_prefix": "Tie-Break:", # Added for rule tie-break logging
        # --- Added/Modified for Finder logging ---
        "status_test_connection_step": "Testing connection by attempting to list root directory ('/')...",
        "status_scan_finished_duration": "Scan finished in {duration:.2f} seconds.",
        "status_scan_summary_items": "Total items encountered: {count}. Video files processed: {video_count}.",
        "status_scan_warnings": "WARNING: {details}.",
        "warning_hash_missing": "Warning: Hash data missing in attributes for '{path}'. KeyError: {key_error}. Skipping.",
        "warning_hash_short": "Warning: Suspiciously short SHA1 hash ('{hash}') found for '{path}'. Skipping.",
        "warning_size_invalid": "Warning: Invalid size value '{size}' for {path}. Using 0.",
        "status_delete_attempting": "Attempting to delete {count} marked files...",
        "status_deleting_file": "Deleting [{current}/{total}]: {path}",
        "warning_delete_failures": "WARNING: Failed to delete {count} file(s):",
        "warning_delete_failures_more": "  ... and {count} more.",
        "status_delete_no_files": "No files provided for deletion.",
        "warning_tie_break": "{prefix} {reason}. Kept '{filename}' ({detail})",
        "warning_rule_no_date": "Warning: {set_id} - Cannot apply '{rule}': No valid dates. Keeping shortest path.",
        "warning_rule_no_suffix_match": "Warning: {set_id} - No files match suffix '{suffix}'. Keeping shortest path.",
        "warning_rule_failed_selection": "Internal Warning: {set_id} - Rule '{rule}' failed to select file to keep. Skipping deletion for this set.",
        "error_rule_application": "Error applying rule '{rule}' to {set_id}: {error}. Skipping deletion for this set.",
        "log_debug_calc_path": "[Debug] Calculated effective cloud scan path: '{fs_path}' from Scan='{scan_raw}', Mount='{mount_raw}'",
    },
    "zh": {
        "window_title": "CloudDrive2 重复视频查找与删除工具",
        "config_title": "配置",
        "address_label": "API 地址:",
        "account_label": "账户:",
        "password_label": "密码:",
        "scan_path_label": "要扫描的根路径:",
        "mount_point_label": "CloudDrive 挂载点:",
        "load_config_button": "加载配置",
        "save_config_button": "保存配置",
        "test_connection_button": "测试连接",
        "find_button": "查找重复项",
        "rules_title": "删除规则 (选择一项以标记文件)", # 澄清目的
        "rule_shortest_path": "保留路径最短的文件",
        "rule_longest_path": "保留路径最长的文件",
        "rule_oldest": "保留最旧的文件 (修改日期)",
        "rule_newest": "保留最新的文件 (修改日期)",
        "rule_keep_suffix": "保留路径以此后缀结尾的文件:",
        "rule_suffix_entry_label": "后缀:",
        "tree_rule_action_col": "操作",
        "tree_path_col": "文件路径",
        "tree_modified_col": "修改时间",
        "tree_size_col": "大小 (MB)",
        "tree_set_col": "集合",
        "tree_action_keep": "保留",
        "tree_action_delete": "删除",
        "tree_set_col_value": "{index}", # 集合编号
        "delete_by_rule_button": "删除标记文件",
        "save_list_button": "保存报告",
        "show_chart_button": "显示文件类型图表",
        "show_chart_button_disabled": "显示图表 (需安装 matplotlib)",
        "log_title": "日志",
        "menu_language": "语言",
        "menu_english": "English",
        "menu_chinese": "中文",
        "status_loading_config": "正在从 {file} 加载配置...",
        "status_saving_config": "正在保存配置到 {file}...",
        "status_config_loaded": "配置已加载。",
        "status_config_saved": "配置已保存。",
        "status_config_not_found": "未找到配置文件 '{file}'。将使用默认/空值。",
        "status_config_section_missing": "配置文件已加载，但缺少 '[config]' 部分。",
        "status_connecting": "正在尝试连接...",
        "status_connect_success": "连接成功。",
        "status_scan_progress": "已扫描 {count} 个项目... 找到 {video_count} 个视频。",
        "status_populating_tree": "正在使用 {count} 个重复集合填充列表...",
        "status_tree_populated": "结果列表已填充。",
        "status_clearing_tree": "正在清除结果列表和规则选择...", # 更新
        "status_applying_rule": "正在对 {count} 个集合应用规则 '{rule_name}'...",
        "status_rule_applied": "规则已应用。{delete_count} 个文件被标记为删除。",
        "find_starting": "开始扫描重复文件...",
        "find_complete_found": "扫描完成。找到 {count} 个重复集合。",
        "find_complete_none": "扫描完成。未根据 SHA1 哈希找到重复的视频文件。",
        "find_error_during": "扫描重复项期间出错: {error}",
        "delete_starting": "开始根据规则删除: {rule_name}...",
        "delete_finished": "删除完成。尝试删除 {total_marked} 个文件。成功删除了 {deleted_count} 个。", # 澄清措辞
        "delete_results_cleared": "删除过程结束。结果列表已清除，如需请重新扫描。", # 新增
        "delete_error_during": "删除过程中出错: {error}",
        "delete_error_file": "删除 {path} 时出错: {error}",
        "delete_cancelled": "用户取消了删除操作。",
        "delete_confirm_title": "确认删除",
        "delete_confirm_msg": "根据规则 '{rule_name}' 删除标记的文件吗？\n此操作无法撤销。",
        "delete_no_rule_selected": "未选择删除规则。",
        "delete_rule_no_files": "根据当前应用的规则，没有文件被标记为删除。", # 更改措辞
        "delete_suffix_missing": "“保留后缀”规则需要填写后缀。",
        "save_report_header": "找到的重复视频文件集合 (基于 SHA1):",
        "save_report_set_header": "集合 {index} (SHA1: {sha1}) - {count} 个文件",
        "save_report_file_label": "  - 文件:",
        "save_report_details_label": "    (路径: {path}, 修改时间: {modified}, 大小: {size_mb:.2f} MB)",
        "save_report_saved": "报告已成功保存到: {file}",
        "save_report_error": "保存报告到 {file} 时出错: {error}",
        "save_report_no_data": "未找到可保存的重复文件集合。",
        "error_title": "错误",
        "error_input_title": "输入错误",
        "error_config_title": "配置错误",
        "error_config_save_title": "配置保存错误",
        "error_rule_title": "规则错误", # 新增
        "error_connect": "连接 CloudDrive2 API '{address}' 时出错: {error}",
        "error_scan_path": "遍历云端路径 '{path}' 时发生严重错误: {error}",
        "error_get_attrs": "获取 '{path}' 的属性/哈希时出错: {error}",
        "error_parse_date": "警告: 无法解析 '{path}' 的修改日期: {error}。将跳过此文件的日期比较。",
        "error_no_duplicates_found": "未找到或显示重复项。无法应用删除规则。",
        "error_not_connected": "错误：未连接到 CloudDrive2。请先测试连接。",
        "error_path_calc_failed": "错误：无法根据扫描根路径和挂载点确定有效的云端扫描路径。请检查输入。",
        "error_input_missing": "API 地址、账户、要扫描的根路径和挂载点为必填项。", # 更新文本
        "error_input_missing_conn": "测试连接需要 API 地址、账户和挂载点。",
        "error_input_missing_chart": "生成图表需要“要扫描的根路径”和“挂载点”。",
        "error_config_read": "读取配置文件时出错: {error}",
        "error_config_save": "无法写入配置文件: {error}",
        "error_unexpected": "意外错误: {error}",
        "error_icon_load": "加载应用程序图标 '{path}' 时出错: {error}",
        "warning_path_mismatch": "警告：无法根据“要扫描的根路径” ('{scan}') 和“挂载点” ('{mount}') 确定有效的云端路径。有效扫描路径可能是'/'。请核对输入。",
        "path_warning_title": "路径输入警告",
        "path_warning_suspicious_chars": "在输入路径中检测到可疑字符！\n这通常是复制粘贴造成的。\n请在图形界面中删除并手动重新输入路径。",
        "conn_test_success_title": "连接测试成功",
        "conn_test_success_msg": "已成功连接到 CloudDrive2。",
        "conn_test_fail_title": "连接测试失败",
        "conn_test_fail_msg": "连接失败。请检查 API 地址、凭据并确保 CloudDrive2 正在运行。详情请查看日志。",
        "chart_error_title": "图表错误",
        "chart_info_title": "图表信息",
        "chart_error_no_matplotlib": "生成图表需要 'matplotlib' 库，但未找到。请安装它 (`pip install matplotlib`)。",
        "chart_error_no_connection": "无法生成图表：未连接到 CloudDrive2。",
        "chart_error_cloud_scan": "扫描云端路径 '{path}' 以获取图表数据时出错: {error}",
        "chart_status_scanning_cloud": "正在扫描云端路径 '{path}' 以获取文件类型 (可能需要一些时间)...",
        "chart_status_generating": "扫描完成。正在为 {count} 种文件类型 ({total} 个文件) 生成图表...",
        "chart_status_no_files_found": "扫描完成。在 '{path}' 中未找到文件。无法生成图表。",
        "chart_window_title": "文件类型分布: {path}",
        "chart_legend_title": "文件扩展名",
        "chart_label_others": "其他",
        "chart_label_no_extension": "[无扩展名]",
        "tie_break_log_prefix": "规则冲突解决:", # 新增，用于规则冲突日志
        # --- Added/Modified for Finder logging ---
        "status_test_connection_step": "正在通过尝试列出根目录 ('/') 来测试连接...",
        "status_scan_finished_duration": "扫描耗时 {duration:.2f} 秒完成。",
        "status_scan_summary_items": "共遇到 {count} 个项目。已处理 {video_count} 个视频文件。",
        "status_scan_warnings": "警告: {details}。",
        "warning_hash_missing": "警告：'{path}' 的属性中缺少哈希数据。KeyError: {key_error}。正在跳过。",
        "warning_hash_short": "警告：为 '{path}' 找到了可疑的短 SHA1 哈希 ('{hash}')。正在跳过。",
        "warning_size_invalid": "警告：{path} 的大小值 '{size}' 无效。使用 0。",
        "status_delete_attempting": "正在尝试删除 {count} 个标记的文件...",
        "status_deleting_file": "正在删除 [{current}/{total}]: {path}",
        "warning_delete_failures": "警告：未能删除 {count} 个文件：",
        "warning_delete_failures_more": "  ... 以及另外 {count} 个。",
        "status_delete_no_files": "没有提供用于删除的文件。",
        "warning_tie_break": "{prefix} {reason}。保留了 '{filename}' ({detail})",
        "warning_rule_no_date": "警告：{set_id} - 无法应用 '{rule}'：无有效日期。保留最短路径。",
        "warning_rule_no_suffix_match": "警告：{set_id} - 没有文件匹配后缀 '{suffix}'。保留最短路径。",
        "warning_rule_failed_selection": "内部警告：{set_id} - 规则 '{rule}' 未能选择要保留的文件。跳过此集合的删除。",
        "error_rule_application": "将规则 '{rule}' 应用于 {set_id} 时出错：{error}。跳过此集合的删除。",
        "log_debug_calc_path": "[调试] 根据 Scan='{scan_raw}', Mount='{mount_raw}' 计算出的有效云扫描路径: '{fs_path}'",
    }
}


# --- Helper Functions ---
def _validate_path_chars(path_str):
    """Checks a single path string for suspicious characters often causing issues."""
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
        # Add others if needed (e.g., specific formatting chars)
    }
    for i, char in enumerate(path_str):
        char_code = ord(char)
        is_suspicious = False
        reason = ""
        # C0 Controls (ASCII 0-31) and DEL (127) - Generally bad in paths
        if 0 <= char_code <= 31 or char_code == 127:
             # Allow standard space (32), but most others are problematic
             is_suspicious = True
             reason = f"Control Char (C0/DEL {char_code})"
        # C1 Controls (often appear from bad encoding/copy-paste: 128-159)
        elif 128 <= char_code <= 159:
            is_suspicious = True
            reason = f"Control Char (C1 {char_code})"
        # Specific known problematic Unicode chars
        elif char_code in KNOWN_INVISIBLE_CODES:
            is_suspicious = True
            reason = f"Known Invisible/Problematic Char (U+{char_code:04X})"
        # Add more checks if needed (e.g., filesystem specific invalid chars)

        if is_suspicious:
            suspicious_codes.append(f"'{char}' (Code {char_code} / {reason}) at position {i}")
    return suspicious_codes

def _build_full_path(parent_path, item_name):
    """Helper to correctly join cloud paths using forward slashes."""
    # Normalize slashes and remove trailing/leading ones inappropriately
    parent_path_norm = parent_path.replace('\\', '/').rstrip('/')
    item_name_norm = item_name.replace('\\', '/').lstrip('/')
    # Handle root case
    if not parent_path_norm or parent_path_norm == '/':
        # Avoid double slash if item name already starts with /
        return '/' + item_name_norm.lstrip('/')
    else:
        return parent_path_norm + '/' + item_name_norm

def _parse_datetime(date_string):
    """
    Parses common datetime string formats from CloudDrive into timezone-aware datetime objects.
    Attempts multiple formats and defaults to UTC if no timezone info is present.
    Returns None if parsing fails.
    """
    if not date_string or not isinstance(date_string, str):
        return None

    original_date_string = date_string # For logging on failure

    # 1. ISO 8601 format with 'Z' (Zulu time = UTC)
    if date_string.endswith('Z'):
        try:
            # Replace Z with +00:00 for standard parsing
            dt = datetime.fromisoformat(date_string.replace('Z', '+00:00'))
            return dt # Already timezone-aware
        except ValueError:
            pass # Try other formats

    # 2. ISO 8601 format with explicit offset (+HH:MM, +HHMM, -HH:MM, -HHMM)
    # Check if last part looks like an offset
    if len(date_string) > 5 and date_string[-6] in '+-' and date_string[-3] == ':' or \
       len(date_string) > 4 and date_string[-5] in '+-' and date_string[-3] != ':': # Handle +-HHMM
        try:
            # Python 3.7+ handles common ISO formats including offsets directly
            # Need to handle potential lack of colon in offset for fromisoformat
            if len(date_string) >= 5 and date_string[-5] in "+-" and date_string[-3] != ':':
                 # Insert colon: 2023-10-27T08:30:00-0500 -> 2023-10-27T08:30:00-05:00
                 date_string_with_colon = date_string[:-2] + ":" + date_string[-2:]
                 dt = datetime.fromisoformat(date_string_with_colon)
                 return dt
            else:
                dt = datetime.fromisoformat(date_string)
                return dt # Already timezone-aware
        except ValueError:
            pass # Try manual parsing below

    # 3. Attempt manual parsing for common formats, including potential missing offset colon
    formats_to_try = [
        '%Y-%m-%dT%H:%M:%S.%f%z',  # With microseconds and offset (no colon handled by strptime in 3.7+)
        '%Y-%m-%dT%H:%M:%S%z',     # Without microseconds and offset (no colon handled by strptime in 3.7+)
        '%Y-%m-%dT%H:%M:%S.%f',    # With microseconds, NO offset (assume UTC)
        '%Y-%m-%dT%H:%M:%S',       # Without microseconds, NO offset (assume UTC)
        '%Y-%m-%d %H:%M:%S.%f%z',  # Space separator variant
        '%Y-%m-%d %H:%M:%S%z',     # Space separator variant
        '%Y-%m-%d %H:%M:%S.%f',    # Space separator variant (assume UTC)
        '%Y-%m-%d %H:%M:%S',       # Space separator variant (assume UTC)
    ]

    parsed_dt = None
    # Try parsing directly first
    for fmt in formats_to_try:
        try:
            parsed_dt = datetime.strptime(date_string, fmt)
            break
        except ValueError:
            continue # Try next format

    if parsed_dt:
        # If parsing succeeded but no timezone info was derived, assume UTC
        if parsed_dt.tzinfo is None or parsed_dt.tzinfo.utcoffset(parsed_dt) is None:
            return parsed_dt.replace(tzinfo=timezone.utc)
        else:
            # Timezone info was parsed correctly
            return parsed_dt

    # --- Fallback if all parsing fails ---
    # Warning is handled by the caller function now
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
        # Default translator returns key if not found
        self._ = lambda key, **kwargs: kwargs.get('default', f"<{key}?>")

    def set_translator(self, translator_func):
        """Sets the translation function to be used for logging."""
        self._ = translator_func

    def log(self, message):
        """ Sends message to the registered progress callback (GUI logger). """
        if self.progress_callback:
            try:
                message_str = str(message) if message is not None else ""
                self.progress_callback(message_str)
            except Exception as e:
                # Avoid crashing the app if the callback fails
                print(f"Error in progress_callback: {e}")
                print(f"Original message: {message}")
        else:
            # Fallback to console if no callback registered
            print(str(message))

    def set_config(
            self,
            clouddrvie2_address,
            clouddrive2_account,
            clouddrive2_passwd,
            raw_scan_path, # Keep raw inputs for context
            raw_mount_point,
            progress_callback=None,
    ):
        """ Sets configuration and attempts to establish+test connection. """
        self.clouddrvie2_address = clouddrvie2_address
        self.clouddrive2_account = clouddrive2_account
        self.clouddrive2_passwd = clouddrive2_passwd
        # Store raw paths as provided by user
        self._raw_scan_path = raw_scan_path
        self._raw_mount_point = raw_mount_point
        self.progress_callback = progress_callback
        self.fs = None # Reset filesystem object on new config/connection attempt

        # Basic Input Validation
        if not self.clouddrvie2_address:
             self.log(self._("error_connect", address="<empty>", error="API Address cannot be empty.", default="Connection Error: API Address missing."))
             return False

        try:
            self.log(self._("status_connecting", default="Attempting connection..."))
            client = CloudDriveClient(
                self.clouddrvie2_address, self.clouddrive2_account, self.clouddrive2_passwd
            )
            # Wrap filesystem creation in try-except as well
            try:
                 self.fs = CloudDriveFileSystem(client)
            except Exception as fs_init_e:
                 error_msg = self._("error_connect", address=self.clouddrvie2_address, error=f"Failed to initialize filesystem: {fs_init_e}", default=f"Connection Error: Filesystem init failed: {fs_init_e}")
                 self.log(error_msg)
                 self.log(f"Filesystem Init Error Details: {traceback.format_exc()}")
                 self.fs = None
                 return False

            # Test connection by attempting a basic operation (listing root)
            self.log(self._("status_test_connection_step", default="Testing connection by attempting to list root directory ('/')..."))
            self.fs.ls('/') # Raises exception on failure
            self.log(self._("status_connect_success", default="Connection successful."))
            return True

        except Exception as e:
            # Catch errors from CloudDriveClient init or fs.ls('/')
            error_msg = self._("error_connect", address=self.clouddrvie2_address, error=e, default=f"Error connecting to {self.clouddrvie2_address}: {e}")
            self.log(error_msg)
            # Log detailed traceback for debugging
            self.log(f"Connection Error Details: {traceback.format_exc()}")
            self.fs = None # Ensure fs is None on error
            return False

    def calculate_fs_path(self, scan_path_raw, mount_point_raw):
        """
        Calculates the effective cloud filesystem path based on scan root and mount point.
        Handles various path formats and normalizes to a forward-slash separated path.
        Returns the calculated cloud path (e.g., '/movies') or None if calculation fails.
        Logs warnings or errors via self.log().
        """
        # --- Input Validation (Character Check) ---
        # Check raw inputs before normalization
        scan_path_issues = _validate_path_chars(scan_path_raw)
        mount_point_issues = _validate_path_chars(mount_point_raw)
        if scan_path_issues or mount_point_issues:
            all_issues = []
            scan_path_label = self._("scan_path_label", default="Scan Path").rstrip(': ')
            mount_point_label = self._("mount_point_label", default="Mount Point").rstrip(': ')
            if scan_path_issues: all_issues.append(f"'{scan_path_label}' ('{scan_path_raw}'): {', '.join(scan_path_issues)}")
            if mount_point_issues: all_issues.append(f"'{mount_point_label}' ('{mount_point_raw}'): {', '.join(mount_point_issues)}")
            log_msg = self._("path_warning_suspicious_chars", default="Suspicious chars detected!").split('\n')[0] # Get first line
            self.log(f"ERROR: {log_msg} Details: {'; '.join(all_issues)}")
            return None # Indicate failure due to bad characters

        # --- Normalization ---
        # Replace backslashes, strip whitespace, remove trailing slashes
        scan_path_norm = scan_path_raw.replace('\\', '/').strip().rstrip('/')
        mount_point_norm = mount_point_raw.replace('\\', '/').strip().rstrip('/')
        fs_dir_path = None # Initialize result

        # --- Core Logic ---
        # Case 1: Mount point looks like a drive letter (e.g., "X:")
        if len(mount_point_norm) == 2 and mount_point_norm[1] == ':' and mount_point_norm[0].isalpha():
            mount_point_drive_prefix = mount_point_norm.lower() + '/'
            scan_path_lower = scan_path_norm.lower()
            if scan_path_lower.startswith(mount_point_drive_prefix):
                # Scan path is inside the mounted drive, get relative part
                relative_part = scan_path_norm[len(mount_point_norm):]
                fs_dir_path = '/' + relative_part.lstrip('/')
            elif scan_path_lower == mount_point_norm.lower():
                # Scan path *is* the mount point drive itself (represents cloud root)
                 fs_dir_path = '/'
            else:
                 # Scan path is not under the mount point drive letter (mismatch)
                 fs_dir_path = None

        # Case 2: Mount point is empty, '/', or just whitespace (effectively root)
        elif not mount_point_norm or mount_point_norm == '/':
            # Treat scan path as relative to the cloud root
            fs_dir_path = '/' + scan_path_norm.lstrip('/')

        # Case 3: Mount point is an absolute path (starts with '/' or drive letter 'X:/')
        # Assumes the mount point represents a *subfolder* within the cloud drive's root mapping.
        elif mount_point_norm.startswith('/') or (len(mount_point_norm) > 1 and mount_point_norm[1] == ':'):
             # If mount point looks like 'X:/', normalize prefix to 'x:/' for comparison
             prefix_to_check = mount_point_norm + '/'
             scan_path_to_check = scan_path_norm
             if len(mount_point_norm) > 1 and mount_point_norm[1] == ':':
                 prefix_to_check = mount_point_norm.lower() + '/'
                 scan_path_to_check = scan_path_norm.lower()

             if scan_path_to_check.startswith(prefix_to_check):
                 # Scan path is deeper than mount point
                 relative_part = scan_path_norm[len(mount_point_norm):]
                 fs_dir_path = '/' + relative_part.lstrip('/')
             elif scan_path_to_check == prefix_to_check.rstrip('/'): # Compare against 'x:' or '/path'
                 # Scan path is exactly the mount point itself (represents cloud root in this context)
                 fs_dir_path = '/'
             else:
                 # Scan path does NOT start with mount_point_prefix (mismatch)
                 fs_dir_path = None

        # Case 4: Mount point is a relative path (e.g., "mycloud", "data/share")
        # Assumes this name corresponds to a folder directly under the cloud root.
        elif '/' not in mount_point_norm and ':' not in mount_point_norm:
            # Check if scan path starts with mount point OR /mount point
            mount_prefix_slash = '/' + mount_point_norm
            mount_prefix_slash_sep = mount_prefix_slash + '/'
            mount_prefix_sep = mount_point_norm + '/'

            if scan_path_norm.startswith(mount_prefix_slash_sep):
                relative_part = scan_path_norm[len(mount_prefix_slash):]
                fs_dir_path = '/' + relative_part.lstrip('/')
            elif scan_path_norm == mount_prefix_slash:
                 fs_dir_path = '/'
            elif scan_path_norm.startswith(mount_prefix_sep):
                 # Treat as relative to root if scan starts like 'mycloud/...'
                 relative_part = scan_path_norm[len(mount_point_norm):]
                 fs_dir_path = '/' + relative_part.lstrip('/')
            elif scan_path_norm == mount_point_norm:
                 # Treat scan path 'mycloud' as root if mount point is 'mycloud'
                 fs_dir_path = '/'
            else:
                # Mismatch
                fs_dir_path = None
        else:
            # Fallback / unhandled case
            fs_dir_path = None


        # --- Result Handling ---
        if fs_dir_path is None:
            # Path calculation failed (likely a mismatch)
            warning_msg = self._("warning_path_mismatch",
                                      scan=scan_path_raw, mount=mount_point_raw,
                                      default=f"Warning: Could not determine cloud path from Scan Path ('{scan_path_raw}') and Mount Point ('{mount_point_raw}'). Check inputs. Assuming '/' as scan path.")
            self.log(warning_msg)
            # Return None to signal the GUI/caller about the failure.
            return None

        # Final normalization: ensure single slashes, starts with '/', doesn't end with / unless it's just '/'
        if fs_dir_path:
            while '//' in fs_dir_path: fs_dir_path = fs_dir_path.replace('//', '/')
            if not fs_dir_path.startswith('/'): fs_dir_path = '/' + fs_dir_path
            if len(fs_dir_path) > 1: fs_dir_path = fs_dir_path.rstrip('/')
        elif not fs_dir_path: # Handle case where inputs led to empty string
            fs_dir_path = '/'

        # Log the calculated path for debugging/confirmation
        self.log(self._("log_debug_calc_path",
                        fs_path=fs_dir_path, scan_raw=scan_path_raw, mount_raw=mount_point_raw,
                        default=f"[Debug] Calculated effective cloud scan path: '{fs_dir_path}' from Scan='{scan_path_raw}', Mount='{mount_point_raw}'"))
        return fs_dir_path


    def find_duplicates(self):
        """
        Scans the configured cloud path for duplicate video files using SHA1 hash.
        """
        if not self.fs:
            self.log(self._("error_not_connected", default="Error: Not connected to CloudDrive. Cannot scan."))
            return {}

        fs_dir_path = self.calculate_fs_path(self._raw_scan_path, self._raw_mount_point)
        if fs_dir_path is None:
            self.log(
                self._("error_path_calc_failed", default="Error: Could not determine cloud scan path. Aborting scan."))
            return {}

        self.log(self._("find_starting", default="Starting duplicate file scan...") + f" (Path: '{fs_dir_path}')")
        start_time = time.time()

        potential_duplicates = defaultdict(list)
        count = 0
        video_files_checked = 0
        errors_getting_attrs = 0
        files_skipped_no_sha1 = 0
        key_errors_getting_hash = 0  # Counter for specific KeyError on hash lookup

        try:
            walk_iterator = self.fs.walk_path(fs_dir_path)

            for foldername, _, filenames in walk_iterator:
                # Ensure foldername is treated as a string path segment
                foldername_str = str(foldername)

                for filename_obj in filenames:
                    count += 1
                    filename_str = str(filename_obj)

                    # Log progress periodically
                    if count % 500 == 0:
                        self.log(self._("status_scan_progress", count=count, video_count=video_files_checked,
                                        default=f"Scanned {count} items... Found {video_files_checked} videos."))

                    # Check if it's a video file based on extension
                    file_extension = os.path.splitext(filename_str)[1].lower()
                    if file_extension in VIDEO_EXTENSIONS:
                        video_files_checked += 1

                        # --- Construct Full Path Correctly ---
                        # Use helper function for proper cloud path joining
                        path_for_storage = _build_full_path(foldername_str, filename_str)

                        attrs = None
                        mod_time_dt = None
                        file_size = 0
                        file_sha1 = None

                        try:
                            # --- Call fs.attr with the normalized path ---
                            attrs = self.fs.attr(path_for_storage)

                            # --- Extract SHA1 Hash with KeyError Handling ---
                            try:
                                file_hashes_dict = attrs.get('fileHashes')
                                if isinstance(file_hashes_dict, dict):
                                    file_sha1 = file_hashes_dict.get('2')  # Key '2' for SHA1
                                    # Validate SHA1 format/presence
                                    if not isinstance(file_sha1, str) or not file_sha1:
                                        file_sha1 = None
                                    elif len(file_sha1) < 40: # Basic sanity check
                                        self.log(self._("warning_hash_short",
                                                         hash=file_sha1, path=path_for_storage,
                                                         default=f"Warning: Suspiciously short SHA1 hash ('{file_sha1}') for '{path_for_storage}'. Skipping."))
                                        file_sha1 = None
                                else:
                                    file_sha1 = None

                            except KeyError as ke:
                                # Specifically catch KeyError if 'fileHashes' or '2' is missing
                                if key_errors_getting_hash < 10 or key_errors_getting_hash % 10 == 0: # Avoid flooding logs
                                    self.log(self._("warning_hash_missing",
                                                      path=path_for_storage, key_error=ke,
                                                      default=f"Warning: Hash data missing for '{path_for_storage}'. KeyError: {ke}. Skipping."))
                                key_errors_getting_hash += 1
                                errors_getting_attrs += 1
                                file_sha1 = None

                            if not file_sha1:
                                files_skipped_no_sha1 += 1
                                continue # Skip this file if no valid SHA1

                            # --- Extract and Parse Modification Time ---
                            mod_time_str = attrs.get('modifiedTime')
                            mod_time_dt = _parse_datetime(mod_time_str)
                            if mod_time_dt is None and mod_time_str:
                                self.log(self._("error_parse_date", path=path_for_storage,
                                                error=f"Unparseable string '{mod_time_str}'",
                                                default=f"Warning: Could not parse date '{mod_time_str}' for {path_for_storage}"))
                                # Do not increment errors_getting_attrs here, only parse error

                            # --- Extract Size ---
                            size_val = attrs.get('size', 0)
                            try:
                                file_size = int(size_val) if size_val is not None else 0
                            except (ValueError, TypeError):
                                self.log(self._("warning_size_invalid",
                                                  size=size_val, path=path_for_storage,
                                                  default=f"Warning: Invalid size value '{size_val}' for {path_for_storage}. Using 0."))
                                file_size = 0
                                # Do not increment errors_getting_attrs here, only size error

                            # --- Store File Info ---
                            file_info = {
                                'path': path_for_storage, # Store the normalized path
                                'modified': mod_time_dt,
                                'size': file_size,
                                'sha1': file_sha1
                            }
                            potential_duplicates[file_sha1].append(file_info)

                        except FileNotFoundError as fnf_e:
                            # Log using the normalized path for consistency
                            err_msg = self._("error_get_attrs", path=path_for_storage, error=fnf_e,
                                             default=f"Error getting attributes for {path_for_storage}: {fnf_e}")
                            self.log(err_msg)
                            errors_getting_attrs += 1
                            # Log gRPC details if available
                            if hasattr(fnf_e, 'args') and len(fnf_e.args) > 1:
                                grpc_details = str(fnf_e.args[1])
                                self.log(f"    FNF Detail: {grpc_details}")
                        except Exception as e:
                            # Log using the normalized path for consistency in logs
                            err_msg = self._("error_get_attrs", path=path_for_storage, error=e,
                                             default=f"Error getting attributes for {path_for_storage}: {e}")
                            self.log(err_msg)
                            self.log(f"Attribute Error Details: {traceback.format_exc(limit=2)}")
                            errors_getting_attrs += 1

            # --- Scan finished ---
            end_time = time.time()
            duration = end_time - start_time
            self.log(self._("status_scan_finished_duration", duration=duration, default=f"Scan finished in {duration:.2f} seconds."))
            self.log(self._("status_scan_summary_items", count=count, video_count=video_files_checked, default=f"Total items: {count}. Videos processed: {video_files_checked}."))

            # Report errors/skips encountered
            warning_parts = []
            if errors_getting_attrs > 0: warning_parts.append(
                f"Encountered {errors_getting_attrs} errors retrieving attributes")
            if key_errors_getting_hash > 0: warning_parts.append(f"{key_errors_getting_hash} files missing hash data")
            if files_skipped_no_sha1 > key_errors_getting_hash:
                other_skips = files_skipped_no_sha1 - key_errors_getting_hash
                warning_parts.append(f"{other_skips} skipped for other SHA1 reasons")

            if warning_parts:
                self.log(self._("status_scan_warnings", details='; '.join(warning_parts), default=f"WARNING: {'; '.join(warning_parts)}."))

            # --- Filter for actual duplicates ---
            actual_duplicates = {sha1: files for sha1, files in potential_duplicates.items() if len(files) > 1}

            # Report findings
            if actual_duplicates:
                num_sets = len(actual_duplicates)
                num_files = sum(len(files) for files in actual_duplicates.values())
                self.log(self._("find_complete_found", count=num_sets,
                                default=f"Found {num_sets} duplicate sets ({num_files} total duplicate files)."))
            else:
                no_dups_msg = self._("find_complete_none",
                                     default="Scan complete. No duplicate video files found based on SHA1 hash.")
                if files_skipped_no_sha1 > 0:
                    no_dups_msg += f" (Note: {files_skipped_no_sha1} video file(s) were skipped due to missing/invalid SHA1 hash.)"
                self.log(no_dups_msg)

            return actual_duplicates

        except Exception as walk_e:
            # Catch errors during the fs.walk_path() iteration itself
            err_msg = self._("error_scan_path", path=fs_dir_path, error=walk_e,
                             default=f"Critical error walking cloud path '{fs_dir_path}': {walk_e}")
            self.log(err_msg)
            self.log(f"Walk Error Details: {traceback.format_exc()}")
            return {}

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
                # Sort sets by SHA1 for consistent report order
                sorted_sha1s = sorted(duplicate_sets.keys())

                for sha1 in sorted_sha1s:
                    files_in_set = duplicate_sets[sha1]
                    # Basic check
                    if not files_in_set or len(files_in_set) < 2: continue

                    set_count += 1
                    set_header = self._("save_report_set_header", index=set_count, sha1=sha1, count=len(files_in_set), default=f"Set {set_count} (SHA1: {sha1}) - {len(files_in_set)} files")
                    f.write(f"{set_header}\n")
                    # Sort files within the set by path for readability
                    sorted_files = sorted(files_in_set, key=lambda item: item.get('path', ''))

                    for file_info in sorted_files:
                         file_label = self._("save_report_file_label", default="  - File:")
                         mod_time_obj = file_info.get('modified')
                         # Format datetime object or use N/A
                         mod_time_str = mod_time_obj.strftime(DATE_FORMAT) if isinstance(mod_time_obj, datetime) else "N/A"
                         size_bytes = file_info.get('size')
                         # Calculate size safely
                         size_mb = size_bytes / (1024 * 1024) if isinstance(size_bytes, (int, float)) and size_bytes > 0 else 0.0
                         details_label = self._("save_report_details_label",
                                                path=file_info.get('path', 'N/A'),
                                                modified=mod_time_str,
                                                size_mb=size_mb,
                                                default=f"    (Path: {file_info.get('path', 'N/A')}, Modified: {mod_time_str}, Size: {size_mb:.2f} MB)")
                         f.write(f"{file_label}\n{details_label}\n")
                    f.write("\n") # Blank line between sets

            save_msg = self._("save_report_saved", file=output_file, default=f"Report saved successfully to: {output_file}")
            self.log(save_msg)
            return True
        except IOError as ioe:
             error_msg = self._("save_report_error", file=output_file, error=ioe, default=f"Error saving report to {output_file}: {ioe}")
             self.log(error_msg)
             self.log(f"Report Save IO Error Details: {traceback.format_exc()}")
             return False
        except Exception as e:
             error_msg = self._("save_report_error", file=output_file, error=e, default=f"Unexpected error saving report to {output_file}: {e}")
             self.log(error_msg)
             self.log(f"Report Save Unexpected Error Details: {traceback.format_exc()}")
             return False


    def delete_files(self, files_to_delete):
        """
        Deletes a list of file paths from the cloud drive.
        Returns tuple: (deleted_count, total_attempted).
        Logs progress and errors using the translator.
        """
        if not self.fs:
            self.log(self._("error_not_connected", default="Error: Not connected to CloudDrive. Cannot delete files."))
            return 0, 0 # Return zero counts

        deleted_count = 0
        total_to_delete = len(files_to_delete)
        errors_deleting = [] # Store paths that failed to delete

        if total_to_delete == 0:
            self.log(self._("status_delete_no_files", default="No files provided for deletion."))
            return 0, 0

        self.log(self._("status_delete_attempting", count=total_to_delete, default=f"Attempting to delete {total_to_delete} marked files..."))

        for i, file_path in enumerate(files_to_delete):
            # Ensure forward slashes for the API call
            cloud_path = file_path.replace('\\', '/')
            self.log(self._("status_deleting_file", current=i+1, total=total_to_delete, path=cloud_path, default=f"Deleting [{i+1}/{total_to_delete}]: {cloud_path}"))
            try:
                self.fs.remove(cloud_path)
                deleted_count += 1
                time.sleep(0.05) # Small delay to potentially avoid rate limiting
            except Exception as e:
                error_log_msg = self._("delete_error_file", path=cloud_path, error=e, default=f"Error deleting {cloud_path}: {e}")
                self.log(error_log_msg)
                errors_deleting.append(cloud_path) # Record the failed path
                # Optional: Add traceback logging here if needed for debugging delete errors
                # self.log(f"Deletion Error Detail ({cloud_path}): {traceback.format_exc(limit=1)}")

        finish_msg = self._("delete_finished", deleted_count=deleted_count, total_marked=total_to_delete, default=f"Deletion complete. Successfully deleted {deleted_count} of {total_to_delete} marked files.")
        self.log(finish_msg)

        # Report any files that failed to delete
        if errors_deleting:
             num_errors = len(errors_deleting)
             self.log(self._("warning_delete_failures", count=num_errors, default=f"WARNING: Failed to delete {num_errors} file(s):"))
             # Log first few failed paths for diagnosis
             for failed_path in errors_deleting[:10]:
                 self.log(f"  - {failed_path}")
             if num_errors > 10:
                 self.log(self._("warning_delete_failures_more", count=num_errors - 10, default=f"  ... and {num_errors - 10} more."))

        return deleted_count, total_to_delete


# --- GUI Application Class ---
class DuplicateFinderApp:
    def __init__(self, master):
        self.master = master
        self.current_language = self.load_language_preference()
        self.finder = DuplicateFileFinder()
        self.finder.set_translator(self._) # Pass translator to finder

        # Application state
        self.duplicate_sets = {} # Stores {sha1: [FileInfo, ...]}
        self.treeview_item_map = {} # Maps tree item ID (file path) -> FileInfo dict
        self.files_to_delete_cache = [] # Stores paths marked for deletion by rule

        # Tkinter variables
        self.widgets = {} # Holds widget references
        self.string_vars = {} # Holds Entry StringVars
        self.entries = {} # Holds Entry widgets
        self.rule_radios = {} # Holds Radiobutton widgets specific to rules
        self.deletion_rule_var = tk.StringVar(value="") # For deletion rule radio buttons
        self.suffix_entry_var = tk.StringVar() # For suffix entry text

        # --- Treeview Sorting State ---
        self._last_sort_col = None # Track last sorted column ID
        self._sort_ascending = True # Track sort direction

        # --- Window Setup ---
        master.title(self._("window_title", default="Duplicate Finder"))
        master.geometry("1000x800") # Initial size
        master.minsize(850, 650) # Minimum size

        # --- Set Application Icon ---
        try:
            icon_path = ICON_FILE
            if os.path.exists(icon_path):
                master.iconbitmap(icon_path)
            else:
                # Log warning if icon not found
                print(f"Warning: Application icon file not found at '{icon_path}'")
                self.log_message(f"Warning: Application icon file not found at '{os.path.basename(icon_path)}'")
        except tk.TclError as e:
            # Handle specific Tkinter error loading icon
            icon_err_msg = self._("error_icon_load", path=os.path.basename(icon_path), error=e,
                                  default=f"Error loading icon '{os.path.basename(icon_path)}': {e}")
            print(icon_err_msg)
            self.log_message(icon_err_msg)
        except Exception as e:
            # Catch any other unexpected error during icon loading
            icon_err_msg = self._("error_icon_load", path=os.path.basename(icon_path), error=f"Unexpected error: {e}",
                                  default=f"Unexpected error loading icon '{os.path.basename(icon_path)}': {e}")
            print(icon_err_msg)
            self.log_message(icon_err_msg)

        # --- Menu Bar ---
        self.menu_bar = Menu(master)
        master.config(menu=self.menu_bar)
        self.create_menus() # Populate the menu bar

        # --- Create Main Layout Structure (PanedWindow) ---
        # Use a PanedWindow for resizable sections
        self.paned_window = ttk.PanedWindow(master, orient=tk.VERTICAL)
        self.paned_window.pack(fill=tk.BOTH, expand=True, padx=10, pady=10) # Added outer padding

        # --- Build the UI Sections within the PanedWindow ---
        self._build_ui_structure()

        # --- Final Setup ---
        self.load_config() # Load settings on startup
        self.update_ui_language() # Set initial UI text
        self.set_ui_state('initial') # Initial state before connection


    def _build_ui_structure(self):
        """Creates and packs/grids all the UI widgets within the main paned_window."""

        # --- Style Configuration ---
        style = ttk.Style()
        try:
            # Configure a style for the delete button (bold red text)
            style.configure("Danger.TButton", foreground="red", font=('TkDefaultFont', 10, 'bold'))
        except tk.TclError:
             style.configure("Danger.TButton", foreground="red") # Fallback if font spec fails

        # --- 1. Top Pane: Configuration and Primary Actions ---
        top_pane = ttk.Frame(self.paned_window, padding=5)
        # Add to PanedWindow, weight=0 -> no vertical stretch by default
        self.paned_window.add(top_pane, weight=0)
        top_pane.columnconfigure(0, weight=1) # Allow content (config frame) to expand horizontally

        # Configuration Input Fields (within a LabelFrame)
        config_frame = ttk.LabelFrame(top_pane, text=self._("config_title"), padding=(10, 5))
        config_frame.grid(row=0, column=0, padx=5, pady=5, sticky="ew") # Fill horizontally
        config_frame.columnconfigure(1, weight=1) # Entries expand
        self.widgets["config_frame"] = config_frame

        # Define config fields: (internal_key, label_translation_key, grid_row)
        config_fields = [
            ("address", "address_label", 0),
            ("account", "account_label", 1),
            ("password", "password_label", 2),
            ("scan_path", "scan_path_label", 3),
            ("mount_point", "mount_point_label", 4),
        ]

        for key, label_key, row in config_fields:
            label = ttk.Label(config_frame, text=self._(label_key))
            label.grid(row=row, column=0, padx=(5, 2), pady=3, sticky=tk.W)
            self.widgets[f"label_{key}"] = label

            var = tk.StringVar()
            self.string_vars[key] = var
            entry_args = {"textvariable": var} # width removed, relies on grid stretch
            if key == "password":
                entry_args["show"] = "*"
            entry = ttk.Entry(config_frame, **entry_args)
            entry.grid(row=row, column=1, padx=(2, 5), pady=3, sticky=tk.EW) # Expand horizontally
            self.entries[key] = entry

        # Action Buttons Frame (Load, Save, Test, Find)
        action_button_frame = ttk.Frame(top_pane, padding=(5, 0))
        action_button_frame.grid(row=1, column=0, padx=5, pady=(5, 10), sticky="ew")
        # Inner frame to keep buttons left-aligned
        btn_frame_inner = ttk.Frame(action_button_frame)
        btn_frame_inner.pack(side=tk.LEFT) # Pack inner frame to the left

        # Define buttons: (internal_key, translation_key, command, initial_state)
        action_buttons_info = [
            ("load", "load_config_button", self.load_config, tk.NORMAL),
            ("save", "save_config_button", self.save_config, tk.NORMAL),
            ("test_conn", "test_connection_button", self.start_test_connection_thread, tk.NORMAL),
            ("find", "find_button", self.start_find_duplicates_thread, tk.DISABLED), # Disabled initially
        ]
        # Use pack within the inner frame for simple left-to-right layout
        for idx, (w_key, t_key, cmd, initial_state) in enumerate(action_buttons_info):
             padx_val = (0, 5) # Space between buttons
             button = ttk.Button(btn_frame_inner, text=self._(t_key), command=cmd, state=initial_state)
             button.pack(side=tk.LEFT, padx=padx_val, pady=5)
             self.widgets[f"{w_key}_button"] = button


        # --- 2. Middle Pane: Deletion Rules and Results TreeView ---
        middle_pane = ttk.Frame(self.paned_window, padding=5)
        # Add to PanedWindow, weight=1 -> WILL EXPAND vertically
        self.paned_window.add(middle_pane, weight=1)
        middle_pane.rowconfigure(1, weight=1) # Treeview frame expands
        middle_pane.columnconfigure(0, weight=1) # Content expands horizontally

        # Deletion Rules (within a LabelFrame)
        rules_frame = ttk.LabelFrame(middle_pane, text=self._("rules_title"), padding=(10, 5))
        rules_frame.grid(row=0, column=0, padx=5, pady=(5, 5), sticky="ew") # Fill horizontally
        rules_frame.columnconfigure(2, weight=1) # Allow suffix entry to expand
        self.widgets["rules_frame"] = rules_frame

        # Define rule radio buttons: (translation_key_suffix, value_constant, grid_row)
        rule_options = [
            ("shortest_path", RULE_KEEP_SHORTEST, 0),
            ("longest_path", RULE_KEEP_LONGEST, 1),
            ("oldest", RULE_KEEP_OLDEST, 2),
            ("newest", RULE_KEEP_NEWEST, 3),
            ("keep_suffix", RULE_KEEP_SUFFIX, 4)
        ]
        # self.rule_radios dictionary defined in __init__
        suffix_row_index = -1

        for row_idx, (t_key_suffix, value, grid_row) in enumerate(rule_options):
            t_key = f"rule_{t_key_suffix}"
            radio = ttk.Radiobutton(rules_frame, text=self._(t_key),
                                    variable=self.deletion_rule_var, value=value,
                                    command=self._on_rule_change, state=tk.DISABLED) # Disabled initially
            radio.grid(row=grid_row, column=0, columnspan=1, padx=5, pady=2, sticky="w")
            self.rule_radios[value] = radio
            self.widgets[f"radio_{value}"] = radio
            if value == RULE_KEEP_SUFFIX:
                suffix_row_index = grid_row

        # Suffix Label and Entry (next to the last radio button)
        lbl = ttk.Label(rules_frame, text=self._("rule_suffix_entry_label"), state=tk.DISABLED)
        lbl.grid(row=suffix_row_index, column=1, padx=(15, 2), pady=2, sticky="e")
        self.widgets["suffix_label"] = lbl

        entry = ttk.Entry(rules_frame, textvariable=self.suffix_entry_var, state=tk.DISABLED) # width removed
        entry.grid(row=suffix_row_index, column=2, padx=(0, 5), pady=2, sticky="ew") # Expand horizontally
        self.widgets["suffix_entry"] = entry
        self.entries["suffix"] = entry

        # Results TreeView Frame (for Treeview and Scrollbars)
        tree_frame = ttk.Frame(middle_pane)
        tree_frame.grid(row=1, column=0, padx=5, pady=(5, 5), sticky="nsew") # Fills expanding area
        tree_frame.rowconfigure(0, weight=1) # Treeview expands vertically
        tree_frame.columnconfigure(0, weight=1) # Treeview expands horizontally
        self.widgets["tree_frame"] = tree_frame

        # Define Treeview Columns
        self.columns = ("action", "path", "modified", "size_mb", "set_id")
        self.tree = ttk.Treeview(tree_frame, columns=self.columns, show="headings", selectmode="none")
        self.widgets["treeview"] = self.tree

        # Define Column Widths and Alignments (adjust as needed)
        self.tree.column("action", width=80, anchor=tk.CENTER, stretch=tk.NO)
        self.tree.column("path", width=550, anchor=tk.W, stretch=tk.YES) # Allow path to stretch most
        self.tree.column("modified", width=150, anchor=tk.W, stretch=tk.NO)
        self.tree.column("size_mb", width=100, anchor=tk.E, stretch=tk.NO)
        self.tree.column("set_id", width=60, anchor=tk.CENTER, stretch=tk.NO)

        # Scrollbars for Treeview
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        # Grid Treeview and Scrollbars
        self.tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')

        # Define tags for highlighting rows
        self.tree.tag_configure('keep', foreground='darkgreen')
        self.tree.tag_configure('delete', foreground='#CC0000', font=('TkDefaultFont', 9, 'bold'))

        # Setup headings text and sorting bind (called later in update_ui_language)


        # --- 3. Bottom Pane: Log Area and Final Actions ---
        bottom_pane = ttk.Frame(self.paned_window, padding=5)
        # Add to PanedWindow, weight=0 -> no vertical stretch by default for the pane itself
        self.paned_window.add(bottom_pane, weight=0)
        # Configure bottom_pane's grid to allow log_frame to expand
        bottom_pane.columnconfigure(0, weight=1) # Allow content (log frame) to expand horizontally
        bottom_pane.rowconfigure(1, weight=1) # Allow Log frame row to expand vertically


        # Final Action Buttons Frame (Delete, Chart, Save Report)
        final_action_frame = ttk.Frame(bottom_pane)
        final_action_frame.grid(row=0, column=0, padx=5, pady=(5,0), sticky="ew")
        # Inner frame to keep buttons left-aligned
        final_btn_inner_frame = ttk.Frame(final_action_frame)
        final_btn_inner_frame.pack(side=tk.LEFT) # Pack inner frame to the left

        # Define final action buttons: (key, t_key, command, state, style)
        final_buttons_info = [
             ("delete", "delete_by_rule_button", self.start_delete_by_rule_thread, tk.DISABLED, "Danger.TButton"),
             ("chart", "show_chart_button", self.show_cloud_file_types, tk.DISABLED, ""),
             ("save_list", "save_list_button", self.save_duplicates_report, tk.DISABLED, ""),
        ]

        # Use pack within the inner frame
        for idx, (w_key, t_key, cmd, initial_state, style_name) in enumerate(final_buttons_info):
            padx_val = (0, 10) # Increased padding between final buttons
            btn_args = {"text": self._(t_key), "command": cmd, "state": initial_state}
            if style_name: btn_args["style"] = style_name
            # Special handling for chart button text/state based on matplotlib availability
            if w_key == "chart":
                effective_t_key = t_key if MATPLOTLIB_AVAILABLE else "show_chart_button_disabled"
                effective_state = initial_state if MATPLOTLIB_AVAILABLE else tk.DISABLED
                btn_args["text"] = self._(effective_t_key)
                btn_args["state"] = effective_state

            button = ttk.Button(final_btn_inner_frame, **btn_args)
            button.pack(side=tk.LEFT, padx=padx_val, pady=5)
            self.widgets[f"{w_key}_button"] = button

        # Log Output Area (within a LabelFrame)
        log_frame = ttk.LabelFrame(bottom_pane, text=self._("log_title"), padding=(5, 5))
        log_frame.grid(row=1, column=0, sticky="nsew", padx=5, pady=(5,5)) # Fill expanding area
        # Configure log_frame's grid for the ScrolledText expansion
        log_frame.rowconfigure(0, weight=1) # Text area expands vertically
        log_frame.columnconfigure(0, weight=1) # Text area expands horizontally
        self.widgets["log_frame"] = log_frame

        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=10, # Initial height
                                                  state='disabled', relief=tk.SOLID, borderwidth=1,
                                                  font=("TkDefaultFont", 9))
        self.log_text.grid(row=0, column=0, sticky="nsew") # Fills log_frame
        self.widgets["log_text"] = self.log_text

    # --- Language Handling ---
    def _(self, key, **kwargs):
        """ Translation helper: current lang -> default lang -> default string -> format. """
        lang_dict = translations.get(self.current_language, translations[DEFAULT_LANG])
        default_val = kwargs.pop('default', f"<{key}?>") # Default fallback
        base_string = lang_dict.get(key, translations[DEFAULT_LANG].get(key, default_val))
        try:
            # Format only if needed and possible
            if '{' in base_string and '}' in base_string and kwargs:
                 return base_string.format(**kwargs)
            else:
                 return base_string
        except KeyError as e:
             print(f"Warning: Formatting KeyError for key '{key}' ({self.current_language}): Missing key {e}. Kwargs: {kwargs}")
             return f"{base_string} [FORMATTING ERROR: Missing key '{e}']" # Indicate error in output
        except Exception as e:
             print(f"Warning: Formatting failed for key '{key}' ({self.current_language}): {e}. Kwargs: {kwargs}")
             return f"{base_string} [FORMATTING ERROR]"

    def save_language_preference(self):
        """ Saves the current language selection to a JSON file. """
        pref_path = LANG_PREF_FILE
        try:
            with open(pref_path, 'w', encoding='utf-8') as f:
                json.dump({"language": self.current_language}, f, indent=2)
        except IOError as e:
            print(f"Warning: Could not save language preference to {os.path.basename(pref_path)}: {e}")
            self.log_message(f"Warning: Could not save language preference: {e}")
        except Exception as e:
             print(f"Error saving language preference: {e}")
             self.log_message(f"Error saving language preference: {e}")

    def load_language_preference(self):
        """ Loads the language preference from JSON file, defaulting to DEFAULT_LANG. """
        pref_path = LANG_PREF_FILE
        try:
            if os.path.exists(pref_path):
                with open(pref_path, 'r', encoding='utf-8') as f:
                    prefs = json.load(f)
                    lang = prefs.get("language")
                    if lang and lang in translations:
                        print(f"Loaded language preference: {lang}")
                        return lang
                    elif lang:
                        print(f"Warning: Loaded language '{lang}' not supported. Falling back to {DEFAULT_LANG}.")
        except (IOError, json.JSONDecodeError) as e:
            print(f"Warning: Could not load or parse language preference from {os.path.basename(pref_path)}: {e}. Using default.")
        except Exception as e:
             print(f"Unexpected error loading language preference: {e}. Using default.")
        print(f"Using default language: {DEFAULT_LANG}")
        return DEFAULT_LANG

    def create_menus(self):
        """ Creates the application's menu bar (currently just Language). """
        lang_menu = Menu(self.menu_bar, tearoff=0)
        self.widgets["lang_menu"] = lang_menu
        lang_menu.add_command(label="English", command=lambda: self.change_language("en"))
        lang_menu.add_command(label="中文", command=lambda: self.change_language("zh"))
        self.menu_bar.add_cascade(label="Language", menu=lang_menu) # Label updated later


    def change_language(self, lang_code):
        """ Changes the application language, updates UI, and saves preference. """
        if lang_code in translations and lang_code != self.current_language:
            print(f"Changing language to: {lang_code}")
            self.current_language = lang_code
            # IMPORTANT: Update the finder's translator
            self.finder.set_translator(self._)
            self.update_ui_language()
            self.save_language_preference()
            self.log_message(f"Language changed to '{lang_code}'.") # This log will use the new language
        elif lang_code == self.current_language:
            self.log_message(f"Language is already set to '{lang_code}'.")
        elif lang_code not in translations:
             print(f"Error: Attempted to change to unsupported language '{lang_code}'.")
             self.log_message(f"Error: Language code '{lang_code}' is not supported.")

    def update_ui_language(self):
        """ Updates the text of all UI elements based on the current language. """
        print(f"Updating UI language to: {self.current_language}")
        if not self.master or not self.master.winfo_exists():
             print("Error: Master window does not exist during UI language update.")
             return

        try:
            # Window Title
            self.master.title(self._("window_title"))

            # Menu Bar
            if self.menu_bar and self.menu_bar.winfo_exists():
                try:
                    self.menu_bar.entryconfig(0, label=self._("menu_language"))
                except (tk.TclError, IndexError): pass
                lang_menu = self.widgets.get("lang_menu")
                if lang_menu and lang_menu.winfo_exists():
                    try: lang_menu.entryconfig(0, label=self._("menu_english"))
                    except (tk.TclError, IndexError): pass
                    try: lang_menu.entryconfig(1, label=self._("menu_chinese"))
                    except (tk.TclError, IndexError): pass

            # LabelFrame Titles
            for key, title_key in [("config_frame", "config_title"),
                                ("rules_frame", "rules_title"),
                                ("log_frame", "log_title")]:
                widget = self.widgets.get(key)
                if widget and widget.winfo_exists():
                    try: widget.config(text=self._(title_key))
                    except tk.TclError: pass

            # Labels
            label_keys = {
                "label_address": "address_label", "label_account": "account_label",
                "label_password": "password_label", "label_scan_path": "scan_path_label",
                "label_mount_point": "mount_point_label",
                "suffix_label": "rule_suffix_entry_label"
            }
            for widget_key, text_key in label_keys.items():
                widget = self.widgets.get(widget_key)
                if widget and widget.winfo_exists():
                    try: widget.config(text=self._(text_key))
                    except tk.TclError: pass

            # Button Texts
            button_keys = {
                "load_button": "load_config_button", "save_button": "save_config_button",
                "test_conn_button": "test_connection_button",
                "find_button": "find_button",
                "delete_button": "delete_by_rule_button",
                "save_list_button": "save_list_button",
            }
            for widget_key, text_key in button_keys.items():
                widget = self.widgets.get(widget_key)
                if widget and widget.winfo_exists():
                    try: widget.config(text=self._(text_key))
                    except tk.TclError: pass

            # Chart Button (special handling)
            chart_button = self.widgets.get("chart_button")
            if chart_button and chart_button.winfo_exists():
                 try:
                     effective_text_key = "show_chart_button" if MATPLOTLIB_AVAILABLE else "show_chart_button_disabled"
                     chart_button.config(text=self._(effective_text_key))
                 except tk.TclError: pass

            # Radio Button Texts
            radio_keys = {
                RULE_KEEP_SHORTEST: "rule_shortest_path", RULE_KEEP_LONGEST: "rule_longest_path",
                RULE_KEEP_OLDEST: "rule_oldest", RULE_KEEP_NEWEST: "rule_newest",
                RULE_KEEP_SUFFIX: "rule_keep_suffix"
            }
            for value, text_key in radio_keys.items():
                # Use the stored radio widgets directly
                widget = self.rule_radios.get(value)
                if widget and widget.winfo_exists():
                    try: widget.config(text=self._(text_key))
                    except tk.TclError: pass

            # Treeview Headings
            self.setup_treeview_headings()

            # Re-apply rule highlighting if data exists (updates Keep/Delete text)
            tree = self.widgets.get("treeview")
            if self.duplicate_sets and tree and tree.winfo_exists():
                 # Check if a rule is actually selected before reapplying
                 if self.deletion_rule_var.get():
                    self._apply_rule_to_treeview(log_update=False) # Avoid redundant logging during language switch

            print(f"UI Language update to {self.current_language} complete.")

        except Exception as e:
             print(f"ERROR during UI language update: {e}")
             self.log_message(f"ERROR: Failed to fully update UI language: {e}")
             self.log_message(traceback.format_exc(limit=3))


    def setup_treeview_headings(self):
        """Sets or updates the text of the treeview headings and binds sorting command."""
        heading_keys = {
            "action": "tree_rule_action_col", "path": "tree_path_col",
            "modified": "tree_modified_col", "size_mb": "tree_size_col",
            "set_id": "tree_set_col"
        }
        tree = self.widgets.get("treeview")
        if not tree or not tree.winfo_exists():
             return

        try:
            current_columns = tree['columns']
            for col_id in current_columns:
                if col_id in heading_keys:
                    text_key = heading_keys[col_id]
                    heading_text = self._(text_key, default=col_id.replace('_',' ').title())

                    sort_indicator = ""
                    if col_id == self._last_sort_col:
                        sort_indicator = ' ▲' if self._sort_ascending else ' ▼'

                    # Align headers left for consistency, except size/set#
                    anchor = tk.W
                    if col_id in ["size_mb", "set_id", "action"]:
                        anchor = tk.CENTER if col_id != "size_mb" else tk.E

                    tree.heading(col_id, text=heading_text + sort_indicator, anchor=anchor,
                                 command=lambda c=col_id: self._treeview_sort_column(c))
        except tk.TclError:
            print("Warning: Treeview widget destroyed during heading update.")
        except Exception as e:
            print(f"Error setting up treeview headings: {e}")
            self.log_message(f"Error configuring treeview headers: {e}")


    def _treeview_sort_column(self, col):
        """Sorts the treeview rows based on the clicked column header."""
        tree = self.widgets.get("treeview")
        if not tree or not tree.winfo_exists() or not self.treeview_item_map:
            self.log_message("No data in the list to sort.")
            return

        # Determine Sort Order
        if col == self._last_sort_col:
            self._sort_ascending = not self._sort_ascending
        else:
            self._sort_ascending = True
            self._last_sort_col = col

        # Prepare Data for Sorting
        items_to_sort = []
        min_datetime_sort = datetime.min.replace(tzinfo=timezone.utc) + timedelta(seconds=1)
        max_datetime_sort = datetime.max.replace(tzinfo=timezone.utc) - timedelta(seconds=1)
        min_str_sort = ""
        max_str_sort = "~" # Sorts after most common characters

        for item_id in tree.get_children(''):
            if item_id in self.treeview_item_map:
                file_info = self.treeview_item_map[item_id]
                sort_value = None
                try:
                    if col == 'path':
                        sort_value = file_info.get('path', '').lower()
                    elif col == 'modified':
                        dt_obj = file_info.get('modified')
                        # Ensure timezone-aware comparison if possible, otherwise naive
                        if dt_obj and dt_obj.tzinfo is None:
                            # Fallback: treat naive as UTC for sorting? Or use min/max?
                            # Using min/max is safer for comparison consistency if mix exists
                             sort_value = min_datetime_sort if self._sort_ascending else max_datetime_sort
                        else:
                            sort_value = dt_obj if isinstance(dt_obj, datetime) else (min_datetime_sort if self._sort_ascending else max_datetime_sort)

                    elif col == 'size_mb':
                        # Sort by actual byte size for accuracy, not formatted MB string
                        sort_value = file_info.get('size', 0) if file_info.get('size') is not None else 0
                    elif col == 'set_id':
                        current_val_str = tree.set(item_id, col)
                        match = re.search(r'\d+', current_val_str)
                        sort_value = int(match.group(0)) if match else 0
                    elif col == 'action':
                        sort_value = tree.set(item_id, col).lower() # Sort by "Keep" / "Delete" text
                    else: # Fallback? Should not happen with defined columns
                        sort_value = tree.set(item_id, col).lower()

                    items_to_sort.append((sort_value, item_id))

                except Exception as e:
                    print(f"Error getting sort value for item {item_id}, col {col}: {e}")
                    default_sort_val = 0
                    if col == 'modified': default_sort_val = min_datetime_sort if self._sort_ascending else max_datetime_sort
                    elif col in ['path', 'action']: default_sort_val = min_str_sort if self._sort_ascending else max_str_sort
                    items_to_sort.append((default_sort_val, item_id))

        # Perform the Sort
        try:
            items_to_sort.sort(key=lambda x: x[0], reverse=not self._sort_ascending)
        except TypeError as te:
            # This often happens comparing timezone-aware and naive datetimes
            self.log_message(f"Error: Could not sort column '{col}'. Inconsistent data types found (e.g., dates with/without timezone). ({te})")
            print(f"Sorting TypeError for column {col}: {te}")
            # Attempt fallback sort by string representation
            try:
                print(f"Attempting fallback string sort for column {col}")
                items_to_sort.sort(key=lambda x: str(x[0]), reverse=not self._sort_ascending)
            except Exception as fallback_e:
                self.log_message(f"Error: Fallback sort for column '{col}' also failed. ({fallback_e})")
                self._last_sort_col = None
                self.setup_treeview_headings() # Reset header visuals
                return

        # Reorder Items in the Treeview
        for i, (_, item_id) in enumerate(items_to_sort):
            try:
                tree.move(item_id, '', i)
            except tk.TclError as e:
                print(f"Error moving tree item {item_id}: {e}")

        # Update Heading Visuals
        self.setup_treeview_headings()


    # --- GUI Logic Methods ---
    def log_message(self, message):
        """ Safely appends a timestamped message to the log ScrolledText widget from any thread. """
        message_str = str(message) if message is not None else ""
        log_widget = self.widgets.get("log_text")

        if hasattr(self, 'master') and self.master and self.master.winfo_exists() and \
           log_widget and log_widget.winfo_exists():
            try:
                # Use `after` to ensure thread safety for Tkinter updates
                self.master.after(0, self._append_log, message_str)
            except (tk.TclError, RuntimeError) as e:
                 # Handle cases where the window might be destroyed between check and call
                 print(f"Log Error (Tcl/Runtime): {e} - Message: {message_str}")
            except Exception as e:
                 # Catch other potential errors during scheduling
                 print(f"Error scheduling log message: {e}\nMessage: {message_str}")
        else:
             # Fallback if GUI elements are not available (e.g., during shutdown)
             timestamp = datetime.now().strftime("%H:%M:%S")
             print(f"[LOG FALLBACK - {timestamp}] {message_str}")

    def _append_log(self, message):
        """ Internal method to append message to log widget (MUST run in main GUI thread). """
        log_widget = self.widgets.get("log_text")
        if not log_widget or not log_widget.winfo_exists():
            return # Widget gone

        try:
            # Temporarily enable, insert text, scroll, disable again
            current_state = log_widget.cget('state')
            log_widget.configure(state='normal')
            timestamp = datetime.now().strftime("%H:%M:%S")
            log_widget.insert(tk.END, f"[{timestamp}] {message}\n")
            log_widget.see(tk.END) # Scroll to the end
            log_widget.configure(state=current_state) # Restore previous state
        except tk.TclError as e:
            # Catch errors if the widget is destroyed mid-operation
            print(f"Log Append TclError: {e} - Message: {message}")
        except Exception as e:
            print(f"Unexpected error appending log: {e}\nMessage: {message}")
            # Try to restore state even on unexpected error
            try: log_widget.configure(state=current_state)
            except: pass

    def load_config(self):
        """ Loads configuration from the ini file into the GUI fields. """
        config_path = CONFIG_FILE
        self.log_message(self._("status_loading_config", file=os.path.basename(config_path), default=f"Loading config from {os.path.basename(config_path)}..."))
        config = configparser.ConfigParser()
        try:
            if not os.path.exists(config_path):
                 self.log_message(self._("status_config_not_found", file=os.path.basename(config_path), default=f"Config file '{os.path.basename(config_path)}' not found."))
                 for key in ["address", "account", "password", "scan_path", "mount_point"]:
                     if key in self.string_vars: self.string_vars[key].set("")
                 return

            read_files = config.read(config_path, encoding='utf-8')
            if not read_files:
                 self.log_message(f"Warning: Config file '{os.path.basename(config_path)}' exists but could not be read or is empty.")
                 return

            if 'config' in config:
                cfg_section = config['config']
                # Load using original keys for compatibility
                self.string_vars["address"].set(cfg_section.get("clouddrvie2_address", ""))
                self.string_vars["account"].set(cfg_section.get("clouddrive2_account", ""))
                self.string_vars["password"].set(cfg_section.get("clouddrive2_passwd", ""))
                self.string_vars["scan_path"].set(cfg_section.get("root_path", ""))
                self.string_vars["mount_point"].set(cfg_section.get("clouddrive2_root_path", ""))
                self.log_message(self._("status_config_loaded", default="Config loaded successfully."))
            else:
                self.log_message(self._("status_config_section_missing", default="Config file loaded, but '[config]' section is missing."))

        except configparser.Error as e:
            error_msg = self._("error_config_read", error=e, default=f"Error reading config file: {e}")
            if self.master.winfo_exists():
                 messagebox.showerror(self._("error_config_title", default="Config Error"), error_msg, master=self.master)
            self.log_message(error_msg)
        except Exception as e:
             error_msg = self._("error_unexpected", error=f"loading config: {e}", default=f"Unexpected error loading config: {e}")
             if self.master.winfo_exists():
                 messagebox.showerror(self._("error_title", default="Error"), error_msg, master=self.master)
             self.log_message(error_msg)
             self.log_message(traceback.format_exc())


    def save_config(self):
        """ Saves current configuration from GUI fields to the ini file. """
        config_path = CONFIG_FILE
        self.log_message(self._("status_saving_config", file=os.path.basename(config_path), default=f"Saving config to {os.path.basename(config_path)}..."))
        config = configparser.ConfigParser()

        # Prepare data using INI keys
        config_data = {
            "clouddrvie2_address": self.string_vars["address"].get(),
            "clouddrive2_account": self.string_vars["account"].get(),
            "clouddrive2_passwd": self.string_vars["password"].get(),
            "root_path": self.string_vars["scan_path"].get(),
            "clouddrive2_root_path": self.string_vars["mount_point"].get(),
        }
        config['config'] = config_data

        # Preserve Other Sections (Best effort)
        try:
            if os.path.exists(config_path):
                 config_old = configparser.ConfigParser()
                 # Read existing config to preserve other sections
                 config_old.read(config_path, encoding='utf-8')
                 for section in config_old.sections():
                     if section != 'config' and not config.has_section(section):
                         config.add_section(section)
                         for key, value in config_old.items(section):
                             config.set(section, key, value)
        except Exception as e:
            print(f"Warning: Could not merge existing config sections during save: {e}")
            self.log_message(f"Warning: Failed to preserve existing non-'config' sections during save: {e}")

        # Write Config File
        try:
            with open(config_path, 'w', encoding='utf-8') as configfile:
                config.write(configfile)
            self.log_message(self._("status_config_saved", default="Config saved successfully."))
        except IOError as e:
            error_msg = self._("error_config_save", error=e, default=f"Could not write config file: {e}")
            if self.master.winfo_exists():
                 messagebox.showerror(self._("error_config_save_title", default="Config Save Error"), error_msg, master=self.master)
            self.log_message(error_msg)
        except Exception as e:
             error_msg = self._("error_unexpected", error=f"saving config: {e}", default=f"Unexpected error saving config: {e}")
             if self.master.winfo_exists():
                 messagebox.showerror(self._("error_title", default="Error"), error_msg, master=self.master)
             self.log_message(error_msg)
             self.log_message(traceback.format_exc())

    def _check_path_chars(self, path_dict):
        """
        Validates characters in specified path inputs using _validate_path_chars helper.
        Logs details and shows a warning popup if suspicious characters are found.
        Returns True if all paths are valid, False otherwise.
        """
        suspicious_char_found = False
        all_details = []

        path_display_names = {
            "address": self._("address_label", default="API Address").rstrip(': '),
            "scan_path": self._("scan_path_label", default="Root Path to Scan").rstrip(': '),
            "mount_point": self._("mount_point_label", default="CloudDrive Mount Point").rstrip(': ')
        }

        for key, path_str in path_dict.items():
             if key not in path_display_names: continue # Only check known path fields

             issues = _validate_path_chars(path_str)
             if issues:
                 suspicious_char_found = True
                 display_name = path_display_names.get(key, key)
                 all_details.append(f"'{display_name}' (value: '{path_str}'): {', '.join(issues)}")

        if suspicious_char_found:
            log_sep = "!" * 70
            self.log_message(log_sep)
            warning_title = self._("path_warning_title", default="Path Input Warning")
            self.log_message(f"*** {warning_title} ***")
            for detail in all_details: self.log_message(f"  -> {detail}")

            warning_msg_template = self._("path_warning_suspicious_chars", default="Suspicious character(s) detected!\nPlease DELETE and MANUALLY RETYPE paths.")
            warning_lines = warning_msg_template.split('\n')
            # Construct popup message from template lines
            popup_msg = f"{warning_lines[0]}\n\n" + "\n".join(warning_lines[1:]) if len(warning_lines) > 1 else warning_lines[0]
            instruction = warning_lines[1] if len(warning_lines) > 1 else "Please check and retype."

            self.log_message(f"Instruction: {instruction}")
            self.log_message(log_sep)

            if self.master.winfo_exists():
                # Schedule the messagebox to ensure it runs in the main GUI thread
                self.master.after(10, lambda wt=warning_title, pm=popup_msg: messagebox.showerror(wt, pm, master=self.master))
            return False # Indicate failure

        return True # All paths checked were valid


    def set_ui_state(self, mode):
        """
        Enable/disable UI elements based on the application's current mode.
        Modes: 'initial', 'normal' (connected/idle), 'testing_connection', 'finding', 'deleting', 'charting'
        """
        is_idle_state = mode in ['initial', 'normal']
        is_connected = mode != 'initial' and self.finder is not None and self.finder.fs is not None

        # Determine derived states
        has_duplicates = is_connected and bool(self.duplicate_sets)
        is_rule_selected = is_connected and bool(self.deletion_rule_var.get())
        has_files_marked_for_deletion = is_connected and bool(self.files_to_delete_cache)

        # Calculate widget states based on mode and derived states
        config_entry_state = tk.NORMAL if is_idle_state else tk.DISABLED
        config_button_state = tk.NORMAL if is_idle_state else tk.DISABLED
        find_button_state = tk.NORMAL if is_idle_state and is_connected else tk.DISABLED
        rules_radio_state = tk.NORMAL if is_idle_state and is_connected and has_duplicates else tk.DISABLED

        # Suffix entry/label only enabled if the suffix rule is selected AND other conditions met
        suffix_widgets_state = tk.DISABLED
        if is_idle_state and is_connected and has_duplicates and self.deletion_rule_var.get() == RULE_KEEP_SUFFIX:
            suffix_widgets_state = tk.NORMAL

        # Delete button requires connection, duplicates, a selected rule, AND files marked by that rule
        delete_button_state = tk.NORMAL if is_idle_state and is_connected and has_duplicates and is_rule_selected and has_files_marked_for_deletion else tk.DISABLED
        save_report_button_state = tk.NORMAL if is_idle_state and is_connected and has_duplicates else tk.DISABLED
        # Chart button requires connection, idle state, and matplotlib installed
        chart_button_state = tk.NORMAL if is_idle_state and is_connected and MATPLOTLIB_AVAILABLE else tk.DISABLED


        # Apply states safely using try-except for widget existence
        for key, entry in self.entries.items():
            if key != "suffix" and entry and entry.winfo_exists(): # Exclude suffix entry here
                try: entry.config(state=config_entry_state)
                except tk.TclError: pass # Widget might be destroyed

        for btn_key in ["load_button", "save_button", "test_conn_button"]:
             widget = self.widgets.get(btn_key)
             if widget and widget.winfo_exists():
                 try: widget.config(state=config_button_state)
                 except tk.TclError: pass

        widget = self.widgets.get("find_button")
        if widget and widget.winfo_exists():
             try: widget.config(state=find_button_state)
             except tk.TclError: pass

        # Rules Frame widgets
        for radio in self.rule_radios.values():
             if radio and radio.winfo_exists():
                 try: radio.config(state=rules_radio_state)
                 except tk.TclError: pass
        widget = self.widgets.get("suffix_label")
        if widget and widget.winfo_exists():
             try: widget.config(state=suffix_widgets_state)
             except tk.TclError: pass
        widget = self.widgets.get("suffix_entry") # Suffix entry widget
        if widget and widget.winfo_exists():
             try: widget.config(state=suffix_widgets_state)
             except tk.TclError: pass

        # Final action buttons
        widget = self.widgets.get("delete_button")
        if widget and widget.winfo_exists():
             try: widget.config(state=delete_button_state)
             except tk.TclError: pass

        widget = self.widgets.get("save_list_button")
        if widget and widget.winfo_exists():
             try: widget.config(state=save_report_button_state)
             except tk.TclError: pass

        widget = self.widgets.get("chart_button")
        if widget and widget.winfo_exists():
            try:
                # Update text based on availability as well as state
                effective_text_key = "show_chart_button" if MATPLOTLIB_AVAILABLE else "show_chart_button_disabled"
                widget.config(state=chart_button_state, text=self._(effective_text_key))
            except tk.TclError: pass


    def start_test_connection_thread(self):
        """ Validates inputs and starts the connection test in a background thread. """
        address = self.string_vars["address"].get()
        account = self.string_vars["account"].get()
        # Mount point is needed by finder.set_config even for test connection logic
        mount_point = self.string_vars["mount_point"].get()
        # Scan path is also passed to set_config, though not used directly for ls('/') test
        scan_path = self.string_vars["scan_path"].get()

        # Validate required fields for connection test itself
        if not all([address, account, mount_point]):
             error_msg = self._("error_input_missing_conn", default="API Address, Account, and Mount Point are required for connection test.")
             self.log_message(error_msg)
             if self.master.winfo_exists(): messagebox.showerror(self._("error_input_title", default="Input Error"), error_msg, master=self.master)
             return

        # Validate characters in paths used by set_config
        paths_to_check_conn = {
            "address": address,         # Check address format implicitly
            "mount_point": mount_point,
            "scan_path": scan_path      # Check scan path as it's passed to set_config
        }
        if not self._check_path_chars(paths_to_check_conn):
            # Error message handled by _check_path_chars
            return

        self.log_message(self._("status_connecting", default="Attempting connection test..."))
        self.set_ui_state("testing_connection") # Disable UI during test
        # Run the blocking network call in a separate thread
        thread = threading.Thread(target=self._test_connection_worker,
                                  args=(address, account, self.string_vars["password"].get(), scan_path, mount_point),
                                  daemon=True) # Daemon thread exits if main app exits
        thread.start()

    def _test_connection_worker(self, address, account, passwd, scan_path, mount_point):
        """ Worker thread for testing the CloudDrive2 connection. """
        connected = False
        try:
            # Pass log_message callback to the finder
            connected = self.finder.set_config(address, account, passwd, scan_path, mount_point, self.log_message)

            # Schedule GUI updates (messagebox, UI state) back on the main thread
            if self.master.winfo_exists():
                 if connected:
                     success_title = self._("conn_test_success_title", default="Connection Test Successful")
                     success_msg = self._("conn_test_success_msg", default="Successfully connected to CloudDrive2.")
                     # Use master.after to queue the messagebox in the GUI thread
                     self.master.after(10, lambda st=success_title, sm=success_msg: messagebox.showinfo(st, sm, master=self.master))
                 else:
                     fail_title = self._("conn_test_fail_title", default="Connection Test Failed")
                     fail_msg = self._("conn_test_fail_msg", default="Failed to connect. Check log for details.")
                     self.master.after(10, lambda ft=fail_title, fm=fail_msg: messagebox.showwarning(ft, fm, master=self.master))

        except Exception as e:
            # Catch unexpected errors during the test itself
            error_msg = self._("error_unexpected", error=f"during connection test: {e}", default=f"Unexpected error during connection test: {e}")
            self.log_message(error_msg)
            self.log_message(traceback.format_exc())
            if self.master.winfo_exists():
                error_title = self._("error_title", default="Error")
                self.master.after(10, lambda et=error_title, em=error_msg: messagebox.showerror(et, em, master=self.master))
        finally:
            # ALWAYS Schedule UI Reset back on the main thread
            if self.master.winfo_exists():
                # Set final state based on whether connection was successful or not
                final_state = 'normal' if connected else 'initial'
                self.master.after(0, self.set_ui_state, final_state)


    def start_find_duplicates_thread(self):
        """ Handles 'Find Duplicates' click. Validates inputs, clears previous results, and starts worker thread. """
        # 1. Check Connection State
        if not self.finder or not self.finder.fs:
             if self.master.winfo_exists(): messagebox.showwarning(self._("error_title", default="Error"),
                                   self._("error_not_connected", default="Not connected. Please test connection first."),
                                   master=self.master)
             self.log_message(self._("error_not_connected", default="Error: Not connected. Cannot start scan."))
             return

        # 2. Validate Required Inputs for Finding
        scan_path_val = self.string_vars["scan_path"].get()
        mount_point_val = self.string_vars["mount_point"].get()
        address_val = self.string_vars["address"].get() # Still needed for finder state
        account_val = self.string_vars["account"].get() # Still needed for finder state

        missing = []
        if not address_val: missing.append(f"'{self._('address_label').rstrip(': ')}'")
        if not account_val: missing.append(f"'{self._('account_label').rstrip(': ')}'")
        if not scan_path_val: missing.append(f"'{self._('scan_path_label').rstrip(': ')}'")
        if not mount_point_val: missing.append(f"'{self._('mount_point_label').rstrip(': ')}'")

        if missing:
              error_msg_base = self._("error_input_missing", default="Required fields missing")
              # Format the message listing the missing fields
              error_msg = f"{error_msg_base}: {', '.join(missing)}."
              if self.master.winfo_exists(): messagebox.showerror(self._("error_input_title", default="Input Error"), error_msg, master=self.master)
              self.log_message(error_msg)
              return

        # 3. Validate Path Characters (Scan Path and Mount Point)
        paths_to_check = {
            "scan_path": scan_path_val,
            "mount_point": mount_point_val
        }
        if not self._check_path_chars(paths_to_check):
            # Error message handled by _check_path_chars
            return

        # 4. Clear Previous Results and Set UI State
        self.clear_results() # Clear tree, cache, rule selection
        self.log_message(self._("find_starting", default="Starting duplicate file scan..."))
        self.set_ui_state("finding") # Disable UI during scan

        # 5. Start Background Thread
        thread = threading.Thread(target=self._find_duplicates_worker, daemon=True)
        thread.start()

    def _find_duplicates_worker(self):
        """ Worker thread for finding duplicates. Calls finder method and schedules GUI update. """
        # Double check connection state within the thread
        if not self.finder or not self.finder.fs:
            self.log_message("Error: Connection lost before Find Duplicates scan could execute.")
            if self.master.winfo_exists(): self.master.after(0, self.set_ui_state, 'normal') # Reset UI state
            return

        found_duplicates = {}
        try:
            start_time = time.time()
            # Call the core logic in the finder class
            found_duplicates = self.finder.find_duplicates()
            end_time = time.time()
            # Log duration using the finder's log method (which uses the translator)
            # Note: finder already logs duration internally now. No need to repeat here.
            # self.log_message(f"Core duplicate finding operation took {end_time - start_time:.2f} seconds.") # Redundant

            # Schedule the processing of results back on the main GUI thread
            if self.master.winfo_exists():
                self.master.after(0, self._process_find_results, found_duplicates)

        except Exception as e:
            # Catch unexpected errors during the find process
            err_msg = self._("find_error_during", error=e, default=f"Unexpected error during scan process: {e}")
            self.log_message(err_msg)
            self.log_message(traceback.format_exc())
            if self.master.winfo_exists():
                 error_title = self._("error_title", default="Scan Error")
                 # Schedule error popup in main thread
                 self.master.after(10, lambda et=error_title, em=err_msg: messagebox.showerror(et, em, master=self.master))
                 # Reset UI state in main thread on error
                 self.master.after(0, self.set_ui_state, 'normal')


    def _process_find_results(self, found_duplicates):
        """ Processes results from find_duplicates worker (runs in main thread). Updates GUI. """
        if not self.master.winfo_exists(): return # Window closed

        self.duplicate_sets = found_duplicates if found_duplicates else {}

        if self.duplicate_sets:
            # Populate the treeview with the found duplicates
            self.populate_treeview()
            # If a rule was already selected, re-apply it visually (though cache is empty initially)
            if self.deletion_rule_var.get():
                self._apply_rule_to_treeview()
        else:
            # Clear treeview if no duplicates found
            tree = self.widgets.get("treeview")
            if tree and tree.winfo_exists():
                try: tree.delete(*tree.get_children())
                except tk.TclError: pass
            self.treeview_item_map.clear()
            # Finder already logs "no duplicates found" message

        # Set UI state back to normal (enabling rules/buttons if duplicates were found)
        # This should happen AFTER populate_treeview potentially enables rules
        self.set_ui_state('normal')


    def clear_results(self):
        """Clears the treeview, stored duplicate data, rule selection, and delete cache."""
        self.log_message(self._("status_clearing_tree", default="Clearing results list and rule selection..."))

        # Reset data structures
        self.duplicate_sets = {}
        self.treeview_item_map = {}
        self.files_to_delete_cache = []

        # Reset UI controls related to results/rules
        self.deletion_rule_var.set("") # Clear rule selection
        self.suffix_entry_var.set("") # Clear suffix entry

        # Reset sorting state
        self._last_sort_col = None
        self._sort_ascending = True

        # Clear the Treeview widget
        tree = self.widgets.get("treeview")
        if tree and tree.winfo_exists():
            try:
                tree.delete(*tree.get_children()) # Remove all items
                # Reset headers (clears sort indicators) after clearing
                self.setup_treeview_headings()
            except tk.TclError: pass # Handle if tree is destroyed

        # Reset UI state (disables rule radios, delete button etc.)
        self.set_ui_state('normal') # Set to 'normal' idle state


    def populate_treeview(self):
        """ Populates the treeview with found duplicate sets. """
        tree = self.widgets.get("treeview")
        if not tree or not tree.winfo_exists():
            self.log_message("Error: Treeview widget not available. Cannot display results.")
            return

        count = len(self.duplicate_sets)
        if count == 0:
             # This case should ideally be handled before calling populate_treeview
             self.log_message("No duplicate sets to display.")
             try: tree.delete(*tree.get_children())
             except tk.TclError: pass
             self.treeview_item_map.clear()
             return

        self.log_message(self._("status_populating_tree", count=count, default=f"Populating list with {count} duplicate sets..."))
        start_time = time.time()

        # --- Clear existing tree content ---
        try:
            tree.delete(*tree.get_children())
        except tk.TclError:
            self.log_message("Error clearing treeview before population.")
            return # Cannot proceed if clearing failed
        self.treeview_item_map.clear() # Clear the mapping cache

        # --- Insert new items ---
        set_index = 0
        items_inserted = 0
        items_failed = 0

        # Sort sets by SHA1 for consistent display order
        sorted_sha1s = sorted(self.duplicate_sets.keys())

        for sha1 in sorted_sha1s:
            files_in_set = self.duplicate_sets[sha1]
            # Ensure it's a valid duplicate set (at least 2 files)
            if not isinstance(files_in_set, list) or len(files_in_set) < 2:
                continue # Skip invalid sets silently or log a warning if needed

            set_index += 1
            # Sort files within the set by path for readability
            sorted_files = sorted(files_in_set, key=lambda x: x.get('path', ''))

            for file_info in sorted_files:
                try:
                    path = file_info.get('path')
                    if not path:
                         self.log_message(f"Warning: Skipping file in set {set_index} (SHA1: {sha1[:8]}...) due to missing path.")
                         items_failed += 1
                         continue

                    mod_time = file_info.get('modified')
                    # Format datetime safely, handling None
                    mod_time_str = mod_time.strftime(DATE_FORMAT) if isinstance(mod_time, datetime) else "N/A"

                    size = file_info.get('size')
                    # Calculate size safely, handling None or non-numeric types
                    size_mb = size / (1024 * 1024) if isinstance(size, (int, float)) and size > 0 else 0.0

                    set_id_str = self._("tree_set_col_value", index=set_index, default=f"{set_index}")

                    # Define values tuple in the order of self.columns
                    # ("action", "path", "modified", "size_mb", "set_id")
                    values = (
                        "",         # Action (initially empty)
                        path,
                        mod_time_str,
                        f"{size_mb:.2f}", # Format size to 2 decimal places
                        set_id_str
                    )
                    # Use the unique file path as the item identifier (iid)
                    item_id = path

                    # Insert into treeview
                    tree.insert("", tk.END, iid=item_id, values=values, tags=()) # No tags initially
                    # Store the full file_info dict mapped to its item ID (path)
                    self.treeview_item_map[item_id] = file_info
                    items_inserted += 1

                except tk.TclError as e:
                     # Handle errors during insertion (e.g., if path is somehow duplicated as iid)
                     self.log_message(f"Error inserting item with path '{path}' into tree: {e}")
                     items_failed += 1
                     # Clean up map if insertion failed for an ID already added
                     if item_id in self.treeview_item_map: del self.treeview_item_map[item_id]
                except Exception as e:
                     # Catch other unexpected errors during processing
                     self.log_message(f"Unexpected error processing file '{file_info.get('path', 'Unknown')}' for treeview: {e}")
                     self.log_message(traceback.format_exc(limit=2))
                     items_failed += 1

        end_time = time.time()
        duration = end_time - start_time
        log_summary = self._("status_tree_populated", default="Results list populated.")
        log_summary += f" ({items_inserted} items displayed"
        if items_failed > 0: log_summary += f", {items_failed} skipped due to errors"
        log_summary += f" in {duration:.2f}s)"
        self.log_message(log_summary)

        # Reset sorting state after population (important!)
        self._last_sort_col = None
        self._sort_ascending = True
        self.setup_treeview_headings() # Apply headers and clear sort indicators

        # Note: Applying the rule is handled separately after population if needed


    def _on_rule_change(self):
        """Called when a deletion rule radio button is selected."""
        selected_rule = self.deletion_rule_var.get()

        # Update suffix entry state immediately based on rule selection and other conditions
        is_suffix_rule = (selected_rule == RULE_KEEP_SUFFIX)
        # Suffix entry should only be enabled if suffix rule is selected AND we are connected AND have duplicates
        can_enable_suffix = (is_suffix_rule and
                             self.finder is not None and self.finder.fs is not None and
                             bool(self.duplicate_sets))
        suffix_widgets_state = tk.NORMAL if can_enable_suffix else tk.DISABLED

        # Apply state to suffix label and entry safely
        suffix_label = self.widgets.get("suffix_label")
        if suffix_label and suffix_label.winfo_exists():
            try: suffix_label.config(state=suffix_widgets_state)
            except tk.TclError: pass
        suffix_entry = self.widgets.get("suffix_entry")
        if suffix_entry and suffix_entry.winfo_exists():
            try: suffix_entry.config(state=suffix_widgets_state)
            except tk.TclError: pass

        # Clear suffix text if a different rule is selected
        if not is_suffix_rule:
            self.suffix_entry_var.set("")

        # Re-apply the selected rule to update Treeview highlighting and delete cache
        self._apply_rule_to_treeview()


    def _apply_rule_to_treeview(self, log_update=True):
        """
        Updates the 'Action' column and highlighting in the treeview based on the
        selected deletion rule. Updates the `files_to_delete_cache`.
        `log_update`: If True, logs the "Applying rule..." message. Set to False during language change.
        """
        tree = self.widgets.get("treeview")
        # Exit if no duplicates, or treeview doesn't exist
        if not self.duplicate_sets or not tree or not tree.winfo_exists():
            self.files_to_delete_cache = [] # Clear cache if no data
            self.set_ui_state('normal') # Ensure delete button is disabled etc.
            return

        selected_rule = self.deletion_rule_var.get()

        # If no rule is selected, clear action column and cache
        if not selected_rule:
            self.files_to_delete_cache = []
            try:
                for item_id in self.treeview_item_map.keys():
                    if tree.exists(item_id):
                         tree.set(item_id, "action", "") # Clear action text
                         tree.item(item_id, tags=())     # Clear tags (highlighting)
            except tk.TclError: pass
            self.set_ui_state('normal') # Reset UI state (likely disables delete button)
            return

        # --- A rule IS selected ---
        rule_name_display_key = f"rule_{selected_rule}" # e.g., rule_shortest
        rule_name_display = self._(rule_name_display_key, default=selected_rule.replace('_', ' ').title())

        if log_update:
            self.log_message(self._("status_applying_rule", rule_name=rule_name_display, count=len(self.duplicate_sets), default=f"Applying rule '{rule_name_display}'..."))
        start_time = time.time()

        keep_text = self._("tree_action_keep", default="Keep")
        delete_text = self._("tree_action_delete", default="Delete")
        suffix_to_keep = self.suffix_entry_var.get() if selected_rule == RULE_KEEP_SUFFIX else None

        delete_count = 0
        application_error = False

        try:
            # Call helper function to determine which files *should* be deleted
            files_to_delete_list = self._determine_files_to_delete(self.duplicate_sets, selected_rule, suffix_to_keep)

            # Update Cache
            self.files_to_delete_cache = files_to_delete_list # Store the paths
            # Create a set for efficient lookup during UI update
            files_to_delete_paths_set = set(files_to_delete_list)
            delete_count = len(files_to_delete_paths_set)

            # --- Update Treeview items visually ---
            tree_update_start = time.time()
            # Check if tree exists *before* iterating (it might be destroyed by user action)
            if not tree.winfo_exists():
                raise tk.TclError("Treeview destroyed during rule application")

            # Iterate over a copy of the keys, as map might change if errors occur
            for item_id in list(self.treeview_item_map.keys()):
                 # Check if the item still exists in the treeview
                 if not tree.exists(item_id): continue

                 file_info = self.treeview_item_map.get(item_id)
                 if not file_info: continue # Should not happen if item exists, but check anyway

                 path = file_info.get('path')
                 if not path: continue # Skip if path somehow missing

                 # Determine action and tag based on whether path is in the delete set
                 is_marked_for_delete = (path in files_to_delete_paths_set)
                 action_text = delete_text if is_marked_for_delete else keep_text
                 item_tags = ('delete',) if is_marked_for_delete else ('keep',)

                 # Update the treeview item's "Action" column and apply the tag
                 tree.set(item_id, "action", action_text)
                 tree.item(item_id, tags=item_tags)

            tree_update_end = time.time()
            if log_update:
                self.log_message(self._("status_rule_applied", delete_count=delete_count, default=f"Rule applied. {delete_count} files marked for deletion."))
                # Optional: Log performance
                # self.log_message(f"Treeview visual update took {tree_update_end - tree_update_start:.3f}s")


        except ValueError as ve:
             # Handle specific errors from _determine_files_to_delete (e.g., missing suffix)
             self.log_message(f"Rule Application Error: {ve}")
             if self.master.winfo_exists(): messagebox.showerror(self._("error_rule_title", default="Rule Error"), str(ve), master=self.master)
             application_error = True
             self.files_to_delete_cache = [] # Clear cache on error
             # Clear visual indicators in tree
             try:
                 if tree.winfo_exists():
                     for item_id in self.treeview_item_map.keys():
                         if tree.exists(item_id):
                             tree.set(item_id, "action", "")
                             tree.item(item_id, tags=())
             except tk.TclError: pass
             delete_count = 0

        except tk.TclError as e:
             # Handle errors if treeview is destroyed during update
             self.log_message(f"Error updating treeview during rule application: {e}")
             self.files_to_delete_cache = [] # Clear cache
             application_error = True
             delete_count = 0
        except Exception as e:
             # Catch any other unexpected errors
             self.log_message(f"Unexpected error applying rule '{selected_rule}' to treeview: {e}")
             self.log_message(traceback.format_exc())
             self.files_to_delete_cache = [] # Clear cache
             application_error = True
             delete_count = 0
        finally:
            # Log total duration if a rule was processed
            if selected_rule and log_update:
                end_time = time.time()
                # Optional: Log total performance
                # self.log_message(f"Rule determination & tree update took {end_time-start_time:.3f}s total")

            # ALWAYS Update UI State after applying rule (or attempting to)
            # This will correctly enable/disable the delete button based on cache state
            self.set_ui_state('normal')


    def _determine_files_to_delete(self, duplicate_sets, rule, suffix_value):
        """
        Determines which files to delete based on the selected rule and duplicate sets.
        Handles tie-breaking using shortest path as default. Logs warnings via self.log_message.

        Returns: list: A list of full file paths (str) to be deleted.
        Raises: ValueError: If rule is invalid or suffix is missing when required.
        """
        if not isinstance(duplicate_sets, dict) or not duplicate_sets:
             return [] # No duplicates to process
        if not rule:
             # Raise error if no rule is provided (should be caught earlier, but good validation)
             raise ValueError(self._("delete_no_rule_selected", default="No deletion rule selected."))
        if rule == RULE_KEEP_SUFFIX and not suffix_value:
             # Raise error if suffix rule is chosen but no suffix provided
             raise ValueError(self._("delete_suffix_missing", default="Suffix is required for the 'Keep Suffix' rule."))
        # Validate rule value against known constants
        if rule not in [RULE_KEEP_SHORTEST, RULE_KEEP_LONGEST, RULE_KEEP_OLDEST, RULE_KEEP_NEWEST, RULE_KEEP_SUFFIX]:
             raise ValueError(f"Internal Error: Unknown deletion rule '{rule}'.")

        files_to_delete = []
        log_func = self.log_message # Use the GUI's log method directly
        tie_break_prefix = self._("tie_break_log_prefix", default="Tie-Break:")

        # --- Tie-breaking Function (using shortest path) ---
        def tie_break_shortest_path(candidates, reason_for_tiebreak):
             if not candidates: return None # No candidates to break tie for
             if len(candidates) == 1: return candidates[0] # Only one candidate, it wins

             # Sort candidates first by path length (ascending), then alphabetically by path for determinism
             sorted_candidates = sorted(candidates, key=lambda f: (len(f.get('path', '')), f.get('path', '')))
             winner = sorted_candidates[0] # Shortest path wins

             # Log the tie-break decision
             log_func(self._("warning_tie_break",
                              prefix=tie_break_prefix,
                              reason=reason_for_tiebreak,
                              filename=os.path.basename(winner.get('path','N/A')),
                              detail=f"Shortest path of {len(candidates)}",
                              default=f"{tie_break_prefix} {reason_for_tiebreak}. Kept '{os.path.basename(winner.get('path','N/A'))}' (Shortest path of {len(candidates)})."))
             return winner

        # --- Iterate through each duplicate set ---
        for set_index, (sha1, files_in_set) in enumerate(duplicate_sets.items()):
            # Basic validation for the set
            if not isinstance(files_in_set, list) or len(files_in_set) < 2:
                continue # Skip sets that aren't actual duplicates

            keep_file_info = None # Info of the file decided to keep in this set
            set_id_for_log = f"Set {set_index+1} (SHA1: {sha1[:8]}...)" # For clearer logs

            try:
                candidates = [] # List of files potentially eligible to be kept based on the rule
                reason_for_tiebreak = "" # Reason if tie-breaking is needed

                # --- Apply the selected rule to find candidates ---
                if rule == RULE_KEEP_SHORTEST:
                    # Find min path length, allow for missing path attribute
                    min_len = min(len(f.get('path', '')) for f in files_in_set)
                    candidates = [f for f in files_in_set if len(f.get('path', '')) == min_len]
                    reason_for_tiebreak = f"Multiple files have min path length ({min_len})"

                elif rule == RULE_KEEP_LONGEST:
                    # Find max path length
                    max_len = max(len(f.get('path', '')) for f in files_in_set)
                    candidates = [f for f in files_in_set if len(f.get('path', '')) == max_len]
                    reason_for_tiebreak = f"Multiple files have max path length ({max_len})"

                elif rule == RULE_KEEP_OLDEST:
                    # Filter for files with valid datetime objects first
                    valid_files = [f for f in files_in_set if isinstance(f.get('modified'), datetime)]
                    if not valid_files:
                        # If no valid dates, log warning and consider all files for tie-break
                        log_func(self._("warning_rule_no_date", set_id=set_id_for_log, rule=rule, default=f"Warning: {set_id_for_log} - Cannot apply '{rule}': No valid dates. Keeping shortest path."))
                        candidates = list(files_in_set) # All files are potential candidates now
                        reason_for_tiebreak = "No valid dates found"
                    else:
                        # Find the minimum (oldest) date among valid files
                        min_date = min(f['modified'] for f in valid_files)
                        candidates = [f for f in valid_files if f['modified'] == min_date]
                        reason_for_tiebreak = f"Multiple files have oldest date ({min_date.strftime(DATE_FORMAT)})"

                elif rule == RULE_KEEP_NEWEST:
                    valid_files = [f for f in files_in_set if isinstance(f.get('modified'), datetime)]
                    if not valid_files:
                        log_func(self._("warning_rule_no_date", set_id=set_id_for_log, rule=rule, default=f"Warning: {set_id_for_log} - Cannot apply '{rule}': No valid dates. Keeping shortest path."))
                        candidates = list(files_in_set)
                        reason_for_tiebreak = "No valid dates found"
                    else:
                        max_date = max(f['modified'] for f in valid_files)
                        candidates = [f for f in valid_files if f['modified'] == max_date]
                        reason_for_tiebreak = f"Multiple files have newest date ({max_date.strftime(DATE_FORMAT)})"

                elif rule == RULE_KEEP_SUFFIX:
                    suffix_lower = suffix_value.lower() # Case-insensitive comparison
                    candidates = [f for f in files_in_set if f.get('path', '').lower().endswith(suffix_lower)]
                    if not candidates:
                        # If no file matches the suffix, log warning and consider all for tie-break
                         log_func(self._("warning_rule_no_suffix_match", set_id=set_id_for_log, suffix=suffix_value, default=f"Warning: {set_id_for_log} - No files match suffix '{suffix_value}'. Keeping shortest path."))
                         candidates = list(files_in_set)
                         reason_for_tiebreak = f"No files match suffix '{suffix_value}'"
                    else:
                         reason_for_tiebreak = f"Multiple files match suffix '{suffix_value}'"

                # --- Apply Tie-breaker if needed ---
                # Pass the specific reason if tie-breaking occurs
                full_reason = f"{set_id_for_log} - {reason_for_tiebreak}" if len(candidates) > 1 else reason_for_tiebreak
                keep_file_info = tie_break_shortest_path(candidates, full_reason)

                # --- Add Files to Delete List ---
                if keep_file_info and keep_file_info.get('path'):
                    keep_path = keep_file_info['path']
                    # Add all *other* files in the set to the deletion list
                    for f_info in files_in_set:
                        path = f_info.get('path')
                        if path and path != keep_path:
                            files_to_delete.append(path)
                else:
                     # Log if no file could be selected to keep (should be rare)
                     log_func(self._("warning_rule_failed_selection", set_id=set_id_for_log, rule=rule, default=f"Internal Warning: {set_id_for_log} - Rule '{rule}' failed to select file to keep. Skipping deletion for this set."))

            except Exception as e:
                 # Catch errors during rule application for a specific set
                 log_func(self._("error_rule_application", set_id=set_id_for_log, rule=rule, error=e, default=f"Error applying rule '{rule}' to {set_id_for_log}: {e}. Skipping deletion for this set."))
                 log_func(traceback.format_exc(limit=2)) # Log traceback for debugging

        return files_to_delete # Return the final list of paths to delete


    def start_delete_by_rule_thread(self):
        """ Handles 'Delete Marked Files' click. Validates, confirms, starts worker thread. """
        selected_rule = self.deletion_rule_var.get()

        # 1. Check if a rule is selected
        if not selected_rule:
             if self.master.winfo_exists(): messagebox.showerror(self._("error_input_title", default="Input Error"),
                                   self._("delete_no_rule_selected", default="Please select a deletion rule first."),
                                   master=self.master)
             return

        # 2. Check if suffix is provided if suffix rule is selected
        if selected_rule == RULE_KEEP_SUFFIX and not self.suffix_entry_var.get():
             if self.master.winfo_exists(): messagebox.showerror(self._("error_input_title", default="Input Error"),
                                   self._("delete_suffix_missing", default="Please enter the suffix for the 'Keep Suffix' rule."),
                                   master=self.master)
             return

        # 3. CRITICAL Check: Use the cached list generated by _apply_rule_to_treeview
        if not self.files_to_delete_cache:
             rule_name_key = f"rule_{selected_rule}"
             rule_name_for_log = self._(rule_name_key, default=selected_rule.title())
             if self.master.winfo_exists(): messagebox.showinfo(self._("delete_by_rule_button", default="Delete Marked Files"),
                                 self._("delete_rule_no_files", default="No files are currently marked for deletion based on the applied rule."),
                                 master=self.master)
             self.log_message(self._("delete_rule_no_files", default="No files marked for deletion.") + f" (Rule: {rule_name_for_log})")
             return

        # 4. Confirmation Dialog
        num_files_to_delete = len(self.files_to_delete_cache)
        rule_name_key = f"rule_{selected_rule}"
        rule_name_for_msg = self._(rule_name_key, default=selected_rule.title())
        confirm_msg_template = self._("delete_confirm_msg",
                                     rule_name=rule_name_for_msg,
                                     default="Delete marked files based on rule: '{rule_name}'?\nTHIS ACTION CANNOT BE UNDONE.")
        # Add file count to confirmation message
        confirm_msg = f"{confirm_msg_template}\n\n({num_files_to_delete} files will be permanently deleted)"

        confirm = False
        if self.master.winfo_exists():
            # Show warning icon and default to 'No'
            confirm = messagebox.askyesno(
                title=self._("delete_confirm_title", default="Confirm Deletion"),
                message=confirm_msg,
                icon='warning',
                default='no',
                master=self.master
            )

        if not confirm:
            self.log_message(self._("delete_cancelled", default="Deletion cancelled by user."))
            return # User cancelled

        # 5. Start Deletion Process
        self.log_message(self._("delete_starting", rule_name=rule_name_for_msg, default=f"Starting deletion (Rule: {rule_name_for_msg})..."))
        self.set_ui_state("deleting") # Disable UI during deletion

        # Pass a *copy* of the cached list to the worker thread
        # This prevents issues if the cache is somehow modified while the thread runs
        files_to_delete_list_copy = list(self.files_to_delete_cache)

        thread = threading.Thread(target=self._delete_worker,
                                  args=(files_to_delete_list_copy, rule_name_for_msg), # Pass rule name for logging context
                                  daemon=True)
        thread.start()


    def _delete_worker(self, files_to_delete, rule_name_for_log):
        """ Worker thread for deleting files based on the provided list. """
        # Check connection state within the thread
        if not self.finder or not self.finder.fs:
            self.log_message(self._("error_not_connected", default="Error: Connection lost before Deletion could execute."))
            if self.master.winfo_exists(): self.master.after(0, self.set_ui_state, 'normal') # Reset UI state
            return

        deleted_count = 0
        total_attempted = len(files_to_delete)
        deletion_error_occurred = False

        try:
            if not files_to_delete:
                 # This check is mostly redundant due to the check before starting thread, but safe to keep
                 self.log_message(self._("delete_rule_no_files", default="No files to delete.") + f" (Worker check; Rule: {rule_name_for_log})")
            else:
                # Call the finder's delete method
                deleted_count, total_attempted = self.finder.delete_files(files_to_delete)
                # Finder logs completion and errors internally now
                if deleted_count < total_attempted:
                    deletion_error_occurred = True # Track if any errors occurred

            # Schedule GUI update after deletion attempt finishes
            if self.master.winfo_exists():
                # IMPORTANT: Always clear results after delete attempt, forcing a re-scan.
                # Use master.after(10,...) to allow potential error popups or final log messages to appear first
                self.master.after(10, self.clear_results)
                # Add a log message confirming results are cleared
                self.master.after(20, lambda: self.log_message(self._("delete_results_cleared", default="Deletion finished. Results cleared.")))

        except Exception as e:
            # Catch unexpected errors during the deletion process itself
            deletion_error_occurred = True
            err_msg = self._("delete_error_during", error=e, default=f"Unexpected error during deletion process: {e}")
            self.log_message(err_msg)
            self.log_message(traceback.format_exc())
            if self.master.winfo_exists():
                 error_title = self._("error_title", default="Deletion Error")
                 # Schedule error popup in main thread
                 self.master.after(10, lambda et=error_title, em=err_msg: messagebox.showerror(et, em, master=self.master))
                 # On high-level error, just reset UI state, don't necessarily clear results
                 # unless the clear_results is called in the finally block of the main try.
                 # Let's keep the clear_results in the successful path only for now.
                 self.master.after(0, self.set_ui_state, 'normal')
        finally:
            # Ensure UI state is reset if not handled by specific paths above
            # However, the successful path already schedules clear_results which calls set_ui_state('normal')
            # And the error path schedules set_ui_state('normal')
            # So, an explicit call here might be redundant but safe.
            if self.master.winfo_exists():
                 # Check if clear_results wasn't already scheduled
                 # This logic is tricky, maybe rely on the paths above.
                 # Let's remove the finally block state reset for now to avoid potential conflicts.
                 pass



    def save_duplicates_report(self):
        """ Saves the report of FOUND duplicate file sets to a text file. """
        # Check if there are any duplicates found and stored
        if not self.duplicate_sets:
            if self.master.winfo_exists(): messagebox.showinfo(self._("save_list_button", default="Save Report"),
                                self._("save_report_no_data", default="No duplicate sets found to save."),
                                master=self.master)
            self.log_message(self._("save_report_no_data", default="No duplicate sets available to save."))
            return

        # Generate a default filename with timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        initial_filename = f"duplicates_report_{timestamp}.txt"

        file_path = None
        try:
             # Open the "Save As" dialog
             file_path = filedialog.asksaveasfilename(
                 defaultextension=".txt",
                 filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
                 title=self._("save_list_button", default="Save Found Duplicates Report As..."),
                 initialfile=initial_filename,
                 parent=self.master # Ensure dialog is modal to the main window
             )
        except Exception as fd_e:
             # Catch errors opening the dialog itself (rare)
             self.log_message(f"Error opening save file dialog: {fd_e}")
             if self.master.winfo_exists(): messagebox.showerror(self._("error_title", default="Error"), f"Could not open save dialog: {fd_e}", master=self.master)
             return

        # If the user selected a file (didn't cancel)
        if file_path:
            # Call the finder's method to write the report
            # Finder handles logging success/failure internally now
            success = self.finder.write_duplicates_report(self.duplicate_sets, file_path)
            # Show confirmation popup on success
            if success and self.master.winfo_exists():
                messagebox.showinfo(
                    self._("save_list_button", default="Report Saved"),
                    # Use the saved message from finder's log if desired, or keep generic one
                    self._("save_report_saved", file=os.path.basename(file_path), default="Report saved successfully."),
                    master=self.master)
        else:
             # User cancelled the save dialog
             self.log_message("Save report operation cancelled by user.")


    def show_cloud_file_types(self):
        """ Handles 'Show Cloud File Types' click. Validates prerequisites and starts worker thread. """
        # 1. Check if Matplotlib is available
        if not MATPLOTLIB_AVAILABLE:
            if self.master.winfo_exists(): messagebox.showwarning(self._("chart_error_title", default="Chart Error"),
                                   self._("chart_error_no_matplotlib", default="Matplotlib library not found. Please install it."),
                                   master=self.master)
            self.log_message(self._("chart_error_no_matplotlib", default="Matplotlib not found, cannot show chart."))
            return

        # 2. Check CloudDrive Connection
        if not self.finder or not self.finder.fs:
            if self.master.winfo_exists(): messagebox.showwarning(self._("chart_error_title", default="Chart Error"),
                                   self._("chart_error_no_connection", default="Not connected to CloudDrive. Cannot scan for chart data."),
                                   master=self.master)
            self.log_message(self._("chart_error_no_connection", default="Not connected, cannot generate chart."))
            return

        # 3. Validate Required Inputs for Chart Scan
        scan_path_raw = self.string_vars["scan_path"].get()
        mount_point_raw = self.string_vars["mount_point"].get()
        if not scan_path_raw or not mount_point_raw:
             error_msg = self._("error_input_missing_chart", default="Root Path to Scan and Mount Point are required to generate the chart.")
             if self.master.winfo_exists(): messagebox.showwarning(self._("error_input_title", default="Input Error"), error_msg, master=self.master)
             self.log_message(error_msg)
             return

        # 4. Validate Path Characters
        paths_to_check_chart = {
            "scan_path": scan_path_raw,
            "mount_point": mount_point_raw
        }
        if not self._check_path_chars(paths_to_check_chart):
             # Error message handled by _check_path_chars
             return

        # 5. Start Worker Thread
        self.log_message("Starting scan for file type chart data...")
        self.set_ui_state("charting") # Disable UI during chart scan

        thread = threading.Thread(target=self._show_cloud_file_types_worker,
                                  args=(scan_path_raw, mount_point_raw), daemon=True)
        thread.start()


    def _show_cloud_file_types_worker(self, scan_path_raw, mount_point_raw):
        """ Worker thread to scan cloud files, count types, and schedule chart creation. """
        fs_dir_path = None
        file_counts = collections.Counter()
        total_files = 0
        scan_error = None

        try:
             # --- Pre-scan Checks ---
             if not self.finder:
                  self.log_message("Error: Finder object not initialized. Cannot perform chart scan.")
                  if self.master.winfo_exists(): self.master.after(0, self.set_ui_state, 'normal')
                  return

             # Calculate the effective filesystem path
             fs_dir_path = self.finder.calculate_fs_path(scan_path_raw, mount_point_raw)
             if fs_dir_path is None:
                 self.log_message("Chart generation aborted due to path calculation error.")
                 if self.master.winfo_exists(): self.master.after(0, self.set_ui_state, 'normal')
                 return

             # Check connection again within thread
             if not self.finder.fs:
                 self.log_message(self._("chart_error_no_connection", default="Cannot chart: Connection lost before scan."))
                 if self.master.winfo_exists(): self.master.after(0, self.set_ui_state, 'normal')
                 return

             # --- Perform Scan ---
             self.log_message(self._("chart_status_scanning_cloud", path=fs_dir_path, default=f"Scanning '{fs_dir_path}' for file types..."))
             scan_start_time = time.time()
             try:
                 # Use walk_path to iterate through files recursively
                 for _, _, filenames in self.finder.fs.walk_path(fs_dir_path):
                     for filename_obj in filenames:
                         try:
                            filename = str(filename_obj) # Ensure filename is a string
                            total_files += 1
                            # Get extension, handle files with no extension
                            _root, ext = os.path.splitext(filename)
                            ext_label = ext.lower() if ext else self._("chart_label_no_extension", default="[No Ext]")
                            file_counts[ext_label] += 1
                         except Exception as inner_e:
                             # Log errors processing individual filenames during scan
                             self.log_message(f"Warning: Error processing filename '{filename_obj}' during chart scan: {inner_e}")

             except Exception as e:
                 # Catch errors during the walk_path operation itself
                 scan_error = e
                 error_msg = self._("chart_error_cloud_scan", path=fs_dir_path, error=e, default=f"Error scanning cloud path '{fs_dir_path}' for chart data: {e}")
                 self.log_message(error_msg)
                 self.log_message(f"Chart Scan Error Details: {traceback.format_exc()}")
                 if self.master.winfo_exists():
                      error_title = self._("chart_error_title", default="Chart Error")
                      # Schedule error popup in main thread
                      self.master.after(10, lambda et=error_title, em=error_msg: messagebox.showerror(et, em, master=self.master))

             scan_duration = time.time() - scan_start_time
             if not scan_error:
                 self.log_message(f"Chart data scan completed in {scan_duration:.2f}s. Found {total_files} files.")

             # --- Schedule GUI Update (Chart Creation or Message) ---
             def update_gui_after_chart_scan():
                 # Check if master window still exists before proceeding
                 if not self.master.winfo_exists(): return
                 # If scan failed, UI state reset is handled in finally block, just return
                 if scan_error: return

                 if not file_counts:
                     # If scan succeeded but found no files
                     no_files_msg = self._("chart_status_no_files_found", path=fs_dir_path, default=f"No files found in '{fs_dir_path}'. Cannot generate chart.")
                     self.log_message(no_files_msg)
                     if self.master.winfo_exists(): messagebox.showinfo(self._("chart_info_title", default="Chart Info"), no_files_msg, master=self.master)
                     return # Don't try to create chart

                 # If scan succeeded and files were found, proceed to create chart
                 self.log_message(self._("chart_status_generating", count=len(file_counts), total=total_files, default=f"Generating chart for {len(file_counts)} types ({total_files} files)..."))
                 self._create_pie_chart_window(file_counts, fs_dir_path)

             # Schedule the update function to run in the main GUI thread
             if self.master.winfo_exists():
                 self.master.after(0, update_gui_after_chart_scan)

        except Exception as e:
             # Catch unexpected errors during worker setup/path calculation
             err_msg = f"Unexpected error during chart worker setup for path '{scan_path_raw}': {e}"
             self.log_message(err_msg)
             self.log_message(traceback.format_exc())
             if self.master.winfo_exists():
                 error_title = self._("chart_error_title", default="Chart Error")
                 self.master.after(10, lambda et=error_title, em=err_msg: messagebox.showerror(et, em, master=self.master))
        finally:
             # ALWAYS schedule UI state reset back to normal after worker finishes or errors out
             if self.master.winfo_exists():
                 self.master.after(0, self.set_ui_state, 'normal')


    def _create_pie_chart_window(self, counts, display_path):
        """ Creates and displays the file type pie chart in a new Toplevel window. """
        if not MATPLOTLIB_AVAILABLE:
            # Double-check availability right before creation
            self.log_message("Error: Matplotlib became unavailable before chart creation.")
            if self.master.winfo_exists():
                messagebox.showerror(self._("chart_error_title", default="Chart Error"), self._("chart_error_no_matplotlib", default="Matplotlib library not found."), master=self.master)
            return

        chart_window = None # Initialize variable to hold the Toplevel window
        try:
            # --- Matplotlib Settings (Attempt, ignore errors if font setting fails) ---
            try:
                # Ensure minus signs display correctly with CJK fonts
                matplotlib.rcParams['axes.unicode_minus'] = False
                # Get current font list and try to prepend preferred CJK fonts
                current_sans_serif = matplotlib.rcParams['font.sans-serif']
                preferred_fonts = ['SimHei', 'Microsoft YaHei', 'MS Gothic', 'Malgun Gothic', 'Arial Unicode MS', 'sans-serif']
                # Build final list, adding preferred only if not already present
                final_font_list = preferred_fonts + [f for f in current_sans_serif if f not in preferred_fonts]
                matplotlib.rcParams['font.sans-serif'] = final_font_list
            except Exception as mpl_set_err:
                self.log_message(f"Warning: Issue setting Matplotlib rcParams (e.g., fonts): {mpl_set_err}")

            # --- Data Preparation for Chart ---
            top_n = 20 # Show top N extensions + "Others"
            total_count = sum(counts.values())
            # Get extensions sorted by count, descending
            sorted_counts = counts.most_common()

            chart_labels = [] # Labels for wedges (extensions or "Others")
            chart_sizes = []  # Corresponding counts for wedges
            legend_labels_with_counts = [] # Labels for the legend (e.g., ".txt (150)")

            others_label = self._("chart_label_others", default="Others")
            others_count = 0
            others_sources = [] # List of extensions grouped into "Others"

            # Group less frequent extensions into "Others" if necessary
            if len(sorted_counts) > top_n:
                top_items = sorted_counts[:top_n]
                chart_labels = [item[0] for item in top_items]
                chart_sizes = [item[1] for item in top_items]
                legend_labels_with_counts = [f'{item[0]} ({item[1]})' for item in top_items]

                other_items = sorted_counts[top_n:]
                others_count = sum(item[1] for item in other_items)
                others_sources = [item[0] for item in other_items]

                if others_count > 0:
                    chart_labels.append(others_label)
                    chart_sizes.append(others_count)
                    legend_labels_with_counts.append(f'{others_label} ({others_count})')
            else:
                # If fewer than top_n extensions, show all
                chart_labels = [item[0] for item in sorted_counts]
                chart_sizes = [item[1] for item in sorted_counts]
                legend_labels_with_counts = [f'{item[0]} ({item[1]})' for item in sorted_counts]

            # Log if grouping occurred
            if others_count > 0:
                 self.log_message(f"Chart Note: Grouped {len(others_sources)} smaller file types ({others_count} files) into '{others_label}'.")

            # --- Create Chart Window (Toplevel) ---
            chart_window = Toplevel(self.master)
            chart_window.title(self._("chart_window_title", path=display_path, default=f"File Types in {display_path}"))
            chart_window.geometry("900x700") # Initial size
            chart_window.minsize(600, 400) # Minimum size

            # --- Create Matplotlib Figure & Axes ---
            fig = Figure(figsize=(9, 7), dpi=100) # Adjust figsize as needed
            ax = fig.add_subplot(111)

            # --- Generate Pie Chart ---
            explode_value = 0.02 # Slight separation between slices
            # Create explode list, same length as labels
            explode_list = [explode_value] * len(chart_labels)
            # Don't explode the 'Others' slice if it exists
            if others_count > 0 and others_label in chart_labels:
                try:
                    others_index = chart_labels.index(others_label)
                    explode_list[others_index] = 0
                except ValueError: pass # Should not happen, but safe check

            # Create the pie chart
            wedges, texts, autotexts = ax.pie(
                chart_sizes,
                explode=explode_list,
                labels=None, # Use legend instead of direct labels on slices
                autopct=lambda pct: f"{pct:.1f}%" if pct > 1.5 else '', # Show percentage only for slices > 1.5%
                startangle=140, # Rotate start position
                pctdistance=0.85, # Position of percentage labels inside slice
                # Add doughnut effect and white edges for visual separation
                wedgeprops=dict(width=0.6, edgecolor='w')
            )
            ax.axis('equal') # Ensure pie is drawn as a circle

            # --- Create Legend ---
            legend = ax.legend(wedges, legend_labels_with_counts, # Use the labels with counts
                      title=self._("chart_legend_title", default="File Extensions"),
                      loc="center left",
                      bbox_to_anchor=(1, 0, 0.5, 1), # Position legend outside the plot area to the right
                      fontsize='small', # Adjust font size
                      frameon=True, # Add frame around legend
                      labelspacing=0.8 # Adjust spacing between legend items
                      )

            # --- Style Percentage Labels (Autotexts) ---
            for autotext in autotexts:
                if autotext.get_text(): # Check if text exists (due to lambda autopct)
                    autotext.set_color('white')
                    autotext.set_fontsize(8)
                    autotext.set_weight('bold')
                    # Add semi-transparent background for readability
                    autotext.set_bbox(dict(facecolor='black', alpha=0.5, pad=1, edgecolor='none'))

            # --- Add Chart Title ---
            chart_title = self._("chart_window_title", path=display_path, default=f"File Types in {display_path}")
            chart_title += f"\n(Total Files: {total_count})" # Add total file count
            ax.set_title(chart_title, pad=20, fontsize=12) # Adjust padding and font size

            # --- Adjust Layout to Prevent Overlap ---
            try:
                 # Adjust right margin to make space for the legend
                 fig.tight_layout(rect=[0, 0, 0.75, 1])
            except Exception as layout_err:
                print(f"Warning: Chart layout adjustment failed: {layout_err}.")
                self.log_message(f"Warning: Chart layout adjustment failed: {layout_err}")

            # --- Embed Matplotlib Chart in Tkinter Window ---
            canvas = FigureCanvasTkAgg(fig, master=chart_window) # Create Tkinter canvas
            canvas_widget = canvas.get_tk_widget() # Get the Tkinter widget from the canvas

            # Add Matplotlib navigation toolbar (optional but useful)
            toolbar = NavigationToolbar2Tk(canvas, chart_window)
            toolbar.update() # Finalize toolbar

            # Pack toolbar and canvas into the Toplevel window
            toolbar.pack(side=tk.BOTTOM, fill=tk.X)
            canvas_widget.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

            canvas.draw() # Draw the chart onto the canvas

            # Bring the chart window to the front
            chart_window.focus_set()
            chart_window.lift()

        except Exception as e:
            # Catch any error during chart creation/display
            error_msg = f"Error creating or displaying chart window: {e}"
            self.log_message(error_msg)
            self.log_message(traceback.format_exc())
            if self.master.winfo_exists():
                messagebox.showerror(self._("chart_error_title", default="Chart Error"), error_msg, master=self.master)
            # Destroy the Toplevel window if it was created but failed mid-process
            if chart_window and chart_window.winfo_exists():
                 try: chart_window.destroy()
                 except: pass


# --- Main Execution Block ---
if __name__ == "__main__":
    # --- Attempt High DPI Awareness (Windows) ---
    try:
        from ctypes import windll
        try:
            # Windows 8.1+
            windll.shcore.SetProcessDpiAwareness(1)
        except AttributeError:
            try:
                # Windows Vista/7
                windll.user32.SetProcessDPIAware()
            except AttributeError:
                # DPI awareness setting not available/needed
                pass
    except (ImportError, AttributeError):
        # Not on Windows or ctypes issue
        pass

    # --- Initialize Tkinter Root Window ---
    root = tk.Tk()
    try:
        # --- Basic Sanity Check for Translations ---
        # Ensure core keys exist to prevent crashing if translation file is broken
        if not translations["en"].get("window_title") or not translations["zh"].get("window_title"):
            print("ERROR: Core translations appear missing. Exiting.")
            # Attempt to show a GUI error message even if translations are broken
            try:
                 root_err = tk.Tk(); root_err.withdraw() # Temporary hidden window
                 messagebox.showerror("Fatal Error", "Core translation strings missing. Cannot start application.", master=root_err)
                 root_err.destroy()
            except: pass # Ignore errors showing the error message itself
            sys.exit(1) # Exit cleanly

        # --- Create and Run the Application ---
        app = DuplicateFinderApp(root)
        root.mainloop() # Start the Tkinter event loop

    except Exception as main_e:
        # --- Catch-all for Fatal Errors during Application Initialization or Runtime ---
        print("FATAL APPLICATION ERROR:")
        print(traceback.format_exc()) # Print full traceback to console
        # Attempt to show a final error message in a popup
        try:
            root_err = tk.Tk(); root_err.withdraw() # Temporary hidden window
            messagebox.showerror("Fatal Error", f"A critical error occurred and the application must close:\n\n{main_e}", master=root_err)
            root_err.destroy()
        except Exception as mb_err:
            # If even the error popup fails, log it
            print(f"Could not display fatal error in GUI: {mb_err}")
        sys.exit(1) # Exit with error status
