# -*- coding: utf-8 -*-
import os
import configparser
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog, Menu, Toplevel
import threading
import time
from datetime import datetime, timezone, timedelta
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
    except AttributeError: # More specific exception
        # If not running as a PyInstaller bundle, use the script's directory
        base_path = os.path.abspath(os.path.dirname(__file__))
    except Exception: # Fallback generic exception
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

# <<< MODIFICATION: Added default API address constant >>>
DEFAULT_API_ADDRESS = "127.0.0.1:19798"
# <<< END MODIFICATION >>>

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
        "scan_paths_label": "Root Paths to Scan:",
        "add_path_button": "Add Path",
        "remove_path_button": "Remove Selected",
        "mount_point_label": "CloudDrive Mount Point:",
        "load_config_button": "Load Config",
        "save_config_button": "Save Config",
        "test_connection_button": "Test Connection",
        "find_button": "Find Duplicates",
        # <<< MODIFICATION: Changed title slightly for clarity >>>
        "rules_title": "Apply Suggestion Rule (Click 'Action' to Manually Change)",
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
        # <<< MODIFICATION: Renamed button slightly >>>
        "delete_selected_button": "Delete Marked Files",
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
        "status_config_not_found": "Config file '{file}' not found. Using defaults.",
        "status_config_section_missing": "Config file loaded, but '[config]' section is missing.",
        "status_connecting": "Attempting connection...",
        "status_connect_success": "Connection successful.",
        "status_scan_progress": "Path '{path}': Scanned {count} items... Found {video_count} videos.",
        "status_populating_tree": "Populating list with {count} duplicate sets...",
        "status_tree_populated": "Results list populated.",
        "status_clearing_tree": "Clearing results list and rule selection...",
        # <<< MODIFICATION: Adjusted rule log message >>>
        "status_applying_rule": "Applying suggestion rule '{rule_name}' to {count} sets...",
        "status_rule_applied": "Rule suggestion applied. {delete_count} files initially marked for deletion. Click 'Action' column to change.",
        "find_starting": "Starting duplicate file scan across {num_paths} path(s)...",
        "find_scan_path_start": "Scanning path: '{path}'...",
        "find_complete_found": "Scan complete. Found {count} duplicate sets across all paths.",
        "find_complete_none": "Scan complete. No duplicate video files found based on SHA1 hash across all paths.",
        "find_error_during": "Error during duplicate scan: {error}",
        "find_error_processing_path": "Error processing scan path '{path}': {error}. Skipping this path.", # Added
        # <<< MODIFICATION: Renamed delete log messages >>>
        "delete_starting_selected": "Starting deletion of manually marked files...",
        "delete_finished": "Deletion complete. Attempted to delete {total_marked} files. Successfully deleted {deleted_count}.",
        "delete_results_cleared": "Deletion process finished. Results list cleared, please re-scan if needed.", # Added completion message
        "delete_error_during": "Error during deletion process: {error}",
        "delete_error_file": "Error deleting {path}: {error}",
        "delete_cancelled": "Deletion cancelled by user.",
        "delete_confirm_title": "Confirm Deletion",
        # <<< MODIFICATION: Changed confirm message wording >>>
        "delete_confirm_msg": "Permanently delete all files marked 'Delete' in the list?\nTHIS ACTION CANNOT BE UNDONE.",
        "delete_no_rule_selected": "No deletion rule selected.", # Still relevant for applying suggestions
        "delete_no_files_marked": "No files are currently marked for deletion in the list.", # New message
        "delete_suffix_missing": "Suffix is required for the 'Keep Suffix' rule suggestion.",
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
        "error_no_duplicates_found": "No duplicates were found or displayed. Cannot apply deletion rule.", # Still used internally for rule logic
        "error_not_connected": "Error: Not connected to CloudDrive2. Please test connection first.",
        "error_path_calc_failed": "Error: Could not determine a valid cloud scan path from Scan Path '{scan}' and Mount Point '{mount}'. Check inputs.",
        "error_input_missing": "API Address, Account, Mount Point are required. At least one Scan Path must be added.",
        "error_input_missing_conn": "API Address, Account, and Mount Point are required for connection test.",
        "error_input_missing_chart": "Mount Point is required and at least one Scan Path must be added to generate the chart.",
        "error_config_read": "Error reading config file: {error}",
        "error_config_save": "Could not write config file: {error}",
        "error_unexpected": "Unexpected error: {error}",
        "error_icon_load": "Error loading application icon '{path}': {error}",
        "warning_path_mismatch": "Warning: Could not determine a valid cloud path based on Scan Path ('{scan}') and Mount Point ('{mount}'). Effective scan path for this entry might be invalid. Please verify inputs.",
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
        "chart_status_scan_paths_start": "Starting scan for chart data across {num_paths} path(s)...", # Added
        "chart_status_scan_path_complete": "Finished scanning path '{path}'.", # Added
        "chart_status_generating": "Scan complete across all paths. Generating chart for {count} file types ({total} files)...",
        "chart_status_no_files_found": "Scan complete across all paths. No files found. Cannot generate chart.",
        "chart_window_title": "File Types in Scanned Paths", # Updated
        "chart_legend_title": "File Extensions",
        "chart_label_others": "Others",
        "chart_label_no_extension": "[No Ext]",
        "tie_break_log_prefix": "Tie-Break:", # Added for rule tie-break logging
        "status_test_connection_step": "Testing connection by attempting to list root directory ('/')...",
        "status_scan_finished_duration": "Scan for path '{path}' finished in {duration:.2f} seconds.",
        "status_scan_summary_items": "Path '{path}': Total items encountered: {count}. Video files processed: {video_count}.",
        "status_scan_warnings": "Path '{path}': WARNING: {details}.",
        "warning_hash_missing": "Warning: Hash data missing in attributes for '{path}'. KeyError: {key_error}. Skipping.",
        "warning_hash_short": "Warning: Suspiciously short SHA1 hash ('{hash}') found for '{path}'. Skipping.",
        "warning_size_invalid": "Warning: Invalid size value '{size}' for {path}. Using 0.",
        "status_delete_attempting": "Attempting to delete {count} marked files...",
        "status_deleting_file": "Deleting [{current}/{total}]: {path}",
        "warning_delete_failures": "WARNING: Failed to delete {count} file(s):",
        "warning_delete_failures_more": "  ... and {count} more.",
        "status_delete_no_files": "No files provided for deletion.",
        "warning_tie_break": "{prefix} {reason}. Kept '{filename}' ({detail})",
        "warning_rule_no_date": "Warning: {set_id} - Cannot apply suggestion '{rule}': No valid dates. Defaulting to shortest path.",
        "warning_rule_no_suffix_match": "Warning: {set_id} - No files match suffix '{suffix}'. Defaulting to shortest path.",
        "warning_rule_failed_selection": "Internal Warning: {set_id} - Rule '{rule}' failed to select file to keep. Skipping suggestion for this set.",
        "error_rule_application": "Error applying suggestion rule '{rule}' to {set_id}: {error}. Skipping suggestion for this set.",
        "log_debug_calc_path": "[Debug] Calculated effective cloud scan path: '{fs_path}' from Scan='{scan_raw}', Mount='{mount_raw}'",
        "log_debug_process_video": "[Debug] Processing Video: {path}",
        "log_debug_attrs_received": "[Debug] Attrs received for {filename}: {attrs}",
        "log_debug_raw_sha1": "[Debug] Raw SHA1 value (key '2'): {sha1}",
        "log_debug_invalid_sha1_type": "[Debug] Invalid SHA1: Not a string ('{sha1}') for {path}. Skipping.",
        "log_debug_invalid_sha1_empty": "[Debug] Invalid SHA1: Empty string for {path}. Skipping.",
        "log_debug_standardized_sha1": "[Debug] Standardized SHA1: {sha1}",
        "log_debug_hash_missing_or_invalid": "[Debug] 'fileHashes' or key '2' missing or hash invalid for {path}. SHA1 is None.",
        "log_debug_skipping_no_sha1": "[Debug] SKIPPING file {filename} due to missing or invalid SHA1.",
        "log_debug_storing_info": "[Debug] Storing file info for {filename} with SHA1: {sha1}",
        "log_debug_merging_results": "[Debug] Merging results from path '{path}'. Current total sets: {count}.",
        "select_scan_path_dialog_title": "Select Root Directory to Scan",
        "error_no_scan_paths_added": "Error: No scan paths have been added. Please add at least one path.",
        "scan_path_already_exists": "Path '{path}' already exists in the list.",
        # <<< MODIFICATION: Added new translation >>>
        "info_last_keep_in_set": "Info: Cannot mark '{filename}' for deletion as it's the only file marked 'Keep' in Set {set_id}.",
    },
    "zh": {
        "window_title": "CloudDrive2 重复视频查找与删除工具",
        "config_title": "配置",
        "address_label": "API 地址:",
        "account_label": "账户:",
        "password_label": "密码:",
        "scan_paths_label": "要扫描的根路径:",
        "add_path_button": "添加路径",
        "remove_path_button": "删除选中",
        "mount_point_label": "CloudDrive 挂载点:",
        "load_config_button": "加载配置",
        "save_config_button": "保存配置",
        "test_connection_button": "测试连接",
        "find_button": "查找重复项",
        # <<< MODIFICATION: Changed title slightly for clarity >>>
        "rules_title": "应用建议规则 (点击'操作'列手动更改)",
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
        # <<< MODIFICATION: Renamed button slightly >>>
        "delete_selected_button": "删除标记文件",
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
        "status_config_not_found": "未找到配置文件 '{file}'。将使用默认值。",
        "status_config_section_missing": "配置文件已加载，但缺少 '[config]' 部分。",
        "status_connecting": "正在尝试连接...",
        "status_connect_success": "连接成功。",
        "status_scan_progress": "路径 '{path}': 已扫描 {count} 个项目... 找到 {video_count} 个视频。",
        "status_populating_tree": "正在使用 {count} 个重复集合填充列表...",
        "status_tree_populated": "结果列表已填充。",
        "status_clearing_tree": "正在清除结果列表和规则选择...",
        # <<< MODIFICATION: Adjusted rule log message >>>
        "status_applying_rule": "正在对 {count} 个集合应用建议规则 '{rule_name}'...",
        "status_rule_applied": "规则建议已应用。初始有 {delete_count} 个文件被标记为删除。点击'操作'列进行更改。",
        "find_starting": "开始在 {num_paths} 个路径中扫描重复文件...",
        "find_scan_path_start": "正在扫描路径: '{path}'...",
        "find_complete_found": "扫描完成。在所有路径中共找到 {count} 个重复集合。",
        "find_complete_none": "扫描完成。未在所有路径中根据 SHA1 哈希找到重复的视频文件。",
        "find_error_during": "扫描重复项期间出错: {error}",
        "find_error_processing_path": "处理扫描路径 '{path}' 时出错: {error}。正在跳过此路径。", # 新增
        # <<< MODIFICATION: Renamed delete log messages >>>
        "delete_starting_selected": "开始删除手动标记的文件...",
        "delete_finished": "删除完成。尝试删除 {total_marked} 个文件。成功删除了 {deleted_count} 个。",
        "delete_results_cleared": "删除过程结束。结果列表已清除，如需请重新扫描。",
        "delete_error_during": "删除过程中出错: {error}",
        "delete_error_file": "删除 {path} 时出错: {error}",
        "delete_cancelled": "用户取消了删除操作。",
        "delete_confirm_title": "确认删除",
        # <<< MODIFICATION: Changed confirm message wording >>>
        "delete_confirm_msg": "永久删除列表中所有标记为“删除”的文件吗？\n此操作无法撤销。",
        "delete_no_rule_selected": "未选择删除规则。", # Still relevant for applying suggestions
        "delete_no_files_marked": "列表中当前没有文件被标记为删除。", # New message
        "delete_suffix_missing": "“保留后缀”规则建议需要填写后缀。",
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
        "error_rule_title": "规则错误",
        "error_connect": "连接 CloudDrive2 API '{address}' 时出错: {error}",
        "error_scan_path": "遍历云端路径 '{path}' 时发生严重错误: {error}",
        "error_get_attrs": "获取 '{path}' 的属性/哈希时出错: {error}",
        "error_parse_date": "警告: 无法解析 '{path}' 的修改日期: {error}。将跳过此文件的日期比较。",
        "error_no_duplicates_found": "未找到或显示重复项。无法应用删除规则。", # Still used internally
        "error_not_connected": "错误：未连接到 CloudDrive2。请先测试连接。",
        "error_path_calc_failed": "错误：无法根据扫描路径 '{scan}' 和挂载点 '{mount}' 确定有效的云端扫描路径。请检查输入。",
        "error_input_missing": "API 地址、账户、挂载点为必填项。必须添加至少一个扫描路径。",
        "error_input_missing_conn": "测试连接需要 API 地址、账户和挂载点。",
        "error_input_missing_chart": "生成图表需要挂载点，且必须添加至少一个扫描路径。",
        "error_config_read": "读取配置文件时出错: {error}",
        "error_config_save": "无法写入配置文件: {error}",
        "error_unexpected": "意外错误: {error}",
        "error_icon_load": "加载应用程序图标 '{path}' 时出错: {error}",
        "warning_path_mismatch": "警告：无法根据扫描路径 ('{scan}') 和挂载点 ('{mount}') 确定有效的云端路径。此条目的有效扫描路径可能无效。请核对输入。",
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
        "chart_status_scan_paths_start": "开始在 {num_paths} 个路径中扫描图表数据...",
        "chart_status_scan_path_complete": "完成扫描路径 '{path}'。",
        "chart_status_generating": "所有路径扫描完成。正在为 {count} 种文件类型 ({total} 个文件) 生成图表...",
        "chart_status_no_files_found": "所有路径扫描完成。未找到任何文件。无法生成图表。",
        "chart_window_title": "扫描路径中的文件类型分布",
        "chart_legend_title": "文件扩展名",
        "chart_label_others": "其他",
        "chart_label_no_extension": "[无扩展名]",
        "tie_break_log_prefix": "规则冲突解决:",
        "status_test_connection_step": "正在通过尝试列出根目录 ('/') 来测试连接...",
        "status_scan_finished_duration": "路径 '{path}' 的扫描耗时 {duration:.2f} 秒完成。",
        "status_scan_summary_items": "路径 '{path}': 共遇到 {count} 个项目。已处理 {video_count} 个视频文件。",
        "status_scan_warnings": "路径 '{path}': 警告: {details}。",
        "warning_hash_missing": "警告：'{path}' 的属性中缺少哈希数据。KeyError: {key_error}。正在跳过。",
        "warning_hash_short": "警告：为 '{path}' 找到了可疑的短 SHA1 哈希 ('{hash}')。正在跳过。",
        "warning_size_invalid": "警告：{path} 的大小值 '{size}' 无效。使用 0。",
        "status_delete_attempting": "正在尝试删除 {count} 个标记的文件...",
        "status_deleting_file": "正在删除 [{current}/{total}]: {path}",
        "warning_delete_failures": "警告：未能删除 {count} 个文件：",
        "warning_delete_failures_more": "  ... 以及另外 {count} 个。",
        "status_delete_no_files": "没有提供用于删除的文件。",
        "warning_tie_break": "{prefix} {reason}。保留了 '{filename}' ({detail})",
        "warning_rule_no_date": "警告：{set_id} - 无法应用建议 '{rule}'：无有效日期。将默认为最短路径。",
        "warning_rule_no_suffix_match": "警告：{set_id} - 没有文件匹配后缀 '{suffix}'。将默认为最短路径。",
        "warning_rule_failed_selection": "内部警告：{set_id} - 规则 '{rule}' 未能选择要保留的文件。跳过此集合的建议。",
        "error_rule_application": "将建议规则 '{rule}' 应用于 {set_id} 时出错：{error}。跳过此集合的建议。",
        "log_debug_calc_path": "[调试] 根据 Scan='{scan_raw}', Mount='{mount_raw}' 计算出的有效云扫描路径: '{fs_path}'",
        "log_debug_process_video": "[调试] 正在处理视频文件: {path}",
        "log_debug_attrs_received": "[调试] 收到 {filename} 的属性: {attrs}",
        "log_debug_raw_sha1": "[调试] 原始 SHA1 值 (键 '2'): {sha1}",
        "log_debug_invalid_sha1_type": "[调试] 无效 SHA1: 不是字符串 ('{sha1}')，路径: {path}。正在跳过。",
        "log_debug_invalid_sha1_empty": "[调试] 无效 SHA1: 空字符串，路径: {path}。正在跳过。",
        "log_debug_standardized_sha1": "[调试] 标准化 SHA1: {sha1}",
        "log_debug_hash_missing_or_invalid": "[调试] 路径 {path} 的 'fileHashes' 或键 '2' 缺失，或哈希无效。SHA1 为 None。",
        "log_debug_skipping_no_sha1": "[调试] 因 SHA1 缺失或无效，正在跳过文件 {filename}。",
        "log_debug_storing_info": "[调试] 正在存储文件 {filename} 的信息，SHA1 为: {sha1}",
        "log_debug_merging_results": "[调试] 正在合并路径 '{path}' 的结果。当前总集合数: {count}。",
        "select_scan_path_dialog_title": "选择要扫描的根目录",
        "error_no_scan_paths_added": "错误：尚未添加扫描路径。请至少添加一个路径。",
        "scan_path_already_exists": "路径 '{path}' 已存在于列表中。",
        # <<< MODIFICATION: Added new translation >>>
        "info_last_keep_in_set": "提示：无法将 '{filename}' 标记为删除，因为它是集合 {set_id} 中唯一标记为“保留”的文件。",
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
        self._raw_scan_paths = []
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
            raw_scan_paths, # Expects a list now
            raw_mount_point,
            progress_callback=None,
    ):
        """ Sets configuration and attempts to establish+test connection. """
        self.clouddrvie2_address = clouddrvie2_address
        self.clouddrive2_account = clouddrive2_account
        self.clouddrive2_passwd = clouddrive2_passwd
        self._raw_scan_paths = list(raw_scan_paths) if raw_scan_paths else [] # Store a copy
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
            # Use generic labels as we don't have the UI context here
            scan_path_label = "Scan Path"
            mount_point_label = "Mount Point"
            if scan_path_issues: all_issues.append(f"'{scan_path_label}' ('{scan_path_raw}'): {', '.join(scan_path_issues)}")
            if mount_point_issues: all_issues.append(f"'{mount_point_label}' ('{mount_point_raw}'): {', '.join(mount_point_issues)}")
            # Log as an error since it prevents calculation
            log_msg = self._("path_warning_suspicious_chars", default="Suspicious chars detected!").split('\n')[0] # Get first line
            self.log(f"ERROR: {log_msg} Details: {'; '.join(all_issues)}")
            return None # Indicate failure due to bad characters

        # --- Normalization ---
        # Replace backslashes, strip whitespace, remove trailing slashes
        scan_path_norm = scan_path_raw.replace('\\', '/').strip().rstrip('/') if scan_path_raw else ''
        mount_point_norm = mount_point_raw.replace('\\', '/').strip().rstrip('/') if mount_point_raw else ''
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
        Scans the configured cloud paths for duplicate video files using SHA1 hash.
        Handles path construction and standardizes SHA1 hash case. Aggregates results.
        """
        if not self.fs:
            self.log(self._("error_not_connected", default="Error: Not connected to CloudDrive. Cannot scan."))
            return {}

        if not self._raw_scan_paths:
            self.log(self._("error_no_scan_paths_added", default="Error: No scan paths specified. Aborting scan."))
            return {}

        self.log(self._("find_starting", num_paths=len(self._raw_scan_paths), default=f"Starting duplicate file scan across {len(self._raw_scan_paths)} path(s)..."))

        # --- Aggregated results across all paths ---
        all_potential_duplicates = defaultdict(list)
        overall_start_time = time.time()
        overall_items_scanned = 0
        overall_videos_processed = 0
        overall_attr_errors = 0
        overall_sha1_skips = 0

        # --- Iterate through each raw scan path provided ---
        for raw_scan_path_entry in self._raw_scan_paths:
            fs_dir_path = self.calculate_fs_path(raw_scan_path_entry, self._raw_mount_point)

            if fs_dir_path is None:
                self.log(self._("error_path_calc_failed", scan=raw_scan_path_entry, mount=self._raw_mount_point, default=f"Error: Could not determine cloud scan path for '{raw_scan_path_entry}'. Skipping this path."))
                continue # Skip this path and proceed to the next one

            self.log(self._("find_scan_path_start", path=fs_dir_path, default=f"Scanning path: '{fs_dir_path}'..."))
            path_start_time = time.time()
            path_count = 0
            path_video_files_checked = 0
            path_errors_getting_attrs = 0
            path_files_skipped_no_sha1 = 0
            path_key_errors_getting_hash = 0

            try:
                walk_iterator = self.fs.walk_path(fs_dir_path)

                for foldername, _, filenames in walk_iterator:
                    foldername_str = str(foldername)

                    for filename_obj in filenames:
                        path_count += 1
                        overall_items_scanned += 1
                        raw_filepath_str = str(filename_obj)

                        if not raw_filepath_str: continue

                        # Path Construction & Normalization (as before)
                        path_for_storage = ""
                        if raw_filepath_str.startswith('/'):
                            path_for_storage = raw_filepath_str
                        elif '/' in raw_filepath_str:
                            path_for_storage = '/' + raw_filepath_str
                        else:
                            path_for_storage = _build_full_path(foldername_str, raw_filepath_str)

                        while '//' in path_for_storage: path_for_storage = path_for_storage.replace('//', '/')
                        if len(path_for_storage) > 1: path_for_storage = path_for_storage.rstrip('/')
                        if path_for_storage and not path_for_storage.startswith('/'):
                             path_for_storage = '/' + path_for_storage
                        if not path_for_storage: continue

                        # Check Video Extension
                        file_extension = os.path.splitext(path_for_storage)[1].lower()
                        if file_extension in VIDEO_EXTENSIONS:
                            path_video_files_checked += 1
                            overall_videos_processed += 1

                            attrs = None
                            mod_time_dt = None
                            file_size = 0
                            file_sha1_standardized = None

                            try:
                                attrs = self.fs.attr(path_for_storage)
                                raw_sha1_value = None
                                try:
                                    file_hashes_dict = attrs.get('fileHashes')
                                    if isinstance(file_hashes_dict, dict):
                                        raw_sha1_value = file_hashes_dict.get('2')
                                        if isinstance(raw_sha1_value, str) and len(raw_sha1_value) >= 40:
                                            file_sha1_standardized = raw_sha1_value.upper() # Standardize case
                                        else:
                                             # Log invalid SHA1 details if needed (as before)
                                             # self.log(...)
                                             file_sha1_standardized = None
                                    else:
                                        file_sha1_standardized = None
                                except KeyError as ke:
                                    if path_key_errors_getting_hash < 5 or path_key_errors_getting_hash % 10 == 0: # Limit logging
                                        self.log(self._("warning_hash_missing", path=path_for_storage, key_error=ke, default=f"Warning (Path: {fs_dir_path}): Hash data missing for '{path_for_storage}'. KeyError: {ke}. Skipping."))
                                    path_key_errors_getting_hash += 1
                                    path_errors_getting_attrs += 1
                                    file_sha1_standardized = None

                                if not file_sha1_standardized:
                                    path_files_skipped_no_sha1 += 1
                                    overall_sha1_skips += 1
                                    continue # Skip this file

                                mod_time_str = attrs.get('modifiedTime')
                                mod_time_dt = _parse_datetime(mod_time_str)
                                if mod_time_dt is None and mod_time_str:
                                    self.log(self._("error_parse_date", path=path_for_storage, error=f"Unparseable string '{mod_time_str}'", default=f"Warning: Could not parse date '{mod_time_str}' for {path_for_storage}"))

                                size_val = attrs.get('size', 0)
                                try: file_size = int(size_val) if size_val is not None else 0
                                except (ValueError, TypeError):
                                    self.log(self._("warning_size_invalid", size=size_val, path=path_for_storage, default=f"Warning: Invalid size value '{size_val}' for {path_for_storage}. Using 0."))
                                    file_size = 0

                                file_info = {
                                    'path': path_for_storage,
                                    'modified': mod_time_dt,
                                    'size': file_size,
                                    'sha1': file_sha1_standardized
                                    # <<< MODIFICATION: Store set ID later when populating tree >>>
                                }
                                # <<< MERGE into the main dictionary >>>
                                all_potential_duplicates[file_sha1_standardized].append(file_info)

                            except FileNotFoundError as fnf_e:
                                err_msg = self._("error_get_attrs", path=path_for_storage, error=fnf_e, default=f"Error getting attributes/hash for '{path_for_storage}': {fnf_e}")
                                self.log(err_msg)
                                path_errors_getting_attrs += 1
                                overall_attr_errors += 1
                            except Exception as e:
                                err_msg = self._("error_get_attrs", path=path_for_storage, error=e, default=f"Error getting attributes/hash for '{path_for_storage}': {e}")
                                self.log(err_msg)
                                self.log(f"Attribute Error Details ({fs_dir_path}): {traceback.format_exc(limit=2)}")
                                path_errors_getting_attrs += 1
                                overall_attr_errors += 1

                        # Log progress periodically per path
                        if path_count % 200 == 0:
                            self.log(self._("status_scan_progress", path=fs_dir_path, count=path_count, video_count=path_video_files_checked, default=f"Path '{fs_dir_path}': Scanned {path_count} items... Found {path_video_files_checked} videos."))

                # --- Path Scan Finished ---
                path_end_time = time.time()
                path_duration = path_end_time - path_start_time
                self.log(self._("status_scan_finished_duration", path=fs_dir_path, duration=path_duration, default=f"Scan for path '{fs_dir_path}' finished in {path_duration:.2f} seconds."))
                self.log(self._("status_scan_summary_items", path=fs_dir_path, count=path_count, video_count=path_video_files_checked, default=f"Path '{fs_dir_path}': Total items encountered: {path_count}. Video files processed: {path_video_files_checked}."))

                # Report errors/skips for this path
                path_warning_parts = []
                if path_errors_getting_attrs > 0: path_warning_parts.append(f"{path_errors_getting_attrs} attribute errors")
                if path_files_skipped_no_sha1 > 0: path_warning_parts.append(f"{path_files_skipped_no_sha1} files skipped (no/invalid SHA1)")
                if path_warning_parts:
                    self.log(self._("status_scan_warnings", path=fs_dir_path, details='; '.join(path_warning_parts), default=f"Path '{fs_dir_path}': WARNING: {'; '.join(path_warning_parts)}."))

            except Exception as walk_e:
                # Catch errors during the fs.walk_path() iteration itself for this path
                err_msg = self._("error_scan_path", path=fs_dir_path, error=walk_e, default=f"Critical error walking cloud path '{fs_dir_path}': {walk_e}")
                self.log(err_msg)
                self.log(f"Walk Error Details ({fs_dir_path}): {traceback.format_exc()}")
                # Log that we are skipping this path due to the error
                self.log(self._("find_error_processing_path", path=fs_dir_path, error=walk_e, default=f"Error processing scan path '{fs_dir_path}': {walk_e}. Skipping this path."))
                continue # Continue to the next raw_scan_path_entry

        # --- All Paths Processed ---
        overall_end_time = time.time()
        overall_duration = overall_end_time - overall_start_time
        self.log(f"Completed scanning all paths in {overall_duration:.2f} seconds.")
        # Log overall summary stats
        self.log(f"Overall Summary: Items Scanned={overall_items_scanned}, Videos Processed={overall_videos_processed}, Attr Errors={overall_attr_errors}, SHA1 Skips={overall_sha1_skips}")

        # --- Filter Aggregated Results for Actual Duplicates ---
        actual_duplicates = {sha1: files for sha1, files in all_potential_duplicates.items() if len(files) > 1}

        # Report final findings
        if actual_duplicates:
             num_sets = len(actual_duplicates)
             num_files = sum(len(files) for files in actual_duplicates.values())
             self.log(self._("find_complete_found", count=num_sets, default=f"Scan complete. Found {num_sets} duplicate sets ({num_files} total duplicate files) across all paths."))
        else:
             no_dups_msg = self._("find_complete_none", default="Scan complete. No duplicate video files found based on SHA1 hash across all paths.")
             # Add skip reason only if skips occurred overall
             if overall_sha1_skips > 0:
                 no_dups_msg += f" (Note: {overall_sha1_skips} video file(s) were skipped overall due to missing/invalid SHA1.)"
             self.log(no_dups_msg)

        return actual_duplicates

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
                # Optional: Small delay to potentially avoid rate limiting (consider removing if not needed)
                # time.sleep(0.05)
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
# --- End of DuplicateFileFinder Class ---


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
        # <<< MODIFICATION: Removed file cache. Treeview is the source of truth >>>
        # self.files_to_delete_cache = []

        # Tkinter variables
        self.widgets = {} # Holds widget references
        self.string_vars = {} # Holds Entry StringVars (for address, account, password, mount, suffix)
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
                print(f"Warning: Application icon file not found at '{icon_path}'")
        except tk.TclError as e:
            icon_err_msg = self._("error_icon_load", path=os.path.basename(icon_path), error=e, default=f"Error loading icon '{os.path.basename(icon_path)}': {e}")
            print(icon_err_msg)
        except Exception as e:
            icon_err_msg = self._("error_icon_load", path=os.path.basename(icon_path), error=f"Unexpected error: {e}", default=f"Unexpected error loading icon '{os.path.basename(icon_path)}': {e}")
            print(icon_err_msg)

        # --- Menu Bar ---
        self.menu_bar = Menu(master)
        master.config(menu=self.menu_bar)
        self.create_menus() # Populate the menu bar

        # --- Build the UI Sections using master.grid ---
        self._build_ui_structure()

        # --- Final Setup ---
        self.load_config() # Load settings on startup
        self.update_ui_language() # Set initial UI text
        self.set_ui_state('initial') # Initial state before connection


    def _build_ui_structure(self):
        """Creates and grids all the UI widgets directly into the main window (master)."""
        master = self.master
        style = ttk.Style()
        try:
            style.configure("Danger.TButton", foreground="red", font=('TkDefaultFont', 10, 'bold'))
        except tk.TclError:
             style.configure("Danger.TButton", foreground="red") # Fallback

        master.columnconfigure(0, weight=1)
        master.rowconfigure(3, weight=1) # Treeview row expands vertically
        master.rowconfigure(5, weight=1) # Log row expands vertically

        # --- 1. Configuration Section ---
        config_frame = ttk.LabelFrame(master, text=self._("config_title"), padding=(10, 5))
        config_frame.grid(row=0, column=0, padx=10, pady=(10, 5), sticky="ew")
        config_frame.columnconfigure(1, weight=1) # Entries/Listbox expand horizontally
        config_frame.columnconfigure(2, weight=0) # Buttons don't expand
        self.widgets["config_frame"] = config_frame

        # Define simple config fields: (internal_key, label_translation_key, grid_row, is_password)
        simple_config_fields = [
            ("address", "address_label", 0, False),
            ("account", "account_label", 1, False),
            ("password", "password_label", 2, True),
            ("mount_point", "mount_point_label", 4, False), # Mount point below scan paths now
        ]

        for key, label_key, row, is_password in simple_config_fields:
            label = ttk.Label(config_frame, text=self._(label_key))
            label.grid(row=row, column=0, padx=(5, 2), pady=3, sticky=tk.W)
            self.widgets[f"label_{key}"] = label

            var = tk.StringVar()
            self.string_vars[key] = var
            entry_args = {"textvariable": var}
            if is_password:
                entry_args["show"] = "*"
            entry = ttk.Entry(config_frame, **entry_args)
            # Make entry span potentially 2 columns if buttons are in col 2
            entry.grid(row=row, column=1, columnspan=2, padx=(2, 5), pady=3, sticky=tk.EW)
            self.entries[key] = entry

        # Scan Paths Listbox and Buttons
        scan_path_row = 3
        # Label for scan paths
        scan_path_label = ttk.Label(config_frame, text=self._("scan_paths_label"))
        scan_path_label.grid(row=scan_path_row, column=0, padx=(5, 2), pady=(10, 2), sticky=tk.NW) # Align top-west
        self.widgets["label_scan_paths"] = scan_path_label

        # Frame to hold listbox and its scrollbar
        listbox_frame = ttk.Frame(config_frame)
        listbox_frame.grid(row=scan_path_row, column=1, padx=(2, 2), pady=(10, 2), sticky=tk.NSEW)
        listbox_frame.rowconfigure(0, weight=1)
        listbox_frame.columnconfigure(0, weight=1)
        config_frame.rowconfigure(scan_path_row, weight=1) # Allow listbox row to expand slightly if needed

        # Listbox for Scan Paths
        scan_path_listbox = tk.Listbox(listbox_frame, height=4, selectmode=tk.EXTENDED, exportselection=False) # Allow multiple selections for removal
        scan_path_listbox.grid(row=0, column=0, sticky=tk.NSEW)
        self.widgets["scan_path_listbox"] = scan_path_listbox

        # Scrollbar for Listbox
        listbox_scrollbar = ttk.Scrollbar(listbox_frame, orient=tk.VERTICAL, command=scan_path_listbox.yview)
        listbox_scrollbar.grid(row=0, column=1, sticky=tk.NS)
        scan_path_listbox.configure(yscrollcommand=listbox_scrollbar.set)

        # Frame for Add/Remove buttons next to listbox
        scan_path_buttons_frame = ttk.Frame(config_frame)
        scan_path_buttons_frame.grid(row=scan_path_row, column=2, padx=(2, 5), pady=(10, 2), sticky=tk.NS)

        # Add Path Button
        add_path_button = ttk.Button(scan_path_buttons_frame, text=self._("add_path_button"), command=self.add_scan_path)
        add_path_button.pack(side=tk.TOP, pady=(0, 5), fill=tk.X)
        self.widgets["add_scan_path_button"] = add_path_button

        # Remove Path Button
        remove_path_button = ttk.Button(scan_path_buttons_frame, text=self._("remove_path_button"), command=self.remove_selected_scan_paths)
        remove_path_button.pack(side=tk.TOP, fill=tk.X)
        self.widgets["remove_scan_path_button"] = remove_path_button

        # --- 2. Action Buttons Frame (Load, Save, Test, Find) ---
        action_button_frame = ttk.Frame(master, padding=(5, 0))
        action_button_frame.grid(row=1, column=0, padx=10, pady=(0, 5), sticky="ew")
        btn_frame_inner = ttk.Frame(action_button_frame)
        btn_frame_inner.pack(side=tk.LEFT) # Keep buttons left-aligned

        action_buttons_info = [
            ("load", "load_config_button", self.load_config, tk.NORMAL),
            ("save", "save_config_button", self.save_config, tk.NORMAL),
            ("test_conn", "test_connection_button", self.start_test_connection_thread, tk.NORMAL),
            ("find", "find_button", self.start_find_duplicates_thread, tk.DISABLED),
        ]
        for idx, (w_key, t_key, cmd, initial_state) in enumerate(action_buttons_info):
             padx_val = (0, 5)
             button = ttk.Button(btn_frame_inner, text=self._(t_key), command=cmd, state=initial_state)
             button.pack(side=tk.LEFT, padx=padx_val, pady=5)
             self.widgets[f"{w_key}_button"] = button

        # --- 3. Deletion Rules Section ---
        rules_frame = ttk.LabelFrame(master, text=self._("rules_title"), padding=(10, 5))
        rules_frame.grid(row=2, column=0, padx=10, pady=(5, 5), sticky="ew")
        rules_frame.columnconfigure(2, weight=1) # Allow suffix entry to expand
        self.widgets["rules_frame"] = rules_frame

        rule_options = [
            ("shortest_path", RULE_KEEP_SHORTEST, 0),
            ("longest_path", RULE_KEEP_LONGEST, 1),
            ("oldest", RULE_KEEP_OLDEST, 2),
            ("newest", RULE_KEEP_NEWEST, 3),
            ("keep_suffix", RULE_KEEP_SUFFIX, 4)
        ]
        suffix_row_index = -1

        for row_idx, (t_key_suffix, value, grid_row) in enumerate(rule_options):
            t_key = f"rule_{t_key_suffix}"
            radio = ttk.Radiobutton(rules_frame, text=self._(t_key),
                                    variable=self.deletion_rule_var, value=value,
                                    command=self._on_rule_change, state=tk.DISABLED)
            radio.grid(row=grid_row, column=0, columnspan=1, padx=5, pady=2, sticky="w")
            self.rule_radios[value] = radio
            self.widgets[f"radio_{value}"] = radio
            if value == RULE_KEEP_SUFFIX:
                suffix_row_index = grid_row

        lbl = ttk.Label(rules_frame, text=self._("rule_suffix_entry_label"), state=tk.DISABLED)
        lbl.grid(row=suffix_row_index, column=1, padx=(15, 2), pady=2, sticky="e")
        self.widgets["suffix_label"] = lbl

        entry = ttk.Entry(rules_frame, textvariable=self.suffix_entry_var, state=tk.DISABLED)
        entry.grid(row=suffix_row_index, column=2, padx=(0, 5), pady=2, sticky="ew")
        self.widgets["suffix_entry"] = entry
        self.entries["suffix"] = entry

        # --- 4. Results TreeView Section ---
        tree_frame = ttk.Frame(master) # Parent is master
        tree_frame.grid(row=3, column=0, padx=10, pady=(5, 5), sticky="nsew") # Fill expanding area
        tree_frame.rowconfigure(0, weight=1) # Treeview expands vertically within frame
        tree_frame.columnconfigure(0, weight=1) # Treeview expands horizontally within frame
        self.widgets["tree_frame"] = tree_frame

        self.columns = ("action", "path", "modified", "size_mb", "set_id")
        # <<< MODIFICATION: Changed selectmode to browse for single click >>>
        self.tree = ttk.Treeview(tree_frame, columns=self.columns, show="headings", selectmode="browse")
        self.widgets["treeview"] = self.tree
        # <<< MODIFICATION: Bind click event handler >>>
        self.tree.bind("<ButtonRelease-1>", self._on_tree_click)

        self.tree.column("action", width=80, anchor=tk.CENTER, stretch=tk.NO)
        self.tree.column("path", width=550, anchor=tk.W, stretch=tk.YES)
        self.tree.column("modified", width=150, anchor=tk.W, stretch=tk.NO)
        self.tree.column("size_mb", width=100, anchor=tk.E, stretch=tk.NO)
        self.tree.column("set_id", width=60, anchor=tk.CENTER, stretch=tk.NO)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')

        self.tree.tag_configure('keep', foreground='darkgreen')
        self.tree.tag_configure('delete', foreground='#CC0000', font=('TkDefaultFont', 9, 'bold'))
        # Headings setup called later

        # --- 5. Final Action Buttons Frame (Delete, Chart, Save Report) ---
        final_action_frame = ttk.Frame(master) # Parent is master
        final_action_frame.grid(row=4, column=0, padx=10, pady=(5, 0), sticky="ew")
        final_btn_inner_frame = ttk.Frame(final_action_frame)
        final_btn_inner_frame.pack(side=tk.LEFT) # Keep buttons left-aligned

        final_buttons_info = [
             # <<< MODIFICATION: Renamed key and translation key, changed command >>>
             ("delete", "delete_selected_button", self.start_delete_selected_thread, tk.DISABLED, "Danger.TButton"),
             ("chart", "show_chart_button", self.show_cloud_file_types, tk.DISABLED, ""),
             ("save_list", "save_list_button", self.save_duplicates_report, tk.DISABLED, ""),
        ]

        for idx, (w_key, t_key, cmd, initial_state, style_name) in enumerate(final_buttons_info):
            padx_val = (0, 10)
            btn_args = {"text": self._(t_key), "command": cmd, "state": initial_state}
            if style_name: btn_args["style"] = style_name
            if w_key == "chart":
                effective_t_key = t_key if MATPLOTLIB_AVAILABLE else "show_chart_button_disabled"
                effective_state = initial_state if MATPLOTLIB_AVAILABLE else tk.DISABLED
                btn_args["text"] = self._(effective_t_key)
                btn_args["state"] = effective_state

            button = ttk.Button(final_btn_inner_frame, **btn_args)
            button.pack(side=tk.LEFT, padx=padx_val, pady=5)
            self.widgets[f"{w_key}_button"] = button

        # --- 6. Log Output Area Section ---
        log_frame = ttk.LabelFrame(master, text=self._("log_title"), padding=(5, 5)) # Parent is master
        log_frame.grid(row=5, column=0, padx=10, pady=(0, 10), sticky="nsew") # Bottom padding
        log_frame.rowconfigure(0, weight=1) # Text area expands vertically within frame
        log_frame.columnconfigure(0, weight=1) # Text area expands horizontally within frame
        self.widgets["log_frame"] = log_frame

        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=10,
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
            if hasattr(self, 'log_message'): self.log_message(f"Warning: Could not save language preference: {e}")
        except Exception as e:
             print(f"Error saving language preference: {e}")
             if hasattr(self, 'log_message'): self.log_message(f"Error saving language preference: {e}")

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
            self.finder.set_translator(self._)
            self.update_ui_language()
            self.save_language_preference()
            self.log_message(f"Language changed to '{lang_code}'.")
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
            self.master.title(self._("window_title"))

            # Menu Bar
            if self.menu_bar and self.menu_bar.winfo_exists():
                try: self.menu_bar.entryconfig(0, label=self._("menu_language"))
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
                "label_password": "password_label",
                "label_scan_paths": "scan_paths_label",
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
                # <<< MODIFICATION: Use new button key >>>
                "delete_button": "delete_selected_button",
                "save_list_button": "save_list_button",
                "add_scan_path_button": "add_path_button",
                "remove_scan_path_button": "remove_path_button",
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
                widget = self.rule_radios.get(value)
                if widget and widget.winfo_exists():
                    try: widget.config(text=self._(text_key))
                    except tk.TclError: pass

            # Treeview Headings
            self.setup_treeview_headings()

            # Re-apply rule highlighting if data exists and a rule was selected
            # (Now also handles manual selections persisting)
            tree = self.widgets.get("treeview")
            if self.duplicate_sets and tree and tree.winfo_exists() and tree.get_children():
                # Need to re-apply tags based on current Action column values
                self._update_treeview_tags_from_action()

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
        now_aware = datetime.now(timezone.utc)
        min_datetime_sort = datetime.min.replace(tzinfo=timezone.utc)
        max_datetime_sort = datetime.max.replace(tzinfo=timezone.utc) - timedelta(seconds=1)

        def get_sortable_datetime(dt_obj):
            if isinstance(dt_obj, datetime):
                if dt_obj.tzinfo is None or dt_obj.tzinfo.utcoffset(dt_obj) is None:
                     return min_datetime_sort if self._sort_ascending else max_datetime_sort
                else:
                    return dt_obj
            else:
                return min_datetime_sort if self._sort_ascending else max_datetime_sort

        for item_id in tree.get_children(''):
            if item_id in self.treeview_item_map:
                file_info = self.treeview_item_map[item_id]
                sort_value = None
                try:
                    if col == 'path':
                        sort_value = file_info.get('path', '').lower()
                    elif col == 'modified':
                        sort_value = get_sortable_datetime(file_info.get('modified'))
                    elif col == 'size_mb':
                        size_str = tree.set(item_id, col)
                        try:
                            sort_value = float(size_str)
                        except ValueError:
                            sort_value = 0.0
                    elif col == 'set_id':
                        current_val_str = tree.set(item_id, col)
                        match = re.search(r'\d+', current_val_str)
                        sort_value = int(match.group(0)) if match else 0
                    elif col == 'action':
                        sort_value = tree.set(item_id, col).lower()
                    else:
                        sort_value = tree.set(item_id, col).lower()
                    items_to_sort.append((sort_value, item_id))
                except Exception as e:
                    print(f"Error getting sort value for item {item_id}, col {col}: {e}")
                    default_sort_val = 0
                    if col == 'modified': default_sort_val = min_datetime_sort if self._sort_ascending else max_datetime_sort
                    elif col in ['path', 'action']: default_sort_val = "" if self._sort_ascending else "~"
                    elif col == 'size_mb': default_sort_val = 0.0
                    items_to_sort.append((default_sort_val, item_id))

        # Perform the Sort
        try:
            items_to_sort.sort(key=lambda x: x[0], reverse=not self._sort_ascending)
        except TypeError as te:
            self.log_message(f"Error: Could not sort column '{col}'. Inconsistent data types found. ({te})")
            print(f"Sorting TypeError for column {col}: {te}")
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
                if tree.exists(item_id):
                    tree.move(item_id, '', i)
            except tk.TclError as e:
                print(f"Error moving tree item {item_id}: {e}")

        self.setup_treeview_headings()


    # --- GUI Logic Methods ---
    def log_message(self, message):
        """ Safely appends a timestamped message to the log ScrolledText widget from any thread. """
        message_str = str(message) if message is not None else ""
        log_widget = self.widgets.get("log_text")
        if hasattr(self, 'master') and self.master and self.master.winfo_exists() and \
           log_widget and log_widget.winfo_exists():
            try:
                self.master.after(0, self._append_log, message_str)
            except (tk.TclError, RuntimeError) as e:
                 print(f"Log Error (Tcl/Runtime): {e} - Message: {message_str}")
            except Exception as e:
                 print(f"Error scheduling log message: {e}\nMessage: {message_str}")
        else:
             timestamp = datetime.now().strftime("%H:%M:%S")
             print(f"[LOG FALLBACK - {timestamp}] {message_str}")

    def _append_log(self, message):
        """ Internal method to append message to log widget (MUST run in main GUI thread). """
        log_widget = self.widgets.get("log_text")
        if not log_widget or not log_widget.winfo_exists():
            return # Widget gone
        try:
            current_state = log_widget.cget('state')
            log_widget.configure(state='normal')
            timestamp = datetime.now().strftime("%H:%M:%S")
            log_widget.insert(tk.END, f"[{timestamp}] {message}\n")
            log_widget.see(tk.END) # Scroll to the end
            log_widget.configure(state=current_state) # Restore previous state
        except tk.TclError as e:
            print(f"Log Append TclError: {e} - Message: {message}")
        except Exception as e:
            print(f"Unexpected error appending log: {e}\nMessage: {message}")
            try:
                if log_widget and log_widget.winfo_exists(): log_widget.configure(state=current_state)
            except: pass

    def load_config(self):
        """ Loads configuration from the ini file into the GUI fields. """
        config_path = CONFIG_FILE
        self.log_message(self._("status_loading_config", file=os.path.basename(config_path), default=f"Loading config from {os.path.basename(config_path)}..."))
        config = configparser.ConfigParser()

        self.string_vars["address"].set(DEFAULT_API_ADDRESS)

        # Clear other fields and listbox before loading
        for key in ["account", "password", "mount_point"]:
            if key in self.string_vars: self.string_vars[key].set("")
        scan_listbox = self.widgets.get("scan_path_listbox")
        if scan_listbox and scan_listbox.winfo_exists():
            try: scan_listbox.delete(0, tk.END)
            except tk.TclError: pass

        try:
            if not os.path.exists(config_path):
                 self.log_message(self._("status_config_not_found", file=os.path.basename(config_path), default=f"Config file '{os.path.basename(config_path)}' not found. Using defaults."))
                 return # Keep the defaults set above

            read_files = config.read(config_path, encoding='utf-8')
            if not read_files:
                 self.log_message(f"Warning: Config file '{os.path.basename(config_path)}' exists but could not be read or is empty.")
                 return

            if 'config' in config:
                cfg_section = config['config']
                # Load simple fields, using default for API Address if key missing
                self.string_vars["address"].set(cfg_section.get("clouddrvie2_address", DEFAULT_API_ADDRESS))
                self.string_vars["account"].set(cfg_section.get("clouddrive2_account", ""))
                self.string_vars["password"].set(cfg_section.get("clouddrive2_passwd", ""))
                self.string_vars["mount_point"].set(cfg_section.get("clouddrive2_root_path", ""))

                # Load scan paths from potentially multi-line string
                root_path_str = cfg_section.get("root_path", "")
                if root_path_str and scan_listbox and scan_listbox.winfo_exists():
                    paths = [p.strip() for p in root_path_str.split('\n') if p.strip()]
                    for path in paths:
                        scan_listbox.insert(tk.END, path)

                self.log_message(self._("status_config_loaded", default="Config loaded successfully."))
            else:
                self.log_message(self._("status_config_section_missing", default="Config file loaded, but '[config]' section is missing."))

        except configparser.Error as e:
            error_msg = self._("error_config_read", error=e, default=f"Error reading config file: {e}")
            if self.master.winfo_exists(): messagebox.showerror(self._("error_config_title", default="Config Error"), error_msg, master=self.master)
            self.log_message(error_msg)
        except Exception as e:
             error_msg = self._("error_unexpected", error=f"loading config: {e}", default=f"Unexpected error loading config: {e}")
             if self.master.winfo_exists(): messagebox.showerror(self._("error_title", default="Error"), error_msg, master=self.master)
             self.log_message(error_msg)
             self.log_message(traceback.format_exc())


    def save_config(self):
        """ Saves current configuration from GUI fields to the ini file. """
        config_path = CONFIG_FILE
        self.log_message(self._("status_saving_config", file=os.path.basename(config_path), default=f"Saving config to {os.path.basename(config_path)}..."))
        config = configparser.ConfigParser()

        # Get scan paths from Listbox
        scan_paths = []
        scan_listbox = self.widgets.get("scan_path_listbox")
        if scan_listbox and scan_listbox.winfo_exists():
            scan_paths = list(scan_listbox.get(0, tk.END))
        # Join paths with newline for storage in INI
        root_path_value = "\n".join(p.strip() for p in scan_paths if p.strip())

        config_data = {
            "clouddrvie2_address": self.string_vars["address"].get(),
            "clouddrive2_account": self.string_vars["account"].get(),
            "clouddrive2_passwd": self.string_vars["password"].get(),
            "root_path": root_path_value,
            "clouddrive2_root_path": self.string_vars["mount_point"].get(),
        }
        config['config'] = config_data

        # Preserve Other Sections (Best effort)
        try:
            if os.path.exists(config_path):
                 config_old = configparser.ConfigParser()
                 with open(config_path, 'r', encoding='utf-8') as f_old:
                     config_old.read_file(f_old)
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
            if self.master.winfo_exists(): messagebox.showerror(self._("error_config_save_title", default="Config Save Error"), error_msg, master=self.master)
            self.log_message(error_msg)
        except Exception as e:
             error_msg = self._("error_unexpected", error=f"saving config: {e}", default=f"Unexpected error saving config: {e}")
             if self.master.winfo_exists(): messagebox.showerror(self._("error_title", default="Error"), error_msg, master=self.master)
             self.log_message(error_msg)
             self.log_message(traceback.format_exc())

    def add_scan_path(self):
        """Opens a directory selection dialog and adds the selected path to the listbox."""
        dialog_title = self._("select_scan_path_dialog_title", default="Select Root Scan Directory")
        selected_directory = filedialog.askdirectory(title=dialog_title, parent=self.master)

        if selected_directory:
            scan_listbox = self.widgets.get("scan_path_listbox")
            if scan_listbox and scan_listbox.winfo_exists():
                # Check if path already exists
                current_paths = list(scan_listbox.get(0, tk.END))
                normalized_new_path = os.path.normpath(selected_directory)
                normalized_existing = [os.path.normpath(p) for p in current_paths]

                if normalized_new_path in normalized_existing:
                    self.log_message(self._("scan_path_already_exists", path=selected_directory, default=f"Path '{selected_directory}' already exists."))
                    if self.master.winfo_exists(): messagebox.showwarning("Path Exists", self._("scan_path_already_exists", path=selected_directory), parent=self.master)
                else:
                    scan_listbox.insert(tk.END, selected_directory)
                    self.log_message(f"Added scan path: {selected_directory}")
                    scan_listbox.see(tk.END) # Scroll to the newly added item

    def remove_selected_scan_paths(self):
        """Removes the selected path(s) from the scan path listbox."""
        scan_listbox = self.widgets.get("scan_path_listbox")
        if scan_listbox and scan_listbox.winfo_exists():
            selected_indices = scan_listbox.curselection()
            if not selected_indices:
                self.log_message("No scan paths selected to remove.")
                return

            # Remove items in reverse order to avoid index issues
            for i in reversed(selected_indices):
                removed_path = scan_listbox.get(i)
                scan_listbox.delete(i)
                self.log_message(f"Removed scan path: {removed_path}")

    def _check_path_chars(self, path_dict, check_scan_paths_from_listbox=False):
        """
        Validates characters in specified path inputs using _validate_path_chars helper.
        Logs details and shows a warning popup if suspicious characters are found.
        Optionally checks paths from the scan path listbox.
        Returns True if all paths are valid, False otherwise.
        """
        suspicious_char_found = False
        all_details = []

        path_display_names = {
            "address": self._("address_label", default="API Address").rstrip(': '),
            # "scan_path" key is now used generically for listbox items
            "scan_path": self._("scan_paths_label", default="Scan Path").rstrip(': '),
            "mount_point": self._("mount_point_label", default="Mount Point").rstrip(': ')
        }

        # Check paths passed in the dictionary (e.g., address, mount_point)
        for key, path_str in path_dict.items():
             if key not in path_display_names: continue # Only check known path fields
             issues = _validate_path_chars(path_str)
             if issues:
                 suspicious_char_found = True
                 display_name = path_display_names.get(key, key)
                 all_details.append(f"'{display_name}' (value: '{path_str}'): {', '.join(issues)}")

        # Optionally check paths from the listbox
        if check_scan_paths_from_listbox:
            scan_listbox = self.widgets.get("scan_path_listbox")
            if scan_listbox and scan_listbox.winfo_exists():
                scan_paths = list(scan_listbox.get(0, tk.END))
                display_name = path_display_names.get("scan_path", "Scan Path") # Use generic name
                for i, path_str in enumerate(scan_paths):
                    issues = _validate_path_chars(path_str)
                    if issues:
                        suspicious_char_found = True
                        all_details.append(f"'{display_name}' #{i+1} (value: '{path_str}'): {', '.join(issues)}")

        if suspicious_char_found:
            log_sep = "!" * 70
            self.log_message(log_sep)
            warning_title = self._("path_warning_title", default="Path Input Warning")
            self.log_message(f"*** {warning_title} ***")
            for detail in all_details: self.log_message(f"  -> {detail}")

            warning_msg_template = self._("path_warning_suspicious_chars", default="Suspicious character(s) detected!\nPlease DELETE and MANUALLY RETYPE paths.")
            warning_lines = warning_msg_template.split('\n')
            popup_msg = f"{warning_lines[0]}\n\n" + "\n".join(warning_lines[1:]) if len(warning_lines) > 1 else warning_lines[0]
            instruction = warning_lines[1] if len(warning_lines) > 1 else "Please check and retype."

            self.log_message(f"Instruction: {instruction}")
            self.log_message(log_sep)

            if self.master.winfo_exists():
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
        has_duplicates = is_connected and bool(self.duplicate_sets)
        # <<< MODIFICATION: Check treeview directly for delete state >>>
        has_files_marked_for_deletion = False
        tree = self.widgets.get("treeview")
        if is_connected and has_duplicates and tree and tree.winfo_exists():
             delete_text = self._("tree_action_delete", default="Delete")
             for item_id in tree.get_children(''):
                 try:
                     if tree.exists(item_id) and tree.set(item_id, "action") == delete_text:
                         has_files_marked_for_deletion = True
                         break
                 except tk.TclError: pass # Ignore if item vanishes

        # Calculate widget states
        config_entry_state = tk.NORMAL if is_idle_state else tk.DISABLED
        config_button_state = tk.NORMAL if is_idle_state else tk.DISABLED
        find_button_state = tk.NORMAL if is_idle_state and is_connected else tk.DISABLED
        # <<< MODIFICATION: Rules are enabled once duplicates found >>>
        rules_radio_state = tk.NORMAL if is_idle_state and is_connected and has_duplicates else tk.DISABLED
        scan_path_listbox_state = tk.NORMAL if is_idle_state else tk.DISABLED
        scan_path_button_state = tk.NORMAL if is_idle_state else tk.DISABLED

        suffix_widgets_state = tk.DISABLED
        if rules_radio_state == tk.NORMAL and self.deletion_rule_var.get() == RULE_KEEP_SUFFIX:
            suffix_widgets_state = tk.NORMAL

        # <<< MODIFICATION: Delete button depends on files marked in tree >>>
        delete_button_state = tk.NORMAL if is_idle_state and is_connected and has_duplicates and has_files_marked_for_deletion else tk.DISABLED
        save_report_button_state = tk.NORMAL if is_idle_state and is_connected and has_duplicates else tk.DISABLED
        chart_button_state = tk.NORMAL if is_idle_state and is_connected and MATPLOTLIB_AVAILABLE else tk.DISABLED

        # Apply states safely
        for key, entry in self.entries.items():
            if key != "suffix" and entry and entry.winfo_exists():
                try: entry.config(state=config_entry_state)
                except tk.TclError: pass

        for btn_key in ["load_button", "save_button", "test_conn_button"]:
             widget = self.widgets.get(btn_key)
             if widget and widget.winfo_exists():
                 try: widget.config(state=config_button_state)
                 except tk.TclError: pass

        # Set state for scan path Listbox and buttons
        widget = self.widgets.get("scan_path_listbox")
        if widget and widget.winfo_exists():
            try: widget.config(state=scan_path_listbox_state)
            except tk.TclError: pass
        widget = self.widgets.get("add_scan_path_button")
        if widget and widget.winfo_exists():
            try: widget.config(state=scan_path_button_state)
            except tk.TclError: pass
        widget = self.widgets.get("remove_scan_path_button")
        if widget and widget.winfo_exists():
            try: widget.config(state=scan_path_button_state)
            except tk.TclError: pass

        widget = self.widgets.get("find_button")
        if widget and widget.winfo_exists():
             try: widget.config(state=find_button_state)
             except tk.TclError: pass

        for radio in self.rule_radios.values():
             if radio and radio.winfo_exists():
                 try: radio.config(state=rules_radio_state)
                 except tk.TclError: pass
        widget = self.widgets.get("suffix_label")
        if widget and widget.winfo_exists():
             try: widget.config(state=suffix_widgets_state)
             except tk.TclError: pass
        widget = self.widgets.get("suffix_entry")
        if widget and widget.winfo_exists():
             try: widget.config(state=suffix_widgets_state)
             except tk.TclError: pass

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
                effective_text_key = "show_chart_button" if MATPLOTLIB_AVAILABLE else "show_chart_button_disabled"
                widget.config(state=chart_button_state, text=self._(effective_text_key))
            except tk.TclError: pass


    def start_test_connection_thread(self):
        """ Validates inputs and starts the connection test in a background thread. """
        address = self.string_vars["address"].get()
        account = self.string_vars["account"].get()
        mount_point = self.string_vars["mount_point"].get()
        scan_listbox = self.widgets.get("scan_path_listbox")
        scan_paths = list(scan_listbox.get(0, tk.END)) if scan_listbox else []

        if not all([address, account, mount_point]):
             error_msg = self._("error_input_missing_conn", default="API Address, Account, and Mount Point are required for connection test.")
             self.log_message(error_msg)
             if self.master.winfo_exists(): messagebox.showerror(self._("error_input_title", default="Input Error"), error_msg, master=self.master)
             return

        # Check mount point, but not scan paths for connection test char validation
        paths_to_check_conn = {
            "address": address,         # Implicitly checked by connection attempt
            "mount_point": mount_point,
            # Don't check scan_paths here as they aren't strictly needed for fs.ls('/')
        }
        if not self._check_path_chars(paths_to_check_conn, check_scan_paths_from_listbox=False):
            return

        self.log_message(self._("status_connecting", default="Attempting connection test..."))
        self.set_ui_state("testing_connection")
        thread = threading.Thread(target=self._test_connection_worker,
                                  args=(address, account, self.string_vars["password"].get(), scan_paths, mount_point),
                                  daemon=True)
        thread.start()

    def _test_connection_worker(self, address, account, passwd, scan_paths, mount_point):
        """ Worker thread for testing the CloudDrive2 connection. """
        connected = False
        try:
            connected = self.finder.set_config(address, account, passwd, scan_paths, mount_point, self.log_message)

            if self.master.winfo_exists():
                 if connected:
                     success_title = self._("conn_test_success_title", default="Connection Test Successful")
                     success_msg = self._("conn_test_success_msg", default="Successfully connected to CloudDrive2.")
                     self.master.after(10, lambda st=success_title, sm=success_msg: messagebox.showinfo(st, sm, master=self.master))
                 else:
                     fail_title = self._("conn_test_fail_title", default="Connection Test Failed")
                     fail_msg = self._("conn_test_fail_msg", default="Failed to connect. Check log for details.")
                     self.master.after(10, lambda ft=fail_title, fm=fail_msg: messagebox.showwarning(ft, fm, master=self.master))

        except Exception as e:
            error_msg = self._("error_unexpected", error=f"during connection test: {e}", default=f"Unexpected error during connection test: {e}")
            self.log_message(error_msg)
            self.log_message(traceback.format_exc())
            if self.master.winfo_exists():
                error_title = self._("error_title", default="Error")
                self.master.after(10, lambda et=error_title, em=error_msg: messagebox.showerror(et, em, master=self.master))
        finally:
            if self.master.winfo_exists():
                final_state = 'normal' if connected else 'initial'
                self.master.after(0, self.set_ui_state, final_state)


    def start_find_duplicates_thread(self):
        """ Handles 'Find Duplicates' click. Validates inputs, clears previous results, and starts worker thread. """
        if not self.finder or not self.finder.fs:
             if self.master.winfo_exists(): messagebox.showwarning(self._("error_title", default="Error"), self._("error_not_connected", default="Not connected."), master=self.master)
             self.log_message(self._("error_not_connected", default="Error: Not connected. Cannot start scan."))
             return

        scan_listbox = self.widgets.get("scan_path_listbox")
        scan_paths_val = list(scan_listbox.get(0, tk.END)) if scan_listbox else []
        mount_point_val = self.string_vars["mount_point"].get()
        address_val = self.string_vars["address"].get()
        account_val = self.string_vars["account"].get()

        missing = []
        if not address_val: missing.append(f"'{self._('address_label').rstrip(': ')}'")
        if not account_val: missing.append(f"'{self._('account_label').rstrip(': ')}'")
        if not mount_point_val: missing.append(f"'{self._('mount_point_label').rstrip(': ')}'")
        if not scan_paths_val: missing.append(f"'{self._('scan_paths_label').rstrip(': ')} (at least one)")

        if missing:
              error_msg_base = self._("error_input_missing", default="Required fields missing")
              error_msg = f"{error_msg_base}: {', '.join(missing)}."
              if self.master.winfo_exists(): messagebox.showerror(self._("error_input_title", default="Input Error"), error_msg, master=self.master)
              self.log_message(error_msg)
              return

        # Validate Path Characters (Mount Point AND Scan Paths from listbox)
        paths_to_check = {"mount_point": mount_point_val}
        if not self._check_path_chars(paths_to_check, check_scan_paths_from_listbox=True):
            return

        self.clear_results()
        self.log_message(self._("find_starting", num_paths=len(scan_paths_val), default=f"Starting duplicate scan ({len(scan_paths_val)} paths)...")) # Log included in worker now
        self.set_ui_state("finding")

        # NOTE: finder already has the paths from set_config, no need to pass them again
        thread = threading.Thread(target=self._find_duplicates_worker, daemon=True)
        thread.start()

    def _find_duplicates_worker(self):
        """ Worker thread for finding duplicates. Calls finder method and schedules GUI update. """
        if not self.finder or not self.finder.fs:
            self.log_message("Error: Connection lost before Find Duplicates scan could execute.")
            if self.master.winfo_exists(): self.master.after(0, self.set_ui_state, 'normal')
            return

        found_duplicates = {}
        try:
            # Call the core logic - finder now uses its stored list of paths
            found_duplicates = self.finder.find_duplicates()

            if self.master.winfo_exists():
                self.master.after(0, self._process_find_results, found_duplicates)

        except Exception as e:
            err_msg = self._("find_error_during", error=e, default=f"Unexpected error during scan process: {e}")
            self.log_message(err_msg)
            self.log_message(traceback.format_exc())
            if self.master.winfo_exists():
                 error_title = self._("error_title", default="Scan Error")
                 self.master.after(10, lambda et=error_title, em=err_msg: messagebox.showerror(et, em, master=self.master))
                 self.master.after(0, self.set_ui_state, 'normal')


    def _process_find_results(self, found_duplicates):
        """ Processes results from find_duplicates worker (runs in main thread). Updates GUI. """
        if not self.master.winfo_exists(): return

        self.duplicate_sets = found_duplicates if found_duplicates else {}

        if self.duplicate_sets:
            self.populate_treeview()
            # <<< MODIFICATION: Don't automatically apply rule. User must click radio or manually select. >>>
            # if self.deletion_rule_var.get():
            #    self._apply_rule_to_treeview()
        else:
            tree = self.widgets.get("treeview")
            if tree and tree.winfo_exists():
                try:
                    if tree.get_children(): tree.delete(*tree.get_children())
                except tk.TclError: pass
            self.treeview_item_map.clear()
            # Finder already logged "no duplicates found"

        self.set_ui_state('normal')


    def clear_results(self):
        """Clears the treeview, stored duplicate data, rule selection, and resets sort."""
        self.log_message(self._("status_clearing_tree", default="Clearing results list and rule selection..."))

        self.duplicate_sets = {}
        self.treeview_item_map = {}
        # <<< MODIFICATION: No cache to clear >>>
        # self.files_to_delete_cache = []
        self.deletion_rule_var.set("")
        self.suffix_entry_var.set("")
        self._last_sort_col = None
        self._sort_ascending = True

        tree = self.widgets.get("treeview")
        if tree and tree.winfo_exists():
            try:
                if tree.get_children(): tree.delete(*tree.get_children())
                self.setup_treeview_headings() # Reset headers
            except tk.TclError: pass

        # <<< MODIFICATION: Set state *after* clearing rule var >>>
        self.set_ui_state('normal')


    def populate_treeview(self):
        """ Populates the treeview with found duplicate sets, leaving 'Action' blank initially. """
        tree = self.widgets.get("treeview")
        if not tree or not tree.winfo_exists():
            self.log_message("Error: Treeview widget not available. Cannot display results.")
            return

        count = len(self.duplicate_sets)
        if count == 0:
             self.log_message("No duplicate sets to display.")
             try:
                 if tree.get_children(): tree.delete(*tree.get_children())
             except tk.TclError: pass
             self.treeview_item_map.clear()
             return

        self.log_message(self._("status_populating_tree", count=count, default=f"Populating list with {count} duplicate sets..."))
        start_time = time.time()

        try:
            if tree.get_children(): tree.delete(*tree.get_children())
        except tk.TclError:
            self.log_message("Error clearing treeview before population.")
            return
        self.treeview_item_map.clear()

        set_index = 0
        items_inserted = 0
        items_failed = 0
        sorted_sha1s = sorted(self.duplicate_sets.keys())

        for sha1 in sorted_sha1s:
            files_in_set = self.duplicate_sets[sha1]
            if not isinstance(files_in_set, list) or len(files_in_set) < 2:
                continue

            set_index += 1
            sorted_files = sorted(files_in_set, key=lambda x: x.get('path', ''))

            for file_info in sorted_files:
                try:
                    path = file_info.get('path')
                    if not path:
                         self.log_message(f"Warning: Skipping file in set {set_index} (SHA1: {sha1[:8]}...) due to missing path.")
                         items_failed += 1
                         continue

                    mod_time = file_info.get('modified')
                    mod_time_str = mod_time.strftime(DATE_FORMAT) if isinstance(mod_time, datetime) else "N/A"
                    size = file_info.get('size')
                    size_mb = size / (1024 * 1024) if isinstance(size, (int, float)) and size > 0 else 0.0
                    set_id_str = self._("tree_set_col_value", index=set_index, default=f"{set_index}")

                    # <<< MODIFICATION: Initial Action is empty, no initial tag >>>
                    values = ("", path, mod_time_str, f"{size_mb:.2f}", set_id_str)
                    item_id = path # Use path as the unique item ID

                    if not tree.exists(item_id):
                         # Store the Set ID within the file_info for easier lookup later
                         file_info['_set_id_display'] = set_id_str
                         tree.insert("", tk.END, iid=item_id, values=values, tags=())
                         self.treeview_item_map[item_id] = file_info
                         items_inserted += 1
                    else:
                         self.log_message(f"Warning: Item with path '{path}' already exists in tree. Skipping duplicate insertion.")
                         items_failed += 1

                except tk.TclError as e:
                     # Check if item_id was assigned before error
                     item_id_str = item_id if 'item_id' in locals() and item_id else path if path else 'Unknown'
                     self.log_message(f"Error inserting item with path '{item_id_str}' into tree: {e}")
                     items_failed += 1
                     if 'item_id' in locals() and item_id in self.treeview_item_map:
                         try: del self.treeview_item_map[item_id]
                         except KeyError: pass
                except Exception as e:
                     path_str = file_info.get('path', 'Unknown')
                     self.log_message(f"Unexpected error processing file '{path_str}' for treeview: {e}")
                     self.log_message(traceback.format_exc(limit=2))
                     items_failed += 1

        end_time = time.time()
        duration = end_time - start_time
        log_summary = self._("status_tree_populated", default="Results list populated.")
        log_summary += f" ({items_inserted} items displayed"
        if items_failed > 0: log_summary += f", {items_failed} skipped due to errors"
        log_summary += f" in {duration:.2f}s)"
        self.log_message(log_summary)

        self._last_sort_col = None
        self._sort_ascending = True
        self.setup_treeview_headings()


    def _on_rule_change(self):
        """Called when a deletion rule radio button is selected. Applies the rule as a suggestion."""
        selected_rule = self.deletion_rule_var.get()

        # Update suffix entry state
        is_suffix_rule = (selected_rule == RULE_KEEP_SUFFIX)
        can_enable_suffix = (is_suffix_rule and
                             self.finder is not None and self.finder.fs is not None and
                             bool(self.duplicate_sets))
        suffix_widgets_state = tk.NORMAL if can_enable_suffix else tk.DISABLED

        suffix_label = self.widgets.get("suffix_label")
        if suffix_label and suffix_label.winfo_exists():
            try: suffix_label.config(state=suffix_widgets_state)
            except tk.TclError: pass
        suffix_entry = self.widgets.get("suffix_entry")
        if suffix_entry and suffix_entry.winfo_exists():
            try: suffix_entry.config(state=suffix_widgets_state)
            except tk.TclError: pass

        if not is_suffix_rule:
            self.suffix_entry_var.set("")

        # Apply the selected rule to the treeview as a suggestion
        self._apply_rule_to_treeview()


    def _apply_rule_to_treeview(self, log_update=True):
        """
        Updates the 'Action' column and highlighting in the treeview based on the
        selected deletion rule suggestion.
        """
        tree = self.widgets.get("treeview")
        if not self.duplicate_sets or not tree or not tree.winfo_exists():
            # <<< MODIFICATION: No cache to clear >>>
            # self.files_to_delete_cache = []
            self.set_ui_state('normal')
            return

        selected_rule = self.deletion_rule_var.get()

        # If no rule is selected (e.g., after clearing), do nothing to the tree
        if not selected_rule:
            # Don't clear the tree view actions here, only clear them in clear_results()
            # self.files_to_delete_cache = []
            # ... clear tree code removed ...
            self.set_ui_state('normal') # Just update button states
            return

        rule_name_display_key = f"rule_{selected_rule}"
        rule_name_display = self._(rule_name_display_key, default=selected_rule.replace('_', ' ').title())

        if log_update:
            self.log_message(self._("status_applying_rule", rule_name=rule_name_display, count=len(self.duplicate_sets), default=f"Applying suggestion rule '{rule_name_display}'..."))
        start_time = time.time()

        keep_text = self._("tree_action_keep", default="Keep")
        delete_text = self._("tree_action_delete", default="Delete")
        suffix_to_keep = self.suffix_entry_var.get() if selected_rule == RULE_KEEP_SUFFIX else None
        delete_count = 0
        application_error = False
        files_to_delete_list = [] # Store paths suggested for deletion by the rule

        try:
            # Determine the *suggestions* based on the rule
            files_to_delete_list = self._determine_files_to_delete(self.duplicate_sets, selected_rule, suffix_to_keep)
            files_to_delete_paths_set = set(files_to_delete_list)
            delete_count = len(files_to_delete_paths_set)

            tree_update_start = time.time()
            if not tree.winfo_exists():
                raise tk.TclError("Treeview destroyed during rule application")

            # Iterate through all items in the treeview and update Action/Tags
            # Using list() to avoid issues if map changes during iteration (shouldn't here)
            for item_id in list(self.treeview_item_map.keys()):
                 if tree.exists(item_id) and item_id in self.treeview_item_map:
                     file_info = self.treeview_item_map[item_id]
                     path = file_info.get('path')
                     if not path: continue

                     is_marked_for_delete = (path in files_to_delete_paths_set)
                     action_text = delete_text if is_marked_for_delete else keep_text
                     item_tags = ('delete',) if is_marked_for_delete else ('keep',)

                     tree.set(item_id, "action", action_text)
                     tree.item(item_id, tags=item_tags)
                 # else: Item might have been removed already, ignore.

            tree_update_end = time.time()
            if log_update:
                self.log_message(self._("status_rule_applied", delete_count=delete_count, default=f"Rule suggestion applied. {delete_count} files initially marked for deletion. Click 'Action' column to change."))

        except ValueError as ve: # Catch specific error from _determine_files_to_delete
             self.log_message(f"Rule Suggestion Error: {ve}")
             if self.master.winfo_exists(): messagebox.showerror(self._("error_rule_title", default="Rule Error"), str(ve), master=self.master)
             application_error = True
             delete_count = 0
             # Don't clear the tree here, just log the error
        except tk.TclError as e:
             self.log_message(f"Error updating treeview during rule application: {e}")
             application_error = True
             delete_count = 0
        except Exception as e:
             self.log_message(f"Unexpected error applying rule suggestion '{selected_rule}' to treeview: {e}")
             self.log_message(traceback.format_exc())
             application_error = True
             delete_count = 0
        finally:
            # <<< MODIFICATION: No cache to update >>>
            # self.files_to_delete_cache = files_to_delete_list if not application_error else []
            if selected_rule and log_update:
                end_time = time.time()
            # Update UI state which might enable/disable delete button based on tree content
            self.set_ui_state('normal')

    def _get_set_id_for_item(self, item_id):
        """ Helper to safely get the set ID string from a treeview item. """
        tree = self.widgets.get("treeview")
        if tree and tree.winfo_exists() and tree.exists(item_id):
            try:
                return tree.set(item_id, "set_id")
            except tk.TclError:
                return None
        return None

    def _update_treeview_tags_from_action(self):
        """ Updates treeview item tags based on the current value in the 'Action' column. """
        tree = self.widgets.get("treeview")
        if not tree or not tree.winfo_exists():
            return

        keep_text = self._("tree_action_keep", default="Keep")
        delete_text = self._("tree_action_delete", default="Delete")

        try:
            for item_id in tree.get_children(''):
                if not tree.exists(item_id): continue
                action_text = tree.set(item_id, "action")
                if action_text == keep_text:
                    tree.item(item_id, tags=('keep',))
                elif action_text == delete_text:
                    tree.item(item_id, tags=('delete',))
                else:
                    tree.item(item_id, tags=())
        except tk.TclError as e:
            self.log_message(f"Error updating tree tags: {e}")
        except Exception as e:
            self.log_message(f"Unexpected error updating tree tags: {e}")


    def _on_tree_click(self, event):
        """ Handles clicks on the treeview, specifically toggling Keep/Delete in the Action column. """
        tree = self.widgets.get("treeview")
        if not tree or not tree.winfo_exists(): return

        region = tree.identify_region(event.x, event.y)
        if region != "cell":
            return # Click wasn't on a cell

        col_id = tree.identify_column(event.x)
        item_id = tree.identify_row(event.y) # item_id is the file path

        # We only care about clicks on the 'action' column (#1)
        if col_id != "#1" or not item_id:
            return

        try:
            if not tree.exists(item_id): return # Item might have been deleted

            current_action = tree.set(item_id, "action")
            keep_text = self._("tree_action_keep", default="Keep")
            delete_text = self._("tree_action_delete", default="Delete")
            clicked_set_id_str = self._get_set_id_for_item(item_id)

            if not clicked_set_id_str:
                 self.log_message(f"Warning: Could not determine Set ID for clicked item '{item_id}'. Cannot toggle action.")
                 return

            # --- Find siblings in the same set ---
            siblings = []
            current_keep_item = None
            keep_count_in_set = 0
            for sibling_id in tree.get_children(''):
                 if tree.exists(sibling_id):
                     set_id_val = self._get_set_id_for_item(sibling_id)
                     if set_id_val == clicked_set_id_str:
                         siblings.append(sibling_id)
                         action = tree.set(sibling_id, "action")
                         if action == keep_text:
                             keep_count_in_set += 1
                             if sibling_id != item_id: # Don't count self if it's currently Keep
                                 current_keep_item = sibling_id

            # --- Logic: Toggle Keep/Delete, ensuring one Keep per set ---
            new_action = ""
            new_tag = ()

            if current_action == keep_text:
                # Trying to change Keep -> Delete
                if keep_count_in_set <= 1:
                    # Prevent deleting the last 'Keep' item in the set
                    filename = os.path.basename(item_id)
                    set_num_match = re.search(r'\d+', clicked_set_id_str)
                    set_num = set_num_match.group(0) if set_num_match else clicked_set_id_str
                    log_msg = self._("info_last_keep_in_set", filename=filename, set_id=set_num, default=f"Info: Cannot mark '{filename}' for deletion as it's the only file marked 'Keep' in Set {set_num}.")
                    self.log_message(log_msg)
                    return # Do nothing
                else:
                    # Allow change: Keep -> Delete
                    new_action = delete_text
                    new_tag = ('delete',)
            elif current_action == delete_text:
                # Trying to change Delete -> Keep
                new_action = keep_text
                new_tag = ('keep',)
                # Also change the *other* Keep item in this set to Delete
                if current_keep_item and tree.exists(current_keep_item):
                    tree.set(current_keep_item, "action", delete_text)
                    tree.item(current_keep_item, tags=('delete',))
            else: # If current action is blank or something else, treat as changing to Keep
                new_action = keep_text
                new_tag = ('keep',)
                # Also change the *other* Keep item in this set to Delete (if one exists)
                if current_keep_item and tree.exists(current_keep_item):
                    tree.set(current_keep_item, "action", delete_text)
                    tree.item(current_keep_item, tags=('delete',))


            # Apply the change to the clicked item
            if new_action:
                tree.set(item_id, "action", new_action)
                tree.item(item_id, tags=new_tag)

            # Update the UI state (e.g., enable/disable Delete button)
            self.set_ui_state('normal')

        except tk.TclError as e:
            self.log_message(f"Error handling tree click for item '{item_id}': {e}")
        except Exception as e:
            self.log_message(f"Unexpected error handling tree click: {e}")
            self.log_message(traceback.format_exc(limit=2))


    def _determine_files_to_delete(self, duplicate_sets, rule, suffix_value):
        """
        Determines which files to delete based on the selected rule *suggestion* and duplicate sets.
        Handles tie-breaking using shortest path as default. Logs warnings via self.log_message.
        Returns: list: A list of full file paths (str) suggested for deletion.
        Raises: ValueError: If rule is invalid or suffix is missing when required.
        """
        if not isinstance(duplicate_sets, dict) or not duplicate_sets: return []
        if not rule: raise ValueError(self._("delete_no_rule_selected", default="No deletion rule selected."))
        if rule == RULE_KEEP_SUFFIX and not suffix_value: raise ValueError(self._("delete_suffix_missing", default="Suffix is required for the 'Keep Suffix' rule suggestion."))
        valid_rules = {RULE_KEEP_SHORTEST, RULE_KEEP_LONGEST, RULE_KEEP_OLDEST, RULE_KEEP_NEWEST, RULE_KEEP_SUFFIX}
        if rule not in valid_rules: raise ValueError(f"Internal Error: Unknown deletion rule '{rule}'.")

        files_to_delete = []
        log_func = self.log_message
        tie_break_prefix = self._("tie_break_log_prefix", default="Tie-Break:")

        def tie_break_shortest_path(candidates, reason_for_tiebreak):
             if not candidates: return None
             if len(candidates) == 1: return candidates[0]
             # Sort primarily by path length, secondarily by path string itself for stability
             sorted_candidates = sorted(candidates, key=lambda f: (len(f.get('path', '')), f.get('path', '')))
             winner = sorted_candidates[0]
             log_func(self._("warning_tie_break", prefix=tie_break_prefix, reason=reason_for_tiebreak, filename=os.path.basename(winner.get('path','N/A')), detail=f"Shortest path of {len(candidates)}", default=f"{tie_break_prefix} {reason_for_tiebreak}. Kept '{os.path.basename(winner.get('path','N/A'))}' (Shortest path of {len(candidates)})."))
             return winner

        # Determine the display index for each set ID string for logging
        set_id_map = {}
        current_index = 0
        sorted_sha1s_for_index = sorted(duplicate_sets.keys())
        for sha1 in sorted_sha1s_for_index:
             if len(duplicate_sets[sha1]) > 1:
                 current_index += 1
                 set_id_str = self._("tree_set_col_value", index=current_index, default=f"{current_index}")
                 set_id_map[sha1] = set_id_str


        for sha1, files_in_set in duplicate_sets.items():
            if not isinstance(files_in_set, list) or len(files_in_set) < 2: continue
            keep_file_info = None
            # Use the pre-calculated display index for logging consistency
            set_id_for_log = set_id_map.get(sha1, f"SHA1: {sha1[:8]}...")

            try:
                candidates = []
                reason_for_tiebreak = ""

                # --- Apply Rule Logic to find candidate(s) to keep ---
                if rule == RULE_KEEP_SHORTEST:
                    valid_files = [f for f in files_in_set if f.get('path') is not None]
                    if not valid_files: continue
                    min_len = min(len(f.get('path', '')) for f in valid_files)
                    candidates = [f for f in valid_files if len(f.get('path', '')) == min_len]
                    reason_for_tiebreak = f"Multiple files have min path length ({min_len})"
                elif rule == RULE_KEEP_LONGEST:
                    valid_files = [f for f in files_in_set if f.get('path') is not None]
                    if not valid_files: continue
                    max_len = max(len(f.get('path', '')) for f in valid_files)
                    candidates = [f for f in valid_files if len(f.get('path', '')) == max_len]
                    reason_for_tiebreak = f"Multiple files have max path length ({max_len})"
                elif rule == RULE_KEEP_OLDEST:
                    valid_files = [f for f in files_in_set if isinstance(f.get('modified'), datetime)]
                    if not valid_files:
                        log_func(self._("warning_rule_no_date", set_id=set_id_for_log, rule=rule, default=f"Warning: {set_id_for_log} - Cannot apply suggestion '{rule}': No valid dates. Defaulting to shortest path."))
                        candidates = list(files_in_set) # Fallback to all files for tie-break
                        reason_for_tiebreak = "No valid dates found"
                    else:
                        min_date = min(f['modified'] for f in valid_files)
                        candidates = [f for f in valid_files if f['modified'] == min_date]
                        reason_for_tiebreak = f"Multiple files have oldest date ({min_date.strftime(DATE_FORMAT)})"
                elif rule == RULE_KEEP_NEWEST:
                    valid_files = [f for f in files_in_set if isinstance(f.get('modified'), datetime)]
                    if not valid_files:
                        log_func(self._("warning_rule_no_date", set_id=set_id_for_log, rule=rule, default=f"Warning: {set_id_for_log} - Cannot apply suggestion '{rule}': No valid dates. Defaulting to shortest path."))
                        candidates = list(files_in_set) # Fallback to all files for tie-break
                        reason_for_tiebreak = "No valid dates found"
                    else:
                        max_date = max(f['modified'] for f in valid_files)
                        candidates = [f for f in valid_files if f['modified'] == max_date]
                        reason_for_tiebreak = f"Multiple files have newest date ({max_date.strftime(DATE_FORMAT)})"
                elif rule == RULE_KEEP_SUFFIX:
                    suffix_lower = suffix_value.lower()
                    candidates = [f for f in files_in_set if f.get('path', '').lower().endswith(suffix_lower)]
                    if not candidates:
                         log_func(self._("warning_rule_no_suffix_match", set_id=set_id_for_log, suffix=suffix_value, default=f"Warning: {set_id_for_log} - No files match suffix '{suffix_value}'. Defaulting to shortest path."))
                         candidates = list(files_in_set) # Fallback to all files for tie-break
                         reason_for_tiebreak = f"No files match suffix '{suffix_value}'"
                    else:
                         reason_for_tiebreak = f"Multiple files match suffix '{suffix_value}'"

                # --- Tie-breaking or selecting the single candidate ---
                if len(candidates) > 1 or (not candidates and rule in [RULE_KEEP_OLDEST, RULE_KEEP_NEWEST, RULE_KEEP_SUFFIX]):
                    # If rule failed to find *any* candidate (e.g., no dates, no suffix match),
                    # candidates might be empty, so use original files_in_set for tie-break.
                    effective_candidates = candidates if candidates else list(files_in_set)
                    if not effective_candidates: continue # Skip if set was somehow empty
                    full_reason = f"{set_id_for_log} - {reason_for_tiebreak}"
                    keep_file_info = tie_break_shortest_path(effective_candidates, full_reason)
                elif len(candidates) == 1:
                     keep_file_info = candidates[0]
                else: # Should only happen if valid_files was empty initially (e.g., no paths)
                     keep_file_info = None

                # --- Add files *not* kept to the delete list ---
                if keep_file_info and keep_file_info.get('path'):
                    keep_path = keep_file_info['path']
                    for f_info in files_in_set:
                        path = f_info.get('path')
                        if path and path != keep_path:
                            files_to_delete.append(path)
                else:
                     # This case should be rare now due to fallbacks, but log if it happens
                     log_func(self._("warning_rule_failed_selection", set_id=set_id_for_log, rule=rule, default=f"Internal Warning: {set_id_for_log} - Rule '{rule}' failed to select file to keep. Skipping suggestion for this set."))

            except Exception as e:
                 log_func(self._("error_rule_application", set_id=set_id_for_log, rule=rule, error=e, default=f"Error applying suggestion rule '{rule}' to {set_id_for_log}: {e}. Skipping suggestion for this set."))
                 log_func(traceback.format_exc(limit=2))

        return files_to_delete


    # <<< MODIFICATION: Renamed method >>>
    def start_delete_selected_thread(self):
        """ Handles 'Delete Marked Files' click. Validates, confirms, starts worker thread. """
        tree = self.widgets.get("treeview")
        if not tree or not tree.winfo_exists():
            self.log_message("Error: Cannot delete, results list is not available.")
            return

        # <<< MODIFICATION: Collect files to delete directly from treeview >>>
        files_to_delete_list = []
        delete_text = self._("tree_action_delete", default="Delete")
        try:
            for item_id in tree.get_children(''):
                if tree.exists(item_id) and tree.set(item_id, "action") == delete_text:
                    # item_id is the file path
                    files_to_delete_list.append(item_id)
        except tk.TclError as e:
            self.log_message(f"Error reading items to delete from list: {e}")
            return
        except Exception as e:
             self.log_message(f"Unexpected error collecting items for deletion: {e}")
             return

        if not files_to_delete_list:
             if self.master.winfo_exists():
                 messagebox.showinfo(
                     self._("delete_selected_button", default="Delete Marked Files"),
                     self._("delete_no_files_marked", default="No files are currently marked for deletion in the list."),
                     master=self.master)
             self.log_message(self._("delete_no_files_marked", default="No files marked for deletion."))
             return

        num_files_to_delete = len(files_to_delete_list)
        # <<< MODIFICATION: Updated confirmation message >>>
        confirm_msg_template = self._("delete_confirm_msg", default="Permanently delete all files marked 'Delete' in the list?\nTHIS ACTION CANNOT BE UNDONE.")
        confirm_msg = f"{confirm_msg_template}\n\n({num_files_to_delete} files will be permanently deleted)"

        confirm = False
        if self.master.winfo_exists():
            confirm = messagebox.askyesno(
                title=self._("delete_confirm_title", default="Confirm Deletion"),
                message=confirm_msg,
                icon='warning',
                default='no',
                master=self.master)

        if not confirm:
            self.log_message(self._("delete_cancelled", default="Deletion cancelled."))
            return

        # <<< MODIFICATION: Use updated log message key >>>
        self.log_message(self._("delete_starting_selected", default="Starting deletion of manually marked files..."))
        self.set_ui_state("deleting")

        # Pass the collected list (already a copy)
        thread = threading.Thread(target=self._delete_worker, args=(files_to_delete_list,), daemon=True)
        thread.start()

    # <<< MODIFICATION: Removed rule_name arg, no longer needed >>>
    def _delete_worker(self, files_to_delete):
        """ Worker thread for deleting files based on the provided list. """
        if not self.finder or not self.finder.fs:
            self.log_message(self._("error_not_connected", default="Error: Connection lost before Deletion."))
            if self.master.winfo_exists(): self.master.after(0, self.set_ui_state, 'normal')
            return

        deleted_count = 0
        total_attempted = len(files_to_delete)
        deletion_error_occurred = False
        should_clear_results = False

        try:
            if not files_to_delete:
                # <<< MODIFICATION: Simplified log message >>>
                self.log_message(self._("delete_no_files_marked", default="No files to delete.") + " (Worker check)")
            else:
                deleted_count, total_attempted = self.finder.delete_files(files_to_delete)
                if deleted_count < total_attempted: deletion_error_occurred = True
            # Clear results only if deletion was attempted and potentially successful
            if total_attempted > 0: should_clear_results = True

        except Exception as e:
            deletion_error_occurred = True
            err_msg = self._("delete_error_during", error=e, default=f"Unexpected error during deletion process: {e}")
            self.log_message(err_msg)
            self.log_message(traceback.format_exc())
            if self.master.winfo_exists():
                error_title = self._("error_title", default="Deletion Error")
                self.master.after(10, lambda et=error_title, em=err_msg: messagebox.showerror(et, em, master=self.master))
            # Don't clear results if a major error occurred during the process
            should_clear_results = False
        finally:
            if self.master.winfo_exists():
                if should_clear_results:
                     # Use after() to ensure GUI updates happen in the main thread
                    self.master.after(10, self.clear_results) # Clear data and UI elements
                    self.master.after(20, lambda: self.log_message(self._("delete_results_cleared", default="Deletion finished. Results cleared."))) # Log after clearing
                    # Set UI state back to normal *after* clearing is scheduled
                    self.master.after(30, self.set_ui_state, 'normal')
                else:
                    # If deletion wasn't attempted or failed badly, just reset UI state
                    self.master.after(0, self.set_ui_state, 'normal')


    # --- Methods related to Report Saving and Charting ---
    def save_duplicates_report(self):
        """ Saves the report of FOUND duplicate file sets to a text file. """
        if not self.duplicate_sets:
            if self.master.winfo_exists(): messagebox.showinfo(self._("save_list_button", default="Save Report"), self._("save_report_no_data", default="No duplicates to save."), master=self.master)
            self.log_message(self._("save_report_no_data", default="No duplicate sets available."))
            return

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        initial_filename = f"duplicates_report_{timestamp}.txt"
        file_path = None
        try:
            file_path = filedialog.asksaveasfilename(
                defaultextension=".txt", filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
                title=self._("save_list_button", default="Save Duplicates Report As..."),
                initialfile=initial_filename, parent=self.master
            )
        except Exception as fd_e:
            self.log_message(f"Error opening save file dialog: {fd_e}")
            if self.master.winfo_exists(): messagebox.showerror(self._("error_title", default="Error"), f"Could not open save dialog: {fd_e}", master=self.master)
            return

        if file_path:
            success = self.finder.write_duplicates_report(self.duplicate_sets, file_path)
            if success and self.master.winfo_exists():
                messagebox.showinfo(
                    self._("save_list_button", default="Report Saved"),
                    self._("save_report_saved", file=os.path.basename(file_path), default="Report saved."),
                    master=self.master)
        else:
            self.log_message("Save report operation cancelled.")

    def show_cloud_file_types(self):
        """ Handles 'Show Cloud File Types' click. Validates prerequisites and starts worker thread. """
        if not MATPLOTLIB_AVAILABLE:
            if self.master.winfo_exists(): messagebox.showwarning(self._("chart_error_title", default="Chart Error"), self._("chart_error_no_matplotlib", default="Matplotlib not found."), master=self.master)
            self.log_message(self._("chart_error_no_matplotlib", default="Matplotlib not found."))
            return
        if not self.finder or not self.finder.fs:
            if self.master.winfo_exists(): messagebox.showwarning(self._("chart_error_title", default="Chart Error"), self._("chart_error_no_connection", default="Not connected."), master=self.master)
            self.log_message(self._("chart_error_no_connection", default="Not connected, cannot chart."))
            return

        scan_listbox = self.widgets.get("scan_path_listbox")
        scan_paths_raw = list(scan_listbox.get(0, tk.END)) if scan_listbox else []
        mount_point_raw = self.string_vars["mount_point"].get()

        if not mount_point_raw or not scan_paths_raw:
            error_msg = self._("error_input_missing_chart", default="Mount Point and at least one Scan Path required for chart.")
            if self.master.winfo_exists(): messagebox.showwarning(self._("error_input_title", default="Input Error"), error_msg, master=self.master)
            self.log_message(error_msg)
            return

        # Validate Path Characters (Mount Point AND Scan Paths from listbox)
        paths_to_check_chart = {"mount_point": mount_point_raw}
        if not self._check_path_chars(paths_to_check_chart, check_scan_paths_from_listbox=True):
            return

        self.log_message("Starting scan for file type chart data...")
        self.set_ui_state("charting")

        # Pass the list of raw paths to the worker
        thread = threading.Thread(target=self._show_cloud_file_types_worker,
                                    args=(scan_paths_raw, mount_point_raw), daemon=True)
        thread.start()

    def _show_cloud_file_types_worker(self, scan_paths_raw, mount_point_raw):
        """ Worker thread to scan multiple cloud paths, count types, and schedule chart creation. """
        all_file_counts = collections.Counter() # Aggregate counts here
        total_files_overall = 0
        any_scan_error = False
        processed_path_count = 0

        try:
            if not self.finder or not self.finder.fs:
                self.log_message(self._("chart_error_no_connection", default="Cannot chart: Connection lost."))
                if self.master.winfo_exists(): self.master.after(0, self.set_ui_state, 'normal')
                return

            self.log_message(self._("chart_status_scan_paths_start", num_paths=len(scan_paths_raw), default=f"Starting chart scan across {len(scan_paths_raw)} path(s)..."))

            # --- Iterate through each raw scan path ---
            for raw_scan_path_entry in scan_paths_raw:
                fs_dir_path = self.finder.calculate_fs_path(raw_scan_path_entry, mount_point_raw)
                if fs_dir_path is None:
                    self.log_message(f"Chart Scan: Skipping invalid path entry '{raw_scan_path_entry}'.")
                    any_scan_error = True # Treat path calculation error as a scan error for this path
                    continue

                self.log_message(self._("chart_status_scanning_cloud", path=fs_dir_path, default=f"Scanning '{fs_dir_path}' for file types..."))
                scan_start_time = time.time()
                path_files = 0
                path_error = None

                try:
                    for dirpath, _, filenames in self.finder.fs.walk_path(fs_dir_path):
                        for filename_obj in filenames:
                            try:
                                # Use helper to construct path safely
                                full_path = _build_full_path(str(dirpath), str(filename_obj))
                                if not full_path: continue # Skip if path construction fails

                                filename = os.path.basename(full_path)
                                total_files_overall += 1
                                path_files += 1
                                _root, ext = os.path.splitext(filename)
                                ext_label = ext.lower() if ext else self._("chart_label_no_extension", default="[No Ext]")
                                all_file_counts.update([ext_label]) # Use update for Counter
                            except Exception as inner_e:
                                self.log_message(f"Warning: Error processing filename '{filename_obj}' in '{dirpath}' (Chart Scan): {inner_e}")

                except Exception as e:
                    path_error = e
                    any_scan_error = True
                    error_msg = self._("chart_error_cloud_scan", path=fs_dir_path, error=e, default=f"Error scanning '{fs_dir_path}' for chart: {e}")
                    self.log_message(error_msg)
                    self.log_message(f"Chart Scan Error Details ({fs_dir_path}): {traceback.format_exc()}")
                    # Don't show popup here, maybe summarize errors later if desired

                scan_duration = time.time() - scan_start_time
                if not path_error:
                    self.log_message(self._("chart_status_scan_path_complete", path=fs_dir_path, default=f"Finished scanning path '{fs_dir_path}' ({path_files} files found, {scan_duration:.2f}s)."))
                    processed_path_count += 1
                else:
                     self.log_message(f"Finished scanning path '{fs_dir_path}' with errors.")

            # --- Scan finished for all paths ---
            def update_gui_after_chart_scan():
                if not self.master.winfo_exists(): return
                if not all_file_counts:
                    no_files_msg = self._("chart_status_no_files_found", default="No files found in any scanned path. Cannot generate chart.")
                    self.log_message(no_files_msg)
                    if self.master.winfo_exists(): messagebox.showinfo(self._("chart_info_title", default="Chart Info"), no_files_msg, master=self.master)
                    return

                # Generate chart using aggregated counts
                self.log_message(self._("chart_status_generating", count=len(all_file_counts), total=total_files_overall, default=f"Generating chart for {len(all_file_counts)} types ({total_files_overall} files)..."))
                # Use a generic title or list paths if few enough
                chart_display_title = self._("chart_window_title", default="File Types in Scanned Paths") # Generic title now
                self._create_pie_chart_window(all_file_counts, chart_display_title)

            if self.master.winfo_exists():
                self.master.after(0, update_gui_after_chart_scan)

        except Exception as e:
            err_msg = f"Unexpected error during chart worker: {e}"
            self.log_message(err_msg)
            self.log_message(traceback.format_exc())
            if self.master.winfo_exists():
                error_title = self._("chart_error_title", default="Chart Error")
                self.master.after(10, lambda et=error_title, em=err_msg: messagebox.showerror(et, em, master=self.master))
        finally:
            if self.master.winfo_exists():
                self.master.after(0, self.set_ui_state, 'normal')

    def _create_pie_chart_window(self, counts, display_path_or_title):
        """ Creates and displays the file type pie chart in a new Toplevel window. """
        if not MATPLOTLIB_AVAILABLE:
            self.log_message("Error: Matplotlib unavailable.")
            if self.master.winfo_exists(): messagebox.showerror(self._("chart_error_title", default="Chart Error"), self._("chart_error_no_matplotlib", default="Matplotlib not found."), master=self.master)
            return

        chart_window = None
        try:
            try:
                matplotlib.rcParams['axes.unicode_minus'] = False
                current_sans_serif = matplotlib.rcParams['font.sans-serif']
                preferred_fonts = ['SimHei', 'Microsoft YaHei', 'MS Gothic', 'Malgun Gothic', 'Arial Unicode MS', 'sans-serif']
                final_font_list = preferred_fonts + [f for f in current_sans_serif if f not in preferred_fonts]
                matplotlib.rcParams['font.sans-serif'] = final_font_list
            except Exception as mpl_set_err:
                self.log_message(f"Warning: Issue setting Matplotlib rcParams: {mpl_set_err}")

            top_n = 20
            total_count = sum(counts.values())
            sorted_counts = counts.most_common()

            chart_labels, chart_sizes, legend_labels_with_counts = [], [], []
            others_label = self._("chart_label_others", default="Others")
            others_count = 0
            others_sources = []

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
                chart_labels = [item[0] for item in sorted_counts]
                chart_sizes = [item[1] for item in sorted_counts]
                legend_labels_with_counts = [f'{item[0]} ({item[1]})' for item in sorted_counts]

            if others_count > 0: self.log_message(f"Chart Note: Grouped {len(others_sources)} types ({others_count} files) into '{others_label}'.")

            chart_window = Toplevel(self.master)
            # Use the passed title directly
            chart_window.title(display_path_or_title)
            chart_window.geometry("900x700")
            chart_window.minsize(600, 400)

            fig = Figure(figsize=(9, 7), dpi=100)
            ax = fig.add_subplot(111)

            explode_value = 0.02
            explode_list = [explode_value] * len(chart_labels)
            if others_count > 0 and others_label in chart_labels:
                try: explode_list[chart_labels.index(others_label)] = 0
                except ValueError: pass

            wedges, texts, autotexts = ax.pie(
                chart_sizes, explode=explode_list, labels=None,
                autopct=lambda pct: f"{pct:.1f}%" if pct > 1.5 else '',
                startangle=140, pctdistance=0.85,
                wedgeprops=dict(width=0.6, edgecolor='w')
            )
            ax.axis('equal')

            legend = ax.legend(wedges, legend_labels_with_counts,
                                title=self._("chart_legend_title", default="Extensions"),
                                loc="center left", bbox_to_anchor=(1, 0, 0.5, 1),
                                fontsize='small', frameon=True, labelspacing=0.8)

            for autotext in autotexts:
                if autotext.get_text():
                    autotext.set_color('white')
                    autotext.set_fontsize(8)
                    autotext.set_weight('bold')
                    autotext.set_bbox(dict(facecolor='black', alpha=0.5, pad=1, edgecolor='none'))

            # Use the passed title for the chart itself too
            chart_title = f"{display_path_or_title}\n(Total Files: {total_count})"
            ax.set_title(chart_title, pad=20, fontsize=12)

            try: fig.tight_layout(rect=[0, 0, 0.75, 1])
            except Exception as layout_err:
                print(f"Warning: Chart layout adjustment failed: {layout_err}.")
                self.log_message(f"Warning: Chart layout adjustment failed: {layout_err}")

            canvas = FigureCanvasTkAgg(fig, master=chart_window)
            canvas_widget = canvas.get_tk_widget()
            toolbar = NavigationToolbar2Tk(canvas, chart_window)
            toolbar.update()
            toolbar.pack(side=tk.BOTTOM, fill=tk.X)
            canvas_widget.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
            canvas.draw()
            chart_window.focus_set()
            chart_window.lift()

        except Exception as e:
            error_msg = f"Error creating or displaying chart window: {e}"
            self.log_message(error_msg)
            self.log_message(traceback.format_exc())
            if self.master.winfo_exists():
                messagebox.showerror(title=self._("chart_error_title", default="Chart Error"), message=error_msg, master=self.master)
            if chart_window and chart_window.winfo_exists():
                try: chart_window.destroy()
                except Exception: pass
# --- End of DuplicateFinderApp Class ---


# --- Main Execution Block ---
if __name__ == "__main__":
    try:
        from ctypes import windll
        try: windll.shcore.SetProcessDpiAwareness(1)
        except AttributeError:
            try: windll.user32.SetProcessDPIAware()
            except AttributeError: pass
    except (ImportError, AttributeError): pass

    root = tk.Tk()
    try:
        if not translations["en"].get("window_title") or not translations["zh"].get("window_title"):
            print("ERROR: Core translations missing. Exiting.")
            try:
                root_err = tk.Tk()
                root_err.withdraw()
                messagebox.showerror(title="Fatal Error", message="Core translation strings missing.", master=root_err)
                root_err.destroy()
            except Exception: pass
            sys.exit(1)

        app = DuplicateFinderApp(root)
        root.mainloop()

    except Exception as main_e:
        print("\n" + "=" * 30 + " FATAL APPLICATION ERROR " + "=" * 30)
        print(traceback.format_exc())
        print("=" * 80 + "\n")
        try:
            root_err = tk.Tk()
            root_err.withdraw()
            error_details = f"A critical error occurred:\n\n{type(main_e).__name__}: {main_e}"
            messagebox.showerror(title="Fatal Error", message=error_details, master=root_err)
            root_err.destroy()
        except Exception as mb_err:
            print(f"CRITICAL: Could not display fatal error message in GUI: {mb_err}")
        sys.exit(1)
