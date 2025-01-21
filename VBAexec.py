import sys
import os
import re
import datetime
from pathlib import Path

import xlwings as xw
from tkinter import messagebox

def execute_checked():
    # 檢查當前時間是否在可執行範圍內。
    now = datetime.datetime.now()
    start_time = now.replace(hour=9, minute=0, second=0, microsecond=0)
    return now < start_time

def file_matching(folder, pattern):
    # 匹配文件夾中符合模式的文件。
    return [str(file) for file in Path(folder).glob(pattern)]

def select_file(files):
    # 提供用戶交互以選擇文件。
    for idx, file in enumerate(files):
        print(f"{idx}. {file}")
    
    while True:
        try:
            user_input = int(input("Choose file by number: "))
            if 0 <= user_input < len(files):
                return user_input
            else:
                print("Invalid input. Please try again.")
        except ValueError:
            print("Please enter a valid number.")

def run_macro(path):
    # 運行 Excel 宏，處理錯誤並管理資源。
    try:
        with xw.App(visible=False) as app:
            wb = app.books.open(path)
            print("Workbook opened...")
            macro = wb.macro("Module1.main")
            print("Executing macro...")
            try:
                macro()
            except Exception as macro_error:
                # 捕捉 macro 內部錯誤
                raise RuntimeError(f"Macro execution failed: {macro_error}")
    except FileNotFoundError:
        # 處理文件未找到的情況
        messagebox.showwarning('Execution Failed', f'{path} not found.')
    except AttributeError:
        # 處理 macro 不存在的情況
        messagebox.showwarning('Execution Failed', 'Macro not found in the workbook.')
    except RuntimeError as e:
        # 處理 macro 執行時的錯誤
        messagebox.showerror('Execution Failed', str(e))
    except Exception as e:
        # 處理其他未知錯誤
        messagebox.showerror('Execution Failedwat?', f'Unexpected error: {str(e)}')
    finally:
        # 確保關閉工作簿與應用程序
        if 'wb' in locals() and wb:
            wb.close()
        if 'app' in locals() and app:
            app.quit()
        print("Execution completed.")

# 主程序
if __name__ == "__main__":
    # 設為使用者輸入，如送出為空則回傳後者(首個非空值)
    folder = input("Enter the folder path: ").strip() or r"path:\for\VBA\script"
    filePattern = input("Enter the file pattern: ").strip() or r"path:\for\VBA\script"

    matchedFiles = file_matching(folder, filePattern)
    if not matchedFiles:
        print("No files found.")
        sys.exit()

    if len(matchedFiles) == 1:
        selectedFileIndex = 0
    else:
        selectedFileIndex = select_file(matchedFiles)

    selectedFilePath = matchedFiles[selectedFileIndex]
    if execute_checked():
        print("Waiting for NAS4 response...")
        run_macro(selectedFilePath)
    else:
        print("Execution time has passed. Exiting...")
    