import os # 檔案操作
import shutil # 進階檔案控制，移動/壓縮/解壓縮 等等
# from tkinter import messagebox as msgbox # 提示視窗
import datetime # 日期
import re # 正則

def tidyDir(directory, keywords="", removeMatch=True):
    ''' 
    first parameter: work for which folder (required) (str)
    second parameter: kill or keep files when keywords not is empty (optional) (str)
    third parameter: judgement kill or keep files default True (optional) (bool)
    '''
    files = os.listdir(directory)
    os.chdir(directory)
    # 檢查是否設置關鍵字
    if keywords:
        # 刪除符合條件的文件
        return [os.remove(os.path.join(directory,delFile)) for delFile in files if any(keyword in delFile for keyword in keywords) == removeMatch]
    else:
        # 移除資料夾併重建
        try:
            shutil.rmtree(directory)
            os.mkdir(directory)
            return True
        except OSError as e:
            return False, str(e)
def moveToDest(srcDir, destDir, keywords="", moveMatch=True):
    '''
    first parameter: where files from (str)
    second parameter: destination for move (str)
    third parameter: file name for move or freeze (str)
    fourth parameter: judgement third parameter is move or freeze (bool)
    '''
    files = os.listdir(srcDir)
    # 檢查是否設置關鍵字
    if keywords:
        # 移動符合條件的文件到目標文件夾
        for filename in files:
            if any(keyword in filename for keyword in keywords) == moveMatch:
                srcPath = os.path.join(srcDir, filename)
                destPath = os.path.join(destDir, filename)
                shutil.move(srcPath, destPath)
        return True
    else:
        # 移動資料夾併重建
        try:
            shutil.move(srcDir, destDir)
            os.mkdir(srcDir)
            return True
        except Exception as e:
            return False, str(e)
def executeChecked():
    now = datetime.datetime.now()
    fmtNow = now.strftime("%Y/%m/%d %H:%M")
    todayMoring = now.strftime("%Y/%m/%d 09:00")
    return fmtNow < todayMoring
