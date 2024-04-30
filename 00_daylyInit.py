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
today = datetime.datetime.today()
yesterday = datetime.datetime.today() - datetime.timedelta(1)
fmtYesterday = yesterday.strftime("%Y%m%d")
# 下載資料夾
downloadDir = r"D:\Owner\Documents"
# 平台下載的檔案
yahoFolder = r"D:\Tools\每日庫存\資料\Yahoo"
appfolder = r"D:\Tools\每日庫存\資料\APP"
nasYahoFolder = r"\\ileynas4\網路平台備查資料\工具\每日庫存\資料\YAHOO"
# 更新要的資料
updatedir = r"D:\Tools\庫存調整\資料"
# 要上傳的資料夾
uploadDir = r"D:\Tools\庫存調整\完成\每日調整"
trainsitionFile = r"D:\Tools\庫存調整\完成"
# 備查資料夾
myDest = r"D:\Tools\Z.備查資料"
if not yesterday.isoweekday() == 7:
    dstfolder = f"{myDest}\{fmtYesterday}"
else:
    prevfri = (today - datetime.timedelta(days=3)).strftime("%Y%m%d")
    dstfolder = f"{myDest}\{prevfri}"
# 周報表
weeklyReport = r"D:\Tools\週一-週報表"

# execute before 9 am
if executeChecked():
    # delete
    [tidyDir(dir) for dir in [downloadDir, nasYahoFolder]]
    tidyDir(appfolder,"BA",True)
    tidyDir(yahoFolder,"src",True)

    # moving
    os.mkdir(dstfolder)
    [moveToDest(dir,dstfolder,fmtYesterday) for dir in [uploadDir, trainsitionFile]]
    moveToDest(updatedir,dstfolder,["需求", "轉正"],False)

    # create weekly report
    if today.weekday() == 0:
        prevWeek = (today - datetime.timedelta(7)).strftime("%Y%m%d")
        MDYesterday = (today - datetime.timedelta(1)).strftime("%m%d")
        os.mkdir(f"{weeklyReport}/{prevWeek}-{MDYesterday}")