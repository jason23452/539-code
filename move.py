#!/usr/bin/env python3
import os
import subprocess
import sys
import shutil

def main():
    # 定義 AppleScript 原始碼內容 (即你的 AppLauncher.scpt 內容)
    applescript_source = r'''on RunAppTask(theParams)
	-- Split the parameter string by line breaks.
	set paramsList to paragraphs of theParams
	if (count of paramsList) is less than 4 then
		return "Error: Missing parameters."
	end if
	
	set sheetRange to item 1 of paramsList
	set comboSize to item 2 of paramsList
	set wbPath to item 3 of paramsList
	set appPath to item 4 of paramsList
	
	-- Construct the command string.
	-- Make sure the app has the execute permission.
	set cmd to quoted form of appPath & " " & quoted form of sheetRange & " " & quoted form of comboSize & " " & quoted form of wbPath
	
	try
		do shell script cmd
	on error errMsg
		return "Error: " & errMsg
	end try
	return "App launched successfully."
end RunAppTask
'''

    # 暫存目錄與檔案名稱
    tmp_dir = "/tmp"
    source_file = os.path.join(tmp_dir, "AppLauncher.applescript")
    compiled_file = os.path.join(tmp_dir, "AppLauncher.scpt")
    
    # 將原始 AppleScript 寫入暫存檔案
    try:
        with open(source_file, "w") as f:
            f.write(applescript_source)
        print(f"已寫入暫存 AppleScript 檔案: {source_file}")
    except Exception as e:
        print("寫入暫存檔案失敗:", e)
        sys.exit(1)
    
    # 呼叫 osacompile 編譯成 scpt 格式
    compile_cmd = ["osacompile", "-o", compiled_file, source_file]
    print("正在編譯 AppleScript...")
    try:
        subprocess.run(compile_cmd, check=True)
        print(f"已編譯成 scpt 檔案: {compiled_file}")
    except subprocess.CalledProcessError as e:
        print("AppleScript 編譯失敗:", e)
        sys.exit(1)
    
    # 目標資料夾路徑，注意 macOS 的路徑區分大小寫
    target_dir = os.path.expanduser("~/Library/Application Scripts/com.microsoft.Excel")
    target_file = os.path.join(target_dir, "AppLauncher.scpt")

    # 如果目標資料夾不存在則建立之
    os.makedirs(target_dir, exist_ok=True)
    
    # 將編譯後的檔案移動到目標資料夾
    try:
        shutil.move(compiled_file, target_file)
        print(f"已將 AppLauncher.scpt 移動到: {target_file}")
    except Exception as e:
        print("移動檔案失敗:", e)
        sys.exit(1)
    
    # 刪除暫存的原始 AppleScript 檔案（可選）
    try:
        os.remove(source_file)
    except Exception as e:
        print("刪除暫存 AppleScript 檔案失敗，但可忽略此錯誤:", e)
    
    print("安裝程序已完成。")

if __name__ == "__main__":
    main()
