#!/usr/bin/env python3
import os
import subprocess
import sys
import shutil

def main():
    # 定義新的 AppleScript 原始碼：用 open -a 啟動 .app 並傳參
    applescript_source = r'''on RunAppTask(theParams)
	-- 拆分參數
	set paramsList to paragraphs of theParams
	if (count of paramsList) is less than 6 then
		return "Error: Missing parameters."
	end if

	set sheetRange to item 1 of paramsList
	set comboSize  to item 2 of paramsList
	set wbPath     to item 3 of paramsList
	set appBundle  to item 4 of paramsList
	set topN       to item 5 of paramsList
	set maxGap     to item 6 of paramsList

	-- 使用 open -a 啟動 .app 並傳遞後面的參數
	set cmd to "open -a " & quoted form of appBundle & " --args " & ¬
	          quoted form of sheetRange & " " & ¬
	          quoted form of comboSize  & " " & ¬
	          quoted form of wbPath      & " " & ¬
	          quoted form of topN       & " " & ¬
	          quoted form of maxGap

	try
		do shell script cmd
	on error errMsg
		return "Error: " & errMsg
	end try
	return "App launched successfully."
end RunAppTask
'''

    tmp_dir = "/tmp"
    source_file = os.path.join(tmp_dir, "AppLauncher.applescript")
    compiled_file = os.path.join(tmp_dir, "AppLauncher.scpt")

    # 寫入並編譯
    try:
        with open(source_file, "w") as f:
            f.write(applescript_source)
        subprocess.run(["osacompile", "-o", compiled_file, source_file], check=True)
    except Exception as e:
        print("編譯 AppleScript 失敗:", e)
        sys.exit(1)

    # 部署到 Excel 的 AppleScript 目錄
    target_dir = os.path.expanduser("~/Library/Application Scripts/com.microsoft.Excel")
    os.makedirs(target_dir, exist_ok=True)
    target_file = os.path.join(target_dir, "AppLauncher.scpt")
    try:
        shutil.move(compiled_file, target_file)
    except Exception as e:
        print("移動檔案失敗:", e)
        sys.exit(1)

    # 清理
    try: os.remove(source_file)
    except: pass

    print("已安裝 AppLauncher.scpt，可透過 AppleScriptTask 以 .app 形式啟動。")

if __name__ == "__main__":
    main()
