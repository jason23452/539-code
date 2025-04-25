#!/usr/bin/env python3
import os
import subprocess
import sys
import shutil

# 第一个 AppleScript：回测脚本（5 个参数）
applescript_backtest = r'''-- AppLauncher.scpt （5 参数回测）
on RunAppTask(theParams)
    set paramList to paragraphs of theParams
    if (count of paramList) < 5 then
        return "Error: 參數數量不足，需 5 個，但傳入了 " & (count of paramList)
    end if

    set drawsSheet to item 1 of paramList
    set colRange   to item 2 of paramList
    set prizeSheet to item 3 of paramList
    set wbPath     to item 4 of paramList
    set appPath    to item 5 of paramList

    set cmd to "open -a " & quoted form of appPath & " --args " & ¬
              quoted form of wbPath & " " & ¬
              quoted form of drawsSheet & " " & ¬
              quoted form of colRange & " " & ¬
              quoted form of prizeSheet
    try
        do shell script cmd
        return "OK"
    on error errMsg
        return "Error: " & errMsg
    end try
end RunAppTask
'''

# 第二个 AppleScript：Lottery 脚本（6 个参数）
applescript_lottery = r'''-- LotteryLauncher.scpt （6 参数 Lottery）
on RunAppTask(theParams)
    set paramsList to paragraphs of theParams
    if (count of paramsList) < 6 then
        return "Error: Missing parameters."
    end if

    set sheetRange to item 1 of paramsList
    set comboSize  to item 2 of paramsList
    set wbPath     to item 3 of paramsList
    set appBundle  to item 4 of paramsList
    set topN       to item 5 of paramsList
    set maxGap     to item 6 of paramsList

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

def compile_and_deploy(name: str, source: str, target_dir: str):
    ascpt = os.path.join("/tmp", f"{name}.applescript")
    scpt  = os.path.join("/tmp", f"{name}.scpt")
    # 写入源文件
    with open(ascpt, "w") as f:
        f.write(source)
    # 编译
    subprocess.run(["osacompile", "-o", scpt, ascpt], check=True)
    # 部署
    dest = os.path.join(target_dir, f"{name}.scpt")
    shutil.move(scpt, dest)
    os.remove(ascpt)
    print(f"→ 已部署 {name}.scpt")

def main():
    # Excel for Mac AppleScript 存放目录
    target_dir = os.path.expanduser(
        "~/Library/Application Scripts/com.microsoft.Excel"
    )
    os.makedirs(target_dir, exist_ok=True)

    try:
        compile_and_deploy("AppLauncher", applescript_backtest, target_dir)
        compile_and_deploy("LotteryLauncher", applescript_lottery, target_dir)
    except subprocess.CalledProcessError as e:
        print("❌ 编译 AppleScript 失败：", e)
        sys.exit(1)
    except Exception as e:
        print("❌ 部署失败：", e)
        sys.exit(1)

    print("✅ 两个 AppleScriptTask 都已安装完成！")

if __name__ == "__main__":
    main()
