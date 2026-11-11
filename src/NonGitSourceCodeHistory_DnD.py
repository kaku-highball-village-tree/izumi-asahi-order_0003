# -- coding: utf-8 --
###############################################################
#
# NonGitSourceCodeHistory_DnD.py
#
# pip install pywin32
#
###############################################################

import os
import subprocess
import sys

import win32api
import win32con
import win32gui


WINDOW_TITLE: str = "NonGit Source Code History DnD"


###############################################################
#
# show_message_box
#
###############################################################
def show_message_box(pszMessage: str, pszTitle: str) -> None:
    """正常な処理結果を情報アイコン付きメッセージボックスで表示します。"""
    iMessageBoxType: int = win32con.MB_OK | win32con.MB_ICONINFORMATION
    win32gui.MessageBox(0, pszMessage, pszTitle, iMessageBoxType)


###############################################################
#
# show_error_message_box
#
###############################################################
def show_error_message_box(pszMessage: str, pszTitle: str) -> None:
    """エラー内容をエラーアイコン付きメッセージボックスで表示します。"""
    iMessageBoxType: int = win32con.MB_OK | win32con.MB_ICONERROR
    win32gui.MessageBox(0, pszMessage, pszTitle, iMessageBoxType)


###############################################################
#
# run_non_git_source_code_history_cmd
#
###############################################################
def run_non_git_source_code_history_cmd(
    pszInputFileFullPath: str,
) -> tuple[bool, str]:
    """同じフォルダーのCmdプログラムへ対象パスを渡して実行します。"""
    pszCurrentDirectoryFullPath: str = os.path.dirname(os.path.abspath(__file__))
    pszScriptFileName: str = "NonGitSourceCodeHistory_Cmd.py"
    pszScriptFileFullPath: str = os.path.join(
        pszCurrentDirectoryFullPath,
        pszScriptFileName,
    )
    if not os.path.isfile(pszScriptFileFullPath):
        return (
            False,
            "Error: "
            + pszScriptFileName
            + " not found. Path = "
            + pszScriptFileFullPath,
        )

    listCommandArguments: list[str] = [
        sys.executable,
        pszScriptFileFullPath,
        pszInputFileFullPath,
    ]
    try:
        objCompletedProcess: subprocess.CompletedProcess[str] = subprocess.run(
            listCommandArguments,
            check=False,
            capture_output=True,
            text=True,
        )
    except Exception as objException:
        return (
            False,
            "Error: unexpected exception while running "
            + pszScriptFileName
            + ". Detail = "
            + str(objException),
        )

    if objCompletedProcess.returncode != 0:
        pszStdErr: str = objCompletedProcess.stderr
        if pszStdErr.strip() == "":
            pszStdErr = "Process exited with non-zero return code and no stderr output."
        return (
            False,
            "Error: "
            + pszScriptFileName
            + " exited with non-zero return code.\n\nReturn code = "
            + str(objCompletedProcess.returncode)
            + "\n\nstderr:\n"
            + pszStdErr,
        )

    pszStdOut: str = objCompletedProcess.stdout
    if pszStdOut.strip() == "":
        pszStdOut = pszScriptFileName + " finished successfully."
    return True, pszStdOut


###############################################################
#
# draw_instruction_text
#
###############################################################
def draw_instruction_text(iWindowHandle: int) -> None:
    """ウィンドウ内にバックアップ対象を1つドロップする説明を描画します。"""
    iDeviceContextHandle, objPaintStruct = win32gui.BeginPaint(iWindowHandle)
    objClientRect = win32gui.GetClientRect(iWindowHandle)
    iMargin: int = 10
    objClientRect = (
        objClientRect[0] + iMargin,
        objClientRect[1] + iMargin,
        objClientRect[2] - iMargin,
        objClientRect[3] - iMargin,
    )
    pszInstructionText: str = (
        "バックアップしたいソースコードファイルを1つ、\n"
        + "このウィンドウにドラッグ＆ドロップしてください。\n\n"
        + "元ファイルは変更せず、\n同じフォルダーに\n\n"
        + "<元ファイル名>.bk0001.txt\n"
        + "<元ファイル名>.bk0002.txt\n...\n\n"
        + "の形式で履歴バックアップを作成します。"
    )
    iDrawTextFormat: int = win32con.DT_LEFT | win32con.DT_TOP | win32con.DT_WORDBREAK
    win32gui.DrawText(
        iDeviceContextHandle,
        pszInstructionText,
        -1,
        objClientRect,
        iDrawTextFormat,
    )
    win32gui.EndPaint(iWindowHandle, objPaintStruct)


###############################################################
#
# window_proc
#
###############################################################
def window_proc(
    iWindowHandle: int,
    iMessage: int,
    iWparam: int,
    iLparam: int,
) -> int:
    """Windowsメッセージを受け取り、DnD、描画、終了を処理します。"""
    if iMessage == win32con.WM_CREATE:
        win32gui.DragAcceptFiles(iWindowHandle, True)
        return 0
    if iMessage == win32con.WM_DROPFILES:
        iDropHandle: int = iWparam
        try:
            iFileCount: int = win32api.DragQueryFile(iDropHandle, -1)
            if iFileCount != 1:
                show_error_message_box(
                    "Error: バックアップしたいファイルを1つだけドロップしてください。\n"
                    + "Dropped file count = "
                    + str(iFileCount),
                    WINDOW_TITLE,
                )
                return 0
            pszDroppedFilePath: str = win32api.DragQueryFile(iDropHandle, 0)
            bIsSuccess, pszMessage = run_non_git_source_code_history_cmd(
                pszDroppedFilePath
            )
            if bIsSuccess:
                show_message_box(pszMessage, WINDOW_TITLE)
            else:
                show_error_message_box(pszMessage, WINDOW_TITLE)
        finally:
            win32api.DragFinish(iDropHandle)
        return 0
    if iMessage == win32con.WM_PAINT:
        draw_instruction_text(iWindowHandle)
        return 0
    if iMessage == win32con.WM_DESTROY:
        win32gui.PostQuitMessage(0)
        return 0
    return win32gui.DefWindowProc(
        iWindowHandle,
        iMessage,
        iWparam,
        iLparam,
    )


###############################################################
#
# register_window_class
#
###############################################################
def register_window_class(pszWindowClassName: str) -> int:
    """DnDウィンドウに使用するWindowsクラスを登録します。"""
    iInstanceHandle: int = win32api.GetModuleHandle(None)
    objWndClass = win32gui.WNDCLASS()
    objWndClass.hInstance = iInstanceHandle
    objWndClass.lpszClassName = pszWindowClassName
    objWndClass.lpfnWndProc = window_proc
    objWndClass.style = win32con.CS_HREDRAW | win32con.CS_VREDRAW
    objWndClass.hCursor = win32gui.LoadCursor(0, win32con.IDC_ARROW)
    objWndClass.hbrBackground = win32con.COLOR_WINDOW + 1
    return win32gui.RegisterClass(objWndClass)


###############################################################
#
# create_main_window
#
###############################################################
def create_main_window(pszWindowClassName: str, pszWindowTitle: str) -> int:
    """説明を表示する最前面のDnDウィンドウを作成します。"""
    iInstanceHandle: int = win32api.GetModuleHandle(None)
    iWindowStyle: int = (
        win32con.WS_OVERLAPPED
        | win32con.WS_CAPTION
        | win32con.WS_SYSMENU
        | win32con.WS_MINIMIZEBOX
    )
    iWindowHeight: int = 330
    iWindowWidth: int = int(iWindowHeight * 1.618)
    iWindowHandle: int = win32gui.CreateWindowEx(
        win32con.WS_EX_ACCEPTFILES,
        pszWindowClassName,
        pszWindowTitle,
        iWindowStyle,
        win32con.CW_USEDEFAULT,
        win32con.CW_USEDEFAULT,
        iWindowWidth,
        iWindowHeight,
        0,
        0,
        iInstanceHandle,
        None,
    )
    win32gui.ShowWindow(iWindowHandle, win32con.SW_SHOWNORMAL)
    win32gui.UpdateWindow(iWindowHandle)
    win32gui.SetWindowPos(
        iWindowHandle,
        win32con.HWND_TOPMOST,
        0,
        0,
        0,
        0,
        win32con.SWP_NOMOVE | win32con.SWP_NOSIZE,
    )
    win32gui.DragAcceptFiles(iWindowHandle, True)
    return iWindowHandle


###############################################################
#
# main
#
###############################################################
def main() -> None:
    """ウィンドウを準備し、Windowsメッセージループを開始します。"""
    pszWindowClassName: str = "NonGitSourceCodeHistoryDndWindowClass"
    try:
        register_window_class(pszWindowClassName)
    except Exception as objException:
        show_error_message_box(
            "Error: failed to register window class. Detail = " + str(objException),
            WINDOW_TITLE,
        )
        return
    try:
        create_main_window(pszWindowClassName, WINDOW_TITLE)
    except Exception as objException:
        show_error_message_box(
            "Error: failed to create main window. Detail = " + str(objException),
            WINDOW_TITLE,
        )
        return
    try:
        win32gui.PumpMessages()
    except Exception as objException:
        show_error_message_box(
            "Error: unexpected exception in message loop. Detail = "
            + str(objException),
            WINDOW_TITLE,
        )


if __name__ == "__main__":
    main()
