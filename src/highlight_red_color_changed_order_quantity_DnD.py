# -- coding: utf-8 --
###############################################################
#
# highlight_red_color_changed_order_quantity_DnD.py
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


WINDOW_TITLE: str = "Highlight Red Color Changed Order Quantity DnD"
EXCEL_PATH_ENVIRONMENT_VARIABLE: str = "HIGHLIGHT_ORDER_EXCEL_PATH"


###############################################################
#
# show_message_box
#
###############################################################
def show_message_box(
    pszMessage: str,
    pszTitle: str,
) -> None:
    """正常な処理結果を情報アイコン付きメッセージボックスで表示します。"""
    iOwnerWindowHandle: int = 0
    iMessageBoxType: int = win32con.MB_OK | win32con.MB_ICONINFORMATION
    win32gui.MessageBox(
        iOwnerWindowHandle,
        pszMessage,
        pszTitle,
        iMessageBoxType,
    )


###############################################################
#
# show_error_message_box
#
###############################################################
def show_error_message_box(
    pszMessage: str,
    pszTitle: str,
) -> None:
    """エラー内容をエラーアイコン付きメッセージボックスで表示します。"""
    iOwnerWindowHandle: int = 0
    iMessageBoxType: int = win32con.MB_OK | win32con.MB_ICONERROR
    win32gui.MessageBox(
        iOwnerWindowHandle,
        pszMessage,
        pszTitle,
        iMessageBoxType,
    )


###############################################################
#
# run_highlight_red_color_changed_order_quantity_cmd
#
###############################################################
def run_highlight_red_color_changed_order_quantity_cmd(
    pszInputFileFullPath: str,
    pszMode: str = "",
) -> tuple[bool, str]:
    """同じフォルダーのCmdプログラムをsubprocess.run()で実行します。

    Args:
        pszInputFileFullPath: ドロップされたExcelファイルのパスです。
        pszMode: 編集前準備の場合は--prepare、比較処理の場合は空文字です。

    Returns:
        成功したかどうかと、標準出力またはエラー内容を返します。
    """
    pszCurrentDirectoryFullPath: str = os.path.dirname(os.path.abspath(__file__))
    pszScriptFileName: str = "highlight_red_color_changed_order_quantity_Cmd.py"
    pszScriptFileFullPath: str = os.path.join(
        pszCurrentDirectoryFullPath,
        pszScriptFileName,
    )

    if not os.path.exists(pszScriptFileFullPath):
        pszErrorMessage: str = (
            "Error: "
            + pszScriptFileName
            + " not found. Path = "
            + pszScriptFileFullPath
        )
        return False, pszErrorMessage

    pszPythonExecutableFullPath: str = sys.executable
    listCommandArguments: list[str] = [
        pszPythonExecutableFullPath,
        pszScriptFileFullPath,
    ]
    if pszMode != "":
        listCommandArguments.append(pszMode)
    listCommandArguments.append(pszInputFileFullPath)
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
            + " exited with non-zero return code.\n\n"
            + "Return code = "
            + str(objCompletedProcess.returncode)
            + "\n\n"
            + "stderr:\n"
            + pszStdErr,
        )

    pszStdOut: str = objCompletedProcess.stdout
    if pszStdOut.strip() == "":
        pszStdOut = pszScriptFileName + " finished successfully."
    return True, pszStdOut


###############################################################
#
# open_excel_file_and_wait
#
###############################################################
def open_excel_file_and_wait(
    pszInputFileFullPath: str,
) -> tuple[bool, str]:
    """関連付けられたExcelアプリで入力ファイルを開き、終了まで待機します。

    担当者がExcelを編集、保存、終了した後に数量比較を始められるよう、
    PowerShellのStart-Processを使用してExcelのプロセス終了を待ちます。

    Args:
        pszInputFileFullPath: ドロップされたExcelファイルの絶対パスです。

    Returns:
        Excelを正常に開いて終了を待てた場合はTrueと空文字を返します。
        起動または待機に失敗した場合はFalseとエラー内容を返します。
    """
    if not os.path.exists(pszInputFileFullPath):
        return (
            False,
            "Error: input Excel file not found. Path = " + pszInputFileFullPath,
        )
    if not os.path.isfile(pszInputFileFullPath):
        return (
            False,
            "Error: input path is not a file. Path = " + pszInputFileFullPath,
        )
    if os.path.splitext(pszInputFileFullPath)[1].lower() != ".xlsx":
        return (
            False,
            "Error: input file extension is not .xlsx. Path = "
            + pszInputFileFullPath,
        )

    dictEnvironment: dict[str, str] = os.environ.copy()
    dictEnvironment[EXCEL_PATH_ENVIRONMENT_VARIABLE] = os.path.abspath(
        pszInputFileFullPath
    )
    pszPowerShellCommand: str = (
        "$excelPath = $env:"
        + EXCEL_PATH_ENVIRONMENT_VARIABLE
        + "; "
        + "$process = Start-Process -FilePath $excelPath -PassThru; "
        + "$process.WaitForExit(); exit $process.ExitCode"
    )
    try:
        objCompletedProcess: subprocess.CompletedProcess[str] = subprocess.run(
            [
                "powershell.exe",
                "-NoProfile",
                "-Command",
                pszPowerShellCommand,
            ],
            check=False,
            capture_output=True,
            text=True,
            env=dictEnvironment,
        )
    except Exception as objException:
        return (
            False,
            "Error: Excelファイルを開けませんでした。Detail = "
            + str(objException),
        )

    if objCompletedProcess.returncode != 0:
        pszStdErr: str = objCompletedProcess.stderr.strip()
        if pszStdErr == "":
            pszStdErr = "Excel process exited with a non-zero return code."
        return (
            False,
            "Error: Excelファイルの起動または終了待機に失敗しました。\n\n"
            + "Return code = "
            + str(objCompletedProcess.returncode)
            + "\n\n"
            + "stderr:\n"
            + pszStdErr,
        )

    return True, ""


###############################################################
#
# open_excel_file_without_wait
#
###############################################################
def open_excel_file_without_wait(
    pszInputFileFullPath: str,
) -> tuple[bool, str]:
    """処理済みExcelを通常の編集可能な状態で開き、終了を待たずに戻ります。

    Args:
        pszInputFileFullPath: 赤色設定と履歴保存が完了したExcelのパスです。

    Returns:
        起動要求に成功した場合はTrueと空文字を返します。起動に失敗した
        場合はFalseと、担当者へ表示するエラー内容を返します。
    """
    if not os.path.exists(pszInputFileFullPath):
        return (
            False,
            "Error: input Excel file not found. Path = " + pszInputFileFullPath,
        )
    if not os.path.isfile(pszInputFileFullPath):
        return (
            False,
            "Error: input path is not a file. Path = " + pszInputFileFullPath,
        )
    if os.path.splitext(pszInputFileFullPath)[1].lower() != ".xlsx":
        return (
            False,
            "Error: input file extension is not .xlsx. Path = "
            + pszInputFileFullPath,
        )

    dictEnvironment: dict[str, str] = os.environ.copy()
    dictEnvironment[EXCEL_PATH_ENVIRONMENT_VARIABLE] = os.path.abspath(
        pszInputFileFullPath
    )
    pszPowerShellCommand: str = (
        "$excelPath = $env:"
        + EXCEL_PATH_ENVIRONMENT_VARIABLE
        + "; Start-Process -FilePath $excelPath"
    )
    try:
        objCompletedProcess: subprocess.CompletedProcess[str] = subprocess.run(
            [
                "powershell.exe",
                "-NoProfile",
                "-Command",
                pszPowerShellCommand,
            ],
            check=False,
            capture_output=True,
            text=True,
            env=dictEnvironment,
        )
    except Exception as objException:
        return (
            False,
            "Error: 処理済みExcelを再表示できませんでした。Detail = "
            + str(objException),
        )

    if objCompletedProcess.returncode != 0:
        pszStdErr: str = objCompletedProcess.stderr.strip()
        if pszStdErr == "":
            pszStdErr = "PowerShell exited with a non-zero return code."
        return (
            False,
            "Error: 処理済みExcelを再表示できませんでした。\n\n"
            + "Return code = "
            + str(objCompletedProcess.returncode)
            + "\n\n"
            + "stderr:\n"
            + pszStdErr,
        )

    return True, ""


###############################################################
#
# draw_instruction_text
#
###############################################################
def draw_instruction_text(iWindowHandle: int) -> None:
    """ウィンドウ内に、Excelを1つドロップする操作説明を描画します。"""
    iDeviceContextHandle, objPaintStruct = win32gui.BeginPaint(iWindowHandle)
    objClientRect = win32gui.GetClientRect(iWindowHandle)
    iMargin: int = 5
    objClientRect = (
        objClientRect[0] + iMargin,
        objClientRect[1] + iMargin,
        objClientRect[2] - iMargin,
        objClientRect[3] - iMargin,
    )
    pszInstructionText: str = (
        "発注Excelファイルを1つ、このウィンドウにドラッグ＆ドロップしてください。\n"
        "Excelが開いたら数量を編集し、上書き保存してExcelを閉じてください。\n"
        "Excelを閉じた後、変更されたquantityセルを赤色にします。\n"
        "処理後、通常どおり編集できるExcelを自動的に再表示します。\n"
        "エラーが発生した場合は <Excelファイル名>_error.txt を出力します。"
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
                    "Error: 発注Excelファイルを1つだけドロップしてください。\n"
                    + "Dropped file count = "
                    + str(iFileCount),
                    WINDOW_TITLE,
                )
                return 0

            pszDroppedFilePath: str = win32api.DragQueryFile(iDropHandle, 0)
            bIsPrepareSuccess, pszPrepareMessage = (
                run_highlight_red_color_changed_order_quantity_cmd(
                    pszDroppedFilePath,
                    "--prepare",
                )
            )
            if not bIsPrepareSuccess:
                show_error_message_box(pszPrepareMessage, WINDOW_TITLE)
                return 0

            bIsExcelClosed, pszExcelMessage = open_excel_file_and_wait(
                pszDroppedFilePath
            )
            if not bIsExcelClosed:
                show_error_message_box(pszExcelMessage, WINDOW_TITLE)
                return 0

            bIsSuccess, pszMessage = run_highlight_red_color_changed_order_quantity_cmd(
                pszDroppedFilePath
            )
            if not bIsSuccess:
                show_error_message_box(pszMessage, WINDOW_TITLE)
                return 0

            bIsReopenSuccess, pszReopenMessage = open_excel_file_without_wait(
                pszDroppedFilePath
            )
            if not bIsReopenSuccess:
                show_error_message_box(
                    "赤色設定と履歴保存は完了しましたが、結果確認用のExcelを開けませんでした。\n\n"
                    + "Excelファイルを手動で開いて確認してください。\n\n"
                    + "対象ファイル: "
                    + pszDroppedFilePath
                    + "\n\n"
                    + pszReopenMessage,
                    WINDOW_TITLE,
                )
                return 0

            show_message_box(
                pszMessage
                + "\n結果確認用としてExcelを再度開きました。"
                + "\nExcelは通常どおり編集できます。"
                + "\nさらに編集した場合は、保存して閉じた後、改めてDnDしてください。",
                WINDOW_TITLE,
            )
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
    """DnDウィンドウに使用するWindowsクラスを登録し、atomを返します。"""
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
def create_main_window(
    pszWindowClassName: str,
    pszWindowTitle: str,
) -> int:
    """参照DnDと同じ大きさ・スタイルの最前面ウィンドウを作成します。"""
    iInstanceHandle: int = win32api.GetModuleHandle(None)
    iWindowStyle: int = (
        win32con.WS_OVERLAPPED
        | win32con.WS_CAPTION
        | win32con.WS_SYSMENU
        | win32con.WS_MINIMIZEBOX
    )
    iWindowExStyle: int = win32con.WS_EX_ACCEPTFILES
    iWindowPosX: int = win32con.CW_USEDEFAULT
    iWindowPosY: int = win32con.CW_USEDEFAULT
    iWindowHeight: int = 260
    iWindowWidth: int = int(iWindowHeight * 1.618)
    iWindowHandle: int = win32gui.CreateWindowEx(
        iWindowExStyle,
        pszWindowClassName,
        pszWindowTitle,
        iWindowStyle,
        iWindowPosX,
        iWindowPosY,
        iWindowWidth,
        iWindowHeight,
        0,
        0,
        iInstanceHandle,
        None,
    )
    win32gui.ShowWindow(iWindowHandle, win32con.SW_SHOWNORMAL)
    win32gui.UpdateWindow(iWindowHandle)
    iFlags: int = win32con.SWP_NOMOVE | win32con.SWP_NOSIZE
    win32gui.SetWindowPos(
        iWindowHandle,
        win32con.HWND_TOPMOST,
        0,
        0,
        0,
        0,
        iFlags,
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
    pszWindowClassName: str = "HighlightRedColorChangedOrderQuantityDndWindowClass"
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
            "Error: unexpected exception in message loop. Detail = " + str(objException),
            WINDOW_TITLE,
        )


if __name__ == "__main__":
    main()
