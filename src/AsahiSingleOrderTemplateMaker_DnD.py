# -- coding: utf-8 --
###############################################################
#
# AsahiOrderTemplateMaker_DnD.py
#
# pip install openpyxl pywin32 tkcalendar
#
###############################################################

import os
import subprocess
import sys

import win32api
import win32con
import win32gui

from AsahiSingleOrderTemplateMaker_Cmd import (
    PRODUCT_DELETE_PROBABILITY,
    report_processing_error,
    select_start_monday,
)


WINDOW_TITLE: str = "Asahi Order Template Maker (Drag & Drop)"
MAPPING_FILE_NAME: str = "AsahiOrderAreaStoreMapping_対応表.txt"


def show_message_box(pszMessage: str, pszTitle: str) -> None:
    """正常な処理結果を情報アイコン付きメッセージボックスで表示します。"""
    win32gui.MessageBox(
        0, pszMessage, pszTitle, win32con.MB_OK | win32con.MB_ICONINFORMATION
    )


def show_error_message_box(pszMessage: str, pszTitle: str) -> None:
    """エラー内容をエラーアイコン付きメッセージボックスで表示します。"""
    win32gui.MessageBox(0, pszMessage, pszTitle, win32con.MB_OK | win32con.MB_ICONERROR)


def run_asahi_single_order_template_maker_cmd(
    pszInputFileFullPath: str,
    pszStartMonday: str,
    pszMappingFileFullPath: str,
) -> tuple[bool, str]:
    """同じフォルダーのCmdプログラムを実行します。"""
    pszCurrentDirectoryFullPath: str = os.path.dirname(os.path.abspath(__file__))
    pszScriptFileName: str = "AsahiSingleOrderTemplateMaker_Cmd.py"
    pszScriptFileFullPath: str = os.path.join(
        pszCurrentDirectoryFullPath, pszScriptFileName
    )
    if not os.path.exists(pszScriptFileFullPath):
        return (
            False,
            "Error: "
            + pszScriptFileName
            + " not found. Path = "
            + pszScriptFileFullPath,
        )
    try:
        objCompletedProcess: subprocess.CompletedProcess[str] = subprocess.run(
            [
                sys.executable,
                pszScriptFileFullPath,
                "--start-monday",
                pszStartMonday,
                "--mapping-file",
                pszMappingFileFullPath,
                pszInputFileFullPath,
            ],
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


def draw_instruction_text(iWindowHandle: int) -> None:
    """ドラッグ＆ドロップ操作の案内を描画します。"""
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
        "Excel、TSV、またはCSVファイルをこのウィンドウにドラッグ＆ドロップしてください。\n"
        "開始月曜日をカレンダーから選択します。初期選択は来週の月曜日です。\n"
        "step0005では商品ごとに"
        + str(PRODUCT_DELETE_PROBABILITY * 100)
        + "%の独立した削除抽選を行います。\n"
        "プログラムと同じフォルダーの配送センター・店舗対応表を使用します。\n"
        "同じフォルダーにstep0001～step0007のXLSX・TSVを作成します。\n"
        "step0007では配送センターごとに発注のない商品を除外します。\n"
        "既存の出力ファイルは自動的に上書きします。\n"
        "キャンセルまたはエラー時は <元ファイル名>_error.txt を出力します。"
    )
    win32gui.DrawText(
        iDeviceContextHandle,
        pszInstructionText,
        -1,
        objClientRect,
        win32con.DT_LEFT | win32con.DT_TOP | win32con.DT_WORDBREAK,
    )
    win32gui.EndPaint(iWindowHandle, objPaintStruct)


def window_proc(
    iWindowHandle: int, iMessage: int, iWparam: int, iLparam: int
) -> int:
    """Windowsメッセージを処理します。"""
    if iMessage == win32con.WM_CREATE:
        win32gui.DragAcceptFiles(iWindowHandle, True)
        return 0
    if iMessage == win32con.WM_DROPFILES:
        iDropHandle: int = iWparam
        try:
            iFileCount: int = win32api.DragQueryFile(iDropHandle, -1)
            if iFileCount < 1:
                show_error_message_box("Error: no files were dropped.", WINDOW_TITLE)
                return 0
            listDroppedFilePaths: list[str] = [
                win32api.DragQueryFile(iDropHandle, iFileIndex)
                for iFileIndex in range(iFileCount)
            ]
            pszCurrentDirectoryFullPath: str = os.path.dirname(os.path.abspath(__file__))
            pszMappingFileFullPath: str = os.path.join(
                pszCurrentDirectoryFullPath, MAPPING_FILE_NAME
            )
            if not os.path.isfile(pszMappingFileFullPath):
                pszMappingError: str = (
                    MAPPING_FILE_NAME
                    + " が見つかりません。\n\n"
                    + "プログラムと同じフォルダーに配置してください。\n"
                    + "処理0006を完了できないため、処理を中止します。"
                )
                for pszDroppedFilePath in listDroppedFilePaths:
                    report_processing_error(
                        pszDroppedFilePath,
                        "旭注文テンプレート処理0006",
                        pszMappingError,
                    )
                show_error_message_box(pszMappingError, WINDOW_TITLE)
                return 0
            try:
                objStartMonday = select_start_monday()
            except Exception as objException:
                pszSelectionError: str = (
                    "開始月曜日の選択中にエラーが発生しました。Detail = "
                    + str(objException)
                )
                for pszDroppedFilePath in listDroppedFilePaths:
                    report_processing_error(
                        pszDroppedFilePath,
                        "旭注文テンプレート処理0004",
                        pszSelectionError,
                    )
                show_error_message_box(pszSelectionError, WINDOW_TITLE)
                return 0
            if objStartMonday is None:
                pszCancellationDetail: str = (
                    "開始月曜日の選択がキャンセルされたため、処理0004を中止しました。"
                )
                for pszDroppedFilePath in listDroppedFilePaths:
                    report_processing_error(
                        pszDroppedFilePath,
                        "旭注文テンプレート処理0004",
                        pszCancellationDetail,
                    )
                show_error_message_box(
                    "開始月曜日の選択がキャンセルされました。\n"
                    "処理0004を完了できないため、処理を中止します。",
                    WINDOW_TITLE,
                )
                return 0
            pszStartMonday: str = objStartMonday.isoformat()
            listFailedFileNames: list[str] = []
            listSuccessMessages: list[str] = []
            iSuccessCount: int = 0
            for pszDroppedFilePath in listDroppedFilePaths:
                bIsSuccess, pszResultMessage = (
                    run_asahi_single_order_template_maker_cmd(
                        pszDroppedFilePath, pszStartMonday, pszMappingFileFullPath
                    )
                )
                if bIsSuccess:
                    iSuccessCount += 1
                    listWarnings: list[str] = [
                        pszLine for pszLine in pszResultMessage.splitlines()
                        if pszLine.startswith("警告:")
                    ]
                    if listWarnings:
                        listSuccessMessages.extend(listWarnings)
                else:
                    listFailedFileNames.append(os.path.basename(pszDroppedFilePath))
            pszMessage: str = (
                "完了: "
                + str(iFileCount)
                + "件中 "
                + str(iSuccessCount)
                + "件成功 / "
                + str(len(listFailedFileNames))
                + "件失敗"
            )
            if listFailedFileNames:
                pszMessage += "\n失敗: " + ", ".join(listFailedFileNames)
                show_error_message_box(pszMessage, WINDOW_TITLE)
            else:
                if listSuccessMessages:
                    pszMessage += "\n\n" + "\n".join(listSuccessMessages)
                show_message_box(pszMessage, WINDOW_TITLE)
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
        iWindowHandle, iMessage, iWparam, iLparam
    )


def register_window_class(pszWindowClassName: str) -> int:
    """DnDウィンドウ用のWindowsクラスを登録します。"""
    iInstanceHandle: int = win32api.GetModuleHandle(None)
    objWndClass = win32gui.WNDCLASS()
    objWndClass.hInstance = iInstanceHandle
    objWndClass.lpszClassName = pszWindowClassName
    objWndClass.lpfnWndProc = window_proc
    objWndClass.style = win32con.CS_HREDRAW | win32con.CS_VREDRAW
    objWndClass.hCursor = win32gui.LoadCursor(0, win32con.IDC_ARROW)
    objWndClass.hbrBackground = win32con.COLOR_WINDOW + 1
    return win32gui.RegisterClass(objWndClass)


def create_main_window(pszWindowClassName: str, pszWindowTitle: str) -> int:
    """最前面表示のドラッグ＆ドロップ受付ウィンドウを作成します。"""
    iInstanceHandle: int = win32api.GetModuleHandle(None)
    iWindowStyle: int = (
        win32con.WS_OVERLAPPED
        | win32con.WS_CAPTION
        | win32con.WS_SYSMENU
        | win32con.WS_MINIMIZEBOX
    )
    iWindowHeight: int = 260
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


def main() -> None:
    """DnDウィンドウを作成してWindowsメッセージループを開始します。"""
    pszWindowClassName: str = "AsahiOrderTemplateMakerDndWindowClass"
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
