# -- coding: utf-8 --
###############################################################
#
# AsahiOrderAreaStoreMappingMaker_DnD.py
#
# pip install openpyxl pywin32
#
###############################################################

import os
import re
import subprocess
import sys
import tempfile

import win32api
import win32con
import win32gui

from AsahiOrderAreaStoreMappingMaker_Cmd import report_processing_error


WINDOW_TITLE: str = "Asahi Order Area Store Mapping Maker (Drag & Drop)"
MAX_RESULT_BACKUP_NUMBER: int = 9999
FINAL_AREA_STORE_MAPPING_FILE_NAMES: tuple[str, str] = (
    "AsahiOrderAreaStoreMapping_対応表.txt",
    "AsahiOrderAreaStoreMapping_対応表.tsv",
)


def show_message_box(pszMessage: str, pszTitle: str) -> None:
    """正常な処理結果を情報アイコン付きメッセージボックスで表示します。"""
    win32gui.MessageBox(
        0, pszMessage, pszTitle, win32con.MB_OK | win32con.MB_ICONINFORMATION
    )


def show_error_message_box(pszMessage: str, pszTitle: str) -> None:
    """エラー内容をエラーアイコン付きメッセージボックスで表示します。"""
    win32gui.MessageBox(0, pszMessage, pszTitle, win32con.MB_OK | win32con.MB_ICONERROR)


def show_warning_message_box(pszMessage: str, pszTitle: str) -> None:
    """警告内容を警告アイコン付きメッセージボックスで表示します。"""
    win32gui.MessageBox(0, pszMessage, pszTitle, win32con.MB_OK | win32con.MB_ICONWARNING)


def get_result_file_full_path(pszInputFileFullPath: str) -> str:
    """入力Excelと同じフォルダーの<入力Excel名>_result.txtを返します。"""
    pszDirectoryFullPath: str = os.path.dirname(os.path.abspath(pszInputFileFullPath))
    pszBaseNameWithoutExtension: str = os.path.splitext(
        os.path.basename(pszInputFileFullPath)
    )[0]
    return os.path.join(
        pszDirectoryFullPath, pszBaseNameWithoutExtension + "_result.txt"
    )


def get_next_result_backup_full_path(pszResultFileFullPath: str) -> str:
    """既存結果テキストに対応する次の.bk0001.txt形式のパスを返します。"""
    pszDirectoryFullPath: str = os.path.dirname(pszResultFileFullPath)
    pszResultFileName: str = os.path.basename(pszResultFileFullPath)
    objPattern: re.Pattern[str] = re.compile(
        r"^" + re.escape(pszResultFileName) + r"\.bk([0-9]{4})\.txt$"
    )
    listBackupNumbers: list[int] = []
    for pszCandidateFileName in os.listdir(pszDirectoryFullPath):
        objMatch: re.Match[str] | None = objPattern.fullmatch(pszCandidateFileName)
        if objMatch is None:
            continue
        pszCandidateFullPath: str = os.path.join(
            pszDirectoryFullPath, pszCandidateFileName
        )
        if not os.path.isfile(pszCandidateFullPath):
            continue
        iBackupNumber: int = int(objMatch.group(1))
        if 1 <= iBackupNumber <= MAX_RESULT_BACKUP_NUMBER:
            listBackupNumbers.append(iBackupNumber)
    iNextBackupNumber: int = (
        1 if not listBackupNumbers else max(listBackupNumbers) + 1
    )
    if iNextBackupNumber > MAX_RESULT_BACKUP_NUMBER:
        raise ValueError(
            "結果テキストのバックアップ番号が最大値9999に到達しています。Path = "
            + pszResultFileFullPath
        )
    return pszResultFileFullPath + f".bk{iNextBackupNumber:04d}.txt"


def save_result_text(
    pszInputFileFullPath: str, pszResultText: str
) -> tuple[str, str | None]:
    """詳細結果をUTF-8で安全に保存し、既存結果があれば連番バックアップします。"""
    pszResultFileFullPath: str = get_result_file_full_path(pszInputFileFullPath)
    pszNormalizedResultText: str = pszResultText.rstrip("\r\n") + "\n"
    iFileDescriptor, pszTemporaryFileFullPath = tempfile.mkstemp(
        prefix=os.path.basename(pszResultFileFullPath) + "_",
        suffix=".txt",
        dir=os.path.dirname(pszResultFileFullPath),
    )
    os.close(iFileDescriptor)
    pszBackupFileFullPath: str | None = None
    bBackupCreated: bool = False
    try:
        with open(
            pszTemporaryFileFullPath, mode="w", encoding="utf-8", newline=""
        ) as objFile:
            objFile.write(pszNormalizedResultText)
        with open(
            pszTemporaryFileFullPath, mode="r", encoding="utf-8", newline=""
        ) as objFile:
            if objFile.read() != pszNormalizedResultText:
                raise ValueError("詳細結果テキストの保存後検証に失敗しました。")
        if os.path.exists(pszResultFileFullPath):
            if not os.path.isfile(pszResultFileFullPath):
                raise ValueError(
                    "詳細結果の出力先がファイルではありません。Path = "
                    + pszResultFileFullPath
                )
            pszBackupFileFullPath = get_next_result_backup_full_path(
                pszResultFileFullPath
            )
            os.rename(pszResultFileFullPath, pszBackupFileFullPath)
            bBackupCreated = True
        os.replace(pszTemporaryFileFullPath, pszResultFileFullPath)
    except Exception:
        if bBackupCreated and pszBackupFileFullPath is not None:
            if os.path.exists(pszResultFileFullPath):
                os.remove(pszResultFileFullPath)
            if os.path.exists(pszBackupFileFullPath):
                os.rename(pszBackupFileFullPath, pszResultFileFullPath)
        raise
    finally:
        if os.path.exists(pszTemporaryFileFullPath):
            os.remove(pszTemporaryFileFullPath)
    return pszResultFileFullPath, pszBackupFileFullPath


def build_success_summary_message(
    pszInputFileFullPath: str, pszResultFileFullPath: str
) -> str:
    """正常終了MessageBox用の短い要約を返します。"""
    return (
        "朝日注文エリア店舗対応TSVの作成が正常に完了しました。\n\n"
        + "入力ファイル:\n"
        + os.path.basename(pszInputFileFullPath)
        + "\n\n最終出力:\n"
        + "\n".join(FINAL_AREA_STORE_MAPPING_FILE_NAMES)
        + "\n\n詳細結果:\n"
        + os.path.basename(pszResultFileFullPath)
        + "\n\n各ファイルは入力ファイルと同じフォルダーに保存しました。"
    )


def build_result_save_warning_message(
    pszInputFileFullPath: str, objException: Exception
) -> str:
    """TSV成功・詳細結果保存失敗時の短い警告文を返します。"""
    return (
        "朝日注文エリア店舗対応TSVの作成は正常に完了しました。\n\n"
        + "ただし、詳細結果テキストを保存できませんでした。\n\n"
        + "入力ファイル:\n"
        + os.path.basename(pszInputFileFullPath)
        + "\n\n最終出力:\n"
        + "\n".join(FINAL_AREA_STORE_MAPPING_FILE_NAMES)
        + "\n\n詳細結果保存エラー:\n"
        + str(objException).splitlines()[0]
    )


def run_asahi_order_area_store_mapping_maker_cmd(
    pszInputFileFullPath: str,
) -> tuple[bool, str]:
    """同じフォルダーのCmdプログラムを実行します。"""
    pszCurrentDirectoryFullPath: str = os.path.dirname(os.path.abspath(__file__))
    pszScriptFileName: str = "AsahiOrderAreaStoreMappingMaker_Cmd.py"
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
            [sys.executable, pszScriptFileFullPath, pszInputFileFullPath],
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
        ".xlsxファイルを1つだけ、このウィンドウにドラッグ＆ドロップしてください。\n"
        "「本州マグロ(週間)」と「割り」のセル値をTSVへ変換します。\n"
        "存在する対象シートだけをTSVへ変換します。\n"
        "「割り」から割り店舗対応表を作成します。\n"
        "割り店舗対応表から店舗コード・店舗名の不一致TSVを作成します。\n"
        "既存の不一致TSVは.bk0001.tsv形式でバックアップします。\n"
        "正式店舗名.txtは必須で、正式名を反映したstep0002 TSVを作成します。\n"
        "正式店舗名.txtがない場合はstep0002_error.txtを作成します。\n"
        "step0002からAPEXの2列を除いたstep0003 TSVを作成します。\n"
        "週間店舗対応表から配送センター名を付けたstep0004 TSVを作成します。\n"
        "指定エリアを除外し、配送センター名を補完したstep0005 TSVを作成します。\n"
        "step0005からエリア名を除いた最終対応表TXT・TSVを作成します。\n"
        "「本州マグロ(週間)」から週間店舗対応表も作成します。\n"
        "対象シートがない場合は_warning.txtを出力します。\n"
        "存在しないシートの旧TSVは.bk0001.tsv形式へ名前変更します。\n"
        "既存の出力ファイルは自動的に上書きします。\n"
        "エラー時は<元ファイル名>_error.txtを出力します。"
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
            if iFileCount != 1:
                show_error_message_box(
                    "ファイルは1つだけドラッグ＆ドロップしてください。\n"
                    + "ドロップされたファイル数: "
                    + str(iFileCount),
                    WINDOW_TITLE,
                )
                return 0
            pszDroppedFilePath: str = win32api.DragQueryFile(iDropHandle, 0)
            if not os.path.isfile(pszDroppedFilePath):
                pszDetailMessage: str = (
                    "入力パスがファイルではありません。Path = "
                    + pszDroppedFilePath
                )
                report_processing_error(
                    pszDroppedFilePath,
                    "朝日注文エリア店舗対応調査TSV作成処理",
                    pszDetailMessage,
                )
                show_error_message_box(pszDetailMessage, WINDOW_TITLE)
                return 0
            if os.path.splitext(pszDroppedFilePath)[1].lower() != ".xlsx":
                pszDetailMessage = (
                    ".xlsxファイルをドラッグ＆ドロップしてください。Path = "
                    + pszDroppedFilePath
                )
                report_processing_error(
                    pszDroppedFilePath,
                    "朝日注文エリア店舗対応調査TSV作成処理",
                    pszDetailMessage,
                )
                show_error_message_box(pszDetailMessage, WINDOW_TITLE)
                return 0
            bIsSuccess, pszResultMessage = (
                run_asahi_order_area_store_mapping_maker_cmd(pszDroppedFilePath)
            )
            if bIsSuccess:
                try:
                    pszResultFileFullPath, _ = save_result_text(
                        pszDroppedFilePath, pszResultMessage
                    )
                    show_message_box(
                        build_success_summary_message(
                            pszDroppedFilePath, pszResultFileFullPath
                        ),
                        WINDOW_TITLE,
                    )
                except Exception as objException:
                    show_warning_message_box(
                        build_result_save_warning_message(
                            pszDroppedFilePath, objException
                        ),
                        WINDOW_TITLE,
                    )
            else:
                show_error_message_box(pszResultMessage, WINDOW_TITLE)
        finally:
            win32api.DragFinish(iDropHandle)
        return 0
    if iMessage == win32con.WM_PAINT:
        draw_instruction_text(iWindowHandle)
        return 0
    if iMessage == win32con.WM_DESTROY:
        win32gui.PostQuitMessage(0)
        return 0
    return win32gui.DefWindowProc(iWindowHandle, iMessage, iWparam, iLparam)


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
    iWindowHeight: int = 500
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
    pszWindowClassName: str = "AsahiOrderAreaStoreMappingMakerDndWindowClass"
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
