# -- coding: utf-8 --
###############################################################
#
# RotateImageFile_90DegreesLeft_DnD.py
#
# pip install pywin32
#
###############################################################

import os
import sys
import subprocess

import win32api
import win32con
import win32gui


###############################################################
#
# show_message_box
#
###############################################################
def show_message_box(
    pszMessage: str,
    pszTitle: str,
) -> None:
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
# run_rotate_image_file_cmd
#
###############################################################
def run_rotate_image_file_cmd(
    pszInputFileFullPath: str,
) -> tuple[bool, str]:
    pszCurrentDirectoryFullPath: str = os.path.dirname(os.path.abspath(__file__))
    pszScriptFileName: str = "RotateImageFile_90DegreesLeft_Cmd.py"
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
        show_error_message_box(
            pszErrorMessage,
            "RotateImageFile 90DegreesLeft DnD",
        )
        return False, pszErrorMessage

    pszPythonExecutableFullPath: str = sys.executable

    try:
        objCompletedProcess: subprocess.CompletedProcess[str] = subprocess.run(
            [
                pszPythonExecutableFullPath,
                pszScriptFileFullPath,
                pszInputFileFullPath,
            ],
            check=False,
            capture_output=True,
            text=True,
        )
    except Exception as objException:
        pszErrorMessage: str = (
            "Error: unexpected exception while running "
            + pszScriptFileName
            + ". Detail = "
            + str(objException)
        )
        return False, pszErrorMessage

    if objCompletedProcess.returncode != 0:
        pszStdErr: str = objCompletedProcess.stderr
        if pszStdErr.strip() == "":
            pszStdErr = "Process exited with non-zero return code and no stderr output."

        pszErrorMessage: str = (
            "Error: "
            + pszScriptFileName
            + " exited with non-zero return code.\n\n"
            + "Return code = "
            + str(objCompletedProcess.returncode)
            + "\n\n"
            + "stderr:\n"
            + pszStdErr
        )
        return False, pszErrorMessage

    pszStdOut: str = objCompletedProcess.stdout
    if pszStdOut.strip() == "":
        pszStdOut = pszScriptFileName + " finished successfully."

    return True, pszStdOut


###############################################################
#
# draw_instruction_text
#
###############################################################
def draw_instruction_text(
    iWindowHandle: int,
) -> None:
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
        "画像ファイルをこのウィンドウにドラッグ＆ドロップしてください。\n"
        "左に90度回転した画像を同じフォルダに作成します。\n"
        "エラーが発生した場合は <元ファイル名>_error.txt や output_error.txt を出力します。"
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
    if iMessage == win32con.WM_CREATE:
        win32gui.DragAcceptFiles(iWindowHandle, True)
        return 0

    if iMessage == win32con.WM_DROPFILES:
        iDropHandle: int = iWparam
        iFileCount: int = win32api.DragQueryFile(iDropHandle, -1)

        if iFileCount < 1:
            win32api.DragFinish(iDropHandle)
            show_error_message_box("Error: no files were dropped.", "RotateImageFile 90DegreesLeft DnD")
            return 0

        listSucceededFileNames: list[str] = []
        listFailedFileNames: list[str] = []

        for iFileIndex in range(iFileCount):
            pszDroppedFilePath: str = win32api.DragQueryFile(iDropHandle, iFileIndex)
            bIsSuccess, _ = run_rotate_image_file_cmd(pszDroppedFilePath)
            if bIsSuccess:
                listSucceededFileNames.append(os.path.basename(pszDroppedFilePath))
            else:
                listFailedFileNames.append(os.path.basename(pszDroppedFilePath))

        win32api.DragFinish(iDropHandle)

        pszMessage: str = (
            "完了: "
            + str(iFileCount)
            + "件中 "
            + str(len(listSucceededFileNames))
            + "件成功 / "
            + str(len(listFailedFileNames))
            + "件失敗"
        )
        if len(listFailedFileNames) > 0:
            pszMessage = pszMessage + "\n失敗: " + ", ".join(listFailedFileNames)
            show_error_message_box(pszMessage, "RotateImageFile 90DegreesLeft DnD")
        else:
            show_message_box(pszMessage, "RotateImageFile 90DegreesLeft DnD")
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
def register_window_class(
    pszWindowClassName: str,
) -> int:
    iInstanceHandle: int = win32api.GetModuleHandle(None)

    objWndClass = win32gui.WNDCLASS()
    objWndClass.hInstance = iInstanceHandle
    objWndClass.lpszClassName = pszWindowClassName
    objWndClass.lpfnWndProc = window_proc
    objWndClass.style = win32con.CS_HREDRAW | win32con.CS_VREDRAW
    objWndClass.hCursor = win32gui.LoadCursor(0, win32con.IDC_ARROW)
    objWndClass.hbrBackground = win32con.COLOR_WINDOW + 1

    iClassAtom: int = win32gui.RegisterClass(objWndClass)
    return iClassAtom


###############################################################
#
# create_main_window
#
###############################################################
def create_main_window(
    pszWindowClassName: str,
    pszWindowTitle: str,
) -> int:
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
    pszWindowClassName: str = "RotateImageFile90DegreesLeftDndWindowClass"
    pszWindowTitle: str = "RotateImageFile 90DegreesLeft (Drag & Drop)"

    try:
        register_window_class(pszWindowClassName)
    except Exception as objException:
        show_error_message_box(
            "Error: failed to register window class. Detail = " + str(objException),
            "RotateImageFile 90DegreesLeft DnD",
        )
        return

    try:
        create_main_window(pszWindowClassName, pszWindowTitle)
    except Exception as objException:
        show_error_message_box(
            "Error: failed to create main window. Detail = " + str(objException),
            "RotateImageFile 90DegreesLeft DnD",
        )
        return

    try:
        win32gui.PumpMessages()
    except Exception as objException:
        show_error_message_box(
            "Error: unexpected exception in message loop. Detail = " + str(objException),
            "RotateImageFile 90DegreesLeft DnD",
        )


if __name__ == "__main__":
    main()
