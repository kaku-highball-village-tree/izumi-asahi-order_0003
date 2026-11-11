# -- coding: utf-8 --
###############################################################
#
# AsahiSingleOrderProductCodeSelector_DnD.py
#
# pip install pywin32
#
###############################################################

import os
import re
import subprocess
import sys
import tempfile
from pathlib import Path

import win32api
import win32con
import win32gui

WINDOW_TITLE: str = (
    "Asahi Single Order Product Code Selector step0001-step0005 (Drag & Drop)"
)
CMD_FILE_NAME: str = "AsahiSingleOrderProductCodeSelector_Cmd.py"
PRODUCTS_FILE_NAME: str = "products_all_109_readable.tsv"
WEEKLY_TEMPLATE_FILE_NAME: str = "template_イズミ週間予定表_3列.xlsx"
HIROSHIMA_WEEKLY_TEMPLATE_FILE_NAME: str = (
    "template_イズミ週間予定表_2列_広島センター.xlsx"
)
OKAYAMA_SHIKOKU_WEEKLY_TEMPLATE_FILE_NAME: str = (
    "template_イズミ週間予定表_2列_岡山四国センター.xlsx"
)
FRESH_FISH_TEMPLATE_FILE_NAMES: tuple[str, str] = (
    "template_鮮魚_店別納入明細票_広島1列_岡山1列_四国1列.xlsx",
    "template_鮮魚_店別納入明細票_広島2列_岡山四国1列.xlsx",
)
AREA_STORE_MAPPING_FILE_NAME: str = "AsahiOrderAreaStoreMapping_対応表.txt"
DISPLAY_FILE_LIMIT: int = 10


def show_message_box(pszMessage: str, pszTitle: str) -> None:
    """正常な処理結果を情報アイコン付きメッセージボックスで表示します。"""
    win32gui.MessageBox(
        0, pszMessage, pszTitle, win32con.MB_OK | win32con.MB_ICONINFORMATION
    )


def show_error_message_box(pszMessage: str, pszTitle: str) -> None:
    """エラー内容をエラーアイコン付きメッセージボックスで表示します。"""
    win32gui.MessageBox(0, pszMessage, pszTitle, win32con.MB_OK | win32con.MB_ICONERROR)


def get_result_base_path(pszInputFilePath: str) -> Path:
    """入力ファイルの絶対パスから拡張子を除いた結果名の基準を返します。"""
    return Path(os.path.abspath(pszInputFilePath)).with_suffix("")


def get_result_key(pszInputFilePath: str) -> str:
    """Windowsの大文字・小文字を無視したフォルダー＋stemのキーを返します。"""
    return os.path.normcase(str(get_result_base_path(pszInputFilePath))).casefold()


def get_next_result_history_path(objResultPath: Path, pszResult: str) -> Path:
    """既存の最大連番の次の結果履歴パスを返します。"""
    pszBaseStem: str = objResultPath.stem[: -(len(pszResult) + 1)]
    objPattern: re.Pattern[str] = re.compile(
        re.escape(pszBaseStem) + "_" + re.escape(pszResult) + r"_(\d{4,})\.txt$"
    )
    iMaximumSequence: int = 0
    for objPath in objResultPath.parent.glob(pszBaseStem + "_" + pszResult + "_*.txt"):
        objMatch: re.Match[str] | None = objPattern.fullmatch(objPath.name)
        if objMatch is not None:
            iMaximumSequence = max(iMaximumSequence, int(objMatch.group(1)))
    return objResultPath.with_name(
        pszBaseStem + "_" + pszResult + "_" + f"{iMaximumSequence + 1:04d}.txt"
    )


def write_input_error_text(pszInputFilePath: str, pszErrorMessage: str) -> None:
    """過去の結果を履歴化し、DnD事前エラーを安全に保存します。"""
    objBasePath: Path = get_result_base_path(pszInputFilePath)
    objSuccessPath: Path = objBasePath.with_name(objBasePath.name + "_success.txt")
    objErrorPath: Path = objBasePath.with_name(objBasePath.name + "_error.txt")
    pszText: str = (
        "処理名:\nProductCodeSelector step0001～step0005\n\n"
        + "エラー:\n"
        + pszErrorMessage
        + "\n"
    )
    pszCrLfText: str = pszText.replace("\n", "\r\n")
    iFileDescriptor, pszTemporaryPath = tempfile.mkstemp(
        prefix="." + objErrorPath.stem + ".",
        suffix=objErrorPath.suffix,
        dir=objErrorPath.parent,
    )
    os.close(iFileDescriptor)
    objTemporaryPath: Path = Path(pszTemporaryPath)
    objTemporaryPath.write_bytes(pszCrLfText.encode("utf-8"))
    listRenames: list[tuple[Path, Path]] = []
    try:
        for objOldPath, pszOldResult in (
            (objSuccessPath, "success"),
            (objErrorPath, "error"),
        ):
            if not objOldPath.is_file():
                continue
            objHistoryPath: Path = get_next_result_history_path(
                objOldPath, pszOldResult
            )
            objOldPath.rename(objHistoryPath)
            listRenames.append((objOldPath, objHistoryPath))
        os.replace(objTemporaryPath, objErrorPath)
    except Exception:
        for objOriginalPath, objHistoryPath in reversed(listRenames):
            if objHistoryPath.exists() and not objOriginalPath.exists():
                objHistoryPath.rename(objOriginalPath)
        raise
    finally:
        if objTemporaryPath.exists():
            objTemporaryPath.unlink()


def write_dropped_file_error_texts(
    listDroppedFilePaths: list[str], pszErrorMessage: str
) -> list[str]:
    """結果名の異なるドロップ入力ごとに_error.txtを作成します。"""
    listWriteFailures: list[str] = []
    setWrittenKeys: set[str] = set()
    for pszDroppedFilePath in listDroppedFilePaths:
        pszResultKey: str = get_result_key(pszDroppedFilePath)
        if pszResultKey in setWrittenKeys:
            continue
        setWrittenKeys.add(pszResultKey)
        try:
            write_input_error_text(pszDroppedFilePath, pszErrorMessage)
        except Exception as objException:
            listWriteFailures.append(
                str(get_result_base_path(pszDroppedFilePath))
                + "_error.txt: "
                + str(objException)
            )
    return listWriteFailures


def format_limited_file_names(listFileNames: list[str]) -> str:
    """ダイアログ用に最大10件のファイル名と残りの件数を返します。"""
    listDisplayedNames: list[str] = listFileNames[:DISPLAY_FILE_LIMIT]
    pszText: str = "\n".join(listDisplayedNames)
    iRemainingCount: int = len(listFileNames) - len(listDisplayedNames)
    if iRemainingCount > 0:
        pszText += "\nほか" + str(iRemainingCount) + "件"
    return pszText


def report_dropped_files_error(
    listDroppedFilePaths: list[str], pszErrorMessage: str
) -> None:
    """事前エラーを入力別ファイルと要約ダイアログで報告します。"""
    listWriteFailures: list[str] = write_dropped_file_error_texts(
        listDroppedFilePaths, pszErrorMessage
    )
    pszDialogMessage: str = (
        pszErrorMessage
        + "\n\n処理は開始していません。"
        + "\n詳細は各入力と同じフォルダーの_error.txtを確認してください。"
    )
    if listWriteFailures:
        pszDialogMessage += (
            "\n\n_error.txt保存失敗: " + str(len(listWriteFailures)) + "件"
        )
    show_error_message_box(pszDialogMessage, WINDOW_TITLE)


def run_product_code_selector_cmd(
    pszInputFileFullPath: str,
) -> tuple[str, str]:
    """同じフォルダーのCmdプログラムを実行します。"""
    pszCurrentDirectoryFullPath: str = os.path.dirname(os.path.abspath(__file__))
    pszScriptFileFullPath: str = os.path.join(
        pszCurrentDirectoryFullPath, CMD_FILE_NAME
    )
    if not os.path.isfile(pszScriptFileFullPath):
        return (
            "failed",
            "Error: " + CMD_FILE_NAME + " not found. Path = " + pszScriptFileFullPath,
        )
    dictEnvironment: dict[str, str] = os.environ.copy()
    dictEnvironment["PYTHONIOENCODING"] = "utf-8"
    try:
        objCompletedProcess: subprocess.CompletedProcess[str] = subprocess.run(
            [sys.executable, pszScriptFileFullPath, pszInputFileFullPath],
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            env=dictEnvironment,
        )
    except Exception as objException:
        return (
            "failed",
            "Error: unexpected exception while running "
            + CMD_FILE_NAME
            + ". Detail = "
            + str(objException),
        )
    if objCompletedProcess.returncode == 2:
        return "cancelled", objCompletedProcess.stderr
    if objCompletedProcess.returncode == 3:
        return "not_found", objCompletedProcess.stderr
    if objCompletedProcess.returncode != 0:
        pszStdErr: str = objCompletedProcess.stderr
        if pszStdErr.strip() == "":
            pszStdErr = "Process exited with non-zero return code and no stderr output."
        return "failed", pszStdErr
    pszStdOut: str = objCompletedProcess.stdout
    if pszStdOut.strip() == "":
        pszStdOut = CMD_FILE_NAME + " finished successfully."
    return "success", pszStdOut


def draw_instruction_text(iWindowHandle: int) -> None:
    """ドラッグ＆ドロップするファイルと出力内容を描画します。"""
    iDeviceContext, objPaintStruct = win32gui.BeginPaint(iWindowHandle)
    try:
        objClientRect = win32gui.GetClientRect(iWindowHandle)
        iMargin: int = 8
        objClientRect = (
            objClientRect[0] + iMargin,
            objClientRect[1] + iMargin,
            objClientRect[2] - iMargin,
            objClientRect[3] - iMargin,
        )
        pszInstructionText: str = (
            "商品別step0007ファイル（_step0007_なしの名前も可）を\n"
            "このウィンドウにドラッグ＆ドロップしてください。\n\n"
            "対応形式はXLSXとTSVです。\n\n"
            "旧形式（Ｐ品番・APEX品番）と新形式（APEX品番のみ）に対応し、\n"
            "出力はAPEX品番のみの新形式に統一します。\n\n"
            "APEX品番を空欄にして、\n"
            "ProductCodeSelector step0001のXLSXとTSVを作成します。\n\n"
            "products_all_109_readable.tsvのA～C列から\n"
            "products_all_109_readable_ABC.tsvを自動作成します。\n"
            "続いてABC版から同じ魚介カテゴリを中心に関連候補を表示します。\n"
            "完全一致以外の同じカテゴリの商品や全商品も確認でき、\n"
            "最終的な商品は担当者が検索・選択して確定します。\n"
            "選択結果を設定したstep0002のXLSXとTSVを作成します。\n"
            "step0002の両ファイルを再読込した後、\n"
            "N1から最終店舗列の9行を転置し、\n"
            "全店舗・広島・岡山・四国の店舗別TSVを作成します。\n"
            "既存の店舗別TSVは%TEMP%へコピー後、\n"
            "元フォルダーで最終更新日時付きの名前へ変更します。\n"
            "template_イズミ週間予定表_3列.xlsxの作成日を更新し、\n"
            "step0002の納品日とその前日の出荷日を3地区へ設定し、\n"
            "8行目の合計と9行目の曜日別小計を空欄にしたstep0003のXLSXと\n"
            "A1:AB44のTSVを作成します。\n"
            "step0002 TSVのF3の商品名を、3列版はB5・K5、4列版はB5・U5へ\n"
            "XLSX・TSVの両方に設定し、step0004へ引き継ぎます。\n"
            "参照元の商品名の読込失敗・空欄時はstep0003以降を作成しません。\n"
            "step0003の週間予定表とエリア別TSVを再読込し、\n"
            "広島・岡山・四国の店舗別データを各30店舗を上限に転記した\n"
            "step0004のXLSXとTSVを作成して、8行目へ各地区の合計を\n"
            "D8・M8・V8に1つずつ、9行目へ月～日の小計を設定します。\n"
            "合計・小計ラベルはC8・L8・U8とC9・L9・U9です。\n"
            "広島が31～120店舗の場合は、\n"
            "template_イズミ週間予定表_4列_広島広島岡山四国センター.xlsxから\n"
            "広島1～60・広島61～120・岡山・四国の4ブロックへ転記した\n"
            "A1:AL73のstep0003・step0004 XLSX／TSVを1組作成します。\n"
            "4列版では岡山・四国は各60店舗まで、作成日はQ1・AJ1です。\n"
            "店舗欄は14～73行目、8行目に各ブロックの週間合計、\n"
            "9行目に月～日の小計を設定します。\n"
            "出荷日・曜日は10・11行目、納品日・曜日は12・13行目です。\n"
            "広島が121店舗以上の場合は_error.txtを出力して終了します。\n\n"
            "通常版または4列版step0004 TSVの全日付と数量を検証し、\n"
            "発注数量を目方重量1と0.5へ分解して、注文あり店舗ごとに\n"
            "2行で転記した鮮魚店別納入明細票を作成します。\n"
            "広島0～20店舗は1列版、21～40店舗は2列版を使用し、\n"
            "41～60店舗は両方とも2列版の2帳票へ分割します。\n"
            "各曜日の広島注文あり店舗が61店舗以上なら_error.txtを出力します。\n"
            "複数帳票では岡山・四国と全帳票分の箱数集計を1枚目だけに\n"
            "表示し、2枚目の箱数集計欄を空欄にします。\n"
            "1列版は岡山・四国各20店舗、2列版は両地区合計20店舗を\n"
            "上限とし、容量超過時は_error.txtを出力します。\n"
            "容量超過時はstep0005を更新せず、作成済みstep0001～step0004と\n"
            "過去のstep0005を保持します。\n"
            "J5へ出荷日、J6へ納品日を設定したstep0005 XLSX・\n"
            "A1:M60 TSVを月～日分作成します。\n\n"
            "各日・全ページのXLSX・TSVのK9へ、step0002 TSVのF3の商品名を設定します。\n"
            "参照元の読込失敗・商品名空欄時はstep0005を更新しません。\n\n"
            "step0005は「店別納入明細票」シートだけを処理し、\n"
            "テンプレート内のテキストボックス・図形とその他のシートは\n"
            "変更せずに保持します。\n\n"
            "出力ファイルは入力ファイルと同じフォルダーに作成します。\n"
            "既存の出力ファイルは自動的に上書きします。\n"
            "エラー時は_error.txtを出力します。\n"
            "step0005のテンプレートエラー時は差し込み元ファイル名を\n"
            "_error.txtに記録します。"
        )
        win32gui.DrawText(
            iDeviceContext,
            pszInstructionText,
            -1,
            objClientRect,
            win32con.DT_LEFT | win32con.DT_TOP | win32con.DT_WORDBREAK,
        )
    finally:
        win32gui.EndPaint(iWindowHandle, objPaintStruct)


def window_proc(iWindowHandle: int, iMessage: int, iWparam: int, iLparam: int) -> int:
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
            dictInputGroups: dict[str, list[str]] = {}
            for pszDroppedFilePath in listDroppedFilePaths:
                dictInputGroups.setdefault(
                    get_result_key(pszDroppedFilePath), []
                ).append(pszDroppedFilePath)
            listDuplicateGroups: list[list[str]] = [
                listGroupPaths
                for listGroupPaths in dictInputGroups.values()
                if len(listGroupPaths) > 1
            ]
            if listDuplicateGroups:
                listWriteFailures: list[str] = []
                listDuplicatePaths: list[str] = [
                    pszPath
                    for listGroupPaths in listDuplicateGroups
                    for pszPath in listGroupPaths
                ]
                pszDuplicateSummary: str = "\n".join(
                    os.path.basename(pszPath) for pszPath in listDuplicatePaths
                )
                for listGroupPaths in dictInputGroups.values():
                    if len(listGroupPaths) > 1:
                        pszErrorMessage: str = (
                            "同じ結果ファイル名になる入力が複数指定されています。\n\n"
                            + "入力ファイル:\n"
                            + "\n".join(
                                os.path.abspath(pszPath) for pszPath in listGroupPaths
                            )
                            + "\n\n同じフォルダー・同じファイル名本体の"
                            + "XLSXとTSVは同時処理できません。\n"
                            + "どちらか一方をドラッグ＆ドロップしてください。\n\n"
                            + "重複入力があったため、バッチ全体の処理を開始していません。"
                        )
                    else:
                        pszErrorMessage = (
                            "同時にドラッグ＆ドロップされた別の入力に"
                            + "重複があったため、この入力についても処理を"
                            + "開始していません。\n\n入力ファイル:\n"
                            + os.path.abspath(listGroupPaths[0])
                            + "\n\n重複していた入力:\n"
                            + pszDuplicateSummary
                        )
                    try:
                        write_input_error_text(listGroupPaths[0], pszErrorMessage)
                    except Exception as objException:
                        listWriteFailures.append(str(objException))
                listDuplicateNames: list[str] = [
                    os.path.basename(pszPath) for pszPath in listDuplicatePaths
                ]
                show_error_message_box(
                    "入力ファイルが重複しています。\n\n"
                    + format_limited_file_names(listDuplicateNames)
                    + "\n\n同じ結果ファイル名になるため、同時処理できません。"
                    + "\nどちらか一方をドラッグ＆ドロップしてください。"
                    + "\n\nドロップされた全入力の処理は開始していません。"
                    + "\n詳細は各入力と同じフォルダーの_error.txtを確認してください。"
                    + (
                        "\n\n_error.txt保存失敗: " + str(len(listWriteFailures)) + "件"
                        if listWriteFailures
                        else ""
                    ),
                    WINDOW_TITLE,
                )
                return 0
            pszProgramDirectory: str = os.path.dirname(os.path.abspath(__file__))
            pszCmdPath: str = os.path.join(pszProgramDirectory, CMD_FILE_NAME)
            if not os.path.isfile(pszCmdPath):
                report_dropped_files_error(
                    listDroppedFilePaths,
                    CMD_FILE_NAME
                    + " が見つかりません。プログラムと同じフォルダーに配置してください。",
                )
                return 0
            pszProductsPath: str = os.path.join(pszProgramDirectory, PRODUCTS_FILE_NAME)
            if not os.path.isfile(pszProductsPath):
                report_dropped_files_error(
                    listDroppedFilePaths,
                    PRODUCTS_FILE_NAME
                    + " が見つかりません。プログラムと同じフォルダーに配置してください。",
                )
                return 0
            pszMappingPath: str = os.path.join(
                pszProgramDirectory, AREA_STORE_MAPPING_FILE_NAME
            )
            if not os.path.isfile(pszMappingPath):
                pszErrorMessage: str = (
                    AREA_STORE_MAPPING_FILE_NAME
                    + " が見つかりません。"
                    + "プログラムと同じフォルダーに配置してください。"
                )
                report_dropped_files_error(listDroppedFilePaths, pszErrorMessage)
                return 0
            listFailedFileNames: list[str] = []
            listCancelledFileNames: list[str] = []
            listNotFoundFileNames: list[str] = []
            listSuccessDetails: list[str] = []
            iSuccessCount: int = 0
            for pszDroppedFilePath in listDroppedFilePaths:
                pszResult, pszResultMessage = run_product_code_selector_cmd(
                    pszDroppedFilePath
                )
                if pszResult == "success":
                    iSuccessCount += 1
                    listSuccessDetails.append(pszResultMessage.strip())
                elif pszResult == "cancelled":
                    listCancelledFileNames.append(os.path.basename(pszDroppedFilePath))
                elif pszResult == "not_found":
                    listNotFoundFileNames.append(os.path.basename(pszDroppedFilePath))
                else:
                    objErrorPath: Path = get_result_base_path(
                        pszDroppedFilePath
                    ).with_name(
                        get_result_base_path(pszDroppedFilePath).name + "_error.txt"
                    )
                    if not objErrorPath.is_file():
                        try:
                            write_input_error_text(
                                pszDroppedFilePath, pszResultMessage.strip()
                            )
                        except Exception:
                            pass
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
            if listCancelledFileNames:
                pszMessage += " / " + str(len(listCancelledFileNames)) + "件キャンセル"
            if listNotFoundFileNames:
                pszMessage += " / " + str(len(listNotFoundFileNames)) + "件該当商品なし"
            if listFailedFileNames:
                pszMessage += "\n\n失敗:\n" + format_limited_file_names(
                    listFailedFileNames
                )
                if listCancelledFileNames:
                    pszMessage += "\n\nキャンセル:\n" + format_limited_file_names(
                        listCancelledFileNames
                    )
                if listNotFoundFileNames:
                    pszMessage += "\n\n該当商品なし:\n" + format_limited_file_names(
                        listNotFoundFileNames
                    )
                pszMessage += (
                    "\n\n詳細は各入力と同じフォルダーの"
                    + "_success.txtまたは_error.txtを確認してください。"
                )
                show_error_message_box(pszMessage, WINDOW_TITLE)
            elif listCancelledFileNames or listNotFoundFileNames:
                if listCancelledFileNames:
                    pszMessage += "\n\nキャンセル:\n" + format_limited_file_names(
                        listCancelledFileNames
                    )
                if listNotFoundFileNames:
                    pszMessage += "\n\n該当商品なし:\n" + format_limited_file_names(
                        listNotFoundFileNames
                    )
                show_message_box(pszMessage, WINDOW_TITLE)
            else:
                pszMessage += (
                    "\n\nProductCodeSelector step0001～step0005を作成しました。"
                    + "\n\n作成内容:"
                    + "\n・step0001 XLSX・TSV"
                    + "\n・step0002 XLSX・TSV"
                    + "\n・step0003 週間予定表XLSX・TSV"
                    + "\n・step0003 店舗別TSV 4ファイル"
                    + "\n・step0004 店舗別週間予定表XLSX・TSV"
                    + "\n・step0005 鮮魚店別納入明細票XLSX・TSV（月～日）"
                    + "\n\n詳細は各入力と同じフォルダーの"
                    + "_success.txtを確認してください。"
                )
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
    return win32gui.DefWindowProc(iWindowHandle, iMessage, iWparam, iLparam)


def register_window_class(pszWindowClassName: str) -> int:
    """DnDウィンドウクラスを登録します。"""
    objWindowClass = win32gui.WNDCLASS()
    objWindowClass.hInstance = win32api.GetModuleHandle(None)
    objWindowClass.lpszClassName = pszWindowClassName
    objWindowClass.lpfnWndProc = window_proc
    objWindowClass.hCursor = win32gui.LoadCursor(0, win32con.IDC_ARROW)
    objWindowClass.hbrBackground = win32con.COLOR_WINDOW + 1
    return win32gui.RegisterClass(objWindowClass)


def create_main_window(pszWindowClassName: str, pszWindowTitle: str) -> int:
    """ドラッグ＆ドロップを受け付けるメインウィンドウを作成します。"""
    iWindowHandle: int = win32gui.CreateWindowEx(
        win32con.WS_EX_ACCEPTFILES,
        pszWindowClassName,
        pszWindowTitle,
        win32con.WS_OVERLAPPEDWINDOW | win32con.WS_VISIBLE,
        win32con.CW_USEDEFAULT,
        win32con.CW_USEDEFAULT,
        680,
        390,
        0,
        0,
        win32api.GetModuleHandle(None),
        None,
    )
    win32gui.DragAcceptFiles(iWindowHandle, True)
    return iWindowHandle


def main() -> None:
    """DnDウィンドウを作成してWindowsメッセージループを開始します。"""
    pszWindowClassName: str = "AsahiSingleOrderProductCodeSelectorDndWindowClass"
    try:
        register_window_class(pszWindowClassName)
        create_main_window(pszWindowClassName, WINDOW_TITLE)
        win32gui.PumpMessages()
    except Exception as objException:
        show_error_message_box(
            "Error: failed to create the drag-and-drop window. Detail = "
            + str(objException),
            WINDOW_TITLE,
        )


if __name__ == "__main__":
    main()
