# -- coding: utf-8 --
###############################################################
# AsahiSingleOrderQuantityChangeHighlighter_Cmd.py
# pip install openpyxl
###############################################################
"""商品別週間発注Excelの数量変更を赤字化します。

--prepareで編集前履歴を確保し、編集・保存・終了後に通常モードで比較します。
空欄と0は同数量、非0から空欄への変更は赤字の0に補完します。
赤字は累積し、過去の履歴は削除しません。保存失敗時の自動復旧はしません。
"""

import csv
import io
import json
import os
import re
import sys
import tempfile
import uuid
from contextlib import contextmanager
from copy import copy
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.cell.cell import Cell
from openpyxl.styles import Color
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

EXPECTED_HEADERS = (
    "納品日", "曜日", "配送パターン", "便", "APEX品番", "商品名", "産地",
    "仕様", "伝票原価", "伝票売価", "値入", "売価", "単位",
)
WEEKDAYS = ("月", "火", "水", "木", "金", "土", "日")
HISTORY_HEADERS = (
    "formatVersion", "apexProductCode", "deliveryDate", "storeCode", "quantity", "isRed",
)
FORMAT_VERSION = "1"
RED_FONT_COLOR = "FFFF0000"
MAX_HISTORY_NUMBER = 9999


@dataclass
class OrderRow:
    """日付・店舗ごとの数量と累積赤字状態。Excelの場合は実セルも保持。"""

    product_code: str
    delivery_date: date
    store_code: str
    quantity: object
    is_red: bool
    quantity_cell: Cell | None = None

    @property
    def key(self) -> tuple[date, str]:
        return self.delivery_date, self.store_code


def write_error_text(
    pszOutputFileFullPath: str,
    pszErrorMessage: str,
) -> None:
    """エラーメッセージをUTF-8のテキストファイルへ上書き保存します。

    Args:
        pszOutputFileFullPath: エラーファイルの保存先です。
        pszErrorMessage: 利用者へ知らせるエラー内容です。

    Returns:
        戻り値はありません。
    """
    pszDirectoryFullPath: str = os.path.dirname(pszOutputFileFullPath)
    if pszDirectoryFullPath != "":
        os.makedirs(pszDirectoryFullPath, exist_ok=True)
    with open(pszOutputFileFullPath, mode="w", encoding="utf-8", newline="") as objFile:
        objFile.write(pszErrorMessage.rstrip("\n") + "\n")


###############################################################
#
# get_error_file_full_path
#
###############################################################
def get_error_file_full_path(pszInputFileFullPath: str) -> str:
    """入力Excelと同じフォルダーに作る_error.txtのパスを返します。"""
    pszDirectoryFullPath: str = os.path.dirname(os.path.abspath(pszInputFileFullPath))
    pszBaseNameWithoutExtension: str = os.path.splitext(
        os.path.basename(pszInputFileFullPath)
    )[0]
    return os.path.join(
        pszDirectoryFullPath,
        pszBaseNameWithoutExtension + "_error.txt",
    )


###############################################################
#
# report_processing_error
#
###############################################################
def report_processing_error(
    pszInputFileFullPath: str,
    pszProcessName: str,
    pszDetailMessage: str,
) -> None:
    """標準エラーと入力Excel用_error.txtへ同じエラーを出力します。"""
    pszErrorMessage: str = (
        "処理結果: エラー\n"
        + "入力Excel: "
        + os.path.abspath(pszInputFileFullPath)
        + "\n"
        + "発生した処理: "
        + pszProcessName
        + "\n"
        + "エラー内容: "
        + pszDetailMessage
        + "\n"
    )
    print(pszErrorMessage, file=sys.stderr, end="")
    try:
        write_error_text(get_error_file_full_path(pszInputFileFullPath), pszErrorMessage)
    except OSError as objException:
        print(
            "Error: _error.txtの保存にも失敗しました。Detail = " + str(objException),
            file=sys.stderr,
        )


###############################################################
#
# remove_old_error_file
#
###############################################################
def remove_old_error_file(pszInputFileFullPath: str) -> None:
    """正常終了後、以前の処理で作られた_error.txtがあれば削除します。"""
    pszErrorFileFullPath: str = get_error_file_full_path(pszInputFileFullPath)
    if os.path.exists(pszErrorFileFullPath):
        os.remove(pszErrorFileFullPath)


###############################################################
#
# validate_excel_path
#
###############################################################
def validate_excel_path(pszInputFileFullPath: str) -> str:
    """入力パスが存在する.xlsxファイルであることを確認します。

    Returns:
        検証後の絶対パスを返します。

    Raises:
        ValueError: パス、ファイル種別、拡張子に問題がある場合です。
    """
    pszAbsolutePath: str = os.path.abspath(pszInputFileFullPath)
    if not os.path.exists(pszAbsolutePath):
        raise ValueError("入力Excelファイルが見つかりません。Path = " + pszAbsolutePath)
    if not os.path.isfile(pszAbsolutePath):
        raise ValueError("入力パスがファイルではありません。Path = " + pszAbsolutePath)
    if os.path.splitext(pszAbsolutePath)[1].lower() != ".xlsx":
        raise ValueError("入力ファイルの拡張子は.xlsxではありません。Path = " + pszAbsolutePath)
    return pszAbsolutePath


###############################################################
#
# load_excel
#
###############################################################
def load_excel(pszInputFileFullPath: str) -> Workbook:
    """openpyxlでExcelを読み込み、既存書式を含むWorkbookを返します。"""
    return load_workbook(pszInputFileFullPath)


def normalize_text_cell(objValue: object) -> str:
    return "" if objValue is None else str(objValue).strip()


def normalize_store_code(objValue: object) -> str:
    """Excelの数値型・文字列型の店舗コードを同じ整数文字列にそろえます。"""
    pszCode = normalize_text_cell(objValue)
    if not re.fullmatch(r"[0-9]+(?:\.0+)?", pszCode):
        raise ValueError("店舗コードが空欄または不正です。Value = " + pszCode)
    return str(int(pszCode.split(".")[0]))


def normalize_quantity_for_comparison(objValue: object) -> Decimal:
    """空欄と0は同じ。非負の有限数だけを許可します。"""
    pszValue = normalize_text_cell(objValue)
    if pszValue == "":
        return Decimal(0)
    if isinstance(objValue, bool) or pszValue.startswith("="):
        raise ValueError("発注数量に数式・真偽値は使用できません。Value = " + pszValue)
    try:
        objQuantity = Decimal(pszValue)
    except (InvalidOperation, ValueError) as objException:
        raise ValueError("発注数量を数値として解釈できません。Value = " + pszValue) from objException
    if not objQuantity.is_finite() or objQuantity < 0:
        raise ValueError("発注数量は0以上の有限数にしてください。Value = " + pszValue)
    return objQuantity


def normalize_delivery_date(objValue: object) -> date:
    if isinstance(objValue, datetime):
        return objValue.date()
    if isinstance(objValue, date):
        return objValue
    try:
        return datetime.strptime(normalize_text_cell(objValue), "%Y/%m/%d").date()
    except ValueError as objException:
        raise ValueError("納品日が不正です。Value = " + str(objValue)) from objException


def is_red_cell(objCell: Cell) -> bool:
    """直接指定された赤文字を判定します（標準のindexed赤も含みます）。"""
    objColor = objCell.font.color
    if objColor is None:
        return False
    if objColor.type == "rgb":
        return str(objColor.rgb).upper()[-6:] == "FF0000"
    return objColor.type == "indexed" and objColor.indexed in (2, 10)


def find_target_worksheet(objWorkbook: Workbook) -> Worksheet:
    if len(objWorkbook.worksheets) != 1:
        raise ValueError("入力XLSXのワークシート数は1枚にしてください。")
    objWorksheet = objWorkbook.worksheets[0]
    if objWorksheet.max_row != 9 or objWorksheet.max_column < 14:
        raise ValueError("入力はヘッダー2行・月～日7行、N列以降に店舗がある形式にしてください。")
    if any(normalize_text_cell(objWorksheet.cell(1, i).value) for i in range(1, 14)):
        raise ValueError("A1～M1は空欄にしてください。")
    if tuple(normalize_text_cell(objWorksheet.cell(2, i).value) for i in range(1, 14)) != EXPECTED_HEADERS:
        raise ValueError("A2～M2の項目名が仕様と一致しません。")
    return objWorksheet


def read_excel_rows(objWorksheet: Worksheet) -> list[OrderRow]:
    pszProductCode = normalize_text_cell(objWorksheet.cell(3, 5).value)
    if not pszProductCode or objWorksheet.cell(3, 5).data_type == "f":
        raise ValueError("E3のAPEX品番が空欄または数式です。")
    if not normalize_text_cell(objWorksheet.cell(3, 6).value):
        raise ValueError("F3の商品名が空欄です。")
    objMonday = normalize_delivery_date(objWorksheet.cell(3, 1).value)
    if objMonday.weekday() != 0:
        raise ValueError("開始納品日は月曜日にしてください。")
    for iDay, pszWeekday in enumerate(WEEKDAYS):
        if normalize_delivery_date(objWorksheet.cell(iDay + 3, 1).value) != objMonday + timedelta(days=iDay):
            raise ValueError("納品日は月～日の連続した7日間にしてください。")
        if normalize_text_cell(objWorksheet.cell(iDay + 3, 2).value) != pszWeekday:
            raise ValueError("納品日と曜日が一致しません。Row = " + str(iDay + 3))

    listOrderRows = []
    setStoreCodes = set()
    for iColumn in range(14, objWorksheet.max_column + 1):
        pszStoreCode = normalize_store_code(objWorksheet.cell(1, iColumn).value)
        if pszStoreCode in setStoreCodes:
            raise ValueError("店舗コードが重複しています。Code = " + pszStoreCode)
        setStoreCodes.add(pszStoreCode)
        if not normalize_text_cell(objWorksheet.cell(2, iColumn).value):
            raise ValueError("店舗名が空欄です。Column = " + str(iColumn))
        for iDay in range(7):
            objCell = objWorksheet.cell(iDay + 3, iColumn)
            try:
                if objCell.data_type in ("f", "e"):
                    raise ValueError("数式またはExcelエラー値は使用できません。")
                normalize_quantity_for_comparison(objCell.value)
            except ValueError as objException:
                raise ValueError(objCell.coordinate + ": " + str(objException)) from objException
            listOrderRows.append(OrderRow(
                pszProductCode, objMonday + timedelta(days=iDay), pszStoreCode,
                objCell.value, is_red_cell(objCell), objCell,
            ))
    return listOrderRows


def build_history_patterns(pszExcelBaseName: str) -> tuple[re.Pattern[str], re.Pattern[str]]:
    """対象Excel専用の履歴TSVとlatest TSVを判定する正規表現を返します。"""
    pszEscapedBaseName: str = re.escape(pszExcelBaseName)
    objHistoryPattern: re.Pattern[str] = re.compile(
        r"^" + pszEscapedBaseName + r"_(\d{4})\.tsv$"
    )
    objLatestPattern: re.Pattern[str] = re.compile(
        r"^" + pszEscapedBaseName + r"_(\d{4})_latest\.tsv$"
    )
    return objHistoryPattern, objLatestPattern


###############################################################
#
# find_history_files
#
###############################################################
def find_history_files(
    pszInputFileFullPath: str,
) -> tuple[dict[int, Path], dict[int, Path]]:
    """Excelと同じフォルダーから、対応する履歴とlatestを検索します。"""
    objExcelPath: Path = Path(pszInputFileFullPath)
    objHistoryPattern, objLatestPattern = build_history_patterns(objExcelPath.stem)
    dictHistoryFiles: dict[int, Path] = {}
    dictLatestFiles: dict[int, Path] = {}

    for objPath in objExcelPath.parent.iterdir():
        if not objPath.is_file():
            continue
        objHistoryMatch = objHistoryPattern.fullmatch(objPath.name)
        if objHistoryMatch is not None:
            dictHistoryFiles[int(objHistoryMatch.group(1))] = objPath
            continue
        objLatestMatch = objLatestPattern.fullmatch(objPath.name)
        if objLatestMatch is not None:
            dictLatestFiles[int(objLatestMatch.group(1))] = objPath
    return dictHistoryFiles, dictLatestFiles


def validate_history_state(
    dictHistoryFiles: dict[int, Path], dictLatestFiles: dict[int, Path],
) -> tuple[bool, Path | None, int]:
    """欠番・旧形式・数量や赤字履歴の破損も検証し、自動修復しません。"""
    if not dictHistoryFiles and not dictLatestFiles:
        return True, None, 1
    if len(dictLatestFiles) != 1 or not dictHistoryFiles:
        raise ValueError("履歴と比較用latestが不整合です。latestは1つ必要です。")
    iMaximum = max(dictHistoryFiles)
    if set(dictHistoryFiles) != set(range(1, iMaximum + 1)):
        raise ValueError("履歴番号に欠番または不正な番号があります。")
    if set(dictLatestFiles) != {iMaximum}:
        raise ValueError("latestの番号が最新履歴番号と一致しません。")
    objLatestPath = dictLatestFiles[iMaximum]
    if dictHistoryFiles[iMaximum].read_bytes() != objLatestPath.read_bytes():
        raise ValueError("同じ番号の履歴TSVとlatest TSVの内容が異なります。")
    listPreviousRows = None
    for _, objHistoryPath in sorted(dictHistoryFiles.items()):
        listRows = load_previous_tsv(objHistoryPath)
        if listPreviousRows is not None:
            validate_matching_orders(listRows, listPreviousRows)
            dictPrevious = {row.key: row for row in listPreviousRows}
            for row in listRows:
                previous = dictPrevious[row.key]
                if (previous.is_red or normalize_quantity_for_comparison(previous.quantity)
                        != normalize_quantity_for_comparison(row.quantity)) and not row.is_red:
                    raise ValueError("履歴の累積赤字が不整合です。File = " + str(objHistoryPath))
        listPreviousRows = listRows
    if iMaximum >= MAX_HISTORY_NUMBER:
        raise ValueError("履歴番号が上限9999に達しています。")
    return False, objLatestPath, iMaximum + 1


def load_previous_tsv(objTsvPath: Path) -> list[OrderRow]:
    with objTsvPath.open("r", encoding="utf-8-sig", newline="") as objFile:
        listRows = list(csv.reader(objFile, delimiter="\t", strict=True))
    if not listRows or tuple(listRows[0]) != HISTORY_HEADERS:
        raise ValueError("履歴の形式が対応していません。File = " + str(objTsvPath))
    listOrderRows = []
    setKeys = set()
    for listRow in listRows[1:]:
        if len(listRow) != len(HISTORY_HEADERS) or listRow[0] != FORMAT_VERSION:
            raise ValueError("履歴の列数またはバージョンが不正です。File = " + str(objTsvPath))
        _, pszProductCode, pszDate, pszStoreCode, pszQuantity, pszIsRed = listRow
        if not pszProductCode.strip() or pszIsRed not in ("0", "1"):
            raise ValueError("履歴の商品コードまたは赤字情報が不正です。")
        try:
            objDate = date.fromisoformat(pszDate)
        except ValueError as objException:
            raise ValueError("履歴の納品日が不正です。") from objException
        normalize_quantity_for_comparison(pszQuantity)
        objRow = OrderRow(pszProductCode, objDate, normalize_store_code(pszStoreCode), pszQuantity, pszIsRed == "1")
        if objRow.key in setKeys:
            raise ValueError("履歴の日付・店舗コードが重複しています。")
        setKeys.add(objRow.key)
        listOrderRows.append(objRow)
    setProducts = {row.product_code for row in listOrderRows}
    listDates = sorted({row.delivery_date for row in listOrderRows})
    setStores = {row.store_code for row in listOrderRows}
    if len(setProducts) != 1 or len(listDates) != 7 or not setStores:
        raise ValueError("履歴は1商品・7日分・1店舗以上が必要です。")
    if listDates[0].weekday() != 0 or listDates != [listDates[0] + timedelta(days=i) for i in range(7)]:
        raise ValueError("履歴の納品日は月～日の連続した7日間が必要です。")
    if len(listOrderRows) != 7 * len(setStores):
        raise ValueError("履歴に不足している日付・店舗の数量があります。")
    return listOrderRows


def validate_matching_orders(listCurrentRows: list[OrderRow], listPreviousRows: list[OrderRow]) -> None:
    if ({row.product_code for row in listCurrentRows} != {row.product_code for row in listPreviousRows}
            or {row.key for row in listCurrentRows} != {row.key for row in listPreviousRows}):
        raise ValueError(
            "商品・対象日付・店舗構成が履歴と一致しません。"
            "別の入力ファイル名で新しい履歴を開始してください。"
        )


def compare_quantity(listCurrentRows: list[OrderRow], listPreviousRows: list[OrderRow]) -> list[Cell]:
    """同数量の空欄/0は触らず、非0から空欄だけ0補完。赤字を累積します。"""
    validate_matching_orders(listCurrentRows, listPreviousRows)
    dictPrevious = {row.key: row for row in listPreviousRows}
    listChangedCells = []
    for row in listCurrentRows:
        previous = dictPrevious[row.key]
        bChanged = normalize_quantity_for_comparison(row.quantity) != normalize_quantity_for_comparison(previous.quantity)
        if row.quantity_cell is None:
            raise ValueError("発注数量セルを取得できません。")
        if bChanged:
            if normalize_text_cell(row.quantity) == "":
                row.quantity = 0
                row.quantity_cell.value = 0
            listChangedCells.append(row.quantity_cell)
        row.is_red = row.is_red or previous.is_red or bChanged
        if row.is_red:
            objFont = copy(row.quantity_cell.font)
            objFont.color = Color(rgb=RED_FONT_COLOR)
            row.quantity_cell.font = objFont
    return listChangedCells


def serialize_tsv(listOrderRows: list[OrderRow]) -> str:
    objBuffer = io.StringIO(newline="")
    objWriter = csv.writer(objBuffer, delimiter="\t", lineterminator="\n")
    objWriter.writerow(HISTORY_HEADERS)
    for row in sorted(listOrderRows, key=lambda row: (row.delivery_date, int(row.store_code))):
        objWriter.writerow([
            FORMAT_VERSION, row.product_code, row.delivery_date.isoformat(), row.store_code,
            normalize_text_cell(row.quantity), "1" if row.is_red else "0",
        ])
    return objBuffer.getvalue()


def build_snapshot_paths(
    pszInputFileFullPath: str,
    iHistoryNumber: int,
) -> tuple[Path, Path]:
    """指定履歴番号の履歴TSVとlatest TSVの保存先を返します。"""
    objExcelPath: Path = Path(pszInputFileFullPath)
    pszNumber: str = f"{iHistoryNumber:04d}"
    objHistoryPath: Path = objExcelPath.with_name(
        objExcelPath.stem + "_" + pszNumber + ".tsv"
    )
    objLatestPath: Path = objExcelPath.with_name(
        objExcelPath.stem + "_" + pszNumber + "_latest.tsv"
    )
    return objHistoryPath, objLatestPath


def pending_path(objExcelPath: Path) -> Path:
    return objExcelPath.with_name(objExcelPath.stem + "_highlighter_pending.json")


@contextmanager
def processing_lock(objExcelPath: Path):
    """二重実行を防ぎ、未完了保存は手動復旧まで停止します。"""
    objLockPath = objExcelPath.with_name(objExcelPath.stem + "_highlighter.lock")
    try:
        objLock = objLockPath.open("x", encoding="utf-8")
    except FileExistsError as objException:
        raise ValueError(
            "同じ入力の処理中、または前回処理が異常終了しています。Lock = "
            + str(objLockPath)
            + "\n他の処理が終了していることと未完了保存の有無を確認してください。"
            + "\n未完了保存の記録: " + str(pending_path(objExcelPath))
            + "\n復旧後、他の処理が動作していない場合だけロックを削除してください。"
        ) from objException
    try:
        with objLock:
            objLock.write(str(os.getpid()))
        objPendingPath = pending_path(objExcelPath)
        if objPendingPath.exists():
            raise ValueError(
                "前回の保存が未完了です。自動復旧は行いません。"
                + "\n記録ファイル内のbackupDirectoryにあるrecovery.txtを確認してください。"
                + "\n記録ファイル: " + str(objPendingPath)
            )
        if objExcelPath.with_name("~$" + objExcelPath.name).exists():
            raise ValueError("入力Excelが開かれています。保存してExcelを閉じてください。")
        yield
    finally:
        objLockPath.unlink()


def write_exclusive_bytes(objPath: Path, bytesContent: bytes) -> None:
    """既存ファイルを上書きせず作成し、書込みを完了します。"""
    with objPath.open("xb") as objFile:
        objFile.write(bytesContent)
        objFile.flush()
        os.fsync(objFile.fileno())


def save_output_set(
    objExcelPath: Path,
    bytesOriginalExcel: bytes,
    objWorkbook: Workbook,
    listRows: list[OrderRow],
    iHistoryNumber: int,
    objPreviousLatestPath: Path | None,
) -> tuple[Path, Path, Path]:
    """バックアップ・一時保存・検証後に反映。失敗時は元へ自動復旧しません。

    pending記録は正式反映前に作成し、全処理成功時だけ削除します。
    停電等で履歴だけ整合して残っても、記録があれば次回処理を停止します。
    """
    objHistoryPath, objLatestPath = build_snapshot_paths(str(objExcelPath), iHistoryNumber)
    if objHistoryPath.exists() or objLatestPath.exists():
        raise ValueError("今回の保存先の履歴が既に存在します。上書きしません。")
    objBackupDirectory = (
        objExcelPath.parent / (objExcelPath.stem + "_highlighter_backups")
        / (datetime.now().strftime("%Y%m%d_%H%M%S_%f") + "_" + uuid.uuid4().hex)
    )
    objBackupDirectory.mkdir(parents=True, exist_ok=False)
    objBackupExcel = objBackupDirectory / objExcelPath.name
    objPendingPath = pending_path(objExcelPath)
    listTemporaryPaths = []
    pszStage = "バックアップ作成"
    bytesPreviousLatest = None
    try:
        write_exclusive_bytes(objBackupExcel, bytesOriginalExcel)
        if objPreviousLatestPath is not None:
            bytesPreviousLatest = objPreviousLatestPath.read_bytes()
            write_exclusive_bytes(objBackupDirectory / objPreviousLatestPath.name, bytesPreviousLatest)
        pszRecovery = (
            "このバックアップは、担当者の編集・保存後、0補完・赤字化前の状態です。\n"
            "自動復旧は行いません。正常終了時もバックアップを保管します。\n"
            "保存失敗時の復旧手順:\n"
            "1. Highlighterと対象Excelを閉じ、別の処理が動作していないことを確認します。\n"
            "2. 現在のExcel・履歴・pending記録を別の場所に保全します。\n"
            "3. バックアップExcelを次の入力パスへコピーします。\n"
            + str(objBackupExcel) + " -> " + str(objExcelPath) + "\n"
            "4. 今回分の次の履歴があれば削除します（それ以外の履歴は残します）。\n"
            + str(objHistoryPath) + "\n" + str(objLatestPath) + "\n"
        )
        if objPreviousLatestPath is not None:
            pszRecovery += (
                "5. バックアップ内の旧latestを次のパスへコピーします。\n"
                + str(objBackupDirectory / objPreviousLatestPath.name)
                + " -> " + str(objPreviousLatestPath) + "\n"
            )
        else:
            pszRecovery += "5. 初回準備のため、復元する旧latestはありません。\n"
        pszRecovery += (
            "6. Excelと履歴を確認し、復旧が完了した場合だけ次のpending記録を削除します。\n"
            + str(objPendingPath) + "\n"
            "7. 異常終了のロックが残っている場合は、他の処理がないことを確認して削除し、再実行します。\n"
            "復旧時は担当者が保存した修正を残し、保存処理だけをやり直します。\n"
        )
        write_exclusive_bytes(objBackupDirectory / "recovery.txt", pszRecovery.encode("utf-8"))
        write_exclusive_bytes(objPendingPath, json.dumps({
            "inputExcel": str(objExcelPath),
            "backupDirectory": str(objBackupDirectory),
            "newHistory": str(objHistoryPath),
            "newLatest": str(objLatestPath),
            "previousLatest": str(objPreviousLatestPath) if objPreviousLatestPath else None,
        }, ensure_ascii=False, indent=2).encode("utf-8"))

        pszStage = "一時ファイル作成・検証"
        def temporary_path(pszSuffix: str) -> Path:
            iDescriptor, pszPath = tempfile.mkstemp(
                prefix="." + objExcelPath.stem + "_highlighter_", suffix=pszSuffix,
                dir=objExcelPath.parent,
            )
            os.close(iDescriptor)
            objPath = Path(pszPath)
            listTemporaryPaths.append(objPath)
            return objPath

        bytesTsv = serialize_tsv(listRows).encode("utf-8")
        objTemporaryHistory = temporary_path(".tsv")
        objTemporaryLatest = temporary_path(".tsv")
        objTemporaryHistory.write_bytes(bytesTsv)
        objTemporaryLatest.write_bytes(bytesTsv)
        if (serialize_tsv(load_previous_tsv(objTemporaryHistory)).encode("utf-8") != bytesTsv
                or objTemporaryLatest.read_bytes() != bytesTsv):
            raise OSError("一時履歴の検証に失敗しました。")
        objTemporaryExcel = None
        if objPreviousLatestPath is not None:
            objTemporaryExcel = temporary_path(".xlsx")
            objWorkbook.save(objTemporaryExcel)
            objCheckWorkbook = load_excel(str(objTemporaryExcel))
            try:
                listCheckRows = read_excel_rows(find_target_worksheet(objCheckWorkbook))
                if serialize_tsv(listCheckRows) != bytesTsv.decode("utf-8"):
                    raise OSError("一時Excelと履歴の数量・赤字情報が一致しません。")
            finally:
                objCheckWorkbook.close()

        # 処理中にExcelやlatestが外部から更新された場合は上書きしません。
        if objExcelPath.read_bytes() != bytesOriginalExcel:
            raise OSError("処理中に入力Excelが更新されました。正式反映を中止します。")
        if objPreviousLatestPath is not None:
            if objPreviousLatestPath.read_bytes() != bytesPreviousLatest:
                raise OSError("処理中に比較用latestが更新されました。")
            objPreviousHistory = objPreviousLatestPath.with_name(
                objPreviousLatestPath.name.removesuffix("_latest.tsv") + ".tsv"
            )
            if objPreviousHistory.read_bytes() != bytesPreviousLatest:
                raise OSError("処理中に前回履歴が更新されました。")
        if objExcelPath.with_name("~$" + objExcelPath.name).exists():
            raise OSError("入力Excelが開かれました。保存処理を中止します。")

        pszStage = "新しい履歴の保存"
        write_exclusive_bytes(objHistoryPath, objTemporaryHistory.read_bytes())
        pszStage = "新しいlatestの保存"
        write_exclusive_bytes(objLatestPath, objTemporaryLatest.read_bytes())
        if objTemporaryExcel is not None:
            pszStage = "Excelの置き換え"
            os.replace(objTemporaryExcel, objExcelPath)
        if objPreviousLatestPath is not None:
            pszStage = "旧latestの削除"
            objPreviousLatestPath.unlink()
        pszStage = "保存結果の確認"
        if objHistoryPath.read_bytes() != bytesTsv or objLatestPath.read_bytes() != bytesTsv:
            raise OSError("正式な履歴とlatestの内容が一致しません。")
        remove_old_error_file(str(objExcelPath))
        objPendingPath.unlink()
    except Exception as objException:
        raise OSError(
            "保存処理に失敗しました。自動復旧は行っていません。"
            + "\n失敗した段階: " + pszStage
            + "\n入力Excel: " + str(objExcelPath)
            + "\nバックアップ: " + str(objBackupDirectory)
            + "\n未完了保存の記録: " + str(objPendingPath)
            + "\nバックアップ内のrecovery.txtを確認してください。"
            + "\n詳細: " + str(objException)
        ) from objException
    finally:
        for objTemporaryPath in listTemporaryPaths:
            try:
                objTemporaryPath.unlink(missing_ok=True)
            except OSError:
                pass
    return objHistoryPath, objLatestPath, objBackupDirectory


def process_excel_file(pszInputFileFullPath: str) -> None:
    """編集後Excelを前回履歴と比較し、赤字と履歴を累積更新します。"""
    objExcelPath = Path(validate_excel_path(pszInputFileFullPath))
    with processing_lock(objExcelPath):
        bytesOriginalExcel = objExcelPath.read_bytes()
        objWorkbook = load_workbook(io.BytesIO(bytesOriginalExcel))
        try:
            listRows = read_excel_rows(find_target_worksheet(objWorkbook))
            bIsFirstRun, objPreviousLatestPath, iNumber = validate_history_state(*find_history_files(str(objExcelPath)))
            if bIsFirstRun or objPreviousLatestPath is None:
                raise ValueError("比較用履歴がありません。編集前に--prepareを実行してください。")
            listChangedCells = compare_quantity(listRows, load_previous_tsv(objPreviousLatestPath))
            objHistoryPath, objLatestPath, objBackup = save_output_set(
                objExcelPath, bytesOriginalExcel, objWorkbook, listRows, iNumber, objPreviousLatestPath,
            )
        finally:
            objWorkbook.close()
    print("Excelと履歴TSVを更新しました。")
    print("Changed quantity cells: " + str(len(listChangedCells)))
    print("Excel: " + str(objExcelPath))
    print("History: " + str(objHistoryPath))
    print("Latest: " + str(objLatestPath))
    print("Backup: " + str(objBackup))


def prepare_history_before_edit(pszInputFileFullPath: str) -> None:
    """初回だけ履歴を保存し、既存履歴がある場合は検証だけを行います。"""
    objExcelPath = Path(validate_excel_path(pszInputFileFullPath))
    with processing_lock(objExcelPath):
        bytesOriginalExcel = objExcelPath.read_bytes()
        objWorkbook = load_workbook(io.BytesIO(bytesOriginalExcel))
        try:
            listRows = read_excel_rows(find_target_worksheet(objWorkbook))
            bIsFirstRun, objPreviousLatestPath, iNumber = validate_history_state(*find_history_files(str(objExcelPath)))
            if bIsFirstRun:
                objHistoryPath, objLatestPath, objBackup = save_output_set(
                    objExcelPath, bytesOriginalExcel, objWorkbook, listRows, iNumber, None,
                )
                print("編集前の初回履歴を作成しました。")
                print("History: " + str(objHistoryPath))
                print("Latest: " + str(objLatestPath))
                print("Backup: " + str(objBackup))
            else:
                validate_matching_orders(listRows, load_previous_tsv(objPreviousLatestPath))
                remove_old_error_file(str(objExcelPath))
                print("編集前の履歴確認が完了しました。")
        finally:
            objWorkbook.close()


def main() -> int:
    """引数を確認して処理を実行し、成功0・失敗1の終了コードを返します。"""
    iArgumentCount: int = len(sys.argv)
    bIsPrepareMode: bool = iArgumentCount == 3 and sys.argv[1] == "--prepare"
    bIsCompareMode: bool = iArgumentCount == 2
    if not bIsPrepareMode and not bIsCompareMode:
        pszScriptFileName: str = os.path.basename(__file__)
        pszErrorMessage: str = (
            "Error: 実行モードとExcelファイルパスを正しく指定してください。\n"
            + "Usage: python "
            + pszScriptFileName
            + " <excel_file_path>\n"
            + "Prepare: python "
            + pszScriptFileName
            + " --prepare <excel_file_path>\n"
        )
        print(pszErrorMessage, file=sys.stderr, end="")
        pszErrorFileFullPath: str = os.path.splitext(pszScriptFileName)[0] + "_error_argument.txt"
        try:
            write_error_text(pszErrorFileFullPath, pszErrorMessage)
        except OSError as objException:
            print("Error: 引数エラーファイルを保存できません。Detail = " + str(objException), file=sys.stderr)
        return 1

    pszInputFileFullPath: str = sys.argv[2] if bIsPrepareMode else sys.argv[1]
    try:
        if bIsPrepareMode:
            prepare_history_before_edit(pszInputFileFullPath)
        else:
            process_excel_file(pszInputFileFullPath)
    except Exception as objException:
        report_processing_error(
            pszInputFileFullPath,
            "発注数量変更の強調処理",
            str(objException),
        )
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
