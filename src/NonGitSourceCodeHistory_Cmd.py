# -- coding: utf-8 --
###############################################################
#
# NonGitSourceCodeHistory_Cmd.py
#
###############################################################

import os
import re
import sys
from pathlib import Path


MAX_HISTORY_NUMBER: int = 9999
MAX_COPY_ATTEMPTS: int = 3


###############################################################
#
# write_error_text
#
###############################################################
def write_error_text(
    pszOutputFileFullPath: str,
    pszErrorMessage: str,
) -> None:
    """エラーメッセージをUTF-8のテキストファイルへ上書き保存します。"""
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
    """対象ファイルと同じフォルダーに作る_error.txtのパスを返します。"""
    pszAbsolutePath: str = os.path.abspath(pszInputFileFullPath)
    return os.path.join(
        os.path.dirname(pszAbsolutePath),
        os.path.basename(pszAbsolutePath) + "_error.txt",
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
    """標準エラーと対象ファイル用_error.txtへエラーを出力します。"""
    pszErrorMessage: str = (
        "処理結果: エラー\n"
        + "対象ファイル: "
        + os.path.abspath(pszInputFileFullPath)
        + "\n発生した処理: "
        + pszProcessName
        + "\nエラー内容: "
        + pszDetailMessage
        + "\n"
    )
    print(pszErrorMessage, file=sys.stderr, end="")
    try:
        write_error_text(get_error_file_full_path(pszInputFileFullPath), pszErrorMessage)
    except OSError as objException:
        print(
            "Error: エラーファイルの保存にも失敗しました。Detail = "
            + str(objException),
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
# is_backup_file_name
#
###############################################################
def is_backup_file_name(pszFileName: str) -> bool:
    """ファイル名が本プログラムの確定済みバックアップ形式か返します。"""
    objMatch: re.Match[str] | None = re.search(r"\.bk([0-9]{4})\.txt$", pszFileName)
    if objMatch is None:
        return False
    iHistoryNumber: int = int(objMatch.group(1))
    return 1 <= iHistoryNumber <= MAX_HISTORY_NUMBER


###############################################################
#
# validate_input_file_path
#
###############################################################
def validate_input_file_path(pszInputFileFullPath: str) -> str:
    """入力パスがバックアップ可能な通常ファイルであることを確認します。"""
    if pszInputFileFullPath.strip() == "":
        raise ValueError("対象ファイルのパスが空です。")
    pszAbsolutePath: str = os.path.abspath(pszInputFileFullPath)
    if not os.path.exists(pszAbsolutePath):
        raise ValueError("対象ファイルが見つかりません。Path = " + pszAbsolutePath)
    if not os.path.isfile(pszAbsolutePath):
        raise ValueError("対象パスが通常ファイルではありません。Path = " + pszAbsolutePath)
    if is_backup_file_name(os.path.basename(pszAbsolutePath)):
        raise ValueError(
            "このファイルはNonGitSourceCodeHistoryのバックアップファイルです。\n"
            + "元のソースコードファイルを指定してください。Path = "
            + pszAbsolutePath
        )
    return pszAbsolutePath


###############################################################
#
# find_history_numbers
#
###############################################################
def find_history_numbers(pszInputFileFullPath: str) -> list[int]:
    """対象ファイル名に完全一致する既存履歴の番号を返します。"""
    objInputPath: Path = Path(pszInputFileFullPath)
    objPattern: re.Pattern[str] = re.compile(
        r"^" + re.escape(objInputPath.name) + r"\.bk([0-9]{4})\.txt$"
    )
    listHistoryNumbers: list[int] = []
    for objCandidatePath in objInputPath.parent.iterdir():
        objMatch: re.Match[str] | None = objPattern.fullmatch(objCandidatePath.name)
        if objMatch is None or not objCandidatePath.is_file():
            continue
        iHistoryNumber: int = int(objMatch.group(1))
        if 1 <= iHistoryNumber <= MAX_HISTORY_NUMBER:
            listHistoryNumbers.append(iHistoryNumber)
    return listHistoryNumbers


###############################################################
#
# get_next_history_number
#
###############################################################
def get_next_history_number(pszInputFileFullPath: str) -> int:
    """既存履歴の最大番号に1を加えた次番号を返します。"""
    listHistoryNumbers: list[int] = find_history_numbers(pszInputFileFullPath)
    if len(listHistoryNumbers) == 0:
        return 1
    iMaximumHistoryNumber: int = max(listHistoryNumbers)
    if iMaximumHistoryNumber >= MAX_HISTORY_NUMBER:
        raise ValueError(
            "履歴番号が最大値9999に到達しています。新しいバックアップを作成できません。"
        )
    return iMaximumHistoryNumber + 1


###############################################################
#
# remove_unverified_backup
#
###############################################################
def remove_unverified_backup(objBackupPath: Path) -> None:
    """今回作成した未確定バックアップを削除し、失敗時は例外にします。"""
    try:
        objBackupPath.unlink()
    except OSError as objException:
        raise OSError(
            "不一致バックアップを削除できませんでした。Path = "
            + str(objBackupPath)
            + ", Detail = "
            + str(objException)
        ) from objException


###############################################################
#
# copy_and_verify_with_retry
#
###############################################################
def copy_and_verify_with_retry(
    pszInputFileFullPath: str,
    iHistoryNumber: int,
) -> tuple[str, int]:
    """排他的にコピーし、バイト一致まで合計最大3回試行します。"""
    objInputPath: Path = Path(pszInputFileFullPath)
    objBackupPath: Path = objInputPath.with_name(
        objInputPath.name + f".bk{iHistoryNumber:04d}.txt"
    )

    for iCopyAttempt in range(1, MAX_COPY_ATTEMPTS + 1):
        bytesSourceBeforeCopy: bytes = objInputPath.read_bytes()
        with open(objBackupPath, mode="xb") as objBackupFile:
            objBackupFile.write(bytesSourceBeforeCopy)
            objBackupFile.flush()
            os.fsync(objBackupFile.fileno())

        bytesSourceAfterCopy: bytes = objInputPath.read_bytes()
        bytesBackup: bytes = objBackupPath.read_bytes()
        if bytesSourceAfterCopy == bytesBackup:
            return str(objBackupPath), iCopyAttempt

        remove_unverified_backup(objBackupPath)

    raise ValueError(
        "元ファイルとバックアップの内容が3回とも一致しませんでした。"
        + "コピー中に元ファイルが変更された可能性があります。"
        + "元ファイルを使用しているアプリケーションを閉じてから、再度実行してください。"
        + " Backup = "
        + str(objBackupPath)
        + ", Copy attempts = "
        + str(MAX_COPY_ATTEMPTS)
    )


###############################################################
#
# create_source_code_history
#
###############################################################
def create_source_code_history(pszInputFileFullPath: str) -> None:
    """対象ファイルを検証し、次の1世代バックアップを作成します。"""
    pszValidatedPath: str = validate_input_file_path(pszInputFileFullPath)
    iHistoryNumber: int = get_next_history_number(pszValidatedPath)
    pszBackupFileFullPath, iCopyAttempts = copy_and_verify_with_retry(
        pszValidatedPath,
        iHistoryNumber,
    )
    remove_old_error_file(pszValidatedPath)
    print("バックアップを作成しました。")
    print("\nSource:")
    print(pszValidatedPath)
    print("\nBackup:")
    print(pszBackupFileFullPath)
    print("\nHistory number: " + f"{iHistoryNumber:04d}")
    print("\nCopy attempts: " + str(iCopyAttempts))


###############################################################
#
# main
#
###############################################################
def main() -> int:
    """引数を確認してバックアップし、成功0・失敗1を返します。"""
    if len(sys.argv) != 2:
        pszScriptFileName: str = os.path.basename(__file__)
        pszErrorMessage: str = (
            "Error: バックアップ対象ファイルを1つだけ指定してください。\n"
            + "Usage: python "
            + pszScriptFileName
            + " <source_file_path>\n"
        )
        print(pszErrorMessage, file=sys.stderr, end="")
        pszErrorFileFullPath: str = os.path.splitext(pszScriptFileName)[0] + "_error_argument.txt"
        try:
            write_error_text(pszErrorFileFullPath, pszErrorMessage)
        except OSError as objException:
            print(
                "Error: 引数エラーファイルを保存できません。Detail = "
                + str(objException),
                file=sys.stderr,
            )
        return 1

    pszInputFileFullPath: str = sys.argv[1]
    try:
        create_source_code_history(pszInputFileFullPath)
    except Exception as objException:
        report_processing_error(
            pszInputFileFullPath,
            "ソースコード履歴バックアップ作成",
            str(objException),
        )
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
