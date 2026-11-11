# -- coding: utf-8 --
###############################################################
#
# RotateImageFile_90DegreesLeft_Cmd.py
#
# pip install pillow
#
###############################################################

import os
import sys

from PIL import Image


###############################################################
#
# write_error_text
#
###############################################################
def write_error_text(
    pszOutputFileFullPath: str,
    pszErrorMessage: str,
) -> None:
    pszDirectoryFullPath: str = os.path.dirname(pszOutputFileFullPath)
    if pszDirectoryFullPath != "":
        os.makedirs(pszDirectoryFullPath, exist_ok=True)
    with open(pszOutputFileFullPath, mode="w", encoding="utf-8") as objFile:
        objFile.write(pszErrorMessage)


###############################################################
#
# get_output_file_full_path
#
###############################################################
def get_output_file_full_path(
    pszInputFileFullPath: str,
) -> str:
    pszDirectoryFullPath: str = os.path.dirname(pszInputFileFullPath)
    pszBaseName: str = os.path.basename(pszInputFileFullPath)
    pszNameWithoutExtension, pszExtension = os.path.splitext(pszBaseName)
    pszOutputFileName: str = pszNameWithoutExtension + "_90DegreesLeft" + pszExtension
    pszOutputFileFullPath: str = os.path.join(
        pszDirectoryFullPath,
        pszOutputFileName,
    )
    return pszOutputFileFullPath


###############################################################
#
# write_conversion_error_text
#
###############################################################
def write_conversion_error_text(
    pszInputFileFullPath: str,
    pszDetailMessage: str = "",
) -> None:
    pszBaseNameWithoutExtension: str = os.path.splitext(
        os.path.basename(pszInputFileFullPath)
    )[0]
    pszErrorFileFullPath: str = os.path.join(
        os.path.dirname(pszInputFileFullPath),
        pszBaseNameWithoutExtension + "_error.txt",
    )
    pszOutputFileName: str = os.path.basename(get_output_file_full_path(pszInputFileFullPath))
    pszInputFileName: str = os.path.basename(pszInputFileFullPath)
    pszErrorMessage: str = (
        "変換に失敗しました。\n"
        + "元画像ファイル「"
        + pszInputFileName
        + "」から「"
        + pszOutputFileName
        + "」への変換に失敗しました。\n"
    )
    if pszDetailMessage != "":
        pszErrorMessage = pszErrorMessage + "詳細: " + pszDetailMessage + "\n"
    write_error_text(pszErrorFileFullPath, pszErrorMessage)


###############################################################
#
# rotate_image_file
#
###############################################################
def rotate_image_file(
    pszInputFileFullPath: str,
) -> None:
    if not os.path.exists(pszInputFileFullPath):
        write_conversion_error_text(
            pszInputFileFullPath,
            "入力画像ファイルが見つかりません。Path = " + pszInputFileFullPath,
        )
        return

    pszExtension: str = os.path.splitext(pszInputFileFullPath)[1]
    pszLowerExtension: str = pszExtension.lower()
    setAllowedExtensions: set[str] = {".jpg", ".jpeg", ".png", ".bmp", ".gif"}

    if pszLowerExtension not in setAllowedExtensions:
        write_conversion_error_text(
            pszInputFileFullPath,
            "入力ファイルの拡張子は未対応です。Path = " + pszInputFileFullPath,
        )
        return

    try:
        objImage: Image.Image = Image.open(pszInputFileFullPath)
    except Exception as objException:
        write_conversion_error_text(
            pszInputFileFullPath,
            "画像ファイルの読み込み中に予期しない例外が発生しました。Detail = "
            + str(objException),
        )
        return

    try:
        objRotatedImage: Image.Image = objImage.rotate(90, expand=True)
    except Exception as objException:
        write_conversion_error_text(
            pszInputFileFullPath,
            "画像ファイルの回転中に予期しない例外が発生しました。Detail = "
            + str(objException),
        )
        return

    pszOutputFileFullPath: str = get_output_file_full_path(pszInputFileFullPath)

    try:
        if pszLowerExtension in {".jpg", ".jpeg"}:
            if objRotatedImage.mode not in {"RGB", "L"}:
                objRotatedImage = objRotatedImage.convert("RGB")
            objRotatedImage.save(pszOutputFileFullPath, format="JPEG")
        else:
            objRotatedImage.save(pszOutputFileFullPath)
    except Exception as objException:
        write_conversion_error_text(
            pszInputFileFullPath,
            "回転後画像ファイルの書き込み中に予期しない例外が発生しました。Detail = "
            + str(objException),
        )
        return

    print("Output: " + pszOutputFileFullPath)


###############################################################
#
# main
#
###############################################################
def main() -> None:
    iArgumentCount: int = len(sys.argv)
    if iArgumentCount < 2:
        pszScriptFileName: str = os.path.basename(__file__)
        pszLine1: str = "Error: input image file path is not specified (insufficient arguments)."
        pszLine2: str = "Usage: python " + pszScriptFileName + " <input_image_file_path>"
        pszLine3: str = "Example: python " + pszScriptFileName + " C:\\Data\\sample.jpg"
        print(pszLine1)
        print(pszLine2)

        pszErrorMessage: str = pszLine1 + "\n" + pszLine2 + "\n" + pszLine3 + "\n"
        pszOutputFileFullPath: str = os.path.splitext(pszScriptFileName)[0] + "_error_argument.txt"
        write_error_text(pszOutputFileFullPath, pszErrorMessage)
        return

    pszInputFileFullPath: str = sys.argv[1]
    rotate_image_file(pszInputFileFullPath)


if __name__ == "__main__":
    main()
