###
#DOCX files are not true binaries - rather, they consist of a zip archive that contains 
# an xml typsetting source of the file that M$Word renders (among other files)
#
#This python script is a compact way to automate extracting the source from the docx input
# file. Executing it from the command line:
# 		python unzipdocx.py myword.docx --output myword
# will unzip the source archive from myword.docx into directory myword.
###

import zipfile
import argparse
import os

def unzip_docx(docx_path: str, output_dir: str):
    """
    Unzips the entire DOCX file (a ZIP archive) into the specified directory.
    """
    if not os.path.exists(docx_path):
        print(f"Error: DOCX file not found at '{docx_path}'")
        return

    # Ensure the output directory exists
    os.makedirs(output_dir, exist_ok=True)

    try:
        with zipfile.ZipFile(docx_path, 'r') as docx_zip:
            docx_zip.extractall(output_dir)
            print(f"Successfully extracted DOCX contents into '{output_dir}'")

    except zipfile.BadZipFile:
        print(f"Error: '{docx_path}' is not a valid DOCX (zip) file.")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description="Unpack .docx archive into a directory."
    )
    parser.add_argument(
        "input_docx",
        type=str,
        help="Path to the input .docx file."
    )
    parser.add_argument(
        "-o", "--output",
        type=str,
        default="docx_unpacked",
        help="Directory where the files will be extracted (default: ./docx_unpacked)."
    )

    args = parser.parse_args()
    unzip_docx(args.input_docx, args.output)
