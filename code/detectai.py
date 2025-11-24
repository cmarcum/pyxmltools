###
#M$Office files are not true binaries - rather, they consist of a zip archive that contains 
# an xml typsetting source of the file that M$Word renders (among other files)
#
#This python script automates flagging whether docx, xlsx, or pptx files should be reviewed
# for further evaluation as being AI-, machine-, or script-generated.
#
#Executing it from the command line:
# 		python detectai.py foo.bar
# will print some flags to the screen if ai telltales are detected.
###

import zipfile
import os
import argparse
import xml.etree.ElementTree as ET
from datetime import datetime

# ------------------------------------------------------------
# Utility: safe XML load
# ------------------------------------------------------------
def load_xml_from_zip(z, path):
    try:
        with z.open(path) as f:
            return ET.fromstring(f.read())
    except:
        return None


# ------------------------------------------------------------
# Heuristic Tests
# ------------------------------------------------------------

def is_timestamp_suspicious(created, modified):
    try:
        c = datetime.fromisoformat(created.replace("Z", ""))
        m = datetime.fromisoformat(modified.replace("Z", ""))
        return abs((m - c).total_seconds()) < 1
    except:
        return False


def check_metadata(core_xml):
    findings = []
    ns = {
        "dc": "http://purl.org/dc/elements/1.1/",
        "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    }

    creator = core_xml.find("dc:creator", ns)
    modifier = core_xml.find("cp:lastModifiedBy", ns)
    created = core_xml.find("dcterms:created", {"dcterms": "http://purl.org/dc/terms/"})
    modified = core_xml.find("dcterms:modified", {"dcterms": "http://purl.org/dc/terms/"})

    if creator is not None and ("python-docx" in creator.text.lower() or creator.text.strip() == ""):
        findings.append("Creator is empty or indicates python-docx.")

    if modifier is not None and creator is not None and modifier.text == creator.text:
        findings.append("LastModifiedBy equals Creator (common in machine-generated files).")

    if created is not None and modified is not None:
        if is_timestamp_suspicious(created.text, modified.text):
            findings.append("Created and Modified timestamps are identical.")

    return findings


def expected_parts_by_ext(ext):
    if ext == ".docx":
        return [
            "word/document.xml",
            "word/styles.xml",
            "word/settings.xml",
            "word/fontTable.xml",
        ]
    if ext == ".xlsx":
        return [
            "xl/workbook.xml",
            "xl/styles.xml",
            "xl/_rels/workbook.xml.rels",
        ]
    if ext == ".pptx":
        return [
            "ppt/presentation.xml",
            "ppt/slides/slide1.xml",
            "ppt/_rels/presentation.xml.rels",
        ]
    return []


def detect_suspicious(path):
    ext = os.path.splitext(path)[1].lower()

    if ext not in (".docx", ".xlsx", ".pptx"):
        return ["Not an Office OpenXML document."]

    findings = []

    try:
        with zipfile.ZipFile(path, 'r') as z:
            names = z.namelist()

            # Missing expected Office files
            for part in expected_parts_by_ext(ext):
                if part not in names:
                    findings.append(f"Missing expected part: {part}")

            # Core properties
            core_xml = load_xml_from_zip(z, "docProps/core.xml")
            if core_xml is not None:
                findings.extend(check_metadata(core_xml))
            else:
                findings.append("Missing core.xml metadata file (unusual for human-authored docs).")

            # Suspiciously small archives
            if len(names) <= 6:
                findings.append("Archive contains very few files (likely machine-generated).")

            # Check for missing thumbnails
            if "docProps/thumbnail.jpeg" not in names:
                findings.append("No thumbnail found (often missing in machine-generated files).")

    except zipfile.BadZipFile:
        findings.append("Invalid ZIP file (corrupt or not an Office document).")

    return findings


# ------------------------------------------------------------
# Main CLI
# ------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Flag suspicious DOCX, XLSX, and PPTX files.")
    parser.add_argument("targets", nargs="+", help="Files or directories to scan.")

    args = parser.parse_args()

    for target in args.targets:
        if os.path.isdir(target):
            for root, _, files in os.walk(target):
                for f in files:
                    path = os.path.join(root, f)
                    findings = detect_suspicious(path)
                    if len(findings) > 0:
                        print(f"\n {path}")
                        for line in findings:
                            print(f"  - {line}")
        else:
            findings = detect_suspicious(target)
            print(f"\n {target}")
            for line in findings:
                print(f"  - {line}")


if __name__ == "__main__":
    main()
