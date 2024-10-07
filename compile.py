import os
import shutil
from collections.abc import Iterable
from zipfile import ZipFile, ZIP_DEFLATED
from lxml import etree
import tempfile
from typing import Union
from win32com.client import Dispatch

str_or_PathLike = Union[str, os.PathLike]

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PPTM_FILE_NAME = "macros.pptm"
PPAM_FILE_NAME = "macros.ppam"
BAS_DIR = os.path.join(SCRIPT_DIR, "src", "Macros")
OUTPUT_DIR = os.path.join(SCRIPT_DIR, "bin")
RELS_ARC_NAME = "_rels/.rels"

def create_pptm_with_modules(
        ppt_app: Dispatch,
        bas_files: Iterable[str],
        outdir: str):
    presentation = ppt_app.Presentations.Add()

    # Import bas modules
    for bas_file in bas_files:
        import_module(presentation, bas_file)

    # Save and close the presentation
    pptm = os.path.join(outdir, PPTM_FILE_NAME)
    print(f"Saving: {pptm}")
    os.makedirs(outdir, exist_ok=True)
    presentation.SaveAs(pptm, 11)  # 11: ppSaveAsDefault

    # Save ppam file in a trusted dir temporalily and move to outdir
    ppam = os.path.join(os.path.expanduser("~"), "Documents", PPAM_FILE_NAME)
    print(f"Saving: {ppam}")
    presentation.SaveAs(ppam, 11)
    shutil.move(ppam, os.path.join(outdir, PPAM_FILE_NAME))

    print("Closing Presentation")
    presentation.Close()


def import_module(presentation: Dispatch, bas_file: str):
    vba_project = presentation.VBProject
    try:
        vba_project.VBComponents.Import(bas_file)
        print(f"Imported: {bas_file}")
    except Exception as e:
        print(f"Failed to import {bas_file}: {e}")

def generate_rels(base_xml: str):
    tree = etree.fromstring(base_xml)
    new_relationship = etree.Element(
        'Relationship',
        Id='myCustomUI',
        Type='http://schemas.microsoft.com/office/2006/relationships/ui/extensibility',
        Target='customUI/customUI.xml'
    )
    tree.append(new_relationship)
    #modified_xml_bytes = io.BytesIO()
    #tree.write(modified_xml_bytes, encoding='utf-8', xml_declaration=True)
    #modified_xml_bytes.seek(0)
    #return modified_xml_bytes
    return etree.tostring(tree, encoding='unicode')

def replace_rels(ppam_file, rels_arcname):
    tmp_ppam_file = tempfile.NamedTemporaryFile().name
    data = None
    print(f"Modifying: {rels_arcname} in {ppam_file}")
    with ZipFile(ppam_file, 'r') as old_ppam, ZipFile(tmp_ppam_file, 'w', compression=ZIP_DEFLATED) as tmp_ppam:
        # Copy all files except the old rels file
        for item in old_ppam.infolist():
            arcname = item.filename
            if arcname == rels_arcname:
                xml_content = old_ppam.read(arcname)
                data = generate_rels(xml_content)
                # Add the modified XML file to the new ZIP archive
                tmp_ppam.writestr(arcname, data)
            else:
                # copy form old zip file
                tmp_ppam.writestr(arcname, old_ppam.read(arcname))
    shutil.move(tmp_ppam.filename, ppam_file)


def add_files_to_zip(
        zip_file: str_or_PathLike,
        files_to_add: Union[tuple, list]  # [[filename, arc_dir], ...]
    ):
    with ZipFile(zip_file, 'a', compression=ZIP_DEFLATED) as zip_out:
        for filename, arc_dir in files_to_add:
            add_file_to_zip(zip_out, filename, arc_dir)

def add_file_to_zip(zip_out: ZipFile, filename: str, arc_dir: str_or_PathLike):
    zip_out.write(filename, os.path.join(arc_dir, filename))
    print(f'{filename} successfully added into {arc_dir}')


if __name__ == "__main__":
    ppt_app = Dispatch("PowerPoint.Application")
    ppt_app.Visible = True
    bas_files = [os.path.join(BAS_DIR, f) for f in os.listdir(BAS_DIR) if f.endswith('.bas')]
    create_pptm_with_modules(ppt_app, bas_files, OUTPUT_DIR)
    replace_rels(os.path.join(OUTPUT_DIR, PPAM_FILE_NAME), RELS_ARC_NAME)
    print("Quitting PowerPoint")
    ppt_app.Quit()
