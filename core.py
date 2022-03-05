from typing import Union
import os
from pathlib import Path


import pandas as pd
from docxtpl import DocxTemplate
from docx2pdf import convert
import fitz

import config


def read_excel_table(file_path: Union[str, Path]) -> pd.DataFrame:
    frame = pd.read_excel(file_path, header=None)
    return frame


def _render_docx(template_name, context: dict, number: int):
    template_path = config.TEMPLATE_PATH / f"{template_name}.docx"
    report_template = DocxTemplate(str(template_path))
    save_path = config.TEMP_PATH / f"{number}.docx"
    report_template.render(context)
    report_template.save(save_path)


def _generate_pdf(number: int) -> None:
    in_path = config.TEMP_PATH / f"{number}.docx"
    out_path = config.TEMP_PATH / f"{number}.pdf"
    convert(in_path, out_path, keep_active=True)
    _delete_file(in_path)


def _delete_file(file_path: Union[str, Path]) -> None:
    os.remove(file_path)


def _connect_user_pdfs(context: dict, save_path: str) -> None:
    new_pdf = fitz.open()
    for file in sorted(os.listdir(config.TEMP_PATH)):
        path = str(config.TEMP_PATH / file)
        if ".pdf" not in path:
            _delete_file(path)
            continue
        pdf = fitz.open(path)
        if "3.pdf" not in path:
            new_pdf.insert_pdf(pdf)
        else:
            pdf = fitz.open(path)
            report = fitz.open(config.TEMPLATE_PATH / "report.pdf")
            pdf.insert_pdf(report)
            new_pdf.insert_pdf(pdf)
            report.close()
        pdf.close()
        _delete_file(path)
    new_pdf.save(str(save_path + f"/{context['name']}.pdf"))
    new_pdf.close()


def process_user(contex: dict, save_path: str) -> None:
    template_names = ["performance", "certification", "referral", "report"]
    for num, template in enumerate(template_names):
        _render_docx(template, contex, num)
        _generate_pdf(num)
    _connect_user_pdfs(contex, save_path)
