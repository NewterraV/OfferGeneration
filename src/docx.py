import os
from docxtpl import DocxTemplate

def replace_placeholders(document_path, data_dict, output_path='output.docx'):
    """
    Заменяет плейсхолдеры вида {{ключ}} в документе docx на соответствующие значения из словаря.

    Args:
        document_path (str): Путь к исходному файлу docx.
        data_dict (dict): Словарь с данными для подстановки (ключи соответствуют плейсхолдерам).
        output_path (str, optional): Путь для сохранения измененного документа.
                                     По умолчанию 'output.docx'.
    """
    doc = DocxTemplate(document_path)
    doc.render(data_dict)
    doc.save(output_path)
    os.startfile(output_path)