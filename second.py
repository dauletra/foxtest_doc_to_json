from main import paragraph_to_html
import os
import json

import win32com.client


if __name__ == '__main__':
    folder_name = 'test.documents'
    document_name = 'sub_sup_scripts.doc'

    abs_folder_path = os.path.abspath(folder_name)
    assert os.path.isdir(abs_folder_path), 'папка не существует'
    abs_file_path = os.path.join(abs_folder_path, document_name)
    assert os.path.isfile(abs_file_path), 'файл не существует'

    word = win32com.client.Dispatch('Word.Application')
    word.Visible = True

    document = word.Documents.Open(abs_file_path)
    print('-- Файл открыт --')

    my_document = []
    for index, paragraph in enumerate(document.Paragraphs, start=1):
        if len(paragraph.Range.Text) > 1:
            text = paragraph_to_html(document, paragraph).strip()
            my_document.append({
                "id": index,
                "html": text
            })

    print('-- Конвертирован')

    json_name = document_name + '.json'
    json_abs_path = os.path.join(abs_folder_path, json_name)
    with open(json_abs_path, 'w', encoding='utf8') as file:
        json.dump(my_document, file, ensure_ascii=False)
