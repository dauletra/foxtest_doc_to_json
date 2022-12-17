import os
import base64
import json
from io import BytesIO

import win32com.client
from PIL import Image


class WdSaveOptions:
    wdDoNotSaveChanges = 0
    wdPromptToSaveChanges = -2
    wdSaveChanges = -1


class WdFindWrap:
    wdFindAsk = 2
    wdFindContinue = 1
    wdFindStop = 0


class WdReplace:
    wdReplaceAll = 2
    wdReplaceNone = 0
    wdReplaceOne = 1


def open_document(path: str) -> tuple:
    path = path if os.path.isabs(path) else os.path.abspath(path)
    if not os.path.isfile(path):
        raise FileNotFoundError(f'Ошибка: {os.path.basename(path)} не найден или не является файлом. '
                                f'Полный путь: {path}')
    if not os.path.splitext(path)[1] == '.doc':
        raise TypeError(f'Формат файла не является .doc')

    word = win32com.client.Dispatch('Word.Application')
    word.Visible = True
    doc = word.Documents.Open(path)
    return word, doc


def convert(doc1) -> list:
    document_json = []

    # Заменить нечитаемые символы
    symbols = (('^u61477', '%'),
               ('^u61513', 'I'),
               ('^u61472', ' '),
               ('&', '&amp;'),
               ('<', '&lt;'),
               ('>', '&gt;'),
               ('"', '&quot;'),
               ('\'', '&#39;'))
    for code, symbol in symbols:
        doc1.Range().Find.Execute(FindText=code,
                                 MatchCase=False,
                                 MatchWholeWord=False,
                                 MatchWildcards=False,
                                 MatchSoundsLike=False,
                                 MatchAllWordForms=False,
                                 Forward=True,
                                 Wrap=WdFindWrap.wdFindContinue,
                                 Format=False,
                                 ReplaceWith=symbol,
                                 Replace=WdReplace.wdReplaceAll)

    rng = doc1.Range()
    rng.Find.Text = ''
    rng.Find.Font.Superscript = True
    while rng.Find.Execute():
        rng.InsertBefore('<sup>')
        rng.InsertAfter('</sup>')
        rng.Font.Superscript = False
        rng.Find.Text = ''
        rng.Find.Font.Superscript = True
        rng = doc1.Range()

    rng = doc1.Range()
    rng.Find.Font.Subscript = True
    while rng.Find.Execute():
        rng.InsertBefore('<sub>')
        rng.InsertAfter('</sub>')
        rng.Font.Subscript = False
        rng.Find.Text = ''
        rng.Find.Font.Subscript = True
        rng = doc1.Range()

    rng = doc1.Range()
    rng.Find.Font.Underline = True
    while rng.Find.Execute():
        rng.InsertBefore('<u>')
        rng.InsertAfter('</u>')
        rng.Font.Underline = False
        rng.Find.Text = ''
        rng.Find.Font.Underline = True
        rng = doc1.Range()

    rng = doc1.Range()
    rng.Find.Font.Italic = True
    while rng.Find.Execute():
        rng.InsertBefore('<i>')
        rng.InsertAfter('</i>')
        rng.Font.Italic = False
        rng.Find.Text = ''
        rng.Find.Font.Italic = True
        rng = doc1.Range()

    shapeScale = 4
    img_temp = '<img align="Middle" src="data:image/png;base64,{0}" />'
    for shape in doc1.Range().InlineShapes:
        wmfim = Image.open(BytesIO(shape.Range.EnhMetaFileBits))
        wmfim = wmfim.resize(map(lambda heightwidth: heightwidth // shapeScale, wmfim.size))
        bimage = BytesIO()
        wmfim.save(bimage, format='png')
        img_str = base64.b64encode(bimage.getvalue())
        shape.Range.Text = img_temp.format(img_str.decode('ascii'))

    for index, para in enumerate(doc1.Paragraphs, start=1):
        text = para.Range.Text.strip()
        if text:
            document_json.append({
                "id": index,
                "html": text
            })

    return document_json


def paragraph_to_html(doc, para):
    rng = doc.Range(para.Range.Start, para.Range.End - 1)
    rng.Find.Text = ''
    rng.Find.Font.Superscript = True
    while rng.Find.Execute():
        rng.InsertBefore('<sup>')
        rng.InsertAfter('</sup>')
        rng.Font.Superscript = False
        rng = doc.Range(para.Range.Start, para.Range.End - 1)
        rng.Find.Text = ''
        rng.Find.Font.Superscript = True

    rng = doc.Range(para.Range.Start, para.Range.End - 1)
    rng.Find.Text = ''
    rng.Find.Font.Subscript = True
    while rng.Find.Execute():
        rng.InsertBefore('<sub>')
        rng.InsertAfter('</sub>')
        rng.Font.Subscript = False
        rng = doc.Range(para.Range.Start, para.Range.End - 1)
        rng.Find.Text = ''
        rng.Find.Font.Subscript = True

    rng = doc.Range(para.Range.Start, para.Range.End - 1)
    rng.Find.Font.Underline = True
    while rng.Find.Execute():
        rng.InsertBefore('<u>')
        rng.InsertAfter('</u>')
        rng.Font.Underline = False
        rng = doc.Range(para.Range.Start, para.Range.End - 1)
        rng.Find.Text = ''
        rng.Find.Font.Underline = True

    rng = doc.Range(para.Range.Start, para.Range.End - 1)
    rng.Find.Font.Italic = True
    while rng.Find.Execute():
        rng.InsertBefore('<i>')
        rng.InsertAfter('</i>')
        rng.Font.Italic = False
        rng = doc.Range(para.Range.Start, para.Range.End - 1)
        rng.Find.Text = ''
        rng.Find.Font.Italic = True

    para.Range.Find.Execute(FindText='\\n',
                             MatchCase=False,
                             MatchWholeWord=False,
                             MatchWildcards=False,
                             MatchSoundsLike=False,
                             MatchAllWordForms=False,
                             Forward=True,
                             Wrap=WdFindWrap.wdFindContinue,
                             Format=False,
                             ReplaceWith='<br/>',
                             Replace=WdReplace.wdReplaceAll)

    shapeScale = 4
    img_temp = '<img align="Middle" src="data:image/png;base64,{0}" />'
    for shape in para.Range.InlineShapes:
        wmfim = Image.open(BytesIO(shape.Range.EnhMetaFileBits))
        wmfim = wmfim.resize(map(lambda heightwidth: heightwidth // shapeScale, wmfim.size))
        bimage = BytesIO()
        wmfim.save(bimage, format='png')
        img_str = base64.b64encode(bimage.getvalue())
        shape.Range.Text = img_temp.format(img_str.decode('ascii'))

    return para.Range.Text.strip()


def replace_symbols(doc):
    symbols = (('^u61477', '%'),
               ('^u61513', 'I'),
               ('^u61472', ' '),
               ('&', '&amp;'),
               ('<', '&lt;'),
               ('>', '&gt;'),
               ('"', '&quot;'),
               ('\'', '&#39;'))
    for code, symbol in symbols:
        doc.Range().Find.Execute(FindText=code,
                                 MatchCase=False,
                                 MatchWholeWord=False,
                                 MatchWildcards=False,
                                 MatchSoundsLike=False,
                                 MatchAllWordForms=False,
                                 Forward=True,
                                 Wrap=WdFindWrap.wdFindContinue,
                                 Format=False,
                                 ReplaceWith=symbol,
                                 Replace=WdReplace.wdReplaceAll)


folder_name: str = "documents"
prefix: str = 'LS_'
def get_files_to_convert():
    print('+' * 50)
    if not os.path.isdir(folder_name):
        print(f'Создайте папку "{folder_name}", а затем поместите туда документы в формате .doc')
        exit()

    abs_path = os.path.abspath(folder_name)
    document_names = [f for f in os.listdir(abs_path) if
                      os.path.isfile(os.path.join(abs_path, f)) and
                      os.path.splitext(f)[1] == '.doc' and os.path.basename(f)[:2] != '~$']

    json_files = [item for item in os.listdir(abs_path) if
                  os.path.isfile(os.path.join(abs_path, item)) and
                  os.path.splitext(item)[1] == '.json']

    raw_document_names = []

    for doc_name in document_names:
        if prefix + doc_name + '.json' in json_files:
            continue
        raw_document_names.append(doc_name)

    if len(raw_document_names) == 0:
        print('Все документы конвертированы: ', len(raw_document_names))
        exit()

    print('Программа автоматический конвертирует следующие doc файлы в json')
    print('\n'.join(raw_document_names))
    print('*' * 50)

    key = input('Хотите продолжить?(Y/n)')
    if key.upper() != 'Y' and key != '':
        print('Отменено')
        exit()
    return raw_document_names, abs_path


if __name__ == '__main__':
    raw_document_names, abs_path = get_files_to_convert()

    word = win32com.client.Dispatch('Word.Application')
    word.Visible = True

    for document_name in raw_document_names:
        file_path = os.path.join(abs_path, document_name)
        assert os.path.exists(file_path), 'Даулет, файл не существует'

        print('-- Файл ' + document_name)
        document = word.Documents.Open(file_path)
        print('-- Открыт')
        replace_symbols(document)

        my_document = []

        for index, paragraph in enumerate(document.Paragraphs, start=1):
            if len(paragraph.Range.Text.strip()) > 1:
                text = paragraph_to_html(document, paragraph).strip()
                my_document.append({
                    "id": index,
                    "html": text
                })
        print('-- Конвертирован')

        new_name = prefix + document_name + '.json'
        json_abs_path = os.path.join(abs_path, new_name)
        with open(json_abs_path, 'w', encoding='utf8') as file:
            json.dump(my_document, file, ensure_ascii=False)
        print('-- JSON файл сохранен')
        document.Close(WdSaveOptions.wdDoNotSaveChanges)
        print('*' * 30)

    word.Quit(WdSaveOptions.wdDoNotSaveChanges)
    print('*' * 50)
