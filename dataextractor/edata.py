import docx
import glob
import pandas as pd
import numpy as np
import io
import csv
from docx import table


class ProtocolData(object):
    """
    Required data from all protocols in a folder
    """
    def __init__(self, path):
        self.file_list = glob.glob(path + '/*-ПРЧ.docx')
        self.parsed_data = {}
        self.file_names = {}
        for file in self.file_list:
            print(file)
            doc = docx.Document(docx=file)
            try:
                data = parse_docx(doc)
            except ValueError:
                data = None
            if data is not None:
                self.parsed_data[get_device_name(doc)] = data
                self.file_names[get_device_name(doc)] = file

    def get_file_list(self):
        return self.file_list

    def get_latchup_data(self):
        return [device_data[0] for key, device_data in self.parsed_data.items()]

    def get_upset_data(self):
        return [device_data[1] for key, device_data in self.parsed_data.items()]

    def write_to_xlsx(self, path):
        writer = pd.ExcelWriter(path)
        for key, device_data in self.parsed_data.items():
            if device_data[0] is not None:
                device_data[0].to_excel(writer, sheet_name=key+' latchup')
            if device_data[1] is not None:
                device_data[1].to_excel(writer, sheet_name=key+' upset')
        writer.save()


def parse_docx(doc):
    latchup_data = get_test_table(doc, 'ТЭ')
    upset_data = get_test_table(doc, 'ОС')
    if latchup_data is not None or upset_data is not None:
        return [prepare_data(latchup_data), prepare_data(upset_data)]


def prepare_data(data):
    if data is None:
        return None

    df = table_to_df(data)
    processed_data = {}
    for column_title in list(df):
        if 'ЛПЭ' in column_title:
            processed_data['LET'] = df[column_title]
        if 'см-2' in column_title:
            if df[column_title].dtype == object:
                processed_data['fluence'] = [float(value.replace('·10','e').replace(',', '.').replace('Е','E'))
                                             for value in df[column_title]]
            else:
                processed_data['fluence'] = df[column_title]
        if 'N' in column_title or 'ТЭ' in column_title or 'ИО' in column_title:
            if df[column_title].dtype == object:
                processed_data['number'] = [float(value.replace('·10','e').replace(',','.').replace('5)',''))
                                            for value in df[column_title]]
            else:
                processed_data['number'] = df[column_title]
        if '°C' in column_title or '°С' in column_title or 'оС' in column_title:
            processed_data['temperature'] = df[column_title]
    if 'temperature' not in processed_data:
        processed_data['temperature'] = np.zeros(len(processed_data['number']))
    processed_data['cross'] = processed_data['number'] / processed_data['fluence']
    new_df = pd.DataFrame(processed_data)
    return new_df


def get_device_name(doc):
    """
    Finds device name in the document title
    :param doc: docx Document instance
    :return: device name as string
    """
    found = False
    for paragraph in doc.paragraphs:
        if found:
            title = paragraph.text
            for word in str(title).split():
                if any(char.isdigit() for char in word):
                    return word[0:20].replace(',', '').replace('*', '').replace('#', '')
        if 'ЦЕЛЬ ИСПЫТАНИЙ' in paragraph.text or \
           'Цель ИСПЫТАНИЙ' in paragraph.text or \
           'ЦЕЛЬ РАБОТЫ' in paragraph.text:
            found = True


def table_to_df(tab):
    """
    Transforms a docx Table instance into a pandas DataFrame instance
    :param tab: docx Table instance
    :return: pandas DataFrame instance
    """
    vf = io.StringIO()
    writer = csv.writer(vf)
    for row in tab.rows:
        writer.writerow(cell.text for cell in row.cells)
    vf.seek(0)
    return pd.read_csv(vf, decimal=",")


def get_test_table(doc, effect_type):
    """
    Looks for a Table instance in the document based on the title of a table containing
    the phrase 'Результаты испытаний' and the given effect_type name
    :param doc: Document instance
    :param str effect_type: name of the effect to search for, for example, 'ТЭ' or 'ОС'
    :return: Table instance
    """
    ion_table = None
    is_ion_table = False
    blocks = iter_block_items(doc)
    for block in blocks:
        if is_ion_table and isinstance(block, docx.table.Table):
            ion_table = block
            break
        if isinstance(block, docx.text.paragraph.Paragraph):
            if 'Результаты испытаний' in block.text \
                    and 'Таблица' in block.text \
                    and effect_type in block.text\
                    and 'протон' not in block.text:
                is_ion_table = True
    return ion_table


def iter_block_items(parent):
    """
    Yield each paragraph and table child within *parent*, in document order.
    Each returned value is an instance of either Table or Paragraph. *parent*
    would most commonly be a reference to a main Document object, but
    also works for a _Cell object, which itself can contain paragraphs and tables.
    """
    if isinstance(parent, docx.document.Document):
        parent_elm = parent.element.body
    elif isinstance(parent, docx.table._Cell):
        parent_elm = parent._tc
    else:
        raise ValueError("something's not right")

    for child in parent_elm.iterchildren():
        if isinstance(child, docx.oxml.text.paragraph.CT_P):
            yield docx.text.paragraph.Paragraph(child, parent)
        elif isinstance(child, docx.oxml.table.CT_Tbl):
            yield docx.table.Table(child, parent)
