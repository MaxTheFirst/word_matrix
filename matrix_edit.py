from win32com import __gen_path__
import win32com.client as win32
import re
import sys
from os import walk
from shutil import rmtree


EMPTY_MATRIX = '(■())'
START_MATRIX = '(■('
MATRIX_BODY_RE = r'(?<=\(■\()\S+(?=\)\))'
LINE_SYMBOL = '@'
COLUN_SYMBOL = '&'

def get_doc():
    try:
        word = win32.gencache.EnsureDispatch('Word.Application')
        word.ActiveDocument
        return word
    except AttributeError:
        paths = list(walk(__gen_path__))
        rmtree(paths[0][0] + '\\' + paths[0][1][0])
        return get_doc()


def get_text(doc, x, y):
    return doc.ActiveDocument.Range(x, y).Text


def add_text(doc, text, x, y=None):
    if y == None:
        y = x
    doc.ActiveDocument.Range(x, y).Text = text


def set_cursor(doc, position):
    doc.Selection.SetRange(position, position)


def add_line_text(text: str, cursor_position):
    start_index = 0
    if LINE_SYMBOL in text:
        start_index = text.rfind(LINE_SYMBOL)
    return LINE_SYMBOL + COLUN_SYMBOL * text[start_index:cursor_position].count(COLUN_SYMBOL)


def add_colun_text(text: str):
    if LINE_SYMBOL not in text:
        return text + COLUN_SYMBOL
    return text.replace(LINE_SYMBOL, COLUN_SYMBOL+LINE_SYMBOL) + COLUN_SYMBOL


def add_line(doc, cursor_position, matrix_body):
    add_text(doc, add_line_text(matrix_body, cursor_position), cursor_position+1)


def add_colun(doc, start_position, matrix_body):
    text = add_colun_text(matrix_body)
    add_text(doc, text, start_position, start_position+len(matrix_body))


def get_matrix_body(text):
    matrix_body = re.findall(MATRIX_BODY_RE, text)
    if matrix_body:
        return matrix_body[0]
    return ''


def get_matrix_start_index(omath):
    start_math_index = omath.Range.Start
    start_matrix_index = omath.Range.Text.rfind(START_MATRIX)
    if start_matrix_index != -1:
        return start_matrix_index + start_math_index 


def get_matrix(doc, is_line=True):
    selection = doc.Selection
    cursor_position = selection.Start
    omath = selection.OMaths(1)
    omath.Linearize()
    start_index = get_matrix_start_index(omath)
    matrix_text = get_text(doc, start_index, omath.Range.End)
    matrix_body = get_matrix_body(matrix_text)
    if is_line:
        add_line(doc, cursor_position, matrix_body)
        cursor_position += 1
    else:
        add_colun(doc, start_index+3, matrix_body)
        cursor_position += get_text(doc, start_index, cursor_position).count(LINE_SYMBOL) + 1
    omath.BuildUp()
    set_cursor(get_doc(), cursor_position)


if __name__ == '__main__':
    get_matrix(get_doc(), sys.argv[1] == 'add_line')