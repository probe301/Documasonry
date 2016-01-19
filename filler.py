












from pylon import AutoDelegator
from pylon import puts
import re
from pylon import dedupe
import win32com.client


class Filler(AutoDelegator):
  """ docstring for Filler

      filler = Filler.from_template(path='')
      filler.detect_required_fields()
      filler.render(info=yaml_info)
      filler.save(folder='path/to/', close=True)

  """
  def __init__(self, template_path, app=None):
    self.template_path = template_path
    if template_path.endswith(['.xls', '.xlsx']):
      excel = win32com.client.Dispatch('Excel.Application')
      # excel.Visible = False
      self.app_name = 'Office Excel'
      filler = ExcelFiller(template_path, excel)

    elif template_path.endswith(['.doc', '.docx']):
      word = win32com.client.Dispatch('Word.Application')
      # word.Visible = False
      self.app_name = 'Office Word'
      filler = WordFiller(template_path, word)

    elif template_path.endswith(['.dwg', ]):
      cad = win32com.client.Dispatch('AutoCAD.Application')
      # word.Visible = False
      self.app_name = 'Autodesk AutoCAD'
      filler = WordFiller(template_path, cad)


    self.delegates = [filler]
    # delegates:
    # filler.detect_required_fields()
    # filler.render(info=yaml_info)
    # filler.save(folder='path/to/', close=True)

  def __str__(self):
    return '<Filler type={} template_path={}>'.format(self.app_name, self.template_path)








class WordFiller:
  def __init__(self, template_path, app):
    self.template_path = template_path
    self.app = app



  def detect_required_fields(self):

    self.document = self.app.Documents.Open(self.template_path)
    self.app.Selection.HomeKey(6)
    text_range = self.document.Content
    text = "\{{*\}}"
    # pfunc(text_range.Find.Execute)
    # ArgSpec(args=['self', 'FindText', 'MatchCase', 'MatchWholeWord',
    # 'MatchWildcards', 'MatchSoundsLike', 'MatchAllWordForms', 'Forward',
    # 'Wrap', 'Format', 'ReplaceWith', 'Replace', 'MatchKashida',
    # 'MatchDiacritics', 'MatchAlefHamza', 'MatchControl'],
    # varargs=None, keywords=None, defaults=())


    field_names = re.findall(r'{{.+?}}', self.template_path)
    while text_range.Find.Execute(text, False, False, True, False, False, True, 0, True, 'NewStr', 0):
      field_names.append(text_range.Text)
    result = dedupe([name.strip().split('|')[0] for name in field_names])
    self.document.Close()
    return result



  def fill_fields(self):

    for field in self.detect_fields(raw=True, close=False):
      self.word_app.Selection.HomeKey(6)
      text_range = self.document.Content
      val = field.calculate(self.content)
      if not val:
        self.record('无法找到字段的值 {}'.format(field.raw), important=True)
        continue
      while text_range.Find.Execute(field.raw, False, False, False, False, False, True, 0, True, val, 2):
      # 'Replace' => 0 no replace ?
      # 'Replace' => 1 replace once ?
      # 'Replace' => 2 replace all ?
        self.record('replace done {} ---> {}'.format(field.raw, val))
      # print(field_names)
    self.record('all fields replace done')
      # w.Selection.Find.Execute(OldStr, False, False, False, False, False, True, 1, True, NewStr, 2)







class ExcelFiller:
  def __init__(self, template_path, app):
    self.template_path = template_path
    self.app = app







class AutoCADFiller:
  def __init__(self, template_path, app):
    self.template_path = template_path
    self.app = app
