
from pylon import AutoDelegator
from pylon import puts
import re
import os
from pylon import dedupe
import win32com.client
from jinja2 import Template
from jinja2 import meta
from jinja2 import Environment
import time
from pylon import datalines

def extract_field(field):
  # {{name | lower}} -> ['name'] because lower is pipe method
  # {{(area1 - area0) / area0}} -> ['area1', 'area0'] because two variable in one field
  ast = Environment().parse(field)
  return meta.find_undeclared_variables(ast)


def evalute_field(field, info):
  # return info.content.get(field)
  template = Template(field)
  return template.render(**info.content)













####### ###### ##      ##      ####### ######
##        ##   ##      ##      ##      ##   ##
######    ##   ##      ##      ######  ######
##        ##   ##      ##      ##      ##  ##
##      ###### ####### ####### ####### ##   ##

class Filler(AutoDelegator):
  """ docstring for Filler

      filler = Filler.from_template(path='')
      filler.detect_required_fields()
      filler.render(info=yaml_info)
      filler.save(folder='path/to/', close=True)

  """
  def __init__(self, template_path, app=None):
    self.template_path = template_path
    self.info = None
    self.output_name = None
    if template_path.endswith(('.xls', '.xlsx')):
      excel = win32com.client.Dispatch('Excel.Application')
      # excel.Visible = False
      self.app_name = 'Office Excel'
      filler = ExcelFiller(template_path, excel)

    elif template_path.endswith(('.doc', '.docx')):
      word = win32com.client.Dispatch('Word.Application')
      # word.Visible = False
      self.app_name = 'Office Word'
      filler = WordFiller(template_path, word)

    elif template_path.endswith(('.dwg', '.dxf')):
      cad = win32com.client.Dispatch('AutoCAD.Application')
      # cad.Visible = False
      self.app_name = 'Autodesk AutoCAD'
      filler = AutoCADFiller(template_path, cad)


    self.delegates = [filler]
    # delegates:
    # filler.detect_required_fields()
    # filler.render(info=yaml_info)
    # filler.save(folder='path/to/', close=True)

  def __str__(self):
    return '<Filler type={} template_path={}>'.format(self.app_name, self.template_path)






  def save(self, info, folder, close=True):
    self.output_name = evalute_field(os.path.basename(self.template_path), info)
    output_path = os.path.join(folder, self.output_name)

    if os.path.exists(output_path):
      fix = time.strftime('.backup-%Y%m%d-%H%M%S')
      os.rename(output_path, fix.join(os.path.splitext(output_path)))
    try:
      self.document.SaveAs(output_path)
    except Exception:
      t = 'Word Filler can not save document: <{}>'.format(output_path)
      raise SaveDocumentError(t)

    if close:
      self.document.Close()















class NoInfoKeyError(Exception):
  pass

class SaveDocumentError(Exception):
  pass












##   ##  #####  ######  ######  ####### ###### ##      ##      ####### ######
##   ## ##   ## ##   ## ##   ## ##        ##   ##      ##      ##      ##   ##
## # ## ##   ## ######  ##   ## ######    ##   ##      ##      ######  ######
### ### ##   ## ##  ##  ##   ## ##        ##   ##      ##      ##      ##  ##
##   ##  #####  ##   ## ######  ##      ###### ####### ####### ####### ##   ##


class WordFiller:
  def __init__(self, template_path, app):
    self.template_path = template_path
    self.app = app
    self.document = self.app.Documents.Open(self.template_path)


  def detect_required_fields(self, close=True, unique=False):

    self.app.Selection.HomeKey(6)
    text_range = self.document.Content
    text = "\{\{*\}\}"
    # pfunc(text_range.Find.Execute)
    # ArgSpec(args=['self', 'FindText', 'MatchCase', 'MatchWholeWord',
    # 'MatchWildcards', 'MatchSoundsLike', 'MatchAllWordForms', 'Forward',
    # 'Wrap', 'Format', 'ReplaceWith', 'Replace', 'MatchKashida',
    # 'MatchDiacritics', 'MatchAlefHamza', 'MatchControl'],
    # varargs=None, keywords=None, defaults=())


    field_names = re.findall(r'{{.+?}}', self.template_path)
    while text_range.Find.Execute(text, False, False, True, False, False, True, 0, True, 'NewStr', 0):
      field_names.append(text_range.Text)
    self.app.Selection.HomeKey(6)
    if close:
      self.document.Close()

    if unique:
      unique_names = []
      for name in field_names:
        unique_names.extend(extract_field(name))
      return list(dedupe(unique_names))
    else:
      return field_names




  def render(self, info):

    self.info = info
    self.app.Visible = True
    for field in self.detect_required_fields(close=False, unique=False):
      self.app.Selection.HomeKey(6)
      text_range = self.document.Content
      val = evalute_field(field=field, info=info)
      # puts('val')
      if val in (None, ''):
        raise NoInfoKeyError('无法找到字段的值 {}'.format(field))


      while text_range.Find.Execute(field, False, False, False, False, False, True, 0, True, val, 2):
        # 'Replace' => 0 no replace ?
        # 'Replace' => 1 replace once ?
        # 'Replace' => 2 replace all ?
        # self.record('replace done {} ---> {}'.format(field.raw, val))
        pass
      # print(field_names)










def test_word_filler_detect_fields():


  t1 = os.getcwd() + '/test/test_templates/test_{{name}}_面积计算表.doc'
  t2 = os.getcwd() + '/test/test_templates/test_no_field_面积计算表.doc'
  from information import Information
  yaml_info = Information.from_yaml(os.getcwd() + '/test/测试单位.inf')


  filler = Filler(template_path=t1)
  filler.detect_required_fields(unique=False) | puts()
  filler.detect_required_fields(unique=True) | puts()


  filler = Filler(template_path=t2)
  filler.detect_required_fields(unique=False) | puts()
  filler.detect_required_fields(unique=True) | puts()


  filler.render(info=yaml_info)
  # filler.save(folder='path/to/', close=True)


def test_word_filler_render():


  t1 = os.getcwd() + '/test/test_templates/test_{{name}}_面积计算表.doc'
  t2 = os.getcwd() + '/test/test_templates/test_no_field_面积计算表.doc'
  from information import Information

  text = '''
    单位名称: 测试单位
    name: 测试单位name
    项目名称: 测试项目
    项目编号: 2015-项目编号-001
    面积90: 12345.600
    面积80: 12345.300
    地籍号: 1234567890010010000
    四至: 测试路1;测试街2;测试路3;测试街4
    土地坐落: 测试路以东,测试街以南
    area: 1000
    已设定值: value
  '''

  info = Information.from_string(text)

  filler = Filler(template_path=t1)
  filler.detect_required_fields(unique=False) | puts()
  filler.detect_required_fields(unique=True) | puts()

  filler.render(info=info)
  # filler.save(folder='path/to/', close=True)




def test_word_filler_render_and_save():

  t1 = os.getcwd() + '/test/test_templates/test_{{name}}_面积计算表.doc'
  from information import Information
  text = '''
    单位名称: 测试单位
    name: 测试单位name
    项目名称: 测试项目
    项目编号: 2015-项目编号-001
    面积90: 12345.600
    面积80: 12345.300
    地籍号: 1234567890010010000
    四至: 测试路1;测试街2;测试路3;测试街4
    土地坐落: 测试路以东,测试街以南
    area: 1000
    已设定值: value
  '''
  info = Information.from_string(text)
  filler = Filler(template_path=t1)

  filler.render(info=info)
  filler.save(folder=os.getcwd() + '/test/test_output', close=True)








def test_jinja():


  from information import Information

  text = '''
    单位名称: 测试单位
    项目名称: 测试项目
    项目编号: 2015-项目编号-001
    面积90: 12345.600
    面积80: 12345.300
    地籍号: 1234567890010010000
    四至: 测试路1;测试街2;测试路3;测试街4
    土地坐落: 测试路以东,测试街以南
    area: 1000
    已设定值: value
  '''
  info = Information.from_string(text)




  template = '''
  title
  {{单位名称}}                           -> 测试单位
  {{100 - 98}}                           -> 2
  {{'%.2f' | format(面积90 - 面积80)}}   -> 0.30
  {{面积80|round(1)}}                    -> 12345.3

  {{'%.2f'| format(area)}}               -> 1000.00
  {{'%.1f'| format(area)}}               -> 1000.0
  {{'%.4f'| format(area)}}               -> 1000.0000
  {{ "Hello " ~ 单位名称 ~ " !" }}        -> Hello 测试单位 !

  {{ "Hello World"|replace("Hello", "Goodbye") }} -> Goodbye World
  {{ "aaaaargh"|replace("a", "d'oh, ", 2) }} -> d'oh, d'oh, aaargh
  {{area}}
  {area}
  {{未设定值|d(default_value='default')}}  -> default
  {{已设定值|d(default_value='default')}}  -> value

  {{四至|truncate(12, killwords=True)}}  -> 测试路1;测试街2;测试...





  '''

  t = Template(template)
  result = t.render(**info.content)
  for line in datalines(result):
    if '->' in line:
      rendered, spec = line.split('->')
      if rendered.strip() != spec.strip():
        print('! ', line)

  # filler.save(folder='path/to/', close=True)
























####### ##   ## ###### ####### ##      ####### ###### ##      ##      ####### ######
##       ## ## ###     ##      ##      ##        ##   ##      ##      ##      ##   ##
######    ###  ##      ######  ##      ######    ##   ##      ##      ######  ######
##       ## ## ###     ##      ##      ##        ##   ##      ##      ##      ##  ##
####### ##   ## ###### ####### ####### ##      ###### ####### ####### ####### ##   ##




class ExcelFiller:
  def __init__(self, template_path, app):
    self.template_path = template_path
    self.app = app
    self.document = self.app.Workbooks.Open(template_path)


  def detect_required_fields(self, close=True, unique=False):

    # self.document = self.app.Workbooks.Open(self.template_path)
    sheet = self.document.WorkSheets.Item(1)


    field_names = re.findall(r'{{.+?}}', self.template_path)
    for cell in self.field_cells(sheet):
      for match in re.findall(r'{{.+?}}', cell.Value):
        field_names.append(match)

    if close:
      self.document.Close()

    if unique:
      unique_names = []
      for name in field_names:
        unique_names.extend(extract_field(name))
      return list(dedupe(unique_names))
    else:
      return field_names



  def used_cells(self, sheet):
    for cell in sheet.UsedRange.Cells:
      yield cell

  def field_cells(self, sheet):
    for cell in self.used_cells(sheet):
      value = str(cell.Value)
      if value and re.match(r'.*?{{.+?}}', value):
        # puts(value)
        yield cell



  def render(self, info):


    self.info = info
    self.app.Visible = True
    sheet = self.document.WorkSheets.Item(1)
    for cell in self.field_cells(sheet):
      cell_string = cell.Value
      cell.Value = evalute_field(cell_string, info)









def test_excel_filler_detect_fields():


  t1 = os.getcwd() + '/test/test_templates/_test_{{项目名称}}-申请表.xls'

  filler = Filler(template_path=t1)
  filler.detect_required_fields(unique=False) | puts()
  filler.detect_required_fields(unique=True) | puts()


def test_excel_filler_render():


  t1 = os.getcwd() + '/test/test_templates/_test_{{项目名称}}-申请表.xls'
  from information import Information

  text = '''
    单位名称: 测试单位
    name: 测试单位name
    项目名称: 测试项目
    项目编号: 2015-项目编号-001
    面积90: 12345.600
    面积80: 12345.300
    地籍号: 1234567890010010000
    四至: 测试路1;测试街2;测试路3;测试街4
    土地坐落: 测试路以东,测试街以南
    area: 1000
    已设定值: value
  '''

  info = Information.from_string(text)

  filler = Filler(template_path=t1)

  filler.render(info=info)




def test_excel_filler_render_and_save():

  t1 = os.getcwd() + '/test/test_templates/_test_{{项目名称}}-申请表.xls'

  from information import Information
  text = '''
    单位名称: 测试单位
    name: 测试单位name
    项目名称: 测试项目
    项目编号: 2015-项目编号-001
    面积90: 12345.600
    面积80: 12345.300
    地籍号: 1234567890010010000
    四至: 测试路1;测试街2;测试路3;测试街4
    土地坐落: 测试路以东,测试街以南
    area: 1000
    已设定值: value
  '''
  info = Information.from_string(text)
  filler = Filler(template_path=t1)

  filler.render(info=info)
  filler.save(folder=os.getcwd() + '/test/test_output', close=True)















































 ######  #####  ######  ####### ###### ##      ##      ####### ######
###     ##   ## ##   ## ##        ##   ##      ##      ##      ##   ##
##      ####### ##   ## ######    ##   ##      ##      ######  ######
###     ##   ## ##   ## ##        ##   ##      ##      ##      ##  ##
 ###### ##   ## ######  ##      ###### ####### ####### ####### ##   ##


class AutoCADFiller:
  def __init__(self, template_path, app):
    self.template_path = re.sub('/', '\\\\', template_path)
    self.app = app
    self.document = self.app.Documents.Open(template_path)


  def detect_required_fields(self, close=True, unique=False):

    field_names = re.findall(r'{{.+?}}', self.template_path)

    for en in self.field_text_entities():
      for match in re.findall(r'{{.+?}}', en.TextString):
        field_names.append(match)

    if unique:
      unique_names = []
      for name in field_names:
        unique_names.extend(extract_field(name))
      return list(dedupe(unique_names))
    else:
      return field_names

  def field_text_entities(self):
    msitem = self.document.ModelSpace.Item
    entities_count = self.document.ModelSpace.Count
    for i in range(entities_count):
      entity = msitem(i)
      if entity.EntityName[4:] in ['Text', ]:
        yield entity


  def render(self, info):


    self.info = info
    self.app.Visible = True
    for en in self.field_text_entities():
      val = evalute_field(field=en.TextString, info=info)
      if val in (None, ''):
        raise NoInfoKeyError('无法找到字段的值 {}'.format(en.TextString))

      en.TextString = val








def test_cad_filler_detect_fields():

  t1 = os.getcwd() + '/test/test_templates/test_{{测试单位}}-宗地图.dwg'

  filler = Filler(template_path=t1)
  filler.detect_required_fields(unique=False) | puts()
  filler.detect_required_fields(unique=True) | puts()



def test_cad_filler_render():

  t1 = os.getcwd() + '/test/test_templates/test_{{测试单位}}-宗地图.dwg'

  from information import Information
  text = '''
    测试单位: 测试单位name
    # title: testtitle
    project: custom_project
    date: 0160202
    ratio: 1000
    landcode: 200
    area80: 123.4
    area90: 234.5
    地籍号: 10939512
  '''
  info = Information.from_string(text)
  filler = Filler(template_path=t1)
  filler.render(info=info)
  filler.save(folder=os.getcwd() + '/test/test_output', info=info, close=True)

  # TODO: if templates tring has multiple field and just missing 1 value,
  # in this case it cannot catch 'NoInfoKeyError'
  # template '{{a}}:::{{b}}' a=null, b= test -> ':::test' no raise exception here!






