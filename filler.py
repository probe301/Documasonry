
from pylon import AutoDelegator
from pylon import puts
import re
import os
from pylon import dedupe
import win32com.client
from jinja2 import Template



def extract_field(field):
  # {{name | lower}}
  return field[2:-2].split('|')[0].strip()


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

    elif template_path.endswith(('.dwg', )):
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









  def output_name(self):
    tmpl_name = re.split(r'\/|\\', self.tmpl_path)[-1]
    output_name = tmpl_name[:]
    # print(tmpl_name)
    for match in re.finditer(r'{.+?}', tmpl_name):
      field_name = match.group()
      if Field(field_name).calculate(self.content):
        output_name = re.sub(field_name, Field(field_name).calculate(self.content), output_name)
    return output_name
















class NoInfoKeyError(Exception):
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



  def detect_required_fields(self, close=True, unique=False):

    self.document = self.app.Documents.Open(self.template_path)
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
      return list(dedupe(extract_field(name) for name in field_names))
    else:
      return field_names




  def render(self, info):
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
    # self.record('all fields replace done')
      # w.Selection.Find.Execute(OldStr, False, False, False, False, False, True, 1, True, NewStr, 2)



  def save(self, folder, close=True):
    pass




  def save(self, close=True):
    output_name = self.output_name
    # print(tmpl_name)
    workspace = self.workspace
    if (workspace[-1] == '/' or workspace[-1] == '\\'):
      self.output_path = workspace + output_name
    else:
      self.output_path = workspace + '/' + output_name

    # print('!!',self.output_path)
    if os.path.exists(self.output_path):
      fix = time.strftime('.backup-%Y-%m-%d-%H%M%S')
      os.rename(self.output_path, fix.join(os.path.splitext(self.output_path)))
    try:
      self.document.SaveAs(self.output_path)
    except Exception as e:
      self.record('Word Filler can not save document: <{}>'.format(self.output_path), important=True)
    if close:
      self.document.Close()
    # self.record('成功填充了文档 <{}>: \n  <{}>'.format(self.output_name, self.output_path), important=True)





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
  单位名称: 测试单位ANSI
  项目名称: 测试项目
  项目编号: 2015-项目编号-009
  面积90: 10099.900
  area90: 20000.22
  area80: 30000.33
  面积80: 10088.8000
  地籍号: 1234567890010010000
  四至: 测试路1;测试街2;测试路3;测试街4
  土地坐落: 测试路以东,测试街以南
  name: 测试单位ANSI
  '''

  info = Information.from_string(text)

  filler = Filler(template_path=t1)
  filler.detect_required_fields(unique=False) | puts()
  filler.detect_required_fields(unique=True) | puts()

  filler.render(info=info)
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







 ######  #####  ######  ####### ###### ##      ##      ####### ######
###     ##   ## ##   ## ##        ##   ##      ##      ##      ##   ##
##      ####### ##   ## ######    ##   ##      ##      ######  ######
###     ##   ## ##   ## ##        ##   ##      ##      ##      ##  ##
 ###### ##   ## ######  ##      ###### ####### ####### ####### ##   ##


class AutoCADFiller:
  def __init__(self, template_path, app):
    self.template_path = template_path
    self.app = app
