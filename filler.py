
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



class NoInfoKeyError(Exception):
  pass

class SaveDocumentError(Exception):
  pass

class AutoCADCustomFieldError(Exception):
  pass

class ExcelFieldError(Exception):
  pass


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

      filler = Filler(template_path, output_folder)
      filler.detect_required_fields()
      filler.render(info=yaml_info)
      filler.save(info=yaml_info, close=True)

  """
  def __init__(self, template_path, output_folder):
    self.template_path = template_path
    self.output_folder = output_folder
    self.info = None
    self.output_name = None
    if template_path.endswith(('.xls', '.xlsx')):
      excel = win32com.client.Dispatch('Excel.Application')
      # excel.Visible = False
      self.app_name = 'Office Excel'
      filler_app = ExcelFiller(template_path, output_folder, app=excel)
    elif template_path.endswith(('.doc', '.docx')):
      word = win32com.client.Dispatch('Word.Application')
      # word.Visible = False
      self.app_name = 'Office Word'
      filler_app = WordFiller(template_path, output_folder, app=word)
    elif template_path.endswith(('.dwg', '.dxf')):
      cad = win32com.client.Dispatch('AutoCAD.Application')
      # cad.Visible = False
      self.app_name = 'Autodesk AutoCAD'
      filler_app = AutoCADFiller(template_path, output_folder, app=cad)

    self.delegates = [filler_app] # delegates render and detect_required_filds


  def __str__(self):
    return '<Filler type={} template_path={}>'.format(self.app_name, self.template_path)

  def save(self, info, close=True, prefix=''):
    self.output_name = prefix + evalute_field(os.path.basename(self.template_path), info)
    output_path = os.path.join(self.output_folder, self.output_name)
    output_path = output_path.replace('\\', '/')
    output_path = output_path.replace('/', '\\')
    if os.path.exists(output_path):
      fix = time.strftime('.backup-%Y%m%d-%H%M%S')
      os.rename(output_path, fix.join(os.path.splitext(output_path)))
    try:
      self.document.SaveAs(output_path)
      puts('save document done - output_path')
    except Exception:
      raise
      t = 'Word Filler can not save document: <{}>'.format(output_path)
      raise SaveDocumentError(t)

    if close:
      self.document.Close()
























##   ##  #####  ######  ######
##   ## ##   ## ##   ## ##   ##
## # ## ##   ## ######  ##   ##
### ### ##   ## ##  ##  ##   ##
##   ##  #####  ##   ## ######

class WordFiller:
  def __init__(self, template_path, output_folder, app):
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


  filler = Filler(template_path=t1, output_folder=os.getcwd() + '/test/test_output')
  filler.detect_required_fields(unique=False) | puts()
  filler = Filler(template_path=t1, output_folder=os.getcwd() + '/test/test_output')
  filler.detect_required_fields(unique=True) | puts()


  filler = Filler(template_path=t2, output_folder=os.getcwd() + '/test/test_output')
  filler.detect_required_fields(unique=False) | puts()
  filler = Filler(template_path=t2, output_folder=os.getcwd() + '/test/test_output')
  filler.detect_required_fields(unique=True) | puts()


  filler = Filler(template_path=t1, output_folder=os.getcwd() + '/test/test_output')
  filler.render(info=yaml_info)


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

  filler = Filler(template_path=t1, output_folder=os.getcwd() + '/test/test_output')
  filler.detect_required_fields(unique=False) | puts()
  filler = Filler(template_path=t1, output_folder=os.getcwd() + '/test/test_output')
  filler.detect_required_fields(unique=True) | puts()

  filler = Filler(template_path=t1, output_folder=os.getcwd() + '/test/test_output')
  filler.render(info=info)




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
  filler = Filler(template_path=t1, output_folder=os.getcwd() + '/test/test_output')

  filler.render(info=info)
  filler.save(info=info, close=True)








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



def test_jinja_edge_cases():
  from information import Information
  text = '''
    codes: 1231234000050280000
    borders: 空地;空地;空地;空地
  '''
  template = '''
    {{(codes | string)[12:15]}}-01
    {{(codes | string)[12:15]}}-01
  '''

  info = Information.from_string(text)
  t = Template(template)
  result = t.render(**info.content)
  result | puts()






















####### ##   ## ###### ####### ##
##       ## ## ###     ##      ##
######    ###  ##      ######  ##
##       ## ## ###     ##      ##
####### ##   ## ###### ####### #######


class ExcelFiller:
  '''
  特殊字段
  - insert list data

  yaml_text =
    项目名称: test1
    单位名称: test2
    points_x: [100.1, 100.2, 100.3, 100.4]  # 数据为列表
    points_y: [200.1, 200.2, 200.3, 200.4]
    lengths: [10, 15, 20, 30]
    radius: [0, 0, 5.5, 0]

  template excel cell label =
    {points_x[20]} # [20] 表示本页最多容纳20行, 对应数据必须为列表
    {points_y[20]}
    {lengths[15]}

  '''
  def __init__(self, template_path, output_folder, app):
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
    self.fill_array_cells(info)

  def array_cells(self, sheet):
    '''返回所有需填充array的单元格'''
    for cell in self.used_cells(sheet):
      value = cell.Value
      if value and re.match(r'.*?\{.+\}', str(value)):
        if '[' in value and ']' in value:
          # puts('array_cells_ value')
          yield cell


  def fill_array_cells(self, info):
    '''填充列表部分, 单元格可能是普通的, 可能是纵向合并过的'''
    sheet = self.document.Worksheets.Item(1)
    for cell in list(self.array_cells(sheet)):
      field_text = cell.Value
      field_label = field_text.split('[')[0][1:]
      field_max_length = int(field_text.split('[')[1][:-2])
      # puts('field_text field_label field_max_length')
      value_list = info.get(field_label)
      if value_list is None:
        raise NoInfoKeyError('没有找到列表数据 <{}>'.format(field_label))
      if isinstance(value_list, str) or not isinstance(value_list, list):
        raise NoInfoKeyError('列表数据 <{}> 的值必须为数组/序列'.format(field_label))
      if field_max_length < len(value_list):
        raise ExcelFieldError('无法填充序列 <{}> 表格容量太小({}), 数据过多({})'.format(field_text, field_max_length, len(value_list)))

      row = cell.Row
      col = cell.Column
      span_rows = cell.MergeArea.Rows.Count
      # 有些坐标是界址点和界址线信息交错排布,
      # 这时每个 cell 占两行高度, 点线错开一行
      # 使用 cell.MergeArea 判断是否跨行合并过
      for elem in value_list:
        sheet.Cells(row, col).Value = elem
        row += span_rows
      # puts('replace done {} ---> {}'.format(field.raw, val_list))

  def get_sheet(self, index=0):
    return self.document.Worksheets.Item(index+1)

  def copy_sheet(self, source=0, after=-1, label='page'):
    original_sheet = self.get_sheet(source)
    if after < 0:
      after += self.document.WorkSheets.Count
    original_sheet.Copy(None, self.get_sheet(after))
    # VBA 接受 Before After 两种参数
    # python 调用时不能用关键字参数, 必须按照位置传递
    new_sheet = self.get_sheet(after+1)
    new_sheet.name = label
    return new_sheet

  def delete_sheet(self, index=0):
    self.document.Worksheets.Item(index+1).Delete()








def test_excel_filler_detect_fields():
  t1 = os.getcwd() + '/test/test_templates/_test_{{项目名称}}-申请表.xls'
  filler = Filler(template_path=t1, output_folder=os.getcwd() + '/test/test_output')
  filler.detect_required_fields(unique=False) | puts()
  filler = Filler(template_path=t1, output_folder=os.getcwd() + '/test/test_output')
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
  filler = Filler(template_path=t1, output_folder=os.getcwd() + '/test/test_output')
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
  filler = Filler(template_path=t1, output_folder=os.getcwd() + '/test/test_output')

  filler.render(info=info)
  filler.save(info=info, close=True)


def test_excel_fill_subtable():
  from information import Information
  from pylon import relative_path

  text = '''
    项目名称: 测试项目名称
    date: 20150101
    姓名: [100.1, 100.2, 100.3, 100.4]
    年龄: [1,2,3,1,4,1,2,4,1,1,2,3,None,4,]
    电话: [200.1, 200.2, 200.3, 200.4]
  '''
  info = Information.from_string(text)
  t1 = relative_path('test/test_templates/{{项目名称}}-索引表.xls')

  filler = Filler(template_path=t1, output_folder=os.getcwd() + '/test/test_output')

  filler.render(info=info)
  # filler.save(info=info, close=False)






















 #####  ##   ## ####### #####   ######  #####  ######
##   ## ##   ##    ##  ##   ## ###     ##   ## ##   ##
####### ##   ##    ##  ##   ## ##      ####### ##   ##
##   ## ##   ##    ##  ##   ## ###     ##   ## ##   ##
##   ##  #####     ##   #####   ###### ##   ## ######

class AutoCADFiller:
  """ AutoCAD filler 特殊 method / field
  - insert block {{地形}} to back layer
  - 位置校正
    原始模板在(0, 0)处 需要调整到界址线图形所在位置
    polygon边框确定:
      参数 padding比例, 是否方形
      先定位中心 c, 边界框 w, h
      padding = max(w, h) * padding_ratio
      方形 ? 边界框长边 + padding : 边界框长边短边分别 + padding

  yaml_text = '''
    项目名称: test1
    单位名称: test2
    地籍号: 110123122
    name: sjgisdgd
    面积90: 124.1
    面积80: 234.2
    宗地dwg: zd.dwg
    地形dwg: dx.dwg
    日期: today
  '''


  """


  def __init__(self, template_path, output_folder, app):
    self.template_path = re.sub('/', '\\\\', template_path)
    self.output_folder = output_folder
    self.app = app
    self.document = self.app.Documents.Open(template_path)


  def detect_required_fields(self, close=True, unique=False):
    field_names = re.findall(r'{{.+?}}', self.template_path)
    for en in self.text_entities():
      for match in re.findall(r'{{.+?}}', en.TextString):
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


  def text_entities(self):
    '''CAD中文字实体'''
    return list(self.entities(kinds='Text'))


  def border_entities(self, border_layer):
    '''CAD中特定图层的Polygon实体'''
    return list(self.entities(kinds='Polyline', layers=border_layer))


  def entities(self, kinds=None, layers=None):
    msitem = self.document.ModelSpace.Item
    entities_count = self.document.ModelSpace.Count
    if isinstance(kinds, str):
      kinds = [kinds]
    if isinstance(layers, str):
      layers = [layers]
    for i in range(entities_count):
      entity = msitem(i)
      if kinds and entity.EntityName[4:] not in kinds:
        continue
      if layers and entity.Layer not in layers:
        continue
      yield entity


  def render(self, info):
    self.info = info
    self.app.Visible = True
    if self.info.get('target_position'):
      # 如含有 target_position 字段 编辑前需要调整模板全体 object 位置
      target_position = self.info.content['target_position']
      if isinstance(target_position, list) and len(target_position) == 4:
        # target_center and target_size [566371.2180, 4340932.6223, 202.3, 202.3]
        self.fix_position(target_center=target_position[:2],
                          target_size=target_position[2:])
      elif isinstance(target_position, str):
        if not os.path.isfile(target_position):
          target_position = os.path.join(self.output_folder, target_position)

        self.insert_block(dwg_path=target_position)
        last_entity = list(self.entities(kinds='BlockReference'))[0] | puts()
        target_center = self.mid_point(last_entity)
        target_size = self.bounding_box_size(last_entity)
        # print(last_entity.)
        last_entity.Delete()
        self.fix_position(target_center=target_center, target_size=target_size)



    for en in self.text_entities():
      val = evalute_field(field=en.TextString, info=info)
      if val in (None, ''):
        raise NoInfoKeyError('无法找到字段的值 {}'.format(en.TextString))

      if en.TextString.startswith('{{') and en.TextString.endswith('dwg}}'):
        # block field syntax should insert dwg block
        if not os.path.isfile(val):
          val = os.path.join(self.output_folder, val)
        self.insert_block(dwg_path=val)
        en.Delete()
      else:
        en.TextString = val
    self.document.SendCommand('zoom e ')




  def fix_position(self, target_center, target_size):
    ''' 依据特定的存放框架的图层来平移缩放模板位置
    template 中 border_source 层中的第一个矩形框定义模板原始位置
    平移方法: 两个矩形框中心作为平移矢量
    缩放方法: 两个矩形框width比率和height比率中较大者作为缩放因子
             目标矩形框中心作为缩放中心

    '''
    source_borders = self.border_entities('border_source')
    if not source_borders:
      raise AutoCADCustomFieldError('cannot find source border polyline')
    source_entity = source_borders[0]
    source_center = self.mid_point(source_entity)
    source_size = self.bounding_box_size(source_entity)


    offset = target_center[0] - source_center[0], target_center[1] - source_center[1]
    scale_factor = max([target_size[0]/source_size[0], target_size[1]/source_size[1]])

    move_cmd = 'move all  d {},{} '.format(offset[0], offset[1])
    scale_cmd = 'scale all  {},{} {} '.format(target_center[0], target_center[1], scale_factor)
    self.document.SendCommand(move_cmd)
    self.document.SendCommand(scale_cmd)
    self.change_layer(source_entity, 'border_destination')





  def change_layer(self, entity, layer_name):
    for lay in self.document.Layers:
      if layer_name == lay.Name:
        break
    else:
      self.document.Layers.Add(layer_name)
    entity.Layer = layer_name


  def mid_point(self, entity):
    min_point, max_point = entity.GetBoundingBox()
    return (min_point[0] + max_point[0])/2, (min_point[1] + max_point[1])/2


  def bounding_box_size(self, entity):
    min_point, max_point = entity.GetBoundingBox()
    return max_point[0] - min_point[0], max_point[1] - min_point[1]


  def insert_block(self, dwg_path):
    '''嵌入外部dwg图形
    block name 会被自动设为 dwg 文件名(不含扩展名)
    所以需要保证原图没有重名的block
    dwg_path 中允许空格'''
    if not os.path.isfile(dwg_path):
      raise AutoCADCustomFieldError('cannot open file {}'.format(dwg_path))
    insert_cmd = '-insert {dwg_path}\n0,0 1 1 0 '.format(dwg_path=dwg_path)
    self.document.SendCommand(insert_cmd)















def test_cad_filler_detect_fields():
  t1 = os.getcwd() + '/test/test_templates/test_{{测试单位}}-宗地图.dwg'
  filler = Filler(template_path=t1, output_folder=os.getcwd() + '/test/test_output')
  filler.detect_required_fields(unique=False) | puts()
  filler = Filler(template_path=t1, output_folder=os.getcwd() + '/test/test_output')
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
  filler = Filler(template_path=t1, output_folder=os.getcwd() + '/test/test_output')
  filler.render(info=info)
  filler.save(info=info, close=True)

  # TODO: if templates tring has multiple field and just missing 1 value,
  # in this case it cannot catch 'NoInfoKeyError'
  # template '{{a}}:::{{b}}' a=null, b= test -> ':::test' no raise exception here!


def test_cad_filler_fix_position():
  t1 = os.getcwd() + '/test/test_templates/test_cad_insert_block.dwg'

  filler = Filler(template_path=t1, output_folder=os.getcwd() + '/test/test_output')
  # filler.render()
  # filler.fix_position(target_center=(10000, 10000), target_size=(10, 20))
  filler.fix_position()


def test_cad_filler_insert_block():
  t1 = os.getcwd() + '/test/test_templates/test_cad_insert_block.dwg'
  block_path = os.getcwd() + '/test/test_templates/dixing.dwg'
  # block_path = os.getcwd() + '/test/test_templates/gh.dwg' # test block name already exists in cadfile
  # block_path = os.getcwd() + '/test/test_templates/gui  hua.dwg' # test space in file name
  filler = Filler(template_path=t1, output_folder=os.getcwd() + '/test/test_output')
  filler.insert_block(dwg_path=block_path)


def test_cad_filler_insert_block_from_yaml_config_relative_path():
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
    地形file: ../test_templates/dixing.dwg
  '''
  info = Information.from_string(text)
  t1 = os.getcwd() + '/test/test_templates/test_cad_insert_block.dwg'
  filler = Filler(template_path=t1, output_folder=os.getcwd() + '/test/test_output')
  filler.render(info=info)
  filler.save(info=info, close=True)





