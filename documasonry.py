

from pylon import puts
import re
import os

import time
from pylon import datalines
import pylon

from filler import Filler
from information import Information


######   #####   ###### ##   ## ##   ##  #####   ######  #####  ##   ## ######  ##   ##
##   ## ##   ## ###     ##   ## ### ### ##   ## ##      ##   ## ###  ## ##   ## ##   ##
##   ## ##   ## ##      ##   ## ## # ## #######  #####  ##   ## ## # ## ######   #####
##   ## ##   ## ###     ##   ## ##   ## ##   ##      ## ##   ## ##  ### ##  ##     ##
######   #####   ######  #####  ##   ## ##   ## ######   #####  ##   ## ##   ##    ##





class Documasonry(object):
  """docstring for Documasonry

  yaml basic config
  GUI
  如果某个模板被打开且编辑了部分字段怎么办
  - Word 和 Excel 会使用编辑中的模板, 输出包含编辑部分的成果
  - AutoCAD 会开新的只读实例, 按照原模板输出成果


  AutoCAD 特殊 field
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
    zdfile: zd.dwg
    地形file: dx.dwg
    日期: today

  '''



  """
  def __init__(self, target_path, template_paths):
    self.template_paths = template_paths
    self.target_path = target_path


  def generate(self, info, save=True, add_index=True):
    for i, tmpl in enumerate(self.template_paths, 1):
      filler = Filler(template_path=tmpl)
      filler.render(info=info)
      if save:
        prefix = '{:02d}-'.format(i) if add_index else ''
        filler.save(info=info, folder=self.target_path, close=True, prefix=prefix)

  def detect_required_fields(self):
    field_names = []
    for tmpl in self.template_paths:
      # tmpl | puts()
      filler = Filler(template_path=tmpl)
      field_names.extend(filler.detect_required_fields(close=True, unique=True))
    return list(pylon.dedupe(field_names))

  def load_config():
    '''config yaml
    template_require:
      - 'path_to_template field1 field2 field3...'
      - 'path_to_template field1 field2 field3...'
      - 'path_to_template field1 field2 field3...'
      - 'path_to_template field1 field2 field3...'



    '''
    pass



















def test_documasonry_detect_fields():
  template_paths = [os.getcwd() + '/test/test_templates/_test_{{项目名称}}-申请表.xls',
                    os.getcwd() + '/test/test_templates/test_{{name}}_面积计算表.doc',
                    os.getcwd() + '/test/test_templates/test_{{测试单位}}-宗地图.dwg',
                    os.getcwd() + '/test/test_templates/test_no_field_面积计算表.doc',
                    ]
  target_path = os.getcwd() + '/test/test_output'
  mason = Documasonry(target_path=target_path, template_paths=template_paths)
  mason.detect_required_fields() | puts()
  # [项目名称, 单位名称, 地籍号, name, 面积90, 面积80, area, 测试单位, title, project, date, ratio, landcode, area80, area90] <-- list length 15






def test_documasonry_generate():
  template_paths = [os.getcwd() + '/test/test_templates/_test_{{项目名称}}-申请表.xls',
                    os.getcwd() + '/test/test_templates/test_{{name}}_面积计算表.doc',
                    os.getcwd() + '/test/test_templates/test_{{测试单位}}-宗地图.dwg',
                    os.getcwd() + '/test/test_templates/test_no_field_面积计算表.doc',
                    ]
  target_path = os.getcwd() + '/test/test_output'
  mason = Documasonry(target_path=target_path, template_paths=template_paths)
  text = '''
    项目名称: test1
    单位名称: test2
    地籍号: 110123122
    name: sjgisdgd
    面积90: 124.1
    面积80: 234.2
    area: 124.2
    测试单位: testconm
    title: testtitle
    project: pro.
    date: 20124002
    ratio: 2000
    landcode: 235
    area80: 94923
    area90: 3257
  '''
  info = Information.from_string(text)
  mason.generate(info=info, save=True, add_index=True) | puts()









