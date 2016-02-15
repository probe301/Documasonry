

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

  index success
  """
  def __init__(self, target_path, template_paths):
    self.template_paths = template_paths
    self.target_path = target_path
    # self.info

  def generate(self, info, save=True, add_index=True):

    for tmpl in self.template_paths:

      filler = Filler(template_path=tmpl)
      filler.render(info=info)
      if save:
        filler.save(info=info, folder=self.target_path, close=True)


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









