












import os

import yaml
from collections import OrderedDict
from pylon import puts

def yaml_ordered_load(stream, Loader=None, object_pairs_hook=None):
  '''按照有序字典载入yaml'''
  if Loader is None:
    Loader = yaml.Loader
  if object_pairs_hook is None:
    object_pairs_hook = OrderedDict
  class OrderedLoader(Loader):
    pass
  def construct_mapping(loader, node):
    loader.flatten_mapping(node)
    return object_pairs_hook(loader.construct_pairs(node))
  OrderedLoader.add_constructor(yaml.resolver.BaseResolver.DEFAULT_MAPPING_TAG, construct_mapping)
  return yaml.load(stream, OrderedLoader)




class Information:
  """docstring for Information"""
  def __init__(self, content):
    self.content = content



  @classmethod
  def from_yaml(cls, path):
    try:
      with open(path, 'r', encoding='utf-8') as f:
        txt = f.read()
    except UnicodeDecodeError:
      with open(path, 'r', encoding='gbk') as f:
        txt = f.read()

    return cls(yaml_ordered_load(txt))


  @classmethod
  def from_string(cls, text):
    return cls(yaml_ordered_load(text))






  def __str__(self):
    return '\n'.join('{}: {}'.format(k, v) for k, v in self.content.items())












#############################
#############################
#############################

def test_11():
  puts(111, 22)
  puts('12412'.split('|')[0])

def test_info_yaml():
  info = Information.from_yaml(os.getcwd() + '/test/测试单位.inf')
  print(str(info))
  info = Information.from_yaml(os.getcwd() + '/test/测试单位ansi.inf')
  print(str(info))



def test_info_string():

  text = '''
  单位名称: 测试单位ANSI
  项目名称: 测试项目
  项目编号: 2015-项目编号-009
  面积90: 10000.1100
  面积80: 20000.22
  面积80: .1
  地籍号: 1234567890010010000
  四至: 测试路1;测试街2;测试路3;测试街4
  土地坐落: 测试路以东,测试街以南

  '''
  info = Information.from_string(text)
  print(str(info))














def test_figlet():

  from pylon import generate_figlet

  generate_figlet('excelfiller', fonts=['space_op', ])
  generate_figlet('wordfiller', fonts=['space_op', ])
  generate_figlet('autocadfiller', fonts=['space_op', ])



def test_mkdir():
  import os
  os.mkdir('33/11/22')



