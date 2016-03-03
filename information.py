
import os
import time
import re
import yaml
from collections import OrderedDict
from pylon import puts

def fix_colon(text):
  result = ''
  for line in text.splitlines():
    result += re.sub(pattern=':', repl=': ', string=line, count=1) + '\n'
  return result

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
  def __init__(self, content, add_current_date=True):
    if content is None:
      self.content = OrderedDict()
    else:
      self.content = content

    if add_current_date:
      self.add_date()

  def get(self, key):
    return self.content.get(key)

  def add_date(self, date_value=None):
    date_string = date_value or str(time.strftime('%Y:%m:%d'))
    year, month, day = re.split(r'\:|\.|\-', date_string)
    # cases = {'年': year, 'year': year, 'Y': year, 'y': year,
    #          '月': month, 'month': month, 'M': month, 'm': month,
    #          '日': day, 'day': day, 'D': day, 'd': day}
    date_dict = {'current_date': date_string,
                 'current_date_cn': '{}年{}月{}日'.format(year, month, day),
                 'current_year': year,
                 'current_month': month,
                 'current_day': day
                 }
    self.content.update(date_dict)


  @classmethod
  def from_yaml(cls, path):
    try:
      with open(path, 'r', encoding='utf-8') as f:
        text = f.read()
    except UnicodeDecodeError:
      with open(path, 'r', encoding='gbk') as f:
        text = f.read()
    text = fix_colon(text)
    return cls(yaml_ordered_load(text))

  @classmethod
  def from_string(cls, text):
    text = fix_colon(text)
    return cls(yaml_ordered_load(text))


  def __str__(self):
    text = '\n  '.join('{}: {}'.format(k, v) for k, v in self.content.items())
    if text.strip():
      return '<Information>\n  {}'.format(text)
    else:
      return '<Information>\n  {}'.format('(empty)')

  def to_yaml_string(self):
    return '\n'.join('{}: {}'.format(k, v) for k, v in self.content.items())
























def test_info_yaml():
  info = Information.from_yaml(os.getcwd() + '/test/测试单位.inf')
  print(str(info))
  info = Information.from_yaml(os.getcwd() + '/test/测试单位ansi.inf')
  print(str(info))


def test_info_string_spec():
  text = '''
  单位名称: 测试单位
  项目名称: 测试项目
  面积: 10000.11
  四至: 测试路1;测试街2;测试路3;测试街4
  土地坐落: 测试路以东,测试街以南
  '''
  info = Information.from_string(text)
  print(str(info))


def test_info_string_empty():
  text = '''
  # 单位名称: 测试单位ANSI

  '''
  info = Information.from_string(text)
  print(str(info))
  print(info.content)

def test_info_string_some_not_space_after_colon():
  text = '''
  单位名称:测试单位
  项目名称: 测试项目
  土地坐落:测试路以东,测试街以南
  '''
  info = Information.from_string(text)
  print(str(info))
  print(info.content)

def test_info_string_in_list():
  text = '''
    项目名称: test1
    单位名称: test2
    points_x[]: [100.1, 100.2, 100.3, 100.4]
    points_y[]: [200.1, 200.2, 200.3, 200.4]
    lengths[]: [10, 15, 20, 30]
    radius[]: [0, 0, 5.5, 0]
    日期: today
  '''
  info = Information.from_string(text)
  print(str(info))












def test_figlet():
  from pylon import generate_figlet
  text = 'word excel autocad'.split(' ')
  for word in text:
    generate_figlet(word, fonts=['space_op', ])




def test_mkdir():
  import os
  os.mkdir('33/11/22')



