
import os
import time
import re
import io
import yaml
from collections import OrderedDict
from pylon import puts



class IncludeOrderedLoader(yaml.Loader):
  ''' yaml loader
      以有序 dict 替代默认 dict
      值为 !include 开头时, 嵌套另一个 yaml

        -- main.yaml
        key_normal: [foo, bar]
        key_included: !include 'another.yaml'

        -- another.yaml
        foo: bar
        bar: baz

        -- nested result
        key_normal: [foo, bar]
        key_included:
          foo: bar
          bar: baz

      !include 可以是绝对路径或相对路径
      如果嵌套太深, 可能遇到相对路径错乱的问题
  '''
  def __init__(self, stream):
    super(IncludeOrderedLoader, self).__init__(stream)
    self.add_constructor(yaml.resolver.BaseResolver.DEFAULT_MAPPING_TAG,
                         self._construct_mapping)
    self.add_constructor('!include', self._include)
    self._root = os.path.split(stream.name)[0]

  def _include(self, loader, node):
    filename = os.path.join(self._root, self.construct_scalar(node))
    try:
      f = open(filename, 'r', encoding='utf-8').read()
      encoding = 'utf-8'
    except UnicodeDecodeError:
      encoding = 'gbk'
    f = open(filename, 'r', encoding=encoding)
    return yaml.load(f, IncludeOrderedLoader)

  def _construct_mapping(self, loader, node):
    loader.flatten_mapping(node)
    return OrderedDict(loader.construct_pairs(node))

def yaml_load(stream, loader=None):
  '''按照有序字典载入yaml 支持 !include'''
  if loader is None:
    loader = IncludeOrderedLoader
  return yaml.load(stream, loader)







class InfoText:
  """ InfoText

  存放解析文本形式的键值对信息
  基于 yaml 但对格式更加宽容

  例
      name: foo
      age: 56
      语言: 中文
      # 此行为注释
      phone: [12312412, 21414124]

  可以用 utf8 gbk 编码存储
  每个键值对之间可以用冒号, 也可以用等号, 周围空格不强制要求

  例
      name=foo
      age :56
      语言 = 中文
      # 此行为注释
      phone : [12312412, 21414124]

  """



  @classmethod
  def from_yaml(cls, path):
    try:
      with open(path, 'r', encoding='utf-8') as f:
        text = f.read()
    except UnicodeDecodeError:
      with open(path, 'r', encoding='gbk') as f:
        text = f.read()
    text = cls.fix_colon(text)

    virtual_file = io.StringIO(text)  # 虚拟文件可以提供文件名
                                      # yaml loader 里会用到
    virtual_file.name = path
    return cls(yaml_load(virtual_file))

  @classmethod
  def from_string(cls, text):
    text = cls.fix_colon(text)
    virtual_file = io.StringIO(text)
    virtual_file.name = 'virtual'
    return cls(yaml_load(virtual_file))

  @staticmethod
  def fix_colon(text):
    result = ''
    for line in text.splitlines(keepends=True):
      line = re.sub(pattern=r' ?(:|=) ?', repl=r': ', string=line, count=1)
      result += line
    return result

  def __init__(self, content):
    if content is None:
      self.content = OrderedDict()
    else:
      self.content = content

  def __str__(self):
    text = '\n  '.join('{}: {}'.format(k, v) for k, v in self.content.items())
    if text.strip():
      return '<InfoText>\n  {}'.format(text)
    else:
      return '<InfoText>\n  {}'.format('(empty)')

  def to_yaml_string(self):
    return '\n'.join('{}: {}'.format(k, v) for k, v in self.content.items())

  def get(self, key):
    return self.content.get(key) or self.additional_key(key)

  def additional_key(self, key):
    '''未写入文本配置的信息

    日期时间:
    查询含有 '日期', 'current_date' 等文字的 key 时,
    如果没有显式声明日期, 'default' 中也没有日期, 则取当前日期

    prototype:
    如果含有名为 'default', 值为 dict 的 key/value 配对,
    则尝试从中提取相应 key 的值

    '''
    default_dict = self.content.get('default')
    if default_dict:
      if isinstance(default_dict, dict) and default_dict.get(key):
        return default_dict.get(key)

    if any(part in key for part in 'current,date,日期,年,月,日'.split(',')):
      return self.additional_date_key(key)

    return None

  def additional_date_key(self, key):
    year = str(time.strftime('%Y'))
    month = str(time.strftime('%m'))
    day = str(time.strftime('%d'))
    if key in ('current_date', '日期', '当前日期'):
      return '{}年{}月{}日'.format(year, month, day)
    if key in ('current_year', '年'):
      return year
    if key in ('current_month', '月'):
      return month
    if key in ('current_day', '日'):
      return day
    return None

  def merge(self, other):
    ''' 合并两个info, 未额外处理 key<default>
        other 中的 Value None 不会覆盖 self 中的已有值'''
    # self.content.update(other.content)
    for k, v in other.content.items():
      if k in self.content and v is None:
        continue
      self.content[k] = v
    return self























def test_info_yaml():
  info = InfoText.from_yaml(os.getcwd() + '/test/测试单位.inf')
  print(str(info))
  info = InfoText.from_yaml(os.getcwd() + '/test/测试单位ansi.inf')
  print(str(info))


def test_info_string_spec():
  text = '''
  单位名称: 测试单位
  项目名称: 测试项目
  面积: 10000.11
  '''
  info = InfoText.from_string(text)
  print(str(info))


def test_info_string_empty():
  text = '''
  # comment: comment

  '''
  info = InfoText.from_string(text)
  print(str(info))
  # print(info.content)

def test_info_string_some_not_space_after_colon():
  text = '''
  name=foo
  age :56
  语言 = 中文
  # 此行为注释
  phone : [12312412, 21414124]

  '''
  info = InfoText.from_string(text)
  print(str(info))
  # print(info.content)

def test_info_string_in_list():
  text = '''
    points_x[]: [100.1, 100.2, 100.3, 100.4]
    points_y[]: [200.1, 200.2, 200.3, 200.4]
  '''
  info = InfoText.from_string(text)
  print(str(info))


def test_info_nested_by_yaml_load():
  path = os.getcwd() + '/test/nested.inf'
  info = InfoText.from_yaml(path)
  puts(info)
  print('----')
  puts(info.content)
  print('----')
  print(yaml.dump(info.content))


def test_info_additional_keys():
  path = os.getcwd() + '/test/nested.inf'
  info = InfoText.from_yaml(path)
  puts(info.content)
  print(info.get('a'))
  print(info.get('ErrorKey'))
  print(info.get('foo'))           # from key<default>
  print(info.get('current_year'))  # key<default> contains this 1404
  print(info.get('current_date'))  # key<default> does not contain this,
                                   # use todays date



def test_infotext_merge():

  text1 = '''
  单位名称 =name1
  项目名称: name2
  key1:
  key2:
  key3: 456

  '''
  info1 = InfoText.from_string(text1)

  text2 = '''
    项目名称: change1
    单位名称: change2
    key1:
    key2: 123
    key3:
  '''
  info2 = InfoText.from_string(text2)
  print(str(info1))
  print(str(info2))

  info1.merge(info2)
  print(info1)






def test_figlet():
  from pylon import generate_figlet
  text = 'word excel autocad'.split(' ')
  for word in text:
    generate_figlet(word, fonts=['space_op', ])




def test_mkdir():
  import os
  os.mkdir('33/11/22')

