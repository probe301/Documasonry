

from pylon import puts
import pylon
import re
import os
import sys
import time
# from pylon import datalines
import html
import urllib.parse
import logging
from filler import Filler
from information import Information
from documasonry import Documasonry
from PyQt4 import QtCore, QtGui, uic
from PyQt4.QtGui import (QApplication, QMessageBox, QCheckBox, QFileDialog, QWidget)







class QCommonTools(object):
  """PyQt 应 用 通 用 功 能
  """
  def __init__(self):
    super(QCommonTools, self).__init__()

  def clear_and_close(self, event):
    if event.key() == QtCore.Qt.Key_Escape:
      self.close()

  def popup(self, content="alart", title="提示", message_type='Critical'):
    dic = {
      'Critical': QMessageBox.Critical
    }
    msg = QMessageBox(dic[message_type], title, content)
    # 这种 alert 窗口必须一直置顶
    msg.setWindowFlags(QtCore.Qt.WindowStaysOnTopHint)
    if msg.exec_():
      return False

  def select_path(self, title='选择路径', current_path=''):
    options = QFileDialog.DontResolveSymlinks | QFileDialog.ShowDirsOnly
    directory = QFileDialog.getExistingDirectory(self, title,
                                                 current_path, options=options)
    return directory


  def select_file(self, title='选择文件', current_path='', ext=['txt']):
    accept_files = "Accept Files ({})".format('; '.join(['*.'+x for x in ext]))
    file_path = QFileDialog.getOpenFileName(self, title, current_path, accept_files)
    return file_path[0]

  def set_window_order(self, top=True):
    '''设置窗口是否置顶'''
    flags = QtCore.Qt.WindowFlags()
    flags |= QtCore.Qt.Dialog
    if top:
      flags |= QtCore.Qt.WindowStaysOnTopHint
    else:
      flags |= QtCore.Qt.WindowStaysOnBottomHint
    self.setWindowFlags(flags)
    self.show()















class XStream(QtCore.QObject):
  _stdout = None
  _stderr = None
  messageWritten = QtCore.pyqtSignal(str)
  def flush(self):
    pass
  def fileno(self):
    return -1
  def write(self, msg):
    if (not self.signalsBlocked()):
      self.messageWritten.emit(msg)
  @staticmethod
  def stdout():
    if (not XStream._stdout):
      XStream._stdout = XStream()
      sys.stdout = XStream._stdout
    return XStream._stdout
  @staticmethod
  def stderr():
    if (not XStream._stderr):
      XStream._stderr = XStream()
      sys.stderr = XStream._stderr
    return XStream._stderr


class QtHandler(logging.Handler):
  def __init__(self):
    logging.Handler.__init__(self)

  def emit(self, record):
    record = self.format(record)
    if record:
      # print(record)
      XStream.stdout().write('%s' % record)


class QLogger:
  """QLogger
  在 PyQt 的 text_browser Widget 中输出print()

  使用: PyQtWidget 需要有一个 TextBrowser 对象

  在 PyQtWidget 实例中混入 QLogger
    class Darter(QWidget, QLogger):
  在 __init__ 时设置 TextBrowser ID
    self.set_logger()
  """
  def __init__(self):
    super(QLogger, self).__init__()

  def set_logger(self, logger=None, text_browser_id='color_logger'):
    if not logger:
      logger = logging.getLogger(__name__)
    handler = QtHandler()
    handler.setFormatter(logging.Formatter("%(levelname)s: %(message)s"))
    logger.addHandler(handler)
    # logger.addHandler(logging.StreamHandler())
    logger.setLevel(logging.DEBUG)
    self.logger = logger

    text_browser = getattr(self, text_browser_id)
    XStream.stdout().messageWritten.connect(text_browser.append)
    XStream.stderr().messageWritten.connect(text_browser.append)
    self.logger_widget = text_browser

  def log(self, *text, level='INFO', color=None):
    color_dict = {'DEBUG': 'gray',
                  'INFO': 'darkgreen',
                  'WARNING': 'orange',
                  'ERROR': 'red',
                  'CRITICAL': 'red',
                  'SUCCESS': 'green'}
    stamp = time.strftime('%m/%d %H:%M:%S')
    if not color:
      color = color_dict.get(level.upper(), 'gray')
    text = [html.escape(str(t), quote=True) for t in text]
    ret = '{} <font color="{}">{}</font>'.format(stamp, color, ' '.join(text))
    self.logger_widget.append(ret)

  def debug(self, *text):
    self.log(*text, color='gray')
  def warn(self, *text):
    self.log(*text, color='orange')
  def error(self, *text):
    self.log(*text, color='red')
  def success(self, *text):
    self.log(*text, color='green')
  def info(self, *text):
    self.log(*text, color='black')















class DragInArea:
  """能接受拖动文件或文件夹的widget

  将文件或文件夹拖动到按钮上释放
  参数
  main_window = Qt 主程序
  accept_exts = 允许的后缀, 以逗号分隔的字符串, 设为 'folder' 则允许文件夹拖入
  accept_single_path = 设为 True 则只接受拖入的第一个路径
  callback = 回调函数, 传入 widget paths 为参数, 即按钮对象和目前已选择的路径
  """

  def __init__(self, widget_id, main_window,
               accept_exts=['txt', 'py'], accept_single_path=False,
               hover_color='edefff', normal_color='ffffff',
               callback=lambda x: print(x.selecting)

               ):
    self.widget = getattr(main_window.ui, widget_id)
    self.widget_id = widget_id
    self.main_window = main_window
    self.hover_color = hover_color
    self.normal_color = normal_color
    self.accept_single_path = accept_single_path
    self.callback = callback
    self.selecting = []
    self.default_path = ''
    if accept_exts == 'folder':
      self.accept_exts = 'folder'
    else:
      self.accept_exts = [s.lower() for s in accept_exts]

    self.set_drag_and_drop()


  def change_background(self, color):
    css = self.main_window.css + '\n\n#%s { background-color: #%s; }' % (self.widget_id, color)
    self.main_window.ui.setStyleSheet(css)


  def set_drag_and_drop(self):
    def drag_enter(event):

      file_paths = [urllib.parse.unquote(x.toString()[8:]) for x in event.mimeData().urls()]
      valid_paths = [p for p in file_paths if os.path.isdir(p)]
      valid_files = [p for p in file_paths if os.path.splitext(p)[1][1:].lower() in self.accept_exts]
      # puts('valid_files valid_paths', self.accept_exts)
      if (('folder' == self.accept_exts) and valid_paths) or valid_files:
        self.change_background(self.hover_color)
        # 'can' | puts()
        event.acceptProposedAction()


    def drag_leave(event):
      self.change_background(self.normal_color)


    def drag_move(event):
      event.accept()

    def drop(event):
      paths_all = [urllib.parse.unquote(x.toString()[8:]) for x in event.mimeData().urls()]
      file_paths = []
      for p in paths_all:
        if ('folder' == self.accept_exts) and os.path.isdir(p):
          file_paths.append(p)
        elif os.path.splitext(p)[1][1:].lower() in self.accept_exts:
          file_paths.append(p)

      if self.accept_single_path:
        file_paths = file_paths[:1]
      self.selecting = file_paths
      self.callback(self)
      self.change_background(self.normal_color)

    self.widget.setAcceptDrops(True)
    self.widget.dragEnterEvent = drag_enter
    self.widget.dropEvent = drop
    self.widget.dragMoveEvent = drag_move
    self.widget.dragLeaveEvent = drag_leave








class DocumasonryGUI(QWidget, QCommonTools, QLogger):
  ''' GUI

  ######   #####   ###### ##   ## ##   ##  #####   ######  #####  ##   ## ######  ##   ##
  ##   ## ##   ## ###     ##   ## ### ### ##   ## ##      ##   ## ###  ## ##   ## ##   ##
  ##   ## ##   ## ##      ##   ## ## # ## #######  #####  ##   ## ## # ## ######   #####
  ##   ## ##   ## ###     ##   ## ##   ## ##   ##      ## ##   ## ##  ### ##  ##     ##
  ######   #####   ######  #####  ##   ## ##   ## ######   #####  ##   ## ##   ##    ##

  '''

  def __init__(self, parent=None):
    super(QWidget, self).__init__(parent)
    self.ui = uic.loadUi('documasonry_gui.ui', self)
    self.set_bindings()
    self.set_logger(logger=None, text_browser_id='color_logger')


  def set_bindings(self):
    self.documasonry = Documasonry(output_path='', template_paths=[])
    self.init_templates_table()
    self.init_output_path_textedit()
    self.add_templates_from_config()
    self.rebuild_required_info_text(quick=True)
    self.set_window_order(top=True)
    self.keyPressEvent = self.clear_and_close
    self.set_templates_dropper()
    self.set_info_text_dropper()
    self.set_output_path_textedit_dropper()



    self.css = '''
    #templates_table { background-color: #ffffff; }
    #templates_table QCheckBox { padding: 3px auto 3px 10px; }
    #templates_table QCheckBox:hover { background: #E0ECF8; }
    '''
    self.ui.setStyleSheet(self.css)

  def init_output_path_textedit(self):
    path = self.documasonry.read_config().get('default_output_path')
    if path and os.path.isdir(path):
      self.output_path_textedit.setText(path)
    else:
      self.output_path_textedit.setText('C:/')

  def rebuild_required_info_text(self, quick):
    template_paths = self.get_templates_from_table()
    self.documasonry.set_template_paths(template_paths)
    required_text = self.documasonry.generate_required_fields_info_text(quick=quick)
    self.info_textedit.setPlainText(required_text)

  def set_output_path_textedit_dropper(self):
    def output_path_textedit_drop_done(dropper):
      # dropper | puts()
      self.output_path_textedit.setText(dropper.selecting[0])
    DragInArea(widget_id='output_path_textedit',
               main_window=self,
               accept_exts='folder',
               hover_color='ffede9',
               accept_single_path=True,
               callback=output_path_textedit_drop_done)

  def set_templates_dropper(self):
    def templates_drop_done(dropper):
      self.add_templates_from_browser(dropper.selecting)
    DragInArea(widget_id='templates_table',
               main_window=self,
               accept_exts='doc docx dwg xls xlsx txt'.split(' '),
               callback=templates_drop_done)


  def set_info_text_dropper(self):
    def info_text_drop_done(dropper):
      info_add = open(dropper.selecting[0], encoding='utf8').read()
      info_org = dropper.widget.toPlainText()
      info_text_combine = self.documasonry.combine_fields_info_text(info_org, info_add)
      dropper.widget.setPlainText(info_text_combine)
    DragInArea(widget_id='info_textedit',
               main_window=self,
               accept_exts='txt inf ini yaml md'.split(' '),
               accept_single_path=True,
               hover_color='fff8ed',
               callback=info_text_drop_done)

  def get_table_items(self, table, only_checked=False):
    for i in range(table.rowCount()):
      checker = table.cellWidget(i, 0)
      if not only_checked or checker.isChecked():
        yield checker


  def init_templates_table(self):
    templates_table = self.templates_table
    templates_table.verticalHeader().setVisible(False)
    templates_table.horizontalHeader().setVisible(False)
    templates_table.setColumnCount(1)
    templates_table.setRowCount(0)


  def add_templates_from_browser(self, file_paths):
    templates_table = self.templates_table
    for file_path in file_paths:
      exist_checkers = self.get_table_items(table=templates_table)
      if os.path.abspath(file_path) in [c.template_path for c in exist_checkers]:
        continue
      self.add_checker(table=templates_table,
                       template_path=file_path,
                       selected=True)

  def add_templates_from_config(self):
    items = self.documasonry.read_config()['default_templates']
    items = list(pylon.dedupe(items, key=lambda x: x['file']))
    templates_table = self.templates_table
    for item in items:
      self.add_checker(table=templates_table,
                       template_path=item['file'],
                       selected=item['selected'])


  def add_checker(self, table, template_path, selected):
    name = os.path.basename(template_path)
    row_count = table.rowCount()
    table.insertRow(row_count)
    checker = QCheckBox(name)
    checker.setChecked(selected)
    checker.template_path = os.path.abspath(template_path)
    checker.stateChanged.connect(lambda x: self.rebuild_required_info_text(quick=True))
    table.setCellWidget(row_count, 0, checker)
    table.resizeColumnsToContents()
    return checker


  def get_templates_from_table(self):
    checkers = self.get_table_items(table=self.templates_table, only_checked=True)
    return [checker.template_path for checker in checkers]


  @QtCore.pyqtSlot()
  def on_select_all_templates_button_clicked(self):
    for checker in self.get_table_items(table=self.templates_table):
      checker.setChecked(True)

  @QtCore.pyqtSlot()
  def on_invert_select_templates_button_clicked(self):
    for checker in self.get_table_items(table=self.templates_table):
      checker.setChecked(not checker.isChecked())

  @QtCore.pyqtSlot()
  def on_detect_required_fields_button_clicked(self):
    self.rebuild_required_info_text(quick=False)

  @QtCore.pyqtSlot()
  def on_set_output_path_button_clicked(self):
    directory = self.select_path(title='选择成果文件存储路径...', current_path='D:/')
    self.output_path_textedit.setText(directory)

  @QtCore.pyqtSlot()
  def on_generate_button_clicked(self):
    msr = self.documasonry
    msr.set_template_paths(paths=self.get_templates_from_table())
    msr.set_output_path(path=self.output_path_textedit.text())

    info_text = self.info_textedit.toPlainText()
    need_index = self.add_index_checker.isChecked()
    msr.generate(info=Information.from_string(info_text),
                 save=True,
                 add_index=need_index)













if __name__ == '__main__':
  app = QApplication(sys.argv)
  gui = DocumasonryGUI()
  gui.show()
  sys.exit(app.exec_())
