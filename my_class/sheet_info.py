import openpyxl
import row_col_operator

class SheetInfo:
  def __init__(self):
    self.__wb = ''
    self.__ws = ''
    self.__wb_Name = ''
    self.__ws_Name = ''
    self.__start_row = 0
    self.__start_col = 0
    self.__end_row = 0
    self.__end_col = 0
    self.__key_row = 0
    self.__key_col = 0
    self.__target_row = 0
    self.__target_col = 0
    self.__result_row = 0
    self.__result_col = 0
  
  #----------------------------------------------------------------
  #setter
  #----------------------------------------------------------------
  def set_wb(self, wb):
    self.__wb = wb

  def set_ws(self, ws):
    self.__ws = ws

  def set_wb_name(self, wb_name):
    self.__wb_name = wb_name

  def set_ws_name(self, ws_name):
    self.__ws_name = ws_name

  def set_start_row(self, start_row):
    self.__start_row = start_row

  def set_start_col(self, start_col):
    self.__start_col = start_col

  def set_end_row(self, end_row):
    self.__end_row = end_row

  def set_end_col(self, end_col):
    self.__end_col = end_col

  def set_key_row(self, key_row):
    self.__key_row = key_row

  def set_key_col(self, key_col):
    self.__key_col = key_col

  def set_target_row(self, target_row):
    self.__target_row = target_row

  def set_target_col(self, target_col):
    self.__target_col = target_col

  def set_result_row(self, result_row):
    self.__result_row = result_row

  def set_result_col(self, result_col):
    self.__result_col = result_col

  #----------------------------------------------------------------
  #getter
  #----------------------------------------------------------------
  def get_wb(self):
    return self.__wb

  def get_ws(self):
    return self.__ws

  def get_wb_name(self):
    return self.__wb_name

  def get_ws_name(self):
    return self.__ws_name

  def get_start_row(self):
    return self.__start_row

  def get_start_col(self):
    return self.__start_col

  def get_end_row(self):
    return self.__end_row

  def get_end_col(self):
    return self.__end_col

  def get_key_row(self):
    return self.__key_row

  def get_key_col(self):
    return self.__key_col

  def get_target_row(self):
    return self.__target_row

  def set_target_col(self):
    return self.__target_col

  def set_result_row(self):
    return self.__result_row

  def set_result_col(self):
    return self.__result_col
  
  #----------------------------------------------------------------
  #function
  #----------------------------------------------------------------
  def check_my_value(self):
    """check my value and determine to continue

    Returns:
        bool: determine to continue 
    """
    if ('' == self.__wb_name) or ('' == self.__ws_name):
      flag = False
    else:
      flag = True
    return flag

  def check_my_value_row(self):
    """check my value of row

    Returns:
        bool: determine to continue
    """
    if ('' == self.__ws) or (0 == self.__key_row):
      flag = False
    else:
      flag = True
    return flag

  def check_my_value_col(self):
    """check my value of col

    Returns:
        bool: determine to continue
    """
    if ('' == self.__ws) or (0 == self.__key_col):
      flag = False
    else:
      flag = True
    return flag

  def set_sheet_info(self):
    """set sheet info

    Returns:
        nothing
    """
    if self.check_my_value:
      self.set_wb(self.set_wb_name)
      self.set_ws(self.set_ws_name)
    else:
      print('err')

  def set_row_col_info(self):
    """set row and column

    Returns:
        nothing
    """

    RowColOperator = row_col_operator.RowColOperator()
    RowColOperator.set_ws(self.__ws)
    RowColOperator.set_target_row(self.__target_row)
    RowColOperator.set_target_col(self.__target_col)

    if (0 == self.__end_row) and (0 != self.__key_col) \
    and (self.check_my_value_row):
      self.set_end_row(RowColOperator.get_end_row)
    
    if (0 == self.__end_col) and (0 != self.__key_row) \
    and (self.check_my_value_col):
      self.set_end_col(RowColOperator.get_end_col)
