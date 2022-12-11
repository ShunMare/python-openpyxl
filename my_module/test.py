import os
import sys
# sys.path.append(os.path.dirname(os.path.abspath(__file__)) + '\\..\\..\\')
sys.path.append(os.path.dirname(os.path.abspath(__file__)) + '\\..\\')
import openpyxl
from my_class.row_col_operator.row_col_operator import RowColOperator

def main():
  row_col_operator = RowColOperator()
  row_col_operator.set_ws = openpyxl.load_workbook('tool_openpyxl.xlsx')['sample1']
  row_col_operator.set_target_col = 3
  print(row_col_operator.get_end_row())
  row_col_operator.set_target_row = 16
  print(row_col_operator.get_end_col())

if __name__ == "__main__":
  main()