# import openpyxl
# from openpyxl.utils import get_column_letter
#
# import datetime
#
# wb = openpyxl.load_workbook('attendance_reg.xlsx')
# sheet = wb.get_sheet_by_name('February')
# col = sheet.max_column
# col += 1
# column = get_column_letter(col)
# today = datetime.datetime.today().strftime('%Y-%m-%d')
# dtr = "Aprajita Verma"
#
# for row in sheet.iter_rows(min_row=1, min_col=1):
#     for cell in row:
#         if cell.value:
#             if cell.value == dtr:
#                 row_no_of_this_cell = cell.row
#                 cell_no = str(column) + str(row_no_of_this_cell)
#                 cell_to_write = sheet[cell_no]
#                 cell_to_write.value = "Present"
#
#     wb.save('attendance_reg.xlsx')
#
# wb = openpyxl.load_workbook('attendance_reg.xlsx')
# print("Workbook opened")
# sheet = wb.get_sheet_by_name('February')
# print("Got the sheet")
#
#
#
#
#
#
#
#
# #
# #
# # # xl_wb = xlwt.Workbook()
# # #
# # # sheet = xl_wb.add_sheet("Sheet1")
# # #
# # # sheet.write(1, 0, "Nishant Sambyal")
# # # sheet.write(2, 0, "Rajat Mehra")
# # # sheet.write(3, 0, "Aprajita Verma")
# # # sheet.write(4, 0, "Vivek")
# # # sheet.write(5, 0, "Tishu Singh")
# # #
# # # xl_wb.save("attendance_register.xlsx")
# #
# # name = "Nishant Sambyal"
# #
# #
# # def write_to_register(n):
# #     # if n == "Unknown":
# #     #     pass
# #     # else:
# #     xl_workbook = openpyxl.load_workbook("attendance_reg.xlsx")
# #     sheet = xl_workbook.active
# #
# #     for row in sheet.iter_rows(min_row=1, min_col=1):
# #         for cell in row:
# #             if cell.value:
# #                 if cell.value == n:
# #                     print(row)
# #                     # sheet[] = "Present"
# #
# #
# # if __name__ == "__main__":
# #     write_to_register(name)
#
# # import pandas as pd
# #
# #
# # def mark_present(cell):
# #     if cell == '':
# #         return 'Present'
# #     else:
# #         return cell
# #
# #
# # df = pd.read_excel("attendance_reg.xlsx", "Sheet1", converters={
# #     'NAME': mark_present
# # })
# #
# # print(df)
#
#
#
#
