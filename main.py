import openpyxl

wb = openpyxl.load_workbook('2022 YIR Mailing List.xlsx', data_only=True)
print(wb.sheetnames)
sheet = wb['2022 combined list']

# Dynamically create dictionary from the spreadsheet
address_dict = {}
# Set to iterate over the number of columns
n_col = 0
# Remove max_row once ready for production
for col in sheet.iter_cols(min_row=1, max_row=11, min_col=1):
    # Set row_n to be the id number for each row
    row_n = 0
    temp_dict = {}
    for row in sheet.iter_rows(min_row=1, max_row=11, min_col=1):
        if row_n == 0:
            row_head = row[n_col].value
        else:
            temp_dict[row_n] = row[n_col].value
        row_n += 1
    n_col += 1
    temp_dict = {row_head: temp_dict}
    address_dict.update(temp_dict)
    # print(temp_dict)
print(address_dict)


# # Convert data into a dictionary with row number as the identifying key
# address_dict = {}
# # Select the max number of columns to read
# maxcoln = 12
# for n in range(0, maxcoln):
#     row_n = 0
#     temp_dict = {}
#     for row in sheet.iter_rows(min_row=1, max_row=11, min_col=1, max_col=maxcoln):
#         if row_n == 0:
#             row_head = row[n].value
#         else:
#             temp_dict[row_n] = row[n].value
#         row_n += 1
#     temp_dict = {row_head: temp_dict}
#     address_dict.update(temp_dict)
#     # print(temp_dict)
# print(address_dict)





####################################
# import pandas as pd
# from geopy.geocoders import Nominatim
# # messy_address = pd.read_excel("Book1.xlsx")
# # messy_address = pd.read_excel("2022 YIR Mailing List.xlsx", sheet_name="2022 combined list")
# messy_address = pd.read_csv("2022 YIR Mailing List.csv")
# geolocator = Nominatim(user_agent="benwhit44@gmail.com")
#
#
# def extract_clean_address(address):
#     try:
#         location = geolocator.geocode(address)
#         return location.address
#     except:
#         return ''
# messy_address['clean address'] = messy_address.apply(lambda x: extract_clean_address(x['Raw Address']) , axis =1  )