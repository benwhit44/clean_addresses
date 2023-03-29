import openpyxl
import urllib.request
import xml.etree.ElementTree as ET

wb = openpyxl.load_workbook('2022 YIR Mailing List.xlsx', data_only=True)
print(wb.sheetnames)
sheet = wb['2022 combined list']
n_rows = 4

# Dynamically create dictionary from the spreadsheet
address_dict = {}
# Set to iterate over the number of columns
n_col = 0
# Remove max_row once ready for production
for col in sheet.iter_cols(min_row=1, max_row=n_rows, min_col=1):
    # Set row_n to be the id number for each row
    row_n = 0
    temp_dict = {}
    for row in sheet.iter_rows(min_row=1, max_row=n_rows, min_col=1):
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
# print(address_dict['First Name'][3])
list_dict = list(address_dict)
print(list_dict)

for k in address_dict['Addr1']:
    print(k, address_dict['Addr1'][k], address_dict['Addr2'][k], address_dict['City'][k])
#     x = address_dict.get('Addr1', {}).get('jesus', 'nothing')
# print(x)

for k in address_dict['Addr1']:
    requestXML = """
    <?xml version="1.0"?>
    <AddressValidateRequest USERID="292BENWA3717">
        <Revision>1</Revision>
        <Address ID="0">
            <Address1>"""+str(address_dict.get('Addr1', {}).get(k, '') or '')+"""</Address1>
            <Address2>"""+str(address_dict.get('Addr2', {}).get(k, '') or '')+"""</Address2>
            <City>"""+str(address_dict.get('City', {}).get(k, '') or '')+"""</City>
            <State>"""+str(address_dict.get('State', {}).get(k, '') or '')+"""</State>
            <Zip5>"""+str(address_dict.get('Zip', {}).get(k, '') or '')+"""</Zip5>
            <Zip4/>
        </Address>
    </AddressValidateRequest>
    """
    # print("XML\n", requestXML)

    #prepare xml string doc for query string
    docString = requestXML
    docString = docString.replace('\n','').replace('\t','')
    docString = urllib.parse.quote_plus(docString)

    url = "http://production.shippingapis.com/ShippingAPI.dll?API=Verify&XML=" + docString
    print(url + "\n\n")

    response = urllib.request.urlopen(url)
    if response.getcode() != 200:
        print("Error making HTTP call:")
        print(response.info())
        exit()

    contents = response.read()
    print(contents)

# root = ET.fromstring(contents)
# for address in root.findall('Address'):
#     print()
#     print("Address1: " + address.find("Address1").text)
#     print("Address2: " + address.find("Address2").text)
#     print("City:	 " + address.find("City").text)
#     print("State:	" + address.find("State").text)
#     print("Zip5:	 " + address.find("Zip5").text)
#     print("Zip4:	 " + address.find("Zip4").text)

#########################

# requestXML = """
# <?xml version="1.0"?>
# <AddressValidateRequest USERID="292BENWA3717">
#     <Revision>1</Revision>
#     <Address ID="0">
#         <Address1>2335 S State</Address1>
#         <Address2>Suite 300</Address2>
#         <City>Provo</City>
#         <State>UT</State>
#         <Zip5>84604</Zip5>
#         <Zip4/>
#     </Address>
# </AddressValidateRequest>
# """
#
# #prepare xml string doc for query string
# docString = requestXML
# docString = docString.replace('\n','').replace('\t','')
# docString = urllib.parse.quote_plus(docString)
#
# url = "http://production.shippingapis.com/ShippingAPI.dll?API=Verify&XML=" + docString
# print(url + "\n\n")
#
# response = urllib.request.urlopen(url)
# if response.getcode() != 200:
#     print("Error making HTTP call:")
#     print(response.info())
#     exit()
#
# contents = response.read()
# print(contents)
#
# root = ET.fromstring(contents)
# for address in root.findall('Address'):
#     print()
#     print("Address1: " + address.find("Address1").text)
#     print("Address2: " + address.find("Address2").text)
#     print("City:	 " + address.find("City").text)
#     print("State:	" + address.find("State").text)
#     print("Zip5:	 " + address.find("Zip5").text)
#     print("Zip4:	 " + address.find("Zip4").text)





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