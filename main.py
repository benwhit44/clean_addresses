import urllib.request
import xml.etree.ElementTree as ET
import csv
import pandas as pd
import re

# Read xlsx into dataframe
df = pd.read_excel(open('2022 YIR Mailing List_NEW.xlsx', 'rb')
                       , sheet_name='2022 combined list', nrows=200)

# Convert to dictionary
address_dict = df.to_dict()

# TODO: Detect invalid address values, i.e "Mail Drop, and suggest new values"
suggested_addr1 = {}
suggested_addr2 = {}
for k in address_dict['Addr1']:
    # print(k, "Addr1:", address_dict['Addr1'][k], "Addr2:", address_dict['Addr2'][k], "City:", address_dict['City'][k])
    text = address_dict['Addr1'][k]
    reg_match = ",? mail drop.*"
    if re.search(reg_match, text, re.IGNORECASE):
        suggested_addr1[k] = re.sub(reg_match, "", text, flags=re.IGNORECASE)
        suggested_addr2[k] = re.search("mail drop.*", text, flags=re.IGNORECASE).group()
        # suggested_addr2 = temp_split[1]
        # print("Sugg_Addr1:", suggested_addr1[k], "Sugg_Addr2:", suggested_addr2[k])
    else:
        suggested_addr1[k] = address_dict['Addr1'][k]
        suggested_addr2[k] = ''
suggested_addr1 = {"Suggested_Addr1": suggested_addr1}
suggested_addr2 = {"Suggested_Addr2": suggested_addr2}
address_dict.update(suggested_addr1)
address_dict.update(suggested_addr2)

# address_json = json.dumps(address_dict, indent=4)
# print(address_json)
#
# Create the requestXML for USPS API
temp_addr1_dict, temp_addr2_dict, temp_city_dict, temp_state_dict\
    , temp_zip_dict, temp_extzip_dict, temp_err_dict, temp_full_addr = {}, {}, {}, {}, {}, {}, {}, {}

for k in address_dict['Addr1']:
    requestXML = """
    <?xml version="1.0"?>
    <AddressValidateRequest USERID="292BENWA3717">
        <Revision>1</Revision>
        <Address ID="0">
            <Address1>"""+str(address_dict.get('Suggested_Addr1', {}).get(k, '') or '')+"""</Address1>
            <Address2>"""+str(address_dict.get('Suggested_Addr2', {}).get(k, '') or '')+"""</Address2>
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
    # print(url + "\n\n")

    response = urllib.request.urlopen(url)
    if response.getcode() != 200:
        print("Error making HTTP call:")
        print(response.info())
        exit()

    contents = response.read()
    # print(contents)

    root = ET.fromstring(contents)
    # print(root.text)

    # Find the key address values and create new dictionaries
    for address in root.findall('Address'):
        if not address.find("Error"):
            # print("Address1: " + address.find("Address1").text)
            # print("Address2: " + address.find("Address2").text)
            # print("City:	 " + address.find("City").text)
            # print("State:	" + address.find("State").text)
            # print("Zip5:	 " + address.find("Zip5").text)
            # print("Zip4:	 " + address.find("Zip4").text)
            if address.find("Address1") == None:
                # print("Missing Addr1:", url)
                temp_addr1_dict[k] = ''
            else:
                temp_addr1_dict[k] = address.find("Address1").text
            # if address.find("Address2") == None:
            #     print("Missing Addr2:", url)
            #     temp_addr2_dict[k] = ''
            # else:
            # temp_addr1_dict[k] = address.find("Address1").text
            temp_addr2_dict[k] = address.find("Address2").text
            temp_city_dict[k] = address.find("City").text
            temp_state_dict[k] = address.find("State").text
            temp_zip_dict[k] = address.find("Zip5").text
            temp_extzip_dict[k] = address.find("Zip4").text
            temp_err_dict[k] = ""
            # temp_full_addr[k] = address.find("Address1").text + ' ' + address.find("Address2").text \
            #                     + ', ' + address.find("City").text + ', ' + address.find("State").text \
            #                     + '-' + address.find("Zip5").text
            temp_full_addr[k] = str(temp_addr1_dict[k]) + ' ' +str(temp_addr2_dict[k]) + ' ' + \
                                str(temp_city_dict[k]) + ', '+ str(temp_state_dict[k]) + ' ' + \
                                str(temp_zip_dict[k]) + '-' + str(temp_extzip_dict[k])

        else:
            temp_addr1_dict[k] = ""
            temp_addr2_dict[k] = ""
            temp_city_dict[k] = ""
            temp_state_dict[k] = ""
            temp_zip_dict[k] = ""
            temp_extzip_dict[k] = ""
            for err in root.findall('Address/Error'):
                temp_err_dict[k] = err.find("Description").text
                # print("Error: ", err.find("Description").text)
        #     print("get error: ", root.text)


new_addr1_dict = {"New_Addr1": temp_addr1_dict}
new_addr2_dict = {"New_Addr2": temp_addr2_dict}
new_city_dict = {"New_City": temp_city_dict}
new_state_dict = {"New_State": temp_state_dict}
new_zip_dict = {"New_Zip": temp_zip_dict}
new_extzip_dict = {"New_ExtZip": temp_extzip_dict}
new_err_dict = {"Error_Description": temp_err_dict}
new_full_addr = {"Full_Address": temp_full_addr}

address_dict = address_dict | new_addr1_dict | new_addr2_dict | new_city_dict\
               | new_state_dict | new_zip_dict | new_extzip_dict | new_full_addr | new_err_dict
# list_dict = list(address_dict)
# print(list_dict)

# Convert back to dataframe
df = pd.DataFrame.from_dict(address_dict)

# Sort data and identify duplicates
df = df.sort_values(by=['Full Name', 'Full_Address', 'Email Address'])
# print(df[['Full Name', 'Addr1', 'Email Address']])

bool_series = df.duplicated(subset=['Full Name', 'Full_Address'])
# print(bool_series)

# Remove Duplicates
df = df[~bool_series]
# print(df[['Full Name', 'Addr1', 'Email Address']])

# Sort back to the original
df = df.sort_index()
# print(df[['Full Name', 'Addr1', 'Email Address']])
df.to_csv('clean_addresses.csv', encoding='utf-8', index=False)