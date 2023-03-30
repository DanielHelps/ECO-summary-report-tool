
from requests.auth import HTTPBasicAuth
import xlsxwriter
import requests
import copy
import os
from datetime import datetime

username = "daniel.marom@kornit.com"
password = "Kornit@2023"
auth = HTTPBasicAuth(username, password)
ECO = 'ECO-10033-04'
check_assemblies = ['23','24','33','06']
params = {'q':'ChangeNotice='+ECO}
# pb['value'] = 10
active_ECO_items = []
kids_matched_with_active_ECO = []
parents_of_affected_items = []
selves_matched_with_active_ECO = []
parents_matched_with_active_ECO = []
# params_except_ECO = {'q':f"StatusTypeValue=Interim approval;ChangeNotice!={ECO}", 'limit':1000, 'expand':'all'}
params_except_ECO = {'q':f"StatusTypeValue=Interim approval", 'limit':1000, 'expand':'all'} #WRONG ONE!!!!!!!!!!!!!!!!!!!!!!!!!!!
all_ECOs_link = "https://fa-evbp-saasfaprod1.fa.ocs.oraclecloud.com/fscmRestApi/resources/11.13.18.05/productChangeOrdersV2"
response = requests.get(all_ECOs_link, auth=auth,params=params_except_ECO)
# print(response.status_code)
response_json =  response.json()


if os.path.isfile(f'{ECO}_Summary_report.xlsx'):
        os.remove(f'{ECO}_Summary_report.xlsx')
workbook = xlsxwriter.Workbook(f'{ECO}_Conflict_report.xlsx')
worksheet = workbook.add_worksheet()
worksheet.set_column(0, 0, 64.14)
# worksheet.set_column(1, 1, 7.14)
# worksheet.set_column(2, 2, 8.14)
# worksheet.set_column(3, 3, 19.14)
# worksheet.set_column(5, 5, 7.71)
# worksheet.set_column(6, 6, 10.14)
# worksheet.set_column(7, 7, 11.29)


bold = workbook.add_format({'bold': True})
blue = workbook.add_format({'font_color':'blue'})
purple = workbook.add_format({'font_color':'purple'})
red = workbook.add_format({'font_color':'red'})
green = workbook.add_format({'font_color':'green'})

worksheet.write(0,0, f'{ECO} conflict report',bold)
# worksheet.write(1,0, f"ECO name: {response_json['items'][0]['ChangeName']}",bold)
# worksheet.write(2,0, f"ECO description: {response_json['items'][0]['Description']}",bold)

worksheet.write(2,0, 'Conflict',bold)
# worksheet.write(2,1, 'Old Rev',bold)
# worksheet.write(2,2, 'New Rev',bold)
# worksheet.write(4,3, 'Changed Component',bold)
# worksheet.write(4,4, 'Old qty',bold)
# worksheet.write(4,5, 'New qty',bold)
# worksheet.write(4,6, 'Qty change',bold)
# worksheet.write(4,7, 'Change type',bold)
i = 3



for active_ECO in response_json['items']:
    for affected_item in active_ECO['changeOrderAffectedObject']:
        active_ECO_items.append((affected_item['ItemNumber'], active_ECO['ChangeNotice']))
pass
# print(active_ECO_items)
structure_link = 'https://fa-evbp-saasfaprod1.fa.ocs.oraclecloud.com:443/fscmRestApi/resources/11.13.18.05/itemStructures'
response = requests.get('https://fa-evbp-saasfaprod1.fa.ocs.oraclecloud.com/fscmRestApi/resources/11.13.18.05/productChangeOrdersV2', auth=auth,params=params)
response_json =  response.json()
affected_object_link = response_json['items'][0]['links'][2]['href']
response_affected = requests.get(affected_object_link,auth=auth)
response_affected_json = response_affected.json()
pass
# for affected_item in response_affected_json['items']:
for affected_item in response_affected_json['items']:
        
        # print(f"Item {affected_item['ItemNumber']}, changes from Rev {affected_item['OldRevision']} to Rev {affected_item['NewItemRevision']}")
    selves_matched_with_active_ECO.extend([v for v in active_ECO_items if v[0]==affected_item['ItemNumber']])

    if affected_item['ItemNumber'][0:2] in check_assemblies:
        affected_item_structure_link = affected_item['links'][6]['href']
        response_affected_structure = requests.get(affected_item_structure_link,auth=auth).json()
        if response_affected_structure['items'] != []:
            response_affected_structure_comps = requests.get(response_affected_structure['items'][0]['links'][3]['href'],auth=auth,params={'limit':1000}).json()
            # OLD_BOM = f"https://fa-evbp-saasfaprod1.fa.ocs.oraclecloud.com:443/fscmRestApi/resources/11.13.18.05/itemStructures/{response_affected_structure_comps['items'][0]['BillSequenceId']}/child/Component"
            added = []
            removed = []
            updated = []

            for component in response_affected_structure_comps['items']:
                    a = [v for v in active_ECO_items if v[0] == component['ComponentItemNumber']]
                    a = [(affected_item['ItemNumber'],x[0],x[1]) for x in a]
                    kids_matched_with_active_ECO.extend(a)
    
        where_used_params = {'q':f"ItemNumber={affected_item['ItemNumber']}", 'limit':1000}
        structure_affected_item = requests.get(structure_link,where_used_params, auth=auth)
        structure_affected_item = structure_affected_item.json()
        component_link = structure_affected_item['items'][0]['links'][6]['href']
        component_response = requests.get(component_link, auth=auth)
        component_response = component_response.json()
        where_used_link = -1
        for item in component_response['items']:
            if item['ComponentItemNumber'][0:2] != '45':
                    where_used_link = item['links'][15]['href']
                    break
            
        if where_used_link == -1:
            where_used_link = component_response['items'][0]['links'][15]['href']

        where_used_response = requests.get(where_used_link,params={'limit':1000}, auth=auth)
        where_used_response = where_used_response.json()
        parents_of_affected_items.extend(list(set([(affected_item['ItemNumber'], x['ParentItemNumber']) for x in where_used_response['items'] if x['ComponentItemNumber']==affected_item['ItemNumber']])))
pass

for item in parents_of_affected_items:
     parents_matched_with_active_ECO.extend([(item[0], v[0],v[1]) for v in active_ECO_items if v[0] == item[1]])

for item in selves_matched_with_active_ECO:
     worksheet.write(i,0,f"Item {item[0]} exists in {item[1]}")
     i+=1

for item in parents_matched_with_active_ECO:
     worksheet.write(i,0,f"Item {item[0]} has a parent {item[1]} that exists in {item[2]}")
     i+=1

for item in kids_matched_with_active_ECO:
    worksheet.write(i,0,f"Item {item[0]} has a kid {item[1]} that exists in {item[2]}")
    i+=1

workbook.close()
pass