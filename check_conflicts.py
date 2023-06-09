
def create_conflict_report(username, password, ECO, pb, value_text, root):
    
    from requests.auth import HTTPBasicAuth
    import xlsxwriter
    import requests
    import copy
    import os
    from datetime import datetime


    def check_lifecycle_status(items_list,auth,parents_flag):
        
        non_production_items = []
        request_link = "https://fa-evbp-saasfaprod1.fa.ocs.oraclecloud.com/fscmRestApi/resources/11.13.18.05/itemsLOV"
        if parents_flag == 0:
            lifecycle_params = {'q':'ItemNumber='+' or '.join(items_list), 'limit':1000}
        else:
            lifecycle_params = {'q':'ItemNumber='+' or '.join(set([x[1] for x in items_list if x[1] != None])),'limit':1000}
        items_LOVs = requests.get(request_link,params=lifecycle_params,auth=auth, verify=False)
        items_LOVs = items_LOVs.json()
        for item in items_LOVs['items']:
            if item['CurrentPhaseCode'] != 'Production' and item['ItemNumber'] in items_list:
                if parents_flag == 0:
                    non_production_items.append((item['ItemNumber'],item['CurrentPhaseCode']))
                else:
                    try:
                        index_of_item = [x[0] for x in items_list].index(item['ItemNumber'])
                    except:
                        pass
                    else:
                        non_production_items.append((item['ItemNumber'],item['CurrentPhaseCode'],items_list[index_of_item][1]))
                    
        
        return non_production_items

    def get_only_latest_not_disabled(items):
        # sorted_BOM = sorted(items, key=lambda i: i["StartDateTime"], reverse=True)
        items.reverse()
        assembly_BOM = []
        assembly_BOM_components_names = []
        for component in items:
            if component['ComponentItemNumber'] not in assembly_BOM_components_names and component['ACDTypeValue'] != 'Disabled':
                assembly_BOM.append(component)
            assembly_BOM_components_names.append(component['ComponentItemNumber'])

                
        
        return assembly_BOM
        


    # username = "daniel.marom@kornit.com"
    # password = "Kornit@2023"
    auth = HTTPBasicAuth(username, password)
    # ECO = 'ECO-10033-04'
    check_assemblies = ['23','24','33','06']
    params = {'q':'ChangeNotice='+ECO}
    # pb['value'] = 10
    affected_items_list = []
    active_ECO_items = [] 
    kids_matched_with_active_ECO = []
    parents_of_affected_items = []
    selves_matched_with_active_ECO = []
    parents_matched_with_active_ECO = []
    same_revision_updates = []
    non_production_selves = []
    params_except_ECO = {'q':f"StatusTypeValue=Interim approval;ChangeNotice!={ECO}", 'limit':1000, 'expand':'all'}
    # params_except_ECO = {'q':f"StatusTypeValue=Interim approval", 'limit':1000, 'expand':'all'} #WRONG ONE!!!!!!!!!!!!!!!!!!!!!!!!!!!
    all_ECOs_link = "https://fa-evbp-saasfaprod1.fa.ocs.oraclecloud.com/fscmRestApi/resources/11.13.18.05/productChangeOrdersV2"
    response = requests.get(all_ECOs_link, auth=auth,params=params_except_ECO, verify=False)
    if response.status_code == 401:
        return 401
    
    # print(response.status_code)
    
    pb['value'] = 0
    value_text['text'] = "0%"
    root.update()
        
    if response.status_code == 401:
        return 401
    response_json =  response.json()


    if os.path.isfile(f'{ECO}_Summary_report.xlsx'):
            os.remove(f'{ECO}_Summary_report.xlsx')
    workbook = xlsxwriter.Workbook(f'{ECO}_Conflict_report.xlsx')
    worksheet = workbook.add_worksheet()
    worksheet.set_column(0, 0, 73.29)
    worksheet.set_column(1, 1, 15)
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
    worksheet.write(2,1, 'Problematic P/N',bold)
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
            active_ECO_items.append((affected_item['ItemNumber'], active_ECO['ChangeNotice'],affected_item['NewItemRevision']))
    pass
    # print(active_ECO_items)
    structure_link = 'https://fa-evbp-saasfaprod1.fa.ocs.oraclecloud.com:443/fscmRestApi/resources/11.13.18.05/itemStructures'
    response = requests.get('https://fa-evbp-saasfaprod1.fa.ocs.oraclecloud.com/fscmRestApi/resources/11.13.18.05/productChangeOrdersV2', auth=auth,params=params, verify=False)
    response_json =  response.json()
    affected_object_link = response_json['items'][0]['links'][2]['href']
    response_affected = requests.get(affected_object_link,auth=auth, verify=False,params={'limit':1000})
    response_affected_json = response_affected.json()
    
    num_of_items = len(response_affected_json['items'])
    progress_value = 90/num_of_items
    
    pass
    # for affected_item in response_affected_json['items']:
    for affected_item in response_affected_json['items']:
        
            # print(f"Item {affected_item['ItemNumber']}, changes from Rev {affected_item['OldRevision']} to Rev {affected_item['NewItemRevision']}")
        selves_matched_with_active_ECO.extend([v for v in active_ECO_items if v[0]==affected_item['ItemNumber']])
        affected_item_lifecycle_link = affected_item['links'][3]['href']
        affected_item_lifecycle_LOV = requests.get(affected_item_lifecycle_link,auth=auth, verify=False).json()
        non_production_selves.extend([(x['ItemNumber'],x['LifecyclePhaseValue']) for x in affected_item_lifecycle_LOV['items'] if x['LifecyclePhaseValue']!='Production'])
        same_revision_updates.extend([v for v in active_ECO_items if v[0]==affected_item['ItemNumber'] and v[2] == affected_item['NewItemRevision']])    
        affected_items_list.append(affected_item['ItemNumber'])
        if affected_item['ItemNumber'][0:2] in check_assemblies:
            affected_items_components = set()
            affected_item_structure_link = affected_item['links'][6]['href']
            
            response_affected_structure = requests.get(affected_item_structure_link,auth=auth, verify=False).json()
            
            
            if response_affected_structure['items'] != []:
                response_affected_structure_comps = requests.get(response_affected_structure['items'][0]['links'][3]['href'],auth=auth,params={'limit':1000}, verify=False).json()
                # OLD_BOM = f"https://fa-evbp-saasfaprod1.fa.ocs.oraclecloud.com:443/fscmRestApi/resources/11.13.18.05/itemStructures/{response_affected_structure_comps['items'][0]['BillSequenceId']}/child/Component"
                added = []
                removed = []
                updated = []

                for component in response_affected_structure_comps['items']:
                        a = [v for v in active_ECO_items if v[0] == component['ComponentItemNumber']]
                        a = [(affected_item['ItemNumber'],x[0],x[1]) for x in a]
                        kids_matched_with_active_ECO.extend(a)
        
            where_used_params = {'q':f"ItemNumber={affected_item['ItemNumber']}", 'limit':1000}
            structure_affected_item = requests.get(structure_link,where_used_params, auth=auth, verify=False)
            structure_affected_item = structure_affected_item.json()
            if len(structure_affected_item['items']) != 0:
                component_link = structure_affected_item['items'][0]['links'][6]['href']
                component_response = requests.get(component_link, auth=auth, verify=False,params={'limit':1000})
                component_response = component_response.json()
                where_used_link = -1
                # sorted_components = sorted(component_response['items'], key=lambda i: i["StartDateTime"], reverse=True)
                components_updated = get_only_latest_not_disabled(component_response['items'])
                for item in components_updated:
                    where_used_link = item['links'][15]['href']
                    affected_items_components.add(item['ComponentItemNumber'])
                    continue
                if affected_items_components != set():
                    if affected_item['ItemNumber']=='33-PNEM-0048':
                        pass
                    non_production_components = check_lifecycle_status(affected_items_components,auth,0)
                    for item in non_production_components:
                        worksheet.write(i,0,f"Item {item[0]} (kid of {affected_item['ItemNumber']}) is in {item[1]} lifecycle stage")
                        worksheet.write(i,1,item[0])
                        i+=1
        
            
                
                if where_used_link == -1:
                    where_used_link = component_response['items'][0]['links'][15]['href']

                where_used_response = requests.get(where_used_link,params={'limit':1000}, auth=auth, verify=False)
                where_used_response = where_used_response.json()
                parents_of_affected_items.extend(list(set([(affected_item['ItemNumber'], x['ParentItemNumber']) for x in where_used_response['items'] if x['ComponentItemNumber']==affected_item['ItemNumber']])))
            
        pb['value'] += progress_value
        value_text['text'] = f"{round(pb['value'], 2)}%"
        root.update()
            
    non_production_parents = check_lifecycle_status(parents_of_affected_items, auth=auth,parents_flag=1)

    for item in non_production_parents:
        worksheet.write(i,0,f"Item {item[2]} (parent of {item[0]}) is in {item[1]} lifecycle stage")
        worksheet.write(i,1,item[2])
        i+=1


    # non_production_selves = check_lifecycle_status(affected_items_list,auth=auth,parents_flag=0)

    for item in non_production_selves:
        worksheet.write(i,0,f"Item {item[0]} is in {item[1]} lifecycle stage")
        worksheet.write(i,1,item[0])
        i+=1
    pass

    for item in parents_of_affected_items:
        parents_matched_with_active_ECO.extend([(item[0], v[0],v[1]) for v in active_ECO_items if v[0] == item[1]])

    # for item in selves_matched_with_active_ECO:
    #      worksheet.write(i,0,f"Item {item[0]} exists in {item[1]}")
    #      same_revision_updates
    #      i+=1

    for item in same_revision_updates:
        worksheet.write(i,0,f"Item {item[0]} already exists in {item[1]} at the same revision ({item[2]})")
        worksheet.write(i,1,item[0])
        i+=1

    for item in parents_matched_with_active_ECO:
        worksheet.write(i,0,f"Item {item[0]} has a parent {item[1]} that exists in {item[2]}")
        worksheet.write(i,1,item[0])
        i+=1

    for item in kids_matched_with_active_ECO:
        worksheet.write(i,0,f"Item {item[0]} has a kid {item[1]} that exists in {item[2]}")
        worksheet.write(i,1,item[0])
        i+=1
    
    pb['value'] = 100
    value_text['text'] = "100%"
    root.update()
        
    workbook.close()
    return f"{ECO}_Conflict_report.xlsx"


if __name__ == "__main__":
   # stuff only to run when not called via 'import' here
   create_conflict_report(None, None, None, None, None, None)


