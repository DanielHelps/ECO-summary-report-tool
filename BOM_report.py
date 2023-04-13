
def main(username, password, ECO, pb, value_text, root):
    
    from requests.auth import HTTPBasicAuth
    import xlsxwriter
    import requests
    import copy
    import os
    from datetime import datetime

    auth = HTTPBasicAuth(username, password)
    # ECO = 'ECO-10028-23'
    check_assemblies = ['23','33','06']
    params = {'q':'ChangeNotice='+ECO}
    # pb['value'] = 10
    # root.update()
    response = requests.get('https://fa-evbp-saasfaprod1.fa.ocs.oraclecloud.com/fscmRestApi/resources/11.13.18.05/productChangeOrdersV2', auth=auth,params=params,verify=False)
    if response.status_code == 401:
        return 401
    pb['value']=0
    response_json =  response.json()
    affected_object_link = response_json['items'][0]['links'][2]['href']
    response_affected = requests.get(affected_object_link,auth=auth,verify=False)
    response_affected_json = response_affected.json()
    if os.path.isfile(f'{ECO}_Summary_report.xlsx'):
        os.remove(f'{ECO}_Summary_report.xlsx')
    workbook = xlsxwriter.Workbook(f'{ECO}_Summary_report.xlsx')
    worksheet = workbook.add_worksheet()
    worksheet.set_column(0, 0, 27.43)
    worksheet.set_column(1, 1, 7.14)
    worksheet.set_column(2, 2, 8.14)
    worksheet.set_column(3, 3, 19.14)
    worksheet.set_column(5, 5, 7.71)
    worksheet.set_column(6, 6, 10.14)
    worksheet.set_column(7, 7, 11.29)
    

    bold = workbook.add_format({'bold': True})
    blue = workbook.add_format({'font_color':'blue'})
    purple = workbook.add_format({'font_color':'purple'})
    red = workbook.add_format({'font_color':'red'})
    green = workbook.add_format({'font_color':'green'})

    worksheet.write(0,0, f'{ECO} summary report',bold)
    worksheet.write(1,0, f"ECO name: {response_json['items'][0]['ChangeName']}",bold)
    worksheet.write(2,0, f"ECO description: {response_json['items'][0]['Description']}",bold)
    
    worksheet.write(4,0, 'Item P/N',bold)
    worksheet.write(4,1, 'Old Rev',bold)
    worksheet.write(4,2, 'New Rev',bold)
    worksheet.write(4,3, 'Changed Component',bold)
    worksheet.write(4,4, 'Old qty',bold)
    worksheet.write(4,5, 'New qty',bold)
    worksheet.write(4,6, 'Qty change',bold)
    worksheet.write(4,7, 'Change type',bold)


    def remove_a_from_b(a, b):
        a_copy = copy.deepcopy(a)
        b_copy = copy.deepcopy(b)
        for i, b_item in enumerate(b_copy):
                match_in_a = [item for item in a_copy if item[0] == b_item[0]]
                for item in match_in_a:
                    b_copy[i] = (b_copy[i][0], b_copy[i][1]-item[1]) 
        return list(filter(lambda x: x[1] > 0,b_copy))

    def add_comp_to_report(component, qty_change, old_qty, new_qty, change, assembly, i):
        worksheet.write(i,0,assembly['ItemNumber'], purple)
        worksheet.write(i,1,assembly['OldRevision'], purple)
        worksheet.write(i,2,assembly['NewItemRevision'], purple)
        color = green if change == "Added" else red
        worksheet.write(i,3,component, color)
        worksheet.write(i,4,old_qty, color)
        worksheet.write(i,5,new_qty, color)
        worksheet.write(i,6,qty_change, color)
        worksheet.write(i,7,change, color)
        
            
        
        

    def find_BOM_change(OLD_BOM, checked_list, added, removed, ECO_num, change_flag):
        OLD_BOM_ITEMS = requests.get(OLD_BOM,auth=auth, params={'limit':1000}, verify=False).json()
        for item in checked_list:
            matched_old_bom = [v for i, v in enumerate(OLD_BOM_ITEMS['items']) if v['ComponentItemNumber'] == item[0] and v['ChangeNotice']==ECO_num]
            if change_flag == 'Added':
                old_qty = 0
                new_qty = matched_old_bom[0]['Quantity']
            elif change_flag == 'Removed':
                old_qty = matched_old_bom[0]['Quantity']
                new_qty = 0
            else:
                change_in_bom = 0
                if len(matched_old_bom)>1:
                    sorted_date_bom = sorted(matched_old_bom, key=lambda x: datetime.strptime(x['LastUpdateDateTime'].split(".")[0], '%Y-%m-%dT%H:%M:%S'))
                    old_qty = sorted_date_bom[0]['Quantity']
                    new_qty = sorted_date_bom[1]['Quantity']
                elif len(matched_old_bom)==1:
                    matched_old_bom = [v for i, v in enumerate(OLD_BOM_ITEMS['items']) if v['ComponentItemNumber'] == item[0]]
                    if len(matched_old_bom)>1:
                        matched_old_bom = sorted(matched_old_bom, key=lambda x: datetime.strptime(x['LastUpdateDateTime'].split(".")[0], '%Y-%m-%dT%H:%M:%S'))
                        if matched_old_bom[1]['EndDateTime'] == None:
                            old_qty = matched_old_bom[0]['Quantity']
                            new_qty = matched_old_bom[1]['Quantity']
                        else:
                            old_qty = matched_old_bom[1]['Quantity']
                            new_qty = matched_old_bom[0]['Quantity']
                        # if matched_old_bom[0]['LastUpdateDateTime'].split(".")[0]==matched_old_bom[1]['LastUpdateDateTime'].split(".")[0]:
                        #     if removed == []:
                        #         old_qty = matched_old_bom[0]['Quantity']
                        #         new_qty = 0
                        #     else:
                        #         old_qty = 0
                        #         new_qty = matched_old_bom[0]['Quantity']


                            
                    elif len(matched_old_bom)==1:
                        old_qty = 0
                        new_qty = matched_old_bom[0]['Quantity']
            change_in_bom = new_qty - old_qty
            a = item[0],abs(change_in_bom),old_qty,new_qty
            if change_in_bom>0:
                added.append(a)
            elif change_in_bom<0:
                removed.append(a)
            
        # for item in OLD_BOM_ITEMS['items']:
        #     updated_component = [v for i, v in enumerate(updated) if v[0] == item['ComponentItemNumber']]
        #     if updated_component != []:
        #         if item['Quantity'] < updated_component[0][1]:
        #             removed.append(updated_component[0])
        #             updated.remove(updated_component[0])
        #         elif item['Quantity'] > updated_component[0][1]:
        #             added.append(updated_component[0])
        #             updated.remove(updated_component[0])


            # elif updated_component != []:
            #     if item['Quantity'] > updated_component[0]:
            #         removed.append(tuple())
            
        return added, removed
            
    i = 5
    num_of_items = len(response_affected_json['items'])
    progress_value = 100/num_of_items
    for affected_item in response_affected_json['items']:
        
        print(f"Item {affected_item['ItemNumber']}, changes from Rev {affected_item['OldRevision']} to Rev {affected_item['NewItemRevision']}")
        worksheet.write(i,0,affected_item['ItemNumber'], blue)
        worksheet.write(i,1,affected_item['OldRevision'], blue)
        worksheet.write(i,2,affected_item['NewItemRevision'], blue)
        i+=1

        if affected_item['ItemNumber'][0:2] in check_assemblies:
            affected_item_structure_link = affected_item['links'][6]['href']
            response_affected_structure = requests.get(affected_item_structure_link,auth=auth,verify=False).json()
            if response_affected_structure['items'] != []:
                response_affected_structure_comps = requests.get(response_affected_structure['items'][0]['links'][3]['href'],auth=auth,params={'limit':1000},verify=False).json()
                OLD_BOM = f"https://fa-evbp-saasfaprod1.fa.ocs.oraclecloud.com:443/fscmRestApi/resources/11.13.18.05/itemStructures/{response_affected_structure_comps['items'][0]['BillSequenceId']}/child/Component"
                added = []
                removed = []
                updated = []

                for component in response_affected_structure_comps['items']:
                    
                    if component['ChangeNotice'] == ECO:
                        print(component['ComponentItemNumber'])
                        if component['ACDTypeCode'] == 1:
                            added.append(tuple((component['ComponentItemNumber'], component['ComponentQuantity'])))
                        elif component['ACDTypeCode'] == 2:
                            updated.append(tuple((component['ComponentItemNumber'], component['ComponentQuantity'])))
                        else:
                            removed.append(tuple((component['ComponentItemNumber'], component['ComponentQuantity'])))
                pass

                
                
                removed_in_end = remove_a_from_b(added,remove_a_from_b(updated, removed))
                updated_in_end = remove_a_from_b(added,remove_a_from_b(removed,updated))
                added_in_end = remove_a_from_b(removed, remove_a_from_b(updated,added))

                
                if added_in_end != []:
                    added_in_end, removed_in_end = find_BOM_change(OLD_BOM, added_in_end, [], removed_in_end, ECO, 'Added')

                if removed_in_end != []:
                    added_in_end, removed_in_end = find_BOM_change(OLD_BOM, removed_in_end, added_in_end, [], ECO, 'Removed')

                if updated_in_end != []:
                    added_in_end, removed_in_end = find_BOM_change(OLD_BOM, updated_in_end, added_in_end, removed_in_end, ECO, 'Updated')



                for item in removed_in_end:
                    if len(item) == 4:
                        add_comp_to_report(item[0],item[1],item[2],item[3],'Removed',affected_item,i)
                    else:
                        add_comp_to_report(item[0],item[1],"","",'Removed',affected_item,i)
                    i+=1
                
                # for item in updated_in_end:
                #     add_comp_to_report(item[0],item[1],'Qty changed',affected_item,i)
                #     i+=1
                
                for item in added_in_end:
                    if len(item) == 4:
                        add_comp_to_report(item[0],item[1],item[2],item[3],'Added',affected_item,i)
                    else:
                        add_comp_to_report(item[0],item[1],"","",'Added',affected_item,i)
                    i+=1

                print(f'Removed: {removed_in_end}')
                print(f'Added: {added_in_end}')
                print(f'Updated: {updated_in_end}')
                # for removed_item in removed:
                #     if removed_item[1] == 0:
                #         removed.remove(removed_item)
        pb['value'] += progress_value
        value_text['text'] = f"{round(pb['value'], 2)}%"
        root.update()
    
        
    workbook.close()
    return f'{ECO}_Summary_report.xlsx'

if __name__ == "__main__":
   # stuff only to run when not called via 'import' here
   main(None, None, None, None, None, None)


