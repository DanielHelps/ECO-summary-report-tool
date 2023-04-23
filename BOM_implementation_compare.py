def create_BOM_Implementation_report(username, password, ECO, pb, value_text, root, check_where):
        
    import requests
    from requests.auth import HTTPBasicAuth
    import pandas as pd
    import os
    import xlsxwriter
    import openpyxl
    import re

    def get_only_latest(items):
        assembly_BOM_duplicates = sorted(items, key=lambda i: i["StartDateTime"], reverse=True)
        assembly_BOM = []
        for component in assembly_BOM_duplicates:
            if component['ComponentItemNumber'] not in [x[0] for x in assembly_BOM] and component['EndDateTime'] == None:
                assembly_BOM.append(tuple((component['ComponentItemNumber'],component['Quantity'])))
        
        return assembly_BOM
        
    def get_latest_BOM(ItemNumber):
        params = {'q':f'ItemNumber={ItemNumber}'}
        response = requests.get('https://fa-evbp-saasfaprod1.fa.ocs.oraclecloud.com:443/fscmRestApi/resources/11.13.18.05/itemStructures', auth=auth,params=params,verify=False)
        if response.status_code == 401:
            return 401
        response = response.json()

        try:
            assembly_BOM_link = response['items'][0]['links'][6]['href']
        except IndexError:
            return []

        assembly_BOM_response = requests.get(assembly_BOM_link, auth=auth,params={'limit':1000},verify=False)
        assembly_BOM_response = assembly_BOM_response.json()
        assembly_BOM = get_only_latest(assembly_BOM_response['items'])
        # assembly_BOM_duplicates = sorted(assembly_BOM_response['items'], key=lambda i: i["StartDateTime"], reverse=True)
        # assembly_BOM = []
        # for component in assembly_BOM_duplicates:
        #     if component['ComponentItemNumber'] not in [x[0] for x in assembly_BOM] and component['EndDateTime'] == None:
        #         assembly_BOM.append(tuple((component['ComponentItemNumber'],component['Quantity'])))
        
        return assembly_BOM
        

    # urllib3.disable_warnings()
    # username = 'daniel.marom@kornit.com'
    # password = 'Kornit@2023'
    auth = HTTPBasicAuth(username, password)
        

    dir = f"{check_where}\\{ECO}"
    ECO_items = os.listdir(dir)
    check_assemblies = ['06','09','23','24','33','60']

    workbook = xlsxwriter.Workbook(f'{ECO}_BOM_Implementation_Report.xlsx')
    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({'bold': True})
    blue = workbook.add_format({'font_color':'blue'})
    purple = workbook.add_format({'font_color':'purple'})
    red = workbook.add_format({'font_color':'red'})
    green = workbook.add_format({'font_color':'green'})
    gray = workbook.add_format({'font_color':'gray'})
    error_font = workbook.add_format({'font_color':'red', 'bold':True})
    worksheet.write(0,0, f'{ECO} BOM Implementation Report',bold)
        
    worksheet.write(2,0, 'Assembly number',bold)
    worksheet.write(2,1, 'Changed item',bold)
    worksheet.write(2,2, 'Old qty',bold)
    worksheet.write(2,3, 'New qty',bold)
    worksheet.write(2,4, 'Qty change',bold)
    worksheet.write(2,5, 'Change type',bold)
    worksheet.write(2,6, 'Units',bold)

    worksheet.set_column(0, 0, 16.71)
    worksheet.set_column(1, 1, 13.86)
    worksheet.set_column(2, 2, 6.71)
    worksheet.set_column(3, 3, 7.71)
    worksheet.set_column(3, 3, 14)
    worksheet.set_column(4, 4, 10.14)
    worksheet.set_column(5, 5, 11.29)
    
    i=3
    # worksheet.write(4,3, 'Changed Component',bold)
    # worksheet.write(4,4, 'Old qty',bold)
    # worksheet.write(4,5, 'New qty',bold)
    # worksheet.write(4,6, 'Qty change',bold)
    # worksheet.write(4,7, 'Change type',bold)
    pb['value'] = 0
    value_text['text'] = f"0%"
    left_to_write=[]
    was_written = []

    progress_value = 100/len(ECO_items)
    for item in ECO_items:
        try:
            
            excel_PN = re.findall("\d\d-[A-Z][A-Z][A-Z][A-Z]-\d\d\d\d\d",item)
            if not excel_PN:
                excel_PN = re.findall("\d\d-[A-Z][A-Z][A-Z][A-Z]-\d\d\d\d",item)
            if excel_PN and 'xls' in item[-4:]:
                if excel_PN[0][0:2] in check_assemblies:
                    df = pd.read_excel(dir + "\\" + item)
                    assembly_BOM = get_latest_BOM(excel_PN[0])
                    if assembly_BOM == 401:
                        return 401
                    PN_column = -1
                    PN_column = [i for i in range(len(df.columns)) if re.findall("\d\d-[A-Z][A-Z][A-Z][A-Z]",str(df.iloc[1][i]))][0]
                    try:
                        unit_column = [i for i, column in enumerate(df.columns) if 'unit' in column.lower()][0]
                    except IndexError:
                        unit_column = -1
                    if type(df.iloc[0,-1])==str or pd.isnull(df.iloc[1,-1]):
                        NEW_BOM = list(df.apply(lambda x: tuple((x.iloc[PN_column].strip(),x.iloc[-2])), axis=1))
                    else:
                        NEW_BOM = list(df.apply(lambda x: tuple((x.iloc[PN_column].strip(),x.iloc[-1])), axis=1))
                    # BOM_delta = []
                    for comp_row, component in enumerate(NEW_BOM):
                        if component[0] in [x[0] for x in assembly_BOM]:
                            item_index = [x[0] for x in assembly_BOM].index(component[0])
                            component_delta = component[1]-assembly_BOM[item_index][1]
                            if component_delta > 0:
                                was_written.append(excel_PN[0])
                                if excel_PN[0] in left_to_write:
                                    left_to_write.remove(excel_PN[0])
                                worksheet.write(i,0,excel_PN[0],purple)
                                worksheet.write(i,1,component[0],green)
                                worksheet.write(i,2,assembly_BOM[item_index][1],green)
                                worksheet.write(i,3,component[1],green)
                                worksheet.write(i,4,component_delta,green)
                                worksheet.write(i,5,"Updated",green)
                                if unit_column !=-1:
                                    worksheet.write(i,6,df.iloc[comp_row,unit_column],green)
                                i+=1
                                # BOM_delta.append(tuple((component[0],component_delta)))
                            elif component_delta < 0:
                                was_written.append(excel_PN[0])
                                if excel_PN[0] in left_to_write:
                                    left_to_write.remove(excel_PN[0])
                                worksheet.write(i,0,excel_PN[0],purple)
                                worksheet.write(i,1,component[0],red)
                                worksheet.write(i,2,assembly_BOM[item_index][1],red)
                                worksheet.write(i,3,component[1],red)
                                worksheet.write(i,4,component_delta,red)
                                worksheet.write(i,5,"Updated",red)
                                if unit_column !=-1:
                                    worksheet.write(i,6,df.iloc[comp_row,unit_column],red)
                                i+=1
                        else:
                            was_written.append(excel_PN[0])
                            if excel_PN[0] in left_to_write:
                                left_to_write.remove(excel_PN[0])
                            worksheet.write(i,0,excel_PN[0],purple)
                            worksheet.write(i,1,component[0],green)
                            worksheet.write(i,2,0,green)
                            worksheet.write(i,3,component[1],green)
                            worksheet.write(i,4,component[1],green)
                            worksheet.write(i,5,"Added",green)
                            if unit_column !=-1:
                                worksheet.write(i,6,df.iloc[comp_row,unit_column],green)
                            i+=1
                            # BOM_delta.append(tuple((component[0],component[1])))
                            
                    for comp_row, component in enumerate(assembly_BOM):
                        if component[0] not in [x[0] for x in NEW_BOM]:
                            was_written.append(excel_PN[0])
                            if excel_PN[0] in left_to_write:
                                left_to_write.remove(excel_PN[0])
                            worksheet.write(i,0,excel_PN[0],purple)
                            worksheet.write(i,1,component[0],red)
                            worksheet.write(i,2,component[1],red)
                            worksheet.write(i,3,0,red)
                            worksheet.write(i,4,-component[1],red)
                            worksheet.write(i,5,"Removed",red)
                            if unit_column !=-1:
                                worksheet.write(i,6,df.iloc[comp_row,unit_column],red)
                            i+=1
                            # BOM_delta.append(tuple((component[0],-component[1])))
                    
                    # In cases where first file is not xls
                    if excel_PN[0] not in was_written:
                        left_to_write.append(excel_PN[0])
                        was_written.append(excel_PN[0])
                    # print(excel_PN[0])
                    # print(BOM_delta)
                else:
                    pb['value'] += progress_value
                    value_text['text'] = f"{round(pb['value'], 2)}%"
                    root.update()
                    continue
            elif excel_PN and excel_PN[0] not in was_written:
                left_to_write.append(excel_PN[0])
                was_written.append(excel_PN[0])
        except:
            worksheet.write(i,0,f"ERROR with item {item}",error_font)
            i+=1
            pb['value'] += progress_value
            value_text['text'] = f"{round(pb['value'], 2)}%"
            root.update()
            continue
            # return f"Problem with item {item}"
        # if len(item.split('_')) == 1:
        #     pb['value'] += progress_value
        #     value_text['text'] = f"{round(pb['value'], 2)}%"
        #     root.update()
        #     continue
        # if len(item.split('_')[1]) not in [12,13]:
        #     pb['value'] += progress_value
        #     value_text['text'] = f"{round(pb['value'], 2)}%"
        #     root.update()
        #     continue
        
        
        
        pb['value'] += progress_value
        value_text['text'] = f"{round(pb['value'], 2)}%"
        root.update()

    for item_left in left_to_write:
        worksheet.write(i,0,item_left,gray)
        i+=1
    workbook.close()
    return f'{ECO}_BOM_Implementation_Report.xlsx' 
    # df.drop('DOCUMENT PREVIEW',axis=1,inplace=True)




if __name__ == "__main__":
   # stuff only to run when not called via 'import' here
   create_BOM_Implementation_report(None, None, None, None, None, None)

