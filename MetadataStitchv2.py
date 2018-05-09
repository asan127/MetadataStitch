import requests
import os
import MetadataStitch_config

##v2: stitch data into project plan

# Smartsheet API access token
#access_token = '4n4f9q743sr595oirap5l4sze7' #Anthony Sanfilippo's personal token
access_token = MetadataStitch_config.access_token #CPS

sheetCount = 0
affectedSheetCount = 0

headers = {
    'Authorization': 'Bearer {}'.format(access_token),
    'Content-Type': 'application/json',
}

def counter():
    global count
    count = count + 1
h=[]

def clearSheets():
    global h
    h=[]

def FindSheets2(ws):

    if('sheets' in ws):
        h.extend(ws['sheets'])
        
    if('folders' in ws):

        for folder in ws['folders']:
            
            fid = folder["id"]
            folder_url = "https://api.smartsheet.com/2.0/folders/" + str(fid)
            folder_response = requests.request("GET", folder_url, headers=headers)
            folder_object = folder_response.json()

            FindSheets2(folder_object)
    return h


def findColumnIdByIndex(columns, index):
    for col in columns:
        try:
            if(col['index']==index):
                return col
        except ValueError:
            print('Error finding column via index.')

def getColumnById(id, columns):
    for col in columns:
        try:
            if(col['id']==id):
                return col['title']
        except ValueError:
            print('Column not found.')

def getColumnByName(name, columns):
    for col in columns:
        try:
            if(col['title']==name):
                return col['id']
        except ValueError:
            print('Column not found.')

def findPrimary(columns):
    for col in columns:
        try:
            if('primary' in col and col['primary'] == True):
                return col
        except ValueError:
            print('Primary Column not found.')

def getIndex(id, arr):  #gets index of column or row
    for i in arr:
        try:
            if(id == i['id']):
                return i['index']
        except ValueError:
            print("Unable to find index of "+str(id)+" in "+str(arr))
            
    
if __name__ == '__main__':
    intakeSheetId = 5841615085430660


    sheet_url = "https://api.smartsheet.com/2.0/sheets/" + str(intakeSheetId)
    sheet_response = requests.request("GET", sheet_url, headers=headers)
    intakeSheet = sheet_response.json()
    
    
    #Process Intake#
    print('Gathering Intake Data...')

    columnNames = {}
    columnTypes = {}
    columnIds = {}
    
    for col in intakeSheet['columns']:
        columnNames[col['title']]=col['id']
        columnTypes[col['title']]=col['type']
        columnIds[col['id']]=col['title']
    
    #print(columnNames)
    
    projectNameColumn = 'Project Name'
    
    #!!!!!Control Point!!!!!!! Enter Column names to match the desired metadata
    metadataNames = ['Last Updated By','Last Updated On','Health','Health Description','Update Notes','Phase','State','Schedule Start','Release','Delivery','Completion','Schedule % Complete','CIO Approval Status','Approval Status Notes','Scheduled Duration, days','Estimated Work, hrs','Estimated Labor Cost, $','Total Actual Labor Cost, $','Total Planned Labor Cost, $','Total Scheduled Work, hrs','Total Estimated Cost, $','Total Actual Work, hrs','Requirements','Build','Testing','Tech Change Mgmt','Support']
    
    masterMetadata = {}
    projectNameColumnId = getColumnByName(projectNameColumn, intakeSheet['columns'])
    #print(projectNameColumnId)
    
    
    for row in intakeSheet['rows']:

        rowObj = {}
        rowMeta = {}
        projectName = ''
        projectNameCell = next(item for item in row['cells'] if item['columnId']==projectNameColumnId)
        if('value' in projectNameCell):
            
            projectName = projectNameCell['value']
            #print(projectName)
        else: 
            continue
        #print(projectName)
        
        for cell in row['cells']:
            currentColumn = getColumnById(cell['columnId'], intakeSheet['columns'])
            if('value' in cell and currentColumn in metadataNames):
                if(cell['columnId']==projectNameColumnId):
                    projectName = cell['displayValue']

                rowMeta[currentColumn] = str(cell['value'])
                
                rowObj["metadata"]=rowMeta
                masterMetadata[projectName]=rowObj
    #print(masterMetadata)
    
    
    #Start Import#
    print('Stitching...')
    #targetSheetId = 7954817673092
    targetWorkspace = 116553166415748  #!!!!!Critical Control point: target workspace for active projects
    
    ws_url = "https://api.smartsheet.com/2.0/workspaces/" + str(targetWorkspace)
    ws_response = requests.request("GET", ws_url, headers=headers)
    ws = ws_response.json()
    
    allSheets = FindSheets2(ws)
    
    
    for s in allSheets:
        #print(folder['sheets'])
        if(s['name'].endswith('Metadata')):         #!!!!!Critical Control Point: find sheets ending with...
            sheetCount+=1
            
            sheet_url = "https://api.smartsheet.com/2.0/sheets/" + str(s['id'])
            sheet_response = requests.request("GET", sheet_url, headers=headers)
            sheet = sheet_response.json()
            #print(sheet['name'])
            
            metadataKeyCol = findPrimary(sheet['columns'])
            metadataValCol = findColumnIdByIndex(sheet['columns'],metadataKeyCol['index']+1)
            
            metadataDateCol=''
            metadataContactcol=''
            
            for col in sheet['columns']:
                if('Metadata' in col['title']):
                    if('Date' in col['title']):
                        metadataDateCol = col
                    if('Contact' in col['title']):
                        metadataContactCol = col
        #print(masterMetadata)
            
            #find Summary parent row
            metadataHeader = ''
            for row in sheet['rows']:
                for cell in row['cells']:
                    if(cell['columnId']==metadataKeyCol['id']):
                        if('Summary' in cell['value']):   #!!!!!Critical Control Point: Summary keyword
                            metadataHeader = row['id']
                    else:
                        pass
            #print(metadataHeader)
            metadataRows=[]
            
            #find all metadata rows
            for row in sheet['rows']:
                if('parentId' in row and row['parentId']==metadataHeader):
                    metadataRows.append(row['id'])

            rowUpdates = []
            projectName = ''
            
            metaMap = {}
            for r in metadataRows:
                
                targetRow = next(item for item in sheet['rows'] if item['id']==r)
                
                metadataKeyCell = next(item for item in targetRow['cells'] if item['columnId']==metadataKeyCol['id'])
                metadataValCell = next(item for item in targetRow['cells'] if item['columnId']==metadataValCol['id'])
                
                metaMap[metadataKeyCell['value']]=r
                
                if(metadataKeyCell['value']=='Project Name'):
                    projectName = metadataValCell['value']
                

            metaMap.pop('Project Name', None)

            for k,v in metaMap.items():
                if(k in masterMetadata[projectName]['metadata']):
                    targetCol = metadataValCol['id']
                    if(columnTypes[k]=='DATE'):
                        targetCol = metadataDateCol['id']
                    if(columnTypes[k]=='CONTACT_LIST'):
                        targetCol = metadataContactCol['id']
                    
                    rowUpdates.append({'id':v, 'cells':[{'columnId':targetCol, 'value':masterMetadata[projectName]['metadata'][k], 'strict':'false'}]})
                    #print('Found: '+str(k))
                else:
                    #print('Metadata "'+k+'" not found on Intake Sheet.')
                    pass

            print('Updating '+sheet['name']+'...')
            sheet_url = "https://api.smartsheet.com/2.0/sheets/" + str(sheet['id'])+'/rows'
            sheet_response = requests.request("PUT", sheet_url, headers=headers, json=rowUpdates)
            sheet = sheet_response.json()
            print(sheet['message'])
 
    print('Sheets affected: '+str(sheetCount))
    print('Done!')
    