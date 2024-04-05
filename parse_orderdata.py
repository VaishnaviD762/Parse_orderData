import pandas as pd
import json

df = pd.read_excel("orderdata.xlsx")

def parse_text(text):
    lines = text.strip().split('\n')
    parsed_data = {}
    current_section = None
    current_var_type = None
    for line in lines:
        if len(line) < 21:
            continue
        sequenceNumber = line[0:6]
        parent = line[6:12]
        sectionNumber = line[12:14]
        varType = line[14]
        action = line[15]
        varName = line[16:21]
        varValue = line[21:].strip()

        if varType == 'S':
            current_section = varName
            current_var_type = None
            parsed_data[current_section] = []
        if current_section is not None and current_var_type not in {' ', 'N'}:
            if varType in {' ', 'N'}:
                if parsed_data[current_section]:
                    parsed_data[current_section][-1]['varValue'] += ' ' + varValue
            else:
                parsed_data[current_section].append({
                    'sqNo': sequenceNumber.strip(),
                    'parent': parent.strip(),
                    'sectionNo': sectionNumber.strip(),
                    'varType': varType.strip(),
                    'action': action.strip(),
                    'varName': varName.strip(),
                    'varValue': varValue.strip()
                })
        if varType not in {' ', 'N'}:
            current_var_type = varType
    return parsed_data

for index, data in df.iterrows():
    try:
        req = json.loads(data['Baseline_Request'])
        orderData_list = req['TOMRulesRequest']['PCSRUSOOrder']
        for orderData_obj in orderData_list:
            orderData = orderData_obj.get('orderData')
            if orderData:
                print(orderData)
                resp = parse_text(orderData)
                print(resp)
    except json.JSONDecodeError as e:
        print(f"Error decoding JSON in row {index}: {e}")
