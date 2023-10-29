import sys
import os
import pandas as pd

def createEppFile(data):
    f = open(output_filename, "w")
    f.write('[INFO]\n"1.05",3,1250,,,,,,,,,,,,,0,,,,20231028211044,,"PL","PL0000000000",1\n\n[NAGLOWEK]\n"TOWARY"\n[ZAWARTOSC]\n')
    for x in range(count_rows):
        object = data[x]

        f.write(f'1,"{object["symbol"]}","{object["symbol"]}",,"{object["name"]}","{object["description"]}","{object["name"]}",,,"{object["unit"]}","{object["vat"]}",-1.0000,"{object["vat"]}",-1.0000,0.0000,0.0000,,0,,,,{object["state"]}.0000,0,,,0,"{object["unit"]}",0.0000,0.0000,,0,,0,0,,,,,,,,\n')

    f.write('\n[NAGLOWEK]\n"CENNIK"\n[ZAWARTOSC]\n')
    for x in range(count_rows):
        object = data[x]

        f.write(f'"{object["symbol"]}","Detaliczna",{object["netto"]}000,{object["brutto"]}000,10.0000,100.0000,{object["netto"]}000\n')

    f.close()

def fetchData():
    json_data = []

    for x in range(count_rows):
        fetched_cols_data = []

        for y in range(count_cols):
            fetched_cols_data.append(data_in.iat[x, y])

        json_data.append({
            "symbol": '{:010}'.format(fetched_cols_data[0]),
            "name": fetched_cols_data[1],
            "description": fetched_cols_data[2],
            "state": fetched_cols_data[3],
            "brutto": fetched_cols_data[4],
            "netto": fetched_cols_data[5],
            "vat": fetched_cols_data[6],
            "unit": fetched_cols_data[7],
        },)

    return json_data
        

# Main ========================================================

if len(sys.argv) < 2: 
    print("Please type your file path in arguments!")
    exit()

exel_file = sys.argv[1]

output_filename = os.path.basename(exel_file)[:-5] + ".epp"
data_in = pd.read_excel(exel_file)

count_rows, count_cols = data_in.shape

if os.path.exists(output_filename):
    os.remove(output_filename)

createEppFile(fetchData())