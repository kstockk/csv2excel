
import sys
import openpyxl
import csv

def main(csv_file, wb_name, ws_name):
    wb = openpyxl.load_workbook(wb_name) # Must be xlsx

    for row in wb[ws_name]['A2:Z10000']:
        for cell in row:
            cell.value = None
    
    with open(csv_file, newline='') as f_input:
        ws = wb[ws_name]
        
        for rowy, row in enumerate(csv.reader(f_input, delimiter=','), start=1):
            for colx, value in enumerate(row, start=1): 
                try:
                    value = float(value)
                except ValueError:
                    pass                           
                ws.cell(column=colx, row=rowy, value=value)
            
    wb.save(wb_name)

if __name__ == "__main__":
   main(sys.argv[1], sys.argv[2], sys.argv[3])