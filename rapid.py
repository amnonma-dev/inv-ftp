import os
import csv
import easygui
from datetime import datetime


def RepresentsInt(s):
    if s is None:
        return False
    try: 
        int(s)
        return True
    except ValueError:
        return False
def main():
    fields=['SupplierCode','AccountNumber','InvoiceDate','InvoiceNumber','OriginalAmountDue'\
    ,'TrackingNumber','TansportationCharge','NetChargeAmount','ServiceType','ShipmentDate'\
    ,'ActualWeightAmount','RatedWeightAmount','RatedWeightUnits','NumberofPieces','RecipientCompany'\
    ,'RecipientZipCode','RecipientCountry','ShipperZipCode','ShipperCountry','CustomerReference'\
    ,'CustomerReference#2','ZoneCode','OrderCharge1','OrderChargeAmount1','OrderCharge2','OrderChargeAmount2']

    invoice_num= input('Invoice number:')
    # invoice_num='502912750'
    account_num='4LOGRMG'
    invoice_date= input('Invoice date(yyyymmdd):')
    # invoice_date='20190930'
    outdir = os.path.dirname("outfiles/")
    os.makedirs(outdir,exist_ok=True)

    total_amount=0
    dest_rows=[]
    invoice_file_name=easygui.fileopenbox(msg="select txt invoice to upload",)
    if invoice_file_name is None:
        exit()
    invoice_file=open(invoice_file_name,'r')
    for row in invoice_file:
        row = row.strip()
        if len(row)>6 and row[:4].lower()=='road':
            dest_row={}.fromkeys(fields,None)
            entity_num=row[-6:]
            while True:
                if entity_num=='none':
                    break                
                if RepresentsInt(entity_num):
                    if len(str(entity_num))==6:
                        if str(entity_num)[0]=='2':
                            dest_row['CustomerReference#2']=entity_num
                            break
                        elif str(entity_num)[0]=='4':
                            dest_row['CustomerReference']=entity_num  
                            break                        
                entity_num= input("Please enter entity number[invoice entity:"+str(entity_num)+". Type 'none' for no entity:  ")
            dest_row['InvoiceDate']=invoice_date
            dest_row['InvoiceNumber']=invoice_num
            dest_row['ServiceType']="AWF"
            dest_row['OrderCharge1']="AWF"
            dest_row['ShipmentDate']=invoice_date
            dest_row['SupplierCode']='RL'
            dest_row['RecipientCompany']='Rapid Logistics'
            dest_row['TansportationCharge']=0
            dest_row['TrackingNumber']=0
            dest_row['AccountNumber']=account_num
            dest_row['NumberofPieces']=1
            charge_amount=6.5
            charge_amount=float(charge_amount)
            total_amount+=charge_amount
            dest_row['NetChargeAmount']=round(charge_amount,2)
            dest_row['OrderChargeAmount1']=round(charge_amount,2)
            dest_rows.append(dest_row)
    for v in dest_rows:
        v['OriginalAmountDue']=round(total_amount,2)
    outfile_name='outfiles\\'+str(invoice_num) +'-'+ datetime.now().strftime('%Y%m%d%H%M%S')+'.txt'

    with open(outfile_name, 'w',newline='')as outcsv:
        w= csv.DictWriter(outcsv,fields,delimiter='\t')
        w.writeheader()
        w.writerows(dest_rows)
    #remove last line
    with open(outfile_name, 'r') as f:
        data = f.read().rstrip('\n')
    with open(outfile_name, 'w') as f_output:    
        f_output.write(data)