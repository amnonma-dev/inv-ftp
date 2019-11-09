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
    account_num=5029
    invoice_date= input('Invoice date(yyyymmdd):')
    # invoice_date='20190930'
    outdir = os.path.dirname("outfiles/")
    os.makedirs(outdir,exist_ok=True)

    in_order=False
    service_line=False
    entity_line=False
    total_amount=0
    dest_rows=[]
    invoice_file_name=easygui.fileopenbox(msg="select txt invoice to upload",)
    if invoice_file_name is None:
        exit()
    invoice_file=open(invoice_file_name,'r')
    for row in invoice_file:

        if service_line:
            service_line=False
            service_type=row.split(' ')[0].strip()
            if service_type.lower()=='inbound':
                service_type='IHL'
            elif service_type.lower()=='outbound':
                service_type='OHL'
            else:
                print("skipping tracking id "+str(tracking_id)+". Please check")
                in_order=False
                continue
            dest_row['ServiceType']=service_type 
            dest_row['OrderCharge1']=service_type

        if entity_line:
            entity_line=False
            entity_num=row.split(' ')[0].strip()
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
                if entity_num[0:3].upper()=='IDF':
                    dest_row['CustomerReference#2']=entity_num
                    break
                entity_num= input("Please enter entity number[A1 Tracking id:"+str(tracking_id)+\
                    ". Operation date:"+operation_date.strftime('%Y%m%d')+"]. Type 'none' for no entity:  ")
            dest_rows.append(dest_row)
            in_order=False
            continue
            
        if not in_order:
            try:
                operation_date= datetime.strptime(row.split(' ')[0],'%m/%d/%y')
                in_order=True
                dest_row={}.fromkeys(fields,None)
                tracking_id=row.split(' ')[1].strip()
                dest_row['InvoiceDate']=invoice_date
                dest_row['InvoiceNumber']=invoice_num
                dest_row['ShipmentDate']=operation_date.strftime('%Y%m%d')
                dest_row['SupplierCode']='A1'
                dest_row['RecipientCompany']='A1'
                dest_row['TansportationCharge']=0
                dest_row['TrackingNumber']=0
                dest_row['AccountNumber']=account_num
            except:
                pass
            continue
        if len(row)>=3 and row[:3]=='Pcs':
            pieces=row.split(':')[1]
            pieces=pieces.split(' ')[0]
            dest_row['NumberofPieces']=pieces
        elif row[0]=='$':
            other_charge=float(row[1:])
            if other_charge>0:
                print("skipping tracking id "+str(tracking_id)+". Please check")
                in_order=False
                continue
            service_line=True
        if len(row)>=8 and row[:8]=='Packages':
            charge_amount=row.split('$')[1].split(' ')[0]
            charge_amount=float(charge_amount)
            total_amount+=charge_amount
            dest_row['NetChargeAmount']=round(charge_amount,2)
            dest_row['OrderChargeAmount1']=round(charge_amount,2)
            entity_line=True
        
    for v in dest_rows:
        v['OriginalAmountDue']=round(total_amount,2)
    outfile_name='outfiles\\'+str(invoice_num) +'-inbound-'+ datetime.now().strftime('%Y%m%d%H%M%S')+'.txt'

    with open(outfile_name, 'w',newline='')as outcsv:
        w= csv.DictWriter(outcsv,fields,delimiter='\t')
        w.writeheader()
        w.writerows(dest_rows)
    #remove last line
    with open(outfile_name, 'r') as f:
        data = f.read().rstrip('\n')
    with open(outfile_name, 'w') as f_output:    
        f_output.write(data)