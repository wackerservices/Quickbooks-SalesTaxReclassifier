#run Python 32 bit
import win32com.client as win32
import pandas as pd
import json
import xmltodict, json

from datetime import datetime, timedelta
import re
import numpy as np
import warnings
import os
import sys
import requests
warnings.filterwarnings("ignore")
from datetime import datetime
#packages for email
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

#create a folder for JSON output folder when it does not exist
if not os.path.isdir(os.getcwd()+'/output'):ir(os.getcwd()+'/output')

#helper function to convert data types to be Json appropriate 
class NpEncoder(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, np.integer):
            return int(obj)
        if isinstance(obj, np.floating):
            return float(obj)
        if isinstance(obj, np.ndarray):
            return obj.tolist()
        return super(NpEncoder, self).default(obj)


#helper function to query QuickBook SDK
def qb_request(xml_query):
    #initilization and end seesion code only happen once, hence they are put outside of the helper function
    response = qbxml.ProcessRequest(ticket, xml_query)
    return response ## replace this with response parsing

def invoice_ids_query(txn_date_start, txn_date_end):
    #query all invoices
    invoice_query = f'''<?xml version="1.0" encoding="utf-8"?>
    <?qbxml version="8.0"?>
    <QBXML>
        <QBXMLMsgsRq onError="stopOnError">
            <InvoiceQueryRq  metaData="MetaDataAndResponseData" > 
                <TxnDateRangeFilter> <!-- optional -->
                    <FromTxnDate >{txn_date_start}</FromTxnDate> <!-- optional -->
                    <ToTxnDate >{txn_date_end}</ToTxnDate> <!-- optional -->
                </TxnDateRangeFilter>
                <IncludeRetElement >TxnID</IncludeRetElement> <!-- optional, may repeat -->           
            </InvoiceQueryRq>
        </QBXMLMsgsRq>
    </QBXML>'''
    return invoice_query

def single_invoice_query(TxnId):
    #query a single invoice
    invoice_query = f'''<?xml version="1.0" encoding="utf-8"?>
    <?qbxml version="8.0"?>
    <QBXML>
        <QBXMLMsgsRq onError="stopOnError">
            <InvoiceQueryRq  metaData="MetaDataAndResponseData" > 
                <TxnID>{TxnId}</TxnID>   
                <IncludeLineItems >true</IncludeLineItems> <!-- optional -->
                <IncludeLinkedTxns >true</IncludeLinkedTxns>
                <IncludeRetElement >TxnID</IncludeRetElement> <!-- optional, may repeat -->
                <IncludeRetElement >Name</IncludeRetElement> <!-- optional, may repeat -->
                <IncludeRetElement >BillAddress</IncludeRetElement> <!-- optional, may repeat -->
                <IncludeRetElement >ShipAddress</IncludeRetElement> <!-- optional, may repeat -->
                <IncludeRetElement >FullName</IncludeRetElement> <!-- optional, may repeat -->  
                <IncludeRetElement >InvoiceLineRet</IncludeRetElement> 
                <IncludeRetElement >InvoiceLineGroupRet</IncludeRetElement> 
                <IncludeRetElement> EditSequence </IncludeRetElement>
                <IncludeRetElement> RefNumber </IncludeRetElement>            
            </InvoiceQueryRq>
        </QBXMLMsgsRq>
    </QBXML>'''
    return invoice_query

def sales_tax_item_query(name):
    #query item by name
    sales_tax_query = f"""<?xml version="1.0" encoding="utf-8"?>
<?qbxml version="8.0"?>
            <QBXML>
                    <QBXMLMsgsRq onError="stopOnError">
                            <ItemQueryRq metaData="MetaDataAndResponseData">
                                    <FullName>{name}</FullName>
                                    <IncludeRetElement>Name</IncludeRetElement>
                                    <IncludeRetElement>ListID</IncludeRetElement> <!-- optional, may repeat -->
                                    <IncludeRetElement>FullName</IncludeRetElement> <!-- optional, may repeat -->
                                    <IncludeRetElement>IsActive</IncludeRetElement> <!-- optional, may repeat -->
                            </ItemQueryRq>
                    </QBXMLMsgsRq>
            </QBXML>
    """
    return sales_tax_query

def invoice_mod_query(txn_id, edit_sequence, invoice_lines, new_sales_tax_name, new_item_list_id):
    #query to modify invoice, and change the line item
    invoice_mod_query_string = f'''<?xml version="1.0" encoding="utf-8"?>
<?qbxml version="8.0"?>
<QBXML>
        <QBXMLMsgsRq onError="stopOnError">
                <InvoiceModRq>
                    <InvoiceMod>
                        <TxnID >{txn_id}</TxnID>
                        <EditSequence>{edit_sequence}</EditSequence>
    '''
    for line in invoice_lines:
        # include transaction line id for each invoice line you want to keep
        invoice_mod_query_string += f'''
            <InvoiceLineMod>
                <TxnLineID >{line['TxnLineID']}</TxnLineID> '''
        
        if 'ItemRef' in line and line['ItemRef']['FullName'] == 'Sales Tax':
            # add new line name for sales tax 
            invoice_mod_query_string += f'''
            <ItemRef>
                <ListID>{new_item_list_id}</ListID> 
                <FullName>{new_sales_tax_name}</FullName>
            </ItemRef>
            <Rate>{line['Rate']}</Rate>
            <Amount>{line['Amount']}</Amount>
            '''
        
        invoice_mod_query_string += '''
        </InvoiceLineMod>
        '''

    invoice_mod_query_string += '''
            </InvoiceMod>
            <IncludeRetElement>TxnID</IncludeRetElement>
            </InvoiceModRq>
        </QBXMLMsgsRq>
    </QBXML>               
    '''  
    return invoice_mod_query_string

# standard procedure to produce the Json file
def run_scripts(date_range,test_flag = False):
    global qbxml
    global ticket
    global output_path
    global state_codes
    global sales_tax_list_ids

    with open('./static/state_codes.json') as f:
        state_codes = json.load(f)
    sales_tax_list_ids = {}
    start_time = datetime.now().replace(microsecond=0)  
    txn_date_start = date_range['start'].strftime('%Y-%m-%d')
    txn_date_end = date_range['end'].strftime('%Y-%m-%d')
    print('Start: ' + txn_date_start)
    print('End: ' + txn_date_end)
    try:
        with open("static/config.json", "r") as f:
            config = json.load(f)
        #use test comapny file when testing
        if test_flag:
            company_file = config['TEST_FILE']
            # user = config['TEST_USER']
            # password = config['TEST_PASSWORD']
        else: 
            company_file = config['PROD_FILE']
            # user = config['USER']
            # password = config['PASSWORD']
        
        #begin session
        print(company_file)

        #run the request processor
        qbxml = win32.Dispatch('QBXMLRP2.RequestProcessor')
        #open connection
        qbxml.OpenConnection2("", "SalesTaxReclassifier", 1) # Connection type, 1 = localQBD (means local quick books desktop)
        ticket = qbxml.BeginSession(company_file, 2) # Session connect mode, 2 = multi-user mode
        invoice_ids_response = qb_request(invoice_ids_query(txn_date_start, txn_date_end))
        status_code = pd.read_xml(invoice_ids_response, xpath=".//InvoiceQueryRs")['statusCode'][0]
        #check if the response include valid transactions.
        #If status code == 0, then the response is valid. The function will return data frame with all transactions.
        #Else if status code == 1, then the response is invalid. The function will return False
        if status_code == 0:
            # invoices_df = pd.read_xml(invoices_response, xpath=".//InvoiceRet")
            # for column in invoices_df:
            # print(invoices_df[column])
            total_to_update = 0
            invoice_ids_json = xmltodict.parse(invoice_ids_response)
            print(invoice_ids_json['QBXML']['QBXMLMsgsRs']['InvoiceQueryRs']['@retCount'])
            errors = []
            success = []
            invoices_to_be_updated = []
            for invoice_id in invoice_ids_json['QBXML']['QBXMLMsgsRs']['InvoiceQueryRs']['InvoiceRet']:
                invoice = qb_request(single_invoice_query(invoice_id['TxnID']))
                status_code = pd.read_xml(invoice, xpath=".//InvoiceQueryRs")['statusCode'][0]
                # check whether there is a Sales Tax line that needs to be updated
                need_to_update = False
                if status_code == 0:
                    invoice_json = xmltodict.parse(invoice)
                    invoice_ret = invoice_json['QBXML']['QBXMLMsgsRs']['InvoiceQueryRs']['InvoiceRet']
                    invoice_items = invoice_ret['InvoiceLineRet']
                    # invoice line items could be either of type dict or list based on the number of line items associated with the invoice
                    if type(invoice_items) is dict:
                        invoice_items = [invoice_items]
                    for item in invoice_items:
                        if 'ItemRef' in item and item['ItemRef']['FullName'] == 'Sales Tax':
                            need_to_update = True
                    
                    if need_to_update:
                        total_to_update += 1
                        invoices_to_be_updated.append(invoice_ret)
                        # form invoice mod query
                        if 'ShipAddress' in invoice_ret:
                            try:
                                ship_to_state = invoice_ret['ShipAddress']['State']
                                sales_tax_string = 'Sales Tax:' + state_codes.get(ship_to_state)
                            except Exception as err :
                                errors.append({'RefNumber': invoice_ret['RefNumber'], 'message': 'Update failed. Error: ' + 'Key error'})
                                continue
                        else:
                            errors.append({'RefNumber': invoice_ret['RefNumber'], 'message': 'No ShipAddress found.'})
                            continue
                        if sales_tax_string in sales_tax_list_ids:
                            #check if we have already read in the sales tax list id
                            item_ret = sales_tax_list_ids[sales_tax_string]
                        else:
                            #query qb to get the list id
                            sales_tax_item_rsp = qb_request(sales_tax_item_query(sales_tax_string))
                            status_code = pd.read_xml(sales_tax_item_rsp, xpath=".//ItemQueryRs")['statusCode'][0]
                            if status_code != 0:
                                errors.append({'RefNumber': invoice_ret['RefNumber'], 'message': 'Failed to get sales tax item. Response XML: ' + sales_tax_item_rsp})
                                continue
                            sales_tax_item_json = xmltodict.parse(sales_tax_item_rsp)
                            item_ret = sales_tax_item_json['QBXML']['QBXMLMsgsRs']['ItemQueryRs']['ItemOtherChargeRet']
                            sales_tax_list_ids[sales_tax_string] = item_ret
                        invoice_mod_rsp = qb_request(invoice_mod_query(invoice_ret['TxnID'], 
                                                    invoice_ret['EditSequence'],
                                                    invoice_items,
                                                    sales_tax_string,
                                                    sales_tax_list_ids[sales_tax_string]['ListID']))
                        status_code = pd.read_xml(invoice_mod_rsp, xpath=".//InvoiceModRs")['statusCode'][0]
                        if status_code != 0:
                            errors.append({'RefNumber': invoice_ret['RefNumber'], 'message': 'Update failed. Response XML: ' + invoice_mod_rsp})
                        else:
                            success.append({'RefNumber': invoice_ret['RefNumber'], 'message': 'Update successful. Response XML: ' + invoice_mod_rsp})
            #write errors and success messages to file
            print(total_to_update)
            with open(f'./output/errors_month_of_{txn_date_start}.json', 'w') as f:
                f.write(json.dumps({'errors':errors}, indent=4))
            with open(f'./output/success_month_of_{txn_date_start}.json', 'w') as f:
                f.write(json.dumps({'success':success}, indent=4))
            with open(f'./output/invoices_month_of_{txn_date_start}.json', 'w') as f:
                f.write(json.dumps({'invoices_to_be_updated':invoices_to_be_updated}, indent=4))
        else:
            print(invoice_ids_response)

        # response = qb_request(sales_tax_item_query('Sales Tax:Alabama'))
        # with open('out.txt', "w") as f:
        #     f.write(response)
        
    except Exception as e:
        #record log
        end_time = datetime.now().replace(microsecond=0)
        body = f'Program start time: {start_time}. Duration: {end_time - start_time}. Error: {e}, Error line number:{sys.exc_info()[-1].tb_lineno}, '
        print(body)
    finally:
        # end session
        if (ticket != None):
            qbxml.EndSession(ticket)
        # close connection
        qbxml.CloseConnection()

#run the script to update all invoices with Sales Tax items
if __name__ == '__main__':
    date_range_start = datetime(2021,9,28)
    date_range_end = datetime(2021,10, 28)
    while date_range_start < date_range_end:
        date_range = {'start': date_range_start, 'end': date_range_start + timedelta(days=30)}
        run_scripts(date_range,False)
        date_range_start += timedelta(days=30)
