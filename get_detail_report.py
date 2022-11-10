import csv
from tabnanny import verbose
import pandas as pd
from collections import defaultdict
from fpdf import FPDF
import os
import copy
import chardet
import textwrap
from copy import deepcopy
from pyhtml2pdf import converter


#==============================================================================================================================================
#HTML template
#==============================================================================================================================================
html = """<style>
@import "https://fonts.googleapis.com/css?family=Montserrat:300,400,700";
.rwd-table {
  margin: 1em 0;
  min-width: 300px;
  font-family: 'Open Sans', sans-serif;
  border-collapse: collapse;
  font-size: 0.5em;

  
}
.rwd-table tr {
  border-top: 1px black;

}
.rwd-table th {
  display: none;
  border: 1px solid black;
}
.rwd-table td {
  display: block;
}
.rwd-table td:first-child {
  padding-top: .5em;
}
.rwd-table td:last-child {
  padding-bottom: .5em;
}
.rwd-table td:before {
  content: attr(data-th) ": ";
  font-weight: bold;
  width: 6.5em;
  display: inline-block;
}
@media (min-width: 480px) {
  .rwd-table td:before {
    display: none;
  }
}
.rwd-table th, .rwd-table td {
  text-align: left;
}
@media (min-width: 480px) {
  .rwd-table th, .rwd-table td {
    display: table-cell;
    padding: .25em .5em;
  }
  .rwd-table th:first-child, .rwd-table td:first-child {
    padding-left: 0;
  }
  .rwd-table th:last-child, .rwd-table td:last-child {
    padding-right: 0;
  }
}
 
 
h1 {
  font-weight: normal;
  letter-spacing: -1px;
  color: black;
}
 
.rwd-table {
  background: white;
  color: black;
  border-radius: .4em;

}
.rwd-table tr {
  border-color: black;
}
.rwd-table th, .rwd-table td {
  margin: .5em 1em;
}
@media (min-width: 480px) {
  .rwd-table th, .rwd-table td {
    padding: 1em !important;
  }
}
.rwd-table th, .rwd-table td:before {
  color: black;
}
</style>
<script>
  window.console = window.console || function(t) {};
</script>
<script>
  if (document.location.search.match(/type=embed/gi)) {
    window.parent.postMessage("resize", "*");
  }
</script>
"""

#===============================================================================================================================================
# All Function is defined in here
#===============================================================================================================================================
def get_invoice_net_sales(l):
    '''
    get the net sales, (sub_total - cost)
    for each invoice entry(row)
    '''
    cost_index = len(l) - 1
    sales_index = len(l) - 2
    cost_sub_total = float(l[cost_index])
    sales_sub_total = float(l[sales_index])
    netsales = sales_sub_total - cost_sub_total
    return round(netsales,2)

def get_sales_sub_total(l):
    '''
    SubTotal is at index 7
    Discount Amount is at index 8
    sales_sub_total = SubTotal - Discount Amount
    '''
    sales_total = float(l[5]) * float(l[6])
 
    discount_amount = float(l[8])
    sales_subtotal = sales_total - discount_amount
    return round(sales_subtotal,2)

def get_cost_sub_total(invoice_list, item_list):
    '''
    get the sub total for cost
    if there's cf cost, use cf cost
    else use the purchase price
    return a float value
    '''
    quantity = float(invoice_list[5])
    cf_cost = invoice_list[10]
    item_name = invoice_list[4]
    if cf_cost != "":
     
        cf_cost = float(cf_cost)
        cost_sub_total = quantity * cf_cost
        return round(cost_sub_total,2)
    else:
     
        purchase_rate = get_purchase_rate(item_list, item_name)
        cost_sub_total = quantity * purchase_rate
        return round(cost_sub_total,2)
        
def get_purchase_rate(l, item_name):
    '''
    search invoice item name in the item list
    and return the purchase rate
    return float value
    '''
    for item in l:
        if item[0] == item_name:
            return float(item[2])
    return 0
    
def get_credit_note_cost_sub_total(credit_note,item_list):


    item_name = credit_note[4]
    item_quantity = float(credit_note[5])
    discount = float(credit_note[8])
    item_cost = get_purchase_rate(item_list, item_name)

    sub_cost = (item_quantity * item_cost) - discount

    if credit_note[10] != '':
        cf_cost = float(credit_note[10])
        cf_sub_cost = (item_quantity * cf_cost) - discount
        return cf_sub_cost
    
    return sub_cost

def get_sales_person_data(sales_name):
    sales_data = []
    invoice_sales_total = 0
    invoice_cost_total = 0
    credit_sales_total = 0
    credit_cost_total = 0
    marginal_profit = 0

    for invoice in invoice_raw:
        if sales_name == invoice[3]:
            invoice_sales_total += float(invoice[13])
            invoice_cost_total += float(invoice[14])
            marginal_profit += float(invoice[15])
    
    for creditnote in credit_note_data_raw:
        if sales_name == creditnote[3]:
            credit_sales_total += float(creditnote[13])
            credit_cost_total += float(creditnote[14])
            marginal_profit -= float(creditnote[15])
    
    sales_total = invoice_sales_total - credit_sales_total
    sales_total = round(sales_total,2)
    cost_total = invoice_cost_total - credit_cost_total
    cost_total = round(cost_total,2)
    marginal_profit_percentage = marginal_profit / sales_total if sales_total > 0 else 0

    marginal_profit_percentage = str(round((marginal_profit_percentage * 100),2)) + '%'

    sales_data.append(sales_name)
    sales_data.append(sales_total)
    sales_data.append(cost_total)
    sales_data.append(marginal_profit)
    sales_data.append(marginal_profit_percentage)
    return sales_data

def get_grand_total(l):
    sales_total = 0
    cost_total = 0
    marginal_profit = 0
    marginal_profit_percentage = 0
    grand_total = ['Grand Total']

    for sales in l:
        sales_total += sales[1]
        cost_total += sales[2]
        marginal_profit += sales[3]
    marginal_profit_percentage = marginal_profit/sales_total if sales_total > 0 else 0

    marginal_profit_percentage = round(marginal_profit_percentage * 100,2)
    marginal_profit_percentage = str(marginal_profit_percentage) + '%'

    grand_total.append(round(sales_total,2))
    grand_total.append(round(cost_total,2))
    grand_total.append(round(marginal_profit,2))
    grand_total.append(marginal_profit_percentage)
    return grand_total


def get_sales_rows(sales_name):
    '''
    return all the row in the invoice/credit 
    note that realte to particular sales person
    return list(list(invoice/creditnote))
    '''

    sales_rows = []
    total_cost = 0
    total_sales = 0
    net_sales = 0

   
    
    for invoice in invoice_raw:
        if sales_name == invoice[3]:
            sales_rows.append(invoice)
            total_sales  += float(invoice[13])
            total_cost  += float(invoice[14])
            net_sales += float(invoice[15])
    
    for creditnote in credit_note_data_raw:
        if sales_name == creditnote[3]:
            sales_rows.append(creditnote)
            total_sales  -= float(creditnote[13])
            total_cost  -= float(creditnote[14])
            net_sales -= float(creditnote[15])

    sales_rows.sort()
    last_row = ["#","","",sales_name,"","","","","","","","","",total_sales,total_cost,net_sales]
    sales_rows.append(last_row)


    return sales_rows

def make_csv_report(file_path,data):

    with open(file_path, 'w', newline='') as file:
        write = csv.writer(file)
        

        for invoice in data:
            write.writerow(invoice)

def convert_csv_html(csv_file, file_name):
    with open(csv_file, 'rb') as f:
        result = chardet.detect(f.read()) 

    df = pd.read_csv(csv_file,encoding=result['encoding'])
    df = df.fillna('')
    df.to_html(file_name)
    with open(file_name) as file:
     file = file.read()
     file = file.replace("<table ", "<table class='rwd-table'")
    with open(file_name, "w") as file_to_write:
        file_to_write.write(html + file)
  
def convert_table_pdf(html_file, pdf_file):
  
  # path = os.path.abspath('update_invoice.html')
  
    converter.convert(f'file:///{html_file}', pdf_file)

#===============================================================================================================================================================  
# Main operation starts from here 
#================================================================================================================================================================

#make dir for final reports in csv hml pdf, form
CURRENT_PATH = os.getcwd()
pdfpath = 'PDF Reports'
csvpath = 'CSV Reports'
htmlpath = 'HTML Reports'
# if not os.path.exists(pdfpath):
#     os.mkdir('PDF Reports')
if not os.path.exists(pdfpath):
    os.mkdir(pdfpath)
if not os.path.exists(csvpath):
    os.mkdir('CSV Reports')
if not os.path.exists(htmlpath):
    os.mkdir('HTML Reports')

#read item.csv and store the data in to a list
sales_person_set = set()
item_data = []



with open('Item.csv', 'r', newline='', encoding='ISO-8859-1') as csvfile:
    item_reader = csv.reader(csvfile)
    for row in item_reader:
        item_data.append(row)
        
item_header = item_data[0]
item_data_raw = item_data[1:]

#convert the format of each item from ['Shipping Charges*', 'Cad 0.0', 'Cad 0.0']
#to ['Shipping Charges*', 0.0, 0.0]
for item in item_data_raw:
    item[1] = float(item[1][4:])
    item[2] = float(item[2][4:])

#print(item_header)
#print(item_data_raw[-1])
#print (get_purchase_rate(item_data_raw,"GX-SMT-XHD784F-IZ"))


#read invoice.csv and store the data in to a list
invoice_data = []
with open('Invoice.csv', 'r', newline='', encoding='ISO-8859-1') as csvfile:

    reader = csv.reader(csvfile)

    for row in reader:
        if row[11] != "Draft":
            invoice_data.append(row)


#Get rid of header
invoice_original_raw = invoice_data[1:]
invoice_raw = invoice_data[1:]
invoice_header = invoice_data[0]

#put sales sub total and cost sub total in to the column
invoice_header.append("Purchase Rate")
invoice_header.append("Sales Sub Total")
invoice_header.append("Cost Sub Total")
invoice_header.append("Net Sales")

# print(invoice_header)


#calculate the sales sub total and cost subtotal for each row
#and insert them in to the last 2 index of the list
#[... , sales_sub_total, cost_subtotal]

for invoice in invoice_raw:
   item_cost = round(float(invoice[6]),2)
   invoice[6] = str(item_cost)
   invoice[0] =  str(pd.to_datetime(invoice[0]).date())
   sales_person = invoice[3]
   item_name = invoice[4]
   item_cost = get_purchase_rate(item_data_raw, item_name)
   sales_person_set.add(sales_person)
   invoice.append(item_cost)
   invoice.append(get_sales_sub_total(invoice))
   invoice.append(get_cost_sub_total(invoice, item_data_raw))
   invoice.append(get_invoice_net_sales(invoice))

invoice_raw.insert(0, invoice_header)
#update the invoice with 2 new column

update_invoice_path = os.path.join(CURRENT_PATH,"./CSV Reports/update_invoice.csv")
update_invoice_html_path = os.path.join(CURRENT_PATH,"./HTML Reports/update_invoice.html")
update_invoice_pdf_path = os.path.join(CURRENT_PATH,"./PDF Reports/update_invoice.pdf")

update_invoice_abs_html_path = os.path.abspath('./HTML Reports/update_invoice.html')
update_invoice_abs_pdf_path = os.path.abspath('./PDF Reports/update_invoice.pdf')

make_csv_report(update_invoice_path, invoice_raw)
convert_csv_html(update_invoice_path, update_invoice_html_path)

convert_table_pdf(invoice_raw, update_invoice_pdf_path)
convert_table_pdf(update_invoice_abs_html_path, update_invoice_abs_pdf_path)


#===================================================================

#get Credit Note


credit_note_data = []
with open('Credit_Note.csv', 'r', newline='', encoding='ISO-8859-1') as csvfile:

    reader = csv.reader(csvfile)

    for row in reader:
        if row[11] != "Draft":
            credit_note_data.append(row)

credit_note_data_header = credit_note_data[0]
credit_note_data_raw = credit_note_data[1:]
credit_note_data_original_raw = credit_note_data[1:]


for creditnote in credit_note_data_raw:
    sales_person = creditnote[3]
    sales_person_set.add(sales_person)
    creditnote[0] =  str(pd.to_datetime(creditnote[0]).date())
    item_name = creditnote[4]
    quantity = float(creditnote[5])
    item_price = float(creditnote[6])
    discount_amount = round(float(creditnote[8]),2)

    # if creditnote[10] != '':
    #     cf_cost = float(creditnote[10])
    #     credit_note_sub_total = (quantity * cf_cost) - discount_amount
    #     creditnote.append(credit_note_sub_total)
    # else:

    credit_note_sub_total = (quantity * item_price) - discount_amount
    credit_note_cost_sub_total = get_credit_note_cost_sub_total(creditnote, item_data_raw)
    item_cost = get_purchase_rate(item_data_raw,item_name)
    credit_note_net_sales = credit_note_sub_total - credit_note_cost_sub_total
  
    creditnote.append(item_cost)
    creditnote.append(credit_note_sub_total)
    creditnote.append(credit_note_cost_sub_total)
    creditnote.append(credit_note_net_sales)

credit_note_data_raw.insert(0, invoice_header)

#add the new header to credit note  "credit note sub total" and write it to the update credit note csv
credit_note_data_header.append("Purchase Rate")
credit_note_data_header.append("Credit Note Sub Total")
credit_note_data_header.append("Credit Note Sub  Cost Total")
credit_note_data_header.append("CN Net Sales")

update_credit_note_path = os.path.join(CURRENT_PATH,"./CSV Reports/update_credit_note.csv")
update_credit_note_html_path = os.path.join(CURRENT_PATH,"./HTML Reports/update_credit_note.html")

update_credit_abs_html_path = os.path.abspath('./HTML Reports/update_credit_note.html')
update_credit_abs_pdf_path = os.path.abspath('./PDF Reports/update_credit_note.pdf')

make_csv_report(update_credit_note_path, credit_note_data_raw)
convert_csv_html(update_credit_note_path, update_credit_note_html_path)
convert_table_pdf(update_credit_abs_html_path,update_credit_abs_pdf_path )


#================================================================================================================================================================
# generate final report
#================================================================================================================================================================

final_report_header = ['Sales person', 'Sales Total', 'Cost Total', 'Marginal Profit', 'Marginal Profit%']
final_report_data = []

# final_report_data.append(final_report_header)

for sales_person in sales_person_set:
    data = get_sales_person_data(sales_person)
    final_report_data.append(data)
    
grand_total = get_grand_total(final_report_data)
final_report_data.append(grand_total)
final_report_data.insert(0, final_report_header)
final_commission_report_path = os.path.join(CURRENT_PATH,"./CSV Reports/final_commission_report.csv")
final_commission_report_html_path = os.path.join(CURRENT_PATH,"./HTML Reports/final_commission_report.html")

final_commission_report_abs_html_path = os.path.abspath('./HTML Reports/final_commission_report.html')
final_commission_report_abs_pdf_path = os.path.abspath('./PDF Reports/final_commission_report.pdf')

make_csv_report(final_commission_report_path, final_report_data)
convert_csv_html(final_commission_report_path, final_commission_report_html_path)
convert_table_pdf(final_commission_report_abs_html_path, final_commission_report_abs_pdf_path)
#===========================================================
#make verbose reports
#===========================================================

verbose_report = invoice_original_raw + credit_note_data_original_raw

verbose_report.sort()
verbose_report.insert(0, invoice_header)
verbose_commission_report_path = os.path.join(CURRENT_PATH,"./CSV Reports/verbose_commission_report.csv")
verbose_commission_report_html_path = os.path.join(CURRENT_PATH,"./HTML Reports/verbose_commission_report.html")

verbose_commission_report_abs_html_path = os.path.abspath('./HTML Reports/verbose_commission_report.html')
verbose_commission_report_abs_pdf_path = os.path.abspath('./PDF Reports/verbose_commission_report.pdf')

make_csv_report(verbose_commission_report_path, verbose_report)
convert_csv_html(verbose_commission_report_path, verbose_commission_report_html_path)
convert_table_pdf(verbose_commission_report_abs_html_path, verbose_commission_report_abs_pdf_path)




sales_verbose_data = []

for sales_name in sales_person_set:
    
    sales_row = get_sales_rows(sales_name)
    sales_verbose_data += sales_row
    csv_file_path = "./CSV Reports/" + sales_name + " commission_report.csv"
    csv_file_html_path = "./HTML Reports/" + sales_name + " commission_report.html"
    pdf_file_path = "./PDF Reports/" + sales_name + " commission_report.pdf"
    sales_commission_path = os.path.join(CURRENT_PATH, csv_file_path)
    sales_commission_html_path = os.path.join(CURRENT_PATH, csv_file_html_path)
    sales_row.insert(0, invoice_header)

    sales_commission_report_abs_html_path = os.path.abspath(csv_file_html_path)
    sales_commission_report_abs_pdf_path = os.path.abspath(pdf_file_path)

    make_csv_report(sales_commission_path, sales_row)
    convert_csv_html(sales_commission_path, sales_commission_html_path)
    convert_table_pdf(sales_commission_report_abs_html_path, sales_commission_report_abs_pdf_path)

sales_verbose_commission_report_path = os.path.join(CURRENT_PATH,"./CSV Reports/sales_verbose_commission_report.csv")
sales_verbose_commission_report_html_path = os.path.join(CURRENT_PATH,"./HTML Reports/sales_verbose_commission_report.html")


sales_verbose_commission_report_abs_html_path = os.path.abspath("./HTML Reports/sales_verbose_commission_report.html")
sales_verbose_commission_report_abs_pdf_path = os.path.abspath("./PDF Reports/sales_verbose_commission_report.pdf")

sales_verbose_data.insert(0, invoice_header)

make_csv_report(sales_verbose_commission_report_path, sales_verbose_data)
convert_csv_html(sales_verbose_commission_report_path, sales_verbose_commission_report_html_path)
convert_table_pdf(sales_verbose_commission_report_abs_html_path, sales_verbose_commission_report_abs_pdf_path)


#==========================================================================================================
#Make verbose report in sorted sales person format, and output in pdf
#==========================================================================================================

# pdf = FPDF()
# pdf.add_page()
# pdf.set_font("Times", size=10)
# line_height = pdf.font_size * 2.5
# col_width = pdf.epw / 4  # distribute content evenly
print("all sales person")
print(sales_person_set)
print("Operation Complete")

