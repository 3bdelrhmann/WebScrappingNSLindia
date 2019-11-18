from urllib.request import Request, urlopen
from bs4 import BeautifulSoup as BS
import sys,os
import random
import pandas as pd
from pandas import DataFrame
import xlwt

req = Request('https://www.nseindia.com/live_market/dynaContent/live_watch/option_chain/optionKeys.jsp?symbolCode=-10006&symbol=NIFTY&symbol=NIFTY&instrument=-&date=-&segmentLink=17&symbolCount=2&segmentLink=17', headers={'User-Agent': 'Mozilla/5.0'})
webpage = urlopen(req).read()
soup = BS(webpage,'html.parser')

orig_stdout = sys.stdout
randNum = random.randrange(9999)
Html_file_name = 'out'+str(randNum)+'.html'
Open_html = open(Html_file_name, 'w')
sys.stdout = Open_html

reomveCalls = soup.find("th", {"colspan" : "11"})
reomveCalls.decompose()
RemovePuts = soup.find("th", {"colspan" : "11"})
RemovePuts.decompose()
EmptyTH = soup.select(".opttbldata #octable thead tr")
for d in EmptyTH:
    if len(d.get_text(strip=True)) == 0:
        d.extract()
        break

print(soup.prettify().encode("utf-8"))
sys.stdout = orig_stdout
Open_html.close()

wb = xlwt.Workbook()
ws = wb.add_sheet('a test sheet',cell_overwrite_ok=True)
columns = ['Chart',
           'OI', 'Chng in OI', 'Volume', 'IV',
           'LTP', 'Net Chng', 'Bid Qty', 'Bid Price',
           'Ask Price', 'Ask Qty', 'Strike Price', 'Bid Qty',
           'Bid Price', 'E-mail', 'Nascimento', 'Criado em',
           'Bid Price', 'Ask Price', 'Ask Qty', 'Net Chng',
           'LTP', 'IV', 'Volume', 'Chng in OI','OI',
           'Chart'
           ]
row_num = 0           
for col_num in range(len(columns)):
    ws.write(row_num, col_num, columns[col_num])

with open(Html_file_name,'r') as f:
   webpage_contents = f.read()
   webpage_contents_read = BS(webpage_contents,'html.parser')
   table = webpage_contents_read.find("table",id="octable")
   rows = table.findAll("tr")
   
   x = 1
   for tr in rows:
        cols = tr.findAll("td")
        if not cols: 
        # when we hit an empty row, we should not print anything to the workbook
            continue
        y = 0
        columns = ['Chart',
                'OI', 'Chng in OI', 'Volume', 'IV',
                'LTP', 'Net Chng', 'Bid Qty', 'Bid Price',
                'Ask Price', 'Ask Qty', 'Strike Price', 'Bid Qty',
                'Bid Price', 'Ask Price', 'Ask Qty','Net Chng',
                'LTP','IV','Volume','Chng in OI','OI',
                'Chart'
                ]
        row_num = 0           
        for col_num in range(len(columns)):
            ws.write(row_num, col_num, columns[col_num])

        for td in cols:
            texte_bu = td.text
            texte_bu = texte_bu.encode('utf-8')
            texte_bu = texte_bu.strip()
            td_content = td.text
            removeSlashFromData = td_content.replace('\\n','')
            removeSpacesFromData = removeSlashFromData.replace(' ','')
            ws.write(x, y, removeSpacesFromData)
            print(x, y, removeSpacesFromData)
            y = y + 1
    # update the row pointer AFTER a row has been printed
    # this avoids the blank row at the top of your table
        x = x + 1
 
wb.save('BlastResults.xls')   
#    columns = []
#    DataRows = []
#    for column in webpage_contents_read.select('.opttbldata #octable thead th '):
#        columns_name = column.text
#        removeSlash = columns_name.replace('\\n','')
#        removeSpaces = removeSlash.replace(' ','')
#        columns.append(removeSpaces)
   
#    for DataRow in webpage_contents_read.select('.opttbldata #octable tr td'):
#        DataRow_name = DataRow.text
#        removeSlashFromData = DataRow_name.replace('\\n','')
#        removeSpacesFromData = removeSlashFromData.replace(' ','')
#        print(str(removeSpacesFromData))
#        DataRows.append(removeSpacesFromData)
#    print(DataRows)

# Cars = {
#         columns[0]: [5,6,8],
#         }
# df = DataFrame(Cars, columns= columns)
# export_excel = df.to_excel (r'D:\3bdelrhmann\Python Scripts\NsIndia\export_dataframe.xlsx', index = None, header=True) 


# os.remove(Html_file_name)



