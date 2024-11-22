from make_sure_path_exists import make_sure_path_exists

import datetime as dt
import re
import pandas as pd
from xhtml2pdf import pisa

# Utility function
def convert_html_to_pdf(source_html, output_filename):

    # open output file for writing (truncated binary)
    result_file = open(output_filename, "w+b")

    # convert HTML to PDF
    pisa_status = pisa.CreatePDF(
            source_html,                # the HTML to convert
            dest=result_file)           # file handle to recieve result

    # close output file
    result_file.close()                 # close output file

    # return False on success and True on errors
    return pisa_status.err



# HTML template for PDF output
template = '''\
<!DOCTYPE html>
<html>
  <head>
    <title>{0}</title>
    <link href="assets/csv_to_pdf.css" rel="stylesheet">
  </head>
  <body>
    <div class="header_logo">
      <img src="assets/MIFL_LOGO_1PX_LHS.png">
    </div>
    <div class="main_div">
      <table>
        <tr>
          <td>{0} - Cash <b>{1}</b> Instruction</td>
        </tr>
      </table>
      <table class="tableA">
        <tr>
          <td class="onethird">{1} Amount ({2}):</td>
          <td class="twothird">{3}</td>
        </tr>
        <tr>
          <td class="onethird">Trade date:</td>
          <td class="twothird">{4}</td>
        </tr>
        <tr>
          <td class="onethird">Value date:</td>
          <td class="twothird">{5}</td>
        </tr>
        <tr>
          <td class="onethird"></td>
          <td class="twothird"></td>
        </tr>
      </table>
      <table class="tableB">
        <tr>
          <td>Reference:</td>
          <td class="twothird">Manager Transfer</td>
        </tr>
        <tr>
          <td>Debit RBC Custody A/C:</td>
          <td class="twothird">{6} - {0}</td>
        </tr>
        <tr>
          <td>Credit RBC Custody A/C:</td>
          <td class="twothird">{7} - {0}</td>
        </tr>
      </table>
      <table class="tableC">
        <tr>
          <td>_________________</td>
          <td>_________________</td>
        </tr>
        <tr>
          <td>{8}</td>
          <td>{10}</td>
        </tr>
        <tr>
          <td>Authorised signatory ({9})</td>
          <td>Authorised signatory ({11})</td>
        </tr>
      </table>
      <table class="tableD">
        <tr>
          <td class="onethird">
            <span>
Mediolanum International Funds Ltd
4th Floor, The Exchange
Georges Dock
IFSC, Dublin 1
D01 P2V6
Ireland

Tel: +353 1 2310800
Fax: +353 1 2310805</span>
          </td>
          <td class="twothird">
            <span>
Registered in Dublin No: 264023
Directors:  K Zachary, C Bocca (Italian), M Nolan, F Frick,
F Pietribiasi (Managing) (Italian), M Hodson,
C Jaubert (French), E Fontana Rava (Italian), C Bryans.

Mediolanum International Funds Limited is regulated by the Central Bank of Ireland</span>
          </td>
        </tr>
      </table>
    </div>
  </body>
</html>
'''

# load the source CSV file
try:
   df = pd.read_csv('source.csv', parse_dates=['trade_date', 'value_date'], thousands=',', dayfirst=True)
except FileNotFoundError:
   print("No source.csv file")
   quit()

print(df.head())

make_sure_path_exists('output')

for row in df.itertuples(index=True):
    
    try:
      # prepare output file name
      # strip invalid characters
      # add row number at the end to make the file name unique
      n = re.sub(r'[\W_]', '', row.fund_name)
      fname = 'output/{}_{}.pdf'.format(n, row[0])
      
      # fill in template
      s = template.format(row.fund_name, \
                          row.instruction_type, \
                          row.currency_code,
                          '{0:,.0f}'.format(row.amount), \
                          dt.date.strftime(row.trade_date, "%A, %B %d, %Y"), \
                          dt.date.strftime(row.value_date, "%A, %B %d, %Y"), \
                          row.debit_account, row.credit_account, \
                          row.signatory_one, row.list_one,
                          row.signatory_two, row.list_two )

      # save HMTL as PDF
      convert_html_to_pdf(s,  fname)

      # save HTML for debugging
      #with open('{}.html'.format(fname), 'w', encoding='utf-8') as f:
      #    f.write(s)

      print(fname)

    
    except ValueError as e:
      print('line {} - possible incorrect amount {}'.format(row[0], type(e)))
      print(row.amount)
    except TypeError as e:
      print('line {} - possible incorrect date {}'.format(row[0], type(e)))

    #except:
    #  print('line {} - possible empty line'.format(row[0]))

print('done...')
