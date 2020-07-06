from openpyxl import *
from os import remove
import time
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import landscape
from reportlab.lib.units import inch

def getCustomers(customer, envelope):
    customerFilename = customer
    envelopeFilename = envelope

    customerWorkbook = load_workbook(customerFilename)
    customerSheet = customerWorkbook.worksheets[0]

    envelopeWorkbook = Workbook()
    envelopeSheet = envelopeWorkbook.worksheets[0]

    print('Parsing Customers\n')
    time.sleep(1)
    for i in range(customerSheet.max_row):
        sheetEmail = customerSheet.cell(row=i + 1, column=13)
        sheetBalance = customerSheet.cell(row=i + 1, column=12)

        if (sheetEmail.value is None and float(sheetBalance.value) > 0.0):
            try:
                row = []
                row.append(customerSheet.cell(i + 1, column=2).value)
                row.append(customerSheet.cell(row=i + 1, column=3).value)
                row.append(
                    customerSheet.cell(row=i + 1, column=7).value + ', ' + customerSheet.cell(row=i + 1, column=8).value +
                    '  ' + customerSheet.cell(row=i + 1, column=9).value)
                row.append('ATTN: ACCOUNTS PAYABLE')

                envelopeSheet.append(row)
                envelopeWorkbook.save(envelopeFilename)
            except TypeError:
                print('There is an empty column(7, 8, or 9) at row ' + str(i + 1) + ' for customer: ' + row[0])
                print('This line has been skipped and the final envelope file does not include this customer\n')
                time.sleep(3)

    envelopeWorkbook.save(envelopeFilename)


def createEnvelopes(envelope, pdfFile):
    print('Creating Envelope File\n')
    time.sleep(1)

    #Create pdf page and adjust to correct size
    c = canvas.Canvas(pdfFile)
    c.setTitle('Statement Envelopes')
    c.setPageSize(landscape([9.5*inch, 4*inch]))

    #Create Envelopes
    workbook = load_workbook(envelope)
    sheet = workbook.worksheets[0]

    for row in sheet.iter_rows():

        #Sending Address
        c.drawString(275, 150, row[0].value)
        c.drawString(275, 138, row[1].value)
        c.drawString(275, 126, row[2].value)
        c.drawString(275, 114, row[3].value)

        #Return Address
        c.drawString(15, 272, 'Newton Crouch Inc.')
        c.drawString(15, 260, 'P.O. Box 17')
        c.drawString(15, 248, 'Griffin, GA  30224')

        #Go to next page
        c.showPage()

    c.save()


def main():
    customerFilename = 'customers.xlsx'
    envelopeFilename = 'envelope_customers.xlsx'
    pdfEnvelope = 'Envelopes.pdf'

    getCustomers(customerFilename, envelopeFilename)
    createEnvelopes(envelopeFilename, pdfEnvelope)

    print('The program has run succesfully! Goodbye')
    time.sleep(5)

    #Remove created middle files
    # remove(envelopeFilename)


if __name__ == "__main__": main()
