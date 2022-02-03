# A python program to automate th generation of zipped folder of images
# generated from PDFs created using scrpped data from an Excel sheet.

# Importing neccessary libraries
from reportlab.pdfgen import canvas         # Library and method for generating a PDF.
from reportlab.lib.units import inch,cm     # Library and methods for unit conversions.
import pandas as pd                         # Library for data manipulation.                        
from pdf2image import convert_from_path     # Library and method for PDF to image conversion.
from PIL import Image                       # Library and method for image modification.
from zipfile import ZipFile                 # Library and method to generate a zip file.



# Data scrapping
#----------------
# Specifying the path and sheet name of the required Excel sheet.
print('Receipts Generator')
print('Version: 1.0')

path=input('Specify the complete path to the file:')
sheetname = 'clean'

# Forming a pandas dataframe and additional required data structures from the data.
conveyance_df = pd.read_excel ( path,
                                sheet_name=sheetname,
                                header=None,
                                names=[ 'Date',
                                        'Source',
                                        'Destination',
                                        'Amount',
                                        'Vehicle' ])

date_list = list(conveyance_df['Date'].dt.strftime("%d %b %Y"))
billno_list = list(conveyance_df['Date'].dt.strftime("#%d%m%y"))
source_list = list(conveyance_df['Source'])
destination_list = conveyance_df['Destination']
amount_list = conveyance_df['Amount']
records = len(conveyance_df)

# PDF generation.
#-----------------
#Defining the function for the creation of PDFs with the required data.
def gentext(canvas_pdf):
    canvas_pdf.translate(0,29.7*cm)     # Setting origin to top left.


    # Setting values for the header of the PDF.
    canvas_pdf.setFont('Courier-Bold',11)       
    canvas_pdf.drawString(8.75*cm,-1.25*cm,'Kukreja Taxi Stand')
    canvas_pdf.drawString(7.5*cm,-1.75*cm,'Madangir, near Pushpa Bhawan')
    canvas_pdf.drawString(8.75*cm,-2.25*cm,'New Delhi - 110062')
    canvas_pdf.drawString(9*cm,-2.75*cm,'+91 - 7011301693')

    # Setting values for the middle and data element.
    canvas_pdf.setFont('Courier',11)
    canvas_pdf.drawString(5.75*cm,-3.75*cm,f'Date: {date_list[trip]}')
    canvas_pdf.drawString(5.75*cm,-4.25*cm,f'Bill No.: {billno_list[trip]}')
    canvas_pdf.drawString(13.75*cm,-4.75*cm,f'Rs {float(amount_list[trip])} ')
    canvas_pdf.drawString(13.75*cm,-5.25*cm,f'Trans: ')
    canvas_pdf.drawString(13.75*cm,-5.75*cm,f'Auth: ')
    canvas_pdf.drawString(5.75*cm,-7.25*cm,f'Tax: 0.0 ')
    canvas_pdf.drawString(5.75*cm,-7.75*cm,f'Total: {float(amount_list[trip])}')
    canvas_pdf.drawString(5.75*cm,-8.25*cm,f'From: {source_list[trip]}')
    canvas_pdf.drawString(5.75*cm,-8.75*cm,f'To: {destination_list[trip]}')
    canvas_pdf.drawString(5.75*cm,-9.25*cm,'Mode of Payment: Cash')

    # Setting values for 'Thank you' note.
    canvas_pdf.setFont('Courier-Bold',11)
    canvas_pdf.drawString(9.5*cm,-11.75*cm,'Thank you!')
    canvas_pdf.drawString(9.15*cm,-12.25*cm,'Customer copy')
    # End of the function definition.


# Creating a loop to generate required number of PDFs.
for trip in range(records):
    c = canvas.Canvas(f'{date_list[trip]}_{trip}.pdf')      # Defining the file name of the PDF.
    gentext(c)      # Calling the above created 'gentext' function.
    c.showPage()    # Ending the creation of PDF.
    c.save()        # Saving the PDF.

print('Data to PDFs conversion done.')


#PDF to image creation and cropping.
#------------------------------------
# Defining the co-ordinates of the image to be cropped.
left = 400 
top = 50
right = 1200
bottom = 1100

for trip in range(records):
    # Converting the PDFs to images.
    converted_image = convert_from_path(f'D:\All Thing Files\Programs\Python Programs\{date_list[trip]}_{trip}.pdf',
                           output_folder='D:\All Thing Files\Programs\Python Programs\pdf2image',
                           fmt='png',
                           single_file=True,
                           output_file=f'{date_list[trip]}_{trip}',
                         )

    # Cropping the images.
    load_image = Image.open(r'D:\All Thing Files\Programs\Python Programs\pdf2image\{}_{}.png'.format(date_list[trip],trip))       # Accessing the newly formed image.
    cropped_image = load_image.crop((left,top,right,bottom))    # Generating a cropped image by giving the co-ordinates for the desired image crop.
    # cropped_image.show()
    cropped_image.save(f'D:\All Thing Files\Programs\Python Programs\pdf2image\{date_list[trip]}_{trip}.png')       # Saving the cropped image.

print('Images generated successfully.')

# Zip folder creation.
#----------------------
# Defining the path of the zip folder.
zip_path='D:\All Thing Files\Programs\Python Programs\Conveyance_Dec_21.zip'
zip_object = ZipFile(zip_path,'w')      # Creating an object and specifying the write mode.
for trip in range(records):
    # Zipping all the images together.
    zip_object.write(f'D:\All Thing Files\Programs\Python Programs\pdf2image\{date_list[trip]}_{trip}.png',arcname = f'{date_list[trip]}_{trip}.png')
# Closing the zip object.
zip_object.close()

print('Zip folder created successfully.')
# End
#-----