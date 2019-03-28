import re
import requests
from bs4 import BeautifulSoup
from xlsxwriter.workbook import Workbook
from PIL import Image
from io import BytesIO
import os
import glob
from time import sleep

# Parker's site only serves pages if request comes from a legit broswer.
# Setting header's User-Agent to a supported browser.
headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:66.0) Gecko/20100101 Firefox/66.0'}

try:
    # Update links in links.txt file
    linksFile = "links.txt"
    
    # Create Workbook
    workbook = Workbook('Parker Sensor List.xlsx')
    format = workbook.add_format({'text_wrap': True})
    
    # Create sheet
    sheet1 = workbook.add_worksheet()
    #sheet1.set_column(0, 4, 100)
    sheet1.write('A1', 'S. No.')
    sheet1.write('B1', 'Product Name') 
    sheet1.write('C1', 'Product Image') 
    sheet1.write('D1', 'Product Description') 
    sheet1.write('E1', 'Product URL')

    # Update row number to begin filling content
    row_number = 2

    # List to hold max length of contents of various columns for final formatting
    columnSize = [0, 0, 0, 0, 0]

    # Read links.txt file into an iterable object
    with open(linksFile, 'r') as f:
        links = f.readlines()
        for link in links:
            
            # Send GET request to url
            response = requests.get(link, headers=headers)
            
            # Use BeautifulSoup to parse response received above
            soup = BeautifulSoup(response.text, 'html.parser')
            
            # Find all the products listed in this page
            products = soup.findAll("div", {"class": "product"})
            
            # Iterate over individual product and parse contents
            for product in products:
                # Find product name
                productName = product.find("img", alt=True)
                # First productName is shoppingListItemAddedImg and is not valid
                # ie has no 'alt'
                if productName['alt']:
                    # Insert serial number to excel sheet
                    size = len(str(row_number))
                    if columnSize[0] < size:
                        columnSize[0] = size
                    sheet1.write('A'+str(row_number), row_number-1)
                    
                    # Product name
                    size = len(productName['alt'])
                    if columnSize[1] < size:
                        columnSize[1] = size
                    sheet1.write('B'+str(row_number), productName['alt'])
                    print productName['alt']
                    
                    # Product image
                    print 'http://'+productName['src'].split("//")[1]
                    imageResponse = requests.get('http://'+productName['src'].split("//")[1], headers=headers)
                    if (imageResponse.status_code == 200):
                        img = Image.open(BytesIO(imageResponse.content))
                        # Best size that fits really well is 137 pixels x 125 pixels
                        image = img.resize((137,125))
                        width, height = image.size
                        if columnSize[2] < width:
                            columnSize[2] = width
                        image.save('productImage'+str(row_number-1)+'.jpg', format='JPEG')
                        
                        sheet1.set_row(row_number-1, height)
                        sheet1.insert_image('C'+str(row_number), 'productImage'+str(row_number-1)+'.jpg')            
                        print productName['src'][2:]
                    else:
                        sheet1.write('C'+str(row_number), 'No image')
            
                    # Product description
                    productDescription = product.find("div", attrs={'class': 'product_description'}).text
                    sheet1.write('D'+str(row_number), productDescription)
                    print productDescription
            
                    # Product page link
                    productPage = product.find("div", attrs={'class': 'product_name'})
                    productPageUrl = productPage.find("a", href=True)
                    sheet1.write('E'+str(row_number), productPageUrl['href'])
                    size = len(str(productPageUrl['href']))
                    if columnSize[4] < size:
                        columnSize[4] = size
                    print productPageUrl['href']
            
                    # Increment row number for next iteration
                    row_number += 1

                    #break                    
            
    # Adujst column sizes for aesthetics
    sheet1.set_column(0, 0, columnSize[0]+5)
    sheet1.set_column(1, 1, columnSize[1]+5)
    # 7 is required to scale pixels to size of 1 character * number of characters it would take to occupy image's width
    sheet1.set_column(2, 2, int(columnSize[2] / 7))
    sheet1.set_column(3, 3, 100, format)
    sheet1.set_column(4, 4, columnSize[4]+5)
    # Save and close workbook
    workbook.close()
    
    # Remove all downloaded images from this folder
    for i in glob.glob("*.jpg"):
        os.remove(i)

except Exception as e:
    print(e)
    workbook.close()
    # Remove all downloaded images from this folder
    for i in glob.glob("*.jpg"):
        os.remove(i)
    pass

