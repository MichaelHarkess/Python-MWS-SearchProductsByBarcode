#test

import requests, csv, re, sys
from urllib.parse import urlparse
from boto.mws.connection import MWSConnection
from unidecode import unidecode
import xlrd
import tkinter as tk
from tkinter import filedialog

def read_filenames_from_csv(file, col):
    '''
    
    :param file: The CSV file containing the barcodes to search
    :param col: The column number of the barcodes in the CSV file, 0 = A, 1 = B etc.  
    :return: List of barcodes to be searched
    
    '''
    
    search_items = []
    start = 1
    workbook = xlrd.open_workbook(file, on_demand=True)
    worksheet = workbook.sheet_by_index(0)
    numRows = worksheet.nrows

    while start < numRows:
        search_items.append(int(worksheet.cell(start, col).value))
        start += 1

    return search_items

def amazonbarcodesearch(list, mkt, idtype):
    
    '''
    
    :param list: list of barcodes to search inside Amazon
    :param mkt: Marketplace ID
    :param idtype: Barcode Type, EG UPC, EAN
    :return: List of results

    Amazon Marketplace	MarketplaceId
    DE	                A1PA6795UKMFR9
    ES	                A1RKKUPIHCS9HS
    FR	                A13V1IB3VIYZZH
    IT	                APJ6JRA9NG5V4
    UK	                A1F83G8C2ARO7P
    
    '''
    
    counter = 1
    results = []
    # Provide credentials
    conn = MWSConnection(
        aws_access_key_id="", #add your access_key_id
        aws_secret_access_key="", #add your secret_access_key
        Merchant="", #add your merchant id
        host = "" #add your host
    )

    while len(list) > 0:
        # Get a list of orders.
        response = conn.get_matching_product_for_id(
            IdType=idtype,
            # Can only search 5 barcodes at a time
            IdList=list[0:5],
            MarketplaceId=mkt
        )

        product = response.GetMatchingProductForIdResult
        print(product)

        for value in product:
            status = value['status']
            barcode = value['Id']

            if status != 'Success':
                results.append("There is no matching Amazon product for item with {}: {}".format(idtype, barcode))
            else:
                itemtitle = unidecode(value.Products.Product[0].AttributeSets.ItemAttributes[0].Title)
                asin = value.Products.Product[0].Identifiers.MarketplaceASIN.ASIN
                results.append("Found match for Linnworks Item: {}.  {}: {} with matching ASIN: {}".format(itemtitle, idtype, barcode, asin))

        list = list[5:]
        counter += 1

    return results

def output_csv(results, file_name):
    '''

    :param results: List of results from amazonbarcodesearch()
    :param file_name: Name of the CSV file to be created
    :return: None
    '''

    with open(file_name + '.csv', 'w') as myfile:
        wr = csv.writer(myfile, quoting=csv.QUOTE_ALL)
        for result in results:
            wr.writerow([result])
    print("Success")
                
                
                             

root = tk.Tk()
root.withdraw()
file_path = filedialog.askopenfilename()
bcodes = read_filenames_from_csv(file_path)
results = amazonbarcodesearch(bcodes, '', '') #add marketplace ID and barcode type here
output_csv(results, file_path)


        
    

