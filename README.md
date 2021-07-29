# Download and analysis of invoices

## Table of contents
* [General info](#general-info)
* [How it works](#how-it-works)
* [Technologies](#technologies)
* [Setup](#setup)

## General info
The program checks an email box by accessing and downloading the attached invoices sent by a specified sender. Accessed emails are marked as viewed and downloaded attachments are put in a folder in the same location as the script.
The downloaded files are analyzed and an XLSX file is generated with the invoice data.

## How it works
The script uses the IMAP protocol to access the corresponding email and checks the emails from the corresponding sender.
With the os and datetime modules, a folder is created to download the PDFs attached to the emails.
After downloading the files, the reviewed emails are marked as viewed.
If there are no new emails, the created folder is deleted and the program warns that there is nothing to do.
If invoices have been downloaded, they are analyzed in search of the predetermined data, they were established using regular expressions through the re module.
The search of the data is possible through the use of the PyPDF2 module that takes the text of the PDF and converts it into plain text capable of being read by the script.
Using the openpyxl module, the collected data is inserted into an XLSX file, it is saved in the created folder.
	
## Technologies
Project is created with:
* Python 3.8
* PyPDF2 1.26.0
* IMAP (via Imbox 0.9.8)
* openpyxl 3.0.7
	
## Setup
To run this project, install locally:
* pip install PyPDF2
* pip install imbox
* pip install openpyxl
