# Stampinator
This was my first "bigger" Project, made with Python. The Idea behind the Stampinator was to make my life easier and not waste hours preparing pdf invoices
with the stationery stamp. The Process behind the workflow was: 
1. To open the invoice.docm file and export it as a pdf
2. Open the invoic.pdf file and apply the stationery stamp
3. Export the invoice.pdf file Audit-proof / as a picture pdf

This took by hand up to 4 minutes per file, so i made this programm with a gui for me and my coworkers.
When you open the Stampinator application you only have 4 Steps to do:
1. Chose if you want to convert and stamp invoices or Participation confirmation
2. Chose the Input Folder with one or multiple .doc/.docx/.docm files
3. Chose the Output Folder, where the .pdf files will be saved
4. Chose the stationery stamp pdf file
5. Start

The difference between invoice and Participation confirmation mode is, that with the invoice mode the final .pdf files will be stored at a second location
that is defined in the config.ini file.