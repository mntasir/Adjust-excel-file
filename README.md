It's simple code to enter and adjust excel files, i get it from chatgpt when i asked him to make code to write on excel files.
first openpyxl.load_workbook() method used to open the excel file.then the result linked with a variable "workbook = openpyxl.load_workbook(r'C:\Users\yourname\Desktop\FileName.xlsx')".
Then worksheet determened by using the variable we declared above and choosing the sheet as a key value "worksheet = workbook['Sheet1']".
"worksheet" will be used to determine which cell will be updated and value that will be added "worksheet['A3'] = 'Hello, World!'".
Here we will use "workbook" variable again to approach save method to save chenges "workbook.save(r'C:\Users\yourname\Desktop\example.xlsx')".
