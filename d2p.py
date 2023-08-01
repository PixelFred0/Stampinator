def convert(input_folder, output_folder):
    import os
    import win32com.client
    pdf_format_key = 17
    folder_in = os.path.abspath(input_folder)
    folder_out = os.path.abspath(output_folder)
    if os.listdir(input_folder) != 0:
        word = win32com.client.Dispatch('Word.Application')
        for filename in os.listdir(input_folder):
            #for f in filenames:
            print(filename)  
            if filename.lower().endswith(".docx") :
                file_out = filename.replace(".docx", ".pdf")
                doc = word.Documents.Open(folder_in + "/" + filename)
                doc.SaveAs(folder_out + "/" + file_out, FileFormat=pdf_format_key)
                doc.Close(0)
            if filename.lower().endswith(".doc"):
                file_out = filename.replace(".doc", ".pdf")
                doc = word.Documents.Open(folder_in + "/" + filename)
                doc.SaveAs(folder_out + "/" + file_out, FileFormat=pdf_format_key)
                doc.Close(0)
            if filename.lower().endswith(".docm"):
                file_out = filename.replace(".docm", ".pdf")
                doc = word.Documents.Open(folder_in + "/" + filename)
                doc.SaveAs(folder_out + "/" + file_out, FileFormat=pdf_format_key)
                doc.Close(0)
            print(file_out)
        word.Quit()
    else:
        print("Keine Datein im Inputordner")
                