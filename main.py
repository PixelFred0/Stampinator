import tkinter as tk
from tkinter import Toplevel, ttk , messagebox
from tkinter import filedialog
import os
from pdfrw import PdfReader, PdfWriter, PageMerge
from PIL import Image
import configparser
import fitz
from datetime import date
import d2p

#Config
config = configparser.ConfigParser()
user = os.environ['USERPROFILE']
start_path = user+"\Desktop"
path = '/'.join((os.path.abspath(__file__).replace('\\', '/')).split('/')[:-1])#file path
text_beschreibung = "Wähle einen Input-Ordner mit Word-Dateien, einen Output-Ordner für die fertigen PDF-Dateien und den zu verwendenden PDF-Stempel und klicke dann auf START:"
folder_path_doc, folder_path_pdf, file_name_stemp = None, None, None
error_names_list = ["Es wurde kein Input-Ordner markiert!", "Es wurde kein Output-Ordner markiert!", "Es wurde kein PDF-Stempel gewählt!"]
jpg_folder = os.path.join(path, 'jpg')#Temporärer Speicherort für die Bilder
pdf_folder = os.path.join(path, 'pdf')#Temporärer Speicherort für die PDFs # neu
start_text = "START: Je nach Anzahl der Dateien dauert der Vorgang etwas länger...bitte habe ein wenig Geduld und lasse dieses Fenster geöffnet, bis eine Erfolgsmeldung erscheint!"
rechnungen_chosed = False
progress = True

#Variables
licensed_keys = ["707b", "admin_b"]
input_folder_path = output_folder_path = stamp_file_path = doc_folder = finial_pdf_folder = stemp_file = None

def dispatch(app_name:str):
    try:
        from win32com import client
        app = client.gencache.EnsureDispatch(app_name)
    except AttributeError:
        print("raised")
        # Corner case dependencies.
        import os
        import re
        import sys
        import shutil
        # Remove cache and try again.
        MODULE_LIST = [m.__name__ for m in sys.modules.values()]
        for module in MODULE_LIST:
            if re.match(r'win32com\.gen_py\..+', module):
                del sys.modules[module]
        shutil.rmtree(os.path.join(os.environ.get('LOCALAPPDATA'), 'Temp', 'gen_py'))
        from win32com import client
        app = client.gencache.EnsureDispatch(app_name)
    return app

word = dispatch("Word.application")   

def key_check():
    config.read(os.path.join(path, 'config.ini'))
    config_key = config['default']['key']
    print(config_key)
    todays_date = date.today()
    expired_date = date(2023, 3, 30)
    if config_key in licensed_keys:
        print("true")
        return True
    elif todays_date < expired_date:
        print(todays_date)
        print("true")
        return True
    else:
        print("false")
        return False

def secret_folder():#Zweiter Output Ordner für Rechnungen
    global secret_pdf_folder
    config.read(os.path.join(path, 'config.ini'))
    secret_pdf_folder = config['default']['folder']
    return secret_pdf_folder

def check_path_type(Path: str):
    if not ":" in Path:
        Path = None
    return Path

def browse_button_doc():#Button Funktion um den Doc Ordner auszuwählen
    global folder_path_doc, doc_folder
    folder_path_doc = filedialog.askdirectory(initialdir=start_path, title="Input-Ordner markieren")
    doc_folder = check_path_type(folder_path_doc)
    folder_content = folder_content_names = []
    sec_folder_status = True
    for i in os.listdir(folder_path_doc):
        if i.lower().endswith((".doc", ".docm", ".docx")):
            folder_content.append(True)
            i_name = os.path.splitext(os.path.basename(i))
            i_name = i_name[0]
            folder_content_names.append(i_name)
        else:
            folder_content.append(False)
    if False in folder_content:
        messagebox.showinfo("Fehler", "Im Input Ordner sind Ordner oder andere Datei-Typen,\ndie zu Problemen führen! Bitte entferne diese! ")
        folder_content = folder_content_names = []
    else:
        print("True")
    if rechnungen_chosed: #Überprüfung von Dokuemten mit gleichen Namen im Secret Folder
        sec_folder_content_names = []
        for i in os.listdir(secret_folder()):
            if i.lower().endswith((".pdf")):
                i_name = os.path.splitext(os.path.basename(i))
                i_name = i_name[0]
                sec_folder_content_names.append(i_name)
        for i in folder_content_names:
            if i in sec_folder_content_names:
                 messagebox.showinfo("Fehler", f"Im Secret Ordner befindet sich bereits die Rechnung: {i}.\n Um Probleme zu vermeiden, Bitte entferne diese Rechnung aus dem Input Ordner! ")
                 sec_folder_content_names = []
                 doc_folder = None
                 sec_folder_status = False
                 print("doc 1 none")
            elif sec_folder_status:
                doc_folder = folder_path_doc
                print("True-Pass")
    print(start_path)
    

def browse_button_pdf():#Button Funktion um den Pdf Ordner auszuwählen
    global folder_path_pdf, finial_pdf_folder
    folder_path_pdf = filedialog.askdirectory(initialdir=start_path, title="Output-Ordner markieren")
    finial_pdf_folder = check_path_type(folder_path_pdf)
    print(finial_pdf_folder)

def browse_button_stemp():#Button Funktion um die Stempel Datei auszuwählen
    global file_name_stemp, stemp_file
    file_name_stemp = filedialog.askopenfilename(initialdir=start_path, title="Stempel auswählen", filetypes=[("PDF files", "*.pdf")])
    stemp_file = check_path_type(file_name_stemp)
    print(stemp_file)

def doc_to_docx(doc_folder: str):
    doc_counter =  docx_counter = 0
    
    for file_path in os.listdir(doc_folder):
        file_path = os.path.join(doc_folder, file_path)
        if file_path.lower().endswith(".doc"):
            print("Datei ist doc")
            docx_file = '{0}{1}'.format(file_path, 'x')
            file_path = os.path.abspath(file_path)
            docx_file = os.path.abspath(docx_file)
            try:
                wordDoc = word.Documents.Open(file_path)
                wordDoc.SaveAs2(docx_file, FileFormat = 16)
                wordDoc.Close()
                print("docx")
            except Exception as e:
                print('Failed to Convert: {0}'.format(file_path))
                print(e)
            os.remove(file_path)
            doc_counter += 1
        else:
            print("Datei ist docx")
            docx_counter += 1
    print(f"{doc_counter + docx_counter} Datein im Input_Ordner")

def stemping_pdf(pdf_file: str, stemp_file: str, pdf_folder: str):#pdf wird gestempelt
    reader_input = PdfReader(pdf_file)
    writer_output = PdfWriter()
    watermark_input = PdfReader(stemp_file)
    watermark = watermark_input.pages[0]
    for current_page in range(len(reader_input.pages)):
        merger = PageMerge(reader_input.pages[current_page])
        merger.add(watermark).render()
    writer_output.write(pdf_folder, reader_input)
    print("stamped pdf")

def pdf_to_image(pdf_file, jpg_folder):#pdf wird in jpg umgewandelt
    doc = fitz.open(pdf_file)
    for page in doc:  
        pix = page.get_pixmap(dpi=300)  
        pix.save(jpg_folder + "/page%i.jpg" % page.number)  
    print("pdf to image")

def image_to_pdf(jpg_folder: str, finial_pdf_folder: str):#jpg's werden in eine pdf umgewandelt
    image_list = []
    image = Image.open(jpg_folder + "/page0.jpg")
    im_1 = image.convert('RGB')
    if len(os.listdir(jpg_folder)) > 1:
        os.remove(jpg_folder + "/page0.jpg")
        for file in os.listdir(jpg_folder):
            image = Image.open(jpg_folder + "/" + file)
            im = image.convert('RGB')
            image_list.append(im)
            os.remove(jpg_folder+ "/" + file)
    im_1.save(finial_pdf_folder + "/" + pdf_name +".pdf", save_all=True, append_images=image_list, quality=50,resolution=300, optimize=True)# original folder
    if rechnungen_chosed:
        im_1.save(secret_folder() + "/" + pdf_name +".pdf", save_all=True, append_images=image_list, quality=50,resolution=300, optimize=True)# secret folder
    print("image to pdf")

def auto_doc_to_pdf_stamper(doc_folder: str, stemp_file: str, pdf_folder: str, jpg_folder: str, finial_pdf_folder: str):
    global pdf_name, pdf_count, progress
    pdf_count = 0
    for i in os.listdir(pdf_folder):#new
        os.remove(pdf_folder+ "/" + i)
    for i in os.listdir(jpg_folder):
        os.remove(jpg_folder+ "/" + i)
    #doc_to_docx(doc_folder)
    d2p.convert(doc_folder, pdf_folder)
    print("Pdfs saved")
    if os.listdir(pdf_folder) != 0:
        print("file in pdf_folder")
        for file in os.listdir(pdf_folder):#Schleife zum doc zu gestempelten pdf
            pdf_count += 1
            print(pdf_count)
            pdf_name = os.path.splitext(os.path.basename(file))
            pdf_name = pdf_name[0]
            stemping_pdf(pdf_folder + "/" + pdf_name + ".pdf", stemp_file, pdf_folder + "/" + pdf_name + ".pdf")
        for file in os.listdir(pdf_folder):#Schleife um pdf zu jpg und zurück in pdf
            pdf_name = os.path.splitext(os.path.basename(file))
            pdf_name = pdf_name[0]
            pdf_to_image(pdf_folder + "/" + pdf_name + ".pdf", jpg_folder)
            image_to_pdf(jpg_folder, finial_pdf_folder)
            os.remove(pdf_folder+ "/" + file)
        for i in os.listdir(jpg_folder):
            os.remove(jpg_folder+ "/" + i)
        progress = False
    else:
        print("Keine pdfs in pdf_folder")

def start_programm_button():
    global start_text, error_text
    if doc_folder is None:
        print("doc none")
        messagebox.showinfo("Error", error_names_list[0])
    elif finial_pdf_folder is None:
        print("pdf none")
        messagebox.showinfo("Error", error_names_list[1])
    elif stemp_file is None:
        print("stemp none")
        messagebox.showinfo("Error", error_names_list[2])
    else:
        print("start")
        if key_check():
            openNewWindow_working()
            error_text = "none"
            try:
                auto_doc_to_pdf_stamper(doc_folder, stemp_file, pdf_folder, jpg_folder, finial_pdf_folder)
            except IndexError as e:
                error_text = f"---------------Index Fehler---------------\nAußerhalb des Index von einer Liste! Sage bitte deinem Admin Bescheid!\n{e}"
                print(e)
                NewWindow_error()
            except NameError as e:
                error_text = f"---------------Namens Fehler---------------\nVariable oder Datei nicht defeniert! Sage bitte deinem Admin Bescheid!\n{e}"
                NewWindow_error()
                print(e)
            except OSError as e:
                error_text = f"---------------OS Fehler---------------\nFehler mit einer Datei oder Pfadangabe! Sage bitte deinem Admin Bescheid!\n{e}"
                NewWindow_error()
                print(e)
            except RuntimeError or ValueError or FileExistsError as e:
                error_text = f"---------------Runtime Fehler---------------\nEs gab einen unbekannten Fehler! Sage bitte deinem Admin Bescheid!\n{e}"
                NewWindow_error()
                print(e)
            except SyntaxError as e:
                error_text = f"---------------Syntax Fehler---------------\nBeim starten des Prozess gab es ein Fehler! Sage bitte deinem Admin Bescheid!\n{e}"
                NewWindow_error()
                print(e)
            except SystemError as e:
                error_text = f"---------------System Fehler---------------\nEs gab im System einen Fehler! Sage bitte deinem Admin Bescheid!\n{e}"
                NewWindow_error()
                print(e)
            except FileExistsError as e:
                error_text= f"---------------Datei Fehler---------------\nEine Datei gibt es bereits! Sage bitte deinem Admin Bescheid!\n{e}"
            print(error_text)
            if error_text == "none":
                openNewWindow_done()
        else:
            messagebox.showwarning("Abgelaufen", "Der Programm Key ist ausgelaufen")
            
root = tk.Tk()
root.wm_iconbitmap(path+"/pdf.ico")
root.title("Stampinator")#Program Name

def close_button():
    root.destroy()

def clear_doc():
    for i in os.listdir(doc_folder):
        os.remove(doc_folder+ "/" + i)
    messagebox.showinfo("Fertig", "Der Inhalt vom Input Ordner wurde gelöscht!")

def openNewWindow_rechnung():
    global rechnungen_chosed
    global newWindow2
    global text_2
    root.iconify()
    rechnungen_chosed = True
    newWindow2 = Toplevel(root)
    newWindow2.wm_iconbitmap(path+"/pdf.ico")
    newWindow2.title("Rechnungen")
    newWindow2.resizable(False,False)
    doc_button = ttk.Button(newWindow2,text="Input (doc)", width=20, command= browse_button_doc, ).grid(row=2,column=0, padx=10, pady=15)#Button_doc
    pdf_button = ttk.Button(newWindow2,text="Output (pdf)", width=20, command= browse_button_pdf, ).grid(row=2,column=1, padx=10, pady=15)#Button_pdf
    stempel_button = ttk.Button(newWindow2,text="Stempel (pdf)",width=20, command= browse_button_stemp, ).grid(row=2, column=2, padx=10, pady=15)#Button_stempel
    start_button = ttk.Button(newWindow2,text="START",width=20,command=start_programm_button, ).grid(row=3, column=1, pady=15)#Button_start
    text = ttk.Label(newWindow2, text= text_beschreibung).grid(row=0, column=0, columnspan=3, padx=20, pady=15)#Text
    text_2 = ttk.Label(newWindow2, text= start_text).grid(row=4, column=0, columnspan=3,  pady=15)#Text

def openNewWindow_teilnehmer():
    global newWindow2
    global text_2
    root.iconify()
    newWindow2 = Toplevel(root)
    newWindow2.wm_iconbitmap(path+"/pdf.ico")
    newWindow2.title("Teilnahmebestätigungen")
    newWindow2.resizable(False,False)
    doc_button = ttk.Button(newWindow2,text="Input (doc)", width=20, command= browse_button_doc, ).grid(row=2,column=0, padx=10, pady=15)#Button_doc
    pdf_button = ttk.Button(newWindow2,text="Output (pdf)", width=20, command= browse_button_pdf, ).grid(row=2,column=1, padx=10, pady=15)#Button_pdf
    stempel_button = ttk.Button(newWindow2,text="Stempel (pdf)",width=20, command= browse_button_stemp, ).grid(row=2, column=2, padx=10, pady=15)#Button_stempel
    start_button = ttk.Button(newWindow2,text="START",width=20,command=start_programm_button, ).grid(row=3, column=1, pady=15)#Button_start
    text = ttk.Label(newWindow2, text= text_beschreibung).grid(row=0, column=0, columnspan=3, padx=20, pady=15)#Text
    text_2 = ttk.Label(newWindow2, text= start_text).grid(row=4, column=0, columnspan=3,  pady=15)#Text

def openNewWindow_working():
    global newWindow3
    global working_text
    working_text = "In Arbeit . . ."
    print("start working")
    newWindow2.destroy()
    newWindow3 = Toplevel(root)
    newWindow3.wm_iconbitmap(path+"/pdf.ico")
    newWindow3.title("In Arbeit")
    newWindow3.geometry("250x100")
    newWindow3.resizable(False,False)
    text = ttk.Label(newWindow3, text= working_text).pack(pady=30, padx= 10)
    newWindow3.update()

def openNewWindow_done():
    window_text = f'''Es wurden {pdf_count} Dateien erfolgreich erstellt, Du kannst 
denn Input Ordner Inhalt löschen und das Programm schließen!'''
    print("done")
    newWindow3.destroy()
    newWindow = Toplevel(root)
    newWindow.wm_iconbitmap(path+"/pdf.ico")
    newWindow.title("Fertig")
    newWindow.resizable(False,False)
    text = ttk.Label(newWindow, text= window_text).pack(pady= 15, padx= 10)
    button_exit = ttk.Button(newWindow, text= "Schließen", width=20, command= close_button).pack(pady=15)
    button_doc_clear = ttk.Button(newWindow, text= "Input (doc) löschen", width=35, command= clear_doc).pack(pady=15)

def NewWindow_error():
    global newWindow4
    newWindow3.destroy()
    newWindow4 = Toplevel(root)
    newWindow4.wm_iconbitmap(path+"/pdf.ico")
    newWindow4.title("Fehler")
    newWindow4.geometry("300x200")
    newWindow4.resizable(False,False)
    text = ttk.Label(newWindow4, text= error_text).pack(pady= 15, padx= 20)
    button_exit = ttk.Button(newWindow4, text= "Schließen", width=20, command= close_button).pack(pady=15, padx= 20)

text = ttk.Label(root, text= "Wähle eine Funktion:").grid(row=0, column=0, columnspan=3, padx=20, pady=10)#Text
rechnungen_button = ttk.Button(root,text="Rechnung", width=25, command= openNewWindow_rechnung, ).grid(row=2,column=0, padx=10, pady=10)
teilnehmer_button = ttk.Button(root,text="Teilnahmebestätigung", width=25, command= openNewWindow_teilnehmer, ).grid(row=2,column=1, padx=10, pady=10)
text_2 = ttk.Label(root, text= "made by Frederik Schmidt").grid(row=3, column=0, columnspan=3, padx=20, pady=10)#Text

if __name__ == '__main__':
    root.mainloop()#Startet die GUI