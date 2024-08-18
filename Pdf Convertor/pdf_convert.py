import customtkinter as ctk 
from tkinter import filedialog,messagebox
from pdf2docx import Converter
import os
from PyPDF2 import PdfReader


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        ctk.set_appearance_mode("dark") 
        ctk.set_default_color_theme("blue")
        self.title("Pdf Convertor")
        self.minsize(700,500)
        #self.iconbitmap(r"youricon.ico")
        self.state("withdrawn")
        
        # widgets 
        self.widgets()
        self.mainloop()

    def widgets(self):
        # create a frame
        self.main_frame = ctk.CTkFrame(self,corner_radius=10)
        self.main_frame.place(relx=0,rely=0,relwidth=1,relheight=1)
        # headline label
        self.label_text = ctk.CTkLabel(self.main_frame,
                                      text="Convert PDF to Word",
                                      font=("Arial",20),)
        self.label_text.place(relx=0.3,rely=0.1,relwidth= 0.3,relheight=0.09)
        # path label
        self.label_path = ctk.CTkLabel(self.main_frame, corner_radius=10,
                                      text="File Path",
                                      font=("Arial",15),)
        self.label_path.place(relx=0.01,rely=0.25,relwidth= 0.2,relheight=0.09)        
        self.stringvar = ctk.StringVar() # use stringvar to change entry text
        self.path_entry = ctk.CTkEntry(self.main_frame, textvariable=self.stringvar,
                                      border_width=2,
                                      corner_radius=10)
        self.path_entry.place(relx=0.2, rely = 0.25,relwidth=0.5,relheight= 0.09)
        # to open file dialog add a button
        self.open_button = ctk.CTkButton(self.main_frame,text="Open\nFile",
                                         corner_radius=10,
                                         command=self.select_pdf)
        self.open_button.place(relx=0.75,rely=0.25,relwidth = 0.1,relheight=0.09)

        # output label
        self.output_path_lbl = ctk.CTkLabel(self.main_frame, corner_radius=10,
                                      text="Output",
                                      font=("Arial",15),)
        self.output_path_lbl.place(relx=0.01,rely=0.4,relwidth= 0.2,relheight=0.09)        
        self.stringvar2 = ctk.StringVar() # use stringvar to change entry text
        self.output_entry = ctk.CTkEntry(self.main_frame, textvariable=self.stringvar2,
                                      border_width=2,
                                      corner_radius=10)
        self.output_entry.place(relx=0.2, rely = 0.4,relwidth=0.5,relheight= 0.09)
         # choose a path to save output
        self.open_button2 = ctk.CTkButton(self.main_frame,text="Choose\npath",
                                         corner_radius=10,
                                         command=self.select_output)
        self.open_button2.place(relx=0.75,rely=0.4,relwidth = 0.1,relheight=0.09)
        # convert button 
        self.convert_button = ctk.CTkButton(self.main_frame,text="Convert",
                                         corner_radius=10,
                                         command=self.convert)
        self.convert_button.place(relx=0.2, rely = 0.55,relwidth=0.5,relheight= 0.09)
        self.lbl_name = ctk.CTkLabel(self.main_frame,text="Designed by Özkan Yıldız 17.08.2024",font=("Ariel",12))
        self.lbl_name.place(relx=0.25, rely = 0.93,relwidth=0.5,relheight= 0.09)
   
    # to open file dialog function    
    def select_pdf(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if self.file_path:
            self.stringvar.set(self.file_path)
   
    # select output
    def select_output(self):
        self.output_path = filedialog.askdirectory()
        if self.output_path:
            self.stringvar2.set(self.output_path)
   
    # to convert pdf to word function
    def convert(self):
        try:
            pdf_file = f'{self.stringvar.get()}' # file path
            pdf_file_name = os.path.basename(pdf_file)
            docx_file = f'{self.output_entry.get()}/{pdf_file_name.replace(".pdf","")}.docx'
            # open pdf file and convert it to docx
            show_messagebox = messagebox.askyesno("Convert",f"Would you like to convert '{pdf_file_name}', tottally '{self.file_page()}' pages")
            if show_messagebox:
                cv = Converter(pdf_file)
                cv.convert(docx_file)
                cv.close()
                open = messagebox.askyesnocancel("Converted",f"file has converted, and saved to '{docx_file}' ,would you like to open it?")
                if open:
                    os.startfile(docx_file)

        except Exception as a:
            messagebox.showerror("Error",a)

    def file_page(self):
        # open pdf
        pdf_file = PdfReader(f'{self.stringvar.get()}')
        # take pages
        pages = len(pdf_file.pages)
        return pages

        
if __name__=="__main__":
    app = App()

