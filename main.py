from tkinter import PhotoImage, Label, Tk, Button
from tkinter.filedialog import askopenfilenames, asksaveasfilename
import tkinter as tk
import time
import cv2
import pytesseract
from docx import Document
#from doctr.io import DocumentFile
#from doctr.models import ocr_predictor

def work():

    # Allow the user to select files
    filez = askopenfilenames(parent=root, title='Choose a file')

    start_time = time.time()
    curr_time = time.strftime("%H:%M:%S", time.localtime())
    print("Start Time:", curr_time)
    document = Document()


 # text recognition

    # read image
    

    for fileLoc in filez:
        im = cv2.imread(fileLoc)
        # configurations
        config = ('-l eng --oem 1 --psm 3')
        #pytessercat
        text = pytesseract.image_to_string(im, config=config)
        # print text
        text = text.split('\n')
        # Load image file as DocumentFile
        '''doc = DocumentFile.from_images(ii)
        
        # Perform OCR
        result = model(doc)
        jon = result.export()

        # Function to extract text from OCR result
        def textData():
            string = ""
            for i in jon["pages"]:
                for j in i["blocks"]:
                    for k in j["lines"]:
                        for l in k["words"]:
                            string += " " + l["value"]
            return string
        
        # Get extracted text
        ss = textData()'''

        # Add extracted text to the Word document
        para = document.add_paragraph(text)
        para.alignment = 1
        document.add_page_break()
        print("Text extracted from", fileLoc)

    # Allow the user to choose the folder and filename to save the document
    file_path = asksaveasfilename(defaultextension=".docx", filetypes=[("Word Document", "*.docx")], parent=root)

    # Check if the user canceled the save dialog
    if file_path:
        # Save the document
        document.save(file_path)
        print("Document saved as", file_path)
    else:
        print("Save operation canceled.")

    decoded_label = tk.Label(root, text="Completed", bg="#C0B974",font=("Helvetica", 25))
    decoded_label.place(x=626, y=498)

     # Calculate and print elapsed time

    seconds = time.time() - start_time
    print('Time Taken:', time.strftime("%H:%M:%S", time.gmtime(seconds)))

if __name__ == "__main__":

    
    # Initialize Tkinter root window
    root = Tk()
    root.title("Image to Docx")
    root.geometry("1360x710")
    root.config(bg="#F0F0F0")

    # Set background image
    bg_image = PhotoImage(file="./template/bgimage.png")
    background_label = Label(root, image=bg_image)
    background_label.place(x=0, y=0, relwidth=1, relheight=1)

    # Create button for triggering OCR
    encodebtn = PhotoImage(file="./template/btn.png")
    button_encode = Button(image=encodebtn, command=work, bd=0, bg="#C0B974",highlightthickness=0, activebackground="#C0B974")
    button_encode.place(x=531, y=355)

    # Start Tkinter event loop
    root.mainloop()
