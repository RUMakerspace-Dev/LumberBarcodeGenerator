'''
Developer Notes:
Last build done on 3/5/25 By Maitreya Yogeshwar
This should be all the code necessary for the barcode generator. Certain Python libraries need older versions to place nice with the rest of the code.
Edit at your own risk.
All related files will be posted to GitHub at: RUMakerspace-Dev
The excel file used as a config for this software should always be included when deployed to a workstation. The format of the sheet labeled "ConfigSheet"
is important, and only the information within the fields should be modified. 
The printer and label formatting functions were tough to implement, I highly reccommend reading through the documentation for PIL and brother_ql libraries
and testing with the printer before making any changes!
'''
import tkinter as tk
from fractions import Fraction
import pandas as pd
import barcode as bc
from barcode.writer import ImageWriter
from datetime import date
from tkinter import messagebox, ttk
from PIL import Image, ImageDraw, ImageFont
from brother_ql.conversion import convert
from brother_ql.backends.helpers import send
from brother_ql.raster import BrotherQLRaster
'''#Test For fixing exe data packaging
Include libusb0.dll from system32 to include files
pip install usb, pyusb, libusb to environment
Maybe try python 3.9
USE Pyinstaller to compile to exe: 
   >pyinstaller  --onedir --noconsole --add-data "Hardwood UPC Generator.xlsx;." "Lumber_Barcode_Generator.py" --distpath .

'''
#Function for Label Printer connection
def SendToPrinter(barcodeImage, title, pricing, bdft, Thickness_fraction):
    
    im = Image.open(barcodeImage)
    im=im.resize((696, 271), resample=Image.NEAREST) 
    # Initialize the drawing context
    draw = ImageDraw.Draw(im)
   
        # Define the text to be written
    text = f"{bdft:.1f} "+title+"        ("+Thickness_fraction+" inch)"#+" $"+f"{pricing:.2f}" f"{Thickness_fraction:.2f}"
    
    # Specify the font (you can change the font and size as needed)
    #font = ImageFont.truetype("arial.ttf", 36)
    font = ImageFont.load_default(size=48)

    # Determine the position to place the text (x, y)
    # Here, we place it at the bottom left corner: (20, im.height-40)
    position = (15, im.height-70)

    # Define the color of the text (in RGB format)
    text_color = (0, 0, 0)  # Black color

    # Draw the text on the image
    draw.text(position, text, fill=text_color, font=font)
    
    
    # Save or display the modified image
    #im.show()
    im.save(barcodeImage)

    backend = 'pyusb'    # 'pyusb', 'linux_kernal', 'network'
    model = 'QL-800' # your printer model.
    printer = 'usb://0x04f9:0x209b'    # Get these values from the Windows usb driver filter.  Linux/Raspberry Pi uses '/dev/usb/lp0'.


    qlr = BrotherQLRaster(model)
    qlr.exception_on_warning = True

    instructions = convert(

            qlr=qlr, 
            images=[im],    #  Takes a list of file names or PIL objects.
            label='62x29', #should be changed depending on roll Type
            rotate='0',    # 'Auto', '0', '90', '270'
            threshold=70.0,    # Black and white threshold in percent.
            dither=False, 
            compress=False, 
            red=False,    # Only True if using Red/Black 62 mm label tape.
            dpi_600=False, 
            hq=True,    # False for low quality.
            cut=True

    )

    send(instructions=instructions, printer_identifier=printer, backend_identifier=backend, blocking=True)

class Lumber(object):
    #Init function to classify info from option selection
    def __init__(self,species,sku,price):
        self.species= species
        self.sku= sku
        self.price= price
    #creates UPC without check digit for given bdft measurement
    def format_UPC(self,board_ft):
        tenthboardft=int(board_ft*10)
        boardft_str="00000"+str(tenthboardft)
        #print(boardft_str)
        self.UPC_A="2"+str(self.sku)+boardft_str[-5:]
        return self.UPC_A

#gets selection from GUI Dropdown menu and initializes lumber class based on info from excel
def select_dropdown_item(event):
    selected_item = option_var.get() #Selected dropdown item
    if selected_item:
        # Get the row corresponding to the selected item
        selected_row = UPC_df[UPC_df['Product Name'] == selected_item].iloc[0]
        
        # Create an object using the selected row data
        global Product_selection
        Product_selection = Lumber(selected_row['Product Name'], selected_row['SKU'], selected_row['Pricing/BDFT'])
        print(Product_selection.species)
        print(Product_selection.sku)
        print(Product_selection.price)
def select_Thickness(event):
    selected_Thickness = Thick_var.get()#Selects thickness from dropdown
    if selected_Thickness:
        global Thickness
        Thickness = Fraction(selected_Thickness)
        print(float(Thickness))

#when generate output button pressed, selection and dimensions are taken from input and processed to return barcode
def generate_output():
    selected_option=Product_selection
    Width = float(Width_entry.get())
    Length= float(Length_entry.get())
    #Thickness=float(Thickness_entry.get()) deprecated for next revision

     

    # Perform any necessary validation or processing here
    if Width and Length and Thickness and selected_option:
        Board_ft=((Width)*(Thickness)*(Length))/144
        selected_option.format_UPC(Board_ft)
        cost=selected_option.price*Board_ft
        print("Your product has a volume of %.2f Board Ft." % Board_ft)
        barcode_output=bc.UPCA(selected_option.UPC_A)
        #Using barcode library to create PNG
        StickerFileName="Lumber Barcode"+str(date.today())+".png"
        with open(StickerFileName, "wb") as f:
            bc.UPCA(selected_option.UPC_A, writer=ImageWriter()).write(f)
        #once barcode image is created, send file to printer with product info to be printed out
        SendToPrinter(StickerFileName, selected_option.species, cost, Board_ft,Thick_var.get())
        output_string = f"Selected Option: {selected_option.species}\nWidth: {Width}\nLength: {Length}\nThickness: {Thickness}\nYour UPC is: {barcode_output}"
        #messagebox.showinfo("Output", output_string)
    else:
        messagebox.showerror("Error", "Dimension field is empty.")
#Opens Excel file and initializes dropdown list options
UPC_df=pd.read_excel("Hardwood UPC Generator.xlsx", sheet_name= "ConfigSheet")#change to _internal/Hardwood UPC Generator.xlsx
Wood_options=UPC_df['Product Name'].tolist()


# Create the main application window
root = tk.Tk()
root.title("Lumber Barcode Generator")

# Option selection
option_var = tk.StringVar()
option_label = tk.Label(root, text="Select Option:")
option_label.grid(row=0, column=0, padx=10, pady=5)
dropdown_menu = ttk.Combobox(root, textvariable=option_var, values=Wood_options)
dropdown_menu.grid(row=0, column=1, padx=10, pady=5)
dropdown_menu.bind("<<ComboboxSelected>>", select_dropdown_item)

# Measurement input
#Width
Width_label = tk.Label(root, text="Enter Width in inches:")
Width_label.grid(row=2, column=0, padx=10, pady=5)
Width_entry = tk.Entry(root)
Width_entry.grid(row=2, column=1, padx=10, pady=5)
#Length
Length_label = tk.Label(root, text="Enter Length in inches:")
Length_label.grid(row=3, column=0, padx=10, pady=5)
Length_entry = tk.Entry(root)
Length_entry.grid(row=3, column=1, padx=10, pady=5)
#Thickness
# Thickness_label = tk.Label(root, text="Enter Thickness in inches:")
# Thickness_label.grid(row=3, column=0, padx=10, pady=5)
# Thickness_entry = tk.Entry(root)
# Thickness_entry.grid(row=3, column=1, padx=10, pady=5)
# Alt Thickness selection
Thick_var = tk.StringVar()
Thick_label = tk.Label(root, text="Select Thickness (actual):")
Thick_label.grid(row=1, column=0, padx=10, pady=5)
dropdown_menu = ttk.Combobox(root, textvariable=Thick_var, values=(['1/4','2/4','3/4','4/4','5/4','6/4','7/4','8/4']))
dropdown_menu.grid(row=1, column=1, padx=10, pady=5)
dropdown_menu.bind("<<ComboboxSelected>>", select_Thickness)
# Button to generate output
generate_button = tk.Button(root, text="Generate Output", command=generate_output)
generate_button.grid(row=4, column=0, columnspan=2, padx=10, pady=10)

# Start the Tkinter event loop
root.mainloop()