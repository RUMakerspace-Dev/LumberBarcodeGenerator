# LumberBarcodeGenerator
Codebase and Spreadsheet config for the barcode generator app @RUMS woodshop
This project is to generate on-the-fly barcodes for various types and sizes of lumber sold in the woodshop. 
The python file uses pandas to read an excel table, then uses the SKU data to create a UPC-A barcode. This barcode is passed into a library that formats a 
custom label for the Brother ql label printer used with our POS. 

The python file has multiple dependencies, some of which require older versions of the libraries to run properly.
Please read documentation for the Brother_ql raster software CAREFULLY before making edits to that part of the code.
The python script is compiled into an .exe using pyinstaller. This will create a full build, but the .xlsx "config" will need to be kept on the top-level folder along with the exe. 
winUSB-32 will also be needed for each new machine the software is running on. You can read more about that from the Brother_ql library.


