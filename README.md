# Automatic-Inventory-Management-System
Instant print and scan system with data storage on excel for inventory management

1. Used for instant printing of barcode onto labels. Current alignment settings are according to 5 x 2.5cm labels
2. A Start/Stop button is used to initialize the whole process. This button will follow the selected cell so as to not lose track of the button.
3. Currently developed for workbook with one sheet. Aiming to develop for multiple sheets in one book.
4. The setting of the scanner is CR Suffix, such that after scanning the barcode, the cell under the code will be selected. Like hitting Enter after writing the code

Barcode Scanner used: 2D Imaging Scanner RETSOL D-2060

Barcode label Printer used: Xprinter XP-350B

Labels used: 5 x 2.5 cm

STEP 1:Download and install barcode font 
          
STEP 2: Open Excel .xlsm file (Excel Macro Enabled Workbook) 

          	Excel > new sheet > save as > .xlsm
          
STEP 3: Enable developer mode 

			File > Options > Customize Ribbon 
			
			In the section of the right side, tick the ‘Developer’ option > OK 
			
			A new ‘developer’ tab will appear on the tools bar 
          
STEP 4: Go to visual basic to import / edit macro 

			Developer > Visual Basic 
			
			Right click on the Project Explorer section 
			
			Import file > Select the Module2.bas file 
			
			Import file > Select ThisWorkbook1.cls file 
			
			IMPORTANT: Copy code from ThisWorkbook1 and paste in the already present ThisWorkbook file and save it 
        
STEP 5: Close the Excel sheet and reopen it 

        	After reopening the file a START button should be visible 
			
			Press the button to start scanning 


Edits in macros for customization: 

1. To only get the barcode on the excel sheet without printing: 

    Open Module2.bas 
	
    Comment out 'Call PrintSet(cell_row, cell_col)' and 'Cells(cell_row - 1, cell_col + 1).PrintOut' 
	
2. To use scanner at a different suffux mode: 

    Open Module2.bas 
	
    Convert all 'cell_row - 1' to 'cell_row' 

