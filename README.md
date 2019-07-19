# Corel_macroses
VBA macros for Corel
Useful tiny scripts for Corel in geophysics:-)

1. Text_rescale.bas
This small script allow you to get text size and proportions back to the original parameters after rescaling of whole vector picture 

	How to use:

	1. Select all objects that you have rescaled (shapes, texts etc.)
	2. Run script:		
	      Press F5 in a Microsoft Visual Basic window
	3. Enter rescale sizes   	
	      You can find it  in the field at the left upper corner if you select any rescaled object
	4. Done! :-)

2.xy from shape.bas
Script for extracting coordinates of each node of the curve. It is useful for extract well curves from .cdr format to .las

	How to use:
	
	1. Select one curve you want to extract nodes values
	2. Run script:		
	      Press F5 in a Microsoft Visual Basic window
	3. Enter your output filename and path    	
	       For example C:\data\nodes.csv
	4. Done! :-)


      
How to add script to Corel:

	1.download script

	2.press Alt+F11 in a Corel

	3.Import script:		
	      File - Import file - your_script.bas
