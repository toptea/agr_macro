/*!
	Class: 	GlobalMacro
			Useful automation for all programs
*/
class GlobalMacro
{
	/*!
		Constructor: ()
			Creates an GlobalMacro object.
	*/
	__New(){
		
	}
	
	/*!
		Method: __Call()
			There's nothing in this meta-function!
	*/
	__Call(){
	
	}
	
	/*!
		Method: printDate()
			print the current date in dd/MM/yy format
			if Microsoft Excel is active, change the format code to dd/mm/yy
	*/
	printDate(){
		IfWinActive, Microsoft Excel
		{
			Send, +{F10}	; display shortcut menu for this cell (right click)
			Send, f			; Format Cells... dialog box
			Send, n			; 	Number tab
			Send, {Tab}		; 		Category
			Send, {Down 11}	; 		Category:Custom
			Send, {Tab}		; 			Type
			Send, dd/mm/yy	; 			fill in the format code
			Send, {Enter}	; OK
			Send, ^;		; Excel time stamp shortcut
			Send, {Enter}	; Accept
			Send, {Up}		; move the cursor back to its original position
			return	
		}
		else
		{
			Formattime, OutputVar, YYYYMMDDHH24MISS, dd/MM/yy
			Send %OutputVar%
			return
		}
	}
	
	/*!
		Method: viewPO()
			Enter the purchase order number and view it's corresponding pdf file
	*/
	viewPO(){
		InputBox, POnumber, Purchase Order, Please enter the PO number:, , 200, 120
		
		path := "\\Balmoral\Pegasus\Operations II\data\A_PDF\PO_00"
		end := ".pdf"
		if ErrorLevel
			MsgBox, CANCEL was pressed.
		else
		{
			;MsgBox, %path%%POnumber%%end%
			Run, %path%%POnumber%%end%,,UseErrorLevel
			if(ErrorLevel = "ERROR")
			{
				Msgbox,0,Oh no! An Error!, The system cannot find the file specified. Please enter the correct PO number!,60
			}
		}
			
	}
	
	/*!
		Method: viewDeliveryNote()
			Enter the delivery note number and view it's corresponding pdf file
	*/
	viewDeliveryNote(){
		InputBox, deliveyNumber, Delivery Note, Please enter the Delivery Note Number:, , 200, 135
		
		path := "\\Balmoral\Pegasus\Operations II\data\A_PDF\DNOTE_00"
		end := ".pdf"
		if (ErrorLevel == 1)
			MsgBox, CANCEL was pressed.
		else
		{
			
			;MsgBox, %path%%deliveyNumber%%end%
			Run, %path%%deliveyNumber%%end%,,UseErrorLevel
			if(ErrorLevel = "ERROR")
			{
				Msgbox,0,Oh no! An Error!, The system cannot find the file specified. Please enter the correct Delivery Note number!,60
			}
		}
			
	}

}

/*!
	End of class
*/
