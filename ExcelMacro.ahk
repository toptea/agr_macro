/*!
	Class: 	ExcelMacro
			Useful automation for Microsoft Excel
	
	;Inherits: standalone
	
*/

class ExcelMacro
{
	/*!
		Constructor: ()
			Creates an ExcelMacro object.
	*/
	__New(){
		
	}
	
	/*!
		Method: __Call()
			Check if Microsft Excel is active
	*/
	__Call(){
		IfWinNotActive, Excel
		{
			Exit
		}
	}
		
	/*!
		Method: insertRow()
			Insert a new row in Excel
	*/	
	insertRow(){
		Send, +{Space}
		Send, ^{+}
		return
	}
	
	/*!
		Method: deleteItem()
			Strikethru and grey out highlighted cell  
	*/	
	deleteItem(version){
		Send, ^{5}
		Send, +{F10}
		Send, fnf
		Send, {Tab 4}
		Send, {k}
		Send, {Tab 4}
		Send, {Down 6}
		Send, {Left 5}
		Send, {Enter}
		Send, {Tab 2}
		Send, {Enter}
		Sleep, 500
	}
	
	
	/*!
		Method: copyLabel()
			For creating label spreadsheet. Copy the row and delete the qty and reference. 
	*/	
	copyLabel(){
		Send, +{Space}
		Send, ^c
		ExcelMacro.insertRow()
		Send, ^v
		Send, {Down}
		Send, {Left 2}
		Send, {Delete}
		Send, {Left 2}
		Send, {Delete}
		Send, {Right 4}
		Send, +{Space}
		Send, ^5
		Send, {Up}
		return
	}
	
	/*!
		Method: pasteSpecial()
			Macro for Excel's paste Special function.
			Add more info here!
		
		Parameters:
			"a" - "all"
			"f" - "formulas"
			"v" - "values"
			"t" - "formats"
			"c" - "comments"
			"h" - "sourceTheme"
			"x" - "allExceptborders"
			"w" - "columnWidths"
			"r" - "formulas2"
			"u" - "values2"
	*/
	pasteSpecial(pasteType, operationType:="none", skipBlank:=0, transpose:=0){
   
		#Include, pasteSpecial_arrays.ahk
		Send, +{F10}
		Send, {s 2}
		Send, o
	   
		For key, value in PastetypeArray
			if(pasteType = key || pasteType = value)
				Send, % PastetypeArray[key]		
	   
		For key, value in OperationTypeArray
		  if(operationType = key || operationType = value)
				Send, % OperationTypeArray[key]		
	   
		if (skipBlank=1)
			Send, b
	   
		if (transpose=1)
			Send, e
	   
		Send, {Enter}    
		return
	} 
	
	/*!
		Method: printIssue(changeType)
			Print out "Issue [changeType]" and highlight this cell bold.
		
		Parameters:
			changeType - This argument takes a string value. 
						 Usually its "Addition", "Deletion" or "Rev Change".
	*/
	printIssue(changeType){
		Send, {backspace 3}
		Send, Issue  
		Send, % changeType
		Send, {Enter}
		Send, {up}
		Send, ^{b}
		Send, {F2}
		Loop, % StrLen(changeType)
			Send, {left}
		Send, {Space 2}
		Send, {left}
		return
	}
	
	/*!
		Method: formatPainter()
			For Excel 2003, select the format painter icon using AHK's ImageSearch command.
			If 2007 version is used, use the keyboard shorcut instead.
	*/
	formatPainter(){
		MouseGetPos, xpos, ypos
		ImageSearch, img2003posx,img2003posy,0,0,A_ScreenWidth,A_ScreenHeight,brush2003.png
		If (ErrorLevel == 0) ; for excel 2003
		{
			img2003posx += 20
			img2003posy += 20
			Click, %img2003posx%,%img2003posy%
			MouseMove,  %xpos%, %ypos%
		}
		Else if(ErrorLevel == 1) ; for excel 2007
		{
			Send, {Alt}
			Send, hfp
		}
		Return
	}
	
	jobcardStatus(version, material, machined, inspect, platingOut, platingIn, complete){
		
		if (material == 1)
		{
			ExcelMacro.gotoColor(version)
			ExcelMacro.pickColor(version, 0,255,0)
		}	
		
		Send, {Right}
		if (machined == 1)
		{
			ExcelMacro.gotoColor(version)
			ExcelMacro.pickColor(version,255,255,0)
		}
		
		Send, {Right}
		if (inspect == 1)
		{
			ExcelMacro.gotoColor(version)
			ExcelMacro.pickColor(version,255,0,255)
		}
		
		Send, {Right}
		if (platingOut == 1)
		{
			ExcelMacro.gotoColor(version)
			ExcelMacro.pickColor(version,255,102,0)
		}
		
		Send, {Right}
		if (platingIn == 1)
		{
			ExcelMacro.gotoColor(version)
			ExcelMacro.pickColor(version,255,102,0)
		}
		
		Send, {Right}
		if (complete == 1)
		{
			ExcelMacro.gotoColor(version)
			ExcelMacro.pickColor(version,255,0,255)
		}
		
		Send, {left 5}
		
		return
	}
		
	gotoColor(version){
		if (version == 2003)
		{
			Send, +{F10}
			Send, fnp
			Send, {Tab}
			Send, {Down}
		}
		else
		{
			Send, +{F10}
			Send, fnff
			Send, !{m}
			Send, {Right}
			Send, {Tab 4}
		}
		return
	}

	pickColor(version,r,g,b){
		if (version == 2003)
		{
			if (r==0 && g==255 && b==0)
			{
				Send, {Down 3}
				Send, {Right 3}
			}
			if (r==255 && g==255 && b==0)
			{
				Send, {Down 3}
				Send, {Right 2}
			}
			if (r==255 && g==0 && b==255)
			{
				Send, {Down 3}
				
			}
			if (r==255 && g==102 && b==0)
			{
				Send, {Down}
				Send, {Right}
			}
			Send, {Enter}
			Send, {tab 2}
			Send, {Enter}
		}
		else
		{
			Send, %r%
			Send, {Tab}
			Send, %g%
			Send, {Tab}
			Send, %b%
			Send, {Tab}
			Send, {Enter}
			Send, {tab 3}
			Send, {Enter}
		}
		return
	}

		
}

/*!
	End of class
*/