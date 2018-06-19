/*!
	Class: 	InventorMacro
			Useful automation for Autodesk Inventor
*/

class InventorMacro
{
	/*!
		Constructor: ()
			Creates an InventorMacro object.
	*/
	__New(){
		
	}
	
	/*!
		Method: __Call()
			Check if Autodesk Inventor is active
	*/
    
    /*!
	__Call(){
		IfWinNotActive, Autodesk Inventor
		{
			Exit
		}
	}
	*/
    
	/*!
		Method: save_as(filetype)
			save drawing as a specific filetype
	*/
	save_as(filetype){
        MouseGetPos, xpos, ypos
		Click %xpos%, %ypos%,2
        Sleep 500
        Send, {Enter}
        Sleep 2000
		Send, {Alt}
		Send, {f 2}
		Send, {c}
		Sleep, 1000
		Send, {Tab}
        
        if (filetype == "dwg")
        {
            Send, daa
        }
        
        if (filetype == "dwf")
        {
            Send, ad
        }
        
        if (filetype == "dxf")
        {
            Send, addd
        }
        
        if (filetype == "pdf")
        {
            Send, ap
        }
        
        if (filetype == "png")
        {
            Send, app
        }
        
        
		Send, +{Tab}
		Send, {Home}
		Send, C:\Users\GARY\Desktop\cad\
		Sleep 500
        Send, !s
        
        if (filetype == "pdf")
        {
            Sleep 2000
            WinClose, Adobe Acrobat Reader
        }
        
        Send, {Alt}
        Send, {f}
        Send, {c}
        Send, {n}
        Send, {Escape}
        Sleep 500
        WinMinimize, Autodesk Inventor
        WinActivate, BlueCielo Meridian
		return
	}

}

/*!
	End of class
*/