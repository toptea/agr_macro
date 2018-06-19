/*!
	Class: 	MeridianMacro
			Useful automation for BlueCielo Meridian
*/

class MeridianMacro
{
	/*!
		Constructor: ()
			Creates an MeridianMacro object.
	*/
	__New(){
		
	}
	
	/*!
		Method: __Call()
			Check if BlueCielo Meridian is active
	*/
	;__Call(){
	;	IfWinNotActive, BlueCielo Meridian
	;	{
	;		Exit
	;	}
	;}
	
	/*!
		Method: build_report()
			build report to get revision number for the production jobcard
	*/
	build_report(){
		Send, {Alt}
		Send, {d}
		Send, {Up 10}
		Send, {Enter}
		Send, {Down}
		Send, {Enter}
		return
	}
    
    /*!
		Method: save_as(filetype)
			save drawing as a specific filetype
	*/
	save_as(filetype){
        MouseGetPos, xpos, ypos
		Click %xpos%, %ypos%,2
        Sleep 200
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
        
        if (filetype == "stp")
        {
            Send, asss
        }
        
        
		Send, +{Tab}
		Send, {Home}
		Send, C:\Users\GARY\Desktop\CAD\
		Sleep 500
        Send, !s
        Sleep 2000
        
        if (filetype == "pdf")
        {

            WinClose, Adobe Acrobat Reader
            WinActivate, Autodesk Inventor
            Sleep 500
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
    
    autocad_save_as(){
        ;MouseGetPos, xpos, ypos
		;Click %xpos%, %ypos%,2
        ;Sleep 500
        ;Send, {Enter}
        ;Sleep 1000
        ;Send, {Enter}
        ;Sleep 1000
        ;Send, {Enter}
        ;Sleep 2000
		Send, {Alt}
		Send, {f}
		Send, {e}
        Send, {p}
		Sleep, 1000
		Send, {Home}
		Send, C:\Users\GARY\Desktop\CAD\
        Send, !x
        Send, e
        Send, !s
        Sleep 2000
        WinClose, Adobe Acrobat Reader
        WinActivate, AutoCAD
        Sleep 500
        Send, {Alt}
        Send, {f}
        Send, {c}
        Sleep 500
        Send, {n}
        Send, {Escape}
        Sleep 500
        WinMinimize, AutoCAD
        WinActivate, BlueCielo Meridian

    }
	
}

/*!
	End of class
*/