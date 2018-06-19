/*!
	Class: 	WordMacro
			Useful automation for Microsoft Word
*/
class WordMacro
{
	/*!
		Constructor: ()
			Creates an object.
	*/
	__New(){
		
	}
	
	/*!
		Method: __Call()
			Nothing!
	*/
	__Call(){
		
	}
	
	/*!
		Method: createLabels()
			Create labels for purchase parts.
			Need the correct docx and must use 2007 version.
	*/
	createLabel(){
		ifWinNotExist, ahk_class OpusApp
		{
			Run C:\Documents and Settings\GARY\Desktop\PUR Part Labels\Purchase Part Labels.docx
			Sleep, 1000
			Send, !{y}
			Sleep, 1000
		}
		ifWinNotActive, ahk_class OpusApp
		{
			WinActivate, ahk_class OpusApp
			Sleep, 1000
		}
		Send, {Alt}
		Send, mfe
		Send, {Enter}
	}
}

/*!
	End of class
*/