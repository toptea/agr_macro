/*!
	Class: AutoWin	
			Useful automation for Microsoft Word
*/
class AutoWin
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
		Method: runAll()
	*/
	runAll(){
		Autorun.runEmail()
		Sleep, 2000
		Autorun.runOperation()
		Sleep, 2000
		Autorun.runCooperation()
		return
	}
	
	/*!
		Method: runEmail()
	*/
	runEmail(){
		ifWinNotActive, ahk_class rctrl_renwnd32
			WinActivate, ahk_class rctrl_renwnd32
		IfWinNotExist, ahk_class rctrl_renwnd32
			Run, C:\Program Files\Microsoft Office\Office12\OUTLOOK.EXE
			
		return
	}

	/*!
		Method: runOperation()
	*/
	runOperation(){
		ifWinNotActive, ahk_class Operation6c000000
			WinActivate, ahk_class Operation6c000000
		IfWinNotExist, ahk_class Operation6c000000
			;Run, C:\Program Files\Pegasus\Operations II\Operations.exe
			Run, C:\Documents and Settings\GARY\My Documents\Shortcut\Pegasus Operations II.lnk 
			Sleep, 3000
			Send, Gary
			Send, {Tab}
			Send, abc
			Send, {Enter 3}
		return
	}

	/*!
		Method: runCooperation()
	*/
	runCooperation(){
		ifWinNotActive, ahk_class OMain
			WinActivate, ahk_class OMain
		IfWinNotExist, ahk_class OMain
			Run, C:\Documents and Settings\GARY\My Documents\Shortcut\Co-operation II.lnk
			Sleep, 3000
			WinActivate, Microsoft Access
			Sleep, 1000
			Send, user
			Send, {Enter}
		return
	}
}

/*!
	End of class
*/