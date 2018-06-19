/*!
	Class: 	OperationMacro
			Useful automation for Pegasus Operation
*/

class OperationMacro
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
		IfWinNotActive, Pegasus Operation
		{
			Exit
		}
	}
	
	/*!
		Method: routing(code)
			In pegasus operation's stock detail window,
			create the routing process for this assembly.
		
		Parameters:
			code -  We normally use "PMU0552-028-00" code. See operation's help manual for more info.
	*/
	routing(code){
		Send !a				;Action Menu
		Send r				;Routing Window
		Send !a				;	Action Menu
		Send m				;		Model sub-window
		Send %code%			;		Enter routing code in textbox
		Send {Enter}		;		Exit textbox
		Send {Enter}		;		OK in Model sub-window 
		Send !o				;		OK in Routing Layout sub-window
		Send {Enter}		;		OK to "save all changes made to all records" dialog
		Send {Enter 4}		;		Cycle thru and OK in Change Comments sub-window
		return
	}

	/*!
		Method: unpick()
			In pegasus operation's pick list window, 
			unpick the item where the mouse cursor is
	*/
	unpick(){
		MouseGetPos, xpos, ypos
		Click %xpos%, %ypos%,2 	; Click cell for Batch sub-window
		Send 0					;	0
		Send {Enter}			; 	Exit textbox
		Send !o					;	OK
		return
	}
	
	/*!
		Method: updateAssy()
			In pegasus operation's stock detail window,
			update the assembly definition for this assembly
	*/
	updateAssy(){
		;Send, a			;Action Menu
		;Send, a			;Assembly Definition
		Sleep 500		;	wait until the Assy Def Window is loaded
		Send, +{Tab}	;	select assembly part
		Send, {F8}		;	expand assembly
		Send, {Down}	;	select first part
		Send, {F11}		;	Edit (open Line Details)
		Send, {Tab}		;		BOM Quantity textbox
		Send, ^{a}		;		select qty
		Send, ^{c}		;		copy qty
		Send, ^{v}		;		paste qty
		Send, !o		;		OK
		Send, !o		;	OK
		Sleep 500		;		wait until assembly options sub-window is loaded
		Send, {Tab 6}	;		go to change option
		Send, {Down}	;		highlight "Change Base & Jobs"
		Send, {Enter}	;		select "Change Base & Jobs"
		Send, !o		;		OK
		return
	}
	
	/*!
		Method: viewPO()
			In pegasus operation, view purchase order
	*/
	viewPO(){
		InputBox, POnumber, Purchase Order, Please enter the Purchase Order number., , 250, 120
		if ErrorLevel
		{
			MsgBox, CANCEL was pressed.
		}
		else
		{
			Send,{Alt}				; Activiate menu
			Send, p					; Purchase menu
			Send, r					; Print Orders
			Sleep, 500				;	wait for the Print Orders window to load
			loop, 2
			{
				Send, {Tab}			;	goto Lower Order No		
				Send, ^{a}			;	select old PO number
				Send, %POnumber% 	;	enter PO number
			}
			Send, {Tab 3}			;	goto "Include Reprints"
			Send, {Enter}			;	select
			Send, {Tab 3}			;	goto OK
			Send, {Enter}			;	OK
			Sleep, 500
			Send, {left}
			Send, {Tab}
			Send, {Enter}
		}
	}
	
}

/*!
	End of class
*/