/*!
	Library: General Macro 0.40
		This library contain a collection of macro I find useful.
	
	Author: Gary Ip
	
	;License: MIT
*/

;------------------------------------------------------------------------------
;Initialization
;------------------------------------------------------------------------------

#NoEnv
#SingleInstance force
;SendMode Input

SetWorkingDir, %A_ScriptDir%
SetKeyDelay, -1
SetTitleMatchMode, 2
CoordMode, Mouse, Screen
CoordMode, Pixel, Screen
;MouseGetPos, xpos, ypos

Excel := new ExcelMacro
Operation := new OperationMacro
Global := new GlobalMacro
Word := new WordMacro
Auto := new AutoWin
Meridian := new MeridianMacro
Inventor := new InventorMacro
Encompix := new EncompixMacro

;------------------------------------------------------------------------------
;Assigning Hotkeys
;------------------------------------------------------------------------------

;disable the window keys
~LWin Up:: return
~RWin Up:: return


Insert::Excel.insertRow()
;#c::Excel.copyLabel()
;#v::Excel.pasteSpecial("values2")
;#f::Excel.formatPainter()
;#9::Excel.jobcardStatus(2003,1,1,1,1,1,1)
#0::Excel.jobcardStatus(2007,1,1,1,1,1,1)
^5::Excel.deleteItem(2007)

#r::Operation.routing("PMU0552-028-00")
#1::Operation.unpick()
#a::Operation.updateAssy()
; ::Operation.viewPO() <- use Global.viewPO() because its faster and more reliable

^;::Global.printDate()
#p::Global.viewPO()
#n::Global.viewDeliveryNote()

;#F5::Auto.runAll()
;#e::Auto.runEmail()
;#o::Auto.runOperation()
;#c::Auto.runCooperation()

;#s::Word.createLabel()
#t::Meridian.build_report()
#d::Meridian.save_as("dxf")
#s::Meridian.save_as("pdf")
#m::Meridian.save_as("stp")
#z::Meridian.autocad_save_as()
#Esc::ExitApp


;^Left::Encompix.previous_record()
;^Right::Encompix.next_record()
;^+f::Encompix.search_record()
;^f::Encompix.find_record()

;Original Window Shortcuts
;#d::[Toggle all windows]
;#e::[Open Windows Explorer]
;#f::[Open Find dialog]
;#l::[Lock Workstation]
;#m::[Minimize all windows]
;#r::[Open Run dialog]
;#u::[Access Center]


;------------------------------------------------------------------------------
;Assigning Hotstrings
;------------------------------------------------------------------------------

:*bo:cci::Convert\import1.w
:*bo:ccp::C:\code\python\database_transfer\src\data\encompix

:*b0:iia::
	Excel.printIssue("Addition")
	return
	
:*b0:iid::
	Excel.printIssue("Deletion")
	return

:*b0:iir::
	Excel.printIssue("Rev Change")
	return

;------------------------------------------------------------------------------
;Import AHK Scripts
;------------------------------------------------------------------------------

#Include ExcelMacro.ahk
#Include OperationMacro.ahk
#Include GlobalMacro.ahk
#Include WordMacro.ahk
#Include AutoWin.ahk
#Include MeridianMacro.ahk
#Include EncompixMacro.ahk