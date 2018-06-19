/*!
	Class: 	EncompixMacro
			Useful automation for Encompix ERP
*/
class EncompixMacro
{
	/*!
		Constructor: ()
			Creates an object.
	*/
	__New(){
		
	}
	
	/*!
		Method: __Call()
			Check if Encompix is active
	*/
	__Call(){
		IfWinNotActive, Encompix
		{
			Exit
		}
	}
	
    
	previous_record(){
		Send, ^{F1}
	}
    
    next_record(){
		Send, ^{F2}
	}
    
    find_record(){
        Send, ^{F3}
    }
    
    search_record(){
        Send, ^{F4}
    }
}

/*!
	End of class
*/