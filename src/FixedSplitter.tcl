
namespace eval FixedSplitter {
	variable properties
	proc _getState {w} {
		uplevel {
			variable properties
		}
	}
	proc create {w args} {
		_getState $w
		
		frame $w -width 200 -height 200 -relief sunken -bd 1
		
		set properties($w-orient) horizontal
		set properties($w-fixed) A
		set properties($w-windowA) {}
		set properties($w-windowB) {}
		set properties($w-barwidth) 8
		set properties($w-fixedsize) 100

		return $w
	}

}