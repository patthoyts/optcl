
# public interface 
namespace eval TlviewTooltip {
	proc Set {w text} {
		variable properties
		set properties($w-text) $text
		bind $w <Enter> [namespace code "Pending $w %X %Y"]
		bind $w <Leave> [namespace code "Hide"]
		Hide
	}

	proc Unset {w} {
		bind $w <Enter> {}
		bind $w <Leave> {}
		Hide
	}
}


# private stuff
namespace eval TlviewTooltip {
	variable properties
	set properties(window)	.__tlview__tooltip
	set properties(showbh)	1
	set properties(pending) {}
	

	destroy $properties(window)
	toplevel $properties(window) -bg SystemInfoText
	label $properties(window).l -text "Tooltip" -bg SystemInfoBackground -fg SystemInfoText
	pack $properties(window).l -padx 1 -pady 1

	
	wm overrideredirect $properties(window) 1
	wm withdraw $properties(window)
	

	proc Pending {w x y} {
		variable properties
		Cancel
		set properties(pending) [after 1000 [namespace code "Show $w $x $y"]]
	}

	proc Cancel {} {
		variable properties
		if {$properties(pending) != {}} {
			after cancel $properties(pending)
			set properties(pending) {}
		}
	}

	proc Show {w x y} {
		variable properties

		$properties(window).l configure -text $properties($w-text)
		wm deiconify $properties(window)
		incr x 8
		incr y 8
		wm transient $properties(window) $w
		wm geometry $properties(window) +$x+$y
	}

	proc Hide {} {
		variable properties
		Cancel
		wm withdraw $properties(window)
	}
}
