
namespace eval optcl_utils {
namespace export *
# nice and simple error catching
proc try {body catch errvarname catchblock} {
	upvar $errvarname errvar

	if {![string match $catch catch]} {
		error "invalid syntax - should be: try {body} catch {catchblock}"
	}
	if { [catch [list uplevel $body] errvar] } {
		uplevel $catchblock
	} else {
		return $errvar
	}
}

proc throw {errmsg} {
	uplevel [list error $errmsg]
}


proc center_window {w} {
	set width [winfo reqwidth $w]
	set height [winfo reqheight $w]
	set swidth [winfo screenwidth $w]
	set sheight [winfo screenheight $w]

	set x [expr ($swidth - $width) /2 ]
	set y [expr ($sheight - $height) /2 ]
	wm geometry $w +$x+$y
}

# set multiple variables with the contents of the list
proc mset {vars values} {
	foreach var $vars value $values {
		if {$var == {}} {error "not enough variables for mset operation"}
		upvar $var myvar
		set myvar $value
	}
}

}