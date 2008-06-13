#----------------------------------------------------------------------
# Splitter.tcl
#	Beautiful splitter, just the way mr fuzz likes it :-)
#	Proportional or fixed (on one of either two panes). One window
#	per splitter. Horizontal or vertical splits.
#
# Date First Completed: 2000.05.02
# Last Updated: 2000.05.02
# Author: fuzz@sys.uea.ac.uk
#----------------------------------------------------------------------


# Splitter --
# Global level function to create a Splitter - delegates to the 
# Splitter::create namespace command. Requires the window
# name followed by a collection of options
proc Splitter {w args} {
	return [eval Splitter::create $w $args]
}

# Namespace Splitter
# Encapsulates the Splitter functionality
namespace eval Splitter {

	namespace import -force ::optcl_utils::*

	# all splitter properties are held here
	# indexing uses the window name followed by the property in the form $window-$property
	variable properties

	# _GetState --
	# Retrieves state information regarding window w
	# Actually all it does for now is bring the 'properties' variable into scope
	proc _GetState {w} {
		uplevel {
			variable properties
			set bar $w._splitbar
		}
	}


	# create --
	# Give a window name and a list of options, attempts to create a splitter widget
	proc create {w args} {
		_GetState $w
		frame $w -width 200 -height 200 -relief flat -bd 0
		set properties($w-position) 0.5
		set properties($w-orientation) horizontal
		set properties($w-windowA)	{}
		set properties($w-windowB)	{}
		set properties($w-barwidth)	8
		set properties($w-type) prop
		set properties($w-min) 0
		set properties($w-max) -1

        frame $bar -bd 1 -relief raised -width $properties($w-barwidth) -height $properties($w-barwidth) -cursor sb_v_double_arrow
		bind $bar <ButtonPress-1> [namespace code "_StartDrag $w"]
		bind $bar <ButtonRelease-1> [namespace code "_FinishDrag $w"]
		try {
			eval configure $w $args
		} catch {er} {
			destroy $w
			throw $er
		}
		_UpdatePosition $w
		rename ::$w [namespace current]::$w
		proc ::$w {command args} "eval [namespace current]::_Dispatch $w \$command \$args"
		eval ::$w configure $args
		return $w
	}	

	# _Dispatch --
	# First port of call for the widget command. Attempts to find a suitable command
	# (performing command completion), before dispatching the call.
	proc _Dispatch {w command args} {
		set cmd [info commands $command*]
		if {$cmd == {}} {
			throw "no such command: $command"
		}
		return [eval [lindex $cmd 0] $w $args]
	}


	
	# cget --
	# retrieves the settings for this splitter or its underlying frame
	proc cget {w option} {
		_GetState $w
		switch -- $option {
			-orient {return $properties($w-orientation)}
			-position {return $properties($w-position)}
			-windowA {return $properties($w-windowA)}
			-windowB {return $properties($w-windowB)}
			-barwidth {return $properties($w-barwidth)}
			-type {return $properties($w-type)}
			-min {return $properties($w-min)}
			-max {return $properties($w-max)}
			default {return [$w cget $option]}
		}
	}


	# configure --
	# Performs basic widget (re)configuration. Defaults to the frame configuration command
	# if option not found. Doesn't support option value query.
	proc configure {w args} {
		_GetState $w
		if { [expr [llength $args] % 2] != 0 } {
			throw "options must be in pairs"
		}
     	
		set unhandled {}   
		set update 0

		foreach {option value} $args {
			switch -- $option {
            -orient {
				_SetOrientation $w $value
				set update 1
			}
			-position 	{
				set properties($w-position) $value; 
				set update 1
			}
            -windowA	{
				_SetWindow $w windowA $value
				set update 1
			}

            -windowB	{
				_SetWindow $w windowB $value
				set update 1
			}
			-barwidth {
				set properties($w-barwidth) $value
				set update 1
			}
			-type {
				_SetType $w $value
				set update 1
			}
			-min {
				set properties($w-min) $value
				if {$value > $properties($w-position)} {
					set properties($w-position) $value
				}
				set update 1
			}
			-max {
				set properties($w-max) $value
				if {$value < $properties($w-position) && $value >= 0} {
					set properties($w-position) $value
				}
			}
			default {lappend unhandled $option $value}
			}

			if {$unhandled != {}} {
				eval $w config $unhandled
				set update 1
			}

			if {$update} {
				_UpdatePosition $w
			}
		}
	}

	# _UpdatePosition --
	# This is the main proc for setting up the visual representation of the splitter
	proc _UpdatePosition {w} {
		_GetState $w
		_UpdateBar $w
		switch $properties($w-orientation) {
			horizontal {
				_UpdatePositionHorizontal $w
			}
			vertical {
				_UpdatePositionVertical $w
			}
		}
	}


	# _UpdatePositionHorizontal --
	# Called to set window A and B when split is horizontal
	proc _UpdatePositionHorizontal {w} {
        _GetState $w
        set height [winfo height $w]
        set pos $properties($w-position)

        if {$properties($w-type) != "prop"} {
            if {$pos < 0} {
                set pos 0
            } elseif {$pos > $height} {
                set pos $height
            }
        }
        set hbw [expr $properties($w-barwidth) / 2]
        switch $properties($w-type) {
            prop {
                if {$properties($w-windowA) != {} && [winfo exists $properties($w-windowA)]} {
                    place $properties($w-windowA) -in $w -x 0 -y 0 -relx 0 -rely 0 -relheight $pos -relwidth 1.0 -height [expr -$hbw] -width 0
                }
                if {$properties($w-windowB) != {} && [winfo exists $properties($w-windowB)]} {
                    place $properties($w-windowB) -in $w -y $hbw -x 0 -rely $pos -relx 0 -relwidth 1.0 -relheight [expr 1.0 - $pos] -height [expr -$hbw] -width 0
                }
            }
            fixA {
                if {$properties($w-windowA) != {} && [winfo exists $properties($w-windowA)]} {
                    place $properties($w-windowA) -in $w -y 0 -x 0 -rely 0 -relx 0 -relwidth 1.0 -relheight 0 -height [expr $pos - $hbw] -width 0
                }
                if {$properties($w-windowB) != {} && [winfo exists $properties($w-windowB)]} {
                    place $properties($w-windowB) -in $w -y [expr $properties($w-position) + $hbw] -x 0 -rely 0 -relx 0 -relwidth 1.0 -relheight 0 -height [expr $height - $pos - $hbw] -width 0
                }
            }
            fixB {
                if {$properties($w-windowA) != {} && [winfo exists $properties($w-windowA)]} {
                    place $properties($w-windowA) -in $w -y 0 -x 0 -rely 0 -relx 0 -relwidth 1.0 -relheight 0 -height [expr $height - $pos - $hbw] -width 0
                }
                if {$properties($w-windowB) != {} && [winfo exists $properties($w-windowB)]} {
                    place $properties($w-windowB) -in $w -y [expr $height - $pos + $hbw] -x 0 -rely 0 -relx 0 -relwidth 1.0 -relheight 0 -height [expr $pos - $hbw] -width 0
                }
            }
        }
    }



	# _UpdatePositionVertical --
	# Called to place window A and B when split is vertical
	proc _UpdatePositionVertical {w} {
		_GetState $w
		set width [winfo width $w]
		set pos $properties($w-position)
		if {$properties($w-type) != "prop"} {
			if {$pos < 0} {
				set pos 0
			} elseif {$pos > $width} {
				set pos $width
			}	
		}
		set hbw [expr $properties($w-barwidth) / 2]
		switch $properties($w-type) {
			prop {
				if {$properties($w-windowA) != {} && [winfo exists $properties($w-windowA)]} {
                    place $properties($w-windowA) -in $w -x 0 -y 0 -relx 0 -rely 0 -relheight 1.0 -relwidth $pos -width [expr -$hbw] -height 0
				}
				if {$properties($w-windowB) != {} && [winfo exists $properties($w-windowB)]} {
					place $properties($w-windowB) -in $w -x $hbw -y 0 -relx $pos -rely 0 -relheight 1.0 -relwidth [expr 1.0 - $pos] -width [expr -$hbw] -height 0
				}
			}
			fixA {
				if {$properties($w-windowA) != {} && [winfo exists $properties($w-windowA)]} {
                    place $properties($w-windowA) -in $w -x 0 -y 0 -relx 0 -rely 0 -relheight 1.0 -relwidth 0 -width [expr $pos - $hbw] -height 0
				}
				if {$properties($w-windowB) != {} && [winfo exists $properties($w-windowB)]} {
                    place $properties($w-windowB) -in $w -x [expr $pos + $hbw] -y 0 -relx 0 -rely 0 -relheight 1.0 -relwidth 0 -width [expr $width - $pos - $hbw] -height 0
				}
			}
			fixB {
				if {$properties($w-windowA) != {} && [winfo exists $properties($w-windowA)]} {
					place $properties($w-windowA) -in $w -x 0 -y 0 -relx 0 -rely 0 -relheight 1.0 -relwidth 0 -width [expr $width - $pos - $hbw] -height 0
				}
				if {$properties($w-windowB) != {} && [winfo exists $properties($w-windowB)]} {
					place $properties($w-windowB) -in $w -x [expr $width - $pos + $hbw] -y 0 -relx 0 -rely 0 -relheight 1.0 -relwidth 0 -width [expr $pos - $hbw] -height 0
				}
			}
		}
	}


	# _UpdateBar --
	# Main proc for the visual setting of the split bar.
	proc _UpdateBar {w} {
		_GetState $w
		
		switch $properties($w-orientation) {
			horizontal {
				_UpdateBarHorizontal $w
			}
			vertical {
				_UpdateBarVertical $w
			}
		}
	}

	# _UpdateBarHorizontal --
	# Called to set the bar when the split is horizontal
	proc _UpdateBarHorizontal {w} {
        _GetState $w
        set bd [$w cget -bd]
        
        if {$bd >= 1} {
            $bar config -bd $bd
        } else {
            set bd 1
        }
        
        set height [winfo height $w]
        set pos $properties($w-position)
        if {$properties($w-type) != "prop"} {
            if {$pos < 0} {
                set pos 0
            } elseif {$pos > $height} {
                set pos $height
            }
        }
        
        switch $properties($w-type) {
            prop {
                place $bar -y 0 -rely $pos -x [expr -2*$bd] -relx 0 -height $properties($w-barwidth) -relheight 0 -width [expr 4 * $bd] -relwidth 1.0 -anchor w
            }
            fixA {
                place $bar -y $pos -rely 0 -x [expr -2*$bd] -relx 0 -height $properties($w-barwidth) -relheight 0 -width [expr 4 * $bd] -relwidth 1.0 -anchor w
            }
            fixB {
                place $bar -y [expr $height - $pos] -rely 0 -x [expr -2*$bd] -relx 0 -height $properties($w-barwidth) -relheight 0 -width [expr 4 * $bd] -relwidth 1.0 -anchor w
            }
        }
	}

	# _UpdateBarVertical --
	# Called to set the bar when split is vertical.
	proc _UpdateBarVertical {w} {
		_GetState $w
		set bd [$w cget -bd]

		if {$bd >= 1} {
  	  		$bar config -bd $bd
		} else {
			set bd 1
		}

		set width [winfo width $w]
		set pos $properties($w-position)
		if {$properties($w-type) != "prop"} {
			if {$pos < 0} {
				set pos 0
			} elseif {$pos > $width} {
				set pos $width
			}
		}

		switch $properties($w-type) {
			prop {
				place $bar -x 0 -relx $pos -y [expr -2*$bd] -rely 0 -width $properties($w-barwidth) -relwidth 0 -height [expr 4 * $bd] -relheight 1.0 -anchor n
			}
			fixA {

				place $bar -x $pos -relx 0 -y [expr -2*$bd] -rely 0 -width $properties($w-barwidth) -relwidth 0 -height [expr 4 * $bd] -relheight 1.0 -anchor n
			}
			fixB {
				place $bar -x [expr $width - $pos] -relx 0 -y [expr -2*$bd] -rely 0 -width $properties($w-barwidth) -relwidth 0 -height [expr 4 * $bd] -relheight 1.0 -anchor n
			}
		}
	}


	# destroy --
	# Proc to destroy this object
	proc destroy {w} {
		_GetState $w
		destroy $w
	}

	# _SetOrientation --
	# Called to set the orientation (vertical or horizontal) on this splitter.
	proc _SetOrientation {w orientation} {	
		_GetState $w
        if {![regexp ^(vertical|horizontal)$ $orientation]} {
			throw "bad orientation - should be either horizontal or vertical"
		}
		set properties($w-orientation) $orientation
		
		if {[string match $orientation "vertical"]} {
			$bar config -cursor sb_h_double_arrow
		} else {
			$bar config -cursor sb_v_double_arrow
		}  
	}

	# _SetWindow --
	# Called to set the window for pane A or B.
	# w is the splitter window. winX is either windowA or windowB
	# child is the child-window to be placed.
	proc _SetWindow {w winX child} {
		_GetState $w
		if {$properties($w-$winX) != {} && [winfo exists $properties($w-$winX)]} {
			place forget $properties($w-$winX)
		}
		if {$child != {}} {
			lower $child $bar
		}
		set properties($w-$winX) $child
	}

	# _StartDrag --
	# Called when the user clicks on the bar.
	proc _StartDrag {w} {
		_GetState $w
		# create a proxy 
		set proxy [frame $w._proxysplitbar -bd [$bar cget -bd] -relief raised]
		eval place $proxy [place info $bar]
		lower $proxy $bar
		focus $bar
		bind $bar <Motion> [namespace code "_Move $w %X %Y"]	
		bind $bar <KeyPress-Escape> [namespace code "set properties($w-position) $properties($w-position); _FinishDrag $w"]
	}

	# _FinishDrag --
	# Called when the user releases the mouse button from the bar.
	proc _FinishDrag {w} {
		_GetState $w
		bind $bar <Motion> {}
		bind $bar <KeyPress-Escape> {}
		_UpdatePosition $w
		::destroy $w._proxysplitbar
	}
	
	# _Move --
	# Called when 
	proc _Move {w x y} {
		_GetState $w
		
		if {$properties($w-type) == "prop"} {		
			set x [expr double($x - [winfo rootx $w] + 2) / double([winfo width $w])]
			set y [expr double($y - [winfo rooty $w] + 2) / double([winfo height $w])]
		
			if {$x < 0.0} {
				set x 0.0
			} elseif {$x > 1.0} {
				set x 1.0
			}

			if {$y < 0.0} {
				set y 0.0
			} elseif {$y > 1.0} {
				set y 1.0
			}
		} elseif {$properties($w-type) == "fixA"} {
			set x [expr $x - [winfo rootx $w]]
			set y [expr $y - [winfo rooty $w]]
		} else {
			set x [expr [winfo width $w] - ($x - [winfo rootx $w])]
			set y [expr [winfo height $w] - ($y - [winfo rooty $w])]
		}
		if {$x < $properties($w-min)} {
			set x $properties($w-min)
		} elseif {$x > $properties($w-max) && $properties($w-max) >= 0} {
			set x $properties($w-max)
		}
		if {$y < $properties($w-min)} {
			set y $properties($w-min)
		} elseif {$y > $properties($w-max) && $properties($w-max) >= 0} {
			set y $properties($w-max)
		}

		switch $properties($w-orientation) {
		horizontal {set properties($w-position) $y}
		vertical {set properties($w-position) $x}
		}

		_UpdateBar $w
	}

	proc _SetType {w value} {
		_GetState $w
		if {![regexp {^(fixA)|(fixB)|(prop)$} $value]} {
			throw "bad position type '$value': should be fixA, fixB or prop"
		}

		
		if {![string match $properties($w-type) $value]} {
			
			if {[string match $properties($w-orientation) "vertical"]} {
				set d [winfo width $w]
			} else {
				set d [winfo height $w]
			}
			
			
			if {$properties($w-type) == "prop"} {
				# going from prop to fix*
				set v [expr int(double($d) * $properties($w-position))]
				set min [expr int(double($d) * $properties($w-min))]
				set max [expr int(double($d) * $properties($w-max))]
				if {$value == "fixB"} {
					set v [expr $d - $v]
					set min [expr $d - $min]
					set max [expr $d - $max]
				}
				set properties($w-position) $v
				set properties($w-min) $min
				if {$properties($w-max) >= 0} {
					set properties($w-max) $max
				}
			} elseif { ($value == "fixA" && $properties($w-type) == "fixB") ||
					   ($value == "fixB" && $properties($w-type) == "fixA") } {
				# going from fixA to fixB or vice versa
				set properties($w-position) [expr $d - $properties($w-position)]
				set properties($w-min) [expr $d - $properties($w-min)]
				if {$properties($w-max) >= 0} {
					set properties($w-max) [expr $d - $properties($w-max)]
				}
			} else {
				# going from fix* to prop
				if {$properties($w-type) == "fixB"} {
					set properties($w-position) [expr $d - $properties($w-position)]
					set properties($w-min) [expr $d - $properties($w-min)]
					if {$properties($w-max) >= 0} {
						set properties($w-max) [expr $d - $properties($w-max)]
					}
				}
				set properties($w-position) [expr double($properties($w-position)) / double($d)]
				set properties($w-min) [expr double($properties($w-min)) / double($d)]
				if {$properties($w-max) >= 0} {
					set properties($w-max) [expr double($properties($w-max)) / double($d)]
				}
			}
			set properties($w-type) $value
		}

		
		if {$value == "prop"} {
			# we're not interested in window resize events any more 
			bind $w <Configure> {}
		} else {
			# bind to Configure event to watch window size changes
			bind $w <Configure> [namespace code {_OnConfigure %W}]
		}
	}

	proc _OnConfigure {w} {
		_UpdatePosition $w
	}

}
set type prop      
    
proc MakeHorizontal {w b} {
	$w config -orient horizontal
	$b config -command [namespace code "MakeVertical $w $b"]
}

proc MakeVertical {w b} {
	$w config -orient vertical
	$b config -command [namespace code "MakeHorizontal $w $b"]
}

proc ChangeType {w} {
	variable type
	$w config -type $type
}

if {0} {
frame .f -bd 2 -relief groove
pack .f -side bottom -fill x
button .f.orient -text H/V -width 5 -command [namespace code "MakeVertical .s .f.orient"]
checkbutton .f.prop -text prop -variable type -onvalue prop -indicatoron 0 -width 5 -command [namespace code "ChangeType .s"]
checkbutton .f.fixA -text fixA -variable type -onvalue fixA -indicatoron 0 -width 5 -command [namespace code "ChangeType .s"]
checkbutton .f.fixB -text fixB -variable type -onvalue fixB -indicatoron 0 -width 5 -command [namespace code "ChangeType .s"]
pack .f.orient .f.prop .f.fixA .f.fixB -side left

Splitter .s -orient horizontal -type fixB -min 40
pack .s -fill both -expand 1 
Splitter .s.l1 -orient vertical -type fixA -position 100
label .s.l2 -text Two -bg green
.s config -windowA .s.l1 -windowB .s.l2
label .s.l1.l1 -text One.One -bg red
label .s.l1.l2 -text One.Two -bg blue -fg white
.s.l1 config -windowA .s.l1.l1 -windowB .s.l1.l2
console show
}