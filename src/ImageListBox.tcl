proc ImageListBox {args} {
	return [eval ImageListBox::create $args]
}

namespace eval ImageListBox {
	variable properties
	array set properties {}

	proc _getState {w} {
		uplevel {variable properties}
	}

	proc create {w args} {
		_getState $w
		text $w -bd 0 -relief flat -width 30 -height 15 -state disabled -cursor arrow -wrap none
		set properties($w-item) {}
		set properties($w-nextId) 0
		set properties($w-active) 0
		set properties($w-anchor) 0
		set properties($w-selectmode) browse
		set properties($w-uiselection) 0		

		$w tag configure ILB_Selection -background SystemHighlight -foreground SystemHighlightText
		set font [$w cget -font]
		$w tag configure ILB_Active -foreground red

		setBindings $w
		rename ::$w [namespace current]::$w
		proc ::$w {cmd args} "return \[eval [namespace current]::_dispatch $w \$cmd \$args\]"

		eval $w configure $args
		return $w
	}

	proc setBindings {w} {
		foreach binding [bind Text] {
			bind $w $binding "break;"
		}

		foreach binding [bind Listbox] {
			bind $w $binding "[bind Listbox $binding];break;"
		}

		# special bindings
		bind $w <ButtonPress-1> "[namespace current]::OnBeginUISelection $w; [bind Listbox <ButtonPress-1>]; break;"
		bind $w <ButtonRelease-1> "[namespace current]::OnEndUISelection $w; [bind Listbox <ButtonRelease-1>]; break;"
		return
	}

	proc _dispatch {w cmd args} {
		_getState $w
		set cmds [info commands [namespace current]::$cmd*]
		if {$cmds == {}} {
			return [eval $w $cmd $args]
		} else {
			return [eval [lindex $cmds 0] $w $args]
		}
	}



	proc insert {w index args} {
		_getState $w
		
		set bEnd [string match $index end]
		if {!$bEnd} {
			incr index
		} else {
			set index [expr int([$w index end])]
		}

		$w config -state normal
		foreach item $args {
			$w image create $index.0 -align center -name _ILB_IMAGE_$properties($w-nextId)
			$w insert $index.1 $item\n
			$w tag add _ILB_TAG_$properties($w-nextId) $index.0 $index.end 

			incr properties($w-nextId)
			incr index
		}
		$w config -state disabled
	}
	
	proc setimage {w index image} {
		_getState $w
		set index [index $w $index]
		if {$index >= [index $w end]} {
			set index [expr [index $w end] - 1]
		}
		set pos [expr $index + 1].0
		$w image configure $pos -image $image
	}

	proc getimage {w index} {
		_getState $w
		set index [index $w $index]
		set pos [expr $index + 1].0
		$w image cget $pos -image
	}

	proc delete {w first {last {}} } {
		_getState $w
		
		if {$last == {}} {
			set last $first
		}
		set first [index $w $first]
		set last [index $w $last]

		incr first
		incr last 2
		$w config -state normal
		$w delete $first.0 $last.0
		$w config -state disabled
	}

	proc size {w} {
		_getState $w
		return [expr int([$w index end]) - 2]
	}

	proc get {w first {last {}} } {
		_getState $w
		if {$last == {}} {
			set last $first
		}
		set first [index $w $first]
		set last [index $w $last]
		if { [catch {
			incr first
			incr last
		} ]} {
			return {}
		} 
		set result {}
		while {$first <= $last} {
			lappend result [$w get $first.0 $first.end]
			incr first
		}
		return $result
	}

	proc selection {w cmd args} {
		_getState $w
		switch -- $cmd {
			clear {
				eval _selectClear $w $args
			}
			includes {
				eval _selectIncludes $w $args			
			}
			set {
				eval _selectSet $w $args	
			}
			anchor {
				eval _selectAnchor $w $args
			}
			default {error "unknown selection command: $cmd"}
		}
	}


	proc _selectAnchor {w index} {
		_getState $w
		set properties($w-anchor) [index $w $index]
	}

	proc _selectClear {w first {last {}} } {
		
		if {$last == {}} {
			set last $first
		}
		set first [index $w $first]
		set last [index $w $last]

		incr first;
		incr last

		while {$first <= $last} {
			$w tag remove ILB_Selection $first.0 [incr first].0
		}
	}

	proc _selectSet {w args} {
		_getState $w
		$w tag remove ILB_Selection 1.0 end

		foreach index $args {
			set index [index $w $index]
			if {$index < [size $w]} {
				$w tag add ILB_Selection [incr index].0 [incr index].0
			}
		}

		if {!$properties($w-uiselection)} {
			event generate $w <<Select>>
		}
	}

	proc _selectIncludes {w first {last {}}} {
		if {$last == {}} {
			set last $first
		}
		set first [index $w $first]
		set last [index $w $last]
		incr first;
		incr last

		while {$first <= $last} {
			$w tag add ILB_Selection $first.0 [incr first].0
		}

		if {!$properties($w-uiselection)} {
			event generate $w <<Select>>
		}
	}

	
	proc curselection {w} {
		_getState $w
		set index 0.0
		set result {}
		while {[set range [$w tag nextrange ILB_Selection $index]] != {}} {
			lappend result [expr int([lindex $range 0]) - 1]
			set index [lindex $range 1]
		}
		return $result
	}

	proc nearest {w y} {
		set index [$w index @0,$y]
		return [expr int($index) - 1]
	}


	proc see {w index} {
		set index [index $w $index]
		if {![string match $index end]} {
			set index [expr $index + 1].0
		}
		$w see $index
	}

	proc index {w index} {
		_getState $w
		if {$index == {}} {
			error "index can't be an empty string"
		}

		switch -regexp -- $index {
		{^(-)?[0-9]+$} {}
		{^@(-)?[0-9]+,(-)?[0-9]+} { return [expr int([$w index $index]) - 1]}
		active	{return $properties($w-active)}
		anchor	{return $properties($w-anchor)}
		end	{return [size $w]}
		default {error "unknown index value: $index"}
		}
		set size [size $w]
		if {$index > $size} {
			set index $size
		} elseif {$index < 0} {
			set index 0
		}
		return $index
	}

	proc activate {w index} {
		_getState $w
		set index [index $w $index]
		set properties($w-active) $index
		return
	}

	proc bbox {w index} {
		_getState $w 
		set index [index $w $index]
		return [$w bbox $index.0]
	}

	proc cget {w option} {
		_getState $w
		switch -- $option {
		-selectmode {return $properties($w-selectmode)}
		default {return [$w cget $option]}
		}
	}

	proc configure {w args} {
		_getState $w
		if {[llength $args]%2 != 0 && [llength $args] != 1} {
			error "configure requires pairs"
		}
		set def {}
		foreach {option value} $args {
			switch -- $option {
				-selectmode {return [_configSelectMode $w $value]}
				default {
					if {[llength $args] == 1} {
						lappend def $option
					} else {
						lappend def $option $value
					}
				}
			}
		}
		if {$def != {}} {
			eval [namespace current]::$w configure $def
		}
	}


	proc _configSelectMode {w value} {
		_getState $w
		if {$value == {}} {
			return $properties($w-selectmode)
		} else {
			if {[regexp {^single|browse|multiple|extended$} $value]} {
				error "invalid select mode: $value"
			}
			set properties($w-selectmode) $value
		}
	}

	proc OnBeginUISelection {w} {
		_getState $w
		set properties($w-uiselection) 1
	}

	proc OnEndUISelection {w} {
		_getState $w
		set properties($w-uiselection) 0
		event generate $w <<Select>>
	}
}

