package require registry
package provide optcl 3.0

namespace eval typelib {
	variable syslibs
	variable syslibguids
	array set syslibs {}
	array set syslibguids {}


	# -----------------------------------------------------------------------------

	# updatelibs -- called to enumerate and store the system libraries
	proc updatelibs {} {

		variable syslibs;
		catch {unset syslibs}
		array set syslibs {}


		set root {HKEY_CLASSES_ROOT\TypeLib}
		foreach id [registry keys $root] {
			catch {
				foreach v [registry keys $root\\$id] {
					scan $v "%d.%d" maj min;
					if [catch {
						set flags [registry get $root\\$id\\$v\\FLAGS {}];
					}] { set flags 0;}

					# check for restricted or hidden libraries
					if {[expr ($flags & 1) || ($flags & 4)]} {
						continue;
					}

					set name "[registry get $root\\$id\\$v {}] (Ver $maj.$min)"
					set syslibs($name) [list $id $maj $min]
				}
			}
		}
	}

	# -----------------------------------------------------------------------------

	# categories -- returns the component categories
	proc categories {} {

		set alldata {}
		set k "HKEY_CLASSES_ROOT\\Component Categories"
		set cats [registry keys $k]

		foreach cat $cats {
			set values [registry values $k\\$cat]
			set data {}
			foreach value $values {
				lappend data [registry get $k\\$cat $value] 
			}
			lappend alldata $data
		}

		return $alldata
	}




	# -----------------------------------------------------------------------------

	# 	libdetail -- returns a the id, maj and min version number
	#		in a list if it exists, else throws an error
	proc libdetail {name} {
		variable syslibs

		if {[array names syslibs $name] == {}} {
			error "could not find the library '$name'"
		}

		return [lindex [array get syslibs $name] 1]
	}


	#------------------------------------------------------------------------------

	# alllibs -- returns all the registered libraries by name
	proc alllibs {} {
		variable syslibs
		return [array names syslibs]
	}

	proc defaultinterface {classtype} {
		set desc [typelib::typeinfo $classtype]
		if {[llength $desc] != 3} {
			error "$classtype is not a class"
		}
		set implintf [lindex $desc 2]
		foreach intf $implintf {
			if {[lsearch -exact [lindex $intf 0] default] >= 0} {
				return [lindex $intf 1]
			}
		}
		error "object doesn't have a default interface"
	}

	#------------------------------------------------------------------------------
	updatelibs

}





if {[info commands tk] != {}} {
	namespace eval tlview {
		catch {font delete tlviewertext}
		catch {font delete tlviewerhigh}
		catch {font delete tlviewerbold}
		font create tlviewertext -family Arial -size 9 -weight normal
		font create tlviewerhigh -family Arial -size 9 -weight bold
		font create tlviewerbold -family Arial -size 9 -weight bold

		variable bgcolor white
		variable textcolor black
		variable highlightcolor blue
		variable selectcolor red
		variable labelcolor red

		array set viewedtypes {}

		#------------------------------------------------------------------------------
		proc scrltxt {w {sb {x y}}} {
			variable bgcolor;
			frame $w -bd 2 -relief sunken;

			text $w.t -bg $bgcolor -bd 0 -relief flat -cursor arrow -width 40 -height 20
			grid $w.t -column 0 -row 0 -sticky nsew;

			if {[lsearch $sb x] >= 0} {
				scrollbar $w.x -orient horizontal -command [list $w.t xview]
				$w.t config -xscrollcommand [list $w.x set] -wrap none
				grid $w.x -column 0 -row 1 -sticky ew;
			}
			if {[lsearch $sb y] >= 0} {
				scrollbar $w.y -orient vertical -command [list $w.t yview]
				$w.t config -yscrollcommand [list $w.y set] 
				grid $w.y -column 1 -row 0 -sticky ns;
			}
			
			grid columnconfigure $w 0 -weight 1;
			grid rowconfigure $w 0 -weight 1;
		}


		#------------------------------------------------------------------------------
		proc cl_list {w} {
			variable bgcolor
			frame $w -bd 2 -relief sunken
			canvas $w.c -yscrollcommand "$w.v set" -xscrollcommand "$w.h set" -bd 0 -relief flat -cursor arrow -bg $bgcolor -highlightthickness 0
			scrollbar $w.h -orient horizontal -command "$w.c xview"
			scrollbar $w.v -orient vertical -command "$w.c yview"

			grid $w.c -column 0 -row 0 -sticky news
			grid $w.h -column 0 -row 1 -sticky ew
			grid $w.v -column 1 -row 0 -sticky ns
			grid columnconfigure $w 0 -weight 1
			grid rowconfigure $w 0 -weight 1
			bind $w.c <1> { focus %W }
			bind $w.c <Prior> { %W yview scroll -1 pages}
			bind $w.c <Next> { %W yview scroll 1 pages}
			return $w
		}



		proc cl_list_update {w} {
			variable ::typelib::syslibs
			variable bgcolor

			if {![winfo exists $w]} {
				error "expected to find a TypeLib list widget: $w"
			}

			set c $w.c
			$c delete all

			foreach tl [lsort [array names ::typelib::syslibs]] {
				cl_list_addlib $w $tl
			}
		}



		proc cl_list_addlib {w tl} {
			variable bgcolor

			set c $w.c
			set bbox [$c bbox entry]
			if {$bbox == {}} {set bbox {0 0 10 10}}
			set bottom [lindex $bbox 3]
			set bottom [expr int($bottom) + 3]
			set tag [$c create text 10 $bottom -anchor nw -fill black -font tlviewertext -justify left -text $tl -tags entry]
			$c bind $tag <1> [namespace code "cl_list_press $w $tag"]
			cl_list_updatetag $w $tag

			set bbox [$c bbox entry]
			set sr [list 0 0 [lindex $bbox 2] [expr $bottom + 20]]
			$c config -scrollregion $sr
		}


		proc cl_list_updatetag {w tag} {
			variable textcolor
			variable highlightcolor

			set c $w.c
			set tl [$c itemcget $tag -text]

			if {![typelib::isloaded $tl]} {
				$c itemconfig $tag -fill $textcolor  -font tlviewertext
			} else {
				$c itemconfig $tag -fill $highlightcolor -font tlviewerhigh
			}
		}


		proc cl_list_press {w tag} {
			set c $w.c
			set tl [$c itemcget $tag -text]
			set parent [winfo parent $w]

			if {![typelib::isloaded $tl]} {
				# loading typelib
				if {[catch {typelib::load $tl} progname]} {
					puts $progname
					$parent.error config -text [string trim $progname]
				} else {
					puts "loaded $progname"
					$parent.error config -text "loaded $progname"
					loadedlibs_updateall
				}
			} else {
				typelib::unload $tl
				puts "unloaded $tl"
				$parent.error config -text "unloaded $tl"
				loadedlibs_updateall
			}

			cl_list_updatetag $w $tag
		}



		proc refview {w} {
			toplevel $w
			wm title $w "Referenced Type Libraries"
			bind $w <Alt-F4> "destroy $w"
			bind $w <Alt-c> "$w.close invoke"
			bind $w <Alt-r> "$w.refresh config -relief sunken; update; $w.refresh invoke; $w.refresh config -relief raised"
			button $w.close -text Close -width 7 -command "destroy $w" -underline 0
			button $w.refresh -text Refresh -width 7 -command [namespace code "cl_list_update $w.list"] -underline 0
			label $w.error -bd 1 -relief sunken

			grid [cl_list $w.list] -column 0 -row 0 -columnspan 2 -sticky nsew
			grid $w.close -column 0 -row 1 -padx 5 -pady 5
			grid $w.refresh -column 1 -row 1 -padx 5 -pady 5
			grid $w.error -column 0 -row 2 -columnspan 2 -sticky nsew

			grid columnconfig $w 0 -weight 1
			grid columnconfig $w 1 -weight 1
			grid rowconfig $w 0 -weight 1
			
			cl_list_update $w.list
			return $w
		}


		
		#------------------------------------------------------------------------------

		proc loadedlibs_updateall {} {
			foreach w [winfo child .] {
				if {[string compare [winfo class $w] TLLoadedTypeLibs] == 0} {
					loadedlibs_update $w
				}
			}
		}

		proc loadedlibs_update {w} {
			variable bgcolor 
			variable textcolor 
			variable highlightcolor 

			$w.l.t config -state normal	
			$w.l.t delete 1.0 end
			foreach lib [lsort [typelib::loaded]] {
				$w.l.t tag configure tag$lib -foreground $highlightcolor -font tlviewertext -underline 0 
				$w.l.t insert end "$lib\n" tag$lib
				$w.l.t tag bind tag$lib <1> [namespace code "viewlib $lib"]
				$w.l.t tag bind tag$lib <Enter> "$w.l.t config -cursor hand2; $w.l.t tag config tag$lib -underline 1"
				$w.l.t tag bind tag$lib <Leave> "$w.l.t config -cursor arrow; $w.l.t tag config tag$lib -underline 0"
			}
			$w.l.t config -state disabled
		}

		proc loadedlibs {w} {
			toplevel $w -class TLLoadedTypeLibs

			wm title $w "Loaded Libraries"
			scrltxt $w.l

			grid $w.l -column 0 -row 0 -sticky nsew 
			grid columnconfig $w 0 -weight 1
			grid rowconfig $w 0 -weight 1
			loadedlibs_update $w
			bind $w <FocusIn> [namespace code "loadedlibs_update $w"]
		}

		#------------------------------------------------------------------------------
		proc viewlib_onenter {txt tag} {
			$txt config -cursor hand2
			 $txt tag config $tag -underline 1
		}

		proc viewlib_onleave {txt tag} {
			$txt config -cursor arrow
			$txt tag config $tag -underline 0
		}

		proc viewlib_unselect {txt lib} {
			variable viewedtypes
			variable textcolor 
			if {[array name viewedtypes $lib] != {}} {
				set type $viewedtypes($lib)
				$txt tag config tag$type -foreground $textcolor -font tlviewertext
				set viewedtypes($lib) {}
			}
		}



		proc viewlib_select {fulltype } {
			variable viewedtypes
			variable highlightcolor

			puts $fulltype
			set sp [split $fulltype .]
			if {[llength $sp] != 2} {
				return
			}
			
			set lib [lindex $sp 0]
			set type [lindex $sp 1]

			set w [viewlib $lib]
			set txt $w.types.t

			viewlib_unselect $txt $lib
			$txt tag config tag$type -foreground $highlightcolor -font tlviewerhigh

			$txt see [lindex [$txt tag ranges tag$type] 0]
			set viewedtypes($lib) $type
			viewlib_readelems $w $lib $type;
		}
		
		
		proc viewlib_selectelem {w fulltype element} {
			variable viewedtypes
			variable highlightcolor

			puts "$fulltype $element"
			set sp [split $fulltype .]
			set lib [lindex $sp 0]
			set type [lindex $sp 1]

			set txt $w.elems.t

			viewlib_unselect $txt $fulltype
			$txt tag config tag$element -foreground $highlightcolor -font tlviewerhigh
			$txt see [lindex [$txt tag ranges tag$element] 0]
			set viewedtypes($fulltype) $element
			viewlib_readdesc $w $lib $type $element
		}

		###
		# creates a list of types in some library
		proc viewlib_readtypes {w lib} {
			variable textcolor
			set txt $w.types.t

			$txt config -state normal
			$txt del 1.0 end
			
			foreach tdesc [lsort [typelib::types $lib]] {
				$txt insert end "[lindex $tdesc 0]\t"
				set full [lindex $tdesc 1]
				set type [lindex [split $full .] 1] 
				$txt tag configure tag$type -foreground $textcolor -font tlviewertext -underline 0
				$txt insert end "$type\n" tag$type
				$txt tag bind tag$type <1> [namespace code "
											viewlib_select $full;
											"]

				$txt tag bind tag$type <Enter> [namespace code "viewlib_onenter $txt tag$type"]
				$txt tag bind tag$type <Leave> [namespace code "viewlib_onleave $txt tag$type"]
			}
			$txt config -state disabled
		}


		proc viewlib_writetype {txt fulltype} {
			variable highlightcolor
			if {[llength [split $fulltype .]] > 1} {
				$txt tag configure tag$fulltype -foreground $highlightcolor -font tlviewertext -underline 0
				$txt tag bind tag$fulltype <Enter> [namespace code "viewlib_onenter $txt tag$fulltype"]
				$txt tag bind tag$fulltype <Leave> [namespace code "viewlib_onleave $txt tag$fulltype"]
				$txt tag bind tag$fulltype <1> [namespace code "viewlib_select $fulltype"]
				$txt insert end "$fulltype" tag$fulltype
			} else {
				$txt insert end "$fulltype"
			}
		}


		###
		# displays the elements for a type of some library
		proc viewlib_readelems {w lib type} {
			variable labelcolor 
			variable textcolor
			variable highlightcolor

			set txt $w.elems.t
			$txt config -state normal
			$txt del 1.0 end
			set elems [typelib::typeinfo $lib.$type]
			loadedlibs_updateall

			$txt tag configure label -font tlviewerhigh -underline 1 -foreground $labelcolor

			if {[string compare "typedef" [lindex $elems 0]] == 0} {
				# --- we are working with a typedef
				set t [lindex $elems 3]
				$txt insert end "Typedef\n\t" label
				viewlib_writetype $txt $t
			} else {
				if {[llength [lindex $elems 1]] != 0} {
					$txt insert end "Methods\n" label
				}

				foreach method [lsort [lindex $elems 1]] {
					$txt tag configure tag$method -foreground $textcolor -font tlviewertext -underline 0
					$txt tag bind tag$method <Enter> [namespace code "viewlib_onenter $txt tag$method"]
					$txt tag bind tag$method <Leave> [namespace code "viewlib_onleave $txt tag$method"]
					$txt tag bind tag$method <1> [namespace code "viewlib_selectelem $w $lib.$type $method"]
					$txt insert end "\t$method\n" tag$method
				}

				if {[llength [lindex $elems 2]] != 0} {
					$txt insert end "Properties\n" label
				}

				foreach prop [lsort [lindex $elems 2]] {
					$txt tag configure tag$prop -foreground $textcolor -font tlviewertext -underline 0
					$txt tag bind tag$prop <Enter> [namespace code "viewlib_onenter $txt tag$prop"]
					$txt tag bind tag$prop <Leave> [namespace code "viewlib_onleave $txt tag$prop"]
					$txt tag bind tag$prop <1> [namespace code "viewlib_selectelem $w $lib.$type $prop"]
					$txt insert end "\t$prop\n" tag$prop
				}

				if {[llength [lindex $elems 3]] != 0} {
					$txt insert end "Inherited Types\n" label
				}

				foreach impl [lsort -index 1 [lindex $elems 3]] {
				# implemented interfaces
					set t [lindex $impl 1]
					set flags [lindex $impl 0]
					if {[lsearch -exact $flags default] != -1} {
						$txt insert end "*"
					}

					if {[lsearch -exac $flags source] != -1} {
						$txt insert end "event"
					}
					$txt insert end "\t"

					$txt tag configure itag$t -foreground $highlightcolor -font tlviewertext -underline 0
					$txt tag bind itag$t <Enter> [namespace code "viewlib_onenter $txt itag$t"]
					$txt tag bind itag$t <Leave> [namespace code "viewlib_onleave $txt itag$t"]
					$txt tag bind itag$t <1> [namespace code "viewlib_select $t"]
					
					$txt insert end "$t\n" itag$t
				}
			}
			$txt config -state disabled
			viewlib_settypedoc $w [lindex $elems 4]
		}

		
		proc viewlib_settypedoc {w doc} {
			set txt $w.desc.t
			$txt config -state normal
			$txt delete 1.0 end
			$txt insert end $doc
			$txt config -state disabled
		}


		###
		# retrieves the description for an element
		proc viewlib_readdesc {w lib type elem} {
			variable labelcolor

			set txt $w.desc.t
			$txt config -state normal
			$txt delete 1.0 end

			$txt tag configure label -font tlviewerhigh -underline 1 -foreground $labelcolor
			$txt tag configure element -font tlviewerbold
			$txt tag bind element <Enter> [namespace code "viewlib_onenter $txt element"]
			$txt tag bind element <Leave> [namespace code "viewlib_onleave $txt element"]
			
			$txt tag bind element <1> [namespace code "viewlib_select $lib.$type; viewlib_selectelem $w $lib.$type $elem"]

			set desc [typelib::typeinfo $lib.$type $elem]
			set kind [lindex $desc 0]
			switch $kind {
				property {
					$txt insert end "Property" label
					$txt insert end "\t[lindex $desc 2]\n"

					set p [lindex $desc 1]
					# insert the flags
					$txt insert end "[lindex $p 0]\t"
					viewlib_writetype $txt [lindex $p 1]
					$txt insert end "  "
					$txt insert end "[lindex $p 2]" element

					set params [lrange $p 3 end]

					foreach param $params {
						$txt insert end "\n\t"

						if {[llength $param] == 3} {
							$txt insert end "[lindex $param 0]\t"
							set param [lrange $param 1 end]
						}
						viewlib_writetype $txt [lindex $param 0]
						$txt insert end "  [lrange $param 1 end]"
					}
				}

				method	{
					set m [lindex $desc 1]
					$txt insert end "Method" label
					$txt insert end "\t[lindex $desc 2]\n"
					viewlib_writetype $txt [lindex $m 0]
					$txt insert end "  "
					$txt insert end "[lindex $m 1]" element
					set params [lrange $m 2 end]

					foreach param $params {
						$txt insert end "\n\t"

						if {[llength $param] == 3} {
							$txt insert end "[lindex $param 0]\t"
							set param [lrange $param 1 end]
						}
						viewlib_writetype $txt [lindex $param 0]
						$txt insert end "  [lrange $param 1 end]"
					}
				}
			}

			puts [lindex $desc 1]
			$txt config -state disabled
		}


		####
		# Creates a viewer for library
		proc viewlib {lib} {
			set w ._tlview$lib
			if [winfo exists $w] {
				raise $w
				return $w
			}
			toplevel $w -class tlview_$lib
			wm title $w "Type Library: $lib"

			label $w.tl -text Types;
			label $w.el -text Elements;
			label $w.dl -text Description;

			scrltxt $w.types;
			scrltxt $w.elems
			scrltxt $w.desc y
			$w.desc.t config -height 5
			$w.desc.t config -state disabled
			$w.elems.t config -state disabled
			$w.types.t config -state disabled

			grid $w.tl -column 0 -row 0 -sticky nw
			grid $w.types -column 0 -row 1 -sticky news -padx 2 -pady 2
			grid $w.el -column 1 -row 0 -sticky nw
			grid $w.elems -column 1 -row 1 -sticky news -padx 2 -pady 2
			grid $w.dl -column 0 -row 2 -sticky nw
			grid $w.desc -column 0 -row 3 -columnspan 2 -sticky news -padx 2 -pady 2

			grid columnconfigure $w 0 -weight 1
			grid columnconfigure $w 1 -weight 1
			grid rowconfigure $w 1 -weight 1
			#grid rowconfigure $w 3 -weight 1

			viewlib_readtypes $w $lib
			return $w
		}


		proc viewtype {fullname} {
			viewlib_select $fullname
		}
	}
}