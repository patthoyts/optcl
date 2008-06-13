#--------------------------------------------------------------------
# File: TLView.tcl
#	Implements the GUI for Type Library management
# Author:	Farzad Pezeshkpour fuzz@sys.uea.ac.uk
# Date:		May 25th 2000
#--------------------------------------------------------------------

image create bitmap ::tlview::downarrow_img -data {
	#define down_arrow_width 12
	#define down_arrow_height 12
	static char down_arrow_bits[] = {
		0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
		0xfc,0xf1,0xf8,0xf0,0x70,0xf0,0x20,0xf0,
		0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00;
	}
}
image create bitmap ::tlview::leftarrow_img -data {
	#define left_width 12
	#define left_height 12
	static unsigned char down_bits[] = {
	   0x00, 0x00, 0x00, 0x00, 0x80, 0x00, 0xc0, 0x00, 0xe0, 0x00, 0xf0, 0x00,
	   0xe0, 0x00, 0xc0, 0x00, 0x80, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 };
}

image create bitmap ::tlview::rightarrow_img -data {
	#define right_width 12
	#define right_height 12
	static unsigned char right_bits[] = {
	   0x00, 0x00, 0x00, 0x00, 0x10, 0x00, 0x30, 0x00, 0x70, 0x00, 0xf0, 0x00,
	   0x70, 0x00, 0x30, 0x00, 0x10, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 };
}

image create bitmap ::tlview::find_img -data {
	#define find_width 17
	#define find_height 17
	static char find_bits[] = {
		0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x70, 0x1c, 0x00, 0x50, 0x1c, 0x00,
		0x50, 0x1c, 0x00, 0x70, 0x1c, 0x00, 0xf8, 0x3e, 0x00, 0xe8, 0x3a, 0x00,
		0xfc, 0x7f, 0x00, 0x7e, 0xfb, 0x00, 0x7e, 0xfb, 0x00, 0xfe, 0xfb, 0x00,
		0xfe, 0xfe, 0x00, 0x3a, 0xe8, 0x00, 0x3a, 0xe8, 0x00, 0x3e, 0xf8, 0x00,
		0x3e, 0xf8, 0x00, 0x00, 0x00, 0x00;
	}
}

image create bitmap ::tlview::show_img -data {
	#define show_width 17
	#define show_height 17
	static char show_bits[] = {
		0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
		0x00, 0x00, 0x00, 0x10, 0x08, 0x00, 0x20, 0x04, 0x00, 0x40, 0x02, 0x00,
		0x80, 0x01, 0x00, 0x10, 0x08, 0x00, 0x20, 0x04, 0x00, 0x40, 0x02, 0x00,
		0x80, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
		0x00, 0x00, 0x00 
	};
}
image create bitmap ::tlview::hide_img -data {
	#define hide_width 17
	#define hide_height 17
	static char hide_bits[] = {
	   0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
	   0x80, 0x01, 0x00, 0x40, 0x02, 0x00, 0x20, 0x04, 0x00, 0x10, 0x08, 0x00,
	   0x80, 0x01, 0x00, 0x40, 0x02, 0x00, 0x20, 0x04, 0x00, 0x10, 0x08, 0x00,
	   0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
	   0x00, 0x00, 0x00
   };
}

image create photo ::tlview::class_img -file [file join $env(OPTCL_LIBRARY) images class.gif]
image create photo ::tlview::interface_img -file [file join $env(OPTCL_LIBRARY) images interface.gif]
image create photo ::tlview::dispatch_img -file [file join $env(OPTCL_LIBRARY) images dispatch.gif]
image create photo ::tlview::module_img -file [file join $env(OPTCL_LIBRARY) images module.gif]
image create photo ::tlview::struct_img -file [file join $env(OPTCL_LIBRARY) images struct.gif]
image create photo ::tlview::union_img -file [file join $env(OPTCL_LIBRARY) images union.gif]
image create photo ::tlview::enum_img -file [file join $env(OPTCL_LIBRARY) images enum.gif]
image create photo ::tlview::typedef_img -file [file join $env(OPTCL_LIBRARY) images typedef.gif]
image create photo ::tlview::property_img -file [file join $env(OPTCL_LIBRARY) images property.gif]
image create photo ::tlview::method_img -file [file join $env(OPTCL_LIBRARY) images method.gif]
image create photo ::tlview::copy_img -file [file join $env(OPTCL_LIBRARY) images copy.gif]
image create photo ::tlview::select_img -file [file join $env(OPTCL_LIBRARY) images select.gif]
image create photo ::tlview::noselect_img -file [file join $env(OPTCL_LIBRARY) images noselect.gif]
image create photo ::tlview::libselect_img -file [file join $env(OPTCL_LIBRARY) images libselect.gif]


namespace eval tlview {
	namespace import -force ::optcl_utils::*

	catch {font delete tlviewertext}
	catch {font delete tlviewerhigh}
	catch {font delete tlviewerbold}
	catch {font delete tlvieweritalic}

	font create tlviewertext -family Arial -size 9
	font create tlviewerhigh -family Arial -size 9 -weight bold
	font create tlviewerbold -family Arial -size 9 -weight bold
	font create tlvieweritalic -family Arial -size 9 -slant italic

	variable colors
	
	set colors(bgcolor) SystemWindow
	set colors(textcolor) SystemWindowText
	set colors(highlightcolor) blue
	set colors(selectcolor) red
	set colors(labelcolor) brown

	variable navigation_history
	variable search_history
	variable properties
	variable libs

	array set viewedtypes {}
	array set navigation_history {}
	array set search_history {}
	array set properties {}
	array set libs {}


	proc TlviewScrolledListBox {w args} {
		return [eval TlviewScrolledListBox::create $w $args]
	}




	namespace eval TlviewScrolledListBox {
		proc create {w args} {
			frame $w
			label $w.label -text {}
			scrollbar $w.h -orient horizontal -command [namespace code "$w.listbox xview"]
			scrollbar $w.v -orient vertical -command [namespace code "$w.listbox yview"]
			ImageListBox $w.listbox -yscrollcommand "$w.v set" -xscrollcommand "$w.h set"
			grid $w.label -col 0 -row 0 -sticky nw
			grid $w.listbox -col 0 -row 1 -sticky nsew
			grid $w.v -col 1 -row 1 -sticky ns
			grid $w.h -col 0 -row 2 -sticky ew
			grid columnconfigure $w 0 -weight 1
			grid rowconfigure $w 1 -weight 1
			return $w
		}
	}


	namespace eval TlviewPopup {
		variable popupstate
		
		proc Create {w} {
			toplevel $w
			wm withdraw $w
			wm overrideredirect $w 1
			return $w
		}

		proc Show {w x y} {
			
			set parent [winfo parent $w]
			wm geometry $w [join [list +$x $y] +]
			wm deiconify $w
			wm transient $w [winfo toplevel $w]

			#raise $w $parent
			#focus -force $w
			bind $w <ButtonPress-1> [namespace code "Click $w %W"]
			bind $w <KeyPress-Escape> "destroy $w"
			bind $w <FocusOut> "destroy $w"
			grab $w
			return $w
		}
		
		proc Click {pw w} {
			if	{[string compare $pw $w] == 0} {
				destroy $pw
			}
		}
	}




	namespace eval ScrolledList {
		proc Create {w {callback {}}} {
			frame $w -borderwidth 0 -highlightthickness 1 -highlightbackground black
			listbox $w.lb -yscrollcommand "$w.s set" -borderwidth 0 -relief flat -background SystemWindow -fg SystemWindowText -highlightthickness 0 -height 5
			bind $w.lb <ButtonRelease-1> [namespace code "OnSelect $w.lb $callback"]
			scrollbar $w.s -orient vertical -command "$w.lb yview"
			pack $w.s -side right -fill y
			pack $w.lb -side left -fill both -expand 1
			return $w
		}
		
		proc OnSelect {w args} {
			set cs [$w curselection]
			if {$cs != {}} {
				if {$args != {}} {
					eval $args [$w get $cs]
				}
					after 50 "destroy [winfo toplevel $w]"
			}
		}
	}




	namespace eval TlViewDropList {
		proc Create {w contents callback} {
			::tlview::TlviewPopup::Create $w
			set sl [::tlview::ScrolledList::Create $w.sl $callback]
			eval $sl.lb insert end $contents
			pack $sl -fill both -expand 1
		}

		proc Show {w x y} {
			::tlview::TlviewPopup::Show $w $x $y
		}
	}

	namespace eval TlviewCombo {
		proc Ondropdown {w contentsfn} {
			set x [winfo rootx $w]
			set y [expr [winfo rooty $w] + [winfo height $w]]
			set contents [eval $contentsfn]
			::tlview::TlViewDropList::Create $w.dropdown $contents [namespace code "Callback $w"]

			set height [winfo reqheight $w.dropdown.sl.lb]
			set width [winfo width $w]
			wm geometry $w.dropdown [join "$width $height" x]
			if {[expr $y + $height + 20] >= [winfo screenheight $w.dropdown]} {
				set y [expr [winfo rooty $w] - $height]
			}
			::tlview::TlViewDropList::Show $w.dropdown $x $y
		}

		proc Create {w contents args} {
			frame $w -borderwidth 2 -relief sunken -height 10
			eval entry $w.e -borderwidth 0 -relief flat -font {{arial 8}} $args
			label $w.b -image ::tlview::downarrow_img -font {arial 8} -borderwidth 1 -relief raised
			pack $w.b -side right -fill y
			pack $w.e -side left -fill both -expand 1

			bind $w.b <ButtonPress-1> [namespace code [list Ondropdown $w $contents]]
			bind $w.e <KeyPress-Down> [namespace code [list Ondropdown $w $contents]]
			return $w
		}


		proc Callback {w value} {
			set state [$w.e cget -state]
			$w.e config -state normal
			$w.e delete 0 end
			$w.e insert end $value
			$w.e config -state $state
		}
	}




	

	
	proc search_history_add {w item} {
		variable search_history
		set pairs [array get search_history $w]
		if {$pairs == {}} {
			set search_history($w) $item
			return
		}

		set searches [lindex $pairs 1]
		set found [lsearch $searches $item]
		if {$found >= 0} {
			set searches [concat $item [lreplace $searches $found $found]]
		} else {
			set searches [concat $item $searches]
		}
		set search_history($w) $searches
	}

	proc search_history_getlist {w} {
		variable search_history
		set pairs [array get search_history $w]
		if {$pairs == {}} {
			return ""
		} else {
			return [lindex $pairs 1]
		}
	}

	proc history_init {w} {
		variable navigation_history
		
		set navigation_history($w) {}
		set navigation_history($w-index) -1
		set navigation_history($w-changing) 0
	}

	proc history_erase {w} {
		variable navigation_history
		array unset navigation_history $w
		array unset navigation_history $w-*
	}

	proc history_add {w lib {type {}} {elem {}}} {
		variable navigation_history

		if {[history_locked? $w]} return

		set index $navigation_history($w-index)
		set history [lrange $navigation_history($w) 0 $index]


		# check that this isn't already the current item
		set lastitem [lindex $history end]
		if { [string match [lindex $lastitem 0] $lib] &&
			 [string match [lindex $lastitem 1] $type] &&
			 [string match [lindex $lastitem 2] $elem]} {
			
			# we'll just quietly return
			return
		}

		lappend history [list $lib $type $elem]
		set navigation_history($w) $history
		incr navigation_history($w-index)
		viewlib_updatenav $w
	}

	proc history_addwindowstate {w} {
		variable properties
		history_add $w $properties($w-viewedlibrary) $properties($w-type) $properties($w-element)
	}
			
	proc history_back {w} {
		variable navigation_history
		if {![history_back? $w]} return

		incr navigation_history($w-index) -1
		history_current $w
	}

	proc history_forward {w} {
		variable navigation_history
		if {![history_forward? $w]} return
		incr navigation_history($w-index)
		history_current $w
	}

	proc history_current {w} {
		variable navigation_history
		history_lock $w
		set item [lindex $navigation_history($w) $navigation_history($w-index)]
		tlview::viewlib_select $w [lindex $item 0] [lindex $item 1] [lindex $item 2]
		viewlib_updatenav $w
		history_unlock $w
	}

	proc history_last {w} {
		variable navigation_history
		set navigation_history($w-index) [expr [llength navigation_history($w) - 1]
		history_current $w
	}

	proc history_back? {w} {
		variable navigation_history
		return [expr $navigation_history($w-index) > 0] 
	}

	proc history_forward? {w} {
		variable navigation_history
		return [expr $navigation_history($w-index) < ([llength $navigation_history($w)] - 1)]
	}
	
	proc history_lock {w} {
		variable navigation_history
		incr navigation_history($w-changing)
	}

	proc history_unlock {w} {
		variable navigation_history
		if {[incr navigation_history($w-changing) -1] < 0} {
			set navigation_history($w-changing) 0
		}
	}

	proc history_locked? {w} {
		variable navigation_history
		return [expr $navigation_history($w-changing) > 0]
	}

	proc history_clean {} {
		variable navigation_history
		set loadedlibs [typelib::loadedlibs]

		foreach item [array names navigation_history] {
			if {[regexp -- - $item]} continue
			
			if {![winfo exists $item]} {
				history_erase $item
				continue
			}

			set newhistory {}
			foreach history_item $navigation_history($item) {
				set lib [lindex $history_item 0]
				if {[lsearch -exact $loadedlibs $lib] >= 0} {
					lappend newhistory $history_item
				}
			}
			set navigation_history($item) $newhistory
			if {$navigation_history($item-index) >= [llength $newhistory]} {
				set navigation_history($item-index) [expr [llength $newhistory] - 1]
			}
			set history_item [lindex $newhistory $navigation_history($item-index)]
			history_lock $item
			viewlib_select $item [lindex $history_item 0] [lindex $history_item 1] [lindex $history_item 2] false
			history_unlock $item
		}
	}

	#------------------------------------------------------------------------------
	proc scrltxt {w {sb {x y}}} {
		variable colors

		frame $w -bd 2 -relief sunken;

		text $w.t -bg $colors(bgcolor) -bd 0 -relief flat -cursor arrow -width 40 -height 20
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
		return $w
	}





	proc libs {w} {
		toplevel $w -class TLLoadedTypeLibs -height 250 -width 450
		wm transient $w [winfo parent $w]
		wm title $w "Referenced Type Libraries"

		grab set $w

		set closecmd "destroy $w"
		set browsecmd "[namespace current]::libs_browse $w"
		set savecmd {puts [typelib::persistLoaded]}
		button $w.close -text Close -padx 5 -underline 0 -width 10 -command $closecmd
		place $w.close -anchor ne -relx 1.0 -x -10 -y 10

		button $w.browse -text "Browse ..." -padx 5 -underline 0 -width 10 -command $browsecmd
		place $w.browse -anchor ne -relx 1.0 -x -10 -y 50

		button $w.save -text "Save" -padx 5 -underline 0 -width 10 -command $savecmd 
		place $w.save -anchor ne -relx 1.0 -x -10 -y 90


		set liblist [::tlview::TlviewScrolledListBox $w.liblist]
		$liblist config -bd 2 -relief sunken
		
		destroy $liblist.label
		
		label $w.label1 -text "Available Libraries:"
		place $w.label1 -anchor nw -x 10 -y 10
		

		set lb $liblist.listbox
		$lb config -spacing1 3 -takefocus 1
		place $w.liblist -anchor nw -x 10 -y 30 -relheight 1.0 -height -100 -relwidth 1.0 -width -120
		
		frame $w.details -bd 2 -relief groove
		place $w.details -anchor sw -x 10 -rely 1.0 -y -10 -height 50 -relwidth 1.0 -width -20
		label $w.details_label -text {Lib Details}
		place $w.details_label -anchor nw -x 15 -rely 1.0 -y [expr -60 - [winfo reqheight $w.details_label] / 2]
		
		label $w.details.path -text "Path: " -justify left -font {Arial 8 {}}
		place $w.details.path -anchor nw -x 5 -y 10
		
		bind $lb <Double-ButtonPress-1> [namespace code "libs_ondblselect $w %x %y; break;"]
		bind $lb <<Select>> [namespace code "libs_onselect $w"]

		bind $lb <B1-Motion> {break;}
		bind $lb <KeyPress> [namespace code "libs_onkeypress $w %K; continue"]
		
		bind $lb <KeyPress-space> [namespace code "libs_onenterkey $w"]
		
		
		bind $lb <Alt-c> "$closecmd; break"
		bind $lb <Alt-b> $browsecmd
		bind $lb <Alt-s> $savecmd

		bind $lb <Escape> $closecmd
		bind $lb <KeyPress-Return> $closecmd
		bind $w <Destroy> [namespace code "libs_close %W $w"]
		
		focus $lb
		$lb activate 0
		$lb selection set 0

		wm geometry $w 600x300
		wm minsize $w 600 300
		optcl_utils::center_window $w
		libs_update $w

		return $w
	}

	proc libs_close {W w} {
		if {[string match $w $W]} {
			viewlib_clean
		}
	}

	proc libs_update {w} {
		variable libs
		set libs($w) {}

		set loaded {}
		set notloaded {}
		set loadednames {}
		set notloadednames {}
		$w.liblist.listbox delete 0 end

		foreach lib [array names ::typelib::typelibraries] {
			mset {guid maj min} $lib
			mset {name path} $typelib::typelibraries($lib)
			if {[typelib::isloaded $guid $maj $min] != {}} {
				lappend loaded [list $name $guid $maj $min $path]
			} else {
				lappend notloaded [list $name $guid $maj $min $path]
			}
		}

		foreach lib [lsort -dictionary -index 0 $loaded ] {
			$w.liblist.listbox insert end [lindex $lib 0]
			$w.liblist.listbox setimage end ::tlview::select_img
			lappend libs($w) $lib
		}

		foreach lib [lsort -dictionary -index 0 $notloaded ] {
			$w.liblist.listbox insert end [lindex $lib 0]
			$w.liblist.listbox setimage end ::tlview::noselect_img
			lappend libs($w) $lib
		}
		$w.liblist.listbox selection set 0
		libs_onselect $w
	}
	
	proc libs_onkeypress {w key} {
		set lb $w.liblist.listbox
		set key [string tolower $key]
		if {![regexp {^[a-z]$} $key]} return
		
		set currentindex [$lb index active]

		set liblist [$lb get 0 end]
		set searchlist [concat [lrange $liblist [expr $currentindex + 1] end] [lrange $liblist 0 $currentindex]]
		
		set nextindex [lsearch -regexp $searchlist ^($key|[string toupper $key]).*$]
		
		if {$nextindex>=0} {
			if {$nextindex < [expr [llength $liblist] - $currentindex - 1]} {
				set nextindex [expr $nextindex + $currentindex + 1]
			} else {
				set nextindex [expr $nextindex - ([llength $liblist] - $currentindex) + 1]
			}
			$lb selection clear 0 end
			$lb activate $nextindex
			$lb selection set $nextindex
			$lb see active
		}
	}

	proc libs_onenterkey {w} {
		set lb $w.liblist.listbox
		set index [lindex [$lb curselection] 0]
		libs_loader $w $index
	}

	proc libs_loader {w index} {
		variable libs
		set lib [lindex $libs($w) $index]
		mset {name guid maj min path} $lib
		
		set lb $w.liblist.listbox
		set progname [typelib::isloaded $guid $maj $min]
		if {$progname != {}} {
			typelib::unload $progname
			$lb setimage $index ::tlview::noselect_img
		} else {
			if {[catch {typelib::load $path} e]} {
				tk_messageBox -title "Type Library Error" -message $e -icon error -parent $w -type ok
			} else {
				$lb setimage $index ::tlview::select_img
			}
		}
		$lb activate $index
		$lb selection clear 0 end
		$lb selection set $index
	}

	proc libs_ondblselect {w x y} {
		set lb $w.liblist.listbox
		set index [$lb index @$x,$y]
		if {$index == {}} return
		libs_loader $w $index
	}

	proc libs_onselect {w} {
		variable libs
		set lb $w.liblist.listbox
		set index [$lb curselection]

		set lib [lindex $libs($w) $index]
		mset {name guid maj min path} $lib
		$w.details_label config -text $name
		$w.details.path config -text "Path: $path"
	}



	proc libs_browse {w} {
		variable libs

		set types {
			{{Type Libraries} {.tlb .olb .dll}}
			{{Executable Files} {.exe .dll}}
			{{ActiveX Controls} {.ocx}}
			{{All Files} {*}}
		}
		set fname [tk_getOpenFile -filetypes $types -parent $w -title "Add Type Library Reference"]
		if {$fname != {}} {
			try {
				set progname [typelib::load $fname]
				mset {guid maj min path fullname} [typelib::loadedlib_details $progname]
				libs_update $w
				set index [lsearch -exact $libs($w) [list $fullname $guid $maj $min $path]]
				puts $index
				if {$index >= 0} {
					after 50 "$w.liblist.listbox selection set $index; $w.liblist.listbox see $index"
				} 
			} catch {er} {
				tk_messageBox -title "Error in loading library" -message $er -type ok -icon error
			}
		}
	}


	#------------------------------------------------------------------------------


	proc viewlib_clean {} {
		variable properties

		history_clean

		set loadedlibs [typelib::loadedlibs]
		
		foreach item [array names properties] {
			
			if {[regexp -- - $item]} continue
			if {![winfo exists $item]} {
				array unset properties $item
				array unset properties $item-*
				continue
			}

			if {[lsearch -exact $loadedlibs $properites($item-viewedlibrary)] < 0} {
				# library is not loaded ... so revert to the current index of the history
				history_current $item
			}
		}

	}

	proc viewlib_onenter {txt tag} {
		$txt config -cursor hand2
		 $txt tag configure $tag -underline 1
	}

	proc viewlib_onleave {txt tag} {
		$txt config -cursor arrow
		$txt tag configure $tag -underline 0
	}

	
	proc viewlib_updatenav {w} {
		set topbar $w.topbar
		if {[history_back? $w]} {
			$topbar.back config -state normal
		} else {
			$topbar.back config -state disabled
		}

		if {[history_forward? $w]} {
			$topbar.forward config -state normal
		} else {
			$topbar.forward config -state disabled
		}
	}
	


	proc viewlib_select {w lib {type {}} {element {}} {raise true}} {
		variable properties

		history_lock $w
		if {![string match $properties($w-viewedlibrary) $lib]} {
			# try to find a window that is already viewing this library
			foreach tlviewer [array names properties *-viewedlibrary] {
				if {[string match $properties($tlviewer) $lib]} {
					set w [lindex [split $tlviewer -] 0]; break
				}
			}
		}

		set types_lb $w.sp1.sp2.types.listbox
		set elements_lb $w.sp1.sp2.elements.listbox
		
		
		# raise the window and instruct it to view the library
		if {$raise} {
			raise $w
		}

		viewlib_showlibrary $w $lib
		

		if {$type != {}} {
			# now find the type
			set index [lsearch -exact [$types_lb get 0 end] $type]
			if {$index >= 0} {
				$types_lb selection set $index
				$types_lb see $index
			}
		}

		if {$type != {} && $element != {}} {
			set index [lsearch -regexp [$elements_lb get 0 end] "^($element)( .+)?"]
			if {$index >= 0} {
				$elements_lb selection set $index
				$elements_lb see $index
			}
		}
		history_unlock $w
		history_addwindowstate $w
	}

	# browse to a specific library
	proc viewlib_showlibrary {w lib} {
		variable colors
		variable properties

		set types_lb	$w.sp1.sp2.types.listbox
		set elements_lb	$w.sp1.sp2.elements.listbox
		set description	$w.sp1.description.desc.t


		$elements_lb delete 0 end
		#$description config -state normal
		$description delete 1.0 end
		#$description config -state disabled

		# if the viewed library is being changed, redirect through the 
		# $w-library property change event handler
		if {![string match $properties($w-viewedlibrary) $lib] && $lib != {}} {
			set properties($w-viewedlibrary) $lib
			return
		} elseif {[string match $properties($w-library) $lib] || $lib == {}} {
			if {$properties($w-type) != {}} {
				set properties($w-type) {}
				$types_lb selection clear 0 end
			}
			return
		}

		set properties($w-library) $lib
		set properties($w-type) {}
		set properties($w-element) {}

		history_addwindowstate $w
					
		$types_lb delete 0 end
		foreach tdesc [lsort [typelib::types $lib]] {
			set typetype [lindex $tdesc 0]
			set full [lindex $tdesc 1]
			set type [lindex [split $full .] 1] 
			$types_lb insert end $type
			
			switch -- $typetype {
				class {$types_lb setimage end ::tlview::class_img}
				dispatch {$types_lb setimage end ::tlview::dispatch_img}
				interface {$types_lb setimage end ::tlview::interface_img}
				module {$types_lb setimage end ::tlview::module_img}
				struct {$types_lb setimage end ::tlview::struct_img}
				union  {$types_lb setimage end ::tlview::union_img}
				enum  {$types_lb setimage end ::tlview::enum_img}
				typedef {$types_lb setimage end ::tlview::typedef_img}

			}

		}
		bind $types_lb <<Select>> [namespace code "viewlib_showelements_byindex $w $lib \[$types_lb curselection\]"]
		wm title $w "Type Library: $lib"
	}

	

	proc viewlib_writetype {txt fulltype} {
		variable colors
		set split [split $fulltype .]
		set lib [lindex $split 0]
		set type [lindex $split 1]
		set element [lindex $split 2]

		if {[llength [split $fulltype .]] > 1} {
			$txt tag configure tag$fulltype -foreground $colors(highlightcolor) -font tlviewertext -underline 0
			$txt tag bind tag$fulltype <Enter> [namespace code "viewlib_onenter $txt tag$fulltype"]
			$txt tag bind tag$fulltype <Leave> [namespace code "viewlib_onleave $txt tag$fulltype"]
			$txt tag bind tag$fulltype <1> [namespace code "viewlib_select [winfo toplevel $txt] $lib $type $element"]
			if {$element == {}} {
				$txt insert end "$fulltype" tag$fulltype
			} else {
				$txt insert end "$lib.$type   $element" tag$fulltype
			}
		} else {
			$txt insert end "$fulltype"
		}
	}


	###
	# displays the elements for a type of some library
	# a gives a brief description of the type
	# This function is by the type index
	#
	proc viewlib_showelements_byindex {w lib typeindex} {
		set lb $w.sp1.sp2.elements.listbox
		
		if {[string trim $typeindex] == {}} {
			viewlib_showlibrary $w $lib 
		} else {
			set type [$w.sp1.sp2.types.listbox get $typeindex]
			return [viewlib_showelements $w $lib $type]
		}
	}


	proc viewlib_getProgIDs {classid} {
		set key "HKEY_CLASSES_ROOT\\CLSID\\$classid\\ProgID"
		set result {}
		if {[catch {registry get $key {}} e]} {return {}}
		lappend result $e
		set key "HKEY_CLASSES_ROOT\\CLSID\\$classid\\VersionIndependentProgID"
		if {[catch {registry get $key {}} e]} {return $result}
		lappend result $e
		return $result
	}

	###
	# displays the elements for a type of some library
	# a gives a brief description of the type
	#
	proc viewlib_showelements {w lib type} {
		variable colors
		variable properties

		set elements_lb $w.sp1.sp2.elements.listbox


		set desc $w.sp1.description.desc.t
		#$desc config -state normal
		$desc delete 1.0 end			

		set properties($w-element) {}
		set elements [typelib::typeinfo $lib.$type]
		set typedesc [lindex $elements 0]
		$desc insert end "$typedesc $type "
		$desc tag add DescriptionLabel 1.[string length $typedesc] 1.end
		$desc insert end "\n"
		set line 2
		if {[lindex $elements 4] != {}} {
			$desc insert end \"[lindex $elements 4]\"\n
			$desc tag add MemberDocumentation $line.0 $line.end
			incr line
		}

		if {[lindex $elements 0] == "class"} {
			set classid [lindex $elements 5]
			$desc insert end "ClassID: [lindex $elements 5]\n"
			$desc tag add MemberDocumentation $line.0 $line.end
			incr line
			set pids [viewlib_getProgIDs $classid]
			if {$pids != {}} {
				$desc insert end "ProgID: [lindex $pids 0]\n"
				$desc tag add MemberDocumentation $line.0 $line.end
				incr line
				$desc insert end "VersionIndependentID: [lindex $pids 1]\n"
				$desc tag add MemberDocumentation $line.0 $line.end
				incr line
			}
		}

		if {[string match "typedef" [lindex $elements 0]]} {
			# --- we are working with a typedef
			set t [lindex $elements 3]
			viewlib_writetype $desc $t
		} 
		if {![string match $properties($w-type) $type] && $type != {}} {
			$elements_lb delete 0 end
			

			set properties($w-type) $type
			history_addwindowstate $w

			if {![string match "typedef" [lindex $elements 0]]} {
				foreach method [lsort [lindex $elements 1]] {
					$elements_lb insert end $method
					$elements_lb setimage end ::tlview::method_img
				}

				foreach prop [lsort [lindex $elements 2]] {
					$elements_lb insert end $prop
					$elements_lb setimage end ::tlview::property_img
				}				
				foreach impl [lsort -index 1 [lindex $elements 3]] {
					set t [lindex $impl 1]
					set flags [lindex $impl 0]
					set item $t
					if {[lsearch -exact $flags default] != -1} {
						lappend item "*"
					}

					if {[lsearch -exac $flags source] != -1} {
						lappend item (event source)
					}
					
					$elements_lb insert end $item
					$elements_lb setimage end ::tlview::typedef_img
				}
			}
			bind $elements_lb <<Select>> [namespace code "viewlib_showelement_description_byindex $w $lib $type \[$elements_lb curselection\]"]
		}
		#$desc config -state disabled
	}

	

	proc viewlib_showelement_description_byindex {w lib type elemindex} {


		set elements_lb $w.sp1.sp2.elements.listbox
		if {$elemindex == {}} {
			viewlib_showelements $w $lib $type
			return
		} else {
			set element [$elements_lb get $elemindex]
			# because we tend to mark up default and source interfaces 
			# with appended symbols and names, we'll strip these off
			set element [lindex [lindex $element 0] 0]
		
			# but in fact, any implemented type needs not an element 
			# description, but a jump to that types description
			if {[regexp {^.+\..+$} $element]} {
				set split [split $element .]
				set newlib [lindex $split 0]
				set newtype [lindex $split 1]
				viewlib_select $w $newlib $newtype
			} else {
				viewlib_showelement_description $w $lib $type $element
			}
		}
	}

	###
	# retrieves the description for an element
	proc viewlib_showelement_description {w lib type elem} {
		variable colors
		variable properties

		set txt $w.sp1.description.desc.t
		#$txt config -state normal

		$txt tag bind element <1> [namespace code "viewlib_select $lib.$type $elem"]

		# if we're not viewing this element already			
		if {$elem != {} && ![string match $properties($w-element) $elem]} {
			$txt delete 1.0 end
			set properties($w-element) $elem
			history_addwindowstate $w

			set elementdesc [typelib::typeinfo $lib.$type $elem]
			set elementkind [lindex $elementdesc 0]

			switch $elementkind {
				property {
					$txt insert end "property "

					set propertydesc [lindex $elementdesc 1]
					# insert the flags
					set flags [lindex $propertydesc 0]
					if {[lsearch -exact $flags read] < 0} {
						set flags {(write only)}
					} elseif {[lsearch -exact $flags write] < 0} {
						set flags {(read only)}
					} elseif {$flags != {}} {
						set flags {(read+write)}
					}
					$txt insert end "$flags\n" FlagDescription
					
					# the property type
					viewlib_writetype $txt [lindex $propertydesc 1]
					$txt insert end " "

					# the property name
					$txt insert end "[lindex $propertydesc 2]" DescriptionLabel

					# now do the params
					set params [lrange $propertydesc 3 end]

					foreach param $params {
						$txt insert end "\n\t"

						if {[llength $param] == 3} {
							$txt insert end "[lindex $param 0]\t"
							set param [lrange $param 1 end]
						}
						viewlib_writetype $txt [lindex $param 0]
						$txt insert end "  [lrange $param 1 end]"
					}
					# the documentation for the property
					set documentation [lindex $elementdesc 2]
					if {$documentation != {}} {
						$txt insert end "\n\n\"$documentation\"" MemberDocumentation
					}
				}

				method	{
					set methodesc [lindex $elementdesc 1]
					$txt insert end "method\n"
					
					# the return type
					viewlib_writetype $txt [lindex $methodesc 0]
					$txt insert end " "
					$txt insert end "[lindex $methodesc 1]" DescriptionLabel
					set params [lrange $methodesc 2 end]

					foreach param $params {
						$txt insert end "\n\t"

						if {[llength $param] == 3} {
							$txt insert end "[lindex $param 0]\t"
							set param [lrange $param 1 end]
						}
						viewlib_writetype $txt [lindex $param 0]
						$txt insert end "  [lrange $param 1 end]"
					}
					set documentation [lindex $elementdesc 2]
					if {$documentation != {}} {
						$txt insert end "\n\n\"$documentation\"" MemberDocumentation
					}
				}
			}
		}
		#$txt config -state disabled				

	}


	proc viewlib_copy {w} {
		variable properties
		
		set str $properties($w-viewedlibrary)
		if {$properties($w-type) != {}} {
			set str "$str.$properties($w-type)"
			if {$properties($w-element)!={}} {
				set str "$str $properties($w-element)"
			}
		}
		clipboard clear
		clipboard append -format STRING -- $str
	}

	####
	# Creates a viewer for library
	proc viewlib {{w {}} {lib {}}} {
		variable colors
		variable properties

		if {$w == {}} {

			# iterate over the current windows to find one that is viewing this library
			foreach viewedlib [array names properties *-viewedlibrary] {
				if {[string match $properties($viewedlib) $lib]} {
					set w [lindex [split $viewedlib -] 0]
					break;
				}
			}
		}

		if {$w == {}} {
			# make a unique name
			set count 0
			
			set w ._tlview_$count
			while {[winfo exists $w]} {
				set w ._tlview_[incr count]
			}
		}


		if [winfo exists $w] {
			raise $w
			return $w
		}

		toplevel $w -class TypeLibraryViewer -width 400 -height 300
		wm title $w "Type Library:"

		# top bar - search stuff 
		set topbar [frame $w.topbar]
		pack $topbar -side top -fill x -pady 2
		pack [label $topbar.liblabel -text Library -underline 0 -width 6] -side left -anchor nw
		
		::tlview::TlviewCombo::Create $topbar.libs {::typelib::loadedlibs} -textvariable [namespace current]::properties($w-viewedlibrary)

		$topbar.libs.e config -state disabled
		pack $topbar.libs -side left -padx 3
		pack [button $topbar.back -image ::tlview::leftarrow_img -bd 1 -height 16 -width 16 -command [namespace code "history_back $w"]] -side left 
		pack [button $topbar.forward -image ::tlview::rightarrow_img -bd 1 -height 16 -width 16 -command [namespace code "history_forward $w"]] -side left 
		pack [button $topbar.copy -image ::tlview::copy_img -bd 1 -height 16 -width 16 -command [namespace code "viewlib_copy $w"]] -padx 3 -side left 
		pack [button $topbar.libselect -image ::tlview::libselect_img -bd 1 -height 16 -width 16 -command [namespace code "libs $w.selectlibs"]] -padx 0 -side left 
		
		TlviewTooltip::Set $topbar.libs "Loaded Libraries"
		TlviewTooltip::Set $topbar.back "Previous in History"
		TlviewTooltip::Set $topbar.forward "Next in History"
		TlviewTooltip::Set $topbar.copy "Copy"
		TlviewTooltip::Set $topbar.libselect "Referenced Type Libraries"

		searchbox $w.searchbox
		pack $w.searchbox -side top -fill x -pady 2

		# splitters
		set sp1 [Splitter $w.sp1 -orient horizontal -type fixB -position 160 -barwidth 5 -width 460 -height 380]
		set sp2 [Splitter $w.sp1.sp2 -orient vertical -type fixA -position 200 -barwidth 5]
		pack $sp1 -fill both -expand 1 -side bottom
		$sp1 config -windowA $sp2
		
		# description frame
		set desc [frame $sp1.description]
		$sp1 config -windowB $desc
		pack [label $desc.label -text Description] -side top -anchor nw
		pack [scrltxt $desc.desc] -side top -anchor nw -fill both -expand 1
		$desc.desc.t tag configure DescriptionLabel -foreground $colors(labelcolor) -font tlviewerbold
		$desc.desc.t tag configure FlagDescription -font tlvieweritalic
		$desc.desc.t tag configure MemberDocumentation -font tlvieweritalic -foreground $colors(labelcolor)
		$desc.desc.t config -exportselection 1 -cursor xterm -insertontime 0 -selectforeground SystemHighlightText
		bind $desc.desc.t <Alt-F4> {continue}
		bind $desc.desc.t <KeyPress> "break;"
		
		# types frame
		set types [::tlview::TlviewScrolledListBox $sp2.types]
		$types.listbox config -font {Arial 10 {}}
		$sp2 config -windowA $types
		$types.label config -text Types

		set elements [::tlview::TlviewScrolledListBox $sp2.elements]
		$elements.listbox config -font {Arial 10 {}}
		$sp2 config -windowB $elements
		$elements.label config -text "Elements"

		if {$lib == {}} {
			set lib [lindex [typelib::loadedlibs] 0]
		}

		trace vdelete [namespace current]::properties($w-viewedlibrary) w [namespace code viewlib_libchanged]
		set properties($w-viewedlibrary) {}
		set properties($w-library) {}
		set properties($w-type) {}
		set properties($w-element) {}
		trace variable [namespace current]::properties($w-viewedlibrary) w [namespace code viewlib_libchanged]

		bind $w <Destroy> [namespace code "viewlib_ondestroy %W"]
		history_init $w
		viewlib_showlibrary $w $lib
		
		return $w
	}

	proc viewlib_ondestroy {w} {
		variable properties

		if {[winfo toplevel $w] == $w} {
			history_erase $w
			array unset properties $w
			array unset properties $w-*
		}
	}

	proc viewlib_libchanged {n1 n2 command} {
		variable properties
		set lib $properties($n2)
		set w [lindex [split $n2 -] 0]
		viewlib_showlibrary $w $lib
	}

	proc viewtype {fullname {elem {}} {history 1}} {
		set split [split $fullname .]
		set lib [lindex $split 0]
		set type [lindex $split 1]

		viewlib_select $lib $type $elem $history
	}

	### -- Search box code
	

	proc searchbox {w} {
		variable properties

		destroy $w 

		frame $w 
		set splitter [winfo parent $w]

		frame $w.top 
		pack [label $w.top.searchlabel -text Search -underline 0 -width 6] -side left -anchor nw
		::tlview::TlviewCombo::Create $w.top.searchterm [namespace code "search_history_getlist $w"]
		bind $w.top.searchterm.e <Return> [namespace code "searchbox_search $w"]

		button $w.top.search -image ::tlview::find_img -borderwidth 1 -command [namespace code "searchbox_search $w"]
		button $w.top.showhide -image ::tlview::show_img -borderwidth 1 -command [namespace code "searchbox_showhide $w"]
		pack $w.top.showhide $w.top.search -side right -padx 2
		pack $w.top.searchterm  -side left -fill x -expand 1 -padx 3

		TlviewTooltip::Set $w.top.showhide "Show/Hide Search Results"
		TlviewTooltip::Set $w.top.searchterm "Search String"
		TlviewTooltip::Set $w.top.search "Search for String"

		scrltxt $w.searchresults
		$w.searchresults.t config -height 10 -tabs 5c -state disabled

		grid $w.top -sticky nsew
		grid rowconfigure $w 1 -weight 1
		grid columnconfigure $w 0 -weight 1
		update
		set properties($w-collapsed) [expr [winfo reqheight $w] + 2]
		set properties($w-expanded) 200
		set properties($w-min) 90

		#$splitter config -windowB $w -type fixB -position $properties($w-collapsed) -min $properties($w-collapsed) -max $properties($w-collapsed)
		return $w
	}

	proc searchbox_show {w} {
		variable properties
		if {[lsearch [grid slaves $w] $w.searchresults] >= 0} return

		set toplevel [winfo toplevel $w]
		set windowheight [winfo height $toplevel]
		set windowheight [expr $windowheight + [winfo reqheight $w.searchresults] + 5]

		grid $w.searchresults -row 1 -column 0 -sticky nsew -pady 5
		wm geometry $toplevel [winfo width $toplevel]x$windowheight
		$w.top.showhide config -image ::tlview::hide_img -relief sunken
	}

	proc searchbox_search {w} {
		variable properties
		set query [$w.top.searchterm.e get]
		set lib $properties([winfo toplevel $w]-viewedlibrary)

		search_history_add $w $query
		search $w $query
		$w.top.searchterm.e selection clear 
		$w.top.searchterm.e selection range 0 end
	}

	proc searchbox_hide {w} {
		variable properties

		if {[lsearch [grid slaves $w] $w.searchresults] < 0} return

		set textheight [winfo reqheight $w.searchresults] 
		set toplevel [winfo toplevel $w]
		set windowheight [expr [winfo height $toplevel] - $textheight - 5]
		grid forget $w.searchresults
		
		wm geometry $toplevel [winfo width $toplevel]x$windowheight
		$w.top.showhide config -image ::tlview::show_img -relief raised
	}

	proc searchbox_showhide {w} {
		if {[lsearch [grid slaves $w] $w.searchresults] < 0} {
			searchbox_show $w
		} else {
			searchbox_hide $w
		}
	}

	proc search {w query} {
		variable properties
		
		# ensure that the search window exists

		set w [winfo toplevel $w]
		set lib $properties($w-viewedlibrary)
		set searchbox $w.searchbox
		set sr $searchbox.searchresults.t 

		# set up the text box
		$sr config -state normal
		$sr delete 1.0 end
		
		searchbox_show $searchbox

		set query [join [list * $query *] {}]

		foreach desc [typelib::types $lib] {
			set fulltype [lindex $desc 1]
			set reflib [lindex [split $fulltype .] 0]
			set reftype [lindex [split $fulltype .] 1]
			
			# perform search on the type name
			if {[string match -nocase $query $reftype]} {
				viewlib_writetype $sr $fulltype
				$sr insert end "\n"
			}

			# now iterate through its members
			set typeinfo [typelib::typeinfo $fulltype]
			foreach item [lindex $typeinfo 1] {
				if {[string match -nocase $query $item]} {
					viewlib_writetype $sr $fulltype.$item
					$sr insert end "\n"
				}
			}

			foreach item [lindex $typeinfo 2] {
				if {[string match -nocase $query $item]} {
					viewlib_writetype $sr $fulltype.$item
					$sr insert end "\n"
				}
			}
		}
		$sr config -state disabled
	}

	proc viewtype {fulltype {element {}}} {
		variable properties
		set w {}
		set split [split $fulltype .]
		set lib [lindex $split 0]
		set type [lindex $split 1]


		set w [viewlib {} $lib]
		update
		viewlib_select $w $lib $type $element
	}


	proc class {obj} {
		viewtype [optcl::class $obj]
	}

	proc interface {obj} {
		viewtype [optcl::interface $obj]
	}
}
