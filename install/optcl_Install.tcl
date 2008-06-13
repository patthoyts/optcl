
# OpTcl Installer
# Author: Fuzz
# fuzz@sys.uea.ac.uk

package require registry

set piccy ../docs/optcl_medium.gif

set installfolder [file join [info library] .. optcl]
set installname optcl.dll

puts "Install dir: $installfolder"
set version [info tclversion]

if {$version < 8.0} {
	tk_messageBox -message "Sorry, but OpTcl needs Tcl version 8.0.5" -type ok
	exit
} elseif {$version < 8.1} {
	set dll optcl80.dll
} elseif {$version < 9.0} {
	set dll optclstubs.dll
} else {
	tk_messageBox -message "Sorry, but OpTcl was compiled for Tcl major-version 8" -type ok
}

image create photo optclim -file $piccy

proc updategui {} {
	global installfolder installname
	if [file exists [file join $installfolder $installname]] {
		.uninstall config -state normal
		.install config -text "Re-install for Tcl[info tclversion]"
	} else {
		.uninstall config -state disabled
		.install config -text "Install for Tcl[info tclversion]"
	} 
}

proc install {} {
	global installfolder installname dll
	set answer [tk_messageBox -title {} -message "Okay to install $dll in $installfolder\nand register as OpTcl package?" -icon question -type yesno]
	
	switch $answer {
		no {}
		yes {
			set bad [catch {
				file mkdir $installfolder
				file copy -force $dll [file join $installfolder $installname]
				pkg_mkIndex -direct $installfolder
			} err]
			if {$bad} {
				tk_messageBox -type ok -message "Error: $err" -icon error
			} else {
				tk_messageBox -type ok -message "OpTcl successfully installed." -icon info
			}
			exit
		}
	}
}

proc uninstall {} {
	global installfolder installname
	set reply [tk_messageBox -type yesno -message "Delete package OpTcl located at $installfolder?" -icon question]
	if {[string compare $reply yes] != 0} return
	file delete [file join $installfolder $installname] [file join $installfolder pkgIndex.tcl] $installfolder
	updategui
}

wm title . "OpTcl Installer - F2 for console"
bind . <F2> {console show}
bind . <Alt-F4> {exit}

label .im -image optclim -relief flat -bd 0
button .install -text Install... -command install -width 16 -height 1 -bd 2 -font {arial 8 bold}
button .uninstall -text Uninstall -command uninstall -width 16 -height 1 -bd 2 -font {arial 8 bold}
button .quit -text Quit -command exit -bd 2 -font {arial 8 bold} -width 5 -height 1

grid .im -column 0 -row 0 -rowspan 2 -padx 2 -pady 2
grid .install -column 1 -row 0 -padx 2 -pady 2 -sticky nsew
grid .uninstall -column 2 -row 0 -padx 2 -pady 2 -sticky nsew
grid .quit -column 1 -row 1 -columnspan 2 -padx 2 -pady 2 -sticky nsew


wm resizable . 0 0 
updategui
raise .
focus -force .
