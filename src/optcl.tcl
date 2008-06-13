package provide optcl 3.0
set env(OPTCL_LIBRARY) [file dirname [info script]]
source [file join $env(OPTCL_LIBRARY) scripts Utilities.tcl]
source [file join $env(OPTCL_LIBRARY) scripts Splitter.tcl]
source [file join $env(OPTCL_LIBRARY) scripts TypeLib.tcl]
source [file join $env(OPTCL_LIBRARY) scripts ImageListBox.tcl]
source [file join $env(OPTCL_LIBRARY) scripts Tooltip.tcl]
if {[info commands tk] != {}} {source [file join $env(OPTCL_LIBRARY) scripts TLView.tcl]}
load [file join $env(OPTCL_LIBRARY) bin optcl.dll]
typelib::updateLibs