lappend auto_path d:/tcl-lib

package require optcl
bind . <F2> {console show}

wm title . {PDF Document in Tk}
set pdf [optcl::new -window .pdf {C:\Program Files\Adobe\Acrobat 4.0\Help\ENU\acrobat.pdf}]
.pdf config -width 500 -height 300
pack .pdf -fill both -expand 1

# to view the type information for the control
pack [button .b  -text "View TypeLibrary for IE container" -command {
			tlview::viewtype [ optcl::class $pdf ]
	} ] -side bottom

# can't execute these until the document has loaded...
#set doc [$pdf : document]
#tlview::viewtype [ optcl::class $doc ]
