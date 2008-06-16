################################################################
# This file demonstrates the automation MS Word
################################################################


# for debuggin
bind . <F2> {console show}

#load optcl
package require optcl



# with this procedure, closing the document closes wish
proc onclose {obj} {
	# if the document is closing then exit
	# but we can't call exit here as we are processing an event
	# so set up a timer on this
	after 500 {exit}
}

set word [optcl::new word.application]
$word : visible 1

# create a new doc
set doc [$word -with documents add]

# bind to its close event of the document
optcl::bind $doc Close onclose


# gui

button .st -text "Set Text" -command {$doc -with content : text [.f.t get 1.0 end]; $doc : saved 1}
pack .st

frame .f -bd 1 -relief sunken
pack .f -side top -fill both -expand 1
scrollbar .f.ys -orient vertical -command {.f.t yview}
pack .f.ys -side right -fill y
text .f.t -yscrollcommand {.f.ys set} -bd 0 -relief flat
pack .f.t -fill both -expand 1

.f.t insert end "Please type your text here and press 'Set Text'"

