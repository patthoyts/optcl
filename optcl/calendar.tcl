#####################################################
# This file demonstrates the Calendar control being 
# integrated within a Tk widget, and bound to a
# an event handler
#####################################################


# in case we want to do some debugging
bind . <F2> {console show}


# optcl load happens here
package require optcl


##
# called when an AfterUpdate event is raised.
# the first parameter is the object that raised
# the event
proc onupdate {obj} {
	global currentdate
	set currentdate [$obj : value]
}



# main script------


# create a status bar to show the current date
label .cd -bd 1 -relief sunken -textvariable currentdate
pack .cd -side bottom -fill x

# create the calendar object
set cal [optcl::new -window .cal MSCAL.Calendar]
.cal config -width 300 -height 300
pack .cal

# bind to the calendar AfterUpdate event
# routing it to the tcl procedure onupdate
#
#optcl::bind $cal AfterUpdate onupdate


# get the current value
#set currentdate [$cal : value]


# make a button to view the type information of 
# the calendar
button .b -text TypeInfo -command {tlview::viewtype [optcl::class $cal]}
pack .b -side bottom -anchor se


