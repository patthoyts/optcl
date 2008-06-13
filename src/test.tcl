console show
load optcl



proc ie_test {} {
	global ie
	set ie [optcl::new -window .ie {{8856F961-340A-11D0-A96B-00C04FD705A2}}]
	pack .ie
	$ie navigate www.wired.com
}

proc vrml_test {} {
	global vrml
	set vrml [optcl::new -window .vrml {{4B6E3013-6E45-11D0-9309-0020AFE05CC8}}]
	pack .vrml
}

proc tree_test {} {
	global tv
	set tv [optcl::new -window .tv {{C74190B6-8589-11D1-B16A-00C0F0283628}}]
	pack .tv
	set n1 [$tv -with nodes add]
	$n1 : text "Node 1" key "1 Node"
	optcl::unlock $n1
	set n2 [$tv -with nodes add "1 Node" 4 "2 Node" "Node 2"]
	$n2 : text "Node 2.5"
	optcl::unlock $n2
}

proc dp_test {} {
	global dp
	destroy .date
	set dp [optcl::new -window .date MSComCtl2.DTPicker]
	.date config -width 100 -height 20
	pack .date
	tlview::viewtype [optcl::class $dp]
}

proc cal_test {} {
	global cal
	destroy .cal
	set cal [optcl::new -window .cal MSCAL.Calendar]
	pack .cal
}


proc pb_test {} {
	global pb mousedown

	proc PBMouseDown {obj args} {
		global mousedown
		set mousedown $obj
	}	

	proc PBMouseUp {args} {
		global mousedown
		set mousedown {}
	}

	proc PBMouseMove {obj button shift x y} {
		global mousedown
		if {$mousedown == {}} return
		if {[string compare $mousedown $obj]==0} {
			$obj : value $x
		}
	}
	destroy .pb
	set pb [optcl::new -window .pb MSComctlLib.ProgCtrl]
	pack .pb
	.pb config -width 100 -height 10
	optcl::bind $pb MouseDown PBMouseDown
	optcl::bind $pb MouseUp PBMouseUp
	optcl::bind $pb MouseMove PBMouseMove
}




proc word_test {} {
	global word

	set word [optcl::new word.application]
	$word : visible 1
}


proc tl_test {} {
	typelib::load {Microsoft Shell Controls And Automation (Ver 1.0)}
	tlview::refview .r
	tlview::loadedlibs .l
}



proc cosmo_test {} {
	global co
	set co [optcl::new -window .co SGI.CosmoPlayer.2]
	pack .co
}
