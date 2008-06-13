package require registry

namespace eval typelib {

	variable typelibraries
	array set typelibraries {}


	namespace export *
	namespace import -force ::optcl_utils::*

	# -----------------------------------------------------------------------------


	# latest_typelib_version --
	#	For a given registered type library guid, it retrieves the most recent version
	#	number of the library, and returns a string giving a version qualified string
	#	description for the library. Returns {} iff failed.
	#
	proc latest_typelib_version {typelib_guid} {
		set typelibpath HKEY_CLASSES_ROOT\\TypeLib\\$typelib_guid
		if {[registry keys HKEY_CLASSES_ROOT\\TypeLib $typelib_guid] == {}} {
			puts "couldn't find typelib $typelib"
			return {}
		}
		set v [lindex [lsort -decreasing -real [registry keys $typelibpath]] 0]

		if {$v == {}} {
			puts "bad typelib version number: $v for $typelib_guid"
			set result {}
		} else {
			set result [makelibname $typelib $v]
		}
		return $result
	}

	# makelibname --
	#	standard function for creating a library's human readable name from the 
	#	registered guid.
	#
	proc makelibname {typelib_guid ver} {
		set maj 0
		set min 0
		scan $ver "%d.%d" maj min
		return "[registry get HKEY_CLASSES_ROOT\\TypeLib\\$typelib_guid\\$ver {}] (Ver $maj.$min)"
	}

	proc path_from_libname {libname} {
		variable typelibraries
		set r [array get typelibraries $libname]
		if {$r == {}} {
			error "library does not exist: $libname"
		}
		set libsettings [lindex $r 1]
		return [eval path_from_libid $libsettings]
	}

	# updateLibs -- called to enumerate and store the system libraries
	proc updateLibs {} {
		variable typelibraries;

		# enumerate the current type libraries to make sure that they're still there
		foreach library [array names typelibraries] {
			try {
				mset {name path} $typelibraries($library)
				if {![file exists $path]} {throw {}}
			} catch {er} {
				unset typelibraries($library)
			}
		}

		# now iterate over the registered type libraries in the system
		set root {HKEY_CLASSES_ROOT\TypeLib}
		foreach id [registry keys $root] {
			try {
				foreach v [registry keys $root\\$id] {
					scan $v "%d.%d" maj min;
					if [catch {
						set flags [registry get $root\\$id\\$v\\FLAGS {}];
					}] { set flags 0;}

					# check for restricted or hidden libraries
					if {[expr ($flags & 1) || ($flags & 4)]} {
						continue;
					}

					set name [makelibname $id $maj.$min]
					set path [typelib::reglib_path $id $maj $min]
					addLibrary $name $id $maj $min $path
				}
			} catch {e} {
				puts $e
			}
		}
	}


	proc addLibrary {name typelib_id maj min path} {
		variable typelibraries
		set typelibraries([list [string toupper $typelib_id] $maj $min]) [list $name $path]
	}

	proc persistLoaded {} {
		set cmd "typelib::loadLibsFromDetails"
		lappend cmd [getLoadedLibsDetails]
		return $cmd
	}

	# getLoadedLibsDetails --
	#	Retrieves a list of descriptors for the current loaded libraries
	proc getLoadedLibsDetails {} {
		set result {}
		foreach progname [typelib::loadedlibs] {
			lappend result [typelib::loadedlib_details $progname]
		}
		return $result
	}

	proc loadLibsFromDetails {details} {
		foreach libdetail $details {
			loadLibFromDetail $libdetail
		}
	}

	proc loadLibFromDetail {libdetail} {
		variable typelibraries
		mset {guid maj min path fullname} $libdetail

		# if the library is already registered, get the path from the registry
		mset { _ regpath} [lindex [array get typelibraries [list $guid $maj $min]] 1]
		if {$regpath != {}} {
			set path $regpath
		} 
		
		typelib::load $path
		addLibrary $fullname $guid $maj $min $path
	}

	proc load {path} {
		set progname [typelib::_load $path]
		mset {guid maj min path fullname} [typelib::loadedlib_details $progname]
		addLibrary $fullname $guid $maj $min $path
		return $progname
	}


	# -----------------------------------------------------------------------------


	# libdetail -- 
	#	returns the id, maj and min version numbers and 
	#	the path as a list if they exists, else throws an error.
	#
	proc libdetail {name} {
		variable typelibraries

		if {[array names typelibraries $name] == {}} {
			error "could not find the library '$name'"
		}

		return [lindex [array get typelibraries $name] 1]
	}


	#------------------------------------------------------------------------------

	# alllibs -- returns all the registered libraries by {guid maj min} identification
	proc alllibs {} {
		variable typelibraries
		return [array names typelibraries]
	}

	# returns the fully qualified default interface for a com class
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
}




namespace eval COM {
	namespace import -force ::typelib::*

	# categories
	# retrieve a list of all category names
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


	# collate all the category names under the category clsid (parameter 1) into an
	# array passed by name
	proc collate_category_names {category arrname} {
		upvar $arrname categories

		set ck "HKEY_CLASSES_ROOT\\Component Categories\\$category"
		catch {
			foreach value [registry values $ck] {
				catch {set categories([registry get $ck $value]) ""}
			}
		} err
		return $err
	}


	# collates all categories for a given clsid in an array that is passed by name
	proc clsid_categories_to_array {clsid arrname} {
		upvar $arrname categories
		set k "HKEY_CLASSES_ROOT\\CLSID\\$clsid"
		
		# catch if there aren't any implemented categories

			foreach subkey [registry keys "HKEY_CLASSES_ROOT\\CLSID\\$clsid"] {
				switch $subkey {
					{Implemented Categories} {
						foreach category [registry keys "$k\\$subkey"] {
							collate_category_names $category categories
						}
					}

					Programmable {
						array set categories {{Automation Objects} {}}
					}

					Control {
						array set categories {Controls {}}
					}

					DocObject {
						array set categories {{Document Objects} {}}
					}

					Insertable {
						array set categories {{Embeddable Objects} {}}
					}
				}
			}
		
	}

	# retrieves, as a list, the categories for the given clsid
	proc clsid_categories {clsid} {
		array set categories {}
		clsid_categories_to_array $clsid categories
		return [array names categories]
	}


	# retrieves all clsids that match the category name given by the first parameter
	proc clsids {{cat {}}} {
		array set categories {}
		set clsidk "HKEY_CLASSES_ROOT\\CLSID"
		if {$cat == {}} {
			return [registry keys $clsidk]
		}

		# else ...

		set classes {}
		
		foreach clsid [registry keys $clsidk] {
			catch [unset categories]
			array set categories {}
			clsid_categories_to_array $clsid categories
			if {[array names categories $cat]!={}} {
				lappend classes $clsid
			}
		}
		return $classes
	}



	# provides a description for the clsid
	proc describe_clsid {clsid} {
		set clsidk "HKEY_CLASSES_ROOT\\CLSID\\$clsid"
		return [registry get $clsidk {}]
	}



	# retrieves a list of clsid descriptor for all clsids that have the category specified by
	# parameter one. If parameter is {} then all clsids are returned.
	proc describe_all_clsids {{cat {}}} {
		set l {}
		foreach clsid [categories::all_clsids $cat] {
			lappend l [categories::describe_clsid $clsid]
		}
		return [lsort -dictionary $l]
	}

	# retrieve the programmatics identifier for a clsid.
	# If any exist, the result of this procedure is the programmatic identifier for the
	# the clsid, followed by an optional version independent identifier
	proc progid_from_clsid {clsid} {
		set clsidk "HKEY_CLASSES_ROOT\\CLSID\\$clsid"
		set progid {}
		set verindid {}
		catch {set progid [registry get "$clsidk\\ProgID" {}]}
		catch {lappend progid [registry get "$clsidk\\VersionIndependentProgID" {}]}
		return $progid
	}


	proc typelib_from_clsid {clsid} {
		set clsidk "HKEY_CLASSES_ROOT\\CLSID\\$clsid"
		# associated typelib?
		
		if {[registry keys $clsidk TypeLib] == {}} {
			return {}
		}
		set typelib [registry get $clsidk\\TypeLib {}]

		# does it exist?
		if {[registry keys HKEY_CLASSES_ROOT\\TypeLib $typelib] == {}} {
			puts "couldn't find typelib $typelib from clsid $clsid"
			return {}
		}

		# do we have a version number??
		if {[registry keys $clsidk Version] != {}} {
			set ver [registry get $clsidk\\Version {}]
			set result [makelibname $typelib $ver]
		} elseif {[registry keys $clsidk VERSION] != {}} {
			set ver [registry get $clsidk\\VERSION {}]
			set result [makelibname $typelib $ver]
		} else {
			# get the latest version of the type library
			set result [latest_typelib_version $typelib]
		}
		return $result
	}
}
