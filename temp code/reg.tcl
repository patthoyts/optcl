package require registry


namespace eval COM {

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
	proc descrive_all_clsids {{cat {}}} {
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
}

