console show
package require optcl

set xl [optcl::new Excel.Application]
$xl : visible 1
$xl -with workbooks add
$xl -with workbooks.item(1).worksheets.item(1).range(a1,d4) : value 16
set r [$xl -with workbooks.item(1).worksheets.item(1) range a1 d4]
