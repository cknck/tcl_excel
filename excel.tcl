# ----------------------------------------------------------------------------
# Script Name:   Excel.tcl
# Creation Date: 
# Last Modified: 
# Copyright (c) 2007-2010 SunMoon. All Rights Reserved.
# Purpose: Excel API
# ----------------------------------------------------------------------------
package require tcom
if {0} {
set fn {d:/temp/tst.xls}

set appObj [::tcom::ref createobject "Excel.Application"]
#$appObj Visible 1
#set appObj [::tcom::ref getactiveobj "Excel.Application"]

set wbs [$appObj Workbooks]
#set wb 	[$wbs Add]
set wb 	[$wbs Open $fn]

set wss [$wb Worksheets]
set ws 	[$wss Item [expr 1]]

#$appObj Quit

}


#lappend auto_path {F:\tcl\TclExcel}

#source {F:\tcl\TclExcel\biff.tcl}

#Biff::new_table tab BIFF2
#Biff::new_sheet tab cells

if {0} {
    set range [$ws UsedRange]
	set rows [$range Rows]
}

if {0} {
::tcom::import {E:\Program Files\Microsoft Office\OFFICE11\excel.exe}

set range [$ws Range e7]
set range2 [$ws Range e7 [$range End $Excel::XlDirection(xlDown)]]
[$range2 Rows] Group
$range2 ClearOutline

[$ws Outline] SummaryRow $Excel::Constants(xlAbove)
[$ws Outline] SummaryRow $Excel::XlSummaryRow(xlSummaryAbove)

$range2 Select
set sel [$appObj Selection]
$sel ClearOutline

$sel Rows xxxxxx
#$cells Item [expr 2] [expr 3] xxxxxx

set range [$ws Range e7]
set range2 [$ws Range e7 [$range End $Excel::XlDirection(xlDown)]]
$range AutoFill $range2 $Excel::XlAutoFillType(xlFillCopy)
}

proc Alpha2Col {str} {
    set str [string toupper $str]
    if {![regexp {^[A-Z]+$} $str match]} {
        return 0
    }
    
    set alpha_tab {-ABCDEFGHIJKLMNOPQRSTUVWXYZ}
    set power 26
    set p0 1
    set p1 $power
    set col 0
    
    for {set i [expr [string length $str]-1]} {$i>=0} {incr i -1} {
        set ch  [string index $str $i]
        set c   [string first $ch $alpha_tab]
        set col [expr $col + $c*$p0]
        set p0  $p1
        set p1  [expr $p1*$power]
    }
    
    return $col
}

proc Col2Alpha {col} {
    if {![regexp {^[0-9]+$} $col match]} {
        return
    }
    
    set alpha_tab0 {ZABCDEFGHIJKLMNOPQRSTUVWXY}
    set power 26
    
    set str ""
    while {$col} {
        set c   [expr $col%$power]
        set ch  [string index $alpha_tab0 $c]
        set str $ch$str
        
        set col [expr $col/$power]
        if {$c==0} {
            incr col -1
        }
    }
    
    return $str
}

proc word2int {word {alphabet ""}} {
    if {$alphabet eq ""} {set alphabet [a-z]}
    set i [expr {wide(0)}]
    foreach c [split $word ""] {
        set i [expr {$i*[llength $alphabet]+[lsearch $alphabet $c]+1}]
    }
    if {$i<0} {error "word $word too long to fit in integer"}
    set i
}
proc int2word {int {alphabet ""}} {
    if {$alphabet eq ""} {set alphabet [a-z]}
    set word ""
    set la [llength $alphabet]
    while {$int > 0} {
       incr int -1
       set word  [lindex $alphabet [expr {$int % $la}]]$word
       set int   [expr {$int/$la}]
    }
    set word
}

proc a-z {} {list a b c d e f g h i j k l m n o p q r s t u v w x y z}

proc AllocId {id_list} {
    if {$id_list == ""} {
        return 0
    }
    
    set id_list [lsort -dic $id_list]
    
    set id_max [lindex $id_list end]
    
    return [expr $id_max+1]
}

proc dump {obj} {
    set interface [::tcom::info interface $obj]
    puts "interface [$interface name]"

    set properties [$interface properties]
    foreach property $properties {
        puts "property $property"
    }

    set methods [$interface methods]
    foreach method $methods {
	puts "method [lrange $method 0 2] \{"
	set parameters [lindex $method 3]
	foreach parameter $parameters {
            puts "    \{$parameter\}"
	}
	puts "\}"
    }
}

namespace eval excel {
    variable appObj
    variable wbsObj
    
    # workbook对象数
    variable wbCnt
    # workbook索引列表
    variable wbIdxList
    # workbook Name列表
    variable wbNameList
    
    # 激活的workbook索引
    variable wbActIdx
    # 激活的workbook对象
    variable wbAct
    # 激活的worksheet索引
    variable wsActIdx
    # 激活的worksheet对象
    variable wsAct
    
    # wsNameList(0)，worksheet Name列表数组，下标从0开始
    variable wsNameList
    # wsIdxList(0)，worksheet索引列表数组，下标从0开始
    variable wsIdxList
    
    # wbObj(1)，workbook对象数组，下标从0开始
    variable wbObj
    # wssObj(1)，worksheets对象数组，下标从0开始
    variable wssObj
    # wsCnt(1)，workbook的页数数组，下标从0开始
    variable wsCnt
    
    # wsObj(1,1)，worksheet对象数组，下标从0,0开始
    variable wsObj
    # cellsObj(1,1)，cells对象数组，下标从0,0开始
    variable cellsObj
    
    proc show {} {
        if {[info exist excel::appObj]} {
            puts "appObj: \[$excel::appObj\]"
            
            if {$excel::appObj == ""} {
                return
            }
            
            #puts "wbsObj: \[$excel::wbsObj\]"
            puts "wbCnt: \[$excel::wbCnt\] wbIdxList: \[$excel::wbIdxList\]"
            
            foreach i $excel::wbIdxList {
                if {$excel::wbActIdx == $i} {
                    set act "A"
                } else {
                    set act " "
                }
                
                puts "    wbObj($i):  $act   \[[$excel::wbObj($i) Name]\]"
                #puts "     wssObj($i):      \[$excel::wssObj($i)\]"
                #puts "     wsCnt($i):      \[$excel::wsCnt($i)\]"
                
                foreach j $excel::wsIdxList($i) {
                    if {($excel::wbActIdx == $i) && ($excel::wsActIdx == $j)} {
                        set act "A"
                    } else {
                        set act " "
                    }
                    
                    puts "        wsObj($i,$j):  $act   \[[$excel::wsObj($i,$j) Name]\]"
                    #puts  "         cellsObj($i,$j)      \[$excel::cellsObj($i,$j)\]"
                }
            }
        }
    }
    
    proc add_wb {wbObj} {
        puts "== debug: add_wb -> wbIdxList: \[$excel::wbIdxList\]"
        set i [AllocId $excel::wbIdxList]
        
        set excel::wbObj($i) $wbObj
        lappend excel::wbIdxList $i
        lappend excel::wbNameList [$wbObj Name]
        
        set excel::wsIdxList($i) ""
        set excel::wsNameList($i) ""
        
        set excel::wssObj($i) [$excel::wbObj($i) Worksheets]
        set excel::wsCnt($i) [$excel::wssObj($i) Count]
        
        for {set j 0} {$j < $excel::wsCnt($i)} {incr j} {
            set excel::wsObj($i,$j) [$excel::wssObj($i) Item [expr $j+1]]
            set excel::cellsObj($i,$j) [$excel::wsObj($i,$j) Cells]
            
            lappend excel::wsIdxList($i) $j
            lappend excel::wsNameList($i) [$excel::wsObj($i,$j) Name]
        }
    }
    
    proc add_ws {wbObj wsObj} {
        set ix [excel::get_wb_ix [$wbObj Name]]
        set i [lindex $excel::wbIdxList $ix]
        
        puts "== debug: add_ws -> wsIdxList($i) \[$excel::wsIdxList($i)\]"
        set j [AllocId $excel::wsIdxList($i)]
        
        set excel::wsObj($i,$j) $wsObj
        set excel::cellsObj($i,$j) [$excel::wsObj($i,$j) Cells]
        
        lappend excel::wsIdxList($i) $j
        lappend excel::wsNameList($i) [$excel::wsObj($i,$j) Name]
    }
    
    proc get_wb_ix {name} {
        # 不区分大小写
        set cnt [llength $excel::wbNameList]
        
        for {set ix 0} {$ix < $cnt} {incr ix} {
            set wbName [lindex $excel::wbNameList $ix]
            if {[string compare -nocase $wbName $name] == 0} {
                break
            }
        }
        
        if {$ix == $cnt} {
            return -1
        } else {
            return $ix
        }
    }
    
    proc get_ws_ix {wbObj name} {
        set ix [excel::get_wb_ix [$wbObj Name]]
        if {$ix == -1} {
            puts "**** ERROR **** \[get_ws_ix -> need update! \]"
            return -1
        }
        
        set i [lindex $excel::wbIdxList $ix]
        
        # 不区分大小写
        set cnt [llength $excel::wsNameList($i)]
        
        for {set ix 0} {$ix < $cnt} {incr ix} {
            set wsName [lindex $excel::wsNameList($i) $ix]
            if {[string compare -nocase $wsName $name] == 0} {
                break
            }
        }
        
        if {$ix == $cnt} {
            return -1
        } else {
            return $ix
        }
    }
}

proc excel::init {} {
    if { [catch {set excel::appObj [::tcom::ref getactiveobj "Excel.Application"]}] } {
        # 启动Excel
        set excel::appObj [::tcom::ref createobject "Excel.Application"]
        
        if {$excel::appObj == ""} {
            puts "**** ERROR **** \[excel::init -> Excel start failed! \]"
            return
        }
    }
    
    set excel::wbsObj [$excel::appObj Workbooks]
    
    # 更新内部变量
    set excel::wbCnt 0
    
    set excel::wbIdxList ""
    set excel::wbNameList ""

    excel::update
}

proc excel::update {} {
    # appObj, wbsObj没有更新，可能有问题
    
    # 更新wbCnt
    set excel::wbCnt [$excel::wbsObj Count]
    
    # 初始化合法记录
    set ix_list ""
    
    puts "== debug: update -> wbIdxList: \[$excel::wbIdxList\]"
    
    # 根据对wbObj Name检查所有记录
    for {set id 1} {$id <= $excel::wbCnt} {incr id} {
        set wbObj [$excel::wbsObj Item [expr $id]]
        
        set ix [excel::get_wb_ix [$wbObj Name]]
#            puts "== debug: update -> $ix: [$wbObj Name]"
        if {$ix > -1} {
            # 合法记录
            lappend ix_list $ix
            
            continue
            
        } else {
            # 新增项
            excel::add_wb $wbObj
            
            # 合法记录
            lappend ix_list [expr [llength $excel::wbIdxList]-1]
        }
    }
    
    set ix_list [lsort -dic -dec $ix_list]
#        puts "== debug: update -> ix_list: \[$ix_list\]"
    
    # 找出非法记录
    set ix_dellist $excel::wbIdxList
    foreach ix $ix_list {
        set ix_dellist [lreplace $ix_dellist $ix $ix]
    }
    
    set ix_dellist [lsort -dic -dec $ix_dellist]
    puts "== debug: update -> ix_dellist: \[$ix_dellist\]"
    
    # 删除非法记录
    foreach i $ix_dellist {
        unset excel::wbObj($i)
        set ix [lsearch $excel::wbIdxList $i]
        set excel::wbIdxList [lreplace $excel::wbIdxList $ix $ix]
        set excel::wbNameList [lreplace $excel::wbNameList $ix $ix]
    }
    
    # 更新ActiveWorkbook
    set wbAct [$excel::appObj ActiveWorkbook]
    if {$wbAct == ""} {
        set excel::wbActIdx ""
        set excel::wbAct ""
        
        set excel::wsActIdx ""
        set excel::wsAct ""
        return
    }
    
    set ix [excel::get_wb_ix [$wbAct Name]]
    set i [lindex $excel::wbIdxList $ix]
    
    if {[info exist excel::wbActIdx] && ($excel::wbActIdx == $i)} {
        # ok
    } else {
        set excel::wbActIdx $i
        set excel::wbAct $excel::wbObj($i)
        
        puts "== update wb: [$excel::wbAct Name] active"
    }
    
    # 更新ActiveSheet
    set wsAct [$excel::wbAct ActiveSheet]
    if {$wbAct == ""} {
        set excel::wsActIdx ""
        set excel::wsAct ""
        return
    }
    
    set ix [excel::get_ws_ix $excel::wbAct [$wsAct Name]]
    set j [lindex $excel::wsIdxList($i) $ix]
    
    if {[info exist excel::wsActIdx] && ($excel::wsActIdx == $j)} {
        # ok
    } else {
        set excel::wsActIdx $j
        set excel::wsAct $excel::wsObj($i,$j)
        
        puts "    == update ws: [$excel::wsAct Name] active"
    }
}

proc excel::open {filename args} {
    if {$filename == ""} {
        if {[catch {set wbObj [$excel::wbsObj Add]}]} {
            puts "**** ERROR **** \[excel::open -> can not add: $filename\]"
            return
        }
        
        excel::add_wb $wbObj
        
    } elseif {![file exist $filename]} {
        puts "**** ERROR **** \[excel::open -> bad filename: $filename\]"
        return
        
    } else {
        set fn [file tail $filename]
        
        # 检查是否已经打开
        set ix [excel::get_wb_ix $fn]
        if {$ix > -1} {
            # ok
            return
        }
        
        # 未打开
        if {[catch {set wbObj [$excel::wbsObj Open $filename]}]} {
            puts "**** ERROR **** \[excel::open -> can not open: $filename\]"
            return
        }
        
        excel::add_wb $wbObj
    }

    excel::update
}

proc excel::close {filename args} {
}

