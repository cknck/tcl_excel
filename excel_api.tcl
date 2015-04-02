if {0} {
package require tcom
set fn {d:/temp/tst.xls}

#set appObj [::tcom::ref createobject "Excel.Application"]
#$appObj Visible 1
set appObj [::tcom::ref getactiveobj "Excel.Application"]

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

proc show {r_arr} {
    upvar $r_arr arr
    
    puts \t=================
    for {set i 0} {$i < [llength [array names arr]]} {incr i} {
        puts \t$arr($i)
    }
    puts \t=================
}

proc AutoLevel {r_ws cell args} {
    upvar $r_ws ws
    
    variable arr
    array set arr {}
    
    eval {GetLevelArray $ws arr $cell} $args
    
    set rows  [$ws Rows]
    
    for {set i 1} {$i < [llength [array names arr]]} {incr i} {
        foreach rec $arr($i) {
            foreach {r1 r2} $rec {}
            for {set r $r1} {$r <= $r2} {incr r} {
                [$rows Item [expr $r]] Group
            }
        }
    }
}

proc GetLevelArray {ws r_arr cell args} {
    upvar $r_arr arr
    
    array unset arr
    
    foreach {cell_r cell_c} [GetCellRowCol $cell] {}
    
    set count 1
    set collist ""
    
    while {[llength $args]} {
        set key [string toupper [lindex $args 0]]
        switch -- $key {
            -COUNT {
                set count [lindex $args 1]
                set args [lreplace $args 0 1]
            }
            -COL {
                set collist [lindex $args 1]
                set args [lreplace $args 0 1]
            }
            default {
                set args [lreplace $args 0 0]
            }
        }
    }
    
    set col_list $cell_c
    foreach col $collist {
        if {[regexp {\D+} $col match]} {
            set col [Alpha2Col $col]
        }
        lappend col_list $col
    }
    
    set col_list_len [llength $col_list]
    
    foreach {r1 c1 r2 c2} [GetSheetRowCol $ws] {}
    
    for {set idx 0;set ci 0} {$ci < [llength $col_list]} {incr ci;incr idx} {
        set c [lindex $col_list $ci]
        
        set rec ""
        set pre_val ""
        for {set r $cell_r} {$r <= $r2} {incr r} {
            set val [GetValue $ws "$r $c"]
            
            if {$val != ""} {
                if {$rec == ""} {
                    # 记录为空，新记录开始
                    set rec $r
                } elseif {$pre_val == ""} {
                    lappend rec [expr $r-1]
                    lappend arr($idx) $rec
                    set rec $r
                }
            }
            
            set pre_val $val
        }
        
        lappend rec [expr $r-1]
        lappend arr($idx) $rec
        set rec $r
    }
    
    # 初始化比较记录
    set rec_idxlist ""
    for {set i 0} {$i < $col_list_len} {incr i} {
        lappend rec_idxlist 0
        lappend rec_list [lindex $arr($i) 0]
    }
    
    # 逐行比较
    for {set r $cell_r} {$r < $r2} {incr r} {
        # row 以上的记录为无效
        set on_row_list [ReclistOnRow $rec_list $r]
#puts "===== row: $r ==1== \[$rec_list\] : \[$on_row_list\]"

        set rec_change 0
        for {set i 1} {$i < $col_list_len} {incr i} {
            if {[lindex $on_row_list $i] == -1} {
                # record is above row, delete
                set rec_idx [lindex $rec_idxlist $i]
                set arr($i) [lreplace $arr($i) $rec_idx $rec_idx]
                
                # 改变比较记录
                set rec_list [lreplace $rec_list $i $i [lindex $arr($i) $rec_idx]]
                
                set rec_change 1
            }
        }
        if {$rec_change} {
            incr r -1
            continue
        }
        
        # 从最高分级的开始，如果该行为记录顶，返回该分级编号
        set coli [ReclistUpOnRow $rec_list $r]
#puts "===== row: $r ==2== \[$rec_list\] : $coli"

        if {$coli != -1} {
            # 低级记录以该行的下一行开始
            for {set i [expr $coli+1]} {$i < $col_list_len} {incr i} {
                set rec_idx [lindex $rec_idxlist $i]
                set rec [lindex $arr($i) $rec_idx]
                
                foreach {rec_r1 rec_r2} $rec {}
                
                if {$rec_r2 < [expr $r+1]} {
                    # 该记录无效，删除
                    set arr($i) [lreplace $arr($i) $rec_idx $rec_idx]
                    # 取下一个记录
                    set rec [lindex $arr($i) $rec_idx]
                } elseif {$rec_r1 <= $r} {
                    # 改写记录
                    set rec [lreplace $rec 0 0 [expr $r+1]]
                    set arr($i) [lreplace $arr($i) $rec_idx $rec_idx $rec]
                }
                
                # 比较记录更新
                set rec_list [lreplace $rec_list $i $i $rec]
            }
        }
        
        # 从最高分级的开始，如果该行为记录底，返回该分级编号
        set coli [ReclistDownOnRow $rec_list $r]
#puts "===== row: $r ==3== \[$rec_list\] : $coli"

        if {$coli != -1} {
            # 低级记录以该行结束
            for {set i $coli} {$i < $col_list_len} {incr i} {
                set rec_idx [lindex $rec_idxlist $i]
                set rec [lindex $arr($i) $rec_idx]
                
                foreach {rec_r1 rec_r2} $rec {}
                
                # 改写记录
                set rec [lreplace $rec 1 1 $r]
                set arr($i) [lreplace $arr($i) $rec_idx $rec_idx $rec]
                
                # 记录编号增加
                incr rec_idx
                
                if {($rec_r2 >= [expr $r+2]) && ([GetValue $ws "[expr $r+2] [lindex $col_list $i]"] != "")} {
                    # 插入分拆的记录
                    set arr($i) [linsert $arr($i) $rec_idx "[expr $r+2] $rec_r2"]
                }
                
                set rec_idxlist [lreplace $rec_idxlist $i $i $rec_idx]
                # 改变比较记录
                set rec_list [lreplace $rec_list $i $i [lindex $arr($i) $rec_idx]]
                
                #puts ========>$rec_list
            }
        }
    }

    return $col_list
}

proc ReclistDownOnRow {rec_list r} {
    set ret -1
    
    for {set i 0} {$i < [llength $rec_list]} {incr i} {
        foreach {r1 r2} [lindex $rec_list $i] {}
        if {$r == $r2} {
            set ret $i
            break
        }
    }
    
    return $ret
}

proc RecDownOnRow {rec r} {
    set ret 0
    
    foreach {r1 r2} $rec {}
    if {$r == $r2} {
        set ret 1
    }
    
    return $ret
}

proc ReclistUpOnRow {rec_list r} {
    set ret -1
    
    for {set i 0} {$i < [llength $rec_list]} {incr i} {
        foreach {r1 r2} [lindex $rec_list $i] {}
        if {$r == $r1} {
            set ret $i
            break
        }
    }
    
    return $ret
}

proc RecUpOnRow {rec r} {
    set ret 0
    
    foreach {r1 r2} $rec {}
    if {$r == $r1} {
        set ret 1
    }
    
    return $ret
}

proc ReclistOnRow {rec_list r} {
    set on_row_list ""
    foreach rec $rec_list {
        lappend on_row_list [RecOnRow $rec $r]
    }
    
    return $on_row_list
}

proc RecOnRow {rec r} {
    foreach {r1 r2} $rec {}
    
    if {$r2 < $r} {
        # record is above row
        return -1
    } elseif {$r1 > $r} {
        # record is below row
        return 1
    } else {
        # $r1 <= $r <= $r2, record is on row
        return 0
    }
}

proc GetCellRowCol {cell} {
    if {[regexp {^(\D+)(\d+)$} $cell match sub1 sub2]} {
        set r1 $sub2
        set c1 [Alpha2Col $sub1]
    } else {
        foreach {r1 c1} $cell {}
    }
    
    return "$r1 $c1"
}

proc GetValue {ws cell args} {
    
    foreach {r1 c1} [GetCellRowCol $cell] {}
    
    set cells [$ws Cells]
    set cell  [$cells Item [expr $r1] [expr $c1]]
    
    return [$cell Value]
}

# 取WorkSheet使用区域的对角线坐标
proc GetSheetRowCol {ws} {
    
    set range [$ws UsedRange]
    
    return [GetRangeRowCol $range]
}

# 取Range的对角线坐标
proc GetRangeRowCol {range} {
    
    set rows [$range Rows]
    set cols [$range Columns]
    
    set r1 [$range Row]
    set c1 [$range Column]
    
    set range2 [$range Item [$rows Count] [$cols Count]]
    set r2 [$range2 Row]
    set c2 [$range2 Column]
    
    return [list $r1 $c1 $r2 $c2]
}

#source d:/temp/test.tcl