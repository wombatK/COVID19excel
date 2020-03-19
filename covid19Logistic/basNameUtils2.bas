Attribute VB_Name = "basNameUtils2"
'basNameUtils2 - based on basNameUtils with additional new name copy/replace

'(c) K Duffy Mar 2020
'   Licensed under the Apache License, Version 2.0 (the "License");
'   you may not use this file except in compliance with the License.
'   You may obtain a copy of the License at
'
'       http://www.apache.org/licenses/LICENSE-2.0
'
'   Unless required by applicable law or agreed to in writing, software
'   distributed under the License is distributed on an "AS IS" BASIS,
'   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
'   See the License for the specific language governing permissions and
'   limitations under the License.
'------------------------------------

Sub MakeNameFromExample()
    Dim strNameSource As String
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    Dim nmsAll As Names
    Set nmsAll = wb.Names
    Dim nmNext As Name
    Dim strSheet As String
    strSheet = ActiveSheet.Name
    For Each nmNext In nmsAll
        If InStr(1, nmNext.Name, "_") = 0 And InStr(1, nmNext.Name, "solver") = 0 Then
            Dim iLen As Integer
            iLen = WorksheetFunction.Max(0, 20 - Len(nmNext.Name))
            Debug.Print nmNext.Name, String(iLen, " "), nmNext.RefersTo, nmNext.NameLocal
        End If
    Next nmNext
    
    For Each nmNext In nmsAll
        Debug.Print nmNext.Name, nmsAll.Count
        If InStr(1, nmNext.Name, "_") = 0 And InStr(1, nmNext.Name, "solver") = 0 Then
            If InStr(1, nmNext.Name, strSheet & "!") > 0 Then Exit For
        End If
    Next nmNext
    
    strNameSource = InputBox("Enter name to copy formula from", ActiveSheet.Name & "!" & x, "")
End Sub

Public Sub AddSheetWithNameCheckIfExists()
    'http://www.mindspring.com/~tflynn/excelvba3.html#AddSheet
    Dim ws As Worksheet
    Dim newSheetName As String
    'newSheetName = Sheets(1).Range("A1")   '   Substitute your range here
    newSheetName = Trim(InputBox("Enter Name for new Sheet ", "Name of New Sheet"))
    For Each ws In Worksheets
        If ws.Name = newSheetName Or newSheetName = "" Or IsNumeric(newSheetName) Then
            MsgBox "Sheet already exists or name is invalid", vbInformation
            Exit Sub
        End If
    Next
    Sheets.Add Type:="Worksheet"
    With ActiveSheet
        .Move After:=Worksheets(Worksheets.Count)
        .Name = newSheetName
    End With
End Sub

' To add a range name based on a selection
Public Sub AddNameLocal()
    Dim strName As String
    
    If Selection.Rows.Count < 1 Or Selection.Columns.Count < 1 Then
        MsgBox "Nothing selected - First Select a Range that Name applies to", vbOKOnly, "Name Range Failed"
        Exit Sub
    End If
    
    strName = Trim(InputBox("Enter Name for Local Range ", "Name Selected Range"))
    'doesn't handle spaces in sheet name ...
    If Len(strName) > 0 Then
        strName = "'" & ActiveSheet.Name & "'!" & strName
        ' Range("$A$1:$L$12,$A$38:$L$78,$A$86:$E$88,$A$144:$L$157,$A$179:$L$194,$A$203:$L$243,$A$295:$L$302,$A$307:$L$317,$A$363:$L$383,$M$150:$W$167,$M$170:$W$190,$M$200:$W$212,$S$216:$U$220,$M$246:$V$249,$M$256:$S$270,$M$271:$V$274,$M$281:$S$298,$M$299:$W$303").Select
        ' add doesn't have different arguments for range (eg if selection is noncontiguous
        ActiveSheet.Names.Add Name:=strName, RefersTo:=Selection
    End If
End Sub

Public Sub CopyNameToWks()
    ' Copy a local worksheet name to another worksheet
    Dim strNameIn As String
    Dim strNameOut As String
    Dim strSheetIn As String
    Dim strSheetOut As String
    Dim ws As Worksheet
    
    strNameIn = InputBox("Enter Name of Local range", "Name")
    strSheetIn = ActiveSheet.Name
    Dim strA() As String
    Dim strB() As String
    
    Dim strAddress As String
    Dim bFound As Boolean
    
    Dim iX As Integer
    iX = 1
    strNameOut = strNameIn
    
    strNameIn = ActiveSheet.Name & "!" & strNameIn
    While ActiveSheet.Names(iX).Name <> strNameIn And iX < ActiveSheet.Names.Count
        Debug.Print ActiveSheet.Names(iX).Name, ActiveSheet.Name
        iX = iX + 1
    Wend
    
    If iX >= ActiveSheet.Names.Count Then
        MsgBox "No such name: " & strNameIn
        Exit Sub
    End If
    
    strAddress = ActiveSheet.Names(strNameIn)
    If Len(strAddress) < 1 Then Exit Sub
    strA = Split(strAddress, strSheetIn & "!")
    strSheetOut = InputBox("Enter Destination Worksheet Name", "Worksheet")
    
    
    iX = 1
    While ActiveWorkbook.Worksheets(iX).Name <> strSheetOut And iX < ActiveWorkbook.Worksheets.Count
        iX = iX + 1
    Wend
    If iX >= ActiveWorkbook.Worksheets.Count Then
        MsgBox "No Such Worksheet: " & strSheetOut
        Exit Sub
    End If
    
    strNameOut = InputBox("Enter Name in Destination Worksheet " & strSheetOut, "Name", strNameOut)
    
    Dim strAddressO As String
    
    strAddressO = Join(strA, strSheetOut & "!")
    strNameOut = strSheetOut & "!" & strNameOut
    Application.Worksheets(strSheetOut).Names.Add Name:=strNameOut, RefersTo:=Range(strAddressO)
    
End Sub

Public Sub ListNamesOld()
    ' lists names in worksheet to next n rows of currently selected column
    Dim nms As Names
    Dim R As Integer
    Dim sA() As String
    
    Set nms = ActiveWorkbook.Names
    'Set wks = ActiveSheet
    Dim rngLastCell As Range
    'Set rngLastCell = Range("$A$1")
    
    On Error Resume Next
    Set rngLastCell = LastCell(ActiveSheet)
    If rngLastCell = Null Then
        Set rngLastCell = Range("A1:A1")
    End If
    On Error GoTo 0
    
    MsgBox "Last row = " & rngLastCell.Row & " col " & rngLastCell.Column
    Application.GoTo rngLastCell
    Dim rngOut As Range
    ' set to end of worksheet
    Set rngOut = ActiveCell.Offset(2, -rngLastCell.Column + 1)
    
    rngOut.Select
    rngOut.Offset(0, 4).Value = "WorkSheet Names"
    rngOut.Offset(0, 8).Value = "RefersTo"
    rngOut.Offset(0, 12).Value = "RefersToLocal"
    
    Dim strBad As String
    
    For R = 1 To nms.Count
        ' Debug.Print nms(R).Name; " "; nms(R).RefersTo; " "; nms(R).RefersToLocal
        sA = Split(nms(R).Name, "!")
        
        If UBound(sA) > 0 Then
            rngOut.Offset(R, 2).Value = sA(1)
            rngOut.Offset(R, 0).Value = sA(0)
        Else
            rngOut.Offset(R, 2).Value = sA(0)
        End If
        
        
        rngOut.Offset(R + 1, 4).Value = nms(R).Name
        rngOut.Offset(R + 1, 8).Value = "'" & nms(R).RefersTo
        rngOut.Offset(R + 1, 12).Value = "'" & nms(R).RefersToLocal
        If InStr(1, rngOut.Offset(R + 1, 8).Value, "#REF") > 0 Then
            strBad = strBad & vbCrLf & nms(R).Name & vbTab & rngOut.Offset(R + 1, 8).Value
        End If
        
    Next
    
    MsgBox "Listed " & nms.Count & " names" & vbCrLf & strBad, , "All Done"
    
End Sub

Public Sub SelectChartSeries()
    Dim srs As Series
    Dim iSer As Integer
    Dim strMsg As String
    Dim lErrNumber As Long
    Dim strEmsg As String
    
    On Error Resume Next
    iSer = ActiveChart.SeriesCollection.Count
    
    If iSer < 1 Or Err.Number > 0 Then
        MsgBox "Select Chart Series failed, no chart active", vbCritical, "Select Chart Series Error"
        Exit Sub
    End If
    On Error GoTo 0
    
    strMsg = vbCrLf & vbCrLf & "Primary Axis" & vbCrLf
    iOrder = 1
    On Error Resume Next
    Dim intAxisGroup As Integer
    
    For iSer = 1 To ActiveChart.SeriesCollection.Count
        intAxisGroup = ActiveChart.SeriesCollection(iSer).AxisGroup
        If Err.Number = 0 Then
            If intAxisGroup = 1 Then
                strMsg = strMsg & vbCrLf & Format(iSer, "#0") & vbTab & ActiveChart.SeriesCollection(iSer).Name & vbTab & "(axis " & ActiveChart.SeriesCollection(iSer).AxisGroup & ", order " & iOrder & ")"
                iOrder = iOrder + 1
            End If
        Else
            lErrNumber = Err.Number
            strEmsg = Err.Number & " - " & Err.Description
            
            If Err.Number = 1004 Then
                strMsg = strMsg & vbCrLf & "most likely series is not visible, #NA etc.,. - ignore it"
                ' this is a bug in Excel 97, and maybe 2000
                
            End If
            MsgBox strEmsg, vbOKOnly, "Error Accessing Series Definitions"
            strMsg = strMsg & vbCrLf & Format(iSer, "#0") & vbTab & " !!inaccessible!!" & ", order " & iOrder & ")"
            Err.Clear
        End If
    Next iSer
    
    strMsg = strMsg & vbCrLf & vbCrLf & "Secondary Axis" & vbCrLf
    iOrder = 1
    
    For iSer = 1 To ActiveChart.SeriesCollection.Count
        intAxisGroup = ActiveChart.SeriesCollection(iSer).AxisGroup
        If Err.Number = 0 Then
            If intAxisGroup = 2 Then
                strMsg = strMsg & vbCrLf & Format(iSer, "#0") & vbTab & ActiveChart.SeriesCollection(iSer).Name & vbTab & "(axis " & ActiveChart.SeriesCollection(iSer).AxisGroup & ", order " & iOrder & ")"
                iOrder = iOrder + 1
            End If
        Else
            lErrNumber = Err.Number
            strEmsg = Err.Number & " - " & Err.Description
            
            If Err.Number = 1004 Then
                strMsg = strMsg & vbCrLf & "most likely series is not visible, #NA etc.,. - ignore it"
                ' this is a bug in Excel 97, and maybe 2000
                
            End If
            MsgBox strEmsg, vbOKOnly, "Error Accessing Series Definitions"
            strMsg = strMsg & vbCrLf & Format(iSer, "#0") & vbTab & " !!inaccessible!!" & ", order " & iOrder & ")"
            Err.Clear
        End If
    Next iSer
    
    'MsgBox "Series are " & vbCrLf & strMsg
    On Error Resume Next
    Dim varIn As Variant
    
    varIn = InputBox("Select Series from " & ActiveChart.SeriesCollection.Count & " Series" & strMsg, "Series ?")
    If Err.Number > 0 Then
        Exit Sub
    End If
    If Not IsNumeric(varIn) Then
        MsgBox "Number must be specified", vbExclamation, "Bad Series Selection"
        Exit Sub
    End If
    iSer = Val(varIn)
    On Error GoTo 0
    'ActiveChart.SeriesCollection(iSer).Select
    'Dim iOrder As Integer
    Dim strSerName As String
    strSerName = ActiveChart.SeriesCollection(iSer).Name
    
    iOrder = InputBox("Select Order, Currently " & ActiveChart.SeriesCollection(iSer).PlotOrder, "Order", 1)
    ActiveChart.SeriesCollection(strSerName).PlotOrder = iOrder
    ActiveChart.SeriesCollection(strSerName).Select
    Exit Sub
Err_Series:
    
    
    
    lErrNumber = Err.Number
    strEmsg = Err.Number & " - " & Err.Description
    
    If Err.Number = 1004 Then
        strMsg = strMsg & vbCrLf & "most likely series is not visible, #NA etc.,. - ignore it"
        ' this is a bug in Excel 97, and maybe 2000
    End If
    MsgBox strEmsg, vbOKOnly, "Error Accessing Series Definitions"
    Resume Next
End Sub


Public Sub ListAndReplaceInNames()
    ' lists names in worksheet to next n rows of currently selected column
    ' and replaces substrings in Name definitions
    ' for use after copy worksheet to new workbook
    
    Dim nms As Names
    Dim R As Integer
    Dim sA() As String
    
    
    Set nms = ActiveWorkbook.Names
    'Set wks = ActiveSheet
    Dim rngLastCell As Range
    'Set rngLastCell = Range("$A$1")
    
    On Error Resume Next
    Set rngLastCell = LastCell(ActiveSheet)
    If rngLastCell = Null Then
        Set rngLastCell = Range("A1:A1")
    End If
    On Error GoTo 0
    
    MsgBox "Last row = " & rngLastCell.Row & " col " & rngLastCell.Column
    Application.GoTo rngLastCell
    Dim rngOut As Range
    ' set to end of worksheet
    Set rngOut = ActiveCell.Offset(2, -rngLastCell.Column + 1)
    
    rngOut.Select
    rngOut.Offset(0, 1).Value = "WS Scope"
    rngOut.Offset(0, 1).Value = "Name"
    rngOut.Offset(0, 4).Value = "WorkSheet Names"
    rngOut.Offset(0, 8).Value = "RefersTo"
    rngOut.Offset(0, 12).Value = "RefersToLocal"
    rngOut.Offset(0, 14).Value = "Use Count"
    
    Dim strBad As String
    
    Dim bCheckUse As Boolean
    Dim iCheckOpt As Integer
    iCheckOpt = InputBox("Do you want to check use count for each defined name" & vbCrLf & "0. Check none" & vbCrLf & "1. Check only names without underscore in 1st Char (also excludes solver names)" & vbCrLf & "2. Check all names", "Check Use", 1)
    
    
    bExclUScore = False
    
    Dim strFind As String
    
    For R = 1 To nms.Count
        ' Debug.Print nms(R).Name; " "; nms(R).RefersTo; " "; nms(R).RefersToLocal
        sA = Split(nms(R).Name, "!")
        ' If R = nms.count Then Stop
        
        If Len(Trim(sA(0))) < 1 Then
            ' Stop
            ' nms(R).Delete
            ' R = R - 1
        Else
            
            Dim strNsimple As String
            
            If UBound(sA) > 0 Then
                strNsimple = sA(1)
                rngOut.Offset(R, 0).Value = "'" & sA(0)
            Else
                strNsimple = sA(0)
            End If
            
            rngOut.Offset(R, 2).Value = strNsimple
            
            rngOut.Offset(1 * R, 4).Value = nms(R).Name
            rngOut.Offset(1 * R, 8).Value = "'" & nms(R).RefersTo
            rngOut.Offset(1 * R, 12).Value = "'" & nms(R).RefersToLocal
            If InStr(1, rngOut.Offset(R + 1, 8).Value, "#REF") > 0 Then
                strBad = strBad & vbCrLf & nms(R).Name & vbTab & rngOut.Offset(R + 1, 8).Value
            Else
                If iCheckOpt > 0 Then
                    If iCheckOpt > 1 Or Left(strNsimple, 1) <> "_" Then
                        If iCheckOpt <> 1 Or InStr(1, strNsimple, "solver_") = 0 Then
                            
                            Dim strUses As String
                            Application.StatusBar = "Checking for uses of name " & strNsimple
                            
                            strUses = FindInFormulas(strNsimple, ActiveSheet.Name)
                            Dim strU() As String
                            strU = Split(strUses, vbLf)
                            If UBound(strU) > 1 Then
                                rngOut.Offset(1 * R, 14).Value = strU(0)
                                rngOut.Offset(1 * R, 15).Value = strU(1)
                            End If
                        End If
                    End If
                    
                End If
            End If
            
            Dim iB As Integer
            Dim iC As Integer
            
            iB = InStr(1, nms(R).RefersTo, "[")
            If iB > 0 Then
                
                iC = InStr(1, nms(R).RefersTo, "]")
                If iC > 0 Then
                    strFind = Mid(nms(R).RefersTo, iB, iC - iB + 1)
                End If
            End If
        End If ' not zero length name
    Next
    
    
    
    strFind = InputBox("String to Find in Names, * for any [] delimited string, otherwise xxx to find nothing or empty string to skip (not implemented)", "Target", strFind)
    Dim bReplace As Boolean
    If Len(strFind) > 0 Then
        
        bReplace = MsgBox("Do you want to replace out external references in names (prompting for each replace to allow selection)", vbYesNo, "Eliminate Ext Links")
        Dim iReplace As Integer
        
        
        For R = 1 To nms.Count
            Dim strRefersTo As String
            strRefersTo = nms(R).RefersTo
            iB = InStr(1, nms(R).RefersTo, "[")
            If iB > 0 Then
                
                iC = InStr(1, nms(R).RefersTo, "]")
                If iC > 0 Then
                    'strFind = Mid(nms(R).RefersTo, iB, iC - iB + 1)
                End If
            End If
            
            If strFind = "*" And iB > 0 And iC > 0 Then
                Debug.Print Mid(strRefersTo, iB, iC - iB + 1)
                'Stop
                Dim strFrom As String
                strFrom = Mid(nms(R).RefersTo, iB, iC - iB + 1)
                strRefersTo = Replace(strRefersTo, strFrom, "")
                If bReplace Then
                    If vbYes = MsgBox("Confirm Replace in name " & nms(R).Name & vbCrLf & "RefersTo Was " & nms(R).RefersTo & vbCrLf & "Becomes " & strRefersTo & vbCrLf, vbYesNo, "Replace ?") Then
                        
                        iReplace = iReplace + 1
                        nms(R).RefersTo = strRefersTo
                        
                    End If
                End If
                
                
            Else
            End If
            
            strRefersTo = nms(R).RefersToLocal
            Dim iD As Integer
            
            iB = InStr(1, strRefersTo, "[")
            If iB > 0 Then
                
                iC = InStr(1, strRefersTo, "]")
                If iC > 0 Then
                    ' strFind = Mid(nms(R).RefersToLocal, iB, iC - iB + 1)
                End If
                iD = InStr(1, Left(strRefersTo, iB - 1), "\")
                If iD < 1 Then
                    ' probably structured table reference, not an external ref
                    iB = 0
                End If
                
            End If
            
            
            If iB > 0 Then
                strFName = Replace(Replace(Replace(Left(strRefersTo, iC - 1), "[", ""), "'", ""), "=", "") 'replace out the [ and '
                
                On Error Resume Next
                Dim strFPath As String
                strFPath = Dir(strFName, vbNormal)
                If Len(strFPath) = 0 Then
                    Debug.Print "***Warning*** Name " & nms(R).Name & " Refers to " & strRefersTo & " but file " & strFName & " does not exist"
                    MsgBox "Warning: Name " & nms(R).Name & " Refers to " & strRefersTo & vbCrLf & " but file " & strFName & " does not exist", vbOKOnly
                    
                End If
                On Error GoTo 0
            End If
            
            If strFind = "*" And iB > 0 And iC > 0 Then
                
                
                Debug.Print Mid(strRefersTo, iB, iC - iB + 1)
                'Stop
                strFrom = Mid(strRefersTo, iB, iC - iB + 1)
                strRefersTo = Replace(strRefersTo, strFrom, "")
                Debug.Print nms(R).Name & " == " & strFrom & vbTab & nms(R).RefersToLocal
                
                If bReplace Then
                    If vbYes = MsgBox("Confirm Replace in name " & nms(R).Name & vbCrLf & "RefersTo Was " & nms(R).RefersTo & vbCrLf & "Becomes " & strRefersTo & vbCrLf, vbYesNo, "Replace ?") Then
                        
                        iReplace = iReplace + 1
                        nms(R).RefersTo = strRefersTo
                        
                    End If
                    
                End If
                
            Else
            End If
            
        Next R
        
    End If
    
    MsgBox "Listed " & nms.Count & " names" & vbCrLf & strBad & vbCrLf & "Replaced in " & iReplace, , "All Done"
    
End Sub


Function LastCell(ws As Worksheet) As Range
    Dim LastRow&, LastCol%
    
    ' Error-handling is here in case there is not any
    ' data in the worksheet
    'http://www.beyondtechnology.com/geeks012.shtml
    
    On Error Resume Next
    
    With ws
        
        ' Find the last real row
        
        LastRow& = .Cells.Find(What:="*", _
        SearchDirection:=xlPrevious, _
        SearchOrder:=xlByRows).Row
        
        ' Find the last real column
        
        LastCol% = .Cells.Find(What:="*", _
        SearchDirection:=xlPrevious, _
        SearchOrder:=xlByColumns).Column
        
    End With
    
    ' Finally, initialize a Range object variable for
    ' the last populated row.
    
    Set LastCell = ws.Cells(LastRow&, LastCol%)
    
End Function

Sub TestLastCell()
    MsgBox "Last Cell/Column is " & LastCell(ActiveSheet).Column, vbCritical
End Sub

Function FindInFormulas(strTarget As String, strExclWs As String) As String
    'returns a string of count of cells found, first cell found address, and formulas using giving strTarget (typically a defined name)
    'strExclWs is a comma separated list of worksheet names to exclude from search,
    'any sheet named "Link List" is also excluded
    
    Dim sht As Worksheet
    Dim shtName
    Dim myRng As Range
    Dim newRng As Range
    Dim c As Range
    Dim rngFind As Range
    Dim iCountSht As Integer
    Dim iCountBk As Integer
    
    
    Application.ScreenUpdating = False
    Dim strFindBk As String
    Dim iSheetsFound As Integer
    
    strExclWs = strExclWs & ", " & "Link List"
    
    For Each sht In ActiveWorkbook.Worksheets 'loop through the sheets in the workbook
        If InStr(1, strExclWs, sht.Name) < 1 Then 'exclude the sheet just created
            Set myRng = sht.UsedRange 'limit the search to the UsedRange
            On Error Resume Next 'in case there are no formulas
            Set newRng = myRng.SpecialCells(xlCellTypeFormulas) 'use SpecialCells to reduce looping further
            If Not newRng Is Nothing Then
                iCountSht = 0
                Set rngFind = newRng.Find(What:=strTarget, LookIn:=xlFormulas, LookAt:=xlPart, MatchCase:=False)
                If rngFind Is Nothing Then
                    ' does not occur on this worksheet
                Else
                    strFind = "; " & rngFind.Address & "; {" & rngFind.Formula & "} "
                    iSheetsFound = iSheetsFound + 1
                    Dim rngNext As Range
                    Set rngNext = rngFind
                    Do
                        Set rngNext = newRng.FindNext(rngNext)
                        iCountSht = iCountSht + 1
                    Loop Until rngNext Is Nothing Or rngNext.Address = rngFind.Address Or iCountSht > 32000
                    
                    
                    strFind = sht.Name & "; " & iCountSht & strFind
                    iCountBk = iCountBk + iCountSht
                End If
                '  Debug.Print sht.name, iCountSht, strFind
            End If
            
            
        End If
        If Len(strFind) > 0 Then
            strFindBk = strFindBk & vbLf & strFind
            strFind = ""
        End If
    Next sht
    strFindBk = "Found " & iCountBk & " in " & iSheetsFound & " sheets " & strFindBk
    FindInFormulas = strFindBk
    Application.ScreenUpdating = True
End Function


Public Sub ListNoSolverNames()
    
    
    ' lists names in worksheet to next n rows of currently selected column
    ' and replaces substrings in Name definitions
    ' for use after copy worksheet to new workbook
    
    Dim nms As Names
    Dim R As Integer
    Dim sA() As String
    
    
    Set nms = ActiveWorkbook.Names
    'Set wks = ActiveSheet
    Dim rngLastCell As Range
    'Set rngLastCell = Range("$A$1")
    
    On Error Resume Next
    Set rngLastCell = LastCell(ActiveSheet)
    Dim iRow As Integer
    
    If rngLastCell Is Nothing Then
        Set rngLastCell = Range("A1:A1")
        iRow = 0
    Else
        iRow = 1
    End If
    On Error GoTo 0
    
    MsgBox "Last row = " & rngLastCell.Row & " col " & rngLastCell.Column
    Application.GoTo rngLastCell
    Dim rngOut As Range
    ' set to end of worksheet
    Set rngOut = ActiveCell.Offset(iRow, -rngLastCell.Column + 1)
    
    rngOut.Select
    rngOut.Offset(0, 0) = "WS Scope"
    'rngOut.Offset(0, 1).Value = "WS Scope"
    rngOut.Offset(0, 1).Value = "Name"
    rngOut.Offset(0, 2).Value = "WorkSheet Names"
    rngOut.Offset(0, 3).Value = "RefersTo"
    rngOut.Offset(0, 4).Value = "RefersToLocal"
    rngOut.Offset(0, 5).Value = "Use Count"
    
    Dim strBad As String
    
    Dim bCheckUse As Boolean
    Dim iCheckOpt As Integer
    
    iCheckOpt = InputBox("Do you want to check use count for each defined name" & vbCrLf & "0. Check none" & vbCrLf & "1. Check only names without underscore in 1st Char (also excludes solver names)" & vbCrLf & "2. Check all names", "Check Use", 1)
    
    
    bExclUScore = False
    
    Dim strFind As String
    Dim nmNext As Name
    
    Dim iRowOut As Integer
    iRowOut = 1
    
    For R = 1 To nms.Count
        ' Debug.Print nms(R).Name; " "; nms(R).RefersTo; " "; nms(R).RefersToLocal
        Set nmNext = nms(R)
        Debug.Print R, nmNext.Name
        'If R = 95 Then Stop
        
        If InStr(1, nmNext.Name, "_") = 0 And InStr(1, nmNext.Name, "solver") = 0 Then
            iRowOut = iRowOut + 1
            sA = Split(nms(R).Name, "!")
            ' If R = nms.count Then Stop
            
            If Len(Trim(sA(0))) < 1 Then
                ' Stop
                ' nms(R).Delete
                ' R = R - 1
            Else
                
                Dim strNsimple As String
                
                If UBound(sA) > 0 Then
                    strNsimple = sA(1)
                    rngOut.Offset(iRowOut - 1, 0).Value = "'" & sA(0)
                Else
                    strNsimple = sA(0)
                End If
                
                rngOut.Offset(iRowOut - 1, 1).Value = strNsimple
                
                rngOut.Offset(iRowOut - 1, 2).Value = nms(R).Name
                rngOut.Offset(iRowOut - 1, 3).Value = "'" & nms(R).RefersTo
                rngOut.Offset(iRowOut - 1, 4).Value = "'" & nms(R).RefersToLocal
                If InStr(1, rngOut.Offset(iRowOut - 1, 3).Value, "#REF") > 0 Then
                    strBad = strBad & vbCrLf & nms(R).Name & vbTab & rngOut.Offset(iRowOut - 1, 3).Value
                Else
                    If iCheckOpt > 0 Then
                        If iCheckOpt > 1 Or Left(strNsimple, 1) <> "_" Then
                            If iCheckOpt <> 1 Or InStr(1, strNsimple, "solver_") = 0 Then
                                
                                Dim strUses As String
                                Application.StatusBar = "Checking for uses of name " & strNsimple
                                
                                strUses = FindInFormulas(strNsimple, ActiveSheet.Name)
                                Dim strU() As String
                                strU = Split(strUses, vbLf)
                                If UBound(strU) > 1 Then
                                    rngOut.Offset(iRowOut - 1, 5).Value = strU(0)
                                    rngOut.Offset(iRowOut - 1, 6).Value = strU(1)
                                End If
                            End If
                        End If
                        
                    End If
                End If
                
                Dim iB As Integer
                Dim iC As Integer
                
                iB = InStr(1, nms(R).RefersTo, "[")
                If iB > 0 Then
                    
                    iC = InStr(1, nms(R).RefersTo, "]")
                    If iC > 0 Then
                        strFind = Mid(nms(R).RefersTo, iB, iC - iB + 1)
                    End If
                End If
            End If ' not zero length name
        End If
    Next
    MsgBox "ListNoSolverNames Done", vbOKOnly, "Done"
    
End Sub

Sub ReplaceInNames()
    Dim iRow As Integer
    iRow = 2
    Dim ws As Worksheet
    Set ws = Sheets("Names")
    Dim nms As Names
    Dim nmNext As Name
    
    
    Set nms = ActiveWorkbook.Names
    Dim iColOldString As Integer
    iColOldString = 7
    Dim iColRefersTo As Integer
    iColRefersTo = 4
    
    
    While Not IsEmpty(ws.Cells(iRow, 2))
        Dim sName As String
        sName = ws.Cells(iRow, 3)  'name without worksheet qualifier
        Set nmNext = nms(sName)
        If iRow = 26 Then Stop
        If Not IsEmpty(ws.Cells(iRow, iColOldString)) Then
            Dim sOld As String
            sOld = ws.Cells(iRow, iColOldString)
            If Left(sOld, 1) = "'" Then
            sOld = Mid(sOld, 2, 2000)
        End If
        
        Dim sNew As String
        sNew = ws.Cells(iRow, iColOldString + 1)
        If Left(sNew, 1) = "'" Then
        sNew = Mid(sNew, 2, 2000)
    End If
    
    Dim sRefersTo As String
    Dim sRefersToLocal As String
    sRefersTo = ws.Cells(iRow, iColRefersTo)
    Dim sNewRefersTo As String
    sNewRefersTo = Replace(sRefersTo, sOld, sNew)
    Debug.Print nmNext.Name, sRefersTo, sNewRefersTo, sNew, sOld
    Debug.Print nmNext.Name
    If sRefersTo <> sNewRefersTo Then
        If MsgBox("Confirm changing " & nmNext.Name & " replacing " & sOld & " with " & sNew & vbCrLf & "Name.RefersTo: " & nmNext.RefersTo & vbCrLf & "InputRefersTo: " & sRefersTo, vbYesNo, "Replace ?") = vbYes Then
            nmNext.RefersTo = sNewRefersTo
            ws.Cells(iRow, iColOldString) = "'" & sNew 'swap the old and new strings (so can undo if needed)
            ws.Cells(iRow, iColOldString + 1).Value = "'" & sOld
            ws.Cells(iRow, iColOldString + 2) = "'" & sRefersTo ' save the old refers to.
            ws.Cells(iRow, iColOldString + 3) = "'" & sNewRefersTo
        End If
    Else
        MsgBox nmNext.Name & " already refers to " & sNewRefersTo, vbOKOnly, "Information" ' should remove it from sheet ?
        
    End If
    
    Debug.Print iRow, nmNext.Name, nmNext.RefersTo, nmNext.RefersToLocal
    
End If
iRow = iRow + 1
Wend
Stop
MsgBox "Done"
End Sub

