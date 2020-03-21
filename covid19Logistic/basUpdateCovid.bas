Attribute VB_Name = "basUpdateCovid"
Public Const gcsSeriesConf = "time_series_19-covid-Confirmed" 'name of worksheet for output of confirmed cases series

' (c) K Duffy Mar 2020
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

Sub GetJHConfirmedData()
    'gets latest confirmed cases time series from John Hopkins Github
    'loads into the worksheet name as per gcsSeriesConf
    
    'JH data has one column per date, and an increasing number of rows
    ' per country or state with reported cases...
    
    ' need to process this into a time series running down the worksheet
    ' one row per date...
    
    ' this is done in the table named tblData (an excel listobject)
    
    'then processes into a myData worksheet int
    'myData worksheet has a large number of named ranges
    'and a table named tblData
    'tblData has 1 row per day of data
    'there are numerous charts in separate sheets, one per country
    'as its not easy to rearrange excel tables (listobjects) columns,
    'without damaging formulas that depend on them,
    'the columns might see haphazard in order
    
    'tblData has some columns to the right on NSW test data that is
    ' not available in the JH datafeed. These are manually fed in on
    ' a daily basis.
    
    'as well, the JH data is often not the latest available. The loading
    ' macro process will not overwrite any manually entered value that is
    ' larger than that read from JH.
    
    'the myData worksheet has a large set of named ranges, and in particular
    'dynamic named ranges that will expand as the rows of data accumulate.
    
    
    'in addition to tblData, there is a table to the right which is
    ' not a listobject, consisting of two parts.
    ' the top part is a list of the parameters that the Solver process
    ' is estimating, and used in the logistic equation output
    '
    ' below is is an area (growing) for pasting the daily solutions
    ' and then sorting them so that the growth of the asymptotes and
    ' other changes can be easily seen
   
    
    Dim vFiles() As Variant
    Dim sHeaders() As String
    Dim sURL As String
    
    Dim oHTTP As Object
    Dim sResponse As String
    Dim sPath As Variant
    Dim sOutPath As String
    Dim sOutFName As String
    Dim strM As String
    Dim iColN As Integer
    
    sOutPath = InputBox("Getting John Hopkins COVID19 CSV Update Files", "Enter Output Directory", CurDir)
    sOutPath = Replace(sOutPath & "\", "\\", "\")
    
    vFiles = Array("time_series_19-covid-Confirmed.csv", _
    "time_series_19-covid-Deaths.csv", _
    "time_series_19-covid-Recovered.csv")
    
    sHeaders = Split("Accept:application/vnd.github.v3.raw", ":")
    
    Dim sFNext As String
    Dim sDtMax As String
    Dim sDtNext As String
    Dim sFMax As String
    
    For Each sPath In vFiles
        
        sURL = "https://api.github.com/repos/CSSEGISandData/COVID-19/contents/csse_covid_19_data/csse_covid_19_time_series/" & sPath
        
        sFNext = Dir(sOutPath & "*" & sPath, vbNormal)
        While Len(sFNext) > 14
            sDtNext = Left(sFNext, 13)
            If sDtNext > sDtMax Then
                sDtMax = sDtNext
                sFMax = sFNext
            End If
            sFNext = Dir
        Wend
        
        Debug.Print "Latest File " & sFMax
        Dim sLastFile As String
        Dim strFileContent As String
        Dim iFile As Integer: iFile = FreeFile
        Open sFMax For Input As #iFile
        strFileContent = Input(LOF(iFile), iFile)
        Close #iFile
        
        If vbYes = MsgBox("Last file was " & sFMax & ", do you want to check for update", vbYesNo, "JH Data Fetch") Then
        
        
        Set oHTTP = CreateObject("MSXML2.ServerXMLHTTP")
        
        oHTTP.Open "GET", sURL, False
        oHTTP.setrequestheader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
        oHTTP.setrequestheader sHeaders(0), sHeaders(1)
        
        oHTTP.send ("")
        sResponse = DateConvert(oHTTP.responsetext)
        Debug.Print Len(strFileContent), Len(sResponse)
        
            
            sOutFName = sOutPath & Format(Now(), "yyyymmdd_hhmm") & "_raw_" & sPath
            Close 1
            Open sOutFName For Output As #1
            Print #1, sResponse
            Close 1
        
        If Len(strFileContent) = Len(sResponse & vbCrLf) And (sResponse & vbCrLf = strFileContent) Then
            If vbYes <> MsgBox("No Changes since " & sFMax & vbCrLf & "Do you want to reprocess that", vbYesNo, "No Change In Download Data") Then
                ' Stop
                Exit Sub
            Else
                sOutFnam = sFMax
            End If
        Else
            
            sOutFName = sOutPath & Format(Now(), "yyyymmdd_hhmm") & "_" & sPath
            
            Open sOutFName For Output As #1
            Print #1, sResponse
        End If
          strM = strM & sOutFName & " " & Len(sResponse) & " bytes" & vbCrLf
        ' Debug.Print Len(sResponse) & " Characters written to " & sOutFName
        Close 1
        Else
         sResponse = strFileContent
        End If
      
        Exit For
        
    Next sPath
    
    Dim ws As Worksheet
    Dim wsTSeries As Worksheet
    Set wsTSeries = Sheets(gcsSeriesConf)
    
    If InStr(1, wsTSeries.Cells(1, 1), "Province") < 1 Then
        MsgBox "Activesheet does not have Province as top left cell", vbOKOnly
        Stop
    End If
    
    Dim vOut() As Variant
    
    Dim sLines() As String
    If Right(sResponse, 1) = vbLf Then  'there is a trailing blank line in the file
        sResponse = Left(sResponse, Len(sResponse) - 1)
    End If
    
    sLines = Split(sResponse, vbLf)
    Dim vFields As Variant
    vFields = Split(sLines(2), ",")
    Dim lRows As Long
    Dim lCols As Long
    lRows = UBound(sLines) + 1
    lCols = UBound(vFields) + 1
    Debug.Print lRows, lCols
    ReDim vOut(lRows, lCols)
    
    Dim iL As Integer
    Dim iC As Integer
    For iL = 0 To lRows - 1
        vFields = Split(sLines(iL), ",")
        For iC = 0 To lCols - 1
            If iL = 0 And iC > 3 Then
                vOut(iL, iC) = CDate(vFields(iC))
            Else
                vOut(iL, iC) = vFields(iC)
            End If
        Next iC
    Next iL
    
    Dim dtMax As Date
    dtMax = CDate(vOut(0, lCols - 1))
    
    Dim rngOut As Range
    Set rngOut = Range(wsTSeries.Cells(1, 1), wsTSeries.Cells(lRows, lCols))
    rngOut = vOut
    
    Set ws = Sheets("myData")
    ws.Activate
    
    Dim vCountry As Variant
    Dim vCSumOut() As Variant
    
    Dim arrCtry As Variant
    arrCtry = Range("Countries")
    
    Dim dtcountry1st() As Date
    
    ReDim vCountry(UBound(arrCtry, 1) - 4) '-1 as last entry is test, -1 for NSW gap, -1 to skip heading row, -1 for array origin
    ReDim dtcountry1st(UBound(arrCtry, 1) - 4)
    Dim vCountryTbl() As Variant 'the names used in table heading can be different
    ReDim vCountryTbl(UBound(vCountry))
    Dim iC1 As Integer
    Dim iC2 As Integer
    iC1 = 0 'skip the header line with column labels
    For iC2 = 2 To UBound(arrCtry, 1) - 1
     If Not IsEmpty(arrCtry(iC2, 2)) Then
      vCountry(iC1) = arrCtry(iC2, 2)
      vCountryTbl(iC1) = arrCtry(iC2, 1)
      dtcountry1st(iC1) = CDate(arrCtry(iC2, 3))
      iC1 = iC1 + 1
     End If
     
    Next iC2
    'vCountry = Array("Australia", "Italy", "US", "China", "Korea South", "United Kingdom")  'names used in JH data
    'dtcountry1st = Array(#1/1/2020#, #1/1/2020#, #2/10/2020#, #1/1/2020#, #1/1/2020#, #1/1/2020#)   'constants must be in US format mm/dd/yyyy
    ReDim vCSumOut(lCols, UBound(vCountry) + 8) 'add 8 for separate australian states
    Dim sStates() As String 'the state labels expected in the data file
    sStates = Split("From Diamond Princess,Australian Capital Territory,New South Wales,Northern Territory,Queensland,South Australia,Tasmania,Victoria,Western Australia,", ",")
    
    Dim loData As ListObject
    Set loData = ws.ListObjects("tblData")
    Dim iRow1 As Integer 'row1 of table
    Dim rngHead As Range
    Set rngHead = loData.HeaderRowRange
    
    
    Dim iTableStartRow As Integer
    iTableStartRow = rngHead(1, 1).Row
    Dim iTableStartCol As Integer
    iTableStartCol = rngHead(1, 1).Column
    
    Dim nStates As Integer
    nStates = UBound(sStates) - 1
    
    Dim iColHeads() As Integer
    Dim sHeads() As String
    ReDim iColHeads(UBound(vCountry) + nStates + 2)
    ReDim sHeads(UBound(vCountry) + nStates + 2)
    
    Dim iX As Integer
    Dim iSt As Integer
    
    Dim iCountry As Integer
    For iCountry = 0 To UBound(vCountry)
        If iCountry = 0 Then
            For iX = 0 To nStates
                sSeek = sStates(iX)
                iColHeads(iCountry + iX) = ws.Rows(rngHead.Row).Find(What:=sSeek, LookAt:=xlPart).Column
                sHeads(iCountry + iX) = sSeek
            Next iX
        Else
            Dim sFindCtry As String
            sFindCtry = Replace(vCountry(iCountry), "USA", "US")
            sFindCtry = vCountryTbl(iCountry)
            
            iColHeads(iCountry + nStates) = ws.Rows(rngHead.Row).Find(What:=sFindCtry, LookAt:=xlPart).Column
            sHeads(iCountry + nStates) = vCountry(iCountry)
            
        End If
    Next iCountry
    
    Application.Calculation = xlCalculationManual
    
    'have to set the US column to zero so can sum onto it
    Dim iR As Integer
    Dim iNonBlank As Integer
    
    For iCountry = 0 To UBound(vCountry)
        
        If iCountry > 1 Then
            'USA have to zero it's data column.. same for china with state by state data
            iSt = rngHead.Find(What:=vCountryTbl(iCountry), LookAt:=xlWhole).Column
            iR = iTableStartRow + 1
            Debug.Print Range(ws.Cells(iR, iSt), ws.Cells(iR + lCols - 4, iSt)).Address
            Range(ws.Cells(iR, iSt), ws.Cells(iR + lCols - 4, iSt)).Clear
            
            'For iR = iTableStartRow + 1 To lCols + iTableStartRow - 4
            '    ws.Cells(iR, iSt).Clear
            ' Next iR
            
        End If
        
        For iL = 1 To lRows - 1
            If vOut(iL, 1) = vCountry(iCountry) Then  '1
                If iCountry = 0 Then
                    'it's australia, find the state
                    iSt = 0
                    sStates(UBound(sStates)) = vOut(iL, 0)
                    While sStates(iSt) <> vOut(iL, 0)
                        iSt = iSt + 1
                    Wend
                    If iSt = UBound(sStates) Then
                        Stop 'something wrong
                    End If
                Else
                    iSt = rngHead.Find(What:=vCountryTbl(iCountry), LookAt:=xlWhole).Column
                End If
                
                
                Dim iRowN As Integer
                Dim dtTableStart As Date
                dtTableStart = ws.Cells(iTableStartRow + 1, iTableStartCol)
                Dim iColM As Integer
                If iCountry = 0 Then
                    iColM = iColHeads(iSt)
                Else
                    iColM = iSt
                End If
                
                Dim sPlace As String
                Dim iPrev As Integer 'previous US value
                Dim lPrev As Long
                
                'iPrev = 0
                Dim iDateN As Integer
                
                For iDateN = 4 To lCols - 1  'for each date
                    iColN = iDateN + iTableStartRow - 3
                    lPrev = 0
                    If Not IsNumeric(vOut(iL, iDateN)) Then
                    Else
                        
                        If iCountry < 1 Then
                            sPlace = sStates(iSt)
                            
                        Else
                            sPlace = vCountry(iCountry) & " " & vOut(iL, 0)
                            If iCountry > 1 Then
                                If IsNumeric(ws.Cells(iColN, iColM)) Then
                                    lPrev = ws.Cells(iColN, iColM)
                                End If
                            End If
                        End If
                        
                        ' Debug.Print sPlace, vOut(0, iDateN), vOut(iL, iDateN)
                        ' Stop
                        
                        
                        If IsDate(ws.Cells(iColN, iTableStartCol)) Then '2
                            
                            If ws.Cells(iColN, iTableStartCol) = (vOut(0, iDateN)) Then  '3
                                ' it's the same date as row previously labelled with
                                'Stop
                                If IsEmpty(ws.Cells(iColN, iColM)) Then  '4
                                    'Debug.Print sPlace & " Date " & vOut(0, iDateN) & " being set to " & vOut(iL, iDateN)
                                    'doesnt matter if usa, italy or australian state...
                                    'except could get solver to work on USA until I blanked cells up to 10/2/2020
                                    If iCountry < 2 Then
                                        If vOut(iL, iDateN) > 0 Then
                                            ' Debug.Print sPlace & " Date " & vOut(0, iDateN) & " being set to " & vOut(iL, iDateN)
                                            ws.Cells(iColN, iColM).Value = CLng(vOut(iL, iDateN))
                                        Else
                                            '  Debug.Print sPlace & " Date " & vOut(0, iDateN) & " left at blank Row " & lColN
                                        End If
                                    Else
                                        'usa, leave dates before this blank for USA
                                        If CDate(vOut(0, iDateN)) > dtcountry1st(iCountry) Then
                                            If IsEmpty(ws.Cells(iColN - 1, iColM)) Then
                                                ws.Cells(iColN - 1, iColM) = 0
                                            End If
                                            '  Debug.Print sPlace & " Date " & vOut(0, iDateN) & " being set to " & vOut(iL, iDateN)
                                            ws.Cells(iColN, iColM).Value = CLng(vOut(iL, iDateN))
                                        End If
                                    End If
                                Else 'target cell not empty '4
                                    If iCountry < 2 Then '5
                                        ' not usa
                                        
                                        If CLng(ws.Cells(iColN, iColM)) <> CLng(vOut(iL, iDateN)) Then '6
                                            ' its being changed - weird, check <> didn't work
                                            Debug.Print sPlace & " Date " & vOut(0, iDateN) & " being changed from " & ws.Cells(iColN, iColM) & " to " & vOut(iL, iDateN)
                                            If CLng(ws.Cells(iColN, iColM)) > CLng(vOut(iL, iDateN)) Then '7
                                                Debug.Print "Reduction not allowed, review it"
                                            Else '7
                                                ws.Cells(iColN, iColM).Value = CLng(vOut(iL, iDateN))
                                            End If '7
                                        Else '6
                                            ' its same, no update needed
                                            'Stop
                                        End If '6
                                    Else '5
                                        'usa and target cell not empty
                                        If lPrev > 0 Then '6
                                            ws.Cells(iColN, iColM).Value = CLng(vOut(iL, iDateN)) + lPrev
                                        Else '6
                                            ws.Cells(iColN, iColM).Value = CLng(vOut(iL, iDateN))
                                        End If '6
                                        
                                    End If ' 5
                                End If '4
                            Else ' 3
                                'its a different date, error or find it ?
                                Stop
                            End If '3
                        Else '2
                            'probably inserting a new row
                            If IsDate(vOut(0, iDateN)) Then
                                ws.Cells(iColN, iTableStartCol).Value = CDate(vOut(0, iDateN))
                                ws.Cells(iColN, iColHeads(iSt)) = CLng(vOut(iL, iDateN))
                            End If
                        End If '2
                        ' If IsEmpty(ws.Cells(iColN, iColHeads(iSt))) Then
                        '  ws.Cells(iColN, iColHeads(iSt)) = vOut(iL, iDateN)
                        ' Else
                        '  Stop
                        
                        ' End If
                        'Stop
                        ' vCSumOut(iL, iDateN) = vCSumOut(iL, iDateN) + vFields(iDateN)
                        
                    End If 'isnumeric test
                    
                Next iDateN  'next date
            End If ' a country we want
        Next iL
        
        If iCountry = 2 Then
            'usa, cealr rows before the first 0
            For iR = iTableStartRow + 1 To lCols + iTableStartRow - 4
                If CLng(ws.Cells(iR + 1, iSt)) = 0 Then
                    ws.Cells(iR, iSt).Clear
                Else
                    ws.Cells(iR, iSt) = 0
                    Exit For
                End If
            Next iR
        End If
    Next iCountry
    Application.Calculation = xlCalculationAutomatic
    MsgBox strM & " last date " & Format(dtMax, "ddd dd/mm/yyyy"), vbOKOnly, "GetJHConfirmedData Done"
End Sub


Sub TestRange()
    Dim rng As Range
    
    Set rng = Range("myData!ChinaDates")
    Debug.Print rng.Address
    
    
End Sub
