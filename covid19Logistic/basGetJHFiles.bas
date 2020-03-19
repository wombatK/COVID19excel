Attribute VB_Name = "basGetJHFiles"
Option Explicit


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

Sub GetJHData()
    'first version, just gets files. See lata version in module
    'gets latest data from John Hopkins COVID19 github respository, outputs to a chosen folder...
    '16/03/2020
    
    Dim vFiles() As Variant
    Dim sHeaders() As String
    Dim sURL As String
    
    Dim oHTTP As Object
    Dim sResponse As String
    Dim sPath As Variant
    Dim sOutPath As String
    Dim sOutFName As String
    Dim strM As String
    
    sOutPath = InputBox("Getting John Hopkins COVID19 CSV Update Files", "Enter Output Directory", CurDir)
    sOutPath = Replace(sOutPath & "\", "\\", "\")
    
    vFiles = Array("time_series_19-covid-Confirmed.csv", _
    "time_series_19-covid-Deaths.csv", _
    "time_series_19-covid-Recovered.csv")
    
    sHeaders = Split("Accept:application/vnd.github.v3.raw", ":")
    
    
    For Each sPath In vFiles
        
        sURL = "https://api.github.com/repos/CSSEGISandData/COVID-19/contents/csse_covid_19_data/csse_covid_19_time_series/" & sPath
        
        
        Set oHTTP = CreateObject("MSXML2.ServerXMLHTTP")
        
        oHTTP.Open "GET", sURL, False
        oHTTP.setrequestheader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
        oHTTP.setrequestheader sHeaders(0), sHeaders(1)
        
        oHTTP.send ("")
        sResponse = DateConvert(oHTTP.responsetext)
        
        sOutFName = sOutPath & Format(Now(), "yyyymmdd_hhmm") & "_x_" & sPath
        
        Open sOutFName For Output As #1
        Print #1, sResponse
        strM = strM & sOutFName & " " & Len(sResponse) & " bytes" & vbCrLf
        ' Debug.Print Len(sResponse) & " Characters written to " & sOutFName
        Close 1
        
    Next sPath
    MsgBox strM, vbOKOnly, "GetJHData Done"
End Sub


Function DateConvert(strIn As String, Optional iQuoteOut As Integer = 1) As String
    'may require a reference to Microsoft VBScript Regular expressions in project (debugginer tools/References)
    Dim strOut As String
    Dim regEx As Object
    
    Set regEx = CreateObject("VBScript.RegExp"):
    regEx.Global = True
    regEx.MultiLine = True
    'regEx.Pattern = "\b(?<month>\d{1,2})/(?<day>\d{1,2})/(?<year>\d{2,4})\b"
    regEx.Pattern = "\b(0?[1-9]|1[012])([\/\-])(0?[1-9]|[12]\d|3[01])\2(\d{2,4})"
    
    strOut = regEx.Replace(strIn, "$3/$1/$4")
    'github file has opening " that fouls loading in excel
    'also, USA state and region is in form like ""Los Angeles, CA"" with comman in the middle before state
    'excel does not handle this correctly, need to change the comma, just lose it...
    'not interested in county by county or state totals in countries other than Australia
    
    Dim strN As String
    
    If iQuoteOut <> 0 Then
        
        regEx.Pattern = """([^,]*),(\s\w*)"""
        
        Debug.Print Right(strOut, 600), regEx.Pattern
        
        strN = regEx.Replace(strOut, "$1$2")
        
        Debug.Print Left(strN, 800) & vbCrLf & vbCrLf; Right(strN, 8600)
        
        regEx.Pattern = """"
        strOut = regEx.Replace(strN, "")  'celar out the opening quote
    End If
    
    DateConvert = strOut
End Function

Sub TestDateConvert()
    Dim sT As String
    sT = """Province/State,Country/Region,Lat,Long,1/22/20,1/23/20,1/24/20,1/25/20,1/26/20,1/27/20,1/28/20,""""Kitsap, WA"""",US,47.6477,-122.6413,""""Solano, CA"""",US,38.3105,"
    Debug.Print sT
    Debug.Print DateConvert(sT)
    
End Sub

Public Function USADateConvert(strUSDate) As Date
    'converts USA format date to australian, needs to be D/M/Y format
    'less robust method that worked on the John Hopkins data
    Dim sA() As String
    sA = Split(strUSDate, "/")
    USADateConvert = CDate(sA(1) & "/" & sA(0) & "/" & sA(2))  'DateSerial(sA(2), sA(1), sA(0))
End Function

