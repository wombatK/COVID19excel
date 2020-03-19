Attribute VB_Name = "basSolveAll"


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

Sub SolveAll()
    '
    '
    ' Runs Solver for each of a list of countries
    ' need to be defined in a range with name Countries
    
    ' (see array vCtry)
    '
    ' needs a worksheet with a range named SolverMsg, screenscraped from
    ' https://docs.microsoft.com/en-us/office/vba/excel/concepts/functions/solversolve-function
    ' which defines the solver outcomes and messages
    '
    ' needs a reference to be set to SOlver addin see https://peltiertech.com/Excel/SolverVBA.html
    ' peltier has a way of making it less dependent on reference, but too much work !
    '
    ' needs the range of parameters used by solver to be in cells A2:E2
    ' needs CEngnotation function in workbook modules
    '
    Dim arrMsg As Variant
    arrMsg = Range("SolverMsg")  'case sensitive
    
    Dim arrCtry As Variant
    arrCtry = Range("Countries")
    
    Dim vCtry As Variant
    Dim iC As Integer
    Dim rngTest As Range
    Set rngTest = Range("A2")
    
    Dim strPrompt As String
    strPrompt = "Select Country Option " & vbCrLf & "0. For All"
    ReDim vCtry(UBound(arrCtry))
    For iC = 1 To UBound(arrCtry) - 1
        vCtry(iC) = arrCtry(iC + 1, 1)
        strPrompt = strPrompt & vbCrLf & iC & "." & vCtry(iC)
    Next iC
    Dim iOption As Integer
    
    iOption = InputBox(strPrompt, "Countres to Solve", 0)
    
    vCtry(UBound(vCtry)) = Replace(rngTest.Formula, "=beta", "")
    
    Dim rngReplace As Range
    Set rngReplace = Range("A2:E2")
    rngReplace.Select
    'find the country/state
    
    Dim sLookFor As String
    
    sLookFor = Replace(Range("A2").Formula, "=beta", "")
    
    'sLookFor = vCtry(iC)
    
    Dim rngF As Range
    
    For iC = 1 To UBound(vCtry) - 1
        
        If iOption = 0 Or iOption = iC Then
            'update the pararmeters used in solver constraints...
            Dim rngR As Range
            For Each rngR In rngReplace
              rngR.Formula = Replace(rngR.Formula, sLookFor, vCtry(iC), 1, 1, vbTextCompare)
        
            Next rngR
            
            Debug.Print vCtry(iC), rngTest.Formula
            
            Dim rngByChange As Range
            Set rngByChange = Range("myData!beta" & vCtry(iC))
            Set rngByChange = Range(rngByChange, rngByChange.Offset(0, 2)) 'three parameters
            Dim rngSetCell As Range
            Set rngSetCell = Range("myData!msq" & vCtry(iC))
            'peltier article shows how to add constraints too
            'see details of solveradd here
            'https://docs.microsoft.com/en-us/office/vba/excel/concepts/functions/solveradd-function
            
            SolverOk SetCell:=rngSetCell.Address, MaxMinVal:=2, ValueOf:=0, ByChange:=rngByChange.Address, Engine:=1, EngineDesc:="GRG Nonlinear"
            Dim vReturn As Variant
            vReturn = SolverSolve(UserFinish:=True)
            Dim strMsg As String
            strMsg = "Solver for " & vCtry(iC) & " return " & vReturn & ": " & arrMsg(vReturn + 2, 2)
          
            
            
            Dim sDeltaC As String
            sDeltaC = "myData!delta" & vCtry(iC)
            If Range(sDeltaC).Value <= 0 Then
                'solve is sometimes failing to observe the delta > 0 contrains
              Debug.Print strMsg
              MsgBox strMsg & "failed delta contraint, delta= " & Range(sDeltaC).Value, vbOKOnly, "Bad Delta"
              Range(sDeltaC).Value = 0.000001 'try setting an OK value
              vReturn = SolverSolve(UserFinish:=True)
           End If
            
            
            Dim dLimit As Double
            If IsNumeric(Range("myData!Asymptote").Value) Then
            
            dLimit = Range("myData!Asymptote").Value
            Else
             Stop  'maybe should try running Solver  manually now
             dLimit = 1E+15
            End If
            
              Debug.Print strMsg
            MsgBox strMsg, vbOKOnly, vCtry(iC) & " Solver"
            
            'SolverSolve
            ' SolverOk SetCell:="$I$12", MaxMinVal:=2, ValueOf:=0, ByChange:="$E$11:$G$11", _
            '    Engine:=2, EngineDesc:="Simplex LP"
            sLookFor = vCtry(iC) 'look for this on next iteration
            
        End If
    Next iC
    MsgBox "SolveAll Done", vbOKOnly, "Finished"
    
End Sub
