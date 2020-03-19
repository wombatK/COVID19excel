Attribute VB_Name = "basEngNotation"
' basEngNotation

'(c) K Duffy March 2020
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

Public Function CEngNotation(doubleValue As Double, Optional iSigDig As Integer = 3) As String
    Dim x As Double    ' --- Original Double (Floating-point)
    Dim y As Double    ' --- Mantissa
    Dim N As Long      ' --- Exponent
    Dim str As String
    Dim sign As String
    'On Error GoTo error_hander   ' --- uncomment for debug; disable when bug-free!
    x = doubleValue
    If doubleValue = 0 Then
        CEngNotation = "0"
        Exit Function
    End If
    'normalise value, and round significant digits
    Dim n1 As Integer
    
    Dim y1 As Double
    
    'Debug.Print x * 10 ^ -n1, 0.5 * 10 ^ -(iSigDig - 1)
    
    
    
    If x <> 0 Then
        If x < 0 Then
            ' --- x *must* be positive for log function to work
            x = x * -1
            sign = "-"    ' --- we need to preserve the sign for output string
        End If
        n1 = Int(Log(x) / Log(10))
        y1 = x * 10 ^ -n1 + 0.5 * 10 ^ -(iSigDig - 1)
        x = y1 * 10 ^ n1
        
        
        N = 3 * CLng((Log(x) / Log(1000)))   ' --- calculate Exponent...
        '     (Converts: log-base-e to log-base-10)
        y = x / (10 ^ N)                     ' --- calculate Mantissa.
        
        
        If y < 1 Then                        ' --- if Mantissa <1 then...
            N = N - 3                        ' --- ...adjust Exponent and...
            y = x / (10 ^ N)                 ' --- ...recalculate Mantissa.
        End If
        ' --- Create output string (special treatment when Exponent of zero; don't append "e")
        If N <= 24 And N >= -24 Then  'can VBA double range this big
            strSuffix = " " & Mid("yzafpnum kMGTPEZY", N / 3 + 9, 1)
        Else
            strSuffix = IIf(N <> 0, "e" & IIf(N > 0, "+", "") & N, "")
        End If
        
        If Right(Left(y, iSigDig + 1), 1) = "." Then
            str = sign & Left(y, iSigDig) & strSuffix
        Else
            str = sign & Left(y, iSigDig + 1) & strSuffix
        End If
    Else
        ' --- if the value is zero, well, return zero...
        str = "0"
    End If
    CEngNotation = str
    Exit Function
error_hander:
    ' --- this is really just for debugging suspected problems
    Resume Next
End Function

'And here's a function that I used to test my CEngNotation() function:
Private Sub test_ceng2()
    Dim x As Double
    
    x = 149556
    Debug.Print x, CEngNotation(x)
    
    x = 10565.3
    
    x = x * 10 ^ 22
    
    While x > 1E-26
        
        Debug.Print x, CEngNotation(x)
        Debug.Print x, CEngNotation(x, 4)
        Debug.Print x, CEngNotation(x, 5)
        Debug.Print "..."
        x = x / 10
        
    Wend
    
    
End Sub

Private Sub Test_CEngNotation()
    Dim x As Double
    
    x = 105653000000# * 1000000#
    
    Do While x > 0.000000000000001
        Debug.Print x, CEngNotation(x)
        x = x / 10
    Loop
End Sub

