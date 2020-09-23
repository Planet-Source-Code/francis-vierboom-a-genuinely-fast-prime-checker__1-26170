VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Prime Checker"
   ClientHeight    =   3555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   ScaleHeight     =   3555
   ScaleWidth      =   4785
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   495
      Left            =   2880
      TabIndex        =   7
      Top             =   3000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtLimit 
      Height          =   285
      Left            =   2640
      TabIndex        =   3
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdFindPrimes 
      Caption         =   "Find Primes"
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox txtIn 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton cmdCheckPrime 
      Caption         =   "Check if this is a Prime"
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label lblTime 
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label lblRemain 
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "up to:"
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   2160
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Dim primesList() As Double
Dim stopPressed As Boolean


Private Sub cmdCheckPrime_Click()
Dim y As Double, x%, s&, f&, l&

s = timeGetTime
y = isComposite(CDbl(txtIn))
f = timeGetTime
l = f - s
If y = 0 Then
    x = MsgBox(txtIn & " is a prime. Took " & l & "ms")
Else
    x = MsgBox(txtIn & " is not prime; divisible by " & CStr(y) & ". Took " & l & "ms")
End If

End Sub

Private Sub cmdFindPrimes_Click()

Dim i As Double, s&, f&
Dim primes$
Dim limit As Double

limit = CDbl(txtLimit)
cmdStop.Visible = True
primesfile = FreeFile
stopPressed = False
lblRemain.Caption = "Expected time remaining: "
upper = primesList(UBound(primesList))
    'if the highest element of the list is higher than what's been typed in, inform user
    If upper >= limit Then
        x = MsgBox("I've already found primes up to that number.", vbInformation, "Already Done")
    Else
        'now, loop thru every number, from where we left off to where we're going
        s = timeGetTime
        
        For i = upper To limit
            If isComposite(i) = 0 Then
                'if there weren't any factors, then add it to the lists - ie the file and the array in memory
                Open App.Path & "\primes.txt" For Append As #primesfile
                    Print #primesfile, i
                Close #primesfile
                ReDim Preserve primesList(UBound(primesList) + 1)
                primesList(UBound(primesList)) = i
            End If
            'DoEvents < this one really slows things down.
            'one in every hundred numbers, update user on how long to go
            If i Mod 100 = 0 Then
                f = timeGetTime
                l = f - s
                remain = (l / (i - upper + 1)) * (limit - i)
                lblTime.Caption = (Int(remain / 100) / 10) & "s"
                DoEvents
            End If
            If stopPressed Then Exit For
        Next i
        f = timeGetTime
        l = f - s
        MsgBox "Found all primes up to " & i - 1 & " in " & l & "ms."
    End If
    
    lblRemain.Caption = ""
    lblTime.Caption = ""
    stopPressed = False
    cmdStop.Visible = False

End Sub

Function isComposite(number As Double) As Double

Dim i As Double, s&, f&, l&
Dim noofprimes As Double

isComposite = 0
noofprimes = UBound(primesList)
roof = Sqr(number)
    
'first, check if the number is already in the primes list. if its not, but its less than the limit, then obviously its a composite. or the program is screwed.
'For i = 0 To noofprimes
'    If number = primesList(i) Then Exit Function
'    If number < primesList(i) Then
''        isComposite = 1
'        Exit Function
'    End If
'Next i
's = timeGetTime

    'now, loop thru our list to see if any of the primes are factors
    For i = 0 To noofprimes
        If primesList(i) > roof Then Exit Function
        'check if it is divisible by primeslist(i)
        If number / primesList(i) = Int(number / primesList(i)) Then
            'if it was divisible, return the divisor
            isComposite = primesList(i)
            Exit Function
        End If
    Next i
'f = timeGetTime


    'if we've gotten this far, then there weren't any factors found in the primes we know
    'we have to check all the numbers up to the number's square root to see if we've missed any
    'this is the one place i can see where there's a potential for speeding things up, by eliminating a few more basic factors, eg 3, 5 etc... but the question soon arises if it's more time efficient to remove those numbers from the array or just loop through them.
    If primesList(noofprimes) < Int(Sqr(number)) Then
        k = CInt(primesList(noofprimes) Mod 2 = 0) + 1
 '       MsgBox "no known prime factors (" & f - s & "ms). Now looping from " & primesList(noofprimes) + k & " to " & Int(Sqr(number))
        For i = primesList(noofprimes) + k To Int(Sqr(number)) Step 2
            If number / i = Int(number / i) Then
                'if it was divisible, return the divisor
                isComposite = i
                Exit Function
            End If
        Next i
    End If
    
    
    

            
End Function

Private Sub cmdStop_Click()
stopPressed = True
End Sub

Private Sub Form_Load()
Dim fso As New FileSystemObject
Dim strnewprime$, dblnewprime As Double
s = timeGetTime

primesfile = FreeFile
'make sure the files exists - if it doesn't, get it started with 2
If Not fso.FileExists(App.Path & "\primes.txt") Then
    Open App.Path & "\primes.txt" For Output As #primesfile
        Print #primesfile, " 2 "
    Close #primesfile
End If

'prepare the array for the trenches
ReDim primesList(0)

Open App.Path & "\primes.txt" For Input As #primesfile
'get the firstline (special case... bcos the redim thing would rely on ubound being -1 to start off with)
If Not EOF(primesfile) Then
    Line Input #primesfile, strnewprime
    dblnewprime = CDbl(Val(strnewprime))
    If dblnewprime > 0 Then primesList(0) = dblnewprime
End If
'having problems.. so...
primesList(0) = 2
'then get the rest of it... all into one huge array
While Not EOF(primesfile)
    Line Input #primesfile, strnewprime
    dblnewprime = CDbl(Val(strnewprime))
    If dblnewprime > 0 Then
        ReDim Preserve primesList(UBound(primesList) + 1)
        primesList(UBound(primesList)) = dblnewprime
    End If
Wend
Close #primesfile
f = timeGetTime
l = f - s
x = MsgBox("Array creation took " & l & "ms")
End Sub

Private Sub txtIn_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    cmdCheckPrime_Click
End If
End Sub

Private Sub txtLimit_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    cmdFindPrimes_Click
End If
End Sub
