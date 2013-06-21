VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form REVERSI_ADV 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "2 Players Reversi"
   ClientHeight    =   8385
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   8805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd_suggestion 
      Caption         =   "Suggestion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   23
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmd_retainState 
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   21
      Top             =   7920
      Width           =   375
   End
   Begin VB.CommandButton cmd_storeState 
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   20
      Top             =   7920
      Width           =   375
   End
   Begin VB.CommandButton cmd_level 
      Caption         =   "Submit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      TabIndex        =   19
      Top             =   7920
      Width           =   735
   End
   Begin VB.TextBox LEVEL 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   6720
      TabIndex        =   18
      Text            =   "1"
      Top             =   7920
      Width           =   615
   End
   Begin VB.TextBox txt_turn 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11040
      TabIndex        =   9
      Top             =   2040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox deem 
      Height          =   735
      Left            =   1800
      Picture         =   "REVERSI_BEGINNERS.frx":0000
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   7
      Top             =   7320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox tn 
      Height          =   735
      Left            =   7560
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   6
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11040
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   1440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10320
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox wt 
      Height          =   735
      Left            =   3240
      Picture         =   "REVERSI_BEGINNERS.frx":1890
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   3
      Top             =   7320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox bk 
      Height          =   735
      Left            =   3960
      Picture         =   "REVERSI_BEGINNERS.frx":3053
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   2
      Top             =   7320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox bg 
      Height          =   735
      Left            =   2520
      Picture         =   "REVERSI_BEGINNERS.frx":4867
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   1
      Top             =   7320
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSFlexGridLib.MSFlexGrid fg 
      Height          =   5775
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   10186
      _Version        =   393216
      Rows            =   9
      Cols            =   9
      FixedRows       =   0
      FixedCols       =   0
      FocusRect       =   0
      HighLight       =   0
      MousePointer    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Caption         =   "REVERSI ADVANCE"
      ForeColor       =   &H00E0E0E0&
      Height          =   6495
      Left            =   360
      TabIndex        =   15
      Top             =   720
      Width           =   6615
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Caption         =   "REVERSI ADVANCE"
      ForeColor       =   &H00E0E0E0&
      Height          =   975
      Left            =   7440
      TabIndex        =   22
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label LBL_STATUS 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CLICK OVER THE GRID TOSTART YOUR TURN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   615
      Left            =   0
      TabIndex        =   17
      Top             =   7320
      Width           =   8775
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "REVERSI COMPUTER ADVANCE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   16
      Top             =   120
      Width           =   8775
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   8040
      TabIndex        =   14
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   7320
      TabIndex        =   13
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "White"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   8040
      TabIndex        =   12
      Top             =   4920
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Black"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7320
      TabIndex        =   11
      Top             =   4920
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SCORES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   7440
      TabIndex        =   10
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Turn"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   7440
      TabIndex        =   8
      Top             =   2040
      Width           =   975
   End
   Begin VB.Menu MNU_FILE 
      Caption         =   "FILE"
      Begin VB.Menu MNU_NEW 
         Caption         =   "NEW GAME"
      End
      Begin VB.Menu MNU_RESIGN 
         Caption         =   "RESIGN"
      End
      Begin VB.Menu MNU_EXIT 
         Caption         =   "EXIT"
      End
   End
   Begin VB.Menu MNU_LEVEL 
      Caption         =   "LEVEL"
      Begin VB.Menu MNU_2P 
         Caption         =   "TWO PLAYERS"
      End
      Begin VB.Menu MNU_BEG 
         Caption         =   "BEGINNERS"
      End
      Begin VB.Menu MNU_ADV 
         Caption         =   "ADVANCE"
      End
   End
   Begin VB.Menu MNU_HELP 
      Caption         =   "HELP"
      Begin VB.Menu MNU_CREDIT 
         Caption         =   "CREDIT"
      End
   End
End
Attribute VB_Name = "REVERSI_ADV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1

Dim countbk, countwt, posr, posc, turn, counter, counterB, abc, var_level As Integer
Dim bkpic, wtpic, bgpic As Picture
Dim wtc(64, 4), bkc(64, 4), BegArr(64, 3), BegArrOpp(64, 3), state(8, 8) As Variant
Dim deemCounterW, deemCounterB, BegScore, BegScoreOpp, localc, locald As Integer
Dim twoply(64, 6), twoply1(64, 3) As Variant
Dim deci As Double


Sub storeState()
'1 for background
'2 for black
'3 for white
'4 for deem
                                                                
For i = 1 To fg.Rows - 1
    For j = 1 To fg.Cols - 1
            fg.Row = i
            fg.Col = j
            If fg.CellPicture = bg.Picture Then
                state(i, j) = 1
            ElseIf fg.CellPicture = bk.Picture Then
                state(i, j) = 2
            ElseIf fg.CellPicture = wt.Picture Then
                state(i, j) = 3
            Else
                state(i, j) = 4
            End If
    Next j
Next i
                                                                
End Sub

Sub retainState()

For i = 1 To fg.Rows - 1
    For j = 1 To fg.Cols - 1
            fg.Row = i
            fg.Col = j
            If state(i, j) = 1 Then
                Set fg.CellPicture = bg.Picture
            ElseIf state(i, j) = 2 Then
                   Set fg.CellPicture = bk.Picture
            ElseIf state(i, j) = 3 Then
                Set fg.CellPicture = wt.Picture
            Else
                Set fg.CellPicture = deem.Picture
            End If
    Next j
Next i

End Sub

Function BegApp(r As Variant, c As Variant)

fg.Row = r
fg.Col = c
posr = r
posc = c

If fg.CellPicture <> deem.Picture Then
    GoTo endmark
End If

If deemCounterB = 0 And deemCounterW = 0 Then
    GoTo endmark
End If

Skipturn2:

If turn = 1 Then
    Call ruleApp_bk
    Call distroyDeem
    Call searchWt
    
    If deemCounterB = 0 And deemCounterW = 0 Then
        GoTo endmark
    End If
    
    turn = 2
    
    If deemCounterW = 0 Then
        GoTo Skipturn1
    End If
    
    txt_turn.Text = turn
    Set tn.Picture = wtpic
  
    GoTo endmark
End If

Skipturn1:

If turn = 2 Then

    Call ruleApp_wt
    Call distroyDeem
    Call searchBk
    If deemCounterB = 0 And deemCounterW = 0 Then
        GoTo endmark
    End If
        
    turn = 1
    If deemCounterB = 0 Then
        GoTo Skipturn2
    End If
    
    txt_turn.Text = turn
    Set tn.Picture = bkpic
    GoTo endmark
    
    End If
endmark:

Call counting
End Function


Sub BegTop()


For i = 1 To counterB - 1
    rw = bkc(i, 1)
    cm = bkc(i, 2)
    Call Beg(rw, cm)
    BegArr(i, 1) = bkc(i, 1)
    BegArr(i, 2) = bkc(i, 2)
    BegArr(i, 3) = BegScore
Next i

For i = 1 To counterB - 2
     For j = i + 1 To counterB - 1
           
         If BegArr(i, 3) < BegArr(j, 3) Then
                 tempo = BegArr(i, 3)
                 hb = BegArr(i, 2)
                 hstring = BegArr(i, 1)
                 
                 BegArr(i, 3) = BegArr(j, 3)
                 BegArr(i, 2) = BegArr(j, 2)
                 BegArr(i, 1) = BegArr(j, 1)
                 
                 BegArr(j, 3) = tempo
                 BegArr(j, 2) = hb
                 BegArr(j, 1) = hstring
                                
          End If
   Next j
Next i

'Call BegApp(BegArr(1, 1), BegArr(1, 2))

End Sub

Sub BegTop1()
For i = 1 To counterB - 1
    rw = bkc(i, 1)
    cm = bkc(i, 2)
    Call Beg(rw, cm)
    BegArr(i, 1) = bkc(i, 1)
    BegArr(i, 2) = bkc(i, 2)
    BegArr(i, 3) = BegScore
Next i

For i = 1 To counterB - 2
   For j = i + 1 To counterB - 1
   
         If BegArr(i, 3) < BegArr(j, 3) Then
                 tempo = BegArr(i, 3)
                 hb = BegArr(i, 2)
                 hstring = BegArr(i, 1)
                 
                 BegArr(i, 3) = BegArr(j, 3)
                 BegArr(i, 2) = BegArr(j, 2)
                 BegArr(i, 1) = BegArr(j, 1)
                 
                 BegArr(j, 3) = tempo
                 BegArr(j, 2) = hb
                 BegArr(j, 1) = hstring
                                
          End If
   Next j
Next i

Call BegApp(BegArr(1, 1), BegArr(1, 2))

End Sub


Sub twoPlyFun()

For i = 1 To counterB - 1
    
    rw = bkc(i, 1)
    cm = bkc(i, 2)
    
    Call Beg(rw, cm)
    twoply(i, 5) = BegScore
    Call storeState
       
    rw = bkc(i, 1)
    cm = bkc(i, 2)
       
    posr = rw
    posc = cm
    twoply(i, 1) = rw
    twoply(i, 2) = cm
    fg.Row = rw
    fg.Col = cm
    
    
    Call ruleApp_bk
    Call distroyDeem
    Call searchWt
    Call BegTopOpp
    
    twoply(i, 4) = BegArrOpp(1, 3)
    
    If counter = 1 Then
        twoply(i, 3) = BegArrOpp(counter, 3)
    Else
        twoply(i, 3) = BegArrOpp(counter - 1, 3)
    End If
    
    If twoply(i, 4) < 10 Then
            stri = twoply(i, 3) & ".0" & twoply(i, 4)
    Else
             stri = twoply(i, 3) & "." & twoply(i, 4)
    End If
                 
    deci = stri
    twoply(i, 6) = deci
    Call retainState
Next i

REVERSI_ADV.Cls

For i = 1 To counterB - 2
   For j = i + 1 To counterB - 1

         If twoply(i, 6) > twoply(j, 6) Then
                 tempo = twoply(i, 6)
                 hstring1 = twoply(i, 1)
                 hstring2 = twoply(i, 2)
                 hstring3 = twoply(i, 3)
                 hstring4 = twoply(i, 4)
                 hstring5 = twoply(i, 5)

                 twoply(i, 6) = twoply(j, 6)
                 twoply(i, 5) = twoply(j, 5)
                 twoply(i, 4) = twoply(j, 4)
                 twoply(i, 3) = twoply(j, 3)
                 twoply(i, 2) = twoply(j, 2)
                 twoply(i, 1) = twoply(j, 1)

                 twoply(j, 6) = tempo
                 twoply(j, 1) = hstring1
                 twoply(j, 2) = hstring2
                 twoply(j, 3) = hstring3
                 twoply(j, 4) = hstring4
                 twoply(j, 5) = hstring5

          End If
   Next j
Next i

For i = 1 To counterB - 1
    If (twoply(i, 1) = 1 And twoply(i, 2) = 8) Or (twoply(i, 1) = 8 And twoply(i, 2) = 1) Or (twoply(i, 1) = 1 And twoply(i, 2) = 1) Or (twoply(i, 1) = 8 And twoply(i, 2) = 8) Then
        twoply1(1, 1) = twoply(i, 1)
        twoply1(1, 2) = twoply(i, 2)
        twoply1(1, 3) = twoply(i, 5)
    GoTo outer
    End If
'Print twoply(i, 1) & "," & twoply(i, 2) & " **** " & twoply(i, 3) & " --- " & twoply(i, 4) & " -- " & twoply(i, 5) & " -- " & twoply(i, 6)
Next i

localCounter = 0

For i = 1 To counterB - 2
    If twoply(1, 6) = twoply(i, 6) Then
        twoply1(i, 1) = twoply(i, 1)
        twoply1(i, 2) = twoply(i, 2)
        twoply1(i, 3) = twoply(i, 5)
        localCounter = localCounter + 1
    Else
        GoTo out1
    End If
Next i

out1:
If localCounter = 1 Then
    GoTo out
End If

For i = 1 To localCounter - 1
   For j = i + 1 To localCounter
         If twoply1(i, 3) < twoply1(j, 3) Then
                 tempo = twoply1(i, 3)
                 hb = twoply1(i, 2)
                 hstring = twoply1(i, 1)
                 
                 twoply1(i, 3) = twoply1(j, 3)
                 twoply1(i, 2) = twoply1(j, 2)
                 twoply1(i, 1) = twoply1(j, 1)
                 
                 twoply1(j, 3) = tempo
                 twoply1(j, 2) = hb
                 twoply1(j, 1) = hstring
                                
          End If
   Next j
Next i

out:

'For i = 1 To localCounter
'Print twoply1(i, 1) & "," & twoply1(i, 2) & " **** " & twoply1(i, 3)
'Next i
outer:
    Call BegApp(twoply(1, 1), twoply(1, 2))
End Sub

Sub BegTopOpp()

For i = 1 To counter - 1
    rw = wtc(i, 1)
    cm = wtc(i, 2)
    Call BegOpp(rw, cm)
    
    BegArrOpp(i, 1) = wtc(i, 1)
    BegArrOpp(i, 2) = wtc(i, 2)
    BegArrOpp(i, 3) = BegScoreOpp
Next i


For i = 1 To counter - 2
    For j = i + 1 To counter - 1
          If BegArrOpp(i, 3) < BegArrOpp(j, 3) Then
                  tempo = BegArrOpp(i, 3)
                  hb = BegArrOpp(i, 2)
                  hstring = BegArrOpp(i, 1)
    
                  BegArrOpp(i, 3) = BegArrOpp(j, 3)
                  BegArrOpp(i, 2) = BegArrOpp(j, 2)
                  BegArrOpp(i, 1) = BegArrOpp(j, 1)
    
                  BegArrOpp(j, 3) = tempo
                  BegArrOpp(j, 2) = hb
                  BegArrOpp(j, 1) = hstring
    
           End If
    Next j
Next i

End Sub



Function Beg(rw As Variant, cm As Variant)

locald = 0
BegScore = 0
fg.Row = rw
fg.Col = cm

If fg.CellPicture <> deem.Picture Then
    Exit Function
End If

'*******************************************************Chk1 UPWARD
If rw < 3 Then
    ''Msgbox "No rules can be applied 1"
    GoTo chk1exit
End If

fg.Row = rw - 1
fg.Col = cm

If fg.CellPicture = wt.Picture Then
    GoTo chk1
Else
    GoTo chk1exit
End If

chk1:
For i = fg.Row - 1 To 1 Step -1
    fg.Row = i
    If fg.CellPicture = wt.Picture Then
        GoTo last1
    ElseIf fg.CellPicture = bk.Picture Then
        applyr = fg.Row
        applyc = fg.Col
        GoTo chk1sec
    Else
        GoTo chk1exit
    End If
last1:
Next i

If fg.CellPicture <> bk.Picture Then
    GoTo chk1exit
End If

chk1sec:

locald = locald + 1
For k = rw To applyr + 1 Step -1
    fg.Row = k
    BegScore = BegScore + 1
Next k

chk1exit:

'*******************************************************Chk2 DOWNWARD

If rw > 6 Then
    GoTo chk2exit
End If

fg.Row = rw + 1
fg.Col = cm

If fg.CellPicture = wt.Picture Then
    GoTo chk2
Else
    GoTo chk2exit
End If

chk2:

For i = fg.Row + 1 To fg.Rows - 1
    fg.Row = i
        
    If fg.CellPicture = wt.Picture Then
        GoTo last2
    ElseIf fg.CellPicture = bk.Picture Then
        applyr = fg.Row
        applyc = fg.Col
        GoTo chk2sec
    Else
        GoTo chk2exit
    End If
last2:
Next i

If fg.CellPicture <> bk.Picture Then
    GoTo chk2exit
End If


chk2sec:
locald = locald + 1

For k = rw To applyr - 1
fg.Row = k
    BegScore = BegScore + 1

Next k

chk2exit:

'*******************************************************Chk3 RIGHT DIRECTION

If cm > 6 Then
    GoTo chk3exit
End If

fg.Row = rw
fg.Col = cm + 1

If fg.CellPicture = wt.Picture Then
    GoTo chk3
Else
    GoTo chk3exit
End If

chk3:

For i = fg.Col + 1 To fg.Cols - 1
    fg.Col = i
    fg.Row = rw
        
    If fg.CellPicture = wt.Picture Then
        GoTo last3
    ElseIf fg.CellPicture = bk.Picture Then
        
        applyr = fg.Row
        applyc = fg.Col
               
        GoTo chk3sec
    Else
        GoTo chk3exit
    End If
last3:
Next i

If fg.CellPicture <> bk.Picture Then
    GoTo chk3exit
End If

chk3sec:
locald = locald + 1
    For k = cm To applyc - 1
    fg.Col = k
    BegScore = BegScore + 1
Next k

chk3exit:

'*******************************************************Chk3 LEFT DIRECTION

If cm < 3 Then
    GoTo chk4exit
End If

fg.Row = rw
fg.Col = cm - 1

If fg.CellPicture = wt.Picture Then
    GoTo chk4
Else
    GoTo chk4exit
End If

chk4:

For i = fg.Col - 1 To 1 Step -1
    fg.Col = i
    fg.Row = rw
        
    If fg.CellPicture = wt.Picture Then
        GoTo last4
    ElseIf fg.CellPicture = bk.Picture Then
        applyr = fg.Row
        applyc = fg.Col
        GoTo chk4sec
    Else
        GoTo chk4exit
    End If
last4:
Next i

If fg.CellPicture <> bk.Picture Then
    GoTo chk4exit
End If


chk4sec:
locald = locald + 1
For k = cm To applyc + 1 Step -1
    fg.Col = k
    BegScore = BegScore + 1

Next k


chk4exit:

'*******************************************************Chk3 TOP RIGHT DIRECTION

If cm > 6 Or rw < 3 Then
    GoTo chk5exit
End If

fg.Row = rw - 1
fg.Col = cm + 1

If fg.CellPicture = wt.Picture Then
    GoTo chk5
Else
    GoTo chk5exit
End If

chk5:

For i = fg.Row - 1 To 1 Step -1
    If fg.Col + 1 = fg.Cols Then
        GoTo chk5exit
    End If
        
    fg.Col = fg.Col + 1
    fg.Row = i
        
    If fg.CellPicture = wt.Picture Then
        GoTo last5
    ElseIf fg.CellPicture = bk.Picture Then
        applyr = fg.Row
        applyc = fg.Col
        GoTo chk5sec
    Else
        GoTo chk5exit
    End If
last5:
Next i

If fg.CellPicture <> bk.Picture Then
    GoTo chk5exit
End If


chk5sec:
locald = locald + 1
km = cm
For k = rw To applyr + 1 Step -1
    fg.Col = km
    fg.Row = k
    BegScore = BegScore + 1
    km = km + 1
Next k

chk5exit:

'*******************************************************Chk3 TOP LEFT DIRECTION

If cm < 3 Or rw < 3 Then
    GoTo chk6exit
End If

fg.Row = rw - 1
fg.Col = cm - 1

If fg.CellPicture = wt.Picture Then
    GoTo chk6
Else
    GoTo chk6exit
End If

chk6:

For i = fg.Row - 1 To 1 Step -1
    If fg.Col - 1 = 0 Then
        GoTo chk6exit
    End If
        
    fg.Col = fg.Col - 1
    fg.Row = i
        
    If fg.CellPicture = wt.Picture Then
        GoTo last6
    ElseIf fg.CellPicture = bk.Picture Then
        applyr = fg.Row
        applyc = fg.Col
        GoTo chk6sec
    Else
        GoTo chk6exit
    End If
last6:
Next i

If fg.CellPicture <> bk.Picture Then
    GoTo chk6exit
End If


chk6sec:
locald = locald + 1
km = cm
For k = rw To applyr + 1 Step -1
    fg.Col = km
    fg.Row = k
    BegScore = BegScore + 1
    km = km - 1
Next k

chk6exit:


'*******************************************************Chk3 BOTTOM LEFT DIRECTION

If cm < 3 Or rw > 6 Then
    GoTo chk7exit
End If

fg.Row = rw + 1
fg.Col = cm - 1

If fg.CellPicture = wt.Picture Then
    GoTo chk7
Else
    GoTo chk7exit
End If

chk7:

For i = fg.Row + 1 To fg.Rows - 1
    If fg.Col - 1 = 0 Then
        GoTo chk7exit
    End If

    fg.Col = fg.Col - 1
    fg.Row = i

    If fg.CellPicture = wt.Picture Then
        GoTo last7
    ElseIf fg.CellPicture = bk.Picture Then

        applyr = fg.Row
        applyc = fg.Col
        GoTo chk7sec
    Else
        GoTo chk7exit
    End If
last7:
Next i

If fg.CellPicture <> bk.Picture Then
    GoTo chk7exit
End If


chk7sec:
locald = locald + 1
kr = rw
For k = cm To applyc + 1 Step -1
    fg.Row = kr
    fg.Col = k
    BegScore = BegScore + 1
    kr = kr + 1
Next k


chk7exit:

If cm > 6 Or rw > 6 Then
    GoTo chk8exit
End If

fg.Row = rw + 1
fg.Col = cm + 1

If fg.CellPicture = wt.Picture Then
    GoTo chk8
Else
    GoTo chk8exit
End If

chk8:

For i = fg.Row + 1 To fg.Rows - 1
    If fg.Col + 1 = fg.Cols Then
        GoTo chk8exit
    End If

    fg.Col = fg.Col + 1
    fg.Row = i

    If fg.CellPicture = wt.Picture Then
        GoTo last8
    ElseIf fg.CellPicture = bk.Picture Then
        applyr = fg.Row
        applyc = fg.Col
        GoTo chk8sec
    Else
        GoTo chk8exit
    End If
last8:
Next i

If fg.CellPicture <> bk.Picture Then
    GoTo chk8exit
End If


chk8sec:
locald = locald + 1
kr = rw
For k = cm To applyc - 1
    fg.Row = kr
    fg.Col = k
    BegScore = BegScore + 1
    kr = kr + 1
Next k


chk8exit:

If locald > 1 Then
    BegScore = BegScore - (locald - 1)
End If

End Function

Function BegOpp(rw As Variant, cm As Variant)
localc = 0

BegScoreOpp = 0
'reversi_adv.Cls

fg.Row = rw
fg.Col = cm

'BegArr(BegScoreOpp,1) =

If fg.CellPicture <> deem.Picture Then
    Exit Function
End If

'*******************************************************Chk1 UPWARD
If rw < 3 Then
    GoTo chk1exit
End If

fg.Row = rw - 1
fg.Col = cm

If fg.CellPicture = bk.Picture Then
    GoTo chk1
Else
    GoTo chk1exit
End If

chk1:

For i = fg.Row - 1 To 1 Step -1
    fg.Row = i
    
    
    If fg.CellPicture = bk.Picture Then
        GoTo last1
        
    ElseIf fg.CellPicture = wt.Picture Then
        applyr = fg.Row
        applyc = fg.Col
        GoTo chk1sec
    Else
        GoTo chk1exit
    End If
last1:
Next i

If fg.CellPicture <> wt.Picture Then
    GoTo chk1exit
End If

chk1sec:
localc = localc + 1
For k = rw To applyr + 1 Step -1
    fg.Row = k
    BegScoreOpp = BegScoreOpp + 1
Next k

chk1exit:

'*******************************************************Chk2 DOWNWARD

If rw > 6 Then
    GoTo chk2exit
End If

fg.Row = rw + 1
fg.Col = cm

If fg.CellPicture = bk.Picture Then
    GoTo chk2
Else
    GoTo chk2exit
End If

chk2:

For i = fg.Row + 1 To fg.Rows - 1
    fg.Row = i
        
    If fg.CellPicture = bk.Picture Then
        GoTo last2
    ElseIf fg.CellPicture = wt.Picture Then
        applyr = fg.Row
        applyc = fg.Col
        GoTo chk2sec
    Else
        GoTo chk2exit
    End If
last2:
Next i

If fg.CellPicture <> wt.Picture Then
    GoTo chk2exit
End If

chk2sec:
localc = localc + 1
For k = rw To applyr - 1
fg.Row = k
    BegScoreOpp = BegScoreOpp + 1

Next k

chk2exit:

'*******************************************************Chk3 RIGHT DIRECTION

If cm > 6 Then
    GoTo chk3exit
End If

fg.Row = rw
fg.Col = cm + 1

If fg.CellPicture = bk.Picture Then
    GoTo chk3
Else
    GoTo chk3exit
End If

chk3:

For i = fg.Col + 1 To fg.Cols - 1
    fg.Col = i
    fg.Row = rw
        
    If fg.CellPicture = bk.Picture Then
        GoTo last3
    ElseIf fg.CellPicture = wt.Picture Then
        
        applyr = fg.Row
        applyc = fg.Col
               
        GoTo chk3sec
    Else
        GoTo chk3exit
    End If
last3:
Next i

If fg.CellPicture <> wt.Picture Then
    GoTo chk3exit
End If


chk3sec:
localc = localc + 1
    For k = cm To applyc - 1
    fg.Col = k
    BegScoreOpp = BegScoreOpp + 1
Next k

chk3exit:

'*******************************************************Chk3 LEFT DIRECTION

If cm < 3 Then
    GoTo chk4exit
End If

fg.Row = rw
fg.Col = cm - 1

If fg.CellPicture = bk.Picture Then
    GoTo chk4
Else
    GoTo chk4exit
End If

chk4:

For i = fg.Col - 1 To 1 Step -1
    fg.Col = i
    fg.Row = rw
        
    If fg.CellPicture = bk.Picture Then
        GoTo last4
    ElseIf fg.CellPicture = wt.Picture Then
        applyr = fg.Row
        applyc = fg.Col
        GoTo chk4sec
    Else
        GoTo chk4exit
    End If
last4:
Next i

If fg.CellPicture <> wt.Picture Then
    GoTo chk4exit
End If


chk4sec:

localc = localc + 1
For k = cm To applyc + 1 Step -1
    fg.Col = k
    BegScoreOpp = BegScoreOpp + 1

Next k


chk4exit:

'*******************************************************Chk3 TOP RIGHT DIRECTION

If cm > 6 Or rw < 3 Then
    GoTo chk5exit
End If

fg.Row = rw - 1
fg.Col = cm + 1

If fg.CellPicture = bk.Picture Then
    GoTo chk5
Else
    GoTo chk5exit
End If

chk5:

For i = fg.Row - 1 To 1 Step -1
    If fg.Col + 1 = fg.Cols Then
        GoTo chk5exit
    End If
        
    fg.Col = fg.Col + 1
    fg.Row = i
        
    If fg.CellPicture = bk.Picture Then
        GoTo last5
    ElseIf fg.CellPicture = wt.Picture Then
        applyr = fg.Row
        applyc = fg.Col
        GoTo chk5sec
    Else
        GoTo chk5exit
    End If
last5:
Next i

If fg.CellPicture <> wt.Picture Then
    GoTo chk5exit
End If


chk5sec:
localc = localc + 1
kg = cm
For k = rw To applyr + 1 Step -1
    
    fg.Col = kg
    fg.Row = k
    BegScoreOpp = BegScoreOpp + 1
     kg = kg + 1
Next k

chk5exit:

'*******************************************************Chk3 TOP LEFT DIRECTION

If cm < 3 Or rw < 3 Then
    GoTo chk6exit
End If

fg.Row = rw - 1
fg.Col = cm - 1

If fg.CellPicture = bk.Picture Then
    GoTo chk6
Else
    GoTo chk6exit
End If

chk6:

For i = fg.Row - 1 To 1 Step -1
    If fg.Col - 1 = 0 Then
        GoTo chk6exit
    End If
        
    fg.Col = fg.Col - 1
    fg.Row = i
        
    If fg.CellPicture = bk.Picture Then
        GoTo last6
    ElseIf fg.CellPicture = wt.Picture Then
        applyr = fg.Row
        applyc = fg.Col
        GoTo chk6sec
    Else
        GoTo chk6exit
    End If
last6:
Next i

If fg.CellPicture <> wt.Picture Then
    GoTo chk6exit
End If


chk6sec:
localc = localc + 1
kg = cm
For k = rw To applyr + 1 Step -1
    fg.Col = kg
    fg.Row = k
    BegScoreOpp = BegScoreOpp + 1
    kg = kg - 1
Next k

chk6exit:


'*******************************************************Chk3 BOTTOM LEFT DIRECTION

If cm < 3 Or rw > 6 Then
    GoTo chk7exit
End If

fg.Row = rw + 1
fg.Col = cm - 1

If fg.CellPicture = bk.Picture Then
    GoTo chk7
Else
    GoTo chk7exit
End If

chk7:

For i = fg.Row + 1 To fg.Rows - 1
    If fg.Col - 1 = 0 Then
        GoTo chk7exit
    End If

    fg.Col = fg.Col - 1
    fg.Row = i

    If fg.CellPicture = bk.Picture Then
        GoTo last7
    ElseIf fg.CellPicture = wt.Picture Then

        applyr = fg.Row
        applyc = fg.Col
        GoTo chk7sec
    Else
        GoTo chk7exit
    End If
last7:
Next i

If fg.CellPicture <> wt.Picture Then
    GoTo chk7exit
End If


chk7sec:
localc = localc + 1
kr = rw
For k = cm To applyc + 1 Step -1
    fg.Row = kr
    fg.Col = k
    BegScoreOpp = BegScoreOpp + 1
    kr = kr + 1
Next k


chk7exit:

If cm > 6 Or rw > 6 Then
    GoTo chk8exit
End If

fg.Row = rw + 1
fg.Col = cm + 1

If fg.CellPicture = bk.Picture Then
    GoTo chk8
Else
    GoTo chk8exit
End If

chk8:

For i = fg.Row + 1 To fg.Rows - 1
    
    If fg.Col + 1 = fg.Cols Then
        GoTo chk8exit
    End If

    fg.Col = fg.Col + 1
    fg.Row = i

    If fg.CellPicture = bk.Picture Then
        GoTo last8
    ElseIf fg.CellPicture = wt.Picture Then
        applyr = fg.Row
        applyc = fg.Col
        GoTo chk8sec
    Else
        GoTo chk8exit
    End If
last8:
Next i

If fg.CellPicture <> wt.Picture Then
    GoTo chk8exit
End If


chk8sec:

localc = localc + 1
kr = rw
For k = cm To applyc - 1
    fg.Row = kr
    fg.Col = k
    BegScoreOpp = BegScoreOpp + 1
    kr = kr + 1
Next k

chk8exit:


If localc > 1 Then
    BegScoreOpp = BegScoreOpp - (localc - 1)
End If

End Function
Sub turnChanger()

If turn = 1 Then
    turn = 2
    txt_turn.Text = turn
    Set tn.Picture = wtpic
    GoTo endmark
    
End If

If turn = 2 Then
    turn = 1
    txt_turn.Text = turn
    Set tn.Picture = bkpic
    GoTo endmark
End If

endmark:
abc = 0

End Sub
Sub counting()
                                                               'fg.Visible = False
countbk = 0
countwt = 0

For i = 1 To fg.Rows - 1
    For j = 1 To fg.Cols - 1
       fg.Row = i
       fg.Col = j
       
       If fg.CellPicture = bk.Picture Then
            countbk = countbk + 1
        End If
       
       If fg.CellPicture = wt.Picture Then
            countwt = countwt + 1
        End If
       
    Next j
Next i
Label5.Caption = countbk
Label6.Caption = countwt
                                                                'fg.Visible = True
End Sub

Sub distroyDeem()
'fg.visible = false
For i = 1 To fg.Rows - 1
    For j = 1 To fg.Cols - 1
       fg.Row = i
       fg.Col = j
       
       If fg.CellPicture = deem.Picture Then
            Set fg.CellPicture = bg.Picture
        End If
                    
    Next j
Next i

'fg.visible = true


End Sub

Sub searchWt()

deemCounterW = 0
counter = 1
'fg.visible = false
For i = 1 To fg.Rows - 1
    For j = 1 To fg.Cols - 1
    
    fg.Row = i
    fg.Col = j
    
    If fg.CellPicture = wt.Picture Then
        posr = i
        posc = j
        Call chk_white
    End If
    
    Next j
Next i


'reversi_adv.Cls
For k = 1 To counter - 1
    fg.Row = wtc(k, 1)
    fg.Col = wtc(k, 2)
    
    If fg.CellPicture <> deem.Picture Then
        Set fg.CellPicture = deem.Picture
        deemCounterW = deemCounterW + 1
    End If
    
Next k
'fg.visible = true

End Sub


Sub ruleApp_bk()

'reversi_adv.Cls
fg.Row = posr
fg.Col = posc

If fg.CellPicture <> deem.Picture Then
    Exit Sub
End If

'*******************************************************Chk1 UPWARD
If posr < 3 Then
    GoTo chk1exit
End If

fg.Row = posr - 1
fg.Col = posc

If fg.CellPicture = wt.Picture Then
    GoTo chk1
Else
    GoTo chk1exit
End If

chk1:

For i = fg.Row - 1 To 1 Step -1
    fg.Row = i
    
    
    If fg.CellPicture = wt.Picture Then
        GoTo last1
        
    ElseIf fg.CellPicture = bk.Picture Then
        applyr = fg.Row
        applyc = fg.Col
        
        
        GoTo chk1sec
    Else
        GoTo chk1exit
    End If
last1:
Next i

If fg.CellPicture <> bk.Picture Then
    GoTo chk1exit
End If

chk1sec:

For k = posr To applyr + 1 Step -1
fg.Row = k
Set fg.CellPicture = bk.Picture

Next k

chk1exit:

'*******************************************************Chk2 DOWNWARD

If posr > 6 Then
    GoTo chk2exit
End If

fg.Row = posr + 1
fg.Col = posc

If fg.CellPicture = wt.Picture Then
    GoTo chk2
Else
    GoTo chk2exit
End If

chk2:

For i = fg.Row + 1 To fg.Rows - 1
    fg.Row = i
        
    If fg.CellPicture = wt.Picture Then
        GoTo last2
    ElseIf fg.CellPicture = bk.Picture Then
        applyr = fg.Row
        applyc = fg.Col
   
        
        GoTo chk2sec
    Else
        GoTo chk2exit
    End If
last2:
Next i

If fg.CellPicture <> bk.Picture Then
    GoTo chk2exit
End If

chk2sec:

For k = posr To applyr - 1
fg.Row = k
Set fg.CellPicture = bk.Picture

Next k


chk2exit:

'*******************************************************Chk3 RIGHT DIRECTION

If posc > 6 Then
    GoTo chk3exit
End If

fg.Row = posr
fg.Col = posc + 1

If fg.CellPicture = wt.Picture Then
    GoTo chk3
Else
    GoTo chk3exit
End If

chk3:

For i = fg.Col + 1 To fg.Cols - 1
    fg.Col = i
    fg.Row = posr
        
    If fg.CellPicture = wt.Picture Then
        GoTo last3
    ElseIf fg.CellPicture = bk.Picture Then
        
        applyr = fg.Row
        applyc = fg.Col
       
        
        GoTo chk3sec
    Else
        GoTo chk3exit
    End If
last3:
Next i

If fg.CellPicture <> bk.Picture Then
    GoTo chk3exit
End If

chk3sec:
    For k = posc To applyc - 1
    fg.Col = k
    Set fg.CellPicture = bk.Picture

Next k

chk3exit:

'*******************************************************Chk3 LEFT DIRECTION

If posc < 3 Then
  
    GoTo chk4exit
End If

fg.Row = posr
fg.Col = posc - 1

If fg.CellPicture = wt.Picture Then
    GoTo chk4
Else
    GoTo chk4exit
End If

chk4:

For i = fg.Col - 1 To 1 Step -1
    fg.Col = i
    fg.Row = posr
        
    If fg.CellPicture = wt.Picture Then
        GoTo last4
    ElseIf fg.CellPicture = bk.Picture Then
        applyr = fg.Row
        applyc = fg.Col
        
        GoTo chk4sec
    Else
        GoTo chk4exit
    End If
last4:
Next i

If fg.CellPicture <> bk.Picture Then
    GoTo chk4exit
End If


chk4sec:

    For k = posc To applyc + 1 Step -1
    fg.Col = k
    Set fg.CellPicture = bk.Picture

Next k


chk4exit:

'*******************************************************Chk3 TOP RIGHT DIRECTION

If posc > 6 Or posr < 3 Then

    GoTo chk5exit
End If

fg.Row = posr - 1
fg.Col = posc + 1

If fg.CellPicture = wt.Picture Then
    GoTo chk5
Else
    GoTo chk5exit
End If

chk5:

For i = fg.Row - 1 To 1 Step -1
    If fg.Col + 1 = fg.Cols Then
        GoTo chk5exit
    End If
        
    fg.Col = fg.Col + 1
    fg.Row = i
        
    If fg.CellPicture = wt.Picture Then
        GoTo last5
    ElseIf fg.CellPicture = bk.Picture Then
        applyr = fg.Row
        applyc = fg.Col
       
        GoTo chk5sec
    Else
        GoTo chk5exit
    End If
last5:
Next i

If fg.CellPicture <> bk.Picture Then
    GoTo chk5exit
End If


chk5sec:

For k = posr To applyr + 1 Step -1
    fg.Col = posc
    fg.Row = k
    Set fg.CellPicture = bk.Picture
    posc = posc + 1
Next k

chk5exit:

'*******************************************************Chk3 TOP LEFT DIRECTION

If posc < 3 Or posr < 3 Then
  
    GoTo chk6exit
End If

fg.Row = posr - 1
fg.Col = posc - 1

If fg.CellPicture = wt.Picture Then
    GoTo chk6
Else
    GoTo chk6exit
End If

chk6:

For i = fg.Row - 1 To 1 Step -1
    If fg.Col - 1 = 0 Then
        GoTo chk6exit
    End If
        
    fg.Col = fg.Col - 1
    fg.Row = i
        
    If fg.CellPicture = wt.Picture Then
        GoTo last6
    ElseIf fg.CellPicture = bk.Picture Then
        applyr = fg.Row
        applyc = fg.Col
        
        GoTo chk6sec
    Else
        GoTo chk6exit
    End If
last6:
Next i

If fg.CellPicture <> bk.Picture Then
    GoTo chk6exit
End If


chk6sec:

For k = posr To applyr + 1 Step -1
    fg.Col = posc
    fg.Row = k
    Set fg.CellPicture = bk.Picture
    posc = posc - 1
Next k

chk6exit:


'*******************************************************Chk3 BOTTOM LEFT DIRECTION

If posc < 3 Or posr > 6 Then
   
    GoTo chk7exit
End If

fg.Row = posr + 1
fg.Col = posc - 1

If fg.CellPicture = wt.Picture Then
    GoTo chk7
Else
    GoTo chk7exit
End If

chk7:

For i = fg.Row + 1 To fg.Rows - 1
    If fg.Col - 1 = 0 Then
        GoTo chk7exit
    End If

    fg.Col = fg.Col - 1
    fg.Row = i

    If fg.CellPicture = wt.Picture Then
        GoTo last7
    ElseIf fg.CellPicture = bk.Picture Then

        applyr = fg.Row
        applyc = fg.Col
        GoTo chk7sec
    Else
        GoTo chk7exit
    End If
last7:
Next i

If fg.CellPicture <> bk.Picture Then
    GoTo chk7exit
End If


chk7sec:
For k = posc To applyc + 1 Step -1
    fg.Row = posr
    fg.Col = k
    Set fg.CellPicture = bk.Picture
    posr = posr + 1
Next k


chk7exit:

If posc > 6 Or posr > 6 Then
  
    GoTo chk8exit
End If

fg.Row = posr + 1
fg.Col = posc + 1

If fg.CellPicture = wt.Picture Then
    GoTo chk8
Else
    GoTo chk8exit
End If

chk8:

For i = fg.Row + 1 To fg.Rows - 1
    If fg.Col + 1 = fg.Cols Then
        GoTo chk8exit
    End If

    fg.Col = fg.Col + 1
    fg.Row = i

    If fg.CellPicture = wt.Picture Then
        GoTo last8
    ElseIf fg.CellPicture = bk.Picture Then
        applyr = fg.Row
        applyc = fg.Col
        
        GoTo chk8sec
    Else
        GoTo chk8exit
    End If
last8:
Next i

If fg.CellPicture <> bk.Picture Then
    GoTo chk8exit
End If


chk8sec:

For k = posc To applyc - 1
    fg.Row = posr
    fg.Col = k
    Set fg.CellPicture = bk.Picture
    posr = posr + 1
Next k


chk8exit:


End Sub



Sub searchBk()


deemCounterB = 0
counterB = 1
'fg.visible = false
For i = 1 To fg.Rows - 1
    For j = 1 To fg.Cols - 1
    
    fg.Row = i
    fg.Col = j
    
    If fg.CellPicture = bk.Picture Then
        posr = i
        posc = j
        Call chk_black
    End If
    
    Next j
Next i

'reversi_adv.Cls
For k = 1 To counterB - 1
'    Print bkc(k, 1) & " , " & bkc(k, 2) & " ---- " & bkc(k, 3) & " , " & bkc(k, 4)
    fg.Row = bkc(k, 1)
    fg.Col = bkc(k, 2)
    
    If fg.CellPicture <> deem.Picture Then
        Set fg.CellPicture = deem.Picture
        deemCounterB = deemCounterB + 1
    End If
    
    
Next k
'fg.visible = true

End Sub

Sub chk_black()

REVERSI_ADV.Cls
fg.Row = posr
fg.Col = posc

If fg.CellPicture = wt.Picture Or fg.CellPicture = bg.Picture Then
    ''Msgbox "not apllied here"
    GoTo chk8exit
    
End If

'*******************************************************Chk1 UPWARD
If posr < 3 Then
   
    GoTo chk1exit
End If

fg.Row = posr - 1
fg.Col = posc

If fg.CellPicture = wt.Picture Then
    GoTo chk1
Else
    GoTo chk1exit
End If

chk1:

For i = fg.Row - 1 To 1 Step -1
    fg.Row = i
    
    
    If fg.CellPicture = wt.Picture Then
        GoTo last1
    ElseIf fg.CellPicture = bg.Picture Then
        bkc(counterB, 1) = fg.Row
        bkc(counterB, 2) = fg.Col
        bkc(counterB, 3) = posr
        bkc(counterB, 4) = posc
        counterB = counterB + 1
        GoTo chk1exit
    Else
        GoTo chk1exit
    End If
last1:
Next i

chk1exit:

'*******************************************************Chk2 DOWNWARD

If posr > 6 Then
   
    GoTo chk2exit
End If

fg.Row = posr + 1
fg.Col = posc

If fg.CellPicture = wt.Picture Then
    GoTo chk2
Else
    GoTo chk2exit
End If

chk2:

For i = fg.Row + 1 To fg.Rows - 1
    fg.Row = i
        
    If fg.CellPicture = wt.Picture Then
        GoTo last2
    ElseIf fg.CellPicture = bg.Picture Then
        
        bkc(counterB, 1) = fg.Row
        bkc(counterB, 2) = fg.Col
        bkc(counterB, 3) = posr
        bkc(counterB, 4) = posc
        counterB = counterB + 1
        GoTo chk2exit
    Else
        GoTo chk2exit
    End If
last2:
Next i

chk2exit:

'*******************************************************Chk3 RIGHT DIRECTION

If posc > 6 Then
      GoTo chk3exit
End If

fg.Row = posr
fg.Col = posc + 1

If fg.CellPicture = wt.Picture Then
    GoTo chk3
Else
    GoTo chk3exit
End If

chk3:

For i = fg.Col + 1 To fg.Cols - 1
    fg.Col = i
    fg.Row = posr
        
    If fg.CellPicture = wt.Picture Then
        GoTo last3
    ElseIf fg.CellPicture = bg.Picture Then
        bkc(counterB, 1) = fg.Row
        bkc(counterB, 2) = fg.Col
        bkc(counterB, 3) = posr
        bkc(counterB, 4) = posc
        counterB = counterB + 1
    
        GoTo chk3exit
    Else
        GoTo chk3exit
    End If
last3:
Next i

chk3exit:

'*******************************************************Chk3 LEFT DIRECTION

If posc < 3 Then

    GoTo chk4exit
End If

fg.Row = posr
fg.Col = posc - 1

If fg.CellPicture = wt.Picture Then
    GoTo chk4
Else
    GoTo chk4exit
End If

chk4:

For i = fg.Col - 1 To 1 Step -1
    fg.Col = i
    fg.Row = posr
        
    If fg.CellPicture = wt.Picture Then
        GoTo last4
    ElseIf fg.CellPicture = bg.Picture Then
        bkc(counterB, 1) = fg.Row
        bkc(counterB, 2) = fg.Col
        bkc(counterB, 3) = posr
        bkc(counterB, 4) = posc
        counterB = counterB + 1
        GoTo chk4exit
    Else
        GoTo chk4exit
    End If
last4:
Next i

chk4exit:

'*******************************************************Chk3 TOP RIGHT DIRECTION

If posc > 6 Or posr < 3 Then
    GoTo chk5exit
End If

fg.Row = posr - 1
fg.Col = posc + 1

If fg.CellPicture = wt.Picture Then
    GoTo chk5
Else
    GoTo chk5exit
End If

chk5:

For i = fg.Row - 1 To 1 Step -1
    If fg.Col + 1 = fg.Cols Then
        GoTo chk5exit
    End If
        
    fg.Col = fg.Col + 1
    fg.Row = i
        
    If fg.CellPicture = wt.Picture Then
        GoTo last5
    ElseIf fg.CellPicture = bg.Picture Then
        bkc(counterB, 1) = fg.Row
        bkc(counterB, 2) = fg.Col
        bkc(counterB, 3) = posr
        bkc(counterB, 4) = posc
        counterB = counterB + 1
        GoTo chk5exit
    Else
        GoTo chk5exit
    End If
last5:
Next i

chk5exit:

'*******************************************************Chk3 TOP LEFT DIRECTION

If posc < 3 Or posr < 3 Then
    'Msgbox "No rules can be applied 6"
    GoTo chk6exit
End If

fg.Row = posr - 1
fg.Col = posc - 1

If fg.CellPicture = wt.Picture Then
    GoTo chk6
Else
    GoTo chk6exit
End If

chk6:

For i = fg.Row - 1 To 1 Step -1
    If fg.Col - 1 = 0 Then
        GoTo chk6exit
    End If
        
    fg.Col = fg.Col - 1
    fg.Row = i
        
    If fg.CellPicture = wt.Picture Then
        GoTo last6
    ElseIf fg.CellPicture = bg.Picture Then
        bkc(counterB, 1) = fg.Row
        bkc(counterB, 2) = fg.Col
        bkc(counterB, 3) = posr
        bkc(counterB, 4) = posc
        counterB = counterB + 1
        GoTo chk6exit
    Else
        GoTo chk6exit
    End If
last6:
Next i

chk6exit:


'*******************************************************Chk3 BOTTOM LEFT DIRECTION

If posc < 3 Or posr > 6 Then
    'Msgbox "No rules can be applied 7"
    GoTo chk7exit
End If

fg.Row = posr + 1
fg.Col = posc - 1

If fg.CellPicture = wt.Picture Then
    GoTo chk7
Else
    GoTo chk7exit
End If

chk7:

For i = fg.Row + 1 To fg.Rows - 1
    If fg.Col - 1 = 0 Then
        GoTo chk7exit
    End If

    fg.Col = fg.Col - 1
    fg.Row = i

    If fg.CellPicture = wt.Picture Then
        GoTo last7
    ElseIf fg.CellPicture = bg.Picture Then
        bkc(counterB, 1) = fg.Row
        bkc(counterB, 2) = fg.Col
        bkc(counterB, 3) = posr
        bkc(counterB, 4) = posc
        counterB = counterB + 1
        GoTo chk7exit
    Else
        GoTo chk7exit
    End If
last7:
Next i

chk7exit:


'*******************************************************Chk3 BOTTOM RIGHT DIRECTION

If posc > 6 Or posr > 6 Then
    'Msgbox "No rules can be applied 8"
    GoTo chk8exit
End If

fg.Row = posr + 1
fg.Col = posc + 1

If fg.CellPicture = wt.Picture Then
    GoTo chk8
Else
    GoTo chk8exit
End If

chk8:

For i = fg.Row + 1 To fg.Rows - 1
    If fg.Col + 1 = fg.Cols Then
        GoTo chk8exit
    End If

    fg.Col = fg.Col + 1
    fg.Row = i

    If fg.CellPicture = wt.Picture Then
        GoTo last8
    ElseIf fg.CellPicture = bg.Picture Then
        bkc(counterB, 1) = fg.Row
        bkc(counterB, 2) = fg.Col
        bkc(counterB, 3) = posr
        bkc(counterB, 4) = posc
        counterB = counterB + 1
        GoTo chk8exit
    Else
        GoTo chk8exit
    End If
last8:
Next i

chk8exit:


End Sub

Sub ruleApp_wt()

'reversi_adv.Cls
fg.Row = posr
fg.Col = posc

If fg.CellPicture <> deem.Picture Then
    Exit Sub
End If

'*******************************************************Chk1 UPWARD
If posr < 3 Then
    ''Msgbox "No rules can be applied 1"
    GoTo chk1exit
End If

fg.Row = posr - 1
fg.Col = posc

If fg.CellPicture = bk.Picture Then
    GoTo chk1
Else
    GoTo chk1exit
End If

chk1:

For i = fg.Row - 1 To 1 Step -1
    fg.Row = i
    
    
    If fg.CellPicture = bk.Picture Then
        GoTo last1
        
    ElseIf fg.CellPicture = wt.Picture Then
        applyr = fg.Row
        applyc = fg.Col
        
        GoTo chk1sec
    Else
        GoTo chk1exit
    End If
last1:
Next i

If fg.CellPicture <> wt.Picture Then
    GoTo chk1exit
End If


chk1sec:

For k = posr To applyr + 1 Step -1
fg.Row = k
Set fg.CellPicture = wt.Picture

Next k

chk1exit:

'*******************************************************Chk2 DOWNWARD

If posr > 6 Then
    ''Msgbox "No rules can be applied 2"
    GoTo chk2exit
End If

fg.Row = posr + 1
fg.Col = posc

If fg.CellPicture = bk.Picture Then
    GoTo chk2
Else
    GoTo chk2exit
End If

chk2:

For i = fg.Row + 1 To fg.Rows - 1
    fg.Row = i
        
    If fg.CellPicture = bk.Picture Then
        GoTo last2
    ElseIf fg.CellPicture = wt.Picture Then
        applyr = fg.Row
        applyc = fg.Col
   
        GoTo chk2sec
    Else
        GoTo chk2exit
    End If
last2:
Next i

If fg.CellPicture <> wt.Picture Then
    GoTo chk2exit
End If


chk2sec:

For k = posr To applyr - 1
fg.Row = k
Set fg.CellPicture = wt.Picture

Next k


chk2exit:

'*******************************************************Chk3 RIGHT DIRECTION

If posc > 6 Then
    ''Msgbox "No rules can be applied 3"
    GoTo chk3exit
End If

fg.Row = posr
fg.Col = posc + 1

If fg.CellPicture = bk.Picture Then
    GoTo chk3
Else
    GoTo chk3exit
End If

chk3:

For i = fg.Col + 1 To fg.Cols - 1
    fg.Col = i
    fg.Row = posr
        
    If fg.CellPicture = bk.Picture Then
        GoTo last3
    ElseIf fg.CellPicture = wt.Picture Then
        
        applyr = fg.Row
        applyc = fg.Col
        
        GoTo chk3sec
    Else
        GoTo chk3exit
    End If
last3:
Next i

If fg.CellPicture <> wt.Picture Then
    GoTo chk3exit
End If


chk3sec:
    For k = posc To applyc - 1
    fg.Col = k
    Set fg.CellPicture = wt.Picture

Next k

chk3exit:

'*******************************************************Chk3 LEFT DIRECTION

If posc < 3 Then
    ''Msgbox "No rules can be applied 4"
    GoTo chk4exit
End If

fg.Row = posr
fg.Col = posc - 1

If fg.CellPicture = bk.Picture Then
    GoTo chk4
Else
    GoTo chk4exit
End If

chk4:

For i = fg.Col - 1 To 1 Step -1
    fg.Col = i
    fg.Row = posr
        
    If fg.CellPicture = bk.Picture Then
        GoTo last4
    ElseIf fg.CellPicture = wt.Picture Then
        applyr = fg.Row
        applyc = fg.Col
        GoTo chk4sec
    Else
        GoTo chk4exit
    End If
last4:
Next i

If fg.CellPicture <> wt.Picture Then
    GoTo chk4exit
End If


chk4sec:

    For k = posc To applyc + 1 Step -1
    fg.Col = k
    Set fg.CellPicture = wt.Picture

Next k


chk4exit:

'*******************************************************Chk3 TOP RIGHT DIRECTION

If posc > 6 Or posr < 3 Then
    ''Msgbox "No rules can be applied 5"
    GoTo chk5exit
End If

fg.Row = posr - 1
fg.Col = posc + 1

If fg.CellPicture = bk.Picture Then
    GoTo chk5
Else
    GoTo chk5exit
End If

chk5:

For i = fg.Row - 1 To 1 Step -1
    If fg.Col + 1 = fg.Cols Then
        GoTo chk5exit
    End If
        
    fg.Col = fg.Col + 1
    fg.Row = i
        
    If fg.CellPicture = bk.Picture Then
        GoTo last5
    ElseIf fg.CellPicture = wt.Picture Then
        applyr = fg.Row
        applyc = fg.Col
        GoTo chk5sec
    Else
        GoTo chk5exit
    End If
last5:
Next i
    If fg.CellPicture <> wt.Picture Then
        GoTo chk5exit
    End If


chk5sec:

For k = posr To applyr + 1 Step -1
    fg.Col = posc
    fg.Row = k
    Set fg.CellPicture = wt.Picture
    posc = posc + 1
Next k

chk5exit:

'*******************************************************Chk3 TOP LEFT DIRECTION

If posc < 3 Or posr < 3 Then
    ''Msgbox "No rules can be applied 6"
    GoTo chk6exit
End If

fg.Row = posr - 1
fg.Col = posc - 1

If fg.CellPicture = bk.Picture Then
    GoTo chk6
Else
    GoTo chk6exit
End If

chk6:

For i = fg.Row - 1 To 1 Step -1
    If fg.Col - 1 = 0 Then
        GoTo chk6exit
    End If
        
    fg.Col = fg.Col - 1
    fg.Row = i
        
    If fg.CellPicture = bk.Picture Then
        GoTo last6
    ElseIf fg.CellPicture = wt.Picture Then
        applyr = fg.Row
        applyc = fg.Col
        GoTo chk6sec
    Else
        GoTo chk6exit
    End If
last6:
Next i

    If fg.CellPicture <> wt.Picture Then
        GoTo chk6exit
    End If


chk6sec:

For k = posr To applyr + 1 Step -1
    fg.Col = posc
    fg.Row = k
    Set fg.CellPicture = wt.Picture
    posc = posc - 1
Next k

chk6exit:


'*******************************************************Chk3 BOTTOM LEFT DIRECTION

If posc < 3 Or posr > 6 Then
    ''Msgbox "No rules can be applied 7"
    GoTo chk7exit
End If

fg.Row = posr + 1
fg.Col = posc - 1

If fg.CellPicture = bk.Picture Then
    GoTo chk7
Else
    GoTo chk7exit
End If

chk7:

For i = fg.Row + 1 To fg.Rows - 1
    If fg.Col - 1 = 0 Then
        GoTo chk7exit
    End If

    fg.Col = fg.Col - 1
    fg.Row = i

    If fg.CellPicture = bk.Picture Then
        GoTo last7
    ElseIf fg.CellPicture = wt.Picture Then

        applyr = fg.Row
        applyc = fg.Col
        GoTo chk7sec
    Else
        GoTo chk7exit
    End If
last7:
Next i

If fg.CellPicture <> wt.Picture Then
    GoTo chk7exit
End If


chk7sec:
For k = posc To applyc + 1 Step -1
    fg.Row = posr
    fg.Col = k
    Set fg.CellPicture = wt.Picture
    posr = posr + 1
Next k


chk7exit:

If posc > 6 Or posr > 6 Then
    ''Msgbox "No rules can be applied 8"
    GoTo chk8exit
End If

fg.Row = posr + 1
fg.Col = posc + 1

If fg.CellPicture = bk.Picture Then
    GoTo chk8
Else
    GoTo chk8exit
End If

chk8:

For i = fg.Row + 1 To fg.Rows - 1
    If fg.Col + 1 = fg.Cols Then
        GoTo chk8exit
    End If

    fg.Col = fg.Col + 1
    fg.Row = i

    If fg.CellPicture = bk.Picture Then
        GoTo last8
    ElseIf fg.CellPicture = wt.Picture Then
        applyr = fg.Row
        applyc = fg.Col
        
        GoTo chk8sec
    Else
        GoTo chk8exit
    End If
last8:
Next i

If fg.CellPicture <> wt.Picture Then
    GoTo chk8exit
End If



chk8sec:

For k = posc To applyc - 1
    fg.Row = posr
    fg.Col = k
    Set fg.CellPicture = wt.Picture
    posr = posr + 1
Next k


chk8exit:


End Sub


Sub chk_white()
fg.Row = posr
fg.Col = posc


If fg.CellPicture = bk.Picture Or fg.CellPicture = bg.Picture Then
    ''Msgbox "not apllied here"
    GoTo chk8exit
    
End If

'*******************************************************Chk1 UPWARD
If posr < 3 Then
    ''Msgbox "No rules can be applied 1"
    GoTo chk1exit
End If

fg.Row = posr - 1
fg.Col = posc

If fg.CellPicture = bk.Picture Then
    GoTo chk1
Else
    GoTo chk1exit
End If

chk1:

For i = fg.Row - 1 To 1 Step -1
    fg.Row = i
    
    
    If fg.CellPicture = bk.Picture Then
        GoTo last1
    ElseIf fg.CellPicture = bg.Picture Then
        wtc(counter, 1) = fg.Row
        wtc(counter, 2) = fg.Col
        wtc(counter, 3) = posr
        wtc(counter, 4) = posc
        counter = counter + 1
        
        GoTo chk1exit
    Else
        GoTo chk1exit
    End If
last1:
Next i

chk1exit:

'*******************************************************Chk2 DOWNWARD

If posr > 6 Then
    ''Msgbox "No rules can be applied 2"
    GoTo chk2exit
End If

fg.Row = posr + 1
fg.Col = posc

If fg.CellPicture = bk.Picture Then
    GoTo chk2
Else
    GoTo chk2exit
End If

chk2:

For i = fg.Row + 1 To fg.Rows - 1
    fg.Row = i
        
    If fg.CellPicture = bk.Picture Then
        GoTo last2
    ElseIf fg.CellPicture = bg.Picture Then
        
        wtc(counter, 1) = fg.Row
        wtc(counter, 2) = fg.Col
        wtc(counter, 3) = posr
        wtc(counter, 4) = posc
        counter = counter + 1
        
        GoTo chk2exit
    Else
        GoTo chk2exit
    End If
last2:
Next i

chk2exit:

'*******************************************************Chk3 RIGHT DIRECTION

If posc > 6 Then
    ''Msgbox "No rules can be applied 3"
    GoTo chk3exit
End If

fg.Row = posr
fg.Col = posc + 1

If fg.CellPicture = bk.Picture Then
    GoTo chk3
Else
    GoTo chk3exit
End If

chk3:

For i = fg.Col + 1 To fg.Cols - 1
    fg.Col = i
    fg.Row = posr
        
    If fg.CellPicture = bk.Picture Then
        GoTo last3
    ElseIf fg.CellPicture = bg.Picture Then
        
        wtc(counter, 1) = fg.Row
        wtc(counter, 2) = fg.Col
        wtc(counter, 3) = posr
        wtc(counter, 4) = posc
        counter = counter + 1
        
        GoTo chk3exit
    Else
        GoTo chk3exit
    End If
last3:
Next i

chk3exit:

'*******************************************************Chk3 LEFT DIRECTION

If posc < 3 Then
    ''Msgbox "No rules can be applied 4"
    GoTo chk4exit
End If

fg.Row = posr
fg.Col = posc - 1

If fg.CellPicture = bk.Picture Then
    GoTo chk4
Else
    GoTo chk4exit
End If

chk4:

For i = fg.Col - 1 To 1 Step -1
    fg.Col = i
    fg.Row = posr
        
    If fg.CellPicture = bk.Picture Then
        GoTo last4
    ElseIf fg.CellPicture = bg.Picture Then
        wtc(counter, 1) = fg.Row
        wtc(counter, 2) = fg.Col
        wtc(counter, 3) = posr
        wtc(counter, 4) = posc
        counter = counter + 1
        GoTo chk4exit
    Else
        GoTo chk4exit
    End If
last4:
Next i

chk4exit:

'*******************************************************Chk3 TOP RIGHT DIRECTION

If posc > 6 Or posr < 3 Then
    ''Msgbox "No rules can be applied 5"
    GoTo chk5exit
End If

fg.Row = posr - 1
fg.Col = posc + 1

If fg.CellPicture = bk.Picture Then
    GoTo chk5
Else
    GoTo chk5exit
End If

chk5:

For i = fg.Row - 1 To 1 Step -1
    If fg.Col + 1 = fg.Cols Then
        GoTo chk5exit
    End If
        
    fg.Col = fg.Col + 1
    fg.Row = i
        
    If fg.CellPicture = bk.Picture Then
        GoTo last5
    ElseIf fg.CellPicture = bg.Picture Then
        wtc(counter, 1) = fg.Row
        wtc(counter, 2) = fg.Col
        wtc(counter, 3) = posr
        wtc(counter, 4) = posc
        counter = counter + 1
        GoTo chk5exit
    Else
        GoTo chk5exit
    End If
last5:
Next i

chk5exit:

'*******************************************************Chk3 TOP LEFT DIRECTION

If posc < 3 Or posr < 3 Then
    ''Msgbox "No rules can be applied 6"
    GoTo chk6exit
End If

fg.Row = posr - 1
fg.Col = posc - 1

If fg.CellPicture = bk.Picture Then
    GoTo chk6
Else
    GoTo chk6exit
End If

chk6:

For i = fg.Row - 1 To 1 Step -1
    If fg.Col - 1 = 0 Then                         'Error here
        GoTo chk6exit
    End If
        
    fg.Col = fg.Col - 1
    fg.Row = i
        
    If fg.CellPicture = bk.Picture Then
        GoTo last6
    ElseIf fg.CellPicture = bg.Picture Then
        wtc(counter, 1) = fg.Row
        wtc(counter, 2) = fg.Col
        wtc(counter, 3) = posr
        wtc(counter, 4) = posc
        counter = counter + 1
        GoTo chk6exit
    Else
        GoTo chk6exit
    End If
last6:
Next i

chk6exit:


'*******************************************************Chk3 BOTTOM LEFT DIRECTION

If posc < 3 Or posr > 6 Then
    ''Msgbox "No rules can be applied 7"
    GoTo chk7exit
End If

fg.Row = posr + 1
fg.Col = posc - 1

If fg.CellPicture = bk.Picture Then
    GoTo chk7
Else
    GoTo chk7exit
End If

chk7:

For i = fg.Row + 1 To fg.Rows - 1

    If fg.Col - 1 = 0 Then
        GoTo chk7exit
    End If

    fg.Col = fg.Col - 1
    fg.Row = i

    If fg.CellPicture = bk.Picture Then
        GoTo last7
    ElseIf fg.CellPicture = bg.Picture Then
        wtc(counter, 1) = fg.Row
        wtc(counter, 2) = fg.Col
        wtc(counter, 3) = posr
        wtc(counter, 4) = posc
        counter = counter + 1
        GoTo chk7exit
    Else
        GoTo chk7exit
    End If
last7:
Next i

chk7exit:


'*******************************************************Chk3 BOTTOM RIGHT DIRECTION

If posc > 6 Or posr > 6 Then
    ''Msgbox "No rules can be applied 8"
    GoTo chk8exit
End If

fg.Row = posr + 1
fg.Col = posc + 1

If fg.CellPicture = bk.Picture Then
    GoTo chk8
Else
    GoTo chk8exit
End If

chk8:

For i = fg.Row + 1 To fg.Rows - 1
    If fg.Col + 1 = fg.Cols Then
        GoTo chk8exit
    End If

    fg.Col = fg.Col + 1
    fg.Row = i

    If fg.CellPicture = bk.Picture Then
        GoTo last8
    ElseIf fg.CellPicture = bg.Picture Then
        wtc(counter, 1) = fg.Row
        wtc(counter, 2) = fg.Col
        wtc(counter, 3) = posr
        wtc(counter, 4) = posc
        counter = counter + 1
        GoTo chk8exit
    Else
        GoTo chk8exit
    End If
last8:
Next i

chk8exit:

End Sub

Private Sub cmd_applybk_Click()
Call ruleApp_bk

End Sub

Private Sub cmd_applyWt_Click()
Call ruleApp_wt
End Sub

Private Sub cmd_BegOpp_Click()
Call BegTopOpp
End Sub

Private Sub cmd_BegTop_Click()
Call BegTop
End Sub

Private Sub cmd_chkrules_Click()

Call chk_white

End Sub

Private Sub cmd_counting_Click()
Call counting
Label5.Caption = countbk
Label6.Caption = countwt
End Sub

Private Sub cmd_distroyDeem_Click()
Call distroyDeem
End Sub

Private Sub cmd_level_Click()
var_level = Val(LEVEL.Text)

If var_level = 1 Then

    Label7.Caption = "REVERSI ADVANCE"
ElseIf var_level = 2 Then

    Label7.Caption = "REVERSI TWO PLAYERS"
ElseIf var_level = 3 Then

    Label7.Caption = "REVERSI BEGINNERS"

End If

End Sub

Private Sub cmd_restart_Click()
Call Form_Load
End Sub

Private Sub cmd_retainState_Click()
Call retainState
End Sub

Private Sub cmd_storeState_Click()
Call storeState
End Sub



Private Sub cmd_suggestion_Click()
fg.Visible = False
Call BegTopOpp
fg.Visible = True
MsgBox "I would suggest you to play Move at Raw,Column = " & BegArrOpp(1, 1) & "," & BegArrOpp(1, 2)

End Sub

Private Sub fg_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
fg.Visible = False
posc = Int(((X - fg.ColWidth(0)) / fg.ColWidth(1)) + 1)
posr = Int(((Y - fg.RowHeight(0)) / fg.RowHeight(1)) + 1)

fg.Row = posr
fg.Col = posc

If fg.CellPicture <> deem.Picture Then
    GoTo endmark
    
End If

If deemCounterB = 0 And deemCounterW = 0 Then
    MsgBox "GAME OVER"
    GoTo endmark
End If

Skipturn2:
If turn = 1 Then


Call ruleApp_bk


Call distroyDeem
Call searchWt


    If deemCounterB = 0 And deemCounterW = 0 Then
    GoTo endmark
    End If

turn = 2
If deemCounterW = 0 Then
    GoTo Skipturn1
End If

txt_turn.Text = turn
Set tn.Picture = wtpic

GoTo endmark

End If

Skipturn1:
If turn = 2 Then
'Black's Turn
Call ruleApp_wt
Call distroyDeem
'Call chk_black
Call searchBk

    If deemCounterB = 0 And deemCounterW = 0 Then
    
    GoTo endmark
    End If
    
turn = 1
    If deemCounterB = 0 Then
        GoTo Skipturn2
    End If

txt_turn.Text = turn
Set tn.Picture = bkpic


If var_level = 1 Then
    Call twoPlyFun
    Label7.Caption = "REVERSI ADVANCE"
ElseIf var_level = 2 Then
    Call BegTop
    Label7.Caption = "REVERSI TWO PLAYERS"
ElseIf var_level = 3 Then
    Call BegTop1
    Label7.Caption = "REVERSI BEGINNERS"

End If



GoTo endmark

'Call turnChanger
End If


endmark:
'abc = 0

Call counting
'label6.caption = countwt

fg.Visible = True

If deemCounterB = 0 And deemCounterW = 0 Then
    If Val(Label5.Caption) < Val(Label6.Caption) Then
        LBL_STATUS.Caption = "GAME OVER: YOU WIN"
    ElseIf Val(Label5.Caption) > Val(Label6.Caption) Then
        LBL_STATUS.Caption = "GAME OVER: YOU LOSE"
    Else
        LBL_STATUS.Caption = "GAME OVER: MATCH IS DRAWN"
    End If
 End If

End Sub

Private Sub fg_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
                                                                
Text1.Text = Int(((X - fg.ColWidth(0)) / fg.ColWidth(1)) + 1)
Text2.Text = Int(((Y - fg.RowHeight(0)) / fg.RowHeight(1)) + 1)

abc = abc + 1

If abc <> 1 Then
    Exit Sub
End If
    If turn = 1 Then
    
        Call distroyDeem
        Call searchBk
   End If
    If turn = 2 Then
        LBL_STATUS.Caption = "GAME IN PROGRESS"
        fg.Visible = False
        Call distroyDeem
        Call searchWt
        fg.Visible = True
    End If
                    
End Sub

Private Sub Form_Load()
var_level = Val(LEVEL.Text)
abc = 0

Label5.Caption = 2
Label6.Caption = 2

deemCounterB = 0
deemCounterW = 0

Set bkpic = LoadPicture(App.Path & "\black but.jpg")
Set wtpic = LoadPicture(App.Path & "\white but.jpg")
Set bgpic = LoadPicture(App.Path & "\background.jpg")
fg.Visible = False
'Call cmd_random_Click
turn = 2
tn.Picture = wt.Picture
txt_turn.Text = turn
fg.ColWidth(0) = 300

fg.Height = 6250
fg.Width = 6280
        
For i = 1 To fg.Rows - 1
    For j = 1 To fg.Cols - 1

        fg.Row = i
        fg.Col = j
        
        fg.ColWidth(j) = bg.Width
        fg.RowHeight(i) = bg.Height
        
        Set fg.CellPicture = bg.Picture

        
    Next j
Next i


For i = 1 To fg.Cols - 1
    
    fg.TextMatrix(0, i) = "    " & i
    fg.TextMatrix(i, 0) = i
    
   
Next i


'********************************************************Initial Board Pos

fg.Row = 5
fg.Col = 4
Set fg.CellPicture = wt.Picture
fg.Row = 4
fg.Col = 5
Set fg.CellPicture = wt.Picture

fg.Row = 4
fg.Col = 4
Set fg.CellPicture = bk.Picture
fg.Row = 5
fg.Col = 5
Set fg.CellPicture = bk.Picture

fg.Visible = True
End Sub

Private Sub MNU_2P_Click()
yn = MsgBox("This Action Would Reset the Match. Do You Want To Continue?", vbYesNo + vbInformation)
If yn = vbYes Then
Call Form_Load
LEVEL.Text = 2
cmd_level_Click
End If
End Sub

Private Sub MNU_ADV_Click()
yn = MsgBox("This Action Would Reset the Match. Do You Want To Continue?", vbYesNo + vbInformation)
If yn = vbYes Then
Call Form_Load
LEVEL.Text = 1
cmd_level_Click
End If
End Sub

Private Sub MNU_BEG_Click()
yn = MsgBox("This Action Would Reset the Match. Do You Want To Continue?", vbYesNo + vbInformation)
If yn = vbYes Then
Call Form_Load
LEVEL.Text = 3
cmd_level_Click
End If
End Sub

Private Sub MNU_CREDIT_Click()
Form1.Show
End Sub

Private Sub MNU_EXIT_Click()
Unload Me
End Sub

Private Sub MNU_NEW_Click()
fg.Enabled = True
Call Form_Load
End Sub

Private Sub MNU_RESIGN_Click()
yn = MsgBox("Are you sure you would like to RESIGN?", vbYesNo + vbInformation)
If yn = vbYes Then
    LBL_STATUS.Caption = "You Loss The Match... (Resigned)"
    fg.Enabled = False
End If
End Sub
