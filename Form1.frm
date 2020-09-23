VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "BUBBLE SORT"
   ClientHeight    =   5265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   ScaleHeight     =   5265
   ScaleWidth      =   5160
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGo 
      Caption         =   "Do BubbleSort"
      Height          =   375
      Left            =   3240
      TabIndex        =   10
      Top             =   4800
      Width           =   1335
   End
   Begin VB.PictureBox picArw 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   360
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   1
      Top             =   720
      Width           =   495
   End
   Begin VB.Label lblNew 
      Height          =   255
      Index           =   7
      Left            =   3240
      TabIndex        =   26
      Top             =   4200
      Width           =   495
   End
   Begin VB.Label lblNew 
      Height          =   255
      Index           =   6
      Left            =   3240
      TabIndex        =   25
      Top             =   3720
      Width           =   495
   End
   Begin VB.Label lblNew 
      Height          =   255
      Index           =   5
      Left            =   3240
      TabIndex        =   24
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label lblNew 
      Height          =   255
      Index           =   4
      Left            =   3240
      TabIndex        =   23
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label lblNew 
      Height          =   255
      Index           =   3
      Left            =   3240
      TabIndex        =   22
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label lblNew 
      Height          =   255
      Index           =   2
      Left            =   3240
      TabIndex        =   21
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label lblNew 
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   20
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label lblNew 
      Height          =   255
      Index           =   0
      Left            =   3240
      TabIndex        =   19
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label9 
      Caption         =   "(7)"
      Height          =   255
      Left            =   960
      TabIndex        =   18
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Label8 
      Caption         =   "(6)"
      Height          =   255
      Left            =   960
      TabIndex        =   17
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label7 
      Caption         =   "(5)"
      Height          =   255
      Left            =   960
      TabIndex        =   16
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Label6 
      Caption         =   "(4)"
      Height          =   255
      Left            =   960
      TabIndex        =   15
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "(3)"
      Height          =   255
      Left            =   960
      TabIndex        =   14
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "(2)"
      Height          =   255
      Left            =   960
      TabIndex        =   13
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "(1)"
      Height          =   255
      Left            =   960
      TabIndex        =   12
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "(0)"
      Height          =   255
      Left            =   960
      TabIndex        =   11
      Top             =   840
      Width           =   255
   End
   Begin VB.Label lblOrig 
      Caption         =   "77.7"
      Height          =   255
      Index           =   7
      Left            =   2040
      TabIndex        =   9
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label lblOrig 
      Caption         =   "33.3"
      Height          =   255
      Index           =   6
      Left            =   2040
      TabIndex        =   8
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label lblOrig 
      Caption         =   "88.8"
      Height          =   255
      Index           =   5
      Left            =   2040
      TabIndex        =   7
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label lblOrig 
      Caption         =   "44.4"
      Height          =   255
      Index           =   4
      Left            =   2040
      TabIndex        =   6
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label lblOrig 
      Caption         =   "66.6"
      Height          =   255
      Index           =   3
      Left            =   2040
      TabIndex        =   5
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label lblOrig 
      Caption         =   "99.9"
      Height          =   255
      Index           =   2
      Left            =   2040
      TabIndex        =   4
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label lblOrig 
      Caption         =   "22.5"
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   3
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label lblOrig 
      Caption         =   "55.5"
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   2
      Top             =   840
      Width           =   375
   End
   Begin VB.Line Line7 
      X1              =   2880
      X2              =   2880
      Y1              =   720
      Y2              =   4560
   End
   Begin VB.Line Line6 
      X1              =   1560
      X2              =   4200
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line5 
      X1              =   1560
      X2              =   4200
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line4 
      X1              =   1560
      X2              =   4200
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line3 
      X1              =   1560
      X2              =   4200
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line2 
      X1              =   1560
      X2              =   4200
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line1 
      X1              =   1560
      X2              =   4200
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Shape Shape2 
      Height          =   3855
      Left            =   1560
      Top             =   720
      Width           =   2655
   End
   Begin VB.Shape Shape1 
      Height          =   3375
      Left            =   1560
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "BUBBLE SORT"
      Height          =   255
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim arr() 'an array filled with some values to sort
   
Private Sub cmdGo_Click()
BubbleSort arr
End Sub

Private Sub Form_Load()
arr = Array(55.5, 22.5, 99.9, 66.6, 44.4, 88.8, 33.3, 77.7)
End Sub

Private Function BubbleSort(a() As Variant)
For I = 1 To UBound(a)

    DoEvents
    
    For j = 0 To UBound(a) - I
    
        DoEvents
        
        If a(j) > a(j + 1) Then 'if j element is more than the element
                                'in front of j element,
            
            temp = a(j)         'then SWAP EM!
            
            a(j) = a(j + 1)
            lblNew(j).Caption = a(j + 1)
            a(j + 1) = temp
            lblNew(j + 1) = temp
            
            MoveThatArrow j
            Sleep 200
            
            
        End If
        
    Next j
Next I
picArw.Visible = False
End Function

Private Sub MoveThatArrow(ByVal pos As Integer) '0 thru 7
Const movevalue As Integer = 480
picArw.Top = 720 + (480 * pos)
Me.Refresh
End Sub

