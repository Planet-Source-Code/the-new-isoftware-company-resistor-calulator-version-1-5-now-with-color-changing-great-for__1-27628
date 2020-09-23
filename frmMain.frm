VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H0000C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resistor Calculator"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
   BeginProperty Font 
      Name            =   "Terminal"
      Size            =   9
      Charset         =   255
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   6630
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C000&
      Caption         =   "Calculate"
      Height          =   255
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2640
      Width           =   1575
   End
   Begin VB.ComboBox tolerance 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      ItemData        =   "frmMain.frx":058A
      Left            =   5160
      List            =   "frmMain.frx":0597
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2280
      Width           =   1335
   End
   Begin VB.ComboBox multiplier 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      ItemData        =   "frmMain.frx":05B1
      Left            =   3480
      List            =   "frmMain.frx":05D9
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2280
      Width           =   1575
   End
   Begin VB.ComboBox Band2 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      ItemData        =   "frmMain.frx":0630
      Left            =   1800
      List            =   "frmMain.frx":0652
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2280
      Width           =   1575
   End
   Begin VB.ComboBox Band1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      ItemData        =   "frmMain.frx":069B
      Left            =   120
      List            =   "frmMain.frx":06BD
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Note:  If the color is not listed, reverse the resitor."
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   4080
      Width           =   6615
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   6600
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   2640
      TabIndex        =   11
      Top             =   3720
      Width           =   150
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   2640
      TabIndex        =   10
      Top             =   3360
      Width           =   150
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   2640
      TabIndex        =   9
      Top             =   3000
      Width           =   150
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Min. Value (ohms):"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3720
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Max. Value (ohms):"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3360
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Reg. Value (ohms):"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3000
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Resistor Calulator"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   600
      TabIndex        =   3
      Top             =   720
      Width           =   1695
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   15
      X1              =   4440
      X2              =   4440
      Y1              =   360
      Y2              =   1920
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00004080&
      BorderWidth     =   15
      X1              =   3720
      X2              =   3720
      Y1              =   360
      Y2              =   1920
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      BorderWidth     =   15
      X1              =   3000
      X2              =   3000
      Y1              =   360
      Y2              =   1920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00004080&
      BorderWidth     =   15
      X1              =   2280
      X2              =   2280
      Y1              =   360
      Y2              =   1920
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0080C0FF&
      FillStyle       =   0  'Solid
      Height          =   2055
      Left            =   3600
      Shape           =   3  'Circle
      Top             =   120
      Width           =   3375
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0080C0FF&
      FillStyle       =   0  'Solid
      Height          =   2055
      Left            =   -240
      Shape           =   3  'Circle
      Top             =   120
      Width           =   3375
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0080C0FF&
      FillStyle       =   0  'Solid
      Height          =   1815
      Left            =   960
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   4575
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00808080&
      BorderWidth     =   30
      X1              =   480
      X2              =   0
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00808080&
      BorderWidth     =   30
      X1              =   6960
      X2              =   6240
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line7 
      X1              =   2280
      X2              =   960
      Y1              =   2040
      Y2              =   2520
   End
   Begin VB.Line Line8 
      X1              =   2640
      X2              =   3000
      Y1              =   2280
      Y2              =   2040
   End
   Begin VB.Line Line9 
      X1              =   3600
      X2              =   3960
      Y1              =   1920
      Y2              =   2400
   End
   Begin VB.Line Line10 
      X1              =   4440
      X2              =   5520
      Y1              =   2040
      Y2              =   2400
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim value1 As Variant
Dim value2 As Variant
Dim value3 As Variant
Dim value4 As Variant
Dim value5 As Variant
Dim value6 As Variant



Private Sub Band1_Click()
Select Case Band1.Text
Case "Black"
Line1.BorderColor = vbBlack
Case "Brown"
Line1.BorderColor = &H4080&
Case "Red"
Line1.BorderColor = vbRed
Case "Orange"
Line1.BorderColor = &H80FF&
Case "Yellow"
Line1.BorderColor = &HFFFF&
Case "Green"
Line1.BorderColor = &HC000&
Case "Blue"
Line1.BorderColor = &HFF0000
Case "Purple"
Line1.BorderColor = &HC000C0
Case "Gray"
Line1.BorderColor = &HC0C0C0
Case "White"
Line1.BorderColor = &HFFFFFF
End Select
End Sub

Private Sub Band2_Click()
Select Case Band2.Text
Case "Black"
Line2.BorderColor = vbBlack
Case "Brown"
Line2.BorderColor = &H4080&
Case "Red"
Line2.BorderColor = vbRed
Case "Orange"
Line2.BorderColor = &H80FF&
Case "Yellow"
Line2.BorderColor = &HFFFF&
Case "Green"
Line2.BorderColor = &HC000&
Case "Blue"
Line2.BorderColor = &HFF0000
Case "Purple"
Line2.BorderColor = &HC000C0
Case "Gray"
Line2.BorderColor = &HC0C0C0
Case "White"
Line2.BorderColor = &HFFFFFF
End Select
End Sub

Private Sub Command1_Click()
Select Case Band1.Text
Case "Black"
value1 = 0
Case "Brown"
value1 = 1
Case "Red"
value1 = 2
Case "Orange"
value1 = 3
Case "Yellow"
value1 = 4
Case "Green"
value1 = 5
Case "Blue"
value1 = 6
Case "Purple"
value1 = 7
Case "Gray"
value1 = 8
Case "White"
value1 = 9
Case ""
MsgBox "You must fill out the first band field.", vbCritical, "R. C."
Exit Sub
End Select
Select Case Band2.Text
Case "Black"
value2 = 0
Case "Brown"
value2 = 1
Case "Red"
value2 = 2
Case "Orange"
value2 = 3
Case "Yellow"
value2 = 4
Case "Green"
value2 = 5
Case "Blue"
value2 = 6
Case "Purple"
value2 = 7
Case "Gray"
value2 = 8
Case "White"
value2 = 9
Case ""
MsgBox "You must fill out the second band field.", vbCritical, "R. C."
Exit Sub

End Select
value3 = value1 & value2
Select Case Band2.Text
Case "Black"
value2 = 0
Case "Brown"
value2 = 1
Case "Red"
value2 = 2
Case "Orange"
value2 = 3
Case "Yellow"
value2 = 4
Case "Green"
value2 = 5
Case "Blue"
value2 = 6
Case "Purple"
value2 = 7
Case "Gray"
value2 = 8
Case "White"
value2 = 9
Case ""
MsgBox "You must fill out the second band field.", vbCritical, "R. C."
Exit Sub

End Select
Select Case multiplier.Text
Case "Black"
value4 = 1
Case "Brown"
value4 = 10
Case "Red"
value4 = 100
Case "Orange"
value4 = 1000
Case "Yellow"
value4 = 10000
Case "Green"
value4 = 100000
Case "Blue"
value4 = 1000000
Case "Purple"
value4 = 10000000
Case "Gray"
value4 = 100000000
Case "White"
value4 = 1000000000
Case "Gold"
value4 = 0.1
Case "Silver"
value4 = 0.01
Case ""
MsgBox "You must fill out the third band field.", vbCritical, "R. C."
Exit Sub

End Select
Label5.Caption = value3 * value4 & " ê"
Select Case tolerance.Text
Case "(None)"
value5 = value3 * value4 * 0.2
value6 = value5 + value3 * value4
Label6.Caption = value6 & " ê"
Label7.Caption = value3 * value4 - value5 & " ê"
Case "Silver"
value5 = value3 * value4 * 0.1
value6 = value5 + value3 * value4
Label6.Caption = value6 & " ê"
Label7.Caption = value3 * value4 - value5 & " ê"
Case "Gold"
value5 = value3 * value4 * 0.05
value6 = value5 + value3 * value4
Label6.Caption = value6 & " ê"
Label7.Caption = value3 * value4 - value5 & " ê"
Case ""
MsgBox "You must fill out the fourth band field.", vbCritical, "R. C."
End Select

End Sub

Private Sub multiplier_Click()
Select Case multiplier.Text
Case "Black"
Line3.BorderColor = vbBlack
Case "Brown"
Line3.BorderColor = &H4080&
Case "Red"
Line3.BorderColor = vbRed
Case "Orange"
Line3.BorderColor = &H80FF&
Case "Yellow"
Line3.BorderColor = &HFFFF&
Case "Green"
Line3.BorderColor = &HC000&
Case "Blue"
Line3.BorderColor = &HFF0000
Case "Purple"
Line3.BorderColor = &HC000C0
Case "Gray"
Line3.BorderColor = &HC0C0C0
Case "White"
Line3.BorderColor = &HFFFFFF
Case "Silver"
Line4.BorderColor = &HE0E0E0
Case "Gold"
Line4.BorderColor = &H80FFFF
End Select
End Sub

Private Sub tolerance_Click()
Select Case tolerance.Text
Case "(None)"
Line4.Visible = False
Case "Silver"
Line4.Visible = True
Line4.BorderColor = &HE0E0E0
Case "Gold"
Line4.Visible = True
Line4.BorderColor = &H80FFFF

End Select

End Sub
