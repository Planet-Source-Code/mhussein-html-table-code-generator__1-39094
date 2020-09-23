VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HTML Table Code Generator"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10440
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   10440
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Table Code"
      Height          =   2655
      Left            =   120
      TabIndex        =   7
      Top             =   4680
      Width           =   10215
      Begin MSComctlLib.ProgressBar PB1 
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.TextBox txtCode 
         Height          =   1815
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   720
         Width           =   9975
      End
      Begin VB.CommandButton cmdGenerateCode 
         Caption         =   "&Generate Table Code"
         Height          =   375
         Left            =   8280
         TabIndex        =   3
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Table"
      Height          =   4455
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   10215
      Begin MSComctlLib.Slider sldCols 
         Height          =   255
         Left            =   720
         TabIndex        =   14
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Min             =   1
         Max             =   12
         SelStart        =   1
         Value           =   1
      End
      Begin MSComctlLib.Slider sldRows 
         Height          =   255
         Left            =   2640
         TabIndex        =   13
         Top             =   240
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   450
         _Version        =   393216
         Min             =   1
         Max             =   60
         SelStart        =   5
         Value           =   1
      End
      Begin VB.CheckBox chkHeader 
         Caption         =   "Table Header"
         Height          =   255
         Left            =   8760
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdFGUpdate 
         Caption         =   "&Update"
         Height          =   375
         Left            =   8280
         TabIndex        =   2
         Top             =   3960
         Width           =   1815
      End
      Begin VB.TextBox txtFGUpdate 
         Height          =   370
         Left            =   120
         TabIndex        =   1
         Top             =   3960
         Width           =   8055
      End
      Begin MSFlexGridLib.MSFlexGrid FG1 
         Height          =   3015
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   5318
         _Version        =   393216
         Rows            =   1
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         AllowUserResizing=   3
      End
      Begin VB.Label Label2 
         Caption         =   "Rows"
         Height          =   255
         Left            =   2160
         TabIndex        =   9
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Columns"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Cell Text"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   3720
         Width           =   1935
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Mohamed Hussein - IT Department"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   7440
      Width           =   4455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkHeader_Click()
If chkHeader.Value = 1 And FG1.Rows > 1 Then
FG1.FixedRows = 1
ElseIf FG1.Rows > 1 Then
FG1.FixedRows = 0
Else
chkHeader.Value = 0
End If
End Sub

Private Sub cmdFGUpdate_Click()
If FG1.Rows <> 0 And FG1.Cols <> 0 Then FG1.TextMatrix(FG1.Row, FG1.Col) = txtFGUpdate.Text
End Sub

Private Sub cmdGenerateCode_Click()
Dim TempVal As String
Dim T As Integer
Dim R As Integer
If FG1.Cols <> 0 And FG1.Rows <> 0 Then
PB1.Max = FG1.Cols * FG1.Rows
PB1.Value = 0
TempVal = "<table width=" & Chr$(34) & "100%" & Chr$(34) & " border=" & Chr$(34) & "1" & Chr$(34) & " cellspacing=" & Chr$(34) & "0" & Chr$(34) & " cellpadding=" & Chr$(34) & "0" & Chr$(34) & ">"
For T = 0 To FG1.Rows - 1
TempVal = TempVal & "<tr>"
For R = 0 To FG1.Cols - 1
TempVal = TempVal & "<td>"
If T = 0 And chkHeader.Value = 1 Then TempVal = TempVal & "<b>"
TempVal = TempVal & IIf(FG1.TextMatrix(T, R) <> "", FG1.TextMatrix(T, R), "&nbsp;")
If T = 0 And chkHeader.Value = 1 Then TempVal = TempVal & "</b>"
TempVal = TempVal & "</td>"
PB1.Value = PB1.Value + 1
Next R
TempVal = TempVal & "</tr>"
Next T
TempVal = TempVal & "</table>"
txtCode.Text = TempVal
End If
End Sub

Private Sub FG1_Click()
txtFGUpdate.Text = FG1.TextMatrix(FG1.Row, FG1.Col)
End Sub

Private Sub FG1_DblClick()
FG1_Click
txtFGUpdate.SetFocus
End Sub

Private Sub sldCols_Change()
FG1.Cols = sldCols.Value
End Sub

Private Sub sldRows_Change()
FG1.Rows = sldRows.Value
End Sub

Private Sub txtFGUpdate_LostFocus()
cmdFGUpdate.Default = False
End Sub

Private Sub txtFGUpdate_GotFocus()
SendKeys "{End}+{Home}"
cmdFGUpdate.Default = True
End Sub
