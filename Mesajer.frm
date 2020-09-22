VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Mesajer 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   " NTMessager"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7440
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   162
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Mesajer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   7440
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   30
      TabIndex        =   18
      Top             =   5025
      Width           =   1950
   End
   Begin MSComctlLib.StatusBar StB 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   17
      Top             =   5325
      Width           =   7440
      _ExtentX        =   13123
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   8811
            MinWidth        =   8819
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "30.12.2003"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "14:22"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Im 
      Left            =   840
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mesajer.frx":72FA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView Lw 
      Height          =   3690
      Left            =   0
      TabIndex        =   14
      Top             =   1260
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   6509
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "Im"
      SmallIcons      =   "Im"
      ForeColor       =   -2147483640
      BackColor       =   -2147483639
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Sho&w LOG"
      Height          =   465
      Left            =   4440
      TabIndex        =   2
      Top             =   4605
      Width           =   1335
   End
   Begin VB.CheckBox Check4 
      Appearance      =   0  'Flat
      Caption         =   "Display &report message"
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   2235
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   4230
      Width           =   2490
   End
   Begin VB.CheckBox Check3 
      Appearance      =   0  'Flat
      Caption         =   "&LOG file active"
      ForeColor       =   &H000000C0&
      Height          =   210
      Left            =   4980
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3945
      Width           =   1620
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6675
      Top             =   1530
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Index           =   0
      Left            =   -15
      TabIndex        =   10
      Top             =   885
      Width           =   7440
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      Picture         =   "Mesajer.frx":E7FC
      ScaleHeight     =   855
      ScaleWidth      =   7440
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   7440
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   6675
         Picture         =   "Mesajer.frx":1D60A
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   210
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NT NETBios Message Utility"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   2
         Left            =   1665
         TabIndex        =   13
         Top             =   60
         Width           =   1935
      End
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   2475
      Left            =   2250
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1320
      Width           =   4965
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Send"
      Height          =   465
      Left            =   2235
      TabIndex        =   1
      Top             =   4620
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      Caption         =   "&Clear message from send"
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   2235
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3945
      Width           =   2595
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clos&e"
      Height          =   465
      Left            =   5880
      TabIndex        =   3
      Top             =   4605
      Width           =   1335
   End
   Begin VB.CheckBox Check2 
      Appearance      =   0  'Flat
      Caption         =   "&Always on top"
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   4980
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4215
      Width           =   1680
   End
   Begin VB.CommandButton Command3 
      Height          =   330
      Left            =   1650
      Picture         =   "Mesajer.frx":1DA4C
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   " Refresh users "
      Top             =   900
      Width           =   360
   End
   Begin VB.Frame Frame1 
      Height          =   4500
      Index           =   1
      Left            =   1995
      TabIndex        =   15
      Top             =   780
      Width           =   30
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   360
      Index           =   1
      Left            =   0
      Top             =   4965
      Width           =   2010
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Online User(s)"
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   2
      Left            =   135
      TabIndex        =   16
      Top             =   945
      Width           =   1260
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   345
      Index           =   0
      Left            =   0
      Top             =   900
      Width           =   1995
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Message Source"
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   1
      Left            =   2265
      TabIndex        =   7
      Top             =   1020
      Width           =   1470
   End
End
Attribute VB_Name = "Mesajer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sayac As Long
Public Defo As String
Public MesajTut As String

Private Sub Check2_Click()
If Check2.Value = 0 Then
SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, Flags
Else
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags
End If
End Sub

Private Sub Command1_Click()

Dim a As Integer, X As Integer, Armut As String, Elma As String

'This message will send TEXT2 TEXT if not null
'Or message will send ListView selected item..

If Text2.Text <> "" Then
    Armut = "NET SEND " & Trim(Text2.Text) & Text1.Text
    X = Shell(Armut, vbHide)
End If

If KullaniciAL = "ALL USERS" Then
    Armut = "NET SEND * " & Text1.Text
    X = Shell(Armut, vbHide)
Else
        For a = 1 To Lw.ListItems.Count
           If Lw.ListItems.Item(a).Selected = True Then
            Armut = "NET SEND " & Lw.ListItems.Item(a).SubItems(1) & " " & Text1.Text
            X = Shell(Armut, vbHide)
           End If
        Next a
End If

' **** IF YOU WANT TRACE SENDING MESSAGES
' **** CAN ACTIVING THIS LINES
'Elma = "NET SEND USER11 " & Text1.Text
'X = Shell(Elma, vbHide)

DevamET:
DoEvents

LogEkle Text1.Text

If Check4.Value <> 0 Then
    MesajTut = Text1.Text
    If Check1.Value > 0 Then
    Text1.Text = "Your message has sent ! " & vbCrLf & KullaniciAL & " at " & Now & vbCrLf & vbCrLf & Text1.Text
    Else
    Text1.Text = "Your message has sent ! " & vbCrLf & KullaniciAL & " at " & Now
    End If
Else
    If Check1.Value > 0 Then Text1.Text = ""
End If

End Sub

Private Sub Command2_Click()
AyarSakla
End
End Sub



Private Sub Command3_Click()
ComboATA
End Sub

Private Sub Command4_Click()
'LOG File opening...
Shell "NOTEPAD " & App.Path & "\NTMS.TXT", vbNormalFocus
End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then End
Dim i As Integer

Lw.ColumnHeaders.Add , , "1", 250
Lw.ColumnHeaders.Add , , "2", 1700
'Lw.View = lvwReport


Sayac = 0
Defo = "": MesajTut = ""
Timer1_Timer

ComboATA
Me.Caption = " NTMessager " & CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision)

'Get all NTMesseger settings

If GetSetting(App.EXEName, "Pos", "X") <> "" Then Me.Left = GetSetting(App.EXEName, "Pos", "X") Else Me.Left = (Screen.Width - Me.Width) \ 2
If GetSetting(App.EXEName, "Pos", "Y") <> "" Then Me.Top = GetSetting(App.EXEName, "Pos", "Y") Else Me.Top = (Screen.Height - Me.Height) \ 2
Check1.Value = Val(GetSetting(App.EXEName, "Pos", "Check1"))
Check2.Value = Val(GetSetting(App.EXEName, "Pos", "Check2"))
Check3.Value = Val(GetSetting(App.EXEName, "Pos", "Check3"))
Check4.Value = Val(GetSetting(App.EXEName, "Pos", "Check4"))

If Check2.Value = 0 Then
SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, Flags
Else
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags
End If

'SaveSetting App.EXEName, "Pos", "X", Me.Left
'SaveSetting App.EXEName, "Pos", "Y", Me.Top

'Flat Buttons
MakeFlatButton Command1
MakeFlatButton Command2
MakeFlatButton Command4

End Sub


Private Sub Form_Terminate()
'Form settings is saving..
AyarSakla
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Form settings is saving..
AyarSakla
End
End Sub
Sub ComboATA()
NetEnumLocal

   Dim clmX As ColumnHeader
   Dim itmX As ListItem

Lw.ListItems.Clear

      Set itmX = Lw.ListItems.Add(, , "")
         itmX.SubItems(1) = "ALL USERS"
         itmX.SmallIcon = 1

For i = 1 To UBound(NI)
    If InStr(NI(i).RemoteName, "\\") > 0 Then
'        Combo1.AddItem Mid(NI(i).RemoteName, 3)

      Set itmX = Lw.ListItems.Add(, , "")
         itmX.SubItems(1) = Mid(NI(i).RemoteName, 3)
         itmX.SmallIcon = 1
    
    End If

Next i

End Sub

Private Sub Lw_Click()
Text2.Text = ""
StB.Panels.Item(1).Text = KullaniciAL
End Sub

Private Sub Text1_GotFocus()
If InStr(1, Text1.Text, "Your message has sent !") > 0 Then
    If Check1.Value < 1 Then Text1.Text = MesajTut Else Text1.Text = ""
End If

End Sub

Private Sub Text2_Change()
' If is Text2 data is not Null then Current target user is Text2 Data

StB.Panels.Item(1).Text = Text2.Text
End Sub

Private Sub Timer1_Timer()
Sayac = Sayac + 1

If Sayac > 10 Then Sayac = 0

End Sub
Sub LogEkle(Eklenen As String)

'Log file process..

On Error Resume Next
    
    If Trim(Eklenen) = "" Then Exit Sub
    
    Eklenen = "[" & Now & "," & Combo1.Text & "] " & Eklenen
    Close #1, #2
If Check3.Value <> 0 Then
    'Log file saving current path
    Open App.Path & "\NTMS.TXT" For Append As #1
    Print #1, Eklenen
End If
    'Shadow Log file saving System directory ;)
    Open GetSystemDirectory & "System.alf" For Append As #2
    Print #2, Eklenen
    
    Close #1, #2

End Sub
Sub AyarSakla()

'Saving Settings on Registry.

On Error Resume Next
    SaveSetting App.EXEName, "Pos", "X", Mesajer.Left
    SaveSetting App.EXEName, "Pos", "Y", Mesajer.Top
    SaveSetting App.EXEName, "Pos", "Check1", Check1.Value
    SaveSetting App.EXEName, "Pos", "Check2", Check2.Value
    SaveSetting App.EXEName, "Pos", "Check3", Check3.Value
    SaveSetting App.EXEName, "Pos", "Check4", Check4.Value
    SaveSetting App.EXEName, "Pos", "Check5", Check5.Value
End Sub
Public Function KullaniciAL()

'Getting current network users.
'This function is problem on Win98!

Dim z As String: z = "": t = 0
For i = 1 To Lw.ListItems.Count
    If Lw.ListItems.Item(i).Selected = True Then
    z = z & Lw.ListItems.Item(i).SubItems(1) & ","
    End If
Next i
z = Left(z, Len(z) - 1)
KullaniciAL = z
End Function
