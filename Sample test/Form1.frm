VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9165
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12495
   LinkTopic       =   "Form1"
   ScaleHeight     =   9165
   ScaleWidth      =   12495
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   1800
      List            =   "Form1.frx":0002
      TabIndex        =   1
      Text            =   "160x128"
      Top             =   7920
      Width           =   1095
   End
   Begin VB.Timer Timer2 
      Interval        =   5000
      Left            =   480
      Top             =   2280
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   495
      Left            =   5760
      TabIndex        =   0
      Top             =   7920
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   120
      Left            =   360
      Top             =   1440
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   360
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   1
      BaudRate        =   115200
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   7200
      Left            =   1440
      Top             =   360
      Width           =   9600
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SYNC As String
Dim Initial As String
Dim Set_Package_Size As String
Dim Snapshot As String
Dim Get_Picture As String
Dim ACK As String
Dim rxlen As Integer
Dim count1 As Integer
Dim num_of_package As Integer
Dim rxready As Boolean
Dim Rx As String
Dim Tx As String
Dim Data(100) As String
Dim Remainder As Byte
Dim image_size As Integer
Dim Resolution As Byte

Public Sub Send_Sync()
    For try = 1 To 60
        Rx = ""
        MSComm1.Output = SYNC
        Timer1.Enabled = True
        Do
            DoEvents
        Loop Until Timer1.Enabled = False
            If Len(Rx) = 12 Then
                MSComm1.Output = ACK
                GoTo sync_ok
            End If
    Next try
sync_ok:
End Sub

Public Sub Send()
'    For try = 1 To 60
        Rx = ""
        MSComm1.Output = Tx
        Timer1.Enabled = True
        Do
            DoEvents
        Loop Until Timer1.Enabled = False
'            If Len(Rx) = rxlen Then
'                MSComm1.Output = ACK
'                GoTo send_ok
'            End If
'    Next try
send_ok:
End Sub

Private Sub Combo1_Click()
    If Combo1.ListIndex = 0 Then
        Resolution = 3
    End If
    If Combo1.ListIndex = 1 Then
        Resolution = 5
    End If
    If Combo1.ListIndex = 2 Then
        Resolution = 7
    End If
End Sub

Private Sub Command1_Click()
On Error GoTo Skip
Shot:
    Send_Sync
'    rxlen = 12
'    Send
'    Tx = ACK
'    rxlen = 0
'    Send
    Initial = Chr(&HAA) + Chr(&H1) + Chr(&H0) + Chr(&H7) + Chr(&H0) + Chr(Resolution)
    Tx = Initial
    rxlen = 6
    Send
    Tx = Set_Package_Size
    rxlen = 6
    Send
    Tx = Snapshot
    Send
    rxlen = 6
retry:
    Tx = Get_Picture
    rxlen = 12
    Send
    If Len(Rx) = 0 Then
        GoTo retry
    End If
    image_size = (Asc(Mid$(Rx, 11, 1))) * 256 + Asc(Mid$(Rx, 10, 1))
    num_of_package = 0
back:
    If image_size > 506 Then
        num_of_package = num_of_package + 1
        image_size = image_size - 506
        GoTo back
    End If
    If image_size > 0 Then
        num_of_package = num_of_package + 1
    End If
'    count1 = num_of_package Mod 506
    For count1 = 0 To (num_of_package - 1)
        Tx = Chr(&HAA) + Chr(&HE) + Chr(&H0) + Chr(&H0) + Chr(count1) + Chr(&H0)
        rxlen = 512
        Send
        Data(count1) = Rx
    Next count1
    Tx = Chr(&HAA) + Chr(&HE) + Chr(&H0) + Chr(&H0) + Chr(&HF0) + Chr(&HF0)
    Send
    
    If Dir("C:\test.jpg") <> "" Then
        Kill "C:\test.jpg"
    End If
    
    Open "C:\test.jpg" For Binary As #1
    For count1 = 0 To (num_of_package - 2)
        Put #1, , Mid$(Data(count1), 5, 506)
    Next count1
    
    For count1 = 0 To 512
        Put #1, , Mid$(Data(num_of_package - 1), 5 + count1, 1)
        If Mid$(Data(num_of_package - 1), 5 + count1, 1) = Chr(&HD9) And Mid$(Data(num_of_package - 1), 4 + count1, 1) = Chr(&HFF) Then
            GoTo EOF
        End If
    Next count1
EOF:
    Close #1
    On Error GoTo Skip
    If Resolution = 3 Then
        Image1.Left = 5000
        Image1.Top = 3000
    End If
    If Resolution = 5 Then
        Image1.Left = 3800
        Image1.Top = 2000
    End If
    If Resolution = 7 Then
        Image1.Left = 1440
        Image1.Top = 360
    End If
    Image1.Picture = LoadPicture("C:\test.jpg")
Skip:
    GoTo Shot
 '   End If
    
End Sub

Private Sub Command2_Click()
    MSComm1.Output = Chr(&HAA)
End Sub

Private Sub Form_Load()
    SYNC = Chr(&HAA) + Chr(&HD) + Chr(&H0) + Chr(&H0) + Chr(&H0) + Chr(&H0)
    Set_Package_Size = Chr(&HAA) + Chr(&H6) + Chr(&H8) + Chr(&H0) + Chr(&H2) + Chr(&H0)
    Snapshot = Chr(&HAA) + Chr(&H5) + Chr(&H0) + Chr(&H0) + Chr(&H0) + Chr(&H0)
    Get_Picture = Chr(&HAA) + Chr(&H4) + Chr(&H1) + Chr(&H0) + Chr(&H0) + Chr(&H0)
    ACK = Chr(&HAA) + Chr(&HE) + Chr(&HD) + Chr(&H0) + Chr(&H0) + Chr(&H0)
    Resolution = 3
    
    Combo1.AddItem "160x128"
    Combo1.AddItem "320x240"
    Combo1.AddItem "640x480"
    
    MSComm1.PortOpen = True
End Sub

Private Sub MSComm1_OnComm()
    Rx = Rx + MSComm1.Input
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
End Sub

