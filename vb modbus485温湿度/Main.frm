VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ModBus 485 温湿度"
   ClientHeight    =   4515
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   7095
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   7095
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   900
      Left            =   7380
      TabIndex        =   12
      Top             =   180
      Width           =   990
   End
   Begin VB.Timer Timer_GetData 
      Interval        =   100
      Left            =   5805
      Top             =   1890
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   5805
      Top             =   1215
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   9
      BaudRate        =   4800
      InputMode       =   1
   End
   Begin VB.Frame Frame_DeviceList 
      Caption         =   "设备列表"
      Height          =   870
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   6990
      Begin VB.CheckBox Check_DevList 
         Caption         =   "10"
         Height          =   195
         Index           =   9
         Left            =   6300
         TabIndex        =   10
         Top             =   405
         Width           =   600
      End
      Begin VB.CheckBox Check_DevList 
         Caption         =   "9"
         Height          =   195
         Index           =   8
         Left            =   5620
         TabIndex        =   9
         Top             =   405
         Width           =   420
      End
      Begin VB.CheckBox Check_DevList 
         Caption         =   "8"
         Height          =   195
         Index           =   7
         Left            =   4940
         TabIndex        =   8
         Top             =   405
         Width           =   420
      End
      Begin VB.CheckBox Check_DevList 
         Caption         =   "7"
         Height          =   195
         Index           =   6
         Left            =   4260
         TabIndex        =   7
         Top             =   405
         Width           =   420
      End
      Begin VB.CheckBox Check_DevList 
         Caption         =   "6"
         Height          =   195
         Index           =   5
         Left            =   3580
         TabIndex        =   6
         Top             =   405
         Width           =   420
      End
      Begin VB.CheckBox Check_DevList 
         Caption         =   "5"
         Height          =   195
         Index           =   4
         Left            =   2900
         TabIndex        =   5
         Top             =   405
         Width           =   420
      End
      Begin VB.CheckBox Check_DevList 
         Caption         =   "4"
         Height          =   195
         Index           =   3
         Left            =   2220
         TabIndex        =   4
         Top             =   405
         Width           =   420
      End
      Begin VB.CheckBox Check_DevList 
         Caption         =   "3"
         Height          =   195
         Index           =   2
         Left            =   1540
         TabIndex        =   3
         Top             =   405
         Width           =   420
      End
      Begin VB.CheckBox Check_DevList 
         Caption         =   "2"
         Height          =   195
         Index           =   1
         Left            =   860
         TabIndex        =   2
         Top             =   405
         Width           =   420
      End
      Begin VB.CheckBox Check_DevList 
         Caption         =   "1"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   1
         Top             =   405
         Width           =   420
      End
   End
   Begin VB.TextBox Text1 
      Height          =   3480
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   990
      Width           =   6990
   End
   Begin VB.Menu Ref 
      Caption         =   "刷新"
   End
   Begin VB.Menu Uart 
      Caption         =   "串口"
      Begin VB.Menu UartX1 
         Caption         =   "COM1"
         Index           =   0
      End
   End
   Begin VB.Menu Baud 
      Caption         =   "波特率"
      Begin VB.Menu bps 
         Caption         =   "4800bps"
         Index           =   0
      End
      Begin VB.Menu bps 
         Caption         =   "9600bps"
         Index           =   1
      End
      Begin VB.Menu bps 
         Caption         =   "19200bps"
         Index           =   2
      End
      Begin VB.Menu bps 
         Caption         =   "38400bps"
         Index           =   3
      End
   End
   Begin VB.Menu status 
      Caption         =   "status"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim COMBaud As String
Dim COMPort  As Integer

Dim DeviceNumber As Byte
Dim RecordStr As String
Dim Pos As Integer
'

Private Sub bps_Click(Index As Integer)
    Select Case Index
        Case 0
            COMBaud = "4800"
        Case 1
            COMBaud = "9600"
        Case 2
            COMBaud = "19200"
        Case 3
            COMBaud = "38400"
    End Select
    
    Call SetMSCOMM
    
End Sub

Function SetMSCOMM()
    On Error Resume Next
    
    
    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False
    End If
   
    MSComm1.CommPort = COMPort
    MSComm1.Settings = COMBaud & ",n,8,1"
    MSComm1.PortOpen = True
    If Err Then
        MsgBox Err.Number & Err.Description, vbInformation + vbOKOnly
        Err.Clear
        status.Caption = ""
        status.Enabled = False
        status.Visible = False
    End If
    status.Caption = "串口已打开:" & COMBaud & "bps@" & "COM" & COMPort
    status.Visible = True
    status.Enabled = False
    
    Debug.Print MSComm1.Settings
End Function
 
  
Private Function InitDB() As Boolean
    
    
    InitDB = DB_CreateDataBase(App.Path & "\db\temprh.mdb")
    If InitDB = False Then
        Debug.Print "Init DB fail"
        Exit Function
    End If
    
    Dim InitDBTable As New Table
    InitDBTable.Name = "Senser"
    InitDBTable.Columns.Append "RoomNo", adVarWChar, 20
    InitDBTable.Columns.Append "DeviceNo", adVarWChar, 20
    InitDBTable.Columns.Append "RH", adVarWChar, 20
    InitDBTable.Columns.Append "Temp", adVarWChar, 20
    InitDBTable.Columns.Append "Time", adVarWChar, 20
    InitDBTable.Columns.Append "Date", adVarWChar, 20
    
    InitDB = DB_Create_Table(InitDBTable)
    
    If InitDB = False Then
        Debug.Print "Init Table Fail"
        Exit Function
    End If
    Debug.Print "all success"
End Function



Private Sub Form_Load()
    Dim i As Integer, j As Integer
    i = 0
    j = 0
    On Error Resume Next
    COMBaud = "4800"
    MSComm1.Settings = COMBaud & ",n,8,1"
    For j = 1 To 256
        MSComm1.CommPort = j
        MSComm1.PortOpen = True
        If Err.Number = 0 Then
            If i <> 0 Then
                Load UartX1(i)
            End If
            UartX1(i).Caption = "COM" & j
            UartX1(i).Visible = True
            i = i + 1
        End If
        
        If MSComm1.PortOpen = True Then
            MSComm1.PortOpen = False
        End If
        Err.Clear
    Next
    If i = 0 Then '无可用串口
        Uart.Visible = False
        Baud.Visible = False
        status.Caption = 0
        status.Visible = False
        '
        Exit Sub
    End If
    Uart.Visible = True
    Baud.Visible = True
    '
    COMPort = CInt(Replace$(UCase(UartX1(0).Caption), "COM", ""))
    Call SetMSCOMM
    Call InitDB
    RecordStr = vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf
End Sub


Private Sub MSComm1_OnComm()
    Dim Crc1 As Byte
    Dim Crc2 As Byte
    Dim Temper As Double
    Dim Rh As Double
    Dim Records(5) As String
        
    Records(0) = 1
    
    If MSComm1.CommEvent = comEvReceive Then
        Dim intA()  As Byte '十六进制转换中间变量
        
        If MSComm1.InBufferCount <> MSComm1.RThreshold Then
            Exit Sub
        End If
         
        intA() = MSComm1.Input
 
        
        Dim j As Integer
        Debug.Print "____",
        For j = 0 To UBound(intA())
            Debug.Print Hex(intA(j)),
        Next
        Debug.Print
        
        If UBound(intA) = 8 Then
            Call CRC16(intA(), 0, 6, Crc1, Crc2)
            If Crc1 = intA(7) And Crc2 = intA(8) Then
                Debug.Print "数据校验正确!"
                Temper = (intA(5) And &H7F) * 256 + intA(6)
                Temper = Temper / 10
                If (intA(5) And &H7F) <> intA(5) Then '判断最高位,如果是1,则进行换算,否则不需要换算
                    Temper = Temper * -1
                End If
                
                Rh = intA(3) * 256 + intA(4)
                Rh = Rh / 10
                Records(1) = intA(0)
                Records(2) = Rh
                Records(3) = Temper
                Records(4) = Time
                Records(5) = Date
                RecordStr = "Dev:" & intA(0) & vbTab & Format(Temper, "0.0") & "℃" & vbTab & Format(Rh, "0.0") & "%RH" & vbTab & vbTab & Format(Time, "HH:MM:SS") & " " & Format(Date, "YY-MM-DD") & vbCrLf & RecordStr
                Pos = InStrRev(RecordStr, vbCrLf)
                'Pos = InStrRev(RecordStr, vbCrLf, Pos - 1)
                RecordStr = Left(RecordStr, Pos)
                
                Text1 = RecordStr
                Call DB_InsertRecord("Senser", Records())
            End If
        End If
    End If

End Sub

Private Sub Ref_Click()
    Call Form_Load
End Sub


Private Sub Timer_GetData_Timer()
    Dim ModBus485Data(7) As Byte
    Dim i As Byte
    Dim j As Byte
    Dim CRC_ret As String
    

    ModBus485Data(1) = 3
    ModBus485Data(2) = 0
    ModBus485Data(3) = 0
    ModBus485Data(4) = 0
    ModBus485Data(5) = 2
    
    Timer_GetData.Enabled = False
    If DeviceNumber > 9 Then
        DeviceNumber = 0
    End If
    
    For i = DeviceNumber To 9
        If MSComm1.PortOpen = True Then
            DeviceNumber = i + 1
            If Check_DevList(i).Value = 0 Then
                'Debug.Print "未选择设备"; DeviceNumber
            Else
                
                ModBus485Data(0) = DeviceNumber
                Call CRC16(ModBus485Data, 0, 5, ModBus485Data(6), ModBus485Data(7))
                'MSComm1.RThreshold = 9
                MSComm1.Output = ModBus485Data()
                
                Debug.Print "send"; DeviceNumber,
                
                For j = 0 To 7
                    Debug.Print Hex(ModBus485Data(j)),
                Next
                Debug.Print
                Exit For
             End If
        End If
    Next
    Timer_GetData.Enabled = True
End Sub

Private Sub UartX1_Click(Index As Integer)
    COMPort = CInt(Replace$(UCase(UartX1(Index).Caption), "COM", ""))
    Debug.Print COMPort
    Call SetMSCOMM
End Sub
