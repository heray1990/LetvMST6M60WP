VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "Letv Max65 属性烧录工具"
   ClientHeight    =   5175
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   3990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   3990
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CommandUnlock 
      Caption         =   "解锁"
      Height          =   375
      Left            =   1440
      TabIndex        =   18
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton CommandWrite 
      Caption         =   "烧录"
      Height          =   375
      Left            =   2760
      TabIndex        =   17
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton CommandLock 
      Caption         =   "锁定"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Frame FrameBurningMode 
      Caption         =   "老化模式"
      Height          =   735
      Left            =   120
      TabIndex        =   13
      Top             =   3360
      Width           =   3735
      Begin VB.ComboBox ComboBurningMode 
         Height          =   315
         ItemData        =   "Form1.frx":0000
         Left            =   1560
         List            =   "Form1.frx":0002
         TabIndex        =   15
         Text            =   "Color Bar"
         Top             =   300
         Width           =   2000
      End
      Begin VB.CheckBox CheckBurningMode 
         Caption         =   "老化模式"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame FrameProperty 
      Caption         =   "产品属性"
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.ComboBox ComboPanel 
         Height          =   315
         ItemData        =   "Form1.frx":0004
         Left            =   1560
         List            =   "Form1.frx":0006
         TabIndex        =   12
         Text            =   "Panel Model"
         Top             =   2620
         Width           =   2000
      End
      Begin VB.ComboBox Combo2D3D 
         Height          =   315
         ItemData        =   "Form1.frx":0008
         Left            =   1560
         List            =   "Form1.frx":000A
         TabIndex        =   11
         Text            =   "2D/3D"
         Top             =   2160
         Width           =   2000
      End
      Begin VB.ComboBox ComboHwVer 
         Height          =   315
         ItemData        =   "Form1.frx":000C
         Left            =   1560
         List            =   "Form1.frx":000E
         TabIndex        =   10
         Text            =   "Hardware Version"
         Top             =   1680
         Width           =   2000
      End
      Begin VB.ComboBox ComboBoard 
         Height          =   315
         ItemData        =   "Form1.frx":0010
         Left            =   1560
         List            =   "Form1.frx":0012
         TabIndex        =   9
         Text            =   "Board Model"
         Top             =   1240
         Width           =   2000
      End
      Begin VB.ComboBox ComboBacklight 
         Height          =   315
         ItemData        =   "Form1.frx":0014
         Left            =   1560
         List            =   "Form1.frx":0016
         TabIndex        =   8
         Text            =   "Backlight Type"
         Top             =   780
         Width           =   2000
      End
      Begin VB.ComboBox ComboProduct 
         Height          =   315
         ItemData        =   "Form1.frx":0018
         Left            =   1560
         List            =   "Form1.frx":001A
         TabIndex        =   7
         Text            =   "Product Model"
         Top             =   320
         Width           =   2000
      End
      Begin VB.Label Label6 
         Caption         =   "屏型号："
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   2680
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "2D/3D 类型："
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   2220
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "硬件版本："
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1760
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "制版阶段："
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1300
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "背光类型："
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "产品类型："
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   380
         Width           =   1095
      End
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   1680
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Version "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Left            =   3105
      TabIndex        =   19
      Top             =   4800
      Width           =   720
   End
   Begin VB.Menu MenuItemSetting 
      Caption         =   "Setting"
      Begin VB.Menu MenuItemComSetting 
         Caption         =   "COM Setting"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrProductModel
Dim arrBacklightType
Dim arrBoradModel
Dim arrHwVer
Dim arrDimension
Dim arrPanelModel
Dim arrBurningMode

Private Sub SubInit()
    Dim clsConfigData As ProjectConfig

    Set clsConfigData = New ProjectConfig
    clsConfigData.LoadConfigData
    
    glngTVComBaud = clsConfigData.ComBaud
    gintTVComID = clsConfigData.ComID
    SubInitComPort

    gintProductModel = clsConfigData.ProductModel
    gintBacklightType = clsConfigData.BacklightType
    gintBoardModel = clsConfigData.BoardModel
    gintHardwareVersion = clsConfigData.HardwareVersion
    gint2D3DModel = clsConfigData.Dimension
    gintPanelModel = clsConfigData.PanelModel
    gintBurningMode = clsConfigData.BurningMode
    gintBurningModeEnable = clsConfigData.EnableBurningMode

    Set clsConfigData = Nothing
    
    If gintBacklightType < 1 Then
        gintBacklightType = 1
    End If
    If gintBoardModel < 1 Then
        gintBoardModel = 1
    End If
    If gintHardwareVersion < 1 Then
        gintHardwareVersion = 1
    End If
    If gint2D3DModel < 1 Then
        gint2D3DModel = 1
    End If
    If gintPanelModel < 1 Then
        gintPanelModel = 1
    End If
    ComboProduct.Text = arrProductModel(gintProductModel)
    ComboBacklight.Text = arrBacklightType(gintBacklightType - 1)
    ComboBoard.Text = arrBoradModel(gintBoardModel - 1)
    ComboHwVer.Text = arrHwVer(gintHardwareVersion - 1)
    Combo2D3D.Text = arrDimension(gint2D3DModel - 1)
    ComboPanel.Text = arrPanelModel(gintPanelModel - 1)
    ComboBurningMode.Text = arrBurningMode(gintBurningMode)
    CheckBurningMode.Value = gintBurningModeEnable
End Sub

Public Sub SubInitComPort()
    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False
    End If

    MSComm1.CommPort = gintTVComID
    MSComm1.Settings = glngTVComBaud & ",N,8,1"
    MSComm1.InputLen = 0

    MSComm1.InBufferCount = 0
    MSComm1.OutBufferCount = 0
    MSComm1.InputMode = comInputModeBinary

    MSComm1.NullDiscard = False
    MSComm1.DTREnable = False
    MSComm1.EOFEnable = False
    MSComm1.RTSEnable = False
    MSComm1.SThreshold = 1
    MSComm1.RThreshold = 1
    MSComm1.InBufferSize = 1024
    MSComm1.OutBufferSize = 512
End Sub

Private Sub CommandLock_Click()
    ComboProduct.Enabled = False
    ComboBacklight.Enabled = False
    ComboBoard.Enabled = False
    ComboHwVer.Enabled = False
    Combo2D3D.Enabled = False
    ComboPanel.Enabled = False
    CheckBurningMode.Enabled = False
    ComboBurningMode.Enabled = False
    CommandLock.Enabled = False
    CommandUnlock.Enabled = True
    CommandWrite.Enabled = True
End Sub

Private Sub CommandUnlock_Click()
    ComboProduct.Enabled = True
    ComboBacklight.Enabled = True
    ComboBoard.Enabled = True
    ComboHwVer.Enabled = True
    Combo2D3D.Enabled = True
    ComboPanel.Enabled = True
    CheckBurningMode.Enabled = True
    ComboBurningMode.Enabled = True
    CommandLock.Enabled = True
    CommandUnlock.Enabled = False
    CommandWrite.Enabled = False
End Sub

Private Sub CommandWrite_Click()
On Error GoTo ErrExit
    Dim i As Integer
    Dim clsSaveConfigData As ProjectConfig
    
    Set clsSaveConfigData = New ProjectConfig

    clsSaveConfigData.ComBaud = CStr(glngTVComBaud)
    clsSaveConfigData.ComID = gintTVComID
    For i = 0 To 10
        If Trim(ComboProduct.Text) = Trim(arrProductModel(i)) Then
            clsSaveConfigData.ProductModel = i
            gintProductModel = i
            Exit For
        End If
    Next i
    For i = 0 To 1
        If Trim(ComboBacklight.Text) = Trim(arrBacklightType(i)) Then
            clsSaveConfigData.BacklightType = i + 1
            gintBacklightType = i + 1
            Exit For
        End If
    Next i
    For i = 0 To 7
        If Trim(ComboBoard.Text) = Trim(arrBoradModel(i)) Then
            clsSaveConfigData.BoardModel = i + 1
            gintBoardModel = i + 1
            Exit For
        End If
    Next i
    For i = 0 To 4
        If Trim(ComboHwVer.Text) = Trim(arrHwVer(i)) Then
            clsSaveConfigData.HardwareVersion = i + 1
            gintHardwareVersion = i + 1
            Exit For
        End If
    Next i
    For i = 0 To 1
        If Trim(Combo2D3D.Text) = Trim(arrDimension(i)) Then
            clsSaveConfigData.Dimension = i + 1
            gint2D3DModel = i + 1
            Exit For
        End If
    Next i
    For i = 0 To 7
        If Trim(ComboPanel.Text) = Trim(arrPanelModel(i)) Then
            clsSaveConfigData.PanelModel = i + 1
            gintPanelModel = i + 1
            Exit For
        End If
    Next i
    For i = 0 To 2
        If Trim(ComboBurningMode.Text) = Trim(arrBurningMode(i)) Then
            clsSaveConfigData.BurningMode = i
            gintBurningMode = i
            Exit For
        End If
    Next i
    clsSaveConfigData.EnableBurningMode = CheckBurningMode.Value
    gintBurningModeEnable = CheckBurningMode.Value

    clsSaveConfigData.SaveConfigData
    Set clsSaveConfigData = Nothing

    If MSComm1.PortOpen = False Then
        MSComm1.PortOpen = True
    End If
    SetProperty 1, gintProductModel
    SetProperty 2, gintBacklightType
    SetProperty 3, gintBoardModel
    SetProperty 4, gintHardwareVersion
    SetProperty 5, gint2D3DModel
    SetProperty 6, gintPanelModel
    If gintBurningModeEnable = 1 Then
        BurningMode gintBurningMode
    End If
    DelayMS 2000
    Exit Sub

ErrExit:
    MsgBox Err.Description, vbCritical, Err.Source
End Sub

Private Sub Form_Load()
    Dim i As Integer

    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    CommandLock.Enabled = True
    CommandUnlock.Enabled = False
    CommandWrite.Enabled = False

    arrProductModel = Array("UNKNOWN", "Max4_70", "Max4_65C", _
                            "Max4_55B", "Max4_65B", "Max4_75B", _
                            "Max4_70S", "Max4_75S", "Max5_55_938", _
                            "Max4_X70", "Max5_65_938")
    arrBacklightType = Array("PWM", "Local Dimming")
    arrBoradModel = Array("EVT", "EVT2", "EVT3", _
                        "DVT", "DVT2", "DVT3", _
                        "PVT", "MP")
    arrHwVer = Array("H1000", "H2000", "H3000", "H5000", "H6000")
    arrDimension = Array("2D", "3D")
    arrPanelModel = Array("X4_70_2D", "X4_70_3D", "X3_55_120", _
                            "X3_55_60", "X4_65_Curve", "X4_65_Blade", _
                            "X4_70S", "X4_75S")
    arrBurningMode = Array("White Pattern", "Color Bar", "Color Square")

    For i = 0 To 10
        ComboProduct.AddItem arrProductModel(i)
    Next i
    For i = 0 To 1
        ComboBacklight.AddItem arrBacklightType(i)
    Next i
    For i = 0 To 7
        ComboBoard.AddItem arrBoradModel(i)
    Next i
    For i = 0 To 4
        ComboHwVer.AddItem arrHwVer(i)
    Next i
    For i = 0 To 1
        Combo2D3D.AddItem arrDimension(i)
    Next i
    For i = 0 To 7
        ComboPanel.AddItem arrPanelModel(i)
    Next i
    For i = 0 To 2
        ComboBurningMode.AddItem arrBurningMode(i)
    Next i
    
    SubInit
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrExit
    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False
    End If

    End
Exit Sub

ErrExit:
    MsgBox Err.Description, vbCritical, Err.Source
End Sub

Private Sub MenuItemComSetting_Click()
    FrmComPort.Show
End Sub

