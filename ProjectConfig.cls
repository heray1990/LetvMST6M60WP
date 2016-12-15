VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProjectConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**********************************************
' Class module for handling config.xml of the
' application.
'**********************************************

Option Explicit

Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long

Private Type udtConfigData
    strComBaud As String
    intComID As Integer
    intProductModel As Integer
    intBacklightType As Integer
    intBoardModel As Integer
    intHardwareVersion As Integer
    int2D3DModel As Integer
    intPanelModel As Integer
    intBurningModeEnable As Integer
End Type

Private mConfigData As udtConfigData
Private mstrConfigFilePath As String

Private Sub Class_Initialize()
    mstrConfigFilePath = App.Path & "\config.xml"
    mConfigData.strComBaud = "115200"
    mConfigData.intComID = 1
    mConfigData.intProductModel = 1
    mConfigData.intBacklightType = 1
    mConfigData.intBoardModel = 1
    mConfigData.intHardwareVersion = 1
    mConfigData.int2D3DModel = 1
    mConfigData.intPanelModel = 1
    mConfigData.intBurningModeEnable = 1
End Sub

Public Sub LoadConfigData()
    Dim xmlDoc As New MSXML2.DOMDocument
    Dim success As Boolean

    If Not CBool(PathFileExists(mstrConfigFilePath)) Then
        MsgBox "Cannot open " & mstrConfigFilePath & " file."
        End
    End If
    
    success = xmlDoc.Load(mstrConfigFilePath)
    
    If success = False Then
        MsgBox xmlDoc.parseError.reason
    Else
        mConfigData.strComBaud = xmlDoc.selectSingleNode("/config/common").selectSingleNode("@baud").Text
        mConfigData.intComID = Val(xmlDoc.selectSingleNode("/config/common").selectSingleNode("@id").Text)
        mConfigData.intProductModel = Val(xmlDoc.selectSingleNode("/config/MonitorProductModel").Text)
        mConfigData.intBacklightType = Val(xmlDoc.selectSingleNode("/config/MonitorBacklightType").Text)
        mConfigData.intBoardModel = Val(xmlDoc.selectSingleNode("/config/MonitorBoardModel").Text)
        mConfigData.intHardwareVersion = Val(xmlDoc.selectSingleNode("/config/MonitorHardwareVersion").Text)
        mConfigData.int2D3DModel = Val(xmlDoc.selectSingleNode("/config/Monitor2D3DModel").Text)
        mConfigData.intPanelModel = Val(xmlDoc.selectSingleNode("/config/MonitorPanelModel").Text)

        If xmlDoc.selectSingleNode("/config/BurningMode").selectSingleNode("@enable").Text = "True" Then
            mConfigData.intBurningModeEnable = 1
        Else
            mConfigData.intBurningModeEnable = 0
        End If
    End If
End Sub

Public Sub SaveConfigData()
    Dim xmlDoc As New MSXML2.DOMDocument
    Dim success As Boolean
    
    success = xmlDoc.Load(mstrConfigFilePath)
    
    If success = False Then
        MsgBox xmlDoc.parseError.reason
    Else
        xmlDoc.selectSingleNode("/config/common").selectSingleNode("@baud").Text = mConfigData.strComBaud
        xmlDoc.selectSingleNode("/config/common").selectSingleNode("@id").Text = CStr(mConfigData.intComID)
        xmlDoc.selectSingleNode("/config/MonitorProductModel").Text = CStr(mConfigData.intProductModel)
        xmlDoc.selectSingleNode("/config/MonitorBacklightType").Text = CStr(mConfigData.intBacklightType)
        xmlDoc.selectSingleNode("/config/MonitorBoardModel").Text = CStr(mConfigData.intBoardModel)
        xmlDoc.selectSingleNode("/config/MonitorHardwareVersion").Text = CStr(mConfigData.intHardwareVersion)
        xmlDoc.selectSingleNode("/config/Monitor2D3DModel").Text = CStr(mConfigData.int2D3DModel)
        xmlDoc.selectSingleNode("/config/MonitorPanelModel").Text = CStr(mConfigData.intPanelModel)
        
        If mConfigData.intBurningModeEnable = 1 Then
            xmlDoc.selectSingleNode("/config/BurningMode").selectSingleNode("@enable").Text = "True"
        Else
            xmlDoc.selectSingleNode("/config/BurningMode").selectSingleNode("@enable").Text = "False"
        End If
        
        xmlDoc.save mstrConfigFilePath
    End If
End Sub

Public Property Get ComBaud() As String
    ComBaud = mConfigData.strComBaud
End Property

Public Property Let ComBaud(strComBaud As String)
    mConfigData.strComBaud = strComBaud
End Property

Public Property Get ComID() As Integer
    ComID = mConfigData.intComID
End Property

Public Property Let ComID(intComID As Integer)
    mConfigData.intComID = intComID
End Property

Public Property Get ProductModel() As Integer
    ProductModel = mConfigData.intProductModel
End Property

Public Property Let ProductModel(intProductModel As Integer)
    mConfigData.intProductModel = intProductModel
End Property

Public Property Get BacklightType() As Integer
    BacklightType = mConfigData.intBacklightType
End Property

Public Property Let BacklightType(intBacklightType As Integer)
    mConfigData.intBacklightType = intBacklightType
End Property

Public Property Get BoardModel() As Integer
    BoardModel = mConfigData.intBoardModel
End Property

Public Property Let BoardModel(intBoardModel As Integer)
    mConfigData.intBoardModel = intBoardModel
End Property

Public Property Get HardwareVersion() As Integer
    HardwareVersion = mConfigData.intHardwareVersion
End Property

Public Property Let HardwareVersion(intHardwareVersion As Integer)
    mConfigData.intHardwareVersion = intHardwareVersion
End Property

Public Property Get Dimension() As Integer
    Dimension = mConfigData.int2D3DModel
End Property

Public Property Let Dimension(int2D3DModel As Integer)
    mConfigData.int2D3DModel = int2D3DModel
End Property

Public Property Get PanelModel() As Integer
    PanelModel = mConfigData.intPanelModel
End Property

Public Property Let PanelModel(intPanelModel As Integer)
    mConfigData.intPanelModel = intPanelModel
End Property

Public Property Get EnableBurningMode() As Integer
    EnableBurningMode = mConfigData.intBurningModeEnable
End Property

Public Property Let EnableBurningMode(intBurningModeEnable As Integer)
    mConfigData.intBurningModeEnable = intBurningModeEnable
End Property
