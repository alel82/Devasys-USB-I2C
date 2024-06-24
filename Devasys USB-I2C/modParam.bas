Attribute VB_Name = "modParam"
Option Explicit

Global ProcessStart As Boolean

Global Const INVALID_HANDLE_VALUE = -1
Global sDevSymName As String      ' symbolic name of USB device, example: "UsbI2cIo"
Global byDevInstance As Byte     ' currently selected device instance number
Global hDevInstance As Long       ' handle to the currently selected device instance

Global Const vbGray = &HE0E0E0

Global Const PLUS = 1
Global Const MINUS = 2

Global SlaveNo As Integer
Global BusSelected As Integer
Global Const I2C0 = 0
Global Const I2C1 = 1
Global Const NoBus = 2
Global Const SrqLow = 1
Global Const SrqHigh = 2
Global Const IICError = -1
Global Const OK = 0
Global Const NG = 1

Global lpApplicationName As String
Global lpFileName As String
Global lpSecName As String
Global lpName As String
Global lpDefault As String
Global lpReturnedString As String
Global lpKeyValue As String
Global lpKeyName As String
Global nSize As Integer

Global Chassis As String
Global EEPROM_Filename(0 To 1) As String
Global VerifyPath(0 To 1) As String
Global VerifyData(0 To 1, 0 To 65535) As String

Global EEPROMEnable(0 To 1) As Boolean

Global Const CompareEEPROM1 = 1
Global Const CompareEEPROM2 = 2
Global Const CompareFile = 3

Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Integer

Global CompareFilePath As String
Global CompareFileName As String
Global CompareTable1() As String
Global CompareTable2() As String

Global ProcessItem() As String
Global ProcessCount As Integer

Sub delay_ms(ByVal del As Integer)
Dim ST, ET
ST = Timer
Do
    ET = Timer
    If ET - ST < 0 Then ST = ST - 86400
Loop Until (ET - ST) > del / 1000
End Sub

