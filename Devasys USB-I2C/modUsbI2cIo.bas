Attribute VB_Name = "modUsbI2cIo"
Option Explicit

'// UsbI2cIo maximum devices
Public Const USBI2CIO_MAX_DEVICES = 127

'// I2C transaction constants
Public Const USBI2CIO_I2C_HEADER_SIZE = 6
Public Const USBI2CIO_I2C_MAX_DATA = 1088

'// Block of I/O transaction constants
Public Const USBI2CIO_IO_MAX_DATA = 1088

'// I2C transaction types
Public Enum I2C_TRANS_TYPE
    '// standard transaction types
      I2C_TRANS_NOADR = 0       '// read or write with no address cycle
      I2C_TRANS_8ADR = 1        '// read or write with 8 bit address cycle
      I2C_TRANS_16ADR = 2       '// read or write with 16 bit address cycle
      I2C_TRANS_NOADR_NS = 3    '// read or write with no address cycle, stop signaling inhibited
      I2C_TRANS_XICOR = 4       '// read or write with 8 bit instruction cycle, non I2C compliant use of R/W bit
    
      '// customer specific transactions types for accessing rams, only write is mode supported
      I2C_TRANS_8ADR_NONSEQ = 48   '// read or write with 8 bit address cycle, non sequential
      I2C_TRANS_16ADR_NONSEQ = 49  '// read or write with 16 bit address cycle, non sequential
      I2C_TRANS_24ADR_NONSEQ = 50  '// read or write with 24 bit address cycle, non sequential
End Enum

'// BLock of I/O transaction constants
Public Enum IO_MODE
    IO_MODE_SINGLE = 0
    IO_MODE_BLOCK_ABC = 1
    IO_MODE_BLOCK_A = 2
    IO_MODE_BLOCK_B = 3
    IO_MODE_BLOCK_C = 4
    IO_MODE_BLOCK_AB = 5
    IO_MODE_BLOCK_BC = 6
    IO_MODE_BLOCK_AC = 7
End Enum
  
Public Enum USBI2CIO_PROPERTY_INDEX
    WRCOMMAND_RDCOUNT = 0    '// when setting, indicates command, when reading indicates table size
    I2C_CONFIG = 1           '// I2C PROPERTIES (byte 1)
    FAST_XFER_CONFIG = 2     '// FAST TRANSFER PROPERTIES (byte 2)
    IO_CONFIG_GLOBAL = 3     '// I/O PINS GLOBAL PROPERTIES (byte 3)
    IO_CONFIG_PORTA = 4      '// I/O PINS PORT CONFIG PROPERTIES (bytes 4-6)
    IO_CONFIG_PORTB = 5      '
    IO_CONFIG_PORTC = 6      '
    IO_OUTPUT_PORTA = 7      '// I/O PINS PORT OUTPUT PROPERTIES (bytes 7-9)
    IO_OUTPUT_PORTB = 8      '
    IO_OUTPUT_PORTC = 9      '
    DEBUG_CONFIG_GLOBAL = 10 '// DEBUG PROPERTIES (bytes 10-12)
    DEBUG_CONFIG_0 = 11      '
    DEBUG_CONFIG_1 = 12      '
    USER_0 = 13              '// USER PROPERTIES (bytes 13-16)
    USER_1 = 14              '
    USER_2 = 15              '
    USER_3 = 16              '
    IO_CONFIG_PORTD = 17     '// I/O PINS
    IO_OUTPUT_PORTD = 18     '
    IO_OUTPUT_MODE_PORTA = 19    '// OUTPUT MODE (Pusb-Pull or Open-Drain)
    IO_OUTPUT_MODE_PORTB = 20    '
    IO_OUTPUT_MODE_PORTC = 21    '
    IO_OUTPUT_MODE_PORTD = 22    '
    I2C_DEFAULT_CHANNEL = 23 '/ I2C DEFAULT CHANNEL SELECTION AND CLOCK SPEED
    I2C_CHAN0_CLK_LO = 24
    I2C_CHAN0_CLK_HI = 25
    I2C_CHAN1_CLK_LO = 26
    I2C_CHAN1_CLK_HI = 27
    I2C_CHAN2_CLK_LO = 28
    I2C_CHAN2_CLK_HI = 29
    MAX_PROPERTY_BYTES = 30
End Enum

Public Enum USBI2CIO_PROPERTY_COMMAND
    CMD_STORE_TABLE_TO_EEPROM = 0
    CMD_LOAD_TABLE_FROM_EEPROM = 1
    CMD_DISABLE_EEPROM_TABLE = 2
    CMD_RESET_TO_DEFAULTS = 3
End Enum


'// I2C PROPERTIES (byte 1)
'''/#define PROP_I2C_BYTE 0x01
Public Enum I2C_PROP                '''mask
    PROP_I2C_RETRIES_FIELD = &H7 '0000 0111
    PROP_I2C_IGNORE_NAK = &H8 '0000 1000
    PROP_I2C_POLL_EEPROM_ACK = &H10 '0001 0000
    PROP_I2C_AUTO_REDIRECT_A2_REQS = &H20 '0010 0000
    PROP_I2C_RESERVED_FIELD = &HC0 '1100 0000
End Enum
'''''

Public Enum I2C_CLOCK
    PROP_I2C_1MHz = &HF0
    PROP_I2C_400kHz = &HD8
    PROP_I2C_100kHz = &H60
    PROP_I2C_90kHz = &H4E   '''default
End Enum

' Global type definitions for UsbI2cIo API DLL (correspond to values in UsbI2cIo.h)

Public Type WORD          ' provides easy access to high and low bytes of two-byte entity
    lo As Byte
    hi As Byte
End Type

Public Type I2C_TRANS       ' I2C Transaction Structure, used to specify I2C transaction info
    byType As Byte          ' see I2C_TRAN_TYPE enum (above)
    byDevId As Byte         ' bits 7-1 = the I2C device ID, bit 0 is auto set/cleared by call
    wMemAddr As WORD        ' if accessing a device with sub-addressing, sub-address goes here
    wCount As WORD          ' count of bytes in Data array
    Data(1087) As Byte      ' I2C transaction data
End Type

Public Type DEVINFO
    byInstance As Byte      ' instance number of device
    SerialID(8) As Byte     ' 8 bytes Serial ID string of device and a NULL termination
End Type                    ' Note: in Vb, the array size is 1 greater than number specified


'''''

' UsbI2cIo API DLL function declarations

Declare Function DAPI_GetDllVersion Lib "UsbI2cIo.dll" () As WORD

Declare Function DAPI_GetDriverVersion Lib "UsbI2cIo.dll" () As WORD

Declare Function DAPI_GetFirmwareVersion Lib "UsbI2cIo.dll" () As WORD

Declare Function DAPI_GetDeviceCount Lib "UsbI2cIo.dll" ( _
  ByVal lpsDevName As String) _
  As Byte

Declare Function DAPI_GetDeviceInfo Lib "UsbI2cIo.dll" ( _
  ByVal lpsDevName As String, _
  ByRef lpDevInfo As DEVINFO) _
  As Byte

Declare Function DAPI_GetSerialId Lib "UsbI2cIo.dll" ( _
  ByVal hDevInstance As Long, _
  ByVal lpsSerialId As String) _
  As Boolean
  
Declare Function DAPI_DetectDevice Lib "UsbI2cIo.dll" ( _
  ByVal hDevInstance As Long) _
  As Boolean

Declare Function DAPI_OpenDeviceInstance Lib "UsbI2cIo.dll" ( _
  ByVal lpsDevName As String, _
  ByVal byDevInstance As Byte) _
  As Long

Declare Function DAPI_CloseDeviceInstance Lib "UsbI2cIo.dll" ( _
  ByVal hDevInstance As Long) _
  As Boolean

Declare Function DAPI_OpenDeviceBySerialId Lib "UsbI2cIo.dll" ( _
  ByVal lpsDevName As String, _
  ByVal lpsSerialId As String) _
  As Long

Declare Function DAPI_GetIoConfig Lib "UsbI2cIo.dll" ( _
  ByVal hDevInstance As Long, _
  ByRef pulIoPortData As Long) _
  As Byte

Declare Function DAPI_ConfigIoPorts Lib "UsbI2cIo.dll" ( _
  ByVal hDevInstance As Long, _
  ByVal ulIoPortConfig As Long) _
  As Byte

Declare Function DAPI_ReadIoPorts Lib "UsbI2cIo.dll" ( _
  ByVal hDevInstance As Long, _
  ByRef pulIoPortData As Long) _
  As Boolean
  
Declare Function DAPI_WriteIoPorts Lib "UsbI2cIo.dll" ( _
  ByVal hDevInstance As Long, _
  ByVal ulIoPortData As Long, _
  ByVal ulIoPortMask As Long) _
  As Boolean

Declare Function DAPI_ReadI2c Lib "UsbI2cIo.dll" ( _
  ByVal hDevInstance As Long, _
  ByRef TransI2c As I2C_TRANS) _
  As Long

Declare Function DAPI_WriteI2c Lib "UsbI2cIo.dll" ( _
  ByVal hDevInstance As Long, _
  ByRef TransI2c As I2C_TRANS) _
  As Long

Declare Function DAPI_ReadDebugBuffer Lib "UsbI2cIo.dll" ( _
  ByRef DebugBuf As Byte, _
  ByVal hDevInstance As Long, _
  ByVal ulMaxBytes As Long) _
  As Long
  
Declare Function DAPI_SetProperty Lib "UsbI2cIo.dll" ( _
    ByVal hDevInstance As Long, _
    ByVal byPropIndex As Byte, _
    ByVal byPropValue As Byte) _
    As Boolean

Declare Function DAPI_GetProperty Lib "UsbI2cIo.dll" ( _
    ByVal hDevInstance As Long, _
    ByRef byPropValue As Byte, _
    ByVal byPropIndex As Byte) _
    As Boolean

Declare Function DAPI_SetVendorRequest Lib "UsbI2cIo.dll" ( _
    ByVal hDevInstance As Long, _
    ByVal byRequest As Byte, _
    ByVal wValue As Long, _
    ByVal wIndex As Long, _
    ByVal wLength As Long, _
    ByRef pbySetData As Byte) _
    As Boolean

Declare Function DAPI_GetVendorRequest Lib "UsbI2cIo.dll" ( _
    ByVal hDevInstance As Long, _
    ByRef pbySetData As Byte, _
    ByVal byRequest As Byte, _
    ByVal wValue As Long, _
    ByVal wIndex As Long, _
    ByVal wLength As Long) _
    As Boolean

