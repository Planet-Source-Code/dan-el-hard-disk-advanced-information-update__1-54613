Attribute VB_Name = "mod_SmartDrives"
Option Explicit
'=========================================================='
'Thanks to: CodeGuru Forums Xtreme Visual Basic Talk       '
'Web      : http://www.visualbasicforum.com/               '
'Web2     : http://www.codeguru.com/forum/                 '
'Date     : 25-06-2004                                     '
'Name     : mod_SmartDrives.bas                            '
'=========================================================='
'Daniel PC (Daniel Carrasco Olguin)                        '
'Santiago de Chile                                         '
'=========================================================='
Public Const MAX_IDE_DRIVES = 4
Public Const READ_ATTRIBUTE_BUFFER_SIZE = 512
Public Const IDENTIFY_BUFFER_SIZE = 512
Public Const READ_THRESHOLD_BUFFER_SIZE = 512
Public Const OUTPUT_DATA_SIZE = IDENTIFY_BUFFER_SIZE + 16

Public Const DFP_GET_VERSION = &H74080
Public Const DFP_SEND_DRIVE_COMMAND = &H7C084
Public Const DFP_RECEIVE_DRIVE_DATA = &H7C088

Public Type GETVERSIONOUTPARAMS
       bVersion       As Byte
       bRevision      As Byte
       bReserved      As Byte
       bIDEDeviceMap  As Byte
       fCapabilities  As Long
       dwReserved(3)  As Long
End Type

Public Const CAP_IDE_ID_FUNCTION = 1
Public Const CAP_IDE_ATAPI_ID = 2
Public Const CAP_IDE_EXECUTE_SMART_FUNCTION = 4

Public Type IDEREGS
   bFeaturesReg     As Byte
   bSectorCountReg  As Byte
   bSectorNumberReg As Byte
   bCylLowReg       As Byte
   bCylHighReg      As Byte
   bDriveHeadReg    As Byte
   bCommandReg      As Byte
   bReserved        As Byte
End Type

Public Type SENDCMDINPARAMS
   cBufferSize     As Long
   irDriveRegs     As IDEREGS
   bDriveNumber    As Byte
   bReserved(2)    As Byte
   dwReserved(3)   As Long
   bBuffer()      As Byte
End Type

Public Const IDE_ATAPI_ID = &HA1
Public Const IDE_ID_FUNCTION = &HEC
Public Const IDE_EXECUTE_SMART_FUNCTION = &HB0
                                              
Public Const SMART_CYL_LOW = &H4F
Public Const SMART_CYL_HI = &HC2


Public Type DRIVERSTATUS
   bDriverError  As Byte
   bIDEStatus    As Byte
                                  
   bReserved(1)  As Byte
   dwReserved(1) As Long
 End Type

Public Enum DRIVER_ERRORS
       SMART_NO_ERROR = 0
       SMART_IDE_ERROR = 1
       SMART_INVALID_FLAG = 2
       SMART_INVALID_COMMAND = 3
       SMART_INVALID_BUFFER = 4
       SMART_INVALID_DRIVE = 5
       SMART_INVALID_IOCTL = 6
       SMART_ERROR_NO_MEM = 7
       SMART_INVALID_REGISTER = 8
       SMART_NOT_SUPPORTED = 9
       SMART_NO_IDE_DEVICE = 10
End Enum

Public Type IDSECTOR
   wGenConfig                 As Integer
   wNumCyls                   As Integer
   wReserved                  As Integer
   wNumHeads                  As Integer
   wBytesPerTrack             As Integer
   wBytesPerSector            As Integer
   wSectorsPerTrack           As Integer
   wVendorUnique(2)           As Integer
   sSerialNumber(19)          As Byte
   wBufferType                As Integer
   wBufferSize                As Integer
   wECCSize                   As Integer
   sFirmwareRev(7)            As Byte
   sModelNumber(39)           As Byte
   wMoreVendorUnique          As Integer
   wDoubleWordIO              As Integer
   wCapabilities              As Integer
   wReserved1                 As Integer
   wPIOTiming                 As Integer
   wDMATiming                 As Integer
   wBS                        As Integer
   wNumCurrentCyls            As Integer
   wNumCurrentHeads           As Integer
   wNumCurrentSectorsPerTrack As Integer
   ulCurrentSectorCapacity    As Long
   wMultSectorStuff           As Integer
   ulTotalAddressableSectors  As Long
   wSingleWordDMA             As Integer
   wMultiWordDMA              As Integer
   bReserved(127)             As Byte
End Type

Public Type SENDCMDOUTPARAMS
  cBufferSize   As Long
  DRIVERSTATUS  As DRIVERSTATUS
  bBuffer()    As Byte
End Type

Public Const SMART_READ_ATTRIBUTE_VALUES = &HD0
Public Const SMART_READ_ATTRIBUTE_THRESHOLDS = &HD1
Public Const SMART_ENABLE_DISABLE_ATTRIBUTE_AUTOSAVE = &HD2
Public Const SMART_SAVE_ATTRIBUTE_VALUES = &HD3
Public Const SMART_EXECUTE_OFFLINE_IMMEDIATE = &HD4
Public Const SMART_ENABLE_SMART_OPERATIONS = &HD8
Public Const SMART_DISABLE_SMART_OPERATIONS = &HD9
Public Const SMART_RETURN_SMART_STATUS = &HDA

Public Const NUM_ATTRIBUTE_STRUCTS = 30

Public Type DRIVEATTRIBUTE
       bAttrID As Byte
       wStatusFlags As Integer
       bAttrValue As Byte
       bWorstValue As Byte
       bRawValue(5) As Byte
       bReserved As Byte
End Type

Public Enum STATUS_FLAGS
       PRE_FAILURE_WARRANTY = &H1
       ON_LINE_COLLECTION = &H2
       PERFORMANCE_ATTRIBUTE = &H4
       ERROR_RATE_ATTRIBUTE = &H8
       EVENT_COUNT_ATTRIBUTE = &H10
       SELF_PRESERVING_ATTRIBUTE = &H20
End Enum

Public Type ATTRTHRESHOLD
       bAttrID As Byte
       bWarrantyThreshold As Byte
       bReserved(9) As Byte
End Type

Public Enum ATTRIBUTE_ID
       ATTR_INVALID = 0
       ATTR_READ_ERROR_RATE = 1
       ATTR_THROUGHPUT_PERF = 2
       ATTR_SPIN_UP_TIME = 3
       ATTR_START_STOP_COUNT = 4
       ATTR_REALLOC_SECTOR_COUNT = 5
       ATTR_READ_CHANNEL_MARGIN = 6
       ATTR_SEEK_ERROR_RATE = 7
       ATTR_SEEK_TIME_PERF = 8
       ATTR_POWER_ON_HRS_COUNT = 9
       ATTR_SPIN_RETRY_COUNT = 10
       ATTR_CALIBRATION_RETRY_COUNT = 11
       ATTR_POWER_CYCLE_COUNT = 12
       ATTR_SOFT_READ_ERROR_RATE = 13
       ATTR_G_SENSE_ERROR_RATE = 191
       ATTR_POWER_OFF_RETRACT_CYCLE = 192
       ATTR_LOAD_UNLOAD_CYCLE_COUNT = 193
       ATTR_TEMPERATURE = 194
       ATTR_REALLOCATION_EVENTS_COUNT = 196
       ATTR_CURRENT_PENDING_SECTOR_COUNT = 197
       ATTR_UNCORRECTABLE_SECTOR_COUNT = 198
       ATTR_ULTRADMA_CRC_ERROR_RATE = 199
       ATTR_WRITE_ERROR_RATE = 200
       ATTR_DISK_SHIFT = 220
       ATTR_G_SENSE_ERROR_RATEII = 221
       ATTR_LOADED_HOURS = 222
       ATTR_LOAD_UNLOAD_RETRY_COUNT = 223
       ATTR_LOAD_FRICTION = 224
       ATTR_LOAD_UNLOAD_CYCLE_COUNTII = 225
       ATTR_LOAD_IN_TIME = 226
       ATTR_TORQUE_AMPLIFICATION_COUNT = 227
       ATTR_POWER_OFF_RETRACT_COUNT = 228
       ATTR_GMR_HEAD_AMPLITUDE = 230
       ATTR_TEMPERATUREII = 231
       ATTR_READ_ERROR_RETRY_RATE = 250
End Enum

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (LpVersionInformation As OSVERSIONINFO) As Long

Private Type ATTR_DATA
    AttrID As Byte
    AttrName As String
    AttrValue As Byte
    ThresholdValue As Byte
    WorstValue As Byte
    StatusFlags As STATUS_FLAGS
End Type

Public Type DRIVE_INFO
    bDriveType As Byte
    SerialNumber As String
    Model As String
    FirmWare As String
    Cilinders As Long
    Heads As Long
    SecPerTrack As Long
    BytesPerSector As Long
    BytesperTrack As Long
    NumAttributes As Byte
    Attributes() As ATTR_DATA
End Type

Public Enum IDE_DRIVE_NUMBER
    PRIMARY_MASTER
    PRIMARY_SLAVE
    SECONDARY_MASTER
    SECONDARY_SLAVE
End Enum

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function DeviceIoControl Lib "kernel32" (ByVal hDevice As Long, ByVal dwIoControlCode As Long, lpInBuffer As Any, ByVal nInBufferSize As Long, lpOutBuffer As Any, ByVal nOutBufferSize As Long, lpBytesReturned As Long, ByVal lpOverlapped As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000

Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const OPEN_EXISTING = 3
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const CREATE_NEW = 1

Private Const INVALID_HANDLE_VALUE = -1
Dim di As DRIVE_INFO
Dim colAttrNames As Collection

Private Function OpenSmart(drv_num As IDE_DRIVE_NUMBER) As Long
   If IsWindowsNT Then
      OpenSmart = CreateFile("\\.\PhysicalDrive" & CStr(drv_num), GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
   Else
      OpenSmart = CreateFile("\\.\SMARTVSD", 0, 0, ByVal 0&, CREATE_NEW, 0, 0)
   End If
End Function


Private Function CheckSMARTEnable(ByVal hDrive As Long, DriveNum As IDE_DRIVE_NUMBER) As Boolean
   Dim SCIP As SENDCMDINPARAMS
   Dim SCOP As SENDCMDOUTPARAMS
   Dim lpcbBytesReturned As Long
   With SCIP
       .cBufferSize = 0
       With .irDriveRegs
            .bFeaturesReg = SMART_ENABLE_SMART_OPERATIONS
            .bSectorCountReg = 1
            .bSectorNumberReg = 1
            .bCylLowReg = SMART_CYL_LOW
            .bCylHighReg = SMART_CYL_HI

            .bDriveHeadReg = &HA0
            .bCommandReg = IDE_EXECUTE_SMART_FUNCTION
        End With
        .bDriveNumber = DriveNum
   End With
   CheckSMARTEnable = DeviceIoControl(hDrive, DFP_SEND_DRIVE_COMMAND, SCIP, Len(SCIP) - 4, SCOP, Len(SCOP) - 4, lpcbBytesReturned, ByVal 0&)
End Function

Private Function IdentifyDrive(ByVal hDrive As Long, ByVal IDCmd As Byte, ByVal DriveNum As IDE_DRIVE_NUMBER) As Boolean
    Dim SCIP As SENDCMDINPARAMS
    Dim IDSEC As IDSECTOR
    Dim bArrOut(OUTPUT_DATA_SIZE - 1) As Byte
    Dim sMsg As String
    Dim lpcbBytesReturned As Long
    Dim barrfound(100) As Long
    Dim i As Long
    Dim lng As Long

    With SCIP
        .cBufferSize = IDENTIFY_BUFFER_SIZE
        .bDriveNumber = CByte(DriveNum)
        With .irDriveRegs
             .bFeaturesReg = 0
             .bSectorCountReg = 1
             .bSectorNumberReg = 1
             .bCylLowReg = 0
             .bCylHighReg = 0

             .bDriveHeadReg = &HA0
             If Not IsWindowsNT Then .bDriveHeadReg = .bDriveHeadReg Or (DriveNum And 1) * 16

             .bCommandReg = CByte(IDCmd)
        End With
    End With
    If DeviceIoControl(hDrive, DFP_RECEIVE_DRIVE_DATA, SCIP, Len(SCIP) - 4, bArrOut(0), OUTPUT_DATA_SIZE, lpcbBytesReturned, ByVal 0&) Then
       IdentifyDrive = True
       CopyMemory IDSEC, bArrOut(16), Len(IDSEC)
       di.Model = SwapStringBytes(StrConv(IDSEC.sModelNumber, vbUnicode))
       di.FirmWare = SwapStringBytes(StrConv(IDSEC.sFirmwareRev, vbUnicode))
       di.SerialNumber = SwapStringBytes(StrConv(IDSEC.sSerialNumber, vbUnicode))
       di.Cilinders = IDSEC.wNumCyls
       di.Heads = IDSEC.wNumHeads
       di.SecPerTrack = IDSEC.wSectorsPerTrack
    End If
End Function

Private Function ReadAttributesCmd(ByVal hDrive As Long, DriveNum As IDE_DRIVE_NUMBER) As Boolean
   Dim cbBytesReturned As Long
   Dim SCIP As SENDCMDINPARAMS
   Dim drv_attr As DRIVEATTRIBUTE
   Dim bArrOut(OUTPUT_DATA_SIZE - 1) As Byte
   Dim sMsg As String
   Dim i As Long
   With SCIP

       .cBufferSize = READ_ATTRIBUTE_BUFFER_SIZE
       .bDriveNumber = DriveNum
       With .irDriveRegs
            .bFeaturesReg = SMART_READ_ATTRIBUTE_VALUES
            .bSectorCountReg = 1
            .bSectorNumberReg = 1
            .bCylLowReg = SMART_CYL_LOW
            .bCylHighReg = SMART_CYL_HI

            .bDriveHeadReg = &HA0
            If Not IsWindowsNT Then .bDriveHeadReg = .bDriveHeadReg Or (DriveNum And 1) * 16
            .bCommandReg = IDE_EXECUTE_SMART_FUNCTION
       End With
  End With
  ReadAttributesCmd = DeviceIoControl(hDrive, DFP_RECEIVE_DRIVE_DATA, SCIP, Len(SCIP) - 4, bArrOut(0), OUTPUT_DATA_SIZE, cbBytesReturned, ByVal 0&)
  On Error Resume Next
  For i = 0 To NUM_ATTRIBUTE_STRUCTS - 1
      If bArrOut(18 + i * 12) > 0 Then
         di.Attributes(di.NumAttributes).AttrID = bArrOut(18 + i * 12)
         di.Attributes(di.NumAttributes).AttrName = "Unknown value (" & bArrOut(18 + i * 12) & ")"
         di.Attributes(di.NumAttributes).AttrName = colAttrNames(CStr(bArrOut(18 + i * 12)))
         di.NumAttributes = di.NumAttributes + 1
         ReDim Preserve di.Attributes(di.NumAttributes)
         CopyMemory di.Attributes(di.NumAttributes).StatusFlags, bArrOut(19 + i * 12), 2
         di.Attributes(di.NumAttributes).AttrValue = bArrOut(21 + i * 12)
         di.Attributes(di.NumAttributes).WorstValue = bArrOut(22 + i * 12)
      End If
  Next i
End Function

Private Function ReadThresholdsCmd(ByVal hDrive As Long, DriveNum As IDE_DRIVE_NUMBER) As Boolean
   Dim cbBytesReturned As Long
   Dim SCIP As SENDCMDINPARAMS
   Dim IDSEC As IDSECTOR
   Dim bArrOut(OUTPUT_DATA_SIZE - 1) As Byte
   Dim sMsg As String
   Dim thr_attr As ATTRTHRESHOLD
   Dim i As Long, j As Long
   With SCIP

       .cBufferSize = READ_THRESHOLD_BUFFER_SIZE
       .bDriveNumber = DriveNum
       With .irDriveRegs
            .bFeaturesReg = SMART_READ_ATTRIBUTE_THRESHOLDS
            .bSectorCountReg = 1
            .bSectorNumberReg = 1
            .bCylLowReg = SMART_CYL_LOW
            .bCylHighReg = SMART_CYL_HI

            .bDriveHeadReg = &HA0
            If Not IsWindowsNT Then .bDriveHeadReg = .bDriveHeadReg Or (DriveNum And 1) * 16
            .bCommandReg = IDE_EXECUTE_SMART_FUNCTION
       End With
  End With
  ReadThresholdsCmd = DeviceIoControl(hDrive, DFP_RECEIVE_DRIVE_DATA, SCIP, Len(SCIP) - 4, bArrOut(0), OUTPUT_DATA_SIZE, cbBytesReturned, ByVal 0&)
  For i = 0 To NUM_ATTRIBUTE_STRUCTS - 1
      CopyMemory thr_attr, bArrOut(18 + i * Len(thr_attr)), Len(thr_attr)
      If thr_attr.bAttrID > 0 Then
         For j = 0 To UBound(di.Attributes)
             If thr_attr.bAttrID = di.Attributes(j).AttrID Then
                di.Attributes(j).ThresholdValue = thr_attr.bWarrantyThreshold
                Exit For
             End If
         Next j
      End If
  Next i
End Function

Private Function GetSmartVersion(ByVal hDrive As Long, VersionParams As GETVERSIONOUTPARAMS) As Boolean
   Dim cbBytesReturned As Long
   GetSmartVersion = DeviceIoControl(hDrive, DFP_GET_VERSION, ByVal 0&, 0, VersionParams, Len(VersionParams), cbBytesReturned, ByVal 0&)
End Function

Public Function GetDriveInfo(DriveNum As IDE_DRIVE_NUMBER) As DRIVE_INFO
    Dim hDrive As Long
    Dim VerParam As GETVERSIONOUTPARAMS
    Dim cb As Long
    di.bDriveType = 0
    di.NumAttributes = 0
    ReDim di.Attributes(0)
    hDrive = OpenSmart(DriveNum)
    If hDrive = INVALID_HANDLE_VALUE Then Exit Function
    If Not GetSmartVersion(hDrive, VerParam) Then Exit Function
    If Not IsBitSet(VerParam.bIDEDeviceMap, DriveNum) Then Exit Function
    di.bDriveType = 1 + Abs(IsBitSet(VerParam.bIDEDeviceMap, DriveNum + 4))
    If Not CheckSMARTEnable(hDrive, DriveNum) Then Exit Function
    FillAttrNameCollection
    Call IdentifyDrive(hDrive, IDE_ID_FUNCTION, DriveNum)
    Call ReadAttributesCmd(hDrive, DriveNum)
    Call ReadThresholdsCmd(hDrive, DriveNum)
    GetDriveInfo = di
    CloseHandle hDrive
    Set colAttrNames = Nothing
End Function

Private Function IsWindowsNT() As Boolean
   Dim verinfo As OSVERSIONINFO
   verinfo.dwOSVersionInfoSize = Len(verinfo)
   If (GetVersionEx(verinfo)) = 0 Then Exit Function
   If verinfo.dwPlatformId = 2 Then IsWindowsNT = True
End Function

Private Function IsBitSet(iBitString As Byte, ByVal lBitNo As Integer) As Boolean
    If lBitNo = 7 Then
        IsBitSet = iBitString < 0
    Else
        IsBitSet = iBitString And (2 ^ lBitNo)
    End If
End Function

Private Function SwapStringBytes(ByVal sIn As String) As String
   Dim sTemp As String
   Dim i As Integer
   sTemp = Space(Len(sIn))
   For i = 1 To Len(sIn) - 1 Step 2
       Mid(sTemp, i, 1) = Mid(sIn, i + 1, 1)
       Mid(sTemp, i + 1, 1) = Mid(sIn, i, 1)
   Next i
   SwapStringBytes = sTemp
End Function

Public Sub FillAttrNameCollection()
   Set colAttrNames = New Collection
   With colAttrNames
       .Add "ATTR_INVALID", "0"
       .Add "READ_ERROR_RATE", "1"
       .Add "THROUGHPUT_PERF", "2"
       .Add "SPIN_UP_TIME", "3"
       .Add "START_STOP_COUNT", "4"
       .Add "REALLOC_SECTOR_COUNT", "5"
       .Add "READ_CHANNEL_MARGIN", "6"
       .Add "SEEK_ERROR_RATE", "7"
       .Add "SEEK_TIME_PERF", "8"
       .Add "POWER_ON_HRS_COUNT", "9"
       .Add "SPIN_RETRY_COUNT", "10"
       .Add "CALIBRATION_RETRY_COUNT", "11"
       .Add "POWER_CYCLE_COUNT", "12"
       .Add "SOFT_READ_ERROR_RATE", "13"
       .Add "G_SENSE_ERROR_RATE", "191"
       .Add "POWER_OFF_RETRACT_CYCLE", "192"
       .Add "LOAD_UNLOAD_CYCLE_COUNT", "193"
       .Add "TEMPERATURE", "194"
       .Add "REALLOCATION_EVENTS_COUNT", "196"
       .Add "CURRENT_PENDING_SECTOR_COUNT", "197"
       .Add "UNCORRECTABLE_SECTOR_COUNT", "198"
       .Add "ULTRADMA_CRC_ERROR_RATE", "199"
       .Add "WRITE_ERROR_RATE", "200"
       .Add "DISK_SHIFT", "220"
       .Add "G_SENSE_ERROR_RATEII", "221"
       .Add "LOADED_HOURS", "222"
       .Add "LOAD_UNLOAD_RETRY_COUNT", "223"
       .Add "LOAD_FRICTION", "224"
       .Add "LOAD_UNLOAD_CYCLE_COUNTII", "225"
       .Add "LOAD_IN_TIME", "226"
       .Add "TORQUE_AMPLIFICATION_COUNT", "227"
       .Add "POWER_OFF_RETRACT_COUNT", "228"
       .Add "GMR_HEAD_AMPLITUDE", "230"
       .Add "TEMPERATUREII", "231"
       .Add "READ_ERROR_RETRY_RATE", "250"
   End With
End Sub



