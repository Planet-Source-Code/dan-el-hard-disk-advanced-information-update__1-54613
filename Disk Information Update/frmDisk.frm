VERSION 5.00
Begin VB.Form frmDisk 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HDD Information"
   ClientHeight    =   8850
   ClientLeft      =   1470
   ClientTop       =   1020
   ClientWidth     =   13140
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDisk.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   13140
   Begin VB.CheckBox chkCASE_SENSITIVE 
      Height          =   255
      Left            =   12720
      TabIndex        =   61
      Top             =   120
      Width           =   255
   End
   Begin VB.CheckBox chkUNICODE_STORED_ON_DISK 
      Height          =   255
      Left            =   12720
      TabIndex        =   60
      Top             =   600
      Width           =   255
   End
   Begin VB.CheckBox chkPERSISTENT_ACLS 
      Height          =   255
      Left            =   12720
      TabIndex        =   59
      Top             =   2520
      Width           =   255
   End
   Begin VB.CheckBox chkFILE_COMPRESSION 
      Height          =   255
      Left            =   12720
      TabIndex        =   58
      Top             =   1200
      Width           =   255
   End
   Begin VB.CheckBox chkVOL_IS_COMPRESSED 
      Height          =   255
      Left            =   12720
      TabIndex        =   57
      Top             =   960
      Width           =   255
   End
   Begin VB.CheckBox chkCASE_IS_PRESERVED 
      Height          =   255
      Left            =   12720
      TabIndex        =   56
      Top             =   360
      Width           =   255
   End
   Begin VB.CheckBox chkNAMED_STREAMS 
      Height          =   255
      Left            =   12720
      TabIndex        =   55
      Top             =   2040
      Width           =   255
   End
   Begin VB.CheckBox chkSUPPORTS_ENCRYPTION 
      Height          =   255
      Left            =   12720
      TabIndex        =   54
      Top             =   1800
      Width           =   255
   End
   Begin VB.CheckBox chkSUPPORTS_OBJECT_IDS 
      Height          =   255
      Left            =   12720
      TabIndex        =   53
      Top             =   2280
      Width           =   255
   End
   Begin VB.CheckBox chkSUPPORTS_REPARSE_POINTS 
      Height          =   255
      Left            =   12720
      TabIndex        =   52
      Top             =   2760
      Width           =   255
   End
   Begin VB.CheckBox chkSUPPORTS_SPARSE_FILES 
      Height          =   255
      Left            =   12720
      TabIndex        =   51
      Top             =   3000
      Width           =   255
   End
   Begin VB.CheckBox chkVOLUME_QUOTAS 
      Height          =   255
      Left            =   12720
      TabIndex        =   50
      Top             =   1560
      Width           =   255
   End
   Begin VB.CommandButton cmdAboutMe 
      Caption         =   "About Me"
      Height          =   615
      Left            =   6240
      TabIndex        =   48
      Top             =   4560
      Width           =   1335
   End
   Begin VB.TextBox txtAdvancedInfo 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   4560
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   47
      Top             =   6240
      Width           =   4935
   End
   Begin VB.ComboBox cboPrimaryMaster 
      Height          =   315
      Left            =   120
      TabIndex        =   32
      Top             =   5760
      Width           =   3255
   End
   Begin VB.TextBox txtFree 
      Height          =   285
      Left            =   8040
      TabIndex        =   31
      Top             =   6840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtUsed 
      Height          =   285
      Left            =   7800
      TabIndex        =   30
      Top             =   6600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picGraph 
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   4440
      ScaleHeight     =   3375
      ScaleWidth      =   5055
      TabIndex        =   29
      Top             =   120
      Width           =   5055
   End
   Begin VB.CommandButton cmdShowInfo 
      Caption         =   "Show Graph"
      Height          =   615
      Left            =   4680
      TabIndex        =   1
      Top             =   4560
      Width           =   1455
   End
   Begin VB.ComboBox cboDrives 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label25 
      Caption         =   "Vote for me"
      Height          =   255
      Left            =   9720
      TabIndex        =   82
      Top             =   5400
      Width           =   2535
   End
   Begin VB.Label Label24 
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Wingdings 2"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9480
      TabIndex        =   81
      Top             =   5400
      Width           =   255
   End
   Begin VB.Label Label23 
      Caption         =   "Smartvsd.vxd Erased "
      Height          =   255
      Left            =   9720
      TabIndex        =   80
      Top             =   5040
      Width           =   2535
   End
   Begin VB.Label Label21 
      Caption         =   "w"
      BeginProperty Font 
         Name            =   "Wingdings 2"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9480
      TabIndex        =   79
      Top             =   5040
      Width           =   135
   End
   Begin VB.Label Label20 
      Caption         =   "Extended Information of Hard Disk"
      Height          =   255
      Left            =   9720
      TabIndex        =   78
      Top             =   4680
      Width           =   2535
   End
   Begin VB.Label Label19 
      Caption         =   "v"
      BeginProperty Font 
         Name            =   "Wingdings 2"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9480
      TabIndex        =   77
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Label18 
      Caption         =   "Serial Number Hexadecimal"
      Height          =   255
      Left            =   9720
      TabIndex        =   76
      Top             =   4320
      Width           =   2175
   End
   Begin VB.Label Label17 
      Caption         =   "u"
      BeginProperty Font 
         Name            =   "Wingdings 2"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9480
      TabIndex        =   75
      Top             =   4320
      Width           =   255
   End
   Begin VB.Label Label10 
      Caption         =   "Update:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9480
      TabIndex        =   74
      Top             =   3960
      Width           =   2775
   End
   Begin VB.Label lblCASE_SENSITIVE 
      Caption         =   "Case Sensitive Filenames Supported"
      Height          =   255
      Left            =   9840
      TabIndex        =   73
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label lblUNICODE_STORED_ON_DISK 
      Caption         =   "Unicode In Filenames Supported"
      Height          =   255
      Left            =   9840
      TabIndex        =   72
      Top             =   600
      Width           =   2775
   End
   Begin VB.Label lblPERSISTENT_ACLS 
      Caption         =   "Preserve and Enforce ACLs"
      Height          =   255
      Left            =   9840
      TabIndex        =   71
      Top             =   2520
      Width           =   2775
   End
   Begin VB.Label lblFILE_COMPRESSION 
      Caption         =   "File Based Compression Supported"
      Height          =   255
      Left            =   9840
      TabIndex        =   70
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label lblVOL_IS_COMPRESSED 
      Caption         =   "Compressed Volume"
      Height          =   255
      Left            =   9840
      TabIndex        =   69
      Top             =   960
      Width           =   2775
   End
   Begin VB.Label lblCASE_IS_PRESERVED 
      Caption         =   "Preserve Case of Filename Supported"
      Height          =   255
      Left            =   9840
      TabIndex        =   68
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label lblNAMED_STREAMS 
      Caption         =   "Named Streams Supported"
      Height          =   255
      Left            =   9840
      TabIndex        =   67
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Label lblSUPPORTS_ENCRYPTION 
      Caption         =   "Encrypted File System Supported"
      Height          =   255
      Left            =   9840
      TabIndex        =   66
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Label lblSUPPORTS_OBJECT_IDS 
      Caption         =   "Object Identifiers Supported"
      Height          =   255
      Left            =   9840
      TabIndex        =   65
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Label lblSUPPORTS_REPARSE_POINTS 
      Caption         =   "Reparse Points Supported"
      Height          =   255
      Left            =   9840
      TabIndex        =   64
      Top             =   2760
      Width           =   2775
   End
   Begin VB.Label lblSUPPORTS_SPARSE_FILES 
      Caption         =   "Sparse Files Supported"
      Height          =   255
      Left            =   9840
      TabIndex        =   63
      Top             =   3000
      Width           =   2775
   End
   Begin VB.Label lblVOLUME_QUOTAS 
      Caption         =   "Disk Quotas Supported"
      Height          =   255
      Left            =   9840
      TabIndex        =   62
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Label lblSerialNumber2 
      Caption         =   "-"
      Height          =   255
      Left            =   3120
      TabIndex        =   49
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label11 
      Caption         =   "Firmware:"
      Height          =   255
      Left            =   120
      TabIndex        =   46
      Top             =   6600
      Width           =   735
   End
   Begin VB.Label lblFirmware 
      Height          =   255
      Left            =   1440
      TabIndex        =   45
      Top             =   6600
      Width           =   2895
   End
   Begin VB.Label Label22 
      Caption         =   "Model:"
      Height          =   255
      Left            =   120
      TabIndex        =   44
      Top             =   6960
      Width           =   735
   End
   Begin VB.Label lblModelo 
      Height          =   255
      Left            =   1440
      TabIndex        =   43
      Top             =   6960
      Width           =   2895
   End
   Begin VB.Label Label12 
      Caption         =   "Serial Number:"
      Height          =   255
      Left            =   120
      TabIndex        =   42
      Top             =   7320
      Width           =   1095
   End
   Begin VB.Label lblSerialN 
      Height          =   255
      Left            =   1440
      TabIndex        =   41
      Top             =   7320
      Width           =   2895
   End
   Begin VB.Label Label13 
      Caption         =   "Cilinders:"
      Height          =   255
      Left            =   120
      TabIndex        =   40
      Top             =   7680
      Width           =   1095
   End
   Begin VB.Label lblCilinders 
      Height          =   255
      Left            =   1440
      TabIndex        =   39
      Top             =   7680
      Width           =   2895
   End
   Begin VB.Label Label14 
      Caption         =   "Heads:"
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   8040
      Width           =   1095
   End
   Begin VB.Label lblHeads 
      Height          =   255
      Left            =   1440
      TabIndex        =   37
      Top             =   8040
      Width           =   2895
   End
   Begin VB.Label Label15 
      Caption         =   "SecPerTrack:"
      Height          =   255
      Left            =   120
      TabIndex        =   36
      Top             =   8400
      Width           =   1095
   End
   Begin VB.Label lblSecPerTrack 
      Height          =   255
      Left            =   1440
      TabIndex        =   35
      Top             =   8400
      Width           =   2895
   End
   Begin VB.Label Label16 
      Caption         =   "Type:"
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   6240
      Width           =   855
   End
   Begin VB.Label lblType 
      Height          =   255
      Left            =   1440
      TabIndex        =   33
      Top             =   6240
      Width           =   2895
   End
   Begin VB.Label lblINFO0 
      Caption         =   "Basic information of Disks:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   600
      Width           =   3735
   End
   Begin VB.Label lblINFO 
      Caption         =   "Additional information of Disks:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   5400
      Width           =   4335
   End
   Begin VB.Label Label2 
      Caption         =   "Sector per Cluster:"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Bytes per Sectors:"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Free Clusters:"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "Total Clusters:"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label lblSectorPerClusters 
      Height          =   255
      Left            =   2040
      TabIndex        =   22
      Top             =   3720
      Width           =   2295
   End
   Begin VB.Label lblBytesPerClusters 
      Height          =   255
      Left            =   2040
      TabIndex        =   21
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Label lblFreeCluster 
      Height          =   255
      Left            =   2040
      TabIndex        =   20
      Top             =   4440
      Width           =   2295
   End
   Begin VB.Label lblTotalClusters 
      Height          =   255
      Left            =   2040
      TabIndex        =   19
      Top             =   4800
      Width           =   2295
   End
   Begin VB.Label lblPercentUsed 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3120
      TabIndex        =   18
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label lblPercentFree 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3120
      TabIndex        =   17
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label lblPercentTotal 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3120
      TabIndex        =   16
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label lblUsed 
      Height          =   255
      Left            =   2040
      TabIndex        =   15
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Used Space:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Free Space:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Total Space:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label lblVolumeName 
      Height          =   255
      Left            =   2040
      TabIndex        =   11
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Volume Name:"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   10
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label lblFileSystem 
      Height          =   255
      Left            =   2040
      TabIndex        =   9
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label lblSerialNumber 
      Height          =   255
      Left            =   2040
      TabIndex        =   8
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "File System:"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   7
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Serial Number:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label lblFree 
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label lblTotal 
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "Lenght String:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label lblLenghtString 
      Height          =   255
      Left            =   2040
      TabIndex        =   2
      Top             =   3240
      Width           =   2175
   End
End
Attribute VB_Name = "frmDisk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Graph As New clsGraph

Private Sub cboDrives_Click()
picGraph.Refresh

cmdShowInfo_Click
End Sub

Private Sub cboPrimaryMaster_Click()
   Dim Drv_Info As DRIVE_INFO
   Dim sInfo As String
   Dim sInfoAttr As String
   Dim i As Integer
   
   Drv_Info = GetDriveInfo(cboPrimaryMaster.ListIndex)
   
   sInfo = cboPrimaryMaster.List(cboPrimaryMaster.ListIndex)
   
   sInfo = sInfo & ": " & vbCrLf
   
   txtAdvancedInfo.Text = ""
   
With Drv_Info
  
   lblFirmware.Caption = .FirmWare
   lblModelo.Caption = .Model
   lblSerialN.Caption = .SerialNumber
   lblCilinders.Caption = .Cilinders
   lblHeads.Caption = .Heads
   lblSecPerTrack.Caption = .SecPerTrack

   
       If .bDriveType = 0 Then sInfo = sInfo & "[Not present]"
       lblType.Caption = "[Not present]"
      
       
       If .bDriveType = 2 Then sInfo = sInfo & "[ATAPI drive - info not available]"
       lblType.Caption = "[ATAPI drive - info not available]"
  
       If .bDriveType = 1 Then
             
          lblType.Caption = "[IDE drive]"
          
 sInfoAttr = Format("Attribute Name", String(24, "@") & "!") & Format("Value", String(7, "@") & "!") & Format("Threshold", String(11, "@") & "!") & Format("WorstValue", String(12, "@") & "!") & "Status" & vbCrLf
sInfoAttr = sInfoAttr & String(60, "-") & vbCrLf
          
          For i = 1 To .NumAttributes - 1
              sInfoAttr = sInfoAttr & Format(.Attributes(i).AttrName, String(25, "@") & "!") & Format(.Attributes(i).AttrValue, String(3, "@")) & vbTab & Space(2) & Format(.Attributes(i).ThresholdValue, String(3, "@")) & vbTab & Format(.Attributes(i).WorstValue, String(8, "@")) & vbTab & Format("&H" & Hex(.Attributes(i).StatusFlags), String(4, "@"))
              sInfoAttr = sInfoAttr & vbNewLine
          Next i
          
          txtAdvancedInfo.Text = sInfoAttr
       End If
   End With
   
End Sub

Private Sub cmdShowInfo_Click()
Dim Bytes_Avail As LARGE_INTEGER
Dim Bytes_Total As LARGE_INTEGER
Dim Bytes_Free As LARGE_INTEGER

lReturn = GetDiskFreeSpace(sDrive, lSectorsPerCluster, lBytesPerSector, lFreeClusters, lTotalClusters)

lblSectorPerClusters.Caption = lSectorsPerCluster
lblBytesPerClusters.Caption = lBytesPerSector
lblFreeCluster.Caption = lFreeClusters
lblTotalClusters.Caption = lTotalClusters

On Error Resume Next

GetDiskFreeSpaceEx cboDrives.Text, Bytes_Avail, Bytes_Total, Bytes_Free


Dbl_Total = LargeIntegerToDouble(Bytes_Total.Lowpart, Bytes_Total.Highpart)
Dbl_Free = LargeIntegerToDouble(Bytes_Free.Lowpart, Bytes_Free.Highpart)

    lblTotal.Caption = SizeString(Dbl_Total)
    lblFree.Caption = SizeString(Dbl_Free)
    lblUsed.Caption = SizeString(Dbl_Total - Dbl_Free)
    
    lblPercentTotal.Caption = Format$(1, "percent")
    lblPercentFree.Caption = Format$(Dbl_Free / Dbl_Total, "percent")
    lblPercentUsed.Caption = Format$((Dbl_Total - Dbl_Free) / Dbl_Total, "percent")
    
txtFree.Text = Format$(Dbl_Free / Dbl_Total) * 100
txtUsed.Text = Format$((Dbl_Total - Dbl_Free) / Dbl_Total) * 100
    
    Root = cboDrives.Text
    Volume_Name = Space$(1024)
    File_System_Name = Space$(1024)

'*********

If GetVolumeInformation(Root, Volume_Name, Len(Volume_Name), Serial_Number, Max_Component_Length, File_System_Flags, File_System_Name, Len(File_System_Name)) = 0 Then
 
   picGraph.Cls
    
    lblPercentTotal.Caption = ""
    lblPercentFree.Caption = ""
    lblPercentUsed.Caption = ""
    lblVolumeName.Caption = ""
    lblSerialNumber.Caption = ""
    lblFileSystem.Caption = ""
    lblLenghtString.Caption = ""
    lblSectorPerClusters.Caption = ""
    lblBytesPerClusters.Caption = ""
    lblFreeCluster.Caption = ""
    lblTotalClusters.Caption = ""
    lblSerialNumber2.Caption = ""

    chkCASE_IS_PRESERVED.Value = 0
    chkCASE_SENSITIVE.Value = 0
    chkUNICODE_STORED_ON_DISK.Value = 0
    chkPERSISTENT_ACLS.Value = 0
    chkFILE_COMPRESSION.Value = 0
    chkVOL_IS_COMPRESSED.Value = 0
    chkNAMED_STREAMS.Value = 0
    chkSUPPORTS_ENCRYPTION.Value = 0
    chkSUPPORTS_OBJECT_IDS.Value = 0
    chkSUPPORTS_REPARSE_POINTS.Value = 0
    chkSUPPORTS_SPARSE_FILES.Value = 0
    chkVOLUME_QUOTAS.Value = 0
    
    MsgBox "No Disk!", vbExclamation, "Error when Reading the Disk"

Exit Sub
        
        
    End If
    
 '**********
    Dim VolumeNameBuffer As String * 11
    Dim VolumeSerialNumber As Long
    Dim MaximumComponentLength As Long
    Dim FileSystemFlags As Long
    Dim FileSystemNameBuffer As String

   If GetVolumeInformation(Left$(cboDrives, 2) & "\", VolumeNameBuffer, Len(VolumeNameBuffer), VolumeSerialNumber, MaximumComponentLength, FileSystemFlags, FileSystemNameBuffer, Len(FileSystemNameBuffer)) = 0 Then
   
        Exit Sub
    End If

 lblSerialNumber2.Caption = Hex(VolumeSerialNumber)
 
     If FileSystemFlags And FS_CASE_IS_PRESERVED Then chkCASE_IS_PRESERVED.Value = 1
    If FileSystemFlags And FS_CASE_SENSITIVE Then chkCASE_SENSITIVE.Value = 1
    If FileSystemFlags And FS_UNICODE_STORED_ON_DISK Then chkUNICODE_STORED_ON_DISK.Value = 1
    If FileSystemFlags And FS_PERSISTENT_ACLS Then chkPERSISTENT_ACLS.Value = 1
    If FileSystemFlags And FS_FILE_COMPRESSION Then chkFILE_COMPRESSION.Value = 1
    If FileSystemFlags And FS_VOL_IS_COMPRESSED Then chkVOL_IS_COMPRESSED.Value = 1
    'Second half
    If FileSystemFlags And FILE_NAMED_STREAMS Then chkNAMED_STREAMS.Value = 1
    If FileSystemFlags And FILE_SUPPORTS_ENCRYPTION Then chkSUPPORTS_ENCRYPTION.Value = 1
    If FileSystemFlags And FILE_SUPPORTS_OBJECT_IDS Then chkSUPPORTS_OBJECT_IDS.Value = 1
    If FileSystemFlags And FILE_SUPPORTS_REPARSE_POINTS Then chkSUPPORTS_REPARSE_POINTS.Value = 1
    If FileSystemFlags And FILE_SUPPORTS_SPARSE_FILES Then chkSUPPORTS_SPARSE_FILES.Value = 1
    If FileSystemFlags And FILE_VOLUME_QUOTAS Then chkVOLUME_QUOTAS.Value = 1

 
 '********
    Pos = InStr(Volume_Name, Chr$(0))
    Volume_Name = Left$(Volume_Name, Pos - 1)
    lblVolumeName.Caption = Volume_Name

    lblSerialNumber.Caption = Format$(Serial_Number)

    Pos = InStr(File_System_Name, Chr$(0))
    File_System_Name = Left$(File_System_Name, Pos - 1)
    lblFileSystem.Caption = File_System_Name

lblLenghtString.Caption = Format$(Max_Component_Length)

'********************************
    Graph.AddSegment txtFree.Text, "Free Space", &HFF00FF            'Magenta'
    Graph.AddSegment txtUsed.Text, "Used Space", &HFF0000            'Blue'
    Graph.DrawPie picGraph.hdc, picGraph.hwnd, True, "HDD Information"
    
    Graph.Clear
    
'******************************

End Sub

Private Sub Form_Load()
lBuffer = 26 * 4 + 1
sDriveNames = Space$(lBuffer)
lReturn = GetLogicalDriveStrings(lBuffer, sDriveNames)
nOffset = 1

Do

sTempStr = Mid$(sDriveNames, nOffset, 3)
If Left$(sTempStr, 1) = vbNullChar Then Exit Do
cboDrives.AddItem UCase(sTempStr)
nOffset = nOffset + 4

Loop

cboDrives.ListIndex = 1

With cboPrimaryMaster
       .AddItem "PRIMARY_MASTER"
       .AddItem "PRIMARY_SLAVE"
       .AddItem "SECONDARY_MASTER"
       .AddItem "SECONDARY_SLAVE"
End With

cboPrimaryMaster.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Graph = Nothing
End Sub

Private Sub cmdAboutMe_Click()
MsgBox "HDD Information v1.0  25-06-2004" & vbNewLine & vbNewLine & "Info:" & vbNewLine & "HDD Information Show Firmware, Space Free, Model, Name" & vbNewLine & "Space Used, Graph, System Files, Etc of a Hard Disk" & vbNewLine & vbNewLine & "Vote for Me" & vbNewLine & vbNewLine & "Planet Source Code: www.planet-source-code.com/vb" & vbNewLine & vbNewLine & "Thank to vbAcceletaror.com , Planet Source Code, Trevor Newsome for Extract de Code Used in HDD." & vbNewLine & vbNewLine & "Daniel Carrasco Olguin (Daniel PC)" & vbNewLine & vbNewLine & "Santiago de Chile" & vbNewLine & vbNewLine & "Vote for Me", vbInformation + vbOKOnly

End Sub

