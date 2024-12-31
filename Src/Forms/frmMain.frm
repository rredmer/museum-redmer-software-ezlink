VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0E0FF&
   Caption         =   "RSC EZ-LINK v 2.00 r 002"
   ClientHeight    =   7275
   ClientLeft      =   2670
   ClientTop       =   1965
   ClientWidth     =   9630
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7275
   ScaleWidth      =   9630
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar tlbCommLog 
      Height          =   630
      Left            =   60
      TabIndex        =   9
      Top             =   6615
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   1111
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Wrappable       =   0   'False
      ImageList       =   "imgMain"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgMain 
      Left            =   5670
      Top             =   6630
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":062C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":094E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrTimeOut 
      Left            =   8595
      Top             =   6750
   End
   Begin MSDataGridLib.DataGrid dbgLogFile 
      Bindings        =   "frmMain.frx":0C68
      Height          =   3900
      Left            =   75
      TabIndex        =   8
      Top             =   2625
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   6879
      _Version        =   393216
      AllowUpdate     =   -1  'True
      Appearance      =   0
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataMember      =   "cmdLogFile"
      Caption         =   "COMMUNICATIONS LOG"
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "order_id"
         Caption         =   "order_id"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "unit_id"
         Caption         =   "unit_id"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "log_time"
         Caption         =   "log_time"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Number_of_Setups"
         Caption         =   "Number_of_Setups"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Archive"
         Caption         =   "Archive"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "FileName"
         Caption         =   "FileName"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1365.165
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1365.165
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1679.811
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   2085.166
         EndProperty
      EndProperty
   End
   Begin MSDataListLib.DataCombo dbcSerialPort 
      Bindings        =   "frmMain.frx":0C79
      DataField       =   "Serial Port"
      DataMember      =   "cmdConfig"
      DataSource      =   "DE"
      Height          =   315
      Left            =   1185
      TabIndex        =   7
      Top             =   360
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      ListField       =   "SerialPortDescription"
      BoundColumn     =   "SerialPortNumber"
      Text            =   ""
      Object.DataMember      =   "cmdSerialPorts"
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Height          =   285
      Left            =   3180
      TabIndex        =   6
      Top             =   360
      Width           =   765
   End
   Begin VB.DriveListBox drvPath 
      Height          =   315
      Left            =   1185
      TabIndex        =   4
      Top             =   690
      Width           =   4260
   End
   Begin VB.DirListBox dirPath 
      Height          =   1440
      Left            =   1185
      TabIndex        =   3
      Top             =   1035
      Width           =   4275
   End
   Begin VB.TextBox txtCommStatus 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1185
      TabIndex        =   0
      Top             =   60
      Width           =   8385
   End
   Begin MSCommLib.MSComm comPORT1 
      Left            =   9120
      Top             =   6750
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DTREnable       =   -1  'True
      Handshaking     =   2
      RThreshold      =   1
      RTSEnable       =   -1  'True
      BaudRate        =   19200
   End
   Begin VB.Label lblPath 
      BackStyle       =   0  'Transparent
      Caption         =   "Save Files In"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   90
      TabIndex        =   5
      Top             =   735
      Width           =   1035
   End
   Begin VB.Label lblSerialPort 
      BackStyle       =   0  'Transparent
      Caption         =   "Serial Port"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   90
      TabIndex        =   2
      Top             =   375
      Width           =   1065
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   90
      TabIndex        =   1
      Top             =   60
      Width           =   885
   End
   Begin VB.Menu nmuMain 
      Caption         =   "&File"
      Index           =   1
      Begin VB.Menu mnuFile 
         Caption         =   "E&xit"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------------------
'
' Application..:  RSC EZ-LINK(r)
'
' Module.......:  frmMain - RSC EZ-LINK main form
'
' Procedure....:  (General)(declarations)
'
' Description..:  Declare form level variables.
'
' History......:
'                 07-24-96 RDR Designed and Programmed.
'
' (c) 1998 Redmer Software Company.  All Rights Reserved.
'
'--------------------------------------------------------------------------------------------
Option Explicit                                         'Require explicit variable declarations
Private Const EZ_COM_CRC  As String = "C"                    'X-modem CRC Character
Private Const EZ_COM_RETRIES As Integer = 50                 'Number of serial retries
Private Const EZ_COM_XMODEM_PAD As Integer = &H1A            'X-modem Block Padding Character
Private Const EZ_COM_SOH  As Integer = &H1                   'X-modem SOH
Private Const EZ_COM_EOT  As Integer = &H4                   'X-modem EOT
Private Const EZ_COM_ACK  As Integer = &H6                   'X-modem ACK
Private Const EZ_COM_NAK  As Integer = &H15                  'X-modem NAK
Private Const EZ_COM_CAN  As Integer = &H18                  'X-modem CAN
Dim iSecondsElapsed As Integer                          'Timer Variable for xmodem timeout
Dim iTimeOut As Integer                                 'X-Modem Timeout Counter
Dim iTries As Integer                                   'X-Modem Retry Counter

'--------------------------------------------------------------------------------------------
'
' Procedure....:  Form_Load()
'
' Description..:  Set data control recordsets.
'
'--------------------------------------------------------------------------------------------
Private Sub Form_Load()
    
    '--- Initialize local error handler -----------------------------------------------------
    On Error GoTo ErrorHandler
    
    '--- Initialize the form variables ------------------------------------------------------
    txtCommStatus.Text = "Idle."
    Refresh
    
    '--- Initialize data controls -----------------------------------------------------------
    cmdOpen_Click
    
    dirPath.Path = DE.rscmdConfig("StoragePath").Value
    drvPath.Drive = Left$(DE.rscmdConfig("StoragePath").Value, 2)
    
    bTerminalMode = True                                    'Set the terminal more
  
    Exit Sub
    
'--- Local error handler --------------------------------------------------------------------
ErrorHandler:

    Resume Next
    
End Sub

'--------------------------------------------------------------------------------------------
'
' Procedure....:  tlbCommLog_ButtonClick(ByVal Button As ComctlLib.Button)
'
' Description..:  Process communications log toolbar buttons.
'
'--------------------------------------------------------------------------------------------
Private Sub tlbCommLog_ButtonClick(ByVal Button As MSComctlLib.Button)

    '--- Dimension local variables ----------------------------------------------------------
    Dim iAns As Integer                                         'Response to delete confirmation

    '--- Initialize local error handler -----------------------------------------------------
    On Error GoTo ErrorHandler
    
    '--- Process buttons --------------------------------------------------------------------
    Select Case Button.Index
    
        '--- Delete current record ----------------------------------------------------------
        Case 1:
            mnuFile_Click 0
        Case 2:
            iAns = MsgBox("Delete this record?", vbApplicationModal + vbYesNo + vbQuestion, EZ_CAPTION)
            If iAns = vbYes Then
                DE.rscmdLogFile.Delete adAffectCurrent
                dbgLogFile.Refresh
            End If
        
        '--- View file data -----------------------------------------------------------------
        Case 3:
        
            frmView.Show 1
        
    End Select

    Exit Sub
    
ErrorHandler:

    MsgBox "Error in communications log toolbar." & vbCr & "Please select a valid port.", vbInformation + vbApplicationModal + vbOKOnly, EZ_CAPTION

End Sub


'--------------------------------------------------------------------------------------------
'
' Procedure....:  (comPORT1)(OnComm)
'
' Description..:  Handle serial communication events.
'
'--------------------------------------------------------------------------------------------
Private Static Sub comPORT1_OnComm()
    
    '--- Dimension local variables ----------------------------------------------------------
    Dim sSerialBuffer As String                                 'Complete file buffer
    Dim sSNAPVer As String                                      'EZ-SNAP version number
    Dim lReceiveCount As Long                                   'Number of characters received from Symbol
    Dim sOrder_ID As String                                     'Current school being transferred from PDT
    Dim sUnitID As String                                       'Current unit in the Cradle
    Dim sTempSerialBuffer As String                             'Temporary serial input buffer
    Dim strmTmpSerialBuffer As String                           'Temporary serial input buffer
    Dim iFileHandle As Integer                                  'File Handle for individual source files
    Dim iOutputFile As Integer                                  'File Handle for complete output file
    Dim iNumberOfSetups As Integer                              'Number of setups in the setup file
    Dim iTwinCheck As String                                    'Current Twin Check number of film
    Dim ncount As Integer                                       'Index value
    Dim sDataPath As String                                     'Path to data files
    Dim sCopyBuffer As String                                   'Memory copy of buffer
    
    '---- Initialize Local Error Handler ----------------------------------------------------
    On Error GoTo ErrorHandler                                  'Local error handler

    '--- Initialize local variables ---------------------------------------------------------
    iNumberOfSetups = 1                                         'Default to 1 setup record
    iFileHandle = FreeFile                                      'Get a free handle for output file
    iOutputFile = FreeFile                                      'Get a free handle for output file
    
    '--- Handle specific communication control events ---------------------------------------
    Select Case comPORT1.CommEvent
        
        '--- Receive character event --------------------------------------------------------
        Case comEvReceive
        
            sTempSerialBuffer = comPORT1.Input
                    
            '--- If the mode is set to receive info from the symbol -------------------------
            If bTerminalMode = True Then
                
                '--- If the start of text characters are in the input buffer ----------------
                If InStr(sTempSerialBuffer, "S") Then
                    
                    '--- Start receiving the information from the symbol --------------------
                    comPORT1.RThreshold = 0                     'Set comm receive threshold to 0
                    bTerminalMode = False                       'Disable terminal mode
                    txtCommStatus.Text = "Receiving."           'Display receiving message
                    txtCommStatus.Refresh                       'Reshresh controls
                    sSerialBuffer = ""                          'Clear serial buffer
                    lReceiveCount = 0                           'Initialize receive byte counter
                    comPORT1.InputLen = 0                       'Clear the input buffer
                    
                    '--- Download the three data files using XMODEM -------------------------
                    Open "SETFILE.TXT" For Output As #iFileHandle
                    Xmodem_DownLoad (iFileHandle)
                    Close #iFileHandle
                    Open "UNITFILE.TXT" For Output As #iFileHandle
                    Xmodem_DownLoad (iFileHandle)
                    Close #iFileHandle
                    Open "CRDFILE.TXT" For Output As #iFileHandle
                    Xmodem_DownLoad (iFileHandle)
                    Close #iFileHandle
                    
                    '--- Copy Files into single file ----------------------------------------
                    txtCommStatus.Text = "Saving order information."
                    txtCommStatus.Refresh
                    sCopyBuffer = ""
                    
                    Open "OUTFILE.TXT" For Output As #iOutputFile
                    iFileHandle = FreeFile
                    
                    Open "UNITFILE.TXT" For Input Access Read As #iFileHandle
                    Input #iFileHandle, sSerialBuffer
                    sSNAPVer = sSerialBuffer
                    Input #iFileHandle, sSerialBuffer
                    sUnitID = sSerialBuffer
                    Close #iFileHandle
                    Print #iOutputFile, sSNAPVer
                    Print #iOutputFile, sUnitID
                    sCopyBuffer = sSNAPVer & vbCr & vbLf & sUnitID & vbCr & vbLf
                    iFileHandle = FreeFile
                    Open "SETFILE.TXT" For Input Access Read As #iFileHandle
                    Do While Not EOF(iFileHandle)
                        'Load setup string
                        Input #iFileHandle, sSerialBuffer
                        strmTmpSerialBuffer = sSerialBuffer
                        'Delete Underscores - JDM
                        For ncount = 1 To Len(strmTmpSerialBuffer)
                            If Mid$(strmTmpSerialBuffer, ncount, 1) = "_" Then
                               Mid$(strmTmpSerialBuffer, ncount, 1) = " "
                            End If
                        Next
                        'Parse School ID
                        sOrder_ID = Mid$(strmTmpSerialBuffer, 16, 4)
                        'Added twincheck - JDM
                        iTwinCheck = Mid$(strmTmpSerialBuffer, 56, 10)
                        iNumberOfSetups = iNumberOfSetups + 1
                        Print #iOutputFile, sSerialBuffer
                        sCopyBuffer = sCopyBuffer & sSerialBuffer & vbCr & vbLf
                    Loop
                    Close #iFileHandle
                    
                    iFileHandle = FreeFile
                    'Open the Camera Card File
                    Open "CRDFILE.TXT" For Input Access Read As #iFileHandle
                    Do While Not EOF(iFileHandle)
                        'Parse the raw data string into fields
                        Input #iFileHandle, sSerialBuffer
                        Print #iOutputFile, sSerialBuffer
                        sCopyBuffer = sCopyBuffer & sSerialBuffer & vbCr & vbLf
                    Loop
                    Close #iFileHandle
                    Close #iOutputFile
                    'Rename the output file using the specified naming style
                    sDataPath = dirPath.Path & "\" & Trim$(sUnitID) & "_" & Format$(Now, "yymmdd_hhnnss") & ".TXT"
                    If Len(Dir$(sDataPath)) > 0 Then
                        Kill (sDataPath)
                    End If
                    
                    Name "outfile.txt" As sDataPath
                    
                    
                    '--- Add log record to the POS Log Table --------------------------------
                    With DE.rscmdLogFile
                        .AddNew
                        .Fields("order_id") = sOrder_ID
                        .Fields("unit_id") = sUnitID
                        .Fields("log_time") = Now
                        .Fields("number_of_setups") = iNumberOfSetups
                        .Fields("Archive") = sCopyBuffer
                        .Fields("FileName") = sDataPath
                        .UpdateBatch adAffectAllChapters
                    End With
                    dbgLogFile.Refresh
                    
                    txtCommStatus.Text = "Idle."
                    txtCommStatus.Refresh
                    bTerminalMode = True
                    comPORT1.RThreshold = 1
                
                End If
            
            End If
 
            '----------------------------------------------------------------------------
        Case comEvSend
        Case comEvCTS
        Case comEvDSR
        Case comEvCD
        Case comEvRing
        Case comEvEOF
        Case comBreak
        Case comCTSTO
        Case comDSRTO
        Case comFrame
        Case comOverrun
        Case comCDTO
        Case comRxOver
        Case comRxParity
        Case comTxFull
        Case Else
    End Select
    
    Exit Sub

ErrorHandler:

    Select Case Err
        Case 53:    'File not found
            'ignore
        Case Else:
            'an error occurred
            MsgBox Error$, vbApplicationModal + vbInformation + vbOKOnly, EZ_CAPTION
    End Select
    Resume Next
End Sub

'--------------------------------------------------------------------------------------------
'
' Procedure....:  (mnuFile)(Click)
'
' Description..:  Process menu selections
'
'--------------------------------------------------------------------------------------------
Private Sub mnuFile_Click(Index As Integer)

    '--- Dimension local variables ----------------------------------------------------------
    Dim iAns As Integer                                         'Return from MsgBox Prompt

    '--- Initialize local error handler -----------------------------------------------------
    On Error GoTo ErrorHandler

    '--- Process File Menu selections -------------------------------------------------------
    Select Case Index
        
        '--- Exit the Program ---------------------------------------------------------------
        Case 0:
            iAns = MsgBox("Exit the program?", vbApplicationModal + vbYesNo + vbQuestion, EZ_CAPTION)
            If iAns = vbYes Then
                comPORT1.DTREnable = True                       'Toggle DTR to trigger EZSNAP out of comm loop
                comPORT1.DTREnable = False
                comPORT1.DTREnable = True
                comPORT1.DTREnable = False
                End                                             'End the program
            End If

    End Select

    Exit Sub
    
ErrorHandler:

    MsgBox "Error in file menu." & EZ_MSG_TECH_SUPPORT, vbApplicationModal + vbOKOnly + vbInformation, EZ_CAPTION
    Resume Next
    
End Sub

'--------------------------------------------------------------------------------------------
'
' Procedure....:  (General)(Xmodem_DownLoad)
'
' Description..:  Downlaod a file from remote client using Xmodem protocol.
'
'--------------------------------------------------------------------------------------------
Public Sub Xmodem_DownLoad(intvFileHandle)

    '--- Dimension local variables ----------------------------------------------------------
    Dim intmInputBuffer As Integer
    Dim intmCheckSum As Integer
    Dim intmBlock As Integer
    Dim intmBlocksComplement As Integer
    Dim intmRemoteChecksum As Integer
    Dim intmRemoteComplement As Integer
    Dim intmRemoteBlockNumber As Integer
    Dim intmSOHCharacter As Integer
    Dim intmBlockCount As Integer
    Dim intmFilePosition As Integer
    Dim strmBuffer As String
    Dim blnFirstBlock As Boolean


    '--- Initialize local error handler -----------------------------------------------------
    On Error GoTo ErrorHandler

    intmBlock = 0
    iTries = 0
    iTimeOut = 50
    intmSOHCharacter = 0
    blnFirstBlock = True
    bTerminalMode = False                'Disable output to terminal
    comPORT1.InBufferCount = 0  'Flush the Input buffer
    comPORT1.RThreshold = 0     'Disable generation  of OnComm Event
    comPORT1.InputLen = 1       'Receive one char at a time
    
    ' send NAKs until the sender starts sending
    Do While (intmSOHCharacter <> EZ_COM_SOH) And (iTries < EZ_COM_RETRIES)
        iTries = iTries + 1
        comPORT1.Output = Chr$(EZ_COM_NAK)
        Delay 2
        intmSOHCharacter = ReadComm()
        If intmSOHCharacter <> EZ_COM_SOH Then
            Delay 2
        End If
    Loop
    
    iTries = 0
    
    Do While iTries < EZ_COM_RETRIES
        
        ' -- Receive the data and build the file --
        If Not (blnFirstBlock) Then
            
            iTimeOut = 10
            intmSOHCharacter = ReadComm()
            
            If intmSOHCharacter = EZ_COM_CAN Then
                MsgBox "CAN Received", vbApplicationModal + vbInformation + vbOKOnly, EZ_CAPTION
                Exit Do
            End If
         
            If intmSOHCharacter = EZ_COM_EOT Then
                comPORT1.Output = Chr$(EZ_COM_ACK)
                Exit Do
            End If
            
        End If
        blnFirstBlock = False
        
        'iTimeOut = 1                         ' Switch to one sec. timeouts
        iTimeOut = 3                         ' Switch to one sec. timeouts
        
        intmRemoteBlockNumber = ReadComm()      ' Read block number
        intmRemoteComplement = ReadComm()       ' Read 1's complement
        intmCheckSum = 0
        txtCommStatus.Text = "Block: " + Str(intmRemoteBlockNumber)
        txtCommStatus.Refresh
        strmBuffer = ""
        
        ' ---- data block -----
        For intmBlockCount = 1 To 128
            intmInputBuffer = ReadComm()
            strmBuffer = strmBuffer + Chr$(intmInputBuffer)
            intmCheckSum = intmCheckSum + intmInputBuffer
        Next
        ' ---- checksum  from sender ----
        intmRemoteChecksum = ReadComm()
        intmCheckSum = intmCheckSum And 255
        
        'MsgBox "[" & Str(intmRemoteBlockNumber) & "][" & Len(strmBuffer) & "][" & Str(intmRemoteChecksum) & "," & Str(intmCheckSum) & "]" & strmBuffer
        

        ' --- Handle resent blocks ---
        If intmRemoteBlockNumber = intmBlock Then
            intmFilePosition = Seek(intvFileHandle)
            Seek intvFileHandle, intmFilePosition - 128
        
        ' --- handle out of synch block numbers ---
        ElseIf intmRemoteBlockNumber <> (intmBlock + 1) Then
            'MsgBox "No next sequential block"
            'Exit Do
        End If
        intmBlock = intmRemoteBlockNumber
        
        ' --- test the block # 1's complement ---
        intmBlocksComplement = (Not intmRemoteBlockNumber And &HFF)
        If (intmRemoteComplement And &HFF) <> intmBlocksComplement Then
            receive_error "One's complement does not match", EZ_COM_NAK
        End If
        
        ' --- test chksum or crc vs one sent ---
        If intmCheckSum <> intmRemoteChecksum Then
            ' MsgBox "CheckSum = " & Str(intmCheckSum) & " remote = " & intmRemoteChecksum
            ' MsgBox strmBuffer
            ' receive_error "non-matching Checksums", EZ_COM_NAK
        End If
      
        ' --- write the block to disk ---
        For intmBlockCount = 1 To Len(strmBuffer)
            Print #intvFileHandle, Mid(strmBuffer, intmBlockCount, 1);
        Next intmBlockCount
        
        comPORT1.InBufferCount = 0  'Flush the Input buffer
        comPORT1.Output = Chr$(EZ_COM_ACK)
        Delay 1.5
    Loop
    
    If intmSOHCharacter = EZ_COM_EOT Then
        'MsgBox "Transfer Complete"
    Else
        MsgBox "Transfer Aborted   " & Str$(intmSOHCharacter), vbApplicationModal + vbInformation + vbOKOnly, EZ_CAPTION
    End If
    
    iTimeOut = 10
    comPORT1.InBufferCount = 0                  'Flush the buffer
    comPORT1.InputLen = 0                       'Receive all chars in buffer
    comPORT1.RThreshold = 0                     'Enable generation of OnComm Event
    bTerminalMode = True                        'Enable output to terminal

    Exit Sub

ErrorHandler:

        MsgBox "Communications error.", vbApplicationModal + vbInformation + vbOKOnly, EZ_CAPTION
        Resume Next

End Sub



'--------------------------------------------------------------------------------------------
'
' Procedure....:  (tmrTimeOut)(Timer)
'
' Description..:  Increment global timeout counter on timer event.
'
'--------------------------------------------------------------------------------------------
Private Sub tmrTimeOut_Timer()
    
    iSecondsElapsed = iSecondsElapsed + 1

End Sub


'--------------------------------------------------------------------------------------------
'
' Procedure....:  (General)(Delay)
'
' Description..:  Waste time for a specified number of milliseconds.
'
'--------------------------------------------------------------------------------------------
Sub Delay(intvDelaySeconds)
    
    Dim intmEventDivider As Long                        'Process events on this divider
    
    intmEventDivider = 0                                'Initialize event divider
    iSecondsElapsed = 0                                 'Initialize the number of seconds elapsed
    
    If intvDelaySeconds < 1 Then
        tmrTimeOut.Interval = 100 * intvDelaySeconds
    Else
        tmrTimeOut.Interval = 100
    End If
    tmrTimeOut.Enabled = True                           'Enable Timer Control
    
    Do While iSecondsElapsed <= intvDelaySeconds        'While waiting...
        
        If intmEventDivider Mod 10 = 0 Then DoEvents
        
        intmEventDivider = intmEventDivider + 1
    
    Loop
    tmrTimeOut.Enabled = False

End Sub


'--------------------------------------------------------------------------------------------
'
' Procedure....:  (General)(ReadComm)
'
' Description..:  Read characters from the msComm Control Input Property.
'                 ReadComm reads a character from the Comm control's input buffer
'                 and returns the ASCII value of that character. If a null string is
'                 encountered, ReadComm returns 0.
'
'--------------------------------------------------------------------------------------------
Function ReadComm() As Integer
    
    '--- Dimension local variables ----------------------------------------------------------
    Dim strmTempBuffer As String                    'Temporary Buffer
 
    '--- Initialize local error handler -----------------------------------------------------
    On Error GoTo ErrorHandler
    
    If comPORT1.InBufferCount > 0 Then  'If the input bufer has data in it
        strmTempBuffer = comPORT1.Input 'Retrieve characters from the MSComm Control
        If strmTempBuffer <> "" Then                'If the return is not null
            ReadComm = Asc(strmTempBuffer)          'Get the ASCII value of the character
        Else
            ReadComm = 0                            'Set return to empty
        End If
    Else
        ReadComm = 0                                'Set return to empty
    End If

    Exit Function

'--- Local error handler --------------------------------------------------------------------
ErrorHandler:

        MsgBox "Communications error (ReadComm).", vbApplicationModal + vbInformation + vbOKOnly, EZ_CAPTION
        Resume Next


End Function


'--------------------------------------------------------------------------------------------
'
' Procedure....:  (General)(receive_error)
'
' Description..:  Display error message and establish retries
'
'--------------------------------------------------------------------------------------------
Static Sub receive_error(ErrorMsg, Rtn)
    
    iTries = iTries + 1
    
    If iTimeOut = 1 Then
        
        MsgBox "error  " + ErrorMsg, vbApplicationModal + vbInformation + vbOKOnly, EZ_CAPTION
    
    End If

End Sub


'--------------------------------------------------------------------------------------------
'
' Procedure....:  (General)(SendToModem)
'
' Description..:  Send text to serial port.
'
'--------------------------------------------------------------------------------------------
'
Public Sub SendToModem(sSendString As String)

    '--- Dimension local variables ----------------------------------------------------------
    Dim iDoEventsLoop As Integer                            'Loop counting variable

    '--- Initialize local error handler -----------------------------------------------------
    On Error GoTo ErrorHandler                              'Local error handler
    
    '--- Set thresholds to 0-comm event won't be generated ----------------------------------
    comPORT1.RThreshold = 0                                 'Set receive threshold off
    comPORT1.SThreshold = 0                                 'Set send threshold off

    '--- Send command to modem, wait for response -------------------------------------------
    comPORT1.Output = sSendString                           'Send string out serial port
    
    '--- Allow enough doevents to wait for modem --------------------------------------------
    For iDoEventsLoop = 1 To 100                            'Process events waiting for send
        DoEvents                                            'Process windows events
    Next iDoEventsLoop

    '--- Reset thresholds -------------------------------------------------------------------
    comPORT1.RThreshold = 1                                 'Set receive threshold to 1 byte
    comPORT1.SThreshold = 1                                 'Set send threshold to 1 byte

    Exit Sub                                                'Exit routine

'--- Local error handler --------------------------------------------------------------------
ErrorHandler:
    
    MsgBox "Error sending command." & vbCr & _
            EZ_MSG_TECH_SUPPORT, _
            vbApplicationModal + vbInformation + vbOKOnly, EZ_CAPTION

    Resume Next

End Sub


'--------------------------------------------------------------------------------------------
'
' Procedure....:  drvPath_Change()
'
' Description..:  Update path when user changes drive.
'
'--------------------------------------------------------------------------------------------
Private Sub drvPath_Change()
    
    '--- Initialize local error handler -----------------------------------------------------
    On Error GoTo ErrorHandler
    
    dirPath.Path = drvPath.Drive

    Exit Sub

'--- Local error handler --------------------------------------------------------------------
ErrorHandler:

    MsgBox "Error changing drive." & vbCr & _
            "Please make sure selected drive is ready.", _
            vbInformation + vbApplicationModal + vbOKOnly, EZ_CAPTION
            
    Resume Next
    
End Sub

'--------------------------------------------------------------------------------------------
'
' Procedure....:  dirPath_Change()
'
' Description..:  Store the folder path in the database
'
'--------------------------------------------------------------------------------------------
Private Sub dirPath_Change()

    '--- Initialize local error handler -----------------------------------------------------
    On Error GoTo ErrorHandler
    
    '--- Save the Sotrage path in the database ----------------------------------------------
    DE.rscmdConfig("StoragePath").Value = dirPath.Path
    DE.rscmdConfig.UpdateBatch adAffectAllChapters

    Exit Sub
    
'--- Local error handler --------------------------------------------------------------------
ErrorHandler:

    MsgBox "Error saving storage path in database." & vbCr & _
            EZ_MSG_TECH_SUPPORT, _
            vbInformation + vbApplicationModal + vbOKOnly, EZ_CAPTION
    Resume Next
    
End Sub


Private Sub cmdOpen_Click()
    
    '--- Initialize local error handler -----------------------------------------------------
    On Error Resume Next
    
    '--- Open the serial port ---------------------------------------------------------------
    If (comPORT1.PortOpen = True) Then                      'If the port is already open, then close it.
        comPORT1.PortOpen = False                           'Close the serial port
    End If
    DoEvents
    
    comPORT1.CommPort = DE.rscmdConfig("Serial Port").Value 'Set the active serial port number
    comPORT1.Settings = "19200,n,8,1"                       'Set the serial port parameters
    comPORT1.PortOpen = True                                'Open the serial port
    comPORT1.InBufferCount = 0                              'Flush the input buffer
    comPORT1.RThreshold = 1                                 'Set the receive threshold

    DE.rscmdConfig.UpdateBatch adAffectAllChapters
    
End Sub

