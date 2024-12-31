Attribute VB_Name = "basMain"
'--------------------------------------------------------------------------------------------
'
' Application..:  RSC EZSNAP(r) Link
'
' Module.......:  basMain - Application main module.
'
' Procedure....:  (General)(declarations)
'
' Description..:  Declare global variables and DLL Calls.
'
' History......:
'                 07-24-96 RDR Designed and Programmed.
'
' (c) 1996-2000 Redmer Software Company.  All Rights Reserved.
'
'--------------------------------------------------------------------------------------------
Option Explicit                                             'Require Explicit Variable Declaration

'--- Dimension global constants -------------------------------------------------------------
Global Const EZ_CAPTION As String = "RSC EZ-Link"           'Standard MSGBOX window caption
Global Const EZ_APPDATA As String = "EZLINK.MDB"            'The name of the application database
Global Const EZ_APPCONFIG As String = "tbl_Configuration"   'The name of the configuration table
Global Const EZ_APPLOG As String = "tbl_Log_Info"           'The name of the log info table
Global Const EZ_APPPORTS As String = "tbl_Serial_Ports"     'The name of the serial port list table
Global Const EZ_MSG_TECH_SUPPORT As String = "Please contact technical support."
Global Const EZ_CONSTRING = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="

'--- Dimension global variables -------------------------------------------------------------
Global dbApp As ADODB.Connection                            'Application database
Global bTerminalMode As Boolean                             'Mode of operation, true if not receiving data
Global iSerialPort As Integer                               'Serial port number

'--------------------------------------------------------------------------------------------
'
' Procedure....:  Main()
'
' Description..:  Initialize pointers to global database object and call main form.
'
'--------------------------------------------------------------------------------------------
Public Sub Main()
    
    '--- Dimension local variables ----------------------------------------------------------
    Dim sTmpDbName As String                            'Temporary application database name
    Dim sAppDbName As String                            'Name of application database
    
    '--- Initialize local error handler -----------------------------------------------------
    On Error GoTo ErrorHandler                          'Local error handler
    
    '--- Construct path to Symbol database and verify the file exists -----------------------
    If (Right(Trim$(App.Path), 1) = "\") Then           'Check for trailing backslash
        sAppDbName = Trim$(App.Path) + EZ_APPDATA       'If slash present, append db name
    Else                                                'Else no slash in path
        sAppDbName = Trim$(App.Path) + "\" + EZ_APPDATA 'Add slash and db name
    End If
    
    '--- Make sure the application database exists ------------------------------------------
    sTmpDbName = Dir$(sAppDbName, vbNormal)             'Look for the application database
    If (UCase$(sTmpDbName) <> EZ_APPDATA) Then          'Make sure the file exists
        MsgBox "Application data not found." & vbCr & "Please make sure " & sAppDbName & " is in " & App.Path & ".", _
                vbApplicationModal + vbOKOnly + vbExclamation, EZ_CAPTION
        End                                             'End the program
    End If
        
    DE.cnnMain.ConnectionString = EZ_CONSTRING & sAppDbName
    DE.cnnMain.Open
    frmMain.Show                                        'Show the main application form
    Exit Sub                                            'Exit this routine.
    
'--- Local Error Handler --------------------------------------------------------------------
ErrorHandler:

    MsgBox "Error initializing program." & vbCr & EZ_MSG_TECH_SUPPORT, _
            vbApplicationModal + vbOKOnly + vbExclamation, EZ_CAPTION
    End
    
End Sub
