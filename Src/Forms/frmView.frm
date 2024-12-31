VERSION 5.00
Begin VB.Form frmView 
   Caption         =   "View File"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6975
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   6975
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnView 
      Caption         =   "&Save"
      Height          =   465
      Index           =   1
      Left            =   1755
      TabIndex        =   2
      Top             =   4170
      Width           =   1575
   End
   Begin VB.CommandButton btnView 
      Caption         =   "&Cancel"
      Height          =   465
      Index           =   0
      Left            =   75
      TabIndex        =   1
      Top             =   4155
      Width           =   1575
   End
   Begin VB.TextBox txtView 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4050
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   30
      Width           =   6825
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------------------
'
' Procedure....:  btnView_Click(Index As Integer)
'
' Description..:  Set data control recordsets.
'
'--------------------------------------------------------------------------------------------
Private Sub btnView_Click(Index As Integer)
    
    '--- Dimension local variables ----------------------------------------------------------
    Dim iAns As Integer                                         'Response to delete confirmation
    Dim iOutputFile As Integer                                  'Output file

    '--- Initialize local error handler -----------------------------------------------------
    On Error GoTo ErrorHandler
    
    iOutputFile = FreeFile                                      'Get a free handle for output file
    
    '--- Select button pressed --------------------------------------------------------------
    Select Case Index
    
        '--- Cancel -------------------------------------------------------------------------
        Case 0:
    
            Unload Me
        
        '--- Save ---------------------------------------------------------------------------
        Case 1:
    
            '--- Query user to save file ----------------------------------------------------
            iAns = MsgBox("Save changes?", vbApplicationModal + vbYesNo + vbQuestion, EZ_CAPTION)
            If iAns = vbYes Then
                
                '--- Update the archive record ----------------------------------------------
                DE.rscmdLogFile("Archive").Value = txtView.Text
                DE.rscmdLogFile.UpdateBatch adAffectCurrent
                
                '--- Delete the associated text file if it exists ---------------------------
                If Len(Dir$(Trim$(DE.rscmdLogFile("FileName") & " "), vbNormal)) > 0 Then
                    Kill Trim$(DE.rscmdLogFile("FileName"))
                End If
                
                '--- Write the associated text file -----------------------------------------
                Open Trim$(DE.rscmdLogFile("FileName") & " ") For Output As #iOutputFile
                Print #iOutputFile, txtView.Text
                Close #iOutputFile
                
            End If
            
            Unload Me
            
    End Select
    
    
    Exit Sub                                        'Exit routine
    
'--- Local error handler --------------------------------------------------------------------
ErrorHandler:

    MsgBox "Error writing: " & Trim$(DE.rscmdLogFile("FileName") & " ") & vbCr & "Check path name.", _
            vbApplicationModal + vbInformation + vbOKOnly, EZ_CAPTION

    Resume Next

End Sub

'--------------------------------------------------------------------------------------------
'
' Procedure....:  Form_Load()
'
' Description..:  Load text field from database into text control.
'
'--------------------------------------------------------------------------------------------
Private Sub Form_Load()

    '--- Initialize local error handler -----------------------------------------------------
    On Error GoTo ErrorHandler
    
    txtView.Text = DE.rscmdLogFile("Archive")

    Exit Sub                                        'Exit routine
    
'--- Local error handler --------------------------------------------------------------------
ErrorHandler:

    MsgBox "Error retrieving text file.", _
            vbApplicationModal + vbInformation + vbOKOnly, EZ_CAPTION

    Resume Next

End Sub

