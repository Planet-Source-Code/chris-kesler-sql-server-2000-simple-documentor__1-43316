VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SQL Server 2000 Documentor"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   12525
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   12525
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtFileName 
      Height          =   345
      Left            =   2505
      TabIndex        =   9
      Top             =   1665
      Width           =   5025
   End
   Begin VB.CommandButton cmdJobsDoc 
      Caption         =   "Document Jobs"
      Enabled         =   0   'False
      Height          =   390
      Left            =   5040
      TabIndex        =   8
      Top             =   1230
      Width           =   2505
   End
   Begin VB.CommandButton cmdDTSPkgsDoc 
      Caption         =   "Document DTS Packages"
      Enabled         =   0   'False
      Height          =   390
      Left            =   5040
      TabIndex        =   6
      Top             =   825
      Width           =   2505
   End
   Begin VB.CommandButton cmdDocSPs 
      Caption         =   "Document Stored Procedures"
      Enabled         =   0   'False
      Height          =   390
      Left            =   2520
      TabIndex        =   7
      Top             =   1230
      Width           =   2505
   End
   Begin VB.Frame frmOptions 
      Caption         =   "Options"
      Height          =   2025
      Left            =   8295
      TabIndex        =   19
      Top             =   105
      Width           =   4035
      Begin VB.TextBox txtFilter 
         Height          =   330
         Left            =   120
         TabIndex        =   12
         Top             =   1590
         Width           =   3825
      End
      Begin VB.OptionButton opt 
         Caption         =   "Seperate Spreadsheets"
         Height          =   330
         Index           =   1
         Left            =   105
         TabIndex        =   11
         Top             =   855
         Width           =   2070
      End
      Begin VB.OptionButton opt 
         Caption         =   "All in one Spreadsheet"
         Height          =   330
         Index           =   0
         Left            =   105
         TabIndex        =   10
         Top             =   510
         Value           =   -1  'True
         Width           =   2070
      End
      Begin VB.Label lblFilter 
         Caption         =   "Filter for Specific Text in Name:"
         Height          =   225
         Left            =   120
         TabIndex        =   21
         Top             =   1380
         Width           =   2670
      End
      Begin VB.Label Label2 
         Caption         =   "Show Table Results:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   90
         TabIndex        =   20
         Top             =   300
         Width           =   2340
      End
   End
   Begin MSComctlLib.ProgressBar pgTables 
      Height          =   300
      Left            =   120
      TabIndex        =   17
      Top             =   2280
      Width           =   12270
      _ExtentX        =   21643
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdDocDB 
      Caption         =   "Document Chosen Database"
      Enabled         =   0   'False
      Height          =   390
      Left            =   2520
      TabIndex        =   5
      Top             =   825
      Width           =   2505
   End
   Begin VB.ComboBox cboDatabases 
      Height          =   315
      Left            =   30
      TabIndex        =   4
      Top             =   870
      Width           =   2430
   End
   Begin VB.CommandButton cmdConnect 
      Appearance      =   0  'Flat
      Caption         =   "Connect to Server"
      Height          =   330
      Left            =   6435
      TabIndex        =   3
      Top             =   255
      Width           =   1500
   End
   Begin VB.TextBox txtLogin 
      Height          =   330
      Left            =   2490
      TabIndex        =   1
      Text            =   "sa"
      ToolTipText     =   "Login ID for Server"
      Top             =   255
      Width           =   1950
   End
   Begin VB.TextBox txtPwd 
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   4470
      PasswordChar    =   "*"
      TabIndex        =   2
      Text            =   "k3nsingt0n"
      ToolTipText     =   "Password for Login to Server"
      Top             =   255
      Width           =   1950
   End
   Begin VB.ComboBox cboServers 
      Height          =   315
      Left            =   30
      TabIndex        =   0
      Top             =   270
      Width           =   2430
   End
   Begin MSComctlLib.ProgressBar pgColumns 
      Height          =   300
      Left            =   120
      TabIndex        =   18
      Top             =   2955
      Width           =   12270
      _ExtentX        =   21643
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Enter Output File Name:"
      Height          =   195
      Left            =   765
      TabIndex        =   25
      Top             =   1785
      Width           =   1695
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   ".xls"
      Height          =   195
      Left            =   7545
      TabIndex        =   24
      Top             =   1815
      Width           =   225
   End
   Begin VB.Label lblColProgress 
      AutoSize        =   -1  'True
      Caption         =   "Column Progress:"
      Height          =   195
      Left            =   120
      TabIndex        =   23
      Top             =   2745
      Width           =   1230
   End
   Begin VB.Label lblProcessing 
      AutoSize        =   -1  'True
      Caption         =   "Table Progress:"
      Height          =   195
      Left            =   150
      TabIndex        =   22
      Top             =   2070
      Width           =   1110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Database:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   30
      TabIndex        =   16
      Top             =   660
      Width           =   885
   End
   Begin VB.Label lblDTSSvrNAme 
      AutoSize        =   -1  'True
      Caption         =   "SQL Server Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   30
      TabIndex        =   15
      Top             =   45
      Width           =   1590
   End
   Begin VB.Label lblID 
      AutoSize        =   -1  'True
      Caption         =   "Login ID:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2490
      TabIndex        =   14
      Top             =   45
      Width           =   795
   End
   Begin VB.Label lblPassword 
      AutoSize        =   -1  'True
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4470
      TabIndex        =   13
      Top             =   45
      Width           =   885
   End
   Begin VB.Menu mnu_Exit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objApplication          As New SQLDMO.Application
Dim oDBase                      As SQLDMO.Database 'Database in Databases Collection
Dim oTable                      As SQLDMO.Table 'Table in Tables Collection
Dim oSP                         As SQLDMO.StoredProcedure
Dim oProp                       As SQLDMO.Column 'Column in Columns Collection
Dim oSQLServer                  As SQLDMO.SQLServer 'SQL Server connection
Dim oJobServer                  As SQLDMO.JobServer 'Job Server in SQL Server
Dim oJob                        As SQLDMO.Job 'Job in Jobs Collection
Dim oJobSteps                   As SQLDMO.JobSteps 'JobSteps Collection
Dim oStep                       As SQLDMO.JobStep 'Step in JobSteps Collection

Private Sub cmdConnect_Click()
    On Error GoTo LogErr
    
    If cboServers.Text = vbNullString Or cboServers.Text = "" Then
        MsgBox "Please choose a server to continue...", vbOKOnly + vbExclamation, "No Server Chosen"
        cboServers.SetFocus
        Exit Sub
    End If
    
    'If a Server was chosen then connect to the server and set all the proper
    'parameters dependent upon that connection.
    If cmdConnect.Caption = "Connect to Server" Then
        Set oSQLServer = New SQLDMO.SQLServer
        oSQLServer.Connect cboServers.Text, txtLogin.Text, txtPwd.Text
        For Each oDBase In oSQLServer.Databases
            If oDBase.SystemObject = False Then
                cboDatabases.AddItem oDBase.Name
            End If
        Next
        cboServers.Locked = True
        txtLogin.Locked = True
        txtPwd.Locked = True
        cmdDocDB.Enabled = True
        cmdDocSPs.Enabled = True
        cmdDTSPkgsDoc.Enabled = True
        cmdJobsDoc.Enabled = True
        cmdConnect.Caption = "Disconnect Server"
    'When we disconnect from the server we reset all our parameters appropriately.
    Else
        oSQLServer.Disconnect
        Set oSQLServer = Nothing
        cmdConnect.Caption = "Connect to Server"
        cboServers.Locked = False
        cmdDocDB.Enabled = False
        cmdDocSPs.Enabled = False
        cmdDTSPkgsDoc.Enabled = False
        cmdJobsDoc.Enabled = False
        txtLogin.Locked = False
        txtPwd.Locked = False
    End If
    
    Exit Sub
    
LogErr:
    
    If Err.Number = -2147203048 Then
        MsgBox "Please check your login and password and try again!", vbOKOnly + vbExclamation, "Login Failed"
        Set oSQLServer = Nothing
        Exit Sub
    Else
        MsgBox "Error Number: " & vbTab & Err.Number & vbCrLf _
             & "Error Description: " & vbTab & Err.Description & vbCrLf & vbCrLf _
             & "If this error us unknown contact your developer."
        Exit Sub
    End If

End Sub


Private Sub cmdDocDB_Click()
    Dim oADODB As otfADODB, rs As ADODB.Recordset
    Dim cExcel As clsToExcel
    Dim oCol As SQLDMO.Column
    Dim sTableHold As String, sTableInfo As String
    Dim I As Long, x As Long, z As Long
    
    If Len(cboDatabases.Text) = 0 Then MsgBox "No Database Chosen, choose a database to document and try again!", vbOKOnly, "No DB!": Exit Sub
    
    Set cExcel = New clsToExcel
    pgTables.Min = 0
    pgTables.Value = 0
    If opt(0).Value = True Then
        Set oADODB = New otfADODB
    End If
    For Each oDBase In oSQLServer.Databases
        If oDBase.Name = cboDatabases.Text Then
            pgTables.Max = oDBase.Tables.Count
            For Each oTable In oDBase.Tables
                DoEvents
                If opt(1).Value = True Then
                    Set oADODB = New otfADODB
                End If
                If opt(0).Value = True Then
                    oADODB.DefineRecordsetFields "TableName##100;FieldName##255;DataType##150;FieldLength##50;FieldType##3"
                Else
                    oADODB.DefineRecordsetFields "FieldName##255;DataType##150;FieldLength##50;FieldType##3"
                End If
                Set rs = New ADODB.Recordset
                x = 0
                lblProcessing.Caption = "Processing Table: " & oTable.Name
                pgTables.Value = I
                pgColumns.Min = 0
                pgColumns.Max = oTable.Columns.Count
                pgColumns.Value = 0
                If oTable.SystemObject = False Then
                    sTableHold = oTable.Name
                    If InStr(1, oTable.Name, txtFilter.Text) > 0 Then
                        For Each oCol In oTable.Columns
                            lblColProgress.Caption = "Processing Column: " & oCol.Name
                            DoEvents
                            If opt(1).Value = True Then
                                If oCol.InPrimaryKey Then
                                    sTableInfo = oCol.Name & ";" & oCol.DataType & ";" & oCol.Length & ";PK"
                                Else
                                    sTableInfo = oCol.Name & ";" & oCol.DataType & ";" & oCol.Length & ";Null"
                                End If
                            Else
                                If oCol.InPrimaryKey Then
                                    sTableInfo = sTableHold & ";" & oCol.Name & ";" & oCol.DataType & ";" & oCol.Length & ";PK"
                                Else
                                    sTableInfo = sTableHold & ";" & oCol.Name & ";" & oCol.DataType & ";" & oCol.Length & ";Null"
                                End If
                            End If
                            oADODB.AddRecord sTableInfo
                            x = x + 1
                            pgColumns.Value = x
                        Next
                        'Process a spreadsheet for each table.
                        z = z + 1
                        If opt(1).Value = True Then
                            Set rs = oADODB.Recordset
                            cExcel.ADOtoExcel rs, z, oTable.Name
                            Set rs = Nothing
                            Set oADODB = Nothing
                        End If
                    End If
                End If
                I = I + 1
            Next
        End If
    Next
    'Process to one spreadsheet
    If opt(0).Value = True Then
        Set rs = oADODB.Recordset
        cExcel.ADOtoExcel rs, 1, cboDatabases.Text
    End If
    If Len(txtFileName) = 0 Then
        MsgBox "No FileName Entered!  Can not Save File!", vbOKOnly, "No FileName!"
    Else
        cExcel.SaveExcelFileAs Trim(txtFileName.Text)
    End If
    pgColumns.Value = 0
    pgTables.Value = 0
    Set rs = Nothing
    Set cExcel = Nothing
    Set oADODB = Nothing
    Set oTable = Nothing
    Set oCol = Nothing
        
End Sub

Private Sub cmdDocSPs_Click()
    Dim oADODB As otfADODB, rs As ADODB.Recordset
    Dim cExcel As clsToExcel
    Dim oCol As SQLDMO.Column
    Dim oParms As QueryResults
    Dim oParm As String
    Dim sTableHold As String, sTableInfo As String
    Dim I As Long, x As Long, z As Long
    
    If Len(cboDatabases.Text) = 0 Then MsgBox "No Database Chosen, choose a database to document and try again!", vbOKOnly, "No DB!": Exit Sub
    
    Set cExcel = New clsToExcel
    pgTables.Min = 0
    pgTables.Value = 0
    Set oADODB = New otfADODB
    oADODB.DefineRecordsetFields "StoredProc##255;ParmInfo##255"
    For Each oDBase In oSQLServer.Databases
        DoEvents
        If oDBase.Name = cboDatabases.Text Then
            pgTables.Max = oDBase.StoredProcedures.Count
            For Each oSP In oDBase.StoredProcedures
                lblProcessing.Caption = "Processing Stored Proc: " & oSP.Name
                z = z + 1
                pgTables.Value = z
                DoEvents
                If oSP.SystemObject = False Then
                    If Len(txtFilter.Text) > 0 Then
                        If InStr(1, oSP.Name, txtFilter.Text) > 0 Then
                            Set oParms = oSP.EnumParameters
                            For x = 1 To oParms.Rows
                                DoEvents
                                I = 1
                                oParm = oParms.GetColumnString(x, I)
                                I = I + 1
                                For I = 2 To 3
                                    oParm = oParm & ";" & oParms.GetColumnString(x, I)
                                Next
                                oADODB.AddRecord oSP.Name & ";" & oParm
                            Next
                        End If
                    Else
                        Set oParms = oSP.EnumParameters
                        For x = 1 To oParms.Rows
                            DoEvents
                            I = 1
                            oParm = oParms.GetColumnString(x, I)
                            I = I + 1
                            For I = 2 To 3
                                oParm = oParm & ";" & oParms.GetColumnString(x, I)
                            Next
                            oADODB.AddRecord oSP.Name & ";" & oParm
                        Next
                    End If
                End If
            Next
        End If
    Next
    pgTables.Value = 0
    'Process to one spreadsheet
    Set rs = oADODB.Recordset
    cExcel.ADOtoExcel rs, 1, cboDatabases.Text
    If Len(txtFileName) = 0 Then
        MsgBox "No FileName Entered!  Can not Save File!", vbOKOnly, "No FileName!"
    Else
        cExcel.SaveExcelFileAs Trim(txtFileName.Text)
    End If
    pgColumns.Value = 0
    Set rs = Nothing
    Set oDBase = Nothing
    Set oParms = Nothing
    Set cExcel = Nothing
    Set oADODB = Nothing
    Set oTable = Nothing
    Set oCol = Nothing
End Sub

Private Sub Command2_Click()

End Sub

Private Sub cmdDTSPkgsDoc_Click()

    Dim myDTS           As DTS.Package
    Dim oGlobalVars     As DTS.GlobalVariable
    Dim oADODB          As otfADODB
    Dim cExcel          As clsToExcel
    Dim sGlobalVars     As String
    Dim sSQL            As String
    Dim cnn             As ADODB.Connection
    Dim rst             As ADODB.Recordset
    Dim rs              As ADODB.Recordset
    Dim I               As Long
    Dim z               As Long
    
    sSQL = "SELECT DISTINCT name FROM sysdtspackages"
    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    With cnn
        .ConnectionString = "Provider=SQLOLEDB;Data Source=" & cboServers.Text & "; " & _
                 "Initial Catalog=msdb; " & _
                 "User ID=" & txtLogin.Text & "; Password=" & txtPwd.Text
        .ConnectionTimeout = 30
        .Open
        'Execute our SQL Statment from above and return a recordset.
        Set rst = .Execute(sSQL)
    End With
    If opt(0).Value = True Then
        Set oADODB = New otfADODB
        oADODB.DefineRecordsetFields "PackageName##150;GlobalVariable##255;Value##255"
    End If
    Set cExcel = New clsToExcel
    Do Until rst.EOF
        If opt(1).Value Then
            Set oADODB = New otfADODB
            oADODB.DefineRecordsetFields "PackageName##150;GlobalVariable##255;Value##255"
        End If
        Set myDTS = New DTS.Package
        DoEvents
        If Len(txtFilter.Text) > 0 Then
            If InStr(1, rst("name").Value, txtFilter.Text) > 0 Then
                lblProcessing.Caption = "Processing: " & rst("name").Value
                myDTS.LoadFromSQLServer cboServers.Text, txtLogin.Text, txtPwd.Text, , , , , rst("name").Value
                'This error means the package requires an ID.VersionID to be opened.
                
                For Each oGlobalVars In myDTS.GlobalVariables
                    oADODB.AddRecord rst("name").Value & ";" & oGlobalVars.Name & ";" & Replace(oGlobalVars.Value, ";", ":")
                Next
                z = z + 1
                If opt(1).Value = True Then
                    Set rs = oADODB.Recordset
                    cExcel.ADOtoExcel rs, z, Left$(rst("name").Value, 25) & z
                    Set rs = Nothing
                    Set oADODB = Nothing
                End If
            End If
        Else
            lblProcessing.Caption = "Processing: " & rst("name").Value
            myDTS.LoadFromSQLServer cboServers.Text, txtLogin.Text, txtPwd.Text, , , , , rst("name").Value
            'This error means the package requires an ID.VersionID to be opened.
            
            For Each oGlobalVars In myDTS.GlobalVariables
                oADODB.AddRecord rst("name").Value & ";" & oGlobalVars.Name & ";" & oGlobalVars.Value
            Next
            z = z + 1
            If opt(1).Value = True Then
                Set rs = oADODB.Recordset
                cExcel.ADOtoExcel rs, z, Left$(rst("name").Value, 25) & z
                Set rs = Nothing
                Set oADODB = Nothing
            End If
        End If

        rst.MoveNext
        Set myDTS = Nothing
    Loop
    
    If opt(0).Value = True Then
        Set rs = oADODB.Recordset
        cExcel.ADOtoExcel rs, 1, "DTS Packages"
        Set rs = Nothing
    End If
    If Len(txtFileName) = 0 Then
        MsgBox "No FileName Entered!  Can not Save File!", vbOKOnly, "No FileName!"
    Else
        cExcel.SaveExcelFileAs Trim(txtFileName.Text)
    End If
    Set myDTS = Nothing
    Set cnn = Nothing
    Set rst = Nothing
    Set oADODB = Nothing
    Set cExcel = Nothing
    
End Sub

Private Sub cmdJobsDoc_Click()
    Dim oData As otfADODB
    Dim cExcel As clsToExcel
    Dim oJobSched As SQLDMO.JobSchedule
    Dim rs As ADODB.Recordset
    Dim I As Long, x As Long, sTimeHoldVal As String
    Dim sHH As String, sNN As String, sSS As String
    Dim sRunTime As String, sPlaceHolder As String
    Dim sFrequency As String, sFreqEnd As String
    'To retrieve Run History on Job
    Dim cnn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim sSQL As String

    Set cExcel = New clsToExcel
    Set oData = New otfADODB
    Set oSQLServer = New SQLDMO.SQLServer
    oSQLServer.Connect cboServers.Text, txtLogin.Text, txtPwd.Text
    oData.DefineRecordsetFields "JobName##255;JobSteps##3;StartTime##15;EndTime##15;Enabled##3;Frequency##255"
    Set rs = New ADODB.Recordset
    Set oJobServer = oSQLServer.JobServer
    pgTables.Min = 0
    pgTables.Max = oJobServer.Jobs.Count
    pgTables.Value = 0
    For Each oJob In oJobServer.Jobs
        lblProcessing.Caption = "Processing: " & oJob.Name
        pgTables.Value = I
        x = 0
        If InStr(1, oJob.Name, txtFilter.Text) > 0 Or Len(txtFilter.Text) = 0 Then
            If oJob.JobSchedules.Count > 0 Then
                pgColumns.Min = 0
                pgColumns.Max = oJob.JobSchedules.Count
                pgColumns.Value = 0
                For Each oJobSched In oJob.JobSchedules
                    lblColProgress.Caption = "Processing: " & oJobSched.Name
                    Select Case oJobSched.Schedule.FrequencyType
                        
                        Case SQLDMOFreq_Daily
                            sFrequency = "Daily"
                        Case SQLDMOFreq_Weekly
                            sFrequency = "Weekly"
                        Case SQLDMOFreq_Monthly
                            sFrequency = "Monthly"
                        Case Else
                            sFrequency = "Non-Elective"
                            
                    End Select
                    
                    Select Case oJobSched.Schedule.FrequencySubDay
                        
                        Case 0
                            sFrequency = sFrequency & " - Unknown"
                        Case 1
                            If oJobSched.Schedule.FrequencyInterval = 62 Then
                                sFrequency = sFrequency & " - Once Daily - Monday thru Friday"
                            ElseIf oJobSched.Schedule.FrequencyInterval = 124 Then
                                sFrequency = sFrequency & " - Once Daily - Tuesday thru Saturday"
                            ElseIf oJobSched.Schedule.FrequencyInterval = 127 Then
                                sFrequency = sFrequency & " - Once Every Day of the Week"
                            Else
                                sFrequency = sFrequency & " - Once"
                            End If
                        Case 4
                            sFrequency = sFrequency & " - Every " & oJobSched.Schedule.FrequencySubDayInterval & " Minute(s)"
                        Case 8
                            sFrequency = sFrequency & " - Every " & oJobSched.Schedule.FrequencySubDayInterval & " Hour(s)"
                        Case 13
                            sFrequency = sFrequency & " - Valid"
                            
                    End Select
                    
                    sPlaceHolder = oJobSched.Schedule.ActiveStartTimeOfDay
                    If Len(sPlaceHolder) = 5 Then
                        sTimeHoldVal = "0" & sPlaceHolder
                    ElseIf Len(sPlaceHolder) = 4 Then
                        sTimeHoldVal = "00" & sPlaceHolder
                    ElseIf Len(sPlaceHolder) = 3 Then
                        sTimeHoldVal = "000" & sPlaceHolder
                    ElseIf Len(sPlaceHolder) = 2 Then
                        sTimeHoldVal = "0000" & sPlaceHolder
                    ElseIf Len(sPlaceHolder) = 1 Then
                        sTimeHoldVal = "00000" & sPlaceHolder
                    ElseIf sPlaceHolder = 0 Then
                        sTimeHoldVal = "000000"
                    Else
                        sTimeHoldVal = sPlaceHolder
                    End If
                    sHH = Left(sTimeHoldVal, 2)
                    sNN = Mid(sTimeHoldVal, 3, 2)
                    sSS = Right(sTimeHoldVal, 2)
                    sRunTime = sHH & ":" & sNN & ":" & sSS
                    sRunTime = Format(sRunTime, "hh:mm:ss")
                    
                    sPlaceHolder = oJobSched.Schedule.ActiveEndTimeOfDay
                    If Len(sPlaceHolder) = 5 Then
                        sTimeHoldVal = "0" & sPlaceHolder
                    ElseIf Len(sPlaceHolder) = 4 Then
                        sTimeHoldVal = "00" & sPlaceHolder
                    ElseIf Len(sPlaceHolder) = 3 Then
                        sTimeHoldVal = "000" & sPlaceHolder
                    ElseIf Len(sPlaceHolder) = 2 Then
                        sTimeHoldVal = "0000" & sPlaceHolder
                    ElseIf Len(sPlaceHolder) = 1 Then
                        sTimeHoldVal = "00000" & sPlaceHolder
                    ElseIf sPlaceHolder = 0 Then
                        sTimeHoldVal = "000000"
                    Else
                        sTimeHoldVal = sPlaceHolder
                    End If
                    sHH = Left(sTimeHoldVal, 2)
                    sNN = Mid(sTimeHoldVal, 3, 2)
                    sSS = Right(sTimeHoldVal, 2)
                    sFreqEnd = sHH & ":" & sNN & ":" & sSS
                    sFreqEnd = Format(sFreqEnd, "hh:mm:ss")
                    pgColumns.Value = x
                    
                    If oJob.Enabled = True Then
                        oData.AddRecord oJob.Name & ";" & oJob.JobSteps.Count & ";" & sRunTime & ";" & sFreqEnd & ";Yes;" & sFrequency
                    Else
                        oData.AddRecord oJob.Name & ";" & oJob.JobSteps.Count & ";" & sRunTime & ";" & sFreqEnd & ";No;" & sFrequency
                    End If
                    x = x + 1
                Next
            End If
            If oJob.HasSchedule = False Then
                oData.AddRecord oJob.Name & ";" & oJob.JobSteps.Count & ";00:00:00;00:00:00;No;Never -  Has No Schedule"
            End If
        End If
        I = I + 1
    Next
    
    pgTables.Value = I
    pgColumns.Value = 0
    lblProcessing.Caption = "Done..."
    lblColProgress.Caption = "Done..."
    Set rs = oData.Recordset
    'Sort Column based on Col1 = A1, Col2 = B1, Col3 = C1.......  Always Ascending with Headers...
    cExcel.ADOtoExcel rs, 1, "Job Info"
    If Len(txtFileName) = 0 Then
        MsgBox "No FileName Entered!  Can not Save File!", vbOKOnly, "No FileName!"
    Else
        cExcel.SaveExcelFileAs Trim(txtFileName.Text)
    End If
    Set rs = Nothing
    oSQLServer.Disconnect
    Set oSQLServer = Nothing
    Set oJobServer = Nothing
    Set oJob = Nothing
    Set oJobSched = Nothing
    Set cExcel = Nothing
    Set oData = Nothing
    
End Sub

Private Sub Form_Load()
    
    Dim objServerGroup As SQLDMO.ServerGroup
    Dim objRegisteredServer As SQLDMO.RegisteredServer
    Dim I As Integer, j As Integer
    'When the form loads get a list of all SQL Server Groups.
    For Each objServerGroup In objApplication.ServerGroups
        'When the form loads get a list of all registered SQL Servers in those Groups.
        For Each objRegisteredServer In objServerGroup.RegisteredServers
            'Add them to the drop down combo box
            cboServers.AddItem objRegisteredServer.Name
            cboServers.ItemData(cboServers.NewIndex) = CStr(objRegisteredServer.UseTrustedConnection)
        Next objRegisteredServer
    Next objServerGroup
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set oDBase = Nothing                        'Database in Databases Collection
    Set oTable = Nothing                        'Table in Tables Collection
    Set oSP = Nothing
    Set oProp = Nothing                         'Column in Columns Collection
    Set oSQLServer = Nothing                    'SQL Server connection
    Set oJobServer = Nothing                    'Job Server in SQL Server
    Set oJob = Nothing                          'Job in Jobs Collection
    Set oJobSteps = Nothing                     'JobSteps Collection
    Set oStep = Nothing                         'Step in JobSteps Collection
    
End Sub

Private Sub mnu_Exit_Click()
    
    Unload Me
    
End Sub
