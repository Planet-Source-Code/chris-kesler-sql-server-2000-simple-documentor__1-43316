VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   Caption         =   "SQL Server 2000 Documentor"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12525
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   12525
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDocSPs 
      Caption         =   "Document Stored Procedures"
      Height          =   390
      Left            =   2520
      TabIndex        =   16
      Top             =   1230
      Width           =   2505
   End
   Begin VB.Frame frmOptions 
      Caption         =   "Options"
      Height          =   2025
      Left            =   8295
      TabIndex        =   12
      Top             =   105
      Width           =   4035
      Begin VB.TextBox txtFilter 
         Height          =   330
         Left            =   120
         TabIndex        =   17
         Top             =   1590
         Width           =   3825
      End
      Begin VB.OptionButton opt 
         Caption         =   "Seperate Spreadsheets"
         Height          =   330
         Index           =   1
         Left            =   105
         TabIndex        =   14
         Top             =   855
         Width           =   2070
      End
      Begin VB.OptionButton opt 
         Caption         =   "All in one Spreadsheet"
         Height          =   330
         Index           =   0
         Left            =   105
         TabIndex        =   13
         Top             =   510
         Width           =   2070
      End
      Begin VB.Label lblFilter 
         Caption         =   "Filter for Specific Text in Name:"
         Height          =   225
         Left            =   120
         TabIndex        =   18
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
         TabIndex        =   15
         Top             =   300
         Width           =   2340
      End
   End
   Begin MSComctlLib.ProgressBar pgTables 
      Height          =   300
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Width           =   12270
      _ExtentX        =   21643
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdDocDB 
      Caption         =   "Document Chosen Database"
      Height          =   390
      Left            =   2520
      TabIndex        =   9
      Top             =   825
      Width           =   2505
   End
   Begin VB.ComboBox cboDatabases 
      Height          =   315
      Left            =   30
      TabIndex        =   7
      Top             =   870
      Width           =   2430
   End
   Begin VB.CommandButton cmdConnect 
      Appearance      =   0  'Flat
      Caption         =   "Connect to Server"
      Height          =   330
      Left            =   6435
      TabIndex        =   6
      Top             =   255
      Width           =   1500
   End
   Begin VB.TextBox txtLogin 
      Height          =   330
      Left            =   2490
      TabIndex        =   2
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
      TabIndex        =   1
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
      TabIndex        =   11
      Top             =   2955
      Width           =   12270
      _ExtentX        =   21643
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblColProgress 
      AutoSize        =   -1  'True
      Caption         =   "Column Progress:"
      Height          =   195
      Left            =   120
      TabIndex        =   20
      Top             =   2745
      Width           =   1230
   End
   Begin VB.Label lblProcessing 
      AutoSize        =   -1  'True
      Caption         =   "Table Progress:"
      Height          =   195
      Left            =   150
      TabIndex        =   19
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
      TabIndex        =   8
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
      TabIndex        =   5
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
      TabIndex        =   4
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
      TabIndex        =   3
      Top             =   45
      Width           =   885
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
        cmdConnect.Caption = "Disconnect Server"
    'When we disconnect from the server we reset all our parameters appropriately.
    Else
        oSQLServer.Disconnect
        Set oSQLServer = Nothing
        cmdConnect.Caption = "Connect to Server"
        cboServers.Locked = False
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
    Dim i As Long, x As Long, z As Long
    
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
                pgTables.Value = i
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
                i = i + 1
            Next
        End If
    Next
    'Process to one spreadsheet
    If opt(0).Value = True Then
        Set rs = oADODB.Recordset
        cExcel.ADOtoExcel rs, 1, cboDatabases.Text
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
    Dim i As Long, x As Long, z As Long
    
    Set cExcel = New clsToExcel
    pgTables.Min = 0
    pgTables.Value = 0
    Set oADODB = New otfADODB
    oADODB.DefineRecordsetFields "StoredProc##255;ParmName##255;ParmType##50;ParmSize##10"
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
                                i = 1
                                oParm = oParms.GetColumnString(x, i)
                                i = i + 1
                                For i = 2 To 3
                                    oParm = oParm & ";" & oParms.GetColumnString(x, i)
                                Next
                                oADODB.AddRecord oSP.Name & ";" & oParm
                            Next
                        End If
                    Else
                        Set oParms = oSP.EnumParameters
                        For x = 1 To oParms.Rows
                            DoEvents
                            i = 1
                            oParm = oParms.GetColumnString(x, i)
                            i = i + 1
                            For i = 2 To 3
                                oParm = oParm & ";" & oParms.GetColumnString(x, i)
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
    pgColumns.Value = 0
    Set rs = Nothing
    Set oDBase = Nothing
    Set oParms = Nothing
    Set cExcel = Nothing
    Set oADODB = Nothing
    Set oTable = Nothing
    Set oCol = Nothing
End Sub

Private Sub Form_Load()
    
    Dim objServerGroup As SQLDMO.ServerGroup
    Dim objRegisteredServer As SQLDMO.RegisteredServer
    Dim i As Integer, j As Integer
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
