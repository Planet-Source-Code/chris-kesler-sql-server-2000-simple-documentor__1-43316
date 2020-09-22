VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsToExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim xlApp As Excel.Application
Dim xlWb As Excel.Workbook

'*************************************************************
'Procedure:    Public Method ADOtoExcel
'Created on:   02/05/03
'Module:       clsToExcel
'Module File:  D:\Development\JobMonitor\clsToExcel.cls
'Project:      JobList
'Project File: D:\Development\JobMonitor\DTSList.vbp
'Parameters:
'rst
'Optional xlWsName
'*************************************************************
Public Function ADOtoExcel(rst As ADODB.Recordset, ByVal wksArray As Integer, Optional xlWsName As String)
    Dim recArray As Variant
    Dim fldCount As Integer
    Dim recCount As Long
    Dim iCol As Integer
    Dim iRow As Integer
    Dim xlWs As Excel.Worksheet
    Dim x As Integer
    
    If wksArray >= 4 Then
        xlWb.Worksheets.Add.Move after:=xlWb.Worksheets(xlWb.Worksheets.Count)
    End If
    'Create an instance of Excel and add a workbook
    Set xlWs = xlWb.Worksheets(wksArray) '("Sheet1")
  
    'Display Excel and give user control of Excel's lifetime
    
    'Copy field names to the first row of the worksheet
    fldCount = rst.Fields.Count
    For iCol = 1 To fldCount
        xlWs.Cells(1, iCol).Value = rst.Fields(iCol - 1).Name
        'Format Cell Headers...
        xlWs.Cells(1, iCol).Font.Bold = True
        xlWs.Cells(1, iCol).HorizontalAlignment = xlCenter
        xlWs.Cells(1, iCol).Interior.ColorIndex = 15
    Next
    'Set Sheet Name
    If Len(xlWsName) > 0 Then
        xlWs.Name = xlWsName
    Else
        xlWs.Name = "Job Info"
    End If
    'Check version of Excel
    If Val(Mid(xlApp.Version, 1, InStr(1, xlApp.Version, ".") - 1)) > 8 Then
        'EXCEL 2000 or 2002: Use CopyFromRecordset
         
        ' Copy the recordset to the worksheet, starting in cell A2
        xlWs.Cells(2, 1).CopyFromRecordset rst
        xlWs.Cells.Sort Key1:=xlWs.Range("A1"), Order1:=xlAscending, Header:=xlYes
        'Note: CopyFromRecordset will fail if the recordset
        'contains an OLE object field or array data such
        'as hierarchical recordsets
        
    Else
        'EXCEL 97 or earlier: Use GetRows then copy array to Excel
    
        ' Copy recordset to an array
        recArray = rst.GetRows
        'Note: GetRows returns a 0-based array where the first
        'dimension contains fields and the second dimension
        'contains records. We will transpose this array so that
        'the first dimension contains records, allowing the
        'data to appear properly when copied to Excel
        
        ' Determine number of records

        recCount = UBound(recArray, 2) + 1 '+ 1 since 0-based array

        ' Check the array for contents that are not valid when
        ' copying the array to an Excel worksheet
        For iCol = 0 To fldCount - 1
            For iRow = 0 To recCount - 1
                ' Take care of Date fields
                If IsDate(recArray(iCol, iRow)) Then
                    recArray(iCol, iRow) = Format(recArray(iCol, iRow))
                ' Take care of OLE object fields or array fields
                ElseIf IsArray(recArray(iCol, iRow)) Then
                    recArray(iCol, iRow) = "Array Field"
                End If
            Next iRow 'next record
        Next iCol 'next field
            
        ' Transpose and Copy the array to the worksheet,
        ' starting in cell A2
        xlWs.Cells(2, 1).Resize(recCount, fldCount).Value = _
            TransposeDim(recArray)
    End If
    ' Show Number of rows processed.
    xlWs.Cells(1, 1).Value = xlWs.Cells(1, 1).Value & " - Rows: " & rst.RecordCount
    ' Auto-fit the column widths and row heights
    xlApp.Selection.CurrentRegion.Columns.AutoFit
    xlApp.Selection.CurrentRegion.Rows.AutoFit
    ' Close ADO objects
    rst.Close
    Set rst = Nothing
    
    ' Release Excel references
    Set xlWs = Nothing

End Function

'*************************************************************
'Procedure:    Public Method TransposeDim
'Created on:   02/05/03
'Module:       clsToExcel
'Module File:  D:\Development\JobMonitor\clsToExcel.cls
'Project:      JobList
'Project File: D:\Development\JobMonitor\DTSList.vbp
'Parameters:
'v
'*************************************************************

Function TransposeDim(v As Variant) As Variant
' Custom Function to Transpose a 0-based array (v)
    
    Dim x As Long, Y As Long, Xupper As Long, Yupper As Long
    Dim tempArray As Variant
    
    Xupper = UBound(v, 2)
    Yupper = UBound(v, 1)
    
    ReDim tempArray(Xupper, Yupper)
    For x = 0 To Xupper
        For Y = 0 To Yupper
            tempArray(x, Y) = v(Y, x)
        Next Y
    Next x
    
    TransposeDim = tempArray

End Function

Private Sub Class_Initialize()
    
    Set xlApp = New Excel.Application
    Set xlWb = xlApp.Workbooks.Add

End Sub

Private Sub Class_Terminate()

    xlApp.Visible = True
    xlApp.UserControl = True
    Set xlWb = Nothing
    Set xlApp = Nothing

End Sub
