VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "otfADODB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Recordset As Object
Dim RSDefined As Boolean

'*********************************************************************************************
'This method defines the recordset fields
'Inputs : Parameter Fields
'The Parameter "Fields" has to be like this :
'        "NAME OF THE FIELD 1##SIZE OF THE FIELD 1;NAME OF THE FIELD 2##SIZE OF THE FIELD 2"
'        For example : "Name##20;Phone Number##10;Address##40"
'Returns : Boolean
'Returns TRUE if there were no errors in the function execution
'Returns FALSE if there were errors in the function execution
'*********************************************************************************************
Public Function DefineRecordsetFields(Fields As String) As Boolean
On Error GoTo DEFREC_ERR

Dim Arr_Campos() As String
Dim Arr_NomeTipo() As String
Dim I As Integer

Recordset.CursorLocation = 3

Arr_Campos = Split(Fields, ";")

For I = LBound(Arr_Campos()) To UBound(Arr_Campos())
    Arr_NomeTipo = Split(Arr_Campos(I), "##")
    Recordset.Fields.Append Arr_NomeTipo(0), 200, Val(Arr_NomeTipo(1))
Next
Recordset.Open

DefineRecordsetFields = True
RSDefined = True

Exit Function

'.............................................................
DEFREC_ERR:
DefineRecordsetFields = False

End Function

'*********************************************************************************************
'This function adds a new record to the recordset
'Inputs : Parameter Fields
'The Parameter "Fields" has to be like this :
'        "VALUE FOR THE FIELD 1;VALUE FOR THE FIELD 2"
'        For example : "John Doe;8749830"
'Returns : Integer
'Returns 0 if there were errors in the function execution
'Returns -1 if the value of a field is larger than expected
'Returns -2 for any other error
'*********************************************************************************************
Public Function AddRecord(Fields As String) As Integer
On Error GoTo ADDRCD_ERR

Dim Arr_Campos() As String
Dim I As Integer

Arr_Campos = Split(Fields, ";")

Recordset.AddNew
For I = 0 To Recordset.Fields.Count - 1
    Recordset.Fields(I).Value = Arr_Campos(I)
Next
Recordset.Update

AddRecord = 0
Exit Function

'......................................................
ADDRCD_ERR:
If Err = -2147217887 Then
    AddRecord = -1
Else
    AddRecord = -2
End If
End Function
Private Sub Class_Initialize()
    Set Recordset = CreateObject("ADODB.Recordset")
    RSDefined = False
End Sub
Private Sub Class_Terminate()
    Set Recordset = Nothing
End Sub

