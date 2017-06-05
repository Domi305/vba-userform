Option Explicit
Private Sub cmdAdd_Click()
Dim findvalue As Range
Dim cNum As Integer
Dim wks As Worksheet

Application.ScreenUpdating = False

Set wks = Sheet1
'wks.Range("P:P").Font.Bold = False
    If txtNotes.Value = "" Then
    MsgBox "Please fill the notes window"
    Exit Sub
End If

Set findvalue = wks.Range("A:A"). _
    Find(What:=Me.cboCust.Value, LookIn:=xlValues, LookAt:=xlWhole)
'update the values
findvalue = cboCust.Value
findvalue.Offset(0, 1) = txtStore.Value
findvalue.Offset(0, 2) = txtItem1.Value
findvalue.Offset(0, 3) = txtItem2.Value
findvalue.Offset(0, 4) = txtItem3.Value
findvalue.Offset(0, 5) = txtNotes.Value

'copy the data
    If wks.Range("A3").Value = "" Then
    cboCust.RowSource = wks.Range("Table1").Address(external:=False)
    Else: MsgBox "Fields updated/added"
    cboCust.SetFocus
End If

Application.ScreenUpdating = True

End Sub

Private Sub cmdUndo_Click()
UndoAction
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdHide_Click()
frmAdd.Hide
End Sub

Private Sub cboCust_change()
popTextboxes
End Sub

Private Sub UserForm_Initialize()
cboCust.SetFocus
txtNotes.ScrollBars = fmScrollBarsVertical
'fills combobox with list of agents from sheet1
Dim ListRange As Range, cl As Range
Set ListRange = Range("Customers")
        For Each cl In ListRange
        cboCust.AddItem cl.Value
    Next cl

popTextboxes
End Sub

'module1
Option Explicit
Sub popTextboxes()

On Error Resume Next

With frmAdd
'fill text boxes with data from sheet1

'.cboCust = Application.WorksheetFunction.VLookup(frmAdd.cboCust, Sheet1.Range("Table1"), 1, 0)
.txtStore = Application.WorksheetFunction.VLookup(frmAdd.cboCust, Sheet1.Range("Table1"), 2, 0)
.txtItem1 = Application.WorksheetFunction.VLookup(frmAdd.cboCust, Sheet1.Range("Table1"), 3, 0)
.txtItem2 = Application.WorksheetFunction.VLookup(frmAdd.cboCust, Sheet1.Range("Table1"), 4, 0)
.txtItem3 = Application.WorksheetFunction.VLookup(frmAdd.cboCust, Sheet1.Range("Table1"), 5, 0)
.txtNotes = Application.WorksheetFunction.VLookup(frmAdd.cboCust, Sheet1.Range("Table1"), 6, 0)
End With
End Sub
