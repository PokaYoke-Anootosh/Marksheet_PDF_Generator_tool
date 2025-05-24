VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShPDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub BtPrint_Click()

Call Optimize

Call Print_PDF

'Triggering ShImport to Activate
ShPDF.Visible = xlSheetVisible
ShPDF.Activate
ShImport.Activate
ShPDF.Visible = xlSheetHidden

Call De_Optimize

End Sub


Private Sub BtPrintCancel_Click()

'Unprotecting ShPDF and Hiding it
With ShPDF
.Unprotect
.Visible = xlSheetHidden
End With

'Going back to ShImport
With ShImport
.Activate
.Range("A1").Select
End With

End Sub


Private Sub Worksheet_Activate()

Me.Range("A1").Select

End Sub

