VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit


Private Sub BtDwnldPDF_Click()

If ShImport.CbID.Value = "< Select Student ID >" Then

    'Disable Download PDF button
    BtDwnldPDF.Enabled = False
    
    'Display Warning Message in ShImport
    With ShImport.Range("L20")
    .Value = "*Please Select a Student ID to Continue"
    .Font.Color = RGB(255, 0, 0)
    .Font.Bold = True
    End With

ElseIf ShPDF.Range("XEU1").Value = 0 Then

    'Disable Download PDF button
    BtDwnldPDF.Enabled = False
    
    'Display Warning Message in ShImport
    With ShImport.Range("L20")
    .Value = "*Please Import RAW DATA to Continue"
    .Font.Color = RGB(255, 0, 0)
    .Font.Bold = True
    End With

Else

    ShImport.Range("L20").ClearContents
    
    ' Call & Execute Download PDF Sub
    Call Optimize
    
    Call Print_PDF
    
    'Triggering ShImport to Activate
    ShPDF.Visible = xlSheetVisible
    ShPDF.Activate
    ShImport.Activate
    ShPDF.Visible = xlSheetHidden
    
    Call De_Optimize

End If

End Sub

Private Sub BtImport_Click()

Dim FileDir As FileDialog
Dim FilePath As String
Dim FileType As String
Dim Wbk As Workbook
Dim FileNum As Long
Dim ShRAW As Worksheet

Call Delete_Previous_RAWdata

Set FileDir = Application.FileDialog(msoFileDialogFilePicker)

'Toggling back Trigger for CbID
ShPDF.Range("XEU1").Value = 1

With FileDir

    .Title = "Select Report to Import"
    .AllowMultiSelect = False
        
    If .Show = 0 Then
    
        Exit Sub
    
    Else
        
        'Optimizing the code
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        
        'Defining Object variables for the File to be Selected by User
        FilePath = .SelectedItems(.SelectedItems.Count)
        FileType = Right(FilePath, Len(FilePath) - InStrRev(FilePath, "."))
        Set Wbk = Workbooks.Open(FilePath)
        
        'When Selected file is an Excel file (.xls)
        If FileType = "xlsx" Or FileType = "xls" Then
            
            Wbk.Worksheets(1).Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
            
            ThisWorkbook.Worksheets(ThisWorkbook.Sheets.Count).Name = "RAW DATA_Imported"
            
        
        'When Selected file is a txt file (Text/notepad)
        ElseIf FileType = "txt" Then
            
            With Wbk.Sheets(1)
            .Range("A1").CurrentRegion.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
            Semicolon:=False, Comma:=True, Space:=False, Other:=False, FieldInfo _
            :=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
            
            .Range("A1").EntireRow.Font.Bold = True
            .Name = "RAW DATA_Imported"
            .Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
            
            End With
        
        Else
            
            MsgBox Prompt:="Sorry, this File format is not supported. Please select an .xlsx (Excel) or .txt (Text) file", _
            Title:="File Format Warning"
            
            'De-optimizing the code before Exit
            Call De_Optimize
            
            Exit Sub
        
        End If
        
    'De-optimizing the code and Closing external file
    Application.ScreenUpdating = True
    Wbk.Close
    Application.DisplayAlerts = True
    
    End If

End With

'Defining ShRaw as Raw Data sheet
Set ShRAW = ThisWorkbook.Sheets("RAW DATA_Imported")

'Converting ShRaw Data set into a Table
ShRAW.ListObjects.Add(xlSrcRange, ShRAW _
.Range("A1").CurrentRegion, , xlYes).Name = "RAW_DATA"

ShImport.Activate

End Sub


Private Sub BtPreview_Click()

Dim PrevCell As Range

Set PrevCell = ShPDF.Range("A1")

'Clearing Alert Message
Me.Range("L20").ClearContents

'Unhiding ShPDF
ShPDF.Visible = xlSheetVisible

'Activating sheet and selecting A1 to preview
ShPDF.Activate
PrevCell.Select

'Disabling the Preview button
Me.BtPreview.Enabled = False

With ShPDF
.BtPrint.Enabled = True
.BtPrintCancel.Enabled = True
End With

ShPDF.Protect

End Sub

Private Sub Worksheet_Activate()

Dim T_RAW As ListObject
Dim C As Range
Dim i As Long

On Error GoTo ErrorHandle

'To Toggle for Values in ShPDF to not Reset
If ShPDF.Range("XFD1").Value = 1 Then

    Me.CbID.Text = "< Select Student ID >"
    
    Set T_RAW = ThisWorkbook.Sheets("RAW DATA_Imported").ListObjects("RAW_DATA")
    
    'Toggling back CbID change Trigger
    ShPDF.Range("XFD1").Value = 0
    
    'Optimizing Code
    Call Optimize
    
    'Copying Student ID column of raw data
    T_RAW.ListColumns(3).DataBodyRange.Copy ShPDF.Range("xfd54")
    
    'Removing duplicate values from copied List & Setting it up as a Named range
    ShPDF.Range("xfd54").CurrentRegion.RemoveDuplicates 1, xlNo
    Set C = ShPDF.Range("XFD54").CurrentRegion
    
    'Clearing Old values from combobox
    Me.CbID.Clear
    
    'Assigning Values to Combobox from Named Range 'C'
    For i = 1 To C.Rows.Count
    CbID.AddItem C(i)
    Next i
    
    'De-optimimzing Code
    Call De_Optimize

Else

Exit Sub

End If


ErrorHandle:
If Err.Number = 9 Then
Exit Sub
End If

End Sub


Private Sub CbID_Change()

Dim LastRow As Range 'for Fixed Signature'
Dim FixArea As Range
Dim T_Score As ListObject
Dim StartSig As Range 'for Dynamic Signature'
Dim Border1 As Range
Dim border2 As Range
Dim FullArea As Range
Dim cell As Range

    
    If ShPDF.Range("XFD1").Value = 0 Then
    
        'Populating fields in ShPDF format
        With ShPDF
        'Student ID
        .Range("C14").Value = Me.CbID.Value
        
        'Program
        .Range("G6").FormulaR1C1 = "=INDEX(RAW_DATA,MATCH('PDF Preview'!R14C3,RAW_DATA[Student ID],0),4)"
        
        'Year
        .Range("G7").FormulaR1C1 = "=INDEX(RAW_DATA,MATCH('PDF Preview'!R14C3,RAW_DATA[Student ID],0),1)"
    
        'Student Name
        .Range("C11").FormulaR1C1 = "=INDEX(RAW_DATA,MATCH('PDF Preview'!R14C3,RAW_DATA[Student ID],0),2)"
        
        'City
        .Range("G11").FormulaR1C1 = "=INDEX(RAW_DATA,MATCH('PDF Preview'!R14C3,RAW_DATA[Student ID],0),5)"
        
        'Institute
        .Range("G14").FormulaR1C1 = "=INDEX(RAW_DATA,MATCH('PDF Preview'!R14C3,RAW_DATA[Student ID],0),6)"
               
        'Institute Code
        .Range("K11").FormulaR1C1 = "=INDEX(RAW_DATA,MATCH('PDF Preview'!R14C3,RAW_DATA[Student ID],0),7)"
        
        Set LastRow = .Range("C27").EntireRow
        Set T_Score = .ListObjects("Marksheet_Scorecard")
        
        'Refreshing Scorecard Table
        T_Score.Refresh
         
         If T_Score.ListRows.Count < .Range("C17:C" & LastRow.Row).Count Then
         
         'Creating Signature/Footer section (Fixed)
         .Range("C29").Value = "Auth. Signatory : "
         .Range("C29").Font.Bold = True
         
         'Program Director field
         .Range("C31").FormulaR1C1 = "=INDEX(RAW_DATA,MATCH('PDF Preview'!R6C7,RAW_DATA[Course],0),11)"
         .Range("C31").Font.Bold = True
         .Range("C32").FormulaR1C1 = "=""Program Director, ""&R6C7"
         .Range("C32").Font.Italic = True
         
         'Vice Chancellor Sig.
         .Range("G31").Value = "Dr. XYZ"
         .Range("G31").Font.Bold = True
         .Range("G32").Value = "Vice Chancellor, AIIAS"
         .Range("G32").Font.Italic = True
         
         'PDF Background Colour
         Set FixArea = .Range("B2:M37")
         
         'Setting up Anchors for Fixed BG Colour Reset
         .Range("B" & FixArea.Rows.Count).Value = "a"
         .Range("B" & FixArea.Rows.Count).Font.Color = RGB(255, 255, 204)
         .Range("M" & FixArea.Rows.Count).Value = "b"
         .Range("M" & FixArea.Rows.Count).Font.Color = RGB(255, 255, 204)
         
         FixArea.Interior.Color = RGB(255, 255, 204)
         
         'For Putting Thick Border below cell
             With .Range("C32:E32").Borders(xlEdgeBottom)
             .LineStyle = xlContinuous
             .ColorIndex = 0
             .TintAndShade = 0
             .Weight = xlMedium
             End With
         
         'Generating Preview Options
             With ShImport
             .Range("L20").Value = "* Marksheet is Ready for Download."
             .Range("L20").Font.Color = RGB(0, 0, 255)
             .BtPreview.Enabled = True
             End With
         
         Else
         
         'Creating Signature/Footer section (Dynamic)
         Set StartSig = .Range("C" & T_Score.ListRows.Count + 2)
         
         StartSig.Value = "Auth. Signatory : "
         StartSig.Font.Bold = True
         
         'Program Director Sig.
         StartSig.Offset(2, 0).Value = "=INDEX(RAW_DATA,MATCH('PDF Preview'!R6C7,RAW_DATA[Course],0),11)"
         StartSig.Offset(2, 0).Font.Bold = True
         StartSig.Offset(3, 0).Value = "=""Program Director, ""&R6C7"
         StartSig.Offset(3, 0).Font.Italic = True
         
         'Vice Chancellor Sig.
         StartSig.Offset(2, 4).Value = "Dr. XYZ"
         StartSig.Offset(2, 4).Font.Bold = True
         StartSig.Offset(3, 4).Value = "Vice Chancellor, AIIAS"
         StartSig.Offset(3, 4).Font.Italic = True
         
         'PDF Background Colour
         Set Border1 = .Range("B2:M2")
         Set border2 = .Range(StartSig.Offset(9, -1).Row, StartSig.Offset(9, 10))
         Set FullArea = .Range(Border1, border2)
         
         'Setting up Anchors for Dynamic BG Colour Reset
         .Range("B" & FullArea.Rows.Count).Value = "a"
         .Range("B" & FullArea.Rows.Count).Font.Color = RGB(255, 255, 204)
         .Range("M" & FullArea.Rows.Count).Value = "b"
         .Range("M" & FullArea.Rows.Count).Font.Color = RGB(255, 255, 204)
         
         
             For Each cell In FullArea 'To Clear previous BG color
             cell.Interior.Color = xlNone
             Next cell
             
             For Each cell In FullArea 'To Apply BG Colour
             cell.Interior.Color = RGB(255, 255, 204)
             Next cell
         
         'For Putting Thick Border below cell
             With .Range("C" & StartSig.Offset(4, 0).Row, "E" & StartSig.Offset(4, 0).Row).Borders(xlEdgeBottom)
             .LineStyle = xlContinuous
             .ColorIndex = 0
             .TintAndShade = 0
             .Weight = xlMedium
             End With
         
        'Generating Preview Options
             With ShImport
             .Range("L20").Value = "* Marksheet is Ready for Download."
             .Range("L20").Font.Color = RGB(0, 0, 255)
             .BtPreview.Enabled = True
             End With
        
        End If
        
        End With
       
        'Enable Download PDF button
        BtDwnldPDF.Enabled = True

     Else
     
     Exit Sub
     
     End If


End Sub

Sub Test()

'ShPDF.Range("B2:M37").Interior.Color = RGB(255, 255, 204)

End Sub
