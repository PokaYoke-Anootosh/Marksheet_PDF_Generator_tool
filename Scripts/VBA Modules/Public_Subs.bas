Attribute VB_Name = "Public_Subs"
Option Explicit


Sub Optimize()

Application.ScreenUpdating = False
Application.DisplayAlerts = False

End Sub


Sub De_Optimize()

Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub

Sub Optimize_with_Calc()

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual

End Sub

Sub De_Optimize_with_Calc()

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationAutomatic

End Sub


Sub Print_PDF()

Dim StartRng As Range
Dim LastRow As Long
Dim PDFArea As Range
Dim DefFileName As String
Dim DefFilePath As String
Dim RAWRetn As VbMsgBoxResult

'Unhiding Sheet
ShPDF.Visible = xlSheetVisible

'Defining Variables
With ShPDF

Set StartRng = .Range("B2:M2")
LastRow = .Range("B" & Rows.Count).End(xlUp).Row
Set PDFArea = .Range(StartRng, .Range("M" & LastRow))
DefFileName = .Range("G7") & "_" & .Range("K11") & "-" & .Range("C14") & "_" & .Range("C11") & ".pdf"

End With

'Showing Folder picker for User
With Application.FileDialog(msoFileDialogFolderPicker)

.Title = "Select Folder to Save Marksheet PDF"
.ButtonName = "Select"

    If .Show = -1 Then 'If Folder is selected by user
        
        DefFilePath = .SelectedItems(1) & "/" & DefFileName
    
        'Setting up Print PDF properties
        Application.PrintCommunication = False
            
            With ShPDF.PageSetup
                .LeftHeader = ""
                .CenterHeader = ""
                .RightHeader = ""
                .LeftFooter = ""
                .CenterFooter = ""
                .RightFooter = ""
                .LeftMargin = Application.InchesToPoints(0.25)
                .RightMargin = Application.InchesToPoints(0.25)
                .TopMargin = Application.InchesToPoints(0.75)
                .BottomMargin = Application.InchesToPoints(0.75)
                .HeaderMargin = Application.InchesToPoints(0.3)
                .FooterMargin = Application.InchesToPoints(0.3)
                .PrintHeadings = False
                .PrintGridlines = False
                .PrintComments = xlPrintNoComments
                .PrintQuality = 600
                .CenterHorizontally = False
                .CenterVertically = False
                .Orientation = xlPortrait
                .Draft = False
                .PaperSize = xlPaperA4
                .FirstPageNumber = xlAutomatic
                .Order = xlDownThenOver
                .BlackAndWhite = False
                .Zoom = False
                .FitToPagesWide = 1
                .FitToPagesTall = 1
                .PrintErrors = xlPrintErrorsBlank
                .OddAndEvenPagesHeaderFooter = False
                .DifferentFirstPageHeaderFooter = False
                .ScaleWithDocHeaderFooter = True
                .AlignMarginsHeaderFooter = True
                .EvenPage.LeftHeader.Text = ""
                .EvenPage.CenterHeader.Text = ""
                .EvenPage.RightHeader.Text = ""
                .EvenPage.LeftFooter.Text = ""
                .EvenPage.CenterFooter.Text = ""
                .EvenPage.RightFooter.Text = ""
                .FirstPage.LeftHeader.Text = ""
                .FirstPage.CenterHeader.Text = ""
                .FirstPage.RightHeader.Text = ""
                .FirstPage.LeftFooter.Text = ""
                .FirstPage.CenterFooter.Text = ""
                .FirstPage.RightFooter.Text = ""
            End With
            
        Application.PrintCommunication = True
        
        'Saving the PDF with default filename in Selected Folder
        PDFArea.PrintOut Copies:=1, Collate:=True, prtofilename:=DefFilePath
        
        'Unhiding Raw data sheet to be able to Hide ShPDF
        Sheets("RAW DATA_Imported").Visible = xlSheetVisible
        
        'Hiding Sheet, Toggling back Dropdown trigger & Unprotecting Sheet - ShPDF
        With ShPDF
        .Visible = xlSheetHidden
        .Unprotect
        .Range("XFD1") = 1
        End With
        
        'Unhiding ShImport
        ShImport.Visible = xlSheetVisible
        ShImport.Activate
        
        'Disabling Preview button and Clearinig 'Read-for-download' Message text
        With ShImport
        .Range("L20").ClearContents
        .BtPreview.Enabled = False
        End With
        
        'Success Message & Raw Data Retention offer
        RAWRetn = MsgBox("Your PDF has been saved succesfully." & vbNewLine & "Retain RAW DATA to Download Next Marksheet?", _
        vbYesNo, "Success!")

        If RAWRetn = vbYes Then
        
            Exit Sub
            
        ElseIf RAWRetn = vbNo Then
        
            Application.DisplayAlerts = False
            
            ThisWorkbook.Sheets("RAW DATA_Imported").Delete
            
            Application.DisplayAlerts = True
            
            With ShPDF
            .Range("XEU1").Value = 0 'Toggling Trigger for Download Button (BtDwnld)
            .Range("XFD1").Value = 1 'Toggling Trigger for CbID combobox
            End With
'            ShImport.CbID.Value = "< Select Student ID >"
        
        End If
    
    Else ' If user Cancels the Save operation
       
       Application.PrintCommunication = True

        'Hiding Sheet, Toggling back Dropdown trigger & Unprotecting Sheet - ShPDF
        With ShPDF
        .Visible = xlSheetHidden
        .Unprotect
        .Range("XFD1") = 1
        End With
        
        'Disabling Preview button and Clearinig 'Read-for-download' Message text
        With ShImport
        .Range("L20").ClearContents
        .BtPreview.Enabled = False
        End With
        
        MsgBox Prompt:="Your PDF was not Saved", Title:="Thank you"
    
    End If
    
End With

End Sub

Sub Delete_Previous_RAWdata()

Dim PrevSheet As Worksheet

On Error Resume Next

Call Optimize

Set PrevSheet = ThisWorkbook.Sheets("RAW DATA_Imported")

'Check for Previous RAW DATA sheet and Delete it
If Not PrevSheet Is Nothing Then

PrevSheet.Delete

End If

Call De_Optimize

End Sub
