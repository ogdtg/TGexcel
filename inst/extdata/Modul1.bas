Attribute VB_Name = "Modul1"
Sub AddFooterHeaderImage(imagePath as String)
'PURPOSE: Insert Image File into Spreadsheet Header or Footer on every selected worksheet
'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

'Dim Fname As String
Dim sht As Worksheet
Dim headerOne as String
'Dim ImagePath As String
Dim Validation As String
Dim pct as Double

pct = 0.75    '75% of original size
headerOne = "Staatskanzlei"


'Where is Image Located?
  'ImagePath = "Y:\SK\SKStat\Internet\media\image1.png"
  'Fname = ThisWorkbook.FullName
  
'Does the Image File Exist?
  On Error Resume Next
    Validation = Dir(ImagePath)
  On Error GoTo 0

  If Validation = "" Then
    MsgBox "Could not locate the image file located here: " & ImagePath
    Exit Sub
  End If
      
'Add Image To Each Active Sheet
  For Each sht In ThisWorkbook.Sheets

      sht.PageSetup.LeftHeader= "&" & Chr(34) & "Arial" & Chr(34) & "&B&10" & headerOne & Chr(13) & "&B&10" & "Dienststelle " & "f" & ChrW(&H00FC) & "r Statistik"
    
    'Insert Image (ie "LeftFooter","CenterFooter","RightFooter", _
                      "LeftHeader","CenterHeader","RightHeader")
      sht.PageSetup.RightHeader = "&G"
      sht.PageSetup.RightHeaderPicture.Filename = ImagePath
      sht.PageSetup.RightHeaderPicture.Height = sht.PageSetup.RightHeaderPicture.Height * pct
      sht.PageSetup.RightHeaderPicture.Width = sht.PageSetup.RightHeaderPicture.Width * pct
      
    'Ensure Pagebreaks don't show (they are annoying!)
      sht.DisplayPageBreaks = False
  
  Next sht

'ActiveWorkbook.SaveAs Filename:=Fname, FileFormat:=xlWorkbookDefault

End Sub

