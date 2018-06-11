Sub SplitMergedDocument()
' Sourced from: http://www.msofficeforums.com/mail-merge/21803-mailmerge-tips-tricks.html
Application.ScreenUpdating = False
Dim i As Long, j As Long, k As Long, StrTxt As String
Dim Rng As Range, Doc As Document, HdFt As HeaderFooter
Const StrNoChr As String = """*./\:?|"
j = InputBox("How many Section breaks are there per record?", "Split By Sections", 1)
With ActiveDocument
   ' Process each Section
  For i = 1 To .Sections.Count - 1 Step j
    With .Sections(i)
       '*****
       ' Get the 1st paragraph
      StrTxt = Split(.Range.Paragraphs(1).Range.Text, vbCr)(0)
      ' Strip out illegal characters
      For k = 1 To Len(StrNoChr)
        StrTxt = Replace(StrTxt, Mid(StrNoChr, k, 1), "_")
      Next
       ' Construct the destination file path & name
      StrTxt = ActiveDocument.Path & Application.PathSeparator & StrTxt
       '*****
       ' Get the whole Section
      Set Rng = .Range
      With Rng
        If j > 1 Then .MoveEnd wdSection, j - 1
         'Contract the range to exclude the Section break
        .MoveEnd wdCharacter, -1
         ' Copy the range
        .Copy
      End With
    End With
     ' Create the output document
    Set Doc = Documents.Add(Template:=ActiveDocument.AttachedTemplate.FullName, Visible:=False)
    With Doc
       ' Paste contents into the output document, preserving the formatting
      .Range.PasteAndFormat (wdFormatOriginalFormatting)
       ' Delete trailing paragraph breaks & page breaks at the end
      While .Characters.Last.Previous = vbCr Or .Characters.Last.Previous = Chr(12)
        .Characters.Last.Previous = vbNullString
      Wend
       ' Replicate the headers & footers
      For Each HdFt In Rng.Sections(j).Headers
        .Sections(j).Headers(HdFt.Index).Range.FormattedText = HdFt.Range.FormattedText
      Next
      For Each HdFt In Rng.Sections(j).Footers
        .Sections(j).Footers(HdFt.Index).Range.FormattedText = HdFt.Range.FormattedText
      Next
       ' Save & close the output document
      .SaveAs FileName:=StrTxt & ".docx", FileFormat:=wdFormatXMLDocument, AddToRecentFiles:=False
       ' and/or:
      .SaveAs FileName:=StrTxt & ".pdf", FileFormat:=wdFormatPDF, AddToRecentFiles:=False
      .Close SaveChanges:=False
    End With
  Next
End With
Set Rng = Nothing: Set Doc = Nothing
Application.ScreenUpdating = True
End Sub
