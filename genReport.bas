Attribute VB_Name = "NewMacros"
'
' genReport
'
' Copyright (c) 2016 Christopher Crawford
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
'
'
Private Function app()
    Set app = Application
End Function

Private Sub right(n)
    For i = 1 To n
        Call app().Selection.MoveRight
    Next
End Sub

Private Sub left(n)
    For i = 1 To n
        Call app().Selection.MoveLeft
    Next
End Sub

Private Sub up(n)
    For i = 1 To n
        Call app().Selection.MoveUp
    Next
End Sub

Private Sub down(n)
    For i = 1 To n
        Call app().Selection.MoveDown
    Next
End Sub

Private Sub new_paragraph()
    Call app().Selection.Paragraphs.Add
End Sub

Private Sub next_line()
    Call new_paragraph
    Call down(1)
End Sub

Private Sub new_page()
    ' wdPageBreak
    my_wdPageBreak = 7
    Call app().Selection.InsertBreak(my_wdPageBreak)
End Sub

Private Sub text(s)
    app().Selection.text = s
End Sub

Private Sub stylized_text(txt, style)
    Call text(txt)
    app().Selection.style = app().ActiveDocument.Styles(style)
    Call next_line
End Sub

Private Sub bulleted_list(a)
    app().Selection.Range.ListFormat.ApplyBulletDefault
    For i = LBound(a) To UBound(a)
        Call text(a(i))
        Call next_line
    Next
    app().Selection.Range.ListFormat.RemoveNumbers
End Sub

Private Function new_table(rows, cols, border)
    Set t = app().ActiveDocument.Tables.Add(app().Selection.Range, rows, cols)
    t.TopPadding = 0
    t.RightPadding = 0
    t.LeftPadding = 0
    t.BottomPadding = 0
    t.Select
    Selection.style = app().ActiveDocument.Styles("No Spacing")
    Call left(1)
    If border Then
        t.Borders.Enable = True
    Else
        t.Borders.Enable = False
    End If
    Set new_table = t
End Function

Private Function h1(s)
    Call stylized_text(s, "Heading 1")
End Function

Private Function h2(s)
    Call stylized_text(s, "Heading 2")
End Function

Private Function pic(path, border)
    Set p = app().ActiveDocument.Shapes.AddPicture(path, False, True)
    ' wdWrapInline = 7
    my_wdWrapInline = 7
    p.WrapFormat.Type = my_wdWrapInline
    If border Then
        p.Line.Weight = 1
        ' RGB(0,0,0) = 0
        p.Line.ForeColor.RGB = 0
    End If
    Set pic = p
End Function

Private Sub exec_sum_header(logo)
    ' Add heading table
    Set t = new_table(2, 2, False)
    Set c = t.Cell(2, 1)
    Call t.Cell(1, 1).Merge(c)
    Call pic(logo, False)
    
    ' Populate heading table
    Call right(2)
    Call text("<<Report Title>>")
    Selection.style = app().ActiveDocument.Styles("Title")
    Call right(1)
    Call new_paragraph
    Call down(1)
    Call text("UBNETDEF Field Guide")
    Selection.style = app().ActiveDocument.Styles("Subtitle")
    Call down(1)
    Call text("<<Author Name>>")
    Call right(1)
    Call new_paragraph
    Call down(1)
    Call text("<<YYYY-MM-DD>>")
    Call down(1)
End Sub

Private Sub exec_sum_content()
    ' Populate Excutive Summary Page
    Call h1("Executive Summary")
    Call h2("Objective")
    Call text("After completing this guide, the reader will be able to <<finish this statement>>.")
    Call next_line
    
    Call h2("Requirements")
    Call text("In order to complete this guide, the reader will need the following:")
    Call next_line
    
    Call bulleted_list(Array("<<Stuff>>", "<<Things>>", "<<More Things>>"))
    
    Call h2("Time Estimate")
    Call text("The reader can expect the following procedure to take <<X>> minutes.")
    
    Call next_line
    Call new_page
End Sub

Private Sub executive_summary(logo)
    Call exec_sum_header(logo)
    Call exec_sum_content
End Sub

Private Sub table_of_contents()
    ' Add Table of Contents
    Call h1("Table of Contents")
    Call app().ActiveDocument.TablesOfContents.Add(app().Selection.Range)
    Call down(1)
    Call new_page
End Sub

Private Sub addRow(table)
    Call table.rows.Add
    Call down(1)
End Sub

Private Sub addTableHeaders(t)
    Call text("Step")
    Call right(2)
    Call text("Time (minutes)")
End Sub

Private Sub addTableData(t, steps)
    Call t.Cell(1, 1).Select
    For i = LBound(steps) To UBound(steps)
        Call addRow(t)
        Call text(steps(i))
    Next
End Sub

Private Sub addTotalRow(t)
    Call addRow(t)
    Call text("Total Time")
End Sub

Private Sub setColumnWidths(t)
    Call t.Columns(1).SetWidth(404, wdAdjustNone)
    Call t.Columns(2).SetWidth(72, wdAdjustNone)
End Sub

Private Sub centerTable(t)
    t.rows.Alignment = wdAlignRowCenter
End Sub

Private Sub formatHeaderRow(t)
    Set rng = t.rows(1).Range
    rng.Font.Bold = True
    rng.ParagraphFormat.Alignment = wdAlignParagraphCenter
End Sub

Private Sub setTableFonts(t)
    Set rng = t.rows(2).Range
    rng.End = t.rows(t.rows.Count - 1).Range.End
    rng.Font.Name = "Courier New"
    t.Cell(t.rows.Count, 2).Range.Font.Name = "Courier New"
    t.Range.Font.Size = 8
End Sub

Private Sub setAlignmentForTimeData(t As table)
    Set rng = t.Cell(2, 2).Range
    rng.End = t.Cell(t.rows.Count, 2).Range.End
    rng.Select
    app().Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
End Sub

Private Sub italicizeSteps(t)
    Set rng = t.Cell(2, 1).Range
    rng.End = t.Cell(t.rows.Count - 1, 1).Range.End
    rng.Select
    app().Selection.Font.Italic = True
End Sub

Private Sub setTableBorders(t)
    ' Format table borders
    t.Borders.InsideLineWidth = wdLineWidth075pt
    Set rng = t.Cell(2, 1).Range
    rng.End = t.Cell(t.rows.Count - 1, 2).Range.End
    rng.Select
    app().Selection.Borders.OutsideLineWidth = wdLineWidth150pt
        
    Set rng = t.Cell(1, 1).Range
    rng.End = t.Cell(t.rows.Count, 1).Range.End
    rng.Select
    Selection.Borders.OutsideLineStyle = wdLineStyleSingle
    Selection.Borders.OutsideLineWidth = wdLineWidth150pt
    
    t.Borders.OutsideLineStyle = wdLineStyleSingle
    t.Borders.OutsideLineWidth = wdLineWidth225pt
End Sub

Private Sub setTablePadding(t)
    t.LeftPadding = 5
    t.RightPadding = 15
End Sub

Private Sub boldTotalsRow(t)
    Set rng = t.rows(t.rows.Count).Range
    rng.Select
    Selection.Font.Bold = True
End Sub

Private Sub removeItalicsFromTimeData(t)
    ' Make sure time data is not italicized
    t.Columns(2).Select
    Selection.Font.Italic = False
End Sub

Private Sub shadeBandedRows(t)
    For i = 2 To t.rows.Count - 1
        If i Mod 2 = 0 Then
            t.rows(i).Shading.BackgroundPatternColor = wdColorGray20
        End If
    Next
End Sub

Private Sub time_estimate(steps)
    ' Add Time Estimate Section
    Call h1("Time Estimate Table")
    Call next_line
    
    ' Add Time Estimate Table
    Dim t As table
    Set t = new_table(1, 2, True)
        
    Call addTableHeaders(t)
    Call addTableData(t, steps)
    Call addTotalRow(t)
    Call centerTable(t)
    Call formatHeaderRow(t)
    Call setTableFonts(t)
    Call setAlignmentForTimeData(t)
    Call italicizeSteps(t)
    Call setTableBorders(t)
    Call setTablePadding(t)
    Call boldTotalsRow(t)
    Call removeItalicsFromTimeData(t)
    Call setColumnWidths(t)
    Call shadeBandedRows(t)
    Call down(1)
    Call new_page
End Sub

Private Sub procedure_step(i)
    Set t = new_table(6, 1, False)
    Call h2(i)
    Call app().Selection.TypeBackspace
    Call down(2)
    Call text("Estimated Time Required: " & "<<X>>" & " minutes")
    Call down(2)
    Set p = pic("C:\\Users\\Chris\\Desktop\\ubnetdef.png", True)
    ' wdAlignParagraphCenter = 1
    my_wdAlignParagraphCenter = 1
    app().Selection.ParagraphFormat.Alignment = my_wdAlignParagraphCenter
    Call down(2)
    Call new_page
End Sub

Private Sub procedure(steps)
    ' Add Procedure Section
    Call stylized_text("Procedure", "Heading 1")
    For i = LBound(steps) To UBound(steps)
        Call procedure_step(steps(i))
    Next
    Call app().Selection.TypeBackspace
    Call app().Selection.TypeBackspace
    Call app().Selection.TypeBackspace
End Sub

Private Sub update_toc()
    app().ActiveDocument.TablesOfContents(1).Update
End Sub

Private Sub add_page_numbers()
    ' wdHeaderFooterPrimary = 1
    ' wdAlignPageNumberCenter = 1
    my_wdHeaderFooterPrimary = 1
    my_wdAlignPageNumberCenter = 1
    Call app().ActiveDocument.Sections(1).Footers(my_wdHeaderFooterPrimary).PageNumbers.Add(my_wdAlignPageNumberCenter, False)
End Sub

Sub genReport()
    Call add_page_numbers
    logo = "C:\\Users\\Chris\\Desktop\\ubnetdef.png"
    Call executive_summary(logo)
    Call table_of_contents
    steps = Array("One", "Two", "Three")
    Call time_estimate(steps)
    Call procedure(steps)
    Call update_toc
End Sub
