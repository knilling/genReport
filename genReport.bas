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

Attribute VB_Name = "NewMacros"
Private Function app()
    Set app = Application
End Function

Private Sub right()
    Call app().Selection.MoveRight
End Sub

Private Sub left()
    Call app().Selection.MoveLeft
End Sub

Private Sub up()
    Call app().Selection.MoveUp
End Sub

Private Sub down()
    Call app().Selection.MoveDown
End Sub

Private Sub new_paragraph()
    Call app().Selection.Paragraphs.Add
End Sub

Private Sub next_line()
    Call new_paragraph
    Call down
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

Private Function new_table(rows, cols)
    Set new_table = app().ActiveDocument.Tables.Add(app().Selection.Range, rows, cols)
End Function

Private Function h1(s)
    Call stylized_text(s, "Heading 1")
End Function

Private Function h2(s)
    Call stylized_text(s, "Heading 2")
End Function

Private Function pic(path, border)
    Set p = app().ActiveDocument.Shapes.AddPicture(path, False, True)
    p.WrapFormat.Type = wdWrapInline
    If border Then
        p.Line.Weight = 1
        p.Line.ForeColor.RGB = RGB(0, 0, 0)
    End If
    Set pic = p
End Function

Private Sub executive_summary()
    ' Add heading table
    Set t = new_table(2, 2)
    t.Borders.Enable = False
    Set c = t.Cell(2, 1)
    Call t.Cell(1, 1).Merge(c)
    Call pic("C:\\Users\\Chris\\Desktop\\ubnetdef.png", False)
    
    ' Populate heading table
    Call right
    Call right
    Call text("<<Report Title>>")
    Selection.style = app().ActiveDocument.Styles("Title")
    Call right
    Call new_paragraph
    Call down
    Call text("UBNETDEF Field Guide")
    Selection.style = app().ActiveDocument.Styles("Subtitle")
    Call down
    Call text("<<Author Name>>")
    Call right
    Call new_paragraph
    Call down
    Call text("<<YYYY-MM-DD>>")
    Call down
    
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

Private Sub table_of_contents()
    ' Add Table of Contents
    Call h1("Table of Contents")
    Call app().ActiveDocument.TablesOfContents.Add(app().Selection.Range)
    Call down
    Call new_page
End Sub

Private Sub addRow(table)
    Call table.rows.Add
    Call down
End Sub

Private Sub time_estimate(steps)
    ' Add Time Estimate Section
    Call h1("Time Estimate")
    Call next_line
    Set t = new_table(1, 2)
    t.Borders.Enable = True
    Call text("Step")
    Call right
    Call right
    Call text("Estimated Time to Complete")
    Call t.Range.Select
    Call app().Selection.Collapse(1)
    For i = LBound(steps) To UBound(steps)
        Call addRow(t)
        Call text(steps(i))
    Next
    Call addRow(t)
    Call text("Total")
    Call down
    Call new_page
End Sub

Private Sub procedure_step(i)
    Set t = new_table(5, 1)
    t.Borders.Enable = False
    Call h2(i)
    Call app().Selection.TypeBackspace
    Call down
    Call down
    Set p = pic("C:\\Users\\Chris\\Desktop\\ubnetdef.png", True)
    app().Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Call down
    Call down
    Call down
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
End Sub

Private Sub update_toc()
    app().ActiveDocument.TablesOfContents(1).Update
End Sub

Private Sub add_page_numbers()
    Call app().ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary).PageNumbers.Add(wdAlignPageNumberCenter, False)
End Sub

Sub genReport()
    Call add_page_numbers
    Call executive_summary
    Call table_of_contents
    steps = Array("One", "Two", "Three")
    Call time_estimate(steps)
    Call procedure(steps)
    Call update_toc
End Sub
