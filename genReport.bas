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

Private Sub executive_summary()
    ' Add heading table
    Set t = new_table(2, 2)
    Set c = t.Cell(2, 1)
    Call t.Cell(1, 1).Merge(c)
    t.Borders.Enable = False
    
    ' Populate heading table
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

Private Sub procedure(steps)
    ' Add Procedure Section
    Call stylized_text("Procedure", "Heading 1")
    Set t = new_table(4, 1)
    t.Borders.Enable = False
End Sub

Private Sub update_toc()
    app().ActiveDocument.TablesOfContents(1).Update
End Sub

Sub genReport()
    Call executive_summary
    Call table_of_contents
    steps = Array("One", "Two", "Three")
    Call time_estimate(steps)
    Call procedure(steps)
    Call update_toc
End Sub
