Attribute VB_Name = "test"




Sub InsertRowEvery30RowsWithMerging()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim firstColumn As Long
    Dim lastColumn As Long

    ' Set the active worksheet
    Set ws = ThisWorkbook.Worksheets("natega")

    ' Get the first and last columns in the worksheet
    firstColumn = ws.Cells(11, 1).Column  ' Assuming data starts from column A
    ' firstColumn = "A1"  ' Assuming data starts from column A
    MsgBox firstColumn
    ' lastColumn = ws.Cells(14, ws.Columns.Count).End(xlToLeft).Column ' Get the last used column
    lastColumn = 58 ' Get the last used column
    MsgBox lastColumn
    ' Find the last used row in the worksheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ' lastRow=236

    ' Loop through the rows and insert 3 rows after every 30 rows
    For i = 33 To lastRow+76 Step 26
        ' Insert 3 rows after the current 30th row
        ws.Rows(i + 1).Insert Shift:=xlDown
        ws.Rows(i + 2).Insert Shift:=xlDown
        ws.Rows(i + 3).Insert Shift:=xlDown
        ws.Rows(i + 4).Insert Shift:=xlDown
        
        ws.Rows(i + 1).RowHeight = 128  ' Adjust the row height as needed
        ws.Rows(i + 2).RowHeight = 128  ' Adjust the row height as needed
        ws.Rows(i + 3).RowHeight = 128  ' Adjust the row height as needed
        ws.Rows(i + 4).RowHeight = 128  ' Adjust the row height as needed
        
        
        ' First inserted row: Merge all cells from first to last column
        With ws.Range(ws.Cells(i + 1, firstColumn), ws.Cells(i + 1, lastColumn))
            .Merge
            .Value = "„ : «„ Ì«“     Ã‹‹ Ã‹‹ :ÃÌœ Ãœ«     Ã‹‹//  :ÃÌœ     · : „ﬁ»Ê·     —· : —«”» ·«∆Õ…    ÷ : ÷⁄Ì›     ÷ Ã‹ :÷⁄Ì› Ãœ«"
            .Font.Bold = True  ' Apply bold font
            .Font.Size = 55  ' Set font size
            .Font.Name="Calibri"
            ' .Font.Color = RGB(255, 255, 255)  ' Set font color to white
            .Interior.Color = RGB(255, 255, 255)  ' Set background color white
            .HorizontalAlignment = xlCenter  ' Center-align horizontally
            .VerticalAlignment = xlCenter  ' Center-align vertically
        End With

        ' Change the height of the row after formatting

        
        'Second inserted row: Merge cells across three specific columns (e.g., columns A, B, and C)
        With ws.Range(ws.Cells(i + 2, firstColumn), ws.Cells(i + 2, firstColumn+4))
            .Merge
            .Value =  " ÊﬂÌ· «·ﬂ·Ì…" 
            .Font.Bold = True  ' Apply bold font
            .Font.Size = 66  ' Set font size
            .Font.Name="Calibri"
            .WrapText = True  ' Enable text wrapping
            ' .Font.Color = RGB(255, 255, 255)  ' Set font color to white
            ' .Interior.Color = RGB(0, 0, 255)  ' Set background color to blue
            .Interior.Color = RGB(255, 255, 255)  ' Set background color white
            .HorizontalAlignment = xlCenter  ' Center-align horizontally
            .VerticalAlignment = xlCenter  ' Center-align vertically
        End With


        With ws.Range(ws.Cells(i + 2, firstColumn+5), ws.Cells(i + 2, firstColumn+21))
            .Merge
            .Value = "⁄„Ìœ «·ﬂ·Ì…"
            .Font.Bold = True  ' Apply bold font
            .Font.Size = 66  ' Set font size
            .Font.Name="Calibri"
            .WrapText = True  ' Enable text wrapping
            ' .Font.Color = RGB(255, 255, 255)  ' Set font color to white
            ' .Interior.Color = RGB(0, 0, 255)  ' Set background color to blue
            .Interior.Color = RGB(255, 255, 255)  ' Set background color white
            .HorizontalAlignment = xlCenter  ' Center-align horizontally
            .VerticalAlignment = xlCenter  ' Center-align vertically
        End With

        With ws.Range(ws.Cells(i + 2, firstColumn+22), ws.Cells(i + 2, firstColumn+41))
            .Merge
            .Value = "‰«∆» —∆Ì” «·Ã«„⁄… ·‘∆Ê‰ «· ⁄·Ì„ Ê«·ÿ·«» "
            .Font.Bold = True  ' Apply bold font
            .Font.Size = 66  ' Set font size
            .Font.Name="Calibri"
            .WrapText = True  ' Enable text wrapping
            ' .Font.Color = RGB(255, 255, 255)  ' Set font color to white
            ' .Interior.Color = RGB(0, 0, 255)  ' Set background color to blue
            .Interior.Color = RGB(255, 255, 255)  ' Set background color white
            .HorizontalAlignment = xlCenter  ' Center-align horizontally
            .VerticalAlignment = xlCenter  ' Center-align vertically
        End With

        With ws.Range(ws.Cells(i + 2, firstColumn+42), ws.Cells(i + 2, lastColumn))
            .Merge
            .Value = "—∆Ì” «·Ã«„⁄… "
            .Font.Bold = True  ' Apply bold font
            .Font.Size = 66  ' Set font size
            .Font.Name="Calibri"
            .WrapText = True  ' Enable text wrapping
            ' .Font.Color = RGB(255, 255, 255)  ' Set font color to white
            ' .Interior.Color = RGB(0, 0, 255)  ' Set background color to blue
            .Interior.Color = RGB(255, 255, 255)  ' Set background color white
            .HorizontalAlignment = xlCenter  ' Center-align horizontally
            .VerticalAlignment = xlCenter  ' Center-align vertically
        End With

        ' third inserted row
        With ws.Range(ws.Cells(i + 3, firstColumn), ws.Cells(i + 3, firstColumn+4))
            .Merge
            .Value = "√.„.œ/ Â‘«„ “ﬂ—Ì«"
            .Font.Bold = True  ' Apply bold font
            .Font.Size = 66  ' Set font size
            .Font.Name="Calibri"
            .WrapText = True  ' Enable text wrapping
            ' .Font.Color = RGB(255, 255, 255)  ' Set font color to white
            ' .Interior.Color = RGB(0, 0, 255)  ' Set background color to blue
            .Interior.Color = RGB(255, 255, 255)  ' Set background color white
            .HorizontalAlignment = xlCenter  ' Center-align horizontally
            .VerticalAlignment = xlCenter  ' Center-align vertically
        End With


        With ws.Range(ws.Cells(i + 3, firstColumn+5), ws.Cells(i + 3, firstColumn+21))
            .Merge
            .Value = " √.œ / „Õ„œ „’ÿ›Ï "	
            .Font.Bold = True  ' Apply bold font
            .Font.Size = 66  ' Set font size
            .Font.Name="Calibri"
            .WrapText = True  ' Enable text wrapping
            ' .Font.Color = RGB(255, 255, 255)  ' Set font color to white
            ' .Interior.Color = RGB(0, 0, 255)  ' Set background color to blue
            .Interior.Color = RGB(255, 255, 255)  ' Set background color white
            .HorizontalAlignment = xlCenter  ' Center-align horizontally
            .VerticalAlignment = xlCenter  ' Center-align vertically
        End With

        With ws.Range(ws.Cells(i + 3, firstColumn+22), ws.Cells(i + 3, firstColumn+41))
            .Merge
            .Value =  "√.œ / „Õ„œ »·«·"	
            .Font.Bold = True  ' Apply bold font
            .Font.Size = 66  ' Set font size
            .Font.Name="Calibri"
            .WrapText = True  ' Enable text wrapping
            ' .Font.Color = RGB(255, 255, 255)  ' Set font color to white
            ' .Interior.Color = RGB(0, 0, 255)  ' Set background color to blue
            .Interior.Color = RGB(255, 255, 255)  ' Set background color white
            .HorizontalAlignment = xlCenter  ' Center-align horizontally
            .VerticalAlignment = xlCenter  ' Center-align vertically
        End With

        With ws.Range(ws.Cells(i + 3, firstColumn+42), ws.Cells(i + 3, lastColumn))
            .Merge
            .Value ="√.œ/ Â»… ›«—Êﬁ ”«·„"  	
            .Font.Bold = True  ' Apply bold font
            .Font.Size = 66  ' Set font size
            .Font.Name="Calibri"
            .WrapText = True  ' Enable text wrapping
            ' .Font.Color = RGB(255, 255, 255)  ' Set font color to white
            ' .Interior.Color = RGB(0, 0, 255)  ' Set background color to blue
            .Interior.Color = RGB(255, 255, 255)  ' Set background color white
            .HorizontalAlignment = xlCenter  ' Center-align horizontally
            .VerticalAlignment = xlCenter  ' Center-align vertically
        End With



        
        ' forth inserted row: Merge all cells from first to last column
        With ws.Range(ws.Cells(i + 4, firstColumn), ws.Cells(i + 4, firstColumn+4))
            .Merge
            .Value = "«· ÊﬁÌ⁄ ...................."
            .Interior.Color = RGB(255, 255, 255)  ' Set background color white
            .Font.Bold = True
            .Font.Name="Calibri"
            .Font.Size = 66
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        With ws.Range(ws.Cells(i + 4, firstColumn + 5), ws.Cells(i + 4, firstColumn+21))
            .Merge
            .Value =  "«· ÊﬁÌ⁄ ...................."
            .Interior.Color = RGB(255, 255, 255)  ' Set background color white
            .Font.Bold = True
            .Font.Name="Calibri"
            .Font.Size = 66
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        With ws.Range(ws.Cells(i + 4, firstColumn + 22), ws.Cells(i + 4, firstColumn+41))
            .Merge
            .Value =  " .................... «· ÊﬁÌ⁄ "
            .Interior.Color = RGB(255, 255, 255)  ' Set background color white
            .Font.Bold = True
            .Font.Name="Calibri"
            .Font.Size = 66
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        With ws.Range(ws.Cells(i + 4, firstColumn + 42), ws.Cells(i + 4, lastColumn))
            .Merge
            .Value =  "«· ÊﬁÌ⁄ ...................."
            .Interior.Color = RGB(255, 255, 255)  ' Set background color white
            .Font.Bold = True
            .Font.Name="Calibri"
            .Font.Size = 66
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With


        With ws.Rows(i + 2).Borders(xlEdgeBottom)
            .LineStyle = xlNone  ' Remove  bottom border
        End With


    Next i
End Sub




