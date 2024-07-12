# This repository contains the homework for the UNCC Data Analysis BootCamp for:
# Module 2 VBA Challenge

# VBA Sub Quarters()
## This VBA code was written to analyze the Quarterly price change and volume of different stocks in a given excel workbook file(Multiple_year_stock_data.xlm) organized by Quarters in different worksheets. 

## This code section adds headers and formats the column widths for two summary tables in each worksheet. Table 1 - I-L columns, and Table 2- range (N1:P4)


Sub AllQuarters()

Dim Sheet As Worksheet
   
'Loop to cycle through each worksheet in the workbook
For Each Sheet In ActiveWorkbook.Worksheets
    ' Format columns for data summary
    Sheet.Cells(1, 9).Value = "Ticker"
    Sheet.Cells(1, 10).Value = "Quarterly Change"
    Sheet.Cells(1, 10).ColumnWidth = 18
    Sheet.Cells(1, 11).Value = "Percent Change"
    Sheet.Cells(1, 11).ColumnWidth = 18
    Sheet.Cells(1, 12).Value = "Total Stock Volume"
    Sheet.Cells(1, 12).ColumnWidth = 18
    Sheet.Cells(2, 14).Value = "Greatest % Increase"
    Sheet.Cells(1, 14).ColumnWidth = 18
    Sheet.Cells(3, 14).Value = "Greatest % Decrease"
    Sheet.Cells(1, 14).ColumnWidth = 18
    Sheet.Cells(4, 14).Value = "Greatest Total Value"
    Sheet.Cells(1, 14).ColumnWidth = 18
    Sheet.Cells(1, 15).Value = "Ticker"
    Sheet.Cells(1, 16).Value = "Value"
    Sheet.Cells(1, 16).ColumnWidth = 18

## This code section generates the summary table and fills the columns under the headers with: (1)Stock Ticker, (3)Quarterly change from opening price to closing price for a given Quarter, (4) the percent change from opening price to closing price for a given Quarter, (5) total stock volume.
    
    'Summarizing the data for each stock type in one sheet of the worksheet
    Dim lRow As Long
    Dim i As Long

    Dim j As Integer
    Dim Ticker As String
    Dim QOpen As Double
    Dim QClose As Double
    Dim VolTotal As Double
    
    j = 2
    VolTotal = 0
    QOpen = Sheet.Cells(2, 3).Value
    lRow = Sheet.Cells(Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To lRow
        
        If Sheet.Cells(i + 1, 1).Value <> Sheet.Cells(i, 1).Value Then
            Ticker = Sheet.Cells(i, 1).Value
            QClose = Sheet.Cells(i, 6).Value
            VolTotal = VolTotal + Sheet.Cells(i, 7).Value
            
            'fill summary table
            
            Sheet.Cells(j, 9).Value = Ticker
            Sheet.Cells(j, 10).Value = QClose - QOpen
            Sheet.Cells(j, 11).Value = Sheet.Cells(j, 10).Value / QOpen
            Sheet.Cells(j, 11).Value = FormatPercent(Sheet.Cells(j, 11))
            Sheet.Cells(j, 12).Value = VolTotal
            j = j + 1
            QOpen = Sheet.Cells(i + 1, 3).Value
            VolTotal = 0
        Else
            VolTotal = VolTotal + Sheet.Cells(i, 7).Value
        
        End If
           
    Next i

## This code section finds the (1) Greatest % Increase in value and related stock ticker,  (2) Greatest % Decrease in value and related stock ticker, (3) Greatest Total Volume value and related stock ticker. Table range (N1:P4)
    'Finding the maximum and minimum stocks in percent change and total volume
       
    Dim MaxTicker As String
    Dim MinTicker As String
    Dim TotalTicker As String
    Dim MaxChange As Double
    Dim MinChange As Double
    Dim MaxTotal As Double
    
    
    
    lRow = Sheet.Cells(Rows.Count, "I").End(xlUp).Row
    MaxChange = Sheet.Cells(2, 11).Value
    MinChange = Sheet.Cells(2, 11).Value
    MaxTotal = Sheet.Cells(2, 12).Value
    
    For i = 2 To lRow
    
        
        If Sheet.Cells(i + 1, 11).Value > MaxChange Then
            MaxChange = Sheet.Cells(i + 1, 11).Value
            MaxTicker = Sheet.Cells(i + 1, 9).Value
            
        ElseIf Sheet.Cells(i + 1, 11).Value < MinChange Then
            MinChange = Sheet.Cells(i + 1, 11).Value
            MinTicker = Sheet.Cells(i + 1, 9).Value
            
        End If
        
        
        If Sheet.Cells(i + 1, 12).Value > MaxTotal Then
            MaxTotal = Sheet.Cells(i + 1, 12).Value
            TotalTicker = Sheet.Cells(i + 1, 9).Value
        
        End If
        
    Next i
    
    Sheet.Cells(2, 15).Value = MaxTicker
    Sheet.Cells(2, 16).Value = MaxChange
    Sheet.Cells(2, 16).Value = FormatPercent(Sheet.Cells(2, 16))
    Sheet.Cells(3, 15).Value = MinTicker
    Sheet.Cells(3, 16).Value = MinChange
    Sheet.Cells(3, 16).Value = FormatPercent(Sheet.Cells(3, 16))
    Sheet.Cells(4, 15).Value = TotalTicker
    Sheet.Cells(4, 16).Value = MaxTotal
    
    'Color code the stock Quarterly Change summary data
   
    Dim PChange As Double
    
    lRow = Sheet.Cells(Rows.Count, "K").End(xlUp).Row
    
    For i = 2 To lRow
        
        If Sheet.Cells(i, 10).Value > 0 Then
            Sheet.Cells(i, 10).Interior.Color = vbGreen
            
            
        ElseIf Sheet.Cells(i, 10).Value < 0 Then
            Sheet.Cells(i, 10).Interior.Color = vbRed
            
        End If
    
    Next i
        
    
Next Sheet

## A button on the first worksheet is used to run the Macro and a message box is used to let the user know the macro is complete.
MsgBox "Analysis complete!"
    
End Sub

# Attachments
## Excel file with enabled Macro: Multiple_year_stock_data_MXR_final.xlm
## Screen captures: 1- Macro Start.png, 2- Macro End.png, 3- Q1 Worksheet output.png, 4- Q2 Worksheet output.png, 5- Q3 Worksheet output.png, 6- Q4 Worksheet output.png

# Credits
## In addition to the class exercises, I used different online searches to understand better the syntax and coding options to complete the task. In particular I used as a reference the book "Excel VBA in Easy Steps", 3rd edition by Mike Grath.

#
