# VBA-challenge
CWRU Data Analytics Boot Camp VBA Homework: The VBA of Wall Street

' codes created in Moduel 1
'this sub is for allowing the VBA scripts automatically run on every worksheet 2016,2015,2014

'Macro is assigned to "Click to Summarize Stock" Button

    Sub AutomateAdd()
     Dim i As Integer
      i = 1
      Do While i <= Worksheets.Count
         Worksheets(i).Select
         
    'add two previous created subs
        AddHeaderandGreatest
        SummarizeStock
        
      i = i + 1
      
      Columns("I:L").EntireColumn.AutoFit
      Columns("O:Q").EntireColumn.AutoFit
      
     Loop
    End Sub
'=================================================================

'Create the second sub SumarizeStock

    Sub SummarizeStock()
    
    'set variables for all columns in summary_table_row(I:L)
      Dim Ticker As String
      
      Dim TotalStockVolume As Double
          TotalStockVolume = 0
          
      Dim YearlyChange As Double
          YearlyChange = 0
          
      Dim PercentChange As Double
          PercentChange = 0
          
    'values start at row 2    
      Dim Summary_Table_Row As Integer
         Summary_Table_Row = 2
         
    'find the lastrow in Column A
      Dim lastrow As Long
      lastrow = Range("A" & Rows.Count).End(xlUp).Row
      
      For i = 2 To lastrow
      'grab opening price at the beginning of a given year (column C)
       openyearly = Cells(Summary_Table_Row, 3).Value
       
      'set condition:if next row's value is not equal to previous row's value in column A
        If Cells(i + 1, 1) <> Cells(i, 1).Value Then
           Ticker = Cells(i, 1).Value
          
          'just like answer=answer+startnumber, add together whatever is stored n the variable
           TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
           
          'each cell of row do substraction first then add up together
           YearlyChange = YearlyChange + (Cells(i, 6).Value - openyearly) 
           PercentChange = (YearlyChange / openyearly)
           
           Range("I" & Summary_Table_Row).Value = Ticker
           Range("L" & Summary_Table_Row).Value = TotalStockVolume
           Range("J" & Summary_Table_Row).Value = YearlyChange
           Range("K" & Summary_Table_Row).Value = PercentChange
           
          'set the format of column K as percentage
           Range("K" & Summary_Table_Row).Value = "Percent"
           
        Summary_Table_Row = Summary_Table_Row + 1 
        
        'reset values
        TotalStockVolume = 0
        YearlyChange = 0
        PercentChange = 0  
        
       Else
        TotalStockVolume = TotalStockVolume + Cells(i, 7).Value  
        
      End If
     
    Next i
 '---------------------------------------------------------------------------------------------------------------
 
    ' next set conditional formatting color for column J
     Dim lastrowyearly As Double
     lastrowyearly = Range("J" & Rows.Count).End(xlUp).Row
    
     For i = 2 To lastrowyearly
      'any values in column J
       If Cells(i, 10).Value >= 0 Then
          'highlight positive change in green and negative change in red
          Cells(i, 10).Interior.Color = RGB(0, 255, 0)
         
       Else
          Cells(i, 10).Interior.Color = RGB(255, 0, 0)
         
       End If
   
     Next i
  '---------------------------------------------------------------------------------------------------------------
  
    'then fill values into column P and Q
    'set lastrow to find the max/min value from percent change column
     Dim lastrowpercent As Long
     lastrowpercent = Range("K" & Rows.Count).End(xlUp).Row
    
    ' set max value starts at 0
    Dim Greatest_Percent_Increase As Double
        Greatest_Percent_Increase = 0
    
    Dim Greatest_Percent_Decrease As Double
        Greatest_Percent_Decrease = 0
    
    For i = 2 To lastrowpercent
        'compare 0 and max value, grab max value
        If Greatest_Percent_Increase < WorksheetFunction.Max(Cells(i, 11)) Then
           Greatest_Percent_Increase = WorksheetFunction.Max(Cells(i, 11))  
           Cells(2, 17).Value = Greatest_Percent_Increase 
           Cells(2, 16).Value = Cells(i, 9).Value
         
       ElseIf Greatest_Percent_Decrease > WorksheetFunction.Max(Cells(i, 11)) Then
           Greatest_Percent_Decrease = WorksheetFunction.Max(Cells(i, 11))
           Cells(3, 17).Value = Greatest_Percent_Decrease 
           Cells(3, 16).Value = Cells(i, 9).Value  
         
      'set the format of Q2 and Q3 as percentage
       Range("Q2:Q3").Style = "Percent"
      
       End If
     
    Next i
 '---------------------------------------------------------------------------------------------------------------
 
    'after that set lastrow to find the max volume from total volume column
     Dim lastrowvolume As Long
     lastrowvolume = Range("L" & Rows.Count).End(xlUp).Row
    
     Dim Greatest_Total_Volume As Double
         Greatest_Total_Volume = 0
    
     For i = 2 To lastrowvolume
       'compare 0 and max value, grab max value
       If Greatest_Total_Volume < WorksheetFunction.Max(Cells(i, 12)) Then
          Greatest_Total_Volume = WorksheetFunction.Max(Cells(i, 12)) 
          Cells(4, 17).Value = Greatest_Total_Volume
          Cells(4, 16).Value = Cells(i, 9).Value
         
       End If
      
     Next i
    
    End Sub
'=================================================================

'Create the first sub AddHeaderandGreatest

    Sub AddHeaderandGreatest()

    'use Macro Recording function to add Headers for column I:L and O:Q
     Range("02").Select
     ActiveCell.FormulaR1C1 = "Greatest % Increase"
     Range("03").Select
     ActiveCell.FormulaR1C1 = "Greatest % Decrease"
     Range("04").Select
     ActiveCell.FormulaR1C1 = "Greatest Total Volume"
     Range("P1").Select
     ActiveCell.FormulaR1C1 = "Ticker"
     Range("Q1").Select
     ActiveCell.FormulaR1C1 = "Value"
   
     Range("I1").Select
     ActiveCell.FormulaR1C1 = "Ticker"
     Range("J1").Select
     ActiveCell.FormulaR1C1 = "Yearly Change"
     Range("K1").Select
     ActiveCell.FormulaR1C1 = "Percent Change"
     Range("L1").Select
     ActiveCell.FormulaR1C1 = "Total Stock Voume"
   
    End Sub
'=================================================================  

' codes created in Moduel 2

'Macro is assigned to "Clear Enterred Data" Button

    Sub ClearEnterredData()
 
      Dim i As Integer
   
      i = 1
   
       Do While i <= Worksheets.Count
          Worksheets(i).Select
    
      Range("I:Q").ClearContents
      Range("I:Q").Interior.Color = xlNone
   
      i = i + 1
   
     Loop
   
    End Sub
