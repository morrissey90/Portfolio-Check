# Portfolio-Check
Code for daily portfolio upload

Sub PC_all_accounts()

Application.ScreenUpdating = False
Application.DisplayAlerts = False
'clear the previous days data

Sheets("Calculator").Select
Range("Calculator!J5") = Date


        
                'clear position data
                
                Sheets(Array("CROWN", "IBIS", "IBISX2", "TSMA", "ESMA", "MSI")).Select
                Cells.Select
                Selection.ClearContents
                               
                
                'clear Trading levels from calculator tab
                
                 Sheets("Calculator").Select
                 
                
                 
                 Range("F2:F7").Select
                 Selection.ClearContents
                           
        Set wbPortfolioCheck = ActiveWorkbook


'open arbor data which has been downloaded and copy and paste it into relevant tab in portfolio check


                  ChDir "H:\HDrive\Hedge fund\Operations\Portfolio"
                Workbooks.Open Filename:= _
                    "H:\HDrive\Hedge fund\Operations\Portfolio\Portfolio-C.xls"
                Cells.Select
                Selection.Copy
                wbPortfolioCheck.Activate
                Sheets("CROWN").Select
                Range("A1").Select
                ActiveSheet.Paste
                
                 ChDir "H:\HDrive\Hedge fund\Operations\Portfolio"
                Workbooks.Open Filename:= _
                    "H:\HDrive\Hedge fund\Operations\Portfolio\Portfolio-I.xls"
                Cells.Select
                Selection.Copy
                wbPortfolioCheck.Activate
                Sheets("IBIS").Select
                Range("A1").Select
                ActiveSheet.Paste
                
                '   ChDir "H:\HDrive\Hedge fund\Operations\Portfolio"
             '   Workbooks.Open Filename:= _
            '        "H:\HDrive\Hedge fund\Operations\Portfolio\Portfolio-II.xls"
            '    Cells.Select
            '    Selection.Copy
            '    wbPortfolioCheck.Activate
            '    Sheets("IBISX2").Select
            '    Range("A1").Select
            '    ActiveSheet.Paste
                
                 ChDir "H:\HDrive\Hedge fund\Operations\Portfolio"
                Workbooks.Open Filename:= _
                    "H:\HDrive\Hedge fund\Operations\Portfolio\Portfolio-E.xls"
                Cells.Select
                Selection.Copy
                wbPortfolioCheck.Activate
                Sheets("ESMA").Select
                Range("A1").Select
                ActiveSheet.Paste
                
                  ChDir "H:\HDrive\Hedge fund\Operations\Portfolio"
                Workbooks.Open Filename:= _
                    "H:\HDrive\Hedge fund\Operations\Portfolio\Portfolio-T.xls"
                Cells.Select
                Selection.Copy
                wbPortfolioCheck.Activate
                Sheets("TSMA").Select
                Range("A1").Select
                ActiveSheet.Paste
                
                
                    ChDir "H:\HDrive\Hedge fund\Operations\Portfolio"
                Workbooks.Open Filename:= _
                    "H:\HDrive\Hedge fund\Operations\Portfolio\Portfolio-M.xls"
                Cells.Select
                Selection.Copy
                wbPortfolioCheck.Activate
                Sheets("MSI").Select
                Range("A1").Select
                ActiveSheet.Paste
    
'close the arbor portfolio downloads
    
    
                Workbooks("Portfolio-I.xls").Close Savechanges:=False
              '  Workbooks("Portfolio-II.xls").Close Savechanges:=False
                Workbooks("Portfolio-E.xls").Close Savechanges:=False
                Workbooks("Portfolio-T.xls").Close Savechanges:=False
                Workbooks("Portfolio-C.xls").Close Savechanges:=False
                Workbooks("Portfolio-M.xls").Close Savechanges:=False
                
                
'Open the Trading level spreadsheets and copy in the TL figure to the calculator tab


        'ESMA and TSMA
Workbooks.Open Filename:= _
        "H:\HDrive\Hedge fund\Operations\Operations\Incentive Fee Accrual Summary.xlsx"

        
        
         Workbooks.Open Filename:= _
        "H:\HDrive\Hedge fund\Operations\Operations\Daily - Compass Trading Levels.xlsm"
    Sheets("Compass PL").Activate
    Range("A48").Select
    Selection.End(xlToRight).Select
    Selection.Copy
    
     wbPortfolioCheck.Activate
     Sheets("Calculator").Activate
     Range("F5").Select
     Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
      
     Workbooks("Daily - Compass Trading Levels.xlsm").Sheets("Compass PL").Activate
    Range("A49").Select
    Selection.End(xlToRight).Select
    Selection.Copy
    
   wbPortfolioCheck.Activate
     Sheets("Calculator").Activate
     Range("F6").Select
      Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        
        
     Workbooks("Daily - Compass Trading Levels.xlsm").Sheets("Compass PL").Activate
    Range("XFD32").Select
    Selection.End(xlToLeft).Select
    Selection.Copy
    
     wbPortfolioCheck.Activate
     Sheets("Calculator").Activate
     Range("CC8").Select
     Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
               
       Workbooks("Daily - Compass Trading Levels.xlsm").Sheets("Compass PL").Activate
    Range("XFD33").Select
    Selection.End(xlToLeft).Select
    Selection.Copy
    
     wbPortfolioCheck.Activate
     Sheets("Calculator").Activate
     Range("CC5").Select
     Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        
        
        
        Workbooks("Daily - Compass Trading Levels.xlsm").Activate
        Sheets("Compass PL").Activate
        Range("XFD35").Select
        Selection.End(xlToLeft).Select
        Selection.Copy
        
        Workbooks("Incentive Fee Accrual Summary.xlsx").Activate
        Range("C3").PasteSpecial xlPasteValues
      Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
         
         Workbooks("Daily - Compass Trading Levels.xlsm").Activate
        Sheets("Compass PL").Activate
        Range("XFD36").Select
        Selection.End(xlToLeft).Select
        Selection.Copy
                
           Workbooks("Incentive Fee Accrual Summary.xlsx").Activate
        Range("C4").PasteSpecial xlPasteValues
      Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

        
        Workbooks("Daily - Compass Trading Levels.xlsm").Activate
        Sheets("Compass PL").Activate
        Range("XFD33").Select
        Selection.End(xlToLeft).Select
        Selection.Copy
        
        Workbooks("Incentive Fee Accrual Summary.xlsx").Activate
        Range("D3").PasteSpecial xlPasteValues
      Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
         
         Workbooks("Daily - Compass Trading Levels.xlsm").Activate
        Sheets("Compass PL").Activate
        Range("XFD34").Select
        Selection.End(xlToLeft).Select
        Selection.Copy
                
           Workbooks("Incentive Fee Accrual Summary.xlsx").Activate
        Range("D4").PasteSpecial xlPasteValues
      Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        
        
        
        
        
        'Crown
        
          Workbooks.Open Filename:= _
        "H:\HDrive\Hedge fund\Operations\Operations\Daily - Crown Trading Level.xlsm"
    Sheets("CROWN PL").Activate
    Range("A16").Select
    Selection.End(xlToRight).Select
    Selection.Copy
    
    
     wbPortfolioCheck.Activate
     Sheets("Calculator").Activate
     Range("F3").Select
     Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        
        Workbooks("Daily - Crown Trading Level.xlsm").Activate
        Sheets("CROWN PL").Activate
        Range("XFD12").Select
        Selection.End(xlToLeft).Select
        Selection.Copy
                
           Workbooks("Incentive Fee Accrual Summary.xlsx").Activate
        Range("C5").PasteSpecial xlPasteValues
      Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
 
   Workbooks("Daily - Crown Trading Level.xlsm").Activate
        Sheets("CROWN PL").Activate
        Range("XFD14").Select
        Selection.End(xlToLeft).Select
        Selection.Copy
                
           Workbooks("Incentive Fee Accrual Summary.xlsx").Activate
        Range("D5").PasteSpecial xlPasteValues
      Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
                 
        
        
        
        
        'MSI
        
               Workbooks.Open Filename:= _
        "H:\HDrive\Hedge fund\Operations\Operations\Daily - MSI Trading Level.xlsm"
    Sheets("MSI PL").Activate
    Range("A18").Select
    Selection.End(xlToRight).Select
    Selection.Copy
    
       wbPortfolioCheck.Activate
     Sheets("Calculator").Activate
     Range("F7").PasteSpecial xlPasteValues
      Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
  
             Workbooks("Daily - MSI Trading Level.xlsm").Activate
        Sheets("MSI PL").Activate
        Range("XFD12").Select
        Selection.End(xlToLeft).Select
        Selection.Copy
        
        Workbooks("Incentive Fee Accrual Summary.xlsx").Activate
        Range("C8").PasteSpecial xlPasteValues
      Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
         
         Workbooks("Daily - MSI Trading Level.xlsm").Activate
        Sheets("MSI PL").Activate
        Range("XFD14").Select
        Selection.End(xlToLeft).Select
        Selection.Copy
        
        
           Workbooks("Incentive Fee Accrual Summary.xlsx").Activate
        Range("D8").PasteSpecial xlPasteValues
      Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
             
    
    
    'IBIS 1
    
    
     Workbooks.Open Filename:= _
        "H:\HDrive\Hedge fund\Operations\Operations\Daily - IBIS GMF Share Class NAVs.xlsx"
    Sheets(1).Select
      Range("XFD48").Select
    Selection.End(xlToLeft).Select
    Selection.Copy
    
        
     wbPortfolioCheck.Activate
     Sheets("Calculator").Activate
     Range("F2").Select
      Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        
    
        
         Workbooks("Daily - IBIS GMF Share Class NAVs.xlsx").Activate
        Sheets(1).Select
        Range("XFD48").Select
        Selection.End(xlToLeft).Select
        Selection.End(xlToLeft).Select
        Selection.Copy
                
           Workbooks("Incentive Fee Accrual Summary.xlsx").Activate
        Range("D6").PasteSpecial xlPasteValues
      Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
   
   'IBIS 2
   
    
 '   Workbooks.Open Filename:= _
  '      "H:\HDrive\Hedge fund\Operations\Operations\Daily - IBIS GMF II Share Class NAVs.xlsx"
   ' Sheets(1).Select
    '  Range("XFD54").Select
    'Selection.End(xlToLeft).Select
    'Selection.Copy
    
        
     'wbPortfolioCheck.Activate
     'Sheets("Calculator").Activate
     'Range("E4").Select
      'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
       ' :=False, Transpose:=False
        
               
        
      '  Workbooks("Daily - IBIS GMF II Share Class NAVs.xlsx").Activate
       ' Sheets(1).Select
       ' Range("XFD50").Select
       ' Selection.End(xlToLeft).Select
       ' Selection.End(xlToLeft).Select
       ' Selection.Copy
        
        '
         '  Workbooks("Incentive Fee Accrual Summary.xlsx").Activate
        'Range("D7").PasteSpecial xlPasteValues
      'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
       ' :=False, Transpose:=False
        
        
        
        
        'close the workbooks
        
        
                Workbooks("Daily - MSI Trading Level.xlsm").Close Savechanges:=False
                Workbooks("Daily - Crown Trading Level.xlsm").Close Savechanges:=False
                Workbooks("Daily - Compass Trading Levels.xlsm").Close Savechanges:=False
        '        Workbooks("Daily - IBIS GMF II Share Class NAVs.xlsx").Close Savechanges:=False
                Workbooks("Daily - IBIS GMF Share Class NAVs.xlsx").Close Savechanges:=False
                Workbooks("Incentive Fee Accrual Summary.xlsx").Close Savechanges:=True
                
                Application.DisplayAlerts = True
                wbPortfolioCheck.Activate
                  Sheets("Summary").Select
                
        
   
        
      
                
                     Dim myCell As Range, rng As Range
        Set rng = Range("C61:C110")
        For Each myCell In rng
             If IsError(myCell) Then
                myCell.EntireRow.Hidden = True
            End If
        Next myCell
    

      
      
    
    
                
                
                MsgBox ("Portfolio Check completed for all accounts.")
                Sheets("Summary").Select
                
                
End Sub
