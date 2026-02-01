Attribute VB_Name = "LIMPEZA"
Sub limpar()
    Dim wbAtual As Workbook
    Dim wsOrd As Worksheet, wsCons As Worksheet, wsAbs As Worksheet
    
    Set wbAtual = ThisWorkbook
    Set wsOrd = wbAtual.Sheets("ORDENS")
    Set wsCons = wbAtual.Sheets("CONSOLIDADO")
    Set wsAbs = wbAtual.Sheets("ABSOLUTO")

    wsCons.Range("M1:R" & wsCons.Rows.Count).ClearContents
    wsOrd.Range("A2:B" & wsOrd.Rows.Count).ClearContents
    wsOrd.Range("D3:F" & wsOrd.Rows.Count).ClearContents
    wsCons.Range("A2:F" & wsCons.Rows.Count).Interior.ColorIndex = xlNone
    wsAbs.Range("A2:J" & wsAbs.Rows.Count).Interior.ColorIndex = xlNone
    wsAbs.Range("A2:J" & wsAbs.Rows.Count).ClearContents
    
End Sub

