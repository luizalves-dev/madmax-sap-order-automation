Attribute VB_Name = "DIGITADOR"
Sub Digitador()
    
    Dim SAPApp
    Dim Connection
    Dim session
    Dim SapGuiAuto
    Dim wbAtual As Workbook
    Dim wsOrdens As Worksheet, wsCons As Worksheet, wsAbs As Worksheet
    Dim ultimaLinha As Long, ultimaMat As Long
    Dim mat As String, remessa As String, numPed As String, formPgto As String, centro As String
    Dim comeco As Long, alcance As Long, final As Long
    Dim dataFinal As String
    Dim matriculas() As Variant
    Dim ordem As String, ordem2 As String, ordem3 As String, ordem4 As String
    Dim pallet As String
    Dim alv As Object
    Dim SkusSap As Object
    Dim skus() As Variant
    Dim contsup As Long, contsupabs As Long
    Dim textoErro As String
    Dim textoSKUNP As String
    Dim textoSKU() As String
    Dim skunp As String
    Dim Grid As Object
    Dim i As Long, v As Long, j As Long, p As Long, s As Long, h As Long

    Set wbAtual = ThisWorkbook
    Set wsOrdens = wbAtual.Sheets("ORDENS")
    Set wsCons = wbAtual.Sheets("CONSOLIDADO")
    Set wsAbs = wbAtual.Sheets("ABSOLUTO")
    
    wsOrdens.Range("A2:B" & wsCons.Rows.Count).ClearContents
    wsOrdens.Range("D3:F" & wsOrdens.Rows.Count).ClearContents
    wsCons.Range("A2:G" & wsCons.Rows.Count).Interior.ColorIndex = xlNone
    wsCons.Range("M1:R" & wsCons.Rows.Count).ClearContents
    wsAbs.Range("A2:J" & wsAbs.Rows.Count).ClearContents
    
    ultimaLinha = wsCons.Cells(wsCons.Rows.Count, "D").End(xlUp).Row
    
    If ultimaLinha <= 1 Then
        Exit Sub
    End If
    
    wsCons.Cells(1, 13).Formula2 = "=UNIQUE(CONSOLIDADO!D2:D" & ultimaLinha & ")"
    ultimaMat = wsCons.Cells(wsCons.Rows.Count, "M").End(xlUp).Row
    
    With wsCons
        .Cells(1, 14).Formula2 = "=XLOOKUP(M1, $D$2:$D$" & ultimaLinha & ", $G$2:$G$" & ultimaLinha & ", ""ERROR"", 0)"
        'data de entrega
        .Cells(1, 15).Formula2 = "=XLOOKUP(M1, $D$2:$D$" & ultimaLinha & ", $C$2:$C$" & ultimaLinha & ", ""ERROR"", 0)"
        'numero do pedido
        .Cells(1, 16).Formula2 = "=XLOOKUP(M1, $D$2:$D$" & ultimaLinha & ", $B$2:$B$" & ultimaLinha & ", ""ERROR"", 0)"
        'forma de pagamento
        .Cells(1, 17).Formula2 = "=XMATCH(M1, D:D, 0)"
        'comeco
        .Cells(1, 18).Formula2 = "=XLOOKUP(M1, $D$2:$D$" & ultimaLinha & ", $A$2:$A$" & ultimaLinha & ", ""ERROR"", 0)"
        'centro
        .Cells(1, 19).Formula2 = "=COUNTIF($D$2:$D$" & ultimaLinha & ", M1)"
        'alcance
    End With
    
    If ultimaMat >= 2 Then
        wsCons.Range("N1:S" & ultimaMat).FillDown
    End If
    
    ReDim matriculas(1 To ultimaMat, 1 To 7)
    matriculas = wsCons.Range("M1:S" & ultimaMat).Value
        
    wsAbs.Range("A2:C" & ultimaLinha).Value = wsCons.Range("B2:D" & ultimaLinha).Value
    wsAbs.Range("D2:D" & ultimaLinha).Value = wsCons.Range("A2:A" & ultimaLinha).Value
    wsAbs.Range("E2:F" & ultimaLinha).Value = wsCons.Range("E2:F" & ultimaLinha).Value
    wsAbs.Range("G2:J" & ultimaLinha).Value = "-"
    
    If Not IsObject(SAPApp) Then
       Set SapGuiAuto = GetObject("SAPGUI")
       Set SAPApp = SapGuiAuto.GetScriptingEngine
    End If
    If Not IsObject(Connection) Then
       Set Connection = SAPApp.Children(0)
    End If
    If Not IsObject(session) Then
       Set session = Connection.Children(0)
    End If
    If IsObject(WScript) Then
       WScript.ConnectObject session, "on"
       WScript.ConnectObject Application, "on"
    End If
    
    p = 2
    v = 3
    For i = 1 To ultimaMat
        mat = matriculas(i, 1)
        remessa = matriculas(i, 2)
        numPed = matriculas(i, 3)
        formPgto = matriculas(i, 4)
        comeco = matriculas(i, 5)
        centro = CStr(matriculas(i, 6))
        alcance = matriculas(i, 7)

        final = (alcance + comeco - 1)
        
        session.findById("wnd[0]").resizeWorkingPane 135, 37, False
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/NZSDCAPTURABR"
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]/usr/ctxtE_PARAM100-KUNNR").Text = mat
        session.findById("wnd[0]/usr/ctxtE_PARAM100-KUNNR").caretPosition = 10
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]/usr/cmbE_PARAM100-ZTRANSCOM").SetFocus
        session.findById("wnd[0]/usr/cmbE_PARAM100-ZTRANSCOM").Key = "01"
        session.findById("wnd[0]/usr/cmbE_PARAM100-ZESCENARIO").SetFocus
        
        If formPgto = "A VISTA" Then
            session.findById("wnd[0]/usr/cmbE_PARAM100-ZESCENARIO").Key = "01"
        Else
            On Error Resume Next
            session.findById("wnd[0]/usr/cmbE_PARAM100-ZESCENARIO").Key = "29"
            If Err.Number <> 0 Then
                wsOrdens.Cells(p, 1).Value = mat
                wsOrdens.Cells(p, 2).Value = "A VISTA"
                wsAbs.Range("J" & comeco & ":J" & final).Value = "CLIENTE NÃO POSSUI FORMA DE PAGAMENTO A BOLETO"
                Err.Clear
                p = p + 1
                On Error GoTo 0
                GoTo Proximo
            End If
            On Error GoTo 0
        End If
        
        session.findById("wnd[0]").sendVKey 8
        
        If session.Children.Count = 2 Then
            If olhaAPedra(session, "wnd[1]/usr/lbl[0,0]") Then
                wsOrdens.Cells(p, 1).Value = mat
                wsOrdens.Cells(p, 2).Value = "ITEM ABERTO"
                wsAbs.Range("J" & comeco & ":J" & final).Value = "ITEM ABERTO"
                p = p + 1
                GoTo Proximo
            End If
            If olhaAPedra(session, "wnd[1]/usr/txtMESSTXT2") Then
                wsOrdens.Cells(p, 1).Value = mat
                wsOrdens.Cells(p, 2).Value = "INADIMPLENTE"
                wsAbs.Range("J" & comeco & ":J" & final).Value = "INADIMPLENTE"
                p = p + 1
                GoTo Proximo
            End If
        End If
        
        session.findById("wnd[0]").sendVKey 8
        
        If numPed <> "0" Then
            session.findById("wnd[0]/usr/txtE_PARAM200-BSTNK").Text = numPed
            session.findById("wnd[0]/usr/txtE_PARAM200-BSTNK").SetFocus
            session.findById("wnd[0]/usr/txtE_PARAM200-BSTNK").caretPosition = 10
            session.findById("wnd[0]").sendVKey 0
        End If
        
        session.findById("wnd[0]/usr/txtE_PARAM200-FECHA_ENT").Text = remessa
        session.findById("wnd[0]/usr/txtE_PARAM200-FECHA_ENT").SetFocus
        session.findById("wnd[0]/usr/txtE_PARAM200-FECHA_ENT").caretPosition = 10
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]").sendVKey 0
        
        dataFinal = session.findById("wnd[0]/usr/txtE_PARAM200-FECHA_ENT").Text
        
        wsAbs.Range("I" & comeco & ":I" & final).Value = dataFinal
        
        If centro <> "" And centro <> "0" And centro <> " " Then
            session.findById("wnd[0]/usr/cmbE_PARAM200-ZCENTRO").Key = centro
            session.findById("wnd[0]/usr/cmbE_PARAM200-ZCENTRO").SetFocus
            session.findById("wnd[0]").sendVKey 0
        End If
        
        session.findById("wnd[0]/usr/tabsTAB_FICHAS/tabpTAB_FICHAS_FC11/ssubTAB_FICHAS_SCA:ZSDOPBRM001:0310/btnE_PARAM310-BTNCAPTURA").press
        wsCons.Range("E" & comeco & ":F" & final).Copy
        wsCons.Range("E" & comeco & ":F" & final).Copy
        session.findById("wnd[1]/tbar[0]/btn[11]").press
        session.findById("wnd[1]/tbar[0]/btn[8]").press
        
        session.findById("wnd[0]/usr/tabsTAB_FICHAS/tabpTAB_FICHAS_FC11/ssubTAB_FICHAS_SCA:ZSDOPBRM001:0310/cntlCONTEINER310/shellcont/shell").setCurrentCell -1, "UNIDAD"
        session.findById("wnd[0]/usr/tabsTAB_FICHAS/tabpTAB_FICHAS_FC11/ssubTAB_FICHAS_SCA:ZSDOPBRM001:0310/cntlCONTEINER310/shellcont/shell").selectColumn "UNIDAD"
        session.findById("wnd[0]/usr/tabsTAB_FICHAS/tabpTAB_FICHAS_FC11/ssubTAB_FICHAS_SCA:ZSDOPBRM001:0310/cntlCONTEINER310/shellcont/shell").contextMenu
        session.findById("wnd[0]/usr/tabsTAB_FICHAS/tabpTAB_FICHAS_FC11/ssubTAB_FICHAS_SCA:ZSDOPBRM001:0310/cntlCONTEINER310/shellcont/shell").selectContextMenuItem "&FILTER"
        session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").Text = "1"
        session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-HIGH").Text = "100000"
        session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-HIGH").SetFocus
        session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-HIGH").caretPosition = 6
        session.findById("wnd[1]/tbar[0]/btn[0]").press
        
        Set alv = session.findById("wnd[0]/usr/tabsTAB_FICHAS/tabpTAB_FICHAS_FC11/ssubTAB_FICHAS_SCA:ZSDOPBRM001:0310/cntlCONTEINER310/shellcont/shell")
        
        ReDim skus(comeco To final, 1 To 2)
        For s = comeco To final
            skus(s, 1) = CStr(wsCons.Cells(s, 5).Value)
            skus(s, 2) = True
        Next s
        
        skus = pintaClone(skus, comeco, final)
        
        For s = comeco To final
            If Not skus(s, 2) Then
                wsCons.Cells(s, 5).Interior.Color = RGB(255, 252, 120)
            End If
        Next s
        
        Set SkusSap = CreateObject("Scripting.Dictionary")
        For s = 0 To alv.RowCount - 1
            alv.firstVisibleRow = s
            SkusSap.Add alv.GetCellValue(s, "MATNR"), s
        Next s
        
        contsup = 0
        For s = comeco To final
            If Not SkusSap.Exists(skus(s, 1)) Then
                wsOrdens.Cells(v, 4).Value = mat
                wsOrdens.Cells(v, 5).Value = skus(s, 1)
                wsOrdens.Cells(v, 6).Value = wsCons.Cells(s, 6).Value
                wsAbs.Cells(s, 8).Value = "SUPRESSÃO"
                v = v + 1
                contsup = contsup + 1
            Else
                SkusSap(skus(s, 1)) = s
            End If
        Next s
        
        If contsup = alcance Then
            wsOrdens.Cells(p, 1).Value = mat
            wsOrdens.Cells(p, 2).Value = "SUPRESSÃO"
            wsAbs.Range("J" & comeco & ":J" & final).Value = "SUPRESSÃO TOTAL"
            p = p + 1
            GoTo Proximo
        End If
        
        pallet = session.findById("wnd[0]/usr/tabsTAB_FICHAS/tabpTAB_FICHAS_FC11/ssubTAB_FICHAS_SCA:ZSDOPBRM001:0310/txtE_PARAM310-PALLETS").Text
    
        If CLng(pallet) > 6 Then
            wsOrdens.Cells(p, 1).Value = mat
            wsOrdens.Cells(p, 2).Value = "PALLETS"
            p = p + 1
            GoTo Proximo
        End If
        
        session.findById("wnd[0]/tbar[0]/btn[11]").press
        
        opa = olhaAPedra(session, "wnd[1]/usr/lbl[5,2]") And Not olhaAPedra(session, "wnd[1]/usr/cntlCONTEINER204/shellcont/shell")
        Do While opa = True
            textoErro = session.findById("wnd[1]/usr/lbl[5,2]").Text
            If InStr(1, textoErro, "Falha", vbTextCompare) > 0 Then
                
                session.findById("wnd[1]").sendVKey 0
                session.findById("wnd[1]/tbar[0]/btn[71]").press
                session.findById("wnd[2]/usr/chkSCAN_STRING-START").Selected = False
                session.findById("wnd[2]/usr/txtRSYSF-STRING").Text = "material"
                session.findById("wnd[2]/usr/txtRSYSF-STRING").caretPosition = 8
                session.findById("wnd[2]/tbar[0]/btn[0]").press
                    
                textoSKUNP = session.findById("wnd[3]/usr/lbl[17,2]").Text
                textoSKU = Split(textoSKUNP, " ")
                skunp = textoSKU(1)
                    
                session.findById("wnd[3]").sendVKey 12
                session.findById("wnd[2]").sendVKey 12
                session.findById("wnd[1]").sendVKey 12
                    
                alv.setCurrentCell -1, "MATNR"
                alv.selectColumn "MATNR"
                alv.contextMenu
                alv.selectContextMenuItem "&FILTER"
                session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").Text = skunp
                session.findById("wnd[1]/tbar[0]/btn[0]").press
                    
                wsOrdens.Cells(v, 4).Value = mat
                wsOrdens.Cells(v, 5).Value = skunp
                wsOrdens.Cells(v, 6).Value = alv.GetCellValue(0, "UNIDAD")
                v = v + 1
                
                wsAbs.Cells(SkusSap(skunp), 8).Value = "SUPRESSÃO"
                wsAbs.Cells(SkusSap(skunp), 10).Value = "SKU NÃO PERMITIDO"
                    
                alv.modifyCell 0, "UNIDAD", "0"
                alv.currentCellColumn = "UNIDAD"
                alv.pressEnter
                     
                session.findById("wnd[0]/tbar[0]/btn[11]").press
          
            Else
                wsOrdens.Cells(p, 1).Value = mat
                wsOrdens.Cells(p, 2).Value = "PEDIDO MINIMO"
                wsAbs.Range("J" & comeco & ":J" & final).Value = "PEDIDO MINIMO"
                p = p + 1
                GoTo Proximo
            End If
            opa = olhaAPedra(session, "wnd[1]/usr/lbl[5,2]") And Not olhaAPedra(session, "wnd[1]/usr/cntlCONTEINER204/shellcont/shell")
        Loop
        
        If contsup = alcance Then
            wsOrdens.Cells(p, 1).Value = mat
            wsOrdens.Cells(p, 2).Value = "SUPRESSÃO"
            wsAbs.Range("J" & comeco & ":J" & final).Value = "SUPRESSÃO"
            p = p + 1
            GoTo Proximo
        End If
        
        If olhaAPedra(session, "wnd[1]/tbar[0]/btn[12]") = True And olhaAPedra(session, "wnd[1]/usr/cntlCONTEINER204/shellcont/shell") = False Then
            session.findById("wnd[1]/tbar[0]/btn[12]").press
        End If
        
        If olhaAPedra(session, "wnd[1]/usr/btnBUTTON_1") = True And olhaAPedra(session, "wnd[1]/usr/cntlCONTEINER204/shellcont/shell") = False Then
            session.findById("wnd[1]/usr/btnBUTTON_1").press
        End If
        
        If olhaAPedra(session, "wnd[1]/usr/cntlCONTEINER204/shellcont/shell") = True Then
            Set Grid = session.findById("wnd[1]/usr/cntlCONTEINER204/shellcont/shell")
            If Grid.RowCount = 1 Then
                ordem = Grid.GetCellValue(0, "VBELN")
                wsOrdens.Cells(p, 1).Value = mat
                wsOrdens.Cells(p, 2).Value = ordem
                wsAbs.Range("G" & comeco & ":G" & final).Value = ordem
                p = p + 1
            ElseIf Grid.RowCount = 2 Then
                ordem = Grid.GetCellValue(0, "VBELN")
                ordem2 = Grid.GetCellValue(1, "VBELN")
                wsOrdens.Range("A" & p & ":A" & p + 1).Value = mat
                wsOrdens.Cells(p, 2).Value = ordem
                wsOrdens.Cells(p + 1, 2).Value = ordem2
                wsAbs.Range("G" & comeco & ":G" & final).Value = ordem & " / " & ordem2
                p = p + 2
            ElseIf Grid.RowCount = 3 Then
                ordem = Grid.GetCellValue(0, "VBELN")
                ordem2 = Grid.GetCellValue(1, "VBELN")
                ordem3 = Grid.GetCellValue(2, "VBELN")
                wsOrdens.Range("A" & p & ":A" & p + 2).Value = mat
                wsOrdens.Cells(p, 2).Value = ordem
                wsOrdens.Cells(p + 1, 2).Value = ordem2
                wsOrdens.Cells(p + 2, 2).Value = ordem3
                wsAbs.Range("G" & comeco & ":G" & final).Value = ordem & " / " & ordem2 & " / " & ordem3
                p = p + 3
            ElseIf Grid.RowCount = 4 Then
                ordem = Grid.GetCellValue(0, "VBELN")
                ordem2 = Grid.GetCellValue(1, "VBELN")
                ordem3 = Grid.GetCellValue(2, "VBELN")
                ordem4 = Grid.GetCellValue(3, "VBELN")
                wsOrdens.Range("A" & p & ":A" & p + 3).Value = mat
                wsOrdens.Cells(p, 2).Value = ordem
                wsOrdens.Cells(p + 1, 2).Value = ordem2
                wsOrdens.Cells(p + 2, 2).Value = ordem3
                wsOrdens.Cells(p + 3, 2).Value = ordem4
                wsAbs.Range("G" & comeco & ":G" & final).Value = ordem & " / " & ordem2 & " / " & ordem3 & " / " & ordem4
                p = p + 4
            End If
        End If

Proximo:
    Next i
    
End Sub

Public Function olhaAPedra(ByVal session As Object, ByVal targetId As String) As Boolean

    Dim el As Object
    On Error Resume Next
    Set el = session.findById(targetId)
    olhaAPedra = (Err.Number = 0 And Not el Is Nothing)
    Err.Clear
    On Error GoTo 0
    
End Function

Public Function pintaClone(m1() As Variant, c As Long, f As Long) As Variant()

    Dim i As Long
    Dim base As Object

    Set base = CreateObject("Scripting.Dictionary")
    For i = c To f
        If Not base.Exists(m1(i, 1)) Then
            base.Add m1(i, 1), i
        Else
            m1(base(m1(i, 1)), 2) = False
            m1(i, 2) = False
        End If
    Next i
    
    pintaClone = m1

End Function

