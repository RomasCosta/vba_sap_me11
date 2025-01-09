# Código em VBA para automatização de processos em SAP
## Preencher movimentaação ME11

Sub CriarME11()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Planilha1")
    
    Dim ultimaLinha As Long
    ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If Not IsObject(Appl) Then
        Set SapGuiAuto = GetObject("SAPGUI")
        Set Appl = SapGuiAuto.GetScriptingEngine
    End If
    If Not IsObject(Connection) Then
       Set Connection = Appl.Children(0)
    End If
    If Not IsObject(session) Then
       Set session = Connection.Children(0)
    End If
        
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n me11"
    session.findById("wnd[0]").sendVKey 0
    
    For i = 2 To ultimaLinha
    
        On Error Resume Next
        
        'session.findById("wnd[0]/tbar[0]/okcd").Text = "/n me11"
        
        session.findById("wnd[0]/usr/ctxtEINA-MATNR").Text = ws.Cells(i, 1).Value '"66-03416" - material
        session.findById("wnd[0]/usr/ctxtEINA-LIFNR").Text = ws.Cells(i, 2).Value ' "103944" - fornecedor
        session.findById("wnd[0]/usr/ctxtEINE-EKORG").Text = "BR35"
        session.findById("wnd[0]/usr/ctxtEINE-WERKS").Text = "BR35"
        'session.findById("wnd[0]/usr/ctxtEINE-WERKS").SetFocus
        'session.findById("wnd[0]/usr/ctxtEINE-WERKS").caretPosition = 4
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]").sendVKey 0
        
        session.findById("wnd[0]/usr/txtEINE-NORBM").Text = ws.Cells(i, 3).Value '"29428" - MOQ
        session.findById("wnd[0]/usr/txtEINE-MINBM").Text = ws.Cells(i, 3).Value '"29428" - MOQ
        session.findById("wnd[0]/usr/txtEINE-NETPR").Text = ws.Cells(i, 4).Value '"0,01" - preço
        session.findById("wnd[0]/usr/ctxtEINE-WAERS").Text = ws.Cells(i, 5).Value '"BRL" - moeda
        session.findById("wnd[0]/usr/txtEINE-PEINH").Text = ws.Cells(i, 6).Value '"1" - UM. preço
        session.findById("wnd[0]/usr/ctxtEINE-MWSKZ").Text = ws.Cells(i, 7).Value '"I0" - cód.imp
        session.findById("wnd[0]/usr/ctxtEINE-INCO1").Text = "CIF"
        session.findById("wnd[0]/usr/txtEINE-INCO2").Text = "LOUVEIRA"
        'session.findById("wnd[0]/usr/txtEINE-INCO2").SetFocus
        'session.findById("wnd[0]/usr/txtEINE-INCO2").caretPosition = 8
        session.findById("wnd[0]").sendVKey 11
        
        Application.Wait Now + TimeValue("00:00:02")
        
        statusMessage = session.findById("wnd[0]/sbar").Text
        ws.Cells(i, "J").Value = statusMessage
        
        typeMessage = session.findById("wnd[0]/sbar").messagetype
        ws.Cells(i, "K").Value = typeMessage
        
        If typeMessage <> "S" Then
            
            ws.Cells(i, "I").Value = "Erro"
            
            session.findById("wnd[0]/tbar[0]/okcd").Text = "/n me11"
            session.findById("wnd[0]").sendVKey 0
            
        Else
            ws.Cells(i, "I").Value = "Sucesso"
            
        End If
        
    Next i
             
End Sub
