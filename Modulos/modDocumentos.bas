' ============================================
' modDocumentos - Operações dos documentos
' ============================================

Option Explicit

Sub SalvarDocumentos(dados As Collection)
    Dim ws As Worksheet
    Dim linha As Long
    
    Set ws = Sheets("Documentos")
    
    linha = BuscarLinhaDocumentos(dados("ID_Cliente"))
    
    ws.Cells(linha, 3).Value = dados("CNIS")
    ws.Cells(linha, 4).Value = dados("PPP")
    ws.Cells(linha, 5).Value = dados("LTCAT")
    ws.Cells(linha, 6).Value = dados("CTPS")
    ws.Cells(linha, 7).Value = dados("Certidoes")
    ws.Cells(linha, 8).Value = dados("Comprovantes")
    ws.Cells(linha, 9).Value = dados("Observacoes")
End Sub
