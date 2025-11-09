' ---------------------------
' Price Watcher + Telegram (Fiyat değişim bildirimi)
' Hepsiburada sepete özel fiyat + JSON formattedPrice destekli
' ---------------------------
Option Explicit

Private nextRunTime As Date
Private Const DEFAULT_INTERVAL_MINUTES As Long = 1 ' <<< Varsayılan 1 dk

' --- HTTP GET ---
Private Function HttpGet(url As String) As String
    On Error GoTo ErrHandler
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT)"
    http.send
    If http.Status = 200 Then
        HttpGet = http.responseText
    Else
        HttpGet = ""
    End If
    Exit Function
ErrHandler:
    HttpGet = ""
End Function

' --- HTML'den fiyat çekme ---
Private Function ExtractPriceFromHtml(html As String) As Double
    On Error GoTo ErrDefault
    Dim re As Object, m As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    
    Dim raw As String
    
    ' 1) Sepete özel fiyat
    re.Pattern = "<div class=""bWwoI8vknB6COlRVbpRj"">([\d\.,]+)\s*TL</div>"
    Set m = re.Execute(html)
    If m.Count > 0 Then
        raw = Trim(m(0).SubMatches(0))
        GoTo CleanAndConvert
    End If
    
    ' 2) Önceki fiyat (indirim yoksa)
    re.Pattern = "<span class=""uY6qgF91fGtRUWsRau94"">([\d\.,]+)\s*TL</span>"
    Set m = re.Execute(html)
    If m.Count > 0 Then
        raw = Trim(m(0).SubMatches(0))
        GoTo CleanAndConvert
    End If
    
    ' 3) JSON-LD içindeki formattedPrice
    re.Pattern = """formattedPrice""\s*:\s*""([^""]*)"""
    Set m = re.Execute(html)
    If m.Count > 0 Then
        raw = Trim(m(0).SubMatches(0))
        GoTo CleanAndConvert
    End If
    
    ' 4) Alternatif TL metni
    re.Pattern = "([\d]{1,3}(?:[\.\s]\d{3})*(?:,\d+)?)\s*TL"
    Set m = re.Execute(html)
    If m.Count > 0 Then
        raw = Trim(m(0).SubMatches(0))
        GoTo CleanAndConvert
    End If
    
    GoTo ErrDefault

CleanAndConvert:
    ' Nokta ve virgül düzeltme
    raw = Replace(raw, " ", "")
    raw = Replace(raw, ".", "")
    raw = Replace(raw, ",", ".")
    
    ' Başındaki gereksiz sıfırları sil
    Do While Left(raw, 1) = "0" And Len(raw) > 1 And Mid(raw, 2, 1) <> "."
        raw = Mid(raw, 2)
    Loop
    
    Dim price As Double
    price = Val(raw)
    
    ' Büyük fiyat düzeltmesi
    If price > 100000 Then price = price / 10
    If price > 10000 Then price = price / 10
    
    ExtractPriceFromHtml = price
    Exit Function
    
ErrDefault:
    ExtractPriceFromHtml = -1
End Function

' --- Telegram mesaj gönderme ---
Private Function SendTelegramMessage(token As String, chatId As String, message As String) As Boolean
    On Error GoTo ErrHandler
    Dim url As String
    url = "https://api.telegram.org/bot" & token & "/sendMessage"
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    Dim payload As String
    payload = "chat_id=" & chatId & "&text=" & WorksheetFunction.EncodeURL(message)
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    http.send payload
    SendTelegramMessage = (http.Status = 200)
    Exit Function
ErrHandler:
    SendTelegramMessage = False
End Function

' --- Fiyat değişim kontrolü ---
Public Sub CheckPriceOnce()
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Deneme")
    
    Dim token As String: token = Trim(ws.Range("F4").Value)
    Dim chatId As String: chatId = Trim(ws.Range("G4").Value)
    
    If token = "" Or chatId = "" Then
        ws.Range("H1").Value = "Eksik Telegram bilgisi (F4/G4)."
        Exit Sub
    End If
    
    ws.Range("H1").Value = "Son kontrol: " & Now
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim i As Long
    For i = 4 To lastRow
        Dim url As String, productName As String, kisaad As String
        Dim oldPrice As Double, newPrice As Double
        url = Trim(ws.Cells(i, "A").Value)
        productName = Trim(ws.Cells(i, "B").Value)
        oldPrice = Val(ws.Cells(i, "D").Value)
        kisaad = Trim(ws.Cells(i, "I").Value)
        
        If url = "" Then GoTo ContinueLoop
        
        Dim html As String
        html = HttpGet(url)
        If html = "" Then
            ws.Cells(i, "H").Value = "Sayfa alınamadı"
            GoTo ContinueLoop
        End If
        
        newPrice = ExtractPriceFromHtml(html)
        If newPrice < 0 Then
            ws.Cells(i, "H").Value = "Fiyat bulunamadı"
            GoTo ContinueLoop
        End If
        
        ws.Cells(i, "D").Value = newPrice
        
        ' --- Fiyat değiştiyse bildir ---
        If oldPrice > 0 And Abs(newPrice - oldPrice) >= 0.01 Then
            Dim msg As String
            msg = productName & " Ürünün fiyatı değişti!" & vbCrLf & _
                  "Eski fiyat: " & Format(oldPrice, "0.00") & " TL" & vbCrLf & _
                  "Yeni fiyat: " & Format(newPrice, "0.00") & " TL" & vbCrLf & _
                  "Ürüne Git: " & kisaad

            If SendTelegramMessage(token, chatId, msg) Then
                ws.Cells(i, "H").Value = "Fiyat değişti, bildirim gönderildi (" & Now & ")"
            Else
                ws.Cells(i, "H").Value = "Telegram gönderilemedi"
            End If
        ElseIf oldPrice = 0 Then
            ws.Cells(i, "H").Value = "İlk fiyat kaydedildi (" & Format(newPrice, "0.00") & " TL)"
        Else
            ws.Cells(i, "H").Value = "Fiyat aynı (" & Format(newPrice, "0.00") & " TL)"
        End If
        
ContinueLoop:
    Next i
    Exit Sub
ErrHandler:
    ws.Range("H1").Value = "Hata: " & Err.Description
End Sub

' --- Başlat / Durdur ---
Public Sub StartWatcher()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Deneme")
    Dim intervalMinutes As Long
    intervalMinutes = Val(ws.Range("E4").Value)
    If intervalMinutes <= 0 Then intervalMinutes = DEFAULT_INTERVAL_MINUTES
    nextRunTime = Now + TimeSerial(0, intervalMinutes, 0)
    Application.OnTime earliesttime:=nextRunTime, procedure:="RunWatcherTick", schedule:=True
    ws.Range("H1").Value = "Bot başlatıldı (" & intervalMinutes & " dk)"
End Sub

Public Sub StopWatcher()
    On Error Resume Next
    Application.OnTime earliesttime:=nextRunTime, procedure:="RunWatcherTick", schedule:=False
    ThisWorkbook.Sheets("Deneme").Range("H1").Value = "Bot durduruldu."
End Sub

Public Sub RunWatcherTick()
    Call CheckPriceOnce
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Deneme")
    Dim intervalMinutes As Long
    intervalMinutes = Val(ws.Range("E4").Value)
    If intervalMinutes <= 0 Then intervalMinutes = DEFAULT_INTERVAL_MINUTES
    nextRunTime = Now + TimeSerial(0, intervalMinutes, 0)
    Application.OnTime earliesttime:=nextRunTime, procedure:="RunWatcherTick", schedule:=True
End Sub


