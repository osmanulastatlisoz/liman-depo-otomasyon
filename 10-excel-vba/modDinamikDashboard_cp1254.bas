Attribute VB_Name = "modDinamikDashboard"
Option Explicit

' ===============================================================
'  DINAMIK DASHBOARD - TCH STOK TAKIP
'  Veri kaynagi : STOK HAREKETLERI
'  Liste kaynagi: MALZEME LISTESI
'  Giris/Cikis  : GIRIS+IADE -> giris, CIKIS -> cikis, TRANSFER haric
'  Kurulum      : KurDinamikDashboard
'  Yenile       : DashYenile  (YENILE butonu)
' ===============================================================

Private DASH As String
Private VTAB As String
Private STAB As String
Private OZET As String
Private SRCSH As String
Private MLIST As String
Private TUMU As String
Private GIRISK As String
Private CIKISK As String
Private IADEK As String

Private Sub SabitleriKur()
    DASH = "DÝNAMÝK DASHBOARD"
    VTAB = "VERÝ TABANI"
    STAB = "DASH AYAR"
    OZET = "ÖZET VERÝ"
    SRCSH = "STOK HAREKETLERÝ"
    MLIST = "MALZEME LÝSTESÝ"
    TUMU = "TÜMÜ"
    GIRISK = "GÝRÝÞ"
    CIKISK = "ÇIKIÞ"
    IADEK = "ÝADE"
End Sub

' ================================================================
'  KURULUM (tek seferlik)
' ================================================================
Public Sub KurDinamikDashboard()
    SabitleriKur
    Dim eskiCalc As XlCalculation
    eskiCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    SayfaHazirla VTAB
    SayfaHazirla STAB
    SayfaHazirla OZET
    SayfaHazirla DASH

    AyarSayfasiKur
    NamesKur
    DashboardKur
    OzetSayfasiKur

    DashYenile
    GrafikleriYap
    ButonlariYap

    ThisWorkbook.Sheets(STAB).Visible = xlSheetHidden
    On Error Resume Next
    ThisWorkbook.Sheets(DASH).Move Before:=ThisWorkbook.Sheets(1)
    On Error GoTo 0

    ThisWorkbook.Sheets(DASH).Activate
    ThisWorkbook.Sheets(DASH).Range("A1").Select

    Application.Calculation = eskiCalc
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

Private Sub SayfaHazirla(ad As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(ad)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = ad
    Else
        ws.Visible = xlSheetVisible
        Dim co As ChartObject, i As Long
        For Each co In ws.ChartObjects: co.Delete: Next co
        For i = ws.Shapes.Count To 1 Step -1: ws.Shapes(i).Delete: Next i
        ws.Cells.Clear
    End If
End Sub

' ================================================================
'  DASH AYAR (gizli yardimci sayfa)
' ================================================================
Private Sub AyarSayfasiKur()
    Dim wa As Worksheet: Set wa = ThisWorkbook.Sheets(STAB)
    wa.Cells.Clear
    Dim q As String: q = Chr(34)
    wa.Range("A1").Value = "AKTÝF FÝLTRE (ARA butonu doldurur - elle deðiþtirmeyin)"
    Dim lbl As Variant
    lbl = Array("AY", "MALZEME TÜRÜ", "SINIF", "KULLANIM YERÝ", "MALZEME ADI")
    Dim i As Long, rr As Long
    For i = 0 To 4
        rr = 2 + i
        wa.Cells(rr, 1).Value = lbl(i)
        wa.Cells(rr, 2).Value = TUMU
        wa.Cells(rr, 4).Formula = "=IF(B" & rr & "=" & q & TUMU & q & "," & q & "*" & q & ",B" & rr & ")"
    Next i
    wa.Range("F1").Value = "AY": wa.Range("G1").Value = GIRISK: wa.Range("H1").Value = CIKISK
    wa.Range("J1").Value = "SINIF": wa.Range("K1").Value = CIKISK
    wa.Range("M1").Value = "SINIF SEÇENEKLERÝ"
    wa.Range("N1").Value = "MALZEME ADI SEÇENEKLERÝ"
    wa.Range("P1").Value = "TÜR SEÇENEKLERÝ"
    wa.Range("Q1").Value = "YER SEÇENEKLERÝ"
    wa.Range("R1").Value = "AY SEÇENEKLERÝ"
    wa.Rows(1).Font.Bold = True
End Sub

Private Sub NamesKur()
    AddName "LISTE_AY", "='" & STAB & "'!$R$2:$R$2"
    AddName "LISTE_TUR", "='" & STAB & "'!$P$2:$P$2"
    AddName "LISTE_YER", "='" & STAB & "'!$Q$2:$Q$2"
    AddName "LISTE_SINIF", "='" & STAB & "'!$M$2:$M$2"
    AddName "LISTE_ADI", "='" & STAB & "'!$N$2:$N$2"
End Sub

Private Sub AddName(nm As String, ref As String)
    On Error Resume Next
    ThisWorkbook.Names(nm).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=nm, RefersTo:=ref
End Sub

' ================================================================
'  DASHBOARD ARAYUZU
' ================================================================
Private Sub DashboardKur()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(DASH)
    ws.Cells.Clear
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1: ws.Shapes(i).Delete: Next i

    ws.Columns("A").ColumnWidth = 16
    ws.Columns("B").ColumnWidth = 11
    ws.Columns("C").ColumnWidth = 28
    ws.Columns("D").ColumnWidth = 24
    ws.Columns("E").ColumnWidth = 18
    ws.Columns("F").ColumnWidth = 13
    ws.Columns("G").ColumnWidth = 20
    ws.Columns("H").ColumnWidth = 18
    ws.Columns("I").ColumnWidth = 15
    ws.Columns("J").ColumnWidth = 11
    ws.Columns("K").ColumnWidth = 11
    ws.Columns("L").ColumnWidth = 11

    With ws.Range("A1:L2")
        .Merge
        .Value = "TCH STOK YÖNETÝMÝ  •  DÝNAMÝK DASHBOARD"
        .Font.Size = 20: .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 60, 95)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With

    ws.Range("A3:B3").Merge: ws.Range("A3").Value = "AY"
    ws.Range("D3:E3").Merge: ws.Range("D3").Value = "MALZEME TÜRÜ"
    ws.Range("G3:H3").Merge: ws.Range("G3").Value = "SINIF"
    ws.Range("A5:B5").Merge: ws.Range("A5").Value = "KULLANIM YERÝ"
    ws.Range("D5:E5").Merge: ws.Range("D5").Value = "MALZEME ADI"
    ws.Range("A4:B4").Merge: ws.Range("A4").Value = TUMU
    ws.Range("D4:E4").Merge: ws.Range("D4").Value = TUMU
    ws.Range("G4:H4").Merge: ws.Range("G4").Value = TUMU
    ws.Range("A6:B6").Merge: ws.Range("A6").Value = TUMU
    ws.Range("D6:E6").Merge: ws.Range("D6").Value = TUMU

    StilEtiket ws.Range("A3:B3"): StilEtiket ws.Range("D3:E3"): StilEtiket ws.Range("G3:H3")
    StilEtiket ws.Range("A5:B5"): StilEtiket ws.Range("D5:E5")
    StilSecim ws.Range("A4:B4"): StilSecim ws.Range("D4:E4"): StilSecim ws.Range("G4:H4")
    StilSecim ws.Range("A6:B6"): StilSecim ws.Range("D6:E6")

    DvKur ws.Range("A4"), "LISTE_AY"
    DvKur ws.Range("D4"), "LISTE_TUR"
    DvKur ws.Range("G4"), "LISTE_SINIF"
    DvKur ws.Range("A6"), "LISTE_YER"
    DvKur ws.Range("D6"), "LISTE_ADI", False

    ws.Range("A8").Value = "TOPLAM GÝRÝÞ"
    ws.Range("C8").Value = "TOPLAM ÇIKIÞ"
    ws.Range("E8").Value = "HAREKET SAYISI"
    ws.Range("G8").Value = "NET (GÝRÝÞ-ÇIKIÞ)"
    ws.Range("I8").Value = "SEÇÝLEN KRÝTER"
    ws.Range("A9").Formula = KpiSumifs("J")
    ws.Range("C9").Formula = KpiSumifs("K")
    ws.Range("E9").Formula = KpiCountifs()
    ws.Range("G9").Formula = "=A9-C9"
    ws.Range("I9").Formula = KpiKriter()
    StilKpi ws

    ws.Range("A10:L11").Merge
    ws.Range("A10").Value = "Kullaným:  Açýlýr listelerden filtreleri seçin, ARA'ya basýn. Üst kriter (Malzeme Türü, Sýnýf, Malzeme Adý) deðiþince alt listeler otomatik daralýr. TEMÝZLE filtreleri sýfýrlar. YENÝLE stok hareketlerini yeniden okur."
    ws.Range("A10").WrapText = True
    ws.Range("A10").Font.Italic = True
    ws.Range("A10").VerticalAlignment = xlCenter
    ws.Range("A10").Interior.Color = RGB(245, 247, 250)

    ws.Range("A12:L12").Merge
    ws.Range("A12").Value = "SEÇÝME GÖRE FÝLTRELENEN HAREKETLER"
    ws.Range("A12").Font.Bold = True: ws.Range("A12").Font.Size = 12
    ws.Range("A12").Interior.Color = RGB(31, 60, 95): ws.Range("A12").Font.Color = RGB(255, 255, 255)
    ws.Range("A12").HorizontalAlignment = xlCenter

    Dim th As Variant
    th = Array("AY", "ÝÞLEM TÜRÜ", "MALZEME ADI", "PLAKA / AÇIKLAMA", "MÝKTAR", "PERSONEL", "TARÝH", "MALZEME TÜRÜ", "SINIF", "KULLANIM YERÝ", "GÝRÝÞ", "ÇIKIÞ")
    For i = 0 To 11: ws.Cells(13, i + 1).Value = th(i): Next i
    With ws.Range("A13:L13")
        .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(68, 95, 130)
        .HorizontalAlignment = xlCenter
    End With

    ws.Rows("3:6").RowHeight = 22
    ws.Rows("8").RowHeight = 18
    ws.Rows("9").RowHeight = 26

    ws.Activate
    On Error Resume Next
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.FreezePanes = False
    ws.Range("A14").Select
    ActiveWindow.FreezePanes = True
    ws.Range("A1").Select
    On Error GoTo 0
End Sub

Private Sub StilEtiket(r As Range)
    r.Font.Bold = True: r.Font.Color = RGB(255, 255, 255)
    r.Interior.Color = RGB(68, 95, 130)
    r.HorizontalAlignment = xlCenter: r.VerticalAlignment = xlCenter
End Sub

Private Sub StilSecim(r As Range)
    r.Font.Bold = True: r.Font.Size = 11: r.Font.Color = RGB(20, 20, 20)
    r.Interior.Color = RGB(255, 248, 220)
    r.HorizontalAlignment = xlCenter: r.VerticalAlignment = xlCenter
    r.Borders.LineStyle = xlContinuous: r.Borders.Color = RGB(180, 180, 180)
End Sub

Private Sub StilKpi(ws As Worksheet)
    Dim cards As Variant: cards = Array("A", "C", "E", "G", "I")
    Dim cols As Variant
    cols = Array(RGB(215, 240, 220), RGB(252, 220, 220), RGB(225, 235, 248), RGB(255, 240, 220), RGB(235, 235, 235))
    Dim i As Long, lab As Range, val As Range
    For i = 0 To 4
        Set lab = ws.Range(cards(i) & "8")
        Set val = ws.Range(cards(i) & "9")
        lab.Font.Bold = True: lab.Font.Size = 10: lab.HorizontalAlignment = xlCenter
        lab.Interior.Color = RGB(68, 95, 130): lab.Font.Color = RGB(255, 255, 255)
        val.Font.Bold = True: val.Font.Size = 14: val.HorizontalAlignment = xlCenter
        val.Interior.Color = cols(i)
        val.Borders.LineStyle = xlContinuous: val.Borders.Color = RGB(180, 180, 180)
        If cards(i) <> "I" Then
            val.NumberFormat = "#,##0.##"
        Else
            val.Font.Size = 9: val.WrapText = True
        End If
    Next i
End Sub

Private Sub DvKur(rng As Range, nm As String, Optional katiUyari As Boolean = True)
    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=" & nm
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = katiUyari
    End With
End Sub

Private Function KpiSumifs(col As String) As String
    KpiSumifs = "=SUMIFS('" & VTAB & "'!$" & col & "$2:$" & col & "$31000,'" & VTAB & "'!$A$2:$A$31000,'" & STAB & "'!$D$2,'" & VTAB & "'!$G$2:$G$31000,'" & STAB & "'!$D$3,'" & VTAB & "'!$H$2:$H$31000,'" & STAB & "'!$D$4,'" & VTAB & "'!$I$2:$I$31000,'" & STAB & "'!$D$5,'" & VTAB & "'!$C$2:$C$31000,'" & STAB & "'!$D$6)"
End Function

Private Function KpiCountifs() As String
    KpiCountifs = "=COUNTIFS('" & VTAB & "'!$A$2:$A$31000,'" & STAB & "'!$D$2,'" & VTAB & "'!$G$2:$G$31000,'" & STAB & "'!$D$3,'" & VTAB & "'!$H$2:$H$31000,'" & STAB & "'!$D$4,'" & VTAB & "'!$I$2:$I$31000,'" & STAB & "'!$D$5,'" & VTAB & "'!$C$2:$C$31000,'" & STAB & "'!$D$6)"
End Function

Private Function KpiKriter() As String
    Dim q As String: q = Chr(34)
    Dim sep As String: sep = q & " | " & q
    KpiKriter = "='" & STAB & "'!$B$2&" & sep & "&'" & STAB & "'!$B$3&" & sep & "&'" & STAB & "'!$B$4&" & sep & "&'" & STAB & "'!$B$5&" & sep & "&'" & STAB & "'!$B$6"
End Function

' ================================================================
'  OZET VERI sayfasi
' ================================================================
Private Sub OzetSayfasiKur()
    Dim oz As Worksheet: Set oz = ThisWorkbook.Sheets(OZET)
    oz.Cells.Clear
    oz.Range("A1").Value = "ÖZET VERÝ (otomatik güncellenir)"
    oz.Range("A1").Font.Size = 16: oz.Range("A1").Font.Bold = True: oz.Range("A1").Font.Color = RGB(31, 60, 95)
    oz.Range("A3").Value = "AYLIK GÝRÝÞ / ÇIKIÞ"
    oz.Range("A4").Value = "AY": oz.Range("B4").Value = GIRISK: oz.Range("C4").Value = CIKISK
    oz.Range("E3").Value = "MALZEME TÜRÜNE GÖRE ÇIKIÞ"
    oz.Range("E4").Value = "MALZEME TÜRÜ": oz.Range("F4").Value = CIKISK
    oz.Range("A20").Value = "SINIFA GÖRE ÇIKIÞ"
    oz.Range("A21").Value = "SINIF": oz.Range("B21").Value = CIKISK
    oz.Range("E20").Value = "KULLANIM YERÝNE GÖRE ÇIKIÞ"
    oz.Range("E21").Value = "KULLANIM YERÝ": oz.Range("F21").Value = CIKISK
    oz.Range("A3,E3,A20,E20").Font.Bold = True
    oz.Range("A3,E3,A20,E20").Font.Size = 12
    oz.Range("A3,E3,A20,E20").Font.Color = RGB(31, 60, 95)
    oz.Range("A4:C4,E4:F4,A21:B21,E21:F21").Font.Bold = True
    oz.Range("A4:C4,E4:F4,A21:B21,E21:F21").Interior.Color = RGB(68, 95, 130)
    oz.Range("A4:C4,E4:F4,A21:B21,E21:F21").Font.Color = RGB(255, 255, 255)
    oz.Columns("A").ColumnWidth = 20: oz.Columns("E").ColumnWidth = 22
    oz.Columns("B:C").ColumnWidth = 12: oz.Columns("F").ColumnWidth = 12
End Sub

' ================================================================
'  DASH YENILE  (veriyi stok hareketlerinden yeniden olustur)
' ================================================================
Public Sub DashYenile()
    SabitleriKur
    Dim adim As String: adim = "baslangic"
    On Error GoTo EH
    Dim eskiCalc As XlCalculation: eskiCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = False

    Dim src As Worksheet, vt As Worksheet, ml As Worksheet, wa As Worksheet
    Set src = ThisWorkbook.Sheets(SRCSH)
    Set vt = ThisWorkbook.Sheets(VTAB)
    Set ml = ThisWorkbook.Sheets(MLIST)
    Set wa = ThisWorkbook.Sheets(STAB)

    ' --- 1) VERI TABANI <- STOK HAREKETLERI ---
    adim = "1-stok-oku"
    Dim sson As Long: sson = src.Cells(src.Rows.Count, 5).End(xlUp).Row
    Dim cnt As Long: cnt = 0
    If sson >= 2 Then
        Dim s As Variant: s = src.Range("A2:N" & sson).Value
        Dim n As Long: n = UBound(s, 1)
        Dim outv() As Variant: ReDim outv(1 To n, 1 To 12)
        Dim i As Long, tip As String, ad As String, miktar As Double
        For i = 1 To n
            ad = Trim(CStr(s(i, 5)))
            If ad <> "" Then
                cnt = cnt + 1
                tip = IslemNormal(s(i, 3))
                miktar = 0
                If IsNumeric(s(i, 6)) Then miktar = CDbl(s(i, 6))
                outv(cnt, 1) = AyAnahtari(s(i, 1))
                outv(cnt, 2) = tip
                outv(cnt, 3) = ad
                outv(cnt, 4) = miktar
                outv(cnt, 5) = s(i, 4)
                outv(cnt, 6) = s(i, 1)
                outv(cnt, 7) = s(i, 12)
                outv(cnt, 8) = s(i, 14)
                outv(cnt, 9) = s(i, 13)
                outv(cnt, 12) = s(i, 7)
                If tip = GIRISK Or tip = IADEK Then
                    outv(cnt, 10) = miktar: outv(cnt, 11) = 0
                ElseIf tip = CIKISK Then
                    outv(cnt, 10) = 0: outv(cnt, 11) = miktar
                Else
                    outv(cnt, 10) = 0: outv(cnt, 11) = 0
                End If
            End If
        Next i
        Dim vson As Long: vson = vt.Cells(vt.Rows.Count, 1).End(xlUp).Row
        If vson < 2 Then vson = 2
        Dim temizSon As Long: temizSon = vson
        If cnt + 1 > temizSon Then temizSon = cnt + 1
        vt.Range("A2:L" & temizSon).ClearContents
        If cnt > 0 Then vt.Range("A2").Resize(n, 12).Value = outv
    End If
    BasliklariYazVTAB vt

    ' --- 2) Statik listeler ---
    adim = "2-listeler"
    Dim mson As Long: mson = ml.Cells(ml.Rows.Count, 1).End(xlUp).Row
    Dim dTur As Object: Set dTur = CreateObject("Scripting.Dictionary")
    Dim dYer As Object: Set dYer = CreateObject("Scripting.Dictionary")
    If mson >= 2 Then
        Dim m As Variant: m = ml.Range("A2:F" & mson).Value
        Dim j As Long, dd As String
        For j = 1 To UBound(m, 1)
            dd = Trim(CStr(m(j, 4))): If dd <> "" And Not dTur.Exists(dd) Then dTur.Add dd, 1
            dd = Trim(CStr(m(j, 6))): If dd <> "" And Not dYer.Exists(dd) Then dYer.Add dd, 1
        Next j
    End If
    ListeYaz wa, 16, "LISTE_TUR", dTur
    ListeYaz wa, 17, "LISTE_YER", dYer

    Dim dAy As Object: Set dAy = CreateObject("Scripting.Dictionary")
    If cnt > 0 Then
        Dim va As Variant: va = vt.Range("A2:A" & (cnt + 1)).Value
        For i = 1 To UBound(va, 1)
            dd = Trim(CStr(va(i, 1)))
            If dd <> "" And Not dAy.Exists(dd) Then dAy.Add dd, 1
        Next i
    End If
    ListeYaz wa, 18, "LISTE_AY", dAy

    ' --- 3) Aylik grafik kaynagi F/G/H ---
    adim = "3-aylik-grafik"
    wa.Range("F2:H10000").ClearContents
    Dim ayKeys() As String, na As Long
    na = SortedKeys(dAy, ayKeys)
    For i = 1 To na
        Dim r As Long: r = i + 1
        wa.Cells(r, 6).Value = ayKeys(i)
        wa.Cells(r, 7).Formula = "=SUMIFS('" & VTAB & "'!$J$2:$J$31000,'" & VTAB & "'!$A$2:$A$31000,$F" & r & ",'" & VTAB & "'!$G$2:$G$31000,$D$3,'" & VTAB & "'!$H$2:$H$31000,$D$4,'" & VTAB & "'!$I$2:$I$31000,$D$5,'" & VTAB & "'!$C$2:$C$31000,$D$6)"
        wa.Cells(r, 8).Formula = "=SUMIFS('" & VTAB & "'!$K$2:$K$31000,'" & VTAB & "'!$A$2:$A$31000,$F" & r & ",'" & VTAB & "'!$G$2:$G$31000,$D$3,'" & VTAB & "'!$H$2:$H$31000,$D$4,'" & VTAB & "'!$I$2:$I$31000,$D$5,'" & VTAB & "'!$C$2:$C$31000,$D$6)"
    Next i

    ' --- 4) SINIF & ADI ---
    adim = "4-sinif-adi"
    ListeKur "SINIF", TUMU, TUMU
    ListeKur "ADI", TUMU, TUMU

    ' --- 5) OZET VERI ---
    adim = "5-ozet"
    OzetKur

    ' sadece dashboard sayfalarini hesapla (agir DEPO formullerini tetikleme)
    wa.Calculate
    ThisWorkbook.Sheets(DASH).Calculate
    Application.Calculation = eskiCalc
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
EH:
    Application.Calculation = eskiCalc
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Yenileme sirasinda hata olustu." & vbCrLf & _
           "Adim: " & adim & vbCrLf & "Hata " & Err.Number & ": " & Err.Description, _
           vbExclamation, "Dinamik Dashboard"
End Sub

Private Sub BasliklariYazVTAB(vt As Worksheet)
    Dim h As Variant
    h = Array("AY", "ÝÞLEM TÜRÜ", "MALZEME ADI", "MÝKTAR", "PERSONEL", "TARÝH", "MALZEME TÜRÜ", "SINIF", "KULLANIM YERÝ", "GÝRÝÞ", "ÇIKIÞ", "PLAKA / AÇIKLAMA")
    Dim i As Long
    For i = 0 To 11: vt.Cells(1, i + 1).Value = h(i): Next i
    vt.Rows(1).Font.Bold = True
    vt.Rows(1).Interior.Color = RGB(68, 95, 130)
    vt.Rows(1).Font.Color = RGB(255, 255, 255)
End Sub

' ================================================================
'  OZET tablolari
' ================================================================
Private Sub OzetKur()
    SabitleriKur
    Dim vt As Worksheet, oz As Worksheet
    Set vt = ThisWorkbook.Sheets(VTAB)
    Set oz = ThisWorkbook.Sheets(OZET)
    oz.Range("A5:C1000").ClearContents
    oz.Range("E5:F1000").ClearContents
    oz.Range("A22:B1000").ClearContents
    oz.Range("E22:F1000").ClearContents
    Dim vson As Long: vson = vt.Cells(vt.Rows.Count, 1).End(xlUp).Row
    If vson < 2 Then Exit Sub
    Dim arr As Variant: arr = vt.Range("A2:K" & vson).Value
    Dim dAy As Object, dAyC As Object, dTur As Object, dSinif As Object, dYer As Object
    Set dAy = CreateObject("Scripting.Dictionary")
    Set dAyC = CreateObject("Scripting.Dictionary")
    Set dTur = CreateObject("Scripting.Dictionary")
    Set dSinif = CreateObject("Scripting.Dictionary")
    Set dYer = CreateObject("Scripting.Dictionary")
    Dim i As Long, ay As String, tur As String, sinif As String, yer As String
    Dim gir As Double, cik As Double
    For i = 1 To UBound(arr, 1)
        ay = Trim(CStr(arr(i, 1)))
        gir = 0: cik = 0
        If IsNumeric(arr(i, 10)) Then gir = CDbl(arr(i, 10))
        If IsNumeric(arr(i, 11)) Then cik = CDbl(arr(i, 11))
        If ay <> "" Then
            dAy(ay) = dAy(ay) + gir
            dAyC(ay) = dAyC(ay) + cik
        End If
        tur = Trim(CStr(arr(i, 7))): If tur <> "" Then dTur(tur) = dTur(tur) + cik
        sinif = Trim(CStr(arr(i, 8))): If sinif <> "" Then dSinif(sinif) = dSinif(sinif) + cik
        yer = Trim(CStr(arr(i, 9))): If yer <> "" Then dYer(yer) = dYer(yer) + cik
    Next i

    Dim keys() As String, nk As Long
    nk = SortedKeys(dAy, keys)
    For i = 1 To nk
        oz.Cells(4 + i, 1).Value = keys(i)
        oz.Cells(4 + i, 2).Value = dAy(keys(i))
        oz.Cells(4 + i, 3).Value = dAyC(keys(i))
    Next i
    nk = SortedByValueDesc(dTur, keys)
    For i = 1 To nk: oz.Cells(4 + i, 5).Value = keys(i): oz.Cells(4 + i, 6).Value = dTur(keys(i)): Next i
    nk = SortedByValueDesc(dSinif, keys)
    For i = 1 To nk: oz.Cells(20 + i, 1).Value = keys(i): oz.Cells(20 + i, 2).Value = dSinif(keys(i)): Next i
    nk = SortedByValueDesc(dYer, keys)
    For i = 1 To nk: oz.Cells(20 + i, 5).Value = keys(i): oz.Cells(20 + i, 6).Value = dYer(keys(i)): Next i

    oz.Range("B5:C1000").NumberFormat = "#,##0.##"
    oz.Range("F5:F1000").NumberFormat = "#,##0.##"
    oz.Range("B22:B1000").NumberFormat = "#,##0.##"
End Sub

' ================================================================
'  KADEMELI LISTE KURUCU  (MALZEME LISTESI'nden)
' ================================================================
Public Sub ListeKur(hangi As String, turSec As String, Optional sinifSec As String = "")
    SabitleriKur
    Dim ml As Worksheet, wa As Worksheet
    Set ml = ThisWorkbook.Sheets(MLIST)
    Set wa = ThisWorkbook.Sheets(STAB)
    Dim son As Long: son = ml.Cells(ml.Rows.Count, 1).End(xlUp).Row
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    If son >= 2 Then
        Dim arr As Variant: arr = ml.Range("A2:F" & son).Value
        Dim i As Long, deger As String
        For i = 1 To UBound(arr, 1)
            If Esit(arr(i, 4), turSec) Then
                If UCase(hangi) = "SINIF" Then
                    deger = Trim(CStr(arr(i, 5)))
                    If deger <> "" And Not dict.Exists(deger) Then dict.Add deger, 1
                Else
                    If Esit(arr(i, 5), sinifSec) Then
                        deger = Trim(CStr(arr(i, 1)))
                        If deger <> "" And Not dict.Exists(deger) Then dict.Add deger, 1
                    End If
                End If
            End If
        Next i
    End If
    If UCase(hangi) = "SINIF" Then
        ListeYaz wa, 13, "LISTE_SINIF", dict
    Else
        ListeYaz wa, 14, "LISTE_ADI", dict
    End If
End Sub

Private Sub ListeYaz(wa As Worksheet, col As Long, nm As String, dict As Object)
    wa.Range(wa.Cells(2, col), wa.Cells(wa.Rows.Count, col)).ClearContents
    Dim ks() As String, n As Long
    n = SortedKeys(dict, ks)
    wa.Cells(2, col).Value = TUMU
    Dim i As Long
    For i = 1 To n: wa.Cells(2 + i, col).Value = ks(i): Next i
    Dim lastRow As Long: lastRow = 2 + n
    Dim cl As String: cl = SutunHarfi(col)
    On Error Resume Next
    ThisWorkbook.Names(nm).RefersTo = "='" & STAB & "'!$" & cl & "$2:$" & cl & "$" & lastRow
    On Error GoTo 0
End Sub

' ================================================================
'  ARA BUTONU
' ================================================================
Public Sub AraButonu()
    SabitleriKur
    Dim ws As Worksheet, wa As Worksheet, vt As Worksheet
    Set ws = ThisWorkbook.Sheets(DASH)
    Set wa = ThisWorkbook.Sheets(STAB)
    Set vt = ThisWorkbook.Sheets(VTAB)

    Dim eskiCalc As XlCalculation: eskiCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = False

    wa.Range("B2").Value = NormVal(ws.Range("A4").Value)
    wa.Range("B3").Value = NormVal(ws.Range("D4").Value)
    wa.Range("B4").Value = NormVal(ws.Range("G4").Value)
    wa.Range("B5").Value = NormVal(ws.Range("A6").Value)
    wa.Range("B6").Value = NormVal(ws.Range("D6").Value)

    Dim eskiSon As Long: eskiSon = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If eskiSon < 14 Then eskiSon = 14
    ws.Range("A14:L" & eskiSon).Clear
    wa.Range("J2:K11").ClearContents

    Dim son As Long: son = vt.Cells(vt.Rows.Count, 1).End(xlUp).Row
    If son >= 2 Then
        Dim arr As Variant: arr = vt.Range("A2:L" & son).Value
        Dim fAy As String, fTur As String, fSinif As String, fYer As String, fAd As String
        fAy = CStr(wa.Range("B2").Value): fTur = CStr(wa.Range("B3").Value)
        fSinif = CStr(wa.Range("B4").Value): fYer = CStr(wa.Range("B5").Value)
        fAd = CStr(wa.Range("B6").Value)
        Dim n As Long, i As Long, jj As Long, k As Long
        n = UBound(arr, 1)
        k = 0
        For i = 1 To n
            If Esit(arr(i, 1), fAy) And Esit(arr(i, 7), fTur) And Esit(arr(i, 8), fSinif) _
               And Esit(arr(i, 9), fYer) And Esit(arr(i, 3), fAd) Then
                k = k + 1
            End If
        Next i
        Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
        Dim sCls As String, qty As Double
        If k > 0 Then
            Dim out() As Variant: ReDim out(1 To k, 1 To 12)
            Dim rr As Long: rr = 0
            For i = 1 To n
                If Esit(arr(i, 1), fAy) And Esit(arr(i, 7), fTur) And Esit(arr(i, 8), fSinif) _
                   And Esit(arr(i, 9), fYer) And Esit(arr(i, 3), fAd) Then
                    rr = rr + 1
                    out(rr, 1) = arr(i, 1)    ' AY
                    out(rr, 2) = arr(i, 2)    ' ISLEM TURU
                    out(rr, 3) = arr(i, 3)    ' MALZEME ADI
                    out(rr, 4) = arr(i, 12)   ' PLAKA / ACIKLAMA
                    out(rr, 5) = arr(i, 4)    ' MIKTAR
                    out(rr, 6) = arr(i, 5)    ' PERSONEL
                    out(rr, 7) = arr(i, 6)    ' TARIH
                    out(rr, 8) = arr(i, 7)    ' MALZEME TURU
                    out(rr, 9) = arr(i, 8)    ' SINIF
                    out(rr, 10) = arr(i, 9)   ' KULLANIM YERI
                    out(rr, 11) = arr(i, 10)  ' GIRIS
                    out(rr, 12) = arr(i, 11)  ' CIKIS
                    sCls = Trim(CStr(arr(i, 8)))
                    If sCls <> "" Then
                        qty = 0
                        If IsNumeric(arr(i, 11)) Then qty = CDbl(arr(i, 11))
                        dict(sCls) = dict(sCls) + qty
                    End If
                End If
            Next i
            ws.Range("A14").Resize(k, 12).Value = out
        Else
            ws.Range("A14").Value = "Seçilen kritere uygun kayýt bulunamadý."
        End If

        Dim dn As Long: dn = dict.Count
        If dn > 0 Then
            Dim ks() As String, vs() As Double, kk As Variant
            ReDim ks(1 To dn): ReDim vs(1 To dn)
            i = 0
            For Each kk In dict.Keys
                i = i + 1: ks(i) = CStr(kk): vs(i) = dict(kk)
            Next kk
            Dim tD As Double, tS As String
            For i = 1 To dn - 1
                For jj = i + 1 To dn
                    If vs(jj) > vs(i) Then
                        tD = vs(i): vs(i) = vs(jj): vs(jj) = tD
                        tS = ks(i): ks(i) = ks(jj): ks(jj) = tS
                    End If
                Next jj
            Next i
            Dim ust As Long: ust = dn: If ust > 10 Then ust = 10
            For i = 1 To ust
                wa.Cells(1 + i, 10).Value = ks(i)
                wa.Cells(1 + i, 11).Value = vs(i)
            Next i
        End If
    End If

    wa.Calculate
    ws.Calculate
    Application.Calculation = eskiCalc
    TabloRenklendir
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' ================================================================
'  TEMIZLE BUTONU
' ================================================================
Public Sub TemizleButonu()
    SabitleriKur
    Dim ws As Worksheet, wa As Worksheet
    Set ws = ThisWorkbook.Sheets(DASH)
    Set wa = ThisWorkbook.Sheets(STAB)
    Dim eskiCalc As XlCalculation: eskiCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ws.Range("A4").Value = TUMU: ws.Range("D4").Value = TUMU: ws.Range("G4").Value = TUMU
    ws.Range("A6").Value = TUMU: ws.Range("D6").Value = TUMU
    wa.Range("B2").Value = TUMU: wa.Range("B3").Value = TUMU: wa.Range("B4").Value = TUMU
    wa.Range("B5").Value = TUMU: wa.Range("B6").Value = TUMU

    ListeKur "SINIF", TUMU, TUMU
    ListeKur "ADI", TUMU, TUMU

    Dim eskiSon As Long: eskiSon = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If eskiSon >= 14 Then ws.Range("A14:L" & eskiSon).Clear
    wa.Range("J2:K11").ClearContents

    wa.Calculate
    ws.Calculate
    Application.Calculation = eskiCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' ================================================================
'  OTOMATIK MALZEME ADI KONTROLU (D6)  - MALZEME LISTESI'nden
' ================================================================
Public Sub OtomatikAdKontrol(hucre As Range)
    SabitleriKur
    Dim yazilan As String: yazilan = Trim(CStr(hucre.Value))
    If yazilan = "" Or EsittirTumu(yazilan) Then
        hucre.Value = TUMU
        Exit Sub
    End If
    Dim ml As Worksheet: Set ml = ThisWorkbook.Sheets(MLIST)
    Dim son As Long: son = ml.Cells(ml.Rows.Count, 1).End(xlUp).Row
    If son < 2 Then Exit Sub
    Dim arr As Variant: arr = ml.Range("A2:A" & son).Value
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim i As Long, v As String
    For i = 1 To UBound(arr, 1)
        v = Trim(CStr(arr(i, 1)))
        If v <> "" And Not d.Exists(v) Then d.Add v, 1
    Next i
    Dim n As Long: n = d.Count
    If n = 0 Then Exit Sub
    Dim pOrijinal() As String, pNorm() As String
    ReDim pOrijinal(1 To n): ReDim pNorm(1 To n)
    Dim k As Variant: i = 0
    For Each k In d.Keys
        i = i + 1: pOrijinal(i) = CStr(k): pNorm(i) = NormalizeTR(UCase(CStr(k)))
    Next k
    Dim hedef As String: hedef = NormalizeTR(UCase(yazilan))
    Dim j As Long
    For j = 1 To n
        If hedef = pNorm(j) Then
            If yazilan <> pOrijinal(j) Then hucre.Value = pOrijinal(j)
            Exit Sub
        End If
    Next j
    Dim aramaTerimi As String: aramaTerimi = yazilan
ArayipGoster:
    Dim sonucIdx() As Long: ReDim sonucIdx(1 To n)
    Dim sonucSayisi As Long: sonucSayisi = 0
    Dim aramaNorm As String: aramaNorm = NormalizeTR(UCase(aramaTerimi))
    For j = 1 To n
        If InStr(pNorm(j), aramaNorm) > 0 Then
            sonucSayisi = sonucSayisi + 1: sonucIdx(sonucSayisi) = j
        End If
    Next j
    If sonucSayisi = 0 Then
        Dim yeniArama As String
        yeniArama = InputBox("'" & aramaTerimi & "' içeren malzeme bulunamadý." & vbCrLf & vbCrLf & _
                             "Farklý bir kelime yaz" & vbCrLf & "0 = ÝPTAL (TÜMÜ'ye döner)", "Malzeme Bulunamadý")
        If yeniArama = "0" Or yeniArama = "" Then hucre.Value = TUMU: Exit Sub
        aramaTerimi = yeniArama
        GoTo ArayipGoster
    End If
    Dim a As Long, b As Long, gecici As Long
    For a = 1 To sonucSayisi - 1
        For b = a + 1 To sonucSayisi
            If pOrijinal(sonucIdx(b)) < pOrijinal(sonucIdx(a)) Then
                gecici = sonucIdx(a): sonucIdx(a) = sonucIdx(b): sonucIdx(b) = gecici
            End If
        Next b
    Next a
    Const MAX_GOSTER As Long = 25
    Dim gosterSayisi As Long: gosterSayisi = sonucSayisi
    If gosterSayisi > MAX_GOSTER Then gosterSayisi = MAX_GOSTER
    Dim mesaj As String: mesaj = "Aranan: " & Chr(34) & aramaTerimi & Chr(34) & vbCrLf
    If sonucSayisi > MAX_GOSTER Then mesaj = mesaj & "(" & sonucSayisi & " sonuçtan ilk " & MAX_GOSTER & " gösteriliyor)" & vbCrLf
    mesaj = mesaj & vbCrLf
    Dim kk As Long
    For kk = 1 To gosterSayisi
        mesaj = mesaj & "  " & kk & ")  " & pOrijinal(sonucIdx(kk)) & vbCrLf
    Next kk
    mesaj = mesaj & vbCrLf & "Seç (1-" & gosterSayisi & ")" & vbCrLf & _
            "  0 = ÝPTAL (TÜMÜ'ye döner)" & vbCrLf & "  Baþka kelime yazarsan tekrar arar"
    Dim sonucSecim As String: sonucSecim = InputBox(mesaj, "Malzeme Seç")
    If sonucSecim = "0" Or sonucSecim = "" Then hucre.Value = TUMU: Exit Sub
    If IsNumeric(sonucSecim) Then
        Dim sonucNo As Long: sonucNo = Val(sonucSecim)
        If sonucNo >= 1 And sonucNo <= gosterSayisi Then
            hucre.Value = pOrijinal(sonucIdx(sonucNo)): Exit Sub
        End If
    End If
    aramaTerimi = sonucSecim
    GoTo ArayipGoster
End Sub

' ================================================================
'  TABLO RENKLENDIRME
' ================================================================
Private Sub TabloRenklendir()
    SabitleriKur
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(DASH)
    Dim sonSat As Long: sonSat = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If sonSat < 14 Then Exit Sub
    Dim r As Range: Set r = ws.Range("A14:L" & sonSat)
    Application.ScreenUpdating = False
    With r
        .Font.Name = "Calibri": .Font.Size = 9: .Font.Color = RGB(40, 50, 60)
        .VerticalAlignment = xlCenter: .RowHeight = 16
    End With
    Dim bd As Borders: Set bd = r.Borders
    bd(xlEdgeLeft).LineStyle = xlContinuous: bd(xlEdgeLeft).Color = RGB(200, 210, 220)
    bd(xlEdgeRight).LineStyle = xlContinuous: bd(xlEdgeRight).Color = RGB(200, 210, 220)
    bd(xlEdgeTop).LineStyle = xlContinuous: bd(xlEdgeTop).Color = RGB(200, 210, 220)
    bd(xlEdgeBottom).LineStyle = xlContinuous: bd(xlEdgeBottom).Color = RGB(200, 210, 220)
    bd(xlInsideHorizontal).LineStyle = xlContinuous: bd(xlInsideHorizontal).Color = RGB(225, 230, 238)
    bd(xlInsideVertical).LineStyle = xlContinuous: bd(xlInsideVertical).Color = RGB(225, 230, 238)
    ws.Range("A14:A" & sonSat).Interior.Color = RGB(225, 235, 248)   ' AY
    ws.Range("B14:B" & sonSat).Interior.Color = RGB(255, 240, 220)   ' ISLEM TURU
    ws.Range("C14:C" & sonSat).Interior.Color = RGB(255, 252, 225)   ' MALZEME ADI
    ws.Range("D14:D" & sonSat).Interior.Color = RGB(238, 238, 240)   ' PLAKA / ACIKLAMA
    ws.Range("E14:E" & sonSat).Interior.Color = RGB(228, 245, 230)   ' MIKTAR
    ws.Range("F14:F" & sonSat).Interior.Color = RGB(245, 235, 250)   ' PERSONEL
    ws.Range("G14:G" & sonSat).Interior.Color = RGB(235, 245, 252)   ' TARIH
    ws.Range("H14:H" & sonSat).Interior.Color = RGB(252, 235, 240)   ' MALZEME TURU
    ws.Range("I14:I" & sonSat).Interior.Color = RGB(232, 240, 248)   ' SINIF
    ws.Range("J14:J" & sonSat).Interior.Color = RGB(248, 245, 225)   ' KULLANIM YERI
    ws.Range("K14:K" & sonSat).Interior.Color = RGB(215, 240, 220)   ' GIRIS
    ws.Range("L14:L" & sonSat).Interior.Color = RGB(252, 220, 220)   ' CIKIS
    ws.Range("A14:A" & sonSat).HorizontalAlignment = xlCenter
    ws.Range("D14:D" & sonSat).HorizontalAlignment = xlLeft
    ws.Range("E14:E" & sonSat).HorizontalAlignment = xlRight
    ws.Range("E14:E" & sonSat).NumberFormat = "#,##0.##"
    ws.Range("G14:G" & sonSat).HorizontalAlignment = xlCenter
    ws.Range("G14:G" & sonSat).NumberFormat = "yyyy-mm-dd"
    ws.Range("K14:K" & sonSat).HorizontalAlignment = xlRight
    ws.Range("K14:K" & sonSat).NumberFormat = "#,##0.##"
    ws.Range("L14:L" & sonSat).HorizontalAlignment = xlRight
    ws.Range("L14:L" & sonSat).NumberFormat = "#,##0.##"
    Application.ScreenUpdating = True
End Sub

' ================================================================
'  GRAFIKLER
' ================================================================
Public Sub GrafikleriYap()
    SabitleriKur
    Dim ws As Worksheet, wa As Worksheet, oz As Worksheet
    Set ws = ThisWorkbook.Sheets(DASH)
    Set wa = ThisWorkbook.Sheets(STAB)
    Set oz = ThisWorkbook.Sheets(OZET)
    Application.ScreenUpdating = False
    Dim co As ChartObject
    For Each co In ws.ChartObjects: co.Delete: Next co
    For Each co In oz.ChartObjects: co.Delete: Next co

    Dim aySon As Long: aySon = wa.Cells(wa.Rows.Count, 6).End(xlUp).Row
    If aySon < 2 Then aySon = 2

    Dim rng1 As Range: Set rng1 = ws.Range("M3:U21")
    Dim ch1 As ChartObject
    Set ch1 = ws.ChartObjects.Add(rng1.Left, rng1.Top, rng1.Width, rng1.Height)
    ch1.Name = "chDashAylik"
    With ch1.Chart
        .ChartType = xlColumnClustered
        Do While .SeriesCollection.Count > 0: .SeriesCollection(1).Delete: Loop
        With .SeriesCollection.NewSeries
            .Name = "='" & STAB & "'!$G$1"
            .XValues = "='" & STAB & "'!$F$2:$F$" & aySon
            .Values = "='" & STAB & "'!$G$2:$G$" & aySon
        End With
        With .SeriesCollection.NewSeries
            .Name = "='" & STAB & "'!$H$1"
            .XValues = "='" & STAB & "'!$F$2:$F$" & aySon
            .Values = "='" & STAB & "'!$H$2:$H$" & aySon
        End With
        .HasTitle = True: .ChartTitle.Text = "AYLIK GÝRÝÞ / ÇIKIÞ"
        .HasLegend = True
    End With

    Dim rng2 As Range: Set rng2 = ws.Range("M23:U42")
    Dim ch2 As ChartObject
    Set ch2 = ws.ChartObjects.Add(rng2.Left, rng2.Top, rng2.Width, rng2.Height)
    ch2.Name = "chDashSinif"
    With ch2.Chart
        .ChartType = xlBarClustered
        Do While .SeriesCollection.Count > 0: .SeriesCollection(1).Delete: Loop
        With .SeriesCollection.NewSeries
            .Name = "='" & STAB & "'!$K$1"
            .XValues = "='" & STAB & "'!$J$2:$J$11"
            .Values = "='" & STAB & "'!$K$2:$K$11"
        End With
        .HasTitle = True: .ChartTitle.Text = "SINIFA GÖRE ÇIKIÞ (ÝLK 10)"
        .HasLegend = False
    End With

    Dim aEnd As Long: aEnd = 5
    Dim r2 As Long
    For r2 = 5 To 30
        If Trim(CStr(oz.Cells(r2, 1).Value)) = "" Then Exit For
        aEnd = r2
    Next r2
    If aEnd < 5 Then aEnd = 5
    Dim rng3 As Range: Set rng3 = oz.Range("H3:Q22")
    Dim ch3 As ChartObject
    Set ch3 = oz.ChartObjects.Add(rng3.Left, rng3.Top, rng3.Width, rng3.Height)
    ch3.Name = "chOzetAylik"
    With ch3.Chart
        .ChartType = xlColumnClustered
        Do While .SeriesCollection.Count > 0: .SeriesCollection(1).Delete: Loop
        With .SeriesCollection.NewSeries
            .Name = "='" & OZET & "'!$B$4"
            .XValues = "='" & OZET & "'!$A$5:$A$" & aEnd
            .Values = "='" & OZET & "'!$B$5:$B$" & aEnd
        End With
        With .SeriesCollection.NewSeries
            .Name = "='" & OZET & "'!$C$4"
            .XValues = "='" & OZET & "'!$A$5:$A$" & aEnd
            .Values = "='" & OZET & "'!$C$5:$C$" & aEnd
        End With
        .HasTitle = True: .ChartTitle.Text = "AYLIK GÝRÝÞ / ÇIKIÞ (ÖZET)"
        .HasLegend = True
    End With
    Application.ScreenUpdating = True
End Sub

' ================================================================
'  BUTONLAR
' ================================================================
Public Sub ButonlariYap()
    SabitleriKur
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(DASH)
    Application.ScreenUpdating = False
    Dim sh As Shape, i As Long
    For i = ws.Shapes.Count To 1 Step -1
        Set sh = ws.Shapes(i)
        If sh.Name = "btnAra" Or sh.Name = "btnTemizle" Or sh.Name = "btnYenile" Then sh.Delete
    Next i
    ButonYap ws, ws.Range("J3:K4"), "btnAra", "ARA", RGB(67, 160, 71), "AraButonu"
    ButonYap ws, ws.Range("J5:K6"), "btnTemizle", "TEMÝZLE", RGB(245, 124, 0), "TemizleButonu"
    ButonYap ws, ws.Range("J7:K8"), "btnYenile", "YENÝLE", RGB(33, 110, 180), "DashYenile"
    On Error Resume Next
    With ws.Range("D6").Validation
        .ShowError = False
        .ErrorMessage = ""
    End With
    On Error GoTo 0
    Application.ScreenUpdating = True
End Sub

Private Sub ButonYap(ws As Worksheet, hedef As Range, ad As String, metin As String, renk As Long, aksiyon As String)
    Dim btn As Shape
    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, hedef.Left + 2, hedef.Top + 2, hedef.Width - 4, hedef.Height - 4)
    With btn
        .Name = ad
        .Adjustments(1) = 0.3
        .Fill.Visible = msoTrue: .Fill.Solid: .Fill.ForeColor.RGB = renk
        .Line.Visible = msoFalse
        .Shadow.Visible = msoTrue: .Shadow.OffsetX = 1.5: .Shadow.OffsetY = 1.5
        .Shadow.Blur = 4: .Shadow.Transparency = 0.55: .Shadow.ForeColor.RGB = RGB(0, 0, 0)
        With .TextFrame2
            .HorizontalAnchor = msoAnchorCenter: .VerticalAnchor = msoAnchorMiddle
            .MarginLeft = 2: .MarginRight = 2
            With .TextRange
                .Text = metin
                .Font.Name = "Calibri": .Font.Size = 14: .Font.Bold = msoTrue
                .Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
                .ParagraphFormat.Alignment = msoAlignCenter
            End With
        End With
        .OnAction = aksiyon
    End With
End Sub

' ================================================================
'  YARDIMCI FONKSIYONLAR
' ================================================================
Public Function NormalizeTR(s As String) As String
    Dim r As String: r = s
    r = Replace(r, ChrW(199), "C"): r = Replace(r, ChrW(231), "c")
    r = Replace(r, ChrW(350), "S"): r = Replace(r, ChrW(351), "s")
    r = Replace(r, ChrW(286), "G"): r = Replace(r, ChrW(287), "g")
    r = Replace(r, ChrW(220), "U"): r = Replace(r, ChrW(252), "u")
    r = Replace(r, ChrW(214), "O"): r = Replace(r, ChrW(246), "o")
    r = Replace(r, ChrW(304), "I"): r = Replace(r, ChrW(305), "i")
    NormalizeTR = r
End Function

Private Function IslemNormal(v As Variant) As String
    Dim a As String: a = UCase(NormalizeTR(Trim(CStr(v))))
    Select Case a
        Case "GIRIS": IslemNormal = GIRISK
        Case "CIKIS": IslemNormal = CIKISK
        Case "IADE": IslemNormal = IADEK
        Case "TRANSFER": IslemNormal = "TRANSFER"
        Case Else: IslemNormal = Trim(CStr(v))
    End Select
End Function

Private Function AyAnahtari(v As Variant) As String
    On Error GoTo bos
    If IsDate(v) Then
        Dim d As Date: d = CDate(v)
        If Year(d) < 2000 Then GoTo bos
        AyAnahtari = CStr(Year(d)) & "-" & Right$("0" & CStr(Month(d)), 2)
        Exit Function
    End If
bos:
    AyAnahtari = ""
End Function

Private Function NormKw(s As String) As String
    Dim t As String: t = UCase(Trim(s))
    t = Replace(t, ChrW(220), "U")
    NormKw = t
End Function

Private Function EsittirTumu(s As String) As Boolean
    EsittirTumu = (NormKw(s) = "TUMU" Or Trim(s) = "")
End Function

Private Function NormVal(v As Variant) As String
    Dim s As String: s = Trim(CStr(v))
    If EsittirTumu(s) Then NormVal = TUMU Else NormVal = s
End Function

Private Function Esit(hucre As Variant, kriter As String) As Boolean
    If EsittirTumu(kriter) Then
        Esit = True
    Else
        Esit = (Trim(CStr(hucre)) = Trim(kriter))
    End If
End Function

Private Function SutunHarfi(c As Long) As String
    SutunHarfi = Split(ThisWorkbook.Sheets(1).Cells(1, c).Address, "$")(1)
End Function

Private Function SortedKeys(dict As Object, ByRef ks() As String) As Long
    Dim n As Long: n = dict.Count
    If n = 0 Then SortedKeys = 0: Exit Function
    ReDim ks(1 To n)
    Dim i As Long, k As Variant: i = 0
    For Each k In dict.Keys: i = i + 1: ks(i) = CStr(k): Next k
    Dim j As Long, t As String
    For i = 1 To n - 1
        For j = i + 1 To n
            If ks(j) < ks(i) Then t = ks(i): ks(i) = ks(j): ks(j) = t
        Next j
    Next i
    SortedKeys = n
End Function

Private Function SortedByValueDesc(dict As Object, ByRef ks() As String) As Long
    Dim n As Long: n = dict.Count
    If n = 0 Then SortedByValueDesc = 0: Exit Function
    ReDim ks(1 To n)
    Dim vs() As Double: ReDim vs(1 To n)
    Dim i As Long, k As Variant: i = 0
    For Each k In dict.Keys
        i = i + 1: ks(i) = CStr(k): vs(i) = CDbl(dict(k))
    Next k
    Dim j As Long, t As String, td As Double
    For i = 1 To n - 1
        For j = i + 1 To n
            If vs(j) > vs(i) Then
                td = vs(i): vs(i) = vs(j): vs(j) = td
                t = ks(i): ks(i) = ks(j): ks(j) = t
            End If
        Next j
    Next i
    SortedByValueDesc = n
End Function
