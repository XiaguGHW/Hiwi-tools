Attribute VB_Name = "Farben_Makros"
Option Explicit

' =====================================================================
'  Farben fuer die Projektseiten (Projektbericht) - Version 2
'  Verwendet die DESIGNFARBEN von Excel (wie in der Farbpalette ausgewaehlt):
'    Kopfzeile  = Blau, Akzent 5, heller 60 %  + Schrift FRUTIGER, fett
'    Budget     = Gruen, Akzent 6, heller 60 %
'    Obligo     = Blau, Akzent 1, heller 60 %
'    IST        = Orange, Akzent 2, heller 60 %
'    PLAN       = Gold, Akzent 4, heller 60 %
'    Block 3    = Weiss, Hintergrund 1, dunkler 15 %  (Betraege: Gruen, Akzent 6, heller 60 %)
'
'  Bedienung: Bereich markieren (mehrere Bereiche mit Strg) -> Makro starten
'  Achtung: Makro-Aenderungen lassen sich NICHT mit Strg+Z rueckgaengig machen.
'           Vorher eine Kopie der Datei speichern.
' =====================================================================

' ---- Schriftart der Kopfzeile: genauen Namen mit Makro "FarbeAnzeigen" pruefen
Private Const SCHRIFT_KOPF As String = "Frutiger LT 45 Light"

Private Const HELLER_60 As Double = 0.599993896298105   ' Excel-Wert fuer "heller 60 %"
Private Const DUNKLER_15 As Double = -0.149998474074526 ' Excel-Wert fuer "dunkler 15 %"

Private Function AuswahlOK() As Boolean
    If TypeName(Selection) <> "Range" Then
        MsgBox "Bitte zuerst einen Zellbereich markieren.", vbExclamation
        AuswahlOK = False
    Else
        AuswahlOK = True
    End If
End Function

Private Sub Designfarbe(ByVal ziel As Range, ByVal akzent As Long)
    With ziel.Interior
        .Pattern = xlSolid
        .ThemeColor = akzent
        .TintAndShade = HELLER_60
    End With
End Sub

Private Sub FaerbeAuswahl(ByVal akzent As Long)
    If Not AuswahlOK Then Exit Sub
    Designfarbe Selection, akzent
End Sub

' ---------------- Block 2 ----------------
Sub F_Kopfzeile()
    ' Zeile mit "Gesamtbudget", "IST Ausgaben", "PLAN Ausgaben 2026" ... (z. B. A9:O9)
    If Not AuswahlOK Then Exit Sub
    Designfarbe Selection, xlThemeColorAccent5
    Selection.Font.Name = SCHRIFT_KOPF
    Selection.Font.Bold = True
End Sub

Sub F_Budget()
    ' 3 Zeilen unter "Gesamtbudget" + Zeile "gesamt" + gruener Teil in Spalte O
    FaerbeAuswahl xlThemeColorAccent6
End Sub

Sub F_Obligo()
    ' Block unter "Obligo <Jahr>"
    FaerbeAuswahl xlThemeColorAccent1
End Sub

Sub F_IST()
    ' Block unter "Ausg. IST" / "IST Ausgaben"
    FaerbeAuswahl xlThemeColorAccent2
End Sub

Sub F_PLAN()
    ' Bloecke unter "PLAN <Zeitraum>" ... "PLAN Ausgaben Projektende"
    FaerbeAuswahl xlThemeColorAccent4
End Sub

' ---------------- Block 3: "IST - Vorjahre" bis "Projektende" ----------------
Sub F_Restbudget()
    ' Markieren: diese Zeilen von Spalte A bis EINSCHLIESSLICH der Spalte mit den Betraegen
    ' (z. B. A19:F21). Links: Weiss, Hintergrund 1, dunkler 15 %; Betragsspalte: Gruen, Akzent 6, heller 60 %.
    ' Schrift und Zahlenformat bleiben unveraendert.
    Dim bereich As Range, n As Long
    If Not AuswahlOK Then Exit Sub
    For Each bereich In Selection.Areas
        n = bereich.Columns.Count
        If n < 2 Then
            MsgBox "Bitte von Spalte A bis zur Betragsspalte markieren (mind. 2 Spalten).", vbExclamation
            Exit Sub
        End If
        With bereich.Resize(, n - 1).Interior
            .Pattern = xlSolid
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = DUNKLER_15
        End With
        Designfarbe bereich.Columns(n), xlThemeColorAccent6
    Next bereich
End Sub

' ---------------- Legende ----------------
Sub F_Legende()
    ' Fuegt ab der aktiven Zelle die Legende ein (z. B. Zelle O18 unten rechts).
    ' Reihenfolge wie in der Vorlage: Budget / Restbudget, (Obligo), IST, PLAN.
    ' Die Zeile "Obligo" nur, wenn die Seite einen Obligo-Block hat.
    Dim z As Range, mitObligo As Boolean, zeilen As Long, i As Long
    Set z = ActiveCell
    mitObligo = (MsgBox("Legende MIT Obligo?" & vbCrLf & vbCrLf & _
                        "Ja    = Budget / Obligo / IST / PLAN" & vbCrLf & _
                        "Nein  = Budget / IST / PLAN", _
                        vbYesNo + vbQuestion, "Legende") = vbYes)
    zeilen = IIf(mitObligo, 5, 4)
    If Application.WorksheetFunction.CountA(z.Resize(zeilen, 1)) > 0 Then
        If MsgBox("Die Zellen sind nicht leer (z. B. alte Legende). Ueberschreiben?", _
                  vbYesNo + vbQuestion, "Legende") = vbNo Then Exit Sub
        z.Resize(zeilen, 1).Interior.Pattern = xlNone   ' alte Farben entfernen
    End If
    z.Value = "Legende"
    z.Font.Bold = True
    i = 1
    LegendenZeile z.Offset(i, 0), "Budget / Restbudget", xlThemeColorAccent6: i = i + 1
    If mitObligo Then
        LegendenZeile z.Offset(i, 0), "Obligo (gebunden)", xlThemeColorAccent1: i = i + 1
    End If
    LegendenZeile z.Offset(i, 0), "IST Ausgaben (gebucht)", xlThemeColorAccent2: i = i + 1
    LegendenZeile z.Offset(i, 0), "PLAN Ausgaben (Prognose)", xlThemeColorAccent4
    ' Rest einer alten, laengeren Legende entfernen (nur wenn dort wirklich Legendentext steht)
    If Not mitObligo Then
        If LCase$(CStr(z.Offset(zeilen, 0).Value)) Like "plan ausgaben*" Then
            z.Offset(zeilen, 0).ClearContents
            z.Offset(zeilen, 0).Interior.Pattern = xlNone
        End If
    End If
End Sub

Private Sub LegendenZeile(ByVal ziel As Range, ByVal txt As String, ByVal akzent As Long)
    ziel.Value = txt
    Designfarbe ziel, akzent
End Sub

Sub F_LegendeNeuFaerben()
    ' Bestehende Legende behalten, nur die Farben anpassen:
    ' die Zellen der Legende markieren (ohne "Legende"-Ueberschrift) und Makro starten.
    ' Die Farbe wird anhand des Textes gewaehlt.
    Dim c As Range, t As String
    If Not AuswahlOK Then Exit Sub
    For Each c In Selection.Cells
        t = LCase$(Trim$(CStr(c.Value)))
        If t Like "budget*" Or t Like "restbudget*" Then
            Designfarbe c, xlThemeColorAccent6
        ElseIf t Like "obligo*" Then
            Designfarbe c, xlThemeColorAccent1
        ElseIf t Like "ist*" Then
            Designfarbe c, xlThemeColorAccent2
        ElseIf t Like "plan*" Then
            Designfarbe c, xlThemeColorAccent4
        End If
    Next c
End Sub

' ---------------- Hilfen ----------------
Sub FarbeAnzeigen()
    ' Zeigt Farbe und Schrift der aktiven Zelle an (zum Vergleich mit der Vorlage)
    Dim c As Range, f As Long, info As String
    Set c = ActiveCell
    f = c.Interior.Color
    info = "Schriftart: " & c.Font.Name & vbCrLf & _
           "Fett: " & IIf(c.Font.Bold, "ja", "nein") & vbCrLf & _
           "RGB: " & (f Mod 256) & ", " & ((f \ 256) Mod 256) & ", " & (f \ 65536) & vbCrLf
    On Error Resume Next
    info = info & "Designfarbe (Akzent): " & (c.Interior.ThemeColor - 4) & vbCrLf & _
                  "Helligkeit: " & Format(c.Interior.TintAndShade * 100, "0") & " %"
    On Error GoTo 0
    MsgBox info, vbInformation, "Zelle " & c.Address(False, False)
End Sub

Sub Tastenkuerzel_An()
    Dim p As String
    p = "'" & ThisWorkbook.Name & "'!"
    Application.OnKey "^+K", p & "F_Kopfzeile"
    Application.OnKey "^+B", p & "F_Budget"
    Application.OnKey "^+O", p & "F_Obligo"
    Application.OnKey "^+I", p & "F_IST"
    Application.OnKey "^+P", p & "F_PLAN"
    Application.OnKey "^+R", p & "F_Restbudget"
    MsgBox "Tastenkuerzel aktiv:" & vbCrLf & _
           "Strg+Umschalt+K  Kopfzeile (Blau Akzent 5, Frutiger, fett)" & vbCrLf & _
           "Strg+Umschalt+B  Budget (Gruen Akzent 6)" & vbCrLf & _
           "Strg+Umschalt+O  Obligo (Blau Akzent 1)" & vbCrLf & _
           "Strg+Umschalt+I  IST (Orange Akzent 2)" & vbCrLf & _
           "Strg+Umschalt+P  PLAN (Gold Akzent 4)" & vbCrLf & _
           "Strg+Umschalt+R  IST-Vorjahre bis Projektende (grau + Betraege gruen)", vbInformation
End Sub

Sub Tastenkuerzel_Aus()
    Application.OnKey "^+K"
    Application.OnKey "^+B"
    Application.OnKey "^+O"
    Application.OnKey "^+I"
    Application.OnKey "^+P"
    Application.OnKey "^+R"
End Sub
