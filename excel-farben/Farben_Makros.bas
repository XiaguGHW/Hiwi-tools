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
Private Const MAX_ZEILE As Long = 80     ' Automatik: so weit nach unten suchen
Private Const MAX_SPALTE As Long = 30   ' Automatik: so weit nach rechts suchen (30 = AD)

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

' =====================================================================
'  AUTOMATIK: erkennt die Bloecke anhand der Beschriftungen
'    - Kopfzeile  = Zeile, in der Spalte A mit "Gesamtbudget" beginnt
'                   -> ganze Zeile Blau Akzent 5, Schrift SCHRIFT_KOPF, fett
'    - Bloecke    = jede Ueberschrift in der Kopfzeile bestimmt die Farbe der
'                   BEREITS GEFAERBTEN Zellen darunter, bis einschliesslich "gesamt":
'                     ...Budget... / ...Mittel... / Gesamt... -> Gruen Akzent 6
'                       (bei Budget/Mittel zusaetzlich alle Zellen mit Zahlen, auch ohne Farbe)
'                     ...Obligo...              -> Blau Akzent 1
'                     ...PLAN... / ...PLN...    -> Gold Akzent 4
'                     ...IST... / ...Ausg...    -> Orange Akzent 2
'                     sonst -> unveraendert, im Bericht gemeldet
'    - Zeile "gesamt" (Spalte A) -> fett
'    - Restbudget = Zeilen unter "gesamt", deren Text in Spalte A mit
'                   IST / PLAN / PLN / Restbudget / Projektmittel beginnt
'                   -> links Weiss dunkler 15 %, erste Betragsspalte Gruen Akzent 6, alles fett
'    - Legende    = Zellen unter "Legende" werden nach ihrem Text umgefaerbt
'  Nach dem Lauf oeffnet sich eine neue Arbeitsmappe mit einem Bericht.
' =====================================================================

Sub Auto_AktivesBlatt()
    ' Zum Testen: nur das gerade angezeigte Blatt bearbeiten
    Dim wb As Workbook
    Set wb = ZielMappe()
    If wb Is Nothing Then Exit Sub
    BerichtAusgeben wb.Name, BlattFormatieren(ActiveSheet)
End Sub

Sub Auto_AlleProjektseiten()
    ' Bearbeitet alle Blaetter ab dem 3. Blatt der aktiven Arbeitsmappe
    Dim wb As Workbook, i As Long, liste As String, bericht As String
    Set wb = ZielMappe()
    If wb Is Nothing Then Exit Sub
    If wb.Worksheets.Count < 3 Then
        MsgBox "Diese Arbeitsmappe hat weniger als 3 Blaetter.", vbExclamation
        Exit Sub
    End If
    For i = 3 To wb.Worksheets.Count
        liste = liste & "  " & wb.Worksheets(i).Name & vbCrLf
    Next i
    If MsgBox("Arbeitsmappe: " & wb.Name & vbCrLf & vbCrLf & _
              "Folgende Blaetter werden bearbeitet:" & vbCrLf & liste & vbCrLf & _
              "Haben Sie eine Sicherungskopie? (Kein Rueckgaengig moeglich)", _
              vbYesNo + vbQuestion, "Automatisch faerben") = vbNo Then Exit Sub

    On Error GoTo Fehler
    Application.ScreenUpdating = False
    For i = 3 To wb.Worksheets.Count
        bericht = bericht & BlattFormatieren(wb.Worksheets(i)) & vbCrLf
    Next i
    Application.ScreenUpdating = True
    BerichtAusgeben wb.Name, bericht
    Exit Sub
Fehler:
    Application.ScreenUpdating = True
    MsgBox "Fehler bei Blatt " & i & ": " & Err.Description & vbCrLf & _
           "Bisher bearbeitet:" & vbCrLf & Left$(bericht, 700), vbCritical
End Sub

Private Function ZielMappe() As Workbook
    Dim wb As Workbook
    On Error Resume Next
    Set wb = ActiveWorkbook
    On Error GoTo 0
    If wb Is Nothing Then
        MsgBox "Keine bearbeitbare Arbeitsmappe aktiv." & vbCrLf & _
               "Bei 'Geschuetzte Ansicht' zuerst 'Bearbeitung aktivieren' klicken.", vbExclamation
        Exit Function
    End If
    If wb.Name = ThisWorkbook.Name Then
        MsgBox "Bitte zuerst die Arbeitsmappe anklicken, die gefaerbt werden soll " & _
               "(nicht " & ThisWorkbook.Name & ").", vbExclamation
        Exit Function
    End If
    Set ZielMappe = wb
End Function

Private Function BlattFormatieren(ByVal ws As Worksheet) As String
    Dim r As String, kopf As Long, ges As Long, letzte As Long
    Dim c As Long, ende As Long, txt As String, akz As Long, n As Long

    r = "[" & ws.Name & "]" & vbCrLf
    If ws.ProtectContents Then
        BlattFormatieren = r & "  UEBERSPRUNGEN: Blatt ist geschuetzt" & vbCrLf
        Exit Function
    End If

    kopf = FindeZeileA(ws, 4, MAX_ZEILE, "gesamtbudget*")
    If kopf = 0 Then
        BlattFormatieren = r & "  UEBERSPRUNGEN: keine Zeile mit 'Gesamtbudget' in Spalte A" & vbCrLf
        Exit Function
    End If
    ges = FindeZeileA(ws, kopf + 1, kopf + 20, "gesamt")
    If ges = 0 Then
        BlattFormatieren = r & "  UEBERSPRUNGEN: keine Zeile 'gesamt' in Spalte A unter Zeile " & kopf & vbCrLf
        Exit Function
    End If
    letzte = LetzteSpalte(ws, kopf)

    ' Kopfzeile
    With ws.Range(ws.Cells(kopf, 1), ws.Cells(kopf, letzte))
        Designfarbe .Cells, xlThemeColorAccent5
        .Font.Name = SCHRIFT_KOPF
        .Font.Bold = True
    End With
    r = r & "  Kopfzeile: Zeile " & kopf & " (A:" & SpName(letzte) & ")" & vbCrLf

    ' Bloecke unter den Ueberschriften
    c = 1
    Do While c <= letzte
        txt = TextVon(ws.Cells(kopf, c))
        If Len(txt) > 0 Then
            ende = c
            Do While ende + 1 <= letzte
                If Len(TextVon(ws.Cells(kopf, ende + 1))) > 0 Then Exit Do
                ende = ende + 1
            Loop
            akz = Klasse(txt)
            If akz <> 0 Then
                n = FaerbeGefuellte(ws.Range(ws.Cells(kopf + 1, c), ws.Cells(ges, ende)), akz, _
                                    IstBudgetTitel(LCase$(txt)))
                r = r & "  " & SpName(c) & ":" & SpName(ende) & "  '" & txt & "' -> " & _
                    KlassenName(akz) & " (" & n & " Zellen)" & vbCrLf
            Else
                r = r & "  " & SpName(c) & ":" & SpName(ende) & "  '" & txt & _
                    "' -> NICHT gefaerbt (unbekannt), bitte manuell" & vbCrLf
            End If
            c = ende + 1
        Else
            c = c + 1
        End If
    Loop

    ' Zeile "gesamt" fett
    ws.Range(ws.Cells(ges, 1), ws.Cells(ges, letzte)).Font.Bold = True
    r = r & "  Zeile 'gesamt': " & ges & " fett" & vbCrLf

    r = r & RestbudgetFormatieren(ws, ges)
    r = r & LegendeUmfaerben(ws)
    BlattFormatieren = r
End Function

Private Function RestbudgetFormatieren(ByVal ws As Worksheet, ByVal ges As Long) As String
    Dim z As Long, erste As Long, letzteZ As Long, t As String, s As Long, wertSp As Long
    Dim weitere As String

    For z = ges + 1 To ges + 15
        t = LCase$(TextVon(ws.Cells(z, 1)))
        If t Like "ist*" Or t Like "plan*" Or t Like "pln*" Or _
           t Like "restbudget*" Or t Like "projektmittel*" Then
            If erste = 0 Then erste = z
            letzteZ = z
        ElseIf Len(t) > 0 And erste > 0 Then
            Exit For
        End If
    Next z
    If erste = 0 Then
        RestbudgetFormatieren = "  Restbudget: NICHT gefunden, bitte manuell" & vbCrLf
        Exit Function
    End If

    For s = 2 To MAX_SPALTE
        If Application.WorksheetFunction.Count(ws.Range(ws.Cells(erste, s), ws.Cells(letzteZ, s))) > 0 Then
            wertSp = s
            Exit For
        End If
    Next s
    If wertSp = 0 Then
        RestbudgetFormatieren = "  Restbudget: Zeilen " & erste & "-" & letzteZ & _
                                ", aber keine Betraege gefunden - bitte manuell" & vbCrLf
        Exit Function
    End If

    With ws.Range(ws.Cells(erste, 1), ws.Cells(letzteZ, wertSp - 1)).Interior
        .Pattern = xlSolid
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = DUNKLER_15
    End With
    Designfarbe ws.Range(ws.Cells(erste, wertSp), ws.Cells(letzteZ, wertSp)), xlThemeColorAccent6
    ws.Range(ws.Cells(erste, 1), ws.Cells(letzteZ, wertSp)).Font.Bold = True

    For s = wertSp + 1 To MAX_SPALTE
        If Application.WorksheetFunction.Count(ws.Range(ws.Cells(erste, s), ws.Cells(letzteZ, s))) > 0 Then
            weitere = weitere & SpName(s) & " "
        End If
    Next s
    RestbudgetFormatieren = "  Restbudget: Zeilen " & erste & "-" & letzteZ & _
        ", grau A:" & SpName(wertSp - 1) & ", Betraege " & SpName(wertSp) & " gruen, fett" & vbCrLf
    If Len(weitere) > 0 Then
        RestbudgetFormatieren = RestbudgetFormatieren & _
            "    Hinweis: weitere Betraege in Spalte " & weitere & "- nicht gefaerbt, bitte pruefen" & vbCrLf
    End If
End Function

Private Function LegendeUmfaerben(ByVal ws As Worksheet) As String
    Dim f As Range, k As Long, t As String, n As Long, akz As Long
    Set f = ws.UsedRange.Find(What:="Legende", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If f Is Nothing Then
        LegendeUmfaerben = "  Legende: keine gefunden" & vbCrLf
        Exit Function
    End If
    For k = 1 To 6
        t = LCase$(TextVon(f.Offset(k, 0)))
        If Len(t) = 0 Then Exit For
        akz = 0
        If t Like "budget*" Or t Like "restbudget*" Then
            akz = xlThemeColorAccent6
        ElseIf t Like "obligo*" Then
            akz = xlThemeColorAccent1
        ElseIf t Like "ist*" Then
            akz = xlThemeColorAccent2
        ElseIf t Like "plan*" Then
            akz = xlThemeColorAccent4
        End If
        If akz <> 0 Then
            Designfarbe f.Offset(k, 0), akz
            n = n + 1
        End If
    Next k
    LegendeUmfaerben = "  Legende: " & f.Address(False, False) & ", " & n & " Zeilen umgefaerbt" & vbCrLf
End Function

Private Function Klasse(ByVal txt As String) As Long
    Dim t As String
    t = LCase$(txt)
    If IstBudgetTitel(t) Then
        Klasse = xlThemeColorAccent6
    ElseIf InStr(t, "obligo") > 0 Then
        Klasse = xlThemeColorAccent1
    ElseIf InStr(t, "plan") > 0 Or InStr(t, "pln") > 0 Then
        Klasse = xlThemeColorAccent4
    ElseIf t Like "gesamt*" Then
        Klasse = xlThemeColorAccent6
    ElseIf InStr(t, "ist") > 0 Or InStr(t, "ausg") > 0 Then
        Klasse = xlThemeColorAccent2
    Else
        Klasse = 0
    End If
End Function

Private Function KlassenName(ByVal akz As Long) As String
    Select Case akz
        Case xlThemeColorAccent6: KlassenName = "Gruen (Budget)"
        Case xlThemeColorAccent1: KlassenName = "Blau Akzent 1 (Obligo)"
        Case xlThemeColorAccent4: KlassenName = "Gold (PLAN)"
        Case xlThemeColorAccent2: KlassenName = "Orange (IST)"
        Case Else: KlassenName = "?"
    End Select
End Function

Private Function IstBudgetTitel(ByVal t As String) As Boolean
    ' Ueberschriften mit "Budget" oder "Mittel" (z. B. Gesamtbudget, Verfuegbares Budget 2026, Mittel 2026)
    IstBudgetTitel = (InStr(t, "budget") > 0 Or InStr(t, "mittel") > 0)
End Function

Private Function FaerbeGefuellte(ByVal bereich As Range, ByVal akz As Long, _
                                 Optional ByVal auchZahlen As Boolean = False) As Long
    ' Faerbt Zellen um, die bereits eine Fuellfarbe haben (= "Originalfarbblock").
    ' Bei Budget-Bloecken (auchZahlen = True) zusaetzlich alle Zellen mit einer Zahl.
    Dim z As Range, n As Long, mitZahl As Boolean
    For Each z In bereich.Cells
        mitZahl = False
        If auchZahlen Then
            If Not IsError(z.Value) Then
                If Not IsEmpty(z.Value) And IsNumeric(z.Value) And VarType(z.Value) <> vbString Then mitZahl = True
            End If
        End If
        If IstGefuellt(z) Or mitZahl Then
            Designfarbe z, akz
            n = n + 1
        End If
    Next z
    FaerbeGefuellte = n
End Function

Private Function IstGefuellt(ByVal z As Range) As Boolean
    If z.Interior.Pattern = xlNone Then Exit Function
    If z.Interior.Color = RGB(255, 255, 255) Then Exit Function
    IstGefuellt = True
End Function

Private Function FindeZeileA(ByVal ws As Worksheet, ByVal von As Long, ByVal bis As Long, _
                             ByVal muster As String) As Long
    Dim z As Long
    For z = von To bis
        If LCase$(TextVon(ws.Cells(z, 1))) Like muster Then
            FindeZeileA = z
            Exit Function
        End If
    Next z
End Function

Private Function LetzteSpalte(ByVal ws As Worksheet, ByVal zeile As Long) As Long
    Dim c As Long
    For c = MAX_SPALTE To 1 Step -1
        If Len(TextVon(ws.Cells(zeile, c))) > 0 Or IstGefuellt(ws.Cells(zeile, c)) Then
            LetzteSpalte = c
            Exit Function
        End If
    Next c
    LetzteSpalte = 1
End Function

Private Function TextVon(ByVal z As Range) As String
    If IsError(z.Value) Then Exit Function
    TextVon = Trim$(CStr(z.Value))
End Function

Private Function SpName(ByVal spalte As Long) As String
    SpName = Split(Cells(1, spalte).Address(True, False), "$")(0)
End Function

Private Sub BerichtAusgeben(ByVal mappe As String, ByVal bericht As String)
    ' Bericht in eine NEUE Arbeitsmappe schreiben (die bearbeitete Mappe bleibt unberuehrt)
    Dim wbB As Workbook, zeilen() As String, i As Long
    Set wbB = Workbooks.Add
    zeilen = Split(bericht, vbCrLf)
    wbB.Worksheets(1).Cells(1, 1).Value = "Bericht Farben - " & mappe & " - " & Format(Now, "dd.mm.yyyy hh:nn")
    wbB.Worksheets(1).Cells(1, 1).Font.Bold = True
    For i = 0 To UBound(zeilen)
        wbB.Worksheets(1).Cells(i + 3, 1).Value = "'" & zeilen(i)
    Next i
    wbB.Worksheets(1).Columns(1).ColumnWidth = 110
End Sub
