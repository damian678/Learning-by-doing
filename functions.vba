Option Explicit

Function SumaKwadratow(A As Double, B As Double) As Double
    SumaKwadratow = (A) ^ 2 + (B) ^ 2
End Function

Function Wyroznik(A As Double, B As Double, C As Double) As Double
    Wyroznik = (B) ^ 2 - 4 * A * C
End Function

Function Poleprostokata(A As Double, B As Double) As Double
    Poleprostokata = A * B
End Function

Function Polekwadratu(A As Double) As Double
    Polekwadratu = Poleprostokata(A, A)
End Function

Function Poletrapezu(pod1 As Double, pod2 As Double, h As Double) As Double
    Poletrapezu = ((pod1 + pod2) * h) / 2
End Function

Function Poletrojkata(pod As Double, h As Double) As Double
    Poletrojkata = pod * h / 2
End Function

Function PolePC(A As Double, B As Double, C As Double) As Double
    PolePC = 2 * Poleprostokata(A, B) + 2 * Poleprostokata(A, C) + 2 * Poleprostokata(B, C)
End Function

Function PoleKola(r As Double) As Double
    PoleKola = WorksheetFunction.Pi * (r) ^ 2
End Function

Function Objetoscwalca(r As Double, h As Double) As Double
    Objetoscwalca = h * PoleKola(r)
End Function

Function ObjetoscKuli(r As Double)
    ObjetoscKuli = 4 / 3 * WorksheetFunction.Pi * (r) ^ 3
End Function

Function Obwodprostokata(A As Double, B As Double) As Double
    Obwodprostokata = 2 * A + 2 * B
End Function

Function Obwodkola(r As Double) As Double
    Obwodkola = 2 * WorksheetFunction.Pi * r
End Function

Function PolePK(r As Double) As Double
    PolePK = 4 * WorksheetFunction.Pi * (r) ^ 2
End Function

Function StopnieF(C As Double) As Double
    StopnieF = 32 + 9 / 5 * C
End Function

Function StopnieC(F As Double) As Double
    StopnieC = 5 / 9 * (F - 32)
End Function

Function Minimum(A As Double, B As Double) As Double
    If A < B Then
        Minimum = A
    Else
        Minimum = B
    End If
End Function

Function Minimum2(A As Double, B As Double, C As Double) As Double
    If Minimum(A, B) = A And Minimum(A, C) = A Then
        Minimum2 = A
    ElseIf Minimum(B, C) = B Then
        Minimum2 = B
    Else: Minimum2 = C
    End If
End Function

Function przestepny(A As Double) As Boolean
    If A Mod 4 = 0 And A Mod 100 <> 0 Then
        przestepny = True
    ElseIf A Mod 400 = 0 Then
        przestepny = True
    Else
        przestepny = False
    End If
End Function

Function wieksza0(A As Double) As Boolean
    If A > 0 Then
        wieksza0 = True
    Else
        wieksza0 = False
    End If
End Function

Function jedno(A As Double) As Boolean
    If A < 10 And A > -10 Then
        jedno = True
    Else
        jedno = False
    End If
End Function

Function calkowita(A As Double) As Boolean
    If Int(A) = A Then
        calkowita = True
    Else
        calkowita = False
    End If
End Function

Function parzysta(A As Double) As Boolean
    If Int(A) = A And A Mod 2 = 0 Then
        parzysta = True
    Else
        parzysta = False
    End If
End Function

Function podzielna(A As Double) As Boolean
    If A Mod 17 = 0 Then
        podzielna = True
    Else
        podzielna = False
    End If
End Function

Function termin(dzien As Integer, miesiac As Integer, rok As Integer) As String
    Select Case miesiac
        Case 1
        termin = dzien & " Styczeń " & rok
        Case 2
        termin = dzien & " Luty " & rok
        Case 3
        termin = dzien & " Marzec " & rok
        Case 4
        termin = dzien & " Kwiecień " & rok
        Case 5
        termin = dzien & " Maj " & rok
        Case 6
        termin = dzien & " Czerwiec " & rok
        Case 7
        termin = dzien & " Lipiec " & rok
        Case 8
        termin = dzien & " Sierpień " & rok
        Case 9
        termin = dzien & " Wrzesień " & rok
        Case 10
        termin = dzien & " Październik " & rok
        Case 11
        termin = dzien & " Listopad " & rok
        Case 12
        termin = dzien & " Grudzień " & rok
    End Select
End Function

Function Cyfry(A As String) As Integer
    Select Case A
        Case "I"
        Cyfry = 1
        Case "V"
        Cyfry = "5"
        Case "X"
        Cyfry = 10
        Case "L"
        Cyfry = 50
        Case "C"
        Cyfry = 100
        Case "D"
        Cyfry = 500
        Case "M"
        Cyfry = 1000
    End Select
End Function
Function przestepny(A As Double) As Boolean
    If A Mod 4 = 0 And A Mod 100 <> 0 Then
        przestepny = True
    ElseIf A Mod 400 = 0 Then
        przestepny = True
    Else
        przestepny = False
    End If
End Function

Function LiczbaDni(rok As Integer, miesiac As Integer) As Integer
    Select Case miesiac
        Case 1, 3, 5, 7, 8, 10, 12:
            LiczbaDni = 31
        Case 4, 6, 9, 11:
            LiczbaDni = 30
        Case 2:
            If rok Mod 4 = 0 And rok Mod 100 <> 0 Then
                LiczbaDni = 29
            ElseIf rok Mod 400 = 0 Then
                LiczbaDni = 29
            Else
                LiczbaDni = 28
            End If
        End Select
End Function

Function Waluta(Liczba As Integer) As String
    Select Case Liczba
    Case 1
    Waluta = Liczba & " złoty"
    Case Else
        Select Case Liczba Mod 100
        Case 2 To 4, 22 To 24, 32 To 34, 42 To 44, 62 To 64, 72 To 74, 82 To 84, 92 To 94
        Waluta = Liczba & " zlote"
        Case Else
        Waluta = Liczba & " złotych"
        End Select
    End Select
End Function

Function NWD(A As Integer, b As Integer) As Integer
    Dim Tmp As Integer
    A = Abs(A)
    b = Abs(b)
    Do While b > 0
        Tmp = A
        A = b
        b = Tmp Mod b
    Loop
    NWD = A
End Function

Function LiczbaCyfr(A As Long) As Long
    Dim Ldzielen As Integer
    Ldzielen = 0
    A = Abs(A)
    
    Do While A >= 10
        A = A / 10
        Ldzielen = Ldzielen + 1
    Loop
    
    LiczbaCyfr = Ldzielen + 1
End Function

Function SumaCyfr(A As Long) As Long
    Dim S As Integer
    S = 0
    A = Abs(A)
    
    Do While A > 0
        S = S + A Mod 10
        A = A / 10
     Loop
    SumaCyfr = S
End Function

Function Funkcjax(x As Double) As Double
 Funkcjax = IIf(x >= 0, Abs(x) + 1, Abs(x) - 1)
End Function

Function Kwadraty(A As Double, B As Double) As Double
    Kwadraty = IIf(A >= B, (A - B) + 1, (B - A) + 1)
End Function

Function Funkcja(x As Double) As Double
    Funkcja = IIf(x = 0, 777, Tan(x * Abs(x)) + Sin(Abs(x)) + Cos((x) ^ (2) - 1))
End Function

Function LiczbyRzymskie(A As String) As Integer
    LiczbyRzymskie = Switch(A = "I", 1, A = "V", 5, A = "X", 10, A = "L", 50, A = "C", 100, A = "D", 500, A = "M", 1000)
End Function

Function Reszta(A As Integer) As String
    Reszta = IIf(A Mod 5 = 0, "zero", Choose(A Mod 5, "jeden", "dwa", "trzy", "cztery"))
End Function

Function Logarytm(A As Double, P As Double) As Variant
    If A < 0 Or P < 0 Or P = 1 Then
        Logarytm = CVErr(4 + vbObjectError)
    Else
        Logarytm = Log(A) / Log(P)
    End If
End Function

Function Cot(A As Double) As String
    Cot = IIf(A <> WorksheetFunction.Pi, 1 / Tan(A), "Nie można policzyć")
End Function

Function Sinushipe(A As Double) As String
    Sinushipe = IIf(A <> 0, (Exp(A) - Exp(-A)) / 2, "Nie można policzyć")
End Function


Function Losuj(A As Double) As Double
    Losuj = Int(Rnd(A) * 1998 + 2)
End Function

Function Logarytm(Podstawa As Double, ByVal X As Double) As Variant
    If Podstawa <= 0 Or Podstawa = 1 Or X <= 0 Then
        Logarytm = CVErr(xlErrValue)
    Else
        Logarytm = Log(X) / Log(Podstawa)
    End If
End Function

Function Palindrom(Napis As String) As Boolean
    Palindrom = Napis = StrReverse(Napis)
End Function

Function Skrot(Napis As String) As String
   Dim Male As String
   Dim Odpowiednie As String
   Dim I As Integer
    Male = StrConv(Napis, vbLowerCase)
    Odpowiednie = StrConv(Napis, vbProperCase)
    For I = 1 To Len(Napis)
        Znak = Mid(Odpowiednie, I, 1)
        If Mid(Male, I, 1) <> Znak Then
            Skrot = Skrot & Znak
        End If
    Next I
End Function
