Option Explicit

Sub Informacje()
    MsgBox "Informacja 1 - Wyst�pi� b��d", vbAbortRetryIgnore, "Komunikat"
    MsgBox "Informacja 2 - Kolejny b��d", vbApplicationModal, "B��d"
    MsgBox "Informacja 3 - B��d", vbOKOnly, "Okienko"
    MsgBox "Informacja 4 - Koniec b��du", vbCritical, "Okno b��du"
End Sub

Sub Procedura()
    Dim A As Double
    Dim P As Double
    
    A = InputBox("Podaj d�ugo�� boku kwadratu")
    P = Geometria.Polekwadratu(A)
    MsgBox "Pole kwadratu wynosi " & P & "."
End Sub

Sub Sumowanie()
    Dim IleLiczb As Integer
    Dim I As Integer
    Dim Suma, Liczba As Double
    IleLiczb = InputBox("Ile liczb chcesz zsumowa�?")
    Suma = 0
    For I = 1 To IleLiczb
        Liczba = InputBox("Podaj kolejn� liczb�")
        Suma = Suma + Liczba
    Next I
    MsgBox "Wynik: " & Suma
End Sub

Sub TestBledu()
    Dim A As Double, B As Double
    Dim L As Variant
    A = InputBox("A= ")
    B = InputBox("B= ")
    L = Logarytm(A, B)
    If IsError(L) Then
        MsgBox "B��d"
    Else
        MsgBox L
    End If
End Sub

Function ZnakNaPozycji(Napis As String, Pozycja As Integer) As String
    If Pozycja < 1 Or Pozycja > Len(Napis) Then
        Err.Raise vbObjectError + 1, , "Z�a pozycja"
    Else
        ZnakNaPozycji = Mid(Napis, Pozycja, 1)
    End If
End Function

Sub TestBledu2()
    Dim Napis As String
    Dim Poz As Integer
    On Error GoTo ObslugaBledow
    Napis = InputBox("Napis= ")
    Poz = InputBox("Pozycja=")
    MsgBox ZnakNaPozycji(Napis, Poz)
    Exit Sub
ObslugaBledow:
 MsgBox "B��d"
End Sub

Function Skrot(Nazwa As String) As String
    Dim Wyrazy As Variant
    Dim I As Integer
    Wyrazy = Split(Nazwa, " ", 5)
    For I = LBound(Wyrazy) To UBound(Wyrazy)
        Skrot = Skrot & UCase(Left(Wyrazy(I), 1))
    Next I
End Function

Sub TestujSkrot()
    Dim Nazwa As String
    Nazwa = InputBox("Podaj nazw�: ")
    MsgBox Skrot(Nazwa)
End Sub



Sub Data()
    MsgBox Date
End Sub

Sub Czas()
    MsgBox Time
End Sub

Sub DataiCzas()
    MsgBox Now
End Sub

'Dodaj 3 dni do dowolnej daty

Sub dodawanie()
    Dim Data As Variant
    Datadowolna = DateSerial(1994, 7, 31)
    Data = DateAdd("d", 3, Datadowolna)
    MsgBox Data
End Sub

'Dodaj 3 godziny do dowolnego czasu

Sub plusczas()
    Dim Czas As Variant
    DowolnyCzas = TimeSerial(15, 12, 23)
    Czas = DateAdd("h", 3, DowolnyCzas)
    MsgBox Czas
End Sub

' Napisz proceduj� licz�c� r�nic� dni pomi�dzy dwoma datami
Sub Roznica()
    Dim Roznica As Variant
    Data1 = DateSerial(1994, 7, 31)
    Data2 = DateSerial(1999, 1, 26)
    Roznica = DateDiff("d", Data1, Data2)
    MsgBox Roznica
End Sub

'Napisz procedur� m�wi�c� o tym, kt�ry dzie� tygodnia ma data przedstawiona przez u�ytkownika
Sub dzientygodnia()
    Dim R, M, D As Integer
    Dim Data As Variant
    Dim dzientygodnia As Variant
    R = InputBox("Podaj rok")
    M = InputBox("Podaj miesi�c")
    D = InputBox("Podaj dzie�")
    Data = DateSerial(R, M, D)
    dzientygodnia = Weekday(Data, vbMonday)
    MsgBox dzientygodnia
End Sub

Sub test() 'zadanie 68
    LMagazynow = InputBox("Podaj liczb� magazyn�w")
    LTowarow = InputBox("Podaj liczb� towar�w")
    Dim I As Integer
    Dim j As Integer
    Cells(1, 2) = "Suma"
    For I = 3 To LMagazynow + 2
       For j = 2 To LTowarow + 1
    Cells(1, I) = "Magazyn " & (I - 2)
    Cells(j, 1) = "Towar " & (j - 1)
    Cells(I + 1, 2).FormulaLocal _
        = "=suma(C" & (I + 1) & ":" & Chr(Asc("B") + LiczbaMagazynow) & (I + 1) & ")"
        Next j
    Next I
End Sub

Sub test2()
    Dim LiczbaMagazynow As Integer
    Dim LiczbaTowarow As Integer
    Dim I As Integer
    LiczbaMagazynow = InputBox("Podaj liczb� magazyn�w")
    LiczbaTowarow = InputBox("Podaj liczb� towar�w")
    Range("B1") = "Suma"
    For I = 1 To LiczbaMagazynow
        Cells(1, I + 2) = "Magazyn " & I
    Next I
    For I = 1 To LiczbaTowarow
        Cells(I + 1, 1) = "Towar " & I
        Cells(I + 1, 2).FormulaLocal _
        = "=suma(C" & (I + 1) & ":" & Chr(Asc("B") + LiczbaMagazynow) & (I + 1) & ")"
    Next I
End Sub\