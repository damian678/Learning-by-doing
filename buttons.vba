Option Explicit 

Private Sub CommandButton1_Click()
    Dim A As Double
    Dim P As Double
    
    A = InputBox("Podaj warto�� temperatury w stopniach Fahrenheita")
    P = Geometria.StopnieF(A)
    MsgBox "W przeliczeniu na stopnie Celcjusza wynosi: " & P & "."
End Sub

Private Sub CommandButton2_Click()
    Dim A As String
    Dim B As String
    
    A = InputBox("Podaj swoje imi�")
    B = InputBox("Podaj swoje nazwisko")
    MsgBox "Nazywasz si� " & A & " " & B & "."
End Sub

Private Sub CommandButton1_Click() 
Dim A As Double
Dim P As Double

A = InputBox("Podaj promie� ko�a")
P = Geometria.PoleKola(A)
If A > 0 Then
    MsgBox "Pole ko�a wynosi " & P & "."
Else
    MsgBox "Nieprawid�owy promie� ko�a"
End If
End Sub

Private Sub CommandButton2_Click()
Dim A As Double
Dim B As Double
Dim x As Double

A = InputBox("Podaj warto�� parametru a")
B = InputBox("Podaj warto�� parametru b")
x = -B / A
MsgBox "Rozwi�zaniem r�wnania jest x=" & x & "."
End Sub

Private Sub CommandButton1_Click() 
Dim R, M, D As Integer
    R = InputBox("Podaj rok")
    M = InputBox("Podaj miesi�c")
    D = InputBox("Podaj dzie�")
    MsgBox WeekdayName(Weekday(DateSerial(R, M, D), vbMonday))
End Sub


Private Sub CommandButton1_Click()  
    MsgBox Application.Version
End Sub

Private Sub CommandButton2_Click()
     ActiveWorkbook.Save
End Sub

Private Sub CommandButton3_Click()
ActiveWorkbook.SaveAs "C:\Users\komputer\Desktop\Programowanie w j�zyku Visual Basic\zadanie2.xlsm"
End Sub

Private Sub CommandButton4_Click()
    ActiveWorkbook.Close
End Sub

Private Sub CommandButton5_Click()
    Dim A As Integer
    A = InputBox("Do kt�rego arkusza mam przej��?")
    ThisWorkbook.Worksheets(A).Activate
End Sub

Private Sub CommandButton6_Click()
    Dim A As Integer
    A = InputBox("Ile arkuszy mam doda�?")
    ThisWorkbook.Worksheets.Add After:=Worksheets(Worksheets.Count), Count:=A
End Sub


Private Sub CommandButton7_Click()
    Dim A As Integer
    A = InputBox("Kt�ry arkusz mam usun��?")
    ThisWorkbook.Worksheets(A).Delete
End Sub

Private Sub CommandButton8_Click()
    ThisWorkbook.Worksheets(1).Name = Cells(1, 1)
    ThisWorkbook.Worksheets(2).Name = Cells(2, 1)
    ThisWorkbook.Worksheets(3).Name = Cells(3, 1)
End Sub

Private Sub CommandButton9_Click()
    ThisWorkbook.Worksheets(2).Name = Cells(5, 3)
    If Arkusz2.Visible = xlSheetHidden Then
    Arkusz2.Visible = xlSheetVisible
    Else: Arkusz2.Visible = xlSheetHidden
    End If
End Sub

Private Sub CommandButton1_Click() 'zadanie 66
    Cells(7, 4).Formula = Cells(7, 4) * 2
End Sub

Private Sub CommandButton2_Click()
    Dim A As String
    A = InputBox("Podaj nazw� arkusza")
    
    For I = 1 To Worksheets.Count
        If Worksheets(I).Name = A Then
        Istnieje = True
        End If
    Next I
    
        If Istnieje = True Then
        B = InputBox("Podaj now� nazw�")
        Worksheets(A).Name = B
        Else
        ThisWorkbook.Worksheets.Add
        ThisWorkbook.Worksheets(1).Name = A
        End If
End Sub

Private Sub CommandButton3_Click()
    MsgBox "Lewy margines = " & PageSetup.LeftMargin
    MsgBox "Prawy margines = " & PageSetup.RightMargin
    MsgBox "G�rny margines = " & PageSetup.TopMargin
    MsgBox "Dolny margines = " & PageSetup.BottomMargin
End Sub

Private Sub CommandButton4_Click()
    ThisWorkbook.Worksheets(1).Name = Cells(1, 1)
    Worksheets(1).PageSetup.LeftMargin = Application.CentimetersToPoints(Cells(2, 1))
    Worksheets(1).PageSetup.RightMargin = Application.CentimetersToPoints(Cells(3, 1))
    Worksheets(1).PageSetup.TopMargin = Application.CentimetersToPoints(Cells(4, 1))
    Worksheets(1).PageSetup.BottomMargin = Application.CentimetersToPoints(Cells(5, 1))
End Sub