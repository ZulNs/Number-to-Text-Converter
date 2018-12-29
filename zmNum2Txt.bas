Attribute VB_Name = "zmNum2Txt"
Option Explicit

Const cStrZero  As String = "0"
Const cDescZero As String = "Nol"
Const cStrNine  As String = "9"
Const cStrExp   As String = "E"
Const cStrErr   As String = "Err"

Function CNT(Angka As String, Optional Opsi As String = Empty) As String
    If Not IsEnaFunc(Angka) Then
        Beep
        CNT = Angka
        Exit Function
    End If
    CNT = CNum(Angka)
    If Mid(StrReverse(Opsi), 3, 1) = "1" Then
        CNT = Angka
        Exit Function
    End If
    If CNT = cStrErr Then
        If Angka = Empty Then CNT = Empty Else CNT = "# Angka masukan salah..."
        Exit Function
    End If
    If Mid(StrReverse(Opsi), 2, 1) = "1" Then CNT = CNT & " Rupiah"
    Select Case Right(Opsi, 1)
        Case "1"
            CNT = LCase(CNT)
            CNT = UCase(Left(CNT, 1)) & Mid(CNT, 2)
        Case "2"
            CNT = LCase(CNT)
        Case "3"
            CNT = UCase(CNT)
    End Select
End Function

Private Property Get CNum(Num As String) As String
    Dim N As String, Exp As String, ExpPos As Long
    Num = Trim(Num)
    CNum = cStrErr
    If Num = Empty Then Exit Property
    N = Num
    ExpPos = InStr(UCase(Num), cStrExp)
    If ExpPos Then
        Exp = Mid(Num, ExpPos + 1)
        If Not ChkNum(Exp) Then Exit Property
        N = Left(N, ExpPos - 1)
    End If
    If Not ChkNum(N) Then Exit Property
    Num = N
    If N = cStrZero Then
        CNum = cDescZero
    Else
        CNum = CFullNum(N)
        If Exp <> Empty And Exp <> cStrZero Then
            Num = N & cStrExp & Exp
            CNum = CNum & " Kali Sepuluh"
            If Exp <> "1" Then CNum = CNum & " Pangkat " & CFullNum(Exp)
        End If
    End If
End Property

Private Property Get ChkNum(Num As String) As Boolean
    Dim SgnN As String, IntN As String, DecN As String, DPPos As Long, I As Long, _
        NumFl As Boolean
    Num = Trim(Num)
    If Num = Empty Then Exit Property
    For I = 1 To Len(Num)
        Select Case Mid(Num, I, 1)
            Case Application.International(xlDecimalSeparator), cStrZero To cStrNine
                Exit For
        End Select
    Next
    IntN = Mid(Num, I)
    If IntN = Empty Then Exit Property
    SgnN = Left(Num, I - 1)
    SgnN = Replace(Replace(Replace(SgnN, " ", Empty), "+", Empty), "--", Empty)
    If SgnN <> Empty And SgnN <> "-" Then Exit Property
    DPPos = InStr(IntN, Application.International(xlDecimalSeparator))
    If DPPos Then
        DecN = Mid(IntN, DPPos + 1)
        IntN = Left(IntN, DPPos - 1)
        If IntN = Empty And DecN = Empty Then Exit Property
    End If
    For I = 1 To Len(IntN)
        Select Case Mid(IntN, I, 1)
            Case Application.International(xlThousandsSeparator), cStrZero To cStrNine
            Case Else
                Exit Property
        End Select
    Next
    For I = 1 To Len(DecN)
        Select Case Mid(DecN, I, 1)
            Case cStrZero To cStrNine
                NumFl = True
            Case Application.International(xlThousandsSeparator)
                If Not NumFl Then Exit Property
            Case Else
                Exit Property
        End Select
    Next
    IntN = Replace(IntN, Application.International(xlThousandsSeparator), Empty)
    DecN = Replace(DecN, Application.International(xlThousandsSeparator), Empty)
    If Len(SgnN) And IntN & DecN = Empty Then Exit Property
    Do While Left(IntN, 1) = cStrZero
        IntN = Mid(IntN, 2)
    Loop
    Do While Right(DecN, 1) = cStrZero
        DecN = Left(DecN, Len(DecN) - 1)
    Loop
    If IntN = Empty Then Num = cStrZero Else Num = IntN
    If Len(DecN) Then Num = Num & Application.International(xlDecimalSeparator) & DecN
    If Len(SgnN) And Num <> cStrZero Then Num = SgnN & Num
    ChkNum = True
End Property

Private Property Get CFullNum(ByVal Num As String) As String
    Dim SgnNum As Boolean, DecSptPos As Long, DecNum As String, I As Integer
    If Left(Num, 1) = "-" Then
        SgnNum = True
        Num = Mid(Num, 2)
    End If
    DecSptPos = InStr(Num, Application.International(xlDecimalSeparator))
    If DecSptPos Then
        CFullNum = Mid(Num, DecSptPos + 1)
        DecNum = CIntNum(CFullNum)
        If DecNum <> Empty Then
            For I = 1 To Len(CFullNum)
                If Mid(CFullNum, I, 1) = cStrZero Then _
                   DecNum = cDescZero & " " & DecNum Else Exit For
            Next
        End If
        Num = Left(Num, DecSptPos - 1)
    End If
    CFullNum = CIntNum(Num)
    If CFullNum = Empty And DecNum <> Empty Then CFullNum = cDescZero
    If DecNum <> Empty Then CFullNum = CFullNum & " Koma " & DecNum
    If SgnNum And CFullNum <> Empty Then CFullNum = "Minus " & CFullNum
End Property

Private Property Get CIntNum(Num As String) As String
    Dim LN As Long, P As Long
    LN = Len(Num)
    If LN <= 18 Then
        CIntNum = CBaseIntNum(Num)
        Exit Property
    End If
    P = LN Mod 15
    CIntNum = CBaseIntNum(Left(Num, P))
    Do While P < LN
        CIntNum = MergeTxt(AddUnit(CIntNum, "Bilyun"), CBaseIntNum(Mid(Num, P + 1, 15)))
        P = P + 15
    Loop
End Property

Private Property Get CBaseIntNum(ByVal Num As String) As String
    Dim NU(0 To 5) As String, LN As Integer, I As Integer
    NU(1) = "Ribu"
    NU(2) = "Juta"
    NU(3) = "Milyar"
    NU(4) = "Trilyun"
    NU(5) = "Bilyun"
    Num = FixNum(Right(Num, 18))
    LN = Len(Num)
    For I = 1 To LN Step 3
        CBaseIntNum = MergeTxt(CBaseIntNum, _
                               AddUnit(CBaseNum(Mid(Num, I, 3)), NU((LN - I - 2) / 3)))
    Next
End Property

Private Property Get CBaseNum(ByVal Num As String) As String
    Dim TN(0 To 9) As String, N2 As Integer, N1 As Integer
    TN(1) = "Satu"
    TN(2) = "Dua"
    TN(3) = "Tiga"
    TN(4) = "Empat"
    TN(5) = "Lima"
    TN(6) = "Enam"
    TN(7) = "Tujuh"
    TN(8) = "Delapan"
    TN(9) = "Sembilan"
    Num = FixNum(Right(Num, 3))
    N2 = Val(Mid(Num, 2, 1))
    N1 = Val(Mid(Num, 3, 1))
    CBaseNum = AddUnit(TN(Val(Left(Num, 1))), "Ratus")
    If N2 = 1 And N1 > 0 Then
        CBaseNum = MergeTxt(CBaseNum, AddUnit(TN(N1), "Belas"))
        Exit Property
    End If
    CBaseNum = MergeTxt(MergeTxt(CBaseNum, AddUnit(TN(N2), "Puluh")), TN(N1))
End Property

Private Property Get FixNum(Num As String) As String
    Dim LNM3 As Integer
    LNM3 = Len(Num) Mod 3
    If LNM3 Then FixNum = Space(3 - LNM3) & Num Else FixNum = Num
End Property

Private Property Get AddUnit(Num As String, Unit As String) As String
    If Num <> Empty Then
        If Num = "Satu" And ( _
           Unit = "Belas" Or Unit = "Puluh" Or Unit = "Ratus" Or Unit = "Ribu") Then
            AddUnit = "Se" & LCase(Unit)
        Else
            AddUnit = MergeTxt(Num, Unit)
        End If
    End If
End Property

Private Property Get MergeTxt(Text1 As String, Text2 As String) As String
    MergeTxt = Text1 & Space(Sgn(Len(Text1)) * Sgn(Len(Text2))) & Text2
End Property
