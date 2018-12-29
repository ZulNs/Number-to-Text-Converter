Attribute VB_Name = "zmMain"
Option Explicit

Const cPass    As Double = 7895123.19701105
Const cLmtExp  As Date = 20
Const csFst    As Integer = 1
Const csTrl    As Integer = 2
Const csExp    As Integer = 3
Const csReg    As Integer = 4
Const csErr    As Integer = 5
Const csDmg    As Integer = 6
Const csCpy    As Integer = 7
Const cvBase   As Integer = 50
Const cvAuthor As Integer = cvBase + 1
Const cvCallNm As Integer = cvAuthor + 1
Const cvCallNo As Integer = cvCallNm + 1
Const cvTitle  As Integer = cvCallNo + 1
Const cvCmts   As Integer = cvTitle + 1
Const cvCmt2   As Integer = cvCmts + 1
Const cvCmt3   As Integer = cvCmt2 + 1
Const cvFileNm As Integer = cvCmt3 + 1
Const cvUserNm As Integer = cvFileNm + 1
Const cvRegNm  As Integer = cvUserNm + 1
Const cvRegCd  As Integer = cvRegNm + 1
Const cvFstTm  As Integer = cvRegCd + 1
Const cvInitTm As Integer = cvFstTm + 1

Private Status   As Integer
Private Remain   As Date
Private FrmTitle As String

Private Property Let Vars(Ptr As Integer, Txt As String)
    ThisWorkbook.Sheets(1).Cells(Ptr, 1) = Encrypt(Txt)
End Property

Private Property Get Vars(Ptr As Integer) As String
    Vars = Uncrypt(ThisWorkbook.Sheets(1).Cells(Ptr, 1))
End Property

Private Property Get Encrypt(Txt As String) As String
    Dim I As Long, A As Integer, C1 As Integer, C2 As Integer
    For I = 1 To Len(Txt)
        A = Asc(Mid(Txt, I, 1)) Xor 181
        If A < 64 Then
            C1 = A + 64
            C2 = C1
        Else
            C1 = A
            If A < 240 Then C2 = A + 16 Else C2 = A - 176
        End If
        Encrypt = Encrypt & Chr(C1) & Chr(C2)
    Next
End Property

Private Property Get Uncrypt(Txt As String) As String
    Dim I As Long, A As Integer
    For I = 1 To Len(Txt) Step 2
        A = Asc(Mid(Txt, I, 1))
        If A = Asc(Mid(Txt, I + 1, 1)) Then A = A - 64
        Uncrypt = Uncrypt & Chr(A Xor 181)
    Next
End Property

Private Property Get KeyCode(ByVal Txt As String) As String
    If Txt = Empty Then Exit Property
    Dim A As Integer, I As Integer, J As Integer, L As Integer, X As Integer, T As String
    Txt = UCase(Txt)
    For I = 1 To Len(Txt)
        X = X Xor Asc(Mid(Txt, I, 1))
    Next
    A = 10 - Len(Txt) Mod 10
    If A < 10 Then Txt = Txt & Space(A)
    L = Len(Txt) / 10
    For I = 1 To 10
        A = Asc(Mid(Txt, I, 1))
        For J = 2 To L
            A = A Xor Asc(Mid(Txt, J * 10 - 10 + I, 1))
        Next
        A = (A Xor I ^ 2 Xor X) Mod 36
        If A < 10 Then A = A + 48 Else A = A + 55
        KeyCode = KeyCode & Chr(A)
    Next
End Property

Private Property Get AddInCmts() As String
    AddInCmts = Vars(cvCmts) & vbCr & _
                Vars(cvCmt2) & vbCr & _
                "[Dibuat oleh : " & Vars(cvAuthor) & "]" & vbCr
    If (Status = csReg Or Status = csCpy) And Vars(cvRegNm) <> Vars(cvAuthor) Then
        AddInCmts = AddInCmts & "[Dilisensikan untuk : " & Vars(cvRegNm) & "]"
    Else
        AddInCmts = AddInCmts & "[" & Vars(cvCmt3) & "]"
    End If
End Property

Private Property Get ChkPass(Pass As Double) As Boolean
    If Pass = cPass Then ChkPass = True
End Property

Property Get IsEnaFunc(Result As String) As Boolean
    Select Case LCase(Replace(Replace(Trim(Result), ".", Empty), " ", Empty))
        Case LCase(Replace(Replace(Vars(cvAuthor), ".", Empty), " ", Empty)), _
             LCase(Vars(cvCallNm))
            Result = " # He is my creator, and my best regards for him..."
            Exit Property
    End Select
    Select Case Status
        Case csReg
            IsEnaFunc = True
        Case csFst, csTrl
            If Now - CDate(Vars(cvInitTm)) < Remain Then
                IsEnaFunc = True
            Else
                Status = csExp
                Result = "# Add-In kadaluarsa..."
            End If
        Case csExp
            Result = "# Add-In kadaluarsa..."
        Case csErr
            Result = "# Sistem penanggalan kacau..."
        Case csDmg
            Result = "# Add-In rusak..."
        Case csCpy
            Result = "# Hanya untuk : " & Vars(cvRegNm) & "..."
    End Select
End Property

Private Property Get GetStatus() As Integer
    If Vars(cvRegNm) <> Empty And Vars(cvRegCd) <> Empty Then
        If KeyCode(Vars(cvRegNm)) = Vars(cvRegCd) Then
            If Application.UserName = Vars(cvUserNm) Then GetStatus = csReg _
                                                     Else GetStatus = csCpy
            Exit Property
        End If
    End If
    If Vars(cvRegNm) <> Empty Or Vars(cvRegCd) <> Empty Then
        GetStatus = csDmg
        Exit Property
    End If
    Dim DFirst As String, DInit As String
    DFirst = Vars(cvFstTm)
    DInit = Vars(cvInitTm)
    If DFirst = Empty And DInit = Empty Then
        GetStatus = csFst
        Remain = cLmtExp
    Else
        If Not IsDate(DFirst) Or Not IsDate(DInit) Then
            GetStatus = csDmg
        Else
            If CDate(DInit) < CDate(DFirst) Then
                GetStatus = csDmg
            Else
                If Now < CDate(DInit) Then
                    GetStatus = csErr
                Else
                    Remain = cLmtExp + CDate(DFirst) - Now
                    If Remain <= 0 Then
                        GetStatus = csExp
                    Else
                        GetStatus = csTrl
                    End If
                End If
            End If
        End If
    End If
End Property

Private Sub RestoreOrgProperty(Pass As Double)
    If Not ChkPass(Pass) Then Exit Sub
    Application.EnableCancelKey = xlDisabled
    Application.DisplayStatusBar = False
    With ThisWorkbook
        If .Author <> Vars(cvAuthor) Then .Author = Vars(cvAuthor)
        If .Title <> Vars(cvTitle) Then .Title = Vars(cvTitle)
        If .Comments <> AddInCmts Then .Comments = AddInCmts
        If .IsAddin = False Then .IsAddin = True
    End With
    SaveMe Pass
    Application.DisplayStatusBar = True
End Sub

Private Sub SaveMe(Pass As Double)
    If Not ChkPass(Pass) Then Exit Sub
    Application.EnableEvents = False
    Dim S1 As String, S2 As String
    With ThisWorkbook
        If LCase(.Name) <> LCase(Vars(cvFileNm)) Then
            S1 = .Path & Application.PathSeparator & Vars(cvFileNm)
            S2 = .FullName
            .SaveAs S1
            SetAttr S2, vbArchive
            Kill S2
        Else
            If Not .Saved Then
                S1 = .FullName
                S2 = Chr(46) & Chr(122)
                If (GetAttr(S1) And vbReadOnly) Or .ReadOnly Or .ReadOnlyRecommended Then
                    .SaveAs S1 & S2
                    SetAttr S1, vbArchive
                    Kill S1
                    .SaveAs S1
                    Kill S1 & S2
                Else
                    .Save
                End If
            End If
        End If
    End With
    Application.EnableEvents = True
End Sub

Private Property Get UserInp(Arg As String, ByVal Def As String) As String
EnterUserInfo:
    UserInp = Trim(InputBox("Masukkan " & Arg & " anda :", FrmTitle, Def))
    If UserInp = Empty Then If MsgBox("Anda yakin untuk membatalkan registrasi?", vbYesNo _
                               + vbQuestion, FrmTitle) = vbYes Then Exit Property _
                               Else GoTo EnterUserInfo
    If MsgBox("Apakah " & Arg & " yang anda masukkan sudah benar?", vbYesNo + vbQuestion, _
              FrmTitle) = vbNo Then
        Def = UserInp
        GoTo EnterUserInfo
    End If
End Property

Private Sub RegNow()
    Dim Nm As String, Cd As String, Msg As String
    Application.EnableCancelKey = xlDisabled
    FrmTitle = "Registrasi '" & Vars(cvTitle) & "'"
    Select Case Status
        Case Empty, csErr, csDmg
            MsgBox "Maaf!!!" & vbCr & _
                   "Untuk saat ini permintaan registrasi belum dapat dilayani.", _
                   vbExclamation, FrmTitle
            Exit Sub
        Case csReg
            MsgBox "Maaf!!! Permintaan registrasi anda ditolak." & vbCr & _
                   "Add-In ini telah diregistrasikan sebelumnya oleh" & vbCr & _
                   Vars(cvRegNm), vbExclamation, FrmTitle
            Exit Sub
    End Select
    Nm = Application.UserName
    Cd = "XXXXXXXXXX"
EnterName:
    Nm = UserInp("Nama", Nm)
    If Nm = Empty Then GoTo EndMessage
    If Status = csCpy And LCase(Nm) = LCase(Vars(cvRegNm)) Then _
       If MsgBox("Anda tidak dapat memasukkan nama yang sama" & vbCr & _
                 "dengan nama yang pernah diregistrasikan sebelumnya." & vbCr & _
                 "Berniat memasukkan lagi nama anda?", vbCritical + vbYesNo) = vbYes _
                 Then GoTo EnterName Else Exit Sub
    Cd = UCase(UserInp("Kode Kunci", Cd))
    Cd = Replace(Replace(Cd, " ", Empty), "-", Empty)
    If Cd = Empty Then GoTo EndMessage
    If KeyCode(Nm) = Cd Then
        If Application.UserName = Empty Then Application.UserName = Nm
        Vars(cvUserNm) = Application.UserName
        Vars(cvRegNm) = Nm
        Vars(cvRegCd) = Cd
        Status = csReg
        RestoreOrgProperty cPass
        If Nm = Vars(cvAuthor) Then Msg = "Registration complete Boss....." Else _
           Msg = "Anda sukses melakukan registrasi." & vbCr & _
                 "Add-In ini dilisensikan kepada : " & Nm & vbCr & _
                 "Kode lisensi : " & Left(Cd, 5) & "-" & Right(Cd, 5) & vbCr & _
                 "Selamat menggunakan AddIn ini dan" & vbCr & _
                 "semoga sukses selalu menyertai anda." & vbCr & _
                 "Salam saya : " & Vars(cvAuthor)
        MsgBox Msg, vbInformation, FrmTitle
        Exit Sub
    End If
    If MsgBox("Maaf!!!" & vbCr & _
              "Kode Kunci yang anda masukkan masih salah." & vbCr & _
              "Ulangi registrasi?", vbYesNo + vbCritical, FrmTitle) = vbYes _
              Then GoTo EnterName
EndMessage:
    MsgBox "Untuk melakukan registrasi lagi" & vbCr & _
           "tekan [ Ctrl ] , [ Alt ] , [ Shift ] + [ R ]", vbInformation, FrmTitle
End Sub

Private Property Get GetPass() As Boolean
    Application.EnableCancelKey = xlDisabled
    Dim Inp As String
EnterInfo:
    Inp = Trim(InputBox("Enter Password :", FrmTitle))
    If Inp = Empty Then
        If MsgBox("Are you sure want to exit now?", vbYesNo + vbQuestion, FrmTitle) = _
           vbYes Then Exit Property Else GoTo EnterInfo
    End If
    If IsNumeric(Inp) Then
        If ChkPass(CDbl(Inp)) Then
            GetPass = True
            Exit Property
        End If
    End If
    If MsgBox("Wrong Password!!!" & vbCrLf & "Retry again?", vbYesNo + vbCritical, _
              FrmTitle) = vbYes Then GoTo EnterInfo
End Property

Private Sub InitVars()
    FrmTitle = "Initialyzing '" & Vars(cvTitle) & "'"
    If Not GetPass Then Exit Sub
    Dim Inp As String, Def As String, Txt As String, Idx As Integer, _
        V(cvInitTm - cvBase - 1), I As Integer, NotSvd As Boolean
    For I = 0 To cvInitTm - cvAuthor
        V(I) = Vars(I + cvAuthor)
    Next
    Def = "Author"
EnterVarName:
    Inp = StrConv(Replace(LCase(Trim(InputBox("Enter Variable Name :", FrmTitle, Def))), _
                          " ", Empty), vbProperCase)
    Select Case Inp
        Case Empty
            If MsgBox("Are you sure want to exit now?", vbQuestion + vbYesNo, FrmTitle) = _
               vbNo Then GoTo EnterVarName
            If NotSvd Then
                If MsgBox("Save changes?", vbQuestion + vbYesNo, FrmTitle) = vbYes Then
                    For I = cvAuthor To cvInitTm
                        Vars(I) = V(I - cvAuthor)
                    Next
                    SaveMe cPass
                End If
            End If
            Exit Sub
        Case "Author"
            Idx = cvAuthor
        Case "Callname"
            Inp = "CallName"
            Idx = cvCallNm
        Case "Callnumber"
            Inp = "CallNumber"
            Idx = cvCallNo
        Case "Title"
            Idx = cvTitle
        Case "Comments"
            Idx = cvCmts
        Case "Comment2"
            Idx = cvCmt2
        Case "Comment3"
            Idx = cvCmt3
        Case "Filename"
            Inp = "FileName"
            Idx = cvFileNm
        Case "Username"
            Inp = "UserName"
            Idx = cvUserNm
        Case "Registeredname", "Licensedname"
            Inp = "LicensedName"
            Idx = cvRegNm
        Case "Registeredcode", "Licensedcode"
            Inp = "LicensedCode"
            Idx = cvRegCd
        Case "Firsttime"
            Inp = "FirstTime"
            Idx = cvFstTm
        Case "Inittime"
            Inp = "InitTime"
            Idx = cvInitTm
        Case Else
            MsgBox "Variable(" & Inp & ") is not defined!", vbExclamation, FrmTitle
            GoTo SetDefault
    End Select
    Idx = Idx - cvAuthor
    Txt = Trim(InputBox("Variable(" & Inp & ") =", FrmTitle, V(Idx)))
    If Txt <> V(Idx) Then
        Beep
        V(Idx) = Txt
        NotSvd = True
    End If
SetDefault:
    Def = Inp
    GoTo EnterVarName
End Sub

Private Sub GetKeyCode()
    FrmTitle = "Get User Key Code of '" & Vars(cvTitle) & "'"
    If Not GetPass Then Exit Sub
    Dim Inp As String, Def As String
    Def = Application.UserName
EnterUserName:
    Inp = Trim(InputBox("Enter User Name :", FrmTitle, Def))
    If Inp = Empty Then If MsgBox("Are you sure want to exit now?", _
       vbQuestion + vbYesNo, FrmTitle) = vbNo Then GoTo EnterUserName Else Exit Sub
DispUserCode:
    If InputBox("User Key Code for '" & Inp & "' is :", FrmTitle, _
       Left(KeyCode(Inp), 5) & "-" & Right(KeyCode(Inp), 5)) = Empty Then _
       If MsgBox("Are you sure want to exit now?", vbQuestion + vbYesNo, FrmTitle) = _
          vbNo Then GoTo DispUserCode Else Exit Sub
    Def = Inp
    GoTo EnterUserName
End Sub

Private Sub ClearVars()
    FrmTitle = "Clear User Data of '" & Vars(cvTitle) & "'"
    If Not GetPass Then Exit Sub
    If MsgBox("WARNING!!! All User Defined Data will be cleared." & vbCr & _
              "(This Add-In will be reset to First Time Mode)" & vbCr & _
              "Are you sure?", vbExclamation + vbYesNo, FrmTitle) = vbYes Then
        Vars(cvUserNm) = Empty
        Vars(cvRegNm) = Empty
        Vars(cvRegCd) = Empty
        Vars(cvFstTm) = Empty
        Vars(cvInitTm) = Empty
        Status = Empty
        ThisWorkbook.Comments = AddInCmts
        SaveMe cPass
    End If
End Sub

Private Sub WbkOpen(Pass As Double)
    If Not ChkPass(Pass) Then Exit Sub
    Dim Resp As Integer, Msg As String
    Status = GetStatus
    RestoreOrgProperty cPass
    FrmTitle = Vars(cvTitle)
    Select Case Status
        Case csFst
            Msg = "Anda menjalankan Add-In ini untuk pertama kalinya." & vbCr & _
                  "Anda mempunyai bonus waktu pemakaian selama " & _
                      CStr(CInt(cLmtExp)) & " hari," & vbCr & _
                  "selebihnya anda harus meregistrasikannya dahulu." & vbCr & _
                  "Registrasi sekarang?"
            Vars(cvFstTm) = CStr(Now)
        Case csTrl
            Msg = CStr(CInt(Remain)) & _
                      " hari lagi bonus waktu pemakaian Add-In ini akan habis." & vbCr & _
                  "Untuk dapat terus menggunakannya," & vbCr & _
                  "anda harus meregistrasikannya dahulu." & vbCr & _
                  "Registrasi sekarang?"
        Case csExp
            Msg = "Maaf!!! Bonus waktu pemakaian Add-In ini telah habis." & vbCr & _
                  "Untuk dapat terus menggunakannya," & vbCr & _
                  "anda harus meregistrasikannya dahulu." & vbCr & _
                  "Registrasi sekarang?"
        Case csErr
            Msg = "Maaf!!! Periksa dahulu sistem penanggalan dan" & vbCr & _
                  "jam komputer anda sebelum menggunakan Add-In ini."
        Case csDmg
            Msg = "Maaf!!! Oleh karena sesuatu sebab," & vbCr & _
                  "Add-In ini telah mengalami perubahan" & vbCr & _
                  "sehingga tidak dapat digunakan lagi."
        Case csCpy
            Msg = "Maaf!!! Add-In ini terdaftar atas nama orang lain." & vbCr & _
                  "Anda harus meregistrasikannya sendiri bila anda" & vbCr & _
                  "merencanakan hendak menggunakannya." & vbCr & _
                  "Registrasi sekarang?"
    End Select
    Select Case Status
        Case csFst, csTrl, csExp, csCpy
            Vars(cvInitTm) = CStr(Now)
            If MsgBox(Msg, vbQuestion + vbYesNo, FrmTitle) = vbYes Then RegNow
        Case csErr, csDmg
            MsgBox Msg, vbExclamation + vbOKOnly, FrmTitle
    End Select
    Application.OnKey "^%+R", "RegNow"
    Application.OnKey "^%+V", "InitVars"
    Application.OnKey "^%+K", "GetKeyCode"
    Application.OnKey "^%+C", "ClearVars"
End Sub
