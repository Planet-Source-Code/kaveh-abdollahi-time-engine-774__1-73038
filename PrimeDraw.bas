Attribute VB_Name = "PrimeDraw"
'******************************************************************************************
'
'   Copyright(C) 2010 By Kaveh Abdollahi.   Kavehplus@gmal.com
'   Time Eangine
'   March 2010
'
'******************************************************************************************



Option Explicit

Private A As Double, B As Double, coc(0 To 9) As Long, cocS(0 To 255) As Long, BStp As Long
Private Stp As Long, x As Double, y As Double, yen As Long, yst As Long
Private rad As Double, SI1 As Double, SI2 As Double, SI3 As Double, SI4 As Double, SI5 As Double, SI6 As Double
Private pd As Double, cTim As Double, xTm As Double, yTm As Double
Private E As Long, m As Long, Cc As Long, Ti2 As Long, Ti3 As Long, Ti4 As Long, Ti5 As Long
Private xCntr As Double, yCntr As Double, w(1 To 10) As Double, sz As Long
Private bBol(1 To 10) As Byte, tmp1 As Double, tmp2 As Double, tmp3 As Double, tmp4 As Double
Private pot As POINTAPI, co As Long, co2 As Long, co3 As Long, co4 As Long, co5 As Long
Private cA As Long, cR As Long, cG As Long, cb As Long, S As String
Private PCol(0 To 256, 0 To 2) As Long




Public Sub DrawP1()
    
    With frmBase
    
    On Error Resume Next
    If .chkShotAll.Value And .chkAutoShot.Value Then .cmdSF_Click


    LQT = LQT + 1
    If LQT > 148900 Then LQT = 1
    If LQT2 > 148900 Then LQT2 = 1
    .txtspm(13) = LQT
    pd = (.txtspm(2))
    If .chkAvalue Then LQT2 = LQT2 + pd
    .txtspm(11) = LQT2: .txtspm(11).Refresh
    .txtLQT2 = Format$(LQT2, "###,###0") & vbCrLf & _
                Format$(Primes(LQT2), "###,###,###0") & vbCrLf & _
                Format$(PrK(3, LQT2), "####") & vbCrLf & _
                Format$(PrK(2, LQT2), "####") & vbCrLf & _
                Format$(sz, "###,###,###0")
    .txtLQT2.Refresh
    
    .lblFullscr(1) = Round(LQT2, 2): .lblFullscr(1).Refresh
    .lblFullscr(2) = Primes(LQT2): .lblFullscr(2).Refresh
    .lblFullscr(3) = PrK(2, LQT2): .lblFullscr(3).Refresh
    .lblFullscr(4) = PrK(3, LQT2): .lblFullscr(4).Refresh
    
    sz = 0
    '''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''
    
    If .ChkDraw(4) Then BitBlt picTmp.hdc, 0, 0, 1024, 768, picTmp.hdc, 0, 0, 0
    
    If .chkAutoMax Then
          yen = .txtspm(11)
        Else
          yen = .txtspm(21)
    End If
    '''''''''''''''''''''''''''''''''''''''
    If .chkLastP Then
         yst = yen - .txtspm(32)
      Else
         yst = .txtspm(28)
    End If

    cTim = LQT2 / 1000000
    SI1 = 1: SI2 = 1: SI3 = 1: SI4 = 1: SI5 = 1: SI6 = 1
    If IsNumeric(.txtspm(16)) Then SI1 = .txtspm(16) '* 2
    If IsNumeric(.txtspm(17)) Then SI2 = .txtspm(17) '* 2
    If IsNumeric(.txtspm(18)) Then SI3 = .txtspm(18) '* 2
    If IsNumeric(.txtspm(19)) Then SI4 = .txtspm(19) '* 2
    If IsNumeric(.txtspm(20)) Then SI5 = .txtspm(20) '* 2
    If IsNumeric(.txtspm(33)) Then SI6 = .txtspm(33) '* 2

    bBol(1) = .chkCol(1)
    bBol(2) = .chkCol(2)
    bBol(3) = .chkCol(3)
    bBol(4) = .chkCol(4)
    bBol(5) = .chkCol(5)
    bBol(0) = .chkCol(0)
    
    bBol(6) = .chkBox

    xCntr = .txtspm(22)
    yCntr = .txtspm(23)
                                                                                                                
    co5 = RGB((LQT2 Mod 4096) / 16, (LQT2 Mod 2048) / 8, (LQT2 Mod 8192) / 32)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    For B = yst + Stp To yen Step 1
        
        BStp = PrK(3, B)
        Stp = PrK(2, B)
        m = Stp * (SI2)
        Cc = BStp * (SI3)
        If .chkCM Then Cc = Stp * (SI3)
        Cc = Cc * Cc * (SI4)
        E = (m * Cc) * (SI1)
        
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        co4 = RGB(BStp, BStp, BStp) Xor RGB(Stp, Stp, Stp)
        
        If bBol(1) Then
            co = RGB(256 - Stp * 7, 256 \ Stp * 3, Stp * 5)
            co2 = RGB(256 - Stp * 3, 256 - Stp * 5, Stp * 7)
            co3 = RGB(256 - Stp * 7, 256 \ Stp * 7, 256 - Stp * 3)
        ElseIf bBol(2) Then
            co = RGB(256 - Stp * 7, 256 \ Stp * 7, Stp * 5)
            co2 = RGB(256 - Stp * 7, 256 \ Stp * 7, Stp * 5)
            co3 = RGB(256 - Stp * 7, 256 \ Stp * 7, Stp * 5)
        ElseIf bBol(3) Then
            co = RGB(256 - BStp * 7, 256 \ BStp * 3, BStp * 5)
            co2 = RGB(256 - BStp * 3, 256 \ BStp * 5, BStp * 7)
            co3 = RGB(256 - BStp * 7, 256 \ BStp * 3, BStp * 3)
        ElseIf bBol(4) Then
            co = RGB(256 - Stp * 3, 256 \ Stp * 5, Stp * 5)
            co2 = RGB(256 - Stp * 7, 256 \ Stp * 3, Stp * 7)
            co3 = RGB(256 - Stp * 5, 256 \ Stp * 7, Stp * 3)
        ElseIf bBol(5) Then
            co = RGB(256 - Cc \ Stp, Cc \ Stp, 256 - Cc \ Stp)
            co2 = RGB(Cc \ Stp, 256 - E \ Cc, Cc \ Stp)
            co3 = RGB(E \ Stp, E \ Stp * 2, 256 - Cc \ Stp)
        ElseIf bBol(0) Then
            co = RGB(Cc, E, m)
            co2 = RGB(E, m, Cc)
            co3 = RGB(m, Cc, E)
        Else
            co = RGB(Cc \ m, E \ Cc, E \ m) ' RGB(Cc \ Stp, Cc \ Stp, Cc \ Stp)
            co2 = RGB(E \ Cc, E \ m, Cc \ m) ' RGB(Cc \ Stp, 256 - Cc \ Stp * 2, Cc \ Stp)
            co3 = RGB(Cc \ m, Cc \ m, E \ Cc) '  RGB(E \ Stp, 256 - E \ Stp * 2, 256 - Cc \ Stp)
        End If
        
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        If .chkPant(0) Then
            x = Sin(E * cTim) * Cos(Cc * cTim) * Cc \ SI6 + xCntr
            y = Cos(E * cTim) * Cos(Cc * cTim) * Cc \ SI6 + yCntr
        ElseIf .chkPant(1) Then
            x = Sin(E * cTim + B) * Cos(Cc * cTim - B) * Cc \ SI6 + xCntr
            y = Cos(E * cTim - B) * Cos(Cc * cTim + B) * Cc \ SI6 + yCntr
        ElseIf .chkPant(2) Then
            x = Sin(E * cTim) * Cos(Cc * cTim) * Cc \ SI6 + xCntr
            y = Cos(E * cTim) * Sin(Cc * cTim) * Cc \ SI6 + yCntr
        Else
            x = Sin(E * cTim) * Cos(Cc * cTim) * Cc \ SI6 + xCntr
            y = Cos(E * cTim) * Sin(Cc * cTim) * Cc \ SI6 + yCntr
        End If
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If .chkTimeEnable(1) Then
            picTmp.ForeColor = co2 Xor RGB(E \ Cc, Cc \ m, E \ m)
            picTmp.FillColor = co And RGB(E \ Cc, Cc \ m, E \ m)
            xTm = x
            yTm = y
            Ellipse picTmp.hdc, xTm - m \ 2 \ SI6, yTm - m \ 2 \ SI6, xTm + m \ 2 \ SI6, yTm + m \ 2 \ SI6
        End If
        
        ''''''''''''''''''''''''''''''''' Draw Strings '''''''''''''''''''''''''''''''''''''''
        If .chkTimeEnable(0) Then
            If B = yst Then MoveToEx picTmp.hdc, x, y, pot
            picTmp.ForeColor = co3 And co5
            LineTo picTmp.hdc, x, y
        End If
        If .chkTimeEnable(2) Then
            If B = yst Then MoveToEx picTmp.hdc, x, y, pot
            picTmp.ForeColor = co2 Xor co
            LineTo picTmp.hdc, x, y
        End If
        If .chkTimeEnable(4) Then
            If B = yst Then MoveToEx picTmp.hdc, x, y, pot
            picTmp.ForeColor = co3 Xor co4 And co5
            LineTo picTmp.hdc, x, y
        End If
        If .chkTimeEnable(6) Then
            SetPixel picTmp.hdc, x, y, co3 Xor co4 And co5
            sz = sz + 1
        End If
        DoEvents
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       
        If .chkTimeEnable(3) Or .chkTimeEnable(5) Then
            SetPixel picTmp.hdc, x, y, co Xor co4: sz = sz + 1
            For Ti2 = B To B + PrK(2, B)
                
                m = PrK(2, Ti2) * SI2
                Cc = PrK(3, Ti2) * SI3
                Cc = Cc * Cc * SI4
                E = (m * Cc) * SI1
                xTm = Cos(-Ti2) * Cos(m + cTim) * m \ SI6 + x
                yTm = Sin(Ti2) * Cos(m - cTim) * m \ SI6 + y
                
                If .chkTimeEnable(3) Then SetPixel picTmp.hdc, xTm, yTm, co4 Xor PCol(PrK(2, Ti2 * 4) * 2, 0) Xor co3
                
                If .chkTimeEnable(5) Then
                    For Ti3 = Ti2 To Ti2 + PrK(2, Ti2)
                        m = PrK(2, Ti3) * SI2
                        Cc = PrK(3, Ti3) * SI3
                        Cc = Cc * Cc * SI4
                        E = (m * Cc) * SI1
                        x = (Cos(Ti3) + Cos(m - Ti2)) * m \ SI6 + xTm
                        y = (Sin(-Ti3) + Cos(m + Ti2)) * m \ SI6 + yTm
                         
                        SetPixel picTmp.hdc, x, y, PCol(Stp * 8 - PrK(2, Ti3) * 4, 0) Xor co3
                        sz = sz + 1
                     Next Ti3
                 End If
            
            Next Ti2
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        End If
enx:
      B = B + Stp * SI5
    
    Next B
    
    If .chkAutoFix Then .txtspm(33) = BStp \ 8
    If .chkALog Then Loger
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

CycleST

     Blend.SourceConstantAlpha = Val(.txtspm(7))
     Blend.AlphaFormat = 0
     If .chkAlphaEnable Then Blend.AlphaFormat = 1

     If .chkAlpha Then Blend.SourceConstantAlpha = (Stp + Cc \ m + BStp) \ 3
     CopyMemory BlendPtr, Blend, 4

     i = .txtspm(25)
         StretchBlt .picBuffEE.hdc, 0, 0, .Width \ 15 + 1, .Height \ 15 + 1, _
              picBuff.hdc, 0, 0, 1024 \ i, 768 \ i, vbSrcCopy

         AlphaBlend picView.hdc, 0, 0, 1024, 768, _
              .picBuffEE.hdc, 0, 0, 1024, 768, BlendPtr

'''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''

     Nvg(1) = .txtspm(36): Nvg(2) = .txtspm(34): Nvg(3) = .txtspm(35)

     If .fraTelo.Visible = True Then
         AlphaBlend .picTele.hdc, 0, 0, 256, 256, _
         .picBuffEE.hdc, Nvg(2), Nvg(3), Nvg(1), Nvg(1), BlendPtr
     End If

     If bBol(6) Then
        picView.ForeColor = vbYellow
        MoveToEx picView.hdc, Nvg(2), Nvg(3), pot '0

        LineTo picView.hdc, Nvg(2) + Nvg(1), Nvg(3)
        LineTo picView.hdc, Nvg(2) + Nvg(1), Nvg(3) + Nvg(1)

        LineTo picView.hdc, Nvg(2), Nvg(3) + Nvg(1)

        LineTo picView.hdc, Nvg(2), Nvg(3)
        LineTo picView.hdc, Nvg(2), Nvg(3) + Nvg(1)
     End If

     CycleED
     Process(5, 1) = Round(tFa, 2)
     
     
     End With


End Sub

Public Sub SetCols()
Dim x As Integer, y As Integer, co As Long, TR As Long

    For y = 0 To 256
       PCol(y, 0) = RGB(y, y, y)
       PCol(y, 1) = RGB(256 - y, 256 - y, 256 - y)
       PCol(y, 2) = PCol(y, 0) Xor PCol(y, 1)
    Next y

End Sub
