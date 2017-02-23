Attribute VB_Name = "Module1"
Public Sub Start()
    Range("A2").Value = 0
    Call Setup
    Randomize
    it = 0      'iterator
    'granice paletki --> top = 34, bot = 222,
    'center --> top = 166, left = 456
    'granice pi³ki--> poziom 196 / 702, pion 34 / 284
    ' PAD 12/15/20/15/12
    Do
    DoEvents
        it = it + 1
        Call Ball
        Call Paletki

        timeout (0.01)
    Range("L3") = Range("c13").Value - (it / 100)
    Loop Until it = Range("c13").Value * 100
    'MsgBox Arkusz1.Shapes("Lewa").Left
End Sub
Sub timeout(duration_ms As Double)
    Start_Time = Timer
    Do
    DoEvents
    Loop Until (Timer - Start_Time) >= duration_ms
End Sub
Sub Ball()
    bounce = Range("A36").Value
    speed = Range("A37").Value
    mecz = Range("A38").Value
    Arkusz1.Shapes("Pilka").Left = Arkusz1.Shapes("Pilka").Left + speed
    Arkusz1.Shapes("Pilka").Top = Arkusz1.Shapes("Pilka").Top + bounce
    ball_x = Arkusz1.Shapes("Pilka").Left
    ball_y = Arkusz1.Shapes("Pilka").Top
    l_pad = Arkusz1.Shapes("Lewa").Top
    r_pad = Arkusz1.Shapes("Prawa").Top
    'Kolizja z Paletk¹
    a = 30
    If speed < 0 Then   'leci w lewo
        If ball_x < 212 Then
            If ball_y < (l_pad + 74) Then
                If ball_y > l_pad Then
                    speed = speed * (-1)
                    If ball_y < l_pad + 27 Then
                    bounce = bounce + 1
                    End If
                End If
                If ball_y > l_pad + 74 - 27 Then
                    bounce = bounce - 1
                End If
                
            End If
        End If
    Else                'leci w prawo
        If ball_x > 681 Then
            If ball_y < (r_pad + 74) Then
                If ball_y > r_pad Then
                    speed = speed * (-1)
                    If ball_y < r_pad + 27 Then
                    bounce = bounce + 1
                    End If
                If ball_y > r_pad + 74 - 27 Then
                    bounce = bounce - 1
                End If
                
                End If
            End If
        End If
        
    End If
    If ball_y < 34 Then
        bounce = bounce * (-1)
    End If
    If ball_y > 283 Then
        bounce = bounce * (-1)
    End If
    'PUNKTY-------------------
    If ball_x > 715 Then
            Range("L1").Value = Range("L1").Value + 1
            timeout (0.5)
            mecz = 0
            timeout (0.5)
            Range("A37").Value = -1
    End If
    If ball_x < 193 Then
            Range("I1").Value = Range("I1").Value + 1
            timeout (0.5)
            mecz = 0
            timeout (0.5)
            Range("A37").Value = 1
    End If
    'END PUNKTY---------------
    'Restart po zdobyciu punktu
    If mecz = 0 Then
        Call Goal
        mecz = 2
        speed = speed * (-1)
        b = 1 - Int(2 * Rnd) 'bounce
        If b = 1 Then
        Else
            b = -1
        End If
        bounce = b
    End If
    'End Retart
Range("A36").Value = bounce
Range("A37").Value = speed
Range("A38").Value = mecz
End Sub
Sub Setup()
    Call Goal
    szer = 14   ' szerokoœæ paletki
    wys = 74    ' wysokoœæ paletki

    'okreœlanie kszta³tu paletki
        Arkusz1.Shapes("Prawa").Height = wys
        Arkusz1.Shapes("Prawa").Width = szer
        Arkusz1.Shapes("Lewa").Height = wys
        Arkusz1.Shapes("Lewa").Width = szer
    'i pi³ki
        Arkusz1.Shapes("Pilka").Height = szer
        Arkusz1.Shapes("Pilka").Width = szer
    'Data
    b = 1 - Int(2 * Rnd) 'bounce
    If b = 1 Then
    Else
        b = -1
    End If
    s = 1 - Int(2 * Rnd) 'bounce
    If s = 1 Then
        s = 3
    Else
        s = -3
    End If
    Range("A36").Value = b
    Range("A37").Value = s * Range("C16").Value  'speed
    Range("A38").Value = 2  'mecz bool
    Range("L1").Value = 0   'R point
    Range("I1").Value = 0   'L point
End Sub
Sub Goal()
    szer = 14
    wys = 74
    l_bord = 193    'lewa granica
    r_bord = 718    'prawa granica
    Gap = 5         'odleg³oœæ miêdzy paletk¹a granic¹
    set_ball_x = 455 - (szer / 2)
    ball_x = set_ball_x
    set_ball_y = 166 - (szer / 2)
    ball_y = set_ball_y
    set_pad = 166 - (wys / 2)
    l_pad = set_pad
    r_pad = set_pad
    'pozycja paletek
        Arkusz1.Shapes("Lewa").Left = l_bord + Gap
        Arkusz1.Shapes("Prawa").Left = r_bord - Gap - szer
        Arkusz1.Shapes("Lewa").Top = set_pad
        Arkusz1.Shapes("Prawa").Top = set_pad
    'pi³ka
        Arkusz1.Shapes("Pilka").Left = ball_x
        Arkusz1.Shapes("Pilka").Top = ball_y
End Sub
Sub Paletki()
    ball_y = Arkusz1.Shapes("Pilka").Top
    l_pad = Arkusz1.Shapes("Lewa").Top
    r_pad = Arkusz1.Shapes("Prawa").Top
    speed = Range("A37").Value
    pspeed = Range("C18").Value
    If speed < 0 Then
        If l_pad < 224 Then
            If l_pad > 34 Then
                If ball_y > l_pad + 37 Then
                    l_pad = l_pad + pspeed
                Else
                    l_pad = l_pad - pspeed
                End If
            Else
                l_pad = l_pad + pspeed
            End If
        Else
            l_pad = l_pad - pspeed
        End If
    Else
        If r_pad < 224 Then
            If r_pad > 34 Then
                If ball_y > r_pad + 37 Then
                    r_pad = r_pad + pspeed
                Else
                    r_pad = r_pad - pspeed
                End If
            Else
                r_pad = r_pad + pspeed
            End If
        Else
            r_pad = r_pad - pspeed
        End If
    End If
    Arkusz1.Shapes("Lewa").Top = l_pad
    Arkusz1.Shapes("Prawa").Top = r_pad
End Sub


