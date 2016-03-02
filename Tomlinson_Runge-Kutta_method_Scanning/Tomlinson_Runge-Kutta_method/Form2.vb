Public Class Form1
    ' Tomlinson_Runge=Kutta-metdtod
    ' 作成日 : 20150703  ユーザー名 : 棚原翔平
    ' 修正日 : 20150906  ユーザー名 : 棚原翔平
    ' 修正日 : 20160202  ユーザー名 : 棚原翔平


    '----------- Fileの変数 ----------------------------------
    Shared writeFileName As String
    Shared writeFileNum As Integer
    Shared readFileNameN As String ' Ngraph data読み込みファイル
    Shared readFileNumN As Integer
    Shared writeFileNameN As String ' Ngraph data書き込みファイル
    Shared writeFileNumN As Integer
    Shared NewFile As String

    '------パラメータ変数を定義---------
    Shared Ti As Double
    Shared xt As Double
    Shared vt As Double
    Shared xs As Double
    Shared vs As Double
    Shared dTi As Double
    Shared n As Double

    Shared a As Double
    Shared f As Double
    Shared m As Double
    Shared K As Double
    Shared H As Double
    Shared Uo As Double
    Shared x As Double
    Shared Pi As Double


    Shared Hn As Double
    Shared u As Double
    Shared Kk As Double
    Shared xso As Double
    Shared t As Double
    Shared dt As Double
    Shared gamma As Double
    Shared R As Double


    Shared Xxt As Double
    Shared Vvt As Double
    Shared Xxs As Double
    Shared Vvs As Double
    Shared fsub As Double

    Shared fourier_C As Double    'フーリエ係数
    Shared fourier_S As Double
    Shared c As Double          'フィッテングに用いる比例定数

    Shared Scan_width As Double 'スキャン幅
    Shared d As Double          '分割数
    Shared Scan_step As Double 'ステップ幅
    Shared Gap As Double        'ステップ位置
    Shared UU As Double
    '------------------------------



    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Button1.Enabled = False
        Button1.Text = "計算中"

        '-----スキャン条件の設定-------
        Scan_width = TextBox21.Text * 10 ^ -9
        d = TextBox22.Text
        Scan_step = Scan_width / d
        Gap = TextBox24.Text * 10 ^ -9

        '-----初期値の設定-------
        Ti = TextBox18.Text * 10 ^ -6
        Xxt = TextBox17.Text * 10 ^ -9
        vt = TextBox16.Text
        Xxs = TextBox15.Text * 10 ^ -9
        vs = TextBox14.Text
        dTi = TextBox13.Text * 10 ^ -6
        n = TextBox12.Text


        '-----各パラメータの設定-------
        a = TextBox1.Text * 10 ^ -9
        f = TextBox2.Text * 10 ^ 6
        m = TextBox3.Text
        K = TextBox4.Text
        H = TextBox5.Text
        Uo = TextBox6.Text
        x = TextBox7.Text * 10 ^ -9
        gamma = TextBox23.Text   'ローレンツ関数のγ
        R = TextBox25.Text


        Pi = 3.1415926535

        '-----パラメータの無次元化------
        Hn = H / (m * 2 * Pi * f)
        u = 4 * Pi ^ 2 * Uo / (m * (2 * Pi * f * a) ^ 2)
        Kk = K / (m * (2 * Pi * f) ^ 2)
        xso = 2 * Pi * x / a
        t = 2 * Pi * f * Ti
        dt = 2 * Pi * f * dTi
        xt = Xxt * 2 * Pi / a
        xs = Xxs * 2 * Pi / a


        '------無次元化パラメータの表示？-----
        TextBox11.Text = Hn
        TextBox10.Text = u
        TextBox9.Text = Kk
        TextBox8.Text = xso

        '-----フーリエ係数のリセット------
        fourier_C = 0
        fourier_S = 0

        '-----基板が受ける力をリセット--------
        fsub = 0

        ' 日付を取得する
        Dim dtNow As DateTime = DateTime.Today
        ' 指定した書式で日付を文字列に変換する
        Dim Today As String = dtNow.ToString("yyMMdd")
        '-----------ファイルの作成---------------
        Dim sfd As New SaveFileDialog
        sfd.FileName = "Tomlinson_1d-Scanning_" & Today & "_101.dat"

        Dim btn As DialogResult = sfd.ShowDialog()
        If btn = Windows.Forms.DialogResult.OK Then
            MessageBox.Show(sfd.FileName & "に保存します")
        ElseIf btn = Windows.Forms.DialogResult.Cancel Then
            MessageBox.Show("キャンセルされました")
            Button1.Enabled = True
            Button1.Text = "1次元スキャン開始"
            Exit Sub
        End If

        writeFileName = sfd.FileName
        Dim StrArray() As String
        StrArray = Split(writeFileName, "\")
        NewFile = StrArray(UBound(StrArray))
        NewFile = Microsoft.VisualBasic.Left(NewFile, Len(NewFile) - 4)


        '----------　空いているファイル番号を取得 ------------------
        writeFileNum = FreeFile()
        FileOpen(writeFileNum, writeFileName, OpenMode.Output)

        '-------- コメントの記入 ----------------------------------
        PrintLine(writeFileNum, "#書き込みファイル名   :  ", writeFileName)
        PrintLine(writeFileNum, "#日時                :  ", CStr(Now))
        PrintLine(writeFileNum, "#スキャン幅 (nm)　     :  ", Format(Scan_width, "0.00E+00"))
        PrintLine(writeFileNum, "#分割数           　  :  ", Format(d, "0.00"))
        PrintLine(writeFileNum, "#ステップ幅 (nm)　     :  ", Format(Scan_step, "0.00E+00"))
        PrintLine(writeFileNum, "#格子定数a (nm)　     :  ", Format(a, "0.00E+00"))
        PrintLine(writeFileNum, "#共振周波数f (MHz)　  :  ", Format(f, "0.0000"))
        PrintLine(writeFileNum, "#質点の質量m(kg)      :  ", Format(m, "0.00E+00"))
        PrintLine(writeFileNum, "#ばね定数K(N/m)       :  ", Format(K, "0.00E+00"))
        PrintLine(writeFileNum, "#粘性定数H(Ns/m)      :  ", Format(H, "0.0000E+00"))
        PrintLine(writeFileNum, "#基板ポテンシャルUo(J):  ", Format(Uo, "0.0000E+00"))
        PrintLine(writeFileNum, "#基板振幅X (nm)       :  ", Format(x, "0.0000E+00"))
        PrintLine(writeFileNum, "#ローレンチアンgamma   :  ", Format(gamma, "0.0000E+00"))

        'データの記入
        PrintLine(writeFileNum, "#")
        PrintLine(writeFileNum, "# 1.基板位置Xxs(m) 2.xs 3.基板の受ける力fsub 4.逆位相<fsub*cos(t)/xso> 5.同相<fsub*sin(t)/xso> 6.基板ポテンシャルUU(J)")

        FileClose(writeFileNum)


        '--------------------------------------------------------------------------------------------------------
        For j = 0 To d

            '次の基板位置へ移動
            Xxs = j * Scan_step

            'あるところにポテンシャルの異常を作る
            Uo = TextBox6.Text * (1 + 1 / Pi * (gamma / ((Xxs - Gap) ^ 2 + gamma ^ 2)) * R) '基板ポテンシャル(sin波)にローレンチアンを掛ける


                '-----初期値の設定-------
                Ti = TextBox18.Text * 10 ^ -6
                Xxt = TextBox17.Text * 10 ^ -9
                vt = TextBox16.Text
                vs = TextBox14.Text
                dTi = TextBox13.Text * 10 ^ -6
                n = TextBox12.Text


                '-----各パラメータの設定-------
                a = TextBox1.Text * 10 ^ -9
                f = TextBox2.Text * 10 ^ 6
                m = TextBox3.Text
                K = TextBox4.Text
                H = TextBox5.Text
                'Uo = TextBox6.Text
                x = TextBox7.Text * 10 ^ -9

                Pi = 3.1415926535

                '-----パラメータの無次元化------
                Hn = H / (m * 2 * Pi * f)
                u = 4 * Pi ^ 2 * Uo / (m * (2 * Pi * f * a) ^ 2)
                Kk = K / (m * (2 * Pi * f) ^ 2)
                xso = 2 * Pi * x / a
                t = 2 * Pi * f * Ti
                dt = 2 * Pi * f * dTi
                xt = Xxt * 2 * Pi / a
                xs = Xxs * 2 * Pi / a

                '------無次元化パラメータの表示？-----
                TextBox11.Text = Hn
                TextBox10.Text = u
                TextBox9.Text = Kk
                TextBox8.Text = xso

                '-----フーリエ係数のリセット------
                fourier_C = 0
                fourier_S = 0

                '-----基板が受ける力をリセット--------
                fsub = 0

                '------------------ルンゲ=クッタ法による解析-------------------------------
                For i = 0 To n


                    'ルンゲ=クッタ法
                    Dim kxt1 As Double
                    Dim lxt1 As Double
                    Dim kxs1 As Double
                    Dim lxs1 As Double
                    Dim fsub1 As Double
                    Dim kxt2 As Double
                    Dim lxt2 As Double
                    Dim kxs2 As Double
                    Dim lxs2 As Double
                    Dim fsub2 As Double
                    Dim kxt3 As Double
                    Dim lxt3 As Double
                    Dim kxs3 As Double
                    Dim lxs3 As Double
                    Dim fsub3 As Double
                    Dim kxt4 As Double
                    Dim lxt4 As Double
                    Dim kxs4 As Double
                    Dim lxs4 As Double
                    Dim fsub4 As Double


                    kxt1 = dt * Gx(t, xt, vt, xs, vs, fsub)
                    lxt1 = dt * Fx(t, xt, vt, xs, vs, fsub)
                    kxs1 = dt * Gy(t, xt, vt, xs, vs, fsub)
                    lxs1 = dt * Fy(t, xt, vt, xs, vs, fsub)
                    fsub1 = dt * Fs(t, xt, vt, xs, vs, fsub)

                    kxt2 = dt * Gx(t + dt / 2, xt + kxt1 / 2, vt + lxt1 / 2, xs + kxs1 / 2, vs + lxs1 / 2, fsub + fsub1 / 2)
                    lxt2 = dt * Fx(t + dt / 2, xt + kxt1 / 2, vt + lxt1 / 2, xs + kxs1 / 2, vs + lxs1 / 2, fsub + fsub1 / 2)
                    kxs2 = dt * Gy(t + dt / 2, xt + kxt1 / 2, vt + lxt1 / 2, xs + kxs1 / 2, vs + lxs1 / 2, fsub + fsub1 / 2)
                    lxs2 = dt * Fy(t + dt / 2, xt + kxt1 / 2, vt + lxt1 / 2, xs + kxs1 / 2, vs + lxs1 / 2, fsub + fsub1 / 2)
                    fsub2 = dt * Fs(t + dt / 2, xt + kxt1 / 2, vt + lxt1 / 2, xs + kxs1 / 2, vs + lxs1 / 2, fsub + fsub1 / 2)

                    kxt3 = dt * Gx(t + dt / 2, xt + kxt2 / 2, vt + lxt2 / 2, xs + kxs2 / 2, vs + lxs2 / 2, fsub + fsub2 / 2)
                    lxt3 = dt * Fx(t + dt / 2, xt + kxt2 / 2, vt + lxt2 / 2, xs + kxs2 / 2, vs + lxs2 / 2, fsub + fsub2 / 2)
                    kxs3 = dt * Gy(t + dt / 2, xt + kxt2 / 2, vt + lxt2 / 2, xs + kxs2 / 2, vs + lxs2 / 2, fsub + fsub2 / 2)
                    lxs3 = dt * Fy(t + dt / 2, xt + kxt2 / 2, vt + lxt2 / 2, xs + kxs2 / 2, vs + lxs2 / 2, fsub + fsub2 / 2)
                    fsub3 = dt * Fs(t + dt / 2, xt + kxt2 / 2, vt + lxt2 / 2, xs + kxs2 / 2, vs + lxs2 / 2, fsub + fsub2 / 2)

                    kxt4 = dt * Gx(t + dt, xt + kxt3, vt + lxt3, xs + kxs3, vs + lxs3, fsub + fsub3)
                    lxt4 = dt * Fx(t + dt, xt + kxt3, vt + lxt3, xs + kxs3, vs + lxs3, fsub + fsub3)
                    kxs4 = dt * Gy(t + dt, xt + kxt3, vt + lxt3, xs + kxs3, vs + lxs3, fsub + fsub3)
                    lxs4 = dt * Fy(t + dt, xt + kxt3, vt + lxt3, xs + kxs3, vs + lxs3, fsub + fsub3)
                    fsub4 = dt * Fs(t + dt, xt + kxt3, vt + lxt3, xs + kxs3, vs + lxs3, fsub + fsub3)

                    t = t + dt

                    'それぞれの値を計算
                    vt = vt + (kxt1 + 2 * (kxt2 + kxt3) + kxt4) / 6
                    xt = xt + (lxt1 + 2 * (lxt2 + lxt3) + lxt4) / 6
                    vs = vs + (kxs1 + 2 * (kxs2 + kxs3) + kxs4) / 6
                    xs = xs + (lxs1 + 2 * (lxs2 + lxs3) + lxs4) / 6

                    '次元を戻す
                'Ti = t / (2 * Pi * f)
                'Xxt = xt * a / (2 * Pi)
                'Xxs = xs * a / (2 * Pi)


                    '1周期分のフーリエ係数を求める　<fsub*cos(t)/xso>および<fsub*sin(t)/xso>
                    If t > 2 * Pi AndAlso t < 4 * Pi Then
                        fsub = fsub + (fsub1 + 2 * (fsub2 + fsub3) + fsub4) / 6
                        fourier_C = fourier_C + (1 / (2 * Pi)) * fsub / xso * Math.Cos(t) * dt '基板の振幅xsをsinにしているのでこちらが逆位相成分つまりエネルギー散逸変化
                        fourier_S = fourier_S + (1 / (2 * Pi)) * fsub / xso * Math.Sin(t) * dt '基板の振幅xsをsinにしているのでこちらが同相成分つまり周波数変化
                    End If


            Next i

            TextBox19.Text = fourier_C
            TextBox20.Text = fourier_S


            '基板ポテンシャルの変化

            'UU = -u * Math.Sin(xt - xs) / (4 * Pi ^ 2) * (m * (2 * Pi * f * a) ^ 2) '次元を戻す
            UU = -u * Math.Asin(Math.Sin(xs) / (4 * Pi ^ 2) * (m * (2 * Pi * f * a) ^ 2)) '次元を戻す

            '------ファイルへの書き出し-----------
            writeFileNum = FreeFile()
            FileOpen(writeFileNum, writeFileName, OpenMode.Append)

            Dim TmpStr As String
            TmpStr = "  " & Format(Xxs, "0.000E+00") & _
               "  " & Format(xs, "0.000E+00") & _
               "  " & Format(fsub, "0.000E+00") & _
               "  " & Format(fourier_C, "0.000E+00") & _
               "  " & Format(fourier_S, "0.000E+00") & _
               "  " & Format(UU, "0.000E+00")
            PrintLine(writeFileNum, TmpStr)
            FileClose(writeFileNum)


        Next j

        MessageBox.Show("計算が終了しました！")
        Button1.Enabled = True
        Button1.Text = "1次元スキャン開始"
        Exit Sub

    End Sub

    Function Fx(ByVal t, ByVal xt, ByVal vt, ByVal xs, ByVal vs, ByVal fsub)

        'Fx = xt / dt = vt
        Fx = vt

    End Function

    Function Gx(ByVal t, ByVal xt, ByVal vt, ByVal xs, ByVal vs, ByVal fsub)

        'Gx = vt / dt = F
        Gx = -Kk * xt - Hn * (vt - vs) - u * Math.Sin(xt - xs) '通常の基板ポテンシャル(sin)はこの行をコメントイン


    End Function


    Function Fy(ByVal t, ByVal xt, ByVal vt, ByVal xs, ByVal vs, ByVal fsub)

        xs = xso * Math.Sin(t)
        Fy = xs

    End Function

    Function Gy(ByVal t, ByVal xt, ByVal vt, ByVal xs, ByVal vs, ByVal fsub)

        vs = xso * Math.Cos(t)
        Gy = vs

    End Function

    Function Fs(ByVal t, ByVal xt, ByVal vt, ByVal xs, ByVal vs, ByVal fsub)

        fsub = -Hn * (vt - vs) - u * Math.Sin(xt - xs)     '通常の基板ポテンシャル(sin)はこの行をコメントイン
        Fs = fsub                                          '通常の基板ポテンシャル(sin)はこの行をコメントイン

    End Function

End Class