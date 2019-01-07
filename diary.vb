Imports System.IO
Public Class frm_main

    Dim version As String = "Diary Ver 2.07"
    Dim curdate As Date        ' 現在、対象としている日付
    Dim cur_yy As Integer
    Dim cur_mm As Integer
    Dim cur_dd As Integer
    Dim cur_fname As String
    Dim edit_flg As Integer
    Dim olddata(1, 31) As String
    Dim color_data(31) As Integer   ' カラーデータ
    Dim we_data(31) As Integer      ' 天気データ
    Dim we_icon(10) As Object
    Dim flg_enter As Integer
    Dim seldate As Date
    Dim today_yy, today_mm, today_dd As Integer

    Private Structure holiday_type
    Dim yy, mm, dd, htype As Integer
    End Structure
    Dim holiday(300) As holiday_type
    Const max_holiday As Integer = 300  ' 休日データ最大値


    Public Sub New()
        Dim i As Integer

        ' この呼び出しは、Windows フォーム デザイナで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後で初期化を追加します。

        cb_color.Items.Add("黒")
        cb_color.Items.Add("赤")
        cb_color.Items.Add("緑")
        cb_color.Items.Add("青")
        '        cb_color.Text = "黒"

        cb_weather.Items.Add("なし")
        cb_weather.Items.Add("晴")
        cb_weather.Items.Add("晴/曇")
        cb_weather.Items.Add("曇")
        cb_weather.Items.Add("曇/雨")
        cb_weather.Items.Add("雨")
        cb_weather.Items.Add("大雨")
        cb_weather.Items.Add("雪")

        we_icon(0) = My.Resources.blank
        '        we_icon(1) = New Bitmap("fine.ico")
        we_icon(1) = My.Resources.fine1
        we_icon(2) = My.Resources.fine_cloudy
        we_icon(3) = My.Resources.cloudy
        we_icon(4) = My.Resources.rainy_cloudy
        we_icon(5) = My.Resources.rainy
        we_icon(6) = My.Resources.thunder
        we_icon(7) = My.Resources.snow


        For i = 1995 To 2020
            cb_yy.Items.Add(i)
        Next
        For i = 1 To 12
            cb_mm.Items.Add(i)
        Next

        Me.Text = version
        curdate = System.DateTime.Today
        Call date_var()
        today_yy = cur_yy     ' 本日の日付  常に不変とする
        today_mm = cur_mm
        today_dd = cur_dd

        Call read_holiday()
        Call grid_init()
        lbl_sel_date.Text = curdate.ToString("yyyy/MM/dd (ddd)")

        grid.Focus()

        cal_display()
    End Sub

    '  本体の表示
    '      cur_yy cyr_mm の年月の日記を表示する
    '      grid(x,y)   データ本体
    '           x   0 日付  1 .. 曜日  2 ... データ  3 ... 天気
    '           y   日付-1(0始まり)

    Public Sub display_body()
        Dim i As Integer
        Dim ww As Integer
        Dim you As Integer
        Dim tmp As Date
        Dim wkname() As String = {"日", "月", "火", "水", "木", "金", "土"}
        Dim ho As Integer
        Dim mday As Integer

        tmp = System.DateTime.Parse("#" & cur_yy & "/" & cur_mm & "/01#")
        ww = Weekday(tmp)

        Call grid_clear()

        Call data_load()
        Call old_data_load()

        cb_yy.Text = cur_yy
        cb_mm.Text = cur_mm
        grid.Visible = False
        For i = 1 To 31
            mday = month_of_day(cur_yy, cur_mm)
            If i <= mday Then
                grid(0, i - 1).Value = Microsoft.VisualBasic.Right("0" & cur_mm, 2) & "/" & Microsoft.VisualBasic.Right("0" & i, 2)
                you = (ww + i + 5) Mod 7
                grid(1, i - 1).Value = wkname(you)

                ho = check_holiday(cur_yy, cur_mm, i)

                If you = 6 Then
                    grid(0, i - 1).Style.BackColor = Color.LightBlue
                    grid(1, i - 1).Style.BackColor = Color.LightBlue
                    grid(0, i - 1).Style.ForeColor = Color.Blue
                    grid(1, i - 1).Style.ForeColor = Color.Blue

                End If
                If you = 0 Or ho = 1 Then   ' 日曜、または祝日
                    grid(0, i - 1).Style.BackColor = Color.Pink
                    grid(1, i - 1).Style.BackColor = Color.Pink
                    grid(0, i - 1).Style.ForeColor = Color.Red
                    grid(1, i - 1).Style.ForeColor = Color.Red

                End If

                If color_data(i) <> 0 Then
                    Select Case color_data(i)
                        Case 1
                            grid(2, i - 1).Style.ForeColor = Color.Red
                        Case 2
                            grid(2, i - 1).Style.ForeColor = Color.Green
                        Case 3
                            grid(2, i - 1).Style.ForeColor = Color.Blue
                    End Select
                Else
                    grid(2, i - 1).Style.ForeColor = Color.Black
                End If
                '                grid(3, i - 1).Value = New Bitmap("fine.ico")
                grid(3, i - 1).Value = we_icon(we_data(i))
            Else
                grid(0, i - 1).Value = ""
                grid(1, i - 1).Value = ""
                grid(2, i - 1).Value = ""
                grid(3, i - 1).Value = we_icon(0)
            End If
            '            grid.Rows(i - 1).Height = 17
            grid.Rows(i - 1).Height = 18
        Next
        grid.Visible = True

    End Sub

    '  検索 
    Private Sub cmd_search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmd_search.Click
        Dim s As String
        Dim i, mday, start As Integer
        Dim save_curdate As Date

        s = txt_search.Text
        If s = "" Then Exit Sub

        Call data_save()       ' 編集中のデータを保存

        If ck_direc.Checked = True Then
            Call search_foward(s)
        Else
            Call search_backward(s)
        End If

    End Sub
    Private Sub search_backward(ByVal s As String)
        Dim i, mday, start As Integer
        Dim save_curdate As Date

        save_curdate = curdate
        start = grid.CurrentCell.RowIndex - 1
        Do
            If start = -1 Then          ' 月の最初まで検索した
                Call prev_month()
                mday = month_of_day(cur_yy, cur_mm)   ' 月末の日
                start = mday - 1
            End If
            For i = start To 0 Step -1
                If InStr(grid(2, i).Value, s) <> 0 Then
                    grid(2, i).Selected = True
                    Call display_body()
                    Exit Sub
                End If
            Next
            Call prev_month()
            If cur_yy = 1994 Then       ' 検索終了  とりあえず1994年
                curdate = save_curdate      ' 元の画面に戻す
                Call date_var()
                Call display_body()
                Exit Sub
            End If
            mday = month_of_day(cur_yy, cur_mm)
            start = mday - 1
        Loop

    End Sub

    Private Sub search_foward(ByVal s As String)
        Dim i, mday, start As Integer
        Dim save_curdate As Date

        save_curdate = curdate
        start = grid.CurrentCell.RowIndex + 1
        mday = month_of_day(cur_yy, cur_mm)
        Do
            If start >= mday Then
                Call next_month()
                mday = month_of_day(cur_yy, cur_mm)   ' 月末の日
                start = 0
            End If
            For i = start To mday - 1
                If InStr(grid(2, i).Value, s) <> 0 Then
                    grid(2, i).Selected = True
                    Call display_body()
                    Exit Sub
                End If
            Next
            Call next_month()
            If cur_yy * 100 + cur_mm > today_yy * 100 + today_mm Then ' 検索終了
                curdate = save_curdate      ' 元の画面に戻す
                Call date_var()
                Call display_body()
                Exit Sub
            End If
            mday = month_of_day(cur_yy, cur_mm)
            start = 0
        Loop

    End Sub

    '   検索用  日記データ読み込み
    Private Sub prev_month()
        curdate = DateAdd(DateInterval.Month, -1, curdate)
        Call date_var()
        Call data_load()
    End Sub
    Private Sub next_month()
        curdate = DateAdd(DateInterval.Month, 1, curdate)
        Call date_var()
        Call data_load()
    End Sub

    '   cur_yy 年 cur_mm 月の日数を求める
    Private Function month_of_day(ByVal cur_yy As Integer, ByVal cur_mm As Integer)
        Dim mmday() As Integer = {0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31}

        month_of_day = mmday(cur_mm)
        If cur_mm <> 2 Then
            Exit Function
        End If
        If cur_yy Mod 4 = 0 Then
            month_of_day += 1
        End If
    End Function


    Private Sub grid_clear()
        Dim i As Integer
        For i = 0 To 30
            grid(0, i).Style.ForeColor = Color.Black
            grid(0, i).Style.BackColor = Color.White
            grid(1, i).Style.ForeColor = Color.Black
            grid(1, i).Style.BackColor = Color.White
            grid(2, i).Value = ""
            grid(3, i).Value = Nothing
        Next

        For i = 0 To 31
            color_data(i) = 0
            we_data(i) = 0
        Next

    End Sub

    ' グリッドの初期化
    Private Sub grid_init()

        grid.RowCount = 31
        grid.ColumnCount = 4
        grid.Columns(0).Width = 50
        grid.Columns(1).Width = 30
        grid.Columns(2).Width = 390
        grid.Columns(3).Width = 40

        'grid.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        grid.RowHeadersVisible = False

        grid.Columns(0).ReadOnly = True   ' 日付、曜日は編集不可
        grid.Columns(1).ReadOnly = True


        grid.AllowUserToResizeColumns = False  ' ユーザがセルの高さ、幅の編集不可
        grid.AllowUserToResizeRows = False

        grid(2, cur_dd - 1).Selected = True
        '        grid.BeginEdit(True)

        '        grid.Font = New Font("ＭＳ 明朝", 10)
        '        grid.Font = New Font("MS UI Gothic", 10)
        Call display_body()

        Call set_olddata(cur_dd)

    End Sub


    ' グリッドのセルがクリックされた
    Private Sub grid_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles grid.CellClick
        Dim dd As Integer

        dd = grid.CurrentCell.RowIndex + 1
        curdate = System.DateTime.Parse("#" & cur_yy & "/" & cur_mm & "/" & dd & "#")

        lbl_sel_date.Text = curdate
        Call set_olddata(dd)

    End Sub

    ' グリッドのセルにフォーカス
    Private Sub grid_CellEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles grid.CellEnter

        Dim dd As Integer

        dd = grid.CurrentCell.RowIndex + 1

        curdate = System.DateTime.Parse("#" & cur_yy & "/" & cur_mm & "/" & dd & "#")
        lbl_sel_date.Text = curdate
        Call set_olddata(dd)

    End Sub

    Private Sub grid_CellParsing(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellParsingEventArgs) Handles grid.CellParsing

        edit_flg = 1

    End Sub
    '  データの保存
    Private Sub data_save()
        Dim writer As StreamWriter
        Dim i As Integer
        Dim s As String

        grid.EndEdit()  ' 現在編集中のセルをコミットして終了

        If edit_flg = 0 Then Exit Sub ' 編集がなければ保存しない

        writer = New StreamWriter(cur_fname, False, System.Text.Encoding.Default)
        For i = 1 To 31
            s = grid(2, i - 1).Value
            writer.WriteLine(i & vbTab & s & vbTab & color_data(i) & vbTab & we_data(i))
        Next
        writer.Close()

    End Sub
    ' データの読み込み
    '       cur_fname のデータファイルを読み込み grid にセットする
    Private Sub data_load()
        Dim reader As StreamReader
        Dim i As Integer
        Dim s As String
        Dim dt() As String

        edit_flg = 0       ' 修正フラグクリア
        Try
            reader = New StreamReader(cur_fname, System.Text.Encoding.Default)
        Catch E As Exception
            Exit Sub
        End Try

        i = 0
        Do
            s = reader.ReadLine()
            If s = Nothing Then Exit Do
            dt = Split(s, vbTab)
            grid(2, i).Value = dt(1)
            i = i + 1
            color_data(i) = dt(2)
            we_data(i) = dt(3)
        Loop
        reader.Close()
    End Sub

    ' 1ヶ月前、1年前のデータロード
    Private Sub old_data_load()
        Dim olddate As Date
        Dim yy, mm As Integer

        olddate = DateAdd(DateInterval.Month, -1, curdate) ' 1ヶ月前
        yy = Year(olddate)
        mm = Month(olddate)
        Call data_load_ym(yy, mm, 0)

        olddate = DateAdd(DateInterval.Year, -1, curdate) ' 1年前
        yy = Year(olddate)
        mm = Month(olddate)
        Call data_load_ym(yy, mm, 1)

    End Sub

    ' 指定した年月のデータを olddata にロード
    Private Sub data_load_ym(ByVal yy As Integer, ByVal mm As Integer, ByVal type As Integer)
        ' type = 0  1ヶ月前   1 ... 1年前
        Dim reader As StreamReader
        Dim i As Integer
        Dim s As String
        Dim dt() As String
        Dim fname As String

        fname = "NIK" & Format(yy, "00") & Format(mm, "00") & ".txt"

        Try
            reader = New StreamReader(fname, System.Text.Encoding.Default)
        Catch E As Exception
            ' ファイルがない場合は中身をクリア
            For i = 0 To 31
                olddata(type, i) = ""
            Next
            Exit Sub
        End Try

        i = 0
        Do
            s = reader.ReadLine()
            If s = Nothing Then Exit Do
            dt = Split(s, vbTab)
            olddata(type, i) = dt(1)
            i = i + 1
        Loop
        reader.Close()
    End Sub

    Private Sub Form1_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        Call data_save()
    End Sub

    '   1ヶ月前
    Private Sub cmd_prev_m1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmd_prev_m1.Click
        seldate = DateAdd(DateInterval.Month, -1, curdate)
        Call move_date()
    End Sub

    '   1年前
    Private Sub cmd_prev_year_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmd_prev_year.Click
        seldate = DateAdd(DateInterval.Year, -1, curdate)
        Call move_date()
    End Sub

    '   本日
    Private Sub cmd_today_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmd_today.Click
        seldate = System.DateTime.Today
        Call move_date()
    End Sub

    '   1ヶ月後
    Private Sub cmd_next_m1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmd_next_m1.Click
        seldate = DateAdd(DateInterval.Month, 1, curdate)
        Call move_date()
    End Sub

    '   1年後
    Private Sub cmd_next_year_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmd_next_year.Click
        seldate = DateAdd(DateInterval.Year, 1, curdate)
        Call move_date()
    End Sub
    '   指定した年月へ
    Private Sub cmd_jump_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmd_jump.Click
        seldate = System.DateTime.Parse(cb_yy.Text & "年" & cb_mm.Text & "月1日")
        Call move_date()
    End Sub


    '   seldate の年月に移動しデータを表示する
    Private Sub move_date()

        '        Me.SuspendLayout()

        'Panel1.Visible = False
        'Panel2.Visible = False
        'Panel3.Visible = False
        If Month(seldate) <> Month(curdate) Or Year(seldate) <> Year(curdate) Then
            Call data_save()
            curdate = seldate
            Call date_var()
            Call display_body()
            Call cal_display()
        End If
        curdate = seldate
        Call date_var()
        lbl_sel_date.Text = curdate.ToString("yyyy/MM/dd (ddd)")
        Call set_olddata(cur_dd)

        '        lbl_old_mm.Text = olddata(0, cur_dd - 1)
        '        lbl_old_yy.Text = olddata(1, cur_dd - 1)
        grid(2, cur_dd - 1).Selected = True

        '       Me.ResumeLayout()

        'Panel1.Visible = True
        'Panel2.Visible = True
        'Panel3.Visible = True

    End Sub

    '  1ヶ月前、1年前のデータをラベルにセットする
    Private Sub set_olddata(ByVal set_dd As Integer)
        Dim sd, sy As String
        Dim dd As Integer

        dd = set_dd
        If set_dd = 1 Then dd = 2
        sd = (dd - 1) & "日 " & olddata(0, dd - 2) & vbCrLf & dd & "日 " & olddata(0, dd - 1) & vbCrLf & (dd + 1) & "日 " & olddata(0, dd)
        sy = (dd - 1) & "日 " & olddata(1, dd - 2) & vbCrLf & dd & "日 " & olddata(1, dd - 1) & vbCrLf & (dd + 1) & "日 " & olddata(1, dd)
        lbl_old_mm.Text = sd
        lbl_old_yy.Text = sy
    End Sub

    ' 休日データの読み込み
    Private Sub read_holiday()
        Dim reader As StreamReader
        Dim i As Integer
        Dim s, fname, hdate As String
        Dim dt(), ymd() As String

        fname = "holiday.txt"
        Try
            reader = New StreamReader(fname, System.Text.Encoding.Default)
        Catch E As Exception
            ' ファイルがない場合
            Exit Sub
        End Try

        i = 0
        Do
            If i = max_holiday Then
                MsgBox("holiday data overflow")
                Exit Do
            End If
            s = reader.ReadLine()
            If s = Nothing Then Exit Do
            dt = Split(s, ";")
            hdate = dt(0)
            ymd = Split(hdate, "/")
            If ymd(0) = "*" Then
                holiday(i).yy = -1
            Else
                holiday(i).yy = ymd(0)
            End If
            holiday(i).mm = ymd(1)
            holiday(i).dd = ymd(2)
            holiday(i).htype = dt(1)
            i = i + 1
        Loop
        reader.Close()

    End Sub
    ' 休日チェック
    Private Function check_holiday(ByVal yy As Integer, ByVal mm As Integer, ByVal dd As Integer)
        '  0  ...  平日     1 ...  休日
        Dim i As Integer

        For i = 0 To holiday.Length - 1
            If holiday(i).dd <> dd Then
                Continue For
            End If
            If holiday(i).mm <> mm Then
                Continue For
            End If
            If holiday(i).yy = -1 Then  ' 年は無視
                Return 1
            Else
                If holiday(i).yy + 2000 <> yy Then
                    Continue For
                Else
                    Return 1
                End If

            End If
        Next
        Return 0
    End Function

    ' curdate を年・月・日に分割。ファイル名を取得
    Public Sub date_var()
        cur_yy = Year(curdate)
        cur_mm = Month(curdate)
        cur_dd = Microsoft.VisualBasic.DateAndTime.Day(curdate)
        cur_fname = "NIK" & Format(cur_yy, "00") & Format(cur_mm, "00") & ".txt"
    End Sub

    Private Sub cmd_fin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmd_fin.Click
        Me.Close()
        End
    End Sub

    ' 色のセット
    Private Sub cmd_color_set_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim r As Integer
        Dim c As Integer

        r = grid.CurrentCell.RowIndex
        c = cb_color.SelectedIndex

        If grid(2, r).Value = "" Then Exit Sub
        color_data(r + 1) = c

        Select Case c
            Case 0
                grid(2, r).Style.ForeColor = Color.Black
            Case 1
                grid(2, r).Style.ForeColor = Color.Red
            Case 2
                grid(2, r).Style.ForeColor = Color.Green
            Case 3
                grid(2, r).Style.ForeColor = Color.Blue
        End Select
        edit_flg = 1
    End Sub

    ' 天気の設定
    Private Sub cb_weather_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_weather.SelectedIndexChanged

        Dim r As Integer
        Dim c As Integer

        r = grid.CurrentCell.RowIndex
        c = cb_weather.SelectedIndex
        If c = -1 Then Exit Sub
        we_data(r + 1) = c
        edit_flg = 1
        grid(3, r).Value = we_icon(c)

    End Sub

    ' 色の設定
    Private Sub cb_color_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_color.SelectedIndexChanged
        Dim r As Integer
        Dim c As Integer

        r = grid.CurrentCell.RowIndex
        c = cb_color.SelectedIndex

        If grid(2, r).Value = "" Then Exit Sub
        color_data(r + 1) = c

        Select Case c
            Case 0
                grid(2, r).Style.ForeColor = Color.Black
            Case 1
                grid(2, r).Style.ForeColor = Color.Red
            Case 2
                grid(2, r).Style.ForeColor = Color.Green
            Case 3
                grid(2, r).Style.ForeColor = Color.Blue
        End Select
        edit_flg = 1


    End Sub

    ' ************************************  
    ' ***          カレンダー部        ***
    ' ************************************

    Private Sub cal_display()
        Dim yy, mm As Integer

        mm = cur_mm - 1
        yy = cur_yy
        If mm = 0 Then
            mm = 12
            yy = cur_yy - 1
        End If
        Call display(Panel1, yy, mm)
        Call display(Panel2, cur_yy, cur_mm)
        mm = cur_mm + 1
        yy = cur_yy
        If mm = 13 Then
            mm = 1
            yy = cur_yy + 1
        End If
        Call display(Panel3, yy, mm)
    End Sub

    Private Sub display(ByVal pp As Panel, ByVal currentYear As Integer, ByVal currentMonth As Integer)
        Dim lblDays(42) As Label
        Dim fontCell As New Font("ＭＳ 明朝", 9)
        Dim colorCell As Color = Color.Black
        '        Dim colorWeekday As Color = Color.White
        '        Dim colorSunday As Color = Color.Pink
        '       Dim colorSaturday As Color = Color.LightBlue
        Dim color_out_of_cell As Color = Color.FromArgb(220, 220, 220)

        Dim weekday() As String = {"日", "月", "火", "水", "木", "金", "土"}
        Dim widthCell As Integer = 18
        Dim heightCell As Integer = 16
        Dim widthYearMonth As Integer = 100
        Dim x0 As Integer = 5
        Dim y0 As Integer = 0

        Dim x, y As Integer
        Dim firstDay As New DateTime(currentYear, currentMonth, 1)
        Dim firstDayOfWeek As Integer = firstDay.DayOfWeek
        Dim dayMax As Integer = DateTime.DaysInMonth(currentYear, currentMonth)

        Dim dToday As DateTime = DateTime.Today
        Dim colorToday As Color = Color.Green
        Dim colorToday_back As Color = Color.LightGreen

        Dim d As Integer = 0
        Dim p_mm, p_yy, p_dd, last_month_start As Integer
        Dim flg_visible As Integer
        Dim next_mon_dd As Integer
        Dim lblDayOfWeek(7) As Label      ' 曜日のタイトル
        Dim lbltitle As Label

        pp.Controls.Clear()
        pp.Visible = False
        lbltitle = New Label
        With lbltitle
            .Text = currentYear & "年 " & currentMonth & "月"
            .Font = New Font("ＭＳ 明朝", 9, FontStyle.Bold)
            .ForeColor = Color.Blue
            .Height = heightCell
            .Width = widthCell * 7
            .TextAlign = ContentAlignment.MiddleCenter
            .Location = New Point(x0, y0)
        End With
        pp.Controls.Add(lbltitle)

        For i As Integer = 0 To 6
            lblDayOfWeek(i) = New Label
            With lblDayOfWeek(i)
                .Text = weekday(i)
                .Font = fontCell
                .ForeColor = colorCell
                .Width = widthCell
                .Height = heightCell
                .TextAlign = ContentAlignment.MiddleCenter
                .Location = New Point(x0 + widthCell * i, y0 + heightCell)
            End With
            pp.Controls.Add(lblDayOfWeek(i))
        Next

        ' カレンダー部

        flg_visible = 0
        next_mon_dd = 1
        If firstDayOfWeek <> 0 Then   ' 月外の最初の日を求める
            p_mm = currentMonth - 1
            p_yy = currentYear
            If p_mm = 0 Then
                p_mm = 12
                p_yy = p_yy - 1
            End If
            p_dd = DateTime.DaysInMonth(p_yy, p_mm)   ' 先月の最終日
            last_month_start = p_dd - firstDayOfWeek + 1
        End If

        For i As Integer = 0 To 41
            lblDays(i) = New Label
            With lblDays(i)
                .Width = widthCell
                .Height = heightCell
                x = widthCell * (i Mod 7) + x0
                y = heightCell * (i \ 7) + y0 + heightCell * 2
                .Location = New Point(x, y)
                .Tag = i.ToString
                .Cursor = Cursors.Hand
                .TextAlign = ContentAlignment.MiddleCenter

                If i < firstDayOfWeek Then    ' 先月の処理
                    .Text = last_month_start.ToString
                    .Font = fontCell
                    .ForeColor = colorCell
                    last_month_start = last_month_start + 1
                    .BackColor = color_out_of_cell
                    .Visible = True
                End If
                If i >= firstDayOfWeek And i < firstDayOfWeek + dayMax Then    ' 当月の処理
                    d = i - firstDayOfWeek + 1
                    .Text = d.ToString
                    .Font = fontCell

                    Call day_color(lblDays(i), currentYear, currentMonth, d, i)

                    .Visible = True
                End If
                If i >= firstDayOfWeek + dayMax Then        ' 次月の処理

                    If i Mod 7 = DayOfWeek.Sunday Then
                        flg_visible = 1
                    End If
                    If flg_visible = 0 Then
                        .Text = next_mon_dd.ToString
                        .Font = fontCell
                        .ForeColor = colorCell
                        .BackColor = color_out_of_cell
                        .Visible = True
                        next_mon_dd = next_mon_dd + 1
                    Else
                        .Visible = False
                    End If

                End If

            End With
            pp.Controls.Add(lblDays(i))
            '            AddHandler lblDays(i).Click, AddressOf lblDays_Click
        Next
        pp.Visible = True


        '        Call UpdateCalendar()

    End Sub

    ' 日の 文字色、背景色を設定する
    Public Sub day_color(ByVal lab As Label, ByVal yy As Integer, ByVal mm As Integer, ByVal dd As Integer, ByVal i As Integer)
        Dim dToday As DateTime = DateTime.Today

        Dim colorCell As Color = Color.Black
        Dim colorWeekday As Color = Color.White
        Dim colorSunday As Color = Color.Pink
        Dim colorSaturday As Color = Color.LightBlue
        Dim color_out_of_cell As Color = Color.LightGray
        Dim colorToday As Color = Color.Green
        Dim colorToday_back As Color = Color.LightGreen

        With lab
            .ForeColor = colorCell
            .BackColor = colorWeekday

            If i Mod 7 = DayOfWeek.Sunday Then
                .ForeColor = Color.Red
                .BackColor = Color.FromArgb(254, 208, 224)
            End If
            If i Mod 7 = DayOfWeek.Saturday Then
                .ForeColor = Color.Blue
                .BackColor = Color.FromArgb(221, 232, 255)
            End If

            If check_holiday(yy, mm, dd) = 1 Then
                .ForeColor = Color.Red
                .BackColor = Color.FromArgb(254, 208, 224)
            End If
            If dd = dToday.Day And yy = dToday.Year And mm = dToday.Month Then
                '                .Font = New Font("ＭＳ 明朝", 9, FontStyle.Bold)
                .ForeColor = colorToday
                .BackColor = Color.FromArgb(226, 255, 211)
            End If

        End With

    End Sub

    Private Sub cmd_grep_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmd_grep.Click
        Dim f As Form2 = New Form2(txt_search.Text)
        f.Show()

    End Sub
End Class
