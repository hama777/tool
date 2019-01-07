Imports System.IO
Imports System.Xml
Imports System.Text
Imports Microsoft.VisualBasic

Public Class Form1
    Const VERSION As String = "2.01"
    Const C_OUTFILE As String = "itune.htm"
    Const C_ERRLOG As String = "errlog.txt"
    Const C_ITUNELOG As String = "itune.log"
    Const C_BROWSER As String = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    'Const C_BROWSER As String = "D:\oracle\ols\sleipnir\Sleipnir.exe"
    Const EDITOR As String = "D:\ols\hide\Hidemaru.exe"
    Private Structure music_type
        Dim name, artist, player As String
        Dim t_id, size, time, pcount, rate As Integer
        Dim m_date, a_date, p_date As Date

    End Structure

    Dim total_cnt As Integer   ' 全曲数
    Dim rate_cnt As Integer    ' ★がある曲数
    Dim total_time As Integer  ' 全時間
    Dim rate_time As Integer   '★がある曲の時間
    Dim composer As New Dictionary(Of String, Integer)
    Dim com_time As New Dictionary(Of String, Integer)
    Dim com_rate As New Dictionary(Of String, Integer)
    Dim mdata As music_type
    Dim log_flg As Integer

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        log_flg = 0
        Call main_proc()
    End Sub

    Private Sub main_proc()
        Dim Item As New Dictionary(Of String, Integer)
        Dim ppp As New DictionaryEntry

        Call analize_xml()
        If log_flg = 0 Then
            Call display_list()
        Else
            Call output_log()
        End If
        Call display_sammary()

    End Sub

    Private Sub analize_xml()
        Dim key As String
        Dim s, ar As String
        Dim Item As New Dictionary(Of String, Integer)
        Dim ppp As New DictionaryEntry
        Dim r As Integer

        Dim errlog As StreamWriter = New StreamWriter(C_ERRLOG, False, Encoding.Default)
        '    Dim reader As XmlTextReader = New XmlTextReader(New StreamReader("test2.xml"))
        '        Dim reader As XmlTextReader = New XmlTextReader(New StreamReader("C:\Users\yoshi\Music\iTunes\iTunes Music Library.xml"))
        Dim reader As XmlTextReader = New XmlTextReader(New StreamReader("C:\Users\supik\Music\iTunes\iTunes Music Library.xml"))

        total_cnt = 0
        total_time = 0
        rate_cnt = 0
        rate_time = 0
        mdata.rate = 0
        Try
            reader.Read()
            While Not reader.EOF
                If reader.IsStartElement("key") Then
                    key = reader.ReadElementString("key")
                    Select Case key
                        Case "Track ID"
                            mdata.t_id = reader.ReadElementString("integer")
                        Case "Name"
                            mdata.name = reader.ReadElementString("string")

                        Case "Artist"
                            ar = reader.ReadElementString("string")
                            s = ar.Substring(ar.Length - 1)
                            If IsNumeric(s) Then
                                ar = ar.Substring(0, ar.Length - 1)
                            End If
                            mdata.artist = ar
                        Case "Album"
                            mdata.player = reader.ReadElementString("string")
                        Case "Size"
                            mdata.size = reader.ReadElementString("integer")
                        Case "Total Time"
                            mdata.time = reader.ReadElementString("integer") / 1000
                            total_time = total_time + mdata.time
                        Case "Play Count"
                            mdata.pcount = reader.ReadElementString("integer")
                        Case "Rating"
                            r = reader.ReadElementString("integer")
                            If r = 100 Then
                                mdata.rate = 1
                                'rate_cnt = rate_cnt + 1
                                'rate_time = rate_time + mdata.time
                            End If
                        Case "Rating Computed"
                            ' このkeyがある場合は自動的につけられたrateなので無視する
                            If mdata.rate = 1 Then
                                mdata.rate = 0
                            End If
                        Case "Location"
                            's = mdata.t_id & mdata.artist & mdata.name & mdata.player
                            '                            writer.WriteLine("*******************")
                            If mdata.rate = 1 Then
                                rate_cnt = rate_cnt + 1
                                rate_time = rate_time + mdata.time
                            End If
                            total_cnt = total_cnt + 1
                            If composer.ContainsKey(mdata.artist) Then
                                composer(mdata.artist) = composer(mdata.artist) + 1
                                com_time(mdata.artist) = com_time(mdata.artist) + mdata.time
                                If mdata.rate = 1 Then
                                    com_rate(mdata.artist) = com_rate(mdata.artist) + 1
                                End If

                            Else
                                composer(mdata.artist) = 1
                                com_time(mdata.artist) = mdata.time
                                com_rate(mdata.artist) = 0
                                If mdata.rate = 1 Then com_rate(mdata.artist) = 1
                            End If
                            mdata.rate = 0
                    End Select

                    Continue While
                End If
                reader.Read()
            End While

        Catch ex As XmlException
            errlog.WriteLine(ex.ToString())
        Finally
            reader.Close()
            errlog.Close()
        End Try
        total_time = total_time / 3600  ' 時間単位にする
        rate_time = rate_time / 3600

    End Sub
    Private Sub display_list()
        Dim sorted As List(Of KeyValuePair(Of String, Integer))
        Dim outfile As StreamWriter = New StreamWriter(C_OUTFILE, False, Encoding.Default)
        Dim i As Integer
        Dim ptime, hh, mm As Integer
        Dim s As String

        outfile.WriteLine("<html><head><title>一覧</title></head><body>")
        outfile.WriteLine("<center><b>一覧</b>&nbsp;&nbsp;&nbsp;</center><br>")
        outfile.WriteLine("<center>")
        outfile.WriteLine("<table border=0 bgcolor=#edaa28 cellspacing=0 cellpadding=0><tr><td>")
        outfile.WriteLine("<table cellspacing=1 border=0 cellpadding=3><tr bgcolor=#fbf703 >")
        outfile.WriteLine("<td>No.</td><td>作曲者</td><td>曲数</td><td>時間</td><td>時間/曲(分)</td><td>★</td><td>★%</td></tr>")

        sorted = sortByValue(composer)
        i = 0
        For Each kvp As KeyValuePair(Of String, Integer) In sorted
            i = i + 1

            If i Mod 2 = 0 Then
                outfile.WriteLine("<tr bgcolor=#f9f9e2>")
            Else
                outfile.WriteLine("<tr bgcolor=#ffffff>")
            End If

            ptime = com_time(kvp.Key) / 60  ' 秒単位 → 分単位
            hh = Int(ptime / 60)
            mm = ptime Mod 60
            s = "<td align=right>" & i & "</td><td>" & kvp.Key & "</td><td align=right>" & kvp.Value _
                & "</td><td align=right>" & Format(hh, "#0:") & Format(mm, "00") _
                & "</td><td align=right>" & Format(com_time(kvp.Key) / 60 / kvp.Value, " ###.#0 ") _
                & "</td><td align=right>" & com_rate(kvp.Key) & "</td>" _
                & "</td><td align=right>" & Format(com_rate(kvp.Key) / kvp.Value * 100, " ##0.0") & "</td></tr>"

            outfile.WriteLine(s)

        Next
        outfile.WriteLine("</table></td></tr></table></center></body></html>")

        outfile.Close()

        Shell(C_BROWSER & " """ & Application.StartupPath() & "\" & C_OUTFILE & """")

        '        sorted = sortByValue(com_time)

        'writer.WriteLine("*** 時間順 ***")
        'For Each kvp As KeyValuePair(Of String, Integer) In sorted
        '    s = kvp.Key & "  = " & Int(kvp.Value / 60)
        '    writer.WriteLine(s)
        'Next

    End Sub
    Private Sub display_sammary()
        Dim s As String

        lbl_music_cnt.Text = "曲数" & total_cnt.ToString.PadLeft(6) & "曲   " & rate_cnt.ToString.PadLeft(6) & Format(rate_cnt / total_cnt, "曲   (##0.0 %) ")
        lbl_time_cnt.Text = "時間" & total_time.ToString.PadLeft(6) & "時間 " & rate_time.ToString.PadLeft(6) & Format(rate_time / total_time, "時間 (##0.0 %) ")

    End Sub
    Private Sub output_log()
        Dim s As String
        Dim logfile As StreamWriter

        If log_flg = 1 Then
            logfile = New StreamWriter(C_ITUNELOG, True, Encoding.Default)
            s = Format(System.DateTime.Now, "yy/MM/dd HH:mm") & total_cnt.ToString.PadLeft(6) & _
                rate_cnt.ToString.PadLeft(6) & _
                "(" & Format(rate_cnt / total_cnt, "0.0%").PadLeft(6) & ")" & _
                total_time.ToString.PadLeft(6) & rate_time.ToString.PadLeft(6) & _
                "(" & Format(rate_time / total_time, "0.0%").PadLeft(6) & ")"
            logfile.WriteLine(s)
            logfile.Close()
        End If

    End Sub
    Shared Function hikaku( _
      ByVal kvp1 As KeyValuePair(Of String, Integer), _
      ByVal kvp2 As KeyValuePair(Of String, Integer)) As Integer

        Return kvp2.Value - kvp1.Value
    End Function

    Shared Function sortByValue( _
    ByVal item As Dictionary(Of String, Integer)) _
      As List(Of KeyValuePair(Of String, Integer))

        Dim list As New List(Of KeyValuePair(Of String, Integer))(item)

        list.Sort(AddressOf hikaku)
        Return list
    End Function

    Private Sub btm_log_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btm_log.Click
        log_flg = 1
        Call main_proc()
    End Sub

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.Text = "iTuneStat Ver. " & VERSION
        lbl_music_cnt.Text = ""
        lbl_time_cnt.Text = ""
    End Sub

    Private Sub cmd_logdisp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmd_logdisp.Click
        Process.Start(EDITOR, """" & Application.StartupPath() & "\" & C_ITUNELOG)

    End Sub
End Class
