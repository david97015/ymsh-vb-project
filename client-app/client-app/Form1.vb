Imports System.IO
Imports System.IO.Compression
Imports System.Net.Dns

Public Class Form1


    Public Function GetIPaddress() As String '取得IP位置
        Try
            Dim ipEntry As System.Net.IPHostEntry = GetHostByName(Environment.MachineName)
            Dim IpAddr As System.Net.IPAddress() = ipEntry.AddressList
            GetIPaddress = IpAddr(0).ToString
        Catch ex As Exception
            GetIPaddress = "無法取得IP"
        End Try
    End Function

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ToolStripStatusLabel1.Text = ("使用者名稱:" & System.Net.Dns.GetHostName().ToString()) '取得使用者名稱
        ToolStripStatusLabel2.Text = ("IP位置:" & GetIPaddress()) '取得IP位置

        Timer2.Enabled = True
        Timer2.Interval = 150
        Timer2.Start()

        Dim 檔案寫入時間 As String = DateTime.Now.ToString("yyyy-MM-dd") '設定時間變數(日期YYYY-MM-DD)
        Dim 目前小時 As String = DateTime.Now.ToString("HH") '設定時間變數(小時HH)-24小時制
        If 目前小時 >= 10 Then '檢查時間是否正確
            MsgBox("僅開放早上10點前訂餐" & vbCrLf% & "謝謝配合", 64, "訂餐時間錯誤!") '提示時間錯誤
            Me.Close() '關閉訂餐程式
        End If

        Dim TXT檔案寫入位置 As String = ("C:\VB-app\" & 檔案寫入時間 & "\") '設定解壓縮位置
        Dim 資料夾位置 As String = ("C:\VB-app\" & 檔案寫入時間) '檢測ZIP檔是否已存在
        If (System.IO.Directory.Exists(資料夾位置)) Then
            System.IO.Directory.Delete(資料夾位置, True) '刪除舊ZIP檔
        Else
        End If

        ZipFile.ExtractToDirectory("C:\VB-app\vb.zip", TXT檔案寫入位置) '解壓縮

        '時間日期設定
        Dim 本周日期範圍 = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "DATE1.txt") '本周日期範圍
        Dim 國字星期 As String '星期幾(國字表示) 
        Dim 數字星期 = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "DATE3.txt") '星期幾(數字表示)

        '將以星期日為第一天修正為星期一
        If 數字星期 = 1 Then
            國字星期 = "日"
            數字星期 = 7
            GoTo A
        ElseIf 數字星期 = 2 Then
            國字星期 = "一"
            數字星期 = 1
            GoTo A
        ElseIf 數字星期 = 3 Then
            國字星期 = "二"
            數字星期 = 2
            GoTo A
        ElseIf 數字星期 = 4 Then
            國字星期 = "三"
            數字星期 = 3
            GoTo A
        ElseIf 數字星期 = 5 Then
            國字星期 = "四"
            數字星期 = 4
            GoTo A
        ElseIf 數字星期 = 6 Then
            國字星期 = "五"
            數字星期 = 5
            GoTo A
        ElseIf 數字星期 = 7 Then
            國字星期 = "六"
            數字星期 = 6
            GoTo A
        End If

A:
        '檢查是否為一~五可訂餐時間
        If 數字星期 = 6 Then '星期六
            MsgBox("僅開放星期一~五訂餐" & vbCrLf% & "謝謝配合", 64, "訂餐日期錯誤!") '提示日期錯誤
            Me.Close() '關閉訂餐程式
        End If

        If 數字星期 = 7 Then '星期日
            MsgBox("僅開放星期一~五訂餐" & vbCrLf% & "謝謝配合", 64, "訂餐日期錯誤!") '提示日期錯誤
            Me.Close() '關閉訂餐程式
        End If

        '前置標題設定
        Label1.Text = (本周日期範圍 & " (午餐) 品項訂購目錄") '主題列
        GroupBox1.Text = ("午餐主菜(星期" & 國字星期 & ")") '午餐類別名稱
        GroupBox3.Text = ("晚餐副食(星期" & 國字星期 & ")") '晚餐類別名稱
        Label2.Text = ("星期" & 國字星期 & "便當") '午餐介紹
        Label9.Text = ("星期" & 國字星期 & "副食") '晚餐介紹

        Label13.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F301.txt") '綜合羹麵
        Label15.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F302.txt") '綜合羹飯
        Label17.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F303.txt") '綜合羹餃
        Label19.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F300.txt") '羹湯
        Label21.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F401.txt") '酸辣麵
        Label23.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F402.txt") '酸辣飯
        Label25.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F403.txt") '酸辣餃
        Label27.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F400.txt") '酸辣湯
        Label29.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F500.txt") '水餃
        Label31.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F510.txt") '滷肉飯
        Label33.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F521.txt") '湯麵
        Label35.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F522.txt") '乾麵
        Label37.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F530.txt") '湯飯
        Label39.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F600.txt") '貢丸湯
        Label41.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F700.txt") '餡餅
        Label43.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F801.txt") '每週30元的
        Label45.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F802.txt") '每週20元的
        '全日供應價格設定                                   
        Label14.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF301.txt") '綜合羹麵金額
        Label16.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF302.txt") '綜合羹飯金額
        Label18.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF303.txt") '綜合羹餃金額
        Label20.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF300.txt") '羹湯金額
        Label22.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF401.txt") '酸辣麵金額
        Label24.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF402.txt") '酸辣飯金額
        Label26.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF403.txt") '酸辣餃金額
        Label28.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF400.txt") '酸辣湯金額
        Label30.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF500.txt") '水餃金額
        Label32.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF510.txt") '滷肉飯金額
        Label34.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF521.txt") '湯麵金額
        Label36.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF522.txt") '乾麵金額
        Label38.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF530.txt") '湯飯金額
        Label40.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF600.txt") '貢丸湯金額
        Label42.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF700.txt") '餡餅金額
        Label44.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF801.txt") '每週30元的金額
        Label46.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF802.txt") '每週20元的金額

        '每日不同餐點設定
        If 數字星期 = 1 Then '星期一
            Label3.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F111.txt") '星期一便當-1 '午餐便當種類-1
            Label4.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F112.txt") '星期一便當-2 '午餐便當種類-2
            Label10.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F210.txt") '星期一晚上副食
            Label5.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF111.txt") '星期一便當-1金額
            Label6.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF112.txt") '星期一便當-2金額
            Label11.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF210.txt") '星期一晚上副食金額
        End If
        If 數字星期 = 2 Then '星期二
            Label3.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F121.txt") '星期二便當-1 '午餐便當種類-1
            Label4.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F122.txt") '星期二便當-2 '午餐便當種類-2
            Label10.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F220.txt") '星期二晚上副食
            Label5.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF121.txt") '星期二便當-1金額
            Label6.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF122.txt") '星期二便當-2金額
            Label11.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF220.txt") '星期二晚上副食金額
        End If
        If 數字星期 = 3 Then '星期三
            Label3.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F131.txt") '星期三便當-1 '午餐便當種類-1
            Label4.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F132.txt") '星期三便當-2 '午餐便當種類-2
            Label10.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F230.txt") '星期三晚上副食
            Label5.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF131.txt") '星期三便當-1金額
            Label6.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF132.txt") '星期三便當-2金額
            Label11.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF230.txt") '星期三晚上副食金額
        End If
        If 數字星期 = 4 Then '星期四
            Label3.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F141.txt") '星期四便當-1 '午餐便當種類-1
            Label4.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F142.txt") '星期四便當-2 '午餐便當種類-2
            Label10.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F240.txt") '星期四晚上副食
            Label5.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF141.txt") '星期四便當-1金額
            Label6.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF142.txt") '星期四便當-2金額
            Label11.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF240.txt") '星期四晚上副食金額
        End If
        If 數字星期 = 5 Then '星期五
            Label3.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F151.txt") '星期五便當-1  '午餐便當種類-1
            Label4.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F152.txt") '星期五便當-2  '午餐便當種類-2
            Label10.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F250.txt") '星期五晚上副食
            Label5.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF151.txt") '星期五便當-1金額
            Label6.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF152.txt") '星期五便當-2金額
            Label11.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF250.txt") '星期五晚上副食金額
        End If
        If 數字星期 = 6 Then '星期六
            Label3.Text = "今日無餐"
            Label4.Text = "今日無餐"
            Label10.Text = "今日無餐"
            Label5.Text = "N/A"
            Label6.Text = "N/A"
            Label11.Text = "N/A"
        End If
        If 數字星期 = 7 Then '星期日
            Label3.Text = "今日無餐"
            Label4.Text = "今日無餐"
            Label10.Text = "今日無餐"
            Label5.Text = "N/A"
            Label6.Text = "N/A"
            Label11.Text = "N/A"
        End If
        ToolStripStatusLabel3.ForeColor = Color.Red
        ToolStripStatusLabel3.Text = ("注意事項:" & My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "INFO.txt"))

    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Dim Now_time As String = DateTime.Now.ToString("yyyy-MM-dd") '設定時間變數(日期YYYY-MM-DD)
        Label107.Text = Now.ToString("hh:mm:ss")
        Dim txt_location As String = ("C:\VB-app\" & Now_time & "\") '設定解壓縮位置
        '一~五便當價格設定
        Dim [變數MF111] = My.Computer.FileSystem.ReadAllText(txt_location & "MF111.txt")
        Dim [變數MF112] = My.Computer.FileSystem.ReadAllText(txt_location & "MF112.txt")
        Dim [變數MF121] = My.Computer.FileSystem.ReadAllText(txt_location & "MF121.txt")
        Dim [變數MF122] = My.Computer.FileSystem.ReadAllText(txt_location & "MF122.txt")
        Dim [變數MF131] = My.Computer.FileSystem.ReadAllText(txt_location & "MF131.txt")
        Dim [變數MF132] = My.Computer.FileSystem.ReadAllText(txt_location & "MF132.txt")
        Dim [變數MF141] = My.Computer.FileSystem.ReadAllText(txt_location & "MF141.txt")
        Dim [變數MF142] = My.Computer.FileSystem.ReadAllText(txt_location & "MF142.txt")
        Dim [變數MF151] = My.Computer.FileSystem.ReadAllText(txt_location & "MF151.txt")
        Dim [變數MF152] = My.Computer.FileSystem.ReadAllText(txt_location & "MF152.txt")
        '一~五晚餐副食價格設定
        Dim [變數MF210] = My.Computer.FileSystem.ReadAllText(txt_location & "MF210.txt")
        Dim [變數MF220] = My.Computer.FileSystem.ReadAllText(txt_location & "MF220.txt")
        Dim [變數MF230] = My.Computer.FileSystem.ReadAllText(txt_location & "MF230.txt")
        Dim [變數MF240] = My.Computer.FileSystem.ReadAllText(txt_location & "MF240.txt")
        Dim [變數MF250] = My.Computer.FileSystem.ReadAllText(txt_location & "MF250.txt")
        '全日供應價格設定
        Dim [變數MF301] = My.Computer.FileSystem.ReadAllText(txt_location & "MF301.txt")
        Dim [變數MF302] = My.Computer.FileSystem.ReadAllText(txt_location & "MF302.txt")
        Dim [變數MF303] = My.Computer.FileSystem.ReadAllText(txt_location & "MF303.txt")
        Dim [變數MF300] = My.Computer.FileSystem.ReadAllText(txt_location & "MF300.txt")
        Dim [變數MF401] = My.Computer.FileSystem.ReadAllText(txt_location & "MF401.txt")
        Dim [變數MF402] = My.Computer.FileSystem.ReadAllText(txt_location & "MF402.txt")
        Dim [變數MF403] = My.Computer.FileSystem.ReadAllText(txt_location & "MF403.txt")
        Dim [變數MF400] = My.Computer.FileSystem.ReadAllText(txt_location & "MF400.txt")
        Dim [變數MF500] = My.Computer.FileSystem.ReadAllText(txt_location & "MF500.txt")
        Dim [變數MF510] = My.Computer.FileSystem.ReadAllText(txt_location & "MF510.txt")
        Dim [變數MF521] = My.Computer.FileSystem.ReadAllText(txt_location & "MF521.txt")
        Dim [變數MF522] = My.Computer.FileSystem.ReadAllText(txt_location & "MF522.txt")
        Dim [變數MF530] = My.Computer.FileSystem.ReadAllText(txt_location & "MF530.txt")
        Dim [變數MF600] = My.Computer.FileSystem.ReadAllText(txt_location & "MF600.txt")
        Dim [變數MF700] = My.Computer.FileSystem.ReadAllText(txt_location & "MF700.txt")
        Dim [變數MF801] = My.Computer.FileSystem.ReadAllText(txt_location & "MF801.txt")
        Dim [變數MF802] = My.Computer.FileSystem.ReadAllText(txt_location & "MF802.txt")

        '視窗標題設定
        Me.Text = ("新竹縣私立義民高級中學點餐系統-client端 " + "現在時間:" + Now)

        Dim time1 As DateTime = #10:00:00# '剩餘時間計算
        Dim time2 As DateTime = DateTime.Now.ToString("hh:mm:ss")
        Dim ts As System.TimeSpan = time1.Subtract(time2)
        Dim strtimespan As String = ts.Hours.ToString() + ":" + ts.Minutes.ToString() + ":" + ts.Seconds.ToString()
        Label105.Text = strtimespan

        '小計變數設定
        Dim [變數MF111小計] = ([變數MF111] * NumericUpDown1.Value) '午餐
        Dim [變數MF112小計] = ([變數MF112] * NumericUpDown2.Value) '晚餐
        Dim [變數MF210小計] = ([變數MF210] * NumericUpDown3.Value) '晚餐
        Dim [變數MF301午小計] = (NumericUpDown4.Value) * [變數MF301] '午餐
        Dim [變數MF302午小計] = (NumericUpDown5.Value) * [變數MF302] '午餐
        Dim [變數MF303午小計] = (NumericUpDown6.Value) * [變數MF303] '午餐
        Dim [變數MF300午小計] = (NumericUpDown7.Value) * [變數MF300] '午餐
        Dim [變數MF401午小計] = (NumericUpDown8.Value) * [變數MF401] '午餐
        Dim [變數MF402午小計] = (NumericUpDown9.Value) * [變數MF402] '午餐
        Dim [變數MF403午小計] = (NumericUpDown10.Value) * [變數MF403] '午餐
        Dim [變數MF400午小計] = (NumericUpDown11.Value) * [變數MF400] '午餐
        Dim [變數MF500午小計] = (NumericUpDown12.Value) * [變數MF500] '午餐
        Dim [變數MF510午小計] = (NumericUpDown13.Value) * [變數MF510] '午餐
        Dim [變數MF521午小計] = (NumericUpDown14.Value) * [變數MF521] '午餐
        Dim [變數MF522午小計] = (NumericUpDown15.Value) * [變數MF522] '午餐
        Dim [變數MF530午小計] = (NumericUpDown16.Value) * [變數MF530] '午餐
        Dim [變數MF600午小計] = (NumericUpDown17.Value) * [變數MF600] '午餐
        Dim [變數MF700午小計] = (NumericUpDown18.Value) * [變數MF700] '午餐
        Dim [變數MF801午小計] = (NumericUpDown34.Value) * [變數MF801] '午餐
        Dim [變數MF802午小計] = (NumericUpDown35.Value) * [變數MF802] '午餐
        Dim [變數MF301晚小計] = (NumericUpDown33.Value) * [變數MF301] '晚餐
        Dim [變數MF302晚小計] = (NumericUpDown32.Value) * [變數MF302] '晚餐
        Dim [變數MF303晚小計] = (NumericUpDown31.Value) * [變數MF303] '晚餐
        Dim [變數MF300晚小計] = (NumericUpDown30.Value) * [變數MF300] '晚餐
        Dim [變數MF401晚小計] = (NumericUpDown29.Value) * [變數MF401] '晚餐
        Dim [變數MF402晚小計] = (NumericUpDown28.Value) * [變數MF402] '晚餐
        Dim [變數MF403晚小計] = (NumericUpDown27.Value) * [變數MF403] '晚餐
        Dim [變數MF400晚小計] = (NumericUpDown26.Value) * [變數MF400] '晚餐
        Dim [變數MF500晚小計] = (NumericUpDown25.Value) * [變數MF500] '晚餐
        Dim [變數MF510晚小計] = (NumericUpDown24.Value) * [變數MF510] '晚餐
        Dim [變數MF521晚小計] = (NumericUpDown23.Value) * [變數MF521] '晚餐
        Dim [變數MF522晚小計] = (NumericUpDown22.Value) * [變數MF522] '晚餐
        Dim [變數MF530晚小計] = (NumericUpDown21.Value) * [變數MF530] '晚餐
        Dim [變數MF600晚小計] = (NumericUpDown20.Value) * [變數MF600] '晚餐
        Dim [變數MF700晚小計] = (NumericUpDown19.Value) * [變數MF700] '晚餐
        Dim [變數MF801晚小計] = (NumericUpDown36.Value) * [變數MF801] '晚餐
        Dim [變數MF802晚小計] = (NumericUpDown37.Value) * [變數MF802] '晚餐

        U01.A01 = NumericUpDown1.Value
        U02.A01 = NumericUpDown2.Value
        U03.A01 = NumericUpDown3.Value
        U04.A01 = NumericUpDown4.Value
        U05.A01 = NumericUpDown5.Value
        U06.A01 = NumericUpDown6.Value
        U07.A01 = NumericUpDown7.Value
        U08.A01 = NumericUpDown8.Value
        U09.A01 = NumericUpDown9.Value
        U10.A01 = NumericUpDown10.Value
        U11.A01 = NumericUpDown11.Value
        U12.A01 = NumericUpDown12.Value
        U13.A01 = NumericUpDown13.Value
        U14.A01 = NumericUpDown14.Value
        U15.A01 = NumericUpDown15.Value
        U16.A01 = NumericUpDown16.Value
        U17.A01 = NumericUpDown17.Value
        U18.A01 = NumericUpDown18.Value
        U34.A01 = NumericUpDown34.Value
        U35.A01 = NumericUpDown35.Value
        U33.A01 = NumericUpDown33.Value
        U32.A01 = NumericUpDown32.Value
        U31.A01 = NumericUpDown31.Value
        U30.A01 = NumericUpDown30.Value
        U29.A01 = NumericUpDown29.Value
        U28.A01 = NumericUpDown28.Value
        U27.A01 = NumericUpDown27.Value
        U26.A01 = NumericUpDown26.Value
        U25.A01 = NumericUpDown25.Value
        U24.A01 = NumericUpDown24.Value
        U23.A01 = NumericUpDown23.Value
        U22.A01 = NumericUpDown22.Value
        U21.A01 = NumericUpDown21.Value
        U20.A01 = NumericUpDown20.Value
        U19.A01 = NumericUpDown19.Value
        U36.A01 = NumericUpDown36.Value
        U37.A01 = NumericUpDown37.Value


        '小計加總設定
        Label7.Text = ("小計: " & [變數MF111小計] & "元")
        Label8.Text = ("小計: " & [變數MF112小計] & "元")
        Label12.Text = ("小計: " & [變數MF210小計] & "元")
        Label47.Text = ("小計: " & ([變數MF301午小計] + [變數MF301晚小計]) & "元")
        Label48.Text = ("小計: " & ([變數MF302午小計] + [變數MF302晚小計]) & "元")
        Label49.Text = ("小計: " & ([變數MF303午小計] + [變數MF303晚小計]) & "元")
        Label50.Text = ("小計: " & ([變數MF300午小計] + [變數MF300晚小計]) & "元")
        Label51.Text = ("小計: " & ([變數MF401午小計] + [變數MF401晚小計]) & "元")
        Label52.Text = ("小計: " & ([變數MF402午小計] + [變數MF402晚小計]) & "元")
        Label53.Text = ("小計: " & ([變數MF403午小計] + [變數MF403晚小計]) & "元")
        Label54.Text = ("小計: " & ([變數MF400午小計] + [變數MF400晚小計]) & "元")
        Label55.Text = ("小計: " & ([變數MF500午小計] + [變數MF500晚小計]) & "元")
        Label56.Text = ("小計: " & ([變數MF510午小計] + [變數MF510晚小計]) & "元")
        Label57.Text = ("小計: " & ([變數MF521午小計] + [變數MF521晚小計]) & "元")
        Label58.Text = ("小計: " & ([變數MF522午小計] + [變數MF522晚小計]) & "元")
        Label59.Text = ("小計: " & ([變數MF530午小計] + [變數MF530晚小計]) & "元")
        Label60.Text = ("小計: " & ([變數MF600午小計] + [變數MF600晚小計]) & "元")
        Label61.Text = ("小計: " & ([變數MF700午小計] + [變數MF700晚小計]) & "元")
        Label62.Text = ("小計: " & ([變數MF801午小計] + [變數MF801晚小計]) & "元")
        Label63.Text = ("小計: " & ([變數MF802午小計] + [變數MF802晚小計]) & "元")

        Dim [午總計] = (
        [變數MF111小計] + [變數MF112小計] + [變數MF301午小計] + [變數MF302午小計] +
        [變數MF303午小計] + [變數MF300午小計] + [變數MF401午小計] + [變數MF402午小計] +
        [變數MF403午小計] + [變數MF400午小計] + [變數MF500午小計] + [變數MF510午小計] +
        [變數MF521午小計] + [變數MF522午小計] + [變數MF530午小計] + [變數MF600午小計] +
        [變數MF700午小計] + [變數MF801午小計] + [變數MF802午小計])
        Dim [晚總計] = (
        [變數MF210小計] + [變數MF301晚小計] + [變數MF302晚小計] + [變數MF303晚小計] +
        [變數MF300晚小計] + [變數MF401晚小計] + [變數MF402晚小計] + [變數MF403晚小計] +
        [變數MF400晚小計] + [變數MF500晚小計] + [變數MF510晚小計] + [變數MF521晚小計] +
        [變數MF522晚小計] + [變數MF530晚小計] + [變數MF600晚小計] + [變數MF700晚小計] +
        [變數MF801晚小計] + [變數MF802晚小計])

        Label101.Text = ([午總計] & " 元")
        Label102.Text = ([晚總計] & " 元")
        Label103.Text = (([午總計] + [晚總計]) & " 元")

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        MsgBox("#電話:" & vbCrLf% & " 5552020#108", 64, "聯絡福利社")
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        MsgBox("#目前尚未開放使用", 64, "訂餐紀錄")
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        MsgBox("#開發者:" & vbCrLf% & "#聯絡方式:", 64, "關於作者")
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        MsgBox("#已是最新版本!" & vbCrLf% & "#版本資訊:Ver 1.2.0 202001071050.001", 48, "檢查新版本")

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim 檔案寫入時間 As String = DateTime.Now.ToString("yyyy-MM-dd") '設定時間變數(日期YYYY-MM-DD)
        Dim TXT檔案寫入位置 As String = ("C:\VB-app\" & 檔案寫入時間 & "\") '設定解壓縮位置
        Dim 數字星期 = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "DATE3.txt") '星期幾(數字表示)
        '將以星期日為第一天修正為星期一
        If 數字星期 = 1 Then
            數字星期 = 7
            GoTo A
        ElseIf 數字星期 = 2 Then
            數字星期 = 1
            GoTo A
        ElseIf 數字星期 = 3 Then
            數字星期 = 2
            GoTo A
        ElseIf 數字星期 = 4 Then
            數字星期 = 3
            GoTo A
        ElseIf 數字星期 = 5 Then
            數字星期 = 4
            GoTo A
        ElseIf 數字星期 = 6 Then
            數字星期 = 5
            GoTo A
        ElseIf 數字星期 = 7 Then
            數字星期 = 6
            GoTo A
        End If

A:
        Dim 便當1 As String
        Dim 便當2 As String
        Dim 晚餐1 As String

        '每日不同餐點設定
        If 數字星期 = 1 Then '星期一
            便當1 = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F111.txt") '星期一便當-1 '午餐便當種類-1
            便當2 = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F112.txt") '星期一便當-2 '午餐便當種類-2
            晚餐1 = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F210.txt") '星期一晚上副食
        End If
        If 數字星期 = 2 Then '星期二
            便當1 = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F121.txt") '星期二便當-1 '午餐便當種類-1
            便當2 = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F122.txt") '星期二便當-2 '午餐便當種類-2
            晚餐1 = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F220.txt") '星期二晚上副食
        End If
        If 數字星期 = 3 Then '星期三
            便當1 = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F131.txt") '星期三便當-1 '午餐便當種類-1
            便當2 = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F132.txt") '星期三便當-2 '午餐便當種類-2
            晚餐1 = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F230.txt") '星期三晚上副食
        End If
        If 數字星期 = 4 Then '星期四
            便當1 = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F141.txt") '星期四便當-1 '午餐便當種類-1
            便當2 = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F142.txt") '星期四便當-2 '午餐便當種類-2
            晚餐1 = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F240.txt") '星期四晚上副食
        End If
        If 數字星期 = 5 Then '星期五
            便當1 = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F151.txt") '星期五便當-1  '午餐便當種類-1
            便當2 = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F152.txt") '星期五便當-2  '午餐便當種類-2
            晚餐1 = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F250.txt") '星期五晚上副食
        End If

        Dim 訂餐紀錄位置 As String = ("訂餐紀錄" & 檔案寫入時間 & ".txt") '設定儲存資料夾
        If (System.IO.File.Exists(訂餐紀錄位置)) Then '若訂餐紀錄存在
            System.IO.File.Delete(訂餐紀錄位置) '刪除訂餐紀錄
        Else
        End If
        Dim sw As StreamWriter = New StreamWriter("訂餐紀錄" & 檔案寫入時間 & ".txt")
        If U01.A01 <> 0 Then
            sw.Write(便當1 & "*" & U01.A01 & vbCrLf)
        End If
        If U02.A01 <> 0 Then
            sw.Write(便當2 & "*" & U02.A01 & vbCrLf)
        End If
        If U04.A01 <> 0 Then
            sw.Write(My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F301.txt") & "(午)" & "*" & U04.A01 & vbCrLf)
        End If
        If U05.A01 <> 0 Then
            sw.Write(My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F302.txt") & "(午)" & "*" & U05.A01 & vbCrLf)
        End If
        If U06.A01 <> 0 Then
            sw.Write(My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F303.txt") & "(午)" & "*" & U06.A01 & vbCrLf)
        End If
        If U07.A01 <> 0 Then
            sw.Write(My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F300.txt") & "(午)" & "*" & U07.A01 & vbCrLf)
        End If
        If U08.A01 <> 0 Then
            sw.Write(My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F401.txt") & "(午)" & "*" & U08.A01 & vbCrLf)
        End If
        If U09.A01 <> 0 Then
            sw.Write(My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F402.txt") & "(午)" & "*" & U09.A01 & vbCrLf)
        End If
        If U10.A01 <> 0 Then
            sw.Write(My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F403.txt") & "(午)" & "*" & U10.A01 & vbCrLf)
        End If
        If U11.A01 <> 0 Then
            sw.Write(My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F400.txt") & "(午)" & "*" & U11.A01 & vbCrLf)
        End If
        If U12.A01 <> 0 Then
            sw.Write(My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F500.txt") & "(午)" & "*" & U12.A01 & vbCrLf)
        End If
        If U13.A01 <> 0 Then
            sw.Write(My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F510.txt") & "(午)" & "*" & U13.A01 & vbCrLf)
        End If
        If U14.A01 <> 0 Then
            sw.Write(My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F521.txt") & "(午)" & "*" & U14.A01 & vbCrLf)
        End If
        If U15.A01 <> 0 Then
            sw.Write(My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F522.txt") & "(午)" & "*" & U15.A01 & vbCrLf)
        End If
        If U16.A01 <> 0 Then
            sw.Write(My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F530.txt") & "(午)" & "*" & U16.A01 & vbCrLf)
        End If
        If U17.A01 <> 0 Then
            sw.Write(My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F600.txt") & "(午)" & "*" & U17.A01 & vbCrLf)
        End If
        If U18.A01 <> 0 Then
            sw.Write(My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F700.txt") & "(午)" & "*" & U18.A01 & vbCrLf)
        End If
        If U34.A01 <> 0 Then
            sw.Write(My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F801.txt") & "(午)" & "*" & U34.A01 & vbCrLf)
        End If
        If U35.A01 <> 0 Then
            sw.Write(My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F802.txt") & "(午)" & "*" & U35.A01 & vbCrLf)
        End If
        If U03.A01 <> 0 Then
            sw.Write(晚餐1 & "*" & U03.A01 & vbCrLf)
        End If
        If U33.A01 <> 0 Then
            sw.Write(My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F301.txt") & "(晚)" & "*" & U33.A01 & vbCrLf)
        End If
        If U32.A01 <> 0 Then
            sw.Write(My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F302.txt") & "(晚)" & "*" & U32.A01 & vbCrLf)
        End If
        If U31.A01 <> 0 Then
            sw.Write(My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F303.txt") & "(晚)" & "*" & U31.A01 & vbCrLf)
        End If
        If U30.A01 <> 0 Then
            sw.Write(My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F300.txt") & "(晚)" & "*" & U30.A01 & vbCrLf)
        End If
        If U29.A01 <> 0 Then
            sw.Write(My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F401.txt") & "(晚)" & "*" & U29.A01 & vbCrLf)
        End If
        If U28.A01 <> 0 Then
            sw.Write(My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F402.txt") & "(晚)" & "*" & U28.A01 & vbCrLf)
        End If
        If U27.A01 <> 0 Then
            sw.Write(My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F403.txt") & "(晚)" & "*" & U27.A01 & vbCrLf)
        End If
        If U26.A01 <> 0 Then
            sw.Write(My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F400.txt") & "(晚)" & "*" & U26.A01 & vbCrLf)
        End If
        If U25.A01 <> 0 Then
            sw.Write(My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F500.txt") & "(晚)" & "*" & U25.A01 & vbCrLf)
        End If
        If U24.A01 <> 0 Then
            sw.Write(My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F510.txt") & "(晚)" & "*" & U24.A01 & vbCrLf)
        End If
        If U23.A01 <> 0 Then
            sw.Write(My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F521.txt") & "(晚)" & "*" & U23.A01 & vbCrLf)
        End If
        If U22.A01 <> 0 Then
            sw.Write(My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F522.txt") & "(晚)" & "*" & U22.A01 & vbCrLf)
        End If
        If U21.A01 <> 0 Then
            sw.Write(My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F530.txt") & "(晚)" & "*" & U21.A01 & vbCrLf)
        End If
        If U20.A01 <> 0 Then
            sw.Write(My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F600.txt") & "(晚)" & "*" & U20.A01 & vbCrLf)
        End If
        If U19.A01 <> 0 Then
            sw.Write(My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F700.txt") & "(晚)" & "*" & U19.A01 & vbCrLf)
        End If
        If U36.A01 <> 0 Then
            sw.Write(My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F801.txt") & "(晚)" & "*" & U36.A01 & vbCrLf)
        End If
        If U37.A01 <> 0 Then
            sw.Write(My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F802.txt") & "(晚)" & "*" & U37.A01 & vbCrLf)
        End If

        sw.Close()
        Dim X As Integer
        X = MsgBox(“請確認以下餐點是否正確” & vbCrLf & "正確無誤請按下[確認]" & vbCrLf & "有誤請按下[取消]進行修改" & vbCrLf & vbCrLf & My.Computer.FileSystem.ReadAllText(訂餐紀錄位置), 1 + 48, "最終確認,確認後將無法修改!")
        If X = 1 Then
            MsgBox(“訂單已提交,謝謝”, 64, "提交成功")
            If (System.IO.File.Exists("C:\VB-app\訂餐紀錄" & 檔案寫入時間 & ".txt")) Then '若訂餐紀錄存在
                System.IO.File.Delete("C:\VB-app\訂餐紀錄" & 檔案寫入時間 & ".txt") '刪除訂餐紀錄
            Else
            End If
            My.Computer.FileSystem.MoveFile(訂餐紀錄位置, "C:\VB-app\訂餐紀錄" & 檔案寫入時間 & ".txt")
            System.IO.File.Delete(訂餐紀錄位置)
            MsgBox(“訂單紀錄已保存”, 64, "紀錄成功")
        Else
            MsgBox(“訂單尚未送出,請返回修改後重新提交”, 64, "使用者已取消")
        End If
    End Sub
    Public Class U01
        Public Shared A01 As Integer
    End Class
    Public Class U02
        Public Shared A01 As Integer
    End Class
    Public Class U03
        Public Shared A01 As Integer
    End Class
    Public Class U04
        Public Shared A01 As Integer
    End Class
    Public Class U05
        Public Shared A01 As Integer
    End Class
    Public Class U06
        Public Shared A01 As Integer
    End Class
    Public Class U07
        Public Shared A01 As Integer
    End Class
    Public Class U08
        Public Shared A01 As Integer
    End Class
    Public Class U09
        Public Shared A01 As Integer
    End Class
    Public Class U10
        Public Shared A01 As Integer
    End Class
    Public Class U11
        Public Shared A01 As Integer
    End Class
    Public Class U12
        Public Shared A01 As Integer
    End Class
    Public Class U13
        Public Shared A01 As Integer
    End Class
    Public Class U14
        Public Shared A01 As Integer
    End Class
    Public Class U15
        Public Shared A01 As Integer
    End Class
    Public Class U16
        Public Shared A01 As Integer
    End Class
    Public Class U17
        Public Shared A01 As Integer
    End Class
    Public Class U18
        Public Shared A01 As Integer
    End Class
    Public Class U19
        Public Shared A01 As Integer
    End Class
    Public Class U20
        Public Shared A01 As Integer
    End Class
    Public Class U21
        Public Shared A01 As Integer
    End Class
    Public Class U22
        Public Shared A01 As Integer
    End Class
    Public Class U23
        Public Shared A01 As Integer
    End Class
    Public Class U24
        Public Shared A01 As Integer
    End Class
    Public Class U25
        Public Shared A01 As Integer
    End Class
    Public Class U26
        Public Shared A01 As Integer
    End Class
    Public Class U27
        Public Shared A01 As Integer
    End Class
    Public Class U28
        Public Shared A01 As Integer
    End Class
    Public Class U29
        Public Shared A01 As Integer
    End Class
    Public Class U30
        Public Shared A01 As Integer
    End Class
    Public Class U31
        Public Shared A01 As Integer
    End Class
    Public Class U32
        Public Shared A01 As Integer
    End Class
    Public Class U33
        Public Shared A01 As Integer
    End Class
    Public Class U34
        Public Shared A01 As Integer
    End Class
    Public Class U35
        Public Shared A01 As Integer
    End Class
    Public Class U36
        Public Shared A01 As Integer
    End Class
    Public Class U37
        Public Shared A01 As Integer
    End Class

End Class