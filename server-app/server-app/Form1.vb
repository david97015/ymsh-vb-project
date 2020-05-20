Imports System.IO
Imports System.Text
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

        Dim 檔案創建日期 As DateTime = File.GetLastAccessTime("C:\VB-app\vb.zip")
        ToolStripStatusLabel3.Text = ("上次編輯鎖定日期:" & 檔案創建日期)
        Label46.Text = "狀態:尚未鎖定" '尚未鎖定
        Label46.ForeColor = Color.Red '尚未鎖定
        Button2.ForeColor = Color.Red '尚未鎖定
        PictureBox1.Visible = False '尚未鎖定

        ToolStripStatusLabel1.Text = ("使用者名稱:" & System.Net.Dns.GetHostName().ToString()) '取得使用者名稱
        ToolStripStatusLabel2.Text = ("IP位置:" & GetIPaddress()) '取得IP位置
        Dim 檔案寫入時間 As String = DateTime.Now.ToString("yyyy-MM-dd") '設定時間變數(日期YYYY-MM-DD)
        Dim TXT檔案寫入位置 As String = ("C:\VB-app\" & 檔案寫入時間 & "\") '設定解壓縮位置
        '記憶上次餐點
        TextBox25.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F111.txt")
        TextBox24.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F112.txt")
        TextBox23.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F121.txt")
        TextBox22.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F122.txt")
        TextBox21.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F131.txt")
        TextBox20.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F132.txt")
        TextBox19.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F141.txt")
        TextBox18.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F142.txt")
        TextBox17.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F151.txt")
        TextBox16.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F152.txt")
        TextBox32.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F210.txt")
        TextBox31.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F220.txt")
        TextBox30.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F230.txt")
        TextBox29.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F240.txt")
        TextBox28.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F250.txt")
        TextBox1.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F301.txt")
        TextBox2.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F302.txt")
        TextBox3.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F303.txt")
        TextBox4.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F300.txt")
        TextBox5.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F401.txt")
        TextBox6.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F402.txt")
        TextBox7.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F403.txt")
        TextBox8.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F400.txt")
        TextBox9.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F500.txt")
        TextBox10.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F510.txt")
        TextBox11.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F521.txt")
        TextBox12.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F522.txt")
        TextBox13.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F530.txt")
        TextBox14.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F600.txt")
        TextBox15.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F700.txt")
        TextBox27.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F801.txt")
        TextBox26.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "F802.txt")
        TextBox34.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "INFO.txt")
        TextBox33.Text = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "DATE1.txt")

        '記憶上次金額
        NumericUpDown1.Value = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF111.txt")
        NumericUpDown2.Value = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF112.txt")
        NumericUpDown20.Value = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF121.txt")
        NumericUpDown19.Value = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF122.txt")
        NumericUpDown24.Value = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF131.txt")
        NumericUpDown23.Value = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF132.txt")
        NumericUpDown22.Value = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF141.txt")
        NumericUpDown21.Value = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF142.txt")
        NumericUpDown26.Value = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF151.txt")
        NumericUpDown25.Value = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF152.txt")
        NumericUpDown32.Value = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF210.txt")
        NumericUpDown31.Value = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF220.txt")
        NumericUpDown30.Value = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF230.txt")
        NumericUpDown29.Value = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF240.txt")
        NumericUpDown3.Value = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF250.txt")
        NumericUpDown4.Value = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF301.txt")
        NumericUpDown5.Value = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF302.txt")
        NumericUpDown6.Value = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF303.txt")
        NumericUpDown7.Value = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF300.txt")
        NumericUpDown8.Value = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF401.txt")
        NumericUpDown9.Value = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF402.txt")
        NumericUpDown10.Value = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF403.txt")
        NumericUpDown11.Value = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF400.txt")
        NumericUpDown12.Value = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF500.txt")
        NumericUpDown13.Value = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF510.txt")
        NumericUpDown14.Value = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF521.txt")
        NumericUpDown15.Value = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF522.txt")
        NumericUpDown16.Value = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF530.txt")
        NumericUpDown17.Value = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF600.txt")
        NumericUpDown18.Value = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF700.txt")
        NumericUpDown27.Value = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF801.txt")
        NumericUpDown28.Value = My.Computer.FileSystem.ReadAllText(TXT檔案寫入位置 & "MF802.txt")
    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Me.Text = ("新竹縣私立義民高級中學點餐系統-server端 " + "現在時間:" + Now) '編輯標題列
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Dim 檔案寫入時間 As String = DateTime.Now.ToString("yyyy-MM-dd") '設定時間(YYYY-MM-DD)
        Dim 資料夾位置 As String = ("C:\VB-app\" & 檔案寫入時間) '設定儲存資料夾
        If (System.IO.Directory.Exists(資料夾位置)) Then '若資料夾存在
            System.IO.Directory.Delete(資料夾位置, True) '刪除資料夾
        Else
        End If
        My.Computer.FileSystem.CreateDirectory(資料夾位置)
        Dim path01 As String = (資料夾位置 & "\" & "DATE1" & ".txt") '設定檔名
        Dim fs01 As FileStream = File.Create(path01) '新建或是覆蓋
        Dim info01 As Byte() = New UTF8Encoding(True).GetBytes(TextBox33.Text) '在TXT中加入文字
        fs01.Write(info01, 0, info01.Length)
        fs01.Close()

        '午餐便當部分
        Dim path02 As String = (資料夾位置 & "\" & "F111" & ".txt") '設定檔名
        Dim fs02 As FileStream = File.Create(path02) '新建或是覆蓋
        Dim info02 As Byte() = New UTF8Encoding(True).GetBytes(TextBox25.Text) '在TXT中加入文字
        fs02.Write(info02, 0, info02.Length)
        fs02.Close()

        Dim path03 As String = (資料夾位置 & "\" & "F112" & ".txt") '設定檔名
        Dim fs03 As FileStream = File.Create(path03) '新建或是覆蓋
        Dim info03 As Byte() = New UTF8Encoding(True).GetBytes(TextBox24.Text) '在TXT中加入文字
        fs03.Write(info03, 0, info03.Length)
        fs03.Close()

        Dim path04 As String = (資料夾位置 & "\" & "F121" & ".txt") '設定檔名
        Dim fs04 As FileStream = File.Create(path04) '新建或是覆蓋
        Dim info04 As Byte() = New UTF8Encoding(True).GetBytes(TextBox23.Text) '在TXT中加入文字
        fs04.Write(info04, 0, info04.Length)
        fs04.Close()

        Dim path05 As String = (資料夾位置 & "\" & "F122" & ".txt") '設定檔名
        Dim fs05 As FileStream = File.Create(path05) '新建或是覆蓋
        Dim info05 As Byte() = New UTF8Encoding(True).GetBytes(TextBox22.Text) '在TXT中加入文字
        fs05.Write(info05, 0, info05.Length)
        fs05.Close()

        Dim path06 As String = (資料夾位置 & "\" & "F131" & ".txt") '設定檔名
        Dim fs06 As FileStream = File.Create(path06) '新建或是覆蓋
        Dim info06 As Byte() = New UTF8Encoding(True).GetBytes(TextBox21.Text) '在TXT中加入文字
        fs06.Write(info06, 0, info06.Length)
        fs06.Close()

        Dim path07 As String = (資料夾位置 & "\" & "F132" & ".txt") '設定檔名
        Dim fs07 As FileStream = File.Create(path07) '新建或是覆蓋
        Dim info07 As Byte() = New UTF8Encoding(True).GetBytes(TextBox20.Text) '在TXT中加入文字
        fs07.Write(info07, 0, info07.Length)
        fs07.Close()

        Dim path08 As String = (資料夾位置 & "\" & "F141" & ".txt") '設定檔名
        Dim fs08 As FileStream = File.Create(path08) '新建或是覆蓋
        Dim info08 As Byte() = New UTF8Encoding(True).GetBytes(TextBox19.Text) '在TXT中加入文字
        fs08.Write(info08, 0, info08.Length)
        fs08.Close()

        Dim path09 As String = (資料夾位置 & "\" & "F142" & ".txt") '設定檔名
        Dim fs09 As FileStream = File.Create(path09) '新建或是覆蓋
        Dim info09 As Byte() = New UTF8Encoding(True).GetBytes(TextBox18.Text) '在TXT中加入文字
        fs09.Write(info09, 0, info09.Length)
        fs09.Close()

        Dim path10 As String = (資料夾位置 & "\" & "F151" & ".txt") '設定檔名
        Dim fs10 As FileStream = File.Create(path10) '新建或是覆蓋
        Dim info10 As Byte() = New UTF8Encoding(True).GetBytes(TextBox17.Text) '在TXT中加入文字
        fs10.Write(info10, 0, info10.Length)
        fs10.Close()

        Dim path11 As String = (資料夾位置 & "\" & "F152" & ".txt") '設定檔名
        Dim fs11 As FileStream = File.Create(path11) '新建或是覆蓋
        Dim info11 As Byte() = New UTF8Encoding(True).GetBytes(TextBox16.Text) '在TXT中加入文字
        fs11.Write(info11, 0, info11.Length)
        fs11.Close()

        '午餐便當金額
        Dim path12 As String = (資料夾位置 & "\" & "MF111" & ".txt") '設定檔名
        Dim fs12 As FileStream = File.Create(path12) '新建或是覆蓋
        Dim info12 As Byte() = New UTF8Encoding(True).GetBytes(NumericUpDown1.Value) '在TXT中加入文字
        fs12.Write(info12, 0, info12.Length)
        fs12.Close()

        Dim path13 As String = (資料夾位置 & "\" & "MF112" & ".txt") '設定檔名
        Dim fs13 As FileStream = File.Create(path13) '新建或是覆蓋
        Dim info13 As Byte() = New UTF8Encoding(True).GetBytes(NumericUpDown2.Value) '在TXT中加入文字
        fs13.Write(info13, 0, info13.Length)
        fs13.Close()

        Dim path14 As String = (資料夾位置 & "\" & "MF121" & ".txt") '設定檔名
        Dim fs14 As FileStream = File.Create(path14) '新建或是覆蓋
        Dim info14 As Byte() = New UTF8Encoding(True).GetBytes(NumericUpDown20.Value) '在TXT中加入文字
        fs14.Write(info14, 0, info14.Length)
        fs14.Close()

        Dim path15 As String = (資料夾位置 & "\" & "MF122" & ".txt") '設定檔名
        Dim fs15 As FileStream = File.Create(path15) '新建或是覆蓋
        Dim info15 As Byte() = New UTF8Encoding(True).GetBytes(NumericUpDown19.Value) '在TXT中加入文字
        fs15.Write(info15, 0, info15.Length)
        fs15.Close()

        Dim path16 As String = (資料夾位置 & "\" & "MF131" & ".txt") '設定檔名
        Dim fs16 As FileStream = File.Create(path16) '新建或是覆蓋
        Dim info16 As Byte() = New UTF8Encoding(True).GetBytes(NumericUpDown24.Value) '在TXT中加入文字
        fs16.Write(info16, 0, info16.Length)
        fs16.Close()

        Dim path17 As String = (資料夾位置 & "\" & "MF132" & ".txt") '設定檔名
        Dim fs17 As FileStream = File.Create(path17) '新建或是覆蓋
        Dim info17 As Byte() = New UTF8Encoding(True).GetBytes(NumericUpDown23.Value) '在TXT中加入文字
        fs17.Write(info17, 0, info17.Length)
        fs17.Close()

        Dim path18 As String = (資料夾位置 & "\" & "MF141" & ".txt") '設定檔名
        Dim fs18 As FileStream = File.Create(path18) '新建或是覆蓋
        Dim info18 As Byte() = New UTF8Encoding(True).GetBytes(NumericUpDown22.Value) '在TXT中加入文字
        fs18.Write(info18, 0, info18.Length)
        fs18.Close()

        Dim path19 As String = (資料夾位置 & "\" & "MF142" & ".txt") '設定檔名
        Dim fs19 As FileStream = File.Create(path19) '新建或是覆蓋
        Dim info19 As Byte() = New UTF8Encoding(True).GetBytes(NumericUpDown21.Value) '在TXT中加入文字
        fs19.Write(info19, 0, info19.Length)
        fs19.Close()

        Dim path20 As String = (資料夾位置 & "\" & "MF151" & ".txt") '設定檔名
        Dim fs20 As FileStream = File.Create(path20) '新建或是覆蓋
        Dim info20 As Byte() = New UTF8Encoding(True).GetBytes(NumericUpDown26.Value) '在TXT中加入文字
        fs20.Write(info20, 0, info20.Length)
        fs20.Close()

        Dim path21 As String = (資料夾位置 & "\" & "MF152" & ".txt") '設定檔名
        Dim fs21 As FileStream = File.Create(path21) '新建或是覆蓋
        Dim info21 As Byte() = New UTF8Encoding(True).GetBytes(NumericUpDown25.Value) '在TXT中加入文字
        fs21.Write(info21, 0, info21.Length)
        fs21.Close()

        '晚餐副食部分
        Dim path22 As String = (資料夾位置 & "\" & "F210" & ".txt") '設定檔名
        Dim fs22 As FileStream = File.Create(path22) '新建或是覆蓋
        Dim info22 As Byte() = New UTF8Encoding(True).GetBytes(TextBox32.Text) '在TXT中加入文字
        fs22.Write(info22, 0, info22.Length)
        fs22.Close()

        Dim path23 As String = (資料夾位置 & "\" & "F220" & ".txt") '設定檔名
        Dim fs23 As FileStream = File.Create(path23) '新建或是覆蓋
        Dim info23 As Byte() = New UTF8Encoding(True).GetBytes(TextBox31.Text) '在TXT中加入文字
        fs23.Write(info23, 0, info23.Length)
        fs23.Close()

        Dim path24 As String = (資料夾位置 & "\" & "F230" & ".txt") '設定檔名
        Dim fs24 As FileStream = File.Create(path24) '新建或是覆蓋
        Dim info24 As Byte() = New UTF8Encoding(True).GetBytes(TextBox30.Text) '在TXT中加入文字
        fs24.Write(info24, 0, info24.Length)
        fs24.Close()

        Dim path25 As String = (資料夾位置 & "\" & "F240" & ".txt") '設定檔名
        Dim fs25 As FileStream = File.Create(path25) '新建或是覆蓋
        Dim info25 As Byte() = New UTF8Encoding(True).GetBytes(TextBox29.Text) '在TXT中加入文字
        fs25.Write(info25, 0, info25.Length)
        fs25.Close()

        Dim path26 As String = (資料夾位置 & "\" & "F250" & ".txt") '設定檔名
        Dim fs26 As FileStream = File.Create(path26) '新建或是覆蓋
        Dim info26 As Byte() = New UTF8Encoding(True).GetBytes(TextBox28.Text) '在TXT中加入文字
        fs26.Write(info26, 0, info26.Length)
        fs26.Close()

        '晚餐副食金額
        Dim path27 As String = (資料夾位置 & "\" & "MF210" & ".txt") '設定檔名
        Dim fs27 As FileStream = File.Create(path27) '新建或是覆蓋
        Dim info27 As Byte() = New UTF8Encoding(True).GetBytes(NumericUpDown32.Value) '在TXT中加入文字
        fs27.Write(info27, 0, info27.Length)
        fs27.Close()

        Dim path28 As String = (資料夾位置 & "\" & "MF220" & ".txt") '設定檔名
        Dim fs28 As FileStream = File.Create(path28) '新建或是覆蓋
        Dim info28 As Byte() = New UTF8Encoding(True).GetBytes(NumericUpDown31.Value) '在TXT中加入文字
        fs28.Write(info28, 0, info28.Length)
        fs28.Close()

        Dim path29 As String = (資料夾位置 & "\" & "MF230" & ".txt") '設定檔名
        Dim fs29 As FileStream = File.Create(path29) '新建或是覆蓋
        Dim info29 As Byte() = New UTF8Encoding(True).GetBytes(NumericUpDown30.Value) '在TXT中加入文字
        fs29.Write(info29, 0, info29.Length)
        fs29.Close()

        Dim path30 As String = (資料夾位置 & "\" & "MF240" & ".txt") '設定檔名
        Dim fs30 As FileStream = File.Create(path30) '新建或是覆蓋
        Dim info30 As Byte() = New UTF8Encoding(True).GetBytes(NumericUpDown29.Value) '在TXT中加入文字
        fs30.Write(info30, 0, info30.Length)
        fs30.Close()

        Dim path31 As String = (資料夾位置 & "\" & "MF250" & ".txt") '設定檔名
        Dim fs31 As FileStream = File.Create(path31) '新建或是覆蓋
        Dim info31 As Byte() = New UTF8Encoding(True).GetBytes(NumericUpDown3.Value) '在TXT中加入文字
        fs31.Write(info31, 0, info31.Length)
        fs31.Close()

        '全日供應餐點
        Dim path32 As String = (資料夾位置 & "\" & "F301" & ".txt") '設定檔名
        Dim fs32 As FileStream = File.Create(path32) '新建或是覆蓋
        Dim info32 As Byte() = New UTF8Encoding(True).GetBytes(TextBox1.Text) '在TXT中加入文字
        fs32.Write(info32, 0, info32.Length)
        fs32.Close()

        Dim path33 As String = (資料夾位置 & "\" & "F302" & ".txt") '設定檔名
        Dim fs33 As FileStream = File.Create(path33) '新建或是覆蓋
        Dim info33 As Byte() = New UTF8Encoding(True).GetBytes(TextBox2.Text) '在TXT中加入文字
        fs33.Write(info33, 0, info33.Length)
        fs33.Close()

        Dim path34 As String = (資料夾位置 & "\" & "F303" & ".txt") '設定檔名
        Dim fs34 As FileStream = File.Create(path34) '新建或是覆蓋
        Dim info34 As Byte() = New UTF8Encoding(True).GetBytes(TextBox3.Text) '在TXT中加入文字
        fs34.Write(info34, 0, info34.Length)
        fs34.Close()

        Dim path35 As String = (資料夾位置 & "\" & "F300" & ".txt") '設定檔名
        Dim fs35 As FileStream = File.Create(path35) '新建或是覆蓋
        Dim info35 As Byte() = New UTF8Encoding(True).GetBytes(TextBox4.Text) '在TXT中加入文字
        fs35.Write(info35, 0, info35.Length)
        fs35.Close()

        Dim path36 As String = (資料夾位置 & "\" & "F401" & ".txt") '設定檔名
        Dim fs36 As FileStream = File.Create(path36) '新建或是覆蓋
        Dim info36 As Byte() = New UTF8Encoding(True).GetBytes(TextBox5.Text) '在TXT中加入文字
        fs36.Write(info36, 0, info36.Length)
        fs36.Close()

        Dim path37 As String = (資料夾位置 & "\" & "F402" & ".txt") '設定檔名
        Dim fs37 As FileStream = File.Create(path37) '新建或是覆蓋
        Dim info37 As Byte() = New UTF8Encoding(True).GetBytes(TextBox6.Text) '在TXT中加入文字
        fs37.Write(info37, 0, info37.Length)
        fs37.Close()

        Dim path38 As String = (資料夾位置 & "\" & "F403" & ".txt") '設定檔名
        Dim fs38 As FileStream = File.Create(path38) '新建或是覆蓋
        Dim info38 As Byte() = New UTF8Encoding(True).GetBytes(TextBox7.Text) '在TXT中加入文字
        fs38.Write(info38, 0, info38.Length)
        fs38.Close()

        Dim path39 As String = (資料夾位置 & "\" & "F400" & ".txt") '設定檔名
        Dim fs39 As FileStream = File.Create(path39) '新建或是覆蓋
        Dim info39 As Byte() = New UTF8Encoding(True).GetBytes(TextBox8.Text) '在TXT中加入文字
        fs39.Write(info39, 0, info39.Length)
        fs39.Close()

        Dim path40 As String = (資料夾位置 & "\" & "F500" & ".txt") '設定檔名
        Dim fs40 As FileStream = File.Create(path40) '新建或是覆蓋
        Dim info40 As Byte() = New UTF8Encoding(True).GetBytes(TextBox9.Text) '在TXT中加入文字
        fs40.Write(info40, 0, info40.Length)
        fs40.Close()

        Dim path41 As String = (資料夾位置 & "\" & "F510" & ".txt") '設定檔名
        Dim fs41 As FileStream = File.Create(path41) '新建或是覆蓋
        Dim info41 As Byte() = New UTF8Encoding(True).GetBytes(TextBox10.Text) '在TXT中加入文字
        fs41.Write(info41, 0, info41.Length)
        fs41.Close()

        Dim path42 As String = (資料夾位置 & "\" & "F521" & ".txt") '設定檔名
        Dim fs42 As FileStream = File.Create(path42) '新建或是覆蓋
        Dim info42 As Byte() = New UTF8Encoding(True).GetBytes(TextBox11.Text) '在TXT中加入文字
        fs42.Write(info42, 0, info42.Length)
        fs42.Close()

        Dim path43 As String = (資料夾位置 & "\" & "F522" & ".txt") '設定檔名
        Dim fs43 As FileStream = File.Create(path43) '新建或是覆蓋
        Dim info43 As Byte() = New UTF8Encoding(True).GetBytes(TextBox12.Text) '在TXT中加入文字
        fs43.Write(info43, 0, info43.Length)
        fs43.Close()

        Dim path44 As String = (資料夾位置 & "\" & "F530" & ".txt") '設定檔名
        Dim fs44 As FileStream = File.Create(path44) '新建或是覆蓋
        Dim info44 As Byte() = New UTF8Encoding(True).GetBytes(TextBox13.Text) '在TXT中加入文字
        fs44.Write(info44, 0, info44.Length)
        fs44.Close()

        Dim path45 As String = (資料夾位置 & "\" & "F600" & ".txt") '設定檔名
        Dim fs45 As FileStream = File.Create(path45) '新建或是覆蓋
        Dim info45 As Byte() = New UTF8Encoding(True).GetBytes(TextBox14.Text) '在TXT中加入文字
        fs45.Write(info45, 0, info45.Length)
        fs45.Close()

        Dim path46 As String = (資料夾位置 & "\" & "F700" & ".txt") '設定檔名
        Dim fs46 As FileStream = File.Create(path46) '新建或是覆蓋
        Dim info46 As Byte() = New UTF8Encoding(True).GetBytes(TextBox15.Text) '在TXT中加入文字
        fs46.Write(info46, 0, info46.Length)
        fs46.Close()

        Dim path47 As String = (資料夾位置 & "\" & "F801" & ".txt") '設定檔名
        Dim fs47 As FileStream = File.Create(path47) '新建或是覆蓋
        Dim info47 As Byte() = New UTF8Encoding(True).GetBytes(TextBox27.Text) '在TXT中加入文字
        fs47.Write(info47, 0, info47.Length)
        fs47.Close()

        Dim path48 As String = (資料夾位置 & "\" & "F802" & ".txt") '設定檔名
        Dim fs48 As FileStream = File.Create(path48) '新建或是覆蓋
        Dim info48 As Byte() = New UTF8Encoding(True).GetBytes(TextBox26.Text) '在TXT中加入文字
        fs48.Write(info48, 0, info48.Length)
        fs48.Close()

        '全天供應金額
        Dim path49 As String = (資料夾位置 & "\" & "MF301" & ".txt") '設定檔名
        Dim fs49 As FileStream = File.Create(path49) '新建或是覆蓋
        Dim info49 As Byte() = New UTF8Encoding(True).GetBytes(NumericUpDown4.Value) '在TXT中加入文字
        fs49.Write(info49, 0, info49.Length)
        fs49.Close()

        Dim path50 As String = (資料夾位置 & "\" & "MF302" & ".txt") '設定檔名
        Dim fs50 As FileStream = File.Create(path50) '新建或是覆蓋
        Dim info50 As Byte() = New UTF8Encoding(True).GetBytes(NumericUpDown5.Value) '在TXT中加入文字
        fs50.Write(info50, 0, info50.Length)
        fs50.Close()

        Dim path51 As String = (資料夾位置 & "\" & "MF303" & ".txt") '設定檔名
        Dim fs51 As FileStream = File.Create(path51) '新建或是覆蓋
        Dim info51 As Byte() = New UTF8Encoding(True).GetBytes(NumericUpDown6.Value) '在TXT中加入文字
        fs51.Write(info51, 0, info51.Length)
        fs51.Close()

        Dim path52 As String = (資料夾位置 & "\" & "MF300" & ".txt") '設定檔名
        Dim fs52 As FileStream = File.Create(path52) '新建或是覆蓋
        Dim info52 As Byte() = New UTF8Encoding(True).GetBytes(NumericUpDown7.Value) '在TXT中加入文字
        fs52.Write(info52, 0, info52.Length)
        fs52.Close()

        Dim path53 As String = (資料夾位置 & "\" & "MF401" & ".txt") '設定檔名
        Dim fs53 As FileStream = File.Create(path53) '新建或是覆蓋
        Dim info53 As Byte() = New UTF8Encoding(True).GetBytes(NumericUpDown8.Value) '在TXT中加入文字
        fs53.Write(info53, 0, info53.Length)
        fs53.Close()

        Dim path54 As String = (資料夾位置 & "\" & "MF402" & ".txt") '設定檔名
        Dim fs54 As FileStream = File.Create(path54) '新建或是覆蓋
        Dim info54 As Byte() = New UTF8Encoding(True).GetBytes(NumericUpDown9.Value) '在TXT中加入文字
        fs54.Write(info54, 0, info54.Length)
        fs54.Close()

        Dim path55 As String = (資料夾位置 & "\" & "MF403" & ".txt") '設定檔名
        Dim fs55 As FileStream = File.Create(path55) '新建或是覆蓋
        Dim info55 As Byte() = New UTF8Encoding(True).GetBytes(NumericUpDown10.Value) '在TXT中加入文字
        fs55.Write(info55, 0, info55.Length)
        fs55.Close()

        Dim path56 As String = (資料夾位置 & "\" & "MF400" & ".txt") '設定檔名
        Dim fs56 As FileStream = File.Create(path56) '新建或是覆蓋
        Dim info56 As Byte() = New UTF8Encoding(True).GetBytes(NumericUpDown11.Value) '在TXT中加入文字
        fs56.Write(info56, 0, info56.Length)
        fs56.Close()

        Dim path57 As String = (資料夾位置 & "\" & "MF500" & ".txt") '設定檔名
        Dim fs57 As FileStream = File.Create(path57) '新建或是覆蓋
        Dim info57 As Byte() = New UTF8Encoding(True).GetBytes(NumericUpDown12.Value) '在TXT中加入文字
        fs57.Write(info57, 0, info57.Length)
        fs57.Close()

        Dim path58 As String = (資料夾位置 & "\" & "MF510" & ".txt") '設定檔名
        Dim fs58 As FileStream = File.Create(path58) '新建或是覆蓋
        Dim info58 As Byte() = New UTF8Encoding(True).GetBytes(NumericUpDown13.Value) '在TXT中加入文字
        fs58.Write(info58, 0, info58.Length)
        fs58.Close()

        Dim path59 As String = (資料夾位置 & "\" & "MF521" & ".txt") '設定檔名
        Dim fs59 As FileStream = File.Create(path59) '新建或是覆蓋
        Dim info59 As Byte() = New UTF8Encoding(True).GetBytes(NumericUpDown14.Value) '在TXT中加入文字
        fs59.Write(info59, 0, info59.Length)
        fs59.Close()

        Dim path60 As String = (資料夾位置 & "\" & "MF522" & ".txt") '設定檔名
        Dim fs60 As FileStream = File.Create(path60) '新建或是覆蓋
        Dim info60 As Byte() = New UTF8Encoding(True).GetBytes(NumericUpDown15.Value) '在TXT中加入文字
        fs60.Write(info60, 0, info60.Length)
        fs60.Close()

        Dim path61 As String = (資料夾位置 & "\" & "MF530" & ".txt") '設定檔名
        Dim fs61 As FileStream = File.Create(path61) '新建或是覆蓋
        Dim info61 As Byte() = New UTF8Encoding(True).GetBytes(NumericUpDown16.Value) '在TXT中加入文字
        fs61.Write(info61, 0, info61.Length)
        fs61.Close()

        Dim path62 As String = (資料夾位置 & "\" & "MF600" & ".txt") '設定檔名
        Dim fs62 As FileStream = File.Create(path62) '新建或是覆蓋
        Dim info62 As Byte() = New UTF8Encoding(True).GetBytes(NumericUpDown17.Value) '在TXT中加入文字
        fs62.Write(info62, 0, info62.Length)
        fs62.Close()

        Dim path63 As String = (資料夾位置 & "\" & "MF700" & ".txt") '設定檔名
        Dim fs63 As FileStream = File.Create(path63) '新建或是覆蓋
        Dim info63 As Byte() = New UTF8Encoding(True).GetBytes(NumericUpDown18.Value) '在TXT中加入文字
        fs63.Write(info63, 0, info63.Length)
        fs63.Close()

        Dim path64 As String = (資料夾位置 & "\" & "MF801" & ".txt") '設定檔名
        Dim fs64 As FileStream = File.Create(path64) '新建或是覆蓋
        Dim info64 As Byte() = New UTF8Encoding(True).GetBytes(NumericUpDown27.Value) '在TXT中加入文字
        fs64.Write(info64, 0, info64.Length)
        fs64.Close()

        Dim path65 As String = (資料夾位置 & "\" & "MF802" & ".txt") '設定檔名
        Dim fs65 As FileStream = File.Create(path65) '新建或是覆蓋
        Dim info65 As Byte() = New UTF8Encoding(True).GetBytes(NumericUpDown28.Value) '在TXT中加入文字
        fs65.Write(info65, 0, info65.Length)
        fs65.Close()

        Dim DATE3 As Integer = Weekday(Today())

        Dim path66 As String = (資料夾位置 & "\" & "DATE3" & ".txt") '設定檔名
        Dim fs66 As FileStream = File.Create(path66) '新建或是覆蓋
        Dim info66 As Byte() = New UTF8Encoding(True).GetBytes(DATE3) '在TXT中加入文字
        fs66.Write(info66, 0, info66.Length)
        fs66.Close()

        Dim path77 As String = (資料夾位置 & "\" & "INFO" & ".txt") '設定檔名
        Dim fs77 As FileStream = File.Create(path77) '新建或是覆蓋
        Dim info77 As Byte() = New UTF8Encoding(True).GetBytes(TextBox34.Text) '在TXT中加入文字
        fs77.Write(info77, 0, info77.Length)
        fs77.Close()

        If (System.IO.File.Exists("vb.zip")) Then '若ZIP檔已存在(暫存資料夾內)
            System.IO.File.Delete("vb.zip") '刪除ZIP檔(暫存資料夾內)
        Else
        End If
        If (System.IO.File.Exists("C:\VB-app\vb.zip")) Then '若ZIP檔已存在(目標資料夾內)
            System.IO.File.Delete("C:\VB-app\vb.zip") '刪除ZIP檔(目標資料夾內)
        Else
        End If
        ZipFile.CreateFromDirectory(資料夾位置, "vb.zip") '建立ZIP檔案(壓縮)
        My.Computer.FileSystem.MoveFile("vb.zip", "C:\VB-app\vb.zip") '移動ZIP至正確資料夾
        System.IO.File.Delete("vb.zip") '刪除暫存資料夾內ZIP檔

        Dim X As Integer
        X = MsgBox(“請確認餐點是否正確” & vbCrLf & "正確無誤請按下[確認]" & vbCrLf & "有誤請按下[取消]進行修改" & vbCrLf & vbCrLf, 1 + 48, "最終確認")
        If X = 1 Then
            MsgBox(“已成功紀錄餐點,謝謝” & vbCrLf & "請手動關閉程式!", 64, "紀錄成功")
            Label46.Text = "狀態:已鎖定" '已鎖定
            Label46.ForeColor = Color.Green '已未鎖定
            Button2.ForeColor = Color.Green '已未鎖定
            PictureBox1.Visible = True '已未鎖定
        Else
            MsgBox(“餐點尚未儲存,請返回修改後重新鎖定”, 64, "紀錄失敗")
            Application.Restart()
        End If
    End Sub
End Class