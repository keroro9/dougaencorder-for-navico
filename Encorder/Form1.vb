Option Strict On


'--------　下記の位置に名前空間を定義　----------------------
Imports System.Runtime.InteropServices
Public Class Form1


    Inherits System.Windows.Forms.Form

    Dim resollution As String
    Dim encodemethod As String
    Dim bitrate As String
    Private iniFileName As String
    Private ffmpegfilename As String
    '------------------------------------------------------------
    '指定のINIファイルから文字列を取得する(P989)
    <DllImport("KERNEL32.DLL", CharSet:=CharSet.Auto)>
    Public Shared Function GetPrivateProfileString(
       ByVal lpAppName As String,
       ByVal lpKeyName As String,
       ByVal lpDefault As String,
       ByVal lpReturnedString As System.Text.StringBuilder,
       ByVal nSize As Integer,
       ByVal lpFileName As String) As Integer
    End Function
    '指定のINIファイルの指定のキーの文字列を変更する(P994)
    <DllImport("KERNEL32.DLL")>
    Public Shared Function WritePrivateProfileString(
       ByVal lpAppName As String,
       ByVal lpKeyName As String,
       ByVal lpString As String,
       ByVal lpFileName As String) As Integer
    End Function
    '指定のINIファイルから整数値を取得する(P986)
    <DllImport("KERNEL32.DLL", CharSet:=CharSet.Auto)>
    Public Shared Function GetPrivateProfileInt(
       ByVal lpAppName As String,
       ByVal lpKeyName As String,
       ByVal nDefault As Integer,
       ByVal lpFileName As String) As Integer
    End Function


    Private Sub TrackBar1_ValueChanged(ByVal sender As Object,
    ByVal e As System.EventArgs) Handles TrackBar1.ValueChanged
        ' Label3.Text = TrackBar1.Value.ToString + " kbps"
        TextBox5.Text = TrackBar1.Value.ToString
    End Sub

    Private Sub Textbox5_ValueChanged(ByVal sender As Object,
    ByVal e As System.EventArgs) Handles TextBox5.TextChanged
        If TextBox5.Text = "" Then
        Else
            If Integer.Parse(TextBox5.Text) > 499 And Integer.Parse(TextBox5.Text) < 7001 Then
                TrackBar1.Value = Integer.Parse(TextBox5.Text)
            End If
        End If



    End Sub

    Private myLeft, myTop, myWidth, myHeight As Integer, myTxt As String

    Private Sub Button1_Click(ByVal sender As System.Object,
                          ByVal e As System.EventArgs)
        'INI ファイルから読み込み
        '戻り値(文字列)を受け取るバッファーを準備
        Dim strBuffer As New System.Text.StringBuilder
        strBuffer.Capacity = 256   'バッファーのサイズを指定
        Dim ret As Integer
        'INIファイルよりキーの値を読み込み(整数値を取得する場合)
        'myHeight = GetPrivateProfileInt("Form1Size", "Height", 0, iniFileName)
        '文字列の値を取得する場合
        ret = GetPrivateProfileString("ffmpeg", "Text", "",
                                 strBuffer, strBuffer.Capacity, iniFileName)
        myTxt = strBuffer.ToString
        '結果を表示
        'Debug.WriteLine(myHeight)   '135
        Debug.WriteLine(myTxt)      'INIファイルの読み込み・書き込み(75)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim ofd As New OpenFileDialog()

        'はじめのファイル名を指定する
        'はじめに「ファイル名」で表示される文字列を指定する
        ofd.FileName = "ffmpeg.exe"
        'はじめに表示されるフォルダを指定する
        '指定しない（空の文字列）の時は、現在のディレクトリが表示される
        ofd.InitialDirectory = Application.StartupPath
        '[ファイルの種類]に表示される選択肢を指定する
        '指定しないとすべてのファイルが表示される
        ofd.Filter = "exeファイル(*.exe)|*.exe;|すべてのファイル(*.*)|*.*"
        '[ファイルの種類]ではじめに選択されるものを指定する
        '2番目の「すべてのファイル」が選択されているようにする
        ofd.FilterIndex = 2
        'タイトルを設定する
        ofd.Title = "FFmpegの場所を選択"
        'ダイアログボックスを閉じる前に現在のディレクトリを復元するようにする
        ofd.RestoreDirectory = True
        '存在しないファイルの名前が指定されたとき警告を表示する
        'デフォルトでTrueなので指定する必要はない
        ofd.CheckFileExists = True
        '存在しないパスが指定されたとき警告を表示する
        'デフォルトでTrueなので指定する必要はない
        ofd.CheckPathExists = True

        'ダイアログを表示する
        If ofd.ShowDialog() = DialogResult.OK Then
            'OKボタンがクリックされたとき、選択されたファイル名を表示する
            Console.WriteLine(ofd.FileName)
        End If
        If Not ofd.FileName = "ffmpeg.exe" Then
            TextBox1.Text = ofd.FileName
        Else
            TextBox1.Text = TextBox1.Text
        End If


    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object,
                          ByVal e As System.EventArgs)
        'INI ファイルに各キーに対する値を取得して書き込み
        Dim ret As Integer
        '書き込みは文字列も整数値(文字列に変換して)も同じ方法で
        ' ret = WritePrivateProfileString("ffmpeg", "Text", "c:\", iniFileName)

        'ret = WritePrivateProfileString("Form1Size", "Height", CStr(Me.Height), iniFileName)

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim ofd As New OpenFileDialog()
        Dim pos As Integer
        Dim PathName As String
        Dim FileName As String
        Dim outfile As String
        Dim addname As String

        'はじめのファイル名を指定する
        'はじめに「ファイル名」で表示される文字列を指定する
        ofd.FileName = ""
        'はじめに表示されるフォルダを指定する
        '指定しない（空の文字列）の時は、現在のディレクトリが表示される
        ofd.InitialDirectory = Application.StartupPath
        '[ファイルの種類]に表示される選択肢を指定する
        '指定しないとすべてのファイルが表示される
        ofd.Filter = "すべてのファイル(*.*)|*.*"
        '[ファイルの種類]ではじめに選択されるものを指定する
        '2番目の「すべてのファイル」が選択されているようにする
        ofd.FilterIndex = 2
        'タイトルを設定する
        ofd.Title = "入力ファイルの場所を選択"
        'ダイアログボックスを閉じる前に現在のディレクトリを復元するようにする
        ofd.RestoreDirectory = True
        '存在しないファイルの名前が指定されたとき警告を表示する
        'デフォルトでTrueなので指定する必要はない
        ofd.CheckFileExists = True
        '存在しないパスが指定されたとき警告を表示する
        'デフォルトでTrueなので指定する必要はない
        ofd.CheckPathExists = True
        '複数選択を許可
        ofd.Multiselect = True


        'ダイアログを表示する
        If ofd.ShowDialog() = DialogResult.OK Then
            'OKボタンがクリックされたとき、選択されたファイル名を表示する
            Console.WriteLine(ofd.FileName)
            'リストボックスの初期化
            '  ListBox1.Items.Clear()

            '選択されたファイルをテキストボックスに表示する
            For Each strFilePath As String In ofd.FileNames
                ''ファイルパスからファイル名を取得
                'Dim strFileName As String = IO.Path.GetFileName(strFilePath)

                'If Not strFilePath = "" Then
                '    '入力ファイル名に_convertを付加し、出力ファイルパスに指定
                '    pos = CInt(CType(InStrRev(strFilePath, "\"), String))
                '    PathName = VBStrings.Left(strFilePath, pos)
                '    FileName = VBStrings.Mid(strFilePath, pos + 1)

                '    pos = CInt(CType(InStrRev(FileName, "."), String))

                '    FileName = VBStrings.Left(FileName, pos - 1)
                '    FileName = FileName + "_convert.mp4"
                '    outfile = PathName + FileName
                '    strFileName = PathName + FileName
                'End If
                'リストボックスにファイル名を表示
                ListBox1.Items.Add(strFilePath)
            Next
        End If


        'TextBox2.Text = ofd.FileName
        'If Not TextBox2.Text = "" Then
        '    '入力ファイル名に_convertを付加し、出力ファイルパスに指定
        '    pos = CInt(CType(InStrRev(TextBox2.Text, "\"), String))
        '    PathName = VBStrings.Left(TextBox2.Text, pos)
        '    FileName = VBStrings.Mid(TextBox2.Text, pos + 1)

        '    pos = CInt(CType(InStrRev(FileName, "."), String))

        '    FileName = VBStrings.Left(FileName, pos - 1)
        '    FileName = FileName + "_convert.mp4"
        '    outfile = PathName + FileName
        '    TextBox3.Text = PathName + FileName
        'End If


    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs)
        Dim ofd As New OpenFileDialog()

        'はじめのファイル名を指定する
        'はじめに「ファイル名」で表示される文字列を指定する
        ofd.FileName = ""
        'はじめに表示されるフォルダを指定する
        '指定しない（空の文字列）の時は、現在のディレクトリが表示される
        ofd.InitialDirectory = Application.StartupPath
        '[ファイルの種類]に表示される選択肢を指定する
        '指定しないとすべてのファイルが表示される
        ofd.Filter = "すべてのファイル(*.*)|*.*"
        '[ファイルの種類]ではじめに選択されるものを指定する
        '2番目の「すべてのファイル」が選択されているようにする
        ofd.FilterIndex = 2
        'タイトルを設定する
        ofd.Title = "出力ファイルの場所を選択"
        'ダイアログボックスを閉じる前に現在のディレクトリを復元するようにする
        ofd.RestoreDirectory = True
        '存在しないファイルの名前が指定されたとき警告を表示する
        'デフォルトでTrueなので指定する必要はない
        ofd.CheckFileExists = True
        '存在しないパスが指定されたとき警告を表示する
        'デフォルトでTrueなので指定する必要はない
        ofd.CheckPathExists = True

        'ダイアログを表示する
        If ofd.ShowDialog() = DialogResult.OK Then
            'OKボタンがクリックされたとき、選択されたファイル名を表示する
            Console.WriteLine(ofd.FileName)
        End If


    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim outputfilenametale As String
        Dim indir As String
        'GOボタンをクリックしたときの処理
        If TextBox3.Text = "" Then
            Dim result As DialogResult = MessageBox.Show("変換元と同一ディレクトリに変換されますがいいですか？",
                                             "エンコード開始します",
                                             MessageBoxButtons.YesNo,
                                             MessageBoxIcon.Exclamation,
                                             MessageBoxDefaultButton.Button2)
            If result = DialogResult.Yes Then
            ElseIf result = DialogResult.No Then
                GoTo ExitLoop
            End If
        End If

        Dim core As String

        Dim time As String
        Dim ss As String
        core = Environment.ProcessorCount.ToString()
        If CheckBox1.Checked = True Then
            core = " -threads " + core + " "
            Debug.WriteLine(core)
        Else
            core = ""
        End If
        If RadioButton1.Checked = True Then
            resollution = "640x360"
            Debug.WriteLine(resollution)
        ElseIf RadioButton2.Checked = True Then
            resollution = "800x450"
        ElseIf RadioButton3.Checked = True Then
            resollution = "1280x720"
        ElseIf RadioButton4.Checked = True Then
            resollution = "1440x810"
        ElseIf RadioButton5.Checked = True Then
            resollution = "1660x900"
        ElseIf RadioButton6.Checked = True Then
            resollution = "1920x1080"
        ElseIf RadioButton6.Checked = True Then
            resollution = TextBox6.Text + "x" + TextBox7.Text
        End If
        'ビットレート設定
        bitrate = TrackBar1.Value.ToString + "k"

        'エンコード形式の設定
        If RadioButton9.Checked = True Then
            encodemethod = " -c:v libx264 -profile:v main "
        ElseIf RadioButton10.Checked = True Then
            encodemethod = " -vcodec mpeg4 "
        End If
        '切り取り時間の設定
        If CheckBox2.Checked = True Then
            If Not TextBox4.Text = "" Then
                time = " -t" + " " + TextBox4.Text + " "
                ss = " -ss" + " " + TextBox2.Text + " "
            Else
                MsgBox("時間を入力してください")
                GoTo ExitLoop
            End If
        Else
            time = ""
            ss = ""
        End If
        If TextBox1.Text = "" Then
            MsgBox("ffmpegの場所が指定されていません。処理が続行できません。")
            GoTo ExitLoop
        End If
        'アスペクト比設定
        Dim aspect As String
        If RadioButton8.Checked = True Then
            aspect = " -aspect 16:9 "
        ElseIf RadioButton11.Checked = True Then
            aspect = " -aspect 4:3 "
        End If
        'ffmpegエンコードオプションの生成
        Dim encord_opt As String




        'outputfilenametaleの生成
        'ファイル名の生成
        If CheckBox6.Checked = True Then
            outputfilenametale = TextBox8.Text
        Else
            outputfilenametale = outputfilenametale + ""
        End If
        If CheckBox4.Checked = True Then
            If RadioButton9.Checked = True Or RadioButton10.Checked = True Then
                If RadioButton1.Checked = True Then
                    outputfilenametale = outputfilenametale + "_640x360"
                ElseIf RadioButton2.Checked = True Then
                    outputfilenametale = outputfilenametale + "_800x450"
                ElseIf RadioButton3.Checked = True Then
                    outputfilenametale = outputfilenametale + "_1280x720"
                ElseIf RadioButton4.Checked = True Then
                    outputfilenametale = outputfilenametale + "_1440x810"
                ElseIf RadioButton5.Checked = True Then
                    outputfilenametale = outputfilenametale + "_1600x900"
                ElseIf RadioButton6.Checked = True Then
                    outputfilenametale = outputfilenametale + "_1920x1080"
                ElseIf RadioButton7.Checked = True And
                        outputfilenametale = outputfilenametale + "_" + TextBox6.Text + "x" + TextBox7.Text Then
                End If

                If RadioButton9.Checked = True Then
                    outputfilenametale = outputfilenametale + "_h264"
                ElseIf RadioButton10.Checked = True Then
                    outputfilenametale = outputfilenametale + "_mp4"
                End If

                If RadioButton8.Checked = True Then
                    outputfilenametale = outputfilenametale + "_169"
                ElseIf RadioButton11.Checked = True Then
                    outputfilenametale = outputfilenametale + "_43"
                End If

                If CheckBox2.Checked = True Then
                    outputfilenametale = outputfilenametale + "_t" + TextBox2.Text + "-" + TextBox4.Text
                End If
            End If
            outputfilenametale = outputfilenametale + "_b" + TextBox5.Text
        End If





        'ループ処理
        For Each outfilepath As String In ListBox1.Items
            Dim pos As Integer
            Dim PathName As String
            Dim FileName As String
            Dim infilepath As String
            infilepath = outfilepath
            'TextBox2.Text = ofd.FileName

            '    '入力ファイル名に_convertを付加し、出力ファイルパスに指定
            pos = CInt(CType(InStrRev(outfilepath, "\"), String))
            PathName = VBStrings.Left(outfilepath, pos)
            FileName = VBStrings.Mid(outfilepath, pos + 1)

            pos = CInt(CType(InStrRev(FileName, "."), String))
            If Not CheckBox3.Checked = True Then
                PathName = TextBox3.Text
            End If
            FileName = VBStrings.Left(FileName, pos - 1)





            indir = IO.Path.GetFileName(IO.Path.GetDirectoryName(infilepath))
            If CheckBox5.Checked Then
                If CheckBox3.Checked Then
                    If System.IO.Directory.Exists(PathName + indir) Then
                        PathName = PathName + indir + "\"
                    Else
                        'MessageBox.Show(PathName + indir + "は存在しない")
                        System.IO.Directory.CreateDirectory(PathName + indir)
                        PathName = PathName + indir + "\"
                    End If
                Else
                    If Not TextBox3.Text = "" Then
                        If System.IO.Directory.Exists(TextBox3.Text + "\" + indir) Then
                            PathName = TextBox3.Text + indir + "\"
                        Else
                            System.IO.Directory.CreateDirectory(TextBox3.Text + "\" + indir)
                            PathName = TextBox3.Text + indir + "\"
                        End If
                    End If

                End If
            End If





            If RadioButton9.Checked = True Or RadioButton10.Checked = True Then
                FileName = FileName + outputfilenametale + ".mp4"
            ElseIf RadioButton12.Checked = True Then
                FileName = FileName + outputfilenametale + ".mp3"
            ElseIf RadioButton13.Checked = True Then
                FileName = FileName + outputfilenametale + ".flac"
            End If
            ' FileName = FileName + "_convert.mp4"
            outfilepath = PathName + FileName

            ' MsgBox(outfilepath)

            encord_opt = "/c" + Chr(34) + Chr(34) + TextBox1.Text + Chr(34) + " -y -i " + """" + infilepath + """" + encodemethod + " -s " + resollution + " -b:v " + bitrate + aspect + time + ss + core + " " + """" + outfilepath + """" + Chr(34)

            'If RadioButton10.Checked = True Then
            '    encord_opt = "/c" + TextBox1.Text + " -y -i " + TextBox2.Text + " -vcodec mpeg4 -s " + resollution + " -b:v " + bitrate + " -aspect 16:9 " + core + TextBox3.Text
            'ElseIf RadioButton9.Checked = True Then
            '    encord_opt = "/c" + TextBox1.Text + " -y -i " + TextBox2.Text + " -c:v libx264 -profile:v main" + " -s " + resollution + " -b:v " + bitrate + " -aspect 16:9 " + core + TextBox3.Text
            'End If

            Debug.WriteLine(encord_opt)





            'Processオブジェクトを作成
            Dim p As New System.Diagnostics.Process()

            'ComSpec(cmd.exe)のパスを取得して、FileNameプロパティに指定
            p.StartInfo.FileName = System.Environment.GetEnvironmentVariable("ComSpec")
            '出力を読み取れるようにする
            p.StartInfo.UseShellExecute = False
            p.StartInfo.RedirectStandardOutput = True
            p.StartInfo.RedirectStandardInput = False
            'ウィンドウを表示しないようにする
            p.StartInfo.CreateNoWindow = False
            'コマンドラインを指定（"/c"は実行後閉じるために必要）
            p.StartInfo.Arguments = encord_opt

            '起動
            p.Start()

            '出力を読み取る
            Dim results As String = p.StandardOutput.ReadToEnd()

            'プロセス終了まで待機する
            'WaitForExitはReadToEndの後である必要がある
            '(親プロセス、子プロセスでブロック防止のため)
            p.WaitForExit()
            p.Close()

            '出力された結果を表示
            Console.WriteLine(results)
        Next

ExitLoop:

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs)
        Dim f As New howto()
        f.Show()
    End Sub

    Private Sub presetbutton_Click(sender As Object, e As EventArgs) Handles presetbutton.Click
        iniFileName = "settings.ini"  'INIファイル名（直接フルパスを指定してもOK）
        ffmpegfilename = "ffmpeg.exe"
        'INI ファイルをプログラムと同じフォルダーに置く場合
        'ルートディレクトリかの判断
        Dim MyPath As String = Application.StartupPath
        If MyPath.EndsWith("\") = False Then
            MyPath &= "\"
        End If
        iniFileName = MyPath & iniFileName
        'INI ファイルから読み込み
        '戻り値(文字列)を受け取るバッファーを準備
        Dim strBuffer As New System.Text.StringBuilder
        strBuffer.Capacity = 256   'バッファーのサイズを指定
        Dim ret As Integer
        ret = GetPrivateProfileString("preset01", "resolution", "",
                                 strBuffer, strBuffer.Capacity, iniFileName)
        myTxt = strBuffer.ToString
        '結果を表示
        'Debug.WriteLine(myHeight)   '135
        Debug.WriteLine(myTxt)      'INIファイルの読み込み・書き込み(75)
        If myTxt = "640x360" Then
            RadioButton1.Checked = True
        ElseIf myTxt = "800x450" Then
            RadioButton2.Checked = True
        ElseIf myTxt = "1440x810" Then
            RadioButton4.Checked = True
        ElseIf myTxt = "1280x720" Then
            RadioButton3.Checked = True
        ElseIf myTxt = "1600x900" Then
            RadioButton5.Checked = True
        ElseIf myTxt = "1920x1080" Then
            RadioButton6.Checked = True
        End If
        ret = GetPrivateProfileString("preset01", "encordmethod", "",
                                 strBuffer, strBuffer.Capacity, iniFileName)
        myTxt = strBuffer.ToString
        If myTxt = "h264" Then
            RadioButton9.Checked = True
        ElseIf myTxt = "mp4" Then
            RadioButton10.Checked = True
        End If
        'ビットレート設定
        ret = GetPrivateProfileString("preset01", "bitrate", "",
                                 strBuffer, strBuffer.Capacity, iniFileName)
        myTxt = strBuffer.ToString
        TrackBar1.Value = Integer.Parse(myTxt)
    End Sub

    Private Sub preset2button_Click(sender As Object, e As EventArgs) Handles preset2button.Click
        iniFileName = "settings.ini"  'INIファイル名（直接フルパスを指定してもOK）
        ffmpegfilename = "ffmpeg.exe"
        'INI ファイルをプログラムと同じフォルダーに置く場合
        'ルートディレクトリかの判断
        Dim MyPath As String = Application.StartupPath
        If MyPath.EndsWith("\") = False Then
            MyPath &= "\"
        End If
        iniFileName = MyPath & iniFileName
        'INI ファイルから読み込み
        '戻り値(文字列)を受け取るバッファーを準備
        Dim strBuffer As New System.Text.StringBuilder
        strBuffer.Capacity = 256   'バッファーのサイズを指定
        Dim ret As Integer
        ret = GetPrivateProfileString("preset02", "resolution", "",
                                 strBuffer, strBuffer.Capacity, iniFileName)
        myTxt = strBuffer.ToString
        '結果を表示
        'Debug.WriteLine(myHeight)   '135
        Debug.WriteLine(myTxt)      'INIファイルの読み込み・書き込み(75)
        If myTxt = "640x360" Then
            RadioButton1.Checked = True
        ElseIf myTxt = "800x450" Then
            RadioButton2.Checked = True
        ElseIf myTxt = "1440x810" Then
            RadioButton4.Checked = True
        ElseIf myTxt = "1280x720" Then
            RadioButton3.Checked = True
        ElseIf myTxt = "1600x900" Then
            RadioButton5.Checked = True
        ElseIf myTxt = "1920x1080" Then
            RadioButton6.Checked = True
        End If
        ret = GetPrivateProfileString("preset02", "encordmethod", "",
                                 strBuffer, strBuffer.Capacity, iniFileName)
        myTxt = strBuffer.ToString
        If myTxt = "h264" Then
            RadioButton9.Checked = True
        ElseIf myTxt = "mp4" Then
            RadioButton10.Checked = True
        End If
        'ビットレート設定
        ret = GetPrivateProfileString("preset02", "bitrate", "",
                                 strBuffer, strBuffer.Capacity, iniFileName)
        myTxt = strBuffer.ToString
        TrackBar1.Value = Integer.Parse(myTxt)
    End Sub

    Private Sub preset3button_Click(sender As Object, e As EventArgs) Handles preset3button.Click
        iniFileName = "settings.ini"  'INIファイル名（直接フルパスを指定してもOK）
        ffmpegfilename = "ffmpeg.exe"
        'INI ファイルをプログラムと同じフォルダーに置く場合
        'ルートディレクトリかの判断
        Dim MyPath As String = Application.StartupPath
        If MyPath.EndsWith("\") = False Then
            MyPath &= "\"
        End If
        iniFileName = MyPath & iniFileName
        'INI ファイルから読み込み
        '戻り値(文字列)を受け取るバッファーを準備
        Dim strBuffer As New System.Text.StringBuilder
        strBuffer.Capacity = 256   'バッファーのサイズを指定
        Dim ret As Integer
        ret = GetPrivateProfileString("preset03", "resolution", "",
                                 strBuffer, strBuffer.Capacity, iniFileName)
        myTxt = strBuffer.ToString
        '結果を表示
        'Debug.WriteLine(myHeight)   '135
        Debug.WriteLine(myTxt)      'INIファイルの読み込み・書き込み(75)
        If myTxt = "640x360" Then
            RadioButton1.Checked = True
        ElseIf myTxt = "800x450" Then
            RadioButton2.Checked = True
        ElseIf myTxt = "1440x810" Then
            RadioButton4.Checked = True
        ElseIf myTxt = "1280x720" Then
            RadioButton3.Checked = True
        ElseIf myTxt = "1600x900" Then
            RadioButton5.Checked = True
        ElseIf myTxt = "1920x1080" Then
            RadioButton6.Checked = True
        End If
        ret = GetPrivateProfileString("preset03", "encordmethod", "",
                                 strBuffer, strBuffer.Capacity, iniFileName)
        myTxt = strBuffer.ToString
        If myTxt = "h264" Then
            RadioButton9.Checked = True
        ElseIf myTxt = "mp4" Then
            RadioButton10.Checked = True
        End If
        'ビットレート設定
        ret = GetPrivateProfileString("preset03", "bitrate", "",
                                 strBuffer, strBuffer.Capacity, iniFileName)
        myTxt = strBuffer.ToString
        TrackBar1.Value = Integer.Parse(myTxt)
    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        CheckBox3.Checked = False
        Dim fbd As New FolderBrowserDialog
        '上部に表示する説明テキストを指定する
        fbd.Description = "フォルダを指定してください。"
        'ルートフォルダを指定する
        'デフォルトでDesktop
        fbd.RootFolder = Environment.SpecialFolder.Desktop
        '最初に選択するフォルダを指定する
        'RootFolder以下にあるフォルダである必要がある
        fbd.SelectedPath = "C:\Windows"
        'ユーザーが新しいフォルダを作成できるようにする
        'デフォルトでTrue
        fbd.ShowNewFolderButton = True

        'ダイアログを表示する
        If fbd.ShowDialog(Me) = DialogResult.OK Then
            '選択されたフォルダを表示する
            Console.WriteLine(fbd.SelectedPath)
            TextBox3.Text = fbd.SelectedPath + "\"
        End If
    End Sub

    Private Sub Button5_Click_1(sender As Object, e As EventArgs) Handles Button5.Click
        ListBox1.Items.Clear()
    End Sub

    Private Sub RadioButton12_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton12.CheckedChanged
        If RadioButton12.Checked = True Then
            TrackBar1.Minimum = 64
            TrackBar1.Maximum = 360
            TextBox5.Text = "256"
            TrackBar1.Value = 256
            CheckBox5.Checked = True

        Else
            TrackBar1.Minimum = 500
            TrackBar1.Maximum = 7000
            TextBox5.Text = "800"
            TrackBar1.Value = 800
            CheckBox5.Checked = False
        End If
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs)
        Console.WriteLine("CPU 数 : " + Environment.ProcessorCount.ToString())
    End Sub

    Private Sub Button1_Click_2(sender As Object, e As EventArgs) Handles Button1.Click
        '  RadioButton6.Checked = True
        CheckBox1.Checked = True
        ' Label4.Text = "このPCの論理コア数: " + Environment.ProcessorCount.ToString()
        RadioButton9.Checked = True
        TrackBar1.Value = 2000
        ListBox1.Items.Clear()
        CheckBox3.Checked = True
        RadioButton8.Checked = True
        RadioButton4.Checked = True
    End Sub




    Private Sub Form1_Load(ByVal sender As System.Object,
                           ByVal e As System.EventArgs) Handles MyBase.Load
        iniFileName = "settings.ini"  'INIファイル名（直接フルパスを指定してもOK）
        ffmpegfilename = "ffmpeg.exe"

        'INI ファイルをプログラムと同じフォルダーに置く場合
        'ルートディレクトリかの判断
        Dim MyPath As String = Application.StartupPath
        If MyPath.EndsWith("\") = False Then
            MyPath &= "\"
        End If
        iniFileName = MyPath & iniFileName
        ffmpegfilename = MyPath & ffmpegfilename
        'TextBox1にffmpegの既定パスを入力
        TextBox1.Text = ffmpegfilename


        'INI ファイルから読み込み
        '戻り値(文字列)を受け取るバッファーを準備
        Dim strBuffer As New System.Text.StringBuilder
        strBuffer.Capacity = 256   'バッファーのサイズを指定
        Dim ret As Integer
        'INIファイルよりキーの値を読み込み(整数値を取得する場合)
        'myHeight = GetPrivateProfileInt("Form1Size", "Height", 0, iniFileName)
        '文字列の値を取得する場合
        ret = GetPrivateProfileString("ffmpeg", "Text", "",
                                 strBuffer, strBuffer.Capacity, iniFileName)
        myTxt = strBuffer.ToString
        '結果を表示
        'Debug.WriteLine(myHeight)   '135
        Debug.WriteLine(myTxt)      'INIファイルの読み込み・書き込み(75)

        'iniファイルが空の時の処理
        If myTxt = "" Then
            TextBox1.Text = ffmpegfilename
        ElseIf Not ffmpegfilename = myTxt Then

            If System.IO.File.Exists(myTxt) = True Then
                TextBox1.Text = myTxt
            Else
                MsgBox("設定ファイルの場所にffmpeg.exeがありません。初期値に戻します。")
                TextBox1.Text = ffmpegfilename
            End If
        Else
        End If

        If System.IO.File.Exists(TextBox1.Text) = False Then
            MsgBox("設定ファイルの場所にffmpeg.exeがありません。初期値に戻します。")
            TextBox1.Text = ffmpegfilename
        End If
        '  RadioButton6.Checked = True
        CheckBox1.Checked = True
        ' Label4.Text = "このPCの論理コア数: " + Environment.ProcessorCount.ToString()
        RadioButton9.Checked = True
        TrackBar1.Value = 2000
        ListBox1.Items.Clear()
        CheckBox3.Checked = True
        RadioButton8.Checked = True
        RadioButton4.Checked = True

        ret = GetPrivateProfileString("preset01", "button", "",
                                 strBuffer, strBuffer.Capacity, iniFileName)
        myTxt = strBuffer.ToString
        presetbutton.Text = myTxt

        ret = GetPrivateProfileString("preset02", "button", "",
                                 strBuffer, strBuffer.Capacity, iniFileName)
        myTxt = strBuffer.ToString
        preset2button.Text = myTxt

        ret = GetPrivateProfileString("preset03", "button", "",
                                 strBuffer, strBuffer.Capacity, iniFileName)
        myTxt = strBuffer.ToString
        preset3button.Text = myTxt





        CheckBox4.Checked = True
        TextBox8.Text = "_convert"
        CheckBox6.Checked = True
    End Sub

    Private Sub Button7_Click_1(sender As Object, e As EventArgs) Handles Button7.Click
        Call deletelistitem()
    End Sub

    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'INI ファイルに各キーに対する値を取得して書き込み
        Dim ret As Integer
        '書き込みは文字列も整数値(文字列に変換して)も同じ方法で
        ret = WritePrivateProfileString("ffmpeg", "Text", TextBox1.Text, iniFileName)
        'ret = WritePrivateProfileString("Form1Size", "Height", CStr(Me.Height), iniFileName)
    End Sub


    Private Sub ListBox1_DragEnter(ByVal sender As Object,
        ByVal e As System.Windows.Forms.DragEventArgs) _
        Handles ListBox1.DragEnter
        'コントロール内にドラッグされたとき実行される
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            'ドラッグされたデータ形式を調べ、ファイルのときはコピーとする
            e.Effect = DragDropEffects.Copy
        Else
            'ファイル以外は受け付けない
            e.Effect = DragDropEffects.None
        End If
    End Sub

    Private Sub ListBox1_DragDrop(ByVal sender As Object,
            ByVal e As System.Windows.Forms.DragEventArgs) _
            Handles ListBox1.DragDrop
        'コントロール内にドロップされたとき実行される
        'ドロップされたすべてのファイル名を取得する
        Dim fileName As String() = CType(
            e.Data.GetData(DataFormats.FileDrop, False),
            String())
        'ListBoxに追加する
        ListBox1.Items.AddRange(fileName)
    End Sub

    Private Sub deletelistitem()
        Dim i As Integer
        i = 0
        Do While i < ListBox1.Items.Count
            If ListBox1.GetSelected(i) = True Then
                Me.ListBox1.Items.RemoveAt(i)
            Else
                i = i + 1
            End If
        Loop
        'Me.ListBox1.Items.RemoveAt(0)
    End Sub
End Class


Public Class VBStrings

#Region "　Left メソッド　"

    ''' -----------------------------------------------------------------------------------
    ''' <summary>
    '''     文字列の左端から指定された文字数分の文字列を返します。</summary>
    ''' <param name="stTarget">
    '''     取り出す元になる文字列。</param>
    ''' <param name="iLength">
    '''     取り出す文字数。</param>
    ''' <returns>
    '''     左端から指定された文字数分の文字列。
    '''     文字数を超えた場合は、文字列全体が返されます。</returns>
    ''' -----------------------------------------------------------------------------------
    Public Shared Function Left(ByVal stTarget As String, ByVal iLength As Integer) As String
        If iLength <= stTarget.Length Then
            Return stTarget.Substring(0, iLength)
        End If

        Return stTarget
    End Function

#End Region

#Region "　Mid メソッド (+1)　"

    ''' -----------------------------------------------------------------------------------
    ''' <summary>
    '''     文字列の指定された位置以降のすべての文字列を返します。</summary>
    ''' <param name="stTarget">
    '''     取り出す元になる文字列。</param>
    ''' <param name="iStart">
    '''     取り出しを開始する位置。</param>
    ''' <returns>
    '''     指定された位置以降のすべての文字列。</returns>
    ''' -----------------------------------------------------------------------------------
    Public Overloads Shared Function Mid(ByVal stTarget As String, ByVal iStart As Integer) As String
        If iStart <= stTarget.Length Then
            Return stTarget.Substring(iStart - 1)
        End If

        Return String.Empty
    End Function

    ''' -----------------------------------------------------------------------------------
    ''' <summary>
    '''     文字列の指定された位置から、指定された文字数分の文字列を返します。</summary>
    ''' <param name="stTarget">
    '''     取り出す元になる文字列。</param>
    ''' <param name="iStart">
    '''     取り出しを開始する位置。</param>
    ''' <param name="iLength">
    '''     取り出す文字数。</param>
    ''' <returns>
    '''     指定された位置から指定された文字数分の文字列。
    '''     文字数を超えた場合は、指定された位置からすべての文字列が返されます。</returns>
    ''' -----------------------------------------------------------------------------------
    Public Overloads Shared Function Mid(ByVal stTarget As String, ByVal iStart As Integer, ByVal iLength As Integer) As String
        If iStart <= stTarget.Length Then
            If iStart + iLength - 1 <= stTarget.Length Then
                Return stTarget.Substring(iStart - 1, iLength)
            End If

            Return stTarget.Substring(iStart - 1)
        End If

        Return String.Empty
    End Function

#End Region

#Region "　Right メソッド　"

    ''' -----------------------------------------------------------------------------------
    ''' <summary>
    '''     文字列の右端から指定された文字数分の文字列を返します。</summary>
    ''' <param name="stTarget">
    '''     取り出す元になる文字列。</param>
    ''' <param name="iLength">
    '''     取り出す文字数。</param>
    ''' <returns>
    '''     右端から指定された文字数分の文字列。
    '''     文字数を超えた場合は、文字列全体が返されます。</returns>
    ''' -----------------------------------------------------------------------------------
    Public Shared Function Right(ByVal stTarget As String, ByVal iLength As Integer) As String
        If iLength <= stTarget.Length Then
            Return stTarget.Substring(stTarget.Length - iLength)
        End If

        Return stTarget
    End Function

#End Region

End Class