Public Class howto
    Private Sub howto_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label3.Text = "1.ffmpeg.exeの場所を指定します。通常は起動時のままでOKです。ffmpeg.exeの場所を変えたときはその場所を指定します。"
        Label4.Text = "2.変換元動画の場所を指定します。"
        Label5.Text = "3.変換先動画の保存先は、変換元と同じで、ファイル名は_convertを付加したものになりますが、変えたければ指定します。"
    End Sub
End Class