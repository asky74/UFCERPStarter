Public Class Form1
    Public UserName As String, UserPass As String, ini As INIFile

    Private Sub TreeView1_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles TreeView1.MouseDoubleClick
        StartBase()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim SetupIni As String = My.Application.Info.DirectoryPath.ToString & "\setup.ini"

        If IO.File.Exists(SetupIni) = False Then
            MsgBox("Не найден файл настройки: " & SetupIni, MsgBoxStyle.Critical)
            Me.Close()
        End If
        ini = INIFile.Load(SetupIni)
        'ini = INIFile.Load("d:\setup.ini")

        If TypeOf My.User.CurrentPrincipal Is Security.Principal.WindowsPrincipal Then
            ' The application is using Windows authentication.
            ' The name format is DOMAIN\USERNAME.
            Dim parts() As String = Split(My.User.Name, "\")
            UserName = parts(1) & "_" & parts(0)
        Else
            ' The application is using custom authentication.
            UserName = My.User.Name
        End If
        UserPass = StrReverse(ini.Section("1c")("Pass").ToString)
        For Each item In ini.Section("MainList").Keys
            'Dim NameValue As String = item
            'Dim DataValue As String = ini.Section("MainList")(NameValue)
            Dim usrdef As String = ini.Section("MainList")(item).Split((";@").ToCharArray())(0) & "\usrdef\users.usr"
            If IO.File.Exists(usrdef) = False Then
                Continue For
            End If

            Dim allstr = My.Computer.FileSystem.ReadAllText(usrdef, System.Text.Encoding.GetEncoding(1251)).ToString.ToLower
            If InStr(allstr, ",""" & UserName.ToLower & """,") Then
                TreeView1.Nodes.Add(item)
            End If
        Next
    End Sub

    Private Sub TreeView1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TreeView1.KeyPress
        If Asc(e.KeyChar) = 13 Then
            StartBase()
        End If
    End Sub

    Private Sub StartBase()
        Dim baseparam As String = ini.Section("MainList")(TreeView1.SelectedNode.Text)
        Dim spl As String() = baseparam.Split((";@").ToCharArray())
        Dim basedir As String = spl(0)
        Dim exe As String = "s"
        If spl.Length = 2 Then
            exe = spl(1)
        End If

        Dim p As New ProcessStartInfo
        p.FileName = My.Application.Info.DirectoryPath.ToString & "\1cv7" & exe & ".exe"

        ' Use these arguments for the process
        p.Arguments = "enterprise /D""" & basedir & """ /N" & UserName & " /P" & UserPass

        ' Use a hidden window
        'p.WindowStyle = ProcessWindowStyle.Hidden

        ' Start the process
        'MsgBox(p.FileName)

        Process.Start(p)
        Me.Close()
    End Sub
End Class
