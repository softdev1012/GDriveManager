Imports System.Configuration
Imports Google.Apis.Auth
Imports Google.Apis.Download
Imports Google.Apis.Drive.v3
Imports Google.Apis.Auth.OAuth2
Imports Google.Apis.Services
Imports Google.Apis.Drive.v3.Data
Imports System.Threading
Imports MySql.Data.MySqlClient
Imports Google.Apis.Util.Store

Public Module VarGlobal
    Public folderNameArr(50) As String
    Public folderIdArr(50) As String
    Public curIndex As Integer
End Module


Public Class frmMain
    Private service As DriveService = New DriveService

    Private Shared Function MimeType(ByVal ToFind As String) As String 'get the mimetype of the type of file in argument
        Dim mot As Array
        Dim index As Integer
        mot = ToFind.Split(".")
        index = mot.Length
        Dim mimetypeRaw As String
        If index > 0 Then
            mimetypeRaw = mot(index - 1)
        Else
            mimetypeRaw = "txt"
        End If
        mimetypeRaw = LCase(mimetypeRaw) 'case if extension is in uppper case
        Dim result As String
        Select Case mimetypeRaw
            Case "avi"
                result = "video/avi"
            Case "bz"
                result = "application/x-bzip"
            Case "bz2"
                result = "application/x-bzip2"
            Case "c", "c++"
                result = "text/plain"
            Case "css"
                result = "text/css"
            Case "doc"
                result = "application/msword"
            Case "gif"
                result = "image/gif"
            Case "html", "htm", "htmls"
                result = "text/html"
            Case "ico"
                result = "image/x-icon"
            Case "mp3"
                result = "audio/mpeg3"
            Case "mp4"
                result = "video/mp4"
            Case "txt"
                result = "text/plain"
            Case "xls", "xlsx", "cvc"
                result = "application/vnd.ms-excel"
            Case "docx"
                result = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            Case "jpeg", "jpg"
                result = "image/jpg"
            Case "png"
                result = "image/png"
            Case "zip"
                result = "application/zip"
            Case "pdf"
                result = "application/pdf"
            Case Else
                result = "text/plain"
        End Select
        Return result 'return the string
    End Function
    Private Sub CreateService()

        'Dim MyUserCredential As UserCredential = GoogleWebAuthorizationBroker.AuthorizeAsync(New ClientSecrets() With {.ClientId = clientid, .ClientSecret = clientsecret}, {DriveService.Scope.Drive}, "user", CancellationToken.None).Result
        'service = New DriveService(New BaseClientService.Initializer() With {.HttpClientInitializer = MyUserCredential, .ApplicationName = "VisualRack"})

        Dim clientSecretJson = "credit.json"
        Dim applicationName = "GDriveManager"

        Try
            ' Permissions
            Dim scopes As String() = New String() {DriveService.Scope.DriveReadonly}
            Dim credential As UserCredential

            Using stream As New System.IO.FileStream(clientSecretJson, System.IO.FileMode.Open, System.IO.FileAccess.Read)
                Dim credPath As String
                credPath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal)
                credPath = System.IO.Path.Combine(credPath, ".credentials/", System.Reflection.Assembly.GetExecutingAssembly.GetName.Name)

                ' Requesting Authentication
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(GoogleClientSecrets.Load(stream).Secrets,
                                                                         scopes,
                                                                         "user",
                                                                         CancellationToken.None,
                                                                         New FileDataStore(credPath, True)).Result
            End Using

            ' Create Drive API service
            service = New DriveService(New BaseClientService.Initializer() With {.HttpClientInitializer = credential, .ApplicationName = applicationName})
        Catch ex As Exception
            MsgBox("Google Drive Authentication is failed. Check credit.json file.")
            End
        End Try
    End Sub
    Private Function AppendFolder(ByVal folder As String, ByVal parent As String)
        CreateService()
        Dim findrequest As FilesResource.ListRequest = service.Files.List
        If parent <> "" Then
            findrequest.Q = "'" + parent + "' in parents"
        End If
        findrequest.PageSize = 1000
        Dim listFolder As FileList = findrequest.Execute
        While (listFolder.Files IsNot Nothing)
            For Each file As File In listFolder.Files
                If file.MimeType = "application/vnd.google-apps.folder" Then 'check only folder
                    If file.Name = folder Then 'check for the folder name then get the ID of it
                        Return file.Id.ToString 'set global variable
                    End If
                End If
            Next
            If (listFolder.NextPageToken Is Nothing) Then
                Exit While
            End If
            findrequest.PageToken = listFolder.NextPageToken
            listFolder = findrequest.Execute
        End While

        Dim folderMetaData As New File()
        Dim plist As New List(Of String)
        folderMetaData.Name = folder
        folderMetaData.MimeType = "application/vnd.google-apps.folder"
        If parent <> "" Then
            plist.Add(parent)
            folderMetaData.Parents = plist
        End If
        Dim req As FilesResource.CreateRequest = service.Files.Create(folderMetaData)
        req.Fields = "id"
        Dim subfolder As File = req.Execute()
        Return subfolder.Id.ToString
    End Function
    Public Sub UploadFile() 'upload a file in a specific folder
        CreateService() 'initialize drive api service
        Dim plist As New List(Of String)
        Dim mimetypeFinal As String = MimeType(lblBrowse.Text)

        Dim path As String = lblBrowse.Text
        Dim dirs() As String = path.Split(System.IO.Path.DirectorySeparatorChar)
        Dim parent As String
        Dim fileName As String = dirs(dirs.Length - 1)
        parent = service.Files.Get("root").Execute.Id
        Dim i As Integer
        If dirs.Length < 3 Then
            Exit Sub
        End If
        For i = 2 To dirs.Length - 2
            parent = AppendFolder(dirs(i), parent)
        Next
        plist.Add(parent)
        If (System.IO.File.Exists(path)) Then 'each file metadata
            Dim fileMetadata
            If parent = String.Empty Then 'case where the file will be upload in root
                fileMetadata = New File() With {.Name = fileName}
            Else
                fileMetadata = New File() With {
                    .Name = fileName, 'get the original file name from the local drive
                    .Parents = plist
                }
            End If
            Dim request As FilesResource.CreateMediaUpload 'prepare HTTP request

            Using stream = New System.IO.FileStream(path, System.IO.FileMode.Open)
                request = service.Files.Create(fileMetadata, stream, mimetypeFinal)
                If dirs.Length < 4 Then
                    request.Fields = "id" 'upload in root, no parents
                Else
                    request.Fields = "id, parents"
                End If
                request.Upload() 'upload function
            End Using

            Dim file As File = request.ResponseBody 'can be used to get the file ID after upload
        End If
        MsgBox("Upload done!")
    End Sub
    Public Sub ListFile()
        If curIndex < 1 Then
            Exit Sub
        End If

        Dim pathStr As String = ""
        Dim i As Integer
        For i = 0 To curIndex - 1
            pathStr += folderNameArr(i) & System.IO.Path.DirectorySeparatorChar
        Next
        txtPath.Text = pathStr

        lboxFile.Items.Clear()
        CreateService()
        Dim findrequest As FilesResource.ListRequest = service.Files.List 'list all the file in the user drive
        findrequest.Q = "'" + folderIdArr(curIndex - 1) + "' in parents"
        findrequest.PageSize = 1000
        Dim listFolder As FileList = findrequest.Execute
        lboxFile.DisplayMember = "Name"
        lboxFile.ValueMember = "Id"
        While ((listFolder.Files) IsNot Nothing)
            For Each file As File In listFolder.Files
                lboxFile.Items.Add(file)
            Next
            If (listFolder.NextPageToken Is Nothing) Then
                Exit While
            End If

            findrequest.PageToken = listFolder.NextPageToken
            listFolder = findrequest.Execute
        End While
        lboxFile.Sorted = True
    End Sub

    Public Sub EmptyTrash()
        CreateService()
        Dim request = service.Files.EmptyTrash
        request.Execute()
        MsgBox("Empty trash")
    End Sub

    Public Sub DownloadFile()
        Dim destinationFolder As String
        Dim fileID As String = String.Empty
        Dim fileName As String = String.Empty
        CreateService()
        If lboxFile.SelectedIndex < 0 Then
            MsgBox("Select file to download.")
            Exit Sub
        End If
        destinationFolder = "C:" & System.IO.Path.DirectorySeparatorChar & "MyDrive" & System.IO.Path.DirectorySeparatorChar
        Dim i As Integer
        For i = 0 To curIndex - 1
            destinationFolder += folderNameArr(i) & System.IO.Path.DirectorySeparatorChar
        Next
        If Not System.IO.Directory.Exists(destinationFolder) Then
            System.IO.Directory.CreateDirectory(destinationFolder)
        End If
        Dim seled = lboxFile.SelectedItem

        If seled.MimeType = "application/vnd.google-apps.folder" Then
            MsgBox("You selected folder. Please select file to download.")
            Exit Sub
        End If

        fileName = destinationFolder & seled.Name
        If Not (seled.Id.ToString() = String.Empty) Then
            Dim Request = service.Files.Get(seled.Id)
            Dim Stream As New IO.FileStream(fileName, System.IO.FileMode.Create, System.IO.FileAccess.ReadWrite)
            Request.Download(Stream)
            MsgBox("Download done!")
            Stream.Close()
        Else
            MsgBox("File not found in Drive")
        End If
    End Sub

    Public Sub DeleteFile()
        CreateService()
        If lboxFile.SelectedIndex < 0 Then
            MsgBox("Select file to delete.")
            Exit Sub
        End If
        Dim seled = lboxFile.SelectedItem
        Dim findrequest As FilesResource.ListRequest = service.Files.List 'list all the file in the user drive
        Dim deleteReq As FilesResource.DeleteRequest = service.Files.Delete(seled.Id)
        deleteReq.Execute()
        MsgBox("File deleted successfully")
    End Sub

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        folderNameArr(0) = ""
        CreateService()
        folderIdArr(0) = service.Files.Get("root").Execute.Id
        curIndex = 1
        ListFile()
    End Sub
    Private Sub btnBrowse_Click(sender As Object, e As EventArgs) Handles btnBrowse.Click
        Dim pathName As String
        Dim dlgResult As DialogResult = dlgOpenFile.ShowDialog()
        If dlgResult = DialogResult.OK Then
            pathName = dlgOpenFile.FileName
            lblBrowse.Text = pathName
        End If
    End Sub

    Private Sub btnEmpty_Click(sender As Object, e As EventArgs) Handles btnEmpty.Click
        emptyTrash()
    End Sub

    Private Sub btnDownload_Click(sender As Object, e As EventArgs) Handles btnDownload.Click
        DownloadFile()
    End Sub
    Private Sub btnUpload_Click(sender As Object, e As EventArgs) Handles btnUpload.Click
        If lblBrowse.Text = "" Or lblBrowse.Text = String.Empty Then
            MsgBox("Select file to upload")
            Exit Sub
        Else
            UploadFile()
            ListFile()
        End If
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        DeleteFile()
        ListFile()
    End Sub
    Private Sub lboxFile_DoubleClick(sender As Object, e As EventArgs) Handles lboxFile.DoubleClick
        If lboxFile.SelectedIndex < 0 Then
            Exit Sub
        End If
        Dim seled = lboxFile.SelectedItem
        If seled.MimeType = "application/vnd.google-apps.folder" Then
            folderIdArr(curIndex) = seled.Id
            folderNameArr(curIndex) = seled.Name
            curIndex += 1
            ListFile()
        End If
    End Sub

    Private Sub btnUp_Click(sender As Object, e As EventArgs) Handles btnUp.Click
        If curIndex < 2 Then
            Exit Sub
        End If
        curIndex -= 1
        ListFile()
    End Sub















    ''''''''''''''''''''''MySQL Part''''''''''''''''''''''''''''

    Private Sub btnConnect_Click(sender As Object, e As EventArgs) Handles btnConnect.Click
        Dim connString As String = "server=localhost;userid=" & txtUsername.Text & ";password=" & txtPassword.Text
        Try
            Dim dtable As New DataTable
            Using conn As New MySqlConnection(connString),
            cmd As New MySqlCommand("show databases", conn)
                conn.Open()
                dtable.Load(cmd.ExecuteReader)
            End Using
            lboxDatabase.Items.Clear()
            For Each row As DataRow In dtable.Rows
                Dim dbName As String = row("database")
                If dbName = "mysql" Or
                    dbName = "information_schema" Or
                    dbName = "performance_schema" Or
                    dbName = "phpmyadmin" Then
                    Continue For
                End If
                lboxDatabase.Items.Add(dbName)
            Next
            lboxDatabase.Sorted = True
        Catch ex As Exception
        End Try
    End Sub

    Private Sub lboxDatabase_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lboxDatabase.SelectedIndexChanged
        Dim dbName As String = lboxDatabase.SelectedItem
        Dim connString As String = "server=localhost;userid=" & txtUsername.Text & ";password=" & txtPassword.Text & ";database=" & dbName & ";convertzerodatetime=true;"
        Try
            Dim dtable As New DataTable
            Using conn As New MySqlConnection(connString),
            cmd As New MySqlCommand("SHOW TABLES;", conn)
                conn.Open()
                dtable.Load(cmd.ExecuteReader)
            End Using
            lboxTable.Items.Clear()
            For Each row As DataRow In dtable.Rows
                Dim tbName As String = row(0)
                lboxTable.Items.Add(tbName)
            Next
            lboxTable.Sorted = True
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click
        If lboxDatabase.SelectedIndex < 0 Then
            MsgBox("Select database.")
            Exit Sub
        End If

        Dim dlgReault = dlgSaveDB.ShowDialog()
        If dlgReault = DialogResult.OK Then
            Dim dbName As String = lboxDatabase.SelectedItem
            Dim tbNames As List(Of String) = New List(Of String)()
            For Each item In lboxTable.CheckedItems
                tbNames.Add(item.ToString)
            Next
            Dim connString As String = "server=localhost;userid=" & txtUsername.Text & ";password=" & txtPassword.Text & ";database=" & dbName & ";convertzerodatetime=true;"
            Dim exportFile As String = dlgSaveDB.FileName

            Using conn As New MySqlConnection(connString),
                    cmd As New MySqlCommand(),
                    mback As New MySqlBackup(cmd)
                cmd.Connection = conn
                conn.Open()
                If tbNames.Count > 0 Then
                    mback.ExportInfo.TablesToBeExportedList = tbNames
                End If
                mback.ExportToFile(exportFile)
                conn.Close()
                MsgBox(dbName & " database is exported to  " & exportFile)
            End Using
        End If
    End Sub

    Private Sub btnImport_Click(sender As Object, e As EventArgs) Handles btnImport.Click
        If lboxDatabase.SelectedIndex < 0 Then
            MsgBox("Select database.")
            Exit Sub
        End If
        Dim dlgResult = dlgLoadDB.ShowDialog()
        If dlgResult = DialogResult.OK Then
            Dim dbName As String = lboxDatabase.SelectedItem
            Dim connString As String = "server=localhost;userid=" & txtUsername.Text & ";password=" & txtPassword.Text & ";database=" & dbName & ";convertzerodatetime=true;"
            Dim exportFile As String = dlgLoadDB.FileName

            Using conn As New MySqlConnection(connString),
                    cmd As New MySqlCommand(),
                    mback As New MySqlBackup(cmd)
                cmd.Connection = conn
                conn.Open()
                mback.ImportFromFile(exportFile)
                conn.Close()
                MsgBox(dbName & " database is imported from  " & exportFile)
            End Using
        End If
    End Sub
    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        End
    End Sub
End Class
