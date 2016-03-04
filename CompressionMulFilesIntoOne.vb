Imports System.IO
Imports System.IO.Compression
Imports System.IO.Packaging
Imports System.Uri

' 此程式參考網址:
' http://www.emoreau.com/Entries/Articles/2008/08/Introducing-SystemIOPackaging.aspx

Public Class Form1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim errCode As Integer = CompressFilesIntoOne(txtSourcePath.Text, txtDestPath.Text, txtFolderName.Text)

        If errCode = 1 Then
            MessageBox.Show("Source Path is not exist.", "Error occured", MessageBoxButtons.OK, MessageBoxIcon.Error)
        ElseIf errCode = 10 Then
            MessageBox.Show("Add Filename extension error or destination path error", "Error occured", MessageBoxButtons.OK, MessageBoxIcon.Error)
        ElseIf errCode = 100 Then
            MessageBox.Show("Compression error", "Error occured", MessageBoxButtons.OK, MessageBoxIcon.Error)
        ElseIf errCode = 0 Then
            MessageBox.Show("Zip completed!")
        End If

    End Sub

    Private Function CompressFilesIntoOne(ByVal SourcePath As String, ByVal DestinationPath As String, ByVal ZipFileName As String) As Integer
        Dim tempFullDestPath As String = "temp"
        Dim FullDestPath As String = "temp"
        Dim objZip As Package
        Dim di As DirectoryInfo = New DirectoryInfo(SourcePath)
        Dim errCode As Integer
        Dim DesFileExist As Integer

        errCode = 0
        DesFileExist = 0

        ' Check Source Path
        If Not My.Computer.FileSystem.DirectoryExists(SourcePath) Then
            errCode = 1
        End If

        If errCode = 0 Then
            Try
                If (Not ZipFileName.Contains(".zip")) Then
                    ZipFileName = ZipFileName & ".zip"
                End If

                FullDestPath = (DestinationPath & "\" & ZipFileName)

                ' TODO: Sometime, in 'path' will have double '/'. 
                ' I'm not sure it will have a bug or not.
                If (File.Exists(FullDestPath)) Then
                    tempFullDestPath = (DestinationPath & "\" & "_temp" & ZipFileName)
                    DesFileExist = 1
                    objZip = ZipPackage.Open(tempFullDestPath, FileMode.OpenOrCreate, FileAccess.ReadWrite)
                Else
                    objZip = ZipPackage.Open(FullDestPath, FileMode.OpenOrCreate, FileAccess.ReadWrite)
                End If
            Catch ex As Exception
                errCode = 10
            End Try
        End If



        Try
            If errCode = 0 Then
                For Each fi As FileInfo In di.GetFiles()
                    Using inFile As FileStream = fi.OpenRead()
                        If (File.Exists(fi.FullName)) Then
                            If (File.GetAttributes(fi.FullName) And (FileAttributes.Hidden)) <> FileAttributes.Hidden Then '  確認來源路徑下的檔案存在而且非隱藏
                                ' 把檔案讀近來並且用AddFileToZip加入
                                AddFileToZip(objZip, fi.FullName)
                            End If
                        End If
                    End Using
                Next
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error occured", MessageBoxButtons.OK, MessageBoxIcon.Error)
            errCode = 100
        End Try

        'Close the file
        objZip.Close()
        objZip = Nothing

        If DesFileExist = 1 And errCode = 0 Then
            Try
                File.Copy(tempFullDestPath, FullDestPath, True)
                File.Delete(tempFullDestPath)
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Error occured", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

        End If

        Return errCode

    End Function

    Private Sub AddFileToZip(ByVal pZip As Package, ByVal pFileToAdd As String)

        'Create a URI from the filename to zip (to ensure the name is valid)
        Dim partURI As New Uri(CreateUriFromFilename(pFileToAdd), UriKind.Relative)
        'Create a Package Part
        Dim pkgPart As PackagePart = pZip.CreatePart(partURI, _
                                                     Net.Mime.MediaTypeNames.Application.Zip, _
                                                     CompressionOption.Normal)
        'Read the file into a byte array
        Dim arrBuffer As Byte() = File.ReadAllBytes(pFileToAdd)
        'Add the array of byte to the Package
        pkgPart.GetStream().Write(arrBuffer, 0, arrBuffer.Length)
    End Sub

    Private Function CreateUriFromFilename(ByVal pFileName As String) As String
        'Replaces invalid characters
        pFileName = pFileName.Replace(" "c, "_"c)

        'insert a / as the first character (because a Uri MUST start with a /)
        pFileName = String.Concat("/", IO.Path.GetFileName(pFileName))

        Return pFileName
    End Function




End Class
