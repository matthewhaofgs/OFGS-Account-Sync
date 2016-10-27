Imports System.IO
Imports System.DirectoryServices


Module AccountSync

    Class user
        Public firstName
        Public surname
        Public displayName
        Public email
        Public username
        Public profilePath
        Public HomePath
        Public HomeDriveLetter
        Public classOf
        Public employeeID
        Public employeeNumber
        Public password
        Public startDate
        Public endDate
        Public enabled
        Public memberOf As New List(Of String)
        Public userAccountControl
    End Class

    Class configSettings
        Public edumateConnectionString As String
        Public ldapDirectoryEntry As String
    End Class


    Sub Main()
        Dim config As New configSettings()
        config = readConfig()

        Dim dirEntry As DirectoryEntry
        dirEntry = GetDirectoryEntry(config)

        getADUsers(dirEntry)
    End Sub


    Private Function readConfig()
        Dim config As New configSettings()

        Try
            ' Open the file using a stream reader.
            Dim directory As String = My.Application.Info.DirectoryPath



            Using sr As New StreamReader(directory & "\config.ini")
                Dim line As String
                While Not sr.EndOfStream
                    line = sr.ReadLine

                    Select Case True
                        Case Left(line, 24) = "edumateConnectionstring="
                            config.edumateConnectionString = Mid(line, 25)
                        Case Left(line, 19) = "ldapDirectoryEntry="
                            config.edumateConnectionString = Mid(line, 20)
                    End Select

                End While

                readConfig = config
            End Using

        Catch e As Exception
            MsgBox(e.Message)
        End Try
    End Function



    Function getEdumateStudents(config As configSettings)


        Dim ConnectionString As String = config.edumateConnectionString
        Dim commandString As String =
"
select
username,
student_number,
firstname,
surname,
birthdate,
form_name
from schoolbox_students
"


        Dim users As New List(Of user)



        Using conn As New System.Data.Odbc.OdbcConnection(ConnectionString)
            conn.Open()

            'define the command object to execute
            Dim command As New System.Data.Odbc.OdbcCommand(commandString, conn)
            command.Connection = conn
            command.CommandText = commandString

            Dim dr As System.Data.Odbc.OdbcDataReader
            dr = command.ExecuteReader

            Dim i As Integer = 0
            While dr.Read()
                If Not dr.IsDBNull(0) Then
                    users.Add(New user)

                    users.Last.password = ""
                End If
            End While
        End Using
    End Function



    Sub createUsers(usersToCreate As List(Of user))

    End Sub

    ''' <returns>DirectoryEntry</returns>
    Public Function GetDirectoryEntry(config As configSettings) As DirectoryEntry

        Dim dirEntry As New DirectoryEntry(config.ldapDirectoryEntry)
        'Setting username & password to Nothing forces
        'the connection to use your logon credentials
        dirEntry.Username = Nothing
        dirEntry.Password = Nothing
        'Always use a secure connection
        dirEntry.AuthenticationType = AuthenticationTypes.Secure
        Return dirEntry
    End Function


    Sub createUser(dirEntry As DirectoryEntry)


    End Sub

    Function getADUsers(dirEntry As DirectoryEntry)
        Dim searcher As New DirectorySearcher(dirEntry)
        Dim adUsers As New List(Of user)

        searcher.PropertiesToLoad.Add("cn")
        searcher.PropertiesToLoad.Add("employeeID")
        searcher.PropertiesToLoad.Add("distinguishedName")
        searcher.PropertiesToLoad.Add("employeeNumber")
        searcher.PropertiesToLoad.Add("givenName")
        searcher.PropertiesToLoad.Add("homeDirectory")
        searcher.PropertiesToLoad.Add("homeDrive")
        searcher.PropertiesToLoad.Add("mail")
        searcher.PropertiesToLoad.Add("profilePath")
        searcher.PropertiesToLoad.Add("samAccountName")
        searcher.PropertiesToLoad.Add("sn")
        searcher.PropertiesToLoad.Add("userPrincipalName")
        searcher.PropertiesToLoad.Add("memberof")
        searcher.PropertiesToLoad.Add("userAccountControl")


        searcher.Filter = "(objectCategory=person)"
        searcher.ServerTimeLimit = New TimeSpan(0, 0, 60)
        searcher.SizeLimit = 100000000


        Dim queryResults As SearchResultCollection
        queryResults = searcher.FindAll

        Dim result As SearchResult
        For Each result In queryResults
            adUsers.Add(New user)

            If result.Properties("givenName").Count > 0 Then adUsers.Last.firstName = result.Properties("givenName")(0)
            If result.Properties("sn").Count > 0 Then adUsers.Last.surname = result.Properties("sn")(0)
            If result.Properties("cn").Count > 0 Then adUsers.Last.displayName = result.Properties("cn")(0)
            If result.Properties("mail").Count > 0 Then adUsers.Last.email = result.Properties("mail")(0)
            If result.Properties("samAccountName").Count > 0 Then adUsers.Last.username = result.Properties("samAccountName")(0)
            If result.Properties("profilePath").Count > 0 Then adUsers.Last.profilePath = result.Properties("profilePath")(0)
            If result.Properties("homeDirectory").Count > 0 Then adUsers.Last.HomePath = result.Properties("homeDirectory")(0)
            If result.Properties("homeDrive").Count > 0 Then adUsers.Last.HomeDriveLetter = result.Properties("homeDrive")(0)
            If result.Properties("employeeID").Count > 0 Then adUsers.Last.employeeID = result.Properties("employeeID")(0)
            If result.Properties("employeeNumber").Count > 0 Then adUsers.Last.employeeNumber = result.Properties("employeeNumber")(0)
            If result.Properties("userAccountControl").Count > 0 Then adUsers.Last.userAccountControl = result.Properties("userAccountControl")(0)

            If result.Properties("memberof").Count > 0 Then
                For Each group In result.Properties("memberof")
                    adUsers.Last.memberOf.Add(group)
                    Console.WriteLine(group)
                Next
            End If

            If result.Properties("userAccountControl").Count = 66048 Then
                adUsers.Last.enabled = True
            End If
            If result.Properties("userAccountControl").Count = 66050 Then
                adUsers.Last.enabled = False
            End If

        Next

    End Function

End Module
