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

        Dim edumateStudents As List(Of user)
        edumateStudents = getEdumateStudents(config)

        Dim dirEntry As DirectoryEntry
        dirEntry = GetDirectoryEntry(config)

        Dim adUsers As List(Of user)
        adUsers = getADUsers(dirEntry)

        Dim usersToAdd As List(Of user)
        usersToAdd = getEdumateUsersNotInAD(edumateStudents, adUsers)

        usersToAdd = excludeUserOutsideEnrollDate(usersToAdd)

        For Each user In usersToAdd
            Console.WriteLine(user.firstName & " " & user.surname)
        Next

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
SELECT        
contact.firstname, 
contact.surname, 
view_student_start_exit_dates.start_date, 
view_student_start_exit_dates.exit_date, 
student.student_id, 
form.short_name AS grad_form,
YEAR(student_form_run.end_date) as EndYear


FROM            
OFGSODBC.STUDENT, 
contact, 
view_student_start_exit_dates, 
student_form_run, 
form_run, 
form


WHERE (student.contact_id = contact.contact_id) 
AND (student.student_id = view_student_start_exit_dates.student_id) 
AND (student_form_run.student_id = student.student_id) 
AND (form_run.form_id = form.form_id) 
AND (YEAR(view_student_start_exit_dates.exit_date) = YEAR(student_form_run.end_date)) 
AND (student_form_run.form_run_id = form_run.form_run_id)
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

                    users.Last.firstName = dr.GetValue(0)
                    users.Last.surname = dr.GetValue(1)
                    users.Last.startDate = dr.GetValue(2)
                    users.Last.endDate = dr.GetValue(3)
                    users.Last.employeeID = dr.GetValue(4)
                    users.Last.classOf = getYearOf(dr.GetValue(5), dr.GetValue(6))


                End If
            End While
            conn.Close()
        End Using
        Return users
    End Function

    Function getYearOf(ByVal gradForm As String, ByVal endYear As String)

        Select Case gradForm
            Case "K"
                getYearOf = endYear + 12
            Case "01"
                getYearOf = endYear + 11
            Case "02"
                getYearOf = endYear + 10
            Case "03"
                getYearOf = endYear + 9
            Case "04"
                getYearOf = endYear + 8
            Case "05"
                getYearOf = endYear + 7
            Case "06"
                getYearOf = endYear + 6
            Case "07"
                getYearOf = endYear + 5
            Case "08"
                getYearOf = endYear + 4
            Case "09"
                getYearOf = endYear + 3
            Case "10"
                getYearOf = endYear + 2
            Case "11"
                getYearOf = endYear + 1
            Case "12"
                getYearOf = endYear
            Case Else
                getYearOf = ""
        End Select
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
        dirEntry.RefreshCache()
        Return dirEntry

    End Function


    Sub createUser(ByVal dirEntry As DirectoryEntry, ByVal objUserToAdd As user)



        Dim objUser As DirectoryEntry       ' User object.

        Dim strDisplayName As String        ' Display name of user.
        Dim strUser As String               ' User to create.
        Dim strUserPrincipalName As String  ' Principal name of user.

        ' Construct the binding string.


        ' Specify User.
        strUser = "CN=AccTestUser"
        strDisplayName = "Acc Test User"
        strUserPrincipalName = "accTestUser@ofgs.nsw.edu.au"
        Console.WriteLine("Create:  {0}", strUser)

        ' Create User.
        Try
            objUser = dirEntry.Children.Add(strUser, "user")
            objUser.Properties("displayName").Add(strDisplayName)
            objUser.Properties("userPrincipalName").Add(
                    strUserPrincipalName)
            objUser.CommitChanges()
        Catch e As Exception
            Console.WriteLine("Error:   Create failed.")
            Console.WriteLine("         {0}", e.Message)
            Return
        End Try

        ' Output User attributes.
        Console.WriteLine("Success: Create succeeded.")
        Console.WriteLine("Name:    {0}", objUser.Name)
        Console.WriteLine("         {0}",
                objUser.Properties("displayName").Value)
        Console.WriteLine("         {0}",
                objUser.Properties("userPrincipalName").Value)
        Return

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
                    'Console.WriteLine(group)
                Next
            End If

            If result.Properties("userAccountControl").Count = 66048 Then
                adUsers.Last.enabled = True
            End If
            If result.Properties("userAccountControl").Count = 66050 Then
                adUsers.Last.enabled = False
            End If

        Next
        Return adUsers
    End Function

    Function getEdumateUsersNotInAD(ByVal edumateUsers As List(Of user), ByVal adUsers As List(Of user))

        Dim usersToAdd As New List(Of user)

        For Each edumateUser In edumateUsers
            Dim found As Boolean = False
            For Each adUser In adUsers
                If adUser.employeeID = edumateUser.employeeID Then
                    found = True
                End If

            Next
            If found = False Then
                usersToAdd.Add(edumateUser)
            End If
        Next

        Return usersToAdd
    End Function

    Function excludeUserOutsideEnrollDate(ByVal users As List(Of user))

        Dim ReturnUsers As New List(Of user)

        For Each user In users
            If user.endDate > Date.Now() And user.startDate < (Date.Now.AddDays(0)) Then
                ReturnUsers.Add(user)
            End If
        Next
        Return ReturnUsers
    End Function

End Module
