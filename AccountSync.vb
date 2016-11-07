Imports System.IO
Imports System.DirectoryServices
Imports System.Text.RegularExpressions



Module AccountSync

    Class user
        Public firstName As String
        Public surname As String
        Public displayName As String
        Public email As String
        Public username As String
        Public profilePath As String
        Public HomePath As String
        Public HomeDriveLetter As String
        Public classOf As Integer
        Public employeeID As Integer
        Public employeeNumber
        Public password As String
        Public startDate
        Public endDate
        Public enabled
        Public memberOf As New List(Of String)
        Public userAccountControl
        Public userType
        Public children As New List(Of user)
    End Class

    Class configSettings
        Public edumateConnectionString As String
        Public ldapDirectoryEntry As String
        Public daysInAdvanceToCreateAccounts As Integer
        Public studentDomainName As String
        Public studentProfilePath As String
        Public parentOU As String
        Public parentDomainName As String
    End Class



    Sub Main()
        Dim config As New configSettings()
        Console.Clear()
        Console.WriteLine("Reading config...")
        config = readConfig()



        Dim edumateStudents As List(Of user)

        Console.WriteLine("Getting Edumate student data...")
        edumateStudents = getEdumateStudents(config)


        Dim dirEntry As DirectoryEntry

        Console.WriteLine("Connecting to AD...")
        dirEntry = GetDirectoryEntry(config)

        Dim adUsers As List(Of user)
        Console.WriteLine("Loading AD users...")
        Console.WriteLine("")
        Console.WriteLine("")
        adUsers = getADUsers(dirEntry)

        Dim usersToAdd As List(Of user)
        usersToAdd = getEdumateUsersNotInAD(edumateStudents, adUsers)

        usersToAdd = excludeUserOutsideEnrollDate(usersToAdd, config)

        Console.WriteLine("Found " & usersToAdd.Count & " users to add")
        Console.WriteLine("")

        If usersToAdd.Count > 0 Then
            usersToAdd = evaluateUsernames(usersToAdd, adUsers)
            createUsers(dirEntry, usersToAdd, config)
        End If

        Console.WriteLine("Getting Edumate parent data...")
        Console.WriteLine("")
        Dim edumateParents As List(Of user)
        edumateParents = getEdumateParents(config, edumateStudents)

        Dim parentsToAdd As List(Of user)
        parentsToAdd = getEdumateUsersNotInAD(edumateParents, adUsers)

        parentsToAdd = excludeParentsOutsideEnrollDate(config, parentsToAdd)

        Console.WriteLine("Found " & parentsToAdd.Count & " users to add")
        If parentsToAdd.Count > 0 Then
            parentsToAdd = evaluateUsernames(parentsToAdd, adUsers)
            createUsers(dirEntry, parentsToAdd, config)
        End If


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
                            config.ldapDirectoryEntry = Mid(line, 20)
                        Case Left(line, 30) = "daysInAdvanceToCreateAccounts="
                            config.daysInAdvanceToCreateAccounts = Mid(line, 31)
                        Case Left(line, 18) = "studentDomainName="
                            config.studentDomainName = Mid(line, 19)
                        Case Left(line, 19) = "studentProfilePath="
                            config.studentProfilePath = Mid(line, 20)
                        Case Left(line, 9) = "parentOU="
                            config.parentOU = Mid(line, 10)
                        Case Left(line, 17) = "parentDomainName="
                            config.parentDomainName = Mid(line, 18)

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
STUDENT, 
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
                    users.Last.userType = "Student"
                    users.Last.displayName = users.Last.firstName & " " & users.Last.surname
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


    Sub createUsers(dirEntry As DirectoryEntry, ByVal objUsersToAdd As List(Of user), ByVal config As configSettings)

        For Each objUserToAdd In objUsersToAdd


            Dim objUser As DirectoryEntry
            Dim strDisplayName As String        '
            Dim intEmployeeID As Integer
            Dim strUser As String               ' User to create.
            Dim strUserPrincipalName As String  ' Principal name of user.
            Dim strDescription As String

            'common properties for all user types
            intEmployeeID = objUserToAdd.employeeID
            strDisplayName = objUserToAdd.displayName




            Select Case objUserToAdd.userType
                Case "Student"
                    strUser = "CN=" & objUserToAdd.displayName & ",OU=" & objUserToAdd.classOf.ToString & ",OU=Student Users"
                    strUserPrincipalName = objUserToAdd.username & config.studentDomainName
                    strDescription = "Class of " & objUserToAdd.classOf & " Barcode:"
                Case "Staff"
                    'do stuff

                Case "Parent"
                    strUser = "CN=" & objUserToAdd.username & "," & config.parentOU
                    strDescription = objUserToAdd.firstName & " " & objUserToAdd.surname
                    strDisplayName = objUserToAdd.username
                    strUserPrincipalName = objUserToAdd.username & config.parentDomainName
                Case Else
                    'Do Else

            End Select

            Console.WriteLine("Create:  {0}", strUser)

            ' Create User.
            Try
                objUser = dirEntry.Children.Add(strUser, "user")
                objUser.Properties("displayName").Add(strDisplayName)
                objUser.Properties("userPrincipalName").Add(strUserPrincipalName)
                objUser.Properties("EmployeeID").Add(intEmployeeID)

                objUser.Properties("givenName").Add(objUserToAdd.firstName)
                objUser.Properties("samAccountName").Add(objUserToAdd.username)
                objUser.Properties("sn").Add(objUserToAdd.surname)
                objUser.Properties("mail").Add(strUserPrincipalName)
                objUser.Properties("description").Add(strDescription)

                'objUser.Properties("employeeNumber")
                'objUser.Properties("homeDirectory")
                'objUser.Properties("homeDrive")

                objUser.CommitChanges()
            Catch e As Exception
                Console.WriteLine("Error:   Create failed.")
                Console.WriteLine("         {0}", e.Message)
                Return
            End Try

            objUser.Invoke("setPassword", New Object() {createPassword()})
            objUser.CommitChanges()


            Const ADS_UF_ACCOUNTDISABLE = &H10200
            objUser.Properties("userAccountControl").Value = ADS_UF_ACCOUNTDISABLE
            objUser.CommitChanges()


            ' Output User attributes.
            Console.WriteLine("Success: Create succeeded.")
            Console.WriteLine("Name:    {0}", objUser.Name)
            Console.WriteLine("         {0}",
                    objUser.Properties("displayName").Value)
            Console.WriteLine("         {0}",
                    objUser.Properties("userPrincipalName").Value)


            Return

        Next
    End Sub








    Function getADUsers(dirEntry As DirectoryEntry)
        Using searcher As New DirectorySearcher(dirEntry)
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
            searcher.Asynchronous = False
            searcher.ServerPageTimeLimit = New TimeSpan(0, 0, 60)
            searcher.PageSize = 10000

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

                    Next
                End If

                'If result.Properties("userAccountControl")(0) = 66048 Then
                ' adUsers.Last.enabled = True
                ' End If
                ' If result.Properties("userAccountControl")(0) = 66050 Then
                '  adUsers.Last.enabled = False
                '  End If

            Next
            Return adUsers
        End Using
    End Function

    Function getEdumateUsersNotInAD(ByVal edumateUsers As List(Of user), ByVal adUsers As List(Of user))

        Dim usersToAdd As New List(Of user)

        Console.WriteLine("Evaluating users to create:")
        Dim i As Integer = 1

        For Each edumateUser In edumateUsers

            CONSOLE__WRITE(String.Format("Processed {0} of {1}", i, edumateUsers.Count))
            Dim found As Boolean = False
            For Each adUser In adUsers

                If adUser.employeeID = edumateUser.employeeID Then
                    found = True
                End If

            Next
            If found = False Then
                usersToAdd.Add(edumateUser)
            End If
            i = i + 1
        Next
        CONSOLE__CLEAR_EOL()
        Return usersToAdd
    End Function

    Function excludeUserOutsideEnrollDate(ByVal users As List(Of user), config As configSettings)

        Dim ReturnUsers As New List(Of user)

        For Each user In users
            If user.endDate > Date.Now() And user.startDate < (Date.Now.AddDays(config.daysInAdvanceToCreateAccounts)) Then
                ReturnUsers.Add(user)

            End If
        Next
        Return ReturnUsers
    End Function



    Public Sub CONSOLE__WRITE(ByRef szText As String, Optional ByVal bClearEOL As Boolean = True)
        'Output the text
        Console.Write(szText)
        'Optionally clear to end of line (EOL)
        If bClearEOL Then CONSOLE__CLEAR_EOL()
        'Move cursor back to where we started, using Backspaces
        Console.Write(Microsoft.VisualBasic.StrDup(szText.Length(), Chr(8)))
    End Sub

    Public Sub CONSOLE__CLEAR_EOL()
        'Clear to End of line (EOL)
        'Save window and cursor positions
        Dim x As Integer = Console.CursorLeft
        Dim y As Integer = Console.CursorTop
        Dim wx As Integer = Console.WindowLeft
        Dim wy As Integer = Console.WindowTop
        'Write spaces until end of buffer width
        Console.Write(Space(Console.BufferWidth - x))
        'Restore window and cursor position
        Console.SetWindowPosition(wx, wy)
        Console.SetCursorPosition(x, y)
    End Sub


    Function evaluateUsernames(users As List(Of user), adusers As List(Of user))
        Console.WriteLine("Evaluating usernames for new users...")
        Console.WriteLine("")
        For Each user In users



            Console.WriteLine("User:" & user.firstName & " " & user.surname)
            Dim strUsername As String
            Select Case user.userType
                Case "Student"
                    Dim rgx As New Regex("[^a-zA-Z]")
                    Dim availableNameFound As Boolean = False
                    Dim i As Integer = 1

                    While availableNameFound = False And i <= user.firstName.Length

                        strUsername = rgx.Replace(user.surname & Left(user.firstName, i), "").ToLower
                        Console.WriteLine("Trying " & strUsername & "...")
                        Dim duplicate As Boolean = False
                        Dim a As Integer = 1
                        For Each adUser In adusers

                            CONSOLE__WRITE(String.Format("Checking for duplicates {0} of {1}", a, adusers.Count))
                            Try
                                adUser.username = adUser.username.ToLower
                            Catch ex As Exception

                            End Try

                            If strUsername = adUser.username Then
                                duplicate = True
                            End If
                            a = a + 1
                        Next
                        If duplicate = False Then
                            availableNameFound = True
                            user.username = strUsername
                        End If

                        i = i + 1
                    End While

                    If user.username = "" Then
                        Console.WriteLine("No valid username available for " & user.firstName & " " & user.surname)
                    Else
                        Console.WriteLine(user.firstName & " " & user.surname & " will be created as " & user.username)
                    End If


                Case "Staff"
                'Do stuff


                Case "Parent"

                    Dim rgx As New Regex("[^a-zA-Z0-9]")
                    user.username = rgx.Replace(Left(user.surname, 5) & user.employeeID, "").ToLower
                    Console.WriteLine(user.firstName & " " & user.surname & " will be created as " & user.username)
                    Console.WriteLine("")
                Case Else
                    'Do Else
            End Select
        Next
        Return users
    End Function


    Function createPassword()

        Dim wordlist As List(Of String)
        wordlist = getWordList()



        Dim RandomClass As New Random(System.DateTime.Now.Millisecond)
        Dim rndNumber As Integer = RandomClass.Next(10, 99)


        'Dim PasswordPosition As Integer = RandomClass.Next(0, 5)
        Dim PasswordPosition As Integer = 2
        Select Case PasswordPosition
            Case 0 : Return getWord(wordlist) & getWord(wordlist) & rndNumber
            Case 1 : Return rndNumber & getWord(wordlist) & getWord(wordlist)
            Case 2 : Return Mixedcase(getWord(wordlist)) & rndNumber & Mixedcase(getWord(wordlist))
            Case 3 : Return getWord(wordlist) & getWord(wordlist) & rndNumber
            Case 4 : Return rndNumber & getWord(wordlist) & getWord(wordlist)
            Case 5 : Return getWord(wordlist) & rndNumber & getWord(wordlist)
            Case Else : Return getWord(wordlist) & getWord(wordlist)
        End Select


    End Function




    Function getWordList()
        Dim directory As String = My.Application.Info.DirectoryPath
        Dim WordList As New List(Of String)
        Dim word As String

        If My.Computer.FileSystem.FileExists(directory & "\wordList.txt") Then
            Dim fields As String()
            Dim delimiter As String = ","
            Using parser As New Microsoft.VisualBasic.FileIO.TextFieldParser(directory & "\wordList.txt")
                parser.SetDelimiters(delimiter)
                While Not parser.EndOfData
                    fields = parser.ReadFields()
                    For Each word In fields
                        WordList.Add(word)
                    Next
                End While
            End Using

        Else
            Throw New Exception(directory & "\wordList.txt" & " doesn't Exist!")
        End If
        Return WordList
    End Function

    Function getWord(wordlist As List(Of String))
        System.Threading.Thread.CurrentThread.Sleep(1)
        Dim Position As New Random(System.DateTime.Now.Millisecond)
        Dim wordnumber As Integer = Position.Next(0, wordlist.Count - 1)
        Return wordlist(wordnumber)
        Position = Nothing
        wordnumber = Nothing
    End Function


    Private Function Mixedcase(ByVal Word As String) As String
        If Word.Length = 0 Then Return Word
        If Word.Length = 1 Then Return UCase(Word)
        Return Word.Substring(0, 1).ToUpper & Word.Substring(1).ToLower

    End Function

    Function getEdumateParents(config As configSettings, edumateStudents As List(Of user))


        Dim ConnectionString As String = config.edumateConnectionString
        Dim commandString As String =
"
SELECT        
parentcontact.firstname,
parentcontact.surname,
carer.carer_id,
student.student_id



FROM            relationship

INNER JOIN contact as ParentContact
ON relationship.contact_id2 = Parentcontact.contact_id

INNER JOIN contact as StudentContact 
ON relationship.contact_id1 = studentContact.contact_id

INNER JOIN student
ON studentContact.contact_id = student.contact_id

INNER JOIN carer 
ON parentcontact.contact_id = carer.contact_id




WHERE        (relationship.relationship_type_id IN (1, 15, 28, 33)) 
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

                    Dim existingParent As user
                    Dim duplicate As Boolean = False

                    For Each user In users
                        If dr.GetValue(2) = user.employeeID Then
                            existingParent = user
                            duplicate = True
                        End If
                    Next

                    If duplicate Then
                        existingParent.children.Add(getStudentFromID(dr.GetValue(3), edumateStudents))
                    Else
                        users.Add(New user)
                        users.Last.firstName = dr.GetValue(0)
                        users.Last.surname = dr.GetValue(1)
                        users.Last.employeeID = dr.GetValue(2)
                        users.Last.userType = "Parent"
                        users.Last.children.Add(getStudentFromID(dr.GetValue(3), edumateStudents))
                    End If
                End If
            End While
            conn.Close()
        End Using




        commandString =
"
SELECT        
parentcontact.firstname,
parentcontact.surname,
carer.carer_id,
student.student_id



FROM            relationship

INNER JOIN contact as ParentContact
ON relationship.contact_id1 = Parentcontact.contact_id

INNER JOIN contact as StudentContact 
ON relationship.contact_id2 = studentContact.contact_id

INNER JOIN student
ON studentContact.contact_id = student.contact_id

INNER JOIN carer 
ON parentcontact.contact_id = carer.contact_id




WHERE        (relationship.relationship_type_id IN (2, 16, 29, 34)) 
"


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

                    Dim existingParent As user
                    Dim duplicate As Boolean = False

                    For Each user In users
                        If dr.GetValue(2) = user.employeeID Then
                            existingParent = user
                            duplicate = True
                        End If
                    Next

                    If duplicate Then
                        existingParent.children.Add(getStudentFromID(dr.GetValue(3), edumateStudents))
                    Else
                        users.Add(New user)
                        users.Last.firstName = dr.GetValue(0)
                        users.Last.surname = dr.GetValue(1)
                        users.Last.employeeID = dr.GetValue(2)
                        users.Last.userType = "Parent"
                        users.Last.children.Add(getStudentFromID(dr.GetValue(3), edumateStudents))
                    End If
                End If
            End While
            conn.Close()
        End Using

        Return users
    End Function


    Function getStudentFromID(ByVal student_id As String, edumateStudents As List(Of user))

        For Each user In edumateStudents
            If user.employeeID = student_id Then
                Return user
            End If
        Next

    End Function

    Function excludeParentsOutsideEnrollDate(ByVal config As configSettings, users As List(Of user))

        Dim ReturnUsers As New List(Of user)


        For Each user In users
            Dim current As Boolean = False
            For Each student In user.children
                Try
                    If student.endDate > Date.Now() And student.startDate < (Date.Now.AddDays(config.daysInAdvanceToCreateAccounts)) Then
                        current = True
                    End If
                Catch
                End Try
            Next
            If current Then
                ReturnUsers.Add(user)
            End If
        Next

        Return ReturnUsers
    End Function





End Module
