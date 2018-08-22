Imports System.IO
Imports System.DirectoryServices
Imports System.Text.RegularExpressions
Imports System.Net.Mail
Imports MySql.Data.MySqlClient
Imports System.Text
Imports WinSCP



Public Module AccountSync

    Class user
        Public firstName As String
        Public surname As String
        Public displayName As String
        Public email As String
        Public ad_username As String
        Public profilePath As String
        Public HomePath As String
        Public HomeDriveLetter As String
        Public classOf As String
        Public employeeID As String
        Public employeeNumber
        Public password
        Public startDate
        Public endDate
        Public enabled
        Public memberOf As New List(Of String)
        Public userAccountControl
        Public userType
        Public children As New List(Of user)
        Public mailTo As New List(Of String)
        Public currentYear As String
        Public distinguishedName As String
        Public edumateUsername As String
        Public edumateCurrent As String
        Public edumateEmail
        Public smtpProxy As String
        Public edumateLoginActive As String
        Public employmentType
        Public edumateStaffNumber
        Public dob
        Public libraryCard As String
        Public rollClass
        Public bosNumber
        Public edumateGroupMemberships As List(Of String)
        Public contact_id As String
        Public adObject As SearchResult
        Public edumateProperties As New EdumateProperties
        Public enrolledClasses As String
        Public adFirstname As String
        Public adSurname As String



    End Class

    Class EdumateProperties
        Public firstName As String
        Public surname As String
        Public startDate As String
        Public endDate As String
        Public employeeID As String
        Public classOf As String
        Public userType As String 'student etc..
        Public displayName As String
        Public employeeNumber As String
        Public dob As String
        Public libraryCard As String
        Public rollClass As String
        Public bosNumber As String
        Public carer_number As String
    End Class



    Class configSettings
        Public edumateConnectionString As String
        Public ldapDirectoryEntry As String
        Public daysInAdvanceToCreateAccounts As Integer
        Public studentDomainName As String
        Public studentProfilePath As String
        Public parentOU As String
        Public parentDomainName As String
        Public serverAddress As String
        Public serverPort As String
        Public enableSSL As Boolean
        Public username As String
        Public password As String
        Public displayName As String
        Public applyChanges As Boolean
        Public staffDomainName As String
        Public domain As String
        Public studentAlumOU As String
        Public tutorGroupID As Integer
        Public danceTutorGroupID As Integer
        Public staffHomePath As String
        Public formerStaffOU As String
        Public currentStaffOU As String
        Public tutorOU As String
        Public casualStaffOU As String

        Public mailToAll As New List(Of String)
        Public mailToParent As New List(Of String)
        Public mailToK As New List(Of String)
        Public mailTo1 As New List(Of String)
        Public mailTo2 As New List(Of String)
        Public mailTo3 As New List(Of String)
        Public mailTo4 As New List(Of String)
        Public mailTo5 As New List(Of String)
        Public mailTo6 As New List(Of String)
        Public mailTo7 As New List(Of String)
        Public mailTo8 As New List(Of String)
        Public mailTo9 As New List(Of String)
        Public mailTo10 As New List(Of String)
        Public mailTo11 As New List(Of String)
        Public mailTo12 As New List(Of String)


        Public mySQLDatabaseName As String
        Public mySQLserver As String
        Public mySQLUserName As String
        Public mySQLPassword As String

        Public sg_k As String
        Public sg_1 As String
        Public sg_2 As String
        Public sg_3 As String
        Public sg_4 As String
        Public sg_5 As String
        Public sg_6 As String
        Public sg_7 As String
        Public sg_8 As String
        Public sg_9 As String
        Public sg_10 As String
        Public sg_11 As String
        Public sg_12 As String

        Public sg_tutors As String

    End Class

    Class emailNotification
        Public mailTo
        Public body
    End Class

    Class studentParent
        Public student_id As Integer
        Public parent_id As Integer
    End Class

    Class SchoolBoxUser
        Public Delete As String
        Public SchoolboxUserID As String
        Public Username As String
        Public ExternalID As String
        Public Title As String
        Public FirstName As String
        Public Surname As String
        Public Role As String
        Public Campus As String
        Public Password As String
        Public AltEmail As String
        Public Year As String
        Public House As String
        Public ResidentialHouse As String
        Public EPortfolio As String
        Public HideContactDetails As String
        Public HideTimetable As String
        Public EmailAddressFromUsername As String
        Public UseExternalMailClient As String
        Public EnableWebmailTab As String
        Public Superuser As String
        Public AccountEnabled As String
        Public ChildExternalIDs As String
        Public DateOfBirth As String
        Public HomePhone As String
        Public MobilePhone As String
        Public WorkPhone As String
        Public Address As String
        Public Suburb As String
        Public Postcode As String
        Public PositionTitle As String
    End Class

    Class uploadServer
        Public host As String
        Public userName As String
        Public pass As String
        Public rsa As String
    End Class

    Class schoolboxConfigSettings
        Public connectionString As String
        Public uploadServers As List(Of uploadServer)
        Public studentEmailDomain As String
    End Class

    Sub Main()
        Dim config As New configSettings()
        Console.Clear()
        Console.WriteLine("Reading config...")
        config = readConfig()


        'Declare and connect to mySQL Database to text connection is working
        Dim conn As New MySqlConnection
        connect(conn, config)

        'Get ALL AD Data
        Dim dirEntry As DirectoryEntry
        Console.WriteLine("Connecting to AD...")
        dirEntry = GetDirectoryEntry(config.ldapDirectoryEntry)
        Dim adUsers As List(Of user)
        Console.WriteLine("Loading AD users...")
        Console.WriteLine("")
        Console.WriteLine("")
        adUsers = getADUsers(dirEntry)


        'Get Edumate data for students
        Dim edumateStudents As List(Of user)
        Console.WriteLine("Getting Edumate student data...")
        edumateStudents = getEdumateStudents(config)

        'Get student users who do not yet have accounts
        Dim studentUsersToAdd As List(Of user)
        studentUsersToAdd = getEdumateUsersNotInAD(edumateStudents, adUsers)
        studentUsersToAdd = excludeUserOutsideEnrollDate(studentUsersToAdd, config)
        studentUsersToAdd = addMailTo(config, studentUsersToAdd)
        studentUsersToAdd = calculateCurrentYears(studentUsersToAdd)
        Console.WriteLine("Found " & studentUsersToAdd.Count & " users to add")
        Console.WriteLine("")

        'Create student accounts
        If studentUsersToAdd.Count > 0 Then
            studentUsersToAdd = evaluateUsernames(studentUsersToAdd, adUsers)
            createUsers(studentUsersToAdd, config, conn)
        End If

        'Get Edumate data for parents 
        Console.WriteLine("Getting Edumate parent data...")
        Console.WriteLine("")
        Dim edumateParents As List(Of user)
        edumateParents = getEdumateParents(config, edumateStudents)

        'Get parent users who do not yet have accounts
        Dim parentsToAdd As List(Of user)
        parentsToAdd = getEdumateUsersNotInAD(edumateParents, adUsers)
        parentsToAdd = excludeParentsOutsideEnrollDate(config, parentsToAdd)
        parentsToAdd = addMailTo(config, parentsToAdd)
        Console.WriteLine("Found " & parentsToAdd.Count & " users to add")

        'Create Parent Accounts
        If parentsToAdd.Count > 0 Then
            parentsToAdd = evaluateUsernames(parentsToAdd, adUsers)
            createUsers(parentsToAdd, config, conn)
        End If

        'Get Edumate data for staff
        Dim edumateStaff As List(Of user)
        Console.WriteLine("Getting Edumate staff data...")
        edumateStaff = getEdumateStaff(config)

        'Get staff users who do not yet have accounts
        Dim staffToAdd As List(Of user)
        staffToAdd = getEdumateUsersNotInAD(edumateStaff, adUsers)
        staffToAdd = excludeUserOutsideEnrollDate(staffToAdd, config)
        staffToAdd = addMailTo(config, staffToAdd)
        Console.WriteLine("Found " & staffToAdd.Count & " users to add")

        'Create staff accounts
        If staffToAdd.Count > 0 Then
            staffToAdd = evaluateUsernames(staffToAdd, adUsers)
            createUsers(staffToAdd, config, conn)
        End If

        'Add staff to AD groups
        adUsers = addEdumateDetailsToAdUsers(adUsers, edumateStaff)
        adUsers = getEdumateGroups(adUsers, config)
        AddStaffToGroups(adUsers, config)

        'Re Pull AD data after creating new accounts
        adUsers = getADUsers(dirEntry)

        'Add usernames to account objects / refresh data now accountsd are created
        edumateParents = addUsernamesToUsers(edumateParents, adUsers)
        Dim currentEdumateStudents As List(Of user)
        currentEdumateStudents = excludeUserOutsideEnrollDate(edumateStudents, config)
        currentEdumateStudents = addUsernamesToUsers(currentEdumateStudents, adUsers)
        currentEdumateStudents = calculateCurrentYears(currentEdumateStudents)


        'MYSQL Database for student details
        Dim mySQLStudents As List(Of user)
        mySQLStudents = getMySQLStudents(conn)
        updatePasswordsInMysql(mySQLStudents, conn)

        Dim mysqlUsersToAdd As List(Of user)
        mysqlUsersToAdd = getEdumateUsersNotInAD(currentEdumateStudents, mySQLStudents)
        For Each mySQLUserTOAdd In mysqlUsersToAdd
            addUsertoMySQL(conn, mySQLUserTOAdd)
        Next
        updateCurrentFlags(mySQLStudents, currentEdumateStudents, conn, adUsers)

        AddStudentsToYearGroups(currentEdumateStudents, config)
        currentEdumateStudents = addADObjectoToEdumateUser(currentEdumateStudents, adUsers)
        updateMSQLDetails(currentEdumateStudents, mySQLStudents, conn)

        'Schoolbox Stuff
        SchoolboxMain(config, currentEdumateStudents, edumateParents)



        'Staff MYSQL Database
        purgeStaffDB(config)
        updateStaffDatabase(config)


        adUsers = addUserTypeToADUSersFromEdumate(adUsers, edumateStudents)
        adUsers = addUserTypeToAdUsers(adUsers)
        adUsers = addEdumateDetailsToAdUsers(adUsers, edumateStaff)
        adUsers = getEdumateGroups(adUsers, config)



        addParentsToGroups(edumateParents)

        moveUsersToOUs(adUsers, config)




    End Sub

    Function readConfig()
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
                        Case Left(line, 14) = "serverAddress="
                            config.serverAddress = Mid(line, 15)
                        Case Left(line, 11) = "serverPort="
                            config.serverPort = Mid(line, 12)
                        Case Left(line, 10) = "enableSSL="
                            config.enableSSL = Mid(line, 11)
                        Case Left(line, 9) = "username="
                            config.username = Mid(line, 10)
                        Case Left(line, 9) = "password="
                            config.password = Mid(line, 10)
                        Case Left(line, 12) = "displayName="
                            config.displayName = (Mid(line, 13))
                        Case Left(line, 13) = "applyChanges="
                            config.applyChanges = (Mid(line, 14))
                        Case Left(line, 10) = "mailToAll="
                            config.mailToAll.Add(Mid(line, 11))
                        Case Left(line, 13) = "mailToParent="
                            config.mailToParent.Add(Mid(line, 14))
                        Case Left(line, 16) = "staffDomainName="
                            config.staffDomainName = (Mid(line, 17))
                        Case Left(line, 8) = "mailToK="
                            config.mailToK.Add(Mid(line, 9))
                        Case Left(line, 8) = "mailTo1="
                            config.mailTo1.Add(Mid(line, 9))
                        Case Left(line, 8) = "mailTo2="
                            config.mailTo2.Add(Mid(line, 9))
                        Case Left(line, 8) = "mailTo3="
                            config.mailTo3.Add(Mid(line, 9))
                        Case Left(line, 8) = "mailTo4="
                            config.mailTo4.Add(Mid(line, 9))
                        Case Left(line, 8) = "mailTo5="
                            config.mailTo5.Add(Mid(line, 9))
                        Case Left(line, 8) = "mailTo6="
                            config.mailTo6.Add(Mid(line, 9))
                        Case Left(line, 8) = "mailTo7="
                            config.mailTo7.Add(Mid(line, 9))
                        Case Left(line, 8) = "mailTo8="
                            config.mailTo8.Add(Mid(line, 9))
                        Case Left(line, 8) = "mailTo9="
                            config.mailTo9.Add(Mid(line, 9))
                        Case Left(line, 9) = "mailTo10="
                            config.mailTo10.Add(Mid(line, 10))
                        Case Left(line, 9) = "mailTo11="
                            config.mailTo11.Add(Mid(line, 10))
                        Case Left(line, 9) = "mailTo12="
                            config.mailTo12.Add(Mid(line, 10))
                        Case Left(line, 14) = "studentAlumOU="
                            config.studentAlumOU = (Mid(line, 15))
                        Case Left(line, 13) = "tutorGroupId="
                            config.tutorGroupID = (Mid(line, 14))
                        Case Left(line, 18) = "danceTutorGroupId="
                            config.danceTutorGroupID = (Mid(line, 19))




                        Case Left(line, 18) = "mySQLDatabaseName="
                            config.mySQLDatabaseName = (Mid(line, 19))
                        Case Left(line, 12) = "mySQLserver="
                            config.mySQLserver = (Mid(line, 13))
                        Case Left(line, 14) = "mySQLUserName="
                            config.mySQLUserName = (Mid(line, 15))
                        Case Left(line, 14) = "mySQLPassword="
                            config.mySQLPassword = (Mid(line, 15))
                        Case Left(line, 7) = "domain="
                            config.domain = (Mid(line, 8))

                        Case Left(line, 5) = "sg_k="
                            config.sg_k = (Mid(line, 6))
                        Case Left(line, 5) = "sg_1="
                            config.sg_1 = (Mid(line, 6))
                        Case Left(line, 5) = "sg_2="
                            config.sg_2 = (Mid(line, 6))
                        Case Left(line, 5) = "sg_3="
                            config.sg_3 = (Mid(line, 6))
                        Case Left(line, 5) = "sg_4="
                            config.sg_4 = (Mid(line, 6))
                        Case Left(line, 5) = "sg_5="
                            config.sg_5 = (Mid(line, 6))
                        Case Left(line, 5) = "sg_6="
                            config.sg_6 = (Mid(line, 6))
                        Case Left(line, 5) = "sg_7="
                            config.sg_7 = (Mid(line, 6))
                        Case Left(line, 5) = "sg_8="
                            config.sg_8 = (Mid(line, 6))
                        Case Left(line, 5) = "sg_9="
                            config.sg_9 = (Mid(line, 6))
                        Case Left(line, 6) = "sg_10="
                            config.sg_10 = (Mid(line, 7))
                        Case Left(line, 6) = "sg_11="
                            config.sg_11 = (Mid(line, 7))
                        Case Left(line, 6) = "sg_12="
                            config.sg_12 = (Mid(line, 7))
                        Case Left(line, 10) = "sg_tutors="
                            config.sg_tutors = (Mid(line, 11))




                        Case Left(line, 14) = "staffHomePath="
                            config.staffHomePath = (Mid(line, 15))
                        Case Left(line, 14) = "formerStaffOU="
                            config.formerStaffOU = (Mid(line, 15))
                        Case Left(line, 15) = "currentStaffOU="
                            config.currentStaffOU = (Mid(line, 16))
                        Case Left(line, 8) = "tutorOU="
                            config.tutorOU = (Mid(line, 9))
                        Case Left(line, 14) = "casualStaffOU="
                            config.casualStaffOU = (Mid(line, 15))





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
edumate.view_student_start_exit_dates.start_date, 
edumate.view_student_start_exit_dates.exit_date, 
student.student_id, 
form.short_name AS grad_form,
YEAR(student_form_run.end_date) as EndYear,
student.student_number,
contact.birthdate,
stu_school.library_card,
Rollclass.class,
stu_school.bos,
listagg(cast(class.class AS varchar(10000)),',') WITHIN GROUP (ORDER BY class.class ASC) AS classes

FROM            
STUDENT
INNER JOIN contact ON student.contact_id = contact.contact_id
INNER JOIN edumate.view_student_start_exit_dates ON student.student_id = edumate.view_student_start_exit_dates.student_id
INNER JOIN student_form_run ON student_form_run.student_id = student.student_id
INNER JOIN form_run ON student_form_run.form_run_id = form_run.form_run_id
INNER JOIN form ON form_run.form_id = form.form_id
INNER JOIN stu_school ON student.student_id = stu_school.student_id
LEFT JOIN class_enrollment ON student.STUDENT_ID = class_enrollment.STUDENT_ID
LEFT JOIN class ON class_enrollment.class_id = class.class_id 

LEFT JOIN 
	(
	SELECT        
	student.student_id, 
	edumate.view_student_class_enrolment.class

	FROM            
	STUDENT

	INNER JOIN edumate.view_student_class_enrolment ON student.student_id = edumate.view_student_class_enrolment.student_id

	WHERE 
 	(edumate.view_student_class_enrolment.class_type_id = 2)
	AND (edumate.view_student_class_enrolment.academic_year = char(year(current timestamp)))
	) RollClass ON rollclass.student_id = student.student_id

	
WHERE 

(YEAR(edumate.view_student_start_exit_dates.exit_date) = YEAR(student_form_run.end_date)) 

AND (SELECT current date FROM sysibm.sysdummy1) BETWEEN (edumate.view_student_start_exit_dates.start_date - 90 days) AND edumate.view_student_start_exit_dates.exit_date


GROUP BY 

contact.firstname, 
contact.surname, 
edumate.view_student_start_exit_dates.start_date, 
edumate.view_student_start_exit_dates.exit_date, 
student.student_id, 
form.short_name,
YEAR(student_form_run.end_date),
student.student_number,
contact.birthdate,
stu_school.library_card,
Rollclass.class,
stu_school.bos


"


        Dim users As New List(Of user)



        Using conn As New IBM.Data.DB2.DB2Connection(ConnectionString)
            conn.Open()

            'define the command object to execute
            Dim command As New IBM.Data.DB2.DB2Command(commandString, conn)
            command.Connection = conn
            command.CommandText = commandString

            Dim dr As IBM.Data.DB2.DB2DataReader
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

                    users.Last.displayName = Replace(users.Last.firstName, "&#039;", "") & " " & Replace(users.Last.surname, "&#039;", "")


                    users.Last.employeeNumber = dr.GetValue(7)
                    users.Last.dob = dr.GetValue(8)
                    users.Last.libraryCard = dr.GetValue(9)
                    users.Last.rollClass = dr.GetValue(10)
                    users.Last.bosNumber = dr.GetValue(11)
					If Not dr.IsDBNull(12) Then users.Last.enrolledClasses = dr.GetValue(12)

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
            Case "1"
                getYearOf = endYear + 11
            Case "2"
                getYearOf = endYear + 10
            Case "3"
                getYearOf = endYear + 9
            Case "4"
                getYearOf = endYear + 8
            Case "5"
                getYearOf = endYear + 7
            Case "6"
                getYearOf = endYear + 6
            Case "7"
                getYearOf = endYear + 5
            Case "8"
                getYearOf = endYear + 4
            Case "9"





            Case Else
                getYearOf = ""
        End Select
    End Function

    ''' <returns>DirectoryEntry</returns>
    Public Function GetDirectoryEntry(ldapDirectoryEntry As String) As DirectoryEntry

        Dim dirEntry As New DirectoryEntry(ldapDirectoryEntry)
        'Setting username & password to Nothing forces
        'the connection to use your logon credentials
        'dirEntry.Username = Nothing
        'dirEntry.Password = Nothing
        'Always use a secure connection
        dirEntry.AuthenticationType = AuthenticationTypes.Secure
        dirEntry.RefreshCache()
        Return dirEntry

    End Function

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
            searcher.PropertiesToLoad.Add("pwdLastSet")



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

                adUsers.Last.adObject = result

                If result.Properties("givenName").Count > 0 Then adUsers.Last.firstName = result.Properties("givenName")(0)
                If result.Properties("sn").Count > 0 Then adUsers.Last.surname = result.Properties("sn")(0)
                If result.Properties("cn").Count > 0 Then adUsers.Last.displayName = result.Properties("cn")(0)
                If result.Properties("mail").Count > 0 Then adUsers.Last.email = result.Properties("mail")(0)
                If result.Properties("samAccountName").Count > 0 Then adUsers.Last.ad_username = result.Properties("samAccountName")(0)
                If result.Properties("profilePath").Count > 0 Then adUsers.Last.profilePath = result.Properties("profilePath")(0)
                If result.Properties("homeDirectory").Count > 0 Then adUsers.Last.HomePath = result.Properties("homeDirectory")(0)
                If result.Properties("homeDrive").Count > 0 Then adUsers.Last.HomeDriveLetter = result.Properties("homeDrive")(0)
                If result.Properties("employeeID").Count > 0 Then adUsers.Last.employeeID = result.Properties("employeeID")(0)
                If result.Properties("employeeNumber").Count > 0 Then adUsers.Last.employeeNumber = result.Properties("employeeNumber")(0)
                If result.Properties("userAccountControl").Count > 0 Then adUsers.Last.userAccountControl = result.Properties("userAccountControl")(0)
                If result.Properties("distinguishedName").Count > 0 Then adUsers.Last.distinguishedName = result.Properties("distinguishedName")(0)


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

    Function GetADStudents(ldapDirectoryEntry As String)

        Dim dirEntry As New DirectoryEntry(ldapDirectoryEntry)
        'Setting username & password to Nothing forces
        'the connection to use your logon credentials
        dirEntry.Username = Nothing
        dirEntry.Password = Nothing
        'Always use a secure connection
        dirEntry.AuthenticationType = AuthenticationTypes.Secure
        dirEntry.RefreshCache()
        Return dirEntry






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

                adUsers.Last.adObject = result

            Next
            Return adUsers
        End Using
    End Function



    Sub createUsers(ByVal objUsersToAdd As List(Of user), ByVal config As configSettings, conn As MySqlConnection)

        Dim emailsToSend As New List(Of emailNotification)

        For Each objUserToAdd In objUsersToAdd
            If objUserToAdd.ad_username <> "" Then



                Dim objUser As DirectoryEntry
                Dim strDisplayName As String        '
                Dim intEmployeeID As Integer
                Dim strUser As String               ' User to create.
                Dim strUserPrincipalName As String  ' Principal name of user.
                Dim strDescription As String
                Dim intEmployeeNumber As Integer
                Dim strHomeDirectory As String
                Dim strMail As String

                Dim strExt12 As String
                Dim strExt11 As String
                Dim strExt10 As String
                Dim strExt9 As String
                Dim strExt8 As String
                Dim strExt7 As String
                Dim strExt6 As String
                Dim strExt5 As String
                Dim strExt4 As String
                Dim strExt3 As String
                Dim strExt2 As String
                Dim strExt1 As String
                Dim strExt13 As String


                'common properties for all user types
                Console.WriteLine("Creating: " & objUserToAdd.displayName)
                strDisplayName = objUserToAdd.displayName

                Console.WriteLine("EmployeeID: " & objUserToAdd.employeeID)
                intEmployeeID = objUserToAdd.employeeID

                Console.WriteLine("EmployeeNumber: " & objUserToAdd.employeeNumber)
                intEmployeeNumber = objUserToAdd.employeeNumber


                '            Try

                Select Case objUserToAdd.userType
                    Case "Student"

                        Console.WriteLine("CN: " & "CN=" & objUserToAdd.displayName & ",OU=" & objUserToAdd.classOf.ToString & ",OU=Student Users")
                        strUser = "CN=" & objUserToAdd.displayName & ",OU=" & objUserToAdd.classOf.ToString & ",OU=Student Users"

                        Console.WriteLine("UPN: " & objUserToAdd.ad_username & config.studentDomainName)
                        strUserPrincipalName = objUserToAdd.ad_username '& config.domain

                        Console.WriteLine("Class of: " & "Class of " & objUserToAdd.classOf & " Barcode: ")
                        strDescription = "Class of " & objUserToAdd.classOf & " Barcode: "
                        strMail = objUserToAdd.ad_username & config.studentDomainName

                    Case "Staff"
                        Console.WriteLine("CN: " & "CN=" & objUserToAdd.displayName & ",OU=Current Staff,OU=Staff Users")
                        strUser = "CN=" & objUserToAdd.displayName & ",OU=Current Staff,OU=Staff Users"
                        Console.WriteLine("UPN: " & objUserToAdd.ad_username & config.staffDomainName)
                        strUserPrincipalName = objUserToAdd.ad_username '& config.domain
                        strHomeDirectory = config.staffHomePath & objUserToAdd.ad_username
                        strMail = objUserToAdd.ad_username & config.staffDomainName

                    Case "Parent"
                        strUser = "CN=" & objUserToAdd.ad_username & "," & config.parentOU
                        strDescription = objUserToAdd.firstName & " " & objUserToAdd.surname
                        strDisplayName = objUserToAdd.ad_username
                        strUserPrincipalName = objUserToAdd.ad_username '& config.domain
                        strMail = objUserToAdd.ad_username & config.parentDomainName


                    Case Else
                        'Do Else

                End Select

                Console.WriteLine("Create:  {0}", strUser)

                ' Create User.


                Using dirEntry As DirectoryEntry = GetDirectoryEntry(config.ldapDirectoryEntry)
                    dirEntry.RefreshCache()

                    objUser = dirEntry.Children.Add(strUser, "user")
                    '      objUser.Properties("displayName").Add(strDisplayName)




                    If strUserPrincipalName <> "" Then
                        objUser.Properties("mail").Add(strMail)
                    End If



                    If strUserPrincipalName <> "" Then
                        objUser.Properties("homeDrive").Add("H:")
                    End If

                    If strHomeDirectory <> "" Then
                        objUser.Properties("homeDirectory").Add(strHomeDirectory)
                    End If


                    If strUserPrincipalName <> "" Then
                        objUser.Properties("proxyAddresses").Add("SMTP:" & strMail)
                    End If

                    If objUserToAdd.surname <> "" Then
                        objUser.Properties("sn").Add(objUserToAdd.surname)
                    End If
                    If objUserToAdd.ad_username <> "" Then
                        objUser.Properties("samAccountName").Add(objUserToAdd.ad_username)
                    End If
                    If objUserToAdd.firstName <> "" Then
                        objUser.Properties("givenName").Add(objUserToAdd.firstName)
                    End If

                    objUser.Properties("EmployeeID").Add(intEmployeeID)
                    objUser.Properties("EmployeeNumber").Add(intEmployeeNumber)

                    If strUserPrincipalName <> "" Then
                        objUser.Properties("userPrincipalName").Add(strUserPrincipalName)
                    End If
                    If strDisplayName <> "" Then
                        objUser.Properties("displayName").Add(strDisplayName)
                    End If
                    If strDescription <> "" Then
                        objUser.Properties("description").Add(strDescription)
                    End If



                    If config.applyChanges Then
                        objUser.CommitChanges()
                    End If

					Dim pwSet = 0
					While pwSet = 0
						Try
						objUserToAdd.password = createPassword()                   'New Object() {createPassword()}
						If config.applyChanges Then
							objUser.Invoke("setPassword", objUserToAdd.password)
							objUser.CommitChanges()
						End If
							pwSet = 1
						Catch ex As Exception
							pwSet = 0
						End Try
					End While


					'512	Enabled Account
					'514	Disabled Account
					'544	Enabled, Password Not Required
					'546	Disabled, Password Not Required
					'66048	Enabled, Password Doesn't Expire
					'66050	Disabled, Password Doesn't Expire
					'66080	Enabled, Password Doesn't Expire & Not Required
					'66082	Disabled, Password Doesn't Expire & Not Required
					'262656	Enabled, Smartcard Required
					'262658	Disabled, Smartcard Required
					'262688	Enabled, Smartcard Required, Password Not Required
					'262690	Disabled, Smartcard Required, Password Not Required
					'328192	Enabled, Smartcard Required, Password Doesn't Expire
					'328194	Disabled, Smartcard Required, Password Doesn't Expire
					'328224	Enabled, Smartcard Required, Password Doesn't Expire & Not Required
					'328226	Disabled, Smartcard Required, Password Doesn't Expire & Not Required



					Const ADS_UF_ACCOUNTDISABLE = &H10200
                    objUser.Properties("userAccountControl").Value = ADS_UF_ACCOUNTDISABLE
                    If config.applyChanges Then
                        objUser.CommitChanges()
                    End If



                    ' Output User attributes.

                End Using

                Console.WriteLine("Success: Create succeeded.")
                Console.WriteLine("Name:    {0}", objUser.Name)
                Console.WriteLine("         {0}",
                        objUser.Properties("displayName").Value)
                Console.WriteLine("         {0}",
                        objUser.Properties("userPrincipalName").Value)
                Console.WriteLine("")

                For Each mailTo In objUserToAdd.mailTo

                    Dim duplicate As Boolean = False
                    For Each message In emailsToSend
                        If message.mailTo = mailTo Then
                            duplicate = True
                            Select Case objUserToAdd.userType
                                Case "Student"
                                    Dim strMessageBody As String
                                    strMessageBody = "Student account created:  " & objUser.Properties("displayName").Value.ToString & vbCrLf & "Username:" & objUser.Properties("samAccountName").Value.ToString & vbCrLf & "Password:" & objUserToAdd.password.ToString & vbCrLf & "Class Of:" & objUserToAdd.classOf.ToString & vbCrLf & "Start Date: " & objUserToAdd.startDate.ToString & vbCrLf & vbCrLf
                                    message.body = message.body & strMessageBody
                                Case "Parent"
                                    message.body = message.body & "Parent account created:  " & objUser.Properties("description").Value & vbCrLf & "Username:" & objUser.Properties("samAccountName").Value & vbCrLf & "Password:" & objUserToAdd.password.ToString & vbCrLf & vbCrLf
                                Case "Staff"
                                    message.body = message.body & "Staff account created:  " & objUser.Properties("description").Value & vbCrLf & "Username:" & objUser.Properties("samAccountName").Value & vbCrLf & "Password:" & objUserToAdd.password.ToString & vbCrLf & vbCrLf
                            End Select

                        End If
                    Next

                    If duplicate = False Then

                        emailsToSend.Add(New emailNotification)
                        emailsToSend.Last.mailTo = mailTo

                        Select Case objUserToAdd.userType
                            Case "Student"
                                emailsToSend.Last.body = "Student account created:  " & objUser.Properties("displayName").Value.ToString & vbCrLf & "Username:" & objUser.Properties("samAccountName").Value.ToString & vbCrLf & "Password:" & objUserToAdd.password.ToString & vbCrLf & "Class Of:" & objUserToAdd.classOf.ToString & vbCrLf & "Start Date: " & objUserToAdd.startDate.ToString & vbCrLf & vbCrLf
                            Case "Parent"
                                emailsToSend.Last.body = "Parent account created:  " & objUser.Properties("description").Value & vbCrLf & "Username:" & objUser.Properties("samAccountName").Value & vbCrLf & "Password:" & objUserToAdd.password.ToString & vbCrLf & vbCrLf
                            Case "Staff"
                                emailsToSend.Last.body = "Staff account created:  " & objUser.Properties("description").Value & vbCrLf & "Username:" & objUser.Properties("samAccountName").Value & vbCrLf & "Password:" & objUserToAdd.password.ToString & vbCrLf & vbCrLf


                        End Select
                    End If
                Next

            End If

            If objUserToAdd.userType = "Student" Then
                addUsertoMySQL(conn, objUserToAdd)
            End If

        Next

        For Each message In emailsToSend
            Console.WriteLine("Sending email to: " & message.mailTo)
        Next



        sendEmails(config, emailsToSend)
    End Sub

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


            If IsDBNull(user.startDate) Then
                'do nothing
            Else
                If IsDBNull(user.endDate) Then
                    ReturnUsers.Add(user)
                Else
                    If user.endDate > Date.Now() And user.startDate < (Date.Now.AddDays(config.daysInAdvanceToCreateAccounts)) Then
                        ReturnUsers.Add(user)
                    End If
                End If
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

                        user.surname = Replace(user.surname, "&#039;", "")
                        user.firstName = Replace(user.firstName, "&#039;", "")
                        strUsername = rgx.Replace(user.surname & Left(user.firstName, i), "").ToLower

                        Console.WriteLine("Trying " & strUsername & "...")
                        Dim duplicate As Boolean
                        duplicate = False
                        Dim a As Integer = 1
                        For Each adUser In adusers

                            CONSOLE__WRITE(String.Format("Checking for duplicates {0} of {1}", a, adusers.Count))
                            Try
                                adUser.ad_username = adUser.ad_username.ToLower
                            Catch ex As Exception

                            End Try

                            If strUsername = adUser.ad_username Then
                                duplicate = True
                            Else
                                'duplicate = False
                            End If
                            a = a + 1
                        Next
                        If duplicate = False Then
                            availableNameFound = True
                            user.ad_username = strUsername
                        End If

                        i = i + 1
                    End While

                    If user.ad_username = Nothing Then
                        Console.WriteLine("No valid username available for " & user.firstName & " " & user.surname)
                    Else
                        Console.WriteLine(user.firstName & " " & user.surname & " will be created as " & user.ad_username)
                    End If


                Case "Staff"
                    Dim rgx As New Regex("[^a-zA-Z]")
                    Dim availableNameFound As Boolean = False
                    Dim i As Integer = 1

                    While availableNameFound = False And i <= user.surname.Length

                        strUsername = rgx.Replace(user.firstName & Left(user.surname, i), "").ToLower
                        Console.WriteLine("Trying " & strUsername & "...")
                        Dim duplicate As Boolean
                        duplicate = False
                        Dim a As Integer = 1
                        For Each adUser In adusers

                            CONSOLE__WRITE(String.Format("Checking for duplicates {0} of {1}", a, adusers.Count))
                            Try
                                adUser.ad_username = adUser.ad_username.ToLower
                            Catch ex As Exception

                            End Try

                            If strUsername = adUser.ad_username Then
                                duplicate = True
                            Else
                                'duplicate = False
                            End If
                            a = a + 1
                        Next
                        If duplicate = False Then
                            availableNameFound = True
                            user.ad_username = strUsername
                        End If

                        i = i + 1
                    End While

                    If user.ad_username = Nothing Then
                        Console.WriteLine("No valid username available for " & user.firstName & " " & user.surname)
                    Else
                        Console.WriteLine(user.firstName & " " & user.surname & " will be created as " & user.ad_username)
                    End If


                Case "Parent"

                    Dim rgx As New Regex("[^a-zA-Z0-9]")
                    user.ad_username = rgx.Replace(Left(user.surname, 5) & user.employeeID, "").ToLower
                    Console.WriteLine(user.firstName & " " & user.surname & " will be created as " & user.ad_username)
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
		Dim PasswordPosition As Integer = 6   'original was 2
		Select Case PasswordPosition
            Case 0 : Return getWord(wordlist) & getWord(wordlist) & rndNumber
            Case 1 : Return rndNumber & getWord(wordlist) & getWord(wordlist)
            Case 2 : Return Mixedcase(getWord(wordlist)) & rndNumber & Mixedcase(getWord(wordlist))
            Case 3 : Return getWord(wordlist) & getWord(wordlist) & rndNumber
            Case 4 : Return rndNumber & getWord(wordlist) & getWord(wordlist)
            Case 5 : Return getWord(wordlist) & rndNumber & getWord(wordlist)
			Case 6 : Return getWord(wordlist) & " " & getWord(wordlist) & " " & getWord(wordlist)
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

    Function Mixedcase(ByVal Word As String) As String
        If Word.Length = 0 Then Return Word
        If Word.Length = 1 Then Return UCase(Word)
        Return Word.Substring(0, 1).ToUpper & Word.Substring(1).ToLower

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

    Sub sendMail(ByVal config As configSettings, ByVal subject As String, ByVal body As String, ByVal mailTo As String)

        Dim mailClient = New SmtpClient(config.serverAddress)
        mailClient.Port = config.serverPort
        mailClient.EnableSsl = config.enableSSL

        Dim cred = New System.Net.NetworkCredential(config.username, config.password)
        mailClient.Credentials = cred

        Dim Message = New MailMessage()

        Message.From = New MailAddress(config.username, config.displayName)

        Message.To.Add(mailTo)

        Message.Subject = subject
        Message.Body = body

        If Not config.applyChanges Then
            Message.Body = "---Test run only - no accounts created---" & vbCrLf & "These are NOT real accounts. Do not give these details to parents or students. This is a test ONLY" & vbCrLf & Message.Body
            Message.Body = Message.Body & "---Test run only - no accounts created---"
        End If

        'mailClient.Timeout = 600000

        mailClient.Send(Message)

        Message = Nothing
        mailClient = Nothing


    End Sub

    Sub sendEmails(config As configSettings, ByVal emails As List(Of emailNotification))
        For Each mail In emails
            Console.WriteLine("Queueing email " & mail.mailTo)
            sendMail(config, "Account Provisioning", mail.body, mail.mailTo)
        Next
    End Sub

    Function addMailTo(config As configSettings, users As List(Of user))

        For Each user In users
            For Each emailAddress In config.mailToAll
                user.mailTo.Add(emailAddress)
            Next
            Select Case user.currentYear
                Case "K"
                    For Each emailAddress In config.mailToK
                        user.mailTo.Add(emailAddress)
                    Next
                Case "01"
                    For Each emailAddress In config.mailTo1
                        user.mailTo.Add(emailAddress)
                    Next
                Case "02"
                    For Each emailAddress In config.mailTo2
                        user.mailTo.Add(emailAddress)
                    Next
                Case "03"
                    For Each emailAddress In config.mailTo3
                        user.mailTo.Add(emailAddress)
                    Next
                Case "04"
                    For Each emailAddress In config.mailTo4
                        user.mailTo.Add(emailAddress)
                    Next
                Case "05"
                    For Each emailAddress In config.mailTo5
                        user.mailTo.Add(emailAddress)
                    Next
                Case "06"
                    For Each emailAddress In config.mailTo6
                        user.mailTo.Add(emailAddress)
                    Next
                Case "07"
                    For Each emailAddress In config.mailTo7
                        user.mailTo.Add(emailAddress)
                    Next
                Case "08"
                    For Each emailAddress In config.mailTo8
                        user.mailTo.Add(emailAddress)
                    Next
                Case "09"
                    For Each emailAddress In config.mailTo9
                        user.mailTo.Add(emailAddress)
                    Next
                Case "10"
                    For Each emailAddress In config.mailTo10
                        user.mailTo.Add(emailAddress)
                    Next
                Case "11"
                    For Each emailAddress In config.mailTo11
                        user.mailTo.Add(emailAddress)
                    Next
                Case "12"
                    For Each emailAddress In config.mailTo12
                        user.mailTo.Add(emailAddress)
                    Next
            End Select
        Next
        Return users
    End Function

    Function calculateCurrentYears(users As List(Of user))
        For Each user In users
            If Not IsNothing(user) Then

                Select Case user.classOf - Convert.ToInt32(Year(Date.Now))
                    Case 0
                        user.currentYear = "12"
                    Case 1
                        user.currentYear = "11"
                    Case 2
                        user.currentYear = "10"
                    Case 3
                        user.currentYear = "09"
                    Case 4
                        user.currentYear = "08"
                    Case 5
                        user.currentYear = "07"
                    Case 6
                        user.currentYear = "06"
                    Case 7
                        user.currentYear = "05"
                    Case 8
                        user.currentYear = "04"
                    Case 9
                        user.currentYear = "03"
                    Case 10
                        user.currentYear = "02"
                    Case 11
                        user.currentYear = "01"
                    Case 12
                        user.currentYear = "K"
                End Select
            End If
        Next
        Return users
    End Function

    Function getEdumateStaff(config As configSettings)
        Dim ConnectionString As String = config.edumateConnectionString
        Dim commandString As String =
"
SELECT        
contact.firstname,
contact.surname,
staff_employment.start_date,
staff_employment.end_date,
staff.staff_id,
staff.staff_number,
sys_user.username,
contact.email_address,
staff_employment.employment_type_id,
contact.contact_id


FROM            STAFF

INNER JOIN Contact 
  ON staff.contact_id = contact.contact_id 
INNER JOIN staff_employment
  ON staff.staff_id = staff_employment.staff_id
LEFT JOIN sys_user 
  ON contact.contact_id = sys_user.contact_id
"


        Dim users As New List(Of user)



        Using conn As New IBM.Data.DB2.DB2Connection(ConnectionString)
            conn.Open()

            'define the command object to execute
            Dim command As New IBM.Data.DB2.DB2Command(commandString, conn)
            command.Connection = conn
            command.CommandText = commandString

            Dim dr As IBM.Data.DB2.DB2DataReader
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
                    'users.Last.employeeNumber = dr.GetValue(5) (this should be print code, not staff number)
                    users.Last.userType = "Staff"
                    users.Last.displayName = users.Last.firstName & " " & users.Last.surname
                    users.Last.edumateCurrent = 0
                    If Not IsDBNull(dr.GetValue(6)) Then users.Last.edumateUsername = dr.GetValue(6)
                    If Not IsDBNull(dr.GetValue(7)) Then users.Last.edumateEmail = dr.GetValue(7)
                    If Not IsDBNull(dr.GetValue(8)) Then users.Last.employmentType = dr.GetValue(8)
                    users.Last.edumateStaffNumber = dr.GetValue(5)
                    users.Last.contact_id = dr.GetValue(9)

                    If Not IsDBNull(users.Last.startDate) Then

                        If users.Last.startDate < Date.Now.AddDays(10) Then
                            If IsDBNull(users.Last.endDate) Then
                                users.Last.edumateCurrent = 1
                            Else
                                If users.Last.endDate > Date.Now() Then
                                    users.Last.edumateCurrent = 1
                                End If
                            End If
                        End If
                    End If
                End If
            End While
            conn.Close()
        End Using
        Return users
    End Function

    Sub addUsertoMySQL(conn As MySqlConnection, user As user)


        Dim table As String = "student_details"
        Dim studentID As String = user.employeeID
        Dim firstname As String = user.firstName
        Dim surname As String = user.surname
        Dim username As String = user.ad_username
        Dim password As String = user.password
        Dim gradYear As String = user.classOf
        Dim current As String = "1"


        Try
            conn.Open()
        Catch ex As Exception
        End Try





        Dim cmd As New MySqlCommand(String.Format("INSERT INTO `{0}` (`student_id` , `first_name` , `surname` , `username` , `password` , `grad_year` , `current` ) VALUES ('{1}' , '{2}', '{3}', '{4}', '{5}', '{6}', '{7}')", table, studentID, firstname, surname, username, password, gradYear, current), conn)
        cmd.ExecuteNonQuery()






        conn.Close()



    End Sub

    Public Sub connect(conn As MySqlConnection, config As configSettings)
        Dim DatabaseName As String = config.mySQLDatabaseName
        Dim server As String = config.mySQLserver
        Dim userName As String = config.mySQLUserName
        Dim password As String = config.mySQLPassword
        If Not conn Is Nothing Then conn.Close()
        conn.ConnectionString = String.Format("server={0}; user id={1}; password={2}; database={3}; pooling=false", server, userName, password, DatabaseName)
        Try
            conn.Open()

            'MsgBox("Connected")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        conn.Close()

    End Sub

    Function ValidateActiveDirectoryLogin(ByVal Domain As String, ByVal Username As String, ByVal Password As String) As Boolean




        Dim Success As Boolean = False
        Dim Entry As New System.DirectoryServices.DirectoryEntry("LDAP://" & Domain, Username, Password)
        Dim Searcher As New System.DirectoryServices.DirectorySearcher(Entry)
        Searcher.SearchScope = DirectoryServices.SearchScope.OneLevel
        Try
            Dim Results As System.DirectoryServices.SearchResult = Searcher.FindOne
            Success = Not (Results Is Nothing)
        Catch
            Success = False
        End Try
        Return Success
    End Function

    Function getMySQLStudents(conn)

        Dim userTable As String = "student_details"

        Dim users As New List(Of user)

        Dim commandstring As String = ("SELECT student_id, first_name, surname, username, password, grad_year,current FROM " & userTable)
        Dim command As New MySqlCommand(commandstring, conn)

        conn.open

        command.Connection = conn
        command.CommandText = commandstring

        Dim dr As MySqlDataReader
        dr = command.ExecuteReader

        Dim i As Integer = 0
        While dr.Read()
            If Not dr.IsDBNull(0) Then
                users.Add(New user)

                users.Last.employeeID = dr.GetValue(0)
                users.Last.firstName = dr.GetValue(1)
                users.Last.surname = dr.GetValue(2)
                users.Last.ad_username = dr.GetValue(3)
                users.Last.password = dr.GetValue(4)

                If Not dr.IsDBNull(5) Then users.Last.classOf = dr.GetValue(5)
                users.Last.enabled = dr.GetValue(6)
                users.Last.userType = "Student"
                users.Last.displayName = users.Last.firstName & " " & users.Last.surname
            End If
        End While
        conn.Close()
        Return users

    End Function

    Function removeInvalidPasswords(users As List(Of user), domain As String)

        For Each user In users
            If ValidateActiveDirectoryLogin(domain, user.ad_username, user.password) Then

            Else
                user.password = "unknown"
            End If
        Next


        Return users
    End Function

    Sub updatePasswordsInMysql(users As List(Of user), conn As MySqlConnection)

        Dim userTable As String = "student_details"
        Try
            conn.Open()
        Catch ex As Exception
        End Try
        For Each user In users
            Dim cmd As New MySqlCommand(String.Format("UPDATE `{0}` SET password  = '{1}' where student_id = '{2}' ", userTable, user.password, user.employeeID), conn)
            cmd.ExecuteNonQuery()

        Next
        conn.Close()
        For Each user In users
        Next
    End Sub

    Sub updateCurrentFlags(mySQLUsers As List(Of user), edumateStudents As List(Of user), conn As MySqlConnection, adUsers As List(Of user))
        Dim usertable As String = "student_details"

        Try
            conn.Open()
        Catch ex As Exception
        End Try


        For Each user In mySQLUsers
            Dim current As String = 0
            For Each student In edumateStudents
                If user.employeeID = student.employeeID Then
                    current = 1
                End If
            Next
            If Not user.enabled = current Then
                Dim cmd As New MySqlCommand(String.Format("UPDATE `{0}` SET current  = '{1}' where student_id = '{2}' ", usertable, current, user.employeeID), conn)
                cmd.ExecuteNonQuery()
            End If



        Next
        conn.Close()
    End Sub

    Function addUsernamesToUsers(users As List(Of user), adUsers As List(Of user))

        For Each user In users
            For Each adUser In adUsers
                If user.employeeID = adUser.employeeID Then
                    user.ad_username = adUser.ad_username
                    user.distinguishedName = adUser.distinguishedName
                End If
            Next
        Next
        Return users

    End Function

    Sub updateMSQLDetails(usersToUpdate As List(Of user), mySQLUsers As List(Of user), conn As MySqlConnection)

        Dim usertable As String = "student_details"

        Try
            conn.Open()
        Catch ex As Exception
        End Try


        For Each user In usersToUpdate
            For Each mySQLUser In mySQLUsers
                If user.employeeID = mySQLUser.employeeID Then
                    Dim cmd As MySqlCommand
                    If user.ad_username = mySQLUser.ad_username Then
                    Else
                        cmd = New MySqlCommand(String.Format("UPDATE `{0}` SET username  = '{1}' where student_id = '{2}' ", usertable, user.ad_username, user.employeeID), conn)
                        cmd.ExecuteNonQuery()
                    End If
                    cmd = New MySqlCommand(String.Format("UPDATE `{0}` SET student_number  = '{1}' where student_id = '{2}' ", usertable, user.employeeNumber, user.employeeID), conn)
                    cmd.ExecuteNonQuery()

                    cmd = New MySqlCommand(String.Format("UPDATE `{0}` SET current_year  = '{1}' where student_id = '{2}' ", usertable, user.currentYear, user.employeeID), conn)
                    cmd.ExecuteNonQuery()

                    cmd = New MySqlCommand(String.Format("UPDATE `{0}` SET dob  = '{1}' where student_id = '{2}' ", usertable, user.dob, user.employeeID), conn)
                    cmd.ExecuteNonQuery()

                    cmd = New MySqlCommand(String.Format("UPDATE `{0}` SET barcode  = '{1}' where student_id = '{2}' ", usertable, user.libraryCard, user.employeeID), conn)
                    cmd.ExecuteNonQuery()

                    cmd = New MySqlCommand(String.Format("UPDATE `{0}` SET roll_class  = '{1}' where student_id = '{2}' ", usertable, user.rollClass, user.employeeID), conn)
                    cmd.ExecuteNonQuery()

                    cmd = New MySqlCommand(String.Format("UPDATE `{0}` SET bos  = '{1}' where student_id = '{2}' ", usertable, user.bosNumber, user.employeeID), conn)
                    cmd.ExecuteNonQuery()

                    cmd = New MySqlCommand(String.Format("UPDATE `{0}` SET enrolled_classes  = '{1}' where student_id = '{2}' ", usertable, user.enrolledClasses, user.employeeID), conn)
                    cmd.ExecuteNonQuery()

                    cmd = New MySqlCommand(String.Format("UPDATE `{0}` SET adFirstname  = '{1}' where student_id = '{2}' ", usertable, user.adFirstname, user.employeeID), conn)
                    cmd.ExecuteNonQuery()

                    cmd = New MySqlCommand(String.Format("UPDATE `{0}` SET adSurname  = '{1}' where student_id = '{2}' ", usertable, user.adSurname, user.employeeID), conn)
                    cmd.ExecuteNonQuery()

                End If
            Next

        Next
        conn.Close()
    End Sub

    Sub updateEmployeeNumbers(adUsers As List(Of user), edumateStudents As List(Of user), config As configSettings)



        For Each student In edumateStudents
            For Each adUser In adUsers
                If student.employeeID = adUser.employeeID Then

                    If adUser.employeeNumber = "" Then

                        Using user As New DirectoryEntry("LDAP://" & adUser.distinguishedName)
                            'Setting username & password to Nothing forces
                            'the connection to use your logon credentials
                            user.Username = Nothing
                            user.Password = Nothing
                            'Always use a secure connection
                            user.AuthenticationType = AuthenticationTypes.Secure
                            ' user.RefreshCache()

                            user.Properties("employeeNumber").Add(student.employeeNumber)



                            user.CommitChanges()

                        End Using

                    End If
                End If
            Next
        Next


    End Sub

    Sub getStudentAccountsToDisable()

    End Sub

    Sub updateParentStudents(parents As List(Of user), config As configSettings)


        Dim dirEntry As DirectoryEntry

        Console.WriteLine("Connecting to AD...")
        dirEntry = GetDirectoryEntry(config.ldapDirectoryEntry)

        Dim adUsers As List(Of user)
        Console.WriteLine("Loading AD users...")
        Console.WriteLine("")
        Console.WriteLine("")
        adUsers = getADUsers(dirEntry)


        For Each parent In parents

            Dim strExt12 As String
            Dim strExt11 As String
            Dim strExt10 As String
            Dim strExt9 As String
            Dim strExt8 As String
            Dim strExt7 As String
            Dim strExt6 As String
            Dim strExt5 As String
            Dim strExt4 As String
            Dim strExt3 As String
            Dim strExt2 As String
            Dim strExt1 As String
            Dim strExt13 As String

            For Each adUser In adUsers
                If parent.employeeID = adUser.employeeID Then
                    strExt12 = getChildFromChildren(parent.children, "12")
                    strExt11 = getChildFromChildren(parent.children, "11")
                    strExt10 = getChildFromChildren(parent.children, "10")
                    strExt9 = getChildFromChildren(parent.children, "09")
                    strExt8 = getChildFromChildren(parent.children, "08")
                    strExt7 = getChildFromChildren(parent.children, "07")
                    strExt6 = getChildFromChildren(parent.children, "06")
                    strExt5 = getChildFromChildren(parent.children, "05")
                    strExt4 = getChildFromChildren(parent.children, "04")
                    strExt3 = getChildFromChildren(parent.children, "03")
                    strExt2 = getChildFromChildren(parent.children, "02")
                    strExt1 = getChildFromChildren(parent.children, "01")
                    strExt13 = getChildFromChildren(parent.children, "K")

                    '[If CInt(strExt12 & strExt11 & strExt10 & strExt9 & strExt8 & strExt7 & strExt6 & strExt5 & strExt4 & strExt3 & strExt2 & strExt1 & strExt13) > 1 Then

                    Using user As New DirectoryEntry("LDAP://ryze.i.ofgs.nsw.edu.au/" & adUser.distinguishedName)
                        'MsgBox(adUser.distinguishedName)
                        'Setting username & password to Nothing forces
                        'the connection to use your logon credentials
                        user.Username = Nothing
                        user.Password = Nothing
                        'Always use a secure connection
                        user.AuthenticationType = AuthenticationTypes.Secure
                        ' user.RefreshCache()

                        If CInt(strExt12) > 1 Then
                            If user.Properties("extensionAttribute12").Count > 0 Then
                                user.Properties("extensionAttribute12")(0) = (strExt12)
                            Else
                                user.Properties("extensionAttribute12").Add(strExt12)
                            End If
                        Else
                            If user.Properties("extensionAttribute12").Count > 0 Then
                                user.Properties("extensionAttribute12").Clear()
                            End If
                        End If

                        If CInt(strExt11) > 1 Then
                            If user.Properties("extensionAttribute11").Count > 0 Then
                                user.Properties("extensionAttribute11")(0) = (strExt11)
                            Else
                                user.Properties("extensionAttribute11").Add(strExt11)
                            End If
                        Else
                            If user.Properties("extensionAttribute11").Count > 0 Then
                                user.Properties("extensionAttribute11").Clear()
                            End If
                        End If

                        If CInt(strExt10) > 1 Then
                            If user.Properties("extensionAttribute10").Count > 0 Then
                                user.Properties("extensionAttribute10")(0) = (strExt10)
                            Else
                                user.Properties("extensionAttribute10").Add(strExt10)
                            End If
                        Else
                            If user.Properties("extensionAttribute10").Count > 0 Then
                                user.Properties("extensionAttribute10").Clear()
                            End If
                        End If

                        If CInt(strExt9) > 1 Then
                            If user.Properties("extensionAttribute9").Count > 0 Then
                                user.Properties("extensionAttribute9")(0) = (strExt9)
                            Else
                                user.Properties("extensionAttribute9").Add(strExt9)
                            End If
                        Else
                            If user.Properties("extensionAttribute9").Count > 0 Then
                                user.Properties("extensionAttribute9").Clear()
                            End If
                        End If

                        If CInt(strExt8) > 1 Then
                            If user.Properties("extensionAttribute8").Count > 0 Then
                                user.Properties("extensionAttribute8")(0) = (strExt8)
                            Else
                                user.Properties("extensionAttribute8").Add(strExt8)
                            End If
                        Else
                            If user.Properties("extensionAttribute8").Count > 0 Then
                                user.Properties("extensionAttribute8").Clear()
                            End If
                        End If

                        If CInt(strExt7) > 1 Then
                            If user.Properties("extensionAttribute7").Count > 0 Then
                                user.Properties("extensionAttribute7")(0) = (strExt7)
                            Else
                                user.Properties("extensionAttribute7").Add(strExt7)
                            End If
                        Else
                            If user.Properties("extensionAttribute7").Count > 0 Then
                                user.Properties("extensionAttribute7").Clear()
                            End If
                        End If

                        If CInt(strExt6) > 1 Then
                            If user.Properties("extensionAttribute6").Count > 0 Then
                                user.Properties("extensionAttribute6")(0) = (strExt6)
                            Else
                                user.Properties("extensionAttribute6").Add(strExt6)
                            End If
                        Else
                            If user.Properties("extensionAttribute6").Count > 0 Then
                                user.Properties("extensionAttribute6").Clear()
                            End If
                        End If

                        If CInt(strExt5) > 1 Then
                            If user.Properties("extensionAttribute5").Count > 0 Then
                                user.Properties("extensionAttribute5")(0) = (strExt5)
                            Else
                                user.Properties("extensionAttribute5").Add(strExt5)
                            End If
                        Else
                            If user.Properties("extensionAttribute5").Count > 0 Then
                                user.Properties("extensionAttribute5").Clear()
                            End If
                        End If

                        If CInt(strExt4) > 1 Then
                            If user.Properties("extensionAttribute4").Count > 0 Then
                                user.Properties("extensionAttribute4")(0) = (strExt4)
                            Else
                                user.Properties("extensionAttribute4").Add(strExt4)
                            End If
                        Else
                            If user.Properties("extensionAttribute4").Count > 0 Then
                                user.Properties("extensionAttribute4").Clear()
                            End If
                        End If

                        If CInt(strExt3) > 1 Then
                            If user.Properties("extensionAttribute3").Count > 0 Then
                                user.Properties("extensionAttribute3")(0) = (strExt3)
                            Else
                                user.Properties("extensionAttribute3").Add(strExt3)
                            End If
                        Else
                            If user.Properties("extensionAttribute3").Count > 0 Then
                                user.Properties("extensionAttribute3").Clear()
                            End If
                        End If

                        If CInt(strExt2) > 1 Then
                            If user.Properties("extensionAttribute2").Count > 0 Then
                                user.Properties("extensionAttribute2")(0) = (strExt2)
                            Else
                                user.Properties("extensionAttribute2").Add(strExt2)
                            End If
                        Else
                            If user.Properties("extensionAttribute2").Count > 0 Then
                                user.Properties("extensionAttribute2").Clear()
                            End If
                        End If

                        If CInt(strExt1) > 1 Then
                            If user.Properties("extensionAttribute1").Count > 0 Then
                                user.Properties("extensionAttribute1")(0) = (strExt1)
                            Else
                                user.Properties("extensionAttribute1").Add(strExt1)
                            End If
                        Else
                            If user.Properties("extensionAttribute1").Count > 0 Then
                                user.Properties("extensionAttribute1").Clear()
                            End If
                        End If

                        If CInt(strExt13) > 1 Then
                            If user.Properties("extensionAttribute13").Count > 0 Then
                                user.Properties("extensionAttribute13")(0) = (strExt13)
                            Else
                                user.Properties("extensionAttribute13").Add(strExt13)
                            End If
                        Else
                            If user.Properties("extensionAttribute13").Count > 0 Then
                                user.Properties("extensionAttribute13").Clear()
                            End If
                        End If





                        user.CommitChanges()

                    End Using
                End If
                ' End If

            Next
        Next




    End Sub

    Function getChildFromChildren(children As List(Of user), yearToFind As String)

        children = calculateCurrentYears(children)

        Dim found As Boolean = False

        For Each child In children
            '  MsgBox("CY: " & child.currentYear & "YTF: " & yearToFind)

            Try
                If child.currentYear = yearToFind Then
                    Return child.employeeID
                    found = True
                    '    MsgBox("FoundMatch")
                End If
            Catch
            End Try
        Next

        If found = False Then
            Return "0"
        End If
    End Function





    Function ddMMYYYY_to_yyyyMMdd(inString As String)
        ddMMYYYY_to_yyyyMMdd = Strings.Right(inString, 4) & "-" & Left(Mid(inString, Strings.InStr(inString, "/") + 1), 2) & "-" & Left(inString, InStr(inString, "/") - 1)

    End Function








    Function getUsernameFromID(userID As String, adusers As List(Of user))
        For Each user In adusers
            If user.employeeID = userID Then
                Return user.ad_username
            End If
        Next
        Return "noUsername"
    End Function

    Function getMySQLStaff(conn)

        Dim userTable As String = "staff_details"

        Dim users As New List(Of user)

        Dim commandstring As String = ("SELECT staff_id, surname, firstname, ad_username, edumate_username, edumate_current, ad_active, ad_email, edumate_email, smtp_proxy_set, init_password, staff_number, distinguished_name, edumate_login_active, edumate_start_date, edumate_end_date FROM " & userTable)
        Dim command As New MySqlCommand(commandstring, conn)

        conn.open

        command.Connection = conn
        command.CommandText = commandstring

        Dim dr As MySqlDataReader
        dr = command.ExecuteReader

        Dim i As Integer = 0
        While dr.Read()
            If Not dr.IsDBNull(0) Then
                users.Add(New user)
                users.Last.employeeID = dr.GetValue(0)
                users.Last.surname = dr.GetValue(1)
                users.Last.firstName = dr.GetValue(2)
                users.Last.ad_username = dr.GetValue(3)
                users.Last.edumateUsername = dr.GetValue(4)
                users.Last.edumateCurrent = dr.GetValue(5)
                users.Last.enabled = dr.GetValue(6)
                users.Last.email = dr.GetValue(7)
                users.Last.edumateEmail = dr.GetValue(8)
                users.Last.smtpProxy = dr.GetValue(9)
                users.Last.password = dr.GetValue(10)
                users.Last.employeeNumber = dr.GetValue(11)
                users.Last.distinguishedName = dr.GetValue(12)
                users.Last.edumateLoginActive = dr.GetValue(13)
                users.Last.startDate = dr.GetValue(14)
                users.Last.endDate = dr.GetValue(15)
                users.Last.userType = "Staff"
            End If
        End While
        conn.Close()
        Return users

    End Function

    Function addUserTypeToAdUsers(users As List(Of user))
        For Each user In users
            If user.distinguishedName.Contains("OU=Admin") Then
                user.userType = "Staff"
            End If
            If user.distinguishedName.Contains("OU=Current Staff") Then
                user.userType = "Staff"
            End If
            If user.distinguishedName.Contains("OU=Former Staff") Then
                user.userType = "Former Staff"
            End If
            If user.distinguishedName.Contains("OU=Grounds") Then
                user.userType = "Staff"
            End If
            If user.distinguishedName.Contains("OU=Teachers") Then
                user.userType = "Staff"
            End If
            If user.distinguishedName.Contains("OU=Student Users") Then
                user.userType = "Student"
            End If
            If user.distinguishedName.Contains("OU=@ofgsfamily.com") Then
                user.userType = "Parent"
            End If
            If user.distinguishedName.Contains("OU=Casual Staff") Then
                user.userType = "Staff"
            End If
			If user.distinguishedName.Contains("OU=Music") Then
				user.userType = "Staff"
			End If

		Next

        Return users
    End Function

    Sub updateStaffDatabase(config As configSettings)

        Dim conn As New MySqlConnection
        connect(conn, config)
        Dim dirEntry As DirectoryEntry



        Console.WriteLine("Connecting to AD...")
        dirEntry = GetDirectoryEntry(config.ldapDirectoryEntry)
        Console.WriteLine(Chr(8) & "Done")
        Dim adUsers As List(Of user)
        Console.WriteLine("Loading AD users for staff DB...")
        adUsers = getADUsers(dirEntry)
        Console.WriteLine(Chr(8) & "Done")
        Console.WriteLine("Adding user types to AD Users...")
        adUsers = addUserTypeToAdUsers(adUsers)
        Console.WriteLine(Chr(8) & "Done")

        Console.WriteLine("Loading mySQL Staff...")
        Dim mySQLUsers As List(Of user)
        mySQLUsers = getMySQLStaff(conn)
        Console.WriteLine(Chr(8) & "Done")

        Console.WriteLine("Loading edumate Staff...")
        Dim edumateUsers As List(Of user)
        edumateUsers = getEdumateStaff(config)
        Console.WriteLine(Chr(8) & "Done")

        Console.WriteLine("Adding edumate details to AD staff...")
        adUsers = addEdumateDetailsToAdUsers(adUsers, edumateUsers)
        adUsers = getEdumateGroups(adUsers, config)
        Console.WriteLine(Chr(8) & "Done")

        Console.WriteLine("Inserting staff to mySQL database...")
        For Each aduser In adUsers
            If aduser.userType = "Staff" Then

                Dim found As Boolean = False
                For Each mysqlUser In mySQLUsers
                    If aduser.employeeID = mysqlUser.employeeID Then
                        found = True
                    End If
                Next
                If found = False Then
                    insertUserToStaffDB(conn, aduser, config.tutorGroupID, config.danceTutorGroupID)
                End If
            End If

        Next


    End Sub

    Function addEdumateDetailsToAdUsers(adUsers As List(Of user), edumateUsers As List(Of user))
        For Each aduser In adUsers
            For Each edumateUser In edumateUsers
                If aduser.employeeID = edumateUser.employeeID Then
                    aduser.edumateCurrent = edumateUser.edumateCurrent
                    aduser.edumateEmail = edumateUser.edumateEmail
                    aduser.edumateLoginActive = edumateUser.edumateLoginActive
                    aduser.edumateUsername = edumateUser.edumateUsername
                    aduser.edumateStaffNumber = edumateUser.edumateStaffNumber
                    aduser.employmentType = edumateUser.employmentType
                    aduser.contact_id = edumateUser.contact_id
                End If
            Next
        Next
        Return adUsers
    End Function

    Sub insertUserToStaffDB(conn As MySqlConnection, user As user, tutorGroupID As Integer, danceTutorGroupID As Integer)

        Dim table As String = "staff_details"


        '512	Enabled Account
        '514	Disabled Account
        '544	Enabled, Password Not Required
        '546	Disabled, Password Not Required
        '66048	Enabled, Password Doesn't Expire
        '66050	Disabled, Password Doesn't Expire
        '66080	Enabled, Password Doesn't Expire & Not Required
        '66082	Disabled, Password Doesn't Expire & Not Required
        '262656	Enabled, Smartcard Required
        '262658	Disabled, Smartcard Required
        '262688	Enabled, Smartcard Required, Password Not Required
        '262690	Disabled, Smartcard Required, Password Not Required
        '328192	Enabled, Smartcard Required, Password Doesn't Expire
        '328194	Disabled, Smartcard Required, Password Doesn't Expire
        '328224	Enabled, Smartcard Required, Password Doesn't Expire & Not Required
        '328226	Disabled, Smartcard Required, Password Doesn't Expire & Not Required

        Dim accountStatus As String

        Select Case user.userAccountControl
            Case 512
                accountStatus = "Enabled Account"
            Case 514
                accountStatus = "Disabled Account"
            Case 544
                accountStatus = "Enabled, Password Not Required"
            Case 546
                accountStatus = "Disabled, Password Not Required"
            Case 66048
                accountStatus = "Enabled, Password Doesnt Expire"
            Case 66050
                accountStatus = "Disabled, Password Doesnt Expire"
            Case 66080
                accountStatus = "Enabled, Password Doesnt Expire + Not Required"
            Case 66082
                accountStatus = "Disabled, Password Doesnt Expire + Not Required"


        End Select



        Dim musicTutor As Integer = 0
        If Not IsNothing(user.edumateGroupMemberships) Then
            For Each group In user.edumateGroupMemberships


                If group = tutorGroupID Then
                    musicTutor = 1
                End If
                If group = danceTutorGroupID Then
                    musicTutor = 1
                End If

            Next
        End If
        Try
            conn.Open()
        Catch ex As Exception
        End Try
        Dim sanitizedSurname As String
        sanitizedSurname = Replace(user.surname, "'", "\'")
        Dim sanitizedDn
        sanitizedDn = Replace(user.distinguishedName, "'", "\'")
        Dim datePasswordSet As String



        If user.adObject.Properties("pwdLastSet").Count > 0 Then
            datePasswordSet = user.adObject.Properties("pwdLastSet")(0)
        Else
            datePasswordSet = "Never"
        End If







        Dim cmd As New MySqlCommand(String.Format("INSERT INTO `{0}` (`staff_id`,`surname`,`firstname`, `ad_username`,`edumate_username`,`edumate_current`,`ad_active`,`ad_email`,`edumate_email`,`smtp_proxy_set`,`init_password`,`staff_number`,`distinguished_name`,`edumate_login_active`,`edumate_start_date`,`edumate_end_date`,`employment_type`,`edumate_staff_number`,`tutor`,`datePwdSet`) VALUES ('{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}','{17}','{18}','{19}','{20}')", table, user.employeeID, sanitizedSurname, user.firstName, user.ad_username, user.edumateUsername, user.edumateCurrent, accountStatus, user.email, user.edumateEmail, user.smtpProxy, user.password, user.employeeNumber, sanitizedDn, user.edumateLoginActive, user.startDate, user.endDate, user.employmentType, user.edumateStaffNumber, musicTutor, datePasswordSet), conn)
        cmd.ExecuteNonQuery()

        conn.Close()

    End Sub

    Sub updateUserInStaffDB(conn As MySqlConnection, user As user)

    End Sub

    Function getADGroups(dirEntry As DirectoryEntry)
        Using searcher As New DirectorySearcher(dirEntry)
            Dim adUsers As New List(Of user)

            searcher.PropertiesToLoad.Add("cn")
            searcher.PropertiesToLoad.Add("employeeID")
            searcher.PropertiesToLoad.Add("distinguishedName")
            searcher.PropertiesToLoad.Add("mail")
            searcher.PropertiesToLoad.Add("memberof")
            searcher.PropertiesToLoad.Add("userAccountControl")


            searcher.Filter = "(objectCategory=group)"
            searcher.ServerTimeLimit = New TimeSpan(0, 0, 60)
            searcher.SizeLimit = 100000000
            searcher.Asynchronous = False
            searcher.ServerPageTimeLimit = New TimeSpan(0, 0, 60)
            searcher.PageSize = 10000

            Dim queryResults As SearchResultCollection
            queryResults = searcher.FindAll

            Dim result As SearchResult

            For Each result In queryResults

                If result.Properties("members").Count > 0 Then
                    For Each user In result.Properties("members")
                        adUsers.Last.memberOf.Add(user)

                    Next
                End If
            Next
            Return queryResults
        End Using
    End Function

    Sub AddStudentsToYearGroups(users As List(Of user), config As configSettings)

        Dim kindyUsers As New List(Of user)
        Dim Year01Users As New List(Of user)
        Dim Year02Users As New List(Of user)
        Dim Year03Users As New List(Of user)
        Dim Year04Users As New List(Of user)
        Dim Year05Users As New List(Of user)
        Dim Year06Users As New List(Of user)
        Dim Year07Users As New List(Of user)
        Dim Year08Users As New List(Of user)
        Dim Year09Users As New List(Of user)
        Dim Year10Users As New List(Of user)
        Dim Year11Users As New List(Of user)
        Dim Year12Users As New List(Of user)

        For Each user In users
            Select Case user.currentYear
                Case "K"
                    kindyUsers.Add(user)
                Case "01"
                    Year01Users.Add(user)
                Case "02"
                    Year02Users.Add(user)
                Case "03"
                    Year03Users.Add(user)
                Case "04"
                    Year04Users.Add(user)
                Case "05"
                    Year05Users.Add(user)
                Case "06"
                    Year06Users.Add(user)
                Case "07"
                    Year07Users.Add(user)
                Case "08"
                    Year08Users.Add(user)
                Case "09"
                    Year09Users.Add(user)
                Case "10"
                    Year10Users.Add(user)
                Case "11"
                    Year11Users.Add(user)
                Case "12"
                    Year12Users.Add(user)
            End Select
        Next

        addUsersToGroup(kindyUsers, config.sg_k)    '<================================================================
        addUsersToGroup(Year01Users, config.sg_1)
        addUsersToGroup(Year02Users, config.sg_2)
        addUsersToGroup(Year03Users, config.sg_3)
        addUsersToGroup(Year04Users, config.sg_4)
        addUsersToGroup(Year05Users, config.sg_5)
        addUsersToGroup(Year06Users, config.sg_6)
        addUsersToGroup(Year07Users, config.sg_7)
        addUsersToGroup(Year08Users, config.sg_8)
        addUsersToGroup(Year09Users, config.sg_9)
        addUsersToGroup(Year10Users, config.sg_10)
        addUsersToGroup(Year11Users, config.sg_11)
        addUsersToGroup(Year12Users, config.sg_12)


    End Sub

    Sub addParentsToGroups(users As List(Of user))
        Dim kindyParents As New List(Of user)
        Dim Year01Parents As New List(Of user)
        Dim Year02Parents As New List(Of user)
        Dim Year03Parents As New List(Of user)
        Dim Year04Parents As New List(Of user)
        Dim Year05Parents As New List(Of user)
        Dim Year06Parents As New List(Of user)
        Dim Year07Parents As New List(Of user)
        Dim Year08Parents As New List(Of user)
        Dim Year09Parents As New List(Of user)
        Dim Year10Parents As New List(Of user)
        Dim Year11Parents As New List(Of user)
        Dim Year12Parents As New List(Of user)

        For Each parent In users
            For Each child In parent.children
                Try
                    If child.currentYear = "K" Then
                        kindyParents.Add(parent)
                    End If
                    If child.currentYear = "01" Then
                        Year01Parents.Add(parent)
                    End If
                    If child.currentYear = "02" Then
                        Year02Parents.Add(parent)
                    End If
                    If child.currentYear = "03" Then
                        Year03Parents.Add(parent)
                    End If
                    If child.currentYear = "04" Then
                        Year04Parents.Add(parent)
                    End If
                    If child.currentYear = "05" Then
                        Year05Parents.Add(parent)
                    End If
                    If child.currentYear = "06" Then
                        Year06Parents.Add(parent)
                    End If
                    If child.currentYear = "07" Then
                        Year07Parents.Add(parent)
                    End If
                    If child.currentYear = "08" Then
                        Year08Parents.Add(parent)
                    End If
                    If child.currentYear = "09" Then
                        Year09Parents.Add(parent)
                    End If
                    If child.currentYear = "10" Then
                        Year10Parents.Add(parent)
                    End If
                    If child.currentYear = "11" Then
                        Year11Parents.Add(parent)
                    End If
                    If child.currentYear = "12" Then
                        Year12Parents.Add(parent)
                    End If
                Catch
                End Try
            Next
        Next


        Console.WriteLine("Kindy Parent Count: " & kindyParents.Count)

        addUsersToGroup(kindyParents, "CN=SG_Parent_YearK,OU=_Distribution Groups,OU=All,DC=i,DC=ofgs,DC=nsw,DC=edu,DC=au")

        addUsersToGroup(Year01Parents, "CN=SG_Parent_Year1,OU=_Distribution Groups,OU=All,DC=i,DC=ofgs,DC=nsw,DC=edu,DC=au")
        addUsersToGroup(Year02Parents, "CN=SG_Parent_Year2,OU=_Distribution Groups,OU=All,DC=i,DC=ofgs,DC=nsw,DC=edu,DC=au")
        addUsersToGroup(Year03Parents, "CN=SG_Parent_Year3,OU=_Distribution Groups,OU=All,DC=i,DC=ofgs,DC=nsw,DC=edu,DC=au")
        addUsersToGroup(Year04Parents, "CN=SG_Parent_Year4,OU=_Distribution Groups,OU=All,DC=i,DC=ofgs,DC=nsw,DC=edu,DC=au")
        addUsersToGroup(Year05Parents, "CN=SG_Parent_Year5,OU=_Distribution Groups,OU=All,DC=i,DC=ofgs,DC=nsw,DC=edu,DC=au")
        addUsersToGroup(Year06Parents, "CN=SG_Parent_Year6,OU=_Distribution Groups,OU=All,DC=i,DC=ofgs,DC=nsw,DC=edu,DC=au")
        addUsersToGroup(Year07Parents, "CN=SG_Parent_Year7,OU=_Distribution Groups,OU=All,DC=i,DC=ofgs,DC=nsw,DC=edu,DC=au")
        addUsersToGroup(Year08Parents, "CN=SG_Parent_Year8,OU=_Distribution Groups,OU=All,DC=i,DC=ofgs,DC=nsw,DC=edu,DC=au")
        addUsersToGroup(Year09Parents, "CN=SG_Parent_Year9,OU=_Distribution Groups,OU=All,DC=i,DC=ofgs,DC=nsw,DC=edu,DC=au")
        addUsersToGroup(Year10Parents, "CN=SG_Parent_Year10,OU=_Distribution Groups,OU=All,DC=i,DC=ofgs,DC=nsw,DC=edu,DC=au")
        addUsersToGroup(Year11Parents, "CN=SG_Parent_Year11,OU=_Distribution Groups,OU=All,DC=i,DC=ofgs,DC=nsw,DC=edu,DC=au")
        addUsersToGroup(Year12Parents, "CN=SG_Parent_Year12,OU=_Distribution Groups,OU=All,DC=i,DC=ofgs,DC=nsw,DC=edu,DC=au")

    End Sub




    Sub moveUserToOU(user As user, targetOU As String)

        Using ADuser As New DirectoryEntry("LDAP://" & user.distinguishedName)
            'Setting username & password to Nothing forces
            'the connection to use your logon credentials
            ADuser.Username = Nothing
            ADuser.Password = Nothing
            'Always use a secure connection
            ADuser.AuthenticationType = AuthenticationTypes.Secure
            ADuser.RefreshCache()

            ADuser.MoveTo(New DirectoryEntry(("LDAP://" & targetOU)))

        End Using

    End Sub

    Sub moveStudentToAlum(user As user, alumOU As String)
        'MsgBox("move fired")

        Dim targetOU As String
        targetOU = user.distinguishedName.Substring(user.distinguishedName.IndexOf(",") + 1).Replace("OU=Student Users", "OU=Alumni,OU=Student Users")



        moveUserToOU(user, targetOU)

    End Sub

    Sub moveUsersToOUs(adUsers As List(Of user), config As configSettings)



        Console.WriteLine("")
        Console.WriteLine("Moving users...")
        Console.WriteLine("")

        For Each adUser In adUsers
            'MsgBox(adUser.userAccountControl)
            Dim userAccountEnabled As Boolean


            Select Case adUser.userAccountControl
                Case Is = 512
                    userAccountEnabled = True
                Case Is = 514
                    userAccountEnabled = False
                Case Is = 544
                    userAccountEnabled = True
                Case Is = 546
                    userAccountEnabled = False
                Case Is = 66048
                    userAccountEnabled = True
                Case Is = 66050
                    userAccountEnabled = False
                Case Is = 66080
                    userAccountEnabled = True
                Case Is = 66082
                    userAccountEnabled = False
                Case Is = 262656
                    userAccountEnabled = True
                Case Is = 262658
                    userAccountEnabled = False
                Case Is = 262688
                    userAccountEnabled = True
                Case Is = 262690
                    userAccountEnabled = False
                Case Is = 328192
                    userAccountEnabled = True
                Case Is = 328194
                    userAccountEnabled = False
                Case Is = 328224
                    userAccountEnabled = True
                Case Is = 328226
                    userAccountEnabled = False
            End Select

            Dim targetOU As String
            targetOU = Nothing

            'Move former students to Alumni OU
            If adUser.distinguishedName.Contains("Student Users") And Not adUser.distinguishedName.Contains("Alumni") And Not adUser.distinguishedName.Contains("Generic") And Not userAccountEnabled Then
                'moveStudentToAlum(adUser, config.studentAlumOU)
                targetOU = "alum"
            End If

            'Move former staff 
            If adUser.edumateCurrent = 0 And adUser.distinguishedName.Contains("Staff Users") And Not adUser.distinguishedName.Contains("Generic") And Not adUser.distinguishedName.Contains("Domain") And Not adUser.distinguishedName.Contains("Former") And Not adUser.distinguishedName.Contains("@ofgsfamily.com") And Not adUser.distinguishedName.Contains("test") Then
                targetOU = config.formerStaffOU
            End If

            'Move current staff
            If adUser.edumateCurrent = 1 And adUser.distinguishedName.Contains("Staff Users") And Not adUser.distinguishedName.Contains("Generic") And Not adUser.distinguishedName.Contains("Domain") And Not adUser.distinguishedName.Contains("@ofgsfamily.com") And Not adUser.distinguishedName.Contains("test") Then

                If adUser.employmentType = 3 And Not adUser.distinguishedName.Contains("Casual") Then
                    targetOU = config.casualStaffOU
                End If
                If adUser.employmentType < 3 And Not adUser.distinguishedName.Contains("Current") Then
                    targetOU = config.currentStaffOU
                End If

            End If

            For Each group In adUser.memberOf
                If group = config.sg_tutors And adUser.edumateCurrent = 1 And adUser.distinguishedName.Contains("Staff Users") And Not adUser.distinguishedName.Contains("Generic") And Not adUser.distinguishedName.Contains("Domain") And Not adUser.distinguishedName.Contains("@ofgsfamily.com") And Not adUser.distinguishedName.Contains("test") Then
                    targetOU = config.tutorOU
                End If
            Next


            Console.WriteLine(" ThenMoving User: " & adUser.displayName)
            Console.WriteLine("Old OU: " & adUser.distinguishedName)
            Console.WriteLine("New OU: " & targetOU)

            If Not IsNothing(targetOU) Then
                If targetOU = "alum" Then
                    moveStudentToAlum(adUser, config.studentAlumOU)
                Else
                    moveUserToOU(adUser, targetOU)
                End If
            End If




        Next
    End Sub


    Function addUserTypeToADUSersFromEdumate(adUsers As List(Of user), edumateUsers As List(Of user))
        For Each adUser In adUsers
            For Each edumateUser In edumateUsers
                If adUser.employeeID = edumateUser.employeeID Then
                    adUser.userType = edumateUser.userType
                End If
            Next
        Next
        Return adUsers
    End Function


    Function getEdumateGroups(users As List(Of user), config As configSettings)

        Dim ConnectionString As String = config.edumateConnectionString
        Dim commandString As String =
"
SELECT        
group_membership.groups_id,
group_membership.contact_id


FROM            group_membership
"


        Using conn As New IBM.Data.DB2.DB2Connection(ConnectionString)
            conn.Open()

            'define the command object to execute
            Dim command As New IBM.Data.DB2.DB2Command(commandString, conn)
            command.Connection = conn
            command.CommandText = commandString

            Dim dr As IBM.Data.DB2.DB2DataReader
            dr = command.ExecuteReader

            Dim i As Integer = 0
            While dr.Read()
                If Not dr.IsDBNull(0) Then

                    For Each user In users


                        If user.contact_id = dr.GetValue(1) Then
                            If IsNothing(user.edumateGroupMemberships) Then
                                user.edumateGroupMemberships = New List(Of String)
                            End If
                            user.edumateGroupMemberships.Add(dr.GetValue(0))
                        End If
                    Next



                End If
            End While
            conn.Close()
        End Using
        Return users
    End Function



    Sub purgeStaffDB(config As configSettings)
        Dim conn As New MySqlConnection
        connect(conn, config)

        Try
            conn.Open()
        Catch ex As Exception
        End Try

        Dim cmd As New MySqlCommand(String.Format("DELETE FROM `staff_details` WHERE 1"), conn)
        cmd.ExecuteNonQuery()

        conn.Close()

    End Sub

    Function addADObjectoToEdumateUser(users As List(Of user), adUsers As List(Of user))
        For Each user In users
            For Each adUser In adUsers
                If user.employeeID = adUser.employeeID Then
                    user.adFirstname = adUser.firstName
                    user.adSurname = Replace(adUser.surname, "'", "")
                End If
            Next
        Next
        addADObjectoToEdumateUser = users



    End Function



End Module


