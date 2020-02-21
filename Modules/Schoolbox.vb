Imports System.IO
Imports System.DirectoryServices
Imports System.Text.RegularExpressions
Imports System.Net.Mail
Imports MySql.Data.MySqlClient
Imports System.Text
Imports WinSCP
Imports System.Net.Http


Public Module Schoolbox

    Public Sub SchoolboxMain(adconfig As configSettings, currentEdumateStudents As List(Of user), edumateParents As List(Of user))

        Console.WriteLine("Doing Schoolbox stuff")
        Dim config As schoolboxConfigSettings
        config = SchoolboxReadConfig()

		Console.WriteLine("Creating user.csv")
		Call writeUserCSV(config, adconfig, currentEdumateStudents, edumateParents)
		Console.WriteLine("User.csv done")
        'Console.WriteLine("")
        'Console.WriteLine("")

        Call timetableStructure(config)
        Call timetable(config)
        Call enrollment(config)
        Call events(config)

        Call uploadFiles(config)
        Call relationships(config)

        Console.WriteLine("Schoolbox stuff done")
    End Sub

    Function SchoolboxReadConfig()
        Dim config As New schoolboxConfigSettings()
        config.uploadServers = New List(Of uploadServer)

        Try
            ' Open the file using a stream reader.
            Dim directory As String = My.Application.Info.DirectoryPath



            Using sr As New StreamReader(directory & "\config.ini")
                Dim line As String
                While Not sr.EndOfStream
                    line = sr.ReadLine

                    Select Case True
                        Case Left(line, 17) = "connectionstring="
                            config.connectionString = Mid(line, 18)
                        Case Left(line, 13) = "uploadserver="
                            line = Mid(line, 14)
                            Dim split As String() = line.Split(";")
                            config.uploadServers.Add(New uploadServer)
                            config.uploadServers.Last.host = split(0)
                            config.uploadServers.Last.userName = split(1)
                            config.uploadServers.Last.pass = split(2)
                            config.uploadServers.Last.rsa = split(3)
                        Case Left(line, 19) = "studentemaildomain="
                            config.studentEmailDomain = Mid(line, 20)
                    End Select

                End While

                SchoolboxReadConfig = config
            End Using

        Catch e As Exception
            MsgBox(e.Message)
        End Try
    End Function

    Sub writeUserCSV(config As schoolboxConfigSettings, adconfig As configSettings, currentEdumateStudents As List(Of user), edumateParents As List(Of user))


        Dim dirEntry As DirectoryEntry

        Console.WriteLine("Connecting to AD...")
        dirEntry = GetDirectoryEntry(adconfig.ldapDirectoryEntry)

        Dim adUsers As List(Of user)
        Console.Write("Loading AD users...")
        adUsers = getADUsers(dirEntry)
        Console.Write("Done!" & Chr(13) & Chr(10))

        ' Students ****************
        Dim ConnectionString As String = config.connectionString
        Dim commandString As String

        Dim users As New List(Of SchoolBoxUser)

        Console.WriteLine("Loading Schoolbox Data... ")
        Console.Write("Students... ")


        For Each edumateStudent In currentEdumateStudents
            users.Add(New SchoolBoxUser)

            users.Last.Delete = ""
            users.Last.SchoolboxUserID = ""
            users.Last.Title = ""

            Try
                If CInt(edumateStudent.currentYear) < 7 Then
                    users.Last.Campus = "Junior"
                    users.Last.Role = "Junior Students"
                Else
                    users.Last.Campus = "Senior"
                    users.Last.Role = "Senior Students"
                End If
            Catch
                users.Last.Campus = "Junior"
                users.Last.Role = "Junior Students"
            End Try

            users.Last.Password = ""

            users.Last.Year = edumateStudent.currentYear
            users.Last.House = ""
            users.Last.ResidentialHouse = ""
            users.Last.EPortfolio = "Y"
            users.Last.HideContactDetails = "Y"
            users.Last.HideTimetable = "N"
            users.Last.EmailAddressFromUsername = "N"
            users.Last.UseExternalMailClient = "N"
            users.Last.EnableWebmailTab = "Y"
            users.Last.Superuser = "N"
            users.Last.AccountEnabled = "Y"
            users.Last.ChildExternalIDs = ""
            users.Last.HomePhone = ""
            users.Last.MobilePhone = ""
            users.Last.WorkPhone = ""
            users.Last.Address = ""
            users.Last.Suburb = ""
            users.Last.Postcode = ""
            users.Last.Username = edumateStudent.ad_username
            users.Last.AltEmail = (users.Last.Username & config.studentEmailDomain)
            users.Last.ExternalID = edumateStudent.employeeNumber
            users.Last.FirstName = """" & Replace(edumateStudent.firstName, "&#039;", "'") & """"
            users.Last.Surname = """" & Replace(edumateStudent.surname, "&#039;", "'") & """"
			Try
				users.Last.DateOfBirth = ddMMYYYY_to_yyyyMMdd(edumateStudent.dob)
			Catch
			End Try



		Next

        Console.Write("Parents... ")

        'Parents **********************

        For Each edumateParent In edumateParents
            users.Add(New SchoolBoxUser)

            users.Last.Delete = ""
            users.Last.SchoolboxUserID = ""
            users.Last.Title = ""
            users.Last.Role = "Parents"
            users.Last.Password = ""
            users.Last.Year = "Parent"
            users.Last.ResidentialHouse = ""
            users.Last.EPortfolio = "N"
            users.Last.HideContactDetails = "Y"
            users.Last.HideTimetable = "Y"
            users.Last.EmailAddressFromUsername = "N"
            users.Last.UseExternalMailClient = "Y"
            users.Last.EnableWebmailTab = "Y"
            users.Last.Superuser = "N"
            users.Last.AccountEnabled = "Y"
            users.Last.HomePhone = ""
            users.Last.MobilePhone = ""
            users.Last.WorkPhone = ""
            users.Last.DateOfBirth = ""
            users.Last.Address = ""
            users.Last.Suburb = ""
            users.Last.Postcode = ""
            users.Last.Username = edumateParent.ad_username
            users.Last.AltEmail = users.Last.Username & adconfig.parentDomainName
            users.Last.ExternalID = edumateParent.edumateProperties.carer_number
            users.Last.FirstName = """" & Replace(edumateParent.firstName, "&#039;", "'") & """"
            users.Last.Surname = """" & Replace(edumateParent.surname, "&#039;", "'") & """"

            For Each child In edumateParent.children
                If Not IsNothing(child) Then

					Select Case child.currentYear
						Case "12"
							Select Case users.Last.Campus
								Case ""
									users.Last.Campus = "Senior"
								Case "Junior"
									users.Last.Campus = "Junior, Senior"
								Case "Senior"
									users.Last.Campus = "Senior"
								Case "Junior, Senior"
									users.Last.Campus = "Junior, Senior"
							End Select
						Case "11"
							Select Case users.Last.Campus
								Case ""
									users.Last.Campus = "Senior"
								Case "Junior"
									users.Last.Campus = "Junior, Senior"
								Case "Senior"
									users.Last.Campus = "Senior"
								Case "Junior, Senior"
									users.Last.Campus = "Junior, Senior"
							End Select
						Case "10"
							Select Case users.Last.Campus
								Case ""
									users.Last.Campus = "Senior"
								Case "Junior"
									users.Last.Campus = "Junior, Senior"
								Case "Senior"
									users.Last.Campus = "Senior"
								Case "Junior, Senior"
									users.Last.Campus = "Junior, Senior"
							End Select
                        Case "9"
                            Select Case users.Last.Campus
								Case ""
									users.Last.Campus = "Senior"
								Case "Junior"
									users.Last.Campus = "Junior, Senior"
								Case "Senior"
									users.Last.Campus = "Senior"
								Case "Junior, Senior"
									users.Last.Campus = "Junior, Senior"
							End Select
                        Case "8"
                            Select Case users.Last.Campus
								Case ""
									users.Last.Campus = "Senior"
								Case "Junior"
									users.Last.Campus = "Junior, Senior"
								Case "Senior"
									users.Last.Campus = "Senior"
								Case "Junior, Senior"
									users.Last.Campus = "Junior, Senior"
							End Select
                        Case "7"
                            Select Case users.Last.Campus
								Case ""
									users.Last.Campus = "Senior"
								Case "Junior"
									users.Last.Campus = "Junior, Senior"
								Case "Senior"
									users.Last.Campus = "Senior"
								Case "Junior, Senior"
									users.Last.Campus = "Junior, Senior"
							End Select
                        Case "6"
                            Select Case users.Last.Campus
								Case ""
									users.Last.Campus = "Junior"
								Case "Junior"
									users.Last.Campus = "Junior"
								Case "Senior"
									users.Last.Campus = "Junior, Senior"
								Case "Junior, Senior"
									users.Last.Campus = "Junior, Senior"
							End Select
                        Case "5"
                            Select Case users.Last.Campus
								Case ""
									users.Last.Campus = "Junior"
								Case "Junior"
									users.Last.Campus = "Junior"
								Case "Senior"
									users.Last.Campus = "Junior, Senior"
								Case "Junior, Senior"
									users.Last.Campus = "Junior, Senior"
							End Select
                        Case "4"
                            Select Case users.Last.Campus
								Case ""
									users.Last.Campus = "Junior"
								Case "Junior"
									users.Last.Campus = "Junior"
								Case "Senior"
									users.Last.Campus = "Junior, Senior"
								Case "Junior, Senior"
									users.Last.Campus = "Junior, Senior"
							End Select
                        Case "3"
                            Select Case users.Last.Campus
								Case ""
									users.Last.Campus = "Junior"
								Case "Junior"
									users.Last.Campus = "Junior"
								Case "Senior"
									users.Last.Campus = "Junior, Senior"
								Case "Junior, Senior"
									users.Last.Campus = "Junior, Senior"
							End Select
                        Case "2"
                            Select Case users.Last.Campus
								Case ""
									users.Last.Campus = "Junior"
								Case "Junior"
									users.Last.Campus = "Junior"
								Case "Senior"
									users.Last.Campus = "Junior, Senior"
								Case "Junior, Senior"
									users.Last.Campus = "Junior, Senior"
							End Select
                        Case "1"
                            Select Case users.Last.Campus
								Case ""
									users.Last.Campus = "Junior"
								Case "Junior"
									users.Last.Campus = "Junior"
								Case "Senior"
									users.Last.Campus = "Junior, Senior"
								Case "Junior, Senior"
									users.Last.Campus = "Junior, Senior"
							End Select
						Case "K"
							Select Case users.Last.Campus
								Case ""
									users.Last.Campus = "Junior"
								Case "Junior"
									users.Last.Campus = "Junior"
								Case "Senior"
									users.Last.Campus = "Junior, Senior"
								Case "Junior, Senior"
									users.Last.Campus = "Junior, Senior"
							End Select
						Case Else
					End Select

					If users.Last.ChildExternalIDs = "" Then
                        users.Last.ChildExternalIDs = child.employeeNumber
                    Else
                        users.Last.ChildExternalIDs = users.Last.ChildExternalIDs & ", " & child.employeeNumber
                    End If
                End If
            Next
            users.Last.ChildExternalIDs = """" & users.Last.ChildExternalIDs & """"


			If edumateParent.children.Count > 0 And users.Last.Campus = "" Then
				users.Last.Campus = "Junior, Senior"
			End If



		Next


		Console.Write("Done!" & Chr(13) & Chr(10))

        Console.Write("Staff... ")

        'Staff **********************
        commandString = "
select
schoolbox_staff2.username1,
schoolbox_staff2.staff_number,
schoolbox_staff2.salutation,
schoolbox_staff2.firstname,
schoolbox_staff2.surname,
schoolbox_staff2.house,
staff.staff_id,
case when staff.staff_number in (

select distinct
schoolbox_staff1.staff_number

from
(
select
staff.staff_number,
salutation.salutation,
coalesce(replace(contact.preferred_name,'&#0'||'39;',''''), replace(contact.firstname,'&#0'||'39;','''')) as firstname,
replace(contact.surname,'&#039;','''') as surname,
sys_user.username as username1,
contact.email_address,
house.house,
campus.campus,
replace(work_detail.title,'&#039;','''') as title
from staff
inner join contact on contact.contact_id = staff.contact_id
left join staff_employment on staff_employment.staff_id = staff.staff_id
left join work_detail on work_detail.contact_id=contact.contact_id
left join salutation on salutation.salutation_id = contact.salutation_id
left join sys_user on sys_user.contact_id = contact.contact_id
left join house on house.house_id = staff.house_id
left join campus on campus.campus_id = staff.campus_id
where (staff_employment.end_date is null or staff_employment.end_date >= current date)
and staff_employment.start_date <= (current date +90 DAYS)
and (contact.pronounced_name is null or contact.pronounced_name != 'NOT STAFF')

) schoolbox_staff1

inner join staff on schoolbox_staff1.staff_number = staff.staff_number

inner join contact on staff.contact_id = contact.contact_id

left join teacher on contact.contact_id = teacher.contact_id

left join class_teacher on class_teacher.teacher_id = teacher.teacher_id

left join class on class.class_id = class_teacher.class_id


left join 
(
	select max_student_class.class_id, form.short_name

	from 
	(
		select max(student_id) as randomStudentNumber, class_id

		from class_enrollment

		where 

		(SELECT current date FROM sysibm.sysdummy1) between class_enrollment.start_date and class_enrollment.end_date

		group by class_id
	) max_student_class


	INNER JOIN 
	(
		select student_id, max(form_run_id) as max_form_run_id

		from student_form_run 

		where  
		(SELECT current date FROM sysibm.sysdummy1) between student_form_run.start_date and student_form_run.end_date
	
		group by student_id
	) max_form_run
	ON max_form_run.student_id = max_student_class.randomStudentNumber

	INNER JOIN form_run on max_form_run.max_form_run_id = form_run.form_run_id

	INNER JOIN form on form_run.form_id = form.form_id

) class_short_names

on class.class_id = class_short_names.class_id


where class_short_names.short_name = 'K'
and class.class_type_id = '2'

) then 'true' else 'false' END AS kindy,
schoolbox_staff2.title as title

from (

select
staff.staff_number,
salutation.salutation,
coalesce(replace(contact.preferred_name,'&#0'||'39;',''''), replace(contact.firstname,'&#0'||'39;','''')) as firstname,
replace(contact.surname,'&#039;','''') as surname,
sys_user.username as username1,
contact.email_address,
house.house,
campus.campus,
replace(work_detail.title,'&#039;','''') as title
from staff
inner join contact on contact.contact_id = staff.contact_id
left join staff_employment on staff_employment.staff_id = staff.staff_id
left join work_detail on work_detail.contact_id=contact.contact_id
left join salutation on salutation.salutation_id = contact.salutation_id
left join sys_user on sys_user.contact_id = contact.contact_id
left join house on house.house_id = staff.house_id
left join campus on campus.campus_id = staff.campus_id
where (staff_employment.end_date is null or staff_employment.end_date >= current date)
and staff_employment.start_date <= (current date + 90 DAYS)
and (contact.pronounced_name is null or contact.pronounced_name != 'NOT STAFF')

)  schoolbox_staff2

inner join staff on schoolbox_staff2.staff_number = staff.staff_number

"

        Using conn As New IBM.Data.DB2.DB2Connection(ConnectionString)
            conn.Open()

            'define the command object to execute
            Dim command As New IBM.Data.DB2.DB2Command(commandString, conn)
            command.Connection = conn
            command.CommandText = commandString

            Dim dr As IBM.Data.DB2.DB2DataReader
            dr = command.ExecuteReader

            While dr.Read()

                users.Add(New SchoolBoxUser)

                users.Last.Delete = ""
                users.Last.SchoolboxUserID = ""
                'users.Last.Title = ""
                If Not dr.IsDBNull(8) Then
                    users.Last.PositionTitle = """" & dr.GetValue(8) & """"
                Else users.Last.PositionTitle = """" & "Staff" & """"
                End If

                users.Last.Role = "Staff"

				users.Last.Campus = "Junior, Senior"
				users.Last.Password = ""

                users.Last.Year = "Staff"
                users.Last.ResidentialHouse = ""
                users.Last.EPortfolio = "Y"
                users.Last.HideContactDetails = "Y"
                users.Last.HideTimetable = "N"
                users.Last.EmailAddressFromUsername = "Y"
                users.Last.UseExternalMailClient = "Y"
                users.Last.EnableWebmailTab = "Y"
                users.Last.Superuser = "N"
                users.Last.AccountEnabled = "Y"
                users.Last.ChildExternalIDs = " "
                users.Last.HomePhone = ""
                users.Last.MobilePhone = ""
                users.Last.WorkPhone = ""
                users.Last.Address = ""
                users.Last.Suburb = ""
                users.Last.Postcode = ""
                users.Last.DateOfBirth = ""

                '*******************  all this needs cleaning up ******************

                'If Not dr.IsDBNull(0) Then users.Last.Username = dr.GetValue(0)
                If Not dr.IsDBNull(6) Then users.Last.Username = getUsernameFromID(dr.GetValue(6), adUsers)

                If users.Last.Username.ToLower = "jenniferlu" Then
                    users.Last.Role = "Administration"
                End If
                If users.Last.Username.ToLower = "juliet" Then
                    users.Last.Role = "Administration"
                End If

                If users.Last.Username.ToLower = "selinam" Then
                    users.Last.Role = "Administration"
                End If

                If users.Last.Username.ToLower = "kathys" Then
                    users.Last.Role = "Administration"
                End If

                If users.Last.Username.ToLower = "katrinaj" Then
                    users.Last.Role = "Administration"
                End If

                If users.Last.Username.ToLower = "fionar" Then
                    users.Last.Role = "Administration"
                End If

                If users.Last.Username.ToLower = "jacquib" Then
                    users.Last.Role = "Administration"
                End If

                If users.Last.Username.ToLower = "matthewha" Then
                    users.Last.Role = "Administration"
                End If
                If users.Last.Username.ToLower = "michaelp" Then
                    users.Last.Role = "Administration"
                End If
                If users.Last.Username.ToLower = "janineg" Then
                    users.Last.Role = "Administration"
                End If
				If users.Last.Username.ToLower = "jillianp" Then
					users.Last.Role = "Administration"
				End If
				If users.Last.Username.ToLower = "julies" Then
					users.Last.Role = "Administration"
				End If
                If users.Last.Username.ToLower = "donnal" Then
                    users.Last.Role = "Administration"
                End If

                If dr.GetValue(7) = "true" Then
                    users.Last.AltEmail = "donotemail@ofgs.nsw.edu.au"
                    users.Last.EmailAddressFromUsername = "N"
                Else
                    users.Last.AltEmail = users.Last.Username & adconfig.staffDomainName
                End If

                If users.Last.AltEmail = "pddowney@ofgs.nsw.edu.au" Then
                    users.Last.EmailAddressFromUsername = "N"
                    users.Last.AltEmail = "principal@ofgs.nsw.edu.au"
                End If

                If users.Last.AltEmail = "dannyrav@ofgs.nsw.edu.au" Then
                    users.Last.EmailAddressFromUsername = "N"
                    users.Last.AltEmail = "sport@ofgs.nsw.edu.au"
                End If

                ' *******************************************************************************




                If Not dr.IsDBNull(1) Then users.Last.ExternalID = dr.GetValue(1)
                If Not dr.IsDBNull(2) Then users.Last.Title = dr.GetValue(2)
                If Not dr.IsDBNull(3) Then users.Last.FirstName = """" & dr.GetValue(3) & """"
                If Not dr.IsDBNull(4) Then users.Last.Surname = """" & dr.GetValue(4) & """"
                If Not dr.IsDBNull(5) Then users.Last.House = dr.GetValue(5)





            End While
            conn.Close()
        End Using

        Console.Write("Done!" & Chr(13) & Chr(10))


        Console.WriteLine("Saving to CSV...")



        Dim sw As New StreamWriter(".\user.csv")
        sw.WriteLine("Delete?,Schoolbox User ID,Username,External ID,Title,First Name,Surname,Role,Campus,Password,Alt Email,Year,House,Residential House,E-Portfolio?,Hide Contact Details?,Hide Timetable?,Email Address From Username?,Use External Mail Client?,Enable Webmail Tab?,Account Enabled?,Child External IDs,Date of Birth,Home Phone,Mobile Phone,Work Phone,Address,Suburb,Postcode,Position Title")
        For Each i In users

            ' If i.Campus = "Junior" Or i.Campus = "Senior" Or i.Campus = """Junior, Senior""" Then
            If Not IsNothing(i.Campus) Then
                Select Case i.Year
                    Case "K"
                    Case "1"
                    Case "2"
                    Case "3"
                    Case "4"
                    Case "5"
                    Case "6"
                    Case "7"
                    Case "8"
                    Case "9"
                    Case "10"
                    Case "11"
                    Case "12"
                    Case "Parent"
                    Case "Staff"
                    Case Else
                        i.Year = "K"
                End Select



				If Not IsNothing(i.Username) Then
					sw.WriteLine(i.Delete & "," & i.SchoolboxUserID & "," & i.Username & "," & i.ExternalID & "," & i.Title & "," & i.FirstName & "," & i.Surname & "," & i.Role & ",""" & i.Campus & """," & i.Password & "," & i.AltEmail & "," & i.Year & "," & i.House & "," & i.ResidentialHouse & "," & i.EPortfolio & "," & i.HideContactDetails & "," & i.HideTimetable & "," & i.EmailAddressFromUsername & "," & i.UseExternalMailClient & "," & i.EnableWebmailTab & "," & i.AccountEnabled & "," & i.ChildExternalIDs & "," & i.DateOfBirth & "," & i.HomePhone & "," & i.MobilePhone & "," & i.WorkPhone & "," & i.Address & "," & i.Suburb & "," & i.Postcode & "," & i.PositionTitle)
				End If
			End If


		Next
        sw.Close()
        Console.WriteLine("Done!" & Chr(13) & Chr(10))

    End Sub

    Sub timetableStructure(config As schoolboxConfigSettings)

        Dim sep As String = ","
        Dim commandString As String
        commandString = "
SELECT DISTINCT 
                         CASE WHEN substr(timetable.timetable, 6, 6) = 'Year 1' THEN 'Senior' ELSE substr(timetable.timetable, 6, 6) END AS Expr1, 
                         REPLACE(CONCAT(CONCAT(term.term, ' '), substr(timetable.timetable, 1, 4)), 'Term 0', 'Term 4') AS Expr2, term.v_start_date, term.end_date, term.cycle_start_day, 
                         cycle_day.day_index, period.period, period.start_time, period.end_time
FROM            TERM_GROUP, cycle_day, period_cycle_day, period, term, timetable
WHERE        (start_date > '01/01/2020') AND (end_date < '01/01/2021') AND (term_group.cycle_id = cycle_day.cycle_id) AND 
                         (cycle_day.cycle_day_id = period_cycle_day.cycle_day_id) AND (period_cycle_day.period_id = period.period_id) AND (term_group.term_id = term.term_id) AND 
                         (term.timetable_id = timetable.timetable_id)"





        Dim sw As New StreamWriter(".\timetableStructure.csv")

        Dim ConnectionString As String = config.connectionString
        Using conn As New IBM.Data.DB2.DB2Connection(ConnectionString)
            conn.Open()

            'define the command object to execute
            Dim command As New IBM.Data.DB2.DB2Command(commandString, conn)
            command.Connection = conn
            command.CommandText = commandString

            Dim dr As IBM.Data.DB2.DB2DataReader
            dr = command.ExecuteReader


            sw.WriteLine("Term Campus,Term Title,Term Start,Term Finish,Term Start Day Number,Period Day,Period Title,Period Start,Period Finish")

            Dim fields As Integer = dr.FieldCount - 1
            While dr.Read()
                Dim sb As New StringBuilder()

                Dim outLine As String
                Dim strTermTitle As String

				strTermTitle = Replace(dr.GetValue(1), "2020", "2019")

				outLine = (dr.GetValue(0) & "," & strTermTitle & "," & Format(dr.GetValue(2), "yyyy-MM-dd") & "," & Format(dr.GetValue(3), "yyyy-MM-dd") & "," & dr.GetValue(4) & "," & dr.GetValue(5).ToString & "," & dr.GetValue(6).ToString & "," & dr.GetValue(7).ToString & "," & dr.GetValue(8).ToString)
                sw.WriteLine(outLine)
            End While
            conn.Close()
        End Using


        sw.Close()





    End Sub

    Sub timetable(config As schoolboxConfigSettings)
        Dim commandstring As String
        commandstring = "
SELECT DISTINCT
 substr(timetable.timetable, 6, 6) as CAMPUS1,	
CONCAT(CONCAT(term.term, ' '), substr(timetable.timetable, 1, 4)) AS Expr2,
cycle_day.day_index as DAY_NUMBER,
	period.period as PERIOD_NUMBER,
	concat(course.code,class.identifier) AS CLASS_CODE,
	class.class,
	room.code AS ROOM,
staff.staff_number


FROM period_class
INNER JOIN period_cycle_day ON period_cycle_day.period_cycle_day_id = period_class.period_cycle_day_id
INNER JOIN cycle_day ON cycle_day.cycle_day_id = period_cycle_day.cycle_day_id
INNER JOIN period ON period.period_id = period_cycle_day.period_id
INNER JOIN class ON class.class_id = period_class.class_id
INNER JOIN course ON course.course_id = class.course_id
INNER JOIN perd_cls_teacher ON perd_cls_teacher.period_class_id = period_class.period_class_id 
	AND perd_cls_teacher.is_primary = 1
INNER JOIN teacher ON teacher.teacher_id = perd_cls_teacher.teacher_id
INNER JOIN staff ON staff.contact_id = teacher.contact_id
INNER JOIN room ON room.room_id = period_class.room_id
INNER JOIN timetable ON timetable.timetable_id = period_class.timetable_id
INNER JOIN contact ON staff.contact_id = contact.contact_id
INNER JOIN term_group on cycle_day.cycle_id = term_group.cycle_id
INNER JOIN term ON term_group.term_id = term.term_id
WHERE
(
	current date BETWEEN (
	CASE
		WHEN period_class.effective_start = timetable.computed_start_date
		THEN timetable.computed_v_start_date
		ELSE timetable.computed_v_start_date
	END
	)
	AND period_class.effective_end
)
AND
(
	current date BETWEEN (
	CASE
		WHEN period_class.effective_start = timetable.computed_start_date
		THEN timetable.computed_v_start_date
		ELSE timetable.computed_v_start_date
	END
	)
	AND timetable.computed_end_date
)
AND
(
period_class.effective_start BETWEEN term.start_date AND term.end_date
OR
period_class.effective_end BETWEEN term.start_date AND term.end_date
OR
(period_class.effective_start <= term.start_date AND period_class.effective_end >= term.end_date)
)
"

        Dim sw As New StreamWriter(".\timetable.csv")

        Dim ConnectionString As String = config.connectionString
        Using conn As New IBM.Data.DB2.DB2Connection(ConnectionString)
            conn.Open()

            'define the command object to execute
            Dim command As New IBM.Data.DB2.DB2Command(commandstring, conn)
            command.Connection = conn
            command.CommandText = commandstring

            Dim dr As IBM.Data.DB2.DB2DataReader
            dr = command.ExecuteReader


            sw.WriteLine("Term Campus,Term Title,Period Day,Period Title,Class Code,Class Title,Class Room,Teacher Code")

            Dim fields As Integer = dr.FieldCount - 1
            While dr.Read()
                Dim sb As New StringBuilder()

                Dim outLine As String
                Dim tempStr As String
                Dim campus As String
                Dim strTerm As String

                tempStr = dr.GetValue(5)
                tempStr = Replace(tempStr, "&#039;", "'")
                tempStr = Replace(tempStr, "&amp;", "&")

                campus = Replace(dr.GetValue(0), "Year 1", "Senior")
                strTerm = Replace(dr.GetValue(1), "Term 0", "Term 4")
				strTerm = Replace(strTerm, "2020", "2019")
				If True Then
                    outLine = (campus & "," & strTerm & "," & dr.GetValue(2) & "," & dr.GetValue(3) & "," & dr.GetValue(4) & ",""" & tempStr & """," & dr.GetValue(6) & "," & dr.GetValue(7))
                    sw.WriteLine(outLine)
                End If
            End While



            conn.Close()
        End Using
        sw.Close()
    End Sub

    Sub enrollment(config As schoolboxConfigSettings)
        Dim commandstring As String
        commandstring = "
SELECT DISTINCT 

CONCAT(course.code, class.identifier) AS CLASS_CODE, 
class.class, 
student.student_number

FROM            CLASS_ENROLLMENT

INNER JOIN STUDENT 
	ON class_enrollment.student_id = student.student_id

INNER JOIN class 
	ON class_enrollment.class_id = class.class_id
	
INNER JOIN COURSE
	ON class.course_id = course.course_id
	
INNER JOIN ACADEMIC_YEAR
	ON class.academic_year_id = academic_year.academic_year_id
	

 WHERE (academic_year.academic_year = CAST((YEAR(current_date))AS VARCHAR(10)) OR academic_year.academic_year =CAST((YEAR(current_date+ 1 years)) AS varchar(10))) AND ((SELECT current date FROM sysibm.sysdummy1) between (class_enrollment.start_date -10 DAYS) AND class_enrollment.end_date)

 UNION
 
SELECT DISTINCT 

replace(CONCAT(course.code, class.identifier),'12','13') AS CLASS_CODE, 
replace(class.class,'12','13') AS CLASS,
student.student_number

FROM            CLASS_ENROLLMENT

INNER JOIN STUDENT 
	ON class_enrollment.student_id = student.student_id

INNER JOIN class 
	ON class_enrollment.class_id = class.class_id
	
INNER JOIN COURSE
	ON class.course_id = course.course_id
	
INNER JOIN ACADEMIC_YEAR
	ON class.academic_year_id = academic_year.academic_year_id
	

 WHERE (academic_year.academic_year = CAST((YEAR(current_date))AS VARCHAR(10)))  
 	--AND ((SELECT current date FROM sysibm.sysdummy1) between (class_enrollment.start_date -10 DAYS) 
	--AND class_enrollment.end_date) 
 	AND student.student_number IN 
 	(
 	SELECT  distinct      

student.student_number

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
INNER JOIN TIMETABLE ON form_run.TIMETABLE_ID = timetable.TIMETABLE_ID



	
WHERE 

(
YEAR(edumate.view_student_start_exit_dates.exit_date) = YEAR(student_form_run.end_date)) 
AND YEAR(edumate.view_student_start_exit_dates.exit_date) = year(current_date)
AND form.SHORT_NAME = '12'
AND class.CLASS LIKE '12%'
AND class_enrollment.END_DATE < current_date
)
 
 
"




        Dim sw As New StreamWriter(".\enrollment.csv")

        Dim ConnectionString As String = config.connectionString
        Using conn As New IBM.Data.DB2.DB2Connection(ConnectionString)
            conn.Open()

            'define the command object to execute
            Dim command As New IBM.Data.DB2.DB2Command(commandstring, conn)
            command.Connection = conn
            command.CommandText = commandstring

            Dim dr As IBM.Data.DB2.DB2DataReader
            dr = command.ExecuteReader


            sw.WriteLine("Class Code,Class Title,Student Code")

            Dim fields As Integer = dr.FieldCount - 1
            While dr.Read()
                Dim sb As New StringBuilder()

                Dim outLine As String
                Dim tempStr As String

                tempStr = dr.GetValue(1)
                tempStr = Replace(tempStr, "&#039;", "'")
                tempStr = Replace(tempStr, "&amp;", "&")

                outLine = (dr.GetValue(0) & ",""" & tempStr & """," & dr.GetValue(2))
                sw.WriteLine(outLine)
            End While



            conn.Close()
        End Using




        '&#039;
        '&#039;
        '&amp;

        sw.Close()

    End Sub


    Sub events(config As schoolboxConfigSettings)
        Dim commandstring As String
        commandstring = "



SELECT 
DATE(event.start_date) as ""Start Date"", 
varchar_format(event.start_date, 'HH24:MI')  ""Start Time"",
DATE(event.end_date) as ""Finish Date"",
varchar_format(event.end_date, 'HH24:MI')  ""Finish Time"",
0 as ""All Day"",
event.event as ""Name"",
event.description as ""Detail"",
CASE
	when event.location IS NOT NULL then (event.location)
	when (event_rooms.room_count > 1) then	('Various')
	when (event_rooms.room_count = 1) then  (room.room)
end as Location,

1 as ""Type"",
NULL as ""Publish Date"",
0 as ""Attendance"",
event.audience_type_id as ""Audience Type""

FROM
  event
left join
	( select event_id, max(room_id) as max_room_id, count(room_id) as room_count
	from event_room
	group by event_id
	)  event_rooms
on event.event_id = event_rooms.event_id

left join room on event_rooms.max_room_id = room.room_id



WHERE
event.start_date >  '01/01/2018' 
AND event.end_date < '12/31/2019' 
AND event.publish_flag = 1
AND event.recurring_id is not null and event.recurring_id > 0
AND event.event_id = (select min(event_id) from event e2
                                     where e2.recurring_id = event.recurring_id)


UNION



SELECT 
DATE(event.start_date) as ""Start Date"", 
varchar_format(event.start_date, 'HH24:MI')  ""Start Time"",
DATE(event.end_date) as ""Finish Date"",
varchar_format(event.end_date, 'HH24:MI')  ""Finish Time"",
0 as ""All Day"",
event.event as ""Name"",
event.description as ""Detail"",
CASE
	when event.location IS NOT NULL then (event.location)
	when (event_rooms.room_count > 1) then	('Various')
	when (event_rooms.room_count = 1) then  (room.room)
end as Location,

1 as ""Type"",
NULL as ""Publish Date"",
0 as ""Attendance"",
event.audience_type_id as ""Audience Type""
FROM
  event

left join
	( 
	select event_id, max(room_id) as max_room_id, count(room_id) as room_count
	from event_room
	group by event_id
	)  event_rooms
on event.event_id = event_rooms.event_id

left join room on event_rooms.max_room_id = room.room_id

WHERE
event.start_date >  '01/01/2020' 
AND event.end_date < '12/31/2021' 
AND event.publish_flag = 1
AND event.recurring_id is null

"



        Dim swAll As New StreamWriter(".\calendarAll.csv")
        Dim swJnr As New StreamWriter(".\calendarJnr.csv")
        Dim swSnr As New StreamWriter(".\calendarSnr.csv")


        Dim ConnectionString As String = config.connectionString
        Using conn As New IBM.Data.DB2.DB2Connection(ConnectionString)
            conn.Open()

            'define the command object to execute
            Dim command As New IBM.Data.DB2.DB2Command(commandstring, conn)
            command.Connection = conn
            command.CommandText = commandstring

            Dim dr As IBM.Data.DB2.DB2DataReader
            dr = command.ExecuteReader

            swAll.WriteLine("Start Date,Start Time,Finish Date,Finish Time,All Day,Name,Detail,Location,Type,Publish Date,Attendance")
            swJnr.WriteLine("Start Date,Start Time,Finish Date,Finish Time,All Day,Name,Detail,Location,Type,Publish Date,Attendance")
            swSnr.WriteLine("Start Date,Start Time,Finish Date,Finish Time,All Day,Name,Detail,Location,Type,Publish Date,Attendance")

            While dr.Read()

                Dim output(10) As String
                Dim outLine As String

                outLine = """"

                For i = 0 To 10
                    If IsDBNull(dr.GetValue(i)) Then
                        If i = 9 Then
                        Else
                            output(i) = "."
                        End If
                    Else
                        output(i) = dr.GetValue(i)
                    End If
                    output(i) = Replace(output(i), "&#039;", "'")
                    output(i) = Replace(output(i), "&amp;", "&")

                    outLine = outLine & output(i) & ""","""
                Next

                outLine = Left(outLine, outLine.Length - 2)


                If Not IsDBNull(dr.GetValue(11)) Then
                    Select Case dr.GetValue(11)
                        Case 1
                            swAll.WriteLine(outLine)
                        Case 61
                            swJnr.WriteLine(outLine)
                        Case 63
                            swSnr.WriteLine(outLine)
                        Case Else
                            swAll.WriteLine(outLine)
                    End Select
                End If


                'sw.WriteLine(outLine)
            End While
            conn.Close()
        End Using
        swAll.Close()
        swJnr.Close()
        swSnr.Close()

    End Sub


    Sub upload(host As String, userName As String, pass As String, rsa As String)
        Try
            ' Setup session options
            Dim sessionOptions As New SessionOptions
            With sessionOptions
                .Protocol = Protocol.Sftp
                .HostName = host
                .UserName = userName
                .Password = pass
                .SshHostKeyFingerprint = rsa '"ssh-rsa 2048 xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx"
            End With

            Using session As New Session
                ' Connect
                session.Open(sessionOptions)

                ' Upload files
                Dim transferOptions As New TransferOptions
                transferOptions.TransferMode = TransferMode.Binary

                Dim transferResult As TransferOperationResult
                transferResult = session.PutFiles(".\*.csv", "./", False, transferOptions)

                ' Throw on any error
                transferResult.Check()

                ' Print results
                For Each transfer In transferResult.Transfers
                    Console.WriteLine("Upload of {0} succeeded", transfer.FileName)
                Next
            End Using


        Catch e As Exception
            Console.WriteLine("Error: {0}", e)

        End Try

    End Sub

    Sub uploadFiles(config As schoolboxConfigSettings)
        For Each i In config.uploadServers
            upload(i.host, i.userName, i.pass, i.rsa)
        Next
    End Sub


    Sub relationships(config As schoolboxConfigSettings)


        Dim commandString As String
        commandString = "


SELECT        
parentcontact.firstname,
parentcontact.surname,
carer.carer_id,
student.student_id,
carer.carer_number,
student.STUDENT_NUMBER 


FROM            relationship

INNER JOIN contact as ParentContact
ON relationship.contact_id1 = Parentcontact.contact_id

INNER JOIN contact as StudentContact 
ON relationship.contact_id2 = studentContact.contact_id

INNER JOIN student
ON studentContact.contact_id = student.contact_id AND
student.student_id in (select student_id from student s1 where current_date BETWEEN (select min(start_date) from student_form_run sfr1a where s1.student_id = sfr1a.student_id) AND 
                                                                                 (select max(end_date) from student_form_run sfr2a where s1.student_id = sfr2a.student_id))

INNER JOIN carer 
ON parentcontact.contact_id = carer.contact_id


WHERE        (relationship.relationship_type_id IN (2, 5, 9, 16, 29, 34)) 


UNION 

SELECT        
parentcontact.firstname,
parentcontact.surname,
carer.carer_id,
student.student_id,
carer.carer_number,
student.STUDENT_NUMBER 


FROM            relationship

INNER JOIN contact as ParentContact
ON relationship.contact_id2 = Parentcontact.contact_id

INNER JOIN contact as StudentContact 
ON relationship.contact_id1 = studentContact.contact_id

INNER JOIN student
ON studentContact.contact_id = student.contact_id AND
student.student_id in (select student_id from student s1 where current_date BETWEEN (select min(start_date) from student_form_run sfr1a where s1.student_id = sfr1a.student_id) AND 
                                                                                 (select max(end_date) from student_form_run sfr2a where s1.student_id = sfr2a.student_id))
INNER JOIN carer 
ON parentcontact.contact_id = carer.contact_id


WHERE        (relationship.relationship_type_id IN (1, 4, 15, 28, 33)) 





"





        Dim sw As New StreamWriter(".\relationship.csv")

        Dim ConnectionString As String = config.connectionString
        Using conn As New IBM.Data.DB2.DB2Connection(ConnectionString)
            conn.Open()

            'define the command object to execute
            Dim command As New IBM.Data.DB2.DB2Command(commandString, conn)
            command.Connection = conn
            command.CommandText = commandString

            Dim dr As IBM.Data.DB2.DB2DataReader
            dr = command.ExecuteReader


            sw.WriteLine("Guardian ID,Student ID,Type")

            Dim fields As Integer = dr.FieldCount - 1
            While dr.Read()
                Dim sb As New StringBuilder()

                Dim outLine As String

                outLine = (dr.GetValue(4) & "," & dr.GetValue(5) & ",Parent")
                sw.WriteLine(outLine)
            End While
            conn.Close()
        End Using


        sw.Close()





    End Sub



End Module
