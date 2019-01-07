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

		'Call timetableStructure(config)
		Call timetable(config)
		Call enrollment(config)
		Call events(config)

        Call uploadFiles(config)


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
						Case "09"
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
						Case "08"
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
						Case "07"
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
						Case "06"
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
						Case "05"
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
						Case "04"
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
						Case "03"
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
						Case "02"
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
						Case "01"
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
                    users.Last.PositionTitle = dr.GetValue(8)
                Else users.Last.PositionTitle = "Staff"
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
                    Case "01"
                    Case "02"
                    Case "03"
                    Case "04"
                    Case "05"
                    Case "06"
                    Case "07"
                    Case "08"
                    Case "09"
                    Case "10"
                    Case "11"
                    Case "12"
                    Case "Parent"
                    Case "Staff"
                    Case Else
                        i.Year = "K"
                End Select




				sw.WriteLine(i.Delete & "," & i.SchoolboxUserID & "," & i.Username & "," & i.ExternalID & "," & i.Title & "," & i.FirstName & "," & i.Surname & "," & i.Role & ",""" & i.Campus & """," & i.Password & "," & i.AltEmail & "," & i.Year & "," & i.House & "," & i.ResidentialHouse & "," & i.EPortfolio & "," & i.HideContactDetails & "," & i.HideTimetable & "," & i.EmailAddressFromUsername & "," & i.UseExternalMailClient & "," & i.EnableWebmailTab & "," & i.AccountEnabled & "," & i.ChildExternalIDs & "," & i.DateOfBirth & "," & i.HomePhone & "," & i.MobilePhone & "," & i.WorkPhone & "," & i.Address & "," & i.Suburb & "," & i.Postcode & "," & i.PositionTitle)

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
WHERE        (start_date > '01/01/2018') AND (end_date < '01/01/2019') AND (term_group.cycle_id = cycle_day.cycle_id) AND 
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

                strTermTitle = Replace(dr.GetValue(1), "2019", "2018")

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
                strTerm = Replace(strTerm, "2019", "2018")
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
SELECT DISTINCT CONCAT(course.code, class.identifier) AS CLASS_CODE, class.class, student.student_number
FROM            CLASS_ENROLLMENT, STUDENT, class, course, academic_year
WHERE        (class_enrollment.student_id = student.student_id) AND (class_enrollment.class_id = class.class_id) AND (class.course_id = course.course_id) AND (class.academic_year_id = academic_year.academic_year_id) AND (academic_year.academic_year ='" & Date.Today.Year & "' OR academic_year.academic_year ='" & Date.Today.Year + 1 & "' ) AND ((SELECT current date FROM sysibm.sysdummy1) between (class_enrollment.start_date -10 DAYS) AND class_enrollment.end_date)
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



			sw.WriteLine("13BUS3,13 Business Studies 3,10019")
			sw.WriteLine("13CHP1,13 Chapel 1,10019")
			sw.WriteLine("13ENS2,13 English Standard 2,10019")
			sw.WriteLine("13EXT3,13 External Studies 3,10019")
			sw.WriteLine("13MAG2,13 General Mathematics 2,10019")
			sw.WriteLine("13MG2,13 Mentor Group 2,10019")
			sw.WriteLine("13MGBE,13 Mentor Group BE,10019")
			sw.WriteLine("13PDH2,13 PDHPE 2,10019")
			sw.WriteLine("13YM2,13 Year Meeting 2,10019")
			sw.WriteLine("13BUS3,13 Business Studies 3,10117")
			sw.WriteLine("13CHP1,13 Chapel 1,10117")
			sw.WriteLine("13CHE1,13 Chemistry 1,10117")
			sw.WriteLine("13ENA1,13 English Advanced 1,10117")
			sw.WriteLine("13MAT1,13 Mathematics 1,10117")
			sw.WriteLine("13MAX1,13 Mathematics Extension 1 1,10117")
			sw.WriteLine("13MG1,13 Mentor Group 1,10117")
			sw.WriteLine("13MGWA,13 Mentor Group WA,10117")
			sw.WriteLine("13PHY1,13 Physics 1,10117")
			sw.WriteLine("13YM1,13 Year Meeting 1,10117")
			sw.WriteLine("13CHP1,13 Chapel 1,10647")
			sw.WriteLine("13ENS2,13 English Standard 2,10647")
			sw.WriteLine("13EXT3,13 External Studies 3,10647")
			sw.WriteLine("13EXT4,13 External Studies 4,10647")
			sw.WriteLine("13MAG2,13 General Mathematics 2,10647")
			sw.WriteLine("13MG1,13 Mentor Group 1,10647")
			sw.WriteLine("13MGMA,13 Mentor Group MA,10647")
			sw.WriteLine("13PDH1,13 PDHPE 1,10647")
			sw.WriteLine("13TAFET,13 TAFE T,10647")
			sw.WriteLine("13YM1,13 Year Meeting 1,10647")
			sw.WriteLine("13BUS3,13 Business Studies 3,10684")
			sw.WriteLine("13CHP1,13 Chapel 1,10684")
			sw.WriteLine("13ENS1,13 English Standard 1,10684")
			sw.WriteLine("13MAG1,13 General Mathematics 1,10684")
			sw.WriteLine("13ITM1,13 ITMM 1,10684")
			sw.WriteLine("13MG1,13 Mentor Group 1,10684")
			sw.WriteLine("13MGBE,13 Mentor Group BE,10684")
			sw.WriteLine("13PDH2,13 PDHPE 2,10684")
			sw.WriteLine("13YM1,13 Year Meeting 1,10684")
			sw.WriteLine("13CHP1,13 Chapel 1,12821")
			sw.WriteLine("13TX1,13 DT TX 1,12821")
			sw.WriteLine("13ENS2,13 English Standard 2,12821")
			sw.WriteLine("13MG2,13 Mentor Group 2,12821")
			sw.WriteLine("13MGBE,13 Mentor Group BE,12821")
			sw.WriteLine("13ART1,13 Visual Arts 1,12821")
			sw.WriteLine("13YM2,13 Year Meeting 2,12821")
			sw.WriteLine("13BUS2,13 Business Studies 2,13443")
			sw.WriteLine("13CHP1,13 Chapel 1,13443")
			sw.WriteLine("13ENA1,13 English Advanced 1,13443")
			sw.WriteLine("13ENX1,13 English Extension 1 1,13443")
			sw.WriteLine("13MAG1,13 General Mathematics 1,13443")
			sw.WriteLine("13LEG1,13 Legal Studies 1,13443")
			sw.WriteLine("13MG1,13 Mentor Group 1,13443")
			sw.WriteLine("13MGBR,13 Mentor Group BR,13443")
			sw.WriteLine("13SOR1,13 Studies of Religion 1,13443")
			sw.WriteLine("13YM1,13 Year Meeting 1,13443")
			sw.WriteLine("13BUS1,13 Business Studies 1,13751")
			sw.WriteLine("13CHP1,13 Chapel 1,13751")
			sw.WriteLine("13DAT1,13 Design and Technology 1,13751")
			sw.WriteLine("13DRA1,13 Drama 1,13751")
			sw.WriteLine("13ENA3,13 English Advanced 3,13751")
			sw.WriteLine("13ENX1,13 English Extension 1 1,13751")
			sw.WriteLine("13MAG1,13 General Mathematics 1,13751")
			sw.WriteLine("13MG1,13 Mentor Group 1,13751")
			sw.WriteLine("13MGBR,13 Mentor Group BR,13751")
			sw.WriteLine("13YM1,13 Year Meeting 1,13751")
			sw.WriteLine("13BUS2,13 Business Studies 2,15459")
			sw.WriteLine("13CHP1,13 Chapel 1,15459")
			sw.WriteLine("13ENA2,13 English Advanced 2,15459")
			sw.WriteLine("13ENX1,13 English Extension 1 1,15459")
			sw.WriteLine("13FTE1,13 Food Technology 1,15459")
			sw.WriteLine("13MAG1,13 General Mathematics 1,15459")
			sw.WriteLine("13MG2,13 Mentor Group 2,15459")
			sw.WriteLine("13MGBE,13 Mentor Group BE,15459")
			sw.WriteLine("13ART1,13 Visual Arts 1,15459")
			sw.WriteLine("13YM2,13 Year Meeting 2,15459")
			sw.WriteLine("13ANC1,13 Ancient History 1,15566")
			sw.WriteLine("13CHP1,13 Chapel 1,15566")
			sw.WriteLine("13ENS2,13 English Standard 2,15566")
			sw.WriteLine("13MAG1,13 General Mathematics 1,15566")
			sw.WriteLine("13ITM1,13 ITMM 1,15566")
			sw.WriteLine("13MG2,13 Mentor Group 2,15566")
			sw.WriteLine("13MGBE,13 Mentor Group BE,15566")
			sw.WriteLine("13ART1,13 Visual Arts 1,15566")
			sw.WriteLine("13YM2,13 Year Meeting 2,15566")
			sw.WriteLine("13BUS3,13 Business Studies 3,16102")
			sw.WriteLine("13CHP1,13 Chapel 1,16102")
			sw.WriteLine("13DAT1,13 Design and Technology 1,16102")
			sw.WriteLine("13ENS1,13 English Standard 1,16102")
			sw.WriteLine("13MAG2,13 General Mathematics 2,16102")
			sw.WriteLine("13MG1,13 Mentor Group 1,16102")
			sw.WriteLine("13MGBE,13 Mentor Group BE,16102")
			sw.WriteLine("13PDH2,13 PDHPE 2,16102")
			sw.WriteLine("13YM1,13 Year Meeting 1,16102")
			sw.WriteLine("13CHP1,13 Chapel 1,16104")
			sw.WriteLine("13ENA1,13 English Advanced 1,16104")
			sw.WriteLine("13ENX1,13 English Extension 1 1,16104")
			sw.WriteLine("13EXX2,13 English Extension 2 2,16104")
			sw.WriteLine("13EXX2jf,13 English Extension 2 2jf,16104")
			sw.WriteLine("13LEG1,13 Legal Studies 1,16104")
			sw.WriteLine("13MG1,13 Mentor Group 1,16104")
			sw.WriteLine("13MGMA,13 Mentor Group MA,16104")
			sw.WriteLine("13MU11,13 Music 1 1,16104")
			sw.WriteLine("13OHS1,13 Open High School 1,16104")
			sw.WriteLine("13SOR1,13 Studies of Religion 1,16104")
			sw.WriteLine("13YM1,13 Year Meeting 1,16104")
			sw.WriteLine("13BUS1,13 Business Studies 1,3405")
			sw.WriteLine("13CHP1,13 Chapel 1,3405")
			sw.WriteLine("13ENA3,13 English Advanced 3,3405")
			sw.WriteLine("13MAG2,13 General Mathematics 2,3405")
			sw.WriteLine("13GEO2,13 Geography 2,3405")
			sw.WriteLine("13HISX1,13 History Extension 1,3405")
			sw.WriteLine("13MG1,13 Mentor Group 1,3405")
			sw.WriteLine("13MGMA,13 Mentor Group MA,3405")
			sw.WriteLine("13MOD1,13 Modern History 1,3405")
			sw.WriteLine("13YM1,13 Year Meeting 1,3405")
			sw.WriteLine("13CHP1,13 Chapel 1,3610")
			sw.WriteLine("13DAT1,13 Design and Technology 1,3610")
			sw.WriteLine("13ENA1,13 English Advanced 1,3610")
			sw.WriteLine("13EXT5,13 External Studies 5,3610")
			sw.WriteLine("13ITM1,13 ITMM 1,3610")
			sw.WriteLine("13MAT1,13 Mathematics 1,3610")
			sw.WriteLine("13MAX1,13 Mathematics Extension 1 1,3610")
			sw.WriteLine("13MG1,13 Mentor Group 1,3610")
			sw.WriteLine("13MGBE,13 Mentor Group BE,3610")
			sw.WriteLine("13YM1,13 Year Meeting 1,3610")
			sw.WriteLine("13BUS1,13 Business Studies 1,3659")
			sw.WriteLine("13CHP1,13 Chapel 1,3659")
			sw.WriteLine("13DRA1,13 Drama 1,3659")
			sw.WriteLine("13ENA1,13 English Advanced 1,3659")
			sw.WriteLine("13MAG1,13 General Mathematics 1,3659")
			sw.WriteLine("13MG1,13 Mentor Group 1,3659")
			sw.WriteLine("13MGBE,13 Mentor Group BE,3659")
			sw.WriteLine("13MU11,13 Music 1 1,3659")
			sw.WriteLine("13YM1,13 Year Meeting 1,3659")
			sw.WriteLine("13CHP1,13 Chapel 1,3705")
			sw.WriteLine("13ENA2,13 English Advanced 2,3705")
			sw.WriteLine("13ENX1,13 English Extension 1 1,3705")
			sw.WriteLine("13MAG1,13 General Mathematics 1,3705")
			sw.WriteLine("13GEO2,13 Geography 2,3705")
			sw.WriteLine("13ITM1,13 ITMM 1,3705")
			sw.WriteLine("13MG2,13 Mentor Group 2,3705")
			sw.WriteLine("13MGBE,13 Mentor Group BE,3705")
			sw.WriteLine("13YM2,13 Year Meeting 2,3705")
			sw.WriteLine("13BIO2,13 Biology 2,3948")
			sw.WriteLine("13CHP1,13 Chapel 1,3948")
			sw.WriteLine("13CHE1,13 Chemistry 1,3948")
			sw.WriteLine("13ENA1,13 English Advanced 1,3948")
			sw.WriteLine("13MAT2,13 Mathematics 2,3948")
			sw.WriteLine("13MG1,13 Mentor Group 1,3948")
			sw.WriteLine("13MGWA,13 Mentor Group WA,3948")
			sw.WriteLine("13PDH2,13 PDHPE 2,3948")
			sw.WriteLine("13YM1,13 Year Meeting 1,3948")
			sw.WriteLine("13BUS3,13 Business Studies 3,4177")
			sw.WriteLine("13CHP1,13 Chapel 1,4177")
			sw.WriteLine("13ENS2,13 English Standard 2,4177")
			sw.WriteLine("13MAG1,13 General Mathematics 1,4177")
			sw.WriteLine("13GEO2,13 Geography 2,4177")
			sw.WriteLine("13ITM1,13 ITMM 1,4177")
			sw.WriteLine("13MG1,13 Mentor Group 1,4177")
			sw.WriteLine("13MGWA,13 Mentor Group WA,4177")
			sw.WriteLine("13YM1,13 Year Meeting 1,4177")
			sw.WriteLine("13BUS2,13 Business Studies 2,4178")
			sw.WriteLine("13CHP1,13 Chapel 1,4178")
			sw.WriteLine("13ENS1,13 English Standard 1,4178")
			sw.WriteLine("13ITM1,13 ITMM 1,4178")
			sw.WriteLine("13MG1,13 Mentor Group 1,4178")
			sw.WriteLine("13MGWA,13 Mentor Group WA,4178")
			sw.WriteLine("13PDH1,13 PDHPE 1,4178")
			sw.WriteLine("13TAFET,13 TAFE T,4178")
			sw.WriteLine("13YM1,13 Year Meeting 1,4178")
			sw.WriteLine("13ANC1,13 Ancient History 1,4179")
			sw.WriteLine("13BUS2,13 Business Studies 2,4179")
			sw.WriteLine("13CHP1,13 Chapel 1,4179")
			sw.WriteLine("13DRA1,13 Drama 1,4179")
			sw.WriteLine("13ENS1,13 English Standard 1,4179")
			sw.WriteLine("13MAG1,13 General Mathematics 1,4179")
			sw.WriteLine("13MG1,13 Mentor Group 1,4179")
			sw.WriteLine("13MGBR,13 Mentor Group BR,4179")
			sw.WriteLine("13YM1,13 Year Meeting 1,4179")
			sw.WriteLine("13CHP1,13 Chapel 1,4180")
			sw.WriteLine("13DRA1,13 Drama 1,4180")
			sw.WriteLine("13ENS1,13 English Standard 1,4180")
			sw.WriteLine("13MAG2,13 General Mathematics 2,4180")
			sw.WriteLine("13MG1,13 Mentor Group 1,4180")
			sw.WriteLine("13MGBE,13 Mentor Group BE,4180")
			sw.WriteLine("13TAFET,13 TAFE T,4180")
			sw.WriteLine("13YM1,13 Year Meeting 1,4180")
			sw.WriteLine("13BUS3,13 Business Studies 3,4185")
			sw.WriteLine("13CHP1,13 Chapel 1,4185")
			sw.WriteLine("13ENA1,13 English Advanced 1,4185")
			sw.WriteLine("13EXT3,13 External Studies 3,4185")
			sw.WriteLine("13MAG1,13 General Mathematics 1,4185")
			sw.WriteLine("13MG2,13 Mentor Group 2,4185")
			sw.WriteLine("13MGBR,13 Mentor Group BR,4185")
			sw.WriteLine("13MOD1,13 Modern History 1,4185")
			sw.WriteLine("13YM2,13 Year Meeting 2,4185")
			sw.WriteLine("13BUS1,13 Business Studies 1,4191")
			sw.WriteLine("13CHP1,13 Chapel 1,4191")
			sw.WriteLine("13ECO1,13 Economics 1,4191")
			sw.WriteLine("13ENS1,13 English Standard 1,4191")
			sw.WriteLine("13MAG1,13 General Mathematics 1,4191")
			sw.WriteLine("13HISX1,13 History Extension 1,4191")
			sw.WriteLine("13MG1,13 Mentor Group 1,4191")
			sw.WriteLine("13MGWA,13 Mentor Group WA,4191")
			sw.WriteLine("13MOD1,13 Modern History 1,4191")
			sw.WriteLine("13YM1,13 Year Meeting 1,4191")
			sw.WriteLine("13ANC1,13 Ancient History 1,4199")
			sw.WriteLine("13CHP1,13 Chapel 1,4199")
			sw.WriteLine("13ECO1,13 Economics 1,4199")
			sw.WriteLine("13ENA2,13 English Advanced 2,4199")
			sw.WriteLine("13FRB1,13 French Beginners 1,4199")
			sw.WriteLine("13MAG2,13 General Mathematics 2,4199")
			sw.WriteLine("13MG2,13 Mentor Group 2,4199")
			sw.WriteLine("13MGMA,13 Mentor Group MA,4199")
			sw.WriteLine("13YM2,13 Year Meeting 2,4199")
			sw.WriteLine("13BIO2,13 Biology 2,4230")
			sw.WriteLine("13CHP1,13 Chapel 1,4230")
			sw.WriteLine("13CHE1,13 Chemistry 1,4230")
			sw.WriteLine("13ENS1,13 English Standard 1,4230")
			sw.WriteLine("13EXT4,13 External Studies 4,4230")
			sw.WriteLine("13MAT1,13 Mathematics 1,4230")
			sw.WriteLine("13MAX1,13 Mathematics Extension 1 1,4230")
			sw.WriteLine("JSSCI1,13 Mathematics Extension 2 1,4230")
			sw.WriteLine("13MXX2,13 Mathematics Extension 2 2,4230")
			sw.WriteLine("13MG1,13 Mentor Group 1,4230")
			sw.WriteLine("13MGBR,13 Mentor Group BR,4230")
			sw.WriteLine("13PHY1,13 Physics 1,4230")
			sw.WriteLine("13YM1,13 Year Meeting 1,4230")
			sw.WriteLine("13BUS1,13 Business Studies 1,4237")
			sw.WriteLine("13CHP1,13 Chapel 1,4237")
			sw.WriteLine("13ENA3,13 English Advanced 3,4237")
			sw.WriteLine("13MAG1,13 General Mathematics 1,4237")
			sw.WriteLine("13LEG1,13 Legal Studies 1,4237")
			sw.WriteLine("13MG1,13 Mentor Group 1,4237")
			sw.WriteLine("13MGBR,13 Mentor Group BR,4237")
			sw.WriteLine("13MOD1,13 Modern History 1,4237")
			sw.WriteLine("13YM1,13 Year Meeting 1,4237")
			sw.WriteLine("13BIO2,13 Biology 2,4238")
			sw.WriteLine("13CHP1,13 Chapel 1,4238")
			sw.WriteLine("13ENA3,13 English Advanced 3,4238")
			sw.WriteLine("13MAG1,13 General Mathematics 1,4238")
			sw.WriteLine("13MG2,13 Mentor Group 2,4238")
			sw.WriteLine("13MGBR,13 Mentor Group BR,4238")
			sw.WriteLine("13MU11,13 Music 1 1,4238")
			sw.WriteLine("13PDH2,13 PDHPE 2,4238")
			sw.WriteLine("13YM2,13 Year Meeting 2,4238")
			sw.WriteLine("13BUS3,13 Business Studies 3,4440")
			sw.WriteLine("13CHP1,13 Chapel 1,4440")
			sw.WriteLine("13ENS2,13 English Standard 2,4440")
			sw.WriteLine("13MAG1,13 General Mathematics 1,4440")
			sw.WriteLine("13ITM1,13 ITMM 1,4440")
			sw.WriteLine("13MG2,13 Mentor Group 2,4440")
			sw.WriteLine("13MGBE,13 Mentor Group BE,4440")
			sw.WriteLine("13MOD1,13 Modern History 1,4440")
			sw.WriteLine("13YM2,13 Year Meeting 2,4440")
			sw.WriteLine("13BIO1,13 Biology 1,4715")
			sw.WriteLine("13BUS3,13 Business Studies 3,4715")
			sw.WriteLine("13CHP1,13 Chapel 1,4715")
			sw.WriteLine("13ENS1,13 English Standard 1,4715")
			sw.WriteLine("13MAG2,13 General Mathematics 2,4715")
			sw.WriteLine("13GEO2,13 Geography 2,4715")
			sw.WriteLine("13MG1,13 Mentor Group 1,4715")
			sw.WriteLine("13MGWA,13 Mentor Group WA,4715")
			sw.WriteLine("13YM1,13 Year Meeting 1,4715")
			sw.WriteLine("13BUS3,13 Business Studies 3,5030")
			sw.WriteLine("13CHP1,13 Chapel 1,5030")
			sw.WriteLine("13ENS2,13 English Standard 2,5030")
			sw.WriteLine("13FTE1,13 Food Technology 1,5030")
			sw.WriteLine("13MG2,13 Mentor Group 2,5030")
			sw.WriteLine("13MGWA,13 Mentor Group WA,5030")
			sw.WriteLine("13MOD1,13 Modern History 1,5030")
			sw.WriteLine("13PDH2,13 PDHPE 2,5030")
			sw.WriteLine("13YM2,13 Year Meeting 2,5030")
			sw.WriteLine("13BIO1,13 Biology 1,5172")
			sw.WriteLine("13CHP1,13 Chapel 1,5172")
			sw.WriteLine("13DRA1,13 Drama 1,5172")
			sw.WriteLine("13ENA2,13 English Advanced 2,5172")
			sw.WriteLine("13ENX1,13 English Extension 1 1,5172")
			sw.WriteLine("13MAG1,13 General Mathematics 1,5172")
			sw.WriteLine("13LEG1,13 Legal Studies 1,5172")
			sw.WriteLine("13MG2,13 Mentor Group 2,5172")
			sw.WriteLine("13MGBR,13 Mentor Group BR,5172")
			sw.WriteLine("13SOR1,13 Studies of Religion 1,5172")
			sw.WriteLine("13YM2,13 Year Meeting 2,5172")
			sw.WriteLine("13BUS1,13 Business Studies 1,5553")
			sw.WriteLine("13CHP1,13 Chapel 1,5553")
			sw.WriteLine("13DRA1,13 Drama 1,5553")
			sw.WriteLine("13ENA1,13 English Advanced 1,5553")
			sw.WriteLine("13ENX1,13 English Extension 1 1,5553")
			sw.WriteLine("13EXX1as,13 English Extension 2 1as,5553")
			sw.WriteLine("13MG2,13 Mentor Group 2,5553")
			sw.WriteLine("13MGMA,13 Mentor Group MA,5553")
			sw.WriteLine("13MU11,13 Music 1 1,5553")
			sw.WriteLine("13SOR1,13 Studies of Religion 1,5553")
			sw.WriteLine("13YM2,13 Year Meeting 2,5553")
			sw.WriteLine("13BUS2,13 Business Studies 2,6036")
			sw.WriteLine("13CHP1,13 Chapel 1,6036")
			sw.WriteLine("13ECO1,13 Economics 1,6036")
			sw.WriteLine("13ENA2,13 English Advanced 2,6036")
			sw.WriteLine("13MAG2,13 General Mathematics 2,6036")
			sw.WriteLine("13ITM1,13 ITMM 1,6036")
			sw.WriteLine("13LEG1,13 Legal Studies 1,6036")
			sw.WriteLine("13MG1,13 Mentor Group 1,6036")
			sw.WriteLine("13MGBE,13 Mentor Group BE,6036")
			sw.WriteLine("13YM1,13 Year Meeting 1,6036")
			sw.WriteLine("13BIO1,13 Biology 1,6219")
			sw.WriteLine("13BUS3,13 Business Studies 3,6219")
			sw.WriteLine("13CHP1,13 Chapel 1,6219")
			sw.WriteLine("13ENS2,13 English Standard 2,6219")
			sw.WriteLine("13MAG2,13 General Mathematics 2,6219")
			sw.WriteLine("13MG1,13 Mentor Group 1,6219")
			sw.WriteLine("13MGMA,13 Mentor Group MA,6219")
			sw.WriteLine("13PDH2,13 PDHPE 2,6219")
			sw.WriteLine("13YM1,13 Year Meeting 1,6219")
			sw.WriteLine("13BIO1,13 Biology 1,6272")
			sw.WriteLine("13BUS3,13 Business Studies 3,6272")
			sw.WriteLine("13CHP1,13 Chapel 1,6272")
			sw.WriteLine("13ENA2,13 English Advanced 2,6272")
			sw.WriteLine("13MAG1,13 General Mathematics 1,6272")
			sw.WriteLine("13LEG1,13 Legal Studies 1,6272")
			sw.WriteLine("13MG2,13 Mentor Group 2,6272")
			sw.WriteLine("13MGBR,13 Mentor Group BR,6272")
			sw.WriteLine("13YM2,13 Year Meeting 2,6272")
			sw.WriteLine("13ANC1,13 Ancient History 1,6278")
			sw.WriteLine("13CHP1,13 Chapel 1,6278")
			sw.WriteLine("13ENA3,13 English Advanced 3,6278")
			sw.WriteLine("13MAT1,13 Mathematics 1,6278")
			sw.WriteLine("13MAX1,13 Mathematics Extension 1 1,6278")
			sw.WriteLine("13MG2,13 Mentor Group 2,6278")
			sw.WriteLine("13MGMA,13 Mentor Group MA,6278")
			sw.WriteLine("13MOD1,13 Modern History 1,6278")
			sw.WriteLine("13YM2,13 Year Meeting 2,6278")
			sw.WriteLine("13BIO1,13 Biology 1,6570")
			sw.WriteLine("13CHP1,13 Chapel 1,6570")
			sw.WriteLine("13CHE1,13 Chemistry 1,6570")
			sw.WriteLine("13ENA2,13 English Advanced 2,6570")
			sw.WriteLine("13MAG1,13 General Mathematics 1,6570")
			sw.WriteLine("13MG1,13 Mentor Group 1,6570")
			sw.WriteLine("13MGBE,13 Mentor Group BE,6570")
			sw.WriteLine("13PDH2,13 PDHPE 2,6570")
			sw.WriteLine("13YM1,13 Year Meeting 1,6570")
			sw.WriteLine("13BUS2,13 Business Studies 2,6626")
			sw.WriteLine("13CHP1,13 Chapel 1,6626")
			sw.WriteLine("13ENA2,13 English Advanced 2,6626")
			sw.WriteLine("13LEG1,13 Legal Studies 1,6626")
			sw.WriteLine("13MAT1,13 Mathematics 1,6626")
			sw.WriteLine("13MAX1,13 Mathematics Extension 1 1,6626")
			sw.WriteLine("JSSCI1,13 Mathematics Extension 2 1,6626")
			sw.WriteLine("13MXX2,13 Mathematics Extension 2 2,6626")
			sw.WriteLine("13MG1,13 Mentor Group 1,6626")
			sw.WriteLine("13MGWA,13 Mentor Group WA,6626")
			sw.WriteLine("13YM1,13 Year Meeting 1,6626")
			sw.WriteLine("13ANC1,13 Ancient History 1,6877")
			sw.WriteLine("13BIO2,13 Biology 2,6877")
			sw.WriteLine("13CHP1,13 Chapel 1,6877")
			sw.WriteLine("13ENA2,13 English Advanced 2,6877")
			sw.WriteLine("13GEO1,13 Geography 1,6877")
			sw.WriteLine("13MAT2,13 Mathematics 2,6877")
			sw.WriteLine("13MG2,13 Mentor Group 2,6877")
			sw.WriteLine("13MGWA,13 Mentor Group WA,6877")
			sw.WriteLine("13YM2,13 Year Meeting 2,6877")
			sw.WriteLine("13BIO1,13 Biology 1,7072")
			sw.WriteLine("13BUS3,13 Business Studies 3,7072")
			sw.WriteLine("13CHP1,13 Chapel 1,7072")
			sw.WriteLine("13CHE1,13 Chemistry 1,7072")
			sw.WriteLine("13ENA2,13 English Advanced 2,7072")
			sw.WriteLine("13MAG2,13 General Mathematics 2,7072")
			sw.WriteLine("13MG2,13 Mentor Group 2,7072")
			sw.WriteLine("13MGBR,13 Mentor Group BR,7072")
			sw.WriteLine("13YM2,13 Year Meeting 2,7072")
			sw.WriteLine("13CHP1,13 Chapel 1,7442")
			sw.WriteLine("13DAT1,13 Design and Technology 1,7442")
			sw.WriteLine("13ENS1,13 English Standard 1,7442")
			sw.WriteLine("13GEO1,13 Geography 1,7442")
			sw.WriteLine("13MAT2,13 Mathematics 2,7442")
			sw.WriteLine("13MG1,13 Mentor Group 1,7442")
			sw.WriteLine("13MGBE,13 Mentor Group BE,7442")
			sw.WriteLine("13PDH1,13 PDHPE 1,7442")
			sw.WriteLine("13YM1,13 Year Meeting 1,7442")
			sw.WriteLine("13CHP1,13 Chapel 1,7598")
			sw.WriteLine("13CHE1,13 Chemistry 1,7598")
			sw.WriteLine("13DAT1,13 Design and Technology 1,7598")
			sw.WriteLine("13ENA3,13 English Advanced 3,7598")
			sw.WriteLine("13MAT2,13 Mathematics 2,7598")
			sw.WriteLine("13MG1,13 Mentor Group 1,7598")
			sw.WriteLine("13MGBR,13 Mentor Group BR,7598")
			sw.WriteLine("13PHY1,13 Physics 1,7598")
			sw.WriteLine("13YM1,13 Year Meeting 1,7598")
			sw.WriteLine("13BIO1,13 Biology 1,7769")
			sw.WriteLine("13CHP1,13 Chapel 1,7769")
			sw.WriteLine("13CHE1,13 Chemistry 1,7769")
			sw.WriteLine("13ENS1,13 English Standard 1,7769")
			sw.WriteLine("13MAT2,13 Mathematics 2,7769")
			sw.WriteLine("13MG1,13 Mentor Group 1,7769")
			sw.WriteLine("13MGBE,13 Mentor Group BE,7769")
			sw.WriteLine("13PDH1,13 PDHPE 1,7769")
			sw.WriteLine("13YM1,13 Year Meeting 1,7769")
			sw.WriteLine("13ANC1,13 Ancient History 1,7800")
			sw.WriteLine("13BIO2,13 Biology 2,7800")
			sw.WriteLine("13CHP1,13 Chapel 1,7800")
			sw.WriteLine("13ENA3,13 English Advanced 3,7800")
			sw.WriteLine("13ENX1,13 English Extension 1 1,7800")
			sw.WriteLine("13MAG1,13 General Mathematics 1,7800")
			sw.WriteLine("13HISX1,13 History Extension 1,7800")
			sw.WriteLine("13MG2,13 Mentor Group 2,7800")
			sw.WriteLine("13MGWA,13 Mentor Group WA,7800")
			sw.WriteLine("13YM2,13 Year Meeting 2,7800")
			sw.WriteLine("13CHP1,13 Chapel 1,7930")
			sw.WriteLine("13DRA1,13 Drama 1,7930")
			sw.WriteLine("13ENA3,13 English Advanced 3,7930")
			sw.WriteLine("13FRB1,13 French Beginners 1,7930")
			sw.WriteLine("13MAG1,13 General Mathematics 1,7930")
			sw.WriteLine("13MG2,13 Mentor Group 2,7930")
			sw.WriteLine("13MGWA,13 Mentor Group WA,7930")
			sw.WriteLine("13MOD1,13 Modern History 1,7930")
			sw.WriteLine("13YM2,13 Year Meeting 2,7930")
			sw.WriteLine("13ANC1,13 Ancient History 1,7996")
			sw.WriteLine("13BIO2,13 Biology 2,7996")
			sw.WriteLine("13CHP1,13 Chapel 1,7996")
			sw.WriteLine("13CHE1,13 Chemistry 1,7996")
			sw.WriteLine("13ENA2,13 English Advanced 2,7996")
			sw.WriteLine("13MAT2,13 Mathematics 2,7996")
			sw.WriteLine("13MG2,13 Mentor Group 2,7996")
			sw.WriteLine("13MGMA,13 Mentor Group MA,7996")
			sw.WriteLine("13SOR1,13 Studies of Religion 1,7996")
			sw.WriteLine("13YM2,13 Year Meeting 2,7996")
			sw.WriteLine("13BUS2,13 Business Studies 2,8084")
			sw.WriteLine("13CHP1,13 Chapel 1,8084")
			sw.WriteLine("13ECO1,13 Economics 1,8084")
			sw.WriteLine("13ENS1,13 English Standard 1,8084")
			sw.WriteLine("13GEO1,13 Geography 1,8084")
			sw.WriteLine("13MAT2,13 Mathematics 2,8084")
			sw.WriteLine("13MG1,13 Mentor Group 1,8084")
			sw.WriteLine("13MGMA,13 Mentor Group MA,8084")
			sw.WriteLine("13YM1,13 Year Meeting 1,8084")
			sw.WriteLine("13ANC1,13 Ancient History 1,8179")
			sw.WriteLine("13CHP1,13 Chapel 1,8179")
			sw.WriteLine("13ENS1,13 English Standard 1,8179")
			sw.WriteLine("13MAG2,13 General Mathematics 2,8179")
			sw.WriteLine("13MG1,13 Mentor Group 1,8179")
			sw.WriteLine("13MGBR,13 Mentor Group BR,8179")
			sw.WriteLine("13MU11,13 Music 1 1,8179")
			sw.WriteLine("13ART1,13 Visual Arts 1,8179")
			sw.WriteLine("13YM1,13 Year Meeting 1,8179")
			sw.WriteLine("13BUS3,13 Business Studies 3,8273")
			sw.WriteLine("13CHP1,13 Chapel 1,8273")
			sw.WriteLine("13ENS1,13 English Standard 1,8273")
			sw.WriteLine("13MAG1,13 General Mathematics 1,8273")
			sw.WriteLine("13MG1,13 Mentor Group 1,8273")
			sw.WriteLine("13MGWA,13 Mentor Group WA,8273")
			sw.WriteLine("13MOD1,13 Modern History 1,8273")
			sw.WriteLine("13TAFET,13 TAFE T,8273")
			sw.WriteLine("13YM1,13 Year Meeting 1,8273")
			sw.WriteLine("13CHP1,13 Chapel 1,8281")
			sw.WriteLine("13DAT1,13 Design and Technology 1,8281")
			sw.WriteLine("13ENS2,13 English Standard 2,8281")
			sw.WriteLine("13EXT1,13 External Studies 1,8281")
			sw.WriteLine("13FTE1,13 Food Technology 1,8281")
			sw.WriteLine("13MG2,13 Mentor Group 2,8281")
			sw.WriteLine("13MGBR,13 Mentor Group BR,8281")
			sw.WriteLine("13PDH2,13 PDHPE 2,8281")
			sw.WriteLine("13YM2,13 Year Meeting 2,8281")
			sw.WriteLine("13BUS1,13 Business Studies 1,8393")
			sw.WriteLine("13CHP1,13 Chapel 1,8393")
			sw.WriteLine("13ENA3,13 English Advanced 3,8393")
			sw.WriteLine("13EXT1,13 External Studies 1,8393")
			sw.WriteLine("13EXT6,13 External Studies 6,8393")
			sw.WriteLine("13MG2,13 Mentor Group 2,8393")
			sw.WriteLine("13MGMA,13 Mentor Group MA,8393")
			sw.WriteLine("13MOD1,13 Modern History 1,8393")
			sw.WriteLine("13ART1,13 Visual Arts 1,8393")
			sw.WriteLine("13YM2,13 Year Meeting 2,8393")
			sw.WriteLine("13CHP1,13 Chapel 1,8609")
			sw.WriteLine("13ECO1,13 Economics 1,8609")
			sw.WriteLine("13ENA2,13 English Advanced 2,8609")
			sw.WriteLine("13MAT1,13 Mathematics 1,8609")
			sw.WriteLine("13MAX1,13 Mathematics Extension 1 1,8609")
			sw.WriteLine("JSSCI1,13 Mathematics Extension 2 1,8609")
			sw.WriteLine("13MXX2,13 Mathematics Extension 2 2,8609")
			sw.WriteLine("13MG1,13 Mentor Group 1,8609")
			sw.WriteLine("13MGBE,13 Mentor Group BE,8609")
			sw.WriteLine("13PHY1,13 Physics 1,8609")
			sw.WriteLine("13YM1,13 Year Meeting 1,8609")
			sw.WriteLine("13CHP1,13 Chapel 1,8710")
			sw.WriteLine("13CHE1,13 Chemistry 1,8710")
			sw.WriteLine("13ENA1,13 English Advanced 1,8710")
			sw.WriteLine("13MAT1,13 Mathematics 1,8710")
			sw.WriteLine("13MAX1,13 Mathematics Extension 1 1,8710")
			sw.WriteLine("13MG1,13 Mentor Group 1,8710")
			sw.WriteLine("13MGMA,13 Mentor Group MA,8710")
			sw.WriteLine("13MU21,13 Music 2 1,8710")
			sw.WriteLine("13PHY1,13 Physics 1,8710")
			sw.WriteLine("13YM1,13 Year Meeting 1,8710")
			sw.WriteLine("13BIO1,13 Biology 1,8810")
			sw.WriteLine("13BUS1,13 Business Studies 1,8810")
			sw.WriteLine("13CHP1,13 Chapel 1,8810")
			sw.WriteLine("13ECO1,13 Economics 1,8810")
			sw.WriteLine("13ENA2,13 English Advanced 2,8810")
			sw.WriteLine("13MAG2,13 General Mathematics 2,8810")
			sw.WriteLine("13MG1,13 Mentor Group 1,8810")
			sw.WriteLine("13MGMA,13 Mentor Group MA,8810")
			sw.WriteLine("13YM1,13 Year Meeting 1,8810")
			sw.WriteLine("13ANC1,13 Ancient History 1,8822")
			sw.WriteLine("13CHP1,13 Chapel 1,8822")
			sw.WriteLine("13TX1,13 DT TX 1,8822")
			sw.WriteLine("13ENA3,13 English Advanced 3,8822")
			sw.WriteLine("13MAG2,13 General Mathematics 2,8822")
			sw.WriteLine("13MG2,13 Mentor Group 2,8822")
			sw.WriteLine("13MGBE,13 Mentor Group BE,8822")
			sw.WriteLine("13ART1,13 Visual Arts 1,8822")
			sw.WriteLine("13YM2,13 Year Meeting 2,8822")
			sw.WriteLine("13BUS3,13 Business Studies 3,8881")
			sw.WriteLine("13CHP1,13 Chapel 1,8881")
			sw.WriteLine("13ENA2,13 English Advanced 2,8881")
			sw.WriteLine("13ENX1,13 English Extension 1 1,8881")
			sw.WriteLine("13GEO1,13 Geography 1,8881")
			sw.WriteLine("13LEG1,13 Legal Studies 1,8881")
			sw.WriteLine("13MG2,13 Mentor Group 2,8881")
			sw.WriteLine("13MGBR,13 Mentor Group BR,8881")
			sw.WriteLine("13MOD1,13 Modern History 1,8881")
			sw.WriteLine("13YM2,13 Year Meeting 2,8881")
			sw.WriteLine("13BUS2,13 Business Studies 2,9133")
			sw.WriteLine("13CHP1,13 Chapel 1,9133")
			sw.WriteLine("13ENS2,13 English Standard 2,9133")
			sw.WriteLine("13MAG2,13 General Mathematics 2,9133")
			sw.WriteLine("13ITM1,13 ITMM 1,9133")
			sw.WriteLine("13MG1,13 Mentor Group 1,9133")
			sw.WriteLine("13MGBE,13 Mentor Group BE,9133")
			sw.WriteLine("13PHY1,13 Physics 1,9133")
			sw.WriteLine("13TAFET,13 TAFE T,9133")
			sw.WriteLine("13YM1,13 Year Meeting 1,9133")
			sw.WriteLine("13BUS2,13 Business Studies 2,9143")
			sw.WriteLine("13CHP1,13 Chapel 1,9143")
			sw.WriteLine("13ENA1,13 English Advanced 1,9143")
			sw.WriteLine("13MAG1,13 General Mathematics 1,9143")
			sw.WriteLine("13GEO1,13 Geography 1,9143")
			sw.WriteLine("13MG2,13 Mentor Group 2,9143")
			sw.WriteLine("13MGWA,13 Mentor Group WA,9143")
			sw.WriteLine("13ART1,13 Visual Arts 1,9143")
			sw.WriteLine("13YM2,13 Year Meeting 2,9143")
			sw.WriteLine("13BUS2,13 Business Studies 2,9291")
			sw.WriteLine("13CHP1,13 Chapel 1,9291")
			sw.WriteLine("13ENA1,13 English Advanced 1,9291")
			sw.WriteLine("13FTE1,13 Food Technology 1,9291")
			sw.WriteLine("13MAT2,13 Mathematics 2,9291")
			sw.WriteLine("13MG2,13 Mentor Group 2,9291")
			sw.WriteLine("13MGBE,13 Mentor Group BE,9291")
			sw.WriteLine("13PDH2,13 PDHPE 2,9291")
			sw.WriteLine("13YM2,13 Year Meeting 2,9291")
			sw.WriteLine("13CHP1,13 Chapel 1,9297")
			sw.WriteLine("13ECO1,13 Economics 1,9297")
			sw.WriteLine("13ENS2,13 English Standard 2,9297")
			sw.WriteLine("13GEO2,13 Geography 2,9297")
			sw.WriteLine("13ITM1,13 ITMM 1,9297")
			sw.WriteLine("13MAT1,13 Mathematics 1,9297")
			sw.WriteLine("13MAX1,13 Mathematics Extension 1 1,9297")
			sw.WriteLine("13MG1,13 Mentor Group 1,9297")
			sw.WriteLine("13MGBR,13 Mentor Group BR,9297")
			sw.WriteLine("13YM1,13 Year Meeting 1,9297")
			sw.WriteLine("13ANC1,13 Ancient History 1,9317")
			sw.WriteLine("13CHP1,13 Chapel 1,9317")
			sw.WriteLine("13ENS2,13 English Standard 2,9317")
			sw.WriteLine("13MAG1,13 General Mathematics 1,9317")
			sw.WriteLine("13MG1,13 Mentor Group 1,9317")
			sw.WriteLine("13MGBR,13 Mentor Group BR,9317")
			sw.WriteLine("13MU11,13 Music 1 1,9317")
			sw.WriteLine("13ART1,13 Visual Arts 1,9317")
			sw.WriteLine("13YM1,13 Year Meeting 1,9317")
			sw.WriteLine("13CHP1,13 Chapel 1,9441")
			sw.WriteLine("13ENA3,13 English Advanced 3,9441")
			sw.WriteLine("13MAG2,13 General Mathematics 2,9441")
			sw.WriteLine("13GEO1,13 Geography 1,9441")
			sw.WriteLine("13MG2,13 Mentor Group 2,9441")
			sw.WriteLine("13MGBR,13 Mentor Group BR,9441")
			sw.WriteLine("13PDH2,13 PDHPE 2,9441")
			sw.WriteLine("13ART1,13 Visual Arts 1,9441")
			sw.WriteLine("13YM2,13 Year Meeting 2,9441")
			sw.WriteLine("13CHP1,13 Chapel 1,9476")
			sw.WriteLine("13ENS1,13 English Standard 1,9476")
			sw.WriteLine("13MAG2,13 General Mathematics 2,9476")
			sw.WriteLine("13GEO1,13 Geography 1,9476")
			sw.WriteLine("13MG1,13 Mentor Group 1,9476")
			sw.WriteLine("13MGMA,13 Mentor Group MA,9476")
			sw.WriteLine("13MOD1,13 Modern History 1,9476")
			sw.WriteLine("13PDH1,13 PDHPE 1,9476")
			sw.WriteLine("13YM1,13 Year Meeting 1,9476")
			sw.WriteLine("13BUS3,13 Business Studies 3,9531")
			sw.WriteLine("13CHP1,13 Chapel 1,9531")
			sw.WriteLine("13DAT1,13 Design and Technology 1,9531")
			sw.WriteLine("13ENS2,13 English Standard 2,9531")
			sw.WriteLine("13MAG2,13 General Mathematics 2,9531")
			sw.WriteLine("13GEO1,13 Geography 1,9531")
			sw.WriteLine("13MG1,13 Mentor Group 1,9531")
			sw.WriteLine("13MGBR,13 Mentor Group BR,9531")
			sw.WriteLine("13YM1,13 Year Meeting 1,9531")
			sw.WriteLine("13ANC1,13 Ancient History 1,9549")
			sw.WriteLine("13CHP1,13 Chapel 1,9549")
			sw.WriteLine("13TX1,13 DT TX 1,9549")
			sw.WriteLine("13ENA1,13 English Advanced 1,9549")
			sw.WriteLine("13ENX1,13 English Extension 1 1,9549")
			sw.WriteLine("13EXX2oa,13 English Extension 2 2oa,9549")
			sw.WriteLine("13GEO1,13 Geography 1,9549")
			sw.WriteLine("13HISX1,13 History Extension 1,9549")
			sw.WriteLine("13MG2,13 Mentor Group 2,9549")
			sw.WriteLine("13MGWA,13 Mentor Group WA,9549")
			sw.WriteLine("13SOR1,13 Studies of Religion 1,9549")
			sw.WriteLine("13YM2,13 Year Meeting 2,9549")
			sw.WriteLine("13ANC1,13 Ancient History 1,9627")
			sw.WriteLine("13BUS3,13 Business Studies 3,9627")
			sw.WriteLine("13CHP1,13 Chapel 1,9627")
			sw.WriteLine("13ENA1,13 English Advanced 1,9627")
			sw.WriteLine("13ENX1,13 English Extension 1 1,9627")
			sw.WriteLine("13EXT4,13 External Studies 4,9627")
			sw.WriteLine("13MG2,13 Mentor Group 2,9627")
			sw.WriteLine("13MGBE,13 Mentor Group BE,9627")
			sw.WriteLine("13OHS1,13 Open High School 1,9627")
			sw.WriteLine("13YM2,13 Year Meeting 2,9627")
			sw.WriteLine("13BUS2,13 Business Studies 2,9714")
			sw.WriteLine("13CHP1,13 Chapel 1,9714")
			sw.WriteLine("13DRA1,13 Drama 1,9714")
			sw.WriteLine("13ENS1,13 English Standard 1,9714")
			sw.WriteLine("13MAG2,13 General Mathematics 2,9714")
			sw.WriteLine("13ITM1,13 ITMM 1,9714")
			sw.WriteLine("13MG1,13 Mentor Group 1,9714")
			sw.WriteLine("13MGWA,13 Mentor Group WA,9714")
			sw.WriteLine("13YM1,13 Year Meeting 1,9714")
			sw.WriteLine("13CHP1,13 Chapel 1,9772")
			sw.WriteLine("13ENS2,13 English Standard 2,9772")
			sw.WriteLine("13FTE1,13 Food Technology 1,9772")
			sw.WriteLine("13MAG2,13 General Mathematics 2,9772")
			sw.WriteLine("13MG2,13 Mentor Group 2,9772")
			sw.WriteLine("13MGBR,13 Mentor Group BR,9772")
			sw.WriteLine("13MOD1,13 Modern History 1,9772")
			sw.WriteLine("13PDH1,13 PDHPE 1,9772")
			sw.WriteLine("13YM2,13 Year Meeting 2,9772")
			sw.WriteLine("13ANC1,13 Ancient History 1,9797")
			sw.WriteLine("13CHP1,13 Chapel 1,9797")
			sw.WriteLine("13ENA1,13 English Advanced 1,9797")
			sw.WriteLine("13ENX1,13 English Extension 1 1,9797")
			sw.WriteLine("13EXX1,13 English Extension 2 1,9797")
			sw.WriteLine("JSSCI1ab,13 English Extension 2 1ab,9797")
			sw.WriteLine("13EXXab,13 English Extension 2 ab,9797")
			sw.WriteLine("13HISX1,13 History Extension 1,9797")
			sw.WriteLine("13MG2,13 Mentor Group 2,9797")
			sw.WriteLine("13MGBE,13 Mentor Group BE,9797")
			sw.WriteLine("13MOD1,13 Modern History 1,9797")
			sw.WriteLine("13ART1,13 Visual Arts 1,9797")
			sw.WriteLine("13YM2,13 Year Meeting 2,9797")
			sw.WriteLine("13CHP1,13 Chapel 1,9915")
			sw.WriteLine("13DRA1,13 Drama 1,9915")
			sw.WriteLine("13ENA3,13 English Advanced 3,9915")
			sw.WriteLine("13ENX1,13 English Extension 1 1,9915")
			sw.WriteLine("13FRB1,13 French Beginners 1,9915")
			sw.WriteLine("13MG2,13 Mentor Group 2,9915")
			sw.WriteLine("13MGMA,13 Mentor Group MA,9915")
			sw.WriteLine("13MU21,13 Music 2 1,9915")
			sw.WriteLine("13MUX1,13 Music Extension 1,9915")
			sw.WriteLine("13YM2,13 Year Meeting 2,9915")
			sw.WriteLine("13BUS1,13 Business Studies 1,997443155")
			sw.WriteLine("13CHP1,13 Chapel 1,997443155")
			sw.WriteLine("13ENA1,13 English Advanced 1,997443155")
			sw.WriteLine("13ENX1,13 English Extension 1 1,997443155")
			sw.WriteLine("13MG2,13 Mentor Group 2,997443155")
			sw.WriteLine("13MGBR,13 Mentor Group BR,997443155")
			sw.WriteLine("13MOD1,13 Modern History 1,997443155")
			sw.WriteLine("13SOR1,13 Studies of Religion 1,997443155")
			sw.WriteLine("13ART1,13 Visual Arts 1,997443155")
			sw.WriteLine("13YM2,13 Year Meeting 2,997443155")
			sw.WriteLine("13CHP1,13 Chapel 1,997443316")
			sw.WriteLine("13CHE1,13 Chemistry 1,997443316")
			sw.WriteLine("13ENA3,13 English Advanced 3,997443316")
			sw.WriteLine("13MAT2,13 Mathematics 2,997443316")
			sw.WriteLine("13MG1,13 Mentor Group 1,997443316")
			sw.WriteLine("13MGWA,13 Mentor Group WA,997443316")
			sw.WriteLine("13MU11,13 Music 1 1,997443316")
			sw.WriteLine("13PHY1,13 Physics 1,997443316")
			sw.WriteLine("13SOR1,13 Studies of Religion 1,997443316")
			sw.WriteLine("13YM1,13 Year Meeting 1,997443316")
			sw.WriteLine("13BIO2,13 Biology 2,997482819")
			sw.WriteLine("13CHP1,13 Chapel 1,997482819")
			sw.WriteLine("13ENS1,13 English Standard 1,997482819")
			sw.WriteLine("13MAG2,13 General Mathematics 2,997482819")
			sw.WriteLine("13ITM1,13 ITMM 1,997482819")
			sw.WriteLine("13MG1,13 Mentor Group 1,997482819")
			sw.WriteLine("13MGMA,13 Mentor Group MA,997482819")
			sw.WriteLine("13MU11,13 Music 1 1,997482819")
			sw.WriteLine("13YM1,13 Year Meeting 1,997482819")
			sw.WriteLine("13CHP1,13 Chapel 1,997488118")
			sw.WriteLine("13DAT1,13 Design and Technology 1,997488118")
			sw.WriteLine("13ENS2,13 English Standard 2,997488118")
			sw.WriteLine("13MAG2,13 General Mathematics 2,997488118")
			sw.WriteLine("13MG2,13 Mentor Group 2,997488118")
			sw.WriteLine("13MGWA,13 Mentor Group WA,997488118")
			sw.WriteLine("13OHS1,13 Open High School 1,997488118")
			sw.WriteLine("13PDH1,13 PDHPE 1,997488118")
			sw.WriteLine("13YM2,13 Year Meeting 2,997488118")
			sw.WriteLine("13BUS2,13 Business Studies 2,997489795")
			sw.WriteLine("13CHP1,13 Chapel 1,997489795")
			sw.WriteLine("13ENS2,13 English Standard 2,997489795")
			sw.WriteLine("13EXT7,13 External Studies 7,997489795")
			sw.WriteLine("13FTE1,13 Food Technology 1,997489795")
			sw.WriteLine("13MG2,13 Mentor Group 2,997489795")
			sw.WriteLine("13MGBE,13 Mentor Group BE,997489795")
			sw.WriteLine("13PDH1,13 PDHPE 1,997489795")
			sw.WriteLine("13YM2,13 Year Meeting 2,997489795")
			sw.WriteLine("13CHP1,13 Chapel 1,997491312")
			sw.WriteLine("13DAT1,13 Design and Technology 1,997491312")
			sw.WriteLine("13ENS2,13 English Standard 2,997491312")
			sw.WriteLine("13EXT1,13 External Studies 1,997491312")
			sw.WriteLine("13FTE1,13 Food Technology 1,997491312")
			sw.WriteLine("13MG2,13 Mentor Group 2,997491312")
			sw.WriteLine("13MGWA,13 Mentor Group WA,997491312")
			sw.WriteLine("13PDH1,13 PDHPE 1,997491312")
			sw.WriteLine("13YM2,13 Year Meeting 2,997491312")
			sw.WriteLine("13BIO2,13 Biology 2,997491554")
			sw.WriteLine("13CHP1,13 Chapel 1,997491554")
			sw.WriteLine("13TX1,13 DT TX 1,997491554")
			sw.WriteLine("13ENA3,13 English Advanced 3,997491554")
			sw.WriteLine("13MAG2,13 General Mathematics 2,997491554")
			sw.WriteLine("13GEO1,13 Geography 1,997491554")
			sw.WriteLine("13MG2,13 Mentor Group 2,997491554")
			sw.WriteLine("13MGWA,13 Mentor Group WA,997491554")
			sw.WriteLine("13YM2,13 Year Meeting 2,997491554")
			sw.WriteLine("13CHP1,13 Chapel 1,997538047")
			sw.WriteLine("13ECO1,13 Economics 1,997538047")
			sw.WriteLine("13ENA2,13 English Advanced 2,997538047")
			sw.WriteLine("13FRB1,13 French Beginners 1,997538047")
			sw.WriteLine("13MAG1,13 General Mathematics 1,997538047")
			sw.WriteLine("13MG2,13 Mentor Group 2,997538047")
			sw.WriteLine("13MGMA,13 Mentor Group MA,997538047")
			sw.WriteLine("13MOD1,13 Modern History 1,997538047")
			sw.WriteLine("13YM2,13 Year Meeting 2,997538047")
			sw.WriteLine("13CHP1,13 Chapel 1,9983")
			sw.WriteLine("13TX1,13 DT TX 1,9983")
			sw.WriteLine("13ENA1,13 English Advanced 1,9983")
			sw.WriteLine("13ENX1,13 English Extension 1 1,9983")
			sw.WriteLine("13EXX1,13 English Extension 2 1,9983")
			sw.WriteLine("13EXX1rd,13 English Extension 2 1rd,9983")
			sw.WriteLine("13MAT2,13 Mathematics 2,9983")
			sw.WriteLine("13MG2,13 Mentor Group 2,9983")
			sw.WriteLine("13MGBR,13 Mentor Group BR,9983")
			sw.WriteLine("13ART1,13 Visual Arts 1,9983")
			sw.WriteLine("13YM2,13 Year Meeting 2,9983")









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
event.start_date >  '01/01/2018' 
AND event.end_date < '12/31/2019' 
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






End Module
