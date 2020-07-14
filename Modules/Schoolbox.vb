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
edumate.staff.staff_id,
case when edumate.staff.staff_number in (

select distinct
schoolbox_staff1.staff_number

from
(
select
edumate.staff.staff_number,
edumate.salutation.salutation,
coalesce(replace(edumate.contact.preferred_name,'&#0'||'39;',''''), replace(edumate.contact.firstname,'&#0'||'39;','''')) as firstname,
replace(edumate.contact.surname,'&#039;','''') as surname,
edumate.sys_user.username as username1,
edumate.contact.email_address,
edumate.house.house,
edumate.campus.campus,
replace(edumate.work_detail.title,'&#039;','''') as title
from edumate.staff
inner join edumate.contact on edumate.contact.contact_id = edumate.staff.contact_id
left join edumate.staff_employment on edumate.staff_employment.staff_id = edumate.staff.staff_id
left join edumate.work_detail on edumate.work_detail.contact_id=edumate.contact.contact_id
left join edumate.salutation on edumate.salutation.salutation_id = edumate.contact.salutation_id
left join edumate.sys_user on edumate.sys_user.contact_id = edumate.contact.contact_id
left join edumate.house on edumate.house.house_id = edumate.staff.house_id
left join edumate.campus on edumate.campus.campus_id = edumate.staff.campus_id
where (edumate.staff_employment.end_date is null or edumate.staff_employment.end_date >= current date)
and edumate.staff_employment.start_date <= (current date +90 DAYS)
and (edumate.contact.pronounced_name is null or edumate.contact.pronounced_name != 'NOT STAFF')

) schoolbox_staff1

inner join edumate.staff on schoolbox_staff1.staff_number = edumate.staff.staff_number

inner join edumate.contact on edumate.staff.contact_id = edumate.contact.contact_id

left join edumate.teacher on edumate.contact.contact_id = edumate.teacher.contact_id

left join edumate.class_teacher on edumate.class_teacher.teacher_id = edumate.teacher.teacher_id

left join edumate.class on edumate.class.class_id = edumate.class_teacher.class_id


left join 
(
	select max_student_class.class_id, edumate.form.short_name

	from 
	(
		select max(student_id) as randomStudentNumber, class_id

		from edumate.class_enrollment

		where 

		(SELECT current date FROM sysibm.sysdummy1) between edumate.class_enrollment.start_date and edumate.class_enrollment.end_date

		group by class_id
	) max_student_class


	INNER JOIN 
	(
		select student_id, max(form_run_id) as max_form_run_id

		from edumate.student_form_run 

		where  
		(SELECT current date FROM sysibm.sysdummy1) between edumate.student_form_run.start_date and edumate.student_form_run.end_date
	
		group by student_id
	) max_form_run
	ON max_form_run.student_id = max_student_class.randomStudentNumber

	INNER JOIN edumate.form_run on max_form_run.max_form_run_id = edumate.form_run.form_run_id

	INNER JOIN edumate.form on edumate.form_run.form_id = edumate.form.form_id

) class_short_names

on edumate.class.class_id = class_short_names.class_id


where class_short_names.short_name = 'K'
and edumate.class.class_type_id = '2'

) then 'true' else 'false' END AS kindy,
schoolbox_staff2.title as title

from (

select
edumate.staff.staff_number,
edumate.salutation.salutation,
coalesce(replace(edumate.contact.preferred_name,'&#0'||'39;',''''), replace(edumate.contact.firstname,'&#0'||'39;','''')) as firstname,
replace(edumate.contact.surname,'&#039;','''') as surname,
edumate.sys_user.username as username1,
edumate.contact.email_address,
edumate.house.house,
edumate.campus.campus,
replace(edumate.work_detail.title,'&#039;','''') as title
from edumate.staff
inner join edumate.contact on edumate.contact.contact_id = edumate.staff.contact_id
left join edumate.staff_employment on edumate.staff_employment.staff_id = edumate.staff.staff_id
left join edumate.work_detail on edumate.work_detail.contact_id=edumate.contact.contact_id
left join edumate.salutation on edumate.salutation.salutation_id = edumate.contact.salutation_id
left join edumate.sys_user on edumate.sys_user.contact_id = edumate.contact.contact_id
left join edumate.house on edumate.house.house_id = edumate.staff.house_id
left join edumate.campus on edumate.campus.campus_id = edumate.staff.campus_id
where (edumate.staff_employment.end_date is null or edumate.staff_employment.end_date >= current date)
and edumate.staff_employment.start_date <= (current date + 90 DAYS)
and (edumate.contact.pronounced_name is null or edumate.contact.pronounced_name != 'NOT STAFF')

)  schoolbox_staff2

inner join edumate.staff on schoolbox_staff2.staff_number = edumate.staff.staff_number
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
CASE WHEN substr(edumate.timetable.timetable, 6, 6) = 'Year 1' THEN 'Senior' ELSE substr(edumate.timetable.timetable, 6, 6) END AS Expr1, 
REPLACE(CONCAT(CONCAT(edumate.term.term, ' '), 
substr(edumate.timetable.timetable, 1, 4)), 'Term 0', 'Term 4') AS Expr2, 
edumate.term.start_date, 
edumate.term.end_date, 
edumate.term.cycle_start_day, 
edumate.cycle_day.day_index, 
edumate.period.period, 
edumate.period.start_time, 
edumate.period.end_time

FROM edumate.TERM_GROUP, edumate.cycle_day, edumate.period_cycle_day, edumate.period, edumate.term, edumate.timetable

WHERE (start_date > '01/01/2020') 
AND (end_date < '01/01/2021') 
AND (edumate.term_group.cycle_id = edumate.cycle_day.cycle_id) 
AND (edumate.cycle_day.cycle_day_id = edumate.period_cycle_day.cycle_day_id) 
AND (edumate.period_cycle_day.period_id = edumate.period.period_id) 
AND (edumate.term_group.term_id = edumate.term.term_id) 
AND (edumate.term.timetable_id = edumate.timetable.timetable_id)

"





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

                'strTermTitle = Replace(dr.GetValue(1), "2021", "2010")
                strTermTitle = (dr.GetValue(1))

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
 substr(edumate.timetable.timetable, 6, 6) as CAMPUS1,	
CONCAT(CONCAT(edumate.term.term, ' '), substr(edumate.timetable.timetable, 1, 4)) AS Expr2,
edumate.cycle_day.day_index as DAY_NUMBER,
	edumate.period.period as PERIOD_NUMBER,
	concat(edumate.course.code,edumate.class.identifier) AS CLASS_CODE,
	edumate.class.class,
	edumate.room.code AS ROOM,
edumate.staff.staff_number


FROM edumate.period_class
INNER JOIN edumate.period_cycle_day ON edumate.period_cycle_day.period_cycle_day_id = edumate.period_class.period_cycle_day_id
INNER JOIN edumate.cycle_day ON edumate.cycle_day.cycle_day_id = edumate.period_cycle_day.cycle_day_id
INNER JOIN edumate.period ON edumate.period.period_id = edumate.period_cycle_day.period_id
INNER JOIN edumate.class ON edumate.class.class_id = edumate.period_class.class_id
INNER JOIN edumate.course ON edumate.course.course_id = edumate.class.course_id
INNER JOIN edumate.perd_cls_teacher ON edumate.perd_cls_teacher.period_class_id = edumate.period_class.period_class_id 
	AND edumate.perd_cls_teacher.is_primary = 1
INNER JOIN edumate.teacher ON edumate.teacher.teacher_id = edumate.perd_cls_teacher.teacher_id
INNER JOIN edumate.staff ON edumate.staff.contact_id = edumate.teacher.contact_id
INNER JOIN edumate.room ON edumate.room.room_id = edumate.period_class.room_id
INNER JOIN edumate.timetable ON edumate.timetable.timetable_id = edumate.period_class.timetable_id
INNER JOIN edumate.contact ON edumate.staff.contact_id = edumate.contact.contact_id
INNER JOIN edumate.term_group on edumate.cycle_day.cycle_id = edumate.term_group.cycle_id
INNER JOIN edumate.term ON edumate.term_group.term_id = edumate.term.term_id
WHERE
(
	current date BETWEEN (
	CASE
		WHEN edumate.period_class.effective_start = edumate.timetable.computed_start_date
		THEN edumate.timetable.computed_v_start_date
		ELSE edumate.timetable.computed_v_start_date
	END
	)
	AND edumate.period_class.effective_end
)
AND
(
	current date BETWEEN (
	CASE
		WHEN edumate.period_class.effective_start = edumate.timetable.computed_start_date
		THEN edumate.timetable.computed_v_start_date
		ELSE edumate.timetable.computed_v_start_date
	END
	)
	AND edumate.timetable.computed_end_date
)
AND
(
edumate.period_class.effective_start BETWEEN edumate.term.start_date AND edumate.term.end_date
OR
edumate.period_class.effective_end BETWEEN edumate.term.start_date AND edumate.term.end_date
OR
(edumate.period_class.effective_start <= edumate.term.start_date AND edumate.period_class.effective_end >= edumate.term.end_date)
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
                'strTerm = Replace(strTerm, "2020", "2019")
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

CONCAT(edumate.course.code, edumate.class.identifier) AS CLASS_CODE, 
edumate.class.class, 
edumate.student.student_number

FROM            edumate.CLASS_ENROLLMENT

INNER JOIN edumate.STUDENT 
	ON edumate.class_enrollment.student_id = edumate.student.student_id

INNER JOIN edumate.class 
	ON edumate.class_enrollment.class_id = edumate.class.class_id
	
INNER JOIN edumate.COURSE
	ON edumate.class.course_id = edumate.course.course_id
	
INNER JOIN edumate.ACADEMIC_YEAR
	ON edumate.class.academic_year_id = edumate.academic_year.academic_year_id
	

 WHERE (edumate.academic_year.academic_year = CAST((YEAR(current_date))AS VARCHAR(10)) OR edumate.academic_year.academic_year =CAST((YEAR(current_date+ 1 years)) AS varchar(10))) AND ((SELECT current date FROM sysibm.sysdummy1) between (edumate.class_enrollment.start_date -10 DAYS) AND edumate.class_enrollment.end_date)

 UNION
 
SELECT DISTINCT 

replace(CONCAT(edumate.course.code, edumate.class.identifier),'12','13') AS CLASS_CODE, 
replace(edumate.class.class,'12','13') AS CLASS,
edumate.student.student_number

FROM            edumate.CLASS_ENROLLMENT

INNER JOIN edumate.STUDENT 
	ON edumate.class_enrollment.student_id = edumate.student.student_id

INNER JOIN edumate.class 
	ON edumate.class_enrollment.class_id = edumate.class.class_id
	
INNER JOIN edumate.COURSE
	ON edumate.class.course_id = edumate.course.course_id
	
INNER JOIN edumate.ACADEMIC_YEAR
	ON edumate.class.academic_year_id = edumate.academic_year.academic_year_id
	

 WHERE (edumate.academic_year.academic_year = CAST((YEAR(current_date))AS VARCHAR(10)))  
 	--AND ((SELECT current date FROM sysibm.sysdummy1) between (class_enrollment.start_date -10 DAYS) 
	--AND class_enrollment.end_date) 
 	AND edumate.student.student_number IN 
 	(
 	SELECT  distinct      

edumate.student.student_number

FROM            
edumate.STUDENT
INNER JOIN edumate.contact ON edumate.student.contact_id = edumate.contact.contact_id
INNER JOIN edumate.view_student_start_exit_dates ON edumate.student.student_id = edumate.view_student_start_exit_dates.student_id
INNER JOIN edumate.student_form_run ON edumate.student_form_run.student_id = edumate.student.student_id
INNER JOIN edumate.form_run ON edumate.student_form_run.form_run_id = edumate.form_run.form_run_id
INNER JOIN edumate.form ON edumate.form_run.form_id = edumate.form.form_id
INNER JOIN edumate.stu_school ON edumate.student.student_id = edumate.stu_school.student_id
LEFT JOIN edumate.class_enrollment ON edumate.student.STUDENT_ID = edumate.class_enrollment.STUDENT_ID
LEFT JOIN edumate.class ON edumate.class_enrollment.class_id = edumate.class.class_id 
INNER JOIN edumate.TIMETABLE ON edumate.form_run.TIMETABLE_ID = edumate.timetable.TIMETABLE_ID



	
WHERE 

(
YEAR(edumate.view_student_start_exit_dates.exit_date) = YEAR(edumate.student_form_run.end_date)) 
AND YEAR(edumate.view_student_start_exit_dates.exit_date) = year(current_date)
AND edumate.form.SHORT_NAME = '12'
AND edumate.class.CLASS LIKE '12%'
AND edumate.class_enrollment.END_DATE < current_date
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
DATE(edumate.event.start_date) as ""Start Date"", 
varchar_format(edumate.event.start_date, 'HH24:MI')  ""Start Time"",
DATE(edumate.event.end_date) as ""Finish Date"",
varchar_format(edumate.event.end_date, 'HH24:MI')  ""Finish Time"",
0 as ""All Day"",
edumate.event.event as ""Name"",
edumate.event.description as ""Detail"",
CASE
	when edumate.event.location IS NOT NULL then (edumate.event.location)
	when (event_rooms.room_count > 1) then	('Various')
	when (event_rooms.room_count = 1) then  (edumate.room.room)
end as Location,

1 as ""Type"",
NULL as ""Publish Date"",
0 as ""Attendance"",
edumate.event.audience_type_id as ""Audience Type""

FROM
  edumate.event
left join
	( select event_id, max(room_id) as max_room_id, count(room_id) as room_count
	from edumate.event_room
	group by event_id
	)  event_rooms
on edumate.event.event_id = event_rooms.event_id

left join edumate.room on event_rooms.max_room_id = edumate.room.room_id



WHERE
edumate.event.start_date >  '01/01/2020' 
AND edumate.event.end_date < '12/31/2021' 
AND edumate.event.publish_flag = 1
AND edumate.event.recurring_id is not null and edumate.event.recurring_id > 0
AND edumate.event.event_id = (select min(event_id) from edumate.event e2
                                     where e2.recurring_id = edumate.event.recurring_id)


UNION



SELECT 
DATE(edumate.event.start_date) as ""Start Date"", 
varchar_format(edumate.event.start_date, 'HH24:MI')  ""Start Time"",
DATE(edumate.event.end_date) as ""Finish Date"",
varchar_format(edumate.event.end_date, 'HH24:MI')  ""Finish Time"",
0 as ""All Day"",
edumate.event.event as ""Name"",
edumate.event.description as ""Detail"",
CASE
	when edumate.event.location IS NOT NULL then (edumate.event.location)
	when (event_rooms.room_count > 1) then	('Various')
	when (event_rooms.room_count = 1) then  (edumate.room.room)
end as Location,

1 as ""Type"",
NULL as ""Publish Date"",
0 as ""Attendance"",
edumate.event.audience_type_id as ""Audience Type""
FROM
  edumate.event

left join
	( 
	select event_id, max(room_id) as max_room_id, count(room_id) as room_count
	from edumate.event_room
	group by event_id
	)  event_rooms
on edumate.event.event_id = event_rooms.event_id

left join edumate.room on event_rooms.max_room_id = edumate.room.room_id

WHERE
edumate.event.start_date >  '01/01/2020' 
AND edumate.event.end_date < '12/31/2021' 
AND edumate.event.publish_flag = 1
AND edumate.event.recurring_id is null

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
edumate.carer.carer_id,
edumate.student.student_id,
edumate.carer.carer_number,
edumate.student.STUDENT_NUMBER 


FROM            edumate.relationship

INNER JOIN edumate.contact as ParentContact
ON edumate.relationship.contact_id1 = Parentcontact.contact_id

INNER JOIN edumate.contact as StudentContact 
ON edumate.relationship.contact_id2 = studentContact.contact_id

INNER JOIN edumate.student
ON studentContact.contact_id = edumate.student.contact_id AND
edumate.student.student_id in (select student_id from edumate.student s1 where current_date BETWEEN (select min(start_date) from edumate.student_form_run sfr1a where s1.student_id = sfr1a.student_id) AND 
                                                                                 (select max(end_date) from edumate.student_form_run sfr2a where s1.student_id = sfr2a.student_id))

INNER JOIN edumate.carer 
ON parentcontact.contact_id = edumate.carer.contact_id


WHERE        (edumate.relationship.relationship_type_id IN (2, 5, 9, 16, 29, 34)) 


UNION 

SELECT        
parentcontact.firstname,
parentcontact.surname,
edumate.carer.carer_id,
edumate.student.student_id,
edumate.carer.carer_number,
edumate.student.STUDENT_NUMBER 


FROM            edumate.relationship

INNER JOIN edumate.contact as ParentContact
ON edumate.relationship.contact_id2 = Parentcontact.contact_id

INNER JOIN edumate.contact as StudentContact 
ON edumate.relationship.contact_id1 = studentContact.contact_id

INNER JOIN edumate.student
ON studentContact.contact_id = edumate.student.contact_id AND
edumate.student.student_id in (select student_id from edumate.student s1 where current_date BETWEEN (select min(start_date) from edumate.student_form_run sfr1a where s1.student_id = sfr1a.student_id) AND 
                                                                                 (select max(end_date) from edumate.student_form_run sfr2a where s1.student_id = sfr2a.student_id))
INNER JOIN edumate.carer 
ON parentcontact.contact_id = edumate.carer.contact_id


WHERE        (edumate.relationship.relationship_type_id IN (1, 4, 15, 28, 33)) 



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
