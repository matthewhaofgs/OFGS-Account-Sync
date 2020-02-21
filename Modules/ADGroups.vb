Imports System.DirectoryServices


Public Class department
	Public name As String
	Public members As New List(Of user)
End Class


Module ADGroups

    Sub AddStaffToGroups(users As List(Of user), config As configSettings)

		Dim musicTutors As New List(Of user)
		Dim currentStaff As New List(Of user)

		For Each user In users

            Dim musicTutor As Integer = 0
            If Not IsNothing(user.edumateGroupMemberships) Then
                For Each group In user.edumateGroupMemberships


                    If group = config.tutorGroupID Then
                        musicTutor = 1
                    End If
                    If group = config.danceTutorGroupID Then
                        musicTutor = 1
                    End If

                Next
            End If

            If musicTutor = 1 Then
                musicTutors.Add(user)

            End If

			If user.edumateCurrent = 1 And musicTutor = 0 Then
				currentStaff.Add(user)
			End If


		Next

		addUsersToGroup(musicTutors, config.sg_tutors)
		addUsersToGroup(currentStaff, config.sg_currentStaff)
		addUsersToGroup(currentStaff, "CN=SG_Adobe_Staff,OU=_Security Groups,OU=All,DC=i,DC=ofgs,DC=nsw,DC=edu,DC=au")


	End Sub

    Sub addUsersToGroup(users As List(Of user), group As String)



        Using ADgroup As New DirectoryEntry("LDAP://" & group)
            'Setting username & password to Nothing forces
            'the connection to use your logon credentials
            ADgroup.Username = Nothing
            ADgroup.Password = Nothing
            'Always use a secure connection
            ADgroup.AuthenticationType = AuthenticationTypes.Secure
            ADgroup.RefreshCache()


            ADgroup.Properties("member").Clear()
            ADgroup.CommitChanges()
            For Each user In users

                ADgroup.Properties("member").Add(user.distinguishedName)
				ADgroup.Properties("mail").Add(ADgroup.Properties("cn").Value & "@ofgs.nsw.edu.au")

			Next
            ADgroup.CommitChanges()
			'MsgBox(ADgroup.Properties("cn").Value & "@ofgs.nsw.edu.au")



		End Using






    End Sub

	Sub addUserToDepartmentGroups(users As List(Of user), dirEntry As DirectoryEntry)

		Dim departments As New List(Of department)
		Dim existing As Boolean
		Dim existingGroupNames As List(Of String)


		For Each user In users
			If Not IsNothing(user.edumateDepartmentMemberships) Then
				For Each edumateDepartment In user.edumateDepartmentMemberships
					existing = False
					For Each objDepartment In departments
						If edumateDepartment = objDepartment.name Then
							existing = True
							objDepartment.members.Add(user)
						End If
					Next
					If existing = False Then
						Dim objDepartment = New department
						objDepartment.name = edumateDepartment
						objDepartment.members.Add(user)
						departments.Add(objDepartment)
					End If
				Next
			End If

		Next

		existingGroupNames = getADGroups(dirEntry)

		For Each objDepartment In departments
			existing = False
			For Each existingGroupName In existingGroupNames
				If "Edumate_Department_" & objDepartment.name = existingGroupName Then
					existing = True

				End If
			Next
			If existing = False Then
				'MsgBox("Break")
				createADGroup("Edumate_Department_" & objDepartment.name)

			End If
		Next

		For Each objDepartment In departments
			addUsersToGroup(objDepartment.members, ("CN=Edumate_Department_" & objDepartment.name & ",OU=_Edumate Groups,OU=All,DC=i,DC=ofgs,DC=nsw,DC=edu,DC=au"))
		Next

	End Sub

	Function getADGroups(direntry As DirectoryEntry)

		Dim groupNames As New List(Of String)

		Using searcher As New DirectorySearcher("LDAP://" & "OU=_Edumate Groups,OU=All,DC=i,DC=ofgs,DC=nsw,DC=edu,DC=au")

			searcher.PropertiesToLoad.Add("cn")
			searcher.Filter = "(objectCategory=Group)"
			searcher.ServerTimeLimit = New TimeSpan(0, 0, 60)
			searcher.SizeLimit = 100000000
			searcher.Asynchronous = False
			searcher.ServerPageTimeLimit = New TimeSpan(0, 0, 60)
			searcher.PageSize = 10000


			Dim queryResults As SearchResultCollection
			queryResults = searcher.FindAll

			Dim result As SearchResult

			For Each result In queryResults
				groupNames.Add(result.Properties("cn")(0))
			Next
		End Using

		Return groupNames

	End Function

	Sub createADGroup(groupName As String)

		Try
			Using dirEntry As New DirectoryEntry("LDAP://" & "OU=_Edumate Groups,OU=All,DC=i,DC=ofgs,DC=nsw,DC=edu,DC=au")
				Dim objGroup
				'Setting username & password to Nothing forces
				'the connection to use your logon credentials
				dirEntry.Username = Nothing
				dirEntry.Password = Nothing
				'Always use a secure connection
				dirEntry.AuthenticationType = AuthenticationTypes.Secure
				dirEntry.RefreshCache()
				objGroup = dirEntry.Children.Add("CN=" & groupName, "Group")
				objGroup.CommitChanges()

			End Using

		Catch
		End Try

	End Sub

	Sub addUserToRoleGroups(users As List(Of user), dirEntry As DirectoryEntry)

		Dim existingGroupNames As List(Of String)
		Dim jobRoles As New List(Of department)
		Dim existing As Boolean



		For Each user In users

			For Each workTitle In user.workTitles
				If workTitle <> "" Then
					existing = False
					For Each objJobRole In jobRoles
						If (workTitle).ToLower = (objJobRole.name).ToLower Then
							existing = True
							objJobRole.members.Add(user)
						End If
					Next

					If existing = False Then
						Dim objJobRole = New department
						objJobRole.name = workTitle
						objJobRole.members.Add(user)
						jobRoles.Add(objJobRole)
					End If
				End If
			Next
		Next



			existingGroupNames = getADGroups(dirEntry)
		'MsgBox("Break")
		For Each objJobRole In jobRoles
			existing = False
			For Each existingGroupName In existingGroupNames
				If objJobRole.name = existingGroupName Then
					existing = True

				End If
			Next
			If existing = False Then
				'MsgBox("Break")
				createADGroup(objJobRole.name)

			End If
		Next


		Dim emptyUserList As New List(Of user)

		'For Each group In existingGroupNames
		'MsgBox(group)
		'Next


		For Each objJobRole In jobRoles
			addUsersToGroup(objJobRole.members, ("CN=" & objJobRole.name & ",OU=_Edumate Groups,OU=All,DC=i,DC=ofgs,DC=nsw,DC=edu,DC=au"))
		Next





	End Sub


	Function getEdumateManagedGroups()

	End Function


	Sub addUsersToYearGroups(users As List(Of user), dirEntry As DirectoryEntry)
		Dim departments As New List(Of department)
		Dim existing As Boolean
		Dim existingGroupNames As List(Of String)


		Dim yearGroup As New department
		yearGroup.name = "Year_7_Teachers"
		departments.Add(yearGroup)
		yearGroup = Nothing

		yearGroup = New department
		yearGroup.name = "Year_8_Teachers"
		departments.Add(yearGroup)
		yearGroup = Nothing

		yearGroup = New department
		yearGroup.name = "Year_9_Teachers"
		departments.Add(yearGroup)
		yearGroup = Nothing

		yearGroup = New department
		yearGroup.name = "Year_10_Teachers"
		departments.Add(yearGroup)
		yearGroup = Nothing

		yearGroup = New department
		yearGroup.name = "Year_11_Teachers"
		departments.Add(yearGroup)
		yearGroup = Nothing

		yearGroup = New department
		yearGroup.name = "Year_12_Teachers"
		departments.Add(yearGroup)
		yearGroup = Nothing


		For Each user In users
			If Not IsNothing(user.edumateProperties.yearsTeaching) Then





				If user.edumateProperties.yearsTeaching.Contains("07") Then
					departments(0).members.Add(user)
				End If
				If user.edumateProperties.yearsTeaching.Contains("08") Then
					departments(1).members.Add(user)
				End If
				If user.edumateProperties.yearsTeaching.Contains("09") Then
					departments(2).members.Add(user)
				End If
				If user.edumateProperties.yearsTeaching.Contains("10") Then
					departments(3).members.Add(user)
				End If
				If user.edumateProperties.yearsTeaching.Contains("11") Then
					departments(4).members.Add(user)
				End If
				If user.edumateProperties.yearsTeaching.Contains("12") Then
					departments(5).members.Add(user)
				End If

			End If
		Next


		existingGroupNames = getADGroups(dirEntry)

		For Each objDepartment In departments
			existing = False
			For Each existingGroupName In existingGroupNames
						If "Edumate_" & objDepartment.name = existingGroupName Then
							existing = True

						End If
					Next
			If existing = False Then

				createADGroup("Edumate_" & objDepartment.name)

			End If
		Next

		For Each objDepartment In departments
			addUsersToGroup(objDepartment.members, ("CN=Edumate_" & objDepartment.name & ",OU=_Edumate Groups,OU=All,DC=i,DC=ofgs,DC=nsw,DC=edu,DC=au"))

		Next
	End Sub


End Module
