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



            Next
            ADgroup.CommitChanges()

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

		Using searcher As New DirectorySearcher(direntry)

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



	End Sub

End Module
