Imports System.DirectoryServices

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




End Module
