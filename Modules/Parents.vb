Imports System.DirectoryServices

Module Parents

    Function getEdumateParents(config As configSettings, edumateStudents As List(Of user))


        Dim ConnectionString As String = config.edumateConnectionString
        Dim commandString As String =
"
SELECT        
parentcontact.firstname,
parentcontact.surname,
edumate.carer.carer_id,
edumate.student.student_id,
edumate.carer.carer_number,
edumate.relationship_type.relationship_type

FROM            edumate.relationship

INNER JOIN edumate.contact as ParentContact
ON edumate.relationship.contact_id2 = Parentcontact.contact_id

INNER JOIN edumate.contact as StudentContact 
ON edumate.relationship.contact_id1 = studentContact.contact_id

INNER JOIN edumate.student
ON studentContact.contact_id = edumate.student.contact_id

INNER JOIN edumate.carer 
ON parentcontact.contact_id = edumate.carer.contact_id

INNER JOIN edumate.relationship_type
ON edumate.relationship.relationship_type_id = edumate.relationship_type.relationship_type_id


WHERE        (edumate.relationship.relationship_type_id IN (1, 4, 15, 28, 33, 10)) 
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
                        users.Last.edumateProperties.carer_number = dr.GetValue(4)
                        users.Last.children.Add(getStudentFromID(dr.GetValue(3), edumateStudents))
                        users.Last.relationshipType = dr.GetValue(5)
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
edumate.carer.carer_id,
edumate.student.student_id,
edumate.carer.carer_number



FROM            edumate.relationship

INNER JOIN edumate.contact as ParentContact
ON edumate.relationship.contact_id1 = Parentcontact.contact_id

INNER JOIN edumate.contact as StudentContact 
ON edumate.relationship.contact_id2 = studentContact.contact_id

INNER JOIN edumate.student
ON studentContact.contact_id = edumate.student.contact_id

INNER JOIN edumate.carer 
ON parentcontact.contact_id = edumate.carer.contact_id




WHERE        (edumate.relationship.relationship_type_id IN (2, 5, 9, 16, 29, 34, 11)) 


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
                        users.Last.edumateProperties.carer_number = dr.GetValue(4)
                        users.Last.children.Add(getStudentFromID(dr.GetValue(3), edumateStudents))
                    End If
                End If
            End While
            conn.Close()
        End Using

        Return users
    End Function



    Sub DisableFomerParents(direntry As DirectoryEntry, currentEdumateParents As List(Of user))

        Using searcher As New DirectorySearcher(direntry)

            Dim ParentsToDisable As New List(Of SearchResult)

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
                Dim active As Boolean = False
                For Each user In currentEdumateParents
                    If result.Properties("employeeID").Count > 0 Then
                        If result.Properties("employeeID")(0) = user.employeeID Then
                            active = True
                        End If
                    End If
                Next
                If active = False Then
                    ParentsToDisable.Add(result)
                    Console.WriteLine(result.Properties("cn")(0))
                End If
            Next


            MsgBox("break")

            For Each parentToDisable In ParentsToDisable

                Using ADuser As New DirectoryEntry("LDAP://" & parentToDisable.Properties("distinguishedName")(0))
                    'Setting username & password to Nothing forces
                    'the connection to use your logon credentials
                    ADuser.Username = Nothing
                    ADuser.Password = Nothing
                    'Always use a secure connection
                    ADuser.AuthenticationType = AuthenticationTypes.Secure
                    ADuser.RefreshCache()


                    ADuser.Properties("userAccountControl").Value = "66082"


                    ADuser.CommitChanges()


                    ADuser.MoveTo(New DirectoryEntry(("LDAP://" & "OU=@ofgsfamily.com-disabled,OU=Staff Users,OU=All,DC=i,DC=ofgs,DC=nsw,DC=edu,DC=au")))

                End Using






            Next

        End Using

    End Sub






End Module
