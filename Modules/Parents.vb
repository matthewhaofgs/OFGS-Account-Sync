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


End Module
