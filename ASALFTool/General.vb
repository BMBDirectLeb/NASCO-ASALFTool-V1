Imports LFSO102Lib
Imports LFIMAGEENABLE80Lib
Imports System.IO
Module General

    Public ie As New ImageEnable

    Public Client
    Public LFapp As New LFApplication
    Public LFServer As New LFServer
    Public DB As New Object
    Public Con As New Object
    Public ClientConnected As Boolean = False
    Public LFVolume As LFVolume
    Public LFDocID As Long

    Public LFServerName As String
    Public LFDBName As String
    Public LFUserName As String
    Public LFUserPassw As String


    Public LFDestinationFolder As String
    Public LFDocumentName As String
    Public LFVolumeName As String
    Public LFTemplateName As String
    Public LFSearch As String
    Public NewType As String

    ''''Modification applied on 20180619
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Public FieldsList As New List(Of String)

    ''' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''




    Public CountriesList As New ArrayList
    Public AccountMgrsList As New ArrayList
    Public LFAccountManager As String


    Public Sub Log(ByVal text As String)
        Dim sr As New StreamWriter(Application.StartupPath & "\Logfile.log", True)
        sr.WriteLine(text)
        sr.Close()
    End Sub

    Public Function ReadParam(ByVal DocumentType As String)
        Dim a As Array
        Dim i As Integer
        Dim line As String
        Try

            Dim sr As New StreamReader(Application.StartupPath & "\Param.ini")

            a = Nothing
            a = Split(sr.ReadLine, "=", -1, CompareMethod.Text)
            LFServerName = Trim(a(1).trim)

            a = Nothing
            a = Split(sr.ReadLine, "=", -1, CompareMethod.Text)
            LFDBName = Trim(a(1).trim)

            a = Nothing
            a = Split(sr.ReadLine, "=", -1, CompareMethod.Text)
            LFUserName = Trim(a(1).trim)

            a = Nothing
            a = Split(sr.ReadLine, "=", -1, CompareMethod.Text)
            LFUserPassw = DecryptPassword(Trim(a(1).trim))

            ''''Before modification on 20180619
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'a = Nothing
            'a = Split(sr.ReadLine, "=", -1, CompareMethod.Text)
            'LFDestinationFolder = Trim(a(1).trim)

            'a = Nothing
            'a = Split(sr.ReadLine, "=", -1, CompareMethod.Text)
            'LFDocumentName = Trim(a(1).trim)

            'a = Nothing
            'a = Split(sr.ReadLine, "=", -1, CompareMethod.Text)
            'LFVolumeName = Trim(a(1).trim)

            'a = Nothing
            'a = Split(sr.ReadLine, "=", -1, CompareMethod.Text)
            'LFPTemplateName = Trim(a(1).trim)

            'a = Nothing
            'a = Split(sr.ReadLine, "=", -1, CompareMethod.Text)
            'LFOTemplateName = Trim(a(1).trim)

            'sr.Close()
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ''''Modification applied on 20180619
            'Split Reading Document based on Document Type
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            sr.Close()
        Catch ex As Exception
            MsgBox("Sorry, cannot read Param.ini. Please check this file")
            End
        End Try
        Dim sr_DocType As StreamReader
        Try
            Try
                sr_DocType = New StreamReader(Application.StartupPath & "\DocumentTypes\" & DocumentType & ".ini")
            Catch ex As Exception
                Try
                    sr_DocType = New StreamReader(Application.StartupPath & "\DocumentTypes\Default.ini")
                Catch ex2 As Exception
                    MsgBox("Sorry, cannot read DocumentTypes\" & DocumentType & ".ini. Please check this file")
                    End
                End Try

            End Try
            a = Nothing
            a = Split(sr_DocType.ReadLine, "=", -1, CompareMethod.Text)
            LFDestinationFolder = Trim(a(1).trim)
            'MsgBox(LFDestinationFolder)
            a = Nothing
            a = Split(sr_DocType.ReadLine, "=", -1, CompareMethod.Text)
            LFDocumentName = Trim(a(1).trim)
            'MsgBox(LFDocumentName)
            a = Nothing
            a = Split(sr_DocType.ReadLine, "=", -1, CompareMethod.Text)
            LFVolumeName = Trim(a(1).trim)
            'MsgBox(LFVolumeName)
            a = Nothing
            a = Split(sr_DocType.ReadLine, "=", -1, CompareMethod.Text)
            LFTemplateName = Trim(a(1).trim)
            'MsgBox(LFTemplateName)


            a = Nothing
            a = Split(sr_DocType.ReadLine, ";", -1, CompareMethod.Text)
            LFSearch = Trim(a(1).trim)
            ' MsgBox(LFSearch)


            sr_DocType.ReadLine()    ' ''''''''''''''''''''''
            sr_DocType.ReadLine()   ' '''Fields Settings''''
            sr_DocType.ReadLine()   ' ''''''''''''''''''''''

            Do While sr_DocType.Peek() >= 0
                FieldsList.Add(sr_DocType.ReadLine)

            Loop

            sr_DocType.Close()
        Catch ex As Exception
            MsgBox("Sorry, error while reading " & DocumentType & ".ini. Please check this file")
            Log("Sorry, error while reading " & DocumentType & ".ini. Please check this file")
            End
        End Try
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


    End Function
    Public Function DecryptPassword(ByVal EP As String) As String
        Dim i As Integer
        Dim PWD As String
        PWD = ""
        For i = 1 To Len(EP)
            PWD = PWD & Chr((Asc(Mid(EP, i, 1)) - i) - 12)
        Next
        DecryptPassword = PWD
    End Function
    Public Sub LoadLaserfiche()

        Dim rc As Integer
        Try
Stablish:
            rc = rc + 1
            If rc > 5 Then
                MsgBox("Cannot connect to repository, please check if there is an available connection")
                Log(Now & " Cannot connect to repository, please check if there is and available connection")
                Try
                    Con.Terminate()
                Catch ex As Exception
                End Try

                Try
                    Dim ShellCmd As String = Application.StartupPath & "\NirCmd.exe win close title " & """" & LFDBName & " - Laserfiche" & """"
                    'Track(ShellCmd)
                    Shell(ShellCmd)
                Catch ex As Exception
                    Log(Now & " NirCmd Error closing window for Repository")
                End Try

                End
            End If
            ie.LoadLaserFiche(False)

            ie.Login(LFDBName, LFServerName, LFUserName, LFUserPassw, 2)

            Dim counter As Integer
            'Wait until login is complete
            'While ie.GetUserName = ""
            '    Log(rc & "-" & counter)
            '    counter = counter + 1
            '    'MsgBox("a")
            'End While
            'MsgBox(rc)
            'ClientConnect()                        'It connected using ImageEnable, but needs to get the LFSO connection
            Exit Sub
        Catch ex As Exception
            Dim pa As String = Err.Description
            Log(Now & " try number = " & rc & " error = " & pa)
            GoTo Stablish
        End Try


    End Sub
    Public Sub ClientConnect()

        Try
            Client = GetObject(Nothing, "LFClient.Document")
        Catch ex As Exception
            'MsgBox("Sorry Dear, the Laserfiche client is not open.")
            Exit Sub
        End Try

        'Get its connection
        Try
            DB = Client.getdatabase()
            Con = DB.CurrentConnection

            'to solve the RTTI problem
            Dim contransfer As New LFConnection
            Dim SerializedCon = DB.CurrentConnection.SerializedConnection
            contransfer.CloneFromSerializedConnection(SerializedCon)
            DB = contransfer.Database

        Catch ex As Exception
            Log(Now & " Not logged into a database, try to open new connection.")
            Exit Sub
        End Try

        ClientConnected = True

    End Sub
    Public Sub connect()
        Try
            LFServer = LFapp.GetServerByName(LFServerName)
            DB = LFServer.GetDatabaseByName(LFDBName)
            Con.UserName = LFUserName
            Con.Password = LFUserPassw
            Con.Shared = True
            Con.Create(DB)
        Catch ex As Exception
            MsgBox("Cannot connect to Laserfiche Repository." & vbCrLf & Err.Description & vbCrLf & "The application now will close")
            Log(Now & " Could not connect to Laserfiche Repository " & Err.Description)

            Try
                Dim ShellCmd As String = Application.StartupPath & "\NirCmd.exe win close title " & """" & LFDBName & " - Laserfiche" & """"
                'Track(ShellCmd)
                Shell(ShellCmd)
            Catch ex1 As Exception
                Log(Now & " NirCmd Error closing window for Repository")
            End Try

            End
        End Try
    End Sub
    Public Sub Disconnect()
        Try
            'Con.Terminate()
        Catch ex As Exception

        End Try
        Try
            'ie.Logout()
        Catch ex As Exception

        End Try
    End Sub
    Public Sub CreateFolderByPath(ByVal folderpath As String)

        Dim a As Array
        Dim f As New LFFolder
        Dim p As New LFFolder
        Dim name As String
        a = Split(folderpath, "\", "-1", CompareMethod.Text)
        Dim i As Integer
        Dim cumm As String
        cumm = "\"
        For i = 0 To a.Length - 1
            Try
                f = New LFFolder
                name = a(i)
                p = DB.GetEntryByPath(cumm)
                f.Create(name, p, False)
                f.Dispose()
                p.Dispose()
            Catch ex As System.Runtime.InteropServices.COMException
                f.Dispose()
                p.Dispose()
            End Try
            If i + 1 < a.Length Then
                name = a(i + 1)
                cumm = cumm & a(i) & "\"
            End If
        Next
    End Sub

    Public Sub FieldListNamesUpdate(ByVal TemplateName As String, ByVal FieldName As String, ByVal NewName As String)

        Dim LFTargetTemplate As New LFTemplate
        LFTargetTemplate = DB.GetTemplateByName(TemplateName)
        Dim LFField As LFTemplateField
        LFField = LFTargetTemplate.ItemByName(FieldName)

        'LFTemplate.Create(DB, "RoulitaLinda", True)
        'Dim LFField As New LFTemplateField
        'LFField.Create(DB, "Danielita", Field_Type.FIELD_TYPE_LIST, True)
        'LFField.Append("Maiita")
        'LFField.Update()
        'LFTemplate.AddTemplateField(1, LFField)
        'LFTemplate.Update()
        'LFField.Item(2) = "Ritita"
        'LFField.Update()

        If LFField.Size < NewName.Length Then
            'Enlarge the field lenght to host the new name
            LFField.Size = NewName.Length + 1
            LFField.Update()
        End If

        Dim FieldListCount As Integer = LFField.Count
        Dim FieldLastName As String = LFField.Item(FieldListCount)
        If NewName > FieldLastName Then
            'Just put the new name at the end of the list of clients
            LFField.Append(NewName)
            LFField.Update()
        Else
            'We need to recreate the list because Toolkit has no sort method for the list values
            Dim LFFieldListCollection As New ArrayList
            Try
                For jh As Integer = 1 To FieldListCount
                    'Create a ArrayList of current Client Names
                    Dim FieldListName As String = LFField.Item(jh)
                    If FieldListName <> NewName Then
                        LFFieldListCollection.Add(FieldListName)
                    Else
                        'The Client already exists, so dont continue
                        MsgBox("This Name already exists in the list of clients")
                        Exit Sub
                    End If
                Next
                'Add the new name to the list of clients
                LFFieldListCollection.Add(NewName)

                'Sort the list
                LFFieldListCollection.Sort()

                'Remove the list from the Template field
                LFField.ClearDropDownList()
                LFField.Update()

                'Recreate the list with the sorted names
                For Each LFClientName As String In LFFieldListCollection
                    LFField.Append(LFClientName)
                    LFField.Update()
                Next
            Catch ex As Exception
                'In case of error, write the list of Clients to correct it manually
                Dim sr As New StreamWriter(Application.StartupPath & "\LastClientList.txt", True)
                For Each LFClientName As String In LFFieldListCollection
                    sr.WriteLine(LFClientName)
                Next
                sr.Close()
            End Try
        End If

    End Sub


End Module
