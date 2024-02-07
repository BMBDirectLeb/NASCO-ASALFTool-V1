Imports LFSO102Lib
Imports LFIMAGEENABLE80Lib
Public Class Form1

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Log("-----------------------------------")
        Log(Now)
        'Read Parameters
        Dim DocumentType       'Attrib01
        Dim LFSuffix           'Attrib02
        Dim DocumentID         'Attrib03
        Dim Branch             'Attrib04
        Dim InsurerCode        'Attrib05
        Dim DocumentNumber     'Attrib06
        Dim SubscriberCode     'Attrib07

        Try
            DocumentType = System.Environment.GetCommandLineArgs(1)
            LFSuffix = System.Environment.GetCommandLineArgs(2)
            DocumentID = System.Environment.GetCommandLineArgs(3)
            Branch = System.Environment.GetCommandLineArgs(4)
            InsurerCode = System.Environment.GetCommandLineArgs(5)
            DocumentNumber = System.Environment.GetCommandLineArgs(6)
            SubscriberCode = System.Environment.GetCommandLineArgs(7)

            Log("DocumentType   = " & System.Environment.GetCommandLineArgs(1) & vbCrLf &
            "LFSuffix       = " & System.Environment.GetCommandLineArgs(2) & vbCrLf &
            "DocumentID     = " & System.Environment.GetCommandLineArgs(3) & vbCrLf &
            "Branch         = " & System.Environment.GetCommandLineArgs(4) & vbCrLf &
            "InsurerCode    = " & System.Environment.GetCommandLineArgs(5) & vbCrLf &
            "DocumentNumber = " & System.Environment.GetCommandLineArgs(6) & vbCrLf &
            "SubscriberCode = " & System.Environment.GetCommandLineArgs(7))

        Catch ex As Exception

            Log(Now & " Error reading parameters" & vbCrLf &
            "DocumentType   = " & DocumentType & vbCrLf &
            "LFSuffix       = " & LFSuffix & vbCrLf &
            "DocumentID     = " & DocumentID & vbCrLf &
            "Branch         = " & Branch & vbCrLf &
            "InsurerCode    = " & InsurerCode & vbCrLf &
            "DocumentNumber = " & DocumentNumber & vbCrLf &
            "SubscriberCode = " & SubscriberCode)

            If UCase(DocumentType) = "CLAIM" Then
                'its OK, no need to log error and can continue
            Else
                MsgBox("Error reading parameters" & vbCrLf & Err.Description)
                MsgBox("DocumentType = " & DocumentType & vbCrLf &
                                "LFSuffix = " & LFSuffix & vbCrLf &
                                "DocumentID = " & DocumentID & vbCrLf &
                                "Branch = " & Branch & vbCrLf &
                                "InsurerCode = " & InsurerCode & vbCrLf &
                                "DocumentNumber = " & DocumentNumber & vbCrLf &
                                "SubscriberCode = " & SubscriberCode)
                End
            End If


        End Try


        'ReadParam
        ReadParam(DocumentType)








        'if Client is not open, create a LFSO connection to work
        'If ClientConnected = False Then
        '    Connect()
        'End If

        'Get Target Document
        'Find Destination Folder
        Dim DestFolder As New LFFolder
        Dim LFDestinationFolderPath As String
        LFDestinationFolderPath = Replace(LFDestinationFolder, "Attrib01", DocumentType)
        LFDestinationFolderPath = Replace(LFDestinationFolderPath, "Attrib02", LFSuffix)
        LFDestinationFolderPath = Replace(LFDestinationFolderPath, "Attrib03", DocumentID)
        LFDestinationFolderPath = Replace(LFDestinationFolderPath, "Attrib04", Branch)
        LFDestinationFolderPath = Replace(LFDestinationFolderPath, "Attrib05", InsurerCode)
        LFDestinationFolderPath = Replace(LFDestinationFolderPath, "Attrib06", DocumentNumber)
        LFDestinationFolderPath = Replace(LFDestinationFolderPath, "Attrib07", SubscriberCode)
        Log(Now & " DestFolder = " & LFDestinationFolderPath)

        'Select Document Name
        Dim DocName As String
        DocName = Replace(LFDocumentName, "Attrib01", DocumentType)
        DocName = Replace(DocName, "Attrib02", LFSuffix)
        DocName = Replace(DocName, "Attrib03", DocumentID)
        DocName = Replace(DocName, "Attrib04", Branch)
        DocName = Replace(DocName, "Attrib05", InsurerCode)
        DocName = Replace(DocName, "Attrib06", DocumentNumber)
        DocName = Replace(DocName, "Attrib07", SubscriberCode)

        ''''Before modification on 20180619
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'If UCase(DocumentType) <> "OFFER" And UCase(DocumentType) <> "POLICY" Then
        '    DocName = DocumentType & DocName
        'End If
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Log(Now & " DocName = " & DocName)
        If UCase(DocumentType) = "TITLE DEAD" Then

            Dim finalsearch As String = ""

            For Each str As String In DocumentID.split(";")
                'Replace(LFSearch, "Attrib01", DocumentType)
                finalsearch = finalsearch + Replace(LFSearch, "Attrib03", str) + " | "
            Next
            finalsearch = finalsearch.Substring(0, finalsearch.Length - 3)
            'MsgBox(finalsearch)
            If UCase(LFSuffix) = "VIEW ALL" Then

                'If LF was not open, open it to show the document
                If ClientConnected = False Then
                    Try
                        Log("Laserfiche Client Interface is Closed, Trying to load it")
                        ie.LoadLaserFiche(True)

                        ie.Login(LFDBName, LFServerName, LFUserName, LFUserPassw, 2)

                    Catch ex3 As Exception
                        MsgBox("Error loading Laserfiche = " & Err.Description)
                        Log(Now & " Error loading Laserfiche = " & Err.Description)
                        GoTo Finish
                    End Try
                End If
                ie.Search(finalsearch)

            Else
                ClientConnect()
                Dim TitleDeadSearch As LFSearch = DB.CreateSearch()
                TitleDeadSearch.Command = finalsearch
                TitleDeadSearch.BeginSearch(True)
                Dim doc As New LFDocument
                Dim Hits As ILFCollection = TitleDeadSearch.GetSearchHits()
                Dim maxPolicyNb As Integer = 0
                If Hits.Count > 0 Then
                    For Each Hit As LFSearchHit In Hits
                        Dim FD As LFFieldData = Hit.Entry.FieldData()
                        If CInt(FD.Field(FieldsList(0).Split("=")(0).Trim)) > maxPolicyNb Then
                            maxPolicyNb = CInt(FD.Field(FieldsList(0).Split("=")(0).Trim))
                            doc = Hit.Entry
                        End If
                    Next
                    ie.ViewPage(doc.ID, 1)
                    GoTo Finish

                Else

                    'Get Destination Volume
                    Try
                        LFVolume = DB.GetVolumeByName(LFVolumeName)
                    Catch ex As Exception
                        MsgBox("Cannot find volume " & LFVolumeName)
                        Log(Now & " Cannot find volume " & LFVolumeName)
                        GoTo Finish
                    End Try

                    'Get Parent Folder
                    Try
                        DestFolder = DB.GetEntryByPath(LFDestinationFolderPath)
                    Catch ex As Exception
                        CreateFolderByPath(LFDestinationFolderPath)
                        Try
                            DestFolder = DB.GetEntryByPath(LFDestinationFolderPath)
                        Catch ex1 As Exception
                            MsgBox("Cannot create destination folder" & vbCrLf & ex1.Message)
                            Log(Now & " Cannot create destination folder = " & ex1.Message)
                            GoTo Finish
                        End Try
                    End Try


                    'Create Doc
                    Try
                        doc.Create(DocName, DestFolder, LFVolume, False)
                        LFDocID = doc.ID
                        DestFolder.Dispose()
                        Log(Now & " Document created = " & LFDocID)
                    Catch ex As Exception
                        DestFolder.Dispose()
                        MsgBox("Cannot create new document" & vbCrLf & ex.Message)
                        Log(Now & " Cannot create new document = " & ex.Message)
                        GoTo Finish
                    End Try


                    Log(Now & " Template name = " & LFTemplateName)

                    Dim DocTemplate As New LFTemplate
                    Try
                        DocTemplate = DB.GetTemplateByName(LFTemplateName)
                    Catch ex As Exception
                        MsgBox("Cannot find Template " & LFTemplateName & " (DocType = " & DocumentType)
                        Log(Now & " Cannot find Template " & LFTemplateName & " (DocType = " & DocumentType)
                        GoTo Finish
                    End Try


                    Dim LFIndex As LFFieldData

                    Try
                        LFIndex = doc.FieldData
                        LFIndex.LockObject(Lock_Type.LOCK_TYPE_WRITE)
                        LFIndex.Template = DocTemplate
                        Dim FieldName As String = ""
                        Dim FieldValue As String = ""
                        Try

                            Log(Now & " Preparing Fields Values")
                            If FieldsList.Count > 0 Then

                                For Each fieldSeting As String In FieldsList

                                    FieldName = fieldSeting.Split("=")(0).Trim
                                    FieldValue = fieldSeting.Split("=")(1).Trim

                                    FieldValue = Replace(FieldValue, "Attrib01", DocumentType)
                                    FieldValue = Replace(FieldValue, "Attrib02", LFSuffix)
                                    FieldValue = Replace(FieldValue, "Attrib03", DocumentID)
                                    FieldValue = Replace(FieldValue, "Attrib04", Branch)
                                    FieldValue = Replace(FieldValue, "Attrib05", InsurerCode)
                                    FieldValue = Replace(FieldValue, "Attrib06", DocumentNumber)
                                    FieldValue = Replace(FieldValue, "Attrib07", SubscriberCode)

                                    LFIndex.Field(FieldName) = FieldValue
                                Next
                            End If
                        Catch ex As Exception
                            Log(Now & " Error while setting Fields Values = " & Err.Description)
                        End Try
                        ''''''
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



                        LFIndex.Update()
                        LFIndex.UnlockObject()
                    Catch ex As Exception
                        MsgBox("Error indexing Document" & vbCrLf & Err.Description)
                        Log(Now & " Error indexing Document = " & Err.Description)
                        GoTo Finish
                    End Try


                    doc.Dispose()

                    'Open Scanning Interface
                    Log(Now & " Opening Scanning Interface")
                    ie.ViewPage(LFDocID, 1)

                End If


            End If

        Else



            Dim LFDoc As New LFDocument

            Try
                'If LF was not open, open it to show the document
                If ClientConnected = False Then
                    Try
                        Log("Laserfiche Client Interface is Closed, Trying to load it")
                        LoadLaserfiche()

                    Catch ex3 As Exception
                        MsgBox("Error loading Laserfiche = " & Err.Description)
                        Log(Now & " Error loading Laserfiche = " & Err.Description)
                        GoTo Finish
                    End Try
                End If
                Log("Search Syntax: " & LFSearch)

                If LFSearch <> "" Then

                    LFSearch = Replace(LFSearch, "Attrib01", DocumentType)
                    LFSearch = Replace(LFSearch, "Attrib02", LFSuffix)
                    LFSearch = Replace(LFSearch, "Attrib03", DocumentID)
                    LFSearch = Replace(LFSearch, "Attrib04", Branch)
                    LFSearch = Replace(LFSearch, "Attrib05", InsurerCode)
                    LFSearch = Replace(LFSearch, "Attrib06", DocumentNumber)
                    LFSearch = Replace(LFSearch, "Attrib07", SubscriberCode)
                    Log("Search Syntax after replace: " & LFSearch)
                    'Search LFSO
                    Dim NewSearch As LFSearch = DB.CreateSearch()
                    NewSearch.Command = LFSearch
                    NewSearch.BeginSearch(True)
                    Dim Hits As ILFCollection = NewSearch.GetSearchHits()

                    For Each Hit As LFSearchHit In Hits
                        LFDoc = Hit.Entry
                    Next
                    Log("Search result: " & LFDoc.Name)
                    If LFDoc.Name <> DocName Then
                        DocName = LFDoc.Name
                        Log(Now & " DocName after search = " & DocName)
                    End If

                End If

                'LFDoc = DB.GetEntryByPath(LFDestinationFolderPath & "\" & DocName)
                LFDocID = LFDoc.ID
                Log(Now & " DocID = " & LFDocID)
                'LFDoc.Dispose()
                Log(Now & " Opening Laserfiche Document Viewer")
                ie.ViewPage(LFDocID, 1)
                Log(Now & " Laserfiche Document Viewer opened")
                'Close the connection
                GoTo Finish

            Catch ex As Exception
                'Document doesnt exist so continue to create it
                If LFUserName = "VIEWONLY" Then
                    MsgBox("Sorry but this document does not exist in Laserfiche")
                    GoTo Finish
                Else
                    Log(Now & "Error: " & ex.Message)
                End If
            End Try


            'Get Destination Volume
            Try
                LFVolume = DB.GetVolumeByName(LFVolumeName)
            Catch ex As Exception
                MsgBox("Cannot find volume " & LFVolumeName)
                Log(Now & " Cannot find volume " & LFVolumeName)
                GoTo Finish
            End Try

            'Get Parent Folder
            Try
                DestFolder = DB.GetEntryByPath(LFDestinationFolderPath)
            Catch ex As Exception
                CreateFolderByPath(LFDestinationFolderPath)
                Try
                    DestFolder = DB.GetEntryByPath(LFDestinationFolderPath)
                Catch ex1 As Exception
                    MsgBox("Cannot create destination folder" & vbCrLf & ex1.Message)
                    Log(Now & " Cannot create destination folder = " & ex1.Message)
                    GoTo Finish
                End Try
            End Try


            'Create Doc
            Try
                LFDoc.Create(DocName, DestFolder, LFVolume, False)
                LFDocID = LFDoc.ID
                DestFolder.Dispose()
                Log(Now & " Document created = " & LFDocID)
            Catch ex As Exception
                DestFolder.Dispose()
                MsgBox("Cannot create new document" & vbCrLf & ex.Message)
                Log(Now & " Cannot create new document = " & ex.Message)
                GoTo Finish
            End Try


            'Index it

            'Get Destination Template


            ''''Before modification on 20180619
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'Dim LFTemplateName As String
            'If UCase(DocumentType) = "OFFER" Then
            '    LFTemplateName = LFOTemplateName
            'ElseIf UCase(DocumentType) = "POLICY" Then
            '    LFTemplateName = LFPTemplateName
            'Else
            '    LFTemplateName = "General"
            'End If

            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ''''Modification applied on 20180619
            'Set Fields Value based on DocType config File
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



            Log(Now & " Template name = " & LFTemplateName)

            Dim DocTemplate As New LFTemplate
            Try
                DocTemplate = DB.GetTemplateByName(LFTemplateName)
            Catch ex As Exception
                MsgBox("Cannot find Template " & LFTemplateName & " (DocType = " & DocumentType)
                Log(Now & " Cannot find Template " & LFTemplateName & " (DocType = " & DocumentType)
                GoTo Finish
            End Try


            Dim LFIndex As LFFieldData

            Try
                LFIndex = LFDoc.FieldData
                LFIndex.LockObject(Lock_Type.LOCK_TYPE_WRITE)
                LFIndex.Template = DocTemplate
                ''''Before modification on 20180619
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                'If UCase(DocumentType) = "OFFER" Then

                '    LFIndex.Field("OfferOID") = DocumentNumber & "-" & LFSuffix
                '    LFIndex.Field("Branch") = Branch
                '    LFIndex.Field("OfferNumber") = DocumentID
                '    LFIndex.Field("SubscriberCode") = SubscriberCode

                'ElseIf UCase(DocumentType) = "POLICY" Then

                '    LFIndex.Field("PolicyOID") = DocumentNumber & "-" & LFSuffix
                '    LFIndex.Field("Branch") = Branch
                '    LFIndex.Field("InsurerCode") = InsurerCode
                '    LFIndex.Field("PolicyNumber") = DocumentID
                '    LFIndex.Field("SubscriberCode") = SubscriberCode

                'End If
                ''''Modification applied on 20180619
                'Set Fields Values
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                Dim FieldName As String = ""
                Dim FieldValue As String = ""
                Try

                    Log(Now & " Preparing Fields Values")
                    If FieldsList.Count > 0 Then

                        For Each fieldSeting As String In FieldsList

                            FieldName = fieldSeting.Split("=")(0).Trim
                            FieldValue = fieldSeting.Split("=")(1).Trim

                            FieldValue = Replace(FieldValue, "Attrib01", DocumentType)
                            FieldValue = Replace(FieldValue, "Attrib02", LFSuffix)
                            FieldValue = Replace(FieldValue, "Attrib03", DocumentID)
                            FieldValue = Replace(FieldValue, "Attrib04", Branch)
                            FieldValue = Replace(FieldValue, "Attrib05", InsurerCode)
                            FieldValue = Replace(FieldValue, "Attrib06", DocumentNumber)
                            FieldValue = Replace(FieldValue, "Attrib07", SubscriberCode)

                            LFIndex.Field(FieldName) = FieldValue
                        Next
                    End If
                Catch ex As Exception
                    Log(Now & " Error while setting Fields Values = " & Err.Description)
                End Try
                ''''''
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



                LFIndex.Update()
                LFIndex.UnlockObject()
            Catch ex As Exception
                MsgBox("Error indexing Document" & vbCrLf & Err.Description)
                Log(Now & " Error indexing Document = " & Err.Description)
                GoTo Finish
            End Try


            LFDoc.Dispose()

            'Open Scanning Interface
            Log(Now & " Opening Scanning Interface")
            ie.ShowScanWindow(LFDocID)
        End If
Finish:
        Try
            Dim ShellCmd As String = Application.StartupPath & "\NirCmd.exe win close title " & """" & LFDBName & " - Laserfiche" & """"
            'Track(ShellCmd)
            Shell(ShellCmd)
        Catch ex As Exception
            Log(Now & " NirCmd Error closing window for Repository")
        End Try
        Disconnect()

        End
    End Sub
End Class
