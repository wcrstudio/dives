Public Class MruList
    Private m_ApplicationName As String
    Private m_FileMenu As RibbonOrbMenuItem
    Private m_NumEntries As Integer
    Private m_FileNames As Collection
    Private m_MenuItems As Collection

    Public Event OpenFile(ByVal file_name As String)

    Public Sub New(ByVal application_name As String, ByVal num_entries As Integer)
        m_ApplicationName = application_name
        m_NumEntries = num_entries
        m_FileNames = New Collection
        m_MenuItems = New Collection

        ' Load saved file names from the Registry.
        LoadMruList()

        ' Display the MRU list.
        DisplayMruList()
    End Sub

    ' Load previously saved file names from the Registry.
    Private Sub LoadMruList()
        Dim file_name As String
        For i As Integer = 1 To m_NumEntries
            ' Get the next file name and title.
            file_name = IniApp.IniReadValue("MruList", "FileName" & i)

            ' See if we got anything.
            If file_name.Length > 0 Then
                ' Save this file name.
                m_FileNames.Add(file_name, file_name)
                MainForm.rOrbRecentItem1.Text = file_name

            End If
        Next i
    End Sub

    ' Save the MRU list into the Registry.
    Private Sub SaveMruList()

        ' Make the new entries.
        For i As Integer = 1 To m_FileNames.Count
            IniApp.IniWriteValue("MruList", "FileName" & i, _
                m_FileNames(i).ToString)
        Next i
    End Sub

    ' MRU menu item event handler.
    Private Sub MruItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim mnu As MenuItem = DirectCast(sender, MenuItem)

        ' Find the menu item that raised this event.
        For i As Integer = 1 To m_FileNames.Count
            ' See if this is the item. (Add 1 for the separator.)
            If m_MenuItems(i + 1) Is mnu Then
                ' This is the item. Raise the OpenFile 
                ' event for its file name.
                RaiseEvent OpenFile(m_FileNames(i).ToString)
                Exit For
            End If
        Next i
    End Sub

    ' Display the MRU list.
    Private Sub DisplayMruList()
        ' Remove old menu items from the File menu.
        For Each mnu As RibbonOrbMenuItem In m_MenuItems
        Next mnu
        m_MenuItems = New Collection

        ' See if we have any file names.
        If m_FileNames.Count > 0 Then
            ' Make the separator.
            
            ' Make the other menu items.
            For i As Integer = 1 To m_FileNames.Count
                Try
                    If i = 1 Then
                        MainForm.rOrbRecentItem1.Text = FileTitle(m_FileNames(1).ToString)
                        MainForm.rOrbRecentItem1.ToolTip = m_FileNames(1).ToString
                        MainForm.rOrbRecentItem1.Enabled = True
                        MainForm.rOrbRecentItem1.FlashSmallImage = My.Resources.docdives4_16
                        MainForm.rOrbRecentItem1.ShowFlashImage = True
                    End If
                    If i = 2 Then
                        MainForm.rOrbRecentItem2.Text = FileTitle(m_FileNames(2).ToString)
                        MainForm.rOrbRecentItem2.ToolTip = m_FileNames(2).ToString
                        MainForm.rOrbRecentItem2.Enabled = True
                    End If
                    If i = 3 Then
                        MainForm.rOrbRecentItem3.Text = FileTitle(m_FileNames(3).ToString)
                        MainForm.rOrbRecentItem3.ToolTip = m_FileNames(3).ToString
                        MainForm.rOrbRecentItem3.Enabled = True
                    End If
                    If i = 4 Then
                        MainForm.rOrbRecentItem4.Text = FileTitle(m_FileNames(4).ToString)
                        MainForm.rOrbRecentItem4.ToolTip = m_FileNames(4).ToString
                        MainForm.rOrbRecentItem4.Enabled = True
                    End If
                    If i = 5 Then
                        MainForm.rOrbRecentItem5.Text = FileTitle(m_FileNames(5).ToString)
                        MainForm.rOrbRecentItem5.ToolTip = m_FileNames(5).ToString
                        MainForm.rOrbRecentItem5.Enabled = True
                    End If
                Catch ex As Exception
                    LibErro.ErrorViewSave(ex)
                End Try
            Next i
        End If
    End Sub

    ' Add a file to the MRU list.
    Public Sub Add(ByVal file_name As String)
        ' Remove this file from the MRU list
        ' if it is present.
        Dim i As Integer = FileNameIndex(file_name)
        If i > 0 Then m_FileNames.Remove(i)

        ' Add the item to the begining of the list.
        If m_FileNames.Count > 0 Then
            m_FileNames.Add(file_name, file_name, m_FileNames.Item(1))
        Else
            m_FileNames.Add(file_name, file_name)
        End If

        ' If the list is too long, remove the last item.
        If m_FileNames.Count > m_NumEntries Then
            m_FileNames.Remove(m_NumEntries + 1)
        End If

        ' Display the list.
        DisplayMruList()

        ' Save the updated list.
        SaveMruList()
    End Sub

    ' Return the index of this file in the list.
    Private Function FileNameIndex(ByVal file_name As String) As Integer
        For i As Integer = 1 To m_FileNames.Count
            If m_FileNames(i).ToString = file_name Then Return i
        Next i
        Return 0
    End Function

    ' Remove a file from the MRU list.
    Public Sub Remove(ByVal file_name As String)
        ' See if the file is present.
        Dim i As Integer = FileNameIndex(file_name)
        If i > 0 Then
            ' Remove the file.
            m_FileNames.Remove(i)

            ' Display the list.
            DisplayMruList()

            ' Save the updated list.
            SaveMruList()
        End If
    End Sub

End Class
