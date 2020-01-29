Imports System.IO

'Main UI
'Building the UI programatically 
Public Class DirectoryExplorer
    Inherits Form
    Private exDirectoryTreeView As ExtendedDirectoryTreeView
    Private list As New ListView()
    Private selectedNode As TreeNode

    Sub New()
        Dim split As New Splitter()

        Text = "Directory Explorer"
        BackColor = SystemColors.Window
        ForeColor = SystemColors.WindowText

        ' Add ListView Control
        ' Detail View
        list.View = View.Details
        list.Parent = Me
        ' Create columns.
        list.Columns.Add("File Name", 100, HorizontalAlignment.Left)
        list.Columns.Add("Size (KB)", 100, HorizontalAlignment.Right)
        list.Columns.Add("Modified", 100, HorizontalAlignment.Left)
        list.Columns.Add("Attribute", 100, HorizontalAlignment.Left)
        list.Dock = DockStyle.Fill

        ' Addding the split 
        split.Parent = Me
        Split.Dock = DockStyle.Left
        Split.BackColor = SystemColors.Control

        'Assigning the treeview and docking to the left side of the UI
        exDirectoryTreeView = New ExtendedDirectoryTreeView()
        exDirectoryTreeView.Parent = Me
        exDirectoryTreeView.Dock = DockStyle.Left

        'Adding Event Handler when AfterSelect Event fires
        AddHandler exDirectoryTreeView.AfterSelect,
                                AddressOf DirectoryTreeViewOnAfterSelect

        ' Create menu with one item.
        Menu = New MainMenu()
        Menu.MenuItems.Add("View")
        Dim mi As New MenuItem("Refresh",
                        AddressOf RefreshTreeView, Shortcut.F5)
        Menu.MenuItems(0).MenuItems.Add(mi)
    End Sub


    Private Sub DirectoryTreeViewOnAfterSelect(ByVal obj As Object,
                                    ByVal args As TreeViewEventArgs)
        'Assign the selected node
        selectedNode = args.Node

        ' Check whether the List contains any items
        ' If so, clear them first
        If Not IsNothing(list.Items) Then
            list.Items.Clear()
        End If

        'Declare variables for processing the file information
        Dim currentDirectoryInfo As New DirectoryInfo(selectedNode.FullPath)
        Dim arrayOfFiles() As FileInfo

        Try
            ' Get all the file names from the current directory
            arrayOfFiles = currentDirectoryInfo.GetFiles()
        Catch
            Return
        End Try

        For Each fileInfo In arrayOfFiles
            ' Create ListViewItem.
            Dim lvi As New ListViewItem(fileInfo.Name)

            ' Assign ImageIndex based on filename extension.
            If Path.GetExtension(fileInfo.Name).ToUpper() = ".EXE" Then
                lvi.ImageIndex = 1
            Else
                lvi.ImageIndex = 0
            End If

            ' Add file length and modified time sub-items.
            lvi.SubItems.Add(fileInfo.Length.ToString("N0"))
            lvi.SubItems.Add(fileInfo.LastWriteTime.ToString())

            ' Add attribute subitem.
            Dim strAttr As String = ""
            If (fileInfo.Attributes And FileAttributes.Archive) <> 0 Then
                strAttr &= "A"
            End If
            If (fileInfo.Attributes And FileAttributes.Hidden) <> 0 Then
                strAttr &= "H"
            End If
            If (fileInfo.Attributes And FileAttributes.ReadOnly) <> 0 Then
                strAttr &= "R"
            End If
            If (fileInfo.Attributes And FileAttributes.System) <> 0 Then
                strAttr &= "S"
            End If
            lvi.SubItems.Add(strAttr)

            ' Add completed ListViewItem to ListView.
            list.Items.Add(lvi)
        Next fileInfo

    End Sub

    'Refresh Event Handler
    Private Sub RefreshTreeView(ByVal obj As Object, ByVal ea As EventArgs)
        exDirectoryTreeView.BuildTree()
    End Sub

End Class
