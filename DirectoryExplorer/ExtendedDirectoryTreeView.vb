Imports System.Drawing
Imports System.IO

'Inherits from TreeView Control. This Class will populate all the drives, directories and sub-directories
Public Class ExtendedDirectoryTreeView
    Inherits TreeView
    Sub New()

        ' Get images for tree.
        ' The images are embedded as resource
        ImageList = New ImageList()
        ImageList.Images.Add(New Bitmap(Me.GetType(), "HARDDRIVE.png"))
        ImageList.Images.Add(New Bitmap(Me.GetType(), "CLOSEFOLDER.png"))
        ImageList.Images.Add(New Bitmap(Me.GetType(), "OPENFOLDER.png"))
        BuildTree()

    End Sub

    ' BuildTree method builds the initial tree structure and populates with drives and directories as Nodes
    Sub BuildTree()
        ' Make disk drives the root nodes. Get the logical drive names
        Dim arrayOfDrives() As String = Directory.GetLogicalDrives()
        Dim strDriveName As String

        ' Turn off UI updating and clear tree.
        ' Adding this one line improved the UI 
        BeginUpdate()
        ' I read clearing the tree view is good practice
        Nodes.Clear()

        ' Loop through all the drives
        For Each strDriveName In arrayOfDrives
            Dim driveNode As New TreeNode(strDriveName, 0, 0)
            ' Add each drive to the Tree as Node
            Nodes.Add(driveNode)
            ' Add directories for each drive to the Tree 
            AddDirectories(driveNode)
            ' Initial selection of Drive
            If strDriveName = "C:\" Then
                'Assign the drive to SelectedNode property of TreeView
                SelectedNode = driveNode
            End If
        Next strDriveName
        ' Finish Update
        EndUpdate()
    End Sub

    ' This method populates all directories of a given node in the TreeNode
    Private Sub AddDirectories(ByVal currentNode As TreeNode)
        'Get the current Drive Full Path
        Dim strDriveFullPath As String = currentNode.FullPath
        'Get the Root Directory Info
        Dim rootDirectoryInfo As New DirectoryInfo(strDriveFullPath)
        Dim arrayOfSubDirectories() As DirectoryInfo
        Dim currentDirectoryInfo As DirectoryInfo

        'Clear the nodes
        currentNode.Nodes.Clear()

        ' Check whether Directory Exists 
        If Not rootDirectoryInfo.Exists Then Return

        Try
            'Get Sub-Directories, there are any
            arrayOfSubDirectories = rootDirectoryInfo.GetDirectories()
        Catch
            'If any error occurs just return
            Return
        End Try

        'Loop through sub directories and get the name of the directory and create node
        For Each currentDirectoryInfo In arrayOfSubDirectories
            Dim directoryNode As New TreeNode(currentDirectoryInfo.Name, 1, 2)
            currentNode.Nodes.Add(directoryNode)
        Next currentDirectoryInfo

    End Sub

    'Overriding method
    Protected Overrides Sub OnBeforeExpand(ByVal args As TreeViewCancelEventArgs)
        MyBase.OnBeforeExpand(args)
        BeginUpdate()
        'When a root node expanded , add directores as child nodes
        For Each currentTreeNode In args.Node.Nodes
            AddDirectories(currentTreeNode)
        Next currentTreeNode
        EndUpdate()
    End Sub
End Class