Imports System.ComponentModel
Imports System.IO
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Public Class Form1
    Dim selectedValfAdasi As String
    Dim selectedValf As String
    Dim selectedComponent As String
    Dim ValfAdalariName As String = "Valf Adalari"
    Dim myImageList As New ImageList()
    Public ValfAdasisList As New List(Of ValfAdasi)
    Public ValfAdasisList2 As New List(Of ValfAdasi)
    Public redNames As Integer = 0
    Public redNames2 As Integer = 0
    Public myTree As TreeView
    Public firstTree As Boolean = True
    Public myList As New List(Of ValfAdasi)
    Public partsFolderPath As String = ""
    Public sensorsFolderPath As String = ""
    Public folderPath As String = ""
    Public excelFileName As String = ""
    Private Sub UpdateStripMenuNames()
        If selectedValfAdasi = "" Then
            UpdateToolStripMenuItem.Enabled = True
            UpdateToolStripMenuItem.Text = "Add Valf Adasi"
            RenameToolStripMenuItem.Text = "Rename Valf Adalari"
            DeleteToolStripMenuItem.Enabled = False
            RenameToolStripMenuItem.Enabled = True
            DeleteToolStripMenuItem.Text = "Delete"
            ShowScreenshotToolStripMenuItem.Enabled = False
        ElseIf selectedValf = "" Then
            DeleteToolStripMenuItem.Enabled = False
            DeleteToolStripMenuItem.Text = "Delete"
            UpdateToolStripMenuItem.Text = "Add Valf"
            UpdateToolStripMenuItem.Enabled = True
            RenameToolStripMenuItem.Enabled = True
            RenameToolStripMenuItem.Text = "Rename Valf Adasi"
            ShowScreenshotToolStripMenuItem.Enabled = False
        ElseIf selectedComponent = "" Then
            ShowScreenshotToolStripMenuItem.Enabled = True
            UpdateToolStripMenuItem.Enabled = True
            RenameToolStripMenuItem.Enabled = True
            DeleteToolStripMenuItem.Enabled = True
            If firstTree Then
                UpdateToolStripMenuItem.Text = "Add Components"
            Else
                UpdateToolStripMenuItem.Text = "Add Sensor"
            End If
            RenameToolStripMenuItem.Text = "Rename Valf"
            DeleteToolStripMenuItem.Text = "Delete Valf"
        Else
            RenameToolStripMenuItem.Text = "Rename"
            UpdateToolStripMenuItem.Text = "Add"
            UpdateToolStripMenuItem.Enabled = False
            RenameToolStripMenuItem.Enabled = False
            DeleteToolStripMenuItem.Enabled = False
            DeleteToolStripMenuItem.Text = "Delete"
            ShowScreenshotToolStripMenuItem.Enabled = False
        End If
    End Sub
    Public Sub countRedNames()
        Dim redNamesCount As Integer = 0
        For Each myValfAdasi As ValfAdasi In myList
            For Each myValf As Valf In myValfAdasi.ValfsList
                If firstTree Then
                    For Each myComponent As Component In myValf.ComponentsList
                        If myComponent.MakeRed = True Then
                            redNamesCount += 1
                        End If
                    Next
                Else ' new 
                    For Each mySensor As Sensor In myValf.SensorsList
                        If mySensor.MakeRed = True Then
                            redNamesCount += 1
                        End If
                    Next
                End If
            Next
        Next
        If firstTree Then
            redNames = redNamesCount
        Else
            redNames2 = redNamesCount
        End If
    End Sub
    Private Sub UpdateTreeView(m As Integer)
        countRedNames()
        myTree.Nodes.Clear()
        Dim ValfAdalariNode As TreeNode = myTree.TopNode
        ValfAdalariNode = myTree.Nodes.Add(ValfAdalariName)
        ValfAdalariNode.ImageIndex = 0
        ValfAdalariNode.SelectedImageIndex = 0
        For Each myValfAdasi As ValfAdasi In myList
            Dim ValfAdasis As TreeNode
            ValfAdasis = ValfAdalariNode.Nodes.Add(myValfAdasi.Name)
            ValfAdasis.ImageIndex = 1
            ValfAdasis.SelectedImageIndex = 1
            For Each myValf As Valf In myValfAdasi.ValfsList
                Dim Valfs As TreeNode
                Valfs = ValfAdasis.Nodes.Add(myValf.Name)
                Valfs.ImageIndex = 2
                Valfs.SelectedImageIndex = 2
                If firstTree Then
                    For Each myComponent As Component In myValf.ComponentsList
                        Dim ComponentsNode As TreeNode
                        ComponentsNode = Valfs.Nodes.Add(myComponent.Name)
                        If myComponent.MakeRed = vbNull Then
                        ElseIf myComponent.MakeRed = True Then
                            ComponentsNode.BackColor = Color.Tomato
                        ElseIf myComponent.MakeRed = False Then
                            ComponentsNode.BackColor = Color.Empty
                        End If
                        ComponentsNode.ImageIndex = 3
                        ComponentsNode.SelectedImageIndex = 3
                    Next
                Else ' new 
                    For Each mySensor As Sensor In myValf.SensorsList
                        Dim ComponentsNode As TreeNode
                        ComponentsNode = Valfs.Nodes.Add(mySensor.Name)
                        If mySensor.MakeRed = vbNull Then
                        ElseIf mySensor.MakeRed = True Then
                            ComponentsNode.BackColor = Color.Tomato
                        ElseIf mySensor.MakeRed = False Then
                            ComponentsNode.BackColor = Color.Empty
                        End If
                        ComponentsNode.ImageIndex = 3
                        ComponentsNode.SelectedImageIndex = 3
                    Next

                End If
                If firstTree Then
                    If Valfs.Nodes.Count < 6 Then
                        For i As Integer = Valfs.Nodes.Count To 5
                            Dim aNode As TreeNode
                            aNode = Valfs.Nodes.Add("Empty")
                            aNode.ImageIndex = 4
                            aNode.SelectedImageIndex = 4
                        Next
                    End If
                End If
                If ValfAdasis.LastNode IsNot Nothing Then
                    If ValfAdasis.Text = selectedValfAdasi And Valfs.Text = selectedValf Then
                        Valfs.Expand()
                    End If
                End If
            Next
            ValfAdasis.Expand()
            If ValfAdasis.LastNode IsNot Nothing Then
                If ValfAdasis.Text = selectedValfAdasi And m = 1 Then
                    ValfAdasis.LastNode.Expand()
                End If
            End If
        Next
        ValfAdalariNode.Expand()
    End Sub
    Private Sub DeleteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteToolStripMenuItem.Click '''
        delete()
    End Sub
    Private Sub TreeView1_NodeMouseClick(sender As Object, e As System.Windows.Forms.TreeNodeMouseClickEventArgs) Handles TreeView1.NodeMouseClick
        ''''''''
        If e.Button = System.Windows.Forms.MouseButtons.Right Then
            TreeView1.SelectedNode = e.Node
            Dim fullPath As String
            Dim slashCount As Integer = 0
            fullPath = TreeView1.SelectedNode.FullPath
            slashCount = CountCharacter(fullPath, "\")
            Select Case (slashCount)
                Case 0
                    selectedValfAdasi = ""
                    selectedValf = ""
                    selectedComponent = ""
                Case 1
                    selectedValfAdasi = TreeView1.SelectedNode.Text
                    selectedValf = ""
                    selectedComponent = ""
                Case 2
                    selectedValfAdasi = TreeView1.SelectedNode.Parent.Text
                    selectedValf = TreeView1.SelectedNode.Text
                    selectedComponent = ""
                Case 3
                    selectedValfAdasi = TreeView1.SelectedNode.Parent.Parent.Text
                    selectedValf = TreeView1.SelectedNode.Parent.Text
                    selectedComponent = TreeView1.SelectedNode.Text
            End Select
            UpdateStripMenuNames()
        End If
        firstTree = True
        myTree = TreeView1
        myList = ValfAdasisList
        If PictureBox1.Visible = True Then
            PictureBox1.Visible = False
        End If
    End Sub
    Public Sub delete()
        Highlight()
        Dim msgBoxAnswer As String
        msgBoxAnswer = MsgBox("Delete highlighted item(s)?", 33, "Delete").ToString
        If msgBoxAnswer = "Ok" Then
            If selectedValfAdasi = "" Then
            ElseIf selectedValf = "" Then
            ElseIf selectedComponent = "" Then

                DeleteValf(SearchForValfAdasi(selectedValfAdasi), SearchForValf(selectedValfAdasi, selectedValf))
            End If
        End If
        UnMakeRed()
        UpdateTreeView(0)
    End Sub
    Public Sub UnMakeRed()
        For Each myvalfadasi As ValfAdasi In myList
            For Each myvalf As Valf In myvalfadasi.ValfsList
                If firstTree Then
                    For Each mycomponent As Component In myvalf.ComponentsList
                        If mycomponent.MakeRed = True Then
                            Dim duplica = duplicateComponent(myvalfadasi, myvalf, mycomponent.Name)
                            If duplica Is Nothing Then
                                mycomponent.MakeRed = False
                            End If
                        End If
                    Next
                Else ' new 
                    For Each mySensor As Sensor In myvalf.SensorsList
                        If mySensor.MakeRed = True Then
                            Dim duplica = duplicateSensor(myvalfadasi, myvalf, mySensor.Name)
                            If duplica Is Nothing Then
                                mySensor.MakeRed = False
                            End If
                        End If
                    Next
                End If

            Next
        Next
    End Sub
    Public Sub UpdateValf()
        Dim myvalf As Valf = SearchForValf(selectedValfAdasi, selectedValf)
        Dim mycomponent As Component = New Component
        If firstTree Then
            Dim i As Integer = myvalf.ComponentsList.Count
            While i > 0
                myvalf.ComponentsList.RemoveAt(i - 1)
                i -= 1
            End While
        Else ' new 
            Dim i As Integer = myvalf.SensorsList.Count
            While i > 0
                myvalf.SensorsList.RemoveAt(i - 1)
                i -= 1
            End While
        End If
        add()
        UnMakeRed()
        UpdateTreeView(0)
    End Sub
    Public Sub add()
        If selectedValfAdasi = "" Then
            AddValfAdasi()
        ElseIf selectedValf = "" Then
            AddValf(SearchForValfAdasi(selectedValfAdasi))
        ElseIf selectedComponent = "" Then
            Dim numText As String = InputBox("number of components: ", "Enter")
            Dim num As Integer
            Try
                num = Convert.ToInt32(numText)
            Catch
                num = 0
            End Try
            For i As Integer = 1 To num
                Dim ExportName As String = InputBox("Enter a Component's name")
                Dim GroupeName As String = InputBox("Enter a Group's name")
                Dim ComponentName = ExportName & " : " & GroupeName
                If firstTree Then
                    SearchForComponent(ComponentName)
                    AddComponent(SearchForValf(selectedValfAdasi, selectedValf), ComponentName)
                    Dim myComponent As Component = SearchForComponent(ComponentName)
                    If Not (myComponent Is Nothing) Then
                        myComponent.GroupName = GroupeName
                        myComponent.ExportName = ExportName
                    End If
                Else ' new 
                    SearchForSensor(ComponentName)
                    AddSensor(SearchForValf(selectedValfAdasi, selectedValf), ComponentName)
                    Dim mySensor As Sensor = SearchForSensor(ComponentName)
                    If Not (mySensor Is Nothing) Then
                        mySensor.GroupName = GroupeName
                        mySensor.ExportName = ExportName
                    End If
                End If
            Next
            SaveScreenShot(SearchForValf(selectedValfAdasi, selectedValf))
            'getScreenShot().Save("C:\Users\ahmad\Documents\ahmad.jpg", System.Drawing.Imaging.ImageFormat.Jpeg)
            UpdateTreeView(1)
            showScreenshot(SearchForValf(selectedValfAdasi, selectedValf))
        End If
        UpdateTreeView(1)
    End Sub
    Private Sub Highlight()
        Dim aNode As TreeNode = myTree.SelectedNode
        aNode.BackColor = Color.Red
        For Each bNode As TreeNode In aNode.Nodes
            bNode.BackColor = Color.Red
            For Each cNode As TreeNode In bNode.Nodes
                cNode.BackColor = Color.Red
                For Each dNode As TreeNode In cNode.Nodes
                    dNode.BackColor = Color.Red
                Next
            Next
        Next
    End Sub
    Public Sub AddValfAdasi()
        Dim newName As String = "Valf Adasi " & (myList.Count + 1).ToString
        If isValfAdasiValid(newName) And ((myList.Count < 3 And firstTree) Or Not firstTree) Then
            Dim myValfAdasi As ValfAdasi = New ValfAdasi()
            myValfAdasi.Name = newName
            myList.Add(myValfAdasi)
        ElseIf Not isValfAdasiValid(newName) Then
            MsgBox(newName & " already exists.", 48, "Invalid Operation")
        Else

            MsgBox("Maximum number of valve islands reached.", 48, "Invalid Operation")
        End If
    End Sub
    Public Function SearchForValfAdasi(ValfAdasisName As String) As ValfAdasi
        For Each myValfAdasi As ValfAdasi In myList
            If myValfAdasi.Name = ValfAdasisName Then
                Return myValfAdasi
            End If
        Next
        Return Nothing
    End Function
    Public Function SearchForValf(ValfAdasisName As String, ValfsName As String) As Valf
        For Each myValfAdasi As ValfAdasi In myList
            If myValfAdasi.Name = ValfAdasisName Then
                For Each myValf As Valf In myValfAdasi.ValfsList
                    If myValf.Name = ValfsName Then
                        Return myValf
                    End If
                Next
            End If
        Next
        Return Nothing
    End Function
    Public Function SearchForComponent(ComponentsName As String) As Component
        For Each myValfAdasi As ValfAdasi In myList
            For Each myValf As Valf In myValfAdasi.ValfsList
                For Each myComponent As Component In myValf.ComponentsList
                    If myComponent.Name = ComponentsName Then
                        Return myComponent
                    End If
                Next
            Next
        Next
        Return Nothing
    End Function
    Public Function SearchForSensor(ComponentsName As String) As Sensor ' new 
        For Each myValfAdasi As ValfAdasi In myList
            For Each myValf As Valf In myValfAdasi.ValfsList
                For Each mySensor As Sensor In myValf.SensorsList
                    If mySensor.Name = ComponentsName Then
                        Return mySensor
                    End If
                Next
            Next
        Next
        Return Nothing
    End Function
    Public Sub AddValf(ByRef myValfAdasi As ValfAdasi)
        Dim newName As String = "Valf " & (myValfAdasi.ValfsList.Count + 1).ToString
        If (isValfValid(myValfAdasi.Name, newName) And ((myValfAdasi.ValfsList.Count < 12 And firstTree) Or Not firstTree)) Then
            Dim myValf As Valf = New Valf()
            myValf.Name = newName
            setValfPath(myValfAdasi, myValf)
            myValfAdasi.ValfsList.Add(myValf)
        ElseIf Not isValfValid(myValfAdasi.Name, newName) Then
            MsgBox(newName & " already exists.", 48, "Invalid Operation")
        Else
            MsgBox("Maximum number of valves reached.", 48, "Invalid Operation")
        End If
    End Sub
    Public Sub AddComponent(ByRef myValf As Valf, ByVal ComponentName As String)
        Dim duplica As Component = duplicateComponent(ComponentName)
        If firstTree And myValf.ComponentsList.Count > 5 Then
            MsgBox("Maximum number of components reached.", 48, "Invalid Operation")
        ElseIf isComponentNameinValf(myValf, ComponentName) = True Then
            MsgBox(ComponentName & " already exists in this Valf.", 48, "Invalid Operation")
        ElseIf duplica Is Nothing Then
            Dim myComponent As Component = New Component()
            myComponent.Name = ComponentName
            myValf.ComponentsList.Add(myComponent)
        ElseIf (duplica IsNot Nothing) Then
            duplica.MakeRed = True
            If firstTree Then
                redNames += 1
            Else
                redNames2 += 1
            End If
            Dim myComponent As Component = New Component()
            myComponent.Name = ComponentName
            myValf.ComponentsList.Add(myComponent)
            myComponent.MakeRed = True
            If firstTree Then
                redNames += 1
            Else
                redNames2 += 1
            End If
        End If
    End Sub
    Public Sub AddSensor(ByRef myValf As Valf, ByVal ComponentName As String) ' new 
        Dim duplica As Sensor = duplicateSensor(ComponentName)
        If firstTree And myValf.SensorsList.Count > 5 Then
            MsgBox("Maximum number of components reached.", 48, "Invalid Operation")
        ElseIf isComponentNameinValf(myValf, ComponentName) = True Then
            MsgBox(ComponentName & " already exists in this Valf.", 48, "Invalid Operation")
        ElseIf duplica Is Nothing Then
            Dim mySensor As Sensor = New Sensor()
            mySensor.Name = ComponentName
            myValf.SensorsList.Add(mySensor)
        ElseIf (duplica IsNot Nothing) Then
            duplica.MakeRed = True
            If firstTree Then
                redNames += 1
            Else
                redNames2 += 1
            End If
            Dim mySensor As Sensor = New Sensor()
            mySensor.Name = ComponentName
            myValf.SensorsList.Add(mySensor)
            mySensor.MakeRed = True
            If firstTree Then
                redNames += 1
            Else
                redNames2 += 1
            End If
        End If
    End Sub
    Public Function isComponentNameinValf(ByRef myValf, componentName) As Boolean
        If firstTree Then
            For Each myComponent As Component In myValf.ComponentsList
                If myComponent.Name = componentName Then
                    Return True
                End If
            Next
        Else ' new 
            For Each mySensor As Sensor In myValf.SensorsList
                If mySensor.Name = componentName Then
                    Return True
                End If
            Next
        End If
        Return False
    End Function

    Public Function duplicateComponent(componentName) As Component

        For Each myValfAdasi As ValfAdasi In myList
            For Each myValf As Valf In myValfAdasi.ValfsList
                For Each myComponent As Component In myValf.ComponentsList
                    If componentName = myComponent.Name Then
                        Return myComponent
                    End If
                Next
            Next
        Next
        Return Nothing
    End Function
    Public Function duplicateSensor(componentName) As Sensor ' new 
        For Each myValfAdasi As ValfAdasi In myList
            For Each myValf As Valf In myValfAdasi.ValfsList
                For Each mySensor As Sensor In myValf.SensorsList
                    If componentName = mySensor.Name Then
                        Return mySensor
                    End If
                Next
            Next
        Next
        Return Nothing
    End Function
    Public Function duplicateComponent(ByRef myValfAdasi As ValfAdasi, ByRef myValf As Valf, componentName As String) As Component
        For Each myva As ValfAdasi In myList
            For Each myv As Valf In myva.ValfsList
                For Each mycomponent As Component In myv.ComponentsList
                    If mycomponent.Name = componentName And myva.Name <> myValfAdasi.Name Then
                        Return mycomponent
                    ElseIf mycomponent.Name = componentName And myva.Name = myValfAdasi.Name And myv.Name <> myValf.Name Then
                        Return mycomponent
                    End If
                Next
            Next
        Next
        Return Nothing
    End Function
    Public Function duplicateSensor(ByRef myValfAdasi As ValfAdasi, ByRef myValf As Valf, componentName As String) As Sensor ' new 
        For Each myva As ValfAdasi In myList
            For Each myv As Valf In myva.ValfsList
                For Each mySensor As Sensor In myv.SensorsList
                    If mySensor.Name = componentName And myva.Name <> myValfAdasi.Name Then
                        Return mySensor
                    ElseIf mySensor.Name = componentName And myva.Name = myValfAdasi.Name And myv.Name <> myValf.Name Then
                        Return mySensor
                    End If
                Next
            Next
        Next
        Return Nothing
    End Function
    Public Function CountCharacter(ByVal value As String, ByVal ch As Char) As Integer
        Dim cnt As Integer = 0
        For Each c As Char In value
            If c = ch Then
                cnt += 1
            End If
        Next
        Return cnt
    End Function
    Public Function isValfAdasiValid(valfAdasiName As String) As Boolean
        For Each myValfAdasi As ValfAdasi In myList
            If myValfAdasi.Name = valfAdasiName Then
                Return False
            End If
        Next
        Return True
    End Function
    Public Function isValfValid(valfAdasiName As String, valfName As String) As Boolean
        Dim myValfAdasi As ValfAdasi = SearchForValfAdasi(valfAdasiName)
        For Each myValf As Valf In myValfAdasi.ValfsList
            If myValf.Name = valfName Then
                Return False
            End If
        Next
        Return True
    End Function
    Public Sub DeleteValfAdasi(ByRef myValfAdasi As ValfAdasi)
        myList.Remove(myValfAdasi)
    End Sub
    Public Sub DeleteValf(ByRef myValfAdasi As ValfAdasi, ByRef myValf As Valf)
        If firstTree Then
            My.Computer.FileSystem.DeleteFile(partsFolderPath & myValf.Path & ".jpg")
        Else
            My.Computer.FileSystem.DeleteFile(sensorsFolderPath & myValf.Path & ".jpg")
        End If

        myValfAdasi.ValfsList.Remove(myValf)
        'delete its image
    End Sub
    Public Sub setValfPath(ByRef myvalfadasi As ValfAdasi, ByRef myValf As Valf)
        If firstTree Then
            myValf.Path = (myvalfadasi.Name).Replace(" ", "_") + "_" + myValf.Name.Replace(" ", "_")
        Else
            myValf.Path = (myvalfadasi.Name).Replace(" ", "_") + "_" + myValf.Name.Replace(" ", "_")
        End If
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load '''''
        Me.Text = "Select & Import"
        'UpdateToolStripMenuItem.Image = Bitmap.FromFile("C:\Users\ahmad\Downloads\add.png")
        'RenameToolStripMenuItem.Image = Bitmap.FromFile("C:\Users\ahmad\Downloads\edit.png")
        'DeleteToolStripMenuItem.Image = Bitmap.FromFile("C:\Users\ahmad\Downloads\delete.png")
        'myImageList.Images.Add(Image.FromFile("C:\Users\ahmad\Downloads\root1.png"))
        'myImageList.Images.Add(Image.FromFile("C:\Users\ahmad\Downloads\root2.png"))
        'myImageList.Images.Add(Image.FromFile("C:\Users\ahmad\Downloads\valve.png"))
        'myImageList.Images.Add(Image.FromFile("C:\Users\ahmad\Downloads\component.png"))
        'myImageList.Images.Add(Image.FromFile("C:\Users\ahmad\Downloads\empty.png"))
        'TreeView1.ImageList = myImageList
        InitiatePaths() ' new 

    End Sub
    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click '''''
        Me.Close()
    End Sub
    Private Sub Form1_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing ''''
        Dim msgBoxAnswer As String
        msgBoxAnswer = MsgBox("Are you sure you want to quit?", 17, "Close").ToString
        If msgBoxAnswer = "No" Or msgBoxAnswer = "Cancel" Or msgBoxAnswer = "Abort" Or msgBoxAnswer = "Ignore" Then
            e.Cancel = True
        End If
    End Sub
    Private Sub RenameToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RenameToolStripMenuItem.Click '''
        If selectedValfAdasi = "" Then
            Dim tempName As String = InputBox("Choose a name for " + myTree.SelectedNode.Text, "Rename")
            If tempName <> "" Then
                ValfAdalariName = tempName
            End If
            '-------------------------------
        ElseIf selectedValf = "" Then
            Dim myValfAdasi As ValfAdasi = SearchForValfAdasi(selectedValfAdasi)
            Dim tempName As String = InputBox("Choose a name for " + myTree.SelectedNode.Text, "Rename")

            If (tempName <> "") And isValfAdasiValid(tempName) Then
                myValfAdasi.Name = tempName
                Try
                    updatePaths(myValfAdasi)
                Catch ex As Exception
                    MsgBox("No images set yet.", 64, "info")
                End Try

            Else
                MsgBox("Empty string or " & tempName & " already exists", 48, "Invalid Operation")
            End If
            '-------------------------------
        ElseIf selectedComponent = "" Then
            Dim myValfAdasi As ValfAdasi = SearchForValfAdasi(selectedValfAdasi)
            Dim myValf As Valf = SearchForValf(selectedValfAdasi, selectedValf)
            Dim tempName As String = InputBox("Choose a name for " + myTree.SelectedNode.Text, "Rename")

            If (tempName <> "") And isValfValid(myValfAdasi.Name, tempName) Then
                Try
                    updateImageName(myValfAdasi, myValf, tempName)
                Catch ex As Exception
                    MsgBox("Not image set to path yet.", 64, "info")
                End Try

                'rename 
                'My.Computer.FileSystem.RenameFile("C:\Users\ahmad\Documents\ahmad.jpg", "a.jpg")
                myValf.Name = tempName
                setValfPath(myValfAdasi, myValf)
            Else
                MsgBox("Empty string or " & tempName & " already exists", 48, "Invalid Operation")
            End If
        End If
        UpdateTreeView(0)
    End Sub
    Private Sub UpdateToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UpdateToolStripMenuItem.Click ''''
        If (firstTree And redNames > 0) Or (Not firstTree And redNames2 > 0) Then
            MsgBox("Resolve conflicts first!", 48, "Invalid Operation")
        Else
            If selectedValf = "" Then
                add()
            Else
                Dim question As String = MsgBox("Delete parts and start over?", 33, "Add").ToString
                If question = "Ok" Then
                    UpdateValf()
                End If
            End If
        End If
    End Sub
    Public Sub InitiatePaths() 'new
        Dim fd As OpenFileDialog = New OpenFileDialog()


        fd.Title = "Choose Excel File"
        fd.Filter = "All files (*.*)|*.*|All files (*.*)|*.*"
        fd.FilterIndex = 2
        fd.RestoreDirectory = True

        If fd.ShowDialog() = DialogResult.OK Then

            Dim folders() As String = fd.FileName.Split("\")
            For i As Integer = 0 To folders.Length - 2
                folderPath += folders(i) & "\"

            Next
            excelFileName = folders(folders.Length - 1)
            If Not System.IO.Directory.Exists(folderPath & "Parts") Then
                System.IO.Directory.CreateDirectory(folderPath & "Parts")
            End If
            If Not System.IO.Directory.Exists(folderPath & "Sensors") Then
                System.IO.Directory.CreateDirectory(folderPath & "Sensors")
            End If
            partsFolderPath = folderPath & "Parts\"
            sensorsFolderPath = folderPath & "Sensors\"
        Else
            Me.Close()
        End If



    End Sub
    Public Sub showScreenshot(ByRef myvalf As Valf)
        If selectedValfAdasi = "" Then
        ElseIf selectedValf = "" Then
        ElseIf selectedComponent = "" Then
            Dim curNode As TreeNode
            Dim firstNode As TreeNode = myTree.TopNode
            For Each secondLevel As TreeNode In firstNode.Nodes
                If secondLevel.Text = selectedValfAdasi Then
                    For Each thirdLevel As TreeNode In secondLevel.Nodes
                        If thirdLevel.Text = selectedValf Then
                            curNode = thirdLevel
                        End If
                    Next
                End If
            Next
            Dim maxX As Integer = curNode.Bounds.Width + curNode.Bounds.X
            curNode.Expand()
            For Each n As TreeNode In curNode.Nodes
                If maxX < (n.Bounds.X + n.Bounds.Width) Then
                    maxX = n.Bounds.X + n.Bounds.Width
                End If
            Next
            Dim newLocation As New System.Drawing.Point(maxX + 5, curNode.Bounds.Y + 5)
            Dim mypb As PictureBox
            Try
                If firstTree Then
                    mypb = PictureBox1
                    mypb.Image = SaveImageFromFile(partsFolderPath & myvalf.Path & ".jpg")
                Else
                    mypb = PictureBox2
                    mypb.Image = SaveImageFromFile(sensorsFolderPath & myvalf.Path & ".jpg")
                End If
            Catch ex As Exception
                MsgBox("No images set yet.", 64, "info")
                Exit Sub
            End Try

            mypb.Height = curNode.Bounds.Height * 7
                mypb.Width = mypb.Height * 4.0 / 3.0
                mypb.Location = newLocation
                mypb.Visible = True

        End If
    End Sub
    'Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
    'Save Image
    ' Dim newImage As Bitmap = PictureBox1.Image
    'newImage.Save("C:\Users\ahmad\Documents\ahmad.jpg", System.Drawing.Imaging.ImageFormat.Jpeg)
    'End Sub
    Private Sub ShowScreenshotToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ShowScreenshotToolStripMenuItem.Click
        showScreenshot(SearchForValf(selectedValfAdasi, selectedValf))
    End Sub
    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click ''''''

        Process.Start(partsFolderPath & SearchForValf(selectedValfAdasi, selectedValf).Path & ".jpg")
    End Sub
    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click ''''''
        Process.Start(sensorsFolderPath & SearchForValf(selectedValfAdasi, selectedValf).Path & ".jpg")
    End Sub
    Private Sub TreeView1_Click(sender As Object, e As EventArgs) Handles TreeView1.Click ''''''
        firstTree = True
        myTree = TreeView1
        myList = ValfAdasisList
        If PictureBox1.Visible = True Then
            PictureBox1.Visible = False
        End If
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click ''''
        'cells .value.tostring
        'Dim xls As New Excel.Application
        'Dim book As Excel.Workbook
        'Dim sheet As Excel.Worksheet
        'xls.Workbooks.Open("C:\Users\ahmad\Documents\test2.xlsx")
        'get references to first workbook and worksheet
        'xls.Visible = True
        'book = xls.ActiveWorkbook
        'sheet = book.ActiveSheet
        'Dim rowCount As Integer = 1
        'While sheet.Cells(rowCount, 1).Value IsNot Nothing
        'rowCount += 1
        'End While
        'Dim firstAvailableRow As Integer = rowCount
        'Dim columnCount As Integer = 1
        'While sheet.Cells(1, columnCount).Value IsNot Nothing
        'columnCount += 1
        'End While
        'Dim firstAvailableColumn As Integer = columnCount
        'MsgBox("columnCount: " & columnCount)  'or whatever value you want
        'save the workbook and clean up
        'releaseObject(sheet)
        'releaseObject(book)
        'releaseObject(xls)

    End Sub
    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub
    Private Sub TreeView2_NodeMouseClick(sender As Object, e As TreeNodeMouseClickEventArgs) Handles TreeView2.NodeMouseClick
        If e.Button = System.Windows.Forms.MouseButtons.Right Then
            TreeView2.SelectedNode = e.Node
            Dim fullPath As String
            Dim slashCount As Integer = 0
            fullPath = TreeView2.SelectedNode.FullPath
            slashCount = CountCharacter(fullPath, "\")
            Select Case (slashCount)
                Case 0
                    selectedValfAdasi = ""
                    selectedValf = ""
                    selectedComponent = ""
                Case 1
                    selectedValfAdasi = TreeView2.SelectedNode.Text
                    selectedValf = ""
                    selectedComponent = ""
                Case 2
                    selectedValfAdasi = TreeView2.SelectedNode.Parent.Text
                    selectedValf = TreeView2.SelectedNode.Text
                    selectedComponent = ""
                Case 3
                    selectedValfAdasi = TreeView2.SelectedNode.Parent.Parent.Text
                    selectedValf = TreeView2.SelectedNode.Parent.Text
                    selectedComponent = TreeView2.SelectedNode.Text
            End Select
            UpdateStripMenuNames()
        End If
        firstTree = False
        myTree = TreeView2
        myList = ValfAdasisList2
        If PictureBox2.Visible = True Then
            PictureBox2.Visible = False
        End If
    End Sub
    Private Sub TreeView2_Click(sender As Object, e As EventArgs) Handles TreeView2.Click
        firstTree = False
        myTree = TreeView2
        myList = ValfAdasisList2
        If PictureBox2.Visible = True Then
            PictureBox2.Visible = False
        End If
    End Sub
    'delete file
    'My.Computer.FileSystem.DeleteFile("C:\Users\ahmad\Documents\avatar.jpg")
    Public Sub updateImageName(ByRef myvalfadasi As ValfAdasi, ByRef myvalf As Valf, ByVal newname As String)
        If firstTree Then
            My.Computer.FileSystem.RenameFile(partsFolderPath & myvalf.Path & ".jpg", (myvalfadasi.Name.Replace(" ", "_") & "_" + newname.Replace(" ", "_") & ".jpg"))
        Else
            My.Computer.FileSystem.RenameFile(sensorsFolderPath & myvalf.Path & ".jpg", (myvalfadasi.Name.Replace(" ", "_") & "_" + newname.Replace(" ", "_") & ".jpg"))

        End If
    End Sub
    Public Sub updatePaths(ByRef myvalfadasi As ValfAdasi)
        For Each myvalf As Valf In myvalfadasi.ValfsList
            Dim oldPath As String = myvalf.Path

            setValfPath(myvalfadasi, myvalf)

            If firstTree Then
                My.Computer.FileSystem.RenameFile(partsFolderPath & oldPath & ".jpg", myvalf.Path & ".jpg")

            Else
                My.Computer.FileSystem.RenameFile(sensorsFolderPath & oldPath & ".jpg", myvalf.Path & ".jpg")

            End If
        Next
    End Sub
    Public Sub SaveScreenShot(ByVal myvalf As Valf)
        SendKeys.Send("{PRTSC}")

        Dim Screenshot As Image = Clipboard.GetImage()
        If firstTree Then
            Screenshot.Save(partsFolderPath & myvalf.Path & ".jpg", System.Drawing.Imaging.ImageFormat.Jpeg)

        Else
            Screenshot.Save(sensorsFolderPath & myvalf.Path & ".jpg", System.Drawing.Imaging.ImageFormat.Jpeg)

        End If
    End Sub

    Private Sub TreeView2_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TreeView2.AfterSelect

    End Sub
    Public Shared Function SaveImageFromFile(path As String) As Image
        Using fs As New FileStream(path, FileMode.Open, FileAccess.Read)
            Dim img = Image.FromStream(fs)
            Return img
        End Using
    End Function
End Class

Public Class ValfAdasi
    Public Property ValfsList As New List(Of Valf)
    Public Property Name As String
End Class
Public Class Valf
    Public Property ComponentsList As New List(Of Component)
    Public Property SensorsList As New List(Of Sensor) ' new 
    Public Property Name As String
    Public Property Path As String
End Class
Public Class Component
    Public Property Name As String
    Public Property ExportName As String
    Public Property GroupName As String
    Public Property MakeRed As Boolean
End Class
Public Class Sensor
    Public Property Name As String
    Public Property ExportName As String
    Public Property GroupName As String
    Public Property MakeRed As Boolean
    Public Property Brand As String
End Class