Option Strict Off
Imports System
Imports NXOpen
Imports NXOpenUI
Imports System.Windows.Forms
Imports System.ComponentModel
Imports System.Windows.Media.Color.Red
Imports System.Collections.Generic
Imports System.Linq
Imports NXOpen.UF
Imports NXOpen.Assemblies
Imports Microsoft.VisualBasic
Imports System.Drawing
Imports System.IO
Imports System.Globalization
Imports System.Diagnostics
Imports Microsoft.VisualBasic.FileIO
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel

Module NX
	Dim theSession As Session = Session.GetSession()
	public tempCount as integer = 0
	Public theUfSession As UFSession = UFSession.GetUFSession()
	Public ufs As UFSession = UFSession.GetUFSession()
	Public myResponse As Selection.Response
	Dim displayPart As Part = theSession.Parts.Display

	Dim tagList() As NXOpen.Tag
	Dim resultStr As String = ""

	
	
	Dim workPart As Part = theSession.Parts.Work
	Dim lw As ListingWindow = theSession.ListingWindow
	Public mySelectedObjects As NXObject()
	Public ValfAdasisList As New List(Of ValfAdasi)
	Public ValfAdasisList2 As New List(Of ValfAdasi)
	Public FittingInfoList As New List(Of FittingInfo)
	Public firstTree As Boolean = True
	Public partsFolderPath As String = ""
	Public sensorsFolderPath As String = ""
	Public folderPath As String = ""
	Public subFolder As String = ""
	Public excelFileName As String = ""
	Public fileNameWithoutExtension As String = ""
	Public tempDirectory = System.IO.Path.GetTempPath()


	Sub Main()
		
    	Application.Run(New Form1())
	End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Public Function SelectObjects(ByRef selobj() As NXObject) As Selection.Response
		
	End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetManufacturer(anObject As Object) As String

End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetEOTGRNO(anObject As Object) As String
	
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Public Sub clearSelectedObjects()
		
		
	End Sub
'**************************************************************************************************
'**************************************************************************************************
Public Function GetMaxValfAdasi() As Integer
	
End Function


'**************************************************************************************************
'**************************************************************************************************
Public Sub CopyAndPasteCells(max As Integer)
        
    End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ChangeIDtoVxxID()
	
	
End Sub
''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetValvesRowsAndExportToExcel()
	
End Sub
''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetSensorsRowsAndExportToExcel()
	
End Sub
''''''''''''''''''''''''''''''''''''''''''''''
		Public Sub exportToExcel()
        
    End Sub
'**************************************************************************************************
'**************************************************************************************************
	Public Function FindStars(currentRow As Integer,objExce as object ) As Integer
		
	End Function
	'**************************************************************************************************
'**************************************************************************************************
	Public Function FindNOs(currentRow As Integer,objExce as object ) As Integer
		
	End Function
'**************************************************************************************************
'**************************************************************************************************
	Public Sub releaseObject(ByVal obj As Object)
		
	End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub InitiateSaving(byref myvalf as valf)
		
		
    End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	public Sub ExportScreenshot(ByVal filename As String)

       
	End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Public Sub DeleteFile(ByVal filePath As String)

        
		

    End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function GetUnloadOption(ByVal dummy As String) As Integer

        
    End Function
	
	
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Module 


Public Class Form1
	Dim selectedValfAdasi As String
	Dim selectedValf As String
	Dim selectedComponent As String
	Dim ValfAdalariName As String = "Valf Adalari"
	Dim myImageList As New ImageList()

	Public redNames As Integer = 0
	Public redNames2 As Integer = 0
	Public myTree As TreeView
	Public ID As Integer = 0

	Public myList As New List(Of ValfAdasi)

	Private Sub UpdateStripMenuNames()
    
	End Sub
	Public Sub countRedNames()
    	
	End Sub
	Public Sub InitiateTemp()
		

	End Sub
	Public Sub GetExcelPath() 'new

    	End If



	End Sub
	Private Sub UpdateTreeView(m As Integer)
    
	End Sub
	Private Sub DeleteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteToolStripMenuItem.Click '''
    	
	End Sub
	Private Sub TreeView1_NodeMouseClick(sender As Object, e As System.Windows.Forms.TreeNodeMouseClickEventArgs) Handles TreeView1.NodeMouseClick
 
	End Sub
	Public Sub delete()
    
	End Sub
	Public Sub UnMakeRed()
  
	End Sub
	Public Sub UpdateValf()
    	
	End Sub
	Public Sub add_component()

    	End Sub
	Function GetFitting(myComponent As Component) as string
		

	End Function
	Public Sub add_sensor()
       
    	End Sub
	Private Sub Highlight()

	End Sub
	Public Sub AddValfAdasi()

	End Sub
	Public Function SearchForValfAdasi(ValfAdasisName As String) As ValfAdasi

	End Function
	Public Function SearchForValf(ValfAdasisName As String, ValfsName As String) As Valf

	End Function
	Public Function SearchForSensor(ComponentsName As String) As Sensor ' new

	End Function
	Public Function SearchForComponent(ComponentsName As String) As Component

	End Function
	Public Sub AddValf(ByRef myValfAdasi As ValfAdasi)
    	
	End Sub
	Public Sub AddSensor(ByRef myValf As Valf, ByVal ComponentName As String) ' new

	End Sub
	Public Sub AddComponent(ByRef myValf As Valf, ByVal ComponentName As String)

	End Sub
	Public Function isComponentNameinValf(ByRef myValf, componentName) As Boolean

	End Function
	Public Function duplicateSensor(ByRef myValfAdasi As ValfAdasi, ByRef myValf As Valf, componentName As String) As Sensor ' new
    	
	End Function
	Public Function duplicateComponent(componentName) As Component
    	
	End Function
	Public Function duplicateSensor(componentName) As Sensor ' new

	End Function
	Public Function duplicateComponent(ByRef myValfAdasi As ValfAdasi, ByRef myValf As Valf, componentName As String) As Component
    	
	End Function
	Public Function CountCharacter(ByVal value As String, ByVal ch As Char) As Integer

	End Function
	Public Function isValfAdasiValid(valfAdasiName As String) As Boolean

	End Function
	Public Function isValfValid(valfAdasiName As String, valfName As String) As Boolean
 
	End Function
	Public Sub DeleteValfAdasi(ByRef myValfAdasi As ValfAdasi)
    
	End Sub
	Public Sub DeleteValf(ByRef myValfAdasi As ValfAdasi, ByRef myValf As Valf)

	End Sub
	Public Sub setValfPath(ByRef myvalfadasi As ValfAdasi, ByRef myValf As Valf)

	End Sub
	Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load '''''

	End Sub
	Sub InitiateFitting()
	
		
	End Sub
	Sub PopulateFittingInfoClass(path as string, rowCount as integer)

	End Sub
	Function GetSheetRowCount(path as String) 
    
	End Function
	Function GetFittingFile() as string
		
	End Function
	Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click '''''
    	
	End Sub
	Private Sub Form1_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing ''''
 
	End Sub
	Private Sub RenameToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RenameToolStripMenuItem.Click '''
    
	End Sub
	Private Sub UpdateToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UpdateToolStripMenuItem.Click ''''
    
	End Sub
	Public Sub OpenImageFile()
		
	End Sub
	Public Sub InitiateFilesAndFolders()

	End Sub
	
	Public Sub showScreenshot(ByRef myvalf As Valf)
 
    	End Sub
	Private Sub ShowScreenshotToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ShowScreenshotToolStripMenuItem.Click
    
	End Sub
	Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click 
		
	End Sub
	Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click 
		
	End Sub
	Private Sub TreeView1_Click(sender As Object, e As EventArgs) Handles TreeView1.Click 
    	
	End Sub
	Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click 

	End Sub
	Private Sub releaseObject(ByVal obj As Object)
    	
	End Sub

	Private Sub TreeView2_NodeMouseClick(sender As Object, e As TreeNodeMouseClickEventArgs) Handles TreeView2.NodeMouseClick
 
	End Sub
	Private Sub TreeView2_Click(sender As Object, e As EventArgs) Handles TreeView2.Click

	End Sub

	Public Sub updateImageName(ByRef myvalfadasi As ValfAdasi, ByRef myvalf As Valf, ByVal newname As String)
    	
	End Sub
	Public Sub updatePaths(ByRef myvalfadasi As ValfAdasi)
    	
	End Sub
	Public Sub SaveScreenShot(ByVal myvalf As Valf)

	End Sub
	Private Sub TreeView2_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TreeView2.AfterSelect

	End Sub
	Public Shared Function SaveImageFromFile(path As String) As Image
    	
	End Function
End Class


Public Class ValfAdasi
	Public Property ValfsList As New List(Of Valf)
	Public Property Name As String
End Class
Public Class Valf
	Public Property ComponentsList As New List(Of Component)
	Public Property SensorsList As New List(Of Sensor)
	Public Property Name As String
	Public Property Path As String
End Class
Public Class Component
	Public Property Name As String
	Public Property ExportName As String
	Public Property GroupName As String
	Public Property MakeRed As Boolean
	Public Property EOTGRNO As String 
	Public Property ID As String 
	Public Property Brand As String
	Public Property Fitting As String
End Class
Public Class Sensor
	Public Property Name As String
	Public Property ExportName As String
	Public Property GroupName As String
	Public Property MakeRed As Boolean
	Public Property Brand As String
End Class
	


Public Class FittingInfo
	Public Property Brand As String
	Public Property Content As String()
	Public Property ContentNum As String
	Public Property Fitting As String
End Class

End Class


