Option Explicit 'force all variables to be declared

Const folderName = "E:\FreeTool.git\VisioExportPNG"
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

Dim objFolder, objFile
Set objFolder = objFSO.GetFolder(folderName) 

For Each objFile In objFolder.Files
        If LCase(objFSO.GetExtensionName(objFile.Name)) = "vsdx" Then
			ConvertVisioToPNG(objfile.Path)
        End If
    Next
	

Function ConvertVisioToPNG(VisioFile)  'This function is to convert a Visio file to PNG file 
    Dim objshell,ParentFolder,BaseName,Visioapp,Visio 
  
    Set Visioapp = CreateObject("Visio.Application") 
    Visioapp.Visible = False 
    Set Visio = Visioapp.Documents.Open(VisioFile) 
    Set Pages = Visioapp.ActiveDocument.Pages 
     
    Set objshell= CreateObject("scripting.filesystemobject") 
    ParentFolder = objshell.GetParentFolderName(VisioFile) 'Get the current folder path 
    BaseName = objshell.GetBaseName(VisioFile) 'Get the file name 
     
    Dim PageName,Page,Pages 
    For Each Page In Pages 
        PageName = Page.Name 
        Page.Export(parentFolder & "\" & BaseName & "-" & PageName & ".png") 
    Next 
     
    Visio.Close 
    Visioapp.Quit 
    Set objshell = Nothing  
End Function 