' Date Created: January 17, 2019
' Author: Ronknight
' Script: search-move-files.vbs
' Description: Search all files that has a filename as the keyword from a source folder and move files to the new destination folder

dim objFSO, SourceFolder, DestFolder, Keyword, i

'Path
    'change sourcepath example "C:\Users\Ron\Downloads\"
    SourceFolder = "Enter new path here"
    'change sourcepath example "C:\Users\Ron\Desktop\Backup\"
    DestFolder = "Enter new path here"

'Variables
    'Enter search keyword here, example "wc-product-export" or ".xls" or ".jpg"
    Keyword = "Enter keyword here"
    i = 0

'Objects
set objFSO = CreateObject("scripting.Filesystemobject") 
set objFolder = objFSO.GetFolder(SourceFolder)  

'Execute Main Function
	Main objFSO.GetFolder(SourceFolder)
	'Prompt end of Loop
	Wscript.Echo "Exiting program..."

Sub Main(objFolder)
	'List all Subfolder form Source Directory
        For Each File in objFolder.Files
            x = objFSO.GetBasename(File)
            'Search Keyword
                If instr(lcase(x), Keyword) > 0 then
                    i = i+1
                    'Delete file if it already exists
                    If objFSO.Fileexists(DestFolder & "\" & File.name) then
                        objFSO.deleteFile DestFolder & "\" & File.name, true
                    End If
                'Move file to new location
                Wscript.Echo "File moved: " & File.Name
                objFSO.MoveFile SourceFolder & "\" & File.name, DestFolder
                 End If
        Next

        If i>0 then
            'Prompt end of process and show destination folder
            Wscript.Echo "Moving of files completed..."
            Wscript.Echo i &" Files moved to : " & DestFolder
            Wscript.quit()
        End If
    'Prompt if keyword not found
    Wscript.Echo "No match found..."
End Sub