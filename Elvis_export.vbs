Set objArgs = WScript.Arguments 

if  objArgs.Count > 0 then 
param = objArgs(0) 

FolderTo = objArgs(1)
writelog "Beginen"
FTPUpload param

BackUpFile param, FolderTo
writelog "End"
end if





Sub WriteLog(LogMessage)
log_file = "E:\Schnittstellen\log\"
Const ForAppending = 8
A = year(Now)
B=Month(Now)
if len(B) < 2  then b="0" & B end if
C = day (Now)
if len(C)<2 then C ="0" & C end if
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objLogFile = objFSO.OpenTextFile(log_file & A & B & C & "Elvis_export_Status.log" , ForAppending, TRUE)
objLogFile.WriteLine("[" & Now() & "_ELVIS_export] " & LogMessage)
End Sub

'

Sub FTPUpload(param )
Set oShell = CreateObject("Shell.Application")
Set objFSO = CreateObject("Scripting.FileSystemObject")
On Error Resume Next

Const copyType = 16

'FTP Wait Time in ms
waitTime = 8000

FTPUser = Kennwort("ELVISFTPUser")
FTPPass = Kennwort("ELVISFTPPass")
FTPHost = Kennwort("ELVISPHost")
FTPDir = kennwort("ELVISFTPSTATUS")


strFTP = "ftp://" & FTPUser & ":" & FTPPass & "@" & FTPHost & FTPDir
'writelog strFTP 
Set objFTP = oShell.NameSpace(strFTP)

'writelog strFTP
'Upload single file       
If objFSO.FileExists(param ) Then

Set objFile = objFSO.getFile(param )
strParent = objFile.ParentFolder
Set objFolder = oShell.NameSpace(strParent)

Set objItem = objFolder.ParseName(objFile.Name)

'writelog "Uploading file " & objItem.Name & " to " & strFTP
 objFTP.CopyHere objItem, copyType


End If


'Upload all files in folder
If objFSO.FolderExists(param ) Then

'Entire folder
Set objFolder = oShell.NameSpace(param )

'writelog "Uploading folder " & param  & " to " & strFTP
objFTP.CopyHere objFolder.Items, copyType
BackUpFile param , FolderTo
End If


If Err.Number <> 0 Then
WriteLog "Error: " & Err.Description
End If

'Wait for upload
Wscript.Sleep waitTime

End Sub


sub BackUpFile (param ,FolderTo )
A = year(Now)
B=Month(Now)
if len(B) < 2  then b="0" & B end if
C = day (Now)
if len(C)<2 then C ="0" & C end if

Set FSO=CreateObject("Scripting.FileSystemObject")
Set fldr= FSO.GetFolder(param)
Set Collec_Files= fldr.Files
set fso1=CreateObject ("Scripting.FileSystemObject") 
For Each File in Collec_Files
   If Collec_Files.count = 0 then
      Writelog Patch & " ist leer "
    Else
       
      Year_yy = year(Now)
      Month_MM=Month(Now)
      if len(Month_MM) < 2  then Month_MM = "0" & Month_MM 
      Day_dd = day (Now)
      if len(Day_dd)<2 then Day_dd ="0" & Day_dd 
      Hour_HH= (Hour(Now))
      Minute_mm = (Minute(Now))
      Datatime_YY = Year_yy & Month_MM & Day_dd & Hour_HH & Minute_mm 
       Writelog param & File.Name & " " & FolderTo & Datatime_YY & "_"& File.Name
      FSO.MoveFile param & File.Name , FolderTo & Datatime_YY & "_"& File.Name
    End If
Next


End sub



Function Kennwort (PArramm)
'Kennwort= ""
filename = "\\srv-dc2\DATEN$\IT\script\Key\key.zip"
set goFS = CreateObject("Scripting.FileSystemObject")
set ts = goFS.OpenTextFile( filename)
'MsgBox par
do until ts.AtEndOfStream

    strl = ts.ReadLine
    if left (strl, len(PArramm)) = PArramm then Kennwort= right (strl, len(strl) - len(PArramm) - 3) 
   'MsgBox Kennwort
loop
End Function