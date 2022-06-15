Set objArgs = WScript.Arguments 

if  objArgs.Count > 0 then 
param = objArgs(0) 

DEr_Alrer = CInt(objArgs(1))

if DEr_Alrer=0 then DEr_Alrer=2

WriteLog "löschen wir alte Dateien aus " & param & " älter als "& DEr_Alrer & " Tage"
Set FSO = CreateObject("Scripting.FileSystemObject")
sDir = param
Set objDir = GetFolder(sDir)
DeleteOlderFiles objDir, DEr_Alrer 
writelog "Es fertig"
end if


Function GetFile(sFile)
 On Error Resume Next
 
 Set FSO = CreateObject("Scripting.FileSystemObject")
 Set GetFile = FSO.GetFile(sFile)
 if err.number <> 0 then
    writelog "Error Opening file " & sFile & VBlf & "["&Err.Description&"]"
    Wscript.Quit Err.number
 end if
End Function 

' Ordner erhalten
Function GetFolder (sFolder)
 On Error Resume Next
 
 Set GetFolder = FSO.GetFolder(sFolder)
 if err.number <> 0 then
    writelog "Error Opening folder " & sFolder & VBlf & "["&Err.Description&"]"
    Wscript.Quit Err.number
 end if
End Function 

'eine Datei löschen (Dateiname wird an sFile übergeben)

Sub DeleteFile(sFile)
 On Error Resume Next

 FSO.DeleteFile sFile, True
 if err.number <> 0 then
    WScript.Echo "Error Deleteing file " & sFile & VBlf & "["&Err.Description&"]"
    Wscript.Quit Err.number
 end if
End Sub 

'Löschen Sie Dateien, die älter als DEr_Alrer Tage sind
Sub DeleteOlderFiles(objDir, DEr_Alrer )
    ' alle Dateien in einem Verzeichnis anzeigen
    for each efile in objDir.Files        
         ' DateLastModified und nicht DateCreated verwenden, weil
         ' DateCreated gibt nicht immer das richtige Datum zurück

       FileDate = efile.DateLastModified 
        Age = DateDiff("d",Now,FileDate)        
        ' in diesem Fall beträgt das Alter der Datei nicht mehr als DEr_Alrer Tage
                
        If Abs(Age)> DEr_Alrer Then
	    writelog efile
            DeleteFile(efile)
        End If        
    next 
End Sub

Sub WriteLog(LogMessage)
log_file = "E:\Schnittstellen\log\"
Const ForAppending = 8
A = year(Now)
B=Month(Now)
if len(B) < 2  then b="0" & B end if
C = day (Now)
if len(C)<2 then C ="0" & C end if
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objLogFile = objFSO.OpenTextFile(log_file & A & B & C & "delete_alte_fiele.log" , ForAppending, TRUE)
objLogFile.WriteLine("[" & Now() & "_delete_alte_fiele] " & LogMessage)
End Sub