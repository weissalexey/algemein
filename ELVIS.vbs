'FolderTo ="c:\Users\aw\Desktop\winspied\ELVIS\POD_BCK\"
'PatchPDF ="c:\Users\aw\Desktop\winspied\ELVIS\"

FolderTo ="E:\Schnittstellen\Elvis\OUT\POD_BCK\"
PatchPDF ="E:\Schnittstellen\Elvis\OUT\POD\"

sql = "select * from V_POD_ELVIS_941"
WriteLog "Start"
Set con = CreateObject("ADODB.Connection")
	With con
			.Provider = "SQLOLEDB"
			.Properties("Data Source") = Kennwort("Data_Source")
			.ConnectionString = Kennwort("ConnectionString")
			.Open
			.DefaultDatabase = Kennwort("DefaultDatabase")
End With
   
   Set result = con.Execute(sql)
   

    If Not result.EOF  Then
          Var_nul=0 
	  result.MoveFirst
	  While Not result.EOF
	  	dim LN, naim, FE, DOCDAT
	  	  SWort = result.Fields("SWort").Value 
	  	  ZusText_E = result.Fields("ZusText_E").Value
	writelog ZusText_E 
        if ZusText_E = "" then 
        Var_nul = Var_nul +1
	end if
	'ZusText_E = replace (ZusText_E, "/", Var_nul)
	'ZusText_E = replace (ZusText_E, "+", Var_nul)
        'ZusText_E = replace (ZusText_E, "\", Var_nul)
        'ZusText_E = replace (ZusText_E, "#", Var_nul)
	

        DOCDAT = result.Fields("DocumentData").Value
        ArcDocINr = result.Fields("ArcDocINr").Value
        'MsgBox ArcDocINr
        'SaveBinaryData  PatchPDF &"POD_" & SWort & "_" & ZusText_E & "_1.PDF", DOCDAT
	 SaveBinaryData  PatchPDF &"POD_" & SWort  & "_1.PDF", DOCDAT
	    'WriteLog PatchPDF &"POD_" & SWort & "_" & ZusText_E & "_1.PDF"
		WriteLog PatchPDF &"POD_" & SWort & "_1.PDF"
    
	   result.movenext
     UpdateSQlDatai ArcDocINr
	  wend
      FTPUpload(PatchPDF) 

	end if
 
writelog " End"

Function SaveBinaryData(FileName, ByteArray)
  Const adTypeBinary = 1
  Const adSaveCreateOverWrite = 2
  
  'Create Stream object
  Dim BinaryStream
  Set BinaryStream = CreateObject("ADODB.Stream")
  
  'Specify stream type - we want To save binary data.
  BinaryStream.Type = adTypeBinary
  
  'Open the stream And write binary data To the object
  BinaryStream.Open
  BinaryStream.Write ByteArray
  
  'Save binary data To disk
  
  BinaryStream.SaveToFile FileName, adSaveCreateOverWrite
End Function

Sub WriteLog(LogMessage)
log_file = "E:\Schnittstellen\log\"
Const ForAppending = 8
A = year(Now)
B=Month(Now)
if len(B) < 2  then b="0" & B end if
C = day (Now)
if len(C)<2 then C ="0" & C end if
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objLogFile = objFSO.OpenTextFile(log_file & A & B & C & "schaden_ELVIS.log" , ForAppending, TRUE)
objLogFile.WriteLine("[" & Now() & "_ELVIS] " & LogMessage)
End Sub

'

Sub FTPUpload(path)
Set oShell = CreateObject("Shell.Application")
Set objFSO = CreateObject("Scripting.FileSystemObject")
On Error Resume Next

Const copyType = 16

'FTP Wait Time in ms
waitTime = 8000

FTPUser = Kennwort("ELVISFTPUser")
FTPPass = Kennwort("ELVISFTPPass")
FTPHost = Kennwort("ELVISPHost")
FTPDir = Kennwort("ELVISOUtDir")

strFTP = "ftp://" & FTPUser & ":" & FTPPass & "@" & FTPHost & FTPDir
'writelog strFTP 
Set objFTP = oShell.NameSpace(strFTP)


'Upload single file       
If objFSO.FileExists(path) Then

Set objFile = objFSO.getFile(path)
strParent = objFile.ParentFolder
Set objFolder = oShell.NameSpace(strParent)

Set objItem = objFolder.ParseName(objFile.Name)

'writelog "Uploading file " & objItem.Name & " to " & strFTP
 objFTP.CopyHere objItem, copyType


End If


'Upload all files in folder
If objFSO.FolderExists(path) Then

'Entire folder
Set objFolder = oShell.NameSpace(path)

'writelog "Uploading folder " & path & " to " & strFTP
objFTP.CopyHere objFolder.Items, copyType
BackUpFile PatchPDF, FolderTo
End If


If Err.Number <> 0 Then
WriteLog "Error: " & Err.Description
End If

'Wait for upload
Wscript.Sleep waitTime

End Sub


sub BackUpFile (PatchPDF,FolderTo )
A = year(Now)
B=Month(Now)
if len(B) < 2  then b="0" & B end if
C = day (Now)
if len(C)<2 then C ="0" & C end if

Set FSO=CreateObject("Scripting.FileSystemObject")
Set fldr= FSO.GetFolder(PatchPDF)
Set Collec_Files= fldr.Files
set fso1=CreateObject ("Scripting.FileSystemObject") 
For Each File in Collec_Files
    If Collec_Files.count = 0 then
      Writelog PatchPDF & " ist leer "
    Else
      Year_yy = year(Now)
      Month_MM=Month(Now)
      if len(Month_MM) < 2  then Month_MM = "0" & Month_MM 
      Day_dd = day (Now)
      if len(Day_dd)<2 then Day_dd ="0" & Day_dd 
      Hour_HH= (Hour(Now))
      Minute_mm = (Minute(Now))
      Datatime_YY = Year_yy & Month_MM & Day_dd & Hour_HH & Minute_mm 
      writelog PatchPDF & File.Name & " " & FolderTo & Datatime_YY & "_"& File.Name
      'set new_folder=fso1.CreateFolder(FolderTo & A & B & C & "\")
       FSO.MoveFile PatchPDF & File.Name, FolderTo & Datatime_YY & "_"& File.Name
    End If
Next


End sub

sub UpdateSQlDatai (ArcDocINr)

    Set conn = CreateObject("ADODB.Connection")
    
	With conn
			.Provider = "SQLOLEDB"
			.Properties("Data Source") = Kennwort("Data_Source")
			.ConnectionString = Kennwort("ConnectionString")
			.Open
			.DefaultDatabase = Kennwort("DefaultDatabase")
	End With
   
   sqlAnfr = "UPDATE xxaArcDoc SET Exportiert = 1 WHERE ArcDocINr = " & ArcDocINr
   WriteLog sqlAnfr
   conn.Execute(sqlAnfr)
   
end Sub

Function Kennwort (PArramm)
'Kennwort= ""
filename = "\\srv-dc2\DATEN$\IT\script\Key\key.zip"
set goFS = CreateObject("Scripting.FileSystemObject")
set ts = goFS.OpenTextFile( filename)
'MsgBox par
do until ts.AtEndOfStream

    strl = ts.ReadLine

    if left (strl, len(PArramm)) = PArramm then Kennwort = right (strl, len(strl) - len(PArramm) - 3) 
   'writelog  Kennwort
loop
End Function