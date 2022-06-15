pathcXML ="E:\Schnittstellen\VTL\Handewitt\Fulda\Export\Schadensfotos\XML\"
PatchJpg ="E:\Schnittstellen\VTL\Handewitt\Fulda\Export\Schadensfotos\JPG\"
pathcarj ="E:\Schnittstellen\VTL\Handewitt\Fulda\Export\Schadensfotos\BAK\"

'pathcXML ="c:\Users\aw\Desktop\winspied\Schadensfotos\xml\"
'PatchJpg ="c:\Users\aw\Desktop\winspied\Schadensfotos\xml\"
'pathcarj = "c:\Users\aw\Desktop\winspied\Schadensfotos\"

Y = year(Now)
M=Month(Now)
if len(M) < 2  then M="0" & M end if
D = day (Now)
if len(D)<2 then D ="0" & D end if
YMD = Y & M & D & "_"

WriteTEXTLog "Start"

Set con = CreateObject("ADODB.Connection")
With con
			.Provider = "SQLOLEDB"
			.Properties("Data Source") = Kennwort("Data_Source")
			.ConnectionString = Kennwort("ConnectionString")
			.Open
			.DefaultDatabase = Kennwort("DefaultDatabase")
End With
   sql = "select * from V_Schaden_VTL_941"
   Set result = con.Execute(sql)
   

    If Not result.EOF  Then
      
	  result.MoveFirst
	  While Not result.EOF
	  	dim LN, naim, FE, DOCDAT
	  	  LN = result.Fields("LiefNr").Value 
	  	  naim = result.Fields("SW_NVE").Value
          FE = result.Fields("FileExtension").Value 
	  	  DOCDAT = result.Fields("DocumentData").Value
          ArcDocINr = result.Fields("ArcDocINr").Value
               '  MsgBox DOCDAT

        SaveBinaryData PatchJpg & YMD  & ArcDocINr & ".jpg", DOCDAT
        
        WriteTEXTLog PatchJpg & YMD  & ArcDocINr & ".jpg"
        
            ITOg = _ 
            "<?xml version=""1.0"" encoding=""utf-8""?>"& chr(13) & chr(10) _
            &"<DamageReports>"& chr(13) & chr(10) _
            &"    <Sender>04245</Sender>"& chr(13) & chr(10) _
            &"    <DamageReport>"& chr(13) & chr(10) _
            &"        <Author>04245</Author>"& chr(13) & chr(10) _
            &"        <Contact>kj</Contact>"& chr(13) & chr(10) _
            &"        <Fon>0461 95707 0</Fon>"& chr(13) & chr(10) _
            &"        <Email>jm@carstensen.eu</Email>"& chr(13) & chr(10) _
            &"        <NOC>"& LN &"</NOC>"& chr(13) & chr(10) _
            &"        <Details>"& chr(13) & chr(10) _
            &"            <Detail>"& chr(13) & chr(10) _
            &"                <NVE>"& naim &"</NVE>"& chr(13) & chr(10) _
            &"                <Description>Siehe passende NVE Statusmeldungen</Description>"& chr(13) & chr(10) _
            &"                <Documents>"& chr(13) & chr(10) _
            &"                    <Document>"& chr(13) & chr(10) _
            &"                        <File>" 
        
        
        FileNameIN = YMD  & ArcDocINr& ".jpg"
        FileNameOUT = YMD  & ArcDocINr& ".xml"
        base64_Entcod PatchJpg&FileNameIN, pathcXML&FileNameOUT, ITOg
        
 	    
        

        
        
	     ITOg ="</File>"& chr(13) & chr(10) _
            &"                        <FileType>"& FE &"</FileType>"& chr(13) & chr(10) _
            &"                    </Document>"& chr(13) & chr(10) _
            &"                </Documents>"& chr(13) & chr(10) _
            &"            </Detail>"& chr(13) & chr(10) _
            &"        </Details>"& chr(13) & chr(10) _
            &"    </DamageReport>"& chr(13) & chr(10) _
            &"</DamageReports>"& chr(13) & chr(10) 

            
  
	    writelog ITOg, pathcXML&FileNameOUT
        WriteTEXTLog pathcXML&FileNameOUT
	    DeleteFile PatchJpg & FileNameIN 
        UpdateSQlDatai ArcDocINr
	    result.movenext
	  wend
FTPUpload pathcXML
BackUpFile pathcXML, pathcarj
	end if
WriteTEXTLog  "END"

sub WriteLog(  logstr , FileNameOUT )

Const ForAppending = 8
A = year(Now)
B=Month(Now)
if len(B) < 2  then b="0" & B end if
C = day (Now)
if len(C)<2 then C ="0" & C end if
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objLogFile = objFSO.OpenTextFile( FileNameOUT, ForAppending, TRUE)
objLogFile.Write(logstr)
'
end Sub

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

Function base64_Entcod(FileNameIN, FileNameOUT, ITOg)
  
  'Option Explicit

Const fsDoOverwrite     = true  ' Overwrite file with base64 code
Const fsAsASCII         = false ' Create base64 code file as ASCII file
Const adTypeBinary      = 1     ' Binary file is encoded

' Variables for writing base64 code to file
Dim objFSO
Dim objFileOut

' Variables for encoding
Dim objXML
Dim objDocElem

' Variable for reading binary picture
Dim objStream
WriteTEXTLog FileNameIN
' Open data stream from picture
Set objStream = CreateObject("ADODB.Stream")
objStream.Type = adTypeBinary
objStream.Open()
objStream.LoadFromFile(FileNameIN)

' Create XML Document object and root node
' that will contain the data
Set objXML = CreateObject("MSXml2.DOMDocument")
Set objDocElem = objXML.createElement("Base64Data")
objDocElem.dataType = "bin.base64"

' Set binary value
objDocElem.nodeTypedValue = objStream.Read()
'msgbox FileNameOUT
' Open data stream to base64 code file
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFileOut = objFSO.CreateTextFile(FileNameOUT, fsDoOverwrite, fsAsASCII)

objFileOut.Write ITOg 


' Get base64 value and write to file
objFileOut.Write objDocElem.text
objFileOut.Close()

' Clean all
Set objFSO = Nothing
Set objFileOut = Nothing
Set objXML = Nothing
Set objDocElem = Nothing
Set objStream = Nothing
End Function

Sub DeleteFile(FileToDelete)
     Set fso = CreateObject("Scripting.FileSystemObject") 
     fso.DeleteFile FileToDelete 
End Sub


Sub WriteTEXTLog(LogMessage)
log_file = "E:\Schnittstellen\log\"
Const ForAppending = 8
A = year(Now)
B=Month(Now)
if len(B) < 2  then B="0" & B end if
C = day (Now)
if len(C)<2 then C ="0" & C end if
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objLogFile = objFSO.OpenTextFile(log_file & A & B & C & "Schadens_photos_VTL.log" , ForAppending, TRUE)
objLogFile.WriteLine("[" & Now() & " Schadens_photos_VTL_Update] " & LogMessage)
End Sub

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
   WriteTEXTLog sqlAnfr
   conn.Execute(sqlAnfr)
   
end Sub



Sub FTPUpload(path)

Set oShell = CreateObject("Shell.Application")
Set objFSO = CreateObject("Scripting.FileSystemObject")
On Error Resume Next

Const copyType = 16

'FTP Wait Time in ms
waitTime = 80000

FTPUser = Kennwort ("VTLFTPUser")
FTPPass = Kennwort ("VTLFTPPass")
FTPHost = Kennwort ("VTLFTPHost")
FTPDir = Kennwort ("VTLFTPDir")

strFTP = "ftp://" & FTPUser & ":" & FTPPass & "@" & FTPHost & FTPDir
Set objFTP = oShell.NameSpace(strFTP)


'Upload single file       
If objFSO.FileExists(path) Then

Set objFile = objFSO.getFile(path)
strParent = objFile.ParentFolder
Set objFolder = oShell.NameSpace(strParent)

Set objItem = objFolder.ParseName(objFile.Name)

'Wscript.Echo "Uploading file " & objItem.Name & " to " & strFTP
 objFTP.CopyHere objItem, copyType


End If


'Upload all files in folder
If objFSO.FolderExists(path) Then

'Entire folder
Set objFolder = oShell.NameSpace(path)

'Wscript.Echo "Uploading folder " & path & " to " & strFTP
objFTP.CopyHere objFolder.Items, copyType

End If


If Err.Number <> 0 Then
Wscript.Echo "Error: " & Err.Description
End If

'Wait for upload
Wscript.Sleep waitTime

End Sub

sub BackUpFile (PatchXML,FolderTo )

Set FSO=CreateObject("Scripting.FileSystemObject")
Set fldr= FSO.GetFolder(PatchXML)
Set Collec_Files= fldr.Files
For Each File in Collec_Files
    If Collec_Files.count = 0 then
      WriteTEXTLog PatchXML & " ist leer "
    Else
       Year_yy = year(Now)
      Month_MM=Month(Now)
      if len(Month_MM) < 2  then Month_MM = "0" & Month_MM 
      Day_dd = day (Now)
      if len(Day_dd)<2 then Day_dd ="0" & Day_dd 
      Hour_HH= (Hour(Now))
      Minute_mm = (Minute(Now))
      Datatime_YY = Year_yy & Month_MM & Day_dd & Hour_HH & Minute_mm 
      WriteTEXTLog PatchXML & File.Name & " " & FolderTo & Datatime_YY & "_"& File.Name  
      FSO.MoveFile PatchXML & File.Name, FolderTo & Datatime_YY & "_"& File.Name  
    End If
Next


End sub

Function Kennwort (PArramm)
'Kennwort= ""
filename = "\\srv-dc2\DATEN$\IT\script\Key\key.zip"
set goFS = CreateObject("Scripting.FileSystemObject")
set ts = goFS.OpenTextFile( filename)
do until ts.AtEndOfStream
    strl = ts.ReadLine
    if left (strl, len(PArramm)) = PArramm then Kennwort= right (strl, len(strl) - len(PArramm) - 3) 
loop
End Function