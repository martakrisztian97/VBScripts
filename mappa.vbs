Dim objFSO, fajl
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set fajl = objFSO.GetFile("D:\Egyetem\2021-22-1\GUI programozas\VBS\msgbox2.vbs")
Set objFolder = objFSO.GetFolder(fajl.ParentFolder)
MsgBox("Szulo mappa: "&fajl.ParentFolder), 64
MsgBox("Eleresi ut: "&fajl.Path), 64
MsgBox("Fajlok szama: "&objFolder.Files.Count)
MsgBox("Almappak szama: "&objFolder.subfolders.count)