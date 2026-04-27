' pybasil Regression Test Suite - Scripting.FileSystemObject
' Creation date: 2026-04-27
' Purpose: Comprehensive regression test of FSO file system operations

Dim passCount, failCount, totalCount
passCount = 0
failCount = 0
totalCount = 0

Sub AssertEqual(testName, actual, expected)
    totalCount = totalCount + 1
    If actual = expected Then
        passCount = passCount + 1
    Else
        failCount = failCount + 1
        WScript.Echo "FAIL: " & testName & " - Expected: " & CStr(expected) & " Got: " & CStr(actual)
    End If
End Sub

Sub AssertTrue(testName, condition)
    totalCount = totalCount + 1
    If condition Then
        passCount = passCount + 1
    Else
        failCount = failCount + 1
        WScript.Echo "FAIL: " & testName & " - Expected True Got False"
    End If
End Sub

Sub AssertFalse(testName, condition)
    totalCount = totalCount + 1
    If Not condition Then
        passCount = passCount + 1
    Else
        failCount = failCount + 1
        WScript.Echo "FAIL: " & testName & " - Expected False Got True"
    End If
End Sub

' Create the FSO and set up a temp directory for all tests
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

Dim tempRoot
tempRoot = fso.BuildPath(fso.GetSpecialFolder(2).Path, "pybasil_fso_regression")

' Clean up from any previous failed run
If fso.FolderExists(tempRoot) Then fso.DeleteFolder tempRoot
fso.CreateFolder tempRoot

' =========================================================================
'  PATH HELPERS
' =========================================================================
WScript.Echo "--- Path Helpers ---"

Call AssertEqual("BuildPath basic", fso.BuildPath("/tmp", "file.txt"), "/tmp/file.txt")
Call AssertEqual("GetFileName", fso.GetFileName("/path/to/document.pdf"), "document.pdf")
Call AssertEqual("GetBaseName", fso.GetBaseName("/path/to/document.pdf"), "document")
Call AssertEqual("GetExtensionName", fso.GetExtensionName("/path/to/document.pdf"), "pdf")
Call AssertEqual("GetExtensionName no ext", fso.GetExtensionName("README"), "")
Call AssertEqual("GetParentFolderName", fso.GetParentFolderName("/path/to/file.txt"), "/path/to")

Dim absPath
absPath = fso.GetAbsolutePathName(".")
Call AssertTrue("GetAbsolutePathName non-empty", Len(absPath) > 0)

Dim tmpName
tmpName = fso.GetTempName
Call AssertTrue("GetTempName non-empty", Len(tmpName) > 0)

' =========================================================================
'  FILE EXISTS / FOLDER EXISTS
' =========================================================================
WScript.Echo "--- Exists Checks ---"

Call AssertTrue("FolderExists tempRoot", fso.FolderExists(tempRoot))
Call AssertFalse("FolderExists nonexistent", fso.FolderExists(tempRoot & "/nope"))
Call AssertFalse("FileExists nonexistent", fso.FileExists(tempRoot & "/nope.txt"))

' =========================================================================
'  CREATE TEXT FILE & WRITE
' =========================================================================
WScript.Echo "--- CreateTextFile & Write ---"

Dim writePath
writePath = fso.BuildPath(tempRoot, "write_test.txt")

Dim tsWrite
Set tsWrite = fso.CreateTextFile(writePath)
tsWrite.Write "Hello"
tsWrite.Close

Call AssertTrue("CreateTextFile creates file", fso.FileExists(writePath))

' =========================================================================
'  OPEN TEXT FILE & READ
' =========================================================================
WScript.Echo "--- OpenTextFile & Read ---"

Dim tsRead
Set tsRead = fso.OpenTextFile(writePath)
Dim readContent
readContent = tsRead.ReadAll
tsRead.Close

Call AssertEqual("ReadAll content", readContent, "Hello")

' =========================================================================
'  WRITELINE & READLINE
' =========================================================================
WScript.Echo "--- WriteLine & ReadLine ---"

Dim linesPath
linesPath = fso.BuildPath(tempRoot, "lines_test.txt")

Dim tsLines
Set tsLines = fso.CreateTextFile(linesPath)
tsLines.WriteLine "Alpha"
tsLines.WriteLine "Beta"
tsLines.WriteLine "Gamma"
tsLines.Close

Set tsLines = fso.OpenTextFile(linesPath)
Dim line1, line2, line3
line1 = tsLines.ReadLine
line2 = tsLines.ReadLine
line3 = tsLines.ReadLine
tsLines.Close

Call AssertEqual("ReadLine 1", line1, "Alpha")
Call AssertEqual("ReadLine 2", line2, "Beta")
Call AssertEqual("ReadLine 3", line3, "Gamma")

' =========================================================================
'  ATENDOFSTREAM LOOP
' =========================================================================
WScript.Echo "--- AtEndOfStream ---"

Set tsLines = fso.OpenTextFile(linesPath)
Dim lineCount
lineCount = 0
Do While Not tsLines.AtEndOfStream
    tsLines.ReadLine
    lineCount = lineCount + 1
Loop
tsLines.Close

Call AssertEqual("AtEndOfStream line count", lineCount, 3)

' =========================================================================
'  READ(n) - PARTIAL READ
' =========================================================================
WScript.Echo "--- Read(n) ---"

Set tsRead = fso.OpenTextFile(writePath)
Dim partial
partial = tsRead.Read(3)
tsRead.Close

Call AssertEqual("Read(3) from Hello", partial, "Hel")

' =========================================================================
'  WRITEBLANKLINES
' =========================================================================
WScript.Echo "--- WriteBlankLines ---"

Dim blankPath
blankPath = fso.BuildPath(tempRoot, "blank_test.txt")

Dim tsBlank
Set tsBlank = fso.CreateTextFile(blankPath)
tsBlank.Write "A"
tsBlank.WriteBlankLines 2
tsBlank.Write "B"
tsBlank.Close

Set tsBlank = fso.OpenTextFile(blankPath)
Dim blankContent
blankContent = tsBlank.ReadAll
tsBlank.Close

Call AssertEqual("WriteBlankLines content", blankContent, "A" & vbLf & vbLf & "B")

' =========================================================================
'  TEXTSTREAM LINE PROPERTY
' =========================================================================
WScript.Echo "--- TextStream Line Property ---"

Set tsLines = fso.OpenTextFile(linesPath)
tsLines.ReadLine
tsLines.ReadLine
Dim lineNum
lineNum = tsLines.Line
tsLines.Close

Call AssertEqual("Line after 2 ReadLines", lineNum, 3)

' =========================================================================
'  APPEND MODE
' =========================================================================
WScript.Echo "--- Append Mode ---"

Dim appendPath
appendPath = fso.BuildPath(tempRoot, "append_test.txt")

Dim tsAppend
Set tsAppend = fso.CreateTextFile(appendPath)
tsAppend.Write "First"
tsAppend.Close

Set tsAppend = fso.OpenTextFile(appendPath, 8)
tsAppend.Write "Second"
tsAppend.Close

Set tsAppend = fso.OpenTextFile(appendPath)
Dim appendContent
appendContent = tsAppend.ReadAll
tsAppend.Close

Call AssertEqual("Append mode", appendContent, "FirstSecond")

' =========================================================================
'  GETFILE PROPERTIES
' =========================================================================
WScript.Echo "--- GetFile Properties ---"

Dim fileObj
Set fileObj = fso.GetFile(writePath)

Call AssertEqual("File.Name", fileObj.Name, "write_test.txt")
Call AssertEqual("File.Size", fileObj.Size, 5)
Call AssertTrue("File.Path ends correctly", Right(fileObj.Path, Len("write_test.txt")) = "write_test.txt")
Call AssertEqual("File.Type", fileObj.Type, ".txt")
Call AssertEqual("TypeName DateCreated", TypeName(fileObj.DateCreated), "Date")
Call AssertEqual("TypeName DateLastModified", TypeName(fileObj.DateLastModified), "Date")
Call AssertTrue("File.ParentFolder name", Len(fileObj.ParentFolder.Name) > 0)

' =========================================================================
'  FILE OPENASTEXTSTREAM
' =========================================================================
WScript.Echo "--- File.OpenAsTextStream ---"

Dim tsFromFile
Set tsFromFile = fileObj.OpenAsTextStream(1)
Dim oatContent
oatContent = tsFromFile.ReadAll
tsFromFile.Close

Call AssertEqual("OpenAsTextStream ReadAll", oatContent, "Hello")

' =========================================================================
'  GETFOLDER PROPERTIES
' =========================================================================
WScript.Echo "--- GetFolder Properties ---"

Dim folderObj
Set folderObj = fso.GetFolder(tempRoot)

Call AssertEqual("Folder.Name", folderObj.Name, "pybasil_fso_regression")
Call AssertEqual("Folder.Type", folderObj.Type, "File Folder")
Call AssertTrue("Folder.Size >= 0", folderObj.Size >= 0)
Call AssertFalse("Folder.IsRootFolder", folderObj.IsRootFolder)
Call AssertTrue("Folder.ParentFolder exists", Len(folderObj.ParentFolder.Path) > 0)

' =========================================================================
'  FOLDER FILES COLLECTION
' =========================================================================
WScript.Echo "--- Folder.Files ---"

Dim filesCount
filesCount = folderObj.Files.Count
Call AssertTrue("Folder.Files.Count > 0", filesCount > 0)

' Iterate files with For Each
Dim f, fileNames
fileNames = ""
For Each f In folderObj.Files
    If fileNames <> "" Then fileNames = fileNames & ","
    fileNames = fileNames & f.Name
Next
Call AssertTrue("For Each files non-empty", Len(fileNames) > 0)

' =========================================================================
'  SUBFOLDERS
' =========================================================================
WScript.Echo "--- SubFolders ---"

Dim subPath
subPath = fso.BuildPath(tempRoot, "sub1")
fso.CreateFolder subPath

Dim subPath2
subPath2 = fso.BuildPath(tempRoot, "sub2")
fso.CreateFolder subPath2

Set folderObj = fso.GetFolder(tempRoot)
Call AssertEqual("SubFolders.Count", folderObj.SubFolders.Count, 2)

Dim sf, subNames
subNames = ""
For Each sf In folderObj.SubFolders
    If subNames <> "" Then subNames = subNames & ","
    subNames = subNames & sf.Name
Next
Call AssertTrue("For Each subfolders", InStr(subNames, "sub1") > 0)
Call AssertTrue("For Each subfolders 2", InStr(subNames, "sub2") > 0)

' =========================================================================
'  COPYFILE
' =========================================================================
WScript.Echo "--- CopyFile ---"

Dim copySrc, copyDest
copySrc = writePath
copyDest = fso.BuildPath(tempRoot, "copy_of_write.txt")

fso.CopyFile copySrc, copyDest
Call AssertTrue("CopyFile dest exists", fso.FileExists(copyDest))

Dim tsCopy
Set tsCopy = fso.OpenTextFile(copyDest)
Call AssertEqual("CopyFile content", tsCopy.ReadAll, "Hello")
tsCopy.Close

' =========================================================================
'  MOVEFILE
' =========================================================================
WScript.Echo "--- MoveFile ---"

Dim moveSrc, moveDest
moveSrc = fso.BuildPath(tempRoot, "move_src.txt")
moveDest = fso.BuildPath(tempRoot, "move_dest.txt")

Dim tsMove
Set tsMove = fso.CreateTextFile(moveSrc)
tsMove.Write "moving"
tsMove.Close

fso.MoveFile moveSrc, moveDest
Call AssertFalse("MoveFile src gone", fso.FileExists(moveSrc))
Call AssertTrue("MoveFile dest exists", fso.FileExists(moveDest))

' =========================================================================
'  DELETEFILE
' =========================================================================
WScript.Echo "--- DeleteFile ---"

fso.DeleteFile moveDest
Call AssertFalse("DeleteFile removed", fso.FileExists(moveDest))

' =========================================================================
'  FILE.DELETE
' =========================================================================
WScript.Echo "--- File.Delete ---"

Dim delFilePath
delFilePath = fso.BuildPath(tempRoot, "file_delete_test.txt")
Set tsWrite = fso.CreateTextFile(delFilePath)
tsWrite.Write "delete me"
tsWrite.Close

Dim delFileObj
Set delFileObj = fso.GetFile(delFilePath)
delFileObj.Delete
Call AssertFalse("File.Delete removed", fso.FileExists(delFilePath))

' =========================================================================
'  COPYFOLDER
' =========================================================================
WScript.Echo "--- CopyFolder ---"

Dim copySrcDir, copyDestDir
copySrcDir = fso.BuildPath(tempRoot, "copy_src_dir")
copyDestDir = fso.BuildPath(tempRoot, "copy_dest_dir")

fso.CreateFolder copySrcDir
Set tsWrite = fso.CreateTextFile(fso.BuildPath(copySrcDir, "inner.txt"))
tsWrite.Write "inner content"
tsWrite.Close

fso.CopyFolder copySrcDir, copyDestDir
Call AssertTrue("CopyFolder dest exists", fso.FolderExists(copyDestDir))
Call AssertTrue("CopyFolder inner file", fso.FileExists(fso.BuildPath(copyDestDir, "inner.txt")))

' =========================================================================
'  MOVEFOLDER
' =========================================================================
WScript.Echo "--- MoveFolder ---"

Dim moveSrcDir, moveDestDir
moveSrcDir = fso.BuildPath(tempRoot, "move_src_dir")
moveDestDir = fso.BuildPath(tempRoot, "move_dest_dir")

fso.CreateFolder moveSrcDir
fso.MoveFolder moveSrcDir, moveDestDir
Call AssertFalse("MoveFolder src gone", fso.FolderExists(moveSrcDir))
Call AssertTrue("MoveFolder dest exists", fso.FolderExists(moveDestDir))

' =========================================================================
'  DELETEFOLDER
' =========================================================================
WScript.Echo "--- DeleteFolder ---"

fso.DeleteFolder moveDestDir
Call AssertFalse("DeleteFolder removed", fso.FolderExists(moveDestDir))

' =========================================================================
'  CREATEFOLDER ALREADY EXISTS ERROR
' =========================================================================
WScript.Echo "--- CreateFolder Error ---"

Dim errCreateFolder
On Error Resume Next
fso.CreateFolder tempRoot
errCreateFolder = Err.Number
On Error GoTo 0
Call AssertTrue("CreateFolder existing raises error", errCreateFolder <> 0)

' =========================================================================
'  GETSPECIALFOLDER
' =========================================================================
WScript.Echo "--- GetSpecialFolder ---"

Dim tempFolder
Set tempFolder = fso.GetSpecialFolder(2)
Call AssertTrue("GetSpecialFolder(2) exists", fso.FolderExists(tempFolder.Path))

' Invalid folder number
Dim errSpecial
On Error Resume Next
fso.GetSpecialFolder 99
errSpecial = Err.Number
On Error GoTo 0
Call AssertTrue("GetSpecialFolder(99) raises error", errSpecial <> 0)

' =========================================================================
'  DRIVES
' =========================================================================
WScript.Echo "--- Drives ---"

Call AssertTrue("Drives.Count >= 1", fso.Drives.Count >= 1)

' =========================================================================
'  CREATETEXTFILE OVERWRITE=FALSE ERROR
' =========================================================================
WScript.Echo "--- CreateTextFile Overwrite Error ---"

Dim errOverwrite
On Error Resume Next
fso.CreateTextFile writePath, False
errOverwrite = Err.Number
On Error GoTo 0
Call AssertTrue("CreateTextFile no-overwrite error", errOverwrite <> 0)

' =========================================================================
'  OPEN NONEXISTENT FILE ERROR
' =========================================================================
WScript.Echo "--- OpenTextFile Error ---"

Dim errOpen
On Error Resume Next
fso.OpenTextFile fso.BuildPath(tempRoot, "does_not_exist.txt")
errOpen = Err.Number
On Error GoTo 0
Call AssertTrue("OpenTextFile nonexistent error", errOpen <> 0)

' =========================================================================
'  DELETEFILE NONEXISTENT ERROR
' =========================================================================
WScript.Echo "--- DeleteFile Error ---"

Dim errDelete
On Error Resume Next
fso.DeleteFile fso.BuildPath(tempRoot, "does_not_exist.txt")
errDelete = Err.Number
On Error GoTo 0
Call AssertTrue("DeleteFile nonexistent error", errDelete <> 0)

' =========================================================================
'  INTEGRATION: BUILD PATH, CREATE, WRITE, READ, DELETE
' =========================================================================
WScript.Echo "--- Integration ---"

Dim intPath
intPath = fso.BuildPath(tempRoot, "integration.txt")

Dim tsInt
Set tsInt = fso.CreateTextFile(intPath)
Dim i
For i = 1 To 5
    tsInt.WriteLine "Line " & CStr(i)
Next
tsInt.Close

Set tsInt = fso.OpenTextFile(intPath)
Dim intLines
intLines = 0
Do While Not tsInt.AtEndOfStream
    Dim intLine
    intLine = tsInt.ReadLine
    intLines = intLines + 1
Loop
tsInt.Close

Call AssertEqual("Integration line count", intLines, 5)

Dim intFile
Set intFile = fso.GetFile(intPath)
Call AssertTrue("Integration file size > 0", intFile.Size > 0)
Call AssertEqual("Integration GetBaseName", fso.GetBaseName(intPath), "integration")
Call AssertEqual("Integration GetExtensionName", fso.GetExtensionName(intPath), "txt")

fso.DeleteFile intPath
Call AssertFalse("Integration deleted", fso.FileExists(intPath))

' =========================================================================
'  CLEANUP
' =========================================================================
WScript.Echo "--- Cleanup ---"

fso.DeleteFolder tempRoot
Call AssertFalse("Cleanup tempRoot deleted", fso.FolderExists(tempRoot))

' =========================================================================
'  SUMMARY
' =========================================================================

WScript.Echo ""
WScript.Echo "FileSystemObject regression test summary"
WScript.Echo "Total:  " & totalCount
WScript.Echo "Passed: " & passCount
WScript.Echo "Failed: " & failCount
If failCount = 0 Then
    WScript.Echo "Result:  ALL TESTS PASSED"
Else
    WScript.Echo "Result:  SOME TESTS FAILED"
End If