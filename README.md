# USB Data Backup Script

## Description

This script is written in VBScript and is designed to automatically detect and copy the contents of any USB drive connected to the system. It continuously monitors for USB drives and, upon detection, creates a backup of the USB drive's contents into a designated folder on the local system.

## Features

✅  Automatically detects USB drives.

✅ Copies all files and folders from the USB to a local directory.

✅ Creates a separate folder for each USB device using its volume name and timestamp.

✅ Runs continuously, checking for USB drives every 5 seconds.

✅ Ensures that no duplicate backups are created for the same drive at the same time.

## Prerequisites

✅ Windows operating system

✅ Windows Script Host enabled

## How It Works

**1.** The script starts by defining the main destination folder (C:\user). If the folder does not exist, it is created.

**2.** It continuously checks for USB drives (DriveType = 2) using Windows Management Instrumentation (WMI).

**3.** When a USB drive is detected:

✅ The script retrieves the volume name of the USB drive.

✅ It creates a unique folder using the volume name and a timestamp (YYYY-MM-DD_HH-MM).

✅ It copies all files and folders from the USB drive to the created backup folder.

**4.** The script runs indefinitely, checking for new USB drives every 5 seconds.

## Installation

**1.** Copy the script into a .vbs file (e.g., usb_backup.vbs).

**2.** Place the script in a secure location on your system.

**3.** Run the script by double-clicking it or executing it via the command prompt using:

cscript usb_backup.vbs

## Code Explanation

Option Explicit

Dim fso, wmi, usb, drive, destination, objFolder, objFile
Set fso = CreateObject("Scripting.FileSystemObject")
Set wmi = GetObject("winmgmts:\\.\root\cimv2")

' Define main folder
destination = "C:\\user"
If Not fso.FolderExists(destination) Then
    fso.CreateFolder(destination)
End If

Do
    Set usb = wmi.ExecQuery("Select * From Win32_LogicalDisk Where DriveType = 2")
    For Each drive In usb
        On Error Resume Next
        If fso.DriveExists(drive.DeviceID) Then
            Dim usbName, usbFolder
            usbName = GetUSBName(drive.DeviceID)
            usbFolder = destination & "\\" & usbName & "_" & GetTimestamp()
            
            If Not fso.FolderExists(usbFolder) Then
                fso.CreateFolder(usbFolder)
            End If
            
            CopyDrive drive.DeviceID, usbFolder
        End If
        On Error GoTo 0
    Next
    WScript.Sleep 5000 ' Check every 5 seconds 
Loop

## Functions

**✅  CopyDrive:** Copies all files and folders from the USB drive to the designated backup location.

**✅ GetUSBName:** Retrieves the volume name of the USB drive.

**✅ GetTimestamp:** Generates a timestamp for unique folder naming.

## Notes

✅ Ensure that the script is run with appropriate permissions.

✅ The backup location can be changed by modifying the destination variable.

✅ The script will continue running until manually terminated.

## License

This script is provided for educational and personal use. Use responsibly and ensure you have permission to copy USB contents before running the script.

