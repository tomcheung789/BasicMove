VERSION 5.00
Begin VB.Form Main 
   Caption         =   "Basic Move"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9015
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   9015
   StartUpPosition =   3  '系統預設值
   Begin VB.TextBox txtscript 
      Height          =   4335
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  '兩者皆有
      TabIndex        =   5
      Top             =   1320
      Width           =   4095
   End
   Begin VB.Frame bxresult 
      Caption         =   "Result"
      Height          =   5535
      Left            =   4320
      TabIndex        =   3
      Top             =   120
      Width           =   4695
      Begin VB.TextBox txtresult 
         Height          =   5175
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  '兩者皆有
         TabIndex        =   4
         Top             =   240
         Width           =   4455
      End
   End
   Begin VB.Frame Process 
      Caption         =   "Process"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.CommandButton btnclean 
         Caption         =   "Clean Backup"
         Height          =   495
         Left            =   2040
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton btnmove 
         Caption         =   "Run Script"
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ExitWindows Lib "User32" _
Alias "ExitWindowsEx" _
(ByVal dwOptions As Long, _
ByVal dwReserved As Long) As Long

Private Declare Function MoveFileEx Lib "kernel32" _
Alias "MoveFileExA" _
(ByVal lpExistingFileName As String, _
ByVal lpNewFileName As String, _
ByVal dwFlags As Long) As Long

Private Declare Function PathIsDirectory Lib "shlwapi.dll" _
Alias "PathIsDirectoryA" _
(ByVal pszPath As String) As Long

Private Declare Function GetSystemDirectory Lib "kernel32" _
Alias "GetSystemDirectoryA" _
(ByVal lpBuffer As String, _
ByVal nSize As Long) As Long

Private Type SERVICE_STATUS
dwServiceType As Long
dwCurrentState As Long
dwControlsAccepted As Long
dwWin32ExitCode As Long
dwServiceSpecificExitCode As Long
dwCheckPoint As Long
dwWaitHint As Long
End Type

Private Declare Function OpenSCManager Lib "advapi32.dll" _
Alias "OpenSCManagerA" _
(ByVal lpMachineName As String, _
ByVal lpDatabaseName As String, _
ByVal dwDesiredAccess As Long) As Long

Private Declare Function OpenService Lib "advapi32.dll" _
Alias "OpenServiceA" _
(ByVal hSCManager As Long, _
ByVal lpServiceName As String, _
ByVal dwDesiredAccess As Long) As Long

Private Declare Function ControlService Lib "advapi32.dll" _
(ByVal hService As Long, _
ByVal dwControl As Long, _
 lpServiceStatus As SERVICE_STATUS) As Long

Private Declare Function DeleteService Lib "advapi32.dll" _
(ByVal hService As Long) As Long

Private Declare Function CloseServiceHandle Lib "advapi32.dll" _
(ByVal hSCObject As Long) As Long


Dim reboot As Boolean


Private Sub btnclean_Click()
On Error GoTo err
Kill (GetSysDisk() + "BasicMove\*")
MsgBox ("Finished!")
err:

End Sub

Private Sub btnmove_Click()
If (txtscript.Text = "") Then
Call MsgBox("Script Is Empty", vbOKOnly, "Warning!")
Else
If Not (FolderExist(GetSysDisk() + "BasicMove")) Then
MkDir (GetSysDisk() + "BasicMove")
End If
Call ScriptProcess
End If
End Sub

Private Function GetSysDisk() As String
Dim sSysDir As String
sSysDir = String(255, Chr(0))
Call GetSystemDirectory(sSysDir, 255)
GetSysDisk = Left(sSysDir, 3)
End Function

Private Sub ScriptProcess()
Dim Line() As String, i As Integer
Dim service As String, file As String, folder As String, Reg As String
Dim nowIn As Integer
nowIn = 9
reboot = False
service = ""
file = ""
folder = ""
Reg = ""
Line = Split(txtscript.Text, vbNewLine)

For i = 0 To UBound(Line) Step 1

If (Line(i) = "::services") Then
nowIn = 0
GoTo Continue
End If
If (Line(i) = "::files") Then
nowIn = 1
GoTo Continue
End If
If (Line(i) = "::folders") Then
nowIn = 2
GoTo Continue
End If
If (Line(i) = "::regs") Then
nowIn = 3
GoTo Continue
End If

Select Case nowIn
Case 0
service = service + vbNewLine + Line(i)
Case 1
file = file + vbNewLine + Line(i)
Case 2
folder = folder + vbNewLine + Line(i)
Case 3
Reg = Reg + vbNewLine + Line(i)
End Select

Continue:
Next

txtresult.Text = "Basic Move Result:"

If (service <> "") Then
ServiceProcess (service)
End If
If (file <> "") Then
FileProcess (file)
End If
If (folder <> "") Then
FolderProcess (folder)
End If
If (Reg <> "") Then
RegProcess (Reg)
End If

If (reboot) Then
If (MsgBox("Need Reboot", vbYesNo, "BasicMove") = vbYes) Then
Call ExitWindows(2, &HFFFFFFFF)
End If
Else
Call MsgBox("Finished!", , "BasicMove")
End If
End Sub

Private Sub ServiceProcess(ByRef service As String)
Dim Line() As String
Line = Split(service, vbNewLine)
For i = 0 To UBound(Line) Step 1
If (Line(i) <> "") Then
If (ServiceDelete(Line(i)) = True) Then
txtresult.Text = txtresult.Text + vbNewLine + "Service : " + Line(i) + " Deleted"
Else
txtresult.Text = txtresult.Text + vbNewLine + "Service : " + Line(i) + " Not Deleted or Not Exist"
End If
End If
Next
End Sub

Public Function ServiceDelete(ServiceName As String) As Boolean
Dim hSCManager As Long, hService As Long, STATUS As SERVICE_STATUS
hSCManager = OpenSCManager(vbNullString, vbNullString, 983103)
hService = OpenService(hSCManager, ServiceName, 983551)
Call ControlService(hService, 1, STATUS)
ServiceDelete = DeleteService(hService)
CloseServiceHandle (hService)
CloseServiceHandle (hSCManager)
End Function

Private Sub FileProcess(ByRef file As String)
Dim Line() As String
Dim tempstr() As String
Line = Split(file, vbNewLine)
For i = 0 To UBound(Line) Step 1
If (Line(i) <> "") Then
If (FileExists(Line(i))) Then
tempstr = Split(Line(i), "\")
If (MoveFileEx(Line(i), GetSysDisk() + "BasicMove\" + tempstr(UBound(tempstr)) + ".bak", 2) <> 0) Then
txtresult.Text = txtresult.Text + vbNewLine + "File : " + Line(i) + " Moved"
Else
Call MoveFileEx(Line(i), GetSysDisk() + "BasicMove\" + tempstr(UBound(tempstr)) + ".bak", 4)
txtresult.Text = txtresult.Text + vbNewLine + "File : " + Line(i) + " Move at reboot"
reboot = True
End If
Else
txtresult.Text = txtresult.Text + vbNewLine + "File : " + Line(i) + " Not Exist"
End If
End If
Next
End Sub

Private Function FileExists(ByVal sFileName As String) As Boolean
Dim intReturn As Integer
On Error GoTo FalseFile
intReturn = GetAttr(sFileName)
FileExists = True
Exit Function
FalseFile:
FileExists = False
End Function

Private Sub FolderProcess(ByRef folder As String)
Dim Line() As String
Line = Split(folder, vbNewLine)
For i = 0 To UBound(Line) Step 1
If (Line(i) <> "") Then
If (FolderExist(Line(i))) Then
On Error GoTo err
Kill (Line(i) + "\*")
err:
On Error GoTo err1
RmDir (Line(i))
If (FolderExist(Line(i))) Then
err1:
txtresult.Text = txtresult.Text + vbNewLine + "Folder : " + Line(i) + " Not Deleted"
Else
txtresult.Text = txtresult.Text + vbNewLine + "Folder : " + Line(i) + " Deleted"
End If
Else
txtresult.Text = txtresult.Text + vbNewLine + "Folder : " + Line(i) + " Not Exist"
End If
End If
Next
End Sub

Private Function FolderExist(ByVal Path As String) As Boolean
If (PathIsDirectory(Path) <> 0) Then
FolderExist = True
Else
FolderExist = False
End If
End Function

Private Sub RegProcess(ByRef Reg As String)
Open GetSysDisk() + "BasicMove\Tmp.reg" For Output As #1
Print #1, "Windows Registry Editor Version 5.00" + vbNewLine + Reg
Close #1
Shell "regedit.exe /s " + GetSysDisk() + "BasicMove\Tmp.reg"
On Error GoTo err2
Kill (GetSysDisk() + "BasicMove\Tmp.reg")
err2:
End Sub

Private Sub WriteLog()
Open GetSysDisk() + "BasicMove\BasicMove.Log" For Output As #1
Print #1, txtresult.Text
Close #1
End Sub












