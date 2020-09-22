VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Base Address Locator by Chris Vega"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6000
   Icon            =   "GetAddress.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   6000
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFind 
      Caption         =   "Search API-Function"
      Height          =   345
      Left            =   120
      TabIndex        =   13
      Top             =   2550
      Width           =   1845
   End
   Begin VB.ComboBox cboModuleName 
      Height          =   315
      ItemData        =   "GetAddress.frx":08CA
      Left            =   1170
      List            =   "GetAddress.frx":147A
      TabIndex        =   12
      Top             =   60
      Width           =   4095
   End
   Begin VB.CommandButton cmdGO 
      Caption         =   "Load"
      Height          =   315
      Left            =   5310
      TabIndex        =   11
      Top             =   60
      Width           =   645
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   345
      Left            =   4110
      TabIndex        =   10
      Top             =   2550
      Width           =   1755
   End
   Begin VB.Frame frX 
      Height          =   1995
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   5805
      Begin VB.CommandButton cmdGEP 
         Caption         =   "GEP"
         Height          =   315
         Left            =   5160
         TabIndex        =   14
         Top             =   1500
         Width           =   495
      End
      Begin VB.TextBox txtAPIHandle 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1500
         Width           =   3105
      End
      Begin VB.TextBox txtAPIName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         TabIndex        =   7
         Top             =   1140
         Width           =   3615
      End
      Begin VB.TextBox txtModuleHandle 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   570
         Width           =   3615
      End
      Begin VB.TextBox txtModulePath 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   210
         Width           =   3615
      End
      Begin VB.Label lblAPIHandle 
         AutoSize        =   -1  'True
         Caption         =   "API Function Entry-Point"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1560
         Width           =   1725
      End
      Begin VB.Label lblAPIName 
         Caption         =   "API Function Name"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1170
         Width           =   1665
      End
      Begin VB.Label lblModulePath 
         Caption         =   "Dynamic Link Library Path"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   270
         Width           =   1935
      End
      Begin VB.Label lblModuleHandle 
         Caption         =   "Module Base Handle"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1665
      End
   End
   Begin VB.Label lblModuleName 
      AutoSize        =   -1  'True
      Caption         =   "Module Name"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   990
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===================================------------
'=======-
'==-
'-       Base Locator (A Dynamic Link Libary Tool)
'
')       Written by: Chris Vega
'-)                  gwapo@models.com
'--)
'---)______-_===^

' API functions declarations
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Function LoadLibraryEx Lib "kernel32" Alias "LoadLibraryExA" (ByVal lpLibFileName As String, ByVal hFile As Long, ByVal dwFlags As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

' API Constant Declaration
Private Const DONT_RESOLVE_DLL_REFERENCES = &H1

' Used Variables
Private DLLFiles                ' DLL Searching
Private ModHandle As Long       ' ImageBase for Module
Private ModName As String        ' Image Name

Private Sub cboModuleName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And Len(Trim(cboModuleName.Text)) > 0 Then cmdGO_Click
End Sub

'===================
' Search Entry Point
'===================
Private Sub cmdFind_Click()
    If Len(Trim(txtAPIName)) = 0 Then
        MsgBox "Please Enter API-Function-Name to Search", vbInformation, "Chris Vega"
        txtAPIName.SetFocus
        Exit Sub
    End If
    
    ' We are here to search for API-Function with our list
    ' of DLL Modules on the cboModuleName Object
    Dim ModNames() As String, ModCount As Integer
    Dim ModHandle As Long, MsgText As String
    Dim ProcHandle As Long
    
    If MsgBox("Your system currently loaded with " & cboModuleName.ListCount & " Modules; and searching API-Functions to all these modules requires sometime," & vbCrLf & vbCrLf & "Do you wish to proceed?", vbQuestion + vbYesNo, "Chris Vega") = vbYes Then
        Me.MousePointer = 11    ' HourGlass MousePointer
        ModCount = 0
        lblModuleName.Caption = "Searching"

        For i = 0 To cboModuleName.ListCount - 1
            DoEvents    ' Dont Freeze the System
            cboModuleName.Text = cboModuleName.List(i)  ' Show Status
            ' Use LoadLibraryEx to Get the ImageBase of this DLL
            ModHandle = LoadLibraryEx(cboModuleName.List(i), 0, DONT_RESOLVE_DLL_REFERENCES)
            ' User GetProcAddress API to get its Entry-Point on Export Table (Virtual Address)
            ProcHandle = GetProcAddress(ModHandle, Trim(txtAPIName))
            ' Release the Module; We dont need it anymore
            FreeLibrary ModHandle
            ' Does we have an Export Entry containing the
            ' API-Function, user has trying to locate?
            If ProcHandle > 0 Then
                ' Yes, Increase the Counter
                ModCount = ModCount + 1
                ' And the Name Table
                ReDim Preserve ModNames(ModCount)
                ModNames(ModCount) = cboModuleName.Text
            End If
        Next
        
        ' Does we got Matches; Counter is above zero
        If ModCount > 0 Then
            ' Yes; then build message for the User
            MsgText = "API-Function " & txtAPIName & " has been located at these module(s):" & vbCrLf
            For ProcHandle = LBound(ModNames) To UBound(ModNames)
                MsgText = MsgText & ModNames(ProcHandle) & vbCrLf
            Next
            ' Using the Name Table expanded a while ago
            cboModuleName.Text = ModNames(1)
            cmdGO_Click
            ' Show it!
            MsgBox MsgText, vbInformation, "Chris Vega"
        Else
            ' No; Then Show it either way
            MsgBox "Unfortunately, With all " & cboModuleName.ListCount & " modules; the API-Function you are looking for has not been located." & vbCrLf & vbCrLf & "Please try searching again (API-Functions are Case Sensitive)", vbExclamation, "Chris Vega"
        End If
        
        ' Resume Normal
        Me.MousePointer = 0     ' Default MousePointer
        lblModuleName.Caption = "Module Name"
    End If
End Sub

'==================
' Get Entry Point
'==================
Private Sub cmdGEP_Click()
    Dim ProcHandle As Long
    ' GetProcAddress returns the Entry-Point; using the before saved ImageBase
    cmdGO_Click
    ProcHandle = GetProcAddress(ModHandle, Trim(txtAPIName))
    If ProcHandle > 0 Then
        ' We are displaying in Hexadecimal Format
        txtAPIHandle = Hex(ProcHandle) & "h"
    Else
        ' Clear on error
        txtAPIHandle = ""
        ' Alert User
        MsgBox txtAPIName & " API-Function is not exported by this module", vbInformation, "Chris Vega"
    End If
End Sub

'===================
' Load Module
'===================
Private Sub cmdGO_Click()
    ' Again, LoadLibrary is used instead of GetModuleHandle;
    ' because most DLL we are using is not refferenced on our
    ' Running Space (ImageBase)
    ModHandle = LoadLibraryEx(Trim(cboModuleName.Text), 0, DONT_RESOLVE_DLL_REFERENCES)
    If ModHandle > 0 Then
        On Error Resume Next
        ModName = cboModuleName.Text
        txtModulePath = ModuleName(ModHandle)
        ' On Hex Please
        txtModuleHandle = Hex(ModHandle) & "h"
        ' Release
        FreeLibrary ModHandle
        txtAPIName.SetFocus
    Else
        MsgBox "It doesn't appear that this is a valid Module (DLL) file", vbExclamation, "Chris Vega"
        cboModuleName.Text = ModName
    End If
End Sub

Private Sub Form_Load()
    ' Load with all the DLLs in this System
    DLLFiles = GetAllDLLFiles
    ' Move the Data Collected to cboModuleName object
    For i = LBound(DLLFiles) To UBound(DLLFiles)
        cboModuleName.AddItem DLLFiles(i)
    Next
    ' KERNEL32.DLL is our Default
    cboModuleName.Text = "kernel32"
    ' Load it!
    cmdGO_Click
End Sub

' Returns the System Root Directory %SysRoot%
Private Function SystemDirectory() As String
    Dim Str As String, lngRet As Long
    Str = String(128, 0)
    lngRet = GetSystemDirectory(Str, 128)
    SystemDirectory = Left(Str, lngRet)
End Function

' Returns the Complete Path of the Module refferenced
' by its ImageBase (ModHandle)
Private Function ModuleName(ModHandle As Long) As String
    Dim Str As String, lngRet As Long
    Str = String(128, 0)
    lngRet = GetModuleFileName(ModHandle, Str, 128)
    ModuleName = Left(Str, lngRet)
End Function

' Get All the list of DLL Files by manual searching through
' %SysRoot% Folder (%WinDIR% will be skipped for now)
Private Function GetAllDLLFiles()
    On Error Resume Next

    Dim xFiles() As String
    Dim xDir As String, xDirCnt As Integer
    Dim xMod As Long
    
    ' Only Search %System Root% Folder for DLL Files
    ' No Test will be done since all DLL loaded in this
    ' Root Directory will be considered as Valid Module
    ' unless somehow your system get messed-up ;)
    xDir = Dir(SystemDirectory & "\*.dll", vbArchive)
    xDirCnt = 0

    While Len(xDir)
        ReDim Preserve xFiles(xDirCnt)
        ' Do not include File Extension; So it will
        ' appear that the Search returns Module by
        ' name and not by Filename
        xFiles(xDirCnt) = Left(xDir, Len(xDir) - 4)
        ' Next Entry
        xDir = Dir
    Wend
    
    ' Return the Search Result
    GetAllDLLFiles = xFiles
End Function

Private Sub cmdClose_Click()
    ' No Rights Reserved, Use Without Permission
    MsgBox "(C) 2001 by Chris Vega", vbInformation, "Chris Vega"
    ' Fin ;)
    End
End Sub

Private Sub txtAPIName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And Len(Trim(txtAPIName.Text)) > 0 Then cmdGEP_Click
End Sub
