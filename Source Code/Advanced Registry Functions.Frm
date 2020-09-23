VERSION 5.00
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Advanced Registry Functions v0.88.16.0"
   ClientHeight    =   3405
   ClientLeft      =   2130
   ClientTop       =   1440
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3405
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   915
      Left            =   3270
      TabIndex        =   11
      Top             =   2355
      Width           =   1170
   End
   Begin VB.Frame Keys 
      Caption         =   "Create And Delete Keys"
      Height          =   975
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   3015
      Begin VB.CommandButton Deletekeys 
         Caption         =   "&Delete Key"
         Height          =   375
         Left            =   1560
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Createkeys 
         Caption         =   "&Create Key"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Write 
      Caption         =   "Write To The Registry"
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   90
      Width           =   4335
      Begin VB.CommandButton Writestring 
         Caption         =   "Write String"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Writebinary 
         Caption         =   "Write Binary"
         Height          =   375
         Left            =   1560
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Writedword 
         Caption         =   "Write Dword"
         Height          =   375
         Left            =   2880
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Read 
      Caption         =   "Reads From Registry"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   1185
      Width           =   4335
      Begin VB.CommandButton Readdword 
         Caption         =   "Read Dword"
         Height          =   375
         Left            =   2880
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Readbinary 
         Caption         =   "Read Binary"
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Readstring 
         Caption         =   "Read String"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RetMsg As String
Dim Reg As New AdvancedRegistryFunctions

Private Sub Command1_Click()

'// unloads Main form

Unload Me

End Sub

Private Sub Createkeys_Click()

'// This creates a new key in the registry

RetMsg = Reg.CreateKey("HKEY_LOCAL_MACHINE\Software\Test")

MsgBox RetMsg, vbInformation, App.Title

End Sub

Private Sub Deletekeys_Click()

'// This deletes the key from the registry

RetMsg = Reg.DeleteKey("HKEY_LOCAL_MACHINE\Software\Test")

MsgBox RetMsg, vbInformation, App.Title

End Sub

Private Sub Readbinary_Click()

'// reads Binary Value from the registry

RetMsg = Reg.GetBinaryValue("HKEY_LOCAL_MACHINE\Software\Test", "Binary Value")

If RetMsg = Chr$(&H1) + Chr$(&H2) + Chr$(&H3) + Chr$(&H4) Then

MsgBox "Binary Value Successfully Retrieved", vbInformation, App.Title

 Else
 
 MsgBox "Binary Value Not Successfully Retrieved", vbInformation, App.Title
 
End If

End Sub

Private Sub Readdword_Click()

'// reads Dword Value from the registry

RetMsg = Reg.GetDWordValue("HKEY_LOCAL_MACHINE\Software\Test", "DWORD Value")

MsgBox "Dword Value is " & RetMsg, vbInformation, App.Title

End Sub

Private Sub Readstring_Click()

'// reads String Value from the registry

RetMsg = Reg.GetStringValue("HKEY_LOCAL_MACHINE\Software\Test", "String Value")

MsgBox "String Value is " & RetMsg, vbInformation, App.Title

End Sub


Private Sub Writebinary_Click()

'// writes Binary Value to the registry

RetMsg = Reg.SetBinaryValue("HKEY_LOCAL_MACHINE\Software\Test", "Binary Value", Chr$(&H1) + Chr$(&H2) + Chr$(&H3) + Chr$(&H4))

MsgBox RetMsg, vbInformation, App.Title

End Sub

Private Sub Writedword_Click()

'// writes Dword Value to the registry

RetMsg = Reg.SetDWordValue("HKEY_LOCAL_MACHINE\Software\Test", "DWORD Value", "1")

MsgBox RetMsg, vbInformation, App.Title

End Sub

Private Sub Writestring_Click()

'// writes String Value to the registry

RetMsg = Reg.SetStringValue("HKEY_LOCAL_MACHINE\Software\Test", "String Value", "Test String")

MsgBox RetMsg, vbInformation, App.Title

End Sub
