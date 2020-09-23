VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows-Standard
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const SHFD_CAPACITY_DEFAULT = 0 ' default drive capacity
Const SHFD_CAPACITY_360 = 3 ' 360KB, applies to 5.25" drives only
Const SHFD_CAPACITY_720 = 5 ' 720KB, applies to 3.5" drives only
Const SHFD_FORMAT_QUICK = 0 ' quick format
Const SHFD_FORMAT_FULL = 1 ' full format
Const SHFD_FORMAT_SYSONLY = 2 ' copies system files only (Win95 Only!)
Private Declare Function SHFormatDrive Lib "shell32" (ByVal hwndOwner As Long, ByVal iDrive As Long, ByVal iCapacity As Long, ByVal iFormatType As Long) As Long
Private Sub Form_Load()
   'iDrive = The drive number to format. Drive A=0, B=1 (if present, otherwise C=1), and so on.
    SHFormatDrive Me.hWnd, 0, SHFD_CAPACITY_DEFAULT, SHFD_FORMAT_QUICK
End Sub


