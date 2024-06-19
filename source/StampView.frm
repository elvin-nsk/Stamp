VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} StampView 
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3345
   OleObjectBlob   =   "StampView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "StampView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'===============================================================================
' # State

Public IsOk As Boolean
Public IsCancel As Boolean

Public FirstContourHandler As TextBoxHandler
Public AddToTextOutlineHandler As TextBoxHandler
Public AddToVectorOutlineHandler As TextBoxHandler

'===============================================================================
' # Constructor

Private Sub UserForm_Initialize()
    Caption = APP_DISPLAYNAME & " (v" & APP_VERSION & ")"
    btnOk.Default = True
    Set FirstContourHandler = _
        TextBoxHandler.New_(FirstContour, TextBoxTypeDouble, 0)
    Set AddToTextOutlineHandler = _
        TextBoxHandler.New_(AddToTextOutline, TextBoxTypeDouble, 0)
    Set AddToVectorOutlineHandler = _
        TextBoxHandler.New_(AddToVectorOutline, TextBoxTypeDouble, 0)
End Sub

'===============================================================================
' # Handlers

Private Sub UserForm_Activate()
    '
End Sub

Private Sub btnOk_Click()
    FormŒ 
End Sub

Private Sub btnCancel_Click()
    FormCancel
End Sub

'===============================================================================
' # Logic

Private Sub FormŒ ()
    Hide
    IsOk = True
End Sub

Private Sub FormCancel()
    Hide
    IsCancel = True
End Sub

'===============================================================================
' # Helpers



'===============================================================================
' # Boilerplate

Private Sub UserForm_QueryClose(—ancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        —ancel = True
        FormCancel
    End If
End Sub
