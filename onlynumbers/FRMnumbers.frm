VERSION 5.00
Begin VB.Form FRMnumbers 
   Caption         =   "Allow Numbers ONLY!"
   ClientHeight    =   675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   675
   ScaleWidth      =   4125
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Limit only numbers in textbox"
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4095
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3855
      End
   End
End
Attribute VB_Name = "FRMnumbers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 47 Or KeyAscii = 48 Or KeyAscii = 49 Or KeyAscii = 50 Or KeyAscii = 51 Or KeyAscii = 52 Or KeyAscii = 53 Or KeyAscii = 54 Or KeyAscii = 55 Or KeyAscii = 56 Or KeyAscii = 57 Or KeyAscii = 8 Then
        Else: KeyAscii = 0 'sets it to 0 which tells computer nothing was typed so nothing shows
    End If
    
End Sub
