VERSION 5.00
Begin VB.Form frm_main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Page Balloon"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2535
   Icon            =   "frm_main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   2535
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fra_options 
      Caption         =   "Balloon Options"
      Height          =   1035
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   2415
      Begin VB.CheckBox chk_wrap 
         Caption         =   "Wrap edges of screen"
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   2175
      End
      Begin VB.CheckBox chk_ontop 
         Caption         =   "Keep balloons ontop"
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.CommandButton cmd_new 
      Caption         =   "Give me a balloon"
      Height          =   615
      Left            =   60
      TabIndex        =   0
      Top             =   1140
      Width           =   2415
   End
End
Attribute VB_Name = "frm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*****************************************
'CREATE A NEW BALLOON
'*****************************************
Private Sub cmd_new_Click()
    If Not busy Then
        Dim Newballoon As Form
        Set Newballoon = New frm_balloon
        Load Newballoon
    End If
End Sub

'*****************************************
'UNLOAD THE MAIN FORM AND RAISE FLAG TO
'NOTIFY ALL OTHER FORMS
'*****************************************
Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
