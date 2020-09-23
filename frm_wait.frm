VERSION 5.00
Begin VB.Form frm_wait 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Creating new balloon"
   ClientHeight    =   435
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3675
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   435
   ScaleWidth      =   3675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lbl_wait 
      Alignment       =   2  'Center
      Caption         =   "Creating Page Balloon..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3555
   End
End
Attribute VB_Name = "frm_wait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
