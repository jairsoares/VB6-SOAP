VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Teste Integracao SOC"
   ClientHeight    =   4410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12150
   LinkTopic       =   "Form1"
   ScaleHeight     =   4410
   ScaleWidth      =   12150
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtStatus 
      Height          =   4215
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Form1.frx":0000
      Top             =   120
      Width           =   10335
   End
   Begin VB.CommandButton btnConsultar 
      Caption         =   "Consultar WS"
      Height          =   615
      Left            =   10560
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnConsultar_Click()
x = ws_SOC
End Sub
