VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4815
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   4815
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CD 
      Left            =   1800
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "ini"
      DialogTitle     =   "Please, select what file you want to open:"
      Filter          =   "INI Files (*.ini)|*.ini|All Files (*.*)|*.*"
      FontName        =   "Tahoma"
   End
   Begin VB.CommandButton cmdOpenFile 
      Caption         =   "&Open File"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   3120
      Width           =   1455
   End
   Begin ComctlLib.TreeView TV 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   5106
      _Version        =   327682
      Indentation     =   706
      LineStyle       =   1
      Style           =   6
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ARINI As New ARINIManager

Private Sub cmdOpenFile_Click()
    CD.ShowOpen
    If CD.FileName = "" Then
        Exit Sub
    Else
        ARINI.INIFile = CD.FileName
        TV.Nodes.Clear
        Dim N As Long, M As Long
        For N = 1 To ARINI.Sections.Count
            TV.Nodes.Add , , ARINI.Sections(N).Name, ARINI.Sections(N).Name
            For M = 1 To ARINI.Sections(N).Values.Count
                TV.Nodes.Add ARINI.Sections(N).Name, tvwChild, , ARINI.Sections(N).Values(M).Name & " = " & ARINI.Sections(N).Values(M).Value
            Next
        Next
    End If
End Sub
