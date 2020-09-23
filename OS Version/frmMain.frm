VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Operating System"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   89
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   417
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   120
      ScaleHeight     =   71
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   399
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      Begin VB.Label lblServicePack 
         BackStyle       =   0  'Transparent
         Caption         =   "%SERVICE PACK%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1080
         TabIndex        =   3
         Top             =   600
         Width           =   4815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Operating System"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1080
         TabIndex        =   2
         Top             =   75
         Width           =   1740
      End
      Begin VB.Image Image1 
         Height          =   1080
         Left            =   0
         Picture         =   "frmMain.frx":08CA
         Top             =   0
         Width           =   1080
      End
      Begin VB.Label lblVersion 
         BackStyle       =   0  'Transparent
         Caption         =   "%VERSION%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1080
         TabIndex        =   1
         Top             =   360
         Width           =   4740
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OperatingSystem As AP_OperatingSystem

Private Sub Form_Load()

    If OperatingSystem Is Nothing Then
        Set OperatingSystem = New AP_OperatingSystem
    End If
    
    lblVersion = OperatingSystem.Platform_ProductDescription()
    If OperatingSystem.Platform_ServicePack() = "" Then
        lblServicePack = "Build: " & OperatingSystem.Platform_BuildNumber()
    Else
        lblServicePack = OperatingSystem.Platform_ServicePack() & " (Build: " & OperatingSystem.Platform_BuildNumber() & ")"
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Not OperatingSystem Is Nothing Then
        Set OperatingSystem = Nothing
    End If

End Sub
