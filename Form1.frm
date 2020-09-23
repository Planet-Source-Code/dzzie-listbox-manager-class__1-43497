VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listbox Extender Class Demo"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   5310
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLoadfromFile 
      Caption         =   "Load From File"
      Height          =   375
      Left            =   3900
      TabIndex        =   6
      Top             =   1440
      Width           =   1275
   End
   Begin VB.CommandButton cmdFilterList 
      Caption         =   "Filter List"
      Height          =   315
      Left            =   3900
      TabIndex        =   5
      Top             =   1080
      Width           =   1275
   End
   Begin VB.CommandButton cmdListToArray 
      Caption         =   "List to Array"
      Height          =   315
      Left            =   3900
      TabIndex        =   4
      Top             =   720
      Width           =   1275
   End
   Begin VB.CommandButton cmdEditSelected 
      Caption         =   "Edit Item"
      Height          =   315
      Left            =   3900
      TabIndex        =   3
      Top             =   360
      Width           =   1275
   End
   Begin VB.CommandButton cmdDeleteItem 
      Caption         =   "Delete  Item"
      Height          =   315
      Left            =   3900
      TabIndex        =   2
      Top             =   0
      Width           =   1275
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "More Stuff at :"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1980
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "http://sandSprite.com"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   1380
      TabIndex        =   1
      Top             =   1920
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

' This is a quick demo of a small class that will make your dealings
' with listboxes a little bit nicer.
'
' This demo doesnt cover all teh features of the class, but it should be
' enough to make you intrested enough to check it out in the obj browser :P
'
' Anyway, these macros streamline your code nicely, and will make it more it
' more readable and efficient to create.
'
' enjoy.
'
' -dzzie
'
' http://sandsprite.com


'create variable for class
Private clsList As clsListBox
Attribute clsList.VB_VarHelpID = -1



Private Sub Form_Load()

    'create an instance of the class
    Set clsList = New clsListBox

    'tell the class which form object to work on
    clsList.ListObject = List1

    clsList.LoadDelimitedString "Dave,George,Fred,Bill,Vivek,Marsha,Marsha,Marsha,And Finally Jan", ","
    
End Sub


Private Sub cmdDeleteItem_Click()
    
    If Not clsList.EnsureSelection Then Exit Sub
    
    clsList.Remove clsList.SelectedIndex
    
End Sub


Private Sub cmdEditSelected_Click()
    Dim s As String
    
    If Not clsList.EnsureSelection Then Exit Sub
    
    s = InputBox("Enter the new val", , clsList.SelectedText)
    If Len(s) > 0 Then
        clsList.SelectedText = s
    End If
    
End Sub


Private Sub cmdListToArray_Click()
    MsgBox Join(clsList.GetListToArray, vbCrLf)
End Sub


Private Sub cmdFilterList_Click()
    Dim s As String
    
    s = InputBox("Enter filter match expression", , "*av*")
    If Len(s) > 0 Then clsList.FilterList s
    
End Sub


Private Sub cmdLoadfromFile_Click()
    clsList.LoadFile App.path & "\list.txt", vbCrLf
End Sub




Private Sub Label2_Click()
     ShellExecute Me.hwnd, vbNullString, "http://sandsprite.com/", vbNullString, "C:\", 1
End Sub
