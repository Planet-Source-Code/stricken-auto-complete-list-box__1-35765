VERSION 5.00
Begin VB.Form frmCompleter 
   Caption         =   "Auto Complete"
   ClientHeight    =   5475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   2175
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstNames 
      Height          =   4740
      ItemData        =   "autocomplete.frx":0000
      Left            =   120
      List            =   "autocomplete.frx":0097
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox txtTestComplete 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmCompleter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public KeyPressed As Integer
    
Private Sub Form_Load()
    lstNames.ListIndex = -1
End Sub

Private Sub lstNames_Click()
    txtTestComplete = lstNames
End Sub

Private Sub txtTestComplete_Change()
   
    If KeyPressed = 8 Or KeyPressed = vbKeyDelete Then
        If txtTestComplete.Text = "" Then
            lstNames.ListIndex = -1
        Else
        End If
    Else
        Dim i As Integer
        Dim strEntry As String
        Dim strStored As String
        Dim placeholder As Integer
                
        strEntry = txtTestComplete.Text
    
        If strEntry <> "" Then
            For i = 0 To lstNames.ListCount - 1
                
                strStored = lstNames.List(i)
                If LCase(strEntry) = LCase(Left(strStored, Len(strEntry))) Then
                    txtTestComplete.Text = strStored
                    txtTestComplete.SelStart = Len(strEntry)
                    txtTestComplete.SelLength = Len(strStored) - Len(strEntry)
                    Exit For
                End If
            Next
                placeholder = lstNames.ListIndex
                
                If i = lstNames.ListCount Then
                    lstNames.ListIndex = placeholder
                Else
                    lstNames.ListIndex = i
                End If
        End If
    End If
End Sub

Private Sub txtTestComplete_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyPressed = KeyCode
End Sub

Private Sub txtTestComplete_KeyPress(KeyAscii As Integer)
    KeyPressed = KeyAscii
End Sub
