VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Counting Word"
   ClientHeight    =   1455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   ScaleHeight     =   1455
   ScaleWidth      =   6660
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CountOfWord 
      Caption         =   "Count Word"
      Height          =   375
      Left            =   2663
      TabIndex        =   4
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox txtNumber 
      Height          =   285
      Left            =   2040
      TabIndex        =   2
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox txtString 
      Height          =   285
      Left            =   2040
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Number Of Word :"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Enter String :"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author :       Ryan Nilo Ybañez
'Description :  Count of word
'Date Created:  November 3,2006

Private Sub CountOfWord_Click()
    txtNumber = NUMBEROFWORD(Trim(txtString))
End Sub
Private Function NUMBEROFWORD(str As String) As Integer
    Dim i, c
    Dim size, numWord, strStart, pos As Integer
    
    NUMBEROFWORD = 1
    
    strStart = 1
    numWord = 1
    
    size = Len(str)

    If size <> 0 Then
    
        For i = 0 To size
            c = Mid(str, strStart, 1)
            If c = " " Then
                pos = InStr(strStart, str, " ")
                pos = pos - 1
                
                c = Mid(str, pos, 1)
                If c <> " " Then
                    NUMBEROFWORD = NUMBEROFWORD + 1
                    strStart = pos + 1
                End If
            End If
            strStart = strStart + 1
        Next i
    Else
        NUMBEROFWORD = 0
    End If
End Function
