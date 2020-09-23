VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Prompt"
   ClientHeight    =   1335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6390
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   1335
   ScaleWidth      =   6390
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExec 
      Caption         =   "Exec"
      Default         =   -1  'True
      Height          =   300
      Left            =   5400
      TabIndex        =   2
      Top             =   900
      Width           =   855
   End
   Begin VB.ComboBox ComboSys 
      Height          =   315
      ItemData        =   "main.frx":0000
      Left            =   120
      List            =   "main.frx":000A
      TabIndex        =   1
      Text            =   "Cmd.exe(WinNT-Win2000)"
      Top             =   300
      Width           =   2295
   End
   Begin VB.TextBox txtPrompt 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "type a ms-dos command here"
      Top             =   900
      Width           =   5175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "System:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   80
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Prompt:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   680
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************
'Free Open Source :-)))
'Coded by [Jarret*Scream]
'****************************************
'Ciao, ti faccio l'esempio per sfruttare le righe di comando MS-DOS
'sotto qualsiasi tipo di sistema Windows.
'Sotto Win2000 e NT si usa Cdm.exe e sotto
'Win9x si usa Command.com.
'Il comboBox gestisce le variabili, cioè il tipo
'di sistema sotto cui stai operando (ma non in automatico
'quindi la scelta devi farla tu).
'In base a quello che selezioni utilizzi
'un diverso prompt di comandi con tutti i loro diversi parametri.

Private Sub cmdExec_Click()
On Error Resume Next
Dim comando As String
comando = txtPrompt.Text
'qui dici che quello che scriverai
'dentro a txtPrompt è una stringa da eseguire come comando
'dos all'interno della shell selezionata

Select Case ComboSys
    Case Is = "Command.com(Win9x)"
        If ComboSys.Text = "Command.com(Win9x)" Then
    Shell "Command.com /c" & comando, vbNormalFocus
    'this commnd line display the shell of MS-DOS for win9x
        End If
    Case Is = "Cmd.exe(WinNT-Win2000)"
        If ComboSys.Text = "Cmd.exe(WinNT-Win2000)" Then
    Shell "Cmd.exe /k" & comando, vbNormalFocus
    'this commnd line display the shell of MS-DOS for Win NT and Win2yK
        End If
    End Select
End Sub
'Nota Bene:
'Sotto Win9x L'esecuzione della shell command.com potrebbe
'essere non visualizzata (è da testare)
'Sotto Win2000 e NT, dopo cmd.exe se metti /k la shell rimane aperta
'se metti l'opzione /c la shell dopo il comando si richiude.
'se vuoi fare una cosa giusta e cioè eseguire il comando senza andare
'a cliccare con il mouse sopra al pulsante exec, ma premendo
'semplicemente INVIO, devi sempre impostare il cmdExec (come in questo caso)
'come DEFAULT = True
'Cia[X]
'Thanx very well to Planet-source-code.com!!!
'Mail: jarret.scream@libero.it


