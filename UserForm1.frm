VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

Dim Waux As Worksheet
Set Waux = Worksheets("AUX")

V_QUEMINCLUI = ComboBox1.Value
Waux.Cells(4, 3) = V_QUEMINCLUI

Unload Me
UserForm3.Show


End Sub

Private Sub UserForm_Initialize()

Dim Waux As Worksheet
Set Waux = Worksheets("AUX")

ComboBox1.AddItem "------"
ComboBox1.AddItem "Maximo"
ComboBox1.AddItem "Rita"
ComboBox1.AddItem "Maria"
ComboBox1.AddItem "Leonardo"
ComboBox1.AddItem "Leocir"
ComboBox1.AddItem "Leandro"
ComboBox1.AddItem "Jackson"
ComboBox1.AddItem "Estagiario"

ComboBox1.Text = ComboBox1.List(0)
ComboBox1.Text = Waux.Cells(4, 3)


End Sub



