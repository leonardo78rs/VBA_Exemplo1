VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "UserForm3"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14910
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CheckBox5_Click()
If CheckBox5 = False Then
   TextBox2 = Empty
   TextBox3 = Empty
End If

End Sub

Private Sub CheckBox6_Click()
If CheckBox6 = False Then TextBox4 = Empty
CommandButton3_Click

End Sub

Private Sub ComboBox3_Change()
CheckBox5 = False
CheckBox6 = False
TextBox2 = Empty
TextBox3 = Empty
TextBox4 = Empty

Call MONTAR_LISTBOX1(2)


End Sub
Private Sub ComboBox4_Change()
CheckBox5 = False
CheckBox6 = False
TextBox2 = Empty
TextBox3 = Empty
TextBox4 = Empty

Call MONTAR_LISTBOX1(3)
End Sub

Private Sub CommandButton1_Click()
Unload Me
UserForm1.Show

End Sub



Private Sub CommandButton4_click()
CheckBox5 = False
CheckBox6 = False
TextBox2 = Empty
TextBox3 = Empty
TextBox4 = Empty
ComboBox4.Value = TextBox5

Call MONTAR_LISTBOX1(3)

End Sub

Private Sub CommandButton6_Click()

Unload Me
UserForm4.Show

End Sub

Private Sub UserForm_Initialize()

With Worksheets("AUX")

If .Cells(4, 3) = "Maximo" Or .Cells(4, 3) = "Leonardo" Or Cells(4, 3) = "Master" Then
   ComboBox4.Visible = True
   Frame3.Caption = "Pesquisa por Pessoa"
   CommandButton4.Visible = False
   TextBox5.Visible = False
   checkbox4.Visible = True
Else
   Frame3.Caption = "Pesquisa Meus Itens"
   ComboBox4.Visible = False
   TextBox5.Visible = True
   TextBox5.Value = .Cells(4, 3)
   TextBox5.Locked = True
   TextBox5.BackColor = &H80000000
   CommandButton4.Visible = True
   checkbox4.Visible = False
End If

End With

ComboBox4.AddItem "------"
ComboBox4.AddItem "Maximo"
ComboBox4.AddItem "Rita"
ComboBox4.AddItem "Maria"
ComboBox4.AddItem "Leonardo"
ComboBox4.AddItem "Leocir"
ComboBox4.AddItem "Leandro"
ComboBox4.AddItem "Jackson"
ComboBox4.AddItem "Estagiario"

ComboBox3.AddItem "------"
ComboBox3.AddItem "CONTABIL"
ComboBox3.AddItem "FISCAL"
ComboBox3.AddItem "PESSOAL"
ComboBox3.AddItem "REPARTI��ES"
ComboBox3.AddItem "DECLARA��ES"
ComboBox3.AddItem "INFORMATICA"
ComboBox3.AddItem "OUTROS"
ComboBox3.AddItem "COBRAN�A**"

End Sub

Private Sub CommandButton2_Click()
'Bot�o da data
If (TextBox2.Value = Empty And TextBox3.Value = Empty) Then
   CheckBox5.Value = False
Else
   CheckBox5.Value = True
End If

Call MONTAR_LISTBOX1(1)

End Sub

Function MONTAR_LISTBOX1(x As Integer)
If x = 1 Then ComboBox4 = Empty And ComboBox3 = Empty
If x = 2 Then ComboBox4 = Empty
If x = 3 Then ComboBox3 = Empty


Dim campo As String
Dim tanto
Dim i, linha As Integer

With Me.ListBox1
    .Clear
    .ColumnHeads = True
    .ColumnWidths = "40;20;40;40;28;70;50;50;28;200;100;70"
    .ColumnCount = 15
End With
linha = -1

With Worksheets("BDE")
For i = 12 To 42
    'If campo1 = "fim" Then Exit For
    
    campo1 = Trim(CStr(.Cells(i, 1)))  '4    num lcto
    campo2 = Trim(CStr(.Cells(i, 2)))  '4    emp
    campo3 = .Cells(i, 3)              '5    nome
    campo4 = .Cells(i, 4)              '12   quem incl
    campo5 = CStr(.Cells(i, 5))        '10   data incl
    campo6 = .Cells(i, 6)              '12    origem evento
    campo7 = .Cells(i, 7)              '13
    campo8 = .Cells(i, 8)              '14
    campo9 = .Cells(i, 9).Value
    campo10 = .Cells(i, 10).Value
    campo11 = .Cells(i, 11).Value
  
    If (campo5 > TextBox2.Value And campo5 < TextBox3.Value) Or (TextBox2.Value = Empty And TextBox3.Value = Empty) Then
        If campo2 = TextBox4.Value Or CheckBox6 = False Then
            If x = 2 And campo7 <> ComboBox3.Value Then
            Else
                If x = 3 And campo8 <> ComboBox4.Value Then
                Else
                        linha = linha + 1
                        With Me.ListBox1
                                .AddItem campo1
                                .List(linha, 1) = campo11
                                .List(linha, 2) = campo2
                                .List(linha, 3) = campo3
                                .List(linha, 4) = campo4
                                .List(linha, 5) = campo5
                                .List(linha, 6) = campo6
                                .List(linha, 7) = campo7
                                .List(linha, 8) = campo8
                                .List(linha, 9) = campo9
                              '  .List(linha, 10) = campo10
                              '  .List(linha, 10) = campo11
                            End With
                  End If
                  
                End If
            End If
    End If
    
Next

End With

End Function

Private Sub CommandButton3_Click()
'pesquisa empresa

With Worksheets("BD")

If Val(TextBox4.Value) = Error Or TextBox4 = Empty Then
   Frame4.Caption = "Pesquisa Empresa (digite o n�mero)"
   TextBox4.Value = Empty
   CheckBox6.Value = False
  Else
   .Cells(2, 2).Value = Val(TextBox4.Value)
   Frame4.Caption = .Cells(2, 3).Value
   CheckBox6.Value = True
End If

End With

Call MONTAR_LISTBOX1(1)


End Sub

Private Sub ListBox1_Click()


'MsgBox (ListBox1.Value) 'd� a linha que est� selecionada
'MsgBox (ListBox1.ListCount) d� a quantidade de linhas
'MsgBox (ListBox1.ListIndex) 'diz qual o n� da linha atual
 



End Sub

Private Sub OptionButton1_Click()
OptionButton2 = False
OptionButton3 = False
OptionButton4 = False
End Sub
Private Sub OptionButton2_Click()
OptionButton1 = False
OptionButton3 = False
OptionButton4 = False
End Sub

Private Sub OptionButton3_Click()
OptionButton1 = False
OptionButton2 = False
OptionButton4 = False
End Sub
Private Sub OptionButton4_Click()
OptionButton1 = False
OptionButton2 = False
OptionButton3 = False
End Sub

Private Sub TextBox4_Change()
OptionButton1 = False
OptionButton2 = False
OptionButton3 = False
OptionButton4 = True
End Sub
