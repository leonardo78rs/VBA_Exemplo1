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



Private Sub CommandButton4_Click()
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
UserForm2.Show

End Sub


Private Sub UserForm_Initialize()

With Worksheets("AUX")

If .Cells(4, 3) = "Maximo" Or .Cells(4, 3) = "Leonardo" Or Cells(4, 3) = "Master" Then
   ComboBox4.Visible = True
   Frame3.Caption = "Pesquisa por Pessoa"
   CommandButton4.Visible = False
   TextBox5.Visible = False
   CheckBox4.Visible = True
Else
   Frame3.Caption = "Pesquisa Meus Itens"
   ComboBox4.Visible = False
   TextBox5.Visible = True
   TextBox5.Value = .Cells(4, 3)
   TextBox5.Locked = True
   TextBox5.BackColor = &H80000000
   CommandButton4.Visible = True
   CheckBox4.Visible = False
End If
.Cells(10, 12) = Empty         'zerando campo da consulta por numero de lancamento
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
ComboBox3.AddItem "REPARTIÇÕES"
ComboBox3.AddItem "DECLARAÇÕES"
ComboBox3.AddItem "INFORMATICA"
ComboBox3.AddItem "OUTROS"
If Worksheets("AUX").Cells(4, 4) = "permite" Then
   ComboBox3.AddItem "COBRANÇA**"
Else
   ComboBox3.AddItem "**********"
End If



End Sub

Private Sub CommandButton2_Click()
'Botão da data
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
     ' .ListStyle = fmListStyleOption   'com selecionador inicial
     '.BoundColumn = 3                  'define a coluna padrão, quando pedir list.value retorna o desta coluna
     
    .Clear
    .ColumnHeads = True
    .ColumnWidths = "40;20;40;40;28;70;50;50;28;200;100;70"
    .ColumnCount = 15
End With

linha = -1

With Worksheets("BDE")
For i = 12 To 77
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
                                '.List(linha, 10) = campo10
                                '.List(linha, 10) = campo11
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
   Frame4.Caption = "Pesquisa Empresa (digite o número)"
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
Dim Waux As Worksheet
Set Waux = Worksheets("AUX")

'MsgBox (ListBox1.Value) 'dá a linha que está selecionada
'MsgBox (ListBox1.ListCount) dá a quantidade de linhas
'MsgBox (ListBox1.ListIndex) 'diz qual o nº da linha atual

With ListBox1
 LIN = .ListIndex
 
 frase = "Lcto: " + .List(LIN, 0) + "    "
  If .List(LIN, 1) = 1 Then frase = frase + "RESOLVIDO" Else frase = frase + "PENDENTE"
   
 frase = frase + Chr(13) + .List(LIN, 2) + "-"
 frase = frase + .List(LIN, 3) + Chr(13)
 
 frase = frase + "Quem Incluiu: " + .List(LIN, 4) + Chr(13)
 frase = frase + .List(LIN, 5) + Chr(13)
 frase = frase + .List(LIN, 6) + Chr(13)
 frase = frase + .List(LIN, 7) + Chr(13)
 frase = frase + .List(LIN, 9) + Chr(13)
 
  x = MsgBox(frase, vbInformation)
 
  
  End With
  
 Waux.Cells(10, 12) = ListBox1.List(LIN, 0)
 Unload Me
 
 UserForm2.Show
 
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
