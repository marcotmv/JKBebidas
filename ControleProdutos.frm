VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ControleProdutos 
   Caption         =   "Controle de Estoque"
   ClientHeight    =   7875
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6600
   OleObjectBlob   =   "ControleProdutos.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ControleProdutos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub caixaAddProduto_Change()

caixaAddID = ""

End Sub

Private Sub CommandButton4_Click()

If caixaAddID = "" Then
    MsgBox ("Favor selecionar o produto a ser excluído!")
    Exit Sub
End If

Linha = Sheets("Controle_de_Produtos").Range("A:A").Find(caixaAddID.Value).Row
Sheets("Controle_de_Produtos").Range(Linha & ":" & Linha).Delete

caixaAddProduto.Value = ""
caixaAddCusto.Value = ""
caixaAddPrecoDeVenda.Value = ""
caixaAddID.Value = ""

MsgBox ("Produto removido com sucesso do estoque da JK Bebidas!")
End Sub

Private Sub Label1_Click()

End Sub

Private Sub CommandButton3_Click()

If caixaAddProduto = "" Then
    MsgBox ("Favor preencha o nome do produto a ser cadastrado!")
    Exit Sub
End If

If caixaAddCusto = "" Then
    MsgBox ("Favor preencha o custo do produto a ser cadastrado!")
    Exit Sub
End If

If caixaAddPrecoDeVenda = "" Then
    MsgBox ("Favor preencha o preço de venda do produto a ser cadastrado!")
    Exit Sub
End If

Linha = Sheets("Controle_de_Produtos").Range("A1048576").End(xlUp).Row + 1

Sheets("Controle_de_Produtos").Cells(Linha, 1).Value = WorksheetFunction.Max(Sheets("Controle_de_Produtos").Range("A:A")) + 1
Sheets("Controle_de_Produtos").Cells(Linha, 2).Value = caixaAddProduto.Value
Sheets("Controle_de_Produtos").Cells(Linha, 3).Value = caixaAddCusto.Value + 0
Sheets("Controle_de_Produtos").Cells(Linha, 4).Value = caixaAddPrecoDeVenda.Value + 0

caixaAddProduto.Value = ""
caixaAddCusto.Value = ""
caixaAddPrecoDeVenda.Value = ""

MsgBox ("Produto cadastrado com sucesso!")

Call mostraProdutos

End Sub

Private Sub Label2_Click()

End Sub

Private Sub Label4_Click()

End Sub

Private Sub Label9_Click()

End Sub

Private Sub listaProdutos_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

caixaAddProduto.Value = listaProdutos.List(listaProdutos.ListIndex, 1)
caixaAddCusto.Value = listaProdutos.List(listaProdutos.ListIndex, 2)
caixaAddPrecoDeVenda.Value = listaProdutos.List(listaProdutos.ListIndex, 3)
caixaAddID.Value = listaProdutos.List(listaProdutos.ListIndex, 0)

End Sub

Private Sub TextBox4_Change()

End Sub

Sub mostraProdutos()

Linha = Sheets("Controle_de_Produtos").Range("A1048576").End(xlUp).Row

If Linha = 1 Then Linha = 2

ControleProdutos.listaProdutos.ColumnCount = 4
ControleProdutos.listaProdutos.ColumnHeads = True
ControleProdutos.listaProdutos.ColumnWidths = "40;150;55;55"
ControleProdutos.listaProdutos.RowSource = "Controle_de_Produtos!A2:D" & Linha
End Sub

Private Sub UserForm_Initialize()

Call mostraProdutos

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

Call FormControleEstoque.atualizaCaixaListaDeMovimentacoes
Call FormControleEstoque.Atualiza_Produtos

End Sub
