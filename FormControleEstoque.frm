VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormControleEstoque 
   Caption         =   "Controle de Estoque"
   ClientHeight    =   9210.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15705
   OleObjectBlob   =   "FormControleEstoque.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormControleEstoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub atualizaCaixaListaDeMovimentacoes()

Sheets("Estoque").Cells.Clear
Sheets("Controle_de_Produtos").Range("B:B").Copy Sheets("Estoque").Range("A1")

Sheets("Estoque").Range("B1").Value = "Compras"
Sheets("Estoque").Range("C1").Value = "Vendas"
Sheets("Estoque").Range("D1").Value = "Estoque"
Sheets("Estoque").Range("E1").Value = "ID"

Linha = Sheets("Estoque").Range("A1048576").End(xlUp).Row

If Linha > 1 Then
    Sheets("Estoque").Range("B2").FormulaLocal = "=SOMASES(Compras_e_Vendas!C:C;Compras_e_Vendas!B:B;A2;Compras_e_Vendas!D:D;""Compra"")"
    Sheets("Estoque").Range("C2").FormulaLocal = "=SOMASES(Compras_e_Vendas!C:C;Compras_e_Vendas!B:B;A2;Compras_e_Vendas!D:D;""Venda"")"
    Sheets("Estoque").Range("D2").FormulaLocal = "=B2-C2"
    Sheets("Estoque").Range("E2").FormulaLocal = "=PROCV(A2;Controle_de_Produtos!B:F;5;0)"
    If Linha > 2 Then
        Sheets("Estoque").Range("B2:E" & Linha).FillDown
    End If
    Sheets("Estoque").Calculate
End If

Sheets("Estoque").UsedRange.Copy
Sheets("Estoque").UsedRange.PasteSpecial xlPasteValues
Application.CutCopyMode = False

Sheets("Box_Estoque").Cells.Clear
Sheets("Estoque").AutoFilterMode = False

If caixaProdutoDisponivel.Value <> "" Then
    Sheets("Estoque").UsedRange.AutoFilter 1, "*" & caixaProdutoDisponivel.Value & "*"
End If

Sheets("Estoque").UsedRange.Copy Sheets("Box_Estoque").Range("A1")
Sheets("Estoque").AutoFilterMode = False

Linha = Sheets("Box_Estoque").Range("A1048576").End(xlUp).Row

If Linha = 1 Then Linha = 2

FormControleEstoque.caixaListagemEstoque.ColumnCount = 5
FormControleEstoque.caixaListagemEstoque.ColumnHeads = True
FormControleEstoque.caixaListagemEstoque.ColumnWidths = "140;47;37;37;10"
FormControleEstoque.caixaListagemEstoque.RowSource = "Box_Estoque!A2:E" & Linha

End Sub

Sub atualizaListaDeMovimentacoes()

Sheets("Compras_e_Vendas").AutoFilterMode = False

If botaoCompras.Value = True Then
    Sheets("Compras_e_Vendas").UsedRange.AutoFilter 4, "Compra"
ElseIf botaoVendas.Value = True Then
    Sheets("Compras_e_Vendas").UsedRange.AutoFilter 4, "Venda"
End If

Sheets("Box_Compras_e_Vendas").UsedRange.Clear
Sheets("Compras_e_Vendas").UsedRange.Copy
Sheets("Box_Compras_e_Vendas").Range("a1").PasteSpecial

Sheets("Compras_e_Vendas").AutoFilterMode = False

Linha = Sheets("Box_Compras_e_Vendas").Range("A1048576").End(xlUp).Row

If Linha = 1 Then Linha = 2

FormControleEstoque.caixaListagemTransacoes.ColumnCount = 6
FormControleEstoque.caixaListagemTransacoes.ColumnHeads = True
FormControleEstoque.caixaListagemTransacoes.ColumnWidths = "0;200;55;45;45;25"
FormControleEstoque.caixaListagemTransacoes.RowSource = "Box_Compras_e_Vendas!A2:F" & Linha

End Sub

Private Sub botaoCompras_Click()

Call atualizaListaDeMovimentacoes

End Sub

Private Sub botaoProcurar_Click()

Call atualizaCaixaListaDeMovimentacoes

End Sub

Private Sub botaoTodas_Click()

Call atualizaListaDeMovimentacoes

End Sub

Private Sub botaoVendas_Click()

Call atualizaListaDeMovimentacoes

End Sub

Private Sub buscaRapida2_Exit(ByVal Cancel As MSForms.ReturnBoolean)

Dim Linha As Double, Ver As Double
Dim plan As String
Dim C As Variant

plan = Planilha1.Name

Ver = WorksheetFunction.CountIf(Planilha1.Range("A:A"), buscaRapida2.Text)

If Ver = "1" Then

    With Worksheets(plan).Range("A:A")
    
        Set C = .Find(buscaRapida2.Text, LookIn:=xlValues, Lookat:=xlWhole)
           
        If Not C Is Nothing Then
            
            Linha = C.Row
            
            With Planilha1
                caixaProduto.Value = .Cells(Linha, 2).Value
                
            End With
                        
        End If
    
    End With
    
Set C = Nothing

End If

End Sub

Private Sub caixaData_Change()

caixaID.Value = ""

End Sub

Private Sub caixaListagemTransacoes_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

caixaProduto.Value = caixaListagemTransacoes.List(caixaListagemTransacoes.ListIndex, 1)
caixaQuantidade.Value = caixaListagemTransacoes.List(caixaListagemTransacoes.ListIndex, 2)
caixaTipo.Value = caixaListagemTransacoes.List(caixaListagemTransacoes.ListIndex, 3)
caixaValor.Value = Format(caixaListagemTransacoes.List(caixaListagemTransacoes.ListIndex, 4), "R$ #,##0.00")
caixaData.Value = CDate(caixaListagemTransacoes.List(caixaListagemTransacoes.ListIndex, 5))
caixaID.Value = caixaListagemTransacoes.List(caixaListagemTransacoes.ListIndex, 0)

End Sub

Private Sub caixaProduto_Change()

If caixaTipo.Value = "" Then
    caixaValor.Value = ""
    Exit Sub
End If

If caixaTipo.Value = "Compra" Then
    caixaValor.Value = Format(Sheets("Controle_de_Produtos").Range("B:B").Find(caixaProduto.Value).Offset(0, 1).Value, "R$ #,##0.00")
Else
    caixaValor.Value = Format(Sheets("Controle_de_Produtos").Range("B:B").Find(caixaProduto.Value).Offset(0, 2).Value, "R$ #,##0.00")
End If

caixaID.Value = ""

End Sub

Private Sub caixaQuantidade_Change()

caixaID.Value = ""

End Sub

Private Sub caixaTipo_Change()

If caixaTipo.Value = "" Then
    caixaValor.Value = ""
    Exit Sub
End If

If caixaTipo.Value = "Compra" Then
    caixaValor.Value = Format(Sheets("Controle_de_Produtos").Range("B:B").Find(caixaProduto.Value).Offset(0, 1).Value, "R$ #,##0.00")
Else
    caixaValor.Value = Format(Sheets("Controle_de_Produtos").Range("B:B").Find(caixaProduto.Value).Offset(0, 2).Value, "R$ #,##0.00")
End If

caixaID.Value = ""

End Sub

Private Sub caixaValor_Change()

caixaID.Value = ""

End Sub

Private Sub CommandButton1_Click()

ControleProdutos.Show

End Sub

Private Sub CommandButton2_Click()

ThisWorkbook.Save

MsgBox ("Planilha salva com sucesso!")

End Sub

Private Sub CommandButton3_Click()

If caixaProduto = "" Then
    MsgBox ("Favor preencha o nome do produto a ser movimentado!")
    Exit Sub
End If

If caixaQuantidade = "" Then
    MsgBox ("Favor preencha a quantidade do produto a ser movimentada!")
    Exit Sub
End If

If caixaTipo = "" Then
    MsgBox ("Favor preencha o tipo de transação a ser movimentada!")
    Exit Sub
End If

If caixaValor = "" Then
    MsgBox ("Favor preencha o valor unitário a ser movimentada!")
    Exit Sub
End If

If caixaData = "" Then
    MsgBox ("Favor preencha a data que a transação aconteceu!")
    Exit Sub
End If

Linha = Sheets("Compras_e_Vendas").Range("A1048576").End(xlUp).Row + 1

Sheets("Compras_e_Vendas").Cells(Linha, 1).Value = WorksheetFunction.Max(Sheets("Compras_e_Vendas").Range("A:A")) + 1
Sheets("Compras_e_Vendas").Cells(Linha, 2).Value = caixaProduto.Value
Sheets("Compras_e_Vendas").Cells(Linha, 3).Value = caixaQuantidade.Value + 0
Sheets("Compras_e_Vendas").Cells(Linha, 4).Value = caixaTipo.Value
Sheets("Compras_e_Vendas").Cells(Linha, 5).Value = caixaValor.Value + 0
Sheets("Compras_e_Vendas").Cells(Linha, 6).Value = CDate(caixaData.Value)

caixaProduto.Value = ""
caixaQuantidade.Value = ""
caixaTipo.Value = ""
caixaValor.Value = ""
caixaData.Value = Format(Date, "dd/mm/yyyy")

MsgBox ("Movimentação adicionada com sucesso!")

Call atualizaCaixaListaDeMovimentacoes
Call atualizaListaDeMovimentacoes


End Sub

Private Sub CommandButton4_Click()

If caixaID.Value = "" Then
    MsgBox ("Favor selecionar a transação que deverá ser excluída")
    
End If

Linha = Sheets("Compras_e_Vendas").Range("A:A").Find(caixaID.Value).Row
Sheets("Compras_e_Vendas").Range(Linha & ":" & Linha).Delete

caixaProduto.Value = ""
caixaQuantidade.Value = ""
caixaTipo.Value = ""
caixaValor.Value = ""
caixaData.Value = Format(Date, "dd/mm/yyyy")
caixaID.Value = ""

Call atualizaCaixaListaDeMovimentacoes
Call atualizaListaDeMovimentacoes

MsgBox ("Movimentação removida com sucesso do estoque da JK Bebidas!")

End Sub

Private Sub Frame2_Click()

End Sub

Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub


Private Sub UserForm_Initialize()

caixaTipo.AddItem "Compra"
caixaTipo.AddItem "Venda"

caixaData.Value = Format(Date, "dd/mm/yyyy")

Call Atualiza_Produtos
Call atualizaListaDeMovimentacoes
Call atualizaCaixaListaDeMovimentacoes

End Sub

Sub Atualiza_Produtos()

Linha = Sheets("Controle_de_Produtos").Range("A1048576").End(xlUp).Row
caixaProduto.RowSource = "Controle_de_Produtos!B2:B" & Linha

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

Application.DisplayFullScreen = False
ActiveWindow.DisplayGridlines = True
ActiveWindow.DisplayHeadings = True
ActiveWindow.DisplayWorkbookTabs = True
Application.DisplayFormulaBar = True

Application.ScreenUpdating = True

End Sub
