' Anderson rocha de sousa



Imports System.Globalization
Imports System.Net.Mail
Imports System
Imports System.Net
Imports System.Web
Imports System.Data.SqlClient
Imports Excel = Microsoft.Office.Interop.Excel
Imports relatorio_bit.clDALSQL

Module Module1

    Public obDal As New clDALSQL
    Public DT As New DataTable
    Public DT2 As New DataTable
    Public xlApp As New Excel.Application
    Public xlWorkBook As Excel.Workbook
    Public xlworkSheet As Excel.Worksheet
    Public Dbl_Linha As Double



    Sub Main()

        obDal.Ambiente = AmbienteExecucao.Producao
        Buscar_dados()

    End Sub

    Sub Buscar_dados()

        'DECLARA TODAS AS VARIAVEIS PARA UTILIZAR NO EXCEL DA PRIMEIRA PARCELA 

        Dim COD_ESTIPULANTE As String = ""
        Dim NUM_PROPOSTA As String = ""
        Dim NUM_PARCELA As String = ""
        Dim DAT_RETORNO_AGENTE As String = ""
        Dim COD_AGENTE As String = ""
        Dim COD_RETORNO As String = ""
        Dim Diretorio As String = "caminho que sera salvo o arquivo, com nome do arquivo"


        'BUSCAR DADOS NO BANCO - UTILIZANDO PROCEDURE SEM PARAMETROS 
        obDal.ClearParameters()
        DT2.Dispose()
        DT2 = obDal.RetornaTabela("Nome da procedure", "Nome do banco")


        'ABRIR PLANILHA DE RETORNO 
        xlWorkBook = xlApp.Workbooks.Open("caminho do excel formatado")

        'SELECIONAR ABA
        xlworkSheet = xlWorkBook.Sheets("Arrecadação BIT")

        Dbl_Linha = 2

        'fazer o for para prencher varial 

        For Each li In DT2.Rows
            'LIMPANDO AS VARIAVEIS

            COD_ESTIPULANTE = ""
            NUM_PROPOSTA = ""
            NUM_PARCELA = ""
            DAT_RETORNO_AGENTE = ""
            COD_AGENTE = ""
            COD_RETORNO = ""

            'INSERINDO VALOR AS VARIAVEIS

            COD_ESTIPULANTE = li("COD_ESTIPULANTE").ToString
            NUM_PROPOSTA = li("NUM_PROPOSTA").ToString
            NUM_PARCELA = li("NUM_PARCELA").ToString
            DAT_RETORNO_AGENTE = li("DAT_RETORNO_AGENTE").ToString
            COD_AGENTE = li("COD_AGENTE").ToString
            COD_RETORNO = li("COD_RETORNO").ToString


            'INSERIR OS DADoS NA PLANILHA RETORNO
            With xlworkSheet

                .Range("A" & Dbl_Linha).Value = COD_ESTIPULANTE
                .Range("B" & Dbl_Linha).Value = NUM_PROPOSTA
                .Range("C" & Dbl_Linha).Value = NUM_PARCELA
                .Range("D" & Dbl_Linha).Value = DAT_RETORNO_AGENTE
                .Range("E" & Dbl_Linha).Value = COD_AGENTE
                .Range("F" & Dbl_Linha).Value = COD_RETORNO

                'PULAR LINHA
                Dbl_Linha = Dbl_Linha + 1

            End With

        Next

        'SALVAR PLANILHA NO DIRETORIO 
        xlworkSheet.SaveAs(Diretorio & Format(Now, "yyyy_MM_dd_ss") & ".xlsx")

        'LIMPAR MEMORIA
        xlworkSheet.ClearArrows()
        xlWorkBook.Close()
        xlApp.Quit()


    End Sub


End Module










