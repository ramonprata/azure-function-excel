using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using AzureFunctionsDemo.Arquivos;

namespace AzureFunctionsDemo.Excel
{
  public static class ExcelUtil
  {
    public static MemoryStream ObterPlanilhaPorLista<TObject>(IEnumerable<TObject> lista, string tituloPlanilha, string subTitulo, string[] nomesAtributosExibicao, int tipoTabela)
    {
      var workbook = new HSSFWorkbook();
      var sheet = workbook.CreateSheet();

      // sheet = IncluirLogoVLI(workbook, sheet);

      /*Utilizar o mesmo objeto de linha ao criar uma nova célula da ---mesma--- linha.*/

      int linha = 0;
      RetornarLinhaBranca(workbook, sheet, linha);

      linha = 1;
      IRow linhaTituloVLI = RetornarLinhaBranca(workbook, sheet, linha);
      sheet = IncluirTextoVLI(workbook, sheet, linhaTituloVLI);

      linha = 2;
      IRow linhaTituloETotal = RetornarLinhaBranca(workbook, sheet, linha);
      sheet = IncluirTituloPlanilha(workbook, sheet, tituloPlanilha, linhaTituloETotal);

      linha = 3;
      IRow linhaSubTituloESubTotal = RetornarLinhaBranca(workbook, sheet, linha);
      sheet = IncluirSubTituloPlanilha(workbook, sheet, subTitulo, linhaSubTituloESubTotal);

      linha = 4;
      RetornarLinhaBranca(workbook, sheet, linha);

      sheet = IncluirCabecalhoPorNomePropriedadeObjeto(workbook, sheet, nomesAtributosExibicao);

      linha = 6;
      string linhaInicioSomatorio = linha.ToString();


      ICellStyle estiloPrimeiraColuna = RetornarEstiloCelulaFundoBranco(workbook);
      estiloPrimeiraColuna.BorderRight = BorderStyle.Thin;
      estiloPrimeiraColuna.RightBorderColor = HSSFColor.Black.Index;

      ICellStyle estiloUltimaColuna = RetornarEstiloCelulaFundoBranco(workbook);
      estiloUltimaColuna.BorderLeft = BorderStyle.Thin;
      estiloUltimaColuna.LeftBorderColor = HSSFColor.Black.Index;

      var estiloConteudo = RetornarEstiloCelulaConteudo(workbook);
      ICellStyle dateStyle = RetornarEstiloCelulaConteudo(workbook);

      List<PropertyInfo> nomesPropriedadesExibicao = RetornarPropriedadesOrdemExibicao<TObject>(nomesAtributosExibicao);

      foreach (TObject objeto in lista)
      {
        /* Não colocar o estilo da célula dentro da função de colocar conteúdo - Trava o desempenho e não funciona em todas as células */
        if (objeto != null)
        {
          sheet = IncluirLinhaPorConteudoDeObjeto(objeto, workbook, sheet, linha, estiloConteudo, dateStyle, estiloPrimeiraColuna, estiloUltimaColuna, nomesPropriedadesExibicao);
          sheet.GetRow(linha).Height = 300;
          linha++;
        }
      }

      ICellStyle estiloUltimaLinha = RetornarEstiloCelulaFundoBranco(workbook);
      estiloUltimaLinha.BorderTop = BorderStyle.Thin;
      estiloUltimaLinha.TopBorderColor = HSSFColor.Black.Index;

      IRow ultimaLinha = sheet.CreateRow(linha);

      for (int colunaUltimaLinha = 1; colunaUltimaLinha <= nomesAtributosExibicao.Length; colunaUltimaLinha++)
      {
        ICell celulaUltimaLinha = ultimaLinha.CreateCell(colunaUltimaLinha);
        celulaUltimaLinha.CellStyle = estiloUltimaLinha;
      }

      string linhaFimSomatorio = linha.ToString();

      sheet = InsereSomatorioSubSomatorio(workbook, sheet, linhaTituloETotal, linhaSubTituloESubTotal, linhaInicioSomatorio, linhaFimSomatorio, tipoTabela);

      AjustarLarguraColunas(sheet, nomesAtributosExibicao);

      MemoryStream output = new MemoryStream();
      workbook.Write(output);

      return output;
    }

    private static IRow RetornarLinhaBranca(HSSFWorkbook workbook, ISheet sheet, int linha)
    {
      IRow linhaPlanilhaBranca = sheet.CreateRow(linha);
      linhaPlanilhaBranca.RowStyle = RetornarEstiloCelulaFundoBranco(workbook);

      return linhaPlanilhaBranca;
    }

    private static ISheet IncluirCabecalhoPorNomePropriedadeObjeto(HSSFWorkbook workbook, ISheet sheet, string[] nomesAtributosExibicao)
    {
      var estiloCelulaCabecalhoCinza = RetornarEstiloCelulaCabecalhoCinza(workbook);

      int column = 0;

      var linhaTituloPropriedade = sheet.CreateRow(5);
      linhaTituloPropriedade.Height = 400;

      ICell celulaCabecalhoNomePropriedadesBrancoPrimeiro = linhaTituloPropriedade.CreateCell(column);
      celulaCabecalhoNomePropriedadesBrancoPrimeiro.CellStyle = RetornarEstiloCelulaFundoBranco(workbook);

      column++;

      foreach (string nomeAtributo in nomesAtributosExibicao)
      {
        ICell celulaCabecalhoConteudo = linhaTituloPropriedade.CreateCell(column);

        celulaCabecalhoConteudo.CellStyle = estiloCelulaCabecalhoCinza;

        celulaCabecalhoConteudo.SetCellValue(nomeAtributo);

        column++;
      }

      ICell celulaCabecalhoNomePropriedadesBrancoUltimo = linhaTituloPropriedade.CreateCell(column);
      celulaCabecalhoNomePropriedadesBrancoUltimo.CellStyle = RetornarEstiloCelulaFundoBranco(workbook);

      return sheet;
    }

    private static ISheet IncluirLinhaPorConteudoDeObjeto<TObject>(TObject objeto, HSSFWorkbook workbook, ISheet sheet, int row, ICellStyle estiloCelulaConteudo, ICellStyle dateStyle,
        ICellStyle estiloPrimeiraColuna, ICellStyle estiloUltimaColuna, List<PropertyInfo> propriedadesOrdemExibicao)
    {

      int column = 0;
      var linha = sheet.CreateRow(row);

      ICell celulaPrimeiraColuna = linha.CreateCell(column);
      celulaPrimeiraColuna.CellStyle = estiloPrimeiraColuna;

      column++;

      foreach (PropertyInfo prop in propriedadesOrdemExibicao)
      {
        var tipo = prop.PropertyType;
        var valor = prop.GetValue(objeto, null);

        ICell celulaConteudo = linha.CreateCell(column);
        celulaConteudo.CellStyle = estiloCelulaConteudo;

        if (valor != null)
        {
          if (tipo.Equals(typeof(DateTime)) || tipo.Equals(typeof(DateTime?)))
          {

            IDataFormat format = workbook.CreateDataFormat();
            dateStyle.DataFormat = format.GetFormat("dd/MM/yyyy");

            ICell celulaData = linha.CreateCell(column);
            celulaData.SetCellValue((DateTime)valor);
            celulaData.CellStyle = dateStyle;
          }

          else if (tipo.Equals(typeof(double)) || tipo.Equals(typeof(double?)))
          {
            ICell celulaDouble = linha.CreateCell(column); /* C irá interpretar Date e Double como o mesmo tipo. Deverá ser mantido uma nova célula. */
            celulaDouble.SetCellValue(Math.Round((double)valor, 2));
            celulaDouble.CellStyle = estiloCelulaConteudo;
          }

          else if (tipo.Equals(typeof(int)) || tipo.Equals(typeof(int?)))
          {
            celulaConteudo.SetCellValue((int)valor);
          }

          else if (tipo.Equals(typeof(bool)))
          {
            celulaConteudo.SetCellValue((bool)valor);
          }

          else
          {
            string valorStringUpperCase = valor.ToString().ToUpper();
            celulaConteudo.SetCellValue(valorStringUpperCase);
          }
        }
        column++;
      }

      ICell celulaUltimaColuna = linha.CreateCell(column);
      celulaUltimaColuna.CellStyle = estiloUltimaColuna;

      return sheet;
    }

    private static ISheet IncluirLogoVLI(HSSFWorkbook workbook, ISheet sheet)
    {
      var merge = new NPOI.SS.Util.CellRangeAddress(1, 2, 1, 2);
      sheet.AddMergedRegion(merge);

      var diretorioAtual = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

      var caminho = $"{diretorioAtual}/Recursos/logoVLI.png";

      byte[] data = ArquivosUtil.RetornarArquivo(caminho);

      int pictureIndex = workbook.AddPicture(data, PictureType.JPEG);
      ICreationHelper helper = workbook.GetCreationHelper();
      IDrawing drawing = sheet.CreateDrawingPatriarch();

      IClientAnchor anchor = helper.CreateClientAnchor();
      anchor.Col1 = 1;
      anchor.Row1 = 1;
      IPicture picture = drawing.CreatePicture(anchor, pictureIndex);

      picture.Resize(1.8, 1.8); /*Não mudar o tamanho da imagem física. Aparecerá sobrepondo as outras células ou fixa apenas na célula alocada(mesmo sendo mesclada)*/

      return sheet;
    }

    private static ISheet IncluirTextoVLI(HSSFWorkbook workbook, ISheet sheet, IRow linha)
    {
      var merge = new NPOI.SS.Util.CellRangeAddress(1, 1, 3, 5);
      sheet.AddMergedRegion(merge);

      var textoVLI = "VLI - Valor da Logística Integrada";

      IncluirNaPlanilha(workbook, textoVLI, linha, 15, "Calibri");

      return sheet;
    }

    private static ISheet IncluirTituloPlanilha(HSSFWorkbook workbook, ISheet sheet, string tituloPlanilha, IRow linha)
    {
      var merge = new NPOI.SS.Util.CellRangeAddress(2, 2, 3, 5);
      sheet.AddMergedRegion(merge);

      IncluirNaPlanilha(workbook, tituloPlanilha, linha, 16, "Calibri");

      return sheet;
    }

    private static ISheet IncluirSubTituloPlanilha(HSSFWorkbook workbook, ISheet sheet, string tituloPlanilha, IRow linha)
    {
      var merge = new NPOI.SS.Util.CellRangeAddress(3, 3, 3, 5);
      sheet.AddMergedRegion(merge);

      IncluirNaPlanilha(workbook, tituloPlanilha, linha, 16, "Calibri");

      return sheet;
    }

    private static void IncluirNaPlanilha(HSSFWorkbook workbook, string tituloPlanilha, IRow linha, double fontHeight, string fontName)
    {
      var fonte = workbook.CreateFont();
      fonte.IsBold = true;
      fonte.FontHeightInPoints = fontHeight;
      fonte.FontName = fontName;

      var estiloCabecalho = RetornarEstiloCelulaFundoBranco(workbook);
      estiloCabecalho.VerticalAlignment = VerticalAlignment.Center;
      estiloCabecalho.Alignment = HorizontalAlignment.Center;
      estiloCabecalho.SetFont(fonte);

      ICell tituloPlanilhaCelula = linha.CreateCell(3);
      tituloPlanilhaCelula.CellStyle = estiloCabecalho;
      tituloPlanilhaCelula.SetCellValue(tituloPlanilha);
    }

    private static ICellStyle RetornarEstiloCelulaConteudo(HSSFWorkbook workbook)
    {
      var fonteConteudo = workbook.CreateFont();
      fonteConteudo.FontHeightInPoints = 11;
      fonteConteudo.FontName = "Calibri";

      ICellStyle estiloConteudo = workbook.CreateCellStyle();
      estiloConteudo.Alignment = HorizontalAlignment.Right;
      estiloConteudo.SetFont(fonteConteudo);

      estiloConteudo.BorderTop = BorderStyle.Dotted;
      estiloConteudo.TopBorderColor = HSSFColor.Grey80Percent.Index;

      estiloConteudo.BorderRight = BorderStyle.Dotted;
      estiloConteudo.RightBorderColor = HSSFColor.Grey80Percent.Index;

      return estiloConteudo;
    }

    private static ICellStyle RetornarEstiloCelulaCabecalhoCinza(HSSFWorkbook workbook)
    {
      var fonteCabecalho = workbook.CreateFont();
      fonteCabecalho.Color = IndexedColors.White.Index;
      fonteCabecalho.IsBold = true;
      fonteCabecalho.FontHeightInPoints = 11;
      fonteCabecalho.FontName = "Calibri";
      fonteCabecalho.IsBold = true;

      var estiloCabecalho = workbook.CreateCellStyle();
      estiloCabecalho.VerticalAlignment = VerticalAlignment.Center;
      estiloCabecalho.Alignment = HorizontalAlignment.Center;
      estiloCabecalho.SetFont(fonteCabecalho);
      estiloCabecalho.FillForegroundColor = HSSFColor.Grey50Percent.Index;
      estiloCabecalho.FillPattern = FillPattern.SolidForeground;

      estiloCabecalho.SetFont(fonteCabecalho);

      return estiloCabecalho;
    }

    private static ICellStyle RetornarEstiloCelulaFundoBranco(HSSFWorkbook workbook)
    {
      ICellStyle estilo = workbook.CreateCellStyle();
      estilo.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.White.Index;
      estilo.FillPattern = FillPattern.SolidForeground;
      return estilo;
    }

    private static ICellStyle RetornarEstiloCelulaBordaLinhaSuperiorBordaPontilhadaInferior(HSSFWorkbook workbook)
    {
      var fonteConteudo = workbook.CreateFont();
      fonteConteudo.FontHeightInPoints = 11;
      fonteConteudo.FontName = "Calibri";

      ICellStyle estiloCelulaBordaLinhaSuperiorBordaPontilhadaInferior = workbook.CreateCellStyle();
      estiloCelulaBordaLinhaSuperiorBordaPontilhadaInferior.Alignment = HorizontalAlignment.Center;
      estiloCelulaBordaLinhaSuperiorBordaPontilhadaInferior.SetFont(fonteConteudo);

      estiloCelulaBordaLinhaSuperiorBordaPontilhadaInferior.BorderTop = BorderStyle.Thin;
      estiloCelulaBordaLinhaSuperiorBordaPontilhadaInferior.TopBorderColor = HSSFColor.Black.Index;

      estiloCelulaBordaLinhaSuperiorBordaPontilhadaInferior.BorderBottom = BorderStyle.Dotted;
      estiloCelulaBordaLinhaSuperiorBordaPontilhadaInferior.BottomBorderColor = HSSFColor.Grey80Percent.Index;

      estiloCelulaBordaLinhaSuperiorBordaPontilhadaInferior.BorderRight = BorderStyle.Dotted;
      estiloCelulaBordaLinhaSuperiorBordaPontilhadaInferior.RightBorderColor = HSSFColor.Grey80Percent.Index;

      return estiloCelulaBordaLinhaSuperiorBordaPontilhadaInferior;
    }

    private static ICellStyle RetornarEstiloCelulaBordaLinhaInferior(HSSFWorkbook workbook)
    {
      var fonteConteudo = workbook.CreateFont();
      fonteConteudo.FontHeightInPoints = 11;
      fonteConteudo.FontName = "Calibri";

      ICellStyle estiloCelulaBordaLinhaInferior = workbook.CreateCellStyle();
      estiloCelulaBordaLinhaInferior.Alignment = HorizontalAlignment.Center;
      estiloCelulaBordaLinhaInferior.SetFont(fonteConteudo);

      estiloCelulaBordaLinhaInferior.BorderBottom = BorderStyle.Thin;
      estiloCelulaBordaLinhaInferior.BottomBorderColor = HSSFColor.Black.Index;

      estiloCelulaBordaLinhaInferior.BorderRight = BorderStyle.Dotted;
      estiloCelulaBordaLinhaInferior.RightBorderColor = HSSFColor.Grey80Percent.Index;

      return estiloCelulaBordaLinhaInferior;
    }

    private static ISheet AjustarLarguraColunas(ISheet sheet, string[] nomesAtributosExibicao)
    {
      int column = 1;
      foreach (string nomeAtributo in nomesAtributosExibicao)
      {
        sheet.AutoSizeColumn(column);
        column++;
      }

      sheet.SetColumnWidth(0, 500);

      return sheet;
    }

    private static List<PropertyInfo> RetornarPropriedadesOrdemExibicao<TObject>(string[] nomesAtributosExibicao)
    {
      Type tipoObjeto = typeof(TObject);
      List<PropertyInfo> propriedades = new List<PropertyInfo>(tipoObjeto.GetProperties());
      List<PropertyInfo> propriedadesEmOrdemDeExibicao = new List<PropertyInfo>();

      foreach (string nomeAtributo in nomesAtributosExibicao)
      {
        propriedadesEmOrdemDeExibicao.Add(propriedades[propriedades.FindIndex(prop => prop.Name == nomeAtributo)]);
      }

      return propriedadesEmOrdemDeExibicao;
    }

    private static ISheet InsereSomatorioSubSomatorio(HSSFWorkbook workbook, ISheet sheet, IRow linhaTotalSomatorio, IRow linhaSubSomatorio,
        string linhaInicioSomatorio, string linhaFimSomatorio, int tipoTabela)
    {
      switch (tipoTabela)
      {
        case 1:
          string[] letrasSomatorioFerrovias = new string[] { "K", "L", "M", "N", "O", "P" };
          int[] indiceLetrasSomatorioFerrovias = new int[] { 10, 11, 12, 13, 14, 15 };

          sheet = IncluirCelulasTextoSomatorioSubSomatorio(workbook, sheet, linhaTotalSomatorio, (indiceLetrasSomatorioFerrovias[0] - 1), linhaSubSomatorio, (indiceLetrasSomatorioFerrovias[0] - 1));

          sheet = IncluirSomatorio(workbook, sheet, linhaTotalSomatorio, linhaInicioSomatorio, linhaFimSomatorio,
          letrasSomatorioFerrovias, indiceLetrasSomatorioFerrovias);

          sheet = IncluirSubSomatorio(workbook, sheet, linhaSubSomatorio, linhaInicioSomatorio, linhaFimSomatorio,
              letrasSomatorioFerrovias, indiceLetrasSomatorioFerrovias);
          return sheet;

        case 2:
          string[] letrasSomatorioPortosTerminais = new string[] { "H", "I", "J" };
          int[] indiceLetrasSomatorioPortosTerminais = new int[] { 7, 8, 9 };

          sheet = IncluirCelulasTextoSomatorioSubSomatorio(workbook, sheet, linhaTotalSomatorio, (indiceLetrasSomatorioPortosTerminais[0] - 1), linhaSubSomatorio, (indiceLetrasSomatorioPortosTerminais[0] - 1));

          sheet = IncluirSomatorio(workbook, sheet, linhaTotalSomatorio, linhaInicioSomatorio, linhaFimSomatorio,
              letrasSomatorioPortosTerminais, indiceLetrasSomatorioPortosTerminais);

          sheet = IncluirSubSomatorio(workbook, sheet, linhaSubSomatorio, linhaInicioSomatorio, linhaFimSomatorio,
              letrasSomatorioPortosTerminais, indiceLetrasSomatorioPortosTerminais);
          return sheet;
        default:
          return sheet;
      }
    }

    private static ISheet IncluirSomatorio(HSSFWorkbook workbook, ISheet sheet, IRow linhaTotalSomatorio, string linhaInicioSomatorio,
        string linhaFimSomatorio, string[] letras, int[] indiceLetras)
    {
      var fonte = workbook.CreateFont();
      fonte.IsBold = true;
      fonte.FontHeightInPoints = 10;
      fonte.FontName = "Calibri";

      var estiloCabecalho = RetornarEstiloCelulaBordaLinhaSuperiorBordaPontilhadaInferior(workbook);
      estiloCabecalho.VerticalAlignment = VerticalAlignment.Center;
      estiloCabecalho.Alignment = HorizontalAlignment.Center;
      estiloCabecalho.SetFont(fonte);

      /*Tem que duplicar se não ele concatena*/

      var estiloUltimaCelulaCabecalho = RetornarEstiloCelulaBordaLinhaSuperiorBordaPontilhadaInferior(workbook);
      estiloUltimaCelulaCabecalho.VerticalAlignment = VerticalAlignment.Center;
      estiloUltimaCelulaCabecalho.Alignment = HorizontalAlignment.Center;
      estiloUltimaCelulaCabecalho.BorderRight = BorderStyle.Thin;
      estiloUltimaCelulaCabecalho.RightBorderColor = HSSFColor.Black.Index;
      estiloUltimaCelulaCabecalho.SetFont(fonte);

      for (int letraIndiceFormulas = 0; letraIndiceFormulas < indiceLetras.Length; letraIndiceFormulas++)
      {
        ICell cell = linhaTotalSomatorio.CreateCell(indiceLetras[letraIndiceFormulas]);
        cell.SetCellType(CellType.Formula);
        cell.SetCellFormula("SUM(" + letras[letraIndiceFormulas] + linhaInicioSomatorio + ":" +
            letras[letraIndiceFormulas] + linhaFimSomatorio + ")");

        if (letraIndiceFormulas == (indiceLetras.Length - 1))
        {
          cell.CellStyle = estiloUltimaCelulaCabecalho;
        }
        else
        {
          cell.CellStyle = estiloCabecalho;
        }
      }

      return sheet;
    }

    private static ISheet IncluirSubSomatorio(HSSFWorkbook workbook, ISheet sheet, IRow linhaTotalSomatorio,
        string linhaInicioSomatorio, string linhaFimSomatorio, string[] letras, int[] indiceLetras)
    {
      var fonte = workbook.CreateFont();
      fonte.IsBold = true;
      fonte.FontHeightInPoints = 10;
      fonte.FontName = "Calibri";

      var estiloCabecalho = RetornarEstiloCelulaBordaLinhaInferior(workbook);
      estiloCabecalho.VerticalAlignment = VerticalAlignment.Center;
      estiloCabecalho.Alignment = HorizontalAlignment.Center;
      estiloCabecalho.SetFont(fonte);

      /*Tem que duplicar se não ele concatena*/

      var estiloUltimaCelulaCabecalho = RetornarEstiloCelulaBordaLinhaInferior(workbook);
      estiloUltimaCelulaCabecalho.VerticalAlignment = VerticalAlignment.Center;
      estiloUltimaCelulaCabecalho.Alignment = HorizontalAlignment.Center;
      estiloUltimaCelulaCabecalho.BorderRight = BorderStyle.Thin;
      estiloUltimaCelulaCabecalho.RightBorderColor = HSSFColor.Black.Index;
      estiloUltimaCelulaCabecalho.SetFont(fonte);

      for (int letraIndiceFormulas = 0; letraIndiceFormulas < indiceLetras.Length; letraIndiceFormulas++)
      {
        ICell cell = linhaTotalSomatorio.CreateCell(indiceLetras[letraIndiceFormulas]);
        cell.SetCellType(CellType.Formula);
        cell.SetCellFormula("SUBTOTAL(9," + letras[letraIndiceFormulas] + linhaInicioSomatorio + ":" +
            letras[letraIndiceFormulas] + linhaFimSomatorio + ")");

        if (letraIndiceFormulas == (indiceLetras.Length - 1))
        {
          cell.CellStyle = estiloUltimaCelulaCabecalho;
        }
        else
        {
          cell.CellStyle = estiloCabecalho;
        }
      }

      return sheet;
    }

    private static ISheet IncluirCelulasTextoSomatorioSubSomatorio(HSSFWorkbook workbook, ISheet sheet, IRow linhaTextoSomatorio, int colunaTextoSomatorio,
        IRow linhaTextoSubSomatorio, int colunaTextoSubSomatorio)
    {
      var textoTotal = "Total:";
      var textoSubTotal = "Subtotal:";

      var fonte = workbook.CreateFont();
      fonte.IsBold = true;
      fonte.FontHeightInPoints = 10;
      fonte.FontName = "Calibri";

      var estiloTotal = RetornarEstiloCelulaBordaLinhaSuperiorBordaPontilhadaInferior(workbook);
      estiloTotal.VerticalAlignment = VerticalAlignment.Center;
      estiloTotal.Alignment = HorizontalAlignment.Right;
      estiloTotal.BorderLeft = BorderStyle.Thin;
      estiloTotal.LeftBorderColor = HSSFColor.Black.Index;
      estiloTotal.SetFont(fonte);

      ICell textoTotalCelula = linhaTextoSomatorio.CreateCell(colunaTextoSomatorio);
      textoTotalCelula.SetCellValue(textoTotal);
      textoTotalCelula.CellStyle = estiloTotal;

      var estiloSubTotal = RetornarEstiloCelulaBordaLinhaInferior(workbook);
      estiloSubTotal.VerticalAlignment = VerticalAlignment.Center;
      estiloSubTotal.Alignment = HorizontalAlignment.Right;
      estiloSubTotal.BorderLeft = BorderStyle.Thin;
      estiloSubTotal.LeftBorderColor = HSSFColor.Black.Index;

      estiloSubTotal.SetFont(fonte);

      ICell textoSubTotalCelula = linhaTextoSubSomatorio.CreateCell(colunaTextoSubSomatorio);
      textoSubTotalCelula.SetCellValue(textoSubTotal);
      textoSubTotalCelula.CellStyle = estiloSubTotal;

      return sheet;

    }
  }
}
