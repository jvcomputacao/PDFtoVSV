using System.Text;
using GrapeCity.Documents.Pdf;
using GrapeCity.Documents.Pdf.Recognition;
using System.Linq;
using System.Drawing;

const float DPI = 72;
var tableData = new List<List<string>>();
using (var fileStream = File.OpenRead(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\PDFexemplo.pdf")))
{
    // Cria uma instância do objeto GcPdfDocument
    var pdfDocument = new GcPdfDocument();
    // Carrega o documento PDF
    pdfDocument.Load(fileStream);

    // Define uma área aproximada onde a tabela está localizada:
    var tableArea = new RectangleF(0, 2.5f * DPI, 8.5f * DPI, 3.75f * DPI);

    // TableExtractOptions permite ajustar o reconhecimento da tabela com base em
    // características específicas do formato da tabela:
    var extractOptions = new TableExtractOptions();
    var calculateMinRowDistance = extractOptions.GetMinimumDistanceBetweenRows;

    //  aumentama ligeiramente a distância mínima entre as linhas para garantir que células com texto quebrado em várias linhas não sejam interpretadas como duas células distintas:
    extractOptions.GetMinimumDistanceBetweenRows = (list) =>
    {
        var result = calculateMinRowDistance(list);
        return result * 1.2f;
    };

    for (int pageIndex = 0; pageIndex < pdfDocument.Pages.Count; ++pageIndex)
    {
        // Extrai a tabela localizada na área especificada:
        var table = pdfDocument.Pages[pageIndex].GetTable(tableArea, extractOptions);
        if (table != null)
        {
            for (int rowIndex = 0; rowIndex < table.Rows.Count; ++rowIndex)
            {
                // Adiciona a próxima linha de dados, ignorando os cabeçalhos:
                if (rowIndex > 0)
                    tableData.Add(new List<string>());

                for (int colIndex = 0; colIndex < table.Cols.Count; ++colIndex)
                {
                    var cell = table.GetCell(rowIndex, colIndex);
                    if (cell == null && rowIndex > 0)
                        tableData.Last().Add("");
                    else
                    {
                        if (cell != null && rowIndex > 0)
                            tableData.Last().Add($"\"{cell.Text}\"");
                    }
                }
            }
        }
    }

    // Necessário para codificar caracteres não ASCII nos dados
    Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    File.Delete("pdfdata.CSV");
    File.AppendAllLines(
            "pdfdata.CSV",
            tableData.Where(row => row.Any(col => !string.IsNullOrEmpty(col))).Select(row => string.Join(',', row)),
            Encoding.GetEncoding(1252));
}