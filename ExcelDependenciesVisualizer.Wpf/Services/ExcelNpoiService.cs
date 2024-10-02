using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;

namespace ExcelDependenciesVisualizer.Wpf.Services;

public class ExcelNpoiService : IExcelService
{
    public IEnumerable<ExcelCell> GetExcelCells(string filePath)
    {
        using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
        {
            stream.Position = 0;
            IWorkbook workBook;
            if (filePath.EndsWith(".xls"))
            {
                workBook = new HSSFWorkbook(stream);
            }
            else
            {
                workBook = new XSSFWorkbook(stream);
            }
            List<string> sheetNameList = new();
            for (int i = 0; i < workBook.NumberOfSheets; i++)
            {
                sheetNameList.Add(workBook.GetSheetName(i));
            }
            foreach (var sheetName in sheetNameList)
            {
                var sheet = workBook.GetSheet(sheetName);
                for (int i = 0; i <= sheet.LastRowNum; i++)
                {
                    var row = sheet.GetRow(i);
                    if (row == null)
                    {
                        continue;
                    }
                    for (int j = 0; j < row.LastCellNum; j++)
                    {
                        var cell = row.GetCell(j);
                        if (cell == null)
                        {
                            continue;
                        }
                        if (cell.CellType == CellType.Formula)
                        {
                            var excelCell = new ExcelCell
                            {
                                SheetName = sheetName,
                                CellAddress = cell.Address.FormatAsString(),
                                Formula = cell.CellFormula.Replace("$", string.Empty),
                            };
                            if(!excelCell.ReferencingCells.Any() || excelCell.ReferencingCells.All(x=>x.Contains("REF")))
                            {
                                continue;
                            }
                            yield return excelCell; 
                        }
                    }
                }
            }
        }
    }
}
