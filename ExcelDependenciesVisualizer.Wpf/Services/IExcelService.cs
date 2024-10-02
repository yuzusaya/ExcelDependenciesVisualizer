namespace ExcelDependenciesVisualizer.Wpf.Services;

public interface IExcelService
{
    IEnumerable<ExcelCell> GetExcelCells(string filePath);
}
