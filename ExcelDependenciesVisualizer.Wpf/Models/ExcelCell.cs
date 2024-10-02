using ExcelDependenciesVisualizer.Utilities;
using NPOI.SS.Formula;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelDependenciesVisualizer.Models;

public record ExcelCell
{
    public string SheetName { get; set; } = string.Empty;
    public string CellAddress { get; set; } = string.Empty;
    private string _formula = string.Empty;
    public string Formula
    {
        get => _formula;
        set
        {
            _formula = value;
            ReferencingCells = ExtractCellsFromFormula();
        }
    }
    /// <summary>
    /// Full Address in format SheetName!CellAddress
    /// </summary>
    public List<string> ReferencingCells { get; internal set; } = new();

    private List<string> ExtractCellsFromFormula()
    {
        HashSet<string> cellReferences = new HashSet<string>();
        var tokens = Formula.SplitMultiDelims("+-*/()=, \"");

        // Regex pattern to match valid cell references
        string pattern = @"^[A-Z]+[0-9]+$"; // e.g., A1, B12, AA30
        Regex regex = new Regex(pattern, RegexOptions.IgnoreCase);

        foreach (var token in tokens)
        {
            // Check if the token includes a sheet reference
            if (token.Contains("!"))
            {
                // It's possible it's a reference to another sheet
                cellReferences.Add(token); // Add directly to collection
            }
            else if (regex.IsMatch(token)) // Check if it's a valid cell reference without a sheet name
            {
                //cellReferences.Add(token); // Add valid reference to the collection
                cellReferences.Add($"{SheetName}!{token}"); // Add valid reference with sheet name to the collection
            }
        }

        return new List<string>(cellReferences);
    }
}

public record ExcelCellDto
{
    public ExcelCell Cell { get; set; }
    //public List<ExcelCell> ReferencedByCells { get; set; } = new();
    public List<ExcelCellDto> ReferencedByCells { get; set; } = new();
    public override string ToString()
    {
        return ToString(0); // Start with an initial indent level of 0
    }

    private string ToString(int indentCount)
    {
        StringBuilder sb = new StringBuilder();

        // Add indentation based on the level
        //string indent = new string(' ', indentCount * 4); // 4 spaces per level
        string indent = new string('\t', indentCount);
        sb.AppendLine($"{indent}{Cell.SheetName}!{Cell.CellAddress}");

        // Recursively append ReferencedByCells with increased indent level
        foreach (var referencedCell in ReferencedByCells)
        {
            sb.Append(referencedCell.ToString(indentCount + 1));
        }

        return sb.ToString();
    }
}