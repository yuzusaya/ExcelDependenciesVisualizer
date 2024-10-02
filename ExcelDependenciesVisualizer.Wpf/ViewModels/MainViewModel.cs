using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ExcelDependenciesVisualizer.Wpf.Services;
using Microsoft.Win32;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Input;
using ToastNotifications;
using ToastNotifications.Lifetime;
using ToastNotifications.Messages;
using ToastNotifications.Position;

namespace ExcelDependenciesVisualizer.Wpf.ViewModels;

public partial class MainViewModel : BaseModel
{
    private IExcelService _excelService;
    private Notifier _notifier;
    [ObservableProperty]
    private string _selectedExcelFilePath = string.Empty;
    private List<ExcelCell> _excelCells = new();
    [ObservableProperty]
    private List<string> _sheetNameList = new();
    [ObservableProperty]
    private string _selectedSheetName = string.Empty;
    [ObservableProperty]
    private List<ExcelCellDto> _referencingCells = new();
    [ObservableProperty]
    private Visibility _formulaVisibility = Visibility.Hidden;
    [ObservableProperty]
    private string _selectedCellAddress = string.Empty;
    public ExcelCellDto SelectedCell { get; set; }

    [RelayCommand(AllowConcurrentExecutions = false)]
    private async Task LoadExcel()
    {
        var openFileDialog = new OpenFileDialog
        {
            Title = "Select Excel File",
            Filter = "Excel Files (*.xls, *.xlsx, *.xlsm)|*.xls;*.xlsx;*.xlsm",
        };
        var result = openFileDialog.ShowDialog();
        if (result == true)
        {
            try
            {
                SelectedExcelFilePath = openFileDialog.FileName;
                var excelCells = await Task.Run(() => _excelService.GetExcelCells(openFileDialog.FileName));
                _excelCells = excelCells.ToList();
                SheetNameList = _excelCells.Select(x => x.SheetName).Distinct().ToList();
            }
            catch (System.IO.IOException)
            {
                _notifier.ShowError($"{openFileDialog.FileName}が使用されています。閉じてからもう一度開いてください。");
            }
            catch (Exception ex)
            {
                _notifier.ShowError(ex.Message);
            }
        }
    }
    [RelayCommand]
    private void CellSelected(ExcelCellDto excelCellDto)
    {
        SelectedCell = excelCellDto;
    }
    [RelayCommand]
    private void ChangeFormulaVisibility()
    {
        //show or hide formula
        FormulaVisibility = FormulaVisibility == Visibility.Visible ? Visibility.Hidden : Visibility.Visible;
    }
    [RelayCommand]
    private void CopySelectedCell()
    {
        if (SelectedCell != null)
        {
            Clipboard.SetText(SelectedCell.ToString());
            _notifier.ShowInformation("コピーしました。");
        }
    }
    [RelayCommand]
    private void SelectedCellDoubleClicked(ExcelCellDto excelCellDto)
    {
        SelectedCellAddress = excelCellDto.Cell.CellAddress;
        SelectedSheetName = excelCellDto.Cell.SheetName;
    }
    [RelayCommand]
    private void SearchReferencedByCells()
    {
        try
        {
            SearchExecuted?.Invoke(this, EventArgs.Empty);
            if (string.IsNullOrWhiteSpace(SelectedSheetName) || string.IsNullOrWhiteSpace(SelectedCellAddress))
            {
                _notifier.ShowWarning("シート名とセルアドレスを選択してください。");
                return;
            }
            SelectedCellAddress = SelectedCellAddress.ToUpper();
            //check format of cell address (e.g., A1, B12, AA30)
            if (!Regex.IsMatch(SelectedCellAddress, @"^[A-Z]+[0-9]+$"))
            {
                _notifier.ShowWarning("セルアドレスの形式が正しくありません。(e.g., A1, AB30)");
                return;
            }

            ReferencingCells = GetReferencingCells(SelectedSheetName, SelectedCellAddress);
            if (ReferencingCells.Count == 0)
            {
                _notifier.ShowInformation("参照しているセルがありませんよ。");
                return;
            }
            //var parentCell = new ExcelCellDto
            //{
            //    Cell = new ExcelCell
            //    {
            //        SheetName = SelectedSheetName,
            //        CellAddress = SelectedCellAddress,
            //        Formula = _excelCells.FirstOrDefault(x => x.SheetName == SelectedSheetName && x.CellAddress == SelectedCellAddress)?.Formula ?? string.Empty,
            //    },
            //    ReferencedByCells = ReferencingCells
            //};
        }
        catch (Exception ex)
        {
            _notifier.ShowError(ex.Message);
        }
    }

    private List<ExcelCellDto> GetReferencingCells(string sheetName, string cellAddress)
    {
        var referencingCells = new List<ExcelCellDto>();
        var referencingCellAddresses = _excelCells
            .Where(x => x.ReferencingCells.Contains($"{sheetName}!{cellAddress}"))
            .ToList();
        //recursive search
        foreach (var referencingCellAddress in referencingCellAddresses)
        {
            var referencingCell = new ExcelCellDto()
            {
                Cell = referencingCellAddress,
                ReferencedByCells = GetReferencingCells(referencingCellAddress.SheetName, referencingCellAddress.CellAddress)
            };
            referencingCells.Add(referencingCell);
        }
        return referencingCells;
    }
    
    public event EventHandler SearchExecuted;
    public MainViewModel()
    {
        _excelService = new ExcelNpoiService();
        _notifier = new Notifier(cfg =>
        {
            cfg.PositionProvider = new WindowPositionProvider(
                parentWindow: Application.Current.MainWindow,
                corner: Corner.BottomLeft,
                offsetX: 10,
                offsetY: 10);

            cfg.LifetimeSupervisor = new TimeAndCountBasedLifetimeSupervisor(
                notificationLifetime: TimeSpan.FromSeconds(3),
                maximumNotificationCount: MaximumNotificationCount.FromCount(5));

            cfg.Dispatcher = Application.Current.Dispatcher;
        });
    }
}