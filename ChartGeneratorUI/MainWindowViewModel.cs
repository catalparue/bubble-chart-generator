using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Input;
using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.CommandWpf;
using Microsoft.Win32;

namespace ChartGeneratorUI
{
    public class MainWindowViewModel : ViewModelBase
    {
        private const double DefaultChartWidth = 1200;
        private const double DefaultChartHeight = 700;
        private const int WorksheetToFetchFrom = 2;
        private const int MaxStage = 6;

        public double MinChartWidth { get; } = 1;
        public double MinChartHeight { get; } = 1;

        private bool _isAppEnabled = true;
        private string _sourceFilePath;
        private string _destinationFilePath;
        private string _statusMessage;
        private double _chartWidth = DefaultChartWidth;
        private double _chartHeight = DefaultChartHeight;
        private double _stageFraction;

        private readonly ExcelBubbleChartGenerator.ExcelBubbleChartGenerator _excelBubbleChartGenerator;

        public ICommand GenerateBubbleChartCommand { get; set; }
        public ICommand SelectSourceFileCommand { get; set; }
        public ICommand SelectDestinationFileCommand { get; set; }

        public bool IsAppEnabled
        {
            get => _isAppEnabled;
            set => Set(ref _isAppEnabled, value);
        }

        public string SourceFilePath
        {
            get => _sourceFilePath;
            set => Set(ref _sourceFilePath, value);
        }

        public string DestinationFilePath
        {
            get => _destinationFilePath;
            set => Set(ref _destinationFilePath, value);
        }

        public string StatusMessage
        {
            get => _statusMessage;
            private set => Set(ref _statusMessage, value);
        }

        public double ChartWidth
        {
            get => _chartWidth;
            set => Set(ref _chartWidth, Math.Max(value, MinChartWidth));
        }

        public double ChartHeight
        {
            get => _chartHeight;
            set => Set(ref _chartHeight, Math.Max(value, MinChartHeight));
        }

        public double StageFraction
        {
            get => _stageFraction;
            set => Set(ref _stageFraction, value);
        }

        public MainWindowViewModel()
        {
            GenerateBubbleChartCommand = new RelayCommand(GenerateChart, IsAppEnabled);
            SelectSourceFileCommand = new RelayCommand(SelectSourceFile, IsAppEnabled);
            SelectDestinationFileCommand = new RelayCommand(SelectDestinationFile, IsAppEnabled);
            _excelBubbleChartGenerator = new ExcelBubbleChartGenerator.ExcelBubbleChartGenerator();
            _excelBubbleChartGenerator.StatusUpdated += (sender, args) =>
            {
                StatusMessage = args.StatusMessage;
                StageFraction = (double) args.Stage / MaxStage;
            };
        }

        public void SelectSourceFile()
        {
            var openFileDialog = new OpenFileDialog
            {
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                Filter = "Excel files (*.xls, *.xlsx)|*.xls;*.xlsx",
                RestoreDirectory = true
            };

            if (openFileDialog.ShowDialog() == true) SourceFilePath = openFileDialog.FileName;
        }

        public void SelectDestinationFile()
        {
            var saveFileDialog = new SaveFileDialog
            {
                OverwritePrompt = true,
                FileName = "bubblechart.png",
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures),
                Filter = "PNG files (*.png)|*.png",
                RestoreDirectory = true
            };

            if (saveFileDialog.ShowDialog() == true) DestinationFilePath = saveFileDialog.FileName;
        }

        public async void GenerateChart()
        {
            if (!File.Exists(SourceFilePath))
            {
                StatusMessage = "Kunde inte hitta den angivna källfilen!";
                return;
            }

            if (!Directory.Exists(Path.GetDirectoryName(DestinationFilePath)))
            {
                StatusMessage = "Kan inte spara diagrammet till vald plats!";
                return;
            }

            if (Path.GetExtension(DestinationFilePath)?.ToLower() != ".png")
            {
                StatusMessage = "Diagrammet måste sparas som PNG-fil!";
                return;
            }

            IsAppEnabled = false;
            StatusMessage = "Arbetar";
            try
            {
                await Task.Run(() => _excelBubbleChartGenerator.GenerateBubbleChart(
                    ChartWidth,
                    ChartHeight,
                    GetChartScaleFactor(),
                    SourceFilePath,
                    DestinationFilePath,
                    WorksheetToFetchFrom));
                StatusMessage = "Bilden har sparats till:\n" + DestinationFilePath;
            }
            catch (Exception exception)
            {
                StatusMessage = "Något gick fel:\n" + exception.Message;
            }
            finally
            {
                IsAppEnabled = true;
            }
        }

        private double GetChartScaleFactor()
        {
            var xFactor = DefaultChartWidth / ChartWidth;
            var yFactor = DefaultChartHeight / ChartHeight;
            return xFactor * yFactor;
        }
    }
}
