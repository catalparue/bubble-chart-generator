using System;
using System.Threading.Tasks;
using System.Windows.Input;
using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.CommandWpf;
using Microsoft.Win32;

namespace ChartGeneratorUI
{
    public class MainWindowViewModel : ViewModelBase
    {
        private const double DefaultChartWidth = 800;
        private const double DefaultChartHeight = 500;

        public double MinChartWidth { get; } = 400;
        public double MinChartHeight { get; } = 400;

        private bool _isAppEnabled = true;
        private string _stringMessage;
        private double _chartWidth = DefaultChartWidth;
        private double _chartHeight = DefaultChartHeight;

        public ICommand GenerateBubbleChartCommand { get; set; }

        public bool IsAppEnabled
        {
            get => _isAppEnabled;
            set => Set(ref _isAppEnabled, value);
        }

        public string StatusMessage
        {
            get => _stringMessage;
            private set => Set(ref _stringMessage, value);
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

        public MainWindowViewModel()
        {
            GenerateBubbleChartCommand = new RelayCommand(GenerateChart, IsAppEnabled);
        }

        public async void GenerateChart()
        {
            var saveFileDialog = new SaveFileDialog
            {
                OverwritePrompt = true,
                FileName = "bubblechart.png",
                Filter = "PNG files (*.png)|*.png|All files (*.*)|*.*",
                RestoreDirectory = true
            };

            if (saveFileDialog.ShowDialog() != true) return;

            IsAppEnabled = false;
            StatusMessage = "Arbetar";
            try
            {
                await Task.Run(() => ExcelBubbleChartGenerator.ExcelBubbleChartGenerator.GenerateBubbleChart(ChartWidth, ChartHeight, GetChartScaleFactor(),
                    saveFileDialog.FileName));
                StatusMessage = "Bilden har sparats till:\n" + saveFileDialog.FileName;
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
            var xFactor = ChartWidth / DefaultChartWidth;
            var yFactor = ChartHeight / DefaultChartWidth;
            return xFactor * yFactor;
        }
    }
}
