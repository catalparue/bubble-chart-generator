using System;
using System.Collections.Generic;
using System.Diagnostics;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelBubbleChartGenerator
{
    public static class ExcelBubbleChartGenerator
    {
        private const double DataLabelDistanceMargin = 5;
        private const double RotationIncrement = Math.PI / 8;

        private static double _bubbleScaleFactor = 1;

        private static Excel.Application _excelApp;
        private static Excel.Workbook _excelWorkbook;
        private static Excel.Worksheet _excelWorksheet;
        private static Excel.Chart _bubbleChart;
        private static Excel.SeriesCollection SeriesCollection => _bubbleChart?.SeriesCollection();

        private static readonly Dictionary<int, string> ProjectTypeNames = new Dictionary<int, string>
        {
            {1, "Lösning/projekt"},
            {2, "Konsulttjänst"},
            {3, "Indirekt"}
        };

        private static readonly Dictionary<int, int> ProjectTypeColors = new Dictionary<int, int>
        {
            {1, (int) Excel.XlRgbColor.rgbTan},
            {2, (int) Excel.XlRgbColor.rgbLimeGreen},
            {3, (int) Excel.XlRgbColor.rgbRed}
        };

        public static void SetExcelData()
        {

        }

        public static void GenerateBubbleChart(double chartWidth, double chartHeight, double bubbleScaleFactor, string fileName)
        {
            Debug.WriteLine("Starting up...");
            _excelApp = new Excel.Application();
            _excelWorkbook = _excelApp.Workbooks.Add(1);
            _excelWorksheet = (Excel.Worksheet)_excelWorkbook.Sheets[1];

            try
            {
                // Add data
                _excelWorksheet.Cells[1, 1] = "Attentec"; // Projektnamn
                _excelWorksheet.Cells[1, 2] = "600"; // Medeltimpris
                _excelWorksheet.Cells[1, 3] = "20%"; // TG2 Medel
                _excelWorksheet.Cells[1, 4] = "150"; // Omsättning kategori
                _excelWorksheet.Cells[1, 5] = "1"; // Lösning/konsulttjänst/indirekt kategori

                _excelWorksheet.Cells[2, 1] = "Umbrella Corporation";
                _excelWorksheet.Cells[2, 2] = "800";
                _excelWorksheet.Cells[2, 3] = "11%";
                _excelWorksheet.Cells[2, 4] = "300";
                _excelWorksheet.Cells[2, 5] = "2";

                _excelWorksheet.Cells[3, 1] = "LexCorp";
                _excelWorksheet.Cells[3, 2] = "1000";
                _excelWorksheet.Cells[3, 3] = "6%";
                _excelWorksheet.Cells[3, 4] = "900";
                _excelWorksheet.Cells[3, 5] = "3";

                _excelWorksheet.Cells[4, 1] = "Aperture Science";
                _excelWorksheet.Cells[4, 2] = "1300";
                _excelWorksheet.Cells[4, 3] = "12%";
                _excelWorksheet.Cells[4, 4] = "805";
                _excelWorksheet.Cells[4, 5] = "2";

                _excelWorksheet.Cells[5, 1] = "Cyberdyne Systems";
                _excelWorksheet.Cells[5, 2] = "550";
                _excelWorksheet.Cells[5, 3] = "10%";
                _excelWorksheet.Cells[5, 4] = "50";
                _excelWorksheet.Cells[5, 5] = "2";

                _excelWorksheet.Cells[6, 1] = "Weyland-Yutani";
                _excelWorksheet.Cells[6, 2] = "807";
                _excelWorksheet.Cells[6, 3] = "11%";
                _excelWorksheet.Cells[6, 4] = "120";
                _excelWorksheet.Cells[6, 5] = "1";

                _excelWorksheet.Cells[7, 1] = "Wayne Enterprises";
                _excelWorksheet.Cells[7, 2] = "780";
                _excelWorksheet.Cells[7, 3] = "11%";
                _excelWorksheet.Cells[7, 4] = "1000";
                _excelWorksheet.Cells[7, 5] = "3";

                _excelWorksheet.Cells[8, 1] = "Soylent";
                _excelWorksheet.Cells[8, 2] = "650";
                _excelWorksheet.Cells[8, 3] = "14%";
                _excelWorksheet.Cells[8, 4] = "700";
                _excelWorksheet.Cells[8, 5] = "1";

                _excelWorksheet.Cells[9, 1] = "Tyrell Corporation";
                _excelWorksheet.Cells[9, 2] = "150";
                _excelWorksheet.Cells[9, 3] = "-5%";
                _excelWorksheet.Cells[9, 4] = "50";
                _excelWorksheet.Cells[9, 5] = "1";

                //Setup chart
                Debug.WriteLine("Setting up chart properties...");
                SetupChartProperties(chartWidth, chartHeight, bubbleScaleFactor);
                Debug.WriteLine("Adding data points...");
                AddDataPoints();
                Debug.WriteLine("Adding legend...");
                AddLegendAndClearDummySeriesNames();
                Debug.WriteLine("Placing data labels...");
                SpreadOutDataLabels();

                Debug.WriteLine("Exporting image...");
                _bubbleChart.Export(fileName, "PNG");
                Process.Start(fileName);

                Debug.Write("Success! ");
            }
            finally
            {
                _excelWorkbook.Saved = true; // A lie! This is to avoid prompting the user to save before closing.
                _excelWorkbook.Close();
                _excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(_excelWorksheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(_excelWorkbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(_excelApp);
            }
        }

        private static void SetupChartProperties(double chartWidth, double chartHeight, double bubbleScaleFactor)
        {
            var excelChartObjects = (Excel.ChartObjects)_excelWorksheet.ChartObjects();
            var chartObject = excelChartObjects.Add(10, 80, chartWidth, chartHeight);
            _bubbleChart = chartObject.Chart;

            _bubbleChart.ChartType = Excel.XlChartType.xlBubble3DEffect;

            _bubbleScaleFactor = bubbleScaleFactor;

            var xAxis = (Excel.Axis)_bubbleChart.Axes(Excel.XlAxisType.xlCategory);
            xAxis.HasTitle = true;
            xAxis.AxisTitle.Caption = "Medeltimpris";
            xAxis.HasMajorGridlines = true;
            xAxis.MinimumScale = 0;
            xAxis.MajorGridlines.Format.Line.ForeColor.RGB = (int)Excel.XlRgbColor.rgbGainsboro;

            var yAxis = (Excel.Axis)_bubbleChart.Axes(Excel.XlAxisType.xlValue);
            yAxis.HasTitle = true;
            yAxis.AxisTitle.Caption = "TG2 Medel";
            yAxis.HasMajorGridlines = true;
            yAxis.TickLabels.NumberFormat = "0%";
            yAxis.MajorGridlines.Format.Line.ForeColor.RGB = (int)Excel.XlRgbColor.rgbGainsboro;
        }

        private static void AddLegendAndClearDummySeriesNames()
        {
            Excel.Legend legend = _bubbleChart.Legend;

            for (var i = 1; i <= ProjectTypeNames.Count; i++)
            {
                var series = SeriesCollection.NewSeries();
                series.Name = ProjectTypeNames[i];
                series.Format.Fill.ForeColor.RGB = ProjectTypeColors[i];
            }

            while (legend.LegendEntries().Count > 3)
            {
                legend.LegendEntries(1).Delete();
            }

            legend.Position = Excel.XlLegendPosition.xlLegendPositionBottom;
        }

        private static void AddDataPoints()
        {
            var dataMatrix = GetDataMatrixSortedByBubbleSize();

            foreach (var row in dataMatrix)
            {
                var series = SeriesCollection.NewSeries();
                series.Has3DEffect = true;
                series.HasDataLabels = true;

                var dataLabelText = row[0];
                var xValue = Convert.ToDouble(row[1]);
                var yValue = Convert.ToDouble(row[2]);
                var revenue = Convert.ToDouble(row[3]);
                var projectTypeId = Convert.ToInt32(row[4]);

                var bubbleSize = GetBubbleSizeFromRevenue(revenue);

                series.Format.Fill.ForeColor.RGB = ProjectTypeColors[projectTypeId];

                series.XValues = new[] { xValue };
                series.Values = new[] { yValue };
                series.BubbleSizes = new[] { bubbleSize };

                Excel.DataLabel dataLabel = series.DataLabels(1);
                dataLabel.Text = dataLabelText;
            }
        }

        private static void SpreadOutDataLabels()
        {
            var occupiedRectangles = new List<Rectangle>();
            var occupiedCircles = new List<Circle>();
            var leaderLineAttachingPoints = new List<double[]>();

            // First populated the list with all data bubbles...
            double i = 0;
            foreach (Excel.Series series in SeriesCollection)
            {
                foreach (Excel.Point point in series.Points())
                {
                    var circle = new Circle(point.Left + point.Width / 2, point.Top + point.Height / 2, point.Width / 2, i);
                    occupiedCircles.Add(circle);
                    Debug.WriteLine("Added bubble " + point.DataLabel.Text + " at " + circle.CenterX + ", " + circle.CenterY + " with radius " + circle.Radius);

                    i++;
                }
            }

            // ... then make another pass, this time placing the data labels
            i = 0;
            foreach (Excel.Series series in SeriesCollection)
            {
                foreach (Excel.Point point in series.Points())
                {
                    Excel.DataLabel dataLabel = series.DataLabels(1);

                    var bubble = new Circle(point.Left + point.Width / 2, point.Top + point.Height / 2,
                        point.Width / 2, i);

                    Debug.WriteLine("Now finding placement for data label " + dataLabel.Text + "...");

                    var unoccupiedRectangle = FindUnoccupiedRectangleNearCircle(dataLabel.Width, dataLabel.Height,
                        occupiedRectangles, occupiedCircles, bubble,
                        out var leaderLineAttachingCoordinates);

                    dataLabel.Left = unoccupiedRectangle.MinX;
                    dataLabel.Top = unoccupiedRectangle.MinY;

                    leaderLineAttachingPoints.Add(leaderLineAttachingCoordinates);

                    occupiedRectangles.Add(unoccupiedRectangle);
                    Debug.WriteLine("Added data label " + point.DataLabel.Text + " at " + unoccupiedRectangle.MinX + ", " + unoccupiedRectangle.MinY);

                    i++;
                }
            }

            DrawLeaderLines(leaderLineAttachingPoints);
        }

        private static List<List<String>> GetDataMatrixSortedByBubbleSize()
        {
            var dataMatrix = new List<List<string>>();

            foreach (Excel.Range row in _excelWorksheet.UsedRange.Rows)
            {
                var stringList = ConvertRowArrayToStringList(row.Value2);
                if (stringList == null) continue;
                dataMatrix.Add(stringList);
            }

            dataMatrix.Sort((x, y) => Convert.ToDouble(y[3]).CompareTo(Convert.ToDouble(x[3])));

            return dataMatrix;
        }

        private static List<String> ConvertRowArrayToStringList(System.Array array)
        {
            var stringList = new List<string>();
            for (int i = 1; i <= array.Length; i++)
            {
                if (array.GetValue(1, i) == null)
                {
                    return null; //If any cell is empty, ignore this row
                }
                stringList.Add(array.GetValue(1, i).ToString());
            }

            return stringList;
        }

        private static void DrawLeaderLines(List<double[]> leaderLineAttachingPoints)
        {
            var i = 0;
            foreach (Excel.Series series in SeriesCollection)
            {
                foreach (Excel.Point point in series.Points())
                {
                    float bubbleX;
                    float bubbleY;
                    if (leaderLineAttachingPoints[i] != null)
                    {
                        bubbleX = (float) leaderLineAttachingPoints[i][0];
                        bubbleY = (float) leaderLineAttachingPoints[i][1];
                    }
                    else
                    {
                        bubbleX = (float) (point.Left + point.Width / 2);
                        bubbleY = (float) (point.Top + point.Height / 2);
                    }

                    var labelX = (float) point.DataLabel.Left;
                    var labelY = (float) (point.DataLabel.Top + point.DataLabel.Height / 2);
                    var connector = _bubbleChart.Shapes.AddConnector(MsoConnectorType.msoConnectorStraight, bubbleX,
                        bubbleY, labelX, labelY);

                    Debug.WriteLine("Drawing leader line " + point.DataLabel.Text + " from " + bubbleX + ", " + bubbleY + " to " + labelX + ", " + labelY);
                    Debug.WriteLine("The bubble is now at: " + (point.Left + point.Width / 2) + ", " + (point.Top + point.Height / 2));
                    connector.Line.ForeColor.RGB = (int) Excel.XlRgbColor.rgbBlack;
                    i++;
                }
            }
        }

        private static double GetBubbleSizeFromRevenue(double revenue)
        {
            if (revenue <= 200) return 20 *_bubbleScaleFactor;
            if (revenue <= 400) return 50 * _bubbleScaleFactor;
            if (revenue <= 600) return 70 * _bubbleScaleFactor;
            if (revenue <= 800) return 80 * _bubbleScaleFactor;
            return 90 * _bubbleScaleFactor;
        }

        private static Rectangle FindUnoccupiedRectangleNearCircle(double neededWidth, double neededHeight, List<Rectangle> occupiedRectangles, List<Circle> occupiedCircles, Circle bubblePoint, out double[] leaderLineAttachingCoordinates)
        {
            var minX = bubblePoint.CenterX + bubblePoint.Radius + DataLabelDistanceMargin * 2;
            var minY = bubblePoint.CenterY - neededHeight / 2;
            var unrotatedRectangle = new Rectangle(minX, minY, neededWidth, neededHeight);
            var rectangle = unrotatedRectangle;


            leaderLineAttachingCoordinates = null;
            var isAttemptingLeaderLinePlacement = true;

            var rotationAngles = GetRotationAngles();

            var rotationAngleIndex = 0;
            while (DoesRectangleOverlapAnyOccupiedSpot(rectangle, occupiedRectangles, occupiedCircles) ||
                   (leaderLineAttachingCoordinates == null && isAttemptingLeaderLinePlacement))
            {
                if (rotationAngleIndex >= rotationAngles.Count)
                {
                    Debug.WriteLine("Increasing distance from bubble by 5...");
                    minX += 5;
                    unrotatedRectangle = new Rectangle(minX, minY, neededWidth, neededHeight);
                    rectangle = unrotatedRectangle;
                    rotationAngleIndex = 0;
                }
                else
                {
                    Debug.WriteLine("Rotating by " + RotationIncrement + " radians...");
                    rectangle = GetRectangleRotatedAroundPoint(unrotatedRectangle, bubblePoint.CenterX, bubblePoint.CenterY, rotationAngles[rotationAngleIndex]);
                    leaderLineAttachingCoordinates = GetLeaderLinePointBetweenRectangleAndCircle(rectangle, bubblePoint, occupiedCircles);
                    if (leaderLineAttachingCoordinates == null && isAttemptingLeaderLinePlacement)
                    {
                        Debug.WriteLine("No leader line attachment possible at this angle. Continuing...");
                        rotationAngles.RemoveAt(rotationAngleIndex);
                        if (rotationAngles.Count == 0)
                        {
                            Debug.WriteLine("Impossible to place leader line at any angle. Giving up on it...");
                            rotationAngles = GetRotationAngles();
                            rotationAngleIndex = 0;
                            isAttemptingLeaderLinePlacement = false;
                        }
                    }
                    else
                    {
                        rotationAngleIndex++;
                    }
                }
            }

            Debug.WriteLine("Data label placement found!");
            return rectangle;
        }

        private static List<double> GetRotationAngles()
        {
            var rotationAngles = new List<double>();

            for (double v = 0; v <= 2 * Math.PI; v += RotationIncrement)
            {
                rotationAngles.Add(v);
            }

            return rotationAngles;
        }

        private static bool DoesRectangleOverlapAnyOccupiedSpot(Rectangle rectangle, List<Rectangle> occupiedRectangles, List<Circle> occupiedCircles)
        {

            foreach (var occupiedRectangle in occupiedRectangles)
            {
                var overlapsInX = !(rectangle.MaxX + DataLabelDistanceMargin < occupiedRectangle.MinX || rectangle.MinX - DataLabelDistanceMargin > occupiedRectangle.MaxX);
                var overlapsInY = !(rectangle.MaxY + DataLabelDistanceMargin < occupiedRectangle.MinY || rectangle.MinY - DataLabelDistanceMargin > occupiedRectangle.MaxY);

                if (overlapsInX && overlapsInY)
                {
                    return true;
                }
            }
            foreach (var occupiedCircle in occupiedCircles)
            {
                // Circle center is inside rectangle
                if (rectangle.MinX <= occupiedCircle.CenterX &&
                    occupiedCircle.CenterX <= rectangle.MaxX &&
                    rectangle.MinY <= occupiedCircle.CenterY &&
                    occupiedCircle.CenterY <= rectangle.MaxY)
                {
                    return true;
                }
                // Rectangle side intersects circle
                if (IsPointInsideCircle(rectangle.MinX, Clamp(occupiedCircle.CenterY, rectangle.MinY, rectangle.MaxY), occupiedCircle) ||
                    IsPointInsideCircle(rectangle.MaxX, Clamp(occupiedCircle.CenterY, rectangle.MinY, rectangle.MaxY), occupiedCircle) ||
                    IsPointInsideCircle(Clamp(occupiedCircle.CenterX, rectangle.MinX, rectangle.MaxX), rectangle.MinY, occupiedCircle) ||
                    IsPointInsideCircle(Clamp(occupiedCircle.CenterX, rectangle.MinX, rectangle.MaxX), rectangle.MaxY, occupiedCircle))
                {
                    return true;
                }
            }
            return false;
        }

        private static bool IsPointInsideCircle(double x, double y, Circle circle)
        {
            var distanceToCenter = Math.Sqrt(Math.Pow(x - circle.CenterX, 2) + Math.Pow(y - circle.CenterY, 2));
            return distanceToCenter <= circle.Radius;
        }

        private static Rectangle GetRectangleRotatedAroundPoint(Rectangle rectangle, double centerX, double centerY, double rotationAngle)
        {
            var relativeX = rectangle.MinX - centerX;
            var relativeY = rectangle.MinY - centerY;

            var newRelativeX = Math.Cos(rotationAngle) * relativeX - Math.Sin(rotationAngle) * relativeY;
            var newRelativeY = Math.Sin(rotationAngle) * relativeX + Math.Cos(rotationAngle) * relativeY;

            return new Rectangle(centerX + newRelativeX, centerY + newRelativeY, rectangle.Width, rectangle.Height);
        }

        private static double[] GetLeaderLinePointBetweenRectangleAndCircle(Rectangle rectangle, Circle circle,
            List<Circle> occupiedCircles)
        {
            var x = circle.CenterX;
            var y = circle.CenterY;
            var distance = Math.Sqrt(Math.Pow(circle.CenterX - rectangle.MinX, 2) +
                                     Math.Pow(circle.CenterY - (rectangle.MinY + rectangle.Height) / 2, 2));
            var directionX = (circle.CenterX - rectangle.MinX) / distance;
            var directionY = (circle.CenterY - (rectangle.MinY + rectangle.Height / 2)) / distance;

            while (IsPointInsideCircle(x, y, circle))
            {
                var numberOfCirclesOverlappingPoint = 0;
                foreach (var occupiedCircle in occupiedCircles)
                {
                    if (occupiedCircle.ZIndex > circle.ZIndex && IsPointInsideCircle(x, y, occupiedCircle)) numberOfCirclesOverlappingPoint++;
                }

                if (numberOfCirclesOverlappingPoint == 0)
                {
                    return new[] {x, y};
                }

                x -= directionX;
                y -= directionY;
            }

            return null;
        }

        private static double Clamp(double x, double min, double max)
        {
            return Math.Min(Math.Max(x, min), max);
        }

        private struct Rectangle
        {
            public double MinX;
            public double MinY;
            public double MaxX;
            public double MaxY;
            public double Width;
            public double Height;

            public Rectangle(double minX, double minY, double width, double height)
            {
                MinX = minX;
                MinY = minY;
                Width = width;
                Height = height;
                MaxX = minX + width;
                MaxY = minY + height;
            }
        }

        private struct Circle
        {
            public double CenterX;
            public double CenterY;
            public double Radius;
            public double ZIndex;

            public Circle(double x, double y, double radius, double zIndex)
            {
                CenterX = x;
                CenterY = y;
                Radius = radius;
                ZIndex = zIndex;
            }
        }
    }
}
