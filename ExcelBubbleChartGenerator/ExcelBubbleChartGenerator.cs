using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelBubbleChartGenerator
{
    public static class ExcelBubbleChartGenerator
    {
        private const double DataLabelDistanceMargin = 5;
        private const double RotationIncrement = Math.PI / 8;

        private const double DefaultBubbleScale = 1;

        private const double HeuristicExtraWidth = 267;
        private const double HeuristicExtraHeight = 168;

        private const float DotWidth = 2;
        private const double LeaderLineAttachingPointMargin = 5;

        private const float BubbleLegendLabelMargin = 5;
        private const float BubbleLegendLeft = 20;
        private const float BubbleLegendTop = 50;
        private const float RevenueTitleLength = 100;
        private const float RevenueTextLabelLength = 55;
        private const float RevenueTextLabelHeight = 20;

        private static readonly Dictionary<int, string> ProjectTypeNames = new Dictionary<int, string>
        {
            {1, "Lösning/projekt"},
            {2, "Konsulttjänst"},
            {3, "Indirekt"}
        };

        private static readonly Dictionary<int, int> ProjectTypeColors = new Dictionary<int, int>
        {
            {1, (int) Excel.XlRgbColor.rgbBisque},
            {2, (int) Excel.XlRgbColor.rgbYellowGreen},
            {3, (int) Excel.XlRgbColor.rgbRed}
        };

        public static void GenerateBubbleChart(double chartWidth, double chartHeight, double bubbleScaleFactor, string dataFilePath, string outputFilePath, int worksheetToOpen)
        {
            Debug.WriteLine("Starting up...");

            Excel.Application excelApp = null;
            Excel.Workbook excelWorkbook = null;
            Excel.Worksheet excelWorksheet = null;

            Excel.Workbook sourceDataWorkbook = null;

            try
            {
                excelApp = new Excel.Application();
                excelWorkbook = excelApp.Workbooks.Add(1);
                excelWorksheet = (Excel.Worksheet) excelWorkbook.Sheets[1];

                sourceDataWorkbook = excelApp.Workbooks.Open(dataFilePath);

                Debug.WriteLine("Setting up chart properties...");
                var bubbleChart = CreateNewBubbleChart(excelWorksheet, chartWidth, chartHeight, bubbleScaleFactor);
                Debug.WriteLine("Adding data points...");
                AddDataPoints(bubbleChart, sourceDataWorkbook.Sheets[worksheetToOpen]);
                Debug.WriteLine("Adding legend...");
                AddLegendAndClearDummySeriesNames(bubbleChart);
                Debug.WriteLine("Placing data labels...");
                SpreadOutDataLabels(bubbleChart);
                Debug.WriteLine("Exporting image...");
                bubbleChart.Export(outputFilePath, "PNG");
                Process.Start(outputFilePath);

                Debug.Write("Success! ");
            }
            finally
            {
                if (excelWorkbook != null)
                {
                    excelWorkbook.Saved = true; // Avoid prompting user to save
                    excelWorkbook.Close();
                }

                if (sourceDataWorkbook != null)
                {
                    sourceDataWorkbook.Saved = true; // Avoid prompting user to save
                    sourceDataWorkbook.Close();
                }

                excelApp?.Quit();
                if (excelWorksheet != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorksheet);
                if (excelWorkbook != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorkbook);
                if (sourceDataWorkbook != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(sourceDataWorkbook);
                if (excelApp != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            }
        }

        private static Excel.Chart CreateNewBubbleChart(Excel.Worksheet excelWorksheet, double chartWidth, double chartHeight, double bubbleScaleFactor)
        {
            var excelChartObjects = (Excel.ChartObjects)excelWorksheet.ChartObjects();
            var chartObject = excelChartObjects.Add(10, 80, chartWidth - HeuristicExtraWidth, chartHeight - HeuristicExtraHeight);
            var bubbleChart = chartObject.Chart;

            bubbleChart.ChartType = Excel.XlChartType.xlBubble3DEffect;

            bubbleChart.ChartGroups(1).BubbleScale = Clamp(100 * bubbleScaleFactor * DefaultBubbleScale, 0, 300);

            var xAxis = (Excel.Axis) bubbleChart.Axes(Excel.XlAxisType.xlCategory);
            xAxis.HasTitle = true;
            xAxis.AxisTitle.Caption = "Medeltimpris";
            xAxis.HasMajorGridlines = true;
            xAxis.MajorGridlines.Format.Line.ForeColor.RGB = (int)Excel.XlRgbColor.rgbGainsboro;

            var yAxis = (Excel.Axis) bubbleChart.Axes(Excel.XlAxisType.xlValue);
            yAxis.HasTitle = true;
            yAxis.AxisTitle.Caption = "TG2 Medel";
            yAxis.HasMajorGridlines = true;
            yAxis.TickLabels.NumberFormat = "0%";
            yAxis.MajorGridlines.Format.Line.ForeColor.RGB = (int)Excel.XlRgbColor.rgbGainsboro;

            return bubbleChart;
        }

        private static void AddDataPoints(Excel.Chart bubbleChart, Excel.Worksheet sourceWorksheet)
        {
            var dataMatrix = GetDataMatrixFromExcelWorksheet(sourceWorksheet);

            var seriesCollection = bubbleChart.SeriesCollection();

            var minXValue = double.MaxValue;
            var minYValue = double.MaxValue;
            var maxXValue = double.MinValue;
            var maxYValue = double.MinValue;

            foreach (var row in dataMatrix)
            {
                var series = seriesCollection.NewSeries();
                series.Has3DEffect = true;
                series.HasDataLabels = true;

                var dataLabelText = row[0];
                var xValue = Convert.ToDouble(row[1]);
                var yValue = Convert.ToDouble(row[2]);
                var revenue = Convert.ToDouble(row[3]) / 1000;
                var projectTypeId = GetProjectTypeIdFromProjectTypeString(row[4]);

                var bubbleSize = GetBubbleSizeFromRevenue(revenue);

                series.Format.Fill.ForeColor.RGB = ProjectTypeColors[projectTypeId];

                series.XValues = new[] { xValue };
                series.Values = new[] { yValue };
                series.BubbleSizes = new[] { bubbleSize };

                Excel.DataLabel dataLabel = series.DataLabels(1);
                dataLabel.Text = dataLabelText;

                if (xValue < minXValue) minXValue = xValue;
                if (yValue < minYValue) minYValue = yValue;
                if (xValue > maxXValue) maxXValue = xValue;
                if (yValue > maxYValue) maxYValue = yValue;
            }

            var xAxis = (Excel.Axis)bubbleChart.Axes(Excel.XlAxisType.xlCategory);
            xAxis.MinimumScale = (Math.Floor(minXValue / 100) - 1) * 100;
            xAxis.MaximumScale = (Math.Ceiling(maxXValue / 100) + 1) * 100;
            var yAxis = (Excel.Axis)bubbleChart.Axes(Excel.XlAxisType.xlValue);
            yAxis.MinimumScale = (Math.Floor(minYValue * 10) - 1) / 10;
            yAxis.MaximumScale = (Math.Ceiling(maxYValue * 10) + 1) / 10;
        }

        private static void AddLegendAndClearDummySeriesNames(Excel.Chart chart)
        {
            var legend = chart.Legend;

            var seriesCollection = chart.SeriesCollection();

            for (var i = 1; i <= ProjectTypeNames.Count; i++)
            {
                var series = seriesCollection.NewSeries();
                series.Name = ProjectTypeNames[i];
                series.Format.Fill.ForeColor.RGB = ProjectTypeColors[i];
            }

            while (legend.LegendEntries().Count > 3)
            {
                legend.LegendEntries(1).Delete();
            }

            legend.Position = Excel.XlLegendPosition.xlLegendPositionTop;

            DrawBubbleLegend(chart);

            legend.Top = chart.PlotArea.Height + BubbleLegendTop;
            legend.Left = BubbleLegendLeft;
        }

        private static void SpreadOutDataLabels(Excel.Chart chart)
        {
            var occupiedRectangles = new List<Rectangle>();
            var occupiedCircles = new List<Circle>();
            var leaderLineAttachingPoints = new List<double[]>();

            var seriesCollection = chart.SeriesCollection();

            // First populated the list with all data bubbles...
            double i = 0;
            foreach (Excel.Series series in seriesCollection)
            {
                foreach (Excel.Point point in series.Points())
                {
                    var circle = new Circle(point.Left + point.Width / 2, point.Top + point.Height / 2, point.Width / 2, i);
                    occupiedCircles.Add(circle);
                    Debug.WriteLine("Added bubble " + point.DataLabel.Text + " at " + circle.CenterX + ", " + circle.CenterY + " with radius " + circle.Radius);

                    i++;
                }
            }

            // ... and denote the X-axis as an occupied spot...
            var xAxis = (Excel.Axis)chart.Axes(Excel.XlAxisType.xlCategory);
            occupiedRectangles.Add(new Rectangle(xAxis.Left, xAxis.Top, xAxis.Width, xAxis.Height));

            // ... then make another pass, this time placing the data labels
            i = 0;
            foreach (Excel.Series series in seriesCollection)
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

            DrawLeaderLines(chart, leaderLineAttachingPoints);
        }

        private static List<List<string>> GetDataMatrixFromExcelWorksheet(Excel.Worksheet worksheet)
        {
            var dataMatrix = new List<List<string>>();

            foreach (Excel.Range row in worksheet.UsedRange.Rows)
            {
                var stringList = ConvertRowArrayToStringList(row.Value2);
                if (!IsStringListValidData(stringList)) continue;
                dataMatrix.Add(stringList);
            }

            dataMatrix.Sort((x, y) => Convert.ToDouble(y[3]).CompareTo(Convert.ToDouble(x[3])));

            return dataMatrix;
        }

        private static bool IsStringListValidData(List<string> stringList)
        {
            var projectNameString = stringList[0];
            var hourRateString = stringList[1];
            var TG2String = stringList[2];
            var revenueString = stringList[3];
            var projectTypeString = stringList[4];

            return projectNameString != string.Empty &&
                   Double.TryParse(hourRateString, out var _) &&
                   Double.TryParse(TG2String, out _) &&
                   Double.TryParse(revenueString, out var _) &&
                   IsValidProjectType(projectTypeString);
        }

        private static bool IsValidProjectType(string projectType)
        {
            return string.IsNullOrEmpty(projectType) || projectType.ToLower() == "lösning" || projectType.ToLower() == "resurs" ||
                   projectType.ToLower() == "indirekt";
        }

        private static int GetProjectTypeIdFromProjectTypeString(string projectTypeString)
        {
            switch (projectTypeString?.ToLower())
            {
                default:
                case "lösning":
                    return 1;
                case "resurs":
                    return 2;
                case "indirekt":
                    return 3;
            }
        }

        private static List<String> ConvertRowArrayToStringList(Array array)
        {
            var stringList = new List<string>();
            for (var i = 1; i <= array.Length; i++)
            {
                stringList.Add(array.GetValue(1, i)?.ToString());
            }

            return stringList;
        }

        private static double GetBubbleSizeFromRevenue(double revenue)
        {
            return Math.Max(revenue / 100, 0);
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
                    minX += 5;
                    unrotatedRectangle = new Rectangle(minX, minY, neededWidth, neededHeight);
                    rectangle = unrotatedRectangle;
                    rotationAngleIndex = 0;
                }
                else
                {
                    rectangle = GetRectangleRotatedAroundPoint(unrotatedRectangle, bubblePoint.CenterX, bubblePoint.CenterY, rotationAngles[rotationAngleIndex]);
                    leaderLineAttachingCoordinates = GetLeaderLinePointBetweenRectangleAndCircle(rectangle, bubblePoint, occupiedCircles);
                    if (leaderLineAttachingCoordinates == null && isAttemptingLeaderLinePlacement)
                    {
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

            if (leaderLineAttachingCoordinates != null)
            {
                occupiedCircles.Add(new Circle(leaderLineAttachingCoordinates[0], leaderLineAttachingCoordinates[1], LeaderLineAttachingPointMargin, double.MaxValue));
            }

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
                    var closestXToCenter = x;
                    var closestYToCenter = y;
                    while (IsPointInsideCircle(x, y, circle))
                    {
                        x -= directionX;
                        y -= directionY;
                    }

                    return new[] {(x + closestXToCenter) / 2, (y + closestYToCenter) / 2};
                }

                x -= directionX;
                y -= directionY;
            }

            return null;
        }

        private static void DrawLeaderLines(Excel.Chart chart, List<double[]> leaderLineBubbleAttachingPoint)
        {
            var i = 0;
            var seriesCollection = chart.SeriesCollection();
            foreach (Excel.Series series in seriesCollection)
            {
                foreach (Excel.Point point in series.Points())
                {
                    float bubbleX;
                    float bubbleY;
                    if (leaderLineBubbleAttachingPoint[i] != null)
                    {
                        bubbleX = (float)leaderLineBubbleAttachingPoint[i][0];
                        bubbleY = (float)leaderLineBubbleAttachingPoint[i][1];
                    }
                    else
                    {
                        bubbleX = (float)(point.Left + point.Width / 2);
                        bubbleY = (float)(point.Top + point.Height / 2);
                    }

                    var leaderLineLabelAttachingPoint = GetLeaderLineDataLabelAttachingPoint(point.DataLabel, bubbleX, bubbleY);

                    var labelX = (float) leaderLineLabelAttachingPoint[0];
                    var labelY = (float) leaderLineLabelAttachingPoint[1];

                    var connector = chart.Shapes.AddConnector(MsoConnectorType.msoConnectorStraight, bubbleX,
                        bubbleY, labelX, labelY);
                    connector.Line.ForeColor.RGB = (int)Excel.XlRgbColor.rgbBlack;

                    var dotAtLabel = chart.Shapes.AddShape(MsoAutoShapeType.msoShapeOval,
                        labelX - DotWidth / 2,
                        labelY - DotWidth / 2,
                        DotWidth,
                        DotWidth);
                    dotAtLabel.Fill.ForeColor.RGB = (int)Excel.XlRgbColor.rgbBlack;

                    var dotAtBubble = chart.Shapes.AddShape(MsoAutoShapeType.msoShapeOval,
                        bubbleX - DotWidth / 2,
                        bubbleY - DotWidth / 2,
                        DotWidth,
                        DotWidth);
                    dotAtBubble.Fill.ForeColor.RGB = (int)Excel.XlRgbColor.rgbBlack;

                    Debug.WriteLine("Drawing leader line " + point.DataLabel.Text + " from " + bubbleX + ", " + bubbleY + " to " + labelX + ", " + labelY);
                    i++;
                }
            }
        }

        private static double[] GetLeaderLineDataLabelAttachingPoint(Excel.DataLabel dataLabel, double bubbleX, double bubbleY)
        {
            var labelAttachingX = Clamp(bubbleX, dataLabel.Left, dataLabel.Left + dataLabel.Width);
            var labelAttachingY = Clamp(bubbleY, dataLabel.Top, dataLabel.Top + dataLabel.Height);

            return new[] {labelAttachingX, labelAttachingY};
        }

        private static void DrawBubbleLegend(Excel.Chart chart)
        {
            var revenues = new double[] {200, 400, 600, 800, 1000};
            const double revenueIncrement = 200;
            var zeroArray = revenues.Select(x => 0.0).ToArray();
            var bubbleSizes = revenues.Select(x => GetBubbleSizeFromRevenue(x - revenueIncrement / 2)).ToArray();

            Excel.Series dummyBubbleSeries = chart.SeriesCollection().NewSeries();
            dummyBubbleSeries.XValues = zeroArray;
            dummyBubbleSeries.Values = zeroArray;
            dummyBubbleSeries.BubbleSizes = bubbleSizes;

            var bubbleWidths = new List<float>();

            foreach (Excel.Point bubble in dummyBubbleSeries.Points())
            {
                bubbleWidths.Add((float) bubble.Width);
            }

            var bubbleLegendNeededSpace = bubbleWidths.Max() + BubbleLegendTop;

            chart.PlotArea.Height -= bubbleLegendNeededSpace;

            float currentLeft = BubbleLegendLeft + RevenueTitleLength;

            for (var i = 0; i < bubbleWidths.Count; i++)
            {
                var bubbleWidth = bubbleWidths[i];

                var topMargin = (float) chart.PlotArea.Height + BubbleLegendTop + bubbleWidths.Max() / 2 - bubbleWidth / 2;

                var oval = chart.Shapes.AddShape(MsoAutoShapeType.msoShapeOval,
                    currentLeft,
                    topMargin,
                    bubbleWidth,
                    bubbleWidth);
                oval.Fill.ForeColor.RGB = (int)Excel.XlRgbColor.rgbWhite;
                oval.Line.ForeColor.RGB = (int)Excel.XlRgbColor.rgbBlack;

                var label = chart.Shapes.AddLabel(MsoTextOrientation.msoTextOrientationHorizontal,
                    currentLeft + bubbleWidth + BubbleLegendLabelMargin,
                    topMargin + bubbleWidth / 2 - RevenueTextLabelHeight / 2,
                    RevenueTextLabelLength,
                    RevenueTextLabelHeight);

                var revenueText = string.Empty;
                if (i == 0)
                {
                    revenueText = "<" + revenues[i];
                }
                else if (i == revenues.Length - 1)
                {
                    revenueText = ">" + revenues[i-1];
                }
                else
                {
                    revenueText = revenues[i - 1] + "-" + revenues[i];
                }

                label.TextFrame.Characters().Text = revenueText;

                currentLeft += bubbleWidth + RevenueTextLabelLength + BubbleLegendLabelMargin*2;
            }

            var revenueLabel = chart.Shapes.AddLabel(MsoTextOrientation.msoTextOrientationHorizontal,
                BubbleLegendLeft,
                BubbleLegendTop + bubbleWidths.Max() / 2 + (float) chart.PlotArea.Height - RevenueTextLabelHeight / 2,
                RevenueTitleLength,
                RevenueTextLabelHeight);
            revenueLabel.TextFrame.Characters().Text = "Omsättning kkr:";

            dummyBubbleSeries.Delete();
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
