using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace TimeSheetCreator
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Handles the Load event of the Form1 control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs"/> instance containing the event data.</param>
        private void Form1_Load(object sender, EventArgs e)
        {
            PopulateYears();
        }

        /// <summary>
        /// Populates the combo box with 10 years, with the current year in the middle.
        /// </summary>
        private void PopulateYears()
        {
            int year = DateTime.Now.Year;
            _yearComboBox.Items.AddRange(Enumerable.Range(year - 5, 10).Select(x => (object)x).ToArray());
            _yearComboBox.SelectedItem = year + 1;
        }

        /// <summary>
        /// Handles the OnClick event for _saveButton.
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="System.EventArgs"/> instance containing the event data.</param>
        private void SaveButtonClick(object sender, EventArgs e)
        {
            var result = _saveFileDialog.ShowDialog();

            if (result == DialogResult.OK)
            {
                var filename = _saveFileDialog.FileName;

                using (ExcelPackage package = new ExcelPackage())
                {
                    IEnumerable<DateTime> mondays = GetMondays((int)_yearComboBox.SelectedItem);
                    foreach (var monday in mondays)
                    {
                        ExcelWorksheet ws = package.Workbook.Worksheets.Add(monday.ToString("yyyyMMdd"));
                        ws.View.ShowGridLines = false;

                        CreateLabels(ws, monday, _nameTextBox.Text);
                        MergeColumns(ws);
                        CreateLines(ws);
                        SetColumnWidths(ws);
                        AddCellFormats(ws);
                        CreateFormulas(ws);
                        GenerateDefaultData(ws);
                        AddCellFormats(ws); // call again
                    }

                    using (var fs = new FileStream(filename, FileMode.OpenOrCreate))
                    {
                        package.SaveAs(fs);
                    }
                }
            }
        }

        /// <summary>
        /// Sets the column widths to a preset size.
        /// </summary>
        /// <param name="worksheet">The worksheet.</param>
        private void SetColumnWidths(ExcelWorksheet worksheet)
        {
            worksheet.Column(1).Width = ExcelHelper.Pixel2ColumnWidth(worksheet, 87);
            worksheet.Column(2).Width = ExcelHelper.Pixel2ColumnWidth(worksheet, 93);
            worksheet.Column(3).Width = ExcelHelper.Pixel2ColumnWidth(worksheet, 40);
            worksheet.Column(4).Width = ExcelHelper.Pixel2ColumnWidth(worksheet, 40);
            worksheet.Column(5).Width = ExcelHelper.Pixel2ColumnWidth(worksheet, 40);
            worksheet.Column(6).Width = ExcelHelper.Pixel2ColumnWidth(worksheet, 40);
            worksheet.Column(7).Width = ExcelHelper.Pixel2ColumnWidth(worksheet, 55);
            worksheet.Column(8).Width = ExcelHelper.Pixel2ColumnWidth(worksheet, 54);
            worksheet.Column(9).Width = ExcelHelper.Pixel2ColumnWidth(worksheet, 61);
            worksheet.Column(10).Width = ExcelHelper.Pixel2ColumnWidth(worksheet, 72);
        }

        /// <summary>
        /// Generates some default data to fill out the worksheet.
        /// </summary>
        /// <param name="worksheet">The worksheet.</param>
        private void GenerateDefaultData(ExcelWorksheet worksheet)
        {
            //TODO: Factor in holidays
            
            // Sick & Personal hours
            worksheet.Cells["H5:I11"].Value = 0;
            
            // Daily hours
            worksheet.Cells["C5:C9"].Value = DateTime.Parse("7:00 AM").TimeOfDay;
            worksheet.Cells["D5:D9"].Value = DateTime.Parse("12:30 PM").TimeOfDay;
            worksheet.Cells["E5:E9"].Value = DateTime.Parse("1:00 PM").TimeOfDay;
            worksheet.Cells["F5:F9"].Value = DateTime.Parse("4:00 PM").TimeOfDay;

            // Special Wednesdays lunch
            worksheet.Cells["D7"].Value = DateTime.Parse("11:15 AM").TimeOfDay;
            worksheet.Cells["E7"].Value = DateTime.Parse("12:30 PM").TimeOfDay;
        }

        /// <summary>
        /// Creates all the formulas used in a worksheet.
        /// </summary>
        /// <param name="worksheet">The worksheet.</param>
        private void CreateFormulas(ExcelWorksheet worksheet)
        {
            // Dates
            worksheet.Cells["A5"].FormulaR1C1 = "IF(B2 <> \"\", B2, \"\")";
            worksheet.Cells["A6"].FormulaR1C1 = "IF(A5 <> \"\", A5+1, \"\")";
            worksheet.Cells["A7"].FormulaR1C1 = "IF(A6 <> \"\", A6+1, \"\")";
            worksheet.Cells["A8"].FormulaR1C1 = "IF(A7 <> \"\", A7+1, \"\")";
            worksheet.Cells["A9"].FormulaR1C1 = "IF(A8 <> \"\", A8+1, \"\")";
            worksheet.Cells["A10"].FormulaR1C1 = "IF(A9 <> \"\", A9+1, \"\")";
            worksheet.Cells["A11"].FormulaR1C1 = "IF(A10 <> \"\", A10+1, \"\")";

            // Days
            worksheet.Cells["B5"].FormulaR1C1 = "A5";
            worksheet.Cells["B6"].FormulaR1C1 = "A6";
            worksheet.Cells["B7"].FormulaR1C1 = "A7";
            worksheet.Cells["B8"].FormulaR1C1 = "A8";
            worksheet.Cells["B9"].FormulaR1C1 = "A9";
            worksheet.Cells["B10"].FormulaR1C1 = "A10";
            worksheet.Cells["B11"].FormulaR1C1 = "A11";

            // Straight Hours
            worksheet.Cells["G5"].FormulaR1C1 = "((D5-C5)+(F5-E5))*24";
            worksheet.Cells["G6"].FormulaR1C1 = "((D6-C6)+(F6-E6))*24";
            worksheet.Cells["G7"].FormulaR1C1 = "((D7-C7)+(F7-E7))*24";
            worksheet.Cells["G8"].FormulaR1C1 = "((D8-C8)+(F8-E8))*24";
            worksheet.Cells["G9"].FormulaR1C1 = "((D9-C9)+(F9-E9))*24";
            worksheet.Cells["G10"].FormulaR1C1 = "IF(((D10-C10)+(F10-E10))*24>8,8,((D10-C10)+(F10-E10))*24)";
            worksheet.Cells["G11"].FormulaR1C1 = "IF(((D11-C11)+(F11-E11))*24>8,8,((D11-C11)+(F11-E11))*24)";

            // Daily Totals
            worksheet.Cells["J5"].FormulaR1C1 = "SUM(G5:I5)";
            worksheet.Cells["J6"].FormulaR1C1 = "SUM(G6:I6)";
            worksheet.Cells["J7"].FormulaR1C1 = "SUM(G7:I7)";
            worksheet.Cells["J8"].FormulaR1C1 = "SUM(G8:I8)";
            worksheet.Cells["J9"].FormulaR1C1 = "SUM(G9:I9)";
            worksheet.Cells["J10"].FormulaR1C1 = "SUM(G10:I10)";
            worksheet.Cells["J11"].FormulaR1C1 = "SUM(G11:I11)";

            // Hourly Totals
            worksheet.Cells["G12"].FormulaR1C1 = "SUM(G5:G11)";
            worksheet.Cells["H12"].FormulaR1C1 = "SUM(H5:H11)";
            worksheet.Cells["I12"].FormulaR1C1 = "SUM(I5:I11)";
            worksheet.Cells["J12"].FormulaR1C1 = "IF(SUM(G12:I12)=SUM(J5:J11),SUM(G12:I12),\"Error!\")";

        }

        /// <summary>
        /// Adds cell formatting to a worksheet.
        /// </summary>
        /// <param name="worksheet">The worksheet.</param>
        private void AddCellFormats(ExcelWorksheet worksheet)
        {
            // Dates
            worksheet.Cells["B2"].Style.Numberformat.Format = "MM/dd/yyyy";
            worksheet.Cells["A5:A11"].Style.Numberformat.Format = "MM/dd/yyyy";

            // Days
            worksheet.Cells["B5:B11"].Style.Numberformat.Format = "dddd";

            // Times
            //worksheet.Cells[5, 3, 11, 6].Style.Numberformat.Format = "13:30";
            worksheet.Cells["C5:F9"].Style.Numberformat.Format = "hh:mm;@";

        }

        /// <summary>
        /// Creates lines on the worksheet to make it look pretty.
        /// </summary>
        /// <param name="worksheet">The worksheet.</param>
        private void CreateLines(ExcelWorksheet worksheet)
        {
            BlackBoxRange(worksheet.Cells["A3:J4"], worksheet);
            BlackBoxRange(worksheet.Cells["A5:J11"], worksheet);
            BlackBoxRange(worksheet.Cells["A3:A4"], worksheet);
            BlackBoxRange(worksheet.Cells["B3:B4"], worksheet);
            BlackBoxRange(worksheet.Cells["C3:F4"], worksheet);
            BlackBoxRange(worksheet.Cells["G3:I4"], worksheet);
            GraySidesRange(worksheet.Cells["C5:F11"], worksheet);
            GraySidesRange(worksheet.Cells["G5:I11"], worksheet);
            BlackBoxRange(worksheet.Cells["G12:I12"], worksheet);
            BlackBoxRange(worksheet.Cells["J12"], worksheet);
            
            worksheet.Cells["A12:J12"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells["A12:J12"].Style.Fill.BackgroundColor.SetColor(Color.Gray);
        }

        /// <summary>
        /// Creates a black box around a certain range of cells.
        /// </summary>
        /// <param name="range">The range.</param>
        /// <param name="worksheet">The worksheet.</param>
        private void BlackBoxRange(ExcelRange range, ExcelWorksheet worksheet)
        {
            var start = range.Start;
            var end = range.End;

            for (int i = start.Row; i <= end.Row; i++)
            {
                worksheet.Cells[i, start.Column].Style.Border.Left.Style = worksheet.Cells[i, end.Column].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[i, start.Column].Style.Border.Left.Color.SetColor(Color.Black);
                worksheet.Cells[i, end.Column].Style.Border.Right.Color.SetColor(Color.Black);
            }

            for (int i = start.Column; i <= end.Column; i++)
            {
                worksheet.Cells[start.Row, i].Style.Border.Top.Style = worksheet.Cells[end.Row, i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[start.Row, i].Style.Border.Top.Color.SetColor(Color.Black);
                worksheet.Cells[end.Row, i].Style.Border.Bottom.Color.SetColor(Color.Black);
            }
        }

        /// <summary>
        /// Creates gray sides on the left and right of a certain range of cells.
        /// </summary>
        /// <param name="range">The range.</param>
        /// <param name="worksheet">The worksheet.</param>
        private void GraySidesRange(ExcelRange range, ExcelWorksheet worksheet)
        {
            var start = range.Start;
            var end = range.End;

            for (int i = start.Row; i <= end.Row; i++)
            {
                worksheet.Cells[i, start.Column].Style.Border.Left.Style = worksheet.Cells[i, end.Column].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[i, start.Column].Style.Border.Left.Color.SetColor(Color.Gray);
                worksheet.Cells[i, end.Column].Style.Border.Right.Color.SetColor(Color.Gray);
            }
        }

        /// <summary>
        /// Merges several cells in the worksheet to make it look pretty.
        /// </summary>
        /// <param name="worksheet">The worksheet.</param>
        private void MergeColumns(ExcelWorksheet worksheet)
        {
            // Time
            MergeAndCenterRange(worksheet.Cells["C3:F3"]);
            
            // Hourly Breakdown
            MergeAndCenterRange(worksheet.Cells["G3:I3"]);

            // Weekly Totals
            MergeAndCenterRange(worksheet.Cells["A12:F12"]);
        }

        /// <summary>
        /// Merges and centers a range of cells.
        /// </summary>
        /// <param name="range">The range.</param>
        private void MergeAndCenterRange(ExcelRange range)
        {
            range.Merge = true;
            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        }

        /// <summary>
        /// Creates labels for the worksheet.
        /// </summary>
        /// <param name="worksheet">The worksheet.</param>
        /// <param name="monday">The date of the Monday used on this specific worksheet.</param>
        /// <param name="name">The name of the person this worksheet is for.</param>
        private void CreateLabels(ExcelWorksheet worksheet, DateTime monday, string name)
        {
            worksheet.Cells["B1"].Value = name;
            worksheet.Cells["B2"].Value = monday.Date;

            CreateBoldText(worksheet.Cells["A1"], "Name:");
            CreateBoldText(worksheet.Cells["A2"], "Time Period:");
            CreateBoldText(worksheet.Cells["A4"], "Date");
            CreateBoldText(worksheet.Cells["B4"], "Day Of Week");
            CreateBoldText(worksheet.Cells["C4"], "In");
            CreateBoldText(worksheet.Cells["D4"], "Out");
            CreateBoldText(worksheet.Cells["E4"], "In");
            CreateBoldText(worksheet.Cells["F4"], "Out");
            CreateBoldText(worksheet.Cells["G4"], "Straight");
            CreateBoldText(worksheet.Cells["H4"], "Holiday");
            CreateBoldText(worksheet.Cells["I4"], "Personal");
            CreateBoldText(worksheet.Cells["J4"], "Daily Total");
            
            // Merged columns
            CreateBoldText(worksheet.Cells["C3"], "Time");
            CreateBoldText(worksheet.Cells["G3"], "Hourly Breakdown");
            CreateBoldText(worksheet.Cells["A12"], "Weekly Totals");
        }

        /// <summary>
        /// Creates bold text in a certain range of cells.
        /// </summary>
        /// <param name="range">The range.</param>
        /// <param name="text">The text.</param>
        private void CreateBoldText(ExcelRange range, string text)
        {
            range.Value = text;
            range.Style.Font.Bold = true;
        }

        /// <summary>
        /// Gets a list of all the mondays of weeks that have days in the year. So if Jan 1st is a Sunday, the first Monday will be Dec 26th, the previous year.
        /// </summary>
        /// <param name="year">The year.</param>
        private IEnumerable<DateTime> GetMondays(int year)
        {
            List<DateTime> mondays = new List<DateTime>();

            // First, try setting it to January 1st, in case that is a Monday
            DateTime firstMonday = new DateTime(year, 1, 1);

            if (firstMonday.DayOfWeek != DayOfWeek.Monday)
            {
                // If it is not a Monday, let's adjust
                firstMonday = firstMonday.AddDays((int)firstMonday.DayOfWeek - 6);
            }

            // extrapolate all weeks until the end of the year
            DateTime nextMonday = firstMonday;
            while (nextMonday.Year <= year)
            {
                mondays.Add(nextMonday);
                nextMonday = nextMonday.AddDays(7);
            }
            
            return mondays;
        }
    }
}
