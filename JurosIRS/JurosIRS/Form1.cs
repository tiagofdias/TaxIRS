using System.Globalization;
using System.Windows.Forms;
using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using OfficeOpenXml;
using System.Text.RegularExpressions;
using static OfficeOpenXml.ExcelErrorValue;

namespace JurosIRS
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        string exchangeRateCsvFilePath;

        private async Task<double> ProcessCsvFileAsync(string csvFilePath)
        {
            double totalCashInEUR = 0.0;

            using (var reader = new StreamReader(csvFilePath))
            {
                // Skip the header row
                var headerLine = await reader.ReadLineAsync();

                while (!reader.EndOfStream)
                {
                    var line = await reader.ReadLineAsync();
                    var values = line.Split(',');

                    DateTime dateandtime = DateTime.Parse(values[1]).Date;
                    string date = dateandtime.ToString("yyyy-MM-dd");

                    double cash = double.Parse(values[2], CultureInfo.InvariantCulture);

                    string currency = values[3].Trim('"');

                    if (currency == "EUR")
                        totalCashInEUR += cash; 
                    else
                    {
                        double exchangeRate = GetExchangeRate(exchangeRateCsvFilePath, currency, date);
                        exchangeRate = Math.Round(1 / exchangeRate, 5);

                        totalCashInEUR += cash * exchangeRate;
                    }
                }
            }

            return Math.Round(totalCashInEUR, 2); ;
        }

        private void btnBancoCSV_Click(object sender, EventArgs e)
        {
            if (openFileDialog2.ShowDialog() == DialogResult.OK) exchangeRateCsvFilePath = openFileDialog2.FileName;
        }

        public static double GetExchangeRate(string filePath, string currencyName, string date)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Set the license context

            FileInfo fileInfo = new FileInfo(filePath);

            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Dados"]; // Assuming the sheet name is 'Dados'

                int totalColumns = worksheet.Dimension.Columns;
                int dateColumn = 1; // Assuming dates are in the first column

                // Find the column that matches the currency name
                int currencyColumn = -1;
                for (int col = 2; col <= totalColumns; col++)
                {
                    string cellValue = worksheet.Cells[2, col].Text; // Assuming currency names are in the 3rd row

                    if (Regex.IsMatch(cellValue, $@"\b{currencyName}\b", RegexOptions.IgnoreCase))
                    {
                        currencyColumn = col;

                        break;
                    }
                }

                if (currencyColumn == -1) throw new Exception("Currency not found.");

                int totalRows = worksheet.Dimension.Rows;

                // Find the row that matches the date
                for (int row = 5; row <= totalRows; row++) // Assuming data starts from the 4th row
                {
                    string rowDate = worksheet.Cells[row, dateColumn].Text;

                    if (rowDate == date) return double.Parse(worksheet.Cells[row, currencyColumn].Text);
                }

                throw new Exception("Exchange rate not found for the specified date.");
            }

        }

        private async void Form1_Load(object sender, EventArgs e)
        {
            this.Hide();
            // Define the year variable
            int year = DateTime.Now.Year - 1;

            MessageBox.Show($"Welcome, Please select the Banco of Portugal's 1999 - {year} Excel file.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);

            if (openFileDialog2.ShowDialog() == DialogResult.OK) exchangeRateCsvFilePath = openFileDialog2.FileName;

            MessageBox.Show($"Now, select the Trading 212's csv file.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string csvFilePath = openFileDialog.FileName;
                double totalCashInEUR = await ProcessCsvFileAsync(csvFilePath);

                DialogResult result = MessageBox.Show($"Total cash in EUR: {totalCashInEUR}. Use this value in the IRS.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);

                if (result == DialogResult.OK) Environment.Exit(0);      

            }
        }
    }
}
