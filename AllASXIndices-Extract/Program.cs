using System.Data;
using Newtonsoft.Json;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace AllASXIndices_Extract
{
    internal class ExtractIndices
    {
        public ChromeDriver _driver;
        public readonly string testUrl = "https://www.marketindex.com.au/asx-indices";
        private WebDriverWait _wait;

        // Create a new DataTable.
        public DataTable _indicesTable = new DataTable("Indices");
        public DataSet _dtSet = new DataSet();
        public List<string> _columnHeaders = new List<string>();

        public ExtractIndices()
        {
            ChromeOptions chromeOptions = new ChromeOptions();
            chromeOptions.PageLoadStrategy = PageLoadStrategy.None;
            chromeOptions.AddArgument("start-maximized");
            chromeOptions.AddArgument("−−incognito");
            _driver = new ChromeDriver(chromeOptions);
            _driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(20);
            _wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(10));

            _indicesTable.Dispose();
            _indicesTable.Clear();
        }

        private void KillProcess(string processName)
        {
            var command = string.Format("/C taskkill /f /im {0}", processName);
            System.Diagnostics.ProcessStartInfo p;
            p = new System.Diagnostics.ProcessStartInfo("cmd.exe", command);
            System.Diagnostics.Process proc = new System.Diagnostics.Process();
            proc.StartInfo = p;
            proc.Start();
            proc.WaitForExit();
            proc.Close();
        }

        private void CloseBrowser()
        {
            try
            {
                _driver.Close();
            }
            catch { }
            finally
            {
                KillProcess("chromedriver.exe");
            }
        }

        private void AddDataTableColumn(string coulmnName)
        {
            DataColumn _dtColumn = new DataColumn();
            _dtColumn.DataType = typeof(string);
            _dtColumn.ColumnName = coulmnName;
            _dtColumn.Caption = string.Format("Watchlist {0}", coulmnName);
            _dtColumn.ReadOnly = false;
            _dtColumn.Unique = false;
            // Add column to the DataColumnCollection.
            _indicesTable.Columns.Add(_dtColumn);
        }

        private void AddDataTableRow(List<string> columnNames, List<string> rowValues)
        {
            Console.WriteLine("D Start");
            DataRow myDataRow = _indicesTable.NewRow();

            int columnCount = columnNames.Count();

            for (int i = 0; i < columnCount; i++)
            {
                myDataRow[columnNames[i]] = rowValues[i];
            }
            _indicesTable.Rows.Add(myDataRow);
            Console.WriteLine("D Finish");
        }

        private string DataTableToJSONWithJSONNet(DataTable table)
        {
            string jsonString = string.Empty;
            jsonString = JsonConvert.SerializeObject(table);
            return jsonString;
        }

        private string ExtractIndicesLiveData()
        {
            string jsonResult = string.Empty;

            try
            {
                _driver.Navigate().GoToUrl(testUrl);

                string indicesTableXpath = "//table[contains(@class, 'mi-table')]";

                var indicesTable = _wait.Until(el=>el.FindElement(By.XPath(indicesTableXpath)));

                Assert.IsTrue(indicesTable.Displayed, "Failed to load or find Indices Table");

                // Add column
                var tableColumnHeaders = indicesTable.FindElements(By.XPath("//thead/tr/th"));
                Assert.IsNotNull(tableColumnHeaders, "Failed to find Indices table headers");

                foreach (var columnHeader in tableColumnHeaders)
                {
                    var columnName = columnHeader.Text.Trim();
                    AddDataTableColumn(columnName);
                    _columnHeaders.Add(columnName);
                }

                // Add rows
                var tableBody = indicesTable.FindElement(By.XPath("//tbody"));
                var tableBodyRows = tableBody.FindElements(By.XPath("//tr[@selected_period='day']"));

                foreach (var row in tableBodyRows)
                {
                    var rowValues = new List<string>();
                    var rowCells = row.FindElements(By.TagName("td"));
                    foreach (var rowCell in rowCells)
                    {
                        var rowValue = rowCell.Text.Trim().Replace("^", string.Empty).Replace("\r\n", " ");
                        rowValues.Add(rowValue);
                    }
                    AddDataTableRow(_columnHeaders, rowValues);
                }


                // Add custTable to the DataSet.
                _dtSet.Tables.Add(_indicesTable);

                 jsonResult = DataTableToJSONWithJSONNet(_indicesTable);

            }
            catch (Exception ex) { throw new Exception(ex.Message); }
            finally
            {
                CloseBrowser();
            }

            return (jsonResult);
        }

        static void Main(string[] args)
        {
            ExtractIndices extractIndices = new ExtractIndices();
            var result = extractIndices.ExtractIndicesLiveData();
            Console.WriteLine(result.ToString());
        }
    }
}