using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using OfficeOpenXml;
using System.Globalization;
class Program
{
    static void Main()
    {
        // Initialize the ChromeDriver
        using (IWebDriver driver = new ChromeDriver())
        {
            // Navigate to the website
            driver.Navigate().GoToUrl("https://www.g2b.go.kr/index.jsp");

            // Wait for the page to load completely
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

            // Find the search input by its ID and enter a search term
            var searchInput = driver.FindElement(By.Id("bidNm"));
            searchInput.SendKeys("RPA");

            // Calculate the today dates
            DateTime today = DateTime.Today;
            DateTime fiveDaysAgo = today.AddDays(-4);

            // Format the dates as YYYY/MM/DD
            string todayFormatted = today.ToString("yyyy/MM/dd", CultureInfo.InvariantCulture);
            string fiveDaysAgoFormatted = fiveDaysAgo.ToString("yyyy/MM/dd", CultureInfo.InvariantCulture);

            // Find the start date input by its ID and enter the calculated date
            var fromDateInput = driver.FindElement(By.Id("fromBidDt"));
            fromDateInput.Clear(); // Clears any pre-filled values in the input field
            fromDateInput.SendKeys(fiveDaysAgoFormatted);

            // Find the end date input by its ID and enter today's date
            var toDateInput = driver.FindElement(By.Id("toBidDt"));
            toDateInput.Clear(); // Clears any pre-filled values in the input field
            toDateInput.SendKeys(todayFormatted);

            // After entering the dates, click the search button 
            var searchButton = driver.FindElement(By.XPath("//*[@id=\"searchForm\"]/div/fieldset[1]/ul/li[4]/dl/dd[3]/a/strong"));
            searchButton.Click();

            // Switch to main frame
            driver.SwitchTo().Frame("sub");
            driver.SwitchTo().Frame("main");

            // Find table element
            var table = driver.FindElement(By.XPath("//*[@id=\"resultForm\"]/div[2]/table"));

            // Find every table row
            var rows = table.FindElements(By.TagName("tr"));

            // Save table data in 'tableData'
            List<List<string>> tableData = new List<List<string>>();

            // flag to check first row
            bool isFirstRow = true; 

            foreach (var row in rows)
            {
                // Not to add current date and time in first row
                if (isFirstRow)
                {
                    isFirstRow = false;
                    continue; 
                }

                // Find every 'td' in every row
                var cells = row.FindElements(By.TagName("td"));
                List<string> rowData = new List<string>();

                // Use for loop start with 3rd index
                for (int i = 3; i < cells.Count-2; i++)
                {
                    // Extract cell data
                    rowData.Add(cells[i].Text);
                }
                
                // After extracting, add current date and time
                string currentDateTime = DateTime.Now.ToString("yyyy.MM.dd HH:mm");
                rowData.Add(currentDateTime);
                
                // If row data is not empty, then add to list 'tableData'
                if (rowData.Count > 0)
                {
                    tableData.Add(rowData);
                }
            }

        // License setting as NonCommercial
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        // Get folder path of desktop
        string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

        // Set file path
        string filePath = Path.Combine(desktopPath, "장표.xlsx");

        // Create ExcelPackage object
        FileInfo fileInfo = new FileInfo(filePath);

        // Work on Execl file
        using (ExcelPackage package = new ExcelPackage(fileInfo))
        {   
            // Use first worksheet
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; 

            // Find last row to add extracted data
            int lastRow = worksheet.Dimension.End.Row + 1; 

            // Add data
            foreach (var row in tableData)
            {
                for (int i = 0; i < row.Count; i++)
                {
                    worksheet.Cells[lastRow, i + 1].Value = row[i];
                }
                lastRow++;
            }
            // Save changes
            package.Save();
        }
        Console.WriteLine("Extracting is sucessfully worked!");
        }
    }
}
