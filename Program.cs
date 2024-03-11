using System;
using System.Data.SqlClient;
using System.IO;
using System.Net.Http;
using Microsoft.Office.Interop.Excel;

class Program
{
    static void Main()
    {
        string baseUrl = "https://clientportal.jse.co.za/downloadable-files?RequestNode=/YieldX/Derivatives/Docs_DMTM"; // Replace with your actual base URL
        string localFolder = "C:\\Users\\smsomi\\Desktop\\Assessment";

        using (HttpClient client = new HttpClient())
        {
            string[] fileNames = GetFileNamesFromUrl(baseUrl);

            foreach (string fileName in fileNames)
            {
                string filePath = Path.Combine(localFolder, fileName);

                if (!File.Exists(filePath))
                {
                    DownloadFile(client, baseUrl + fileName, filePath);
                    ProcessExcelFile(filePath);
                }
            }
        }
    }

    static string[] GetFileNamesFromUrl(string baseUrl)
    {
        
        throw new NotImplementedException();
    }

    static void DownloadFile(HttpClient client, string url, string destinationPath)
    {
        byte[] fileBytes = client.GetByteArrayAsync(url).Result;
        File.WriteAllBytes(destinationPath, fileBytes);
    }

    static void ProcessExcelFile(string filePath)
    {
        Application excelApp = new Application();
        Workbook workbook = excelApp.Workbooks.Open(filePath);

        try
        {
           
            string contractDetails = ((Range)workbook.Sheets[1].Cells[1, 1]).Text.ToString();

            SaveToDatabase(contractDetails, Path.GetFileName(filePath));
        }
        finally
        {
            workbook.Close(false);
            excelApp.Quit();
            ReleaseObject(workbook);
            ReleaseObject(excelApp);
        }
    }

    static void SaveToDatabase(string contractDetails, string fileName)
    {
        string connectionString = "Data Source=SMSOMI650G4\FLOWCENTRIC;User ID=smsomi;Password=***********;Initial Catalog=your_database;Integrated Security=True;";
        using (SqlConnection connection = new SqlConnection(connectionString))
        {
            connection.Open();

            string checkDuplicateQuery = "SELECT COUNT(*) FROM DailyMTM WHERE FileName = @FileName";
            using (SqlCommand checkCmd = new SqlCommand(checkDuplicateQuery, connection))
            {
                checkCmd.Parameters.AddWithValue("@FileName", fileName);
                int existingCount = (int)checkCmd.ExecuteScalar();

                if (existingCount == 0)
                {
                    string insertQuery = "INSERT INTO DailyMTM (ContractDetails, FileName, DownloadDate) VALUES (@ContractDetails, @FileName, @DownloadDate)";
                    using (SqlCommand cmd = new SqlCommand(insertQuery, connection))
                    {
                        cmd.Parameters.AddWithValue("@ContractDetails", contractDetails);
                        cmd.Parameters.AddWithValue("@FileName", fileName);
                        cmd.Parameters.AddWithValue("@DownloadDate", DateTime.Now);
                        cmd.ExecuteNonQuery();
                    }
                }
                else
                {
                    Console.WriteLine($"File '{fileName}' has already been processed.");
                }
            }
        }
    }

    static void ReleaseObject(object obj)
    {
        try
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error releasing object: " + ex.ToString());
        }
        finally
        {
            GC.Collect();
        }
    }
}


