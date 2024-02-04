using System;
using System.Data.SqlClient;
using System.IO;
using OfficeOpenXml;
using System.Configuration;

class Program
{
    static void Main(string[] args)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        string excelFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "PackageUpload.xlsx");
        string destinationConnectionString = ConfigurationManager.ConnectionStrings["DsdConnectionString"].ConnectionString;

        try
        {
            // Read Excel file
            FileInfo excelFile = new FileInfo(excelFilePath);

            using (ExcelPackage package = new ExcelPackage(excelFile))
            {
                if (package.Workbook.Worksheets.Count > 0)
                {
                    // Assuming the first worksheet contains the data
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[1];

                    // Connect to SQL Server
                    using (SqlConnection connection = new SqlConnection(destinationConnectionString))
                    {
                        connection.Open();

                        // Iterate through rows and insert into SQL Server table
                        for (int row = 2; row <= worksheet.Dimension.Rows; row++) // Assuming the first row is headers
                        {
                            string packageValue = worksheet.Cells[row, 1].Value?.ToString();
                            string duration = worksheet.Cells[row, 2].Value?.ToString();
                            string deno = worksheet.Cells[row, 3].Value?.ToString();
                            string bonus = worksheet.Cells[row, 4].Value?.ToString();
                            // Add more columns as needed

                            // Insert data into SQL Server table
                            string sqlInsert = $"INSERT INTO Package (Package, DurationMonth,Deno, Bonus) VALUES ('{packageValue}', '{duration}','{deno}','{bonus}')";
                            using (SqlCommand command = new SqlCommand(sqlInsert, connection))
                            {
                                command.ExecuteNonQuery();
                            }
                        }
                    }
                }
            }

            Console.WriteLine("Data uploaded successfully!");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
