using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.IO;
using Microsoft.Data.Sqlite;
using SQLitePCL;

namespace Excel_Reader1
{
    
    internal class Program
    {

        static void Main(string[] args)
        {

            SQLitePCL.Batteries_V2.Init();

            Console.OutputEncoding = System.Text.Encoding.UTF8;

            string filePath = @"C:\Users\user1\Documents\EXCEL READER\Excel Reader1\TableForReading.xlsx";
            string dbPath = @"C:\Users\user1\Documents\EXCEL READER\Excel Reader1\database\dbFromExcel.file";
            string connectionString = $"Data Source={dbPath}";

            if (File.Exists(dbPath))
            {
                File.Delete(dbPath);
            }

            Console.WriteLine("Deleting old database...");
            Console.WriteLine("Reading from Excel...");
            Console.WriteLine("Creating new database...");
            Console.WriteLine("Inserting data into database...");
            Console.WriteLine("Displaying data:");

            if (File.Exists(filePath))
            {
                var fileInfo = new FileInfo(filePath);
                using (var package = new ExcelPackage(fileInfo))
                {
                    var workSheet = package.Workbook.Worksheets[0];

                    int rowCount = workSheet.Dimension.Rows;
                    int columnCount = workSheet.Dimension.Columns;

                    using (var connection = new SqliteConnection(connectionString))
                    {
                        connection.Open();

                        var createTableCmd = connection.CreateCommand();
                        createTableCmd.CommandText = @"CREATE TABLE IF NOT EXISTS YourTable 
                        (
                        PurchaseID INTEGER PRIMARY KEY AUTOINCREMENT,
                        FirstName TEXT,
                        LastName TEXT,
                        Birthday DATE,
                        City TEXT,
                        DateOfPurchase DATE,
                        MoneySpent REAL
                        )";
                        createTableCmd.ExecuteNonQuery();

                        for (int row = 2; row <= rowCount; row++)
                        {
                            var column1 = workSheet.Cells[row, 1].Text;
                            var column2 = workSheet.Cells[row, 2].Text;
                            var column3 = workSheet.Cells[row, 3].Text;
                            var column4 = workSheet.Cells[row, 4].Text;
                            var column5 = workSheet.Cells[row, 5].Text;
                            var column6 = workSheet.Cells[row, 6].Text;
                            var column7 = workSheet.Cells[row, 7].Text;

                            var insertCmd = connection.CreateCommand();
                            insertCmd.CommandText = @"
                                    INSERT INTO YourTable (FirstName, LastName, Birthday, City, DateOfPurchase, MoneySpent)
                                    VALUES ($firstName, $lastName, $birthday, $city, $dateOfPurchase, $moneySpent)";

                            insertCmd.Parameters.AddWithValue("$firstName", column2);
                            insertCmd.Parameters.AddWithValue("$lastName", column3);
                            insertCmd.Parameters.AddWithValue("$birthday", column4);
                            insertCmd.Parameters.AddWithValue("$city", column5);
                            insertCmd.Parameters.AddWithValue("$dateOfPurchase", column6);
                            insertCmd.Parameters.AddWithValue("$moneySpent", column7);

                            insertCmd.ExecuteNonQuery();

                        }
                    }

                    Console.WriteLine("");

                    for (int row = 2; row <= rowCount; row++)
                    {
                        StringBuilder userData = new StringBuilder();

                        for (int col = 1; col <= columnCount; col++)
                        {
                            userData.Append(workSheet.Cells[row, col].Text + " ");
                        }
                        Console.WriteLine(userData.ToString());
                    }
                }
                Console.WriteLine("The data was imported succesfully");
            }
            else
            {
                Console.WriteLine("The file could not be found.");
            }
        }
    }
}
