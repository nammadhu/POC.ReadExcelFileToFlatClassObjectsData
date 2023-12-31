﻿using Aspose.Cells;
using Aspose.Cells.Utility;
//using DocumentFormat.OpenXml.Spreadsheet;

//using DocumentFormat.OpenXml.Spreadsheet;
using Newtonsoft.Json;
using System.Data;
using System.Text;

namespace ConsoleAppExcelFromApi
{
    /*
Path.GetDirectoryName(Process.GetCurrentProcess()?.MainModule?.FileName) = /usr/share/dotnet
Path.GetDirectoryName(Environment.ProcessPath)                           = /usr/share/dotnet
Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)          = /home/username/myproject/bin/Debug/net6.0
typeof(SomeType).Assembly.Location                                       = /home/username/myproject/bin/Debug/net6.0
Path.GetDirectoryName(Environment.GetCommandLineArgs()[0])               = /home/username/myproject/bin/Debug/net6.0
AppDomain.CurrentDomain.BaseDirectory                                    = /home/username/myproject/bin/Debug/net6.0/
System.AppContext.BaseDirectory                                          = /home/username/myproject/bin/Debug/net6.0/
     */
    internal class Program
    {
        static void Main(string[] args)
        {

            Console.WriteLine("Hello, World!");

            var data = ReadExcelData<FinalItems>(filePath: @$"D:\Unwell\MedicinesFilterd03Dec2023.xlsx");
            int c = 0;
            foreach (var item in data)
            {
                if (c > 20) break;
                c++;
                Console.WriteLine($"Name: {item.name}, : {item.id}, Email: {item.price}");
            }
        }
        static List<T> ReadExcelData<T>(string filePath, string sheetName = "Sheet1") where T : new()
        {
            var data = new List<T>();

            // Load the Excel file
            var workbook = new Aspose.Cells.Workbook(filePath);
            //workbook.Save("Outputs.json");// to save excel to json you can use this
            // Access the desired worksheet
            var worksheet = workbook.Worksheets[sheetName];

            // Create a mapping of column names to column indices
            var columnIndexMap = new Dictionary<string, int>();
            for (int columnIndex = 0; columnIndex <= worksheet.Cells.MaxDataColumn; columnIndex++)
            {
                var cellValue = worksheet.Cells[0, columnIndex].StringValue;
                columnIndexMap[cellValue] = columnIndex;
            }

            // Iterate over the rows (starting from the second row)
            for (int row = 1; row <= worksheet.Cells.MaxDataRow; row++)
            {
                // Create an instance of the class
                var item = new T();

                // Iterate over the properties of the class
                foreach (var property in typeof(T).GetProperties())
                {
                    // Check if the property name is present in the Excel file
                    if (columnIndexMap.TryGetValue(property.Name, out var columnIndex))
                    {
                        // Get the cell value from the current row and column
                        var cellValue = worksheet.Cells[row, columnIndex].StringValue;

                        // Convert the cell value to the property type and set the value
                        var convertedValue = Convert.ChangeType(cellValue, property.PropertyType);
                        property.SetValue(item, convertedValue);
                    }
                }

                // Add the object to the list
                data.Add(item);
            }

            return data;
        }
    }
}