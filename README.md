# Q101.ExcelLoader

[![NuGet Package](https://img.shields.io/nuget/v/Q101.ExcelLoader.svg?style=for-the-badge&logo=appveyor)](https://www.nuget.org/packages/Q101.ExcelLoader)
[![NuGet Package](https://img.shields.io/nuget/dt/Q101.ExcelLoader.svg?style=for-the-badge&logo=appveyor)](https://www.nuget.org/packages/Q101.ExcelLoader)


### Nuget package link:
[https://www.nuget.org/packages/Q101.ExcelLoader/](https://www.nuget.org/packages/Q101.ExcelLoader/ "Nuget package link")


This assembly (class library) allows you to download the MS Excel file (.xlsx) as an object and work with data from the file as objects. Used famous library EPPlus

 To install this assembly (class library) on the package manager console tab, run the following command:
```bash

    Install-Package Q101.ExcelLoader -Version 1.0.1

```

Example

```
        static void Main(string[] args)
        {
            Console.WriteLine("For start press enter");

            Console.ReadLine();

            var filePath = "d:\\file.xlsx";

            using (var fileStream = File.OpenRead(filePath))
            {
                var excelLoader = new ExcelFileLoader();

                var excelFileModel = excelLoader.Load(fileStream);

                Console.WriteLine("Read");

                foreach (var excelSheetModel in excelFileModel.Sheets)
                {
                    Console.WriteLine($"Sheet name \"{excelSheetModel.SheetName}\"");
                    foreach (var excelRowModel in excelSheetModel.Rows)
                    {
                        Console.Write($"{excelRowModel.RowIndex}. ");
                        foreach (var excelCellModel in excelRowModel.Cells)
                        {
                            Console.Write($"| [{excelCellModel.ColumnIndex}]({excelCellModel.Address}) {excelCellModel.Value} ");
                        }
                        Console.Write("\n\n");
                    }
                }
            }

            Console.ReadLine();
        }
```

When loading an excel file, empty cells are skipped, be careful! Therefore, a cell has indexes on the cell and row.