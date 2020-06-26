using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using Q101.ExcelLoader.Abstract;
using Q101.ExcelLoader.Concrete.Models;

namespace Q101.ExcelLoader.Concrete
{
    /// <summary>
    /// Загрузчик Excel файла
    /// </summary>
    public class ExcelFileLoader : IExcelFileLoader
    {
        /// <summary>
        /// Загрузить файл
        /// </summary>
        /// <param name="fileStream">Файловый поток</param>
        /// <returns></returns>
        public ExcelFileModel Load(Stream fileStream)
        {
            var model = new ExcelFileModel();

            using (var package = new ExcelPackage(fileStream))
            {
                model.Sheets = package.Workbook.Worksheets.Select(ConvertSheet).ToList();
            }

            return model;
        }

        /// <summary>
        /// Конвертировать лист (worksheet)
        /// </summary>
        /// <param name="sheet">Лист</param>
        /// <returns></returns>
        public ExcelSheetModel ConvertSheet(ExcelWorksheet sheet)
        {
            var rows = new List<ExcelRowModel>(sheet.Cells.Rows);
 
            var cells = sheet.Cells;

            cells.Reset();

            int previosRowIndex = 1;

            var currentCells = new List<ExcelCellModel>();

            var currentRow = new ExcelRowModel
            {
                RowIndex = 1,
                Cells = currentCells
            };
            rows.Add(currentRow);

            while (cells.MoveNext())
            {
                if (cells.Current.Start.Row > previosRowIndex)
                {
                    currentCells = new List<ExcelCellModel>();
                    
                    currentRow = new ExcelRowModel
                    {
                        RowIndex = cells.Current.Start.Row,
                        Cells = currentCells
                    };
                    rows.Add(currentRow);
                }
                
                var currentCell = new ExcelCellModel
                {
                    ColumnIndex = cells.Current.Start.Column,
                    Value = cells.Current.Value,
                    Address = cells.Current.Address,
                    FullAddress = cells.Current.FullAddress
                };

                currentCells.Add(currentCell);

                previosRowIndex = cells.Current.Start.Row;
            }

            var sheetModel = new ExcelSheetModel
            {
                SheetName = sheet.Name,
                Rows = rows
            };

            return sheetModel;
        }
    }
}
