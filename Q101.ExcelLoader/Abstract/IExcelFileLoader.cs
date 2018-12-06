using System.IO;
using OfficeOpenXml;
using Q101.ExcelLoader.Concrete.Models;

namespace Q101.ExcelLoader.Abstract
{
    /// <summary>
    /// Загрузчик Excel файла
    /// </summary>
    public interface IExcelFileLoader
    {
        ExcelFileModel Load(Stream fileStream);

        ExcelSheetModel ConvertSheet(ExcelWorksheet sheet);
    }
}
