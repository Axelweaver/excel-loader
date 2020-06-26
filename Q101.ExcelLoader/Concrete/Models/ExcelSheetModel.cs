using System.Collections.Generic;

namespace Q101.ExcelLoader.Concrete.Models
{
    /// <summary>
    /// Модель листа Excel
    /// </summary>
    public class ExcelSheetModel
    {
        /// <summary>
        /// Наименование листа
        /// </summary>
        public string SheetName { get; set; }

        /// <summary>
        /// Строки этого листа
        /// </summary>
        public IEnumerable<ExcelRowModel> Rows { get; set; }
    }
}
