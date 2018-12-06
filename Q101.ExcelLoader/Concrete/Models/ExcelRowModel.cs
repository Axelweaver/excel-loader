using System.Collections.Generic;

namespace Q101.ExcelLoader.Concrete.Models
{
    /// <summary>
    /// Модель строки Excel
    /// </summary>
    public class ExcelRowModel
    {
        /// <summary>
        /// Ячейки строки Excel
        /// </summary>
        public IEnumerable<ExcelCellModel> Cells { get; set; }
    }
}
