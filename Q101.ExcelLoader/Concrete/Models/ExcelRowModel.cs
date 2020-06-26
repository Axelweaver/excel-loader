using System.Collections.Generic;

namespace Q101.ExcelLoader.Concrete.Models
{
    /// <summary>
    /// Модель строки Excel
    /// </summary>
    public class ExcelRowModel
    {
        /// <summary>
        /// Номер строки (начиная с 1)
        /// </summary>
        public int RowIndex { get; set; }

        /// <summary>
        /// Ячейки строки Excel
        /// </summary>
        public IEnumerable<ExcelCellModel> Cells { get; set; }
    }
}
