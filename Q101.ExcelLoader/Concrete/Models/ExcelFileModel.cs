using System.Collections.Generic;

namespace Q101.ExcelLoader.Concrete.Models
{
    /// <summary>
    /// Модель Excel файла
    /// </summary>
    public class ExcelFileModel
    {
        /// <summary>
        /// Листы excel файла
        /// </summary>
        public IEnumerable<ExcelSheetModel> Sheets { get; set; }
    }
}
