namespace Q101.ExcelLoader.Concrete.Models
{
    /// <summary>
    /// Модель ячейки Excel
    /// </summary>
    public class ExcelCellModel
    {
        /// <summary>
        /// Индекс колонки (начиная с 1)
        /// </summary>
        public int ColumnIndex { get; set; }

        /// <summary>
        /// Адрес ячейки
        /// </summary>
        public string Address { get; set; }

        /// <summary>
        /// Полный адрес ячейки
        /// </summary>
        public string FullAddress { get; set; }

        /// <summary>
        /// Значение ячейки
        /// </summary>
        public object Value { get; set; }
    }
}
