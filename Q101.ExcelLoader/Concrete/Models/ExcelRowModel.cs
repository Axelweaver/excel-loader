using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

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

        /// <summary>
        /// Получить значение ячейки как Int32.
        /// </summary>
        /// <param name="cellIndex">Индекс ячейки.</param>
        public int? GetInt(int cellIndex)
        {
            var cellVal = GetCellValue(cellIndex);

            if (int.TryParse(cellVal?.ToString(), out var val))
            {
                return val;
            }

            return null;
        }

        /// <summary>
        /// Получить значение ячейки как Int64.
        /// </summary>
        /// <param name="cellIndex">Индекс ячейки.</param>
        public long? GetLong(int cellIndex)
        {
            var cellVal = GetCellValue(cellIndex);

            if (long.TryParse(cellVal?.ToString(), out var val))
            {
                return val;
            }

            return null;
        }

        /// <summary>
        /// Получить значение ячейки как Double.
        /// </summary>
        /// <param name="cellIndex">Индекс ячейки.</param>
        public double? GetDouble(int cellIndex)
        {
            var cellVal = GetCellValue(cellIndex);

            if (double.TryParse(
                cellVal?.ToString(),
                NumberStyles.Any,
                CultureInfo.InvariantCulture,
                out var val))
            {
                return val;
            }

            return null;
        }

        /// <summary>
        /// Получить значение ячейки как Decimal.
        /// </summary>
        /// <param name="cellIndex">Индекс ячейки.</param>
        public decimal? GetDecimal(int cellIndex)
        {
            var cellVal = GetCellValue(cellIndex);

            if (decimal.TryParse(
                cellVal?.ToString(),
                NumberStyles.Any,
                CultureInfo.InvariantCulture,
                out var val))
            {
                return val;
            }

            return null;
        }

        /// <summary>
        /// Получить значение ячейки как String.
        /// </summary>
        /// <param name="cellIndex">Индекс ячейки.</param>
        public string GetString(int cellIndex)
        {
            var cellVal = GetCellValue(cellIndex);

            return cellVal?.ToString();
        }

        /// <summary>
        /// Получить значение ячейки как DateTime.
        /// </summary>
        /// <param name="cellIndex">Индекс ячейки.</param>
        /// <param name="format">Формат даты (прим. "dd.MM.yyyy").</param>
        public DateTime? GetDateTime(int cellIndex, string format)
        {
            var cellVal = GetCellValue(cellIndex);
            string dateString = cellVal?.ToString();

            if (string.IsNullOrWhiteSpace(dateString))
            {
                return null;
            }

            if (DateTime.TryParseExact(
                dateString,
                format,
                new DateTimeFormatInfo(),
                DateTimeStyles.None,
                out var val))
            {
                return val;
            }

            return null;
        }

        /// <summary>
        /// Получить значение ячейки как TimeSpan.
        /// </summary>
        /// <param name="cellIndex">Индекс ячейки.</param>
        /// <param name="format">Формат даты (прим. "hh\\:mm\\:ss").</param>
        public TimeSpan? GetTimeSpan(int cellIndex, string format)
        {
            var cellVal = GetCellValue(cellIndex);
            string timeString = cellVal?.ToString();

            if (string.IsNullOrWhiteSpace(timeString))
            {
                return null;
            }

            if (TimeSpan.TryParseExact(
                timeString,
                format,
                CultureInfo.InvariantCulture,
                out var val))
            {
                return val;
            }

            return null;
        }

        /// <summary>
        /// Получить значение ячейки по индексу.
        /// </summary>
        /// <param name="cellIndex">Индекс ячейки.</param>
        private object GetCellValue(int cellIndex)
        {
            var cell = Cells.ElementAtOrDefault(cellIndex);
            var cellVal = cell?.Value;

            return cellVal;
        }
    }
}
