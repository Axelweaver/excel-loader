using System;
using System.Globalization;

namespace Q101.ExcelLoader.Concrete.Models
{
    /// <summary>
    /// Модель ячейки Excel
    /// </summary>
    public class ExcelCellModel
    {
        /// <summary>
        /// Значение ячейки
        /// </summary>
        public object Value { get; set; }

        /// <summary>
        /// Получить значение ячейки как Int32.
        /// </summary>
        public int? GetInt()
        {
            if (int.TryParse(Value?.ToString(), out var val))
            {
                return val;
            }

            return null;
        }

        /// <summary>
        /// Получить значение ячейки как Int64.
        /// </summary>
        public long? GetLong()
        {
            if (long.TryParse(Value?.ToString(), out var val))
            {
                return val;
            }

            return null;
        }

        /// <summary>
        /// Получить значение ячейки как Double.
        /// </summary>
        public double? GetDouble()
        {
            if (double.TryParse(
                Value?.ToString(),
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
        public decimal? GetDecimal()
        {
            if (decimal.TryParse(
                Value?.ToString(),
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
        public string GetString()
        {
            return Value?.ToString();
        }

        /// <summary>
        /// Получить значение ячейки как DateTime.
        /// </summary>
        /// <param name="format">Формат даты (прим. "dd.MM.yyyy").</param>
        public DateTime? GetDateTime(string format)
        {
            string dateString = Value?.ToString();

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
        /// <param name="format">Формат даты (прим. "hh\\:mm\\:ss").</param>
        public TimeSpan? GetTimeSpan(string format)
        {
            string timeString = Value?.ToString();

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
    }
}
