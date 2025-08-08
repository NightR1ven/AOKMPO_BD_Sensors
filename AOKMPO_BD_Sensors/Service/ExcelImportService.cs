using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AOKMPO_BD_Sensors.Service
{
    public static class ExcelImportService
    {
        public static List<Sensor> ImportSensorFromExcel(string filePath)
        {
            var sensors = new List<Sensor>();

            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(1);

                // Проверка шапки
                if (worksheet.Cell(5, 1).GetString() != "№")
                    
                {
                    throw new Exception("Неверный формат файла! Используйте шаблон экспорта.");
                }

                // Чтение данных (начиная строки, пропускаем №)
                var rows = worksheet.RowsUsed().Skip(4);

                foreach (var row in rows)
                {
                    // Пропускаем пустые строки или строки без данных (например, подписи в конце)
                    if (row.Cell(2).IsEmpty() ||
                        row.Cell(7).IsEmpty() ||
                        row.Cell(8).IsEmpty())
                        continue;

                    try
                    {
                        // Пытаемся распарсить даты, если не получается — пропускаем строку
                        if (!DateTime.TryParse(row.Cell(7).GetString(), out var placementDate) ||
                            !DateTime.TryParse(row.Cell(8).GetString(), out var expiryDate))
                        {
                            continue;
                        }

                        sensors.Add(new Sensor
                        {
                            Name = row.Cell(2).GetString(),
                            TypeSensor = row.Cell(3).GetString(),
                            SerialNumber = row.Cell(4).GetString(),
                            MeasurementLimits = row.Cell(5).GetString(),
                            ClassForSure = row.Cell(6).GetString(),
                            PlacementDate = DateTime.Parse(row.Cell(7).GetString()),
                            ExpiryDate = DateTime.Parse(row.Cell(8).GetString()),
                            Location = row.Cell(9).GetString(),
                            PlaceOfUse = row.Cell(10).GetString()
                        });
                    }
                    catch (Exception ex)
                    {
                        throw new Exception($"Ошибка в строке {row.RowNumber()}: {ex.Message}");
                    }
                }
            }

            return sensors;
        }
    }
}