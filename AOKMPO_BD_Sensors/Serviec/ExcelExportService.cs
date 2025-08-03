using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AOKMPO_BD_Sensors.Serviec
{
    public static class ExcelExportService
    {
        public static void ExportSensorsToExcel(IEnumerable<Sensor> sensors, string filePath)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Перечень средств измерений");

                // Настройка стилей
                var headerStyle = workbook.Style;
                headerStyle.Font.FontName = "Times New Roman";
                headerStyle.Font.FontSize = 12;
                headerStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                headerStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                // Шапка отчета
                var header = worksheet.Range("A1:E1");
                header.Merge().Value = "ПЕРЕЧЕНЬ";
                header.Style.Font.Bold = true;
                header.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                header.Style.Fill.BackgroundColor = XLColor.LightGray;

                // Подшапка
                worksheet.Cell(2, 1).Value = "Дата формирования:";
                worksheet.Cell(2, 2).Value = DateTime.Now.ToString("dd.MM.yyyy HH:mm");

                // Заголовки столбцов
                string[] columns = {"№", "Наименование СИ", "Тип СИ", "Индивидуальный номер", "Пределы измерения", "Кл. точн", "Дата проверки", "Хранится", "Эксплуатации" };
                for (int i = 0; i < columns.Length; i++)
                {
                    worksheet.Cell(8, i + 1).Value = columns[i];
                    worksheet.Cell(8, i + 1).Style.Font.Bold = true;
                    worksheet.Cell(8, i + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                }

                // Данные
                int row = 9;
                int counter = 1; // Счетчик для нумерации
                foreach (var sensor in sensors)
                {
                    // Нумерация (столбец A)
                    worksheet.Cell(row, 1).Value = counter++;
                    worksheet.Cell(row, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    // Остальные данные (сдвинуты на 1 колонку вправо)
                    worksheet.Cell(row, 2).Value = sensor.Name;
                    worksheet.Cell(row, 3).Value = sensor.TypeSensor;
                    worksheet.Cell(row, 4).Value = sensor.SerialNumber;
                    worksheet.Cell(row, 5).Value = sensor.MeasurementLimits;
                    worksheet.Cell(row, 6).Value = sensor.ClassForSure;
                    worksheet.Cell(row, 7).Value = sensor.ExpiryDate;
                    worksheet.Cell(row, 8).Value = sensor.Location;
                    worksheet.Cell(row, 9).Value = sensor.PlaceOfUse;


                    // Форматирование
                    if (sensor.ExpiryDate < DateTime.Today)
                    {
                        worksheet.Range(row, 1, row, 9).Style.Fill.BackgroundColor = XLColor.Red;
                    }
                    row++;
                }

                // Настройки таблицы
                worksheet.Columns().AdjustToContents();
                worksheet.RangeUsed().SetAutoFilter();
                worksheet.PageSetup.PrintAreas.Add("A1:E" + (row - 1));

                workbook.SaveAs(filePath);
            }
        }
    }
}
