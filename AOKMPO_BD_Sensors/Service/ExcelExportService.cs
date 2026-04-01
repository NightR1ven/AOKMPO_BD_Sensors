using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AOKMPO_BD_Sensors.Service
{
    public static class ExcelExportService
    {
        public static void ExportSensorsToExcel(IEnumerable<Sensor> sensors, string filePath)
        {
            // Убедимся, что папка для сохранения существует
            var directory = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                Directory.CreateDirectory(directory);

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Перечень средств измерений");

                // --- Заголовок документа ---
                var mainHeader1 = worksheet.Range("A1:K1");   // было A1:S1, теперь до K (11 колонок)
                mainHeader1.Merge().Value = "ПЕРЕЧЕНЬ";
                mainHeader1.Style.Font.Bold = true;
                mainHeader1.Style.Font.FontSize = 14;
                mainHeader1.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                var mainHeader2 = worksheet.Range("A2:K2");
                mainHeader2.Merge().Value = "средств измерений для включения в график поверки (калибровки) в ОГМетр";
                mainHeader2.Style.Font.Bold = true;
                mainHeader2.Style.Font.FontSize = 14;
                mainHeader2.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                var mainHeader3 = worksheet.Range("A3:K3");
                mainHeader3.Merge().Value = $"на {DateTime.Now.Year + 1} г.";
                mainHeader3.Style.Font.Bold = true;
                mainHeader3.Style.Font.FontSize = 14;
                mainHeader3.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                // --- Заголовки столбцов (теперь 11) ---
                string[] columnHeaders =
                {
                    "№",
                    "Наименование\nСИ",
                    "Тип\nСИ",
                    "Индивидуальный\nномер",
                    "Пределы\nизмерения",
                    "Кл.\nточн",
                    "Прошлая\nпроверка",
                    "Дата\nпроверки",
                    "Хранится",
                    "Эксплуатации",
                    "Документ"
                };

                // Строка 5 – основные заголовки, строка 6 – пустая (будет объединена с 5-й)
                for (int i = 0; i < columnHeaders.Length; i++)
                {
                    var cell = worksheet.Cell(5, i + 1);
                    cell.Value = columnHeaders[i];
                    cell.Style.Font.Bold = true;
                    cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    cell.Style.Alignment.WrapText = true;
                    cell.Style.Border.SetOutsideBorder(XLBorderStyleValues.Medium);
                }

                // Объединение ячеек для каждой колонки (строки 5-6)
                for (int i = 1; i <= 11; i++)
                {
                    worksheet.Range(5, i, 6, i).Merge();
                }

                // Нумерация в строке 7 (номера столбцов)
                for (int i = 1; i <= 11; i++)
                {
                    var cell = worksheet.Cell(7, i);
                    cell.Value = i;
                    cell.Style.Font.Bold = true;
                    cell.Style.Border.SetOutsideBorder(XLBorderStyleValues.Medium);
                    cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                }

                // --- Данные ---
                int row = 8; // начиная с 8 строки
                int counter = 1;

                // Если нет датчиков – всё равно оставим пустую таблицу
                var sensorList = sensors?.ToList() ?? new List<Sensor>();
                foreach (var sensor in sensorList)
                {
                    worksheet.Cell(row, 1).Value = counter++;
                    worksheet.Cell(row, 2).Value = sensor.Name ?? "";
                    worksheet.Cell(row, 3).Value = sensor.TypeSensor ?? "";
                    worksheet.Cell(row, 4).Value = sensor.SerialNumber ?? "";
                    worksheet.Cell(row, 5).Value = sensor.MeasurementLimits ?? "";
                    worksheet.Cell(row, 6).Value = sensor.ClassForSure ?? "";
                    worksheet.Cell(row, 7).Value = sensor.PlacementDate;
                    worksheet.Cell(row, 8).Value = sensor.ExpiryDate;
                    worksheet.Cell(row, 9).Value = sensor.Location ?? "";
                    worksheet.Cell(row, 10).Value = sensor.PlaceOfUse ?? "";
                    worksheet.Cell(row, 11).Value = sensor.PlaceOfDoc ?? "";

                    // Форматирование строки
                    for (int col = 1; col <= 11; col++)
                    {
                        var cell = worksheet.Cell(row, col);
                        cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        cell.Style.Font.FontSize = 9;
                        cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    }
                    row++;
                }

                // --- Настройка ширины столбцов ---
                worksheet.Column(1).Width = 5;   // №
                worksheet.Column(2).Width = 15;  // Наименование СИ
                worksheet.Column(3).Width = 15;  // Тип СИ
                worksheet.Column(4).Width = 17;  // Индивидуальный номер
                worksheet.Column(5).Width = 10;  // Пределы измерения
                worksheet.Column(6).Width = 8;   // Кл. точн
                worksheet.Column(7).Width = 15;  // Прошлая проверка
                worksheet.Column(8).Width = 15;  // Дата проверки
                worksheet.Column(9).Width = 12;  // Хранится
                worksheet.Column(10).Width = 15; // Эксплуатации
                worksheet.Column(11).Width = 12; // Документ

                // --- Выравнивание текста ---
                worksheet.Rows(8, row - 1).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Rows(8, row - 1).Style.Alignment.WrapText = true;

                // --- Границы всей таблицы ---
                if (row > 8) // если есть данные
                {
                    var tableRange = worksheet.Range(5, 1, row - 1, 11);
                    tableRange.Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
                }

                // --- Подписи под таблицей ---
                int footerRow = row + 1;
                // Сноски
                worksheet.Cell(footerRow + 2, 2).Value = "#   -Устройство САУ и АИиС НК-38СТ   инвентарный № 05Y-00700";
                worksheet.Cell(footerRow + 4, 2).Value = "##   -Стенд автономных испытания опыт. горелок камеры сгорания изд. НК-38СТ, НК-16СТ  инвентарный №СТ-00032";
                worksheet.Cell(footerRow + 6, 2).Value = "###   -Стенд№2 инвентарный №01Y-00404";

                // Подписи должностных лиц
                worksheet.Cell(footerRow + 8, 2).Value = "Начальник уч. 420";
                worksheet.Cell(footerRow + 8, 4).Value = "___________________________";
                worksheet.Cell(footerRow + 8, 7).Value = "___________________________";

                worksheet.Cell(footerRow + 10, 2).Value = "Врио Начальника БАиИ у.420";
                worksheet.Cell(footerRow + 10, 4).Value = "___________________________";
                worksheet.Cell(footerRow + 10, 7).Value = "___________________________";

                worksheet.Cell(footerRow + 12, 2).Value = "Механик уч.420";
                worksheet.Cell(footerRow + 12, 4).Value = "___________________________";
                worksheet.Cell(footerRow + 12, 7).Value = "___________________________";

                worksheet.Cell(footerRow + 14, 2).Value = "Ответственный за СИ по у. 420";
                worksheet.Cell(footerRow + 14, 4).Value = "___________________________";
                worksheet.Cell(footerRow + 14, 7).Value = "___________________________";

                // Объединение ячеек для сносок
                worksheet.Range(footerRow + 2, 2, footerRow + 2, 22).Merge();
                worksheet.Range(footerRow + 4, 2, footerRow + 4, 22).Merge();
                worksheet.Range(footerRow + 6, 2, footerRow + 6, 22).Merge();

                worksheet.Range(footerRow + 8, 2, footerRow + 8, 3).Merge();
                worksheet.Range(footerRow + 10, 2, footerRow + 10, 3).Merge();
                worksheet.Range(footerRow + 12, 2, footerRow + 12, 3).Merge();
                worksheet.Range(footerRow + 14, 2, footerRow + 14, 3).Merge();

                worksheet.Range(footerRow + 8, 4, footerRow + 8, 6).Merge();
                worksheet.Range(footerRow + 10, 4, footerRow + 10, 6).Merge();
                worksheet.Range(footerRow + 12, 4, footerRow + 12, 6).Merge();
                worksheet.Range(footerRow + 14, 4, footerRow + 14, 6).Merge();

                worksheet.Range(footerRow + 8, 7, footerRow + 8, 9).Merge();
                worksheet.Range(footerRow + 10, 7, footerRow + 10, 9).Merge();
                worksheet.Range(footerRow + 12, 7, footerRow + 12, 9).Merge();
                worksheet.Range(footerRow + 14, 7, footerRow + 14, 9).Merge();

                // Жирный шрифт для подписей
                worksheet.Range(footerRow + 2, 2, footerRow + 14, 9).Style.Font.Bold = true;

                // --- Сохранение ---
                workbook.SaveAs(filePath);
            }
        }
    }
}

