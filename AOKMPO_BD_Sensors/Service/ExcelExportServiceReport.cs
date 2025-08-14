using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AOKMPO_BD_Sensors.Service
{
    public static class ExcelExportServiceReport
    {
        public static void ExportReportToExcel(IEnumerable<Sensor> sensors, string filePath)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Перечень средств измерений");

                // Настройка стилей
                var headerStyle = workbook.Style;
                headerStyle.Font.FontName = "Times New Roman";
                headerStyle.Font.FontSize = 6;
                headerStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                headerStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                var mainHeader1 = worksheet.Range("A1:T1");
                mainHeader1.Merge().Value = "ПЕРЕЧЕНЬ";
                mainHeader1.Style.Font.Bold = true;
                mainHeader1.Style.Font.FontSize = 14;
                mainHeader1.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                var mainHeader2 = worksheet.Range("A2:T2");
                mainHeader2.Merge().Value = "средств измерений для включения в график поверки (калибровки) в ОГМетр";
                mainHeader2.Style.Font.Bold = true;
                mainHeader2.Style.Font.FontSize = 14;
                mainHeader2.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                var mainHeader3 = worksheet.Range("A3:T3");
                mainHeader3.Merge().Value = $"на {DateTime.Now.Year + 1} г.";
                mainHeader3.Style.Font.Bold = true;
                mainHeader3.Style.Font.FontSize = 14;
                mainHeader3.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                // Заголовки столбцов основной таблицы
                var columnHeaders = new[]
                {
                    "№",
                    "Наименование\nСИ",
                    "Тип\nСИ",
                    "Индивидуальный\nномер",
                    "Пределы\nизмерения",
                    "Кл.\nточн",
                    "Хранится",
                    "Эксплуатации"
                };

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

                // Выделение ячеек заголовков
                worksheet.Range("A6:A6").Style.Border.SetOutsideBorder(XLBorderStyleValues.Medium);
                worksheet.Range("B6:B6").Style.Border.SetOutsideBorder(XLBorderStyleValues.Medium);
                worksheet.Range("C6:C6").Style.Border.SetOutsideBorder(XLBorderStyleValues.Medium);
                worksheet.Range("D6:D6").Style.Border.SetOutsideBorder(XLBorderStyleValues.Medium);
                worksheet.Range("E6:E6").Style.Border.SetOutsideBorder(XLBorderStyleValues.Medium);
                worksheet.Range("F6:F6").Style.Border.SetOutsideBorder(XLBorderStyleValues.Medium);
                worksheet.Range("G6:G6").Style.Border.SetOutsideBorder(XLBorderStyleValues.Medium);
                worksheet.Range("H6:H6").Style.Border.SetOutsideBorder(XLBorderStyleValues.Medium);

                // Объединение ячеек заголовков
                worksheet.Range("A5:A6").Merge();
                worksheet.Range("B5:B6").Merge();
                worksheet.Range("C5:C6").Merge();
                worksheet.Range("D5:D6").Merge();
                worksheet.Range("E5:E6").Merge();
                worksheet.Range("F5:F6").Merge();
                worksheet.Range("G5:G6").Merge();
                worksheet.Range("H5:H6").Merge();

                // Нумерация столбцов
                for (int i = 1; i <= 8; i++)
                {
                    worksheet.Cell(7, i).Value = i;
                    worksheet.Cell(7, i).Style.Font.Bold = true;
                    worksheet.Cell(7, i).Style.Border.SetOutsideBorder(XLBorderStyleValues.Medium);
                    worksheet.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                }

                // Заголовок для таблицы месяцев
                var monthHeader = worksheet.Range(5, 9, 5, 20);
                monthHeader.Merge().Value = "Месяц поверки";
                monthHeader.Style.Font.Bold = true;
                monthHeader.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                monthHeader.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                monthHeader.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                // Сокращения месяцев
                string[] monthAbbreviations = { "я", "ф", "м", "а", "м", "и", "и", "а", "с", "о", "н", "д" };
                for (int i = 0; i < monthAbbreviations.Length; i++)
                {
                    var cell = worksheet.Cell(6, 9 + i);
                    cell.Value = monthAbbreviations[i];
                    cell.Style.Font.Bold = true;
                    cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                }

                // Номера месяцев
                for (int i = 0; i < 12; i++)
                {
                    var cell = worksheet.Cell(7, 9 + i);
                    cell.Value = i + 1;
                    cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                }

                // Данные
                int row = 8;
                int counter = 1;
                foreach (var sensor in sensors)
                {
                    // Основные данные
                    worksheet.Cell(row, 1).Value = counter++;
                    worksheet.Cell(row, 2).Value = sensor.Name;
                    worksheet.Cell(row, 3).Value = sensor.TypeSensor;
                    worksheet.Cell(row, 4).Value = sensor.SerialNumber;
                    worksheet.Cell(row, 5).Value = sensor.MeasurementLimits;
                    worksheet.Cell(row, 6).Value = sensor.ClassForSure;
                    worksheet.Cell(row, 7).Value = sensor.Location;
                    worksheet.Cell(row, 8).Value = sensor.PlaceOfUse;

                    // Форматирование основных данных
                    for (int col = 1; col <= 8; col++)
                    {
                        worksheet.Cell(row, col).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        worksheet.Cell(row, col).Style.Font.FontSize = 9;
                        worksheet.Cell(row, col).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    }

                    // Рассчитываем месяцы поверки
                    DateTime startDate = sensor.PlacementDate; // прошлая поверка
                    DateTime endDate = sensor.ExpiryDate;      // следующая поверка

                    // Если дата следующей поверки раньше прошлой (например, при переходе через год)
                    if (endDate < startDate)
                    {
                        endDate = endDate.AddYears(1);
                    }

                    // Вычисляем разницу в месяцах
                    int monthsDifference = ((endDate.Year - startDate.Year) * 12) + endDate.Month - startDate.Month;

                    // Если интервал между поверками 6 месяцев или меньше
                    if (monthsDifference <= 6)
                    {
                        // Отмечаем обе даты
                        MarkVerificationMonth(worksheet, row, startDate.Month);
                        MarkVerificationMonth(worksheet, row, endDate.Month);
                    }
                    else
                    {
                        // Если интервал больше 6 месяцев, отмечаем только следующую поверку
                        MarkVerificationMonth(worksheet, row, endDate.Month);
                    }

                    // Рамка для всех ячеек строки месяцев
                    worksheet.Range(row, 9, row, 20).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    worksheet.Range(row, 9, row, 20).Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                    row++;
                }

                // Настройка ширины столбцов
                worksheet.Column(1).Width = 5;
                worksheet.Column(2).Width = 15;
                worksheet.Column(3).Width = 15;
                worksheet.Column(4).Width = 17;
                worksheet.Column(5).Width = 10;
                worksheet.Column(6).Width = 8;
                worksheet.Column(7).Width = 15;
                worksheet.Column(8).Width = 15;

                // Ширина столбцов для месяцев
                for (int col = 9; col <= 20; col++)
                {
                    worksheet.Column(col).Width = 4;
                }

                // Выравнивание текста
                worksheet.Rows(8, row - 1).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Rows(8, row - 1).Style.Alignment.WrapText = true;

                // Границы для всей таблицы
                worksheet.Range(5, 1, row - 1, 20).Style.Border.OutsideBorder = XLBorderStyleValues.Medium;

                // Границы для столбцов таблицы
                worksheet.Range(5, 1, row - 1, 8).Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
                worksheet.Range(5, 1, row - 1, 7).Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
                worksheet.Range(5, 1, row - 1, 6).Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
                worksheet.Range(5, 1, row - 1, 5).Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
                worksheet.Range(5, 1, row - 1, 4).Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
                worksheet.Range(5, 1, row - 1, 3).Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
                worksheet.Range(5, 1, row - 1, 2).Style.Border.OutsideBorder = XLBorderStyleValues.Medium;

                // Добавляем подпись под таблицей
                int footerRow = row + 1;
                worksheet.Cell(footerRow + 2, 2).Value = "#   -Устройство САУ и АИиС НК-38СТ   инвентарный № 05Y-00700";
                worksheet.Cell(footerRow + 4, 2).Value = "##   -Стенд автономных испытания опыт. горелок камеры сгорания изд. НК-38СТ, НК-16СТ  инвентарный №СТ-00032";
                worksheet.Cell(footerRow + 6, 2).Value = "###   -Стенд№2 инвентарный №01Y-00404";

                // Подписи
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

                // Объединение ячеек для подписи
                worksheet.Range(footerRow + 2, 2, footerRow + 2, 20).Merge();
                worksheet.Range(footerRow + 4, 2, footerRow + 4, 20).Merge();
                worksheet.Range(footerRow + 6, 2, footerRow + 6, 20).Merge();

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

                workbook.SaveAs(filePath);
            }
        }

        private static void MarkVerificationMonth(IXLWorksheet worksheet, int row, int month)
        {
            if (month >= 1 && month <= 12)
            {
                var cell = worksheet.Cell(row, 9 + (month - 1));
                cell.Style.Font.Bold = true;
                cell.Style.Border.DiagonalDown = true;
                cell.Style.Border.DiagonalBorder = XLBorderStyleValues.Medium;
            }
        }
    }
}

