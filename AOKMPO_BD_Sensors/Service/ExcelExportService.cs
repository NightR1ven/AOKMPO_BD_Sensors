using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AOKMPO_BD_Sensors.Service
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
                headerStyle.Font.FontSize = 6;
                headerStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                headerStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                var mainHeader1 = worksheet.Range("A1:J1");
                mainHeader1.Merge().Value = "ПЕРЕЧЕНЬ";
                mainHeader1.Style.Font.Bold = true;
                mainHeader1.Style.Font.FontSize = 14;
                mainHeader1.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                var mainHeader2 = worksheet.Range("A2:J2");
                mainHeader2.Merge().Value = "средств измерений для включения в график поверки (калибровки) в ОГМетр";
                mainHeader2.Style.Font.Bold = true;
                mainHeader2.Style.Font.FontSize = 14;
                mainHeader2.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                var mainHeader3 = worksheet.Range("A3:J3");
                mainHeader3.Merge().Value = $"на {DateTime.Now.Year+1} г.";
                mainHeader3.Style.Font.Bold = true;
                mainHeader3.Style.Font.FontSize = 14;
                mainHeader3.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;


                // Заголовки столбцов

                var columnHeaders = new[]
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
                    cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                }

                // Данные
                int row = 6;
                int counter = 1; // Счетчик для нумерации
                foreach (var sensor in sensors)
                {
                    
                    // Нумерация (столбец A)
                    worksheet.Cell(row, 1).Value = counter++;
                    worksheet.Cell(row, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet.Cell(row, 1).Style.Font.FontSize = 9;

                    // Остальные данные (сдвинуты на 1 колонку вправо)
                    worksheet.Cell(row, 2).Value = sensor.Name;
                    worksheet.Cell(row, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet.Cell(row, 2).Style.Font.FontSize = 9;

                    worksheet.Cell(row, 3).Value = sensor.TypeSensor;
                    worksheet.Cell(row, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet.Cell(row, 3).Style.Font.FontSize = 9;

                    worksheet.Cell(row, 4).Value = sensor.SerialNumber;
                    worksheet.Cell(row, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet.Cell(row, 4).Style.Font.FontSize = 9;

                    worksheet.Cell(row, 5).Value = sensor.MeasurementLimits;
                    worksheet.Cell(row, 5).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet.Cell(row, 5).Style.Font.FontSize = 9;

                    worksheet.Cell(row, 6).Value = sensor.ClassForSure;
                    worksheet.Cell(row, 6).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet.Cell(row, 6).Style.Font.FontSize = 9;

                    worksheet.Cell(row, 7).Value = sensor.PlacementDate;
                    worksheet.Cell(row, 8).Value = sensor.ExpiryDate;
                    worksheet.Cell(row, 7).Style.Font.FontSize = 9;
                    worksheet.Cell(row, 8).Style.Font.FontSize = 9;

                    worksheet.Cell(row, 9).Value = sensor.Location;
                    worksheet.Cell(row, 9).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet.Cell(row, 9).Style.Font.FontSize = 9;

                    worksheet.Cell(row, 10).Value = sensor.PlaceOfUse;
                    worksheet.Cell(row, 10).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet.Cell(row, 10).Style.Font.FontSize = 9;


                    // Форматирование
                    for (int col = 1; col <= 10; col++)
                    {
                        worksheet.Cell(row, col).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    }

                    row++;
                }

                // Настройка ширины столбцов
                worksheet.Column(1).Width = 5;   // № п/п
                worksheet.Column(2).Width = 15;  // Наименование СИ
                worksheet.Column(3).Width = 15;  // Тип СИ
                worksheet.Column(4).Width = 17;  // Номер
                worksheet.Column(5).Width = 10;  // Пределы измерений
                worksheet.Column(6).Width = 8;  // Класс точности
                worksheet.Column(7).Width = 15;  // Последняя поверка
                worksheet.Column(8).Width = 15;  // Следующая поверка
                worksheet.Column(9).Width = 12; // Место установки
                worksheet.Column(10).Width = 15; // Место эксплуатации

                // Выравнивание текста в ячейках
                worksheet.Rows(6, row - 1).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Rows(6, row - 1).Style.Alignment.WrapText = true;

                // Границы для всей таблицы
                worksheet.Range(5, 1, row - 1, 10).Style.Border.OutsideBorder = XLBorderStyleValues.Medium;

                // Добавляем подпись под таблицей
                int footerRow = row + 1;
                worksheet.Cell(footerRow + 2, 2).Value = "#   -Устройство САУ и АИиС НК-38СТ   инвентарный № 05Y-00700";
                worksheet.Cell(footerRow + 2, 2).Style.Font.Bold = true;

                worksheet.Cell(footerRow + 4, 2).Value = "##   -Стенд автономных испытания опыт. горелок камеры сгорания изд. НК-38СТ, НК-16СТ  инвентарный №СТ-00032";
                worksheet.Cell(footerRow + 4, 2).Style.Font.Bold = true;

                worksheet.Cell(footerRow + 6, 2).Value = "###   -Стенд№2 инвентарный №01Y-00404";
                worksheet.Cell(footerRow + 6, 2).Style.Font.Bold = true;

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

                // Объединяем ячейки для подписи
                worksheet.Range(footerRow + 2, 2, footerRow + 2, 10).Merge();
                worksheet.Range(footerRow + 4, 2, footerRow + 4, 10).Merge();
                worksheet.Range(footerRow + 6, 2, footerRow + 6, 10).Merge();

                worksheet.Range(footerRow + 8, 2, footerRow + 8, 13).Merge();
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


                workbook.SaveAs(filePath);
            }
        }
    }
}

