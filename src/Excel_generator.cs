using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;

namespace EVM_ECLR.Classes
{
    class ExcelGenerator
    {
        public static void CreateExcelFile(string filePath,
                                           string jitPath,
                                           string forecastPath,
                                           string totalHoursPath,
                                           string routingPath,
                                           string backlogPath)
        {
            XLWorkbook workbook;
            IXLWorksheet worksheet;

            bool fileExist = File.Exists(filePath);

            if (fileExist)
                workbook = new XLWorkbook(filePath);
            else
                workbook = new XLWorkbook();

            // === 1) T_JIT ===
            worksheet = workbook.Worksheets.Contains("T_JIT")
                ? workbook.Worksheet("T_JIT")
                : workbook.Worksheets.Add("T_JIT");

            Fill_T_JIT(worksheet, jitPath);

            // === 2) Forecast ===
            worksheet = workbook.Worksheets.Contains("Forecast")
                ? workbook.Worksheet("Forecast")
                : workbook.Worksheets.Add("Forecast");

            Fill_Forecast(worksheet, forecastPath);

            // === 3) Total Hours ===
            worksheet = workbook.Worksheets.Contains("Total_Hours")
                ? workbook.Worksheet("Total_Hours")
                : workbook.Worksheets.Add("Total_Hours");

            Fill_TotalHours(worksheet, totalHoursPath);

            // === 4) Routing ===
            worksheet = workbook.Worksheets.Contains("Routing")
                ? workbook.Worksheet("Routing")
                : workbook.Worksheets.Add("Routing");

            Fill_Routing(worksheet, routingPath);

            // === 5) Backlog ===
            worksheet = workbook.Worksheets.Contains("Backlog")
                ? workbook.Worksheet("Backlog")
                : workbook.Worksheets.Add("Backlog");

            Fill_Backlog(worksheet, backlogPath);

            workbook.SaveAs(filePath);
        }

        // -----------------------------
        // 1) T_JIT
        // -----------------------------
        private static void Fill_T_JIT(IXLWorksheet ws, string path)
        {
            var source = new XLWorkbook(path).Worksheet(1);

            ws.Cell(1, 1).Value = "trans_date";
            ws.Cell(1, 2).Value = "item";
            ws.Cell(1, 3).Value = "description";
            ws.Cell(1, 4).Value = "qty";

            int row = 2;

            foreach (var r in source.RowsUsed().Skip(1))
            {
                ws.Cell(row, 1).Value = r.Cell("B").Value; // trans_date
                ws.Cell(row, 2).Value = r.Cell("C").Value; // item
                ws.Cell(row, 3).Value = r.Cell("D").Value; // description
                ws.Cell(row, 4).Value = r.Cell("E").Value; // qty
                row++;
            }

            ws.Columns().AdjustToContents();
        }

        // -----------------------------
        // 2) Forecast
        // -----------------------------
        private static void Fill_Forecast(IXLWorksheet ws, string path)
        {
            var source = new XLWorkbook(path).Worksheet(1);

            int colCount = source.Row(1).CellsUsed().Count();

            for (int c = 1; c <= colCount; c++)
                ws.Cell(1, c).Value = source.Cell(1, c).Value;

            int row = 2;

            foreach (var r in source.RowsUsed().Skip(1))
            {
                for (int c = 1; c <= colCount; c++)
                    ws.Cell(row, c).Value = r.Cell(c).Value;

                row++;
            }

            ws.Columns().AdjustToContents();
        }

        // -----------------------------
        // 3) Total Hours
        // -----------------------------
        private static void Fill_TotalHours(IXLWorksheet ws, string path)
        {
            var source = new XLWorkbook(path).Worksheet(1);

            ws.Cell(1, 1).Value = "Month";
            ws.Cell(1, 2).Value = "Gross Hours";

            int row = 2;

            foreach (var r in source.RowsUsed().Skip(1))
            {
                ws.Cell(row, 1).Value = r.Cell("A").Value; // Month
                ws.Cell(row, 2).Value = r.Cell("H").Value; // Gross Hrs
                row++;
            }

            ws.Columns().AdjustToContents();
        }

        // -----------------------------
        // 4) Routing
        // -----------------------------
        private static void Fill_Routing(IXLWorksheet ws, string path)
        {
            var source = new XLWorkbook(path).Worksheet(1);

            ws.Cell(1, 1).Value = "Item";
            ws.Cell(1, 2).Value = "Hours";

            int row = 2;

            foreach (var r in source.RowsUsed().Skip(1))
            {
                ws.Cell(row, 1).Value = r.Cell("A").Value; // Item
                ws.Cell(row, 2).Value = r.Cell("G").Value; // Hours
                row++;
            }

            ws.Columns().AdjustToContents();
        }

        // -----------------------------
        // 5) Backlog
        // -----------------------------
        private static void Fill_Backlog(IXLWorksheet ws, string path)
        {
            var source = new XLWorkbook(path).Worksheet(1);

            ws.Cell(1, 1).Value = "FamilyCode";
            ws.Cell(1, 2).Value = "OrderNumber";
            ws.Cell(1, 3).Value = "Line";
            ws.Cell(1, 4).Value = "DueDate";
            ws.Cell(1, 5).Value = "Description";
            ws.Cell(1, 6).Value = "Item";
            ws.Cell(1, 7).Value = "OpenOrderQty";

            int row = 2;

            foreach (var r in source.RowsUsed().Skip(1))
            {
                ws.Cell(row, 1).Value = r.Cell("A").Value;
                ws.Cell(row, 2).Value = r.Cell("B").Value;
                ws.Cell(row, 3).Value = r.Cell("C").Value;
                ws.Cell(row, 4).Value = r.Cell("D").Value;
                ws.Cell(row, 5).Value = r.Cell("F").Value;
                ws.Cell(row, 6).Value = r.Cell("G").Value;
                ws.Cell(row, 7).Value = r.Cell("H").Value;
                row++;
            }

            ws.Columns().AdjustToContents();
        }
        private static void RepairExcelFile(IXLWorksheet ws, List<string[]> newRows, int keyColumnIndex)
        {
            var existingRows = ws.RowsUsed().Skip(1); // skip header

            foreach (var newRow in newRows)
            {
                string newKey = newRow[keyColumnIndex - 1];

                bool updated = false;

                foreach (var row in existingRows)
                {
                    string existingKey = row.Cell(keyColumnIndex).Value.ToString();

                    if (existingKey == newKey)
                    {
                        // Update row
                        for (int c = 0; c < newRow.Length; c++)
                            row.Cell(c + 1).Value = newRow[c];

                        updated = true;
                        break;
                    }
                }

                if (!updated)
                {
                    // Add new row at bottom
                    int lastRow = ws.LastRowUsed().RowNumber() + 1;

                    for (int c = 0; c < newRow.Length; c++)
                        ws.Cell(lastRow, c + 1).Value = newRow[c];
                }
            }

            ws.Columns().AdjustToContents();
        }

    }
}
