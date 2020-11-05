using Fast.Web.Models;
using Fast.Web.Models.LPH;
using Fast.Web.Models.LPH.PP;
using Fast.Web.Models.Report;
using Fast.Web.Resources;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Drawing;

namespace Fast.Web.Utils
{
    public static class ExcelGenerator
    {
        private static string HEADER_COLOR = "#99ccff";

        #region ::Report Plan Actual Raw Data::
        public static byte[] ExportReportActualPlanningRawData(List<ReportActualPlanningShiftModel> dataList, DateTime startDate, DateTime endDate, string username)
        {
            ExcelPackage Ep = new ExcelPackage();
            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Raw Data");
            AddHeader(Sheet, "Plan vs Actual Report Raw Data", username);

            Color colFromHex = ColorTranslator.FromHtml(HEADER_COLOR);
            Sheet.Cells["A6:J6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Sheet.Cells["A6:J6"].Style.Fill.BackgroundColor.SetColor(colFromHex);
            Sheet.Cells["A6"].Value = "Date";
            Sheet.Cells["B6"].Value = "Shift";
            Sheet.Cells["C6"].Value = "Item Code";
            Sheet.Cells["D6"].Value = "Description";
            Sheet.Cells["E6"].Value = "Market";
            Sheet.Cells["F6"].Value = "Status";
            Sheet.Cells["G6"].Value = "Quantity";
            Sheet.Cells["H6"].Value = "UOM";
            Sheet.Cells["I6"].Value = "LU";
            Sheet.Cells["J6"].Value = "Prod Center";

            Sheet.Cells["A6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["B6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["C6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["D6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["E6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["F6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["G6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["H6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["I6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["J6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            int row = 7;
            foreach (var item in dataList)
            {
                Sheet.Cells[string.Format("A{0}", row)].Value = item.Date.ToString("dd-MMM-yy");
                Sheet.Cells[string.Format("B{0}", row)].Value = item.Shift;
                Sheet.Cells[string.Format("C{0}", row)].Value = item.ItemCode;
                Sheet.Cells[string.Format("D{0}", row)].Value = item.Description;
                Sheet.Cells[string.Format("E{0}", row)].Value = item.Market;
                Sheet.Cells[string.Format("F{0}", row)].Value = item.Type;
                Sheet.Cells[string.Format("G{0}", row)].Value = item.Total;
                Sheet.Cells[string.Format("H{0}", row)].Value = item.UOM;
                Sheet.Cells[string.Format("I{0}", row)].Value = item.LU;
                Sheet.Cells[string.Format("J{0}", row)].Value = item.Location;

                Sheet.Cells[string.Format("A{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("B{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("C{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("D{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("E{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("F{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("G{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("H{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("I{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("J{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                row++;
            }

            Sheet.Cells["A:J"].AutoFitColumns();

            return Ep.GetAsByteArray();
        }
        #endregion

        #region ::Report Actual Planning::
        public static byte[] ExportReportActualPlanningByBrand(List<ReportActualPlanningShiftModel> dataList, DateTime startDate, DateTime endDate, string username, string location, string granularity, List<string> weekList)
        {
            ExcelPackage Ep = new ExcelPackage();

            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add(location);
            AddHeader(Sheet, "Planning vs Actual Report - " + granularity, username);

            if (granularity == "Shiftly")
            {
                Sheet.Cells["A7"].Value = "Item Code";
                Sheet.Cells["B7"].Value = "Description";
                Sheet.Cells["C7"].Value = "Type";
                Sheet.Cells["D7"].Value = "UOM";
            }
            else
            {
                Sheet.Cells["A6"].Value = "Item Code";
                Sheet.Cells["B6"].Value = "Description";
                Sheet.Cells["C6"].Value = "Type";
                Sheet.Cells["D6"].Value = "UOM";
            }

            Sheet.Cells["A6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["B6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["C6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["D6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            if (granularity == "Shiftly")
            {
                Sheet.Cells["A7"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells["B7"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells["C7"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells["D7"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            }

            int dateIndex = 0;
            int maxIndex = 0;
            string columnName = "";

            if (granularity == "Shiftly")
            {
                for (var day = startDate.Date; day.Date <= endDate.Date; day = day.AddDays(1))
                {
                    columnName = Helper.GetExcelColumnName((6 + (dateIndex * 3)));
                    var columnName5 = Helper.GetExcelColumnName((5 + (dateIndex * 3)));
                    var columnName7 = Helper.GetExcelColumnName((7 + (dateIndex * 3)));
                    Sheet.Cells[string.Format("{0}6", columnName)].Value = day.ToString("dd-MMM-yy");
                    Sheet.Cells[string.Format("{0}6:{1}6", columnName5, columnName7)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    columnName = Helper.GetExcelColumnName((5 + (dateIndex * 3)));
                    Sheet.Cells[string.Format("{0}7", columnName)].Value = "Shift 1";
                    Sheet.Cells[string.Format("{0}7", columnName)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    columnName = Helper.GetExcelColumnName((6 + (dateIndex * 3)));
                    Sheet.Cells[string.Format("{0}7", columnName)].Value = "Shift 2";
                    Sheet.Cells[string.Format("{0}7", columnName)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    columnName = Helper.GetExcelColumnName((7 + (dateIndex * 3)));
                    Sheet.Cells[string.Format("{0}7", columnName)].Value = "Shift 3";
                    Sheet.Cells[string.Format("{0}7", columnName)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    maxIndex = 7 + (dateIndex * 3);
                    dateIndex++;
                }
            }
            else if (granularity == "Daily")
            {
                for (var day = startDate.Date; day.Date <= endDate.Date; day = day.AddDays(1))
                {
                    columnName = Helper.GetExcelColumnName(5 + (dateIndex * 1));
                    Sheet.Cells[string.Format("{0}6", columnName)].Value = day.ToString("dd-MMM-yy");
                    Sheet.Cells[string.Format("{0}6", columnName)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    maxIndex = 5 + (dateIndex * 1);
                    dateIndex++;
                }
            }
            else
            {
                foreach (var week in weekList)
                {
                    columnName = Helper.GetExcelColumnName(5 + (dateIndex * 1));
                    Sheet.Cells[string.Format("{0}6", columnName)].Value = "Week - " + week;
                    Sheet.Cells[string.Format("{0}6", columnName)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    maxIndex = 5 + (dateIndex * 1);
                    dateIndex++;
                }
            }

            columnName = Helper.GetExcelColumnName(maxIndex + 1);
            Sheet.Cells[string.Format("{0}6", columnName)].Value = "TOTAL";
            Sheet.Cells[string.Format("{0}6", columnName)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            if (granularity == "Shiftly")
                Sheet.Cells[string.Format("{0}7", columnName)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            string lastcolumnName = Helper.GetExcelColumnName(maxIndex + 2);
            Sheet.Cells[string.Format("{0}6", lastcolumnName)].Value = "Remarks";
            Sheet.Cells[string.Format("{0}6", lastcolumnName)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            if (granularity == "Shiftly")
                Sheet.Cells[string.Format("{0}7", lastcolumnName)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            int rowIndex = granularity == "Shiftly" ? 8 : 7;
            int groupIndex = 1;

            foreach (var item in dataList)
            {
                Sheet.Cells[string.Format("A{0}", rowIndex)].Value = item.ItemCode;
                Sheet.Cells[string.Format("B{0}", rowIndex)].Value = item.Description;
                Sheet.Cells[string.Format("C{0}", rowIndex)].Value = item.Type;
                Sheet.Cells[string.Format("D{0}", rowIndex)].Value = item.UOM;
                Sheet.Cells[string.Format("A{0}", rowIndex)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("B{0}", rowIndex)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("C{0}", rowIndex)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("D{0}", rowIndex)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                dateIndex = 0;
                if (granularity == "Shiftly")
                {
                    for (var day = startDate.Date; day.Date <= endDate.Date; day = day.AddDays(1))
                    {
                        columnName = Helper.GetExcelColumnName((5 + (dateIndex * 3)));
                        if (groupIndex % 3 == 0)
                        {
                            Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Style.Numberformat.Format = "#0\\.00%";
                            Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Value = item.ShiftList[dateIndex].Shift1;
                        }
                        else
                        {
                            Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Style.Numberformat.Format = "0.00";
                            if (item.ShiftList[dateIndex].Shift1 != 0)
                                Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Value = item.ShiftList[dateIndex].Shift1;
                        }

                        Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                        columnName = Helper.GetExcelColumnName((6 + (dateIndex * 3)));
                        if (groupIndex % 3 == 0)
                        {
                            Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Style.Numberformat.Format = "#0\\.00%";
                            Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Value = item.ShiftList[dateIndex].Shift2;
                        }
                        else
                        {
                            Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Style.Numberformat.Format = "0.00";
                            if (item.ShiftList[dateIndex].Shift2 != 0)
                                Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Value = item.ShiftList[dateIndex].Shift2;
                        }
                        Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                        columnName = Helper.GetExcelColumnName((7 + (dateIndex * 3)));
                        if (groupIndex % 3 == 0)
                        {
                            Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Style.Numberformat.Format = "#0\\.00%";
                            Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Value = item.ShiftList[dateIndex].Shift3;
                        }
                        else
                        {
                            Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Style.Numberformat.Format = "0.00";
                            if (item.ShiftList[dateIndex].Shift3 != 0)
                                Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Value = item.ShiftList[dateIndex].Shift3;
                        }
                        Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                        dateIndex++;
                    }
                }
                else if (granularity == "Daily")
                {
                    for (var day = startDate.Date; day.Date <= endDate.Date; day = day.AddDays(1))
                    {
                        columnName = Helper.GetExcelColumnName((5 + (dateIndex * 1)));
                        if (groupIndex % 3 == 0)
                        {
                            Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Style.Numberformat.Format = "#0\\.00%";
                            Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Value = item.ShiftList[dateIndex].AllShift;
                        }
                        else
                        {
                            Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Style.Numberformat.Format = "0.00";
                            if (item.ShiftList[dateIndex].AllShift != 0)
                                Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Value = item.ShiftList[dateIndex].AllShift;
                        }

                        Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                        dateIndex++;
                    }
                }
                else
                {
                    foreach (var week in weekList)
                    {
                        columnName = Helper.GetExcelColumnName((5 + (dateIndex * 1)));
                        if (groupIndex % 3 == 0)
                        {
                            Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Style.Numberformat.Format = "#0\\.00%";
                            Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Value = item.ShiftList[dateIndex].AllShift;
                        }
                        else
                        {
                            Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Style.Numberformat.Format = "0.00";
                            if (item.ShiftList[dateIndex].AllShift != 0)
                                Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Value = item.ShiftList[dateIndex].AllShift;
                        }
                        Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                        dateIndex++;
                    }
                }

                dateIndex--;
                columnName = granularity == "Shiftly" ? Helper.GetExcelColumnName((8 + (dateIndex * 3))) : Helper.GetExcelColumnName((6 + (dateIndex * 1)));
                if (groupIndex % 3 == 0)
                {
                    Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Style.Numberformat.Format = "#0\\.00%";
                    Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Value = item.Total;
                }
                else
                {
                    Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Style.Numberformat.Format = "0.00";
                    if (item.Total != 0)
                        Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Value = item.Total;
                }

                Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                columnName = granularity == "Shiftly" ? Helper.GetExcelColumnName((9 + (dateIndex * 3))) : Helper.GetExcelColumnName((7 + (dateIndex * 1)));
                Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                if (groupIndex > 3)
                {
                    Sheet.Cells[string.Format("A{0}:{1}{2}", rowIndex, lastcolumnName, rowIndex)].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    Sheet.Cells[string.Format("A{0}:{1}{2}", rowIndex, lastcolumnName, rowIndex)].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                }

                rowIndex++;
                groupIndex++;

                if (groupIndex > 6)
                    groupIndex = 1;
            }

            Color colFromHex = ColorTranslator.FromHtml(HEADER_COLOR);
            Sheet.Cells[string.Format("A6:{0}6", lastcolumnName)].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Sheet.Cells[string.Format("A6:{0}6", lastcolumnName)].Style.Fill.BackgroundColor.SetColor(colFromHex);

            if (granularity == "Shiftly")
            {
                Sheet.Cells[string.Format("A7:{0}7", lastcolumnName)].Style.Fill.PatternType = ExcelFillStyle.Solid;
                Sheet.Cells[string.Format("A7:{0}7", lastcolumnName)].Style.Fill.BackgroundColor.SetColor(colFromHex);
            }

            Sheet.Cells[string.Format("A:{0}", lastcolumnName)].AutoFitColumns();

            return Ep.GetAsByteArray();
        }
        #endregion

        #region ::Report Actual Planning by LU::
        public static byte[] ExportReportActualPlanningByLU(List<ReportActualPlanningShiftModel> dataList, DateTime startDate, DateTime endDate, string username, string location, List<long> locationIDList, string granularity, List<string> weekList)
        {
            ExcelPackage Ep = new ExcelPackage();

            int index = 0;

            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add(location);
            AddHeader(Sheet, "Planning vs Actual Report - " + granularity, username);

            if (granularity == "Shiftly")
            {
                Sheet.Cells["A7"].Value = "Production Center";
                Sheet.Cells["B7"].Value = "LU";
                Sheet.Cells["C7"].Value = "Market";
                Sheet.Cells["D7"].Value = "Item Code";
                Sheet.Cells["E7"].Value = "Description";
                Sheet.Cells["F7"].Value = "Type";
                Sheet.Cells["G7"].Value = "UOM";
            }
            else
            {
                Sheet.Cells["A6"].Value = "Production Center";
                Sheet.Cells["B6"].Value = "LU";
                Sheet.Cells["C6"].Value = "Market";
                Sheet.Cells["D6"].Value = "Item Code";
                Sheet.Cells["E6"].Value = "Description";
                Sheet.Cells["F6"].Value = "Type";
                Sheet.Cells["G6"].Value = "UOM";
            }

            Sheet.Cells["A6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["B6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["C6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["D6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["E6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["F6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["G6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            if (granularity == "Shiftly")
            {
                Sheet.Cells["A7"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells["B7"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells["C7"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells["D7"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells["E7"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells["F7"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells["G7"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            }

            int dateIndex = 0;
            int maxIndex = 0;
            string columnName = "";

            if (granularity == "Shiftly")
            {
                for (var day = startDate.Date; day.Date <= endDate.Date; day = day.AddDays(1))
                {
                    columnName = Helper.GetExcelColumnName((9 + (dateIndex * 3)));
                    var columnName5 = Helper.GetExcelColumnName((8 + (dateIndex * 3)));
                    var columnName7 = Helper.GetExcelColumnName((10 + (dateIndex * 3)));
                    Sheet.Cells[string.Format("{0}6", columnName)].Value = day.ToString("dd-MMM-yy");
                    Sheet.Cells[string.Format("{0}6:{1}6", columnName5, columnName7)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    columnName = Helper.GetExcelColumnName((8 + (dateIndex * 3)));
                    Sheet.Cells[string.Format("{0}7", columnName)].Value = "Shift 1";
                    Sheet.Cells[string.Format("{0}7", columnName)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    columnName = Helper.GetExcelColumnName((9 + (dateIndex * 3)));
                    Sheet.Cells[string.Format("{0}7", columnName)].Value = "Shift 2";
                    Sheet.Cells[string.Format("{0}7", columnName)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    columnName = Helper.GetExcelColumnName((10 + (dateIndex * 3)));
                    Sheet.Cells[string.Format("{0}7", columnName)].Value = "Shift 3";
                    Sheet.Cells[string.Format("{0}7", columnName)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    maxIndex = 10 + (dateIndex * 3);
                    dateIndex++;
                }
            }
            else if (granularity == "Daily")
            {
                for (var day = startDate.Date; day.Date <= endDate.Date; day = day.AddDays(1))
                {
                    columnName = Helper.GetExcelColumnName(8 + (dateIndex * 1));
                    Sheet.Cells[string.Format("{0}6", columnName)].Value = day.ToString("dd-MMM-yy");
                    Sheet.Cells[string.Format("{0}6", columnName)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    maxIndex = 8 + (dateIndex * 1);
                    dateIndex++;
                }
            }
            else
            {
                foreach (var week in weekList)
                {
                    columnName = Helper.GetExcelColumnName(8 + (dateIndex * 1));
                    Sheet.Cells[string.Format("{0}6", columnName)].Value = "Week - " + week;
                    Sheet.Cells[string.Format("{0}6", columnName)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    maxIndex = 8 + (dateIndex * 1);
                    dateIndex++;
                }
            }

            columnName = Helper.GetExcelColumnName(maxIndex + 1);
            Sheet.Cells[string.Format("{0}6", columnName)].Value = "TOTAL";
            Sheet.Cells[string.Format("{0}6", columnName)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            if (granularity == "Shiftly")
                Sheet.Cells[string.Format("{0}7", columnName)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            string lastcolumnName = Helper.GetExcelColumnName(maxIndex + 2);
            Sheet.Cells[string.Format("{0}6", lastcolumnName)].Value = "Remarks";
            Sheet.Cells[string.Format("{0}6", lastcolumnName)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            if (granularity == "Shiftly")
                Sheet.Cells[string.Format("{0}7", lastcolumnName)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            int rowIndex = granularity == "Shiftly" ? 8 : 7;
            int groupIndex = 1;

            foreach (var item in dataList)
            {
                Sheet.Cells[string.Format("A{0}", rowIndex)].Value = item.Location;
                Sheet.Cells[string.Format("B{0}", rowIndex)].Value = item.LU;
                Sheet.Cells[string.Format("C{0}", rowIndex)].Value = item.Market;
                Sheet.Cells[string.Format("D{0}", rowIndex)].Value = item.ItemCode;
                Sheet.Cells[string.Format("E{0}", rowIndex)].Value = item.Description;
                Sheet.Cells[string.Format("F{0}", rowIndex)].Value = item.Type;
                Sheet.Cells[string.Format("G{0}", rowIndex)].Value = item.UOM;

                if (item.Location == "Sub Total")
                {
                    Sheet.Cells[string.Format("A{0}", rowIndex)].Style.Font.Bold = true;
                    Sheet.Cells[string.Format("A{0}", rowIndex)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                }
                else
                    Sheet.Cells[string.Format("A{0}", rowIndex)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("B{0}", rowIndex)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("C{0}", rowIndex)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("D{0}", rowIndex)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("E{0}", rowIndex)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                if (item.Location == "Sub Total")
                {
                    Sheet.Cells[string.Format("F{0}", rowIndex)].Style.Font.Bold = true;
                    Sheet.Cells[string.Format("F{0}", rowIndex)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                }
                else
                    Sheet.Cells[string.Format("F{0}", rowIndex)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("G{0}", rowIndex)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                dateIndex = 0;
                if (granularity == "Shiftly")
                {
                    for (var day = startDate.Date; day.Date <= endDate.Date; day = day.AddDays(1))
                    {
                        columnName = Helper.GetExcelColumnName((8 + (dateIndex * 3)));
                        if (groupIndex % 3 == 0)
                        {
                            Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Style.Numberformat.Format = "#0\\.00%";
                            Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Value = item.ShiftList[dateIndex].Shift1;
                        }
                        else
                        {
                            Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Style.Numberformat.Format = "0.00";
                            if (item.ShiftList[dateIndex].Shift1 != 0)
                                Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Value = item.ShiftList[dateIndex].Shift1;
                        }

                        Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                        columnName = Helper.GetExcelColumnName((9 + (dateIndex * 3)));
                        if (groupIndex % 3 == 0)
                        {
                            Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Style.Numberformat.Format = "#0\\.00%";
                            Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Value = item.ShiftList[dateIndex].Shift2;
                        }
                        else
                        {
                            Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Style.Numberformat.Format = "0.00";
                            if (item.ShiftList[dateIndex].Shift2 != 0)
                                Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Value = item.ShiftList[dateIndex].Shift2;
                        }
                        Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                        columnName = Helper.GetExcelColumnName((10 + (dateIndex * 3)));
                        if (groupIndex % 3 == 0)
                        {
                            Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Style.Numberformat.Format = "#0\\.00%";
                            Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Value = item.ShiftList[dateIndex].Shift3;
                        }
                        else
                        {
                            Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Style.Numberformat.Format = "0.00";
                            if (item.ShiftList[dateIndex].Shift3 != 0)
                                Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Value = item.ShiftList[dateIndex].Shift3;
                        }
                        Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                        dateIndex++;
                    }
                }
                else if (granularity == "Daily")
                {
                    for (var day = startDate.Date; day.Date <= endDate.Date; day = day.AddDays(1))
                    {
                        columnName = Helper.GetExcelColumnName((8 + (dateIndex * 1)));
                        if (groupIndex % 3 == 0)
                        {
                            Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Style.Numberformat.Format = "#0\\.00%";
                            Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Value = item.ShiftList[dateIndex].AllShift;
                        }
                        else
                        {
                            Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Style.Numberformat.Format = "0.00";
                            if (item.ShiftList[dateIndex].AllShift != 0)
                                Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Value = item.ShiftList[dateIndex].AllShift;
                        }
                        Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                        dateIndex++;
                    }
                }
                else
                {
                    foreach (var week in weekList)
                    {
                        columnName = Helper.GetExcelColumnName((8 + (dateIndex * 1)));
                        if (groupIndex % 3 == 0)
                        {
                            Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Style.Numberformat.Format = "#0\\.00%";
                            Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Value = item.ShiftList[dateIndex].AllShift;
                        }
                        else
                        {
                            Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Style.Numberformat.Format = "0.00";
                            if (item.ShiftList[dateIndex].AllShift != 0)
                                Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Value = item.ShiftList[dateIndex].AllShift;
                        }
                        Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                        dateIndex++;
                    }
                }

                dateIndex--;
                columnName = granularity == "Shiftly" ? Helper.GetExcelColumnName((11 + (dateIndex * 3))) : Helper.GetExcelColumnName((9 + (dateIndex * 1)));
                if (groupIndex % 3 == 0)
                {
                    Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Style.Numberformat.Format = "#0\\.00%";
                    Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Value = item.Total;
                }
                else
                {
                    Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Style.Numberformat.Format = "0.00";
                    if (item.Total != 0)
                        Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Value = item.Total;
                }
                Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                columnName = granularity == "Shiftly" ? Helper.GetExcelColumnName((12 + (dateIndex * 3))) : Helper.GetExcelColumnName((10 + (dateIndex * 1))); ;
                Sheet.Cells[string.Format("{0}{1}", columnName, rowIndex)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                rowIndex++;
                groupIndex++;

                if (groupIndex > 9)
                    groupIndex = 1;
            }

            Color colFromHex = ColorTranslator.FromHtml(HEADER_COLOR);
            Sheet.Cells[string.Format("A6:{0}6", lastcolumnName)].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Sheet.Cells[string.Format("A6:{0}6", lastcolumnName)].Style.Fill.BackgroundColor.SetColor(colFromHex);
            Sheet.Cells[string.Format("A7:{0}7", lastcolumnName)].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Sheet.Cells[string.Format("A7:{0}7", lastcolumnName)].Style.Fill.BackgroundColor.SetColor(colFromHex);

            Sheet.Cells[string.Format("A:{0}", lastcolumnName)].AutoFitColumns();
            index++;

            return Ep.GetAsByteArray();
        }
        #endregion

        #region ::LPH History Primary::
        public static byte[] ExportLPHHistoryPrimary(List<PPLPHSubmissionsModel> dataList, string username)
        {
            ExcelPackage Ep = new ExcelPackage();
            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("LPH-PP-History");
            AddHeader(Sheet, "LPH PP History", username);

            Color colFromHex = ColorTranslator.FromHtml(HEADER_COLOR);
            Sheet.Cells["A6:I6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Sheet.Cells["A6:I6"].Style.Fill.BackgroundColor.SetColor(colFromHex);
            Sheet.Cells["A6"].Value = "LPH ID";
            Sheet.Cells["B6"].Value = "Date";
            Sheet.Cells["C6"].Value = "Shift";
            Sheet.Cells["D6"].Value = "LPH Type";
            Sheet.Cells["E6"].Value = "Created By";
            Sheet.Cells["F6"].Value = "Location";
            Sheet.Cells["G6"].Value = "Status";
            Sheet.Cells["H6"].Value = "Created at";
            Sheet.Cells["I6"].Value = "Status Changed at";

            Sheet.Cells["A6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["B6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["C6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["D6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["E6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["F6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["G6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["H6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["I6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            int row = 7;
            foreach (var item in dataList)
            {
                Sheet.Cells[string.Format("A{0}", row)].Value = item.LPHID;
                Sheet.Cells[string.Format("B{0}", row)].Value = item.Date.ToString("dd-MMM-yy");
                Sheet.Cells[string.Format("C{0}", row)].Value = item.Shift;
                Sheet.Cells[string.Format("D{0}", row)].Value = item.LPHHeader;
                Sheet.Cells[string.Format("E{0}", row)].Value = item.ModifiedBy;
                Sheet.Cells[string.Format("F{0}", row)].Value = item.Location;
                Sheet.Cells[string.Format("G{0}", row)].Value = item.Status;
                Sheet.Cells[string.Format("H{0}", row)].Value = item.CreatedAt.ToString("dd-MMM-yy HH:mm");
                Sheet.Cells[string.Format("I{0}", row)].Value = item.StatusChangedAt.HasValue ? item.StatusChangedAt.Value.ToString("dd-MMM-yy HH:mm") : string.Empty;

                Sheet.Cells[string.Format("A{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("B{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("C{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("D{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("E{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("F{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("G{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("H{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("I{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                row++;
            }

            Sheet.Cells["A:I"].AutoFitColumns();

            return Ep.GetAsByteArray();
        }
        #endregion

        #region ::LPH History::
        public static byte[] ExportLPHHistory(List<LPHSubmissionsModel> dataList, string username)
        {
            ExcelPackage Ep = new ExcelPackage();
            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("LPH-SP-History");
            AddHeader(Sheet, "LPH SP History", username);

            Color colFromHex = ColorTranslator.FromHtml(HEADER_COLOR);
            Sheet.Cells["A6:J6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Sheet.Cells["A6:J6"].Style.Fill.BackgroundColor.SetColor(colFromHex);
            Sheet.Cells["A6"].Value = "LPH ID";
            Sheet.Cells["B6"].Value = "Date";
            Sheet.Cells["C6"].Value = "Shift";
            Sheet.Cells["D6"].Value = "LPH Type";
            Sheet.Cells["E6"].Value = "Machine";
            Sheet.Cells["F6"].Value = "Created By";
            Sheet.Cells["G6"].Value = "Location";
            Sheet.Cells["H6"].Value = "Status";
            Sheet.Cells["I6"].Value = "Created at";
            Sheet.Cells["J6"].Value = "Status Changed at";

            Sheet.Cells["A6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["B6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["C6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["D6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["E6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["F6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["G6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["H6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["I6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["J6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            int row = 7;
            foreach (var item in dataList)
            {
                Sheet.Cells[string.Format("A{0}", row)].Value = item.LPHID;
                Sheet.Cells[string.Format("B{0}", row)].Value = item.Date.ToString("dd-MMM-yy");
                Sheet.Cells[string.Format("C{0}", row)].Value = item.Shift;
                Sheet.Cells[string.Format("D{0}", row)].Value = item.LPHHeader;
                Sheet.Cells[string.Format("E{0}", row)].Value = item.Machine;
                Sheet.Cells[string.Format("F{0}", row)].Value = item.ModifiedBy;
                Sheet.Cells[string.Format("G{0}", row)].Value = item.Location;
                Sheet.Cells[string.Format("H{0}", row)].Value = item.Status;
                Sheet.Cells[string.Format("I{0}", row)].Value = item.ModifiedDate.HasValue ? item.ModifiedDate.Value.ToString("dd-MMM-yy HH:mm") : string.Empty;
                Sheet.Cells[string.Format("J{0}", row)].Value = item.StatusChangedAt.HasValue ? item.StatusChangedAt.Value.ToString("dd-MMM-yy HH:mm") : string.Empty;

                Sheet.Cells[string.Format("A{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("B{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("C{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("D{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("E{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("F{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("G{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("H{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("I{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("J{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                row++;
            }

            Sheet.Cells["A:J"].AutoFitColumns();

            return Ep.GetAsByteArray();
        }
        #endregion

        #region ::LPH Approval::
        public static byte[] ExportLPHApproval(List<LPHApprovalsModel> dataList, string username)
        {
            ExcelPackage Ep = new ExcelPackage();
            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("LPH-SP-Approval");
            AddHeader(Sheet, "LPH SP Approval", username);

            Color colFromHex = ColorTranslator.FromHtml(HEADER_COLOR);
            Sheet.Cells["A6:H6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Sheet.Cells["A6:H6"].Style.Fill.BackgroundColor.SetColor(colFromHex);
            Sheet.Cells["A6"].Value = "LPH Type";
            Sheet.Cells["B6"].Value = "Submission ID";
            Sheet.Cells["C6"].Value = "User";
            Sheet.Cells["D6"].Value = "Date";
            Sheet.Cells["E6"].Value = "Shift";
            Sheet.Cells["F6"].Value = "Machine";
            Sheet.Cells["G6"].Value = "Status";
            Sheet.Cells["H6"].Value = "Last Update";

            Sheet.Cells["A6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["B6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["C6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["D6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["E6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["F6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["G6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["H6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            int row = 7;
            foreach (var item in dataList)
            {
                Sheet.Cells[string.Format("A{0}", row)].Value = item.LPHType;
                Sheet.Cells[string.Format("B{0}", row)].Value = item.LPHSubmissionID;
                Sheet.Cells[string.Format("C{0}", row)].Value = item.User;
                Sheet.Cells[string.Format("D{0}", row)].Value = item.Date.ToString("dd-MMM-yy");
                Sheet.Cells[string.Format("E{0}", row)].Value = item.Shift;
                Sheet.Cells[string.Format("F{0}", row)].Value = item.Machine;
                Sheet.Cells[string.Format("G{0}", row)].Value = item.Status;
                Sheet.Cells[string.Format("H{0}", row)].Value = item.ModifiedDate.HasValue ? item.ModifiedDate.Value.ToString("dd-MMM-yy HH:mm") : string.Empty;

                Sheet.Cells[string.Format("A{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("B{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("C{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("D{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("E{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("F{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("G{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("H{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                row++;
            }

            Sheet.Cells["A:H"].AutoFitColumns();

            return Ep.GetAsByteArray();
        }
        public static byte[] ExportLPHPPApproval(List<PPLPHApprovalsModel> dataList, string username)
        {
            ExcelPackage Ep = new ExcelPackage();
            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("LPH-PP-Approval");
            AddHeader(Sheet, "LPH PP Approval", username);

            Color colFromHex = ColorTranslator.FromHtml(HEADER_COLOR);
            Sheet.Cells["A6:G6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Sheet.Cells["A6:G6"].Style.Fill.BackgroundColor.SetColor(colFromHex);
            Sheet.Cells["A6"].Value = "LPH Type";
            Sheet.Cells["B6"].Value = "Submission ID";
            Sheet.Cells["C6"].Value = "User";
            Sheet.Cells["D6"].Value = "Date";
            Sheet.Cells["E6"].Value = "Shift";
            Sheet.Cells["F6"].Value = "Status";
            Sheet.Cells["G6"].Value = "Last Update";

            Sheet.Cells["A6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["B6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["C6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["D6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["E6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["F6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["G6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            int row = 7;
            foreach (var item in dataList)
            {
                var lphtype = item.LPHType.Trim();
                if (lphtype == "LPHPrimaryKretekLineAddback")
                    lphtype = "Kretek Line - Addback";
                else if (lphtype == "LPHPrimaryDiet")
                    lphtype = "Intermediate Line - DIET";
                else if (lphtype == "LPHPrimaryCloveInfeedConditioning")
                    lphtype = "Intermediate Line - Clove Feeding & DCCC";
                else if (lphtype == "LPHPrimaryCSFCutDryPacking")
                    lphtype = "Intermediate Line - CSF Cut Dry & Packing";
                else if (lphtype == "LPHPrimaryCSFInfeedConditioning")
                    lphtype = "Intermediate Line - CSF Feeding & DCCC";
                else if (lphtype == "LPHPrimaryCloveCutDryPacking")
                    lphtype = "Intermediate Line - Clove Cut Dry & Packing";
                else if (lphtype == "LPHPrimaryRTC")
                    lphtype = "Intermediate Line - RTC";
                else if (lphtype == "LPHPrimaryKitchen")
                    lphtype = "Intermediate Line - Casing Kitchen";
                else if (lphtype == "LPHPrimaryWhiteLineOTP")
                    lphtype = "White Line OTP - Process Note";
                else if (lphtype == "LPHPrimaryKretekLineFeeding")
                    lphtype = "Kretek Line - Feeding KR & RJ";
                else if (lphtype == "LPHPrimaryKretekLineConditioning")
                    lphtype = "Kretek Line - DCCC KR & RJ";
                else if (lphtype == "LPHPrimaryKretekLineCuttingDrying")
                    lphtype = "Kretek Line - Cut Dry";
                else if (lphtype == "LPHPrimaryKretekLinePacking")
                    lphtype = "Kretek Line - Packing";
                else if (lphtype == "LPHPrimaryCresFeedingConditioning")
                    lphtype = "Kretek Line - CRES Feeding & DCCC";
                else if (lphtype == "LPHPrimaryCresDryingPacking")
                    lphtype = "Kretek Line - CRES Cut Dry & Packing";
                else if (lphtype == "LPHPrimaryWhiteLineFeedingWhite")
                    lphtype = "White Line PMID - Feeding White";
                else if (lphtype == "LPHPrimaryWhiteLineDCCC")
                    lphtype = "White Line PMID - DCCC";
                else if (lphtype == "LPHPrimaryWhiteLineCuttingFTD")
                    lphtype = "White Line PMID - Cutting + FTD";
                else if (lphtype == "LPHPrimaryWhiteLineAddback")
                    lphtype = "White Line PMID - Addback";
                else if (lphtype == "LPHPrimaryWhiteLinePackingWhite")
                    lphtype = "White Line PMID - Packing White";
                else if (lphtype == "LPHPrimaryWhiteLineFeedingSPM")
                    lphtype = "White Line PMID - Feeding SPM";
                else if (lphtype == "LPHPrimaryISWhiteFeeding")
                    lphtype = "White Line PMID - Feeding IS White";
                else if (lphtype == "LPHPrimaryISWhiteCutDry")
                    lphtype = "White Line PMID - Cut Dry IS White";

                Sheet.Cells[string.Format("A{0}", row)].Value = lphtype;
                Sheet.Cells[string.Format("B{0}", row)].Value = item.LPHSubmissionID;
                Sheet.Cells[string.Format("C{0}", row)].Value = item.User;
                Sheet.Cells[string.Format("D{0}", row)].Value = item.Date.ToString("dd-MMM-yy");
                Sheet.Cells[string.Format("E{0}", row)].Value = item.Shift;
                Sheet.Cells[string.Format("F{0}", row)].Value = item.Status;
                Sheet.Cells[string.Format("G{0}", row)].Value = item.ModifiedDate.HasValue ? item.ModifiedDate.Value.ToString("dd-MMM-yy HH:mm") : string.Empty;

                Sheet.Cells[string.Format("A{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("B{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("C{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("D{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("E{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("F{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("G{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                row++;
            }

            Sheet.Cells["A:G"].AutoFitColumns();

            return Ep.GetAsByteArray();
        }
        #endregion

        #region ::Machine Allocation::
        public static byte[] ExportMachineAllocation(List<MachineAllocationModel> dataList, string username)
        {
            ExcelPackage Ep = new ExcelPackage();
            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Planning-MPP-Dashboard");
            AddHeader(Sheet, "Machine Allocation Management", username);

            Color colFromHex = ColorTranslator.FromHtml(HEADER_COLOR);
            Sheet.Cells["A6:D6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Sheet.Cells["A6:D6"].Style.Fill.BackgroundColor.SetColor(colFromHex);
            Sheet.Cells["A6"].Value = UIResources.MachineID;
            Sheet.Cells["B6"].Value = UIResources.Code;
            Sheet.Cells["C6"].Value = UIResources.Category;
            Sheet.Cells["D6"].Value = UIResources.Value;

            Sheet.Cells["A6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["B6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["C6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["D6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            int row = 7;
            foreach (var item in dataList)
            {
                Sheet.Cells[string.Format("A{0}", row)].Value = item.ID;
                Sheet.Cells[string.Format("B{0}", row)].Value = item.MachineCode;
                Sheet.Cells[string.Format("C{0}", row)].Value = item.MachineCategory;
                Sheet.Cells[string.Format("D{0}", row)].Value = item.Value;

                Sheet.Cells[string.Format("A{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("B{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("C{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("D{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                row++;
            }

            Sheet.Cells["A:D"].AutoFitColumns();

            return Ep.GetAsByteArray();
        }
        #endregion

        #region ::Wpp SP:
        public static byte[] ExportWppSP(List<WppStpModel> dataList, DateTime startDate, DateTime endDate)
        {
            ExcelPackage Ep = new ExcelPackage();
            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Planning-Wpp-SP");
            Color colFromHex = ColorTranslator.FromHtml(HEADER_COLOR);

            Sheet.Cells["A2"].Value = "Brand/Blend";
            Sheet.Cells["B2"].Value = "Description";
            Sheet.Cells["C2"].Value = "Machine 1";
            Sheet.Cells["D2"].Value = "Machine 2";
            Sheet.Cells["E2"].Value = "Activity";
            Sheet.Cells["F2"].Value = "NoPO";
            Sheet.Cells["G2"].Value = "NoOPS";
            Sheet.Cells["H2"].Value = "BatchSAP";
            Sheet.Cells["I2"].Value = "Market";

            Sheet.Cells["A1:A2"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["B1:B2"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["C1:C2"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["D1:D2"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["E1:E2"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["F1:F2"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["G1:G2"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["H1:H2"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["I1:I2"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            int dateIndex = 10;
            string columnStartName = ExcelCellAddress.GetColumnLetter(dateIndex);
            for (var day = startDate.Date; day.Date <= endDate.Date; day = day.AddDays(1))
            {
                string columnName1 = ExcelCellAddress.GetColumnLetter(dateIndex++);
                string columnName2 = ExcelCellAddress.GetColumnLetter(dateIndex++);
                string columnName3 = ExcelCellAddress.GetColumnLetter(dateIndex++);

                Sheet.Cells[columnName1 + "1:" + columnName1 + "2"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[columnName2 + "1:" + columnName2 + "2"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[columnName3 + "1:" + columnName3 + "2"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                Sheet.Cells[columnName1 + "1"].Value = day.ToString("yyyyMMdd");
                Sheet.Cells[columnName2 + "1"].Value = day.ToString("yyyyMMdd");
                Sheet.Cells[columnName3 + "1"].Value = day.ToString("yyyyMMdd");
                Sheet.Cells[columnName1 + "2"].Value = "1";
                Sheet.Cells[columnName2 + "2"].Value = "2";
                Sheet.Cells[columnName3 + "2"].Value = "3";

                Sheet.Cells[columnName1 + "1:" + columnName3 + "2"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            }
            string columnEndName = ExcelCellAddress.GetColumnLetter(dateIndex);

            Sheet.Cells[columnStartName + "2:" + columnEndName + "2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            var brandBlendList = dataList.Select(m => new { m.Brand, m.Maker, m.Packer }).Distinct().ToList();

            int row = 3;
            foreach (var item in brandBlendList)
            {
                var tempList = dataList.Where(x => x.Brand == item.Brand && x.Maker == item.Maker && x.Packer == item.Packer).OrderBy(x => x.Date).ToList();

                Sheet.Cells[string.Format("A{0}", row)].Value = tempList[0].Brand;
                Sheet.Cells[string.Format("B{0}", row)].Value = tempList[0].Description;
                Sheet.Cells[string.Format("C{0}", row)].Value = tempList[0].Maker;
                Sheet.Cells[string.Format("D{0}", row)].Value = tempList[0].Packer;
                Sheet.Cells[string.Format("E{0}", row)].Value = tempList[0].Activity;
                Sheet.Cells[string.Format("F{0}", row)].Value = tempList[0].PONumber;
                Sheet.Cells[string.Format("G{0}", row)].Value = tempList[0].OPSNumber;
                Sheet.Cells[string.Format("H{0}", row)].Value = tempList[0].BatchSAP;
                Sheet.Cells[string.Format("I{0}", row)].Value = tempList[0].Others;

                dateIndex = 10;
                for (var day = startDate.Date; day.Date <= endDate.Date; day = day.AddDays(1))
                {
                    var temp = tempList.Where(x => x.Date == day).FirstOrDefault();
                    string columnName1 = ExcelCellAddress.GetColumnLetter(dateIndex++);
                    string columnName2 = ExcelCellAddress.GetColumnLetter(dateIndex++);
                    string columnName3 = ExcelCellAddress.GetColumnLetter(dateIndex++);
                    Sheet.Cells[string.Format(columnName1 + "{0}", row)].Value = temp == null ? string.Empty : string.Format("{0:0.00}", temp.Shift1);
                    Sheet.Cells[string.Format(columnName2 + "{0}", row)].Value = temp == null ? string.Empty : string.Format("{0:0.00}", temp.Shift2);
                    Sheet.Cells[string.Format(columnName3 + "{0}", row)].Value = temp == null ? string.Empty : string.Format("{0:0.00}", temp.Shift3);
                    Sheet.Cells[string.Format(columnName1 + "{0}", row)].Style.Numberformat.Format = "#,##0.00";
                    Sheet.Cells[string.Format(columnName2 + "{0}", row)].Style.Numberformat.Format = "#,##0.00";
                    Sheet.Cells[string.Format(columnName3 + "{0}", row)].Style.Numberformat.Format = "#,##0.00";
                }

                row++;
            }

            Sheet.Cells["A1" + ":" + columnEndName + "2"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Sheet.Cells["A1" + ":" + columnEndName + "2"].Style.Fill.BackgroundColor.SetColor(colFromHex);

            return Ep.GetAsByteArray();
        }
        #endregion

        #region ::Wpp SP View:
        public static byte[] ExportWppSPView(List<WppStpModel> dataList, DateTime startDate, DateTime endDate, string username)
        {
            ExcelPackage Ep = new ExcelPackage();
            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Planning-Wpp-SP");
            AddHeader(Sheet, "Wpp SP Report", username);

            Color colFromHex = ColorTranslator.FromHtml(HEADER_COLOR);
            Sheet.Cells["A6:K6"].Style.Font.Bold = true;
            Sheet.Cells["A6:K6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Sheet.Cells["A6:K6"].Style.Fill.BackgroundColor.SetColor(colFromHex);

            Sheet.Cells["A6"].Value = "Date";
            Sheet.Cells["B6"].Value = "Location";
            Sheet.Cells["C6"].Value = "Brand/Blend";
            Sheet.Cells["D6"].Value = "Description";
            Sheet.Cells["E6"].Value = "Machine 1";
            Sheet.Cells["F6"].Value = "Machine 2";
            Sheet.Cells["G6"].Value = "Shift 1";
            Sheet.Cells["H6"].Value = "Shift 2";
            Sheet.Cells["I6"].Value = "Shift 3";
            Sheet.Cells["J6"].Value = "PO Number";
            Sheet.Cells["K6"].Value = "Activity";

            int row = 7;
            foreach (var item in dataList)
            {
                Sheet.Cells[string.Format("A{0}", row)].Value = item.Date.ToString("dd-MMM-yy");
                Sheet.Cells[string.Format("B{0}", row)].Value = item.Location;
                Sheet.Cells[string.Format("C{0}", row)].Value = item.Brand;
                Sheet.Cells[string.Format("D{0}", row)].Value = item.Description;
                Sheet.Cells[string.Format("E{0}", row)].Value = item.Maker;
                Sheet.Cells[string.Format("F{0}", row)].Value = item.Packer;
                Sheet.Cells[string.Format("G{0}", row)].Value = item.Shift1;
                Sheet.Cells[string.Format("H{0}", row)].Value = item.Shift2;
                Sheet.Cells[string.Format("I{0}", row)].Value = item.Shift3;
                Sheet.Cells[string.Format("J{0}", row)].Value = item.PONumber;
                Sheet.Cells[string.Format("K{0}", row)].Value = item.Activity;

                Sheet.Cells[string.Format("A{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("B{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("C{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("D{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("E{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("F{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("G{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("H{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("I{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("J{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("K{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                row++;
            }

            Sheet.Cells["A:K"].AutoFitColumns();

            return Ep.GetAsByteArray();
        }
        #endregion

        #region ::Wpp Simulation:
        public static byte[] ExportWppSimulation(List<WppPrimaryModel> dataList, string username, DateTime startDate, DateTime endDate)
        {
            ExcelPackage Ep = new ExcelPackage();
            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Planning-Wpp-Simulation");
            AddHeaderSim(Sheet, UIResources.WppPrimaryManagement, username, startDate, endDate);

            Color colFromHex = ColorTranslator.FromHtml(HEADER_COLOR);
            Sheet.Cells["A7:J7"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Sheet.Cells["A7:J7"].Style.Font.Bold = true;
            Sheet.Cells["A7:J7"].Style.Fill.BackgroundColor.SetColor(colFromHex);
            Sheet.Cells["A7"].Value = UIResources.Blend;
            Sheet.Cells["B7"].Value = UIResources.VolPerOps;
            Sheet.Cells["C7"].Value = UIResources.Monday;
            Sheet.Cells["D7"].Value = UIResources.Tuesday;
            Sheet.Cells["E7"].Value = UIResources.Wednesday;
            Sheet.Cells["F7"].Value = UIResources.Thursday;
            Sheet.Cells["G7"].Value = UIResources.Friday;
            Sheet.Cells["H7"].Value = UIResources.Saturday;
            Sheet.Cells["I7"].Value = UIResources.Sunday;
            Sheet.Cells["J7"].Value = UIResources.Total;
            Sheet.Cells["B7"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#000fff"));
            Sheet.Cells["J7"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#ff0000"));

            int row = 8;
            foreach (var data in dataList)
            {
                Sheet.Cells[string.Format("A{0}", row)].Value = data.Blend;
                Sheet.Cells[string.Format("B{0}", row)].Value = data.VolPerOps;
                Sheet.Cells[string.Format("C{0}", row)].Value = data.Monday;
                Sheet.Cells[string.Format("D{0}", row)].Value = data.Tuesday;
                Sheet.Cells[string.Format("E{0}", row)].Value = data.Wednesday;
                Sheet.Cells[string.Format("F{0}", row)].Value = data.Thursday;
                Sheet.Cells[string.Format("G{0}", row)].Value = data.Friday;
                Sheet.Cells[string.Format("H{0}", row)].Value = data.Saturday;
                Sheet.Cells[string.Format("I{0}", row)].Value = data.Sunday;
                Sheet.Cells[string.Format("J{0}", row)].Value = data.Total;

                Sheet.Cells[string.Format("C{0}", row)].Style.Numberformat.Format = "#,##0";
                Sheet.Cells[string.Format("D{0}", row)].Style.Numberformat.Format = "#,##0";
                Sheet.Cells[string.Format("E{0}", row)].Style.Numberformat.Format = "#,##0";
                Sheet.Cells[string.Format("F{0}", row)].Style.Numberformat.Format = "#,##0";
                Sheet.Cells[string.Format("G{0}", row)].Style.Numberformat.Format = "#,##0";
                Sheet.Cells[string.Format("H{0}", row)].Style.Numberformat.Format = "#,##0";
                Sheet.Cells[string.Format("I{0}", row)].Style.Numberformat.Format = "#,##0";
                Sheet.Cells[string.Format("J{0}", row)].Style.Numberformat.Format = "#,##0";

                Sheet.Cells[string.Format("A{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("B{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("B{0}", row)].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#000fff"));
                Sheet.Cells[string.Format("C{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("D{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("E{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("F{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("G{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("H{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("I{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("J{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("J{0}", row)].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#ff0000"));

                if (row == dataList.Count + 7)
                {
                    Sheet.Cells[string.Format("A{0}", row)].Style.Font.Bold = true;
                    Sheet.Cells[string.Format("B{0}", row)].Style.Font.Bold = true;
                    Sheet.Cells[string.Format("C{0}", row)].Style.Font.Bold = true;
                    Sheet.Cells[string.Format("D{0}", row)].Style.Font.Bold = true;
                    Sheet.Cells[string.Format("E{0}", row)].Style.Font.Bold = true;
                    Sheet.Cells[string.Format("F{0}", row)].Style.Font.Bold = true;
                    Sheet.Cells[string.Format("G{0}", row)].Style.Font.Bold = true;
                    Sheet.Cells[string.Format("H{0}", row)].Style.Font.Bold = true;
                    Sheet.Cells[string.Format("I{0}", row)].Style.Font.Bold = true;
                    Sheet.Cells[string.Format("J{0}", row)].Style.Font.Bold = true;
                }

                row++;
            }

            Sheet.Cells["A:J"].AutoFitColumns();

            return Ep.GetAsByteArray();
        }
        #endregion

        #region ::Wpp OPS:
        public static byte[] ExportWppOps(List<WppPrimModel> dataList, DateTime startDate)
        {
            ExcelPackage Ep = new ExcelPackage();
            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Planning-Wpp-OPS");
            Color colFromHex = ColorTranslator.FromHtml("#ffff99");

            DateTime endDate = startDate.AddDays(6);
            int index = 1;
            string dateCell = string.Empty;
            string blendCell = string.Empty;
            string batchLamaCell = string.Empty;
            string POCell = string.Empty;
            string batchSAPCell = string.Empty;

            for (var day = startDate.Date; day.Date <= endDate.Date; day = day.AddDays(1))
            {
                #region Get Cell Index
                if (index == 1)
                {
                    dateCell = "D";
                    blendCell = "B";
                    batchLamaCell = "C";
                    POCell = "D";
                    batchSAPCell = "E";
                }
                else if (index == 2)
                {
                    dateCell = "I";
                    blendCell = "G";
                    batchLamaCell = "H";
                    POCell = "I";
                    batchSAPCell = "J";
                }
                else if (index == 3)
                {
                    dateCell = "N";
                    blendCell = "L";
                    batchLamaCell = "M";
                    POCell = "N";
                    batchSAPCell = "O";
                }
                else if (index == 4)
                {
                    dateCell = "S";
                    blendCell = "Q";
                    batchLamaCell = "R";
                    POCell = "S";
                    batchSAPCell = "T";
                }
                else if (index == 5)
                {
                    dateCell = "X";
                    blendCell = "V";
                    batchLamaCell = "W";
                    POCell = "X";
                    batchSAPCell = "Y";
                }
                else if (index == 6)
                {
                    dateCell = "AC";
                    blendCell = "AA";
                    batchLamaCell = "AB";
                    POCell = "AC";
                    batchSAPCell = "AD";
                }
                else
                {
                    dateCell = "AH";
                    blendCell = "AF";
                    batchLamaCell = "AG";
                    POCell = "AH";
                    batchSAPCell = "AI";
                }
                #endregion

                Sheet.Cells[string.Format("{0}1:{1}2", blendCell, batchSAPCell)].Style.Fill.PatternType = ExcelFillStyle.Solid;
                Sheet.Cells[string.Format("{0}1:{1}2", blendCell, batchSAPCell)].Style.Font.Bold = true;
                Sheet.Cells[string.Format("{0}1:{1}2", blendCell, batchSAPCell)].Style.Fill.BackgroundColor.SetColor(colFromHex);
                Sheet.Cells[string.Format("{0}1:{1}2", blendCell, batchSAPCell)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                Sheet.Cells[string.Format("{0}1", dateCell)].Value = day.ToString("yyyyMMdd");
                Sheet.Cells[string.Format("{0}2", blendCell)].Value = UIResources.Blend;
                Sheet.Cells[string.Format("{0}2", batchLamaCell)].Value = UIResources.OldBatch;
                Sheet.Cells[string.Format("{0}2", POCell)].Value = "PO";
                Sheet.Cells[string.Format("{0}2", batchSAPCell)].Value = UIResources.BatchSAP;
                Sheet.Cells[string.Format("{0}2", blendCell)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("{0}2", batchLamaCell)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("{0}2", POCell)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("{0}2", batchSAPCell)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                int row = 3;
                List<WppPrimModel> dataListTemp = dataList.Where(x => x.Date == day).ToList();
                foreach (var data in dataListTemp)
                {
                    Sheet.Cells[string.Format("{0}{1}", blendCell, row)].Value = data.Blend;
                    Sheet.Cells[string.Format("{0}{1}", batchLamaCell, row)].Value = data.BatchLama;
                    Sheet.Cells[string.Format("{0}{1}", POCell, row)].Value = data.PONumber;
                    Sheet.Cells[string.Format("{0}{1}", batchSAPCell, row)].Value = data.BatchSAP;

                    Sheet.Cells[string.Format("{0}{1}", blendCell, row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    Sheet.Cells[string.Format("{0}{1}", batchLamaCell, row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    Sheet.Cells[string.Format("{0}{1}", POCell, row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    Sheet.Cells[string.Format("{0}{1}", batchSAPCell, row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    Sheet.Cells[string.Format("{0}{1}", blendCell, row)].Style.Font.Bold = true;
                    Sheet.Cells[string.Format("{0}{1}", batchLamaCell, row)].Style.Font.Bold = true;
                    Sheet.Cells[string.Format("{0}{1}", POCell, row)].Style.Font.Bold = true;
                    Sheet.Cells[string.Format("{0}{1}", batchSAPCell, row)].Style.Font.Bold = true;

                    row++;
                }

                Sheet.Cells[string.Format("{0}:{1}", blendCell, batchSAPCell)].AutoFitColumns();

                index++;
            }

            return Ep.GetAsByteArray();
        }
        #endregion

        #region ::Shuttle Request:
        public static byte[] ExportShuttleRequest(List<ShuttleRequestModel> dataList, string username)
        {
            ExcelPackage Ep = new ExcelPackage();
            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Facility-Shuttle-Request");
            AddHeader(Sheet, UIResources.ShuttleRequestManagement, username);

            Color colFromHex = ColorTranslator.FromHtml(HEADER_COLOR);
            Sheet.Cells["A6:L6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Sheet.Cells["A6:L6"].Style.Font.Bold = true;
            Sheet.Cells["A6:L6"].Style.Fill.BackgroundColor.SetColor(colFromHex);
            Sheet.Cells["A6"].Value = UIResources.CostCenter;
            Sheet.Cells["B6"].Value = UIResources.Period;
            Sheet.Cells["C6"].Value = UIResources.Time;
            Sheet.Cells["D6"].Value = UIResources.ProductionCenter;
            Sheet.Cells["E6"].Value = UIResources.Department;
            Sheet.Cells["F6"].Value = UIResources.HmsPIC;
            Sheet.Cells["G6"].Value = UIResources.Phone;
            Sheet.Cells["H6"].Value = UIResources.TotalPassengers;
            Sheet.Cells["I6"].Value = UIResources.GuestType;
            Sheet.Cells["J6"].Value = UIResources.LocationFrom;
            Sheet.Cells["K6"].Value = UIResources.LocationTo;
            Sheet.Cells["L6"].Value = UIResources.Purpose;

            int row = 7;
            foreach (var data in dataList)
            {
                Sheet.Cells[string.Format("A{0}", row)].Value = data.CostCenter;
                Sheet.Cells[string.Format("B{0}", row)].Value = "(" + data.StartDate.ToString("dd-MMM-yy") + ") - (" + data.EndDate.ToString("dd-MMM-yy") + ")";
                Sheet.Cells[string.Format("C{0}", row)].Value = data.Time.ToString(@"hh\:mm");
                Sheet.Cells[string.Format("D{0}", row)].Value = data.ProductionCenter;
                Sheet.Cells[string.Format("E{0}", row)].Value = data.Department;
                Sheet.Cells[string.Format("F{0}", row)].Value = data.EmployeeFullname;
                Sheet.Cells[string.Format("G{0}", row)].Value = data.Phone;
                Sheet.Cells[string.Format("H{0}", row)].Value = data.TotalPassengers;
                Sheet.Cells[string.Format("I{0}", row)].Value = data.GuestType;
                Sheet.Cells[string.Format("J{0}", row)].Value = data.LocationFrom;
                Sheet.Cells[string.Format("K{0}", row)].Value = data.LocationTo;
                Sheet.Cells[string.Format("L{0}", row)].Value = data.Purpose;

                Sheet.Cells[string.Format("A{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("B{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("C{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("D{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("E{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("F{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("G{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("H{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("I{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("J{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("K{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("L{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                row++;
            }

            Sheet.Cells["A:L"].AutoFitColumns();

            return Ep.GetAsByteArray();
        }
        #endregion

        #region ::Shuttle Request Detail Report:
        public static byte[] ExportShuttleRequestDetailReport(List<ShuttleRequestModel> dataList, string username)
        {
            ExcelPackage Ep = new ExcelPackage();
            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Detail-Report");
            AddHeader(Sheet, UIResources.ShuttleRequestManagement, username);

            Color colFromHex = ColorTranslator.FromHtml(HEADER_COLOR);
            Sheet.Cells["A6:F6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Sheet.Cells["A6:F6"].Style.Font.Bold = true;
            Sheet.Cells["A6:F6"].Style.Fill.BackgroundColor.SetColor(colFromHex);
            Sheet.Cells["A6"].Value = UIResources.Date;
            Sheet.Cells["B6"].Value = UIResources.Time;
            Sheet.Cells["C6"].Value = UIResources.LocationFrom;
            Sheet.Cells["D6"].Value = UIResources.LocationTo;
            Sheet.Cells["E6"].Value = UIResources.Qty;
            Sheet.Cells["F6"].Value = UIResources.Name;

            int row = 7;
            foreach (var data in dataList)
            {
                Sheet.Cells[string.Format("A{0}", row)].Value = data.Date.ToString("dd-MMM-yy");
                Sheet.Cells[string.Format("B{0}", row)].Value = data.Time.ToString(@"hh\:mm");
                Sheet.Cells[string.Format("C{0}", row)].Value = data.LocationFrom;
                Sheet.Cells[string.Format("D{0}", row)].Value = data.LocationTo;
                Sheet.Cells[string.Format("E{0}", row)].Value = data.TotalPassengers;
                Sheet.Cells[string.Format("F{0}", row)].Value = data.EmployeeFullname;

                Sheet.Cells[string.Format("A{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("B{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("C{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("D{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("E{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("F{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                row++;
            }

            Sheet.Cells["A:F"].AutoFitColumns();

            return Ep.GetAsByteArray();
        }
        #endregion

        #region ::Shuttle Request Summary Report:
        public static byte[] ExportShuttleRequestSummaryReport(List<ShuttleRequestModel> dataList, string username)
        {
            ExcelPackage Ep = new ExcelPackage();
            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Summary-Report");
            AddHeader(Sheet, UIResources.ShuttleRequestManagement, username);

            Color colFromHex = ColorTranslator.FromHtml(HEADER_COLOR);
            Sheet.Cells["A6:E6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Sheet.Cells["A6:E6"].Style.Font.Bold = true;
            Sheet.Cells["A6:E6"].Style.Fill.BackgroundColor.SetColor(colFromHex);
            Sheet.Cells["A6"].Value = UIResources.Date;
            Sheet.Cells["B6"].Value = UIResources.Time;
            Sheet.Cells["C6"].Value = UIResources.LocationFrom;
            Sheet.Cells["D6"].Value = UIResources.LocationTo;
            Sheet.Cells["E6"].Value = UIResources.Qty;

            int row = 7;
            foreach (var data in dataList)
            {
                Sheet.Cells[string.Format("A{0}", row)].Value = data.Date.ToString("dd-MMM-yy");
                Sheet.Cells[string.Format("B{0}", row)].Value = data.Time.ToString(@"hh\:mm");
                Sheet.Cells[string.Format("C{0}", row)].Value = data.LocationFrom;
                Sheet.Cells[string.Format("D{0}", row)].Value = data.LocationTo;
                Sheet.Cells[string.Format("E{0}", row)].Value = data.TotalPassengers;

                Sheet.Cells[string.Format("A{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("B{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("C{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("D{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("E{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                row++;
            }

            Sheet.Cells["A:E"].AutoFitColumns();

            return Ep.GetAsByteArray();
        }
        #endregion

        #region ::Meal Request:
        public static byte[] ExportMealRequest(List<MealRequestModel> dataList, string username)
        {
            ExcelPackage Ep = new ExcelPackage();
            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Facility-Meal-Request");
            AddHeader(Sheet, UIResources.MealRequestManagement, username);

            Color colFromHex = ColorTranslator.FromHtml(HEADER_COLOR);
            Sheet.Cells["A6:M6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Sheet.Cells["A6:M6"].Style.Font.Bold = true;
            Sheet.Cells["A6:M6"].Style.Fill.BackgroundColor.SetColor(colFromHex);
            Sheet.Cells["A6"].Value = UIResources.CostCenter;
            Sheet.Cells["B6"].Value = UIResources.Period;
            Sheet.Cells["C6"].Value = UIResources.ProductionCenter;
            Sheet.Cells["D6"].Value = UIResources.Canteen;
            Sheet.Cells["E6"].Value = UIResources.HmsPIC;
            Sheet.Cells["F6"].Value = UIResources.Phone;
            Sheet.Cells["G6"].Value = UIResources.TotalGuest;
            Sheet.Cells["H6"].Value = UIResources.GuestType;
            Sheet.Cells["I6"].Value = UIResources.Company;
            Sheet.Cells["J6"].Value = UIResources.Guest;
            Sheet.Cells["K6"].Value = UIResources.Purpose;
            Sheet.Cells["L6"].Value = UIResources.Department;
            Sheet.Cells["M6"].Value = UIResources.Shift;

            int row = 7;
            foreach (var data in dataList)
            {
                Sheet.Cells[string.Format("A{0}", row)].Value = data.CostCenter;
                Sheet.Cells[string.Format("B{0}", row)].Value = "(" + data.StartDate.ToString("dd-MMM-yy") + ") - (" + data.EndDate.ToString("dd-MMM-yy") + ")";
                Sheet.Cells[string.Format("C{0}", row)].Value = data.ProductionCenter;
                Sheet.Cells[string.Format("D{0}", row)].Value = data.Canteen;
                Sheet.Cells[string.Format("E{0}", row)].Value = data.EmployeeFullname;
                Sheet.Cells[string.Format("F{0}", row)].Value = data.Phone;
                Sheet.Cells[string.Format("G{0}", row)].Value = data.TotalGuest;
                Sheet.Cells[string.Format("H{0}", row)].Value = data.GuestType;
                Sheet.Cells[string.Format("I{0}", row)].Value = data.Company;
                Sheet.Cells[string.Format("J{0}", row)].Value = data.Guest;
                Sheet.Cells[string.Format("K{0}", row)].Value = data.Purpose;
                Sheet.Cells[string.Format("L{0}", row)].Value = data.Department;
                Sheet.Cells[string.Format("M{0}", row)].Value = data.Shift;


                Sheet.Cells[string.Format("A{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("B{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("C{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("D{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("E{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("F{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("G{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("H{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("I{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("J{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("K{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("L{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("M{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                row++;
            }

            Sheet.Cells["A:M"].AutoFitColumns();

            return Ep.GetAsByteArray();
        }
        #endregion

        #region ::Meal Request Detail Report:
        public static byte[] ExportMealRequestDetailReport(List<MealRequestModel> dataList, string username)
        {
            ExcelPackage Ep = new ExcelPackage();
            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Detail-Report");
            AddHeader(Sheet, UIResources.MealRequestManagement, username);

            Color colFromHex = ColorTranslator.FromHtml(HEADER_COLOR);
            Sheet.Cells["A6:I6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Sheet.Cells["A6:I6"].Style.Font.Bold = true;
            Sheet.Cells["A6:I6"].Style.Fill.BackgroundColor.SetColor(colFromHex);
            Sheet.Cells["A6"].Value = UIResources.ProductionCenter;
            Sheet.Cells["B6"].Value = UIResources.Canteen;
            Sheet.Cells["C6"].Value = UIResources.CostCenter;
            Sheet.Cells["D6"].Value = UIResources.Name;
            Sheet.Cells["E6"].Value = UIResources.Date;
            Sheet.Cells["F6"].Value = UIResources.Shift;
            Sheet.Cells["G6"].Value = UIResources.Qty;
            Sheet.Cells["H6"].Value = UIResources.Type;
            Sheet.Cells["I6"].Value = UIResources.PIC;

            int row = 7;
            foreach (var data in dataList)
            {
                Sheet.Cells[string.Format("A{0}", row)].Value = data.ProductionCenter;
                Sheet.Cells[string.Format("B{0}", row)].Value = data.Canteen;
                Sheet.Cells[string.Format("C{0}", row)].Value = data.CostCenter;
                Sheet.Cells[string.Format("D{0}", row)].Value = data.Guest;
                Sheet.Cells[string.Format("E{0}", row)].Value = data.Date.ToString("dd-MMM-yy");
                Sheet.Cells[string.Format("F{0}", row)].Value = data.Shift;
                Sheet.Cells[string.Format("G{0}", row)].Value = data.TotalGuest;
                Sheet.Cells[string.Format("H{0}", row)].Value = data.GuestType;
                Sheet.Cells[string.Format("I{0}", row)].Value = data.PIC;

                Sheet.Cells[string.Format("A{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("B{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("C{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("D{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("E{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("F{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("G{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("H{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("I{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                row++;
            }

            Sheet.Cells["A:I"].AutoFitColumns();

            return Ep.GetAsByteArray();
        }
        #endregion

        #region ::Meal Request Summary Report:
        public static byte[] ExportMealRequestSummaryReport(List<MealRequestModel> dataList, string username)
        {
            ExcelPackage Ep = new ExcelPackage();
            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Summary-Report");
            AddHeader(Sheet, UIResources.MealRequestManagement, username);

            Color colFromHex = ColorTranslator.FromHtml(HEADER_COLOR);
            Sheet.Cells["A6:H6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Sheet.Cells["A6:H6"].Style.Font.Bold = true;
            Sheet.Cells["A6:H6"].Style.Fill.BackgroundColor.SetColor(colFromHex);
            Sheet.Cells["A6"].Value = UIResources.ProductionCenter;
            Sheet.Cells["B6"].Value = UIResources.Canteen;
            Sheet.Cells["C6"].Value = UIResources.CostCenter;
            Sheet.Cells["D6"].Value = UIResources.Date;
            Sheet.Cells["E6"].Value = UIResources.Shift1;
            Sheet.Cells["F6"].Value = UIResources.Shift2;
            Sheet.Cells["G6"].Value = UIResources.Shift3;
            Sheet.Cells["H6"].Value = "NS";

            int row = 7;
            foreach (var data in dataList)
            {
                Sheet.Cells[string.Format("A{0}", row)].Value = data.ProductionCenter;
                Sheet.Cells[string.Format("B{0}", row)].Value = data.Canteen;
                Sheet.Cells[string.Format("C{0}", row)].Value = data.CostCenter;
                Sheet.Cells[string.Format("D{0}", row)].Value = data.Date.ToString("dd-MMM-yy");
                Sheet.Cells[string.Format("E{0}", row)].Value = data.Shift1;
                Sheet.Cells[string.Format("F{0}", row)].Value = data.Shift2;
                Sheet.Cells[string.Format("G{0}", row)].Value = data.Shift3;
                Sheet.Cells[string.Format("H{0}", row)].Value = data.NS;

                Sheet.Cells[string.Format("A{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("B{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("C{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("D{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("E{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("F{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("G{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("H{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                row++;
            }

            Sheet.Cells["A:H"].AutoFitColumns();

            return Ep.GetAsByteArray();
        }
        #endregion

        #region ::Training Title Machine Type::
        public static byte[] ExportTrainingTitleMachineType(List<TrainingTitleMachineTypeModel> dataList, string username)
        {
            ExcelPackage Ep = new ExcelPackage();
            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Master-Training-Title-Machine-Type");
            AddHeader(Sheet, "Training Title Machine Type Management", username);

            Color colFromHex = ColorTranslator.FromHtml(HEADER_COLOR);
            Sheet.Cells["A6:B6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Sheet.Cells["A6:B6"].Style.Fill.BackgroundColor.SetColor(colFromHex);
            Sheet.Cells["A6"].Value = UIResources.TrainingTitle;
            Sheet.Cells["B6"].Value = UIResources.MachineTypeList;

            Sheet.Cells["A6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["B6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            int row = 7;
            foreach (var item in dataList)
            {
                Sheet.Cells[string.Format("A{0}", row)].Value = item.TrainingTitle;
                Sheet.Cells[string.Format("B{0}", row)].Value = item.MachineTypeList;

                Sheet.Cells[string.Format("A{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("B{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                row++;
            }

            Sheet.Cells["A:B"].AutoFitColumns();

            return Ep.GetAsByteArray();
        }
        #endregion

        #region ::Training Title::
        public static byte[] ExportTrainingTitle(List<TrainingTitleModel> dataList, string username)
        {
            ExcelPackage Ep = new ExcelPackage();
            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Master-Training-Title");
            AddHeader(Sheet, "Training Title Management", username);

            Color colFromHex = ColorTranslator.FromHtml(HEADER_COLOR);
            Sheet.Cells["A6:D6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Sheet.Cells["A6:D6"].Style.Fill.BackgroundColor.SetColor(colFromHex);
            Sheet.Cells["A6"].Value = UIResources.ID;
            Sheet.Cells["B6"].Value = UIResources.Title;
            Sheet.Cells["C6"].Value = UIResources.Competency;
            Sheet.Cells["D6"].Value = UIResources.Trainees;
            Sheet.Cells["A6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["B6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["C6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["D6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            int row = 7;
            foreach (var item in dataList)
            {
                Sheet.Cells[string.Format("A{0}", row)].Value = item.ID;
                Sheet.Cells[string.Format("B{0}", row)].Value = item.Title;
                Sheet.Cells[string.Format("C{0}", row)].Value = item.Competency;
                Sheet.Cells[string.Format("D{0}", row)].Value = item.Trainees;

                Sheet.Cells[string.Format("A{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("B{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("C{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("D{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                row++;
            }

            Sheet.Cells["A:D"].AutoFitColumns();

            return Ep.GetAsByteArray();
        }
        #endregion

        #region ::MPP Dashboard::
        public static byte[] ExportMPPDashboard(List<MppSummaryModel> dataList, string username)
        {
            ExcelPackage Ep = new ExcelPackage();
            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Planning-MPP-Dashboard");
            AddHeader(Sheet, "MPP Dashboard", username);

            Color colFromHex = ColorTranslator.FromHtml(HEADER_COLOR);
            Sheet.Cells["A6:E6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Sheet.Cells["A6:E6"].Style.Fill.BackgroundColor.SetColor(colFromHex);
            Sheet.Cells["A6"].Value = UIResources.Date;
            Sheet.Cells["B6"].Value = UIResources.JobTitle;
            Sheet.Cells["C6"].Value = UIResources.Total;
            Sheet.Cells["D6"].Value = UIResources.Assigned;
            Sheet.Cells["E6"].Value = UIResources.Idle;
            Sheet.Cells["A6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["B6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["C6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["D6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["E6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            int row = 7;
            foreach (var item in dataList)
            {
                Sheet.Cells[string.Format("A{0}", row)].Value = item.Date.ToString("dd-MMM-yy");
                Sheet.Cells[string.Format("B{0}", row)].Value = item.JobTitle;
                Sheet.Cells[string.Format("C{0}", row)].Value = item.Total;
                Sheet.Cells[string.Format("D{0}", row)].Value = item.Assigned;
                Sheet.Cells[string.Format("E{0}", row)].Value = item.Idle;

                Sheet.Cells[string.Format("A{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("B{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("C{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("D{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("E{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                row++;
            }

            Sheet.Cells["A:E"].AutoFitColumns();

            return Ep.GetAsByteArray();
        }
        #endregion

        #region ::Man Power::
        public static byte[] ExportManPower(List<ManPowerModel> dataList, string username)
        {
            ExcelPackage Ep = new ExcelPackage();
            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Master-Man-Power");
            AddHeader(Sheet, UIResources.ManPowerManagement, username);

            Color colFromHex = ColorTranslator.FromHtml(HEADER_COLOR);
            Sheet.Cells["A6:C6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Sheet.Cells["A6:C6"].Style.Fill.BackgroundColor.SetColor(colFromHex);
            Sheet.Cells["A6"].Value = UIResources.JobTitle;
            Sheet.Cells["B6"].Value = UIResources.MachineType;
            Sheet.Cells["C6"].Value = UIResources.Value;
            Sheet.Cells["A6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["B6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["C6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            int row = 7;
            foreach (var item in dataList)
            {
                Sheet.Cells[string.Format("A{0}", row)].Value = item.JobTitle;
                Sheet.Cells[string.Format("B{0}", row)].Value = item.Location;
                Sheet.Cells[string.Format("C{0}", row)].Value = item.Value;

                Sheet.Cells[string.Format("A{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("B{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("C{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                row++;
            }

            Sheet.Cells["A:C"].AutoFitColumns();

            return Ep.GetAsByteArray();
        }
        #endregion

        #region ::Brands Conversion::
        public static byte[] ExportBrandConversion(List<BrandConversionModel> dataList, string username)
        {
            ExcelPackage Ep = new ExcelPackage();
            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Master-Brands-Conversion");
            AddHeader(Sheet, UIResources.BrandDataManagement, username);

            Color colFromHex = ColorTranslator.FromHtml(HEADER_COLOR);
            Sheet.Cells["A6:F6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Sheet.Cells["A6:F6"].Style.Fill.BackgroundColor.SetColor(colFromHex);
            Sheet.Cells["A6"].Value = "Brand Code";
            Sheet.Cells["B6"].Value = "Value 1";
            Sheet.Cells["C6"].Value = "UOM 1";
            Sheet.Cells["D6"].Value = "Value 2";
            Sheet.Cells["E6"].Value = "UOM 2";
            Sheet.Cells["F6"].Value = "Notes";
            Sheet.Cells["A6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["B6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["C6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["D6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["E6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["F6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            int row = 7;
            foreach (var item in dataList)
            {
                Sheet.Cells[string.Format("A{0}", row)].Value = item.BrandCode;
                Sheet.Cells[string.Format("B{0}", row)].Value = item.Value1;
                Sheet.Cells[string.Format("C{0}", row)].Value = item.UOM1;
                Sheet.Cells[string.Format("D{0}", row)].Value = item.Value2;
                Sheet.Cells[string.Format("E{0}", row)].Value = item.UOM2;
                Sheet.Cells[string.Format("F{0}", row)].Value = item.Notes;

                Sheet.Cells[string.Format("A{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("B{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("C{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("D{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("E{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("F{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                row++;
            }

            Sheet.Cells["A:F"].AutoFitColumns();

            return Ep.GetAsByteArray();
        }
        #endregion

        #region ::Brands::
        public static byte[] ExportBrands(List<BrandModel> dataList, string username)
        {
            ExcelPackage Ep = new ExcelPackage();
            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Master-Brands");
            AddHeader(Sheet, UIResources.BrandDataManagement, username);

            Color colFromHex = ColorTranslator.FromHtml(HEADER_COLOR);
            Sheet.Cells["A6:K6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Sheet.Cells["A6:K6"].Style.Fill.BackgroundColor.SetColor(colFromHex);

            Sheet.Cells["A6"].Value = "Code";
            Sheet.Cells["B6"].Value = "Description";
            Sheet.Cells["C6"].Value = "Location";
            Sheet.Cells["D6"].Value = "Berat Cigarette";
            Sheet.Cells["E6"].Value = "Pack To Stick";
            Sheet.Cells["F6"].Value = "Slof To Pack";
            Sheet.Cells["G6"].Value = "Box To Slof";
            Sheet.Cells["H6"].Value = "CTW";
            Sheet.Cells["I6"].Value = "CTF";
            Sheet.Cells["J6"].Value = "RS Code";
            Sheet.Cells["K6"].Value = "IsActive";

            Sheet.Cells["A6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["B6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["C6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["D6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["E6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["F6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["G6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["H6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["I6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["J6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["K6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            int row = 7;
            foreach (var item in dataList)
            {
                Sheet.Cells[string.Format("A{0}", row)].Value = item.Code;
                Sheet.Cells[string.Format("B{0}", row)].Value = item.Description;
                Sheet.Cells[string.Format("C{0}", row)].Value = item.Location;
                Sheet.Cells[string.Format("D{0}", row)].Value = item.BeratCigarette;
                Sheet.Cells[string.Format("E{0}", row)].Value = item.PackToStick;
                Sheet.Cells[string.Format("F{0}", row)].Value = item.SlofToPack;
                Sheet.Cells[string.Format("G{0}", row)].Value = item.BoxToSlof;
                Sheet.Cells[string.Format("H{0}", row)].Value = item.CTW;
                Sheet.Cells[string.Format("I{0}", row)].Value = item.CTF;
                Sheet.Cells[string.Format("J{0}", row)].Value = item.RSCode;
                Sheet.Cells[string.Format("K{0}", row)].Value = item.IsActive ? "Y" : "N";

                Sheet.Cells[string.Format("A{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("B{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("C{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("D{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("E{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("F{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("G{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("H{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("I{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("J{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("K{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                row++;
            }

            Sheet.Cells["A:K"].AutoFitColumns();

            return Ep.GetAsByteArray();
        }
        #endregion

        #region ::Blends::
        public static byte[] ExportBlends(List<BlendModel> dataList, string username)
        {
            ExcelPackage Ep = new ExcelPackage();
            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Master-Blends");
            AddHeader(Sheet, UIResources.BlendDataManagement, username);

            Color colFromHex = ColorTranslator.FromHtml(HEADER_COLOR);
            Sheet.Cells["A6:E6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Sheet.Cells["A6:E6"].Style.Fill.BackgroundColor.SetColor(colFromHex);
            Sheet.Cells["A6"].Value = "Code";
            Sheet.Cells["B6"].Value = "Description";
            Sheet.Cells["C6"].Value = "Location";
            Sheet.Cells["D6"].Value = "Ops To Kg";
            Sheet.Cells["E6"].Value = "Is Active";
            Sheet.Cells["A6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["B6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["C6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["D6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["E6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            int row = 7;
            foreach (var item in dataList)
            {
                Sheet.Cells[string.Format("A{0}", row)].Value = item.Code;
                Sheet.Cells[string.Format("B{0}", row)].Value = item.Description;
                Sheet.Cells[string.Format("C{0}", row)].Value = item.Location;
                Sheet.Cells[string.Format("D{0}", row)].Value = item.OpsToKg;
                Sheet.Cells[string.Format("E{0}", row)].Value = item.IsActive ? "Y" : "N";

                Sheet.Cells[string.Format("A{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("B{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("C{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("D{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("E{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                row++;
            }

            Sheet.Cells["A:E"].AutoFitColumns();

            return Ep.GetAsByteArray();
        }
        #endregion

        #region ::Material Codes::
        public static byte[] ExportMaterialCodes(List<MaterialCodeModel> dataList, string username)
        {
            ExcelPackage Ep = new ExcelPackage();
            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Master-MaterialCode");
            AddHeader(Sheet, UIResources.MaterialCodesManagement, username);

            Color colFromHex = ColorTranslator.FromHtml(HEADER_COLOR);
            Sheet.Cells["A6:C6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Sheet.Cells["A6:C6"].Style.Fill.BackgroundColor.SetColor(colFromHex);
            Sheet.Cells["A6"].Value = "Code";
            Sheet.Cells["B6"].Value = "Description";
            Sheet.Cells["C6"].Value = "Location";
            Sheet.Cells["A6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["B6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["C6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            int row = 7;
            foreach (var item in dataList)
            {
                Sheet.Cells[string.Format("A{0}", row)].Value = item.Code;
                Sheet.Cells[string.Format("B{0}", row)].Value = item.Description;
                Sheet.Cells[string.Format("C{0}", row)].Value = item.Location;

                Sheet.Cells[string.Format("A{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("B{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("C{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                row++;
            }

            Sheet.Cells["A:C"].AutoFitColumns();

            return Ep.GetAsByteArray();
        }
        #endregion

        #region ::Calendar::
        public static byte[] ExportCalendar(List<CalendarModel> calendarList, string location, string gtype)
        {
            ExcelPackage Ep = new ExcelPackage();

            List<CalendarModel> calList = calendarList.OrderBy(x => x.Date).ToList();
            List<int> calendarMonthList = calList.GroupBy(p => p.Date.Month).Select(g => g.First().Date.Month).ToList();
            calendarMonthList.Sort();

            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add(location + " (" + gtype + ")");
            Color colFromHex = ColorTranslator.FromHtml(HEADER_COLOR);
            Sheet.Cells["A1:AF1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Sheet.Cells["A1:AF1"].Style.Fill.BackgroundColor.SetColor(colFromHex);
            Sheet.Cells["A1"].Value = "Month";
            Sheet.Cells["B1:AF1"].Merge = true;
            Sheet.Cells["B1:AF1"].Value = "Date";
            Sheet.Cells["B1:AF1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            int row = 2;
            foreach (var month in calendarMonthList)
            {
                List<CalendarModel> calPerMonthList = calList.Where(x => x.Date.Month == month).OrderBy(x => x.Date.Day).ToList();
                Sheet.Cells[string.Format("A{0}", row)].Value = calPerMonthList.First().Date.ToString("yyyyMM");
                Sheet.Cells[string.Format("A{0}", row + 1)].Value = "Shift I";
                Sheet.Cells[string.Format("A{0}", row + 2)].Value = "Shift II";
                Sheet.Cells[string.Format("A{0}", row + 3)].Value = "Shift III";

                foreach (var calendar in calPerMonthList)
                {
                    string columnFormat = Helper.Number2String(calendar.Date.Day + 1);

                    Sheet.Cells[string.Format(columnFormat, row)].Value = calendar.Date.Day;
                    Sheet.Cells[string.Format(columnFormat, row + 1)].Value = calendar.Shift1;
                    Sheet.Cells[string.Format(columnFormat, row + 2)].Value = calendar.Shift2;
                    Sheet.Cells[string.Format(columnFormat, row + 3)].Value = calendar.Shift3;

                    Sheet.Cells[string.Format(columnFormat, row)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    Sheet.Cells[string.Format(columnFormat, row + 1)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    Sheet.Cells[string.Format(columnFormat, row + 2)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    Sheet.Cells[string.Format(columnFormat, row + 3)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                }

                row += 5;
            }

            for (int i = 2; i < 33; i++)
            {
                Sheet.Column(i).Width = 5;
            }

            return Ep.GetAsByteArray();
        }
        #endregion

        #region ::Calendar Holiday::
        public static byte[] ExportCalendarHoliday(List<CalendarHolidayModel> calendarList, string username)
        {
            ExcelPackage Ep = new ExcelPackage();
            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Holiday");

            AddHeader(Sheet, UIResources.CalendarHolidayManagement, username);

            Color colFromHex = ColorTranslator.FromHtml(HEADER_COLOR);
            Sheet.Cells["A6:E6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Sheet.Cells["A6:E6"].Style.Fill.BackgroundColor.SetColor(colFromHex);
            Sheet.Cells["A6"].Value = UIResources.Location;
            Sheet.Cells["B6"].Value = UIResources.Year;
            Sheet.Cells["C6"].Value = UIResources.Month;
            Sheet.Cells["D6"].Value = UIResources.Date;
            Sheet.Cells["E6"].Value = UIResources.Information;
            Sheet.Cells["A6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["B6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["C6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["D6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["E6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            int row = 7;
            foreach (var data in calendarList)
            {
                Sheet.Cells[string.Format("A{0}", row)].Value = data.Location;
                Sheet.Cells[string.Format("B{0}", row)].Value = data.Date.ToString("yyyy");
                Sheet.Cells[string.Format("C{0}", row)].Value = data.Date.ToString("MMMM");
                Sheet.Cells[string.Format("D{0}", row)].Value = data.Date.ToString("dd");
                Sheet.Cells[string.Format("E{0}", row)].Value = data.Description;

                Sheet.Cells[string.Format("A{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("B{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("C{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("D{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("E{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                row++;
            }

            Sheet.Cells["A:E"].AutoFitColumns();

            return Ep.GetAsByteArray();
        }
        #endregion

        #region ::Location::
        public static byte[] ExportMasterLocation(LocationTreeModel locationModel, string username)
        {
            ExcelPackage Ep = new ExcelPackage();
            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Master-Location");
            AddHeader(Sheet, UIResources.LocationManagement, username);

            Color colFromHex = ColorTranslator.FromHtml(HEADER_COLOR);
            Sheet.Cells["A6:D6"].Style.Font.Bold = true;
            Sheet.Cells["A6:D6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Sheet.Cells["A6:D6"].Style.Fill.BackgroundColor.SetColor(colFromHex);
            Sheet.Cells["A6"].Value = UIResources.ProductionCenter;
            Sheet.Cells["B6"].Value = UIResources.Department;
            Sheet.Cells["C6"].Value = UIResources.SubDepartment;
            Sheet.Cells["D6"].Value = UIResources.LocationCode;
            Sheet.Cells["A6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["B6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["C6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["D6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            int row = 7;
            foreach (var pc in locationModel.ProductionCenters)
            {
                Sheet.Cells[string.Format("A{0}", row)].Value = pc.Description;
                Sheet.Cells[string.Format("B{0}", row)].Value = string.Empty;
                Sheet.Cells[string.Format("C{0}", row)].Value = string.Empty;
                Sheet.Cells[string.Format("D{0}", row)].Value = "ID-" + pc.Code;

                Sheet.Cells[string.Format("A{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("B{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("C{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("D{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                row++;

                foreach (var dep in pc.Departments)
                {
                    Sheet.Cells[string.Format("A{0}", row)].Value = pc.Description;
                    Sheet.Cells[string.Format("B{0}", row)].Value = dep.Description;
                    Sheet.Cells[string.Format("C{0}", row)].Value = string.Empty;
                    Sheet.Cells[string.Format("D{0}", row)].Value = "ID-" + pc.Code + "-" + dep.Code;

                    Sheet.Cells[string.Format("A{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    Sheet.Cells[string.Format("B{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    Sheet.Cells[string.Format("C{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    Sheet.Cells[string.Format("D{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    row++;

                    foreach (var subdep in dep.SubDepartments)
                    {
                        Sheet.Cells[string.Format("A{0}", row)].Value = pc.Description;
                        Sheet.Cells[string.Format("B{0}", row)].Value = dep.Description;
                        Sheet.Cells[string.Format("C{0}", row)].Value = subdep.Description;
                        Sheet.Cells[string.Format("D{0}", row)].Value = "ID-" + pc.Code + "-" + dep.Code + "-" + subdep.Code;

                        Sheet.Cells[string.Format("A{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        Sheet.Cells[string.Format("B{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        Sheet.Cells[string.Format("C{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        Sheet.Cells[string.Format("D{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                        row++;
                    }
                }
            }

            Sheet.Cells["A:D"].AutoFitColumns();

            return Ep.GetAsByteArray();
        }
        #endregion

        #region ::Reference::
        public static byte[] ExportMasterReference(ReferenceTreeModel refModel, string username)
        {
            ExcelPackage Ep = new ExcelPackage();
            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Master-Reference");
            AddHeader(Sheet, UIResources.LocationManagement, username);

            Color colFromHex = ColorTranslator.FromHtml(HEADER_COLOR);
            Sheet.Cells["A6:C6"].Style.Font.Bold = true;
            Sheet.Cells["A6:C6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Sheet.Cells["A6:C6"].Style.Fill.BackgroundColor.SetColor(colFromHex);
            Sheet.Cells["A6"].Value = UIResources.Purpose;
            Sheet.Cells["B6"].Value = UIResources.Code;
            Sheet.Cells["C6"].Value = UIResources.Description;
            Sheet.Cells["A6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["B6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["C6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            int row = 7;
            foreach (var reference in refModel.Parents)
            {
                Sheet.Cells[string.Format("A{0}", row)].Value = reference.Purpose;

                foreach (var detail in reference.ReferenceDetails)
                {
                    Sheet.Cells[string.Format("A{0}", row)].Value = reference.Purpose;
                    Sheet.Cells[string.Format("B{0}", row)].Value = detail.Code;
                    Sheet.Cells[string.Format("C{0}", row)].Value = detail.Description;

                    Sheet.Cells[string.Format("A{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    Sheet.Cells[string.Format("B{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    Sheet.Cells[string.Format("C{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    row++;
                }
            }

            Sheet.Cells["A:C"].AutoFitColumns();

            return Ep.GetAsByteArray();
        }
        #endregion

        #region ::Machine::
        public static byte[] ExportMasterMachine(List<MachineModel> machineList, string username)
        {
            ExcelPackage Ep = new ExcelPackage();
            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Master-Machine");
            AddHeader(Sheet, UIResources.MachineManagement, username);

            Color colFromHex = ColorTranslator.FromHtml(HEADER_COLOR);
            Sheet.Cells["A6:M6"].Style.Font.Bold = true;
            Sheet.Cells["A6:M6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Sheet.Cells["A6:M6"].Style.Fill.BackgroundColor.SetColor(colFromHex);
            Sheet.Cells["A6"].Value = UIResources.Location;
            Sheet.Cells["B6"].Value = UIResources.Code;
            Sheet.Cells["C6"].Value = UIResources.LegalEntity;
            Sheet.Cells["D6"].Value = UIResources.MachineBrand;
            Sheet.Cells["E6"].Value = UIResources.MachineType;
            Sheet.Cells["F6"].Value = UIResources.MachineSN;
            Sheet.Cells["G6"].Value = UIResources.Area;
            Sheet.Cells["H6"].Value = UIResources.Cluster;
            Sheet.Cells["I6"].Value = UIResources.OrderNumber;
            Sheet.Cells["J6"].Value = UIResources.Items;
            Sheet.Cells["K6"].Value = UIResources.DesignSpeed;
            Sheet.Cells["L6"].Value = UIResources.CellophanerSpeed;
            Sheet.Cells["M6"].Value = UIResources.LinkUp;

            Sheet.Cells["A6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["B6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["C6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["D6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["E6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["F6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["G6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["H6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["I6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["J6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["K6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["L6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["M6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            int row = 7;
            foreach (var data in machineList)
            {
                Sheet.Cells[string.Format("A{0}", row)].Value = data.Location;
                Sheet.Cells[string.Format("B{0}", row)].Value = data.Code;
                Sheet.Cells[string.Format("C{0}", row)].Value = data.LegalEntity;
                Sheet.Cells[string.Format("D{0}", row)].Value = data.MachineBrand;
                Sheet.Cells[string.Format("E{0}", row)].Value = data.MachineType;
                Sheet.Cells[string.Format("F{0}", row)].Value = data.MachineSN;
                Sheet.Cells[string.Format("G{0}", row)].Value = data.Notes;
                Sheet.Cells[string.Format("H{0}", row)].Value = data.Cluster;
                Sheet.Cells[string.Format("I{0}", row)].Value = data.OrderNumber;
                Sheet.Cells[string.Format("J{0}", row)].Value = data.Items;
                Sheet.Cells[string.Format("K{0}", row)].Value = data.DesignSpeed;
                Sheet.Cells[string.Format("L{0}", row)].Value = data.CellophanerSpeed;
                Sheet.Cells[string.Format("M{0}", row)].Value = data.LinkUp;

                Sheet.Cells[string.Format("A{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("B{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("C{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("D{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("E{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("F{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("G{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("H{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("I{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("J{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("K{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("L{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("M{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                row++;
            }

            Sheet.Cells["A:M"].AutoFitColumns();

            return Ep.GetAsByteArray();
        }
        #endregion

        #region ::Machine Type::
        public static byte[] ExportMasterMachineType(List<LocationMachineTypeModel> dataList, string username)
        {
            ExcelPackage Ep = new ExcelPackage();
            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Master-MachineType");
            AddHeader(Sheet, UIResources.MachineTypeManagement, username);

            Color colFromHex = ColorTranslator.FromHtml(HEADER_COLOR);
            Sheet.Cells["A6:B6"].Style.Font.Bold = true;
            Sheet.Cells["A6:B6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Sheet.Cells["A6:B6"].Style.Fill.BackgroundColor.SetColor(colFromHex);
            Sheet.Cells["A6"].Value = UIResources.Location;
            Sheet.Cells["B6"].Value = UIResources.Code;

            Sheet.Cells["A6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["B6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            int row = 7;
            foreach (var data in dataList)
            {
                Sheet.Cells[string.Format("A{0}", row)].Value = data.Location;
                Sheet.Cells[string.Format("B{0}", row)].Value = data.MachineType;

                Sheet.Cells[string.Format("A{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("B{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                row++;
            }

            Sheet.Cells["A:B"].AutoFitColumns();

            return Ep.GetAsByteArray();
        }
        #endregion

        #region ::EmployeeSkill::
        public static byte[] ExportMasterEmployeeSkill(List<UserMachineTypeModel> dataList, string username)
        {
            ExcelPackage Ep = new ExcelPackage();
            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Master-Employee-Skill");
            AddHeader(Sheet, UIResources.EmployeeSkillManagement, username);

            Color colFromHex = ColorTranslator.FromHtml(HEADER_COLOR);
            Sheet.Cells["A6:E6"].Style.Font.Bold = true;
            Sheet.Cells["A6:E6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Sheet.Cells["A6:E6"].Style.Fill.BackgroundColor.SetColor(colFromHex);
            Sheet.Cells["A6"].Value = UIResources.Username;
            Sheet.Cells["B6"].Value = UIResources.EmployeeID;
            Sheet.Cells["C6"].Value = UIResources.FullName;
            Sheet.Cells["D6"].Value = UIResources.PositionDesc;
            Sheet.Cells["E6"].Value = UIResources.Skills;

            Sheet.Cells["A6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["B6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["C6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["D6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["E6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            int row = 7;
            foreach (var data in dataList)
            {
                Sheet.Cells[string.Format("A{0}", row)].Value = data.UserName;
                Sheet.Cells[string.Format("B{0}", row)].Value = data.EmployeeID;
                Sheet.Cells[string.Format("C{0}", row)].Value = data.FullName;
                Sheet.Cells[string.Format("D{0}", row)].Value = data.PositionDesc;
                Sheet.Cells[string.Format("E{0}", row)].Value = data.Skills;

                Sheet.Cells[string.Format("A{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("B{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("C{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("D{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("E{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                row++;
            }

            Sheet.Cells["A:E"].AutoFitColumns();

            return Ep.GetAsByteArray();
        }
        #endregion

        #region ::UserMachine::
        public static byte[] ExportMasterUserMachine(List<UserMachineModel> dataList, string username)
        {
            ExcelPackage Ep = new ExcelPackage();
            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Master-UserMachine");
            AddHeader(Sheet, UIResources.UserMachineManagement, username);

            Color colFromHex = ColorTranslator.FromHtml(HEADER_COLOR);
            Sheet.Cells["A6:E6"].Style.Font.Bold = true;
            Sheet.Cells["A6:E6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Sheet.Cells["A6:E6"].Style.Fill.BackgroundColor.SetColor(colFromHex);
            Sheet.Cells["A6"].Value = UIResources.EmployeeID;
            Sheet.Cells["B6"].Value = UIResources.User;
            Sheet.Cells["C6"].Value = UIResources.PositionDesc;
            Sheet.Cells["D6"].Value = UIResources.MachineList;
            Sheet.Cells["E6"].Value = UIResources.Location;

            Sheet.Cells["A6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["B6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["C6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["D6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["E6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            int row = 7;
            foreach (var data in dataList)
            {
                Sheet.Cells[string.Format("A{0}", row)].Value = data.Employee.EmployeeID;
                Sheet.Cells[string.Format("B{0}", row)].Value = data.Employee.FullName;
                Sheet.Cells[string.Format("C{0}", row)].Value = data.Employee.PositionDesc;
                Sheet.Cells[string.Format("D{0}", row)].Value = data.MachineList;
                Sheet.Cells[string.Format("E{0}", row)].Value = data.Employee.BaseTownLocation;

                Sheet.Cells[string.Format("A{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("B{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("C{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("D{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("E{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                row++;
            }

            Sheet.Cells["A:E"].AutoFitColumns();

            return Ep.GetAsByteArray();
        }
        #endregion

        #region ::UserRole::
        public static byte[] ExportMasterUserRole(List<UserRoleModel> dataList, string username)
        {
            ExcelPackage Ep = new ExcelPackage();
            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Master-UserRole");
            AddHeader(Sheet, UIResources.UserMachineManagement, username);

            Color colFromHex = ColorTranslator.FromHtml(HEADER_COLOR);
            Sheet.Cells["A6:D6"].Style.Font.Bold = true;
            Sheet.Cells["A6:D6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Sheet.Cells["A6:D6"].Style.Fill.BackgroundColor.SetColor(colFromHex);
            Sheet.Cells["A6"].Value = UIResources.User;
            Sheet.Cells["B6"].Value = UIResources.PositionDesc;
            Sheet.Cells["C6"].Value = UIResources.RoleList;
            Sheet.Cells["D6"].Value = UIResources.Location;

            Sheet.Cells["A6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["B6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["C6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["D6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            int row = 7;
            foreach (var data in dataList)
            {
                Sheet.Cells[string.Format("A{0}", row)].Value = data.Employee.FullName;
                Sheet.Cells[string.Format("B{0}", row)].Value = data.Employee.PositionDesc;
                Sheet.Cells[string.Format("C{0}", row)].Value = data.RoleList;
                Sheet.Cells[string.Format("D{0}", row)].Value = data.Employee.BaseTownLocation;

                Sheet.Cells[string.Format("A{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("B{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("C{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("D{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                row++;
            }

            Sheet.Cells["A:D"].AutoFitColumns();

            return Ep.GetAsByteArray();
        }
        #endregion

        #region ::Role::
        public static byte[] ExportMasterRole(List<JobTitleModel> roleList, string username)
        {
            ExcelPackage Ep = new ExcelPackage();
            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Master-Role");
            AddHeader(Sheet, UIResources.RoleManagement, username);

            Color colFromHex = ColorTranslator.FromHtml(HEADER_COLOR);
            Sheet.Cells["A6:D6"].Style.Font.Bold = true;
            Sheet.Cells["A6:D6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Sheet.Cells["A6:D6"].Style.Fill.BackgroundColor.SetColor(colFromHex);
            Sheet.Cells["A6"].Value = UIResources.RoleOrSecurity;
            Sheet.Cells["B6"].Value = UIResources.JobTitle;
            Sheet.Cells["C6"].Value = UIResources.ModifiedBy;
            Sheet.Cells["D6"].Value = UIResources.ModifiedDate;
            Sheet.Cells["A6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["B6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["C6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["D6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            int row = 7;
            foreach (var role in roleList)
            {
                Sheet.Cells[string.Format("A{0}", row)].Value = role.RoleName;
                Sheet.Cells[string.Format("B{0}", row)].Value = role.Title;
                Sheet.Cells[string.Format("C{0}", row)].Value = role.ModifiedBy;
                Sheet.Cells[string.Format("D{0}", row)].Value = role.ModifiedDate.HasValue ? role.ModifiedDate.Value.ToString("dd-MMM-yy HH:mm") : string.Empty;

                Sheet.Cells[string.Format("A{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("B{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("C{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("D{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                row++;
            }

            Sheet.Cells["A:D"].AutoFitColumns();

            return Ep.GetAsByteArray();
        }
        #endregion

        #region ::OvertimeCategory::
        public static byte[] ExportOvertimeCategory(List<EmployeeOvertimeModel> dataList, List<SelectListItem> categoryList)
        {
            ExcelPackage Ep = new ExcelPackage();
            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Planning-Overtime-Category");
            Color colFromHex = ColorTranslator.FromHtml(HEADER_COLOR);
            Sheet.Cells["A1:E1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Sheet.Cells["A1:E1"].Style.Fill.BackgroundColor.SetColor(colFromHex);
            Sheet.Cells["A1"].Value = "Employee ID";
            Sheet.Cells["B1"].Value = "Employee Name";
            Sheet.Cells["C1"].Value = "Overtime Date";
            Sheet.Cells["D1"].Value = "Overtime Category";
            Sheet.Cells["E1"].Value = "Location";
            Sheet.Cells["A1"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["B1"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["C1"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["D1"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            int row = 2;
            foreach (var item in dataList)
            {
                Sheet.Cells[string.Format("A{0}", row)].Value = item.EmployeeID;
                Sheet.Cells[string.Format("B{0}", row)].Value = item.FullName;
                Sheet.Cells[string.Format("C{0}", row)].Value = item.Date.ToString("dd-MMM-yy");
                Sheet.Cells[string.Format("D{0}", row)].Value = item.OvertimeCategory;
                Sheet.Cells[string.Format("E{0}", row)].Value = item.Location;

                Sheet.Cells[string.Format("A{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("B{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("C{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("D{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("E{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                row++;
            }

            Sheet.Cells["A:E"].AutoFitColumns();

            ExcelWorksheet SheetCanteen = Ep.Workbook.Worksheets.Add("Category-List");
            SheetCanteen.Cells["A1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            SheetCanteen.Cells["A1"].Style.Fill.BackgroundColor.SetColor(colFromHex);
            SheetCanteen.Cells["A1"].Value = "Overtime Category";
            SheetCanteen.Cells["A1"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            row = 2;
            foreach (var item in categoryList)
            {
                SheetCanteen.Cells[string.Format("A{0}", row)].Value = item.Value;

                SheetCanteen.Cells[string.Format("A{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                row++;
            }

            SheetCanteen.Cells["A:B"].AutoFitColumns();

            return Ep.GetAsByteArray();
        }
        #endregion

        #region ::UserGroupType::
        public static byte[] ExportUserGroupType(List<EmployeeModel> dataList, List<ReferenceDetailModel> canteenList)
        {
            ExcelPackage Ep = new ExcelPackage();
            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Master-UserGroupType");
            Color colFromHex = ColorTranslator.FromHtml(HEADER_COLOR);
            Sheet.Cells["A1:G1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Sheet.Cells["A1:G1"].Style.Fill.BackgroundColor.SetColor(colFromHex);
            Sheet.Cells["A1"].Value = "Employee ID";
            Sheet.Cells["B1"].Value = "Employee Name";
            Sheet.Cells["C1"].Value = "Group Type";
            Sheet.Cells["D1"].Value = "Group Name";
            Sheet.Cells["E1"].Value = "Location";
            Sheet.Cells["F1"].Value = "Canteen Code";
            Sheet.Cells["G1"].Value = "Supervisor";
            Sheet.Cells["A1"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["B1"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["C1"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["D1"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["E1"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["F1"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["G1"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            int row = 2;
            foreach (var item in dataList)
            {
                Sheet.Cells[string.Format("A{0}", row)].Value = item.EmployeeID;
                Sheet.Cells[string.Format("B{0}", row)].Value = item.FullName;
                Sheet.Cells[string.Format("C{0}", row)].Value = item.GroupType;
                Sheet.Cells[string.Format("D{0}", row)].Value = item.GroupName;
                Sheet.Cells[string.Format("E{0}", row)].Value = item.Location;
                Sheet.Cells[string.Format("F{0}", row)].Value = item.Canteen;
                Sheet.Cells[string.Format("G{0}", row)].Value = item.HomeTownLocation;
                Sheet.Cells[string.Format("A{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("B{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("C{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("D{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("E{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("F{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("G{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                row++;
            }

            Sheet.Cells["H1:M1"].Merge = true;
            Sheet.Cells["H1:M1"].Value = "* Location will be updated, empty location get uploader user's location";
            Sheet.Cells["H1:M1"].Style.Font.Italic = true;
            Sheet.Cells["H1:M1"].Style.Font.Color.SetColor(Color.Red);

            Sheet.Cells["A:M"].AutoFitColumns();

            ExcelWorksheet SheetCanteen = Ep.Workbook.Worksheets.Add("Canteen-List");
            SheetCanteen.Cells["A1:B1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            SheetCanteen.Cells["A1:B1"].Style.Fill.BackgroundColor.SetColor(colFromHex);
            SheetCanteen.Cells["A1"].Value = "Canteen Code";
            SheetCanteen.Cells["B1"].Value = "Canteen Description";
            SheetCanteen.Cells["A1"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            SheetCanteen.Cells["B1"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            row = 2;
            foreach (var item in canteenList)
            {
                SheetCanteen.Cells[string.Format("A{0}", row)].Value = item.Code;
                SheetCanteen.Cells[string.Format("B{0}", row)].Value = item.Description;

                SheetCanteen.Cells[string.Format("A{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                SheetCanteen.Cells[string.Format("B{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                row++;
            }

            SheetCanteen.Cells["A:B"].AutoFitColumns();

            return Ep.GetAsByteArray();
        }
        #endregion

        #region ::Menu:
        public static byte[] ExportMasterMenu(List<MenuModel> dataList, string username)
        {
            ExcelPackage Ep = new ExcelPackage();
            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Master-Menu");
            AddHeader(Sheet, UIResources.MenuManagement, username);

            Color colFromHex = ColorTranslator.FromHtml(HEADER_COLOR);
            Sheet.Cells["A6:E6"].Style.Font.Bold = true;
            Sheet.Cells["A6:G6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Sheet.Cells["A6:G6"].Style.Fill.BackgroundColor.SetColor(colFromHex);
            Sheet.Cells["A6"].Value = UIResources.Menu;
            Sheet.Cells["B6"].Value = UIResources.SubMenu;
            Sheet.Cells["C6"].Value = UIResources.PageController;
            Sheet.Cells["D6"].Value = UIResources.PageAction;
            Sheet.Cells["E6"].Value = UIResources.PageIcon;
            Sheet.Cells["F6"].Value = UIResources.PageSlug;
            Sheet.Cells["G6"].Value = UIResources.DisplayOrder;
            Sheet.Cells["A6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["B6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["C6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["D6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["E6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["F6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["G6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            int row = 7;
            foreach (var data in dataList)
            {
                Sheet.Cells[string.Format("A{0}", row)].Value = data.ParentName;
                Sheet.Cells[string.Format("B{0}", row)].Value = data.Name;
                Sheet.Cells[string.Format("C{0}", row)].Value = data.PageController;
                Sheet.Cells[string.Format("D{0}", row)].Value = data.PageAction;
                Sheet.Cells[string.Format("E{0}", row)].Value = data.PageIcon;
                Sheet.Cells[string.Format("F{0}", row)].Value = data.PageSlug;
                Sheet.Cells[string.Format("G{0}", row)].Value = data.DisplayOrder;

                Sheet.Cells[string.Format("A{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("B{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("C{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("D{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("E{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("F{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("G{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                row++;
            }

            Sheet.Cells["A:G"].AutoFitColumns();

            return Ep.GetAsByteArray();
        }
        #endregion

        #region ::AccessRight::
        public static byte[] ExportMasterAccessRight(List<AccessRightDBModel> dataList, string username, string location)
        {
            ExcelPackage Ep = new ExcelPackage();
            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Master-Access-Right");
            AddHeaderLocation(Sheet, UIResources.AccessRightManagement, username, location);

            Color colFromHex = ColorTranslator.FromHtml(HEADER_COLOR);
            Sheet.Cells["A7:E7"].Style.Font.Bold = true;
            Sheet.Cells["A7:E7"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Sheet.Cells["A7:E7"].Style.Fill.BackgroundColor.SetColor(colFromHex);
            Sheet.Cells["A7"].Value = UIResources.Role;
            Sheet.Cells["B7"].Value = UIResources.Menu;
            Sheet.Cells["C7"].Value = UIResources.Read;
            Sheet.Cells["D7"].Value = UIResources.Write;
            Sheet.Cells["E7"].Value = UIResources.Print;
            Sheet.Cells["A7"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["B7"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["C7"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["D7"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["E7"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            int row = 8;
            foreach (var data in dataList)
            {
                Sheet.Cells[string.Format("A{0}", row)].Value = data.RoleName;
                Sheet.Cells[string.Format("B{0}", row)].Value = data.MenuName;
                Sheet.Cells[string.Format("C{0}", row)].Value = data.ReadName;
                Sheet.Cells[string.Format("D{0}", row)].Value = data.WriteName;
                Sheet.Cells[string.Format("E{0}", row)].Value = data.PrintName;

                Sheet.Cells[string.Format("A{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("B{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("C{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("D{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("E{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                row++;
            }

            Sheet.Cells["A:E"].AutoFitColumns();

            return Ep.GetAsByteArray();
        }
        #endregion

        #region ::User::
        public static byte[] ExportMasterUser(List<UserModel> userList, string username)
        {
            ExcelPackage Ep = new ExcelPackage();
            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Master-User");
            AddHeader(Sheet, UIResources.UserManagement, username);

            Color colFromHex = ColorTranslator.FromHtml(HEADER_COLOR);
            Sheet.Cells["A6:W6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Sheet.Cells["A6:W6"].Style.Font.Bold = true;
            Sheet.Cells["A6:W6"].Style.Fill.BackgroundColor.SetColor(colFromHex);
            Sheet.Cells["A6"].Value = UIResources.Username;
            Sheet.Cells["B6"].Value = UIResources.EmployeeID;
            Sheet.Cells["C6"].Value = UIResources.FullName;
            Sheet.Cells["D6"].Value = UIResources.Role;
            Sheet.Cells["E6"].Value = UIResources.Location;
            Sheet.Cells["F6"].Value = UIResources.Position;
            Sheet.Cells["G6"].Value = UIResources.BusinessUnit;
            Sheet.Cells["H6"].Value = UIResources.Department;
            Sheet.Cells["I6"].Value = UIResources.HmsOrg3;
            Sheet.Cells["J6"].Value = UIResources.CostCenter;
            Sheet.Cells["K6"].Value = UIResources.Canteen;
            Sheet.Cells["L6"].Value = UIResources.GroupType;
            Sheet.Cells["M6"].Value = UIResources.GroupName;
            Sheet.Cells["N6"].Value = UIResources.Type;
            Sheet.Cells["O6"].Value = UIResources.Act;
            Sheet.Cells["P6"].Value = UIResources.SupervisorID;
            Sheet.Cells["Q6"].Value = UIResources.SupervisorName;
            Sheet.Cells["R6"].Value = UIResources.ManagerID;
            Sheet.Cells["S6"].Value = UIResources.Manager;
            Sheet.Cells["T6"].Value = UIResources.Fast;
            Sheet.Cells["U6"].Value = UIResources.Admin;
            Sheet.Cells["V6"].Value = UIResources.ModifiedBy;
            Sheet.Cells["W6"].Value = UIResources.ModifiedDate;

            int row = 7;
            foreach (var user in userList)
            {
                Sheet.Cells[string.Format("A{0}", row)].Value = user.UserName;
                Sheet.Cells[string.Format("B{0}", row)].Value = user.EmployeeID;
                Sheet.Cells[string.Format("C{0}", row)].Value = user.Employee.FullName;
                Sheet.Cells[string.Format("D{0}", row)].Value = user.RoleName;
                Sheet.Cells[string.Format("E{0}", row)].Value = user.Location;
                Sheet.Cells[string.Format("F{0}", row)].Value = user.Employee.PositionDesc;
                Sheet.Cells[string.Format("G{0}", row)].Value = user.Employee.BusinessUnit;
                Sheet.Cells[string.Format("H{0}", row)].Value = user.Employee.DepartmentDesc;
                Sheet.Cells[string.Format("I{0}", row)].Value = user.Employee.HMSOrg3;
                Sheet.Cells[string.Format("J{0}", row)].Value = user.Employee.CostCenter;
                Sheet.Cells[string.Format("K{0}", row)].Value = user.Canteen;
                Sheet.Cells[string.Format("L{0}", row)].Value = user.Employee.GroupType;
                Sheet.Cells[string.Format("M{0}", row)].Value = user.Employee.GroupName;
                Sheet.Cells[string.Format("N{0}", row)].Value = user.Employee.EmployeeType;
                Sheet.Cells[string.Format("O{0}", row)].Value = user.Employee.Status != null && user.Employee.Status.Trim() == "Active" ? "Y" : "N"; ;
                Sheet.Cells[string.Format("P{0}", row)].Value = user.Employee.ReportToID1;
                Sheet.Cells[string.Format("Q{0}", row)].Value = user.SupervisorName;
                Sheet.Cells[string.Format("R{0}", row)].Value = user.Employee.ReportToID2;
                Sheet.Cells[string.Format("S{0}", row)].Value = user.ManagerName;
                Sheet.Cells[string.Format("T{0}", row)].Value = user.IsFast ? "Y" : "N";
                Sheet.Cells[string.Format("U{0}", row)].Value = user.IsAdmin ? "Y" : "N";
                Sheet.Cells[string.Format("V{0}", row)].Value = user.ModifiedBy;
                Sheet.Cells[string.Format("W{0}", row)].Value = user.ModifiedDate.HasValue ? user.ModifiedDate.Value.ToString("dd-MMM-yy HH:mm") : string.Empty;

                Sheet.Cells[string.Format("A{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("B{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("C{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("D{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("E{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("F{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("G{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("H{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("I{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("J{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("K{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("L{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("M{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("N{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("O{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("P{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("Q{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("R{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("S{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("T{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("U{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("V{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("W{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                row++;
            }

            Sheet.Cells["A:W"].AutoFitColumns();

            return Ep.GetAsByteArray();
        }
        #endregion

        #region ::MPP::
        public static byte[] ExportMPPTemplate(List<MppModel> mppList, DateTime startDate, DateTime endDate, string groupType)
        {
            ExcelPackage Ep = new ExcelPackage();
            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("MPPTemplate");
            Sheet.Cells["A1"].Value = "Start Date";
            Sheet.Cells["B1"].Value = startDate.ToString("yyyyMMdd");
            Sheet.Cells["C1"].Value = "End Date";
            Sheet.Cells["D1"].Value = endDate.ToString("yyyyMMdd");
            Sheet.Cells["E1"].Value = "Group Type";
            Sheet.Cells["F1"].Value = groupType;

            Color colFromHex = ColorTranslator.FromHtml(HEADER_COLOR);
            Color greenColFromHex = ColorTranslator.FromHtml("#2ecc71");
            Color orangeColFromHex = ColorTranslator.FromHtml("#f4d03f");
            Color lightblueColFromHex = ColorTranslator.FromHtml("#aed6f1");
            Color whiteColFromHex = ColorTranslator.FromHtml("#FFFFFF");

            DateTime startDateStatic = startDate;
            int index = 2;

            List<string> jobList = mppList.Select(c => c.JobTitle).Distinct().ToList();
            mppList = mppList.GroupBy(x => x.EmployeeID).Select(x => x.FirstOrDefault()).ToList();

            foreach (var jt in jobList)
            {
                Sheet.Cells[string.Format("A{0}", index)].Value = jt;
                Sheet.Cells[string.Format("B{0}:M{0}", index)].Style.Fill.PatternType = ExcelFillStyle.Solid;
                Sheet.Cells[string.Format("B{0}:M{0}", index)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                Color jtColor = colFromHex;
                if (jt == "PRODTECH")
                    jtColor = greenColFromHex;
                else if (jt == "ELECTRICIAN" || jt == "SUPPORT")
                    jtColor = orangeColFromHex;
                else if (jt == "RELIEF" || jt == "FOREMAN")
                    jtColor = lightblueColFromHex;

                if (jt == "SUPPORT")
                {
                    Sheet.Cells[string.Format("B{0}:J{0}", index)].Style.Fill.BackgroundColor.SetColor(jtColor);
                    Sheet.Cells[string.Format("K{0}:M{0}", index)].Style.Fill.BackgroundColor.SetColor(whiteColFromHex);
                }
                else
                    Sheet.Cells[string.Format("B{0}:M{0}", index)].Style.Fill.BackgroundColor.SetColor(jtColor);

                Sheet.Cells[string.Format("B{0}:D{0}", index)].Merge = true;
                Sheet.Cells[string.Format("B{0}:D{0}", index)].Value = "A";
                Sheet.Cells[string.Format("B{0}:D{0}", index)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("E{0}:G{0}", index)].Merge = true;
                Sheet.Cells[string.Format("E{0}:G{0}", index)].Value = "B";
                Sheet.Cells[string.Format("E{0}:G{0}", index)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("H{0}:J{0}", index)].Merge = true;
                Sheet.Cells[string.Format("H{0}:J{0}", index)].Value = "C";
                Sheet.Cells[string.Format("H{0}:J{0}", index)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                if (jt != "SUPPORT")
                {
                    Sheet.Cells[string.Format("K{0}:M{0}", index)].Merge = true;
                    Sheet.Cells[string.Format("K{0}:M{0}", index)].Value = "D";
                    Sheet.Cells[string.Format("K{0}:M{0}", index)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                }

                index++;
                int indexRowA = index;
                int indexRowB = index;
                int indexRowC = index;
                int indexRowD = index;
                startDate = startDateStatic;

                while (startDate <= endDate)
                {
                    List<MppModel> currentDateMppList = mppList.Where(x => x.Date == startDate && x.JobTitle == jt).ToList();
                    List<MppModel> mppAList = currentDateMppList.Where(x => x.GroupName.Trim() == "A").ToList();
                    List<MppModel> mppBList = currentDateMppList.Where(x => x.GroupName.Trim() == "B").ToList();
                    List<MppModel> mppCList = currentDateMppList.Where(x => x.GroupName.Trim() == "C").ToList();
                    List<MppModel> mppDList = currentDateMppList.Where(x => x.GroupName.Trim() == "D").ToList();

                    foreach (var item in mppAList)
                    {
                        Sheet.Cells[string.Format("B{0}", indexRowA)].Value = item.EmployeeID;
                        Sheet.Cells[string.Format("C{0}", indexRowA)].Value = item.EmployeeName;
                        Sheet.Cells[string.Format("D{0}", indexRowA)].Value = item.EmployeeMachine;
                        Sheet.Cells[string.Format("B{0}", indexRowA)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        Sheet.Cells[string.Format("C{0}", indexRowA)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        Sheet.Cells[string.Format("D{0}", indexRowA)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        indexRowA++;
                    }

                    foreach (var item in mppBList)
                    {
                        Sheet.Cells[string.Format("E{0}", indexRowB)].Value = item.EmployeeID;
                        Sheet.Cells[string.Format("F{0}", indexRowB)].Value = item.EmployeeName;
                        Sheet.Cells[string.Format("G{0}", indexRowB)].Value = item.EmployeeMachine;
                        Sheet.Cells[string.Format("E{0}", indexRowB)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        Sheet.Cells[string.Format("F{0}", indexRowB)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        Sheet.Cells[string.Format("G{0}", indexRowB)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        indexRowB++;
                    }

                    foreach (var item in mppCList)
                    {
                        Sheet.Cells[string.Format("H{0}", indexRowC)].Value = item.EmployeeID;
                        Sheet.Cells[string.Format("I{0}", indexRowC)].Value = item.EmployeeName;
                        Sheet.Cells[string.Format("J{0}", indexRowC)].Value = item.EmployeeMachine;
                        Sheet.Cells[string.Format("H{0}", indexRowC)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        Sheet.Cells[string.Format("I{0}", indexRowC)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        Sheet.Cells[string.Format("J{0}", indexRowC)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        indexRowC++;
                    }

                    foreach (var item in mppDList)
                    {
                        Sheet.Cells[string.Format("K{0}", indexRowD)].Value = item.EmployeeID;
                        Sheet.Cells[string.Format("L{0}", indexRowD)].Value = item.EmployeeName;
                        Sheet.Cells[string.Format("M{0}", indexRowD)].Value = item.EmployeeMachine;
                        Sheet.Cells[string.Format("K{0}", indexRowD)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        Sheet.Cells[string.Format("L{0}", indexRowD)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        Sheet.Cells[string.Format("M{0}", indexRowD)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        indexRowD++;
                    }

                    startDate = startDate.AddDays(1);
                }

                int maxRow = indexRowA;
                if (indexRowB > maxRow)
                    maxRow = indexRowB;
                if (indexRowC > maxRow)
                    maxRow = indexRowC;
                if (indexRowD > maxRow)
                    maxRow = indexRowD;

                index = maxRow + 1;
            }

            Sheet.Cells["A:M"].AutoFitColumns();

            return Ep.GetAsByteArray();
        }
        #endregion

        #region ::MPP By Date::
        public static byte[] ExportMPPByDate(List<MppModel> dataList, string username)
        {
            ExcelPackage Ep = new ExcelPackage();
            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Planning-MPP");
            AddHeader(Sheet, UIResources.MppDataManagement, username);

            Color colFromHex = ColorTranslator.FromHtml(HEADER_COLOR);
            Sheet.Cells["A6:J6"].Style.Font.Bold = true;
            Sheet.Cells["A6:J6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Sheet.Cells["A6:J6"].Style.Fill.BackgroundColor.SetColor(colFromHex);
            Sheet.Cells["A6"].Value = "Date";
            Sheet.Cells["B6"].Value = "Job Title";
            Sheet.Cells["C6"].Value = "Shift";
            Sheet.Cells["D6"].Value = "Status";
            Sheet.Cells["E6"].Value = "GroupType";
            Sheet.Cells["F6"].Value = "GroupName";
            Sheet.Cells["G6"].Value = "EmployeeID";
            Sheet.Cells["H6"].Value = "Name";
            Sheet.Cells["I6"].Value = "Machine";
            Sheet.Cells["J6"].Value = "Location";

            Sheet.Cells["A6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["B6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["C6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["D6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["E6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["F6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["G6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["H6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["I6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["J6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            int row = 7;
            foreach (var data in dataList)
            {
                Sheet.Cells[string.Format("A{0}", row)].Value = data.Date.ToString("dd-MMM-yy");
                Sheet.Cells[string.Format("B{0}", row)].Value = data.JobTitle;
                Sheet.Cells[string.Format("C{0}", row)].Value = data.Shift;
                Sheet.Cells[string.Format("D{0}", row)].Value = data.StatusMPP;
                Sheet.Cells[string.Format("E{0}", row)].Value = data.GroupType;
                Sheet.Cells[string.Format("F{0}", row)].Value = data.GroupName;
                Sheet.Cells[string.Format("G{0}", row)].Value = data.EmployeeID;
                Sheet.Cells[string.Format("H{0}", row)].Value = data.EmployeeName;
                Sheet.Cells[string.Format("I{0}", row)].Value = data.EmployeeMachine;
                Sheet.Cells[string.Format("J{0}", row)].Value = data.Location;

                Sheet.Cells[string.Format("A{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("B{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("C{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("D{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("E{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("F{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("G{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("H{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("I{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("J{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                row++;
            }

            Sheet.Cells["A:J"].AutoFitColumns();

            return Ep.GetAsByteArray();
        }
        #endregion

        #region ::Employee All::
        public static byte[] ExportEmployeeAll(List<EmployeeAllModel> dataList, string username)
        {
            ExcelPackage Ep = new ExcelPackage();
            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Master-Employee-All");
            AddHeader(Sheet, UIResources.EmployeeAllManagement, username);

            Color colFromHex = ColorTranslator.FromHtml(HEADER_COLOR);
            Sheet.Cells["A6:L6"].Style.Font.Bold = true;
            Sheet.Cells["A6:L6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Sheet.Cells["A6:L6"].Style.Fill.BackgroundColor.SetColor(colFromHex);

            Sheet.Cells["A6"].Value = "Employee ID";
            Sheet.Cells["B6"].Value = "FullName";
            Sheet.Cells["C6"].Value = "Position Desc";
            Sheet.Cells["D6"].Value = "Business Unit";
            Sheet.Cells["E6"].Value = "Department Desc";
            Sheet.Cells["F6"].Value = "HMSOrg3";
            Sheet.Cells["G6"].Value = "BaseTownLocation";
            Sheet.Cells["H6"].Value = "ReportToID1";
            Sheet.Cells["I6"].Value = "ReportToID2";
            Sheet.Cells["J6"].Value = "Status";
            Sheet.Cells["K6"].Value = "Email";
            Sheet.Cells["L6"].Value = "Employee Type";

            int row = 7;
            foreach (var item in dataList)
            {
                Sheet.Cells[string.Format("A{0}", row)].Value = item.EmployeeID;
                Sheet.Cells[string.Format("B{0}", row)].Value = item.FullName;
                Sheet.Cells[string.Format("C{0}", row)].Value = item.PositionDesc;
                Sheet.Cells[string.Format("D{0}", row)].Value = item.BusinessUnit;
                Sheet.Cells[string.Format("E{0}", row)].Value = item.DepartmentDesc;
                Sheet.Cells[string.Format("F{0}", row)].Value = item.HMSOrg3;
                Sheet.Cells[string.Format("G{0}", row)].Value = item.BaseTownLocation;
                Sheet.Cells[string.Format("H{0}", row)].Value = item.ReportToID1;
                Sheet.Cells[string.Format("I{0}", row)].Value = item.ReportToID2;
                Sheet.Cells[string.Format("J{0}", row)].Value = item.Status;
                Sheet.Cells[string.Format("K{0}", row)].Value = item.Email;
                Sheet.Cells[string.Format("L{0}", row)].Value = item.EmployeeType;

                Sheet.Cells[string.Format("A{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("B{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("C{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("D{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("E{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("F{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("G{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("H{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("I{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("J{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("K{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("L{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                row++;
            }

            Sheet.Cells["A:L"].AutoFitColumns();

            return Ep.GetAsByteArray();
        }
        #endregion

        #region ::EmployeeLeave::
        public static byte[] ExportEmployeeLeave(List<EmployeeLeaveModel> empLeaveModelList, string username)
        {
            ExcelPackage Ep = new ExcelPackage();
            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Planning-Employee-Leave");
            AddHeader(Sheet, UIResources.EmployeeLeaveManagement, username);

            Color colFromHex = ColorTranslator.FromHtml(HEADER_COLOR);
            Sheet.Cells["A6:L6"].Style.Font.Bold = true;
            Sheet.Cells["A6:L6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Sheet.Cells["A6:L6"].Style.Fill.BackgroundColor.SetColor(colFromHex);
            Sheet.Cells["A6"].Value = "Employee ID";
            Sheet.Cells["B6"].Value = "Employee Name";
            Sheet.Cells["C6"].Value = "Start Date";
            Sheet.Cells["D6"].Value = "End Date";
            Sheet.Cells["E6"].Value = "Comments";
            Sheet.Cells["F6"].Value = "Start Half Day";
            Sheet.Cells["G6"].Value = "Start Shift";
            Sheet.Cells["H6"].Value = "End Half Day";
            Sheet.Cells["I6"].Value = "End Shift";
            Sheet.Cells["J6"].Value = "Employee Type";
            Sheet.Cells["K6"].Value = "Leave Type";
            Sheet.Cells["L6"].Value = "Location";


            int row = 7;
            foreach (var item in empLeaveModelList)
            {
                Sheet.Cells[string.Format("A{0}", row)].Value = item.EmployeeID;
                Sheet.Cells[string.Format("B{0}", row)].Value = item.FullName;
                Sheet.Cells[string.Format("C{0}", row)].Value = item.StartDate.HasValue ? item.StartDate.Value.ToString("dd-MMM-yy") : string.Empty;
                Sheet.Cells[string.Format("D{0}", row)].Value = item.EndDate.HasValue ? item.EndDate.Value.ToString("dd-MMM-yy") : string.Empty;
                Sheet.Cells[string.Format("E{0}", row)].Value = item.Comments;
                Sheet.Cells[string.Format("F{0}", row)].Value = item.StartDateHalfDay;
                Sheet.Cells[string.Format("G{0}", row)].Value = item.StartDatePagiSiang;
                Sheet.Cells[string.Format("H{0}", row)].Value = item.EndDateHalfDay;
                Sheet.Cells[string.Format("I{0}", row)].Value = item.EndDatePagiSiang;
                Sheet.Cells[string.Format("J{0}", row)].Value = item.EmployeeType;
                Sheet.Cells[string.Format("K{0}", row)].Value = item.LeaveType;
                Sheet.Cells[string.Format("L{0}", row)].Value = item.Location;

                Sheet.Cells[string.Format("A{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("B{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("C{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("D{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("E{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("F{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("G{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("H{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("I{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("J{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("K{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("L{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                row++;
            }

            Sheet.Cells["A:L"].AutoFitColumns();

            return Ep.GetAsByteArray();
        }
        #endregion

        #region ::EmployeeOvertime::
        public static byte[] ExportEmployeeOvertime(List<EmployeeOvertimeModel> empOvertimeModelList, string username)
        {
            ExcelPackage Ep = new ExcelPackage();
            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Planning-Employee-Overtime");
            AddHeader(Sheet, UIResources.EmployeeOvertimeManagement, username);

            Color colFromHex = ColorTranslator.FromHtml(HEADER_COLOR);
            Sheet.Cells["A6:N6"].Style.Font.Bold = true;
            Sheet.Cells["A6:N6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Sheet.Cells["A6:N6"].Style.Fill.BackgroundColor.SetColor(colFromHex);
            Sheet.Cells["A6"].Value = "Employee ID";
            Sheet.Cells["B6"].Value = "Employee Name";
            Sheet.Cells["C6"].Value = "Department";
            Sheet.Cells["D6"].Value = "Position";
            Sheet.Cells["E6"].Value = "Location";
            Sheet.Cells["F6"].Value = "Cost Center";
            Sheet.Cells["G6"].Value = "Date";
            Sheet.Cells["H6"].Value = "Clock In";
            Sheet.Cells["I6"].Value = "Clock Out";
            Sheet.Cells["J6"].Value = "Actual In";
            Sheet.Cells["K6"].Value = "Actual Out";
            Sheet.Cells["L6"].Value = "Overtime";
            Sheet.Cells["M6"].Value = "Category";
            Sheet.Cells["N6"].Value = "Comment";

            int row = 7;
            foreach (var item in empOvertimeModelList)
            {
                Sheet.Cells[string.Format("A{0}", row)].Value = item.EmployeeID;
                Sheet.Cells[string.Format("B{0}", row)].Value = item.FullName;
                Sheet.Cells[string.Format("C{0}", row)].Value = item.DepartmentDesc;
                Sheet.Cells[string.Format("D{0}", row)].Value = item.PositionDesc;
                Sheet.Cells[string.Format("E{0}", row)].Value = item.Location;
                Sheet.Cells[string.Format("F{0}", row)].Value = item.CostCenter;
                Sheet.Cells[string.Format("G{0}", row)].Value = item.Date.ToString("dd-MMM-yy");
                Sheet.Cells[string.Format("H{0}", row)].Value = item.ClockIn;
                Sheet.Cells[string.Format("I{0}", row)].Value = item.ClockOut;
                Sheet.Cells[string.Format("J{0}", row)].Value = item.ActualIn;
                Sheet.Cells[string.Format("K{0}", row)].Value = item.ActualOut;
                Sheet.Cells[string.Format("L{0}", row)].Value = item.Overtime;
                Sheet.Cells[string.Format("M{0}", row)].Value = item.OvertimeCategory;
                Sheet.Cells[string.Format("N{0}", row)].Value = item.Comments;

                Sheet.Cells[string.Format("A{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("B{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("C{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("D{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("E{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("F{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("G{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("H{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("I{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("J{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("K{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("L{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("M{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("N{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                row++;
            }

            Sheet.Cells["A:AZ"].AutoFitColumns();

            return Ep.GetAsByteArray();
        }
        #endregion

        #region ::Training::
        public static byte[] ExportTraining(List<TrainingModel> trainingModelList, string username)
        {
            ExcelPackage Ep = new ExcelPackage();
            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Master-Training");
            AddHeader(Sheet, UIResources.TrainingManagement, username);

            Color colFromHex = ColorTranslator.FromHtml(HEADER_COLOR);
            Sheet.Cells["A6:N6"].Style.Font.Bold = true;
            Sheet.Cells["A6:N6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Sheet.Cells["A6:N6"].Style.Fill.BackgroundColor.SetColor(colFromHex);
            Sheet.Cells["A6"].Value = "Employee ID";
            Sheet.Cells["B6"].Value = "Employee Name";
            Sheet.Cells["C6"].Value = "Position";
            Sheet.Cells["D6"].Value = "Department";
            Sheet.Cells["E6"].Value = "Business Unit";
            Sheet.Cells["F6"].Value = "BasetownLocation";
            Sheet.Cells["G6"].Value = "TrainingCategory";
            Sheet.Cells["H6"].Value = "TrainingTitle";
            Sheet.Cells["I6"].Value = "Start Date";
            Sheet.Cells["J6"].Value = "End Date";
            Sheet.Cells["K6"].Value = "Trainer";
            Sheet.Cells["L6"].Value = "Score";
            Sheet.Cells["M6"].Value = "StatusTraining";
            Sheet.Cells["N6"].Value = "Machine Type";

            int row = 7;
            foreach (var item in trainingModelList)
            {
                Sheet.Cells[string.Format("A{0}", row)].Value = item.EmployeeID;
                Sheet.Cells[string.Format("B{0}", row)].Value = item.FullName;
                Sheet.Cells[string.Format("C{0}", row)].Value = item.Position;
                Sheet.Cells[string.Format("D{0}", row)].Value = item.Department;
                Sheet.Cells[string.Format("E{0}", row)].Value = item.BU;
                Sheet.Cells[string.Format("F{0}", row)].Value = item.BasetownLocation;
                Sheet.Cells[string.Format("G{0}", row)].Value = item.TrainingCategory;
                Sheet.Cells[string.Format("H{0}", row)].Value = item.TrainingTitle;
                Sheet.Cells[string.Format("I{0}", row)].Value = item.StartDate.HasValue ? item.StartDate.Value.ToString("dd-MMM-yy") : string.Empty;
                Sheet.Cells[string.Format("J{0}", row)].Value = item.EndDate.HasValue ? item.EndDate.Value.ToString("dd-MMM-yy") : string.Empty;
                Sheet.Cells[string.Format("K{0}", row)].Value = item.Trainer;
                Sheet.Cells[string.Format("L{0}", row)].Value = item.Score;
                Sheet.Cells[string.Format("M{0}", row)].Value = item.StatusTraining;
                Sheet.Cells[string.Format("N{0}", row)].Value = item.MachineType;

                Sheet.Cells[string.Format("A{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("B{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("C{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("D{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("E{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("F{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("G{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("H{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("I{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("J{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("K{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("L{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("M{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("N{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                row++;
            }

            Sheet.Cells["A:N"].AutoFitColumns();

            return Ep.GetAsByteArray();
        }
        #endregion

        #region ::User Logs::
        public static byte[] ExportMasterUserLogs(List<UserLogModel> dataList, string username)
        {
            ExcelPackage Ep = new ExcelPackage();
            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Master-UserLogs");
            AddHeader(Sheet, UIResources.UserLogManagement, username);

            Color colFromHex = ColorTranslator.FromHtml(HEADER_COLOR);
            Sheet.Cells["A6:F6"].Style.Font.Bold = true;
            Sheet.Cells["A6:F6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Sheet.Cells["A6:F6"].Style.Fill.BackgroundColor.SetColor(colFromHex);

            Sheet.Cells["A6"].Value = UIResources.Timestamp;
            Sheet.Cells["B6"].Value = UIResources.UserID;
            Sheet.Cells["C6"].Value = UIResources.Username;
            Sheet.Cells["D6"].Value = UIResources.Level;
            Sheet.Cells["E6"].Value = UIResources.Message;
            Sheet.Cells["F6"].Value = UIResources.Stacktrace;

            Sheet.Cells["A6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["B6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["C6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["D6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["E6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["F6"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            int row = 7;
            foreach (var data in dataList)
            {
                Sheet.Cells[string.Format("A{0}", row)].Value = data.Timestamp;
                Sheet.Cells[string.Format("B{0}", row)].Value = data.UserID;
                Sheet.Cells[string.Format("C{0}", row)].Value = data.UserName;
                Sheet.Cells[string.Format("D{0}", row)].Value = data.Level;
                Sheet.Cells[string.Format("E{0}", row)].Value = data.Message;
                Sheet.Cells[string.Format("F{0}", row)].Value = data.Stacktrace;

                Sheet.Cells[string.Format("A{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("B{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("C{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("D{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("E{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("F{0}", row)].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                row++;
            }

            Sheet.Cells["A:F"].AutoFitColumns();

            return Ep.GetAsByteArray();
        }
        #endregion

        #region ::Add Header::
        private static void AddHeader(ExcelWorksheet Sheet, string title, string username)
        {
            using (System.Drawing.Image image = System.Drawing.Image.FromFile(HttpContext.Current.Server.MapPath("~/Content/theme/images/fast-blue.jpg")))
            {
                var excelImage = Sheet.Drawings.AddPicture("Fast Logo", image);
                excelImage.SetPosition(0, 0, 0, 0);
            }

            Sheet.Cells["A3"].Value = UIResources.Title;
            Sheet.Cells["A4"].Value = UIResources.GeneratedBy;
            Sheet.Cells["A5"].Value = UIResources.GeneratedDate;
            Sheet.Cells["B3"].Value = title;
            Sheet.Cells["B4"].Value = username;
            Sheet.Cells["B5"].Value = DateTime.Now.ToString("dd-MMM-yy HH:mm:ss");
        }
        private static void AddHeaderSim(ExcelWorksheet Sheet, string title, string username, DateTime startdate, DateTime enddate)
        {
            using (System.Drawing.Image image = System.Drawing.Image.FromFile(HttpContext.Current.Server.MapPath("~/Content/theme/images/fast-blue.jpg")))
            {
                var excelImage = Sheet.Drawings.AddPicture("Fast Logo", image);
                excelImage.SetPosition(0, 0, 0, 0);
            }

            Sheet.Cells["A3"].Value = UIResources.Title;
            Sheet.Cells["A4"].Value = UIResources.Period;
            Sheet.Cells["A5"].Value = UIResources.GeneratedBy;
            Sheet.Cells["A6"].Value = UIResources.GeneratedDate;
            Sheet.Cells["B3"].Value = title;
            Sheet.Cells["B4"].Value = "[ " + startdate.ToString("dd-MMM-yy") + " ] - [ " + enddate.ToString("dd-MMM-yy") + " ]";
            Sheet.Cells["B5"].Value = username;
            Sheet.Cells["B6"].Value = DateTime.Now.ToString("dd-MMM-yy HH:mm:ss");
        }

        private static void AddHeaderLocation(ExcelWorksheet Sheet, string title, string username, string location)
        {
            using (System.Drawing.Image image = System.Drawing.Image.FromFile(HttpContext.Current.Server.MapPath("~/Content/theme/images/fast-blue.jpg")))
            {
                var excelImage = Sheet.Drawings.AddPicture("Fast Logo", image);
                excelImage.SetPosition(0, 0, 0, 0);
            }

            Sheet.Cells["A3"].Value = UIResources.Title;
            Sheet.Cells["A4"].Value = UIResources.GeneratedBy;
            Sheet.Cells["A5"].Value = UIResources.GeneratedDate;
            Sheet.Cells["A6"].Value = UIResources.Location;
            Sheet.Cells["B3"].Value = title;
            Sheet.Cells["B4"].Value = username;
            Sheet.Cells["B5"].Value = DateTime.Now.ToString("dd-MMM-yy HH:mm:ss");
            Sheet.Cells["B6"].Value = location;
        }
        #endregion

        #region ::Input Daily::
        public static byte[] ExportInputDaily(string username, string pc, List<string> luList) //DateTime date,
        {
            ExcelPackage Ep = new ExcelPackage();

            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Input Daily");

            AddHeader(Sheet, "Input Daily", username);

            Sheet.Cells["A7:AF7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            Sheet.Cells["B7:AF7"].AutoFitColumns();

            Sheet.Cells["A4:A5"].AutoFitColumns();
            //Sheet.Cells["A6"].Value = "Date";
            //Sheet.Cells["B6"].Value = dateVal;
            Sheet.Cells["A7"].Value = pc;
            Sheet.Cells["A7:C7"].Merge = true;
            //Sheet.Cells["A7:C7"].Style.Locked = true;
            //Sheet.Cells["A7:C7"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //Sheet.Cells["A7:C7"].Style.Fill.BackgroundColor.SetColor(Color.Red);

            int rowStart = 8;
            int kpiStartColumn = 2;

            List<string> kpiList = new List<string>() { "MTBF", "CPQI", "VQI", "Working Time", "Uptime", "STRS", "Production Volume", "CRR" };
            //List<string> luList = new List<string>() { "LU51", "LU52", "LU53"};

            foreach (string lu in luList)
            {
                using (var range = Sheet.Cells[rowStart, 1, rowStart, kpiList.Count + 1])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.GreenYellow);
                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                }
                Sheet.Cells[string.Format("A{0}", rowStart)].Value = "Link Up";
                Sheet.Cells[string.Format("B{0}", rowStart++)].Value = lu;

                for (int i = 1; i < 4; i++)
                {
                    string shiftText = "Shift " + i;

                    foreach (var itemKPI in kpiList)
                    {
                        Color colFromHex = ColorTranslator.FromHtml(HEADER_COLOR);
                        Sheet.Cells[string.Format("A{0}", rowStart)].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        Sheet.Cells[string.Format("A{0}", rowStart)].Style.Fill.BackgroundColor.SetColor(colFromHex);
                        Sheet.Cells[string.Format("A{0}", rowStart)].Value = shiftText;
                        Sheet.Cells[string.Format("A{0}", rowStart + 1)].Value = "Value";
                        Sheet.Cells[string.Format("A{0}", rowStart + 2)].Value = "Focus";
                        Sheet.Cells[string.Format("A{0}", rowStart + 3)].Value = "Remarks";

                        string columnFormat = Helper.Number2String(kpiStartColumn);
                        Sheet.Cells[string.Format(columnFormat, rowStart)].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        Sheet.Cells[string.Format(columnFormat, rowStart)].Style.Fill.BackgroundColor.SetColor(colFromHex);
                        Sheet.Cells[string.Format(columnFormat, rowStart)].Value = itemKPI;
                        Sheet.Cells[string.Format(columnFormat, rowStart)].AutoFitColumns();
                        kpiStartColumn += 1;
                    }
                    rowStart = rowStart + 4;
                    kpiStartColumn = 2;
                }
            }
            return Ep.GetAsByteArray();
        }

        #endregion


        #region ::Shift Daily ::
        public static byte[] ExportShiftDaily(string username, Tuple<List<DailyModel>, List<TargetModel>, List<UptimeModel>, WTDModel, MTDModel> model, string pcDesc, string machine, string dateVal)
        {
            string Date_Color = "#eb8034";
            string Mach_Color = "#99ccff";
            string Trend_Color = "#9c4a10";
            string Target_Color = "#30963f";

            ExcelPackage Ep = new ExcelPackage();
            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("ShiftDaily");

            AddHeader(Sheet, "Shift Daily", username);
            Sheet.Cells["A9:AF9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            Sheet.Cells["A4:A5"].AutoFitColumns();
            Sheet.Cells["A6"].Value = "Date";
            Sheet.Cells["B6"].Value = dateVal;
            Sheet.Cells["A7"].Value = "Production Centre";
            Sheet.Cells["B7"].Value = pcDesc;

            List<DailyModel> dailyRes = model.Item1;
            List<TargetModel> targetRes = model.Item2;
            List<UptimeModel> uptimeRes = model.Item3;

            WTDModel wtdModel = model.Item4;
            MTDModel mtdModel = model.Item5;

            decimal target_MTBF = 0;
            decimal target_ProdVol = 0;
            decimal target_CPQI = 0;
            decimal target_VQI = 0;
            decimal target_STRS = 0;
            decimal target_CRR = 0;
            decimal target_Uptime = 0;
            decimal target_Working = 0;

            decimal mtg_ProdVol = 0;
            decimal mtg_CPQI = 0;
            decimal mtg_VQI = 0;
            decimal mtg_STRS = 0;
            decimal mtg_CRR = 0;
            decimal mtg_Uptime = 0;

            #region Get Data, Calculation, Layout

            for (int s = 0; s < targetRes.Count(); s++)
            {
                if (targetRes[s].KPI == "MTBF")
                {
                    target_MTBF = targetRes[s].ValueTarget;
                }
                if (targetRes[s].KPI == "CPQI")
                {
                    target_CPQI = targetRes[s].ValueTarget;
                }
                if (targetRes[s].KPI == "VQI")
                {
                    target_VQI = targetRes[s].ValueTarget;
                }
                if (targetRes[s].KPI == "Working Time")
                {
                    target_Working = targetRes[s].ValueTarget;
                }
                if (targetRes[s].KPI == "STRS")
                {
                    target_STRS = targetRes[s].ValueTarget;
                }
                if (targetRes[s].KPI == "CRR")
                {
                    target_CRR = targetRes[s].ValueTarget;
                }
                if (targetRes[s].KPI == "Production Volume")
                {
                    target_ProdVol = targetRes[s].ValueTarget;
                }
                if (targetRes[s].KPI == "Uptime")
                {
                    target_Uptime = targetRes[s].ValueTarget;
                }
            }

            for (int i = 0; i < targetRes.Count(); i++)
            {
                if (targetRes[i].KPI == "Production Volume")
                {
                    mtg_ProdVol = (targetRes[i].ValueTarget - Convert.ToDecimal(mtdModel.ProdVolume_MTD));
                }
            }

            mtg_CPQI = ((target_ProdVol * target_CPQI) - (Convert.ToDecimal(mtdModel.ProdVolume_MTD) * Convert.ToDecimal(mtdModel.CPQI_MTD))) / mtg_ProdVol;
            mtg_VQI = ((target_ProdVol * target_VQI) - (Convert.ToDecimal(mtdModel.ProdVolume_MTD) * Convert.ToDecimal(mtdModel.VQI_MTD))) / mtg_ProdVol;
            mtg_STRS = ((target_ProdVol * target_STRS) - (Convert.ToDecimal(mtdModel.ProdVolume_MTD) * Convert.ToDecimal(mtdModel.STRS_MTD))) / mtg_ProdVol;
            mtg_CRR = ((target_ProdVol * target_CRR) - (Convert.ToDecimal(mtdModel.ProdVolume_MTD) * Convert.ToDecimal(mtdModel.CRR_MTD))) / mtg_ProdVol;

            decimal mtd_Theo = 0;
            mtd_Theo = (Convert.ToDecimal(mtdModel.ProdVolume_MTD) / Convert.ToDecimal(mtdModel.ProdVolume_MTD)) * 100;

            decimal target_Theo = 0;
            target_Theo = (target_ProdVol / target_Uptime) * 100;

            decimal mtg_Theo = 0;
            mtg_Theo = target_Theo - mtd_Theo;

            mtg_Uptime = (mtg_ProdVol / mtg_Theo) * 100;

            Color colFromHexMach = ColorTranslator.FromHtml(Mach_Color);
            Sheet.Cells["A9"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Sheet.Cells["A9"].Style.Fill.BackgroundColor.SetColor(colFromHexMach);

            for (int i = 9; i <= 41; i++)
            {
                Sheet.Cells[string.Format("A{0}", i)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("B{0}", i)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("C{0}", i)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("D{0}", i)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("E{0}", i)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("F{0}", i)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("G{0}", i)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("H{0}", i)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                Sheet.Cells[string.Format("I{0}", i)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            }

            for (int j = 9; j <= 41; j++)
            {
                Sheet.Cells[string.Format("C{0}", j)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                Sheet.Cells[string.Format("D{0}", j)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                Sheet.Cells[string.Format("E{0}", j)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                Sheet.Cells[string.Format("F{0}", j)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                Sheet.Cells[string.Format("G{0}", j)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                Sheet.Cells[string.Format("H{0}", j)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                Sheet.Cells[string.Format("I{0}", j)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            }
            #endregion

            #region Column Header

            Color colFromHexTrend = ColorTranslator.FromHtml(Trend_Color);
            Sheet.Cells["B9"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Sheet.Cells["B9"].Style.Fill.BackgroundColor.SetColor(colFromHexTrend);

            Sheet.Cells["A9"].Style.Font.Bold = true;
            Sheet.Cells["A9"].Value = machine;

            Sheet.Cells["B9"].Style.Font.Bold = true;
            Sheet.Cells["B9"].Value = "Trend";
            Sheet.Cells["A9"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Sheet.Cells["B9"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            Color colFromHexDate = ColorTranslator.FromHtml(Date_Color);

            int dateColumn = 3;

            for (int i = 0; i < dailyRes.Count(); i++)
            {
                if (i == 0 || i % 3 == 0)
                {
                    string columnFormat = Helper.Number2String(dateColumn);
                    Sheet.Cells[string.Format(columnFormat, 9)].Style.Font.Bold = true;
                    Sheet.Cells[string.Format(columnFormat, 9)].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    Sheet.Cells[string.Format(columnFormat, 9)].Style.Fill.BackgroundColor.SetColor(colFromHexDate);

                    Sheet.Cells[string.Format(columnFormat, 9)].Value = dailyRes[i].DateValue.ToString("dd-MMM");
                    dateColumn += 1;
                }

            }
            Sheet.Cells["F9:H9"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Sheet.Cells["F9:H9"].Style.Fill.BackgroundColor.SetColor(colFromHexDate);

            Sheet.Cells["F9"].Value = "WTD";
            Sheet.Cells["G9"].Value = "MTD";
            Sheet.Cells["H9"].Value = "MTG";

            Sheet.Cells["F9"].Style.Font.Bold = true;
            Sheet.Cells["G9"].Style.Font.Bold = true;
            Sheet.Cells["H9"].Style.Font.Bold = true;

            Color colFromHexTarget = ColorTranslator.FromHtml(Target_Color);
            Sheet.Cells["I9"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            Sheet.Cells["I9"].Style.Fill.BackgroundColor.SetColor(colFromHexTarget);

            Sheet.Cells["I9"].Style.Font.Bold = true;
            Sheet.Cells["I9"].Value = "Target";


            #endregion

            #region MTBF

            Sheet.Cells["A10"].Style.Font.Bold = true;
            Sheet.Cells["A10"].Value = "MTBF";

            Sheet.Cells["A11:A13"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            Sheet.Cells["A11"].Value = "Shift 1";
            Sheet.Cells["A12"].Value = "Shift 2";
            Sheet.Cells["A13"].Value = "Shift 3";

            Sheet.Cells["C10"].Value = (dailyRes[0].MTBFValue + dailyRes[1].MTBFValue + dailyRes[2].MTBFValue).ToString("n");
            Sheet.Cells["D10"].Value = (dailyRes[3].MTBFValue + dailyRes[4].MTBFValue + dailyRes[5].MTBFValue).ToString("n");
            Sheet.Cells["E10"].Value = (dailyRes[6].MTBFValue + dailyRes[7].MTBFValue + dailyRes[8].MTBFValue).ToString("n");

            Sheet.Cells["F10"].Value = wtdModel.MTBF_WTD;
            Sheet.Cells["G10"].Value = mtdModel.MTBF_MTD;
            Sheet.Cells["H10"].Value = ((target_MTBF - Convert.ToDecimal(mtdModel.MTBF_MTD)).ToString("n"));
            Sheet.Cells["I10"].Value = target_MTBF.ToString("n");

            dateColumn = 3;
            for (int j = 0; j < dailyRes.Count(); j++)
            {
                if (j == 0 || j % 3 == 0)
                {
                    if (dailyRes[j].Shift == "1")
                    {
                        string columnFormat = Helper.Number2String(dateColumn);
                        Sheet.Cells[string.Format(columnFormat, 11)].Value = dailyRes[j].MTBFValue.ToString("n");
                        dateColumn += 1;
                    }
                }
            }
            dateColumn = 3;
            for (int k = 0; k < dailyRes.Count(); k++)
            {
                if (k == 1 || k % 3 == 1)
                {
                    if (dailyRes[k].Shift == "2")
                    {
                        string columnFormat = Helper.Number2String(dateColumn);
                        Sheet.Cells[string.Format(columnFormat, 12)].Value = dailyRes[k].MTBFValue.ToString("n");
                        dateColumn += 1;
                    }
                }
            }
            dateColumn = 3;
            for (int l = 0; l < dailyRes.Count(); l++)
            {
                if (l == 2 || l % 3 == 2)
                {
                    if (dailyRes[l].Shift == "3")
                    {
                        string columnFormat = Helper.Number2String(dateColumn);
                        Sheet.Cells[string.Format(columnFormat, 13)].Value = dailyRes[l].MTBFValue.ToString("n");
                        dateColumn += 1;
                    }
                }
            }

            #endregion

            #region CPQI

            Sheet.Cells["A14"].Style.Font.Bold = true;
            Sheet.Cells["A14"].Value = "CPQI";
            Sheet.Cells["A15:A17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            Sheet.Cells["A15"].Value = "Shift 1";
            Sheet.Cells["A16"].Value = "Shift 2";
            Sheet.Cells["A17"].Value = "Shift 3";

            Sheet.Cells["C14"].Value = (((dailyRes[0].ProdVolumeValue * dailyRes[0].CPQIValue)
                                        + (dailyRes[1].ProdVolumeValue * dailyRes[1].CPQIValue)
                                        + (dailyRes[2].ProdVolumeValue * dailyRes[2].CPQIValue))
                                        / (dailyRes[0].ProdVolumeValue + dailyRes[1].ProdVolumeValue + dailyRes[2].ProdVolumeValue)).ToString("n");

            Sheet.Cells["D14"].Value = (((dailyRes[3].ProdVolumeValue * dailyRes[3].CPQIValue)
                                        + (dailyRes[4].ProdVolumeValue * dailyRes[4].CPQIValue)
                                        + (dailyRes[5].ProdVolumeValue * dailyRes[5].CPQIValue))
                                        / (dailyRes[3].ProdVolumeValue + dailyRes[4].ProdVolumeValue + dailyRes[5].ProdVolumeValue)).ToString("n");

            Sheet.Cells["E14"].Value = (((dailyRes[6].ProdVolumeValue * dailyRes[6].CPQIValue)
                                        + (dailyRes[7].ProdVolumeValue * dailyRes[7].CPQIValue)
                                        + (dailyRes[8].ProdVolumeValue * dailyRes[8].CPQIValue))
                                        / (dailyRes[6].ProdVolumeValue + dailyRes[7].ProdVolumeValue + dailyRes[8].ProdVolumeValue)).ToString("n");

            Sheet.Cells["F14"].Value = wtdModel.CPQI_WTD;
            Sheet.Cells["G14"].Value = mtdModel.CPQI_MTD;
            Sheet.Cells["H14"].Value = mtg_CPQI.ToString("n");
            Sheet.Cells["I14"].Value = target_CPQI.ToString("n");

            dateColumn = 3;
            for (int j = 0; j < dailyRes.Count(); j++)
            {
                if (j == 0 || j % 3 == 0)
                {
                    if (dailyRes[j].Shift == "1")
                    {
                        string columnFormat = Helper.Number2String(dateColumn);
                        Sheet.Cells[string.Format(columnFormat, 15)].Value = dailyRes[j].CPQIValue.ToString("n");
                        dateColumn += 1;
                    }
                }
            }
            dateColumn = 3;
            for (int k = 0; k < dailyRes.Count(); k++)
            {
                if (k == 1 || k % 3 == 1)
                {
                    if (dailyRes[k].Shift == "2")
                    {
                        string columnFormat = Helper.Number2String(dateColumn);
                        Sheet.Cells[string.Format(columnFormat, 16)].Value = dailyRes[k].CPQIValue.ToString("n");
                        dateColumn += 1;
                    }
                }
            }
            dateColumn = 3;
            for (int l = 0; l < dailyRes.Count(); l++)
            {
                if (l == 2 || l % 3 == 2)
                {
                    if (dailyRes[l].Shift == "3")
                    {
                        string columnFormat = Helper.Number2String(dateColumn);
                        Sheet.Cells[string.Format(columnFormat, 17)].Value = dailyRes[l].CPQIValue.ToString("n");
                        dateColumn += 1;
                    }
                }
            }
            #endregion

            #region VQI

            Sheet.Cells["A18"].Style.Font.Bold = true;
            Sheet.Cells["A18"].Value = "VQI";
            Sheet.Cells["A19:A21"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            Sheet.Cells["A19"].Value = "Shift 1";
            Sheet.Cells["A20"].Value = "Shift 2";
            Sheet.Cells["A21"].Value = "Shift 3";

            Sheet.Cells["C18"].Value = (((dailyRes[0].ProdVolumeValue * dailyRes[0].VQIValue)
                                       + (dailyRes[1].ProdVolumeValue * dailyRes[1].VQIValue)
                                       + (dailyRes[2].ProdVolumeValue * dailyRes[2].VQIValue))
                                       / (dailyRes[0].ProdVolumeValue + dailyRes[1].ProdVolumeValue + dailyRes[2].ProdVolumeValue)).ToString("n");

            Sheet.Cells["D18"].Value = (((dailyRes[3].ProdVolumeValue * dailyRes[3].VQIValue)
                                        + (dailyRes[4].ProdVolumeValue * dailyRes[4].VQIValue)
                                        + (dailyRes[5].ProdVolumeValue * dailyRes[5].VQIValue))
                                        / (dailyRes[3].ProdVolumeValue + dailyRes[4].ProdVolumeValue + dailyRes[5].ProdVolumeValue)).ToString("n");

            Sheet.Cells["E18"].Value = (((dailyRes[6].ProdVolumeValue * dailyRes[6].VQIValue)
                                        + (dailyRes[7].ProdVolumeValue * dailyRes[7].VQIValue)
                                        + (dailyRes[8].ProdVolumeValue * dailyRes[8].VQIValue))
                                        / (dailyRes[6].ProdVolumeValue + dailyRes[7].ProdVolumeValue + dailyRes[8].ProdVolumeValue)).ToString("n");

            Sheet.Cells["F18"].Value = wtdModel.VQI_WTD;
            Sheet.Cells["G18"].Value = mtdModel.VQI_MTD;
            Sheet.Cells["H18"].Value = mtg_VQI.ToString("n");
            Sheet.Cells["I18"].Value = target_VQI.ToString("n");

            dateColumn = 3;
            for (int j = 0; j < dailyRes.Count(); j++)
            {
                if (j == 0 || j % 3 == 0)
                {
                    if (dailyRes[j].Shift == "1")
                    {
                        string columnFormat = Helper.Number2String(dateColumn);
                        Sheet.Cells[string.Format(columnFormat, 19)].Value = dailyRes[j].VQIValue.ToString("n");
                        dateColumn += 1;
                    }
                }
            }
            dateColumn = 3;
            for (int k = 0; k < dailyRes.Count(); k++)
            {
                if (k == 1 || k % 3 == 1)
                {
                    if (dailyRes[k].Shift == "2")
                    {
                        string columnFormat = Helper.Number2String(dateColumn);
                        Sheet.Cells[string.Format(columnFormat, 20)].Value = dailyRes[k].VQIValue.ToString("n");
                        dateColumn += 1;
                    }
                }
            }
            dateColumn = 3;
            for (int l = 0; l < dailyRes.Count(); l++)
            {
                if (l == 2 || l % 3 == 2)
                {
                    if (dailyRes[l].Shift == "3")
                    {
                        string columnFormat = Helper.Number2String(dateColumn);
                        Sheet.Cells[string.Format(columnFormat, 21)].Value = dailyRes[l].VQIValue.ToString("n");
                        dateColumn += 1;
                    }
                }
            }

            #endregion

            #region Working Time

            Sheet.Cells["A22"].Style.Font.Bold = true;
            Sheet.Cells["A22"].Value = "Working Time";
            Sheet.Cells["A23:A25"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            Sheet.Cells["A23"].Value = "Shift 1";
            Sheet.Cells["A24"].Value = "Shift 2";
            Sheet.Cells["A25"].Value = "Shift 3";

            Sheet.Cells["C22"].Value = (dailyRes[0].WorkingValue + dailyRes[1].WorkingValue + dailyRes[2].WorkingValue).ToString("n");
            Sheet.Cells["D22"].Value = (dailyRes[3].WorkingValue + dailyRes[4].WorkingValue + dailyRes[5].WorkingValue).ToString("n");
            Sheet.Cells["E22"].Value = (dailyRes[6].WorkingValue + dailyRes[7].WorkingValue + dailyRes[8].WorkingValue).ToString("n");

            Sheet.Cells["F22"].Value = wtdModel.Working_WTD;
            Sheet.Cells["G22"].Value = mtdModel.Working_MTD;
            Sheet.Cells["H22"].Value = ((target_Working - Convert.ToDecimal(mtdModel.Working_MTD)).ToString("n"));
            Sheet.Cells["I22"].Value = target_Working.ToString("n");

            dateColumn = 3;
            for (int j = 0; j < dailyRes.Count(); j++)
            {
                if (j == 0 || j % 3 == 0)
                {
                    if (dailyRes[j].Shift == "1")
                    {
                        string columnFormat = Helper.Number2String(dateColumn);
                        Sheet.Cells[string.Format(columnFormat, 23)].Value = dailyRes[j].WorkingValue.ToString("n");
                        dateColumn += 1;
                    }
                }
            }
            dateColumn = 3;
            for (int k = 0; k < dailyRes.Count(); k++)
            {
                if (k == 1 || k % 3 == 1)
                {
                    if (dailyRes[k].Shift == "2")
                    {
                        string columnFormat = Helper.Number2String(dateColumn);
                        Sheet.Cells[string.Format(columnFormat, 24)].Value = dailyRes[k].WorkingValue.ToString("n");
                        dateColumn += 1;
                    }
                }
            }
            dateColumn = 3;
            for (int l = 0; l < dailyRes.Count(); l++)
            {
                if (l == 2 || l % 3 == 2)
                {
                    if (dailyRes[l].Shift == "3")
                    {
                        string columnFormat = Helper.Number2String(dateColumn);
                        Sheet.Cells[string.Format(columnFormat, 25)].Value = dailyRes[l].WorkingValue.ToString("n");
                        dateColumn += 1;
                    }
                }
            }
            #endregion

            #region Uptime

            Sheet.Cells["A26"].Style.Font.Bold = true;
            Sheet.Cells["A26"].Value = "Uptime";
            Sheet.Cells["A27:A29"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            Sheet.Cells["A27"].Value = "Shift 1";
            Sheet.Cells["A28"].Value = "Shift 2";
            Sheet.Cells["A29"].Value = "Shift 3";

            Sheet.Cells["C26"].Value = (((dailyRes[0].ProdVolumeValue + dailyRes[1].ProdVolumeValue + dailyRes[2].ProdVolumeValue)
                                    / (((dailyRes[0].ProdVolumeValue / dailyRes[0].UptimeValue) * 100)
                                    + ((dailyRes[1].ProdVolumeValue / dailyRes[1].UptimeValue) * 100)
                                    + ((dailyRes[2].ProdVolumeValue / dailyRes[2].UptimeValue) * 100))) * 100).ToString("n");

            Sheet.Cells["D26"].Value = (((dailyRes[3].ProdVolumeValue + dailyRes[4].ProdVolumeValue + dailyRes[5].ProdVolumeValue)
                                    / (((dailyRes[3].ProdVolumeValue / dailyRes[3].UptimeValue) * 100)
                                    + ((dailyRes[4].ProdVolumeValue / dailyRes[4].UptimeValue) * 100)
                                    + ((dailyRes[5].ProdVolumeValue / dailyRes[5].UptimeValue) * 100))) * 100).ToString("n");

            Sheet.Cells["E26"].Value = (((dailyRes[6].ProdVolumeValue + dailyRes[7].ProdVolumeValue + dailyRes[8].ProdVolumeValue)
                                    / (((dailyRes[6].ProdVolumeValue / dailyRes[6].UptimeValue) * 100)
                                    + ((dailyRes[7].ProdVolumeValue / dailyRes[7].UptimeValue) * 100)
                                    + ((dailyRes[8].ProdVolumeValue / dailyRes[8].UptimeValue) * 100))) * 100).ToString("n");

            Sheet.Cells["F26"].Value = wtdModel.Uptime_WTD;
            Sheet.Cells["G26"].Value = mtdModel.Uptime_MTD;
            Sheet.Cells["H26"].Value = mtg_Uptime.ToString("n");
            Sheet.Cells["I26"].Value = target_Uptime.ToString("n");

            dateColumn = 3;
            for (int j = 0; j < dailyRes.Count(); j++)
            {
                if (j == 0 || j % 3 == 0)
                {
                    if (dailyRes[j].Shift == "1")
                    {
                        string columnFormat = Helper.Number2String(dateColumn);
                        Sheet.Cells[string.Format(columnFormat, 27)].Value = dailyRes[j].UptimeValue.ToString("n");
                        dateColumn += 1;
                    }
                }
            }
            dateColumn = 3;
            for (int k = 0; k < dailyRes.Count(); k++)
            {
                if (k == 1 || k % 3 == 1)
                {
                    if (dailyRes[k].Shift == "2")
                    {
                        string columnFormat = Helper.Number2String(dateColumn);
                        Sheet.Cells[string.Format(columnFormat, 28)].Value = dailyRes[k].UptimeValue.ToString("n");
                        dateColumn += 1;
                    }
                }
            }
            dateColumn = 3;
            for (int l = 0; l < dailyRes.Count(); l++)
            {
                if (l == 2 || l % 3 == 2)
                {
                    if (dailyRes[l].Shift == "3")
                    {
                        string columnFormat = Helper.Number2String(dateColumn);
                        Sheet.Cells[string.Format(columnFormat, 29)].Value = dailyRes[l].UptimeValue.ToString("n");
                        dateColumn += 1;
                    }
                }
            }
            #endregion

            #region STRS

            Sheet.Cells["A30"].Style.Font.Bold = true;
            Sheet.Cells["A30"].Value = "STRS";
            Sheet.Cells["A31:A33"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            Sheet.Cells["A31"].Value = "Shift 1";
            Sheet.Cells["A32"].Value = "Shift 2";
            Sheet.Cells["A33"].Value = "Shift 3";


            Sheet.Cells["C30"].Value = (((dailyRes[0].ProdVolumeValue * dailyRes[0].STRSValue)
                                       + (dailyRes[1].ProdVolumeValue * dailyRes[1].STRSValue)
                                       + (dailyRes[2].ProdVolumeValue * dailyRes[2].STRSValue))
                                       / (dailyRes[0].ProdVolumeValue + dailyRes[1].ProdVolumeValue + dailyRes[2].ProdVolumeValue)).ToString("n");

            Sheet.Cells["D30"].Value = (((dailyRes[3].ProdVolumeValue * dailyRes[3].STRSValue)
                                        + (dailyRes[4].ProdVolumeValue * dailyRes[4].STRSValue)
                                        + (dailyRes[5].ProdVolumeValue * dailyRes[5].STRSValue))
                                        / (dailyRes[3].ProdVolumeValue + dailyRes[4].ProdVolumeValue + dailyRes[5].ProdVolumeValue)).ToString("n");

            Sheet.Cells["E30"].Value = (((dailyRes[6].ProdVolumeValue * dailyRes[6].STRSValue)
                                        + (dailyRes[7].ProdVolumeValue * dailyRes[7].STRSValue)
                                        + (dailyRes[8].ProdVolumeValue * dailyRes[8].STRSValue))
                                        / (dailyRes[6].ProdVolumeValue + dailyRes[7].ProdVolumeValue + dailyRes[8].ProdVolumeValue)).ToString("n");

            Sheet.Cells["F30"].Value = wtdModel.STRS_WTD;
            Sheet.Cells["G30"].Value = mtdModel.STRS_MTD;
            Sheet.Cells["H30"].Value = mtg_STRS.ToString("n");
            Sheet.Cells["I30"].Value = target_STRS.ToString("n");

            dateColumn = 3;
            for (int j = 0; j < dailyRes.Count(); j++)
            {
                if (j == 0 || j % 3 == 0)
                {
                    if (dailyRes[j].Shift == "1")
                    {
                        string columnFormat = Helper.Number2String(dateColumn);
                        Sheet.Cells[string.Format(columnFormat, 31)].Value = dailyRes[j].STRSValue.ToString("n");
                        dateColumn += 1;
                    }
                }
            }
            dateColumn = 3;
            for (int k = 0; k < dailyRes.Count(); k++)
            {
                if (k == 1 || k % 3 == 1)
                {
                    if (dailyRes[k].Shift == "2")
                    {
                        string columnFormat = Helper.Number2String(dateColumn);
                        Sheet.Cells[string.Format(columnFormat, 32)].Value = dailyRes[k].STRSValue.ToString("n");
                        dateColumn += 1;
                    }
                }
            }
            dateColumn = 3;
            for (int l = 0; l < dailyRes.Count(); l++)
            {
                if (l == 2 || l % 3 == 2)
                {
                    if (dailyRes[l].Shift == "3")
                    {
                        string columnFormat = Helper.Number2String(dateColumn);
                        Sheet.Cells[string.Format(columnFormat, 33)].Value = dailyRes[l].STRSValue.ToString("n");
                        dateColumn += 1;
                    }
                }
            }
            #endregion

            #region Production Volume

            Sheet.Cells["A34"].Style.Font.Bold = true;
            Sheet.Cells["A34"].AutoFitColumns();
            Sheet.Cells["A34"].Value = "Production Volume";
            Sheet.Cells["A35:A37"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            Sheet.Cells["A35"].Value = "Shift 1";
            Sheet.Cells["A36"].Value = "Shift 2";
            Sheet.Cells["A37"].Value = "Shift 3";

            Sheet.Cells["C34"].Value = (dailyRes[0].ProdVolumeValue + dailyRes[1].ProdVolumeValue + dailyRes[2].ProdVolumeValue).ToString("n");
            Sheet.Cells["D34"].Value = (dailyRes[3].ProdVolumeValue + dailyRes[4].ProdVolumeValue + dailyRes[5].ProdVolumeValue).ToString("n");
            Sheet.Cells["E34"].Value = (dailyRes[6].ProdVolumeValue + dailyRes[7].ProdVolumeValue + dailyRes[8].ProdVolumeValue).ToString("n");

            Sheet.Cells["F34"].Value = wtdModel.ProdVolume_WTD;
            Sheet.Cells["G34"].Value = mtdModel.ProdVolume_MTD;
            Sheet.Cells["H34"].Value = ((target_ProdVol - Convert.ToDecimal(mtdModel.ProdVolume_MTD)).ToString("n"));

            dateColumn = 3;
            for (int j = 0; j < dailyRes.Count(); j++)
            {
                if (j == 0 || j % 3 == 0)
                {
                    if (dailyRes[j].Shift == "1")
                    {
                        string columnFormat = Helper.Number2String(dateColumn);
                        Sheet.Cells[string.Format(columnFormat, 35)].Value = dailyRes[j].ProdVolumeValue.ToString("n");
                        dateColumn += 1;
                    }
                }
            }
            dateColumn = 3;
            for (int k = 0; k < dailyRes.Count(); k++)
            {
                if (k == 1 || k % 3 == 1)
                {
                    if (dailyRes[k].Shift == "2")
                    {
                        string columnFormat = Helper.Number2String(dateColumn);
                        Sheet.Cells[string.Format(columnFormat, 36)].Value = dailyRes[k].ProdVolumeValue.ToString("n");
                        dateColumn += 1;
                    }
                }
            }
            dateColumn = 3;
            for (int l = 0; l < dailyRes.Count(); l++)
            {
                if (l == 2 || l % 3 == 2)
                {
                    if (dailyRes[l].Shift == "3")
                    {
                        string columnFormat = Helper.Number2String(dateColumn);
                        Sheet.Cells[string.Format(columnFormat, 37)].Value = dailyRes[l].ProdVolumeValue.ToString("n");
                        dateColumn += 1;
                    }
                }
            }

            Sheet.Cells["I34"].Value = target_ProdVol.ToString("n");
            #endregion

            #region CRR

            Sheet.Cells["A38"].Style.Font.Bold = true;
            Sheet.Cells["A38"].Value = "CRR";
            Sheet.Cells["A39:A41"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            Sheet.Cells["A39"].Value = "Shift 1";
            Sheet.Cells["A40"].Value = "Shift 2";
            Sheet.Cells["A41"].Value = "Shift 3";

            Sheet.Cells["C38"].Value = (((dailyRes[0].ProdVolumeValue * dailyRes[0].CRRValue)
                                       + (dailyRes[1].ProdVolumeValue * dailyRes[1].CRRValue)
                                       + (dailyRes[2].ProdVolumeValue * dailyRes[2].CRRValue))
                                       / (dailyRes[0].ProdVolumeValue + dailyRes[1].ProdVolumeValue + dailyRes[2].ProdVolumeValue)).ToString("n");

            Sheet.Cells["D38"].Value = (((dailyRes[3].ProdVolumeValue * dailyRes[3].CRRValue)
                                        + (dailyRes[4].ProdVolumeValue * dailyRes[4].CRRValue)
                                        + (dailyRes[5].ProdVolumeValue * dailyRes[5].CRRValue))
                                        / (dailyRes[3].ProdVolumeValue + dailyRes[4].ProdVolumeValue + dailyRes[5].ProdVolumeValue)).ToString("n");

            Sheet.Cells["E38"].Value = (((dailyRes[6].ProdVolumeValue * dailyRes[6].CRRValue)
                                        + (dailyRes[7].ProdVolumeValue * dailyRes[7].CRRValue)
                                        + (dailyRes[8].ProdVolumeValue * dailyRes[8].CRRValue))
                                        / (dailyRes[6].ProdVolumeValue + dailyRes[7].ProdVolumeValue + dailyRes[8].ProdVolumeValue)).ToString("n");

            Sheet.Cells["F38"].Value = wtdModel.CRR_WTD;
            Sheet.Cells["G38"].Value = mtdModel.CRR_MTD;
            Sheet.Cells["H38"].Value = mtg_CRR.ToString("n");
            Sheet.Cells["I38"].Value = target_CRR.ToString("n");

            dateColumn = 3;
            for (int j = 0; j < dailyRes.Count(); j++)
            {
                if (j == 0 || j % 3 == 0)
                {
                    if (dailyRes[j].Shift == "1")
                    {
                        string columnFormat = Helper.Number2String(dateColumn);
                        Sheet.Cells[string.Format(columnFormat, 39)].Value = dailyRes[j].CRRValue.ToString("n");
                        dateColumn += 1;
                    }
                }
            }
            dateColumn = 3;
            for (int k = 0; k < dailyRes.Count(); k++)
            {
                if (k == 1 || k % 3 == 1)
                {
                    if (dailyRes[k].Shift == "2")
                    {
                        string columnFormat = Helper.Number2String(dateColumn);
                        Sheet.Cells[string.Format(columnFormat, 40)].Value = dailyRes[k].CRRValue.ToString("n");
                        dateColumn += 1;
                    }
                }
            }
            dateColumn = 3;
            for (int l = 0; l < dailyRes.Count(); l++)
            {
                if (l == 2 || l % 3 == 2)
                {
                    if (dailyRes[l].Shift == "3")
                    {
                        string columnFormat = Helper.Number2String(dateColumn);
                        Sheet.Cells[string.Format(columnFormat, 41)].Value = dailyRes[l].CRRValue.ToString("n");
                        dateColumn += 1;
                    }
                }
            }
            #endregion

            #region Uptime Segment

            string Uptime_Color = "#78c72a";
            for (int i = 0; i < uptimeRes.Count(); i++)
            {
                if (uptimeRes[i].Shift == "1")
                {
                    Color S1 = ColorTranslator.FromHtml(Mach_Color);
                    Sheet.Cells["L9"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    Sheet.Cells["L9"].Style.Fill.BackgroundColor.SetColor(S1);
                    Sheet.Cells["L9"].Style.Font.Bold = true;
                    Sheet.Cells["L9"].Value = "Shift 1";
                    Sheet.Cells["L9:M9"].Merge = true;

                    Color Up1 = ColorTranslator.FromHtml(Uptime_Color);
                    Sheet.Cells["L10"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    Sheet.Cells["L10"].Style.Fill.BackgroundColor.SetColor(Up1);
                    Sheet.Cells["L10"].Value = "Uptime";
                    Sheet.Cells["L10:M10"].Merge = true;

                    Sheet.Cells["L11"].Value = uptimeRes[i].UptimeFocus;
                    Sheet.Cells["M11"].Value = uptimeRes[i].UptimeActPlan;
                    Sheet.Cells["L11:L13"].Merge = true;
                    Sheet.Cells["M11:M13"].Merge = true;

                    Sheet.Column(12).Width = 25;
                    Sheet.Column(13).Width = 25;
                }

                if (uptimeRes[i].Shift == "2")
                {
                    Color S2 = ColorTranslator.FromHtml(Mach_Color);
                    Sheet.Cells["O9"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    Sheet.Cells["O9"].Style.Fill.BackgroundColor.SetColor(S2);
                    Sheet.Cells["O9"].Style.Font.Bold = true;
                    Sheet.Cells["O9"].Value = "Shift 2";
                    Sheet.Cells["O9:P9"].Merge = true;

                    Color Up2 = ColorTranslator.FromHtml(Uptime_Color);
                    Sheet.Cells["O10"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    Sheet.Cells["O10"].Style.Fill.BackgroundColor.SetColor(Up2);
                    Sheet.Cells["O10"].Value = "Uptime";
                    Sheet.Cells["O10:P10"].Merge = true;

                    Sheet.Cells["O11"].Value = uptimeRes[i].UptimeFocus;
                    Sheet.Cells["P11"].Value = uptimeRes[i].UptimeActPlan;
                    Sheet.Cells["O11:O13"].Merge = true;
                    Sheet.Cells["P11:P13"].Merge = true;

                    Sheet.Column(15).Width = 25;
                    Sheet.Column(16).Width = 25;
                }

                if (uptimeRes[i].Shift == "3")
                {
                    Color S2 = ColorTranslator.FromHtml(Mach_Color);
                    Sheet.Cells["R9"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    Sheet.Cells["R9"].Style.Fill.BackgroundColor.SetColor(S2);
                    Sheet.Cells["R9"].Style.Font.Bold = true;
                    Sheet.Cells["R9"].Value = "Shift 3";
                    Sheet.Cells["R9:S9"].Merge = true;

                    Color Up2 = ColorTranslator.FromHtml(Uptime_Color);
                    Sheet.Cells["R10"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    Sheet.Cells["R10"].Style.Fill.BackgroundColor.SetColor(Up2);
                    Sheet.Cells["R10"].Value = "Uptime";
                    Sheet.Cells["R10:S10"].Merge = true;

                    Sheet.Cells["R11"].Value = uptimeRes[i].UptimeFocus;
                    Sheet.Cells["S11"].Value = uptimeRes[i].UptimeActPlan;
                    Sheet.Cells["R11:R13"].Merge = true;
                    Sheet.Cells["S11:S13"].Merge = true;

                    Sheet.Column(18).Width = 25;
                    Sheet.Column(19).Width = 25;
                }
            }

            #endregion

            Sheet.Cells["A:I"].AutoFitColumns();

            return Ep.GetAsByteArray();
        }
        #endregion

        #region ::Downtime Export CSV ::
        public static byte[] ExportDowntimeCSV(string username)
        {
            ExcelPackage Ep = new ExcelPackage();
            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Downtime");

            AddHeader(Sheet, "Downtime", username);

            return Ep.GetAsByteArray();
        }
        #endregion

        #region ::Extract Raw Data SP::
        public static byte[] RawDataExtract(string username, string type, string startDate, string endDate, ArrayList myList)
        {
            int rowNext = 0;
            int nextExt = 2;
            ExcelPackage Ep = new ExcelPackage();
            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add(type);

            int nextIndex = 0;

            object objStart = myList[0];
            var itemStart = objStart as IEnumerable<string>;
            var compStart = itemStart.ToList();

            List<string> headStart = new List<string>();
            foreach (var res in compStart)
            {
                headStart.Add(res);
            }
            nextIndex = headStart.Count();

            for (int i = 0; i < myList.Count; i++)
            {
                if (i == 0)//component header
                {
                    object resComp = myList[i];
                    var itemComp = resComp as IEnumerable<string>;
                    var compList = itemComp.ToList();

                    List<string> head = new List<string>();
                    foreach (var res in compList)
                    {
                        head.Add(res);
                    }
                    int columnStartComp = 1;
                    foreach (var item in head)
                    {
                        string columnFormatComp = Helper.Number2String(columnStartComp);
                        Sheet.Cells[string.Format(columnFormatComp, 1)].Value = item;
                        Sheet.Cells[string.Format(columnFormatComp, 1)].Style.Font.Bold = true;
                        Sheet.Cells[string.Format(columnFormatComp, 1)].AutoFitColumns();
                        columnStartComp += 1;
                    }
                }
                if (i == 1 || i % 3 == 1)// component value
                {
                    object resValue = myList[i];
                    var items = resValue as IEnumerable<string>;
                    var valueList = items.ToList();

                    List<string> valHead = new List<string>();
                    foreach (var res in valueList)
                    {
                        valHead.Add(res);
                    }
                    int columnStartVal = 1;
                    foreach (var item in valHead)
                    {
                        string columnFormatVal = Helper.Number2String(columnStartVal);
                        Sheet.Cells[string.Format(columnFormatVal, 2 + rowNext)].Value = item;
                        Sheet.Cells[string.Format(columnFormatVal, 2 + rowNext)].AutoFitColumns();
                        columnStartVal += 1;
                    }
                    rowNext += 4;
                }
                if (i == 2 || i % 3 == 2)//extras
                {
                    dynamic temp = myList[i];

                    List<Extra> extColl = new List<Extra>();

                    for (int k = 0; k < temp.Count; k++)
                    {
                        var item = temp[k];

                        Extra itemEx = new Extra();
                        itemEx.ID = item.ID;
                        itemEx.Header = item.Header;
                        itemEx.Field = item.Field;
                        itemEx.Value = item.Value;
                        extColl.Add(itemEx);
                    }

                    var groupCollect = extColl.GroupBy(x => x.Field).ToList();
                    int columnStartField = nextIndex + 1;
                    int columnStartValue = nextIndex + 1;

                    foreach (var resExt in groupCollect)
                    {
                        string columnFormatComp = Helper.Number2String(columnStartField);
                        Sheet.Cells[string.Format(columnFormatComp, 1)].Value = resExt.Key;
                        Sheet.Cells[string.Format(columnFormatComp, 1)].Style.Font.Bold = true;
                        Sheet.Cells[string.Format(columnFormatComp, 1)].AutoFitColumns();

                        for (int cVal = 0; cVal < resExt.Count(); cVal++)
                        {
                            Sheet.Cells[string.Format(columnFormatComp, nextExt + cVal)].Value = resExt.ElementAt(cVal).Value;
                            Sheet.Cells[string.Format(columnFormatComp, nextExt + cVal)].AutoFitColumns();
                        }
                        columnStartField += 1;
                    }
                    nextExt += 4;
                }
            }
            return Ep.GetAsByteArray();
        }

        public class Extra
        {
            public long ID { get; set; }
            public string Header { get; set; }
            public string Field { get; set; }
            public string Value { get; set; }
        }
        #endregion

        #region ::Extract Raw PP::
        public static byte[] PPRawDataExtract(Dictionary<string, List<List<string>>> rows, Dictionary<string, List<string>> summaries = null)
        {
            summaries = summaries ?? new Dictionary<string, List<string>>();
            using (ExcelPackage Ep = new ExcelPackage())
            {
                foreach (KeyValuePair<string, List<List<string>>> row in rows)
                {
                    Color colFromHex = ColorTranslator.FromHtml(HEADER_COLOR);
                    ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add(row.Key);
                    for (int j = 0; j < row.Value.Count(); j++)
                        for (int k = 0; k < row.Value[j].Count(); k++)
                            Sheet.Cells[j + 1, k + 1].Value = row.Value[j][k];
                    if (summaries.ContainsKey(row.Key))
                    {
                        int lastRow = row.Value.Count() + 1;
                        for (int j = 0; j < summaries[row.Key].Count(); j++)
                            Sheet.Cells[lastRow, j + 1].Value = summaries[row.Key][j];
                        string footer = "A" + lastRow + ":" + ExcelCellAddress.GetColumnLetter(Sheet.Dimension.End.Column) + lastRow;
                        Sheet.Cells[footer].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        Sheet.Cells[footer].Style.Font.Bold = true;
                        Sheet.Cells[footer].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    }
                    if (row.Value.Count() > 0 && row.Value[0].Count() > 0)
                    {
                        string header = Sheet.Dimension.Start.Address + ":" + ExcelCellAddress.GetColumnLetter(Sheet.Dimension.End.Column) + "1";
                        Sheet.Cells[Sheet.Dimension.Address].AutoFitColumns();
                        Sheet.Cells[Sheet.Dimension.Address].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        Sheet.Cells[header].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        Sheet.Cells[header].Style.Font.Bold = true;
                        Sheet.Cells[header].Style.Fill.BackgroundColor.SetColor(colFromHex);
                    }
                }
                return Ep.GetAsByteArray();
            }
        }
        #endregion

        #region ::Input Target ::
        public static byte[] ExportInputTarget(string username, string version, string kpi, string month, List<SelectListItem> kpiList)
        {

            ExcelPackage Ep = new ExcelPackage();
            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("InputTarget");

            AddHeader(Sheet, "Input Target", username);
            Sheet.Cells["A7:AF7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            Sheet.Cells["A7"].Value = "No.";
            Sheet.Cells["B7"].Value = "Version";
            Sheet.Cells["C7"].Value = "Month";
            Sheet.Cells["D7"].Value = "KPI";
            Sheet.Cells["E7"].Value = "Value";

            Color colFromHex = ColorTranslator.FromHtml(HEADER_COLOR);
            using (var range = Sheet.Cells[7, 1, 7, 5])
            {
                range.Style.Font.Bold = true;
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(colFromHex);
                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            int row = 8;


            List<SelectListItem> listKPI = new List<SelectListItem>();
            if (kpi == "All")
            {
                listKPI = kpiList;
            }
            else
            {
                listKPI.Add(new SelectListItem { Value = kpi });
            }
            List<SelectListItem> listMonth = new List<SelectListItem>();
            if (month == "All")
            {
                for (int i = 1; i <= 12; i++)
                {
                    listMonth.Add(new SelectListItem { Value = getMonthByNumber(i) });
                }
            }
            else
            {
                listMonth.Add(new SelectListItem { Value = month });
            }
            int num = 1;
            foreach (SelectListItem sli in listMonth)
            {
                foreach (SelectListItem sli2 in listKPI)
                {

                    Sheet.Cells[string.Format("A{0}", row)].Value = num++;
                    Sheet.Cells[string.Format("A{0}", row)].AutoFitColumns();
                    Sheet.Cells[string.Format("B{0}", row)].Value = version;
                    Sheet.Cells[string.Format("B{0}", row)].AutoFitColumns();
                    Sheet.Cells[string.Format("C{0}", row)].Value = sli.Value;
                    Sheet.Cells[string.Format("C{0}", row)].AutoFitColumns();
                    Sheet.Cells[string.Format("D{0}", row++)].Value = sli2.Value;
                }
            }
            Sheet.Cells[7, 1, row - 1, 5].Style.Border.BorderAround(ExcelBorderStyle.Thin);


            /*
						ExcelPackage Ep = new ExcelPackage();
						ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("InputTarget");

						AddHeader(Sheet, "Input Target", username);
						Sheet.Cells["A7:AF7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

						Sheet.Cells["A7"].Value = "PC";
						Sheet.Cells["B7"].Value = "KPI";
						Sheet.Cells["C7"].Value = "Version";
						Sheet.Cells["D7"].Value = "Year";
						Sheet.Cells["E7"].Value = "Value";

						int columnStart = 4;
						int columnLast = 16;
						int row = 7;

						if (month != "All")
						{
							Sheet.Cells["D7"].Value = month;
						}

						if (month == "All")
						{
							for (int i = 1; i <= 12; i++)
							{
								string columnFormat = Helper.Number2String(columnStart);
								Sheet.Cells[string.Format(columnFormat, row)].Value = getMonthByNumber(i);
								columnStart += 1;
							}
							string columnFormatLast = Helper.Number2String(columnLast);
						}

						if (pc == "-")
						{
							if (kpi != "All")
							{
								if (version != "All")
								{
									Sheet.Cells["A8"].Value = pcList.ElementAt(0).Text;
									Sheet.Cells["A8"].AutoFitColumns();
									Sheet.Cells["A9"].Value = pcList.ElementAt(1).Text;
									Sheet.Cells["A9"].AutoFitColumns();
									Sheet.Cells["A10"].Value = pcList.ElementAt(2).Text;
									Sheet.Cells["A10"].AutoFitColumns();
									Sheet.Cells["A11"].Value = pcList.ElementAt(3).Text;
									Sheet.Cells["A11"].AutoFitColumns();

									for (int i = 8; i <= 11; i++)
									{
										Sheet.Cells[string.Format("B{0}", i)].Value = kpi;
										Sheet.Cells[string.Format("B{0}", i)].AutoFitColumns();
									}
									for (int j = 8; j <= 11; j++)
									{
										Sheet.Cells[string.Format("C{0}", j)].Value = version;
										Sheet.Cells[string.Format("C{0}", j)].AutoFitColumns();
									}

								}
								if (version == "All")
								{
									int t = 8;
									int elemPC = 0;
									foreach (var item in pcList)
									{

										int elemVersion = 0;
										foreach (var itemV in versionList)
										{
											Sheet.Cells[string.Format("A{0}", t)].Value = pcList.ElementAt(elemPC).Text;
											Sheet.Cells[string.Format("A{0}", t)].AutoFitColumns();
											Sheet.Cells[string.Format("B{0}", t)].Value = kpi;
											Sheet.Cells[string.Format("B{0}", t)].AutoFitColumns();
											Sheet.Cells[string.Format("C{0}", t)].Value = versionList.ElementAt(elemVersion).Text;
											Sheet.Cells[string.Format("C{0}", t)].AutoFitColumns();
											elemVersion++;
											t++;
										}

										elemPC++;
									}
								}
							}
							if (kpi == "All")
							{
								if (version == "All")
								{
									int t = 8;
									int elemPC = 0;
									foreach (var item in pcList)
									{
										int elemKPI = 0;
										foreach (var itemS in kpiList)
										{
											int elemVersion = 0;
											foreach (var itemV in versionList)
											{
												Sheet.Cells[string.Format("A{0}", t)].Value = pcList.ElementAt(elemPC).Text;
												Sheet.Cells[string.Format("A{0}", t)].AutoFitColumns();
												Sheet.Cells[string.Format("B{0}", t)].Value = kpiList.ElementAt(elemKPI).Text;
												Sheet.Cells[string.Format("B{0}", t)].AutoFitColumns();
												Sheet.Cells[string.Format("C{0}", t)].Value = versionList.ElementAt(elemVersion).Text;
												Sheet.Cells[string.Format("C{0}", t)].AutoFitColumns();
												elemVersion++;
												t++;
											}
											elemKPI++;
										}
										elemPC++;
									}
								}
								if (version != "All") // OK
								{
									int t = 8;
									int elemPC = 0;

									foreach (var item in pcList)
									{
										int elemKPI = 0;
										foreach (var itemS in kpiList)
										{
											Sheet.Cells[string.Format("A{0}", t)].Value = pcList.ElementAt(elemPC).Text;
											Sheet.Cells[string.Format("A{0}", t)].AutoFitColumns();
											Sheet.Cells[string.Format("B{0}", t)].Value = kpiList.ElementAt(elemKPI).Text;
											Sheet.Cells[string.Format("B{0}", t)].AutoFitColumns();
											Sheet.Cells[string.Format("C{0}", t)].Value = version;
											Sheet.Cells[string.Format("C{0}", t)].AutoFitColumns();
											elemKPI++;
											t++;
										}
										elemPC++;
									}
								}
							}
						}
						if (pc != "-")
						{
							if (kpi != "All")
							{
								if (version != "All")
								{
									Sheet.Cells["A8"].Value = pc;
									Sheet.Cells["A8"].AutoFitColumns();
									Sheet.Cells["B8"].Value = kpi;
									Sheet.Cells["B8"].AutoFitColumns();
									Sheet.Cells["C8"].Value = version;
									Sheet.Cells["C8"].AutoFitColumns();
								}
								if (version == "All")
								{
									int t = 8;
									int elemVersion = 0;
									foreach (var itemV in versionList)
									{
										Sheet.Cells[string.Format("A{0}", t)].Value = pc;
										Sheet.Cells[string.Format("A{0}", t)].AutoFitColumns();
										Sheet.Cells[string.Format("B{0}", t)].Value = kpi;
										Sheet.Cells[string.Format("B{0}", t)].AutoFitColumns();
										Sheet.Cells[string.Format("C{0}", t)].Value = versionList.ElementAt(elemVersion).Text;
										Sheet.Cells[string.Format("C{0}", t)].AutoFitColumns();
										elemVersion++;
										t++;
									}
								}
							}
							if (kpi == "All")
							{
								if (version == "All")
								{
									int t = 8;
									int elemKPI = 0;
									foreach (var itemS in kpiList)
									{
										int elemVersion = 0;
										foreach (var itemV in versionList)
										{
											Sheet.Cells[string.Format("A{0}", t)].Value = pc;
											Sheet.Cells[string.Format("A{0}", t)].AutoFitColumns();
											Sheet.Cells[string.Format("B{0}", t)].Value = kpiList.ElementAt(elemKPI).Text;
											Sheet.Cells[string.Format("B{0}", t)].AutoFitColumns();
											Sheet.Cells[string.Format("C{0}", t)].Value = versionList.ElementAt(elemVersion).Text;
											Sheet.Cells[string.Format("C{0}", t)].AutoFitColumns();
											elemVersion++;
											t++;
										}
										elemKPI++;
									}
								}
								if (version != "All")
								{
									int t = 8;
									int elemKPI = 0;
									foreach (var itemS in kpiList)
									{

										Sheet.Cells[string.Format("A{0}", t)].Value = pc;
										Sheet.Cells[string.Format("A{0}", t)].AutoFitColumns();
										Sheet.Cells[string.Format("B{0}", t)].Value = kpiList.ElementAt(elemKPI).Text;
										Sheet.Cells[string.Format("B{0}", t)].AutoFitColumns();
										Sheet.Cells[string.Format("C{0}", t)].Value = version;
										Sheet.Cells[string.Format("C{0}", t)].AutoFitColumns();

										t++;
										elemKPI++;
									}
								}
							}
						}
						*/
            return Ep.GetAsByteArray();
        }

        public static byte[] ExportInputTarget(string username, string version, string kpi, string pc, string month, List<SelectListItem> pcList, List<SelectListItem> kpiList, List<SelectListItem> versionList)
        {
            ExcelPackage Ep = new ExcelPackage();
            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("InputTarget");

            AddHeader(Sheet, "Input Target", username);
            Sheet.Cells["A7:AF7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            Sheet.Cells["A7"].Value = "PC";
            Sheet.Cells["B7"].Value = "KPI";
            Sheet.Cells["C7"].Value = "Version";

            int columnStart = 4;
            int columnLast = 16;
            int row = 7;

            if (month != "All")
            {
                Sheet.Cells["D7"].Value = month;
            }

            if (month == "All")
            {
                for (int i = 1; i <= 12; i++)
                {
                    string columnFormat = Helper.Number2String(columnStart);
                    Sheet.Cells[string.Format(columnFormat, row)].Value = getMonthByNumber(i);
                    columnStart += 1;
                }
                string columnFormatLast = Helper.Number2String(columnLast);
                Sheet.Cells[string.Format(columnFormatLast, row)].Value = "Year";
            }

            if (pc == "-")
            {
                if (kpi != "All")
                {
                    if (version != "All")
                    {
                        Sheet.Cells["A8"].Value = pcList.ElementAt(0).Text;
                        Sheet.Cells["A8"].AutoFitColumns();
                        Sheet.Cells["A9"].Value = pcList.ElementAt(1).Text;
                        Sheet.Cells["A9"].AutoFitColumns();
                        Sheet.Cells["A10"].Value = pcList.ElementAt(2).Text;
                        Sheet.Cells["A10"].AutoFitColumns();
                        Sheet.Cells["A11"].Value = pcList.ElementAt(3).Text;
                        Sheet.Cells["A11"].AutoFitColumns();

                        for (int i = 8; i <= 11; i++)
                        {
                            Sheet.Cells[string.Format("B{0}", i)].Value = kpi;
                            Sheet.Cells[string.Format("B{0}", i)].AutoFitColumns();
                        }
                        for (int j = 8; j <= 11; j++)
                        {
                            Sheet.Cells[string.Format("C{0}", j)].Value = version;
                            Sheet.Cells[string.Format("C{0}", j)].AutoFitColumns();
                        }

                    }
                    if (version == "All")
                    {
                        int t = 8;
                        int elemPC = 0;
                        foreach (var item in pcList)
                        {

                            int elemVersion = 0;
                            foreach (var itemV in versionList)
                            {
                                Sheet.Cells[string.Format("A{0}", t)].Value = pcList.ElementAt(elemPC).Text;
                                Sheet.Cells[string.Format("A{0}", t)].AutoFitColumns();
                                Sheet.Cells[string.Format("B{0}", t)].Value = kpi;
                                Sheet.Cells[string.Format("B{0}", t)].AutoFitColumns();
                                Sheet.Cells[string.Format("C{0}", t)].Value = versionList.ElementAt(elemVersion).Text;
                                Sheet.Cells[string.Format("C{0}", t)].AutoFitColumns();
                                elemVersion++;
                                t++;
                            }

                            elemPC++;
                        }
                    }
                }
                if (kpi == "All")
                {
                    if (version == "All")
                    {
                        int t = 8;
                        int elemPC = 0;
                        foreach (var item in pcList)
                        {
                            int elemKPI = 0;
                            foreach (var itemS in kpiList)
                            {
                                int elemVersion = 0;
                                foreach (var itemV in versionList)
                                {
                                    Sheet.Cells[string.Format("A{0}", t)].Value = pcList.ElementAt(elemPC).Text;
                                    Sheet.Cells[string.Format("A{0}", t)].AutoFitColumns();
                                    Sheet.Cells[string.Format("B{0}", t)].Value = kpiList.ElementAt(elemKPI).Text;
                                    Sheet.Cells[string.Format("B{0}", t)].AutoFitColumns();
                                    Sheet.Cells[string.Format("C{0}", t)].Value = versionList.ElementAt(elemVersion).Text;
                                    Sheet.Cells[string.Format("C{0}", t)].AutoFitColumns();
                                    elemVersion++;
                                    t++;
                                }
                                elemKPI++;
                            }
                            elemPC++;
                        }
                    }
                    if (version != "All") // OK
                    {
                        int t = 8;
                        int elemPC = 0;

                        foreach (var item in pcList)
                        {
                            int elemKPI = 0;
                            foreach (var itemS in kpiList)
                            {
                                Sheet.Cells[string.Format("A{0}", t)].Value = pcList.ElementAt(elemPC).Text;
                                Sheet.Cells[string.Format("A{0}", t)].AutoFitColumns();
                                Sheet.Cells[string.Format("B{0}", t)].Value = kpiList.ElementAt(elemKPI).Text;
                                Sheet.Cells[string.Format("B{0}", t)].AutoFitColumns();
                                Sheet.Cells[string.Format("C{0}", t)].Value = version;
                                Sheet.Cells[string.Format("C{0}", t)].AutoFitColumns();
                                elemKPI++;
                                t++;
                            }
                            elemPC++;
                        }
                    }
                }
            }
            if (pc != "-")
            {
                if (kpi != "All")
                {
                    if (version != "All")
                    {
                        Sheet.Cells["A8"].Value = pc;
                        Sheet.Cells["A8"].AutoFitColumns();
                        Sheet.Cells["B8"].Value = kpi;
                        Sheet.Cells["B8"].AutoFitColumns();
                        Sheet.Cells["C8"].Value = version;
                        Sheet.Cells["C8"].AutoFitColumns();
                    }
                    if (version == "All")
                    {
                        int t = 8;
                        int elemVersion = 0;
                        foreach (var itemV in versionList)
                        {
                            Sheet.Cells[string.Format("A{0}", t)].Value = pc;
                            Sheet.Cells[string.Format("A{0}", t)].AutoFitColumns();
                            Sheet.Cells[string.Format("B{0}", t)].Value = kpi;
                            Sheet.Cells[string.Format("B{0}", t)].AutoFitColumns();
                            Sheet.Cells[string.Format("C{0}", t)].Value = versionList.ElementAt(elemVersion).Text;
                            Sheet.Cells[string.Format("C{0}", t)].AutoFitColumns();
                            elemVersion++;
                            t++;
                        }
                    }
                }
                if (kpi == "All")
                {
                    if (version == "All")
                    {
                        int t = 8;
                        int elemKPI = 0;
                        foreach (var itemS in kpiList)
                        {
                            int elemVersion = 0;
                            foreach (var itemV in versionList)
                            {
                                Sheet.Cells[string.Format("A{0}", t)].Value = pc;
                                Sheet.Cells[string.Format("A{0}", t)].AutoFitColumns();
                                Sheet.Cells[string.Format("B{0}", t)].Value = kpiList.ElementAt(elemKPI).Text;
                                Sheet.Cells[string.Format("B{0}", t)].AutoFitColumns();
                                Sheet.Cells[string.Format("C{0}", t)].Value = versionList.ElementAt(elemVersion).Text;
                                Sheet.Cells[string.Format("C{0}", t)].AutoFitColumns();
                                elemVersion++;
                                t++;
                            }
                            elemKPI++;
                        }
                    }
                    if (version != "All")
                    {
                        int t = 8;
                        int elemKPI = 0;
                        foreach (var itemS in kpiList)
                        {

                            Sheet.Cells[string.Format("A{0}", t)].Value = pc;
                            Sheet.Cells[string.Format("A{0}", t)].AutoFitColumns();
                            Sheet.Cells[string.Format("B{0}", t)].Value = kpiList.ElementAt(elemKPI).Text;
                            Sheet.Cells[string.Format("B{0}", t)].AutoFitColumns();
                            Sheet.Cells[string.Format("C{0}", t)].Value = version;
                            Sheet.Cells[string.Format("C{0}", t)].AutoFitColumns();

                            t++;
                            elemKPI++;
                        }
                    }
                }
            }

            return Ep.GetAsByteArray();
        }

        public static string getMonthByNumber(int i)
        {
            string fullMonthName = new DateTime(2020, i, 1).ToString("MMM", CultureInfo.InvariantCulture);
            return fullMonthName;
        }
        #endregion

        #region ::CRR Claimable ::
        public static byte[] ExportCRRClaimable(string username)
        {
            ExcelPackage Ep = new ExcelPackage();
            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("CRR-Claimable");

            AddHeader(Sheet, "CRR Claimable", username);

            return Ep.GetAsByteArray();
        }
        #endregion
    }
}