using CT;
using CT.Logging;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace speard_report
{
    class WriteExcel
    {
        private Ini iniReader = Ini.ProgramIniFile;
        private Dictionary<string, double?[]> m_ListBroker;
        private Dictionary<string, double?[]> m_ListOtherBroker;
        private string[] m_OtherBrokerName;
        private string[] m_BrokerName;
        public WriteExcel(
            ExcelWorksheet workSheet,
            string [] BrokerName, 
            string[] otherBrokerName, 
            Dictionary<string, double?[]> listBroker,
            Dictionary<string, double?[]> listOtherBroker)
        {
            m_ListBroker = listBroker;
            m_ListOtherBroker = listOtherBroker;
            m_OtherBrokerName = otherBrokerName;
            m_BrokerName = BrokerName;
            bool exportNonGBE = iniReader.ReadSetting("Settings", "ExportNonGBE") == "true";
            CreateHeader(workSheet, BrokerName.Length + otherBrokerName.Length);
            // config row excel
 
            if (exportNonGBE)
            {
                // true la only GBE
              
                // flase  la in ra tat  ca
                CreateTable(5, 1, workSheet, true, "XCore vs Xcore");
                CreateTable(5, 1, workSheet, false, "Yesterday's average spreads");
            }
            // false: load 1 table
            else
            {
                // in ra 1 table , set cot bat dau lai
                CreateTable(5, 1 - (otherBrokerName.Length + 2), workSheet, true, "Yesterday's average spreads");
            }
        }
        private void CreateHeader(ExcelWorksheet worksheet, int sizeRow)
        {
            var header = worksheet.Cells[1, 1, 2, sizeRow + 3];
            header.Merge = true;
            header.Value = string.Format("Average Spread {0} {1} ",
            DateTime.UtcNow.Subtract(TimeSpan.FromDays(1)).ToString("dd.MM.yyyy"), worksheet);
            header.Style.Font.Bold = true;
            header.Style.Font.UnderLine = true;
            header.Style.Font.Size = 14;
            header.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            header.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        }
        private void CreateTable(  int rowStart, int columnStart, ExcelWorksheet workSheet, bool isXcore,  string contentHeader)
        {
            Dictionary<string, double?[]> listAvg =  m_ListOtherBroker;
            string[] groupName = m_OtherBrokerName;
            if (isXcore) 
            {
                columnStart = columnStart + m_OtherBrokerName.Length + 2;
                listAvg = m_ListBroker;
                groupName = m_BrokerName;
                HeaderTable(rowStart, columnStart, workSheet, listAvg.Values.First().Length, contentHeader);
            }

            HeaderTable(rowStart, columnStart, workSheet, listAvg.Values.First().Length, contentHeader);

            workSheet.DefaultColWidth = 13;
            // vi tri bat dau hang 2 cot 1  create header Symbol
            workSheet.Cells[rowStart, columnStart].Value = "Symbol";
            workSheet.Cells[rowStart, columnStart].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Cells[rowStart, columnStart].Style.Border.Bottom.Style = workSheet.Cells[rowStart, columnStart].Style.Border.Top.Style = workSheet.Cells[rowStart, columnStart].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            workSheet.Cells[rowStart, columnStart].Style.Font.Bold = true;

            //load header vi tri hang 2 cot 2 Load group broker at the header start to the right the symbol
            for (int i = 0; i < groupName.Length; i++)
            {
                workSheet.Cells[rowStart, columnStart + i + 1].Value = groupName[i];
                workSheet.Cells[rowStart, columnStart + i + 1].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                workSheet.Cells[rowStart, columnStart + i + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[rowStart, columnStart + i + 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells[rowStart, columnStart + i + 1].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                workSheet.Cells[rowStart, columnStart + i + 1].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                workSheet.Cells[rowStart, columnStart + i + 1].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                if (i == groupName.Length - 1)// border right when last item
                {
                    workSheet.Cells[rowStart, columnStart + i + 1].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                }
                if (i == 0) // border left when left item
                {
                    workSheet.Cells[rowStart, columnStart + i + 1].Style.Border.Left.Style = ExcelBorderStyle.Medium;
                }
                workSheet.Cells[rowStart, columnStart + i + 1].Style.Font.Bold = true; 
            }
            rowStart += 1; // row two 
            //loop 1 lap cac key trong rowStarts
            foreach (var item in listAvg)
            {
                //load cac symbol o vi tri hang 3 cot 1
                workSheet.Cells[rowStart, columnStart].Value = item.Key;
                workSheet.Cells[rowStart, columnStart].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Cells[rowStart, columnStart].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells[rowStart, columnStart].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                workSheet.Cells[rowStart, columnStart].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                workSheet.Cells[rowStart, columnStart].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                workSheet.Cells[rowStart, columnStart].Style.Font.Bold = true;

                double? minValue = item.Value.Min();// find min value
                int lengthItem = item.Value.Length;// length value a item
                bool isMin = false; // set min value only once
                for (int i = 0; i < lengthItem; i++)
                {
                    // border right when column end
                    if (i == lengthItem - 1)
                        workSheet.Cells[rowStart, columnStart + i + 1].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    //border to the left when index's current != index end, start
                    if (i > 0 && i < lengthItem)
                        workSheet.Cells[rowStart, columnStart + i + 1].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //border bottom when rowStart end
                    if (item.Key == listAvg.Last().Key)
                        workSheet.Cells[rowStart, columnStart + i + 1].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;

                    if (item.Value[i] == minValue && !isMin)
                    {
                        isMin = true;
                        workSheet.Cells[rowStart, columnStart + i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        workSheet.Cells[rowStart, columnStart + i + 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(146, 208, 80));
                    }

                    workSheet.Cells[rowStart, columnStart + i + 1].Value =
                       item.Value[i].HasValue
                        ? Math.Round(item.Value[i].Value, 5, MidpointRounding.AwayFromZero)
                        : (double?)null;
                    workSheet.Cells[rowStart, columnStart + i + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    workSheet.Cells[rowStart, columnStart + i + 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells[rowStart, columnStart + i + 1].Style.Numberformat.Format = "0.00";
                    if (item.Value[i] == null)
                    {
                        this.LogInfo($"empty AverageSpread to write excel");
                    }
                    else
                    {
                        this.LogInfo($"write AverageSpread {item.Value[i]} in excel");
                    }
                }
                rowStart++;
            }

        }
        private void HeaderTable(int rowStart, int columnStart, ExcelWorksheet workSheet, int length, string strHeader)
        {
            var header = workSheet.Cells[rowStart - 1, columnStart + 1, rowStart - 1, columnStart + length];
            header.Value = strHeader;
            header.Merge = true;
            header.Style.Border.Top.Style = header.Style.Border.Right.Style = header.Style.Border.Left.Style = ExcelBorderStyle.Medium;
            header.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            header.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        }
    }
}
