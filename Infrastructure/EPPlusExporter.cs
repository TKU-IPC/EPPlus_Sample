using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using EPPlusExam.Infrastructure;

using OfficeOpenXml.Style;


namespace EPPlusExam.Infrastructure
{
    public class EPPlusExporter
    {
        //專門用於欄位表頭合併的情況,此FUNCTION功能應該是目前最好用的，彈性也最大
        //它有二種編製表頭的方式，一是利用XlsColsCfgAttribute在套在資料類別上，二也接受彈性編製資料表頭，可用於欄位合併
        //目前有AC650C使用
        public byte[] GenerateXlsx<S, T>(string[] titles, IEnumerable<S> headers, IEnumerable<T> srcClasses, string footer)
        {
            using (ExcelPackage ep = new ExcelPackage())
            {
                ExcelWorksheet sheet = ep.Workbook.Worksheets.Add("sheet1");
                int rowIdx = 1, colIdx = 1, posIdx = 1;
                XlsColsCfgAttribute colCfgAttr;
                Type classType = typeof(T);
                var colInfos = new List<ColHeaderCfg>();
                int colCount = classType.GetProperties().Length;
                int physicalColCount = colCount;
                for (int i = 0; i < colCount; i++)   //計算真實的欄位個數,順便記錄欄位的資訊
                {
                    colCfgAttr = (XlsColsCfgAttribute)classType.GetProperties()[i].GetCustomAttributes(typeof(XlsColsCfgAttribute), true).SingleOrDefault();
                    if (colCfgAttr != null && colCfgAttr.IsGenerate == false)
                    {
                        physicalColCount--;
                        colInfos.Add(new ColHeaderCfg() { IsGenerate = false });
                    }
                    else
                    {
                        colInfos.Add(new ColHeaderCfg()
                        {
                            IsGenerate = (colCfgAttr != null) ? colCfgAttr.IsGenerate : true,
                            IsShowZero = (colCfgAttr != null) ? colCfgAttr.IsShowZero : true,   //資料為0,預設為顯示          
                            IsPercentSymbol = (colCfgAttr != null) ? colCfgAttr.IsPercentSymbol : true,   //預設為顯示百分比符號
                            ColTypeName = (colCfgAttr != null && !string.IsNullOrEmpty(colCfgAttr.ColTypeName)) ? colCfgAttr.ColTypeName : classType.GetProperties()[i].PropertyType.Name,  //欄位資料型態
                            ColName = (colCfgAttr != null && !string.IsNullOrEmpty(colCfgAttr.ColName)) ? colCfgAttr.ColName : classType.GetProperties()[i].Name,  //欄位中文名稱
                            ColWidth = (colCfgAttr != null) ? colCfgAttr.ColWidth : 20,   //欄寬預設為 20                            
                            PointLenNum = (colCfgAttr != null) ? colCfgAttr.PointLenNum : 2,   //預設取小點以下二位
                            ColAlign = (colCfgAttr != null) ? colCfgAttr.ColAlign : 0,   //資料對齊的位置(預設靠左)
                            ColPosIdx = (colCfgAttr != null) ? colCfgAttr.ColPosIdx : 0,   //資料出現的位置(由左靠右)
                        });
                    }
                }
                if (titles != null)
                {
                    for (int k = 0; k < titles.Length; k++)
                    {
                        ExcelRange rng = sheet.Cells[rowIdx, colIdx, rowIdx++, physicalColCount];
                        rng.Style.Font.Bold = true;
                        rng.Merge = true;
                        rng.Value = titles[k];
                    }
                }

                if (headers == null || headers.Count() == 0)   //當沒有Header的任何設定時，則取用預設表頭
                {
                    foreach (var col in colInfos)
                    {
                        if (!col.IsGenerate) continue;
                        colIdx = col.ColPosIdx == 0 ? colIdx : col.ColPosIdx;
                        using (ExcelRange cell = sheet.Cells[rowIdx, colIdx])
                        {
                            cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            cell.Style.Font.Bold = true;
                            cell.Value = col.ColName;
                            cell.Style.WrapText = true;
                            cell.Worksheet.Column(colIdx).Width = col.ColWidth;
                            colIdx++;
                        }
                    }
                }
                else  //有Header設定，則以此設定為主去建構欄位表頭
                {
                    int headerRowCnt = 0;
                    foreach (var d in headers)
                    {
                        var h = (IXlsHeaderSchema)d;
                        var align = (h.align ?? 0);
                        int width = h.width != null ? (int)h.width : 20;
                        using (ExcelRange cell = sheet.Cells[h.y1 + rowIdx - 1, h.x1, h.y2 + rowIdx - 1, h.x2])
                        {
                            cell.Style.HorizontalAlignment = align == 0 ? ExcelHorizontalAlignment.Left : (align == 1 ? ExcelHorizontalAlignment.Center : ExcelHorizontalAlignment.Right);
                            cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            cell.Style.Font.Bold = true;
                            cell.Value = h.name;
                            cell.Merge = true;
                            cell.Style.WrapText = true;
                            cell.Worksheet.Column(h.x1).Width = width;
                        }
                        headerRowCnt = Math.Max(headerRowCnt, h.y2);
                    }
                    rowIdx += headerRowCnt - 1;
                }

                string value;

                foreach (T s in srcClasses)
                {
                    posIdx = 1;
                    rowIdx++;
                    colIdx = -1;
                    foreach (var col in colInfos)
                    {
                        colIdx++;
                        if (!col.IsGenerate) continue;
                        posIdx = col.ColPosIdx == 0 ? posIdx : col.ColPosIdx;
                        value = (classType.GetProperties()[colIdx].GetValue(s, null) ?? "").ToString();
                        switch (col.ColTypeName.ToLower())
                        {
                            case "decimal":
                                sheet.Cells[rowIdx, posIdx].Style.Numberformat.Format = "#,##0 ;(#,##0);" + (col.IsShowZero ? "0;" : "");
                                //是數值且非最大值才顯示(最大值是為了補救columnsInfo[i, 3] 的不足,因columnsInfo[i, 3] 只能規範該欄的所有設定而無法規範單筆的資料)
                                if (value.IsNumeric() && Convert.ToDecimal(value) != Decimal.MaxValue) sheet.Cells[rowIdx, posIdx].Value = Convert.ToDecimal(value);
                                break;
                            case "double":
                                var zeroStr = (col.PointLenNum > 0 ? ".".PadRight(col.PointLenNum + 1, '0') : "") + (col.IsPercentSymbol ? "%" : "");
                                var formatStr = "#0\\" + zeroStr + ";(#0\\" + zeroStr + ");-";
                                sheet.Cells[rowIdx, posIdx].Style.Numberformat.Format = formatStr;
                                if (value.IsNumeric() && Convert.ToDouble(value) != Double.MaxValue) sheet.Cells[rowIdx, posIdx].Value = Convert.ToDouble(value);
                                break;
                            case "int32":
                                if (value.IsNumeric() && Convert.ToInt32(value) != Int32.MaxValue)
                                    if (Convert.ToInt32(value) != 0 || col.IsShowZero) sheet.Cells[rowIdx, posIdx].Value = Convert.ToInt32(value);
                                break;
                            case "int16":
                                //if (IsNumeric(value) && col.IsShowZero && Convert.ToInt16(value) != Int16.MaxValue) sheet.Cells[rowIdx, posIdx].Value = Convert.ToInt16(value);
                                if (value.IsNumeric() && Convert.ToInt16(value) != Int16.MaxValue)
                                    if (Convert.ToInt16(value) != 0 || col.IsShowZero) sheet.Cells[rowIdx, posIdx].Value = Convert.ToInt16(value);
                                break;
                            case "nullable`1":
                                if (value.IsNumeric())
                                {
                                    sheet.Cells[rowIdx, posIdx].Style.Numberformat.Format = "#,##0 ;(#,##0);" + (col.IsShowZero ? "0;" : "");
                                    sheet.Cells[rowIdx, posIdx].Value = Convert.ToDouble(value);
                                }
                                else
                                    sheet.Cells[rowIdx, posIdx].Value = Convert.ToString(classType.GetProperties()[colIdx].GetValue(s, null));
                                break;
                            default:
                                sheet.Cells[rowIdx, posIdx].Value = Convert.ToString(classType.GetProperties()[colIdx].GetValue(s, null));
                                break;
                        }
                        posIdx++;
                    }
                }
                if (footer != null)
                {
                    rowIdx += 2;
                    ExcelRange footerMemo = sheet.Cells[rowIdx, 1, rowIdx, physicalColCount];
                    footerMemo.Merge = true;
                    footerMemo.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    footerMemo.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    footerMemo.Style.WrapText = true;
                    footerMemo.Value = footer;
                    sheet.Row(rowIdx).Height = 60;
                }
                return ep.GetAsByteArray();
            }
        }
    }
}