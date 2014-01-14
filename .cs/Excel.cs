using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;
using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace aps_excel_cs
{

    public class functionProperties
    {
        public string fn { get; set; }
    }

    public class Excel
    {
        private int uniqueKey = 0;
        private int GetKey() {
            return uniqueKey++;
        }
        private readonly Dictionary<int, object> KeyedCache = new Dictionary<int, object>();
        private readonly Dictionary<object, int> CacheKeys = new Dictionary<object, int>();
        private int AddToCache(object itemToCache) {
            int key;

            if (CacheKeys.TryGetValue(itemToCache, out key)) {
                return key;
            } else {
                key = GetKey();
                KeyedCache.Add(key, itemToCache);
                CacheKeys.Add(itemToCache, key);
                return key;
            }
        }


        public async Task<object> Invoke(object input) {
            var parameters = (IDictionary<string, object>)input;
            var functionName = (string)parameters["fn"];

            functionName = functionName.Replace('-', '_');

            return this.GetType().InvokeMember(functionName, BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase | BindingFlags.InvokeMethod, null, this, new object[] {parameters});
        }


        public object Excel_Load(IDictionary<string, object> parameters) {
            var path = (string)parameters["path"];
                
            using (var file = File.Open(path, FileMode.Open, FileAccess.ReadWrite)) {
                var workbook = new HSSFWorkbook(file);

                return AddToCache(workbook);
            }
        }

        public object Workbook_Save(IDictionary<string, object> parameters) {
            var workbook = (HSSFWorkbook)KeyedCache[(int)parameters["workbook"]];
            var path = (string)parameters["path"];

            using (var file = File.Open(path, FileMode.Create, FileAccess.ReadWrite)) {
                workbook.Write(file);
            }

            return null;
        }
        public object Workbook_GetSheetAt(IDictionary<string, object> parameters) {
            var workbook = (HSSFWorkbook)KeyedCache[(int)parameters["workbook"]];
            var index = (int)parameters["index"];

            return AddToCache(workbook.GetSheetAt(index));
        }
        public object Workbook_GetSheet(IDictionary<string, object> parameters) {
            var workbook = (HSSFWorkbook)KeyedCache[(int)parameters["workbook"]];
            var name = (string)parameters["name"];

            return AddToCache(workbook.GetSheet(name));
        }

        public object Sheet_GetRow(IDictionary<string, object> parameters) {
            var sheet = (HSSFSheet)KeyedCache[(int)parameters["sheet"]];
            var row = (int)parameters["row"];

            var rowData = sheet.GetRow(row);
            if (rowData != null) return AddToCache(rowData);
            return null;
        }
        public object Sheet_GetRowExists(IDictionary<string, object> parameters) {
            var sheet = (HSSFSheet)KeyedCache[(int)parameters["sheet"]];
            var row = (int)parameters["row"];

            var rowData = sheet.GetRow(row);
            return rowData != null;
        }
        public object Sheet_GetRowCount(IDictionary<string, object> parameters) {
            var sheet = (HSSFSheet)KeyedCache[(int)parameters["sheet"]];

            return sheet.LastRowNum;
        }
        public object Sheet_CloneRow(IDictionary<string, object> parameters) {
            var sheet = (HSSFSheet)KeyedCache[(int)parameters["sheet"]];
            var sourceRow = (int)parameters["sourceRow"];
            var destRow = (int)parameters["destRow"];

            // Make the destination exist
            if (sheet.GetRow(destRow) == null) sheet.CreateRow(destRow);

            // Clone the source styling onto the destination
            sheet.GetRow(sourceRow).CopyRowTo(destRow);

            return null;
        }
        public object Sheet_Protect(IDictionary<string, object> parameters) {
            var sheet = (HSSFSheet)KeyedCache[(int)parameters["sheet"]];
            var password = (string)parameters["password"];

            sheet.ProtectSheet(password);
            return null;
        }
        public object Sheet_Unprotect(IDictionary<string, object> parameters) {
            var sheet = (HSSFSheet)KeyedCache[(int)parameters["sheet"]];

            sheet.ProtectSheet(null);
            return null;
        }
        public object Sheet_CreateRow(IDictionary<string, object> parameters) {
            var sheet = (ISheet)KeyedCache[(int)parameters["sheet"]];
            var rowIndex = (int)parameters["row"];

            var row = sheet.CreateRow(rowIndex);
            return AddToCache(row);
        }

        //public object Sheet_GetCellValue(IDictionary<string, object> parameters) {
        //    var sheet = (HSSFSheet)KeyedCache[(int)parameters["sheet"]];
        //    var row = (int)parameters["row"];
        //    var column = (int)parameters["column"];

        //    var cell = getCell(sheet, row, column, false);
        //    if (cell == null) return null;

        //    switch (cell.CellType) {
        //        case CellType.Unknown:
        //            return null;
        //        case CellType.Numeric:
        //            if (HSSFDateUtil.IsCellDateFormatted(cell)) {
        //                return cell.DateCellValue;
        //            } else {
        //                return cell.NumericCellValue;
        //            }
        //        case CellType.String:
        //            return cell.StringCellValue;
        //        case CellType.Formula:
        //            switch (cell.CachedFormulaResultType) {
        //                case CellType.Unknown:
        //                    return null;
        //                case CellType.Numeric:
        //                    if (HSSFDateUtil.IsCellDateFormatted(cell)) {
        //                        return cell.DateCellValue;
        //                    } else {
        //                        return cell.NumericCellValue;
        //                    }
        //                case CellType.String:
        //                    return cell.StringCellValue;
        //                case CellType.Blank:
        //                    return null;
        //                case CellType.Boolean:
        //                    return cell.BooleanCellValue;
        //                case CellType.Error:
        //                    return cell.ErrorCellValue;
        //                default:
        //                    return null;
        //            }
        //        case CellType.Blank:
        //            return null;
        //        case CellType.Boolean:
        //            return cell.BooleanCellValue;
        //        case CellType.Error:
        //            return cell.ErrorCellValue;
        //        default:
        //            return null;
        //    }
        //}
        //public object Sheet_SetCellValue(IDictionary<string, object> parameters) {
        //    var sheet = (HSSFSheet)KeyedCache[(int)parameters["sheet"]];
        //    var row = (int)parameters["row"];
        //    var column = (int)parameters["column"];
        //    var value = parameters["value"];

        //    var cell = getCell(sheet, row, column);
        //    var columnType = value.GetType();
        //    //return cell.DateCellValue;

        //    if (columnType == typeof (DateTime)) {
        //        cell.SetCellValue((DateTime)value);

        //    } else if (columnType == typeof (int) || columnType == typeof (double)) {
        //        cell.SetCellValue(Convert.ToDouble(value));

        //    } else if (columnType == typeof (string)) {
        //        cell.SetCellValue((string)value);
                
        //    } else if (columnType == typeof (bool)) {
        //        cell.SetCellValue((bool)value);
        //    }

        //    return null;
        //}
        
        public object Row_GetCell(IDictionary<string, object> parameters) {
            var row = (IRow)KeyedCache[(int)parameters["row"]];
            var cell = (int)parameters["cell"];

            var cellData = row.GetCell(cell);
            if (cellData != null) return AddToCache(cellData);
            return null;
        }
        public object Row_GetCellExists(IDictionary<string, object> parameters) {
            var row = (IRow)KeyedCache[(int)parameters["row"]];
            var cell = (int)parameters["cell"];

            var cellData = row.GetCell(cell);
            return cellData != null;
        }
        public object Row_CreateCell(IDictionary<string, object> parameters) {
            var row = (IRow)KeyedCache[(int)parameters["row"]];
            var cellIndex = (int)parameters["cell"];

            var cell = row.CreateCell(cellIndex);
            return AddToCache(cell);
        }


        public object Cell_GetValue(IDictionary<string, object> parameters) {
            var cell = (ICell)KeyedCache[(int)parameters["cell"]];
            if (cell == null) return null;

            switch (cell.CellType) {
                case CellType.Unknown:
                    return null;
                case CellType.Numeric:
                    if (HSSFDateUtil.IsCellDateFormatted(cell)) {
                        return cell.DateCellValue;
                    } else {
                        return cell.NumericCellValue;
                    }
                case CellType.String:
                    return cell.StringCellValue;
                case CellType.Formula:
                    switch (cell.CachedFormulaResultType) {
                        case CellType.Unknown:
                            return null;
                        case CellType.Numeric:
                            if (HSSFDateUtil.IsCellDateFormatted(cell)) {
                                return cell.DateCellValue;
                            } else {
                                return cell.NumericCellValue;
                            }
                        case CellType.String:
                            return cell.StringCellValue;
                        case CellType.Blank:
                            return null;
                        case CellType.Boolean:
                            return cell.BooleanCellValue;
                        case CellType.Error:
                            return cell.ErrorCellValue;
                        default:
                            return null;
                    }
                case CellType.Blank:
                    return null;
                case CellType.Boolean:
                    return cell.BooleanCellValue;
                case CellType.Error:
                    return cell.ErrorCellValue;
                default:
                    return null;
            }
        }
        public object Cell_SetValue(IDictionary<string, object> parameters) {
            var cell = (ICell)KeyedCache[(int)parameters["cell"]];
            var value = parameters["value"];

            var columnType = value.GetType();
            //return cell.DateCellValue;

            if (columnType == typeof(DateTime)) {
                cell.SetCellValue((DateTime)value);

            } else if (columnType == typeof(int) || columnType == typeof(double)) {
                cell.SetCellValue(Convert.ToDouble(value));

            } else if (columnType == typeof(string)) {
                cell.SetCellValue((string)value);

            } else if (columnType == typeof(bool)) {
                cell.SetCellValue((bool)value);
            }

            return null;
        }
        public object Cell_SetLock(IDictionary<string, object> parameters) {
            var cell = (ICell)KeyedCache[(int)parameters["cell"]];
            var locked = (bool)parameters["lock"];

            var cellStyle = cell.Row.Sheet.Workbook.CreateCellStyle();
            cellStyle.CloneStyleFrom(cell.CellStyle);
            cellStyle.IsLocked = locked;

            cell.CellStyle = cellStyle;
            return null;
        }
        public object Cell_GetLock(IDictionary<string, object> parameters) {
            var cell = (ICell)KeyedCache[(int)parameters["cell"]];
            return cell.CellStyle.IsLocked;
        }
    }

}
