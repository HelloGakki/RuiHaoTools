using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop;
using Application = Microsoft.Office.Interop.Excel.Application;
using System.Runtime.InteropServices;

namespace RuiHaoConvertor
{
    public class ExcelHelper
    {
        #region "Private Variable Definition"

        private Application _excelApp = null;
        private Workbook _workBook = null;
        private Worksheet _workSheet = null;
        private object _miss = System.Reflection.Missing.Value;

        #endregion

        #region "Public Property and Constant Definition"

        /// <summary>
        /// Excel单元格边框的线条的粗细枚举
        /// </summary>
        public enum ExcelBorderWeight
        {
            /// <summary>
            /// 极细的线条
            /// </summary>
            Hairline = XlBorderWeight.xlHairline,
            /// <summary>
            /// 中等的线条
            /// </summary>
            Medium = XlBorderWeight.xlMedium,
            /// <summary>
            /// 粗线条
            /// </summary>
            Thick = XlBorderWeight.xlThick,
            /// <summary>
            /// 细线条
            /// </summary>
            Thin = XlBorderWeight.xlThin
        }

        /// <summary>
        /// Excel单元格边框枚举
        /// </summary>
        public enum ExcelBordersIndex
        {
            /// <summary>
            /// 主对角线从
            /// </summary>
            DiagonalDown = XlBordersIndex.xlDiagonalDown,
            /// <summary>
            /// 辅对角线
            /// </summary>
            DiagonUp = XlBordersIndex.xlDiagonalUp,
            /// <summary>
            ///底边框
            /// </summary>
            EdgeBottom = XlBordersIndex.xlEdgeBottom,
            /// <summary>
            /// 左边框
            /// </summary>
            EdgeLeft = XlBordersIndex.xlEdgeLeft,
            /// <summary>
            /// 右边框
            /// </summary>
            EdgeRight = XlBordersIndex.xlEdgeRight,
            /// <summary>
            /// 顶边框
            /// </summary>
            EdgeTop = XlBordersIndex.xlEdgeTop,
            /// <summary>
            /// 边框内水平横线
            /// </summary>
            InsideHorizontal = XlBordersIndex.xlInsideHorizontal,
            /// <summary>
            /// 边框内垂直竖线
            /// </summary>
            InsideVertical = XlBordersIndex.xlInsideVertical
        }

        /// <summary>
        /// Excel单元格的竖直方法对齐枚举
        /// </summary>
        public enum ExcelVerticalAlignment
        {
            /// <summary>
            /// 居中
            /// </summary>
            Center = Constants.xlCenter,
            /// <summary>
            /// 靠上
            /// </summary>
            Top = Constants.xlTop,
            /// <summary>
            /// 靠下
            /// </summary>
            Bottom = Constants.xlBottom,
            /// <summary>
            /// 两端对齐
            /// </summary>
            Justify = Constants.xlJustify,
            /// <summary>
            /// 分散对齐
            /// </summary>
            Distributed = Constants.xlDistributed

        };

        /// <summary>
        /// Excel 水平方向对齐枚举
        /// </summary>
        public enum ExcelHorizontalAlignment
        {
            /// <summary>
            ///常规
            /// </summary>
            General = Constants.xlGeneral,
            /// <summary>
            /// 靠左
            /// </summary>
            Left = Constants.xlLeft,
            /// <summary>
            /// 居中
            /// </summary>
            Center = Constants.xlCenter,
            /// <summary>
            /// 靠右
            /// </summary>
            Right = Constants.xlRight,
            /// <summary>
            /// 填充
            /// </summary>
            Fill = Constants.xlFill,
            /// <summary>
            /// 两端对齐
            /// </summary>
            Justify = Constants.xlJustify,
            /// <summary>
            /// 跨列居中
            /// </summary>
            CenterAcrossSelection = Constants.xlCenterAcrossSelection,
            /// <summary>
            /// 分散对齐
            /// </summary>
            Distributed = Constants.xlDistributed

        }


        /// <summary>
        /// Excel边框线条的枚举
        /// </summary>
        public enum ExcelStyleLine
        {
            /// <summary>
            /// 没有线条
            /// </summary>
            StyleNone = XlLineStyle.xlLineStyleNone,
            /// <summary>
            /// 连续的细线
            /// </summary>
            Continious = XlLineStyle.xlContinuous,
            /// <summary>
            /// 点状线
            /// </summary>
            Dot = XlLineStyle.xlDot,
            /// <summary>
            /// 双条线
            /// </summary>
            Double = XlLineStyle.xlDouble,
        }

        /// <summary>
        /// 排序的玫举
        /// </summary>
        public enum ExcelSortOrder
        {
            /// <summary>
            /// 升序
            /// </summary>
            Ascending = XlSortOrder.xlAscending,
            /// <summary>
            /// 降序
            /// </summary>
            Descending = XlSortOrder.xlDescending,
        }

        #endregion

        #region "Construction Method"

        /// <summary>
        /// 构造函数
        /// </summary>
        public ExcelHelper(bool excelVisible = false)
        {
            _excelApp = new Application();
            _excelApp.Visible = excelVisible;
        }

        #endregion

        #region "Open and dispose method definition"

        /// <summary>
        /// 打开空白工作文档
        /// </summary>
        public void Open()
        {
            _workBook = _excelApp.Workbooks.Add(true);
        }
        /// <summary>
        /// 打开指定路径文件
        /// </summary>
        /// <param name="xltPath"></param>
        public void Open(string xltPath)
        {
            if (System.IO.File.Exists(xltPath))
            {
                _workBook = _excelApp.Workbooks.Add(xltPath);
                _workSheet = (Worksheet)_workBook.ActiveSheet;
            }
            else
                throw new System.IO.FileNotFoundException(string.Format("{0}不存在，请重新确定文件名", xltPath));

        }
        /// <summary>
        /// 打开指定路径工作文档
        /// </summary>
        /// <param name="xltPath"></param>
        /// <param name="sheetName"></param>
        public void Open(string xltPath, string sheetName)
        {
            if (System.IO.File.Exists(xltPath))
            {
                _workBook = _excelApp.Workbooks.Add(xltPath);
                _workSheet = (Worksheet)_workBook.Worksheets[sheetName];
            }
            else
                throw new System.IO.FileNotFoundException(string.Format("{0}不存在，请重新确定文件名", xltPath));
        }
        /// <summary>
        /// 以指定文件名保存到桌面
        /// </summary>
        /// <param name="newName"></param>
        public void Save(string newName)
        {
            if (_workBook != null)
            {
                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                _workBook.SaveCopyAs(desktopPath + @"\" + newName + @".xlsx");
            }
        }

        /// <summary>
        /// 表示已经保存
        /// </summary>
        public void Saved()
        {
            if (_workBook != null)
                _workBook.Saved = true;
        }
        /// <summary>
        /// 释放资源
        /// </summary>
        /// <param name="hwnd"></param>
        /// <param name="ID"></param>
        /// <returns></returns>
        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        private static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);
        public void Dispose()
        {
            if (_workBook != null)
            {
                _workBook.Close();
                _workBook = null;
            }
            if (_excelApp != null)
            {
                _excelApp.Quit();
                //_excelApp = null;
            }

            // 杀进程
            int id = 0;
            GetWindowThreadProcessId(new IntPtr(_excelApp.Hwnd), out id);
            System.Diagnostics.Process progress = System.Diagnostics.Process.GetProcessById(id);
            progress.Kill();
        }
        #endregion

        #region "Get excel range method definition"

        /// <summary>
        /// 返回指定名称的单元格
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        public Range GetRange(string cell)
        {
            return _workSheet.Range[cell];
        }
        /// <summary>
        /// 返回指定坐标的单元格
        /// </summary>
        /// <param name="row">定位的行</param>
        /// <param name="column">定位的列</param>
        /// <returns></returns>
        public Range GetRange(int row, int column)
        {
            return (Range)_workSheet.Cells[row, column];
        }
        /// <summary>
        /// 返回指定名称到指定名称的范围单元格块
        /// </summary>
        /// <param name="startCell">定位开始的单元格名称</param>
        /// <param name="endCell">定位结束的单元格名称</param>
        /// <returns></returns>
        public Range GetRange(string startCell, string endCell)
        {
            return _workSheet.Range[startCell, endCell];
        }
        /// <summary>
        /// 返回坐标点到坐标点的范围单元格区域
        /// </summary>
        /// <param name="startRow">定位开始的cell行</param>
        /// <param name="startColumn">定位开始的cell列</param>
        /// <param name="endRow">定位结束的cell行</param>
        /// <param name="endColumn">定位结束的cell列</param>
        /// <returns></returns>
        public Range GetRange(int startRow, int startColumn, int endRow, int endColumn)
        {
            return _workSheet.Range[_excelApp.Cells[startRow, startColumn], _excelApp.Cells[endRow, endColumn]];
        }

        #endregion

        #region "Detail control excel method"

        /// <summary>
        /// 隐藏工作簿
        /// </summary>
        public void Hide()
        {
            _excelApp.Visible = false;
        }
        /// <summary>
        /// 显示工作簿
        /// </summary>
        public void Show()
        {
            _excelApp.Visible = true;
        }
        /// <summary>
        /// 设定要操作的工作簿
        /// </summary>
        /// <param name="sheetName">工作簿名称</param>
        public void SetActivitySheet(string sheetName)
        {
            _workSheet = (Worksheet)_workBook.Worksheets[sheetName];
        }
        /// <summary>
        /// 设定工作簿名称
        /// </summary>
        /// <param name="newName">工作簿名称</param>
        public void SetSheetName(string newName)
        {
            _workSheet.Name = newName;
        }
        /// <summary>
        /// 返回指定名称单元格的内容
        /// </summary>
        /// <param name="cell">单元格名称</param>
        /// <returns></returns>
        public object GetCellValue(string cell)
        {
            return GetRange(cell).Value;
        }
        /// <summary>
        /// 返回指定单元格的内容
        /// </summary>
        /// <param name="row">指定单元格的行</param>
        /// <param name="column">指定单元格的列</param>
        /// <returns></returns>
        public object GetCellValue(int row, int column)
        {
            return GetRange(row, column).Value;
        }
        /// <summary>
        /// 返回指定名称到指定名称单元格块的内容
        /// </summary>
        /// <param name="startCell">开始的指定名称单元格</param>
        /// <param name="endCell">结束的指定名称单元格</param>
        /// <returns></returns>
        public object GetCellValue(string startCell, string endCell)
        {
            return GetRange(startCell, endCell).Value;
        }
        /// <summary>
        /// 返回指定单元格范围的内容
        /// </summary>
        /// <param name="startRow">定位开始的行</param>
        /// <param name="startColumn">定位开始的列</param>
        /// <param name="endRow">定位结束的行</param>
        /// <param name="endColumn">定位结束的列</param>
        /// <returns></returns>
        public object GetCellValue(int startRow, int startColumn, int endRow, int endColumn)
        {
            return GetRange(startRow, startColumn, endRow, endColumn).Value;
        }
        /// <summary>
        /// 设置指定名称单元格的内容
        /// </summary>
        /// <param name="cell">指定名称的单元格</param>
        /// <param name="text">填入的内容</param>
        /// 考虑可以公式扩展，使用formula属性
        public void SetCellValue(string cell, string text)
        {
            GetRange(cell).Formula = text;
        }
        /// <summary>
        /// 设置指定名称单元格的内容
        /// </summary>
        /// <param name="cell">指定名称的单元格</param>
        /// <param name="text">填入的内容</param>
        /// 考虑可以公式扩展，使用formula属性
        public void SetCellValue(string cell, int text)
        {
            GetRange(cell).Formula = text;
        }
        /// <summary>
        /// 设置指定名称单元格的内容
        /// </summary>
        /// <param name="cell">指定名称的单元格</param>
        /// <param name="text">填入的内容</param>
        /// 考虑可以公式扩展，使用formula属性
        public void SetCellValue(string cell, double text)
        {
            GetRange(cell).Formula = text;
        }
        /// <summary>
        /// 设置指定坐标单元格的内容
        /// </summary>
        /// <param name="row">指定的单元格行</param>
        /// <param name="column">指定的单元格列</param>
        /// <param name="text">填入的内容</param>
        public void SetCellValue(int row, int column, string text)
        {
            GetRange(row, column).Formula = text;
        }
        /// <summary>
        /// 设置指定坐标单元格的内容
        /// </summary>
        /// <param name="row">指定的单元格行</param>
        /// <param name="column">指定的单元格列</param>
        /// <param name="text">填入的内容</param>
        public void SetCellValue(int row, int column, int text)
        {
            GetRange(row, column).Formula = text;
        }
        /// <summary>
        /// 设置指定坐标单元格的内容
        /// </summary>
        /// <param name="row">指定的单元格行</param>
        /// <param name="column">指定的单元格列</param>
        /// <param name="text">填入的内容</param>
        public void SetCellValue(int row, int column, double text)
        {
            GetRange(row, column).Formula = text;
        }
        /// <summary>
        /// 设置指定范围单元格的内容
        /// </summary>
        /// <param name="startCell">开始的指定名称单元格</param>
        /// <param name="endCell">结束的指定名称单元格</param>
        /// <param name="text">填入的内容</param>
        public void SetCellValue(string startCell, string endCell, string text)
        {
            GetRange(startCell, endCell).Formula = text;
        }
        /// <summary>
        /// 设置指定范围单元格的内容
        /// </summary>
        /// <param name="startCell">开始的指定名称单元格</param>
        /// <param name="endCell">结束的指定名称单元格</param>
        /// <param name="text">填入的内容</param>
        public void SetCellValue(string startCell, string endCell, int text)
        {
            GetRange(startCell, endCell).Formula = text;
        }
        /// <summary>
        /// 设置指定范围单元格的内容
        /// </summary>
        /// <param name="startCell">开始的指定名称单元格</param>
        /// <param name="endCell">结束的指定名称单元格</param>
        /// <param name="text">填入的内容</param>
        public void SetCellValue(string startCell, string endCell, double text)
        {
            GetRange(startCell, endCell).Formula = text;
        }
        /// <summary>
        /// 设置指定范围单元格的内容
        /// </summary>
        /// <param name="startRow">开始的行坐标</param>
        /// <param name="startColumn">开始的列坐标</param>
        /// <param name="endRow">结束的行坐标</param>
        /// <param name="endColumn">结束的列坐标</param>
        /// <param name="text">填入的内容</param>
        public void SetCellValue(int startRow, int startColumn, int endRow, int endColumn, string text)
        {
            GetRange(startRow, startColumn, endRow, endColumn).Formula = text;
        }
        /// <summary>
        /// 设置指定范围单元格的内容
        /// </summary>
        /// <param name="startRow">开始的行坐标</param>
        /// <param name="startColumn">开始的列坐标</param>
        /// <param name="endRow">结束的行坐标</param>
        /// <param name="endColumn">结束的列坐标</param>
        /// <param name="text">填入的内容</param>
        public void SetCellValue(int startRow, int startColumn, int endRow, int endColumn, int text)
        {
            GetRange(startRow, startColumn, endRow, endColumn).Formula = text;
        }
        /// <summary>
        /// 设置指定范围单元格的内容
        /// </summary>
        /// <param name="startRow">开始的行坐标</param>
        /// <param name="startColumn">开始的列坐标</param>
        /// <param name="endRow">结束的行坐标</param>
        /// <param name="endColumn">结束的列坐标</param>
        /// <param name="text">填入的内容</param>
        public void SetCellValue(int startRow, int startColumn, int endRow, int endColumn, double text)
        {
            GetRange(startRow, startColumn, endRow, endColumn).Formula = text;
        }

        #endregion

        #region "Excel range style method definition"

        /// <summary>
        /// 设定单元格的垂直对齐方式
        /// </summary>
        /// <param name="cell">指定单元格名称</param>
        /// <param name="cellAlignment">对齐方式</param>
        public void SetCellVerticalAlignment(string cell, ExcelVerticalAlignment cellAlignment)
        {
            Range range = GetRange(cell);
            range.Select();
            range.VerticalAlignment = cellAlignment;
        }
        /// <summary>
        /// 设置单元格的垂直对齐方式
        /// </summary>
        /// <param name="row">单元格的行坐标</param>
        /// <param name="column">单元格的列坐标</param>
        /// <param name="cellAlignment">对齐方式</param>
        public void SetCellVerticalAlignment(int row, int column, ExcelVerticalAlignment cellAlignment)
        {
            Range range = GetRange(row, column);
            range.Select();
            range.VerticalAlignment = cellAlignment;
        }
        /// <summary>
        /// 设置范围单元格的垂直对齐方式
        /// </summary>
        /// <param name="startCell">开始的单元格名称</param>
        /// <param name="endCell">结束的单元格名称</param>
        /// <param name="cellAlignment">对齐方式</param>
        public void SetCellVerticalAlignment(string startCell, string endCell, ExcelVerticalAlignment cellAlignment)
        {
            Range range = GetRange(startCell, endCell);
            range.Select();
            range.VerticalAlignment = cellAlignment;
        }
        /// <summary>
        /// 设置范围单元格的垂直对齐方式
        /// </summary>
        /// <param name="startRow">开始的单元格行坐标</param>
        /// <param name="startColumn">开始的单元格列坐标</param>
        /// <param name="endRow">结束的单元格行坐标</param>
        /// <param name="endColumn">结束的单元格列坐标</param>
        /// <param name="cellAlignment">对齐方式</param>
        public void SetCellVerticalAlignment(int startRow, int startColumn, int endRow, int endColumn, ExcelVerticalAlignment cellAlignment)
        {
            Range range = GetRange(startRow, startColumn, endRow, endColumn);
            range.Select();
            range.VerticalAlignment = cellAlignment;
        }
        /// <summary>
        /// 设置单元格的水平对齐方式
        /// </summary>
        /// <param name="cell">指定的单元格</param>
        /// <param name="cellAlignment">对齐方式</param>
        public void SetCellHorizontalAlignment(string cell, ExcelHorizontalAlignment cellAlignment)
        {
            Range range = GetRange(cell);
            range.Select();
            range.HorizontalAlignment = cellAlignment;
        }
        /// <summary>
        /// 设置的单元格水平对齐方式
        /// </summary>
        /// <param name="row">指定单元格的行坐标</param>
        /// <param name="column">制定单元格的列坐标</param>
        /// <param name="cellAlignment">对齐方式</param>
        public void SetCellHorizontalAlignment(int row, int column, ExcelHorizontalAlignment cellAlignment)
        {
            Range range = GetRange(row, column);
            range.Select();
            range.HorizontalAlignment = cellAlignment;
        }
        /// <summary>
        /// 设置的单元格水平对齐方式
        /// </summary>
        /// <param name="startCell">开始的单元格名称</param>
        /// <param name="endCell">结束的单元格名称</param>
        /// <param name="cellAlignment">对齐方式</param>
        public void SetCellHorizontalAlignment(string startCell, string endCell, ExcelHorizontalAlignment cellAlignment)
        {
            Range range = GetRange(startCell, endCell);
            range.Select();
            range.HorizontalAlignment = cellAlignment;
        }
        /// <summary>
        /// 设置的单元格水平对齐方式
        /// </summary>
        /// <param name="startRow">开始的单元格行坐标</param>
        /// <param name="startColumn">开始的单元格列坐标</param>
        /// <param name="endRow">结束的单元格行坐标</param>
        /// <param name="endColumn">结束的单元格列坐标</param>
        /// <param name="cellAlignment">对齐方式</param>
        public void SetCellHorizontalAlignment(int startRow, int startColumn, int endRow, int endColumn, ExcelHorizontalAlignment cellAlignment)
        {
            Range range = GetRange(startRow, startColumn, endRow, endColumn);
            range.Select();
            range.HorizontalAlignment = cellAlignment;
        }
        /// <summary>
        /// 设置单个单元格的边框
        /// </summary>
        /// <param name="cell">单元格名称</param>
        /// <param name="styleLine">线条形态</param>
        /// <param name="borderWeight">线条粗细</param>
        /// <param name="constants"></param>
        public void SetCellBorder(string cell,
            ExcelStyleLine styleLine = ExcelStyleLine.Continious, ExcelBorderWeight borderWeight = ExcelBorderWeight.Thin, Constants constants = Constants.xlAutomatic)
        {
            Range range = GetRange(cell);
            range.Select();

            // 上边框
            range.Borders[(XlBordersIndex)ExcelBordersIndex.EdgeTop].LineStyle = styleLine;
            range.Borders[(XlBordersIndex)ExcelBordersIndex.EdgeTop].Weight = borderWeight;
            range.Borders[(XlBordersIndex)ExcelBordersIndex.EdgeTop].ColorIndex = constants;
            // 下边框
            range.Borders[(XlBordersIndex)ExcelBordersIndex.EdgeBottom].LineStyle = styleLine;
            range.Borders[(XlBordersIndex)ExcelBordersIndex.EdgeBottom].Weight = borderWeight;
            range.Borders[(XlBordersIndex)ExcelBordersIndex.EdgeBottom].ColorIndex = constants;
            // 左边框
            range.Borders[(XlBordersIndex)ExcelBordersIndex.EdgeLeft].LineStyle = styleLine;
            range.Borders[(XlBordersIndex)ExcelBordersIndex.EdgeLeft].Weight = borderWeight;
            range.Borders[(XlBordersIndex)ExcelBordersIndex.EdgeLeft].ColorIndex = constants;
            // 右边框
            range.Borders[(XlBordersIndex)ExcelBordersIndex.EdgeRight].LineStyle = styleLine;
            range.Borders[(XlBordersIndex)ExcelBordersIndex.EdgeRight].Weight = borderWeight;
            range.Borders[(XlBordersIndex)ExcelBordersIndex.EdgeRight].ColorIndex = constants;
        }
        /// <summary>
        /// 设置单个单元格边框
        /// </summary>
        /// <param name="row">单元格行坐标</param>
        /// <param name="column">单元格列坐标</param>
        /// <param name="styleLine">线条形态</param>
        /// <param name="borderWeight">线条粗细</param>
        /// <param name="constants"></param>
        public void SetCellBorder(int row, int column,
           ExcelStyleLine styleLine = ExcelStyleLine.Continious, ExcelBorderWeight borderWeight = ExcelBorderWeight.Thin, Constants constants = Constants.xlAutomatic)
        {
            Range range = GetRange(row, column);
            range.Select();

            // 上边框
            range.Borders[(XlBordersIndex)ExcelBordersIndex.EdgeTop].LineStyle = styleLine;
            range.Borders[(XlBordersIndex)ExcelBordersIndex.EdgeTop].Weight = borderWeight;
            range.Borders[(XlBordersIndex)ExcelBordersIndex.EdgeTop].ColorIndex = constants;
            // 下边框
            range.Borders[(XlBordersIndex)ExcelBordersIndex.EdgeBottom].LineStyle = styleLine;
            range.Borders[(XlBordersIndex)ExcelBordersIndex.EdgeBottom].Weight = borderWeight;
            range.Borders[(XlBordersIndex)ExcelBordersIndex.EdgeBottom].ColorIndex = constants;
            // 左边框
            range.Borders[(XlBordersIndex)ExcelBordersIndex.EdgeLeft].LineStyle = styleLine;
            range.Borders[(XlBordersIndex)ExcelBordersIndex.EdgeLeft].Weight = borderWeight;
            range.Borders[(XlBordersIndex)ExcelBordersIndex.EdgeLeft].ColorIndex = constants;
            // 右边框
            range.Borders[(XlBordersIndex)ExcelBordersIndex.EdgeRight].LineStyle = styleLine;
            range.Borders[(XlBordersIndex)ExcelBordersIndex.EdgeRight].Weight = borderWeight;
            range.Borders[(XlBordersIndex)ExcelBordersIndex.EdgeRight].ColorIndex = constants;
        }
        /// <summary>
        /// 设置范围单元格的边框，包括外框与内部
        /// </summary>
        /// <param name="startCell">开始的单元格名称</param>
        /// <param name="endCell">结束的单元格名称</param>
        /// <param name="styleLine">线条形态</param>
        /// <param name="borderWeight">线条粗细</param>
        /// <param name="constants"></param>
        public void SetCellBorder(string startCell, string endCell,
          ExcelStyleLine styleLine = ExcelStyleLine.Continious, ExcelBorderWeight borderWeight = ExcelBorderWeight.Thin, Constants constants = Constants.xlAutomatic)
        {
            Range range = GetRange(startCell, endCell);
            range.Select();

            // 上边框
            range.Borders[(XlBordersIndex)ExcelBordersIndex.EdgeTop].LineStyle = styleLine;
            range.Borders[(XlBordersIndex)ExcelBordersIndex.EdgeTop].Weight = borderWeight;
            range.Borders[(XlBordersIndex)ExcelBordersIndex.EdgeTop].ColorIndex = constants;
            // 下边框
            range.Borders[(XlBordersIndex)ExcelBordersIndex.EdgeBottom].LineStyle = styleLine;
            range.Borders[(XlBordersIndex)ExcelBordersIndex.EdgeBottom].Weight = borderWeight;
            range.Borders[(XlBordersIndex)ExcelBordersIndex.EdgeBottom].ColorIndex = constants;
            // 左边框
            range.Borders[(XlBordersIndex)ExcelBordersIndex.EdgeLeft].LineStyle = styleLine;
            range.Borders[(XlBordersIndex)ExcelBordersIndex.EdgeLeft].Weight = borderWeight;
            range.Borders[(XlBordersIndex)ExcelBordersIndex.EdgeLeft].ColorIndex = constants;
            // 右边框
            range.Borders[(XlBordersIndex)ExcelBordersIndex.EdgeRight].LineStyle = styleLine;
            range.Borders[(XlBordersIndex)ExcelBordersIndex.EdgeRight].Weight = borderWeight;
            range.Borders[(XlBordersIndex)ExcelBordersIndex.EdgeRight].ColorIndex = constants;
            // 内部水平
            range.Borders[(XlBordersIndex)ExcelBordersIndex.InsideHorizontal].LineStyle = styleLine;
            range.Borders[(XlBordersIndex)ExcelBordersIndex.InsideHorizontal].Weight = borderWeight;
            range.Borders[(XlBordersIndex)ExcelBordersIndex.InsideHorizontal].ColorIndex = constants;
            // 内部垂直
            range.Borders[(XlBordersIndex)ExcelBordersIndex.InsideVertical].LineStyle = styleLine;
            range.Borders[(XlBordersIndex)ExcelBordersIndex.InsideVertical].Weight = borderWeight;
            range.Borders[(XlBordersIndex)ExcelBordersIndex.InsideVertical].ColorIndex = constants;
        }
        public void SetCellBorder(int startRow, int startColumn, int endRow, int endColumn,
         ExcelStyleLine styleLine = ExcelStyleLine.Continious, ExcelBorderWeight borderWeight = ExcelBorderWeight.Thin, Constants constants = Constants.xlAutomatic)
        {
            Range range = GetRange(startRow, startColumn, endRow, endColumn);
            range.Select();

            // 上边框
            range.Borders[(XlBordersIndex)ExcelBordersIndex.EdgeTop].LineStyle = styleLine;
            range.Borders[(XlBordersIndex)ExcelBordersIndex.EdgeTop].Weight = borderWeight;
            range.Borders[(XlBordersIndex)ExcelBordersIndex.EdgeTop].ColorIndex = constants;
            // 下边框
            range.Borders[(XlBordersIndex)ExcelBordersIndex.EdgeBottom].LineStyle = styleLine;
            range.Borders[(XlBordersIndex)ExcelBordersIndex.EdgeBottom].Weight = borderWeight;
            range.Borders[(XlBordersIndex)ExcelBordersIndex.EdgeBottom].ColorIndex = constants;
            // 左边框
            range.Borders[(XlBordersIndex)ExcelBordersIndex.EdgeLeft].LineStyle = styleLine;
            range.Borders[(XlBordersIndex)ExcelBordersIndex.EdgeLeft].Weight = borderWeight;
            range.Borders[(XlBordersIndex)ExcelBordersIndex.EdgeLeft].ColorIndex = constants;
            // 右边框
            range.Borders[(XlBordersIndex)ExcelBordersIndex.EdgeRight].LineStyle = styleLine;
            range.Borders[(XlBordersIndex)ExcelBordersIndex.EdgeRight].Weight = borderWeight;
            range.Borders[(XlBordersIndex)ExcelBordersIndex.EdgeRight].ColorIndex = constants;
            // 内部水平
            range.Borders[(XlBordersIndex)ExcelBordersIndex.InsideHorizontal].LineStyle = styleLine;
            range.Borders[(XlBordersIndex)ExcelBordersIndex.InsideHorizontal].Weight = borderWeight;
            range.Borders[(XlBordersIndex)ExcelBordersIndex.InsideHorizontal].ColorIndex = constants;
            // 内部垂直
            range.Borders[(XlBordersIndex)ExcelBordersIndex.InsideVertical].LineStyle = styleLine;
            range.Borders[(XlBordersIndex)ExcelBordersIndex.InsideVertical].Weight = borderWeight;
            range.Borders[(XlBordersIndex)ExcelBordersIndex.InsideVertical].ColorIndex = constants;
        }

        #endregion
    }
}
