/* Copyright (C) 2011  Jan Källman
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "as is" with no warranty either express or implied.
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 *
 * Author Change                      Date
 *******************************************************************************
 * Mats Alm Added		                2016-12-27
 *******************************************************************************/

using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing
{
    /// <summary>
    /// This class should be implemented to be able to deliver excel data
    /// to the formula parser.
    /// </summary>
    public abstract class ExcelDataProvider : IDisposable
    {
        /// <summary>
        /// Information and help methods about a cell
        /// </summary>
        public interface ICellInfo
        {
            string Address { get; }
            int Column { get; }
            string Formula { get; }
            bool IsExcelError { get; }
            bool IsHiddenRow { get; }
            int Row { get; }
            IList<Token> Tokens { get; }
            object Value { get; }
            double ValueDouble { get; }
            double ValueDoubleLogical { get; }
        }

        public interface INameInfo
        {
            string Formula { get; set; }
            ulong Id { get; set; }
            string Name { get; set; }
            IList<Token> Tokens { get; }
            object Value { get; set; }
            string Worksheet { get; set; }
        }

        /// <summary>
        /// A range of cells in a worksheet.
        /// </summary>
        public interface IRangeInfo : IEnumerator<ICellInfo>, IEnumerable<ICellInfo>
        {
            ExcelAddressBase Address { get; }
            bool IsEmpty { get; }
            bool IsMulti { get; }
            ExcelWorksheet Worksheet { get; }

            int GetNCells();

            object GetOffset(int rowOffset, int colOffset);

            object GetValue(int row, int col);
        }

        /// <summary>
        /// Max number of columns in a worksheet that the Excel data provider can handle.
        /// </summary>
        public abstract int ExcelMaxColumns { get; }

        /// <summary>
        /// Max number of rows in a worksheet that the Excel data provider can handle
        /// </summary>
        public abstract int ExcelMaxRows { get; }

        /// <summary>
        /// Use this method to free unmanaged resources.
        /// </summary>
        public abstract void Dispose();

        /// <summary>
        /// Returns a single cell value
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        public abstract object GetCellValue(string sheetName, int row, int col);

        /// <summary>
        /// Returns the address of the lowest rightmost cell on the worksheet.
        /// </summary>
        /// <param name="worksheet"></param>
        /// <returns></returns>
        public abstract ExcelCellAddress GetDimensionEnd(string worksheet);

        public abstract string GetFormat(object value, string format);

        public abstract INameInfo GetName(string worksheet, string name);

        /// <summary>
        /// Returns values from the required range.
        /// </summary>
        /// <param name="worksheetName">The name of the worksheet</param>
        /// <param name="row">Row</param>
        /// <param name="column">Column</param>
        /// <param name="address">The reference address</param>
        /// <returns></returns>
        public abstract IRangeInfo GetRange(string worksheetName, int row, int column, string address);

        /// <summary>
        /// Returns values from the required range.
        /// </summary>
        /// <param name="worksheetName">The name of the worksheet</param>
        /// <param name="address">The reference address</param>
        /// <returns></returns>
        public abstract IRangeInfo GetRange(string worksheetName, string address);

        public abstract IRangeInfo GetRange(string worksheet, int fromRow, int fromCol, int toRow, int toCol);

        public abstract string GetRangeFormula(string worksheetName, int row, int column);

        public abstract List<Token> GetRangeFormulaTokens(string worksheetName, int row, int column);

        ///// <summary>
        ///// Returns a single cell value
        ///// </summary>
        ///// <param name="address"></param>
        ///// <returns></returns>
        //public abstract object GetCellValue(int sheetID, string address);
        ///// <summary>
        ///// Sets the value on the cell
        ///// </summary>
        ///// <param name="address"></param>
        ///// <param name="value"></param>
        //public abstract void SetCellValue(string address, object value);
        public abstract object GetRangeValue(string worksheetName, int row, int column);

        public abstract IEnumerable<object> GetRangeValues(string address);

        /// <summary>
        /// Returns all defined names in a workbook
        /// </summary>
        /// <returns></returns>
        public abstract ExcelNamedRangeCollection GetWorkbookNameValues();

        /// <summary>
        /// Returns the names of all worksheet names
        /// </summary>
        /// <returns></returns>
        public abstract ExcelNamedRangeCollection GetWorksheetNames(string worksheet);

        public abstract bool IsRowHidden(string worksheetName, int row);

        public abstract void Reset();
    }
}