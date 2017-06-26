using System;
using System.Collections.Generic;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Ganss.Excel;
using NPOI.HSSF.UserModel;

namespace Ganss.Excel
{
    /// <ExcelFileReader>
    /// Add mapping columns in Excel to names in Ttype
    /// </ExcelFileReader>
    public class ExcelFileReader<Ttype> where Ttype : new()
    {
        private string excelFilePath;
        private Dictionary<string, string> replacements = null;

        /// <summary>
        /// Initializes a new instance of the ExcelFileReader> class.
        /// </summary>
        private ExcelFileReader() { }

        /// <summary>
        /// Initializes a new instance of the ExcelFileReader> class.
        /// </summary>
        /// <param name="excelFilePath">The path to the Excel file.</param>
        public ExcelFileReader(string excelFilePath)
        {
            this.excelFilePath = excelFilePath;
        }

        /// <summary>
        /// Initializes a new instance of the ExcelFileReader> class.
        /// </summary>
        /// <param name="excelFilePath">The path to the Excel file.</param>
        /// <param name="replacements">Map Excel file names to new names.</param>
        public ExcelFileReader(string excelFilePath, Dictionary<string, string> replacements)
            : this(excelFilePath)
        {
            this.replacements = replacements;

        }

        /// <summary>
        /// Reads the file at excelFilePath and converts it to a list
        /// of SiterraProjects.  Returns null if file is not found, file
        /// is not formatted correctly or file is empty.
        /// 
        /// File must have the data in the first work sheet
        /// File must have a header row that meets the excectations
        /// File must be XLXS
        /// </summary>
        /// <param name="excelFilePath"></param>
        /// <returns></returns>
        public List<Ttype> getProjects(string excelFilePath)
        {

            List<Ttype> projList = new List<Ttype>();

            //First use NOIP to get the workbook
            IWorkbook workbook = getIWorkbook(excelFilePath);

            //Get header row
            IRow headerRow = formatHeaders(getRow(getSheet(workbook, 0), 0));

            if (headerRow != null)
            {
                //Write updated header to file
                writeFile(excelFilePath, workbook);
                //Use ExcelMapper to map the file to list of objects
                projList = getProjectsFromExcelFile(excelFilePath);
            }

            return projList;
        }

        /// <summary>
        /// Only use if the constructor with the file path was used.
        /// </summary>
        /// <returns></returns>
        public List<Ttype> getProjects()
        {
            if (this.excelFilePath != null)
                return getProjects(this.excelFilePath);
            else
            {
                System.Diagnostics.Debug.WriteLine("The path to the excel file must be set.");
                return null;
            }
        }

        /// <summary>
        /// Create workbook from file at filePath.
        /// </summary>
        /// <param name="filePath">The path to the Excel file.</param>
        public IWorkbook getIWorkbook(string filePath)
        {
            try
            {
                IWorkbook workbook = new HSSFWorkbook();

                using (var fs = File.Open(@filePath, FileMode.Open))
                {
                    if (filePath.EndsWith(".xlsx"))
                        workbook = new XSSFWorkbook(fs);
                    else
                        workbook = new HSSFWorkbook(fs);

                }

                return workbook;
            }
            catch (Exception e)
            {
                System.Diagnostics.Debug.WriteLine(e.StackTrace);
                return null;
            }
        }

        /// <summary>
        /// Create worksheet.
        /// </summary>
        /// <param name="workbook">Workbook where sheet is located.</param>
        /// <param name="sheetIndex">Index number of the sheet to get.</param>
        public ISheet getSheet(IWorkbook workbook, int sheetIndex)
        {
            if (workbook != null && workbook.NumberOfSheets > 0)
            {
                ISheet sheet = workbook.GetSheetAt(sheetIndex);
                return sheet;
            }

            return null;
        }

        /// <summary>
        /// Get row rowIndex of sheet.  returns null if sheet has no
        /// rows or the row has not data.
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowIndex"></param>
        /// <returns></returns>
        public IRow getRow(ISheet sheet, int rowIndex)
        {
            if (sheet != null)
            {
                IRow row = sheet.GetRow(rowIndex);
                if (row == null)
                {
                    System.Diagnostics.Debug.WriteLine("Row " + rowIndex + " is not defined for sheet " + sheet.SheetName);
                    return row;
                }
                else if (row.Cells != null && row.Cells.Count > 0)
                {
                    return row;
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("Row " + rowIndex + " does not containt any data for sheet " + sheet.SheetName);
                }
                return row;
            }

            return null;
        }

        /// <summary>
        /// Format a row from a worksheet
        /// </summary>
        /// <param name="row">Row to format.</param>
        public virtual IRow formatHeaders(IRow row)
        {

            if (replacements != null)
            {
                if (row != null)
                {
                    foreach (ICell cell in row)
                    {
                        var val = cell.StringCellValue;
                        if (val != null)
                        {
                            foreach (KeyValuePair<string, string> entry in replacements)
                            {
                                val = val.Replace(entry.Key, entry.Value);
                            }
                        }
                        cell.SetCellValue(val);
                    }
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("Excel file must have a header row.");
                }
            }

            return row;
        }

        /// <summary>
        /// Write out workbook with new header
        /// </summary>
        /// <param name="filePath">Where to write workbook.</param>
        /// <param name="workbook">Workbook to write</param>
        public bool writeFile(string filePath, IWorkbook workbook)
        {
            try
            {
                FileStream file = new FileStream(@filePath, FileMode.Create);
                workbook.Write(file);
                file.Close();
                return true;
            }
            catch (Exception e)
            {
                System.Diagnostics.Debug.WriteLine(e.StackTrace);
                return false;
            }
        }

        /// <summary>
        /// Use Excel mapper to map the Excel to a Type are return it.
        /// </summary>
        /// <param name="filePath">Location of the workbook to map.</param>
        public List<Ttype> getProjectsFromExcelFile(string filePath)
        {
            int i = 0;
            try
            {
                List<Ttype> l = new List<Ttype>();

                ExcelMapper excel = new ExcelMapper(@filePath);
                IEnumerable<Ttype> enumerable = excel.Fetch<Ttype>();

                List<Ttype> list = new List<Ttype>();
                foreach (Ttype tt in enumerable)
                {
                    i++;
                    list.Add(tt);
                }

                return list;
            }
            catch (Exception e)
            {
                System.Diagnostics.Debug.WriteLine(i + " ================\r\n" + e.StackTrace);
            }

            return null;
        }

    }
}