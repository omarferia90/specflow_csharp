using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace SpecFlow_CSharp.Support
{
    /// <summary>
    /// Class to interact with excel files trought Spreadsheetlight framework.
    /// </summary>
    public class DataLoad
    {
        //Variables
        private SLDocument _objFile;
        private int _intCurrentRow;
        private bool _blColumnNames;
        private bool _blSheetInUse;
        private int _intColCount;
        private int _intRowCount;
        private int _intColStartIndex;
        private int _intRowStartIndex;
        private Dictionary<string, int> _dicHeader;

        /// <summary>
        /// Get Current Row Selected
        /// </summary>
        public int CurrentRow
        {
            get { return _intCurrentRow; }
            set { _intCurrentRow = (value > _intRowStartIndex ? value : _intRowStartIndex); }
        }

        /// <summary>
        /// Variable to identify if the spreadsheet has headers/column names
        /// </summary>
        public bool HasColumnNames
        {
            get { return _blColumnNames; }
            set
            {
                _blColumnNames = value;
                _intRowStartIndex = (value ? 2 : 1);
            }
        }

        /// <summary>
        /// Get total number of columns
        /// </summary>
        public int ColumnCount
        {
            get { return _intColCount; }
        }

        /// <summary>
        /// Get total number of rows of the current spreadsheet
        /// </summary>
        public int RowCount
        {
            get { return _intRowCount; }
        }

        /// <summary>
        /// Defines the row start index of the spreadsheet
        /// </summary>
        private int RowStartIndex
        {
            get { return _intRowStartIndex; }
            set
            {
                _intRowStartIndex = value;
                _intCurrentRow = (_intCurrentRow > _intRowStartIndex ? _intCurrentRow : _intRowStartIndex);
            }
        }

        /// <summary>
        /// Defines the column start index of the spreadsheet
        /// </summary>
        private int ColumnStartIndex
        {
            get { return _intColStartIndex; }
            set { _intColStartIndex = value; }
        }

        //Constructor
        public DataLoad()
        {
            _intCurrentRow = 2;
            _blColumnNames = true;
            _blSheetInUse = true;
            _intColCount = -1;
            _intRowCount = -1;
            _intRowStartIndex = 2;
            _intColStartIndex = 1;
        }


        private void _GetRowCount()
        {
            if (_objFile != null && _blSheetInUse)
            {
                _intRowCount = _objFile.GetWorksheetStatistics().NumberOfRows > -1 ? _objFile.GetWorksheetStatistics().NumberOfRows : 0;
            }
        }

        private void _GetColumCount()
        {
            if (_objFile != null && _blSheetInUse)
            {
                _intColCount = _objFile.GetWorksheetStatistics().NumberOfColumns > -1 ? _objFile.GetWorksheetStatistics().NumberOfColumns : 0;
            }
        }

        /// <summary>
        /// Loads a spreadsheet selected
        /// </summary>
        /// <param name="pstrFilePath">File Path</param>
        /// <param name="pstrSheet">Spreadsheet to select</param>
        public void fnLoadFile(string pstrFilePath, string pstrSheet)
        {
            if (!string.IsNullOrEmpty(pstrFilePath) && File.Exists(pstrFilePath))
            {
                _objFile = new SLDocument(pstrFilePath);
                if (!string.IsNullOrEmpty(pstrSheet))
                {
                    if (_objFile.GetSheetNames().Contains(pstrSheet)) { _objFile.SelectWorksheet(pstrSheet); }
                }
                _blColumnNames = true;
                _blSheetInUse = true;
                _intCurrentRow = 2;
                _intRowStartIndex = _objFile.GetWorksheetStatistics().StartRowIndex;
                _intColStartIndex = _objFile.GetWorksheetStatistics().StartColumnIndex;
                _dicHeader = _GetHeaders(_objFile);
                _GetRowCount();
                _GetColumCount();

            }
        }

        /// <summary>
        /// Function to remove special characters
        /// </summary>
        /// <param name="strValue"></param>
        /// <returns></returns>
        private string fnRemoveEspecialChar(string strValue)
        {
            Regex r = new Regex("(?:[^a-z0-9 ]|(?<=['\"])s)", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);
            return r.Replace(strValue, String.Empty);
        }

        /// <summary>
        /// Function to add all the headers into a dictionary
        /// </summary>
        /// <param name="pobjDoc"></param>
        /// <returns></returns>
        private Dictionary<string, int> _GetHeaders(SLDocument pobjDoc)
        {
            if (pobjDoc != null)
            {
                Dictionary<string, int> dicHeaders = new Dictionary<string, int>();
                for (int cols = 1; cols <= pobjDoc.GetWorksheetStatistics().NumberOfColumns; cols++)
                {
                    if (pobjDoc.GetCellValueAsString(1, cols) != "")
                    {
                        dicHeaders.Add(fnRemoveEspecialChar(pobjDoc.GetCellValueAsString(1, cols)), cols);
                    }
                    else
                    {
                        break;
                    }
                }
                return dicHeaders;
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Function to retrive the value of the current sheet selected
        /// </summary>
        /// <param name="pstrColumnName"></param>
        /// <param name="pstrDefaultValue"></param>
        /// <returns></returns>
        public string fnGetValue(string pstrColumnName, [Optional] string pstrDefaultValue)
        {
            string strTempVal = "";
            if (!string.IsNullOrEmpty(pstrDefaultValue)) { strTempVal = pstrDefaultValue; }
            if (string.IsNullOrEmpty(pstrDefaultValue)) { pstrDefaultValue = string.Empty; }
            if (_objFile != null && _blSheetInUse && _dicHeader.Count > 0)
            {
                if (_dicHeader.ContainsKey(pstrColumnName))
                {
                    pstrDefaultValue = _objFile.GetCellValueAsString(_intCurrentRow, _dicHeader[pstrColumnName]);
                    //Adding +Date code
                    if (pstrDefaultValue.StartsWith("+") || pstrDefaultValue.StartsWith("-"))
                    {
                        DateTime dtNewDate;
                        dtNewDate = DateTime.Today.AddDays(Convert.ToDouble(pstrDefaultValue));
                        pstrDefaultValue = dtNewDate.ToString("MM/dd/yyyy");
                    }
                    if (pstrDefaultValue.ToUpper().Equals("TODAY"))
                    {
                        pstrDefaultValue = DateTime.Today.ToString("MM/dd/yyyy");
                    }
                    //Get Default Value
                    if (pstrDefaultValue == "" && strTempVal != "")
                    { pstrDefaultValue = strTempVal; }
                }
                else
                {
                    pstrDefaultValue = "";
                }
            }
            return pstrDefaultValue;
        }

        /// <summary>
        /// Function to save a new value into the sheet provided
        /// </summary>
        /// <param name="pstrPath"></param>
        /// <param name="pstrSheet"></param>
        /// <param name="pstrColumn"></param>
        /// <param name="pintRow"></param>
        /// <param name="pstrValue"></param>
        public void fnSaveValue(string pstrPath, string pstrSheet, string pstrColumn, int pintRow, string pstrValue)
        {
            SLDocument _document = new SLDocument(pstrPath);
            if (!string.IsNullOrEmpty(pstrSheet))
            {
                if (_document.GetSheetNames().Contains(pstrSheet)) { _document.SelectWorksheet(pstrSheet); }
                Dictionary<string, int> _dicHeaderE = _GetHeaders(_document);
                if (_dicHeaderE.ContainsKey(pstrColumn))
                {
                    _document.SetCellValue(pintRow, _dicHeaderE[pstrColumn], pstrValue);
                    _document.Save();
                }
            }
        }

    }
}
