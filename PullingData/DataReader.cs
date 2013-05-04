using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Excel = Microsoft.Office.Interop.Excel;

namespace PullingData
{
    class DataReader
    {
        private Excel.Application app;
        private Excel.Workbook book;
        private Excel.Worksheet sheet;

        private int sheetCount;
        private int currentSheetIndex;

        private int currentRowCount;
        private int currentRow;

        private int currentColCount;
        private int currentCol;

        private Excel.Range range;
        private Excel.Range row;

        public DataReader(String fileName)
        {
            app = new Excel.Application();
            book = app.Workbooks.Open(fileName);

            sheetCount = book.Sheets.Count;
            currentSheetIndex = 0;

            currentRowCount = -1;
            currentRow = 0;
            currentColCount = -1;
            currentCol = 0;
        }

        public String getCurrentSheetName()
        {
            return sheet.Name;
        }

        public bool hasNextSheet()
        {
            return currentSheetIndex < sheetCount;
        }

        public void nextSheet()
        {
            currentSheetIndex++;
            sheet = book.Sheets[currentSheetIndex];
            range = sheet.UsedRange;

            currentRowCount = range.Rows.Count;
            currentRow = 0;
            currentColCount = range.Columns.Count;
            currentCol = currentColCount + 1;
        }

        public bool hasNextRow()
        {
            return currentRow < currentRowCount;
        }

        public void nextRow()
        {
            currentRow++;
            currentCol = 0;
            row = range.Rows[currentRow];
        }

        public bool hasNextCol()
        {
            return currentCol <= row.Columns.Count;
        }

        public string nextCol()
        {
            currentCol++;
            Excel.Range cell = row.Cells[1, currentCol];
            if (cell.Value2 == null)
            {
                return null;
            }
            else
            {
                return cell.Value2.ToString();
            }
        }

        public int nextInt()
        {
            string data = nextCol();
            if (data == null)
            {
                return 0;
            }
            else if (data.Length == 0)
            {
                return 0;
            }
            else
            {
                return Convert.ToInt32(data);
            }
        }

        public bool nextBool()
        {
            string data = nextCol();
            if (data != null)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public void close()
        {
            book.Close(false);
            app.Quit();
        }
    }
}
