using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Koanvi.Excel{
	using Excel = Microsoft.Office.Interop.Excel;

	public class Application :IDisposable{

		public Excel.Application excelApp;
		public Excel.Workbook excelWorkbook;
		public String FileName;

		public Application() {
			excelApp = new Excel.Application();
			excelApp.Visible = true;
		}
		public Application(string filePath):this() {Openfile(filePath);}

		public void CreateWb() {
			excelWorkbook = excelApp.Workbooks.Add(System.Reflection.Missing.Value);
		}

		public void TableToExcel(System.Data.DataTable dataTable) {

			AddSheet(dataTable.TableName);
			var tn = dataTable.TableName;
			if (tn.Length >= 31) { tn = tn.Substring(0, 30); }
			var ws = (Excel.Worksheet)excelWorkbook.Sheets[tn];

			int iRow = 1; int iCol = 1;

			var calculation = excelApp.Calculation;
			var ScreenUpdating = excelApp.ScreenUpdating;
			excelApp.Calculation = Excel.XlCalculation.xlCalculationManual;
			excelApp.ScreenUpdating = false;

			dataTable.Columns.Cast<System.Data.DataColumn>().ToList().ForEach(col => {
				ws.Cells[iRow, iCol] = col.ColumnName;
				if (col.DataType==typeof(string)) {
					((Excel.Range)ws.Columns[iCol]).NumberFormat = "@";
				}
				iCol++;
			});

			iRow = 0; iCol = 0;

			dataTable.Columns.Cast<System.Data.DataColumn>().ToList().ForEach(col => {
				iCol++; iRow = 1;
				dataTable.Rows.Cast<System.Data.DataRow>().ToList().ForEach(row => {
					iRow++;
					//ws.Cells[iRow,iCol]=row[col].ToString();

					if (col.DataType == typeof(DateTime)) {
						if (Convert.ToDateTime(row[col])==new DateTime(0)) {
							return;
						}
					}
					ws.Cells[iRow, iCol] = row[col];

				});

			});

			excelApp.Calculation= calculation;
			excelApp.ScreenUpdating= ScreenUpdating;

		}

		public void AddSheet(string Name) {
			if(excelWorkbook==null) { return; }
			if (Name == string.Empty) { return; }
			if (Name.Length>=31) { Name = Name.Substring(0, 30); }
			
			try {
				var ws = (Excel.Worksheet)excelWorkbook.Sheets.Add();
				ws.Name = Name;
			}
			catch (Exception ex) {

				throw;
			}
		}

		public void Openfile() {
			System.Windows.Forms.OpenFileDialog openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
			if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK) {
				openFileDialog1.Filter= "Excel Files|*.xls;*.xlsx;*.xlsm";
				Openfile(openFileDialog1.FileName);
			}
		}
		public void Openfile(string path) {
			if (!System.IO.File.Exists(path)) {
				this.Close();
				throw new Exception(@"Такого файла нет!");
			}
			this.FileName = System.IO.Path.GetFileNameWithoutExtension(path);
			excelWorkbook = excelApp.Workbooks.Open(path,
				0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "",
				true, false, 0, true, false, false);
			excelApp.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlMinimized;
		}

		private void OpenfileExcel() {
			//Microsoft.Office.Core.MsoFileDialogType.msoFileDialogOpen
			//Microsoft.Office.Core.MsoFileDialogType.msoFileDialogOpen
		}

		public void Close() {Dispose();}

		public System.Data.DataTable GetRange(String WorksheetName, Range range) {

			if (excelWorkbook is null) { return null; }
			var ws = excelWorkbook.Worksheets[WorksheetName] as Excel.Worksheet;
			return GetRange(ws, range);

		}
		public System.Data.DataTable GetRange(String WorksheetName) {

			if (excelWorkbook is null) { return null; }
			var ws = excelWorkbook.Worksheets[WorksheetName] as Excel.Worksheet;

			Excel.Range last = ws.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
			Excel.Range range = ws.get_Range("A1", last);
			var dtbl =  GetRange(ws, range);
			if (dtbl==null) { return null; }
			dtbl.TableName = WorksheetName;
			return dtbl;

		}
		private System.Data.DataTable GetRange(Excel.Worksheet worksheet, Range range) {
			var Range = worksheet.Range[
				worksheet.Cells[range.StartCell.Row, range.StartCell.Col],
				worksheet.Cells[range.EndCell.Row, range.EndCell.Col]
				];
			return GetRange(worksheet, Range);
		}
		private System.Data.DataTable GetRange(Excel.Worksheet worksheet, Excel.Range Range) {

			object[,] arr1 = Range.Value;
			object[,] arr2=new object[0,0];
			//Array.Copy(arr1, arr2, 0);
			if (arr1==null) { return null; }
			var maxIRow = arr1.GetLongLength(0);
			var maxICol=arr1.GetLongLength(1);
			
			var tbl = new System.Data.DataTable();

			for (int iCol = 0; iCol < maxICol; iCol++) {tbl.Columns.Add();}
			for (int iRow = 0; iRow < maxIRow; iRow++) {var tblRow = tbl.NewRow();tbl.Rows.Add(tblRow);}
			for (int iCol = 0; iCol < maxICol; iCol++){
				for (int iRow = 0; iRow < maxIRow; iRow++) {tbl.Rows[iRow ][iCol ] = arr1[iRow +1, iCol+1];}
			}
			return tbl;
		}

		public List<String> Sheets { get {
				if (excelWorkbook==null) { return null; }
				var sheets = excelWorkbook.Sheets;
				int count = sheets.Count;
				var retval = new List<string>();


				foreach (Excel.Worksheet worksheet in excelWorkbook.Worksheets) {
					retval.Add(worksheet.Name);
				}
				return retval;

			} }

		/// <summary>
		/// докрутить диспоуз по человечески
		/// </summary>
		public void Dispose() {
			if (excelWorkbook != null) {
				excelWorkbook.Close();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorkbook);
			}
			if (excelApp != null) {
				excelApp.Quit();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
			}
		}

	}//public class Application :IDisposable

	public class Cell {
		int _Row=1;
		int _Col=1;
		public int Row { get { return _Row; } set { if (value < 1) { throw new Exception(@"Номер строки не может быть менее 1"); } _Row = value; } }
		public int Col { get { return _Col; } set { if (value < 1) { throw new Exception(@"Номер строки не может быть менее 1"); } _Col = value; } }
		public Cell(int Row, int Col) {
			this.Row = Row;
			this.Col = Col;
		}
	}//public class Cell

	public class Range {
		public Cell StartCell;
		public Cell EndCell;
		public Range(Cell startCell, Cell endCell) {
			this.StartCell = startCell;
			this.EndCell = endCell;
		}
	}

}
