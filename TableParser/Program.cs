using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Koanvi.TableProcessor {
	static class Program {
		/// <summary>
		/// Главная точка входа для приложения.
		/// </summary>
		[STAThread]
		static void Main() {
			Application.EnableVisualStyles();
			Application.SetCompatibleTextRenderingDefault(false);

			//TestExcel();
			TestTableProcessor();
			//TestMergeDT();
			//TestJoinDT();
			//TestChangeDatataype();
		}

		public static void TestTableProcessor() {

			var tp = new ABV.TableProcessor.TableProcessor2();
			//var tp = new ABV.TableProcessor.TableProcessor1();
			tp.Process();
			tp.ExportToExcel();
			//var frm = new Koanvi.Test.Forms.Form1();
			//frm.dataGridView1.DataSource = tp.DataSetTables.Tables[@"На импорт новые договоры"];

			//Application.Run(frm);

			MessageBox.Show(@"Готово!",@"Обработка файла");
			
		}

		public static void TestExcel() {

			var dtToImport = TestDT(@"TestDT");

			Koanvi.Excel.Application excel = new Excel.Application();
			//var asd = excel.GetRange(@"123", new Koanvi.Excel.RangePoints(
			//	new Koanvi.Excel.CellPoints(1, 1), new Koanvi.Excel.CellPoints(4, 5)));
			
			excel.CreateWb();

			//var tbl = excel.GetRange(@"123");
			var Sheets = excel.Sheets;

			//excel.AddSheet(dtToImport.TableName);
			excel.TableToExcel(dtToImport);
			//Sheets = excel.Sheets;
			var dtToExport=excel.GetRange(dtToImport.TableName);
			

			excel.Close();

			var frm = new Koanvi.Test.Forms.Form1();
			frm.dataGridView1.DataSource = dtToExport;

			Application.Run(frm);
			
		}

		public static System.Data.DataTable TestDT(string tableName) {
			var dt = new System.Data.DataTable();
			dt.TableName = tableName;
			dt.Columns.Add();
			dt.Columns.Add();
			dt.Columns.Add(@"дата",typeof(DateTime));
			dt.Columns.Add(@"Счет", typeof(string));
			var asd = Enumerable.Range(0, 10).Select(x => dt.NewRow()).ToList();
			int iRow=0;
			asd.ToList().ForEach(row => {
				++iRow;
				row[0] = @"asdasdas" + iRow.ToString();
				row[1] = @"col1" + iRow.ToString();
				row[2] = DateTime.Now;
				row[3] = @"00000000000000000000";
				dt.Rows.Add(row);
			});
			return dt;

		}

		public static void TestMergeDT() {
			var ds = CreateToTest();
			var dt1 = ds.Tables[0];
			var dt2 = ds.Tables[1];

			dt1.Merge(dt2);
			ShowDatatable(dt1);
		}

		public static void ShowDatatable(System.Data.DataTable dt) {

			var frm = new Koanvi.Test.Forms.Form1();
			frm.dataGridView1.DataSource = dt;
			Application.Run(frm);

		}

		public static System.Data.DataSet CreateToTest() {
			var ds = new System.Data.DataSet();

			var dt1 = new System.Data.DataTable();
			dt1.Columns.Add(@"t1c1");
			dt1.Columns.Add(@"t1c2");

			var row = dt1.NewRow();
			row[0] = @"1";
			row[1] = @"01.01.2017";
			dt1.Rows.Add(row);

			row = dt1.NewRow();
			row[0] = @"2";
			row[1] = @"01.01.2017";
			dt1.Rows.Add(row);

			row = dt1.NewRow();
			row[0] = @"3";
			row[1] = @"01.01.2017";
			dt1.Rows.Add(row);

			var dt2 = new System.Data.DataTable();
			dt2.Columns.Add(@"t2c1");
			dt2.Columns.Add(@"t2c2");

			row = dt2.NewRow();
			row[0] = @"1";
			row[1] = @"t2r1c2";
			dt2.Rows.Add(row);

			row = dt2.NewRow();
			row[0] = @"2";
			row[1] = @"t2r2c2";
			dt2.Rows.Add(row);

			row = dt2.NewRow();
			row[0] = @"4";
			row[1] = @"t2r3c2";
			dt2.Rows.Add(row);

			ds.Tables.Add(dt1);
			ds.Tables.Add(dt2);


			return ds;
		}

		public static void TestChangeDatataype() {

			var ds = CreateToTest();
			var dt = ds.Tables[0];

			var type = typeof(DateTime);
			var colName = @"t1c2";
			var iCol = dt.Columns.IndexOf(colName);

			var dt2 = dt.Clone();
			
			dt2.Columns[iCol].DataType = typeof(DateTime);
			dt.Rows.Cast<System.Data.DataRow>().ToList().ForEach(dr => {

				object[] objData = dr.ItemArray;
				objData[iCol]=Convert.ChangeType(objData[iCol], type);
				var nr = dt2.NewRow();
				nr.ItemArray = objData;

			});

		}

		public static void TestChangeDatataype1() {

			var ds = CreateToTest();
			var dt = ds.Tables[0];
			var dt2 = dt.Clone();
			dt2.Columns[0].DataType = typeof(DateTime);
			dt.Rows.Cast<System.Data.DataRow>().ToList().ForEach(x => {
				dt2.ImportRow(x);

			});

		}
		
		/// <summary>
		/// не работает!
		/// </summary>
		public static void TestChangeDatataype2() {

			var ds = CreateToTest();
			var colData = new List<object>();
			ds.Tables[0].Columns[0].AllowDBNull = true;
			ds.Tables[0].Rows.Cast<System.Data.DataRow>().ToList().ForEach(row => {
				colData.Add(row[@"t1c1"]);
				row[@"t1c1"] = DBNull.Value;
			});

			ds.Tables[0].Columns[@"t1c1"].DataType = typeof(DateTime);


		}

	}
}

