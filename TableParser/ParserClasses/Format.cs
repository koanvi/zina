using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Koanvi.TableProcessor {

	/// <summary>
	/// обработка таблиц
	/// </summary>
	public class TableProcessor {
		public System.Data.DataSet DataSetTables;
		public String Name { get; set; }
		public List<string> TablesSheetNames;
		public System.Data.SqlClient.SqlConnection conn;

		public TableProcessor(String Name) {
			this.Name = Name;
			var connStr = System.Configuration.ConfigurationManager.ConnectionStrings["app"].ConnectionString;
			conn = new System.Data.SqlClient.SqlConnection(connStr);
			conn.Open();
			DataSetTables = new System.Data.DataSet();
		}

		public virtual void GetInputFile(String FilePath, List<String> SheetNames) {}

	}
}

namespace ABV.TableProcessor {
	using Koanvi.TableProcessor.Format;
	using Koanvi.TableProcessor;
	using Koanvi.Excel;

	public class TableProcessor : Koanvi.TableProcessor.TableProcessor {
		public IExcelFileFormat excelFileFormat;

		public TableProcessor(string name):base(name) {
			GetSQL();

		}
		public string fileName;
		private DateTime? _ReportDate = null;
		public DateTime? ReportDate {
			get {

				if (_ReportDate != null) { return _ReportDate; }
				_ReportDate = ParseReportDate();
				return _ReportDate;
			}
		}

		private DateTime? ParseReportDate() {
			//Передача_12.05.2017_АБВ-1.xlsx
			if (fileName == null) { return null; }

			var str = fileName.Replace(@"Передача_", @"");
			str = str.Substring(0, str.IndexOf(@"_"));

			return DateTime.Parse(str);

		}

		#region Обработка столбцов

		public string Portfolioname(string status) {
			var retval = @"Передача";

			if (status == @"Банкротство") { retval = retval + @"_Банкроты"; }
			if (status == @"Отзыв персональных данных") { retval = retval + @"_Отзыв ПД"; }

			retval = retval + @"_" + ((DateTime)ReportDate).ToString(@"dd.MM.yyyy");

			return retval;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="OPD">Отзыв персональных данных</param>
		/// <param name="bankrotstvo">банкротство</param>
		/// <returns></returns>
		public string Status(string OPD, string bankrotstvo) {
			var retval = @"";
			if (OPD == @"Y") { retval = @"Отзыв персональных данных"; }
			if (bankrotstvo == @"Y") { retval = @"Банкротство"; }
			return retval;
		}

		public DateTime DataVihodaNaProsrochku(int kolichestvoDneyProsrochki) {
			return ((DateTime)ReportDate).AddDays(-kolichestvoDneyProsrochki);
		}

		public String DatetimeStringToDate(string textDate) {
			int startIndex = textDate.IndexOf(@" ");
			//int length = endIndex - startIndex + 1;
			return textDate.Substring(0,10);

		}

		public DateTime StringToDatetime(string text) {
			return Convert.ToDateTime(text);

		}
		public DateTime StringToDatetime(object text) {
			if (string.IsNullOrEmpty(text.ToString())) { return new DateTime(0); }
			return Convert.ToDateTime(text);
		}

		public string PaspDannie1(string paspDan) {
			return paspDan.Substring(0, 5).Replace(@" ", @"") + @" " + paspDan.ToString().Substring(6, 6);
		}

		public string ConcatFIO(string f, string i, string o) {
			return f + @" " + i + @" " + o;
		}

		#endregion

		public void GetSQL() {
			if (this.DataSetTables == null) { this.DataSetTables = new System.Data.DataSet(); }
			var DataTableAllDept = new Koanvi.Data.Tables.Contact.DataTableAllDept(conn);
			DataTableAllDept.TableName = @"Выгрузка всех долгов";
			DataTableAllDept.Fill();
			if (this.DataSetTables.Tables.Contains(DataTableAllDept.TableName)) {
				this.DataSetTables.Tables.Remove(DataTableAllDept.TableName);
			}
			this.DataSetTables.Tables.Add(DataTableAllDept);
		}

		/// <summary>
		/// достает файл
		/// проверяет файл
		/// проверяет листы
		/// достает таблицы
		/// </summary>
		/// <param name="excel"></param>
		public void GetInputFile(Koanvi.Excel.Application excel) {

			try {
				CreateExcelFileFormat();
				excelFileFormat.Excel = excel;
				excelFileFormat.Check();

				excelFileFormat.Sheets.ForEach(sheet => {
					sheet.Fill(excel);
					sheet.Check();
					DataSetTables.Tables.Add(sheet.Table);
				});
				this.fileName = excel.FileName;
			}
			catch (Exception ex) { excel.Close(); throw; }
			excel.Close();

		}
		public void GetInputFile(string filePath) {
			GetInputFile(new Koanvi.Excel.Application(filePath));
		}
		public void GetInputFile() {

			Koanvi.Excel.Application excel = new Koanvi.Excel.Application();
			excel.Openfile();
			GetInputFile(excel);
		}

		public virtual void CreateExcelFileFormat() {}

		public virtual Koanvi.Excel.Application ExportToExcel() {
			var excel = new Koanvi.Excel.Application();
			excel.CreateWb();
			return excel;
		}

		public virtual void Process() {}

	}

	/// <summary>
	/// передача
	/// </summary>
	public class TableProcessor1 : TableProcessor {
		public TableFormat1 tableFormat1;//Список
		public TableFormat2 tableFormat2;//Досье
		public TableFormat3 tableFormat3;//Доп. досье
		public TableFormat4 tableFormat4;//Изъятие

		public TableProcessor1() : base(@"Передача 1") {


			//GetInputFile(@"C:\Users\koanvi\Downloads\Передача_19.05.2017_АБВ.xlsx");
			//GetInputFile(@"C:\Users\koanvi\Downloads\Передача_05.05.2017_АБВ.xlsx");
			GetInputFile();

		}

		/// <summary>
		/// Делаем список договоров, которые в БД и которых нет в БД
		/// </summary>
		public void GetAtDatabase() {

			var spisok = DataSetTables.Tables[@"Список"].Rows.Cast<System.Data.DataRow>().Select(x =>
			 new {
				 contract = x[@"Номер договора"]
			 }).ToList();

			var sql = DataSetTables.Tables[@"Выгрузка всех долгов"].Rows.Cast<System.Data.DataRow>().Select(x =>
			 new {
				 contract = x[@"№ Договора"],
				 debt_d = x[@"debt_id"],
			 }).ToList();


			var notIn = spisok.Select(x => x.contract).ToList().Except(sql.Select(x => x.contract).ToList()).ToList().Select(x => new { contract = x }).ToList();
			//var In = spisok.Intersect(sql).ToList();

			var In = (
				from sp in spisok
				join sq in sql on sp.contract equals sq.contract
				select new { spisokSelect = sp, sqlSelect = sq }).ToList();


			var dt1 = new System.Data.DataTable();
			dt1.Columns.Add(@"contract");
			dt1.TableName = @"Новые долги";
			DataSetTables.Tables.Add(dt1);
			notIn.ForEach(x => {
				var row = dt1.NewRow();
				dt1.Rows.Add(row);
				row[@"contract"] = x.contract;
			});

			var dt2 = new System.Data.DataTable();
			dt2.Columns.Add(@"contract");
			dt2.Columns.Add(@"debt_id");
			dt2.TableName = @"Уже загружены в БД";
			DataSetTables.Tables.Add(dt2);
			In.ForEach(x => {
				var row = dt2.NewRow();
				dt2.Rows.Add(row);
				row[@"contract"] = x.spisokSelect.contract;
				row[@"debt_id"] = x.sqlSelect.debt_d;
			});
		}

		public override Koanvi.Excel.Application ExportToExcel() {
			var excel = base.ExportToExcel();

			excel.TableToExcel(DataSetTables.Tables[@"Новые долги"]);
			excel.TableToExcel(DataSetTables.Tables[@"Уже загружены в БД"]);
			excel.TableToExcel(DataSetTables.Tables[@"На импорт новые договоры досье"]);
			excel.TableToExcel(DataSetTables.Tables[@"На импорт новые договоры Доп. досье"]);
			excel.TableToExcel(DataSetTables.Tables[@"обнов.доп.досье"]);
			excel.TableToExcel(DataSetTables.Tables[@"обнов.досье"]);
			return excel;
		}

		public override void CreateExcelFileFormat() {
			excelFileFormat = new ExcelFileFormat1();
		}

		#region DataProcessing

		public override void Process() {
			base.Process();
			GetAtDatabase();
			CreateToImportList1();
			CreateToImportAdditioanal();
			CreateToImportAdditioanalUpdate();
			CreateToImportList1Update();
		}

		/// <summary>
		/// досье
		/// </summary>
		public void CreateToImportList1() {

			var ToImport = new System.Data.DataTable();
			ToImport.TableName = @"На импорт новые договоры досье";

			#region Create Table
			ToImport.Columns.Add(@"Данные о персонах / Фамилия");
			ToImport.Columns.Add(@"Данные о персонах / Имя");
			ToImport.Columns.Add(@"Данные о персонах / Отчество");
			ToImport.Columns.Add(@"Данные о персонах / Пол");
			ToImport.Columns.Add(@"Данные о персонах / Дата рождения",typeof(DateTime));
			ToImport.Columns.Add(@"Произвольные атрибуты / MIDAS");
			ToImport.Columns.Add(@"Пасп Данные / Серия и Пасп Данные / Номер (через пробел)");
			ToImport.Columns.Add(@"Адреса фактические / почтовый индекс");
			ToImport.Columns.Add(@"Адреса фактические / Район");
			ToImport.Columns.Add(@"Адреса фактические / Город");
			ToImport.Columns.Add(@"Адреса фактические / Улица");
			ToImport.Columns.Add(@"Адреса фактические / Дом");
			ToImport.Columns.Add(@"Адреса фактические / Корпус");
			ToImport.Columns.Add(@"Адреса фактические / Квартира");
			ToImport.Columns.Add(@"Адреса регистрации / Почтовый индекс");
			ToImport.Columns.Add(@"Адреса регистрации / Район");
			ToImport.Columns.Add(@"Адреса регистрации / Город");
			ToImport.Columns.Add(@"Адреса регистрации / Улица");
			ToImport.Columns.Add(@"Адреса регистрации / Дом");
			ToImport.Columns.Add(@"Адреса регистрации / Корпус");
			ToImport.Columns.Add(@"Адреса регистрации / Квартира");
			ToImport.Columns.Add(@"Произвольные атрибуты / Идентификатор залога");
			ToImport.Columns.Add(@"Произвольные атрибуты / Адрес заложенной недвижимости");
			ToImport.Columns.Add(@"Произвольные атрибуты / Тип заложенной недвижимости");
			ToImport.Columns.Add(@"Произвольные атрибуты / Статус залоговой недвижимости");
			ToImport.Columns.Add(@"Произвольные атрибуты / Оценочная стоимость заложенной недвижимости");
			ToImport.Columns.Add(@"Произвольные атрибуты / Оценочная стоимость валюта");
			ToImport.Columns.Add(@"Произвольные атрибуты / Дата оценки заложенной недвижимости", typeof(DateTime));
			ToImport.Columns.Add(@"Произвольные атрибуты / Дата приобритения недвижииости", typeof(DateTime));
			ToImport.Columns.Add(@"Произвольные атрибуты / Тип владения");
			ToImport.Columns.Add(@"Произвольные атрибуты / Общая площадь заложенной недвижимости, м2");
			ToImport.Columns.Add(@"Произвольные атрибуты / Кадастровый номер заложенной недвижимости");
			ToImport.Columns.Add(@"Произвольные атрибуты / Марка автомобиля");
			ToImport.Columns.Add(@"Произвольные атрибуты / Модель автомобиля");
			ToImport.Columns.Add(@"Произвольные атрибуты / VIN заложенного автомобиля");
			ToImport.Columns.Add(@"Произвольные атрибуты / Оценочная стоимость залогового автомобиля");
			ToImport.Columns.Add(@"Произвольные атрибуты / Оценочная стоимость залогового автомобиля валюта");
			ToImport.Columns.Add(@"Произвольные атрибуты / Дата оценки залогового автомобиля", typeof(DateTime));
			ToImport.Columns.Add(@"Произвольные атрибуты / Год выпуска автомобиля");
			ToImport.Columns.Add(@"Данные о персонах / Доход");
			ToImport.Columns.Add(@"Произвольные атрибуты / Доход семьи Валюта");
			ToImport.Columns.Add(@"Данные о персонах / Семейное положение");
			ToImport.Columns.Add(@"Произвольные атрибуты / Количество детей");
			ToImport.Columns.Add(@"Данные о персонах / Место работы");
			ToImport.Columns.Add(@"Адреса (тексты) рабочие / Текст адреса");
			ToImport.Columns.Add(@"Произвольные атрибуты / Вид деятельности");
			ToImport.Columns.Add(@"Данные о персонах / Должность");
			ToImport.Columns.Add(@"Произвольные атрибуты / Период работы по последнему месту работы");
			ToImport.Columns.Add(@"Телефоны домашние / Номе р телефона (через запятую три)");
			ToImport.Columns.Add(@"Телефоны мобильные / Номе р телефона (через запятую три)");
			ToImport.Columns.Add(@"Телефоны рабочие / Номе р телефона (через запятую три)");
			ToImport.Columns.Add(@"Произвольные атрибуты / Регион должника");
			ToImport.Columns.Add(@"Данные о долгах / Номер договора с банком");
			ToImport.Columns.Add(@"Данные о долгах / Дата выдачи кредита", typeof(DateTime));
			ToImport.Columns.Add(@"Произвольные атрибуты / Дата окончания кредитного договора", typeof(DateTime));
			ToImport.Columns.Add(@"Данные о долгах / Название продукта");
			ToImport.Columns.Add(@"Данные о долгах / Валюта");
			ToImport.Columns.Add(@"Данные о долгах / Лицевой счет должника");
			ToImport.Columns.Add(@"Данные о долгах / Полный размер кредита");
			ToImport.Columns.Add(@"Данные о долгах / Процентная ставка по кредиту");
			ToImport.Columns.Add(@"Данные о долгах / Аннуитентный платеж");
			ToImport.Columns.Add(@"Данные о долгах / Сумма, необходимая к погашению");
			ToImport.Columns.Add(@"Данные о долгах / Основной долг");
			ToImport.Columns.Add(@"Данные о долгах / Просроченный основной долг");
			ToImport.Columns.Add(@"Данные о долгах / Проценты");
			ToImport.Columns.Add(@"Данные о долгах / Просроченные проценты");
			ToImport.Columns.Add(@"Произвольные атрибуты / Перерасход по лимиту кредитной карты");
			ToImport.Columns.Add(@"Данные о долгах / Пени");
			ToImport.Columns.Add(@"Произвольные атрибуты / Процент на просрочку по основному долгу");
			ToImport.Columns.Add(@"Произвольные атрибуты / Штраф за перерасход по лимиту (кредитные карты)");
			ToImport.Columns.Add(@"Произвольные атрибуты / Дата последнего платежа", typeof(DateTime));
			ToImport.Columns.Add(@"Произвольные атрибуты / Сумма последнего платежа");
			ToImport.Columns.Add(@"Произвольные атрибуты / Количество произведенных платежей в счет погашения задолженности");
			ToImport.Columns.Add(@"Произвольные атрибуты / Сумма произведенных платежей в счет погашения");
			ToImport.Columns.Add(@"Произвольные атрибуты / Количество дней от последнего платежа");
			ToImport.Columns.Add(@"Произвольные атрибуты / Количество просроченных дней по кредиту");
			ToImport.Columns.Add(@"Произвольные атрибуты / Наличие других кредитов в банке");
			ToImport.Columns.Add(@"Произвольные атрибуты / Период добровольного страхования с");
			ToImport.Columns.Add(@"Произвольные атрибуты / Период добровольного страхования по");
			ToImport.Columns.Add(@"Произвольные атрибуты / Страховая премия по договору страхования");
			ToImport.Columns.Add(@"Произвольные атрибуты / Страховая премия валюта");
			ToImport.Columns.Add(@"Произвольные атрибуты / наименование страховой организации");
			ToImport.Columns.Add(@"Данные о долгах / Дата выхожа на просрочку", typeof(DateTime));
			ToImport.Columns.Add(@"Данные о долгах / Статус");
			ToImport.Columns.Add(@"Данные о долгах / Стадия долга");
			ToImport.Columns.Add(@"Произвольные атрибуты / Размещение");
			ToImport.Columns.Add(@"Данные о долгах / Регион");
			ToImport.Columns.Add(@"Данные о долгах / тип продукта");
			ToImport.Columns.Add(@"Произвольные атрибуты / Количество поручителей");
			ToImport.Columns.Add(@"Произвольные атрибуты / Количество поручителей родственников");
			ToImport.Columns.Add(@"Поручители должника / ФИО поручителя");
			ToImport.Columns.Add(@"Поручители должника / паспортные данные");
			ToImport.Columns.Add(@"Данные о долгах / Портфель");
			#endregion

			DataSetTables.Tables.Add(ToImport);

			var spisok = DataSetTables.Tables[@"Список"].Rows.Cast<System.Data.DataRow>().ToList();
			var dose = DataSetTables.Tables[@"Досье"].Rows.Cast<System.Data.DataRow>().ToList();
			var newDebts = DataSetTables.Tables[@"Новые долги"].Rows.Cast<System.Data.DataRow>().ToList();

			var toFile = (
				from l1 in spisok
				join d1 in dose on l1[@"Номер договора"] equals d1[@"Номер кредитного договора*"]
				join nd in newDebts on d1[@"Номер кредитного договора*"] equals nd[@"contract"]
				select new { spisok = l1, dose = d1, contract = l1[@"Номер договора"] }).ToList();

			toFile.ForEach(x => {

				var newRow = ToImport.NewRow();
				ToImport.Rows.Add(newRow);
				newRow[@"Данные о персонах / Фамилия"] = x.dose[@"Фамилия"];
				newRow[@"Данные о персонах / Имя"] = x.dose[@"Имя"];
				newRow[@"Данные о персонах / Отчество"] = x.dose[@"Отчество"];
				newRow[@"Данные о персонах / Пол"] = x.dose[@"Пол"];
				newRow[@"Данные о персонах / Дата рождения"] = StringToDatetime(x.dose[@"Дата рождения"].ToString());
				newRow[@"Произвольные атрибуты / MIDAS"] = x.dose[@"Сustomer"];
				newRow[@"Пасп Данные / Серия и Пасп Данные / Номер (через пробел)"] = PaspDannie1(x.dose[@"Номер паспорта"].ToString());
				newRow[@"Адреса фактические / почтовый индекс"] = x.dose[@"Индекс"];
				newRow[@"Адреса фактические / Район"] = x.dose[@"Область (субъект РФ)"];
				newRow[@"Адреса фактические / Город"] = x.dose[@"Город (иной населенный пункт)"];
				newRow[@"Адреса фактические / Улица"] = x.dose[@"Улица"];
				newRow[@"Адреса фактические / Дом"] = x.dose[@"Дом"];
				newRow[@"Адреса фактические / Корпус"] = x.dose[@"Корпус дома"];
				newRow[@"Адреса фактические / Квартира"] = x.dose[@"Квартира"];
				newRow[@"Адреса регистрации / Почтовый индекс"] = x.dose[@"Индекс"];
				newRow[@"Адреса регистрации / Район"] = x.dose[@"Область (субъект РФ)"];
				newRow[@"Адреса регистрации / Город"] = x.dose[@"Город (иной населенный пункт)"];
				newRow[@"Адреса регистрации / Улица"] = x.dose[@"Улица"];
				newRow[@"Адреса регистрации / Дом"] = x.dose[@"Дом"];
				newRow[@"Адреса регистрации / Корпус"] = x.dose[@"Корпус дома"];
				newRow[@"Адреса регистрации / Квартира"] = x.dose[@"Квартира"];
				newRow[@"Произвольные атрибуты / Идентификатор залога"] = x.dose[@"Идентификатор залога"];
				newRow[@"Произвольные атрибуты / Адрес заложенной недвижимости"] = x.dose[@"Адрес объекта недвижимости (Индекс, Область (субъект РФ)*, Город (иной населенный пункт)*, Улица*,Дом*, Корпус дома*, Квартира*)"];
				newRow[@"Произвольные атрибуты / Тип заложенной недвижимости"] = x.dose[@"Тип недвижимости"];
				newRow[@"Произвольные атрибуты / Статус залоговой недвижимости"] = x.dose[@"Статус залоговой недвижимости (идет строительство/ построено, но не сдано/ жилое помещение/ др.)"];
				newRow[@"Произвольные атрибуты / Оценочная стоимость заложенной недвижимости"] = x.dose[@"Оценочная стоимость"];
				newRow[@"Произвольные атрибуты / Оценочная стоимость валюта"] = x.dose[@"Оценочная стоимость Валюта"];
				newRow[@"Произвольные атрибуты / Дата оценки заложенной недвижимости"] = StringToDatetime( x.dose[@"Дата оценки"]);
				newRow[@"Произвольные атрибуты / Дата приобритения недвижииости"] = StringToDatetime(x.dose[@"Дата приобретения"]);
				newRow[@"Произвольные атрибуты / Тип владения"] = x.dose[@"Тип владения"];
				newRow[@"Произвольные атрибуты / Общая площадь заложенной недвижимости, м2"] = x.dose[@"Общая площадь, м2"];
				newRow[@"Произвольные атрибуты / Кадастровый номер заложенной недвижимости"] = x.dose[@"Кадастровый номер"];
				newRow[@"Произвольные атрибуты / Марка автомобиля"] = x.dose[@"Марка"];
				newRow[@"Произвольные атрибуты / Модель автомобиля"] = x.dose[@"Модель"];
				newRow[@"Произвольные атрибуты / VIN заложенного автомобиля"] = x.dose[@"VIN / Идентификационный номер"];
				newRow[@"Произвольные атрибуты / Оценочная стоимость залогового автомобиля"] = x.dose[@"Оценочная стоимость Значение"];
				newRow[@"Произвольные атрибуты / Оценочная стоимость залогового автомобиля валюта"] = x.dose[@"Оценочная стоимость Валюта"];
				newRow[@"Произвольные атрибуты / Дата оценки залогового автомобиля"] = StringToDatetime(x.dose[@"Дата оценки"]);
				newRow[@"Произвольные атрибуты / Год выпуска автомобиля"] = x.dose[@"Год выпуска"];
				newRow[@"Данные о персонах / Доход"] = x.dose[@"Доход семьи (за вычетом налоговых платежей, алиментов и др.) в месяц Значение"];
				newRow[@"Произвольные атрибуты / Доход семьи Валюта"] = x.dose[@"Доход семьи (за вычетом налоговых платежей, алиментов и др.) в месяц Значение (Валюта)"];
				newRow[@"Данные о персонах / Семейное положение"] = x.dose[@"Семейное положение*"];
				newRow[@"Произвольные атрибуты / Количество детей"] = x.dose[@"Количество детей до 18 лет"];
				newRow[@"Данные о персонах / Место работы"] = x.dose[@"Наименование работодателя*"];
				newRow[@"Адреса (тексты) рабочие / Текст адреса"] = x.dose[@"Адрес работы*"];
				newRow[@"Произвольные атрибуты / Вид деятельности"] = x.dose[@"Вид занятости*"];
				newRow[@"Данные о персонах / Должность"] = x.dose[@"Должность*"];
				newRow[@"Произвольные атрибуты / Период работы по последнему месту работы"] = x.dose[@"Период работы по последнему месту работы (с ..по)"];
				newRow[@"Телефоны домашние / Номе р телефона (через запятую три)"] = x.dose[@"Домашний телефон *"];
				newRow[@"Телефоны мобильные / Номе р телефона (через запятую три)"] = x.dose[@"Телефон мобильный*"];
				newRow[@"Телефоны рабочие / Номе р телефона (через запятую три)"] = x.dose[@"Телефон рабочий*"];
				newRow[@"Произвольные атрибуты / Регион должника"] = x.dose[@"Регион должника*"];
				newRow[@"Данные о долгах / Номер договора с банком"] = x.dose[@"Номер кредитного договора*"];
				newRow[@"Данные о долгах / Дата выдачи кредита"] = StringToDatetime(x.dose[@"Дата заключения кредитного договора*"]);
				newRow[@"Произвольные атрибуты / Дата окончания кредитного договора"] = StringToDatetime(x.dose[@"Дата окончания кредитного договора*"]);
				newRow[@"Данные о долгах / Название продукта"] = x.dose[@"Вид кредита*"];
				newRow[@"Данные о долгах / Валюта"] = x.dose[@"Валюта кредита*"];
				newRow[@"Данные о долгах / Лицевой счет должника"] = x.dose[@"Номер счета (ЦБ)"];
				newRow[@"Данные о долгах / Полный размер кредита"] = x.dose[@"Сумма кредита (кредитного лимита)"];
				newRow[@"Данные о долгах / Процентная ставка по кредиту"] = x.dose[@"Процентная ставка"];
				newRow[@"Данные о долгах / Аннуитентный платеж"] = x.dose[@"Сумма аннуитета*"];
				newRow[@"Данные о долгах / Сумма, необходимая к погашению"] = x.spisok[@"Итого"];
				newRow[@"Данные о долгах / Основной долг"] = x.dose[@"Текущий долг (сумма основного долга)*"];
				newRow[@"Данные о долгах / Просроченный основной долг"] = x.dose[@"Просроченный долг (сумма просроченного основного долга)*"];
				newRow[@"Данные о долгах / Проценты"] = x.dose[@"Текущие проценты*"];
				newRow[@"Данные о долгах / Просроченные проценты"] = x.dose[@"Просроченные проценты*"];
				newRow[@"Произвольные атрибуты / Перерасход по лимиту кредитной карты"] = x.dose[@"Перерасход по лимиту (для кредитных карт)*"];
				newRow[@"Данные о долгах / Пени"] = x.dose[@"Рассчитанные на дату передачи пени (штрафы)*"];
				newRow[@"Произвольные атрибуты / Процент на просрочку по основному долгу"] = x.dose[@"Процент на просрочку по основному долгу*"];
				newRow[@"Произвольные атрибуты / Штраф за перерасход по лимиту (кредитные карты)"] = x.dose[@"Штраф за перерасход по лимиту (кредитные карты)"];
				newRow[@"Произвольные атрибуты / Дата последнего платежа"] = StringToDatetime(x.dose[@"Дата последнего платежа*"]);
				newRow[@"Произвольные атрибуты / Сумма последнего платежа"] = x.dose[@"Сумма последнего платежа*"];
				newRow[@"Произвольные атрибуты / Количество произведенных платежей в счет погашения задолженности"] = x.dose[@"Количество произведенных платежей в счет погашения задолженности*"];
				newRow[@"Произвольные атрибуты / Сумма произведенных платежей в счет погашения"] = x.dose[@"Сумма произведенных платежей в счет погашения задолженности, руб.*"];
				newRow[@"Произвольные атрибуты / Количество дней от последнего платежа"] = x.dose[@"Количество дней от последнего платежа"];
				newRow[@"Произвольные атрибуты / Количество просроченных дней по кредиту"] = x.dose[@"Количество просроченыых дней по кредиту*"];
				newRow[@"Произвольные атрибуты / Наличие других кредитов в банке"] = x.dose[@"Наличие других кредитов в Банке*"];
				newRow[@"Произвольные атрибуты / Период добровольного страхования с"] = x.dose[@"Период добровольного страхования с"];
				newRow[@"Произвольные атрибуты / Период добровольного страхования по"] = x.dose[@"Период добровольного страхования по"];
				newRow[@"Произвольные атрибуты / Страховая премия по договору страхования"] = x.dose[@"Страховая премия по договору добровольного страхования"];
				newRow[@"Произвольные атрибуты / Страховая премия валюта"] = x.dose[@"Страховая премия. Валюта"];
				newRow[@"Произвольные атрибуты / наименование страховой организации"] = x.dose[@"Наименование страховой организации"];
				newRow[@"Данные о долгах / Дата выхожа на просрочку"] = DataVihodaNaProsrochku(Convert.ToInt32(x.spisok[@"Количество дней просрочки"]));
				newRow[@"Данные о долгах / Статус"] = Status(x.spisok[@"Отзыв перс. данных"].ToString(), x.spisok[@"Банкротство"].ToString());
				newRow[@"Данные о долгах / Стадия долга"] = x.spisok[@"Стадия"];
				newRow[@"Произвольные атрибуты / Размещение"] = x.spisok[@"Размещение"];
				newRow[@"Данные о долгах / Регион"] = x.spisok[@"Регион"];
				newRow[@"Данные о долгах / тип продукта"] = x.spisok[@"Параметризация"];
				newRow[@"Произвольные атрибуты / Количество поручителей"] = x.dose[@"Количество поручителей*"];
				newRow[@"Произвольные атрибуты / Количество поручителей родственников"] = x.dose[@"Количество поручителей-родственников"];
				newRow[@"Поручители должника / ФИО поручителя"] = ConcatFIO(
					x.dose[@"Фамилия48"].ToString(),
					x.dose[@"Имя49"].ToString(),
					x.dose[@"Отчество50"].ToString());
				newRow[@"Поручители должника / паспортные данные"] = x.dose[@"Паспортные данные"];

				newRow[@"Данные о долгах / Портфель"] = Portfolioname(newRow[@"Данные о долгах / Статус"].ToString());

			});

		}

		/// <summary>
		/// доп досье
		/// </summary>
		public void CreateToImportAdditioanal() {

			var ToImport = new System.Data.DataTable();
			ToImport.TableName = @"На импорт новые договоры Доп. досье";
			DataSetTables.Tables.Add(ToImport);


			var spisok = DataSetTables.Tables[@"Список"].Rows.Cast<System.Data.DataRow>().ToList();
			var dopDose = DataSetTables.Tables[@"Доп. досье"].Rows.Cast<System.Data.DataRow>().ToList();
			var newDebts = DataSetTables.Tables[@"Новые долги"].Rows.Cast<System.Data.DataRow>().ToList();

			var toFile = (
				from l1 in spisok
				join d1 in dopDose on l1[@"Номер договора"] equals d1[@"Номер договора"]
				join nd in newDebts on d1[@"Номер договора"] equals nd[@"contract"]
				select new { spisok = l1, dopDose = d1, contract = l1[@"Номер договора"] }).ToList();

			#region Create table

			ToImport.Columns.Add(@"contract");
			ToImport.Columns.Add(@"Customer");
			ToImport.Columns.Add(@"Регион");
			ToImport.Columns.Add(@"ФИО");
			ToImport.Columns.Add(@"Пол");
			ToImport.Columns.Add(@"Дата рождения", typeof(DateTime));
			ToImport.Columns.Add(@"Семейное положение");
			ToImport.Columns.Add(@"Номер паспорта");
			ToImport.Columns.Add(@"Кем выдан");
			ToImport.Columns.Add(@"Когда выдан", typeof(DateTime));
			ToImport.Columns.Add(@"Место работы");
			ToImport.Columns.Add(@"Должность");
			ToImport.Columns.Add(@"Стадия");
			ToImport.Columns.Add(@"Дней просрочки");
			ToImport.Columns.Add(@"Дата выхода на просрочку", typeof(DateTime));
			ToImport.Columns.Add(@"Сумма просроченной задолж. (RUR)");
			ToImport.Columns.Add(@"Задолженность (RUR)");
			ToImport.Columns.Add(@"Итого");
			ToImport.Columns.Add(@"Дата заключения договора", typeof(DateTime));
			ToImport.Columns.Add(@"День платежа");
			ToImport.Columns.Add(@"Размещение");
			ToImport.Columns.Add(@"Дата передачи на стадию", typeof(DateTime));
			ToImport.Columns.Add(@"Дата передачи в КА", typeof(DateTime));
			ToImport.Columns.Add(@"Статус");
			ToImport.Columns.Add(@"Вид кредита");
			ToImport.Columns.Add(@"Параметризация");
			ToImport.Columns.Add(@"Мобильный телефон");
			ToImport.Columns.Add(@"Домашний телефон");
			ToImport.Columns.Add(@"Рабочий телефон (1)");
			ToImport.Columns.Add(@"Рабочий телефон (2)");
			ToImport.Columns.Add(@"Адрес регистрации");
			ToImport.Columns.Add(@"Фактический адрес");
			ToImport.Columns.Add(@"Портфель");
			#endregion

			toFile.ForEach(x => {

				var newRow = ToImport.NewRow();
				ToImport.Rows.Add(newRow);

				newRow[@"contract"] = x.dopDose[@"Номер договора"];
				newRow[@"Customer"] = x.dopDose[@"Customer"];
				newRow[@"Регион"] = x.dopDose[@"Регион"];
				newRow[@"ФИО"] = x.dopDose[@"ФИО"];
				newRow[@"Пол"] = x.dopDose[@"Пол"];
				newRow[@"Дата рождения"] = StringToDatetime(x.dopDose[@"Дата рождения"]);
				newRow[@"Семейное положение"] = x.dopDose[@"Семейное положение"];
				newRow[@"Номер паспорта"] = PaspDannie1(x.dopDose[@"Номер паспорта"].ToString());
				newRow[@"Кем выдан"] = x.dopDose[@"Кем выдан"];
				newRow[@"Когда выдан"] = StringToDatetime(x.dopDose[@"Когда выдан"]);
				newRow[@"Место работы"] = x.dopDose[@"Место работы"];
				newRow[@"Должность"] = x.dopDose[@"Должность"];
				newRow[@"Стадия"] = x.spisok[@"Стадия"];
				newRow[@"Дней просрочки"] = x.spisok[@"Количество дней просрочки"];
				newRow[@"Дата выхода на просрочку"] = DataVihodaNaProsrochku(Convert.ToInt32(x.spisok[@"Количество дней просрочки"]));
				newRow[@"Сумма просроченной задолж. (RUR)"] = x.spisok[@"Общая сумма просроченной задолженности (RUR)"];
				newRow[@"Задолженность (RUR)"] = x.spisok[@"Основной долг (RUR)"];
				newRow[@"Итого"] = x.spisok[@"Итого"];
				newRow[@"Дата заключения договора"] = StringToDatetime(x.spisok[@"Дата заключения договора"]);
				newRow[@"День платежа"] = x.spisok[@"День платежа"];
				newRow[@"Размещение"] = x.spisok[@"Размещение"];
				newRow[@"Дата передачи на стадию"] = StringToDatetime( x.spisok[@"Дата передачи на стадию"]);
				newRow[@"Дата передачи в КА"] = StringToDatetime( x.spisok[@"Дата передачи в КА"]);
				newRow[@"Статус"] = Status(x.spisok[@"Отзыв перс. данных"].ToString(), x.spisok[@"Банкротство"].ToString());
				newRow[@"Вид кредита"] = x.spisok[@"Вид кредита"];
				newRow[@"Параметризация"] = x.spisok[@"Параметризация"];
				newRow[@"Мобильный телефон"] = x.dopDose[@"Мобильный телефон"];
				newRow[@"Домашний телефон"] = x.dopDose[@"Домашний телефон"];
				newRow[@"Рабочий телефон (1)"] = x.dopDose[@"Рабочий телефон (1)"];
				newRow[@"Рабочий телефон (2)"] = x.dopDose[@"Рабочий телефон (2)"];
				newRow[@"Адрес регистрации"] = x.dopDose[@"Адрес регистрации"];
				newRow[@"Фактический адрес"] = x.dopDose[@"Фактический адрес"];
				newRow[@"Портфель"] = Portfolioname(newRow[@"Статус"].ToString()); ;


			});


		}

		/// <summary>
		/// обнов.доп.досье
		/// </summary>
		public void CreateToImportAdditioanalUpdate() {
			var ToImport = new System.Data.DataTable();
			ToImport.TableName = @"обнов.доп.досье";
			DataSetTables.Tables.Add(ToImport);
			var spisok = DataSetTables.Tables[@"Список"].Rows.Cast<System.Data.DataRow>().ToList();
			var dopDose = DataSetTables.Tables[@"Доп. досье"].Rows.Cast<System.Data.DataRow>().ToList();
			var oldDebts = DataSetTables.Tables[@"Уже загружены в БД"].Rows.Cast<System.Data.DataRow>().ToList();

			var toFile = (
				from spisok1 in spisok
				join dopDose1 in dopDose on spisok1[@"Номер договора"] equals dopDose1[@"Номер договора"]
				join oldDebts1 in oldDebts on dopDose1[@"Номер договора"] equals oldDebts1[@"contract"]
				select new {
					spisok = spisok1,
					dopDose = dopDose1,
					oldDebts = oldDebts1,

				}).ToList();

			#region  create columns

			ToImport.Columns.Add(@"debt_id");
			ToImport.Columns.Add(@"Номер договора");
			ToImport.Columns.Add(@"Customer");
			ToImport.Columns.Add(@"Регион");
			ToImport.Columns.Add(@"ФИО");
			ToImport.Columns.Add(@"Пол");
			ToImport.Columns.Add(@"Дата рождения", typeof(DateTime));
			ToImport.Columns.Add(@"Семейное положение");
			ToImport.Columns.Add(@"Номер паспорта");
			ToImport.Columns.Add(@"Кем выдан");
			ToImport.Columns.Add(@"Когда выдан");
			ToImport.Columns.Add(@"Место работы");
			ToImport.Columns.Add(@"Должность");
			ToImport.Columns.Add(@"Стадия");
			ToImport.Columns.Add(@"Дней просрочки");
			ToImport.Columns.Add(@"Дата выхода на просрочку", typeof(DateTime));
			ToImport.Columns.Add(@"Сумма просроченной задолж. (RUR)");
			ToImport.Columns.Add(@"Задолженность (RUR)");
			ToImport.Columns.Add(@"Итого");
			ToImport.Columns.Add(@"Дата заключения договора", typeof(DateTime));
			ToImport.Columns.Add(@"День платежа");
			ToImport.Columns.Add(@"Размещение");
			ToImport.Columns.Add(@"Дата передачи на стадию", typeof(DateTime));
			ToImport.Columns.Add(@"Дата передачи в КА", typeof(DateTime));
			ToImport.Columns.Add(@"Вид кредита");
			ToImport.Columns.Add(@"Параметризация");
			ToImport.Columns.Add(@"Мобильный телефон");
			ToImport.Columns.Add(@"Домашний телефон");
			ToImport.Columns.Add(@"Рабочий телефон (1)");
			ToImport.Columns.Add(@"Рабочий телефон (2)");
			ToImport.Columns.Add(@"Адрес регистрации");
			ToImport.Columns.Add(@"Фактический адрес");
			ToImport.Columns.Add(@"Портфель");
			ToImport.Columns.Add(@"Статус");

			#endregion

			#region Map columns

			toFile.ForEach(x => {

				var newRow = ToImport.NewRow();
				ToImport.Rows.Add(newRow);


				newRow[@"debt_id"] = x.oldDebts[@"debt_id"];
				newRow[@"Номер договора"] = x.dopDose[@"Номер договора"];
				newRow[@"Customer"] = x.dopDose[@"Customer"];
				newRow[@"Регион"] = x.dopDose[@"Регион"];
				newRow[@"ФИО"] = x.dopDose[@"ФИО"];
				newRow[@"Пол"] = x.dopDose[@"Пол"];
				newRow[@"Дата рождения"] = StringToDatetime(x.dopDose[@"Дата рождения"]);
				newRow[@"Семейное положение"] = x.dopDose[@"Семейное положение"];
				newRow[@"Номер паспорта"] = PaspDannie1(x.dopDose[@"Номер паспорта"].ToString());
				newRow[@"Кем выдан"] = x.dopDose[@"Кем выдан"];
				newRow[@"Когда выдан"] = x.dopDose[@"Когда выдан"];
				newRow[@"Место работы"] = x.dopDose[@"Место работы"];
				newRow[@"Должность"] = x.dopDose[@"Должность"];
				newRow[@"Стадия"] = x.spisok[@"Стадия"];
				newRow[@"Дней просрочки"] = x.spisok[@"Количество дней просрочки"];
				newRow[@"Дата выхода на просрочку"] = DataVihodaNaProsrochku(Convert.ToInt32(x.spisok[@"Количество дней просрочки"]));
				newRow[@"Сумма просроченной задолж. (RUR)"] = x.spisok[@"Общая сумма просроченной задолженности (RUR)"];
				newRow[@"Задолженность (RUR)"] = x.spisok[@"Основной долг (RUR)"];
				newRow[@"Итого"] = x.spisok[@"Итого"];
				newRow[@"Дата заключения договора"] = StringToDatetime(x.spisok[@"Дата заключения договора"]);
				newRow[@"День платежа"] = x.spisok[@"День платежа"];
				newRow[@"Размещение"] = x.spisok[@"Размещение"];
				newRow[@"Дата передачи на стадию"] = StringToDatetime(x.spisok[@"Дата передачи на стадию"]);
				newRow[@"Дата передачи в КА"] = StringToDatetime(x.spisok[@"Дата передачи в КА"]);
				newRow[@"Вид кредита"] = x.spisok[@"Вид кредита"];
				newRow[@"Параметризация"] = x.spisok[@"Параметризация"];
				newRow[@"Мобильный телефон"] = x.dopDose[@"Мобильный телефон"];
				newRow[@"Домашний телефон"] = x.dopDose[@"Домашний телефон"];
				newRow[@"Рабочий телефон (1)"] = x.dopDose[@"Рабочий телефон (1)"];
				newRow[@"Рабочий телефон (2)"] = x.dopDose[@"Рабочий телефон (2)"];
				newRow[@"Адрес регистрации"] = x.dopDose[@"Адрес регистрации"];
				newRow[@"Фактический адрес"] = x.dopDose[@"Фактический адрес"];

				newRow[@"Статус"] = Status(x.spisok[@"Отзыв перс. данных"].ToString(), x.spisok[@"Банкротство"].ToString());
				newRow[@"Портфель"] = Portfolioname(newRow[@"Статус"].ToString());




			});


			#endregion


		}

		/// <summary>
		/// обновление досье
		/// </summary>
		public void CreateToImportList1Update() {

			var ToImport = new System.Data.DataTable();
			ToImport.TableName = @"обнов.досье";
			DataSetTables.Tables.Add(ToImport);
			var spisok = DataSetTables.Tables[@"Список"].Rows.Cast<System.Data.DataRow>().ToList();
			var dose = DataSetTables.Tables[@"Досье"].Rows.Cast<System.Data.DataRow>().ToList();
			var oldDebts = DataSetTables.Tables[@"Уже загружены в БД"].Rows.Cast<System.Data.DataRow>().ToList();

			var toFile = (
				from spisok1 in spisok
				join dose1 in dose on spisok1[@"Номер договора"] equals dose1[@"Номер кредитного договора*"]
				join oldDebts1 in oldDebts on dose1[@"Номер кредитного договора*"] equals oldDebts1[@"contract"]
				select new {
					spisok = spisok1,
					dose = dose1,
					oldDebts = oldDebts1,

				}).ToList();

			#region  create columns

			ToImport.Columns.Add(@"debt_id");
			ToImport.Columns.Add(@"Фамилия");
			ToImport.Columns.Add(@"Имя");
			ToImport.Columns.Add(@"Отчество");
			ToImport.Columns.Add(@"Пол");
			ToImport.Columns.Add(@"Дата рождения", typeof(DateTime));
			ToImport.Columns.Add(@"Customer/midas");
			ToImport.Columns.Add(@"Паспорт");
			ToImport.Columns.Add(@"Залог атрибут");
			ToImport.Columns.Add(@"Адрес заложенной недвижимости");
			ToImport.Columns.Add(@"Тип заложенной недвижимости");
			ToImport.Columns.Add(@"Статус залоговой недвижимости");
			ToImport.Columns.Add(@"Оценочная стоимость заложенной недвижимости");
			ToImport.Columns.Add(@"Оценочная стоимость Валюта");
			ToImport.Columns.Add(@"Дата оценки заложенной недвижимости", typeof(DateTime));
			ToImport.Columns.Add(@"Дата приобритения", typeof(DateTime));
			ToImport.Columns.Add(@"Тип владения");
			ToImport.Columns.Add(@"Общая площадь заложенной недвижимости, м2");
			ToImport.Columns.Add(@"Кадастровый номер заложенной недвижимости");
			ToImport.Columns.Add(@"Марка авто атрибут");
			ToImport.Columns.Add(@"Модель авто");
			ToImport.Columns.Add(@"VIN заложенного автомобиля");
			ToImport.Columns.Add(@"Оценочная стоимость залогового автомобиля");
			ToImport.Columns.Add(@"Оценочная стоимость автомобиля валюта");
			ToImport.Columns.Add(@"Дата оценки заложенного автомобиля", typeof(DateTime));
			ToImport.Columns.Add(@"Год выпуска авто");
			ToImport.Columns.Add(@"Доход семьи в месяц");
			ToImport.Columns.Add(@"Доход Семьи Валюта");
			ToImport.Columns.Add(@"Семейное положение");
			ToImport.Columns.Add(@"Количество детей");
			ToImport.Columns.Add(@"Место работы");
			ToImport.Columns.Add(@"Вид занятости");
			ToImport.Columns.Add(@"Должность");
			ToImport.Columns.Add(@"Период работы по последнему месту работы");
			ToImport.Columns.Add(@"Регион атрибут (Досье)");
			ToImport.Columns.Add(@"Номер договора");
			ToImport.Columns.Add(@"Дата договора", typeof(DateTime));
			ToImport.Columns.Add(@"Дата окончания кредитного договора", typeof(DateTime));
			ToImport.Columns.Add(@"Вид кредита - название продукта");
			ToImport.Columns.Add(@"Валюта кредита");
			ToImport.Columns.Add(@"Номер счета");
			ToImport.Columns.Add(@"Сумма кредита");
			ToImport.Columns.Add(@"Процентная ставка");
			ToImport.Columns.Add(@"Сумма аннуитета");
			ToImport.Columns.Add(@"Итого остаток (список)");
			ToImport.Columns.Add(@"Текущий долг (сумма основного долга)");
			ToImport.Columns.Add(@"Просроченный долг");
			ToImport.Columns.Add(@"Проценты");
			ToImport.Columns.Add(@"Просроченные проценты");
			ToImport.Columns.Add(@"Перерасход по лимиту кредитной карты");
			ToImport.Columns.Add(@"Пени");
			ToImport.Columns.Add(@"Процент на просрочку по основному долгу*");
			ToImport.Columns.Add(@"Штраф за перерасход по лимиту (кредитные карты)");
			ToImport.Columns.Add(@"Дата последнего платежа", typeof(DateTime));
			ToImport.Columns.Add(@"Сумма последнего платежа");
			ToImport.Columns.Add(@"Количество произведенных платежей в счет погашения задолженности");
			ToImport.Columns.Add(@"Сумма произведенных платежей в счет погашения задолженности");
			ToImport.Columns.Add(@"Количество дней от последнего платежа");
			ToImport.Columns.Add(@"Количество просроченыых дней по кредиту*");
			ToImport.Columns.Add(@"Наличие других кредитов в банке");
			ToImport.Columns.Add(@"Период добровольного страхования с");
			ToImport.Columns.Add(@"Период добровольного страхования по");
			ToImport.Columns.Add(@"Страховая премия по договору страхования");
			ToImport.Columns.Add(@"Страховая премия. Валюта");
			ToImport.Columns.Add(@"Наименование страховой организации");
			ToImport.Columns.Add(@"Дата выхода на просрочку", typeof(DateTime));
			ToImport.Columns.Add(@"Статус долга Банкротство/Отзыв ПД");
			ToImport.Columns.Add(@"Стадия долга (список) словарь 101");
			ToImport.Columns.Add(@"Размещение (список)");
			ToImport.Columns.Add(@"Регион долга (лист Список)");
			ToImport.Columns.Add(@"Параметризация (Список Тип продукта)");
			ToImport.Columns.Add(@"Количество поручителей");
			ToImport.Columns.Add(@"Количество поручителей-родственников");


			#endregion


			toFile.ForEach(x => {

				var newRow = ToImport.NewRow();
				ToImport.Rows.Add(newRow);

				newRow[@"debt_id"] = x.oldDebts[@"debt_id"];
				newRow[@"Фамилия"] = x.dose[@"Фамилия"];
				newRow[@"Имя"] = x.dose[@"Имя"];
				newRow[@"Отчество"] = x.dose[@"Отчество"];
				newRow[@"Пол"] = x.dose[@"Пол"];
				newRow[@"Дата рождения"] = StringToDatetime(x.dose[@"Дата рождения"]);
				newRow[@"Customer/midas"] = x.dose[@"Сustomer"];
				newRow[@"Паспорт"] = x.dose[@"Номер паспорта"];
				newRow[@"Залог атрибут"] = x.dose[@"Идентификатор залога"];
				newRow[@"Адрес заложенной недвижимости"] = x.dose[@"Адрес объекта недвижимости (Индекс, Область (субъект РФ)*, Город (иной населенный пункт)*, Улица*,Дом*, Корпус дома*, Квартира*)"];
				newRow[@"Тип заложенной недвижимости"] = x.dose[@"Тип недвижимости"];
				newRow[@"Статус залоговой недвижимости"] = x.dose[@"Статус залоговой недвижимости (идет строительство/ построено, но не сдано/ жилое помещение/ др.)"];
				newRow[@"Оценочная стоимость заложенной недвижимости"] = x.dose[@"Оценочная стоимость"];
				newRow[@"Оценочная стоимость Валюта"] = x.dose[@"Оценочная стоимость Валюта"];
				newRow[@"Дата оценки заложенной недвижимости"] = StringToDatetime(x.dose[@"Дата оценки"]);
				newRow[@"Дата приобритения"] = StringToDatetime(x.dose[@"Дата приобретения"]);
				newRow[@"Тип владения"] = x.dose[@"Тип владения"];
				newRow[@"Общая площадь заложенной недвижимости, м2"] = x.dose[@"Общая площадь, м2"];
				newRow[@"Кадастровый номер заложенной недвижимости"] = x.dose[@"Кадастровый номер"];
				newRow[@"Марка авто атрибут"] = x.dose[@"Марка"];
				newRow[@"Модель авто"] = x.dose[@"Модель"];
				newRow[@"VIN заложенного автомобиля"] = x.dose[@"VIN / Идентификационный номер"];
				newRow[@"Оценочная стоимость залогового автомобиля"] = x.dose[@"Оценочная стоимость Значение"];
				newRow[@"Оценочная стоимость автомобиля валюта"] = x.dose[@"Оценочная стоимость Валюта"];
				newRow[@"Дата оценки заложенного автомобиля"] = StringToDatetime(x.dose[@"Дата оценки"]);
				newRow[@"Год выпуска авто"] = x.dose[@"Год выпуска"];
				newRow[@"Доход семьи в месяц"] = x.dose[@"Доход семьи (за вычетом налоговых платежей, алиментов и др.) в месяц Значение"];
				newRow[@"Доход Семьи Валюта"] = x.dose[@"Доход семьи (за вычетом налоговых платежей, алиментов и др.) в месяц Значение (Валюта)"];
				newRow[@"Семейное положение"] = x.dose[@"Семейное положение*"];
				newRow[@"Количество детей"] = x.dose[@"Количество детей до 18 лет"];
				newRow[@"Место работы"] = x.dose[@"Наименование работодателя*"];
				newRow[@"Вид занятости"] = x.dose[@"Вид занятости*"];
				newRow[@"Должность"] = x.dose[@"Должность*"];
				newRow[@"Период работы по последнему месту работы"] = x.dose[@"Период работы по последнему месту работы (с ..по)"];
				newRow[@"Регион атрибут (Досье)"] = x.dose[@"Регион должника*"];
				newRow[@"Номер договора"] = x.dose[@"Номер кредитного договора*"];
				newRow[@"Дата договора"] = StringToDatetime(x.dose[@"Дата заключения кредитного договора*"]);
				newRow[@"Дата окончания кредитного договора"] = StringToDatetime(x.dose[@"Дата окончания кредитного договора*"]);
				newRow[@"Вид кредита - название продукта"] = x.dose[@"Вид кредита*"];
				newRow[@"Валюта кредита"] = x.dose[@"Валюта кредита*"];
				newRow[@"Номер счета"] = x.dose[@"Номер счета (ЦБ)"];
				newRow[@"Сумма кредита"] = x.dose[@"Сумма кредита (кредитного лимита)"];
				newRow[@"Процентная ставка"] = x.dose[@"Процентная ставка"];
				newRow[@"Сумма аннуитета"] = x.dose[@"Сумма аннуитета*"];
				newRow[@"Итого остаток (список)"] = x.spisok[@"Итого"];
				newRow[@"Текущий долг (сумма основного долга)"] = x.dose[@"Текущий долг (сумма основного долга)*"];
				newRow[@"Просроченный долг"] = x.dose[@"Просроченный долг (сумма просроченного основного долга)*"];
				newRow[@"Проценты"] = x.dose[@"Текущие проценты*"];
				newRow[@"Просроченные проценты"] = x.dose[@"Просроченные проценты*"];
				newRow[@"Перерасход по лимиту кредитной карты"] = x.dose[@"Перерасход по лимиту (для кредитных карт)*"];
				newRow[@"Пени"] = x.dose[@"Рассчитанные на дату передачи пени (штрафы)*"];
				newRow[@"Процент на просрочку по основному долгу*"] = x.dose[@"Процент на просрочку по основному долгу*"];
				newRow[@"Штраф за перерасход по лимиту (кредитные карты)"] = x.dose[@"Штраф за перерасход по лимиту (кредитные карты)"];
				newRow[@"Дата последнего платежа"] = StringToDatetime(x.dose[@"Дата последнего платежа*"]);
				newRow[@"Сумма последнего платежа"] = x.dose[@"Сумма последнего платежа*"];
				newRow[@"Количество произведенных платежей в счет погашения задолженности"] = x.dose[@"Количество произведенных платежей в счет погашения задолженности*"];
				newRow[@"Сумма произведенных платежей в счет погашения задолженности"] = x.dose[@"Сумма произведенных платежей в счет погашения задолженности, руб.*"];
				newRow[@"Количество дней от последнего платежа"] = x.dose[@"Количество дней от последнего платежа"];
				newRow[@"Количество просроченыых дней по кредиту*"] = x.dose[@"Количество просроченыых дней по кредиту*"];
				newRow[@"Наличие других кредитов в банке"] = x.dose[@"Наличие других кредитов в Банке*"];
				newRow[@"Период добровольного страхования с"] = x.dose[@"Период добровольного страхования с"];
				newRow[@"Период добровольного страхования по"] = x.dose[@"Период добровольного страхования по"];
				newRow[@"Страховая премия по договору страхования"] = x.dose[@"Страховая премия по договору добровольного страхования"];
				newRow[@"Страховая премия. Валюта"] = x.dose[@"Страховая премия. Валюта"];
				newRow[@"Наименование страховой организации"] = x.dose[@"Наименование страховой организации"];
				newRow[@"Дата выхода на просрочку"] = DataVihodaNaProsrochku(Convert.ToInt32(x.spisok[@"Количество дней просрочки"]));
				newRow[@"Статус долга Банкротство/Отзыв ПД"] = Status(x.spisok[@"Отзыв перс. данных"].ToString(), x.spisok[@"Банкротство"].ToString());
				newRow[@"Стадия долга (список) словарь 101"] = x.spisok[@"Стадия"];
				newRow[@"Размещение (список)"] = x.spisok[@"Размещение"];
				newRow[@"Регион долга (лист Список)"] = x.spisok[@"Регион"];
				newRow[@"Параметризация (Список Тип продукта)"] = x.spisok[@"Параметризация"];
				newRow[@"Количество поручителей"] = x.dose[@"Количество поручителей*"];
				newRow[@"Количество поручителей-родственников"] = x.dose[@"Количество поручителей-родственников"];

			});

		}

		#endregion

	}

	/// <summary>
	/// АЗ
	/// </summary>
	public class TableProcessor2 : TableProcessor {

		public TableFormat2 tableFormat2;//Досье
		public TableFormat5 tableFormat5;//АЗ
		
		public TableProcessor2():base(@"АЗ") {
			GetInputFile();

		}

		public override void CreateExcelFileFormat() {
			excelFileFormat = new ExcelFileFormat2();
		}

		public override Koanvi.Excel.Application ExportToExcel() {
			var excel = base.ExportToExcel();

			return excel;
		}

		#region DataProcessing
		public override void Process() {
			base.Process();

		}

		/// <summary>
		/// делаем файл загрузки для листа "Досье 2ые кредиты"
		/// </summary>
		public void CreateToToImport1() {

			var ToImport = new Koanvi.Data.Tables.DataTable();
			ToImport.TableName = @"На импорт новые договоры досье";

		}

		/// <summary>
		/// подготовка загрузки обновление остатков
		/// </summary>
		public void CreateToImport2() {

			var ToImport = new Koanvi.Data.Tables.DataTable();
			ToImport.TableName = @"На импорт новые договоры досье";

		}
		#endregion


	}//public class TableProcessor2 : TableProcessor

	/// <summary>
	/// Список
	/// </summary>
	public class TableFormat1 : ExcelTableFormat, IExcelTableFormat {
		public TableFormat1() : base() {

			this.Name = @"Формат листа список";
			this.SheetName = @"Список";
			this.HeaderRow = 0;
			this.ColumnNames =
				new List<string>() {
			 @"Номер договора"
			,@"Customer"
			,@"Валюта"
			,@"Вид кредита"
			,@"Параметризация"
			,@"Тип продукта"
			,@"ФИО"
			,@"Стадия"
			,@"КА"
			,@"Регион"
			,@"Количество дней просрочки"
			,@"Бакет просрочки"
			,@"Основной долг (RUR)"
			,@"Текущие проценты (RUR)"
			,@"Просроченный основной долг (RUR)"
			,@"Просроченные проценты на основной долг (RUR)"
			,@"Штрафы (RUR)"
			,@"Пени (RUR)"
			,@"Перерасход (RUR)"
			,@"Неустойка (кр.карты) / Проценты на просроченную задолженность (RUR)"
			,@"Общая сумма просроченной задолженности (RUR)"
			,@"Итого"
			,@"Дата заключения договора"
			,@"День платежа"
			,@"Размещение"
			,@"Дата передачи на стадию"
			,@"Дата передачи в КА"
			,@"Нет просрочки"
			,@"Отзыв перс. данных"
			,@"Банкротство"
			,@"Статус реализации авто"
			,@"Кол-во контактных звонков за тек.неделю"
			,@"Кол-во контактных звонков за тек.месяц"
			,@"Кол-во отправленных SMS за тек.неделю"
			,@"Кол-во отправленных SMS за тек.месяц"
			,@"Кол-во отправленных писем эл.почты за тек.неделю"
			,@"Кол-во отправленных писем эл.почты за тек.месяц"
			,@"Тип ограничения на взаимодействие"
			,@"Дата начала действия ограничения"
			,@"Дата окончания действия ограничения"
			,@"Представитель"

		};


		}
	}

	/// <summary>
	/// Досье
	/// </summary>
	public class TableFormat2 : ExcelTableFormat, IExcelTableFormat {
		public TableFormat2() : base() {

			this.Name = @"Формат листа досье";
			this.SheetName = @"Досье";
			this.HeaderRow = 1;
			this.ColumnNames =
				new List<string>() {
			 @"КА"
			,@"Фамилия"
			,@"Имя"
			,@"Отчество"
			,@"Пол"
			,@"Дата рождения"
			,@"Сustomer"
			,@"Серия паспорта"
			,@"Номер паспорта"
			,@"Индекс"
			,@"Область (субъект РФ)"
			,@"Город (иной населенный пункт)"
			,@"Улица"
			,@"Дом"
			,@"Корпус дома"
			,@"Квартира"
			,@"Индекс"
			,@"Область (субъект РФ)"
			,@"Город (иной населенный пункт)"
			,@"Улица"
			,@"Дом"
			,@"Корпус дома"
			,@"Квартира"
			,@"Идентификатор залога"
			,@"Адрес объекта недвижимости (Индекс, Область (субъект РФ)*, Город (иной населенный пункт)*, Улица*,Дом*, Корпус дома*, Квартира*)"
			,@"Тип недвижимости"
			,@"Статус залоговой недвижимости (идет строительство/ построено, но не сдано/ жилое помещение/ др.)"
			,@"Оценочная стоимость"
			,@"Оценочная стоимость Валюта"
			,@"Дата оценки"
			,@"Дата приобретения"
			,@"Тип владения"
			,@"Общая площадь, м2"
			,@"Кадастровый номер"
			,@"Марка"
			,@"Модель"
			,@"VIN / Идентификационный номер"
			,@"Оценочная стоимость Значение"
			,@"Оценочная стоимость Валюта"
			,@"Дата оценки"
			,@"Год выпуска"
			,@"Доход семьи (за вычетом налоговых платежей, алиментов и др.) в месяц Значение"
			,@"Доход семьи (за вычетом налоговых платежей, алиментов и др.) в месяц Значение (Валюта)"
			,@"Семейное положение*"
			,@"Количество детей до 18 лет"
			,@"Количество поручителей*"
			,@"Количество поручителей-родственников"
			,@"Фамилия"
			,@"Имя"
			,@"Отчество"
			,@"Паспортные данные"
			,@"Наименование работодателя*"
			,@"Адрес работы*"
			,@"Вид занятости*"
			,@"Должность*"
			,@"Период работы по последнему месту работы (с ..по)"
			,@"Домашний телефон *"
			,@"Телефон мобильный*"
			,@"Телефон рабочий*"
			,@"Регион должника*"
			,@"Номер кредитного договора*"
			,@"Дата заключения кредитного договора*"
			,@"Дата окончания кредитного договора*"
			,@"Вид кредита*"
			,@"Валюта кредита*"
			,@"Номер счета (ЦБ)"
			,@"Сумма кредита (кредитного лимита)"
			,@"Процентная ставка"
			,@"Сумма аннуитета*"
			,@"Текущий долг (сумма основного долга)*"
			,@"Просроченный долг (сумма просроченного основного долга)*"
			,@"Текущие проценты*"
			,@"Просроченные проценты*"
			,@"Просроченные комиссии*"
			,@"Перерасход по лимиту (для кредитных карт)*"
			,@"Рассчитанные на дату передачи пени (штрафы)*"
			,@"Процент на просрочку по основному долгу*"
			,@"Штраф за перерасход по лимиту (кредитные карты)"
			,@"Дата последнего платежа*"
			,@"Сумма последнего платежа*"
			,@"Количество произведенных платежей в счет погашения задолженности*"
			,@"Сумма произведенных платежей в счет погашения задолженности, руб.*"
			,@"Количество дней от последнего платежа"
			,@"Количество просроченыых дней по кредиту*"
			,@"Наличие других кредитов в Банке*"
			,@"Отметка об участие третьих лиц (коллекторских агентств) по взысканию суммы долга по распоряжению Цедента (поле в Siebel Коллекторское агентство:)*"
			,@"должник умер*"
			,@"Дата возбуждения уголовного производства (в Siebel поле: Дата подачи заявления о возбуждении уголовного дела:)"
			,@"Дата судебного приказа (поле в Siebel Дата подачи заявления о выдаче судебного приказа:)"
			,@"Дата исполнительного производства ( поле в Siebel Дата начала исполнительного производства:)"
			,@"Уплаченная Цедентом государственная пошлина, Руб."
			,@"Дата оплаты гос.пошлины"
			,@"Дата возмещения гос.пошлины"
			,@"Госпошлина погашена (Y/N)"
			,@"Стадия работы с задолженностью:"
			,@" Дата передачи в коллекторское агентство:"
			,@"Период добровольного страхования с"
			,@"Период добровольного страхования по"
			,@"Страховая премия по договору добровольного страхования"
			,@"Страховая премия. Валюта"
			,@"Наименование страховой организации"
			,@"Дата формирования отчета"
			,@"Мигрированный кредит"
			,@"Тип фасилити"
			,@"Сиквенс фасилити"

		};


		}
	}

	/// <summary>
	/// Досье 2ые кредиты
	/// </summary>
	public class TableFormat6 : TableFormat2, IExcelTableFormat {
		public TableFormat6():base() {
			this.Name = @"Досье 2ые кредиты";
			this.SheetName = @"Досье 2ые кредиты";
		}

	}

	/// <summary>
	/// Доп. досье
	/// </summary>
	public class TableFormat3 : ExcelTableFormat, IExcelTableFormat {
		public TableFormat3() : base() {

			this.Name = @"Формат листа Доп. досье";
			this.SheetName = @"Доп. досье";
			this.HeaderRow = 0;
			this.ColumnNames =
				new List<string>() {
			 @"КА"
			,@"Номер договора"
			,@"Customer"
			,@"Регион"
			,@"ФИО"
			,@"Пол"
			,@"Дата рождения"
			,@"Семейное положение"
			,@"Номер паспорта"
			,@"Кем выдан"
			,@"Когда выдан"
			,@"Место работы"
			,@"Должность"
			,@"Мобильный телефон"
			,@"Домашний телефон"
			,@"Рабочий телефон (1)"
			,@"Рабочий телефон (2)"
			,@"Адрес регистрации"
			,@"Фактический адрес"


		};

		}
	}

	/// <summary>
	/// Изъятие
	/// </summary>
	public class TableFormat4 : ExcelTableFormat, IExcelTableFormat {
		public TableFormat4() : base() {

			this.Name = @"Формат листа Изъятие";
			this.SheetName = @"Изъятие";
			this.HeaderRow = 0;
			this.ColumnNames =
				new List<string>() {
			 @"Номер договора"
			,@"Customer"
			,@"Регион"
			,@"Изъятие"

		};

		}
	}

	/// <summary>
	/// АЗ
	/// </summary>
	public class TableFormat5: ExcelTableFormat, IExcelTableFormat {
		public TableFormat5() : base() {

			this.Name= @"Формат листа АЗ";
			this.SheetName = @"АЗ";
			this.HeaderRow = 0;
			this.ColumnNames =
				new List<string>() {
			 @"Номер договора"
			,@"Customer"
			,@"Валюта"
			,@"ФИО"
			,@"Тип кредита"
			,@"Дата заключения договора"
			,@"Регион"
			,@"Дата образования просрочки"
			,@"КА"
			,@"Дата передачи клиента КА"
			,@"DPD"
			,@"Бакет просрочки"
			,@"Стадия"
			,@"Основной долг"
			,@"Проценты на основной долг"
			,@"Просроченный основной долг"
			,@"Проcроченные проценты на основной долг"
			,@"Штрафы"
			,@"Пени"
			,@"Перерасход"
			,@"Неустойка (кр.карты) / Проценты на просроченную задолженность"
			,@"Общая сумма просроченной задолженности"
			,@"Общая сумма задолженности"
			,@"Отзыв перс. данных"
			,@"Банкротство"
			,@"Стадия реализации авто"
			,@"Тип ограничения на взаимодействие"
			,@"Дата начала действия ограничения"
			,@"Дата окончания действия ограничения"
			,@"Представитель"
			,@"Кол-во контактных звонков за тек.неделю"
			,@"Кол-во контактных звонков за тек.месяц"
			,@"Кол-во SMS за тек.неделю"
			,@"Кол-во SMS за тек.месяц"
			,@"Кол-во писем за тек.неделю"
			,@"Кол-во писем за тек.месяц"


		};

		}
	}

	/// <summary>
	/// файл передача
	/// </summary>
	public class ExcelFileFormat1 : ExcelFileFormat {
		public ExcelFileFormat1() : base() {
			this.Name = @"файл передача";
			this.Sheets = new List<IExcelTableFormat>();
			this.Sheets.Add(new TableFormat1());//Список
			this.Sheets.Add(new TableFormat2());//Досье
			this.Sheets.Add(new TableFormat3());//Доп. досье
			this.Sheets.Add(new TableFormat4());//Изъятие

		}
	}

	/// <summary>
	/// АЗ
	/// </summary>
	public class ExcelFileFormat2 : ExcelFileFormat {
		public ExcelFileFormat2() : base() {
			this.Name = @"АЗ";
			this.Sheets = new List<IExcelTableFormat>() { };
			this.Sheets.Add(new TableFormat6());
			this.Sheets.Add(new TableFormat5());
		}
	}

}

/// <summary>
/// данные из БД
/// </summary>
namespace Koanvi.Data.Tables.Contact {
	public class DataTableAllDept : Koanvi.Data.Tables.DataTable {
		System.Data.SqlClient.SqlConnection Conn;
		public DataTableAllDept(System.Data.SqlClient.SqlConnection conn) {
			Conn = conn;
		}
		public void Fill() {
			System.Data.SqlClient.SqlCommand cmd = new System.Data.SqlClient.SqlCommand(TableParser.Properties.Resources.AllDebt, Conn);
			var da = new System.Data.SqlClient.SqlDataAdapter(cmd);
			da.Fill(this);


		}

	}
}

/// <summary>
/// операции для процессора
/// </summary>
namespace Koanvi.TableProcessor.Operation {}

/// <summary>
/// для описания формата
/// </summary>
namespace Koanvi.TableProcessor.Format {

	/// <summary>
	/// для проверки
	/// </summary>
	public class Format {
		public string Name { get; set; }
		public Format(string name) { this.Name = name; }
		public Format() { }
		public virtual bool Check() { return true; }
	}

	public class FileFormat : Format {
		public String FileName;

		public FileFormat(string name, String fileName) : base(name) { this.FileName = fileName; }
		public FileFormat() { }
		public override bool Check() {
			return base.Check();

		}

	}

	public interface IExcelFileFormat {
		List<IExcelTableFormat> Sheets { get; set; }
		bool CheckSheets();
		bool Check();
		Koanvi.Excel.Application Excel { get; set; }
	}

	public class ExcelFileFormat : FileFormat, IExcelFileFormat {

		public List<IExcelTableFormat> Sheets { get; set; }
		public Koanvi.Excel.Application Excel { get; set; }

		public ExcelFileFormat(string name, String fileName, List<IExcelTableFormat> sheets, Koanvi.Excel.Application Excel) : base(name, fileName) {
			this.Sheets = sheets;
			this.Excel = Excel;
		}
		public ExcelFileFormat() { }
		
		/// <summary>
		/// пока проверяет только названия листов
		/// </summary>
		/// <returns></returns>
		public override bool Check() {
			if (!base.Check()) { return false; }
			if (!CheckSheets()) { return false; }

			return true;
		}
		
		/// <summary>
		/// проверяет только названия листов
		/// </summary>
		/// <returns></returns>
		public bool CheckSheets() {

			var hasSheets = Excel.Sheets;
			if (hasSheets == null) { return false; }
			var retval = Sheets.Select(x=>x.SheetName).Intersect(hasSheets).Count() == Sheets.Count();

			//Нет в списке:
			var notInList = Sheets.Select(x => x.SheetName).Where(p1 => !hasSheets.Any(p2 => p1 == p2)).ToList();

			try {
				if (retval) { return retval; }
				throw new Exception(@"Внимание! в выбранном фале не обнаружены листы:" + Environment.NewLine +
				string.Join(@";" + Environment.NewLine, notInList)
				+ @".");

			}
			catch (Exception e) { System.Windows.Forms.MessageBox.Show(e.Message); throw; }

		}
		public void Fill(Koanvi.Excel.Application excel) {
			this.Excel = excel;
			this.Sheets.ForEach(sheet => {
				sheet.Fill(excel);
			});
		}

	}

	/// <summary>
	/// проверяет столбцы
	/// </summary>
	public class TableFormat : Format {

		public List<string> ColumnNames { get; set; }

		public Koanvi.Data.Tables.DataTable Table { get; set; }

		public TableFormat() { }

		public override bool Check() {

			var colNames = Table.Columns.Cast<System.Data.DataColumn>().Select(x => x.ColumnName).ToList();

			var except = ColumnNames.Except(colNames).ToList();

			try {
				if (except.Count > 0) {
					throw new Exception($@"На листе {Table.TableName.ToString()} не найдены столбцы:" + Environment.NewLine +
						string.Join(@";" + Environment.NewLine, except)
						);
					//return except.Count > 0;

				}
			}
			catch (Exception ex) {

				throw;
			}

			return except.Count > 0;
		}

	}

	public interface IExcelTableFormat {
		string SheetName { get; set; }
		int HeaderRow { get; set; }
		void SetHeader();
		bool Check();
		void Fill(Koanvi.Excel.Application excel);
		List<string> ColumnNames { get; set; }
		Koanvi.Data.Tables.DataTable Table { get; set; }
	}
	public class ExcelTableFormat : TableFormat, IExcelTableFormat {
		public string SheetName { get; set; }
		public int HeaderRow { get; set; }

		public ExcelTableFormat() { }
		public void SetHeader() {
			this.Table.SetHeader(HeaderRow);

		}
		public override bool Check() {
			var retval= base.Check();
			return base.Check()&& Table.TableName== SheetName;
		}
		public void Fill(Koanvi.Excel.Application excel) {
			this.Table=	new Koanvi.Data.Tables.DataTable( excel.GetRange(this.SheetName));
			this.Table.SetHeader(this.HeaderRow);
			this.Table.Columns.Cast<System.Data.DataColumn>().ToList().ForEach(x => {
				x.ColumnName = x.ColumnName.Replace(@"""", @"");
				x.ColumnName = x.ColumnName.Replace(Environment.NewLine, @" ");
				x.ColumnName = x.ColumnName.Replace("\n", @" ");
			});

		}
	}

}

namespace Koanvi.Data.Tables {

	public class DataTable : System.Data.DataTable {

		public DataTable() { }
		public DataTable(System.Data.DataTable dt ) {

			this.Rows.Clear();
			this.Columns.Clear();

			dt.Columns.Cast<System.Data.DataColumn>().ToList().ForEach(oc => {
				var nc = new System.Data.DataColumn();
				nc.DataType = oc.DataType;
				nc.ColumnName = oc.ColumnName;
				this.Columns.Add(nc);
			});
			dt.Rows.Cast<System.Data.DataRow>().ToList().ForEach(or => {
				var nr = this.NewRow();
				nr.ItemArray = or.ItemArray;
				this.Rows.Add(nr);
			});

			this.TableName = dt.TableName;

		}

		public void SetHeader(int rowNum) {

			var row = this.Rows[rowNum];
			SetHeader(row);

		}
		public void SetHeader(System.Data.DataRow row) {

			// группировка для названий столбцов - попозже доделать:
			var newRowData = row.ItemArray.ToList();
			var group = (
				from data in newRowData
				group data by data into grp
				where grp.Count() > 1
				select new { colName = grp.Key, count = grp.Count(), grp = grp.ToList() }).ToList();

			int i = 0;
			this.Columns.Cast<System.Data.DataColumn>().ToList().ForEach(col => {
				i++;
				if (row[col.ColumnName].ToString() == string.Empty) { return; }
				try {
					col.ColumnName = row[col.ColumnName].ToString();
				}
				catch (System.Data.DuplicateNameException ex) {
					col.ColumnName = row[col.ColumnName].ToString() + i.ToString();
				}
			});
			row.Delete();

		}

		/// <summary>
		/// Меняет тип данных в столбце
		/// </summary>
		/// <param name="colName">название колонки</param>
		/// <param name="type">тип куда менять</param>
		/// <returns></returns>
		public DataTable ChangeDataType(string colName, Type type) {

			var iCol = this.Columns.IndexOf(colName);
			var tmp = new List<object[]>();
			this.Rows.Cast<System.Data.DataRow>().ToList().ForEach(dr => {
				tmp.Add(dr.ItemArray);
			});
			this.Rows.Clear();
			this.Columns[iCol].DataType = typeof(DateTime);

			tmp.ForEach(item => {
				item[iCol] = Convert.ChangeType(item[iCol], type);
				var nr = this.NewRow();
				this.Rows.Add(nr);
				nr.ItemArray = item;
			});
			return (DataTable)this;

		}//public void ChangeDataType(System.Data.DataTable dataTable,string colName, Type type)

		/// <summary>
		/// Копирует значения столбцов в таблицу 1 из таблицы 2
		/// </summary>
		/// <param name="copyFrom">откуда копировать</param>
		/// <param name="joinInfo">информация для джоина</param>
		public void JoinDataTables(System.Data.DataTable copyFrom, JoinInfo joinInfo) {
			var joinData = (
				from rTo in this.Rows.Cast<System.Data.DataRow>().ToList()
				join rFrom in copyFrom.Rows.Cast<System.Data.DataRow>().ToList()
				on rTo[joinInfo.Id.ColNameTo] equals rFrom[joinInfo.Id.ColNameFrom]
				select new { copyTo = rTo, copyFrom = rFrom }
				).ToList();
			joinInfo.Columns.ForEach(colMapping => {

				joinData.ForEach(tables => {
					tables.copyTo[colMapping.ColNameTo] = tables.copyFrom[colMapping.ColNameFrom];
				});


			});

		}

	}

	/// <summary>
	/// информация для объединения таблиц
	/// </summary>
	public class JoinInfo {
		public Mapping Id;
		public List<Mapping> Columns;
		public JoinInfo(Mapping Id, List<Mapping> Columns) {
			this.Id = Id; this.Columns = Columns;

		}
	}

	public class Mapping {
		public string ColNameFrom;
		public string ColNameTo;
		public Mapping(string ColNameFrom, string ColNameTo) { this.ColNameFrom = ColNameFrom; this.ColNameTo = ColNameTo; }
	}

}