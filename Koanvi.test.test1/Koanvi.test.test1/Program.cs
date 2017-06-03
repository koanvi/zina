using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Koanvi.test.test1 {
  static class Program {
    /// <summary>
    /// Главная точка входа для приложения.
    /// </summary>
    [STAThread]
    static void Main() {
      Application.EnableVisualStyles();
      Application.SetCompatibleTextRenderingDefault(false);
      StartApp();
    }
    public static void StartApp() {
      //C:\Users\koanvi\AppData\Local\Microsoft\Microsoft SQL Server Local DB\Instances\MSSQLLocalDB
      var list=Enumerable.Range(1, 12).Select(x =>
      @"https://www.google.ru/search?q=GetResponse&oq=GetResponse&aqs=chrome..69i57.813j0j4&sourceid=chrome&ie=UTF-8#newwindow=1&q=c%23+"+x.ToString()
      ).ToList();
      var asd = new Koanvi.Projects.HtmlHoover.HtmlHoover(list);
      asd.Fill();


    }
  }
}
