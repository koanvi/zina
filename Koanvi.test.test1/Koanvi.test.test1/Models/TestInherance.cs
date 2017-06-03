using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Koanvi.Models.Test {
  public class TestInherance {
    public TestInherance() {

      var ads = a.Create();
      c obj = new c();
      obj.do1();


    }
  }






  public class a {
    public a() { Init(); ; }
    private void Init() { }
    public virtual void do1() { }

    public static a Create() { return new a(); }

  }
  public class b : a {
    public b() { do1(); }
    public override void do1() { }
  }
  public class c : b {
    public c() { Init(); }
    private void Init() { }
    public override void do1() { base.do1(); }
  }
}
