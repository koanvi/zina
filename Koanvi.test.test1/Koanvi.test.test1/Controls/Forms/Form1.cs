using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Koanvi.Controls.Forms.test {
  public partial class Form1 : System.Windows.Forms.Form {

    public Form1() {
      InitializeComponent();
      //this.Controls
    }
    //override createcontr
  }

  public class ControlCollection : System.Windows.Forms.Control.ControlCollection {
    public ControlCollection(System.Windows.Forms. Control owner) : base(owner) {
    }
  }
  
}
