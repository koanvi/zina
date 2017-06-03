using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Koanvi.Controls.v1 {
  public partial class test_custom_control : System.Windows.Forms.Control {
    public test_custom_control() {
      InitializeComponent();
    }

   protected override void OnPaint(PaintEventArgs pe) {
      base.OnPaint(pe);
    }
  }
}
