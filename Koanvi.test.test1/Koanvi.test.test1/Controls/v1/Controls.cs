using System;
using System.Windows.Forms;

namespace Koanvi.Controls {
  public class Control {

    protected System.Windows.Forms.Panel _MainControl;
    public virtual System.Windows.Forms.Panel MainControl { get { return _MainControl; } }
    public virtual void InitMainControl() {
      _MainControl = new Panel();
    }

    public System.Windows.Forms.DockStyle DockStyle { get; set; }

    protected ControlCollection _Controls;
    public virtual ControlCollection Controls { get; }
    public virtual void InitControls() {
      _Controls = new ControlCollection(MainControl);
    }

    public Control() {
      _MainControl = new System.Windows.Forms.Panel();
      Controls = new ControlCollection(MainControl);
    }

  }

  public class Form : Control {

    protected System.Windows.Forms.Form _MainForm;
    public System.Windows.Forms.Form MainForm {get { return _MainForm; }}

    public override ControlCollection Controls { get; }

    public Form() : base() {
      _MainForm = new System.Windows.Forms.Form();
      _MainForm.Controls.Add(MainControl);
      MainControl.Dock = System.Windows.Forms.DockStyle.Fill;

    }
    public virtual DialogResult ShowDialog() {
      return MainForm.ShowDialog();
    }

  }
  public class ControlCollection : System.Windows.Forms.Control.ControlCollection {
    public ControlCollection(System.Windows.Forms.Control owner) : base(owner) {

    }
    public override void Add(System.Windows.Forms.Control value) {
      base.Add(value);
    }
  }
  public class ControlCollection2 : ControlCollection {
    public ControlCollection2(System.Windows.Forms.Control owner) : base(owner) {

    }
    public override void Add(System.Windows.Forms.Control value) {
      base.Add(value);
    }

  }
  public class AppForm: Form {
    public AppForm() :base(){

    }

  }
}
