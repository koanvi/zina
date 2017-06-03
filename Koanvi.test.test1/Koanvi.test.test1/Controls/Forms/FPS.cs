using System;

namespace Koanvi.Controls.Forms {
  //using System.Windows.Forms;
  using System.Drawing;

  public class FPS:System.Windows.Forms.Form {

    public System.Windows.Forms.TextBox TextBox;
    public System.Windows.Forms.Timer timer1;

    public FPS() {

      Init();

      var asd = new Koanvi.Graphic.Common.v2.FPS(TextBox);
      TextBox.Text = asd.GetUpdateDely().ToString();
      //InitTimer();

    }

    public virtual void Init() {

      this.Height = 0;// tb.Height;
      this.Width = 20;
      BackColor = Color.Lime;
      TransparencyKey = Color.Lime;
      this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
      this.TopMost = true;

      this.TextBox = new System.Windows.Forms.TextBox();
      this.Controls.Add(this.TextBox);
      this.TextBox.Dock = System.Windows.Forms.DockStyle.Fill;
      //tb.BackColor = Color.Lime;


    }

    protected override void OnPaintBackground(System.Windows.Forms.PaintEventArgs e) {
      base.OnPaintBackground(e);
    }

    public void InitTimer() {
      timer1 = new System.Windows.Forms.Timer() ;
      timer1.Interval = 2000; // in miliseconds
      timer1.Tick += (object sender, EventArgs e) => 
      //{ this.TextBox.Text = Koanvi.Graphic.Common.v3.Utility.CalculateFrameRate().ToString();};
      timer1.Start();
    }

  }

  public class CustomTextBox : System.Windows.Forms.TextBox {
    public CustomTextBox() {
      
      SetStyle(System.Windows.Forms.ControlStyles.SupportsTransparentBackColor |
               System.Windows.Forms.ControlStyles.OptimizedDoubleBuffer |
               System.Windows.Forms.ControlStyles.AllPaintingInWmPaint |
               System.Windows.Forms.ControlStyles.ResizeRedraw |
               System.Windows.Forms.ControlStyles.UserPaint, true);
      BackColor = Color.Transparent;
    }
  }

}
namespace Koanvi.Graphic.Common.v2 {

  public class FPS {

    public System.Windows.Forms.Control Control;

    public FPS(System.Windows.Forms.Control Control) {
      this.Control = Control;
    }

    public virtual int GetUpdateDely() {
      var Tick1 = System.Environment.TickCount;
      this.Control.Text = @"";
      this.Control.Update();
      this.Control.Refresh();
      var Tick2 = System.Environment.TickCount;
      return Tick2 - Tick1;
    }

  }
}
/// <summary>
/// эта штука для юнити
/// пока не работает
/// </summary>
namespace Koanvi.Graphic.Common.v3.Unity {
  using System.Timers;
  //using UnityEngine;
  public static class FPS {

    public static double? CalculateFrameRate() {
      return null;
      var deltaTime = 0.0;
      var fps = 0.0;
      //deltaTime += Time.deltaTime;
      deltaTime /= 2.0;
      fps = 1.0 / deltaTime;
      return fps;
    }

  }
}

