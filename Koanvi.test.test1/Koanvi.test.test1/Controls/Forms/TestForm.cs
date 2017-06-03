using System;


namespace Koanvi.Controls.Forms.Test{
  public class TestForm:System.Windows.Forms.Form {
    private System.Windows.Forms.OpenFileDialog openFileDialog1;
    private System.Windows.Forms.ToolStripContainer toolStripContainer1;

    public TestForm():base() {


    }

    private void InitializeComponent() {
      this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
      this.toolStripContainer1 = new System.Windows.Forms.ToolStripContainer();
      this.toolStripContainer1.SuspendLayout();
      this.SuspendLayout();
      // 
      // openFileDialog1
      // 
      this.openFileDialog1.FileName = "openFileDialog1";
      // 
      // toolStripContainer1
      // 
      // 
      // toolStripContainer1.ContentPanel
      // 
      this.toolStripContainer1.ContentPanel.Size = new System.Drawing.Size(610, 236);
      this.toolStripContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
      this.toolStripContainer1.Location = new System.Drawing.Point(0, 0);
      this.toolStripContainer1.Name = "toolStripContainer1";
      this.toolStripContainer1.Size = new System.Drawing.Size(610, 261);
      this.toolStripContainer1.TabIndex = 0;
      this.toolStripContainer1.Text = "toolStripContainer1";
      // 
      // TestForm
      // 
      this.ClientSize = new System.Drawing.Size(610, 261);
      this.Controls.Add(this.toolStripContainer1);
      this.Name = "TestForm";
      this.toolStripContainer1.ResumeLayout(false);
      this.toolStripContainer1.PerformLayout();
      this.ResumeLayout(false);

    }
  }
}
namespace Koanvi.Controls.Forms.Test {
  using System.Drawing;
  using System.Drawing.Drawing2D;

  public class TestGrafient: System.Windows.Forms.Form {
    public TestGrafient() {

      this.SetStyle(System.Windows.Forms.ControlStyles.ResizeRedraw, true);
      //this.Controls.Add(new System.Windows.Forms.TransparentPanel());
      this.Controls.Add(new System.Windows.Forms.TextBox() { Text = @"asd" });

    }
    protected override void OnPaintBackground(System.Windows.Forms.PaintEventArgs e) {
      Rectangle rc = new Rectangle(0, 0, this.ClientSize.Width, this.ClientSize.Height);
      using(LinearGradientBrush brush = new LinearGradientBrush(rc, Color.LightBlue, Color.Blue, 45F)) {
        e.Graphics.FillRectangle(brush, rc);
      }
    }
  }

}
