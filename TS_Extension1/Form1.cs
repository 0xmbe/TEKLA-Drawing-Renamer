using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using Tekla.Structures;
using Tekla.Structures.Drawing;
using Tekla.Structures.Model;
namespace TS_Extension1
{
    public class Variables
    {
        public static string caption = "Drawing Renamer v0.2";
        public static string date = "Tekla v21.0";
        public static string title = Variables.caption + " / " + Variables.date;
    }

    public class Form1 : Form
	{
		private IContainer components;
		private Button SelectedDrawing;
		private Button Close_tool;
		private ProgressBar progressBar1;
		private Label CurrentNo;
		private Label label1;
		private Label label2;
		private Model My_model = new Model();
		private DateTime ExpiryDate = new DateTime(2013, 8, 31);
		protected override void Dispose(bool disposing)
		{
			if (disposing && this.components != null)
			{
				this.components.Dispose();
			}
			base.Dispose(disposing);
		}
		private void InitializeComponent()
		{
			ComponentResourceManager componentResourceManager = new ComponentResourceManager(typeof(Form1));
			this.SelectedDrawing = new Button();
			this.Close_tool = new Button();
			this.progressBar1 = new ProgressBar();
			this.CurrentNo = new Label();
			this.label1 = new Label();
			this.label2 = new Label();
			base.SuspendLayout();
			this.SelectedDrawing.Location = new Point(91, 72);
			this.SelectedDrawing.Name = "SelectedDrawing";
			this.SelectedDrawing.Size = new System.Drawing.Size(80, 23);
			this.SelectedDrawing.TabIndex = 1;
			this.SelectedDrawing.Text = "Selected";
			this.SelectedDrawing.UseVisualStyleBackColor = true;
			this.SelectedDrawing.Click += new EventHandler(this.SelectedDrawing_Click);
			this.Close_tool.Location = new Point(172, 72);
			this.Close_tool.Name = "Close_tool";
			this.Close_tool.Size = new System.Drawing.Size(75, 23);
			this.Close_tool.TabIndex = 2;
			this.Close_tool.Text = "Close";
			this.Close_tool.UseVisualStyleBackColor = true;
			this.Close_tool.Click += new EventHandler(this.Close_tool_Click);
			this.progressBar1.Location = new Point(10, 12);
			this.progressBar1.Name = "progressBar1";
			this.progressBar1.Size = new System.Drawing.Size(190, 23);
			this.progressBar1.TabIndex = 3;
			this.CurrentNo.Anchor = (AnchorStyles.Top | AnchorStyles.Right);
			this.CurrentNo.AutoSize = true;
			this.CurrentNo.Location = new Point(203, 13);
			this.CurrentNo.Name = "CurrentNo";
			this.CurrentNo.Size = new System.Drawing.Size(16, 17);
			this.CurrentNo.TabIndex = 5;
			this.CurrentNo.Text = "  ";
			this.label1.AutoSize = true;
			this.label1.Location = new Point(8, 44);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(234, 17);
			this.label1.TabIndex = 6;
			this.label1.Text = "Rename Drawing Names for:";
			this.label2.AutoSize = true;
			this.label2.Location = new Point(8, 100);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(234, 17);
			this.label2.TabIndex = 6;
			this.label2.Text = "------------------------------------------------------------\nMod by Matija Bensa / Skanding\n Based on FL-PL Checker by Domen Zagar";
			base.AutoScaleDimensions = new SizeF(8f, 16f);
			base.AutoScaleMode = AutoScaleMode.Font;
			base.ClientSize = new System.Drawing.Size(420, 133);
			base.Controls.Add(this.label1);
			base.Controls.Add(this.label2);
			base.Controls.Add(this.CurrentNo);
			base.Controls.Add(this.progressBar1);
			base.Controls.Add(this.Close_tool);
			base.Controls.Add(this.SelectedDrawing);
			base.Icon = (Icon)componentResourceManager.GetObject("$this.Icon");
			base.Name = "Form1";
            this.Text = Variables.title;
			base.TopMost = true;
			base.Load += new EventHandler(this.Form1_Load);
			base.ResumeLayout(false);
			base.PerformLayout();
		}
		public Form1()
		{
			this.InitializeComponent();
			if (!this.My_model.GetConnectionStatus())
			{
				MessageBox.Show("Tekla not started");
				Application.Exit();
			}
			Application.Exit();
		}
		private void Form1_Load(object sender, EventArgs e)
		{
			this.progressBar1.Value = 0;
		}
		private void SelectedDrawing_Click(object sender, EventArgs e)
		{
			this.progressBar1.Value = 0;
			DrawingHandler drawingHandler = new DrawingHandler();
			if (drawingHandler.GetConnectionStatus())
			{
				DrawingEnumerator selected = drawingHandler.GetDrawingSelector().GetSelected();
				this.RenameDrawingTitle(selected);
			}
		}
		private void RenameDrawingTitle(DrawingEnumerator DrawingList)
		{
			this.progressBar1.Maximum = DrawingList.GetSize();
            int num = 1;
            int num2 = 0;
            int num3 = 0;
			int num4 = 0;
            bool needsUpdating = false;
			while (DrawingList.MoveNext())
			{
				this.progressBar1.Value++;
				this.CurrentNo.Text = num++.ToString() + '/' + DrawingList.GetSize().ToString();
				this.CurrentNo.Refresh();

                string mainpartname = "";
                string existingDrawingname = "";    // Name of the drawing before modify

				Tekla.Structures.Model.ModelObject modelObject = null;

                Drawing currentDrawing = DrawingList.Current;
                if (currentDrawing.UpToDateStatus.ToString() != "DrawingIsUpToDate")
                {
                    needsUpdating = true;
                    continue;
                }

				if (DrawingList.Current is AssemblyDrawing)
				{
					AssemblyDrawing assemblyDrawing = DrawingList.Current as AssemblyDrawing;
                    Identifier assemblyIdentifier = assemblyDrawing.AssemblyIdentifier;
					modelObject = this.My_model.SelectModelObject(assemblyIdentifier);

                    modelObject.GetReportProperty("ASSEMBLY_NAME", ref mainpartname);

                    num2++;
				}

                if (DrawingList.Current is SinglePartDrawing)
				{
                    SinglePartDrawing singlePartDrawing = DrawingList.Current as SinglePartDrawing;
					Identifier partIdentifier = singlePartDrawing.PartIdentifier;
					modelObject = this.My_model.SelectModelObject(partIdentifier);

                    modelObject.GetReportProperty("NAME", ref mainpartname);
          
				}
				if (modelObject != null)
				{

                    // Check if drawing name already contains the automatic drawing name:
                    existingDrawingname = DrawingList.Current.Name;
                    bool drawingNameMatch = existingDrawingname.Contains(mainpartname);

                    if ((drawingNameMatch == true) || (drawingNameMatch = existingDrawingname.Contains("DS")))
                    {
                        num4++;
                    }                    
                    if (drawingNameMatch == false)
                    {
                        DrawingList.Current.Name = mainpartname;
                        DrawingList.Current.Modify();
                        num3++;
                    }

                }
			}

            if (needsUpdating == true)
            {
                MessageBox.Show("Some of the drawings are not up to date!\n\nNames were not updated for that drawings.", Variables.title);
            }

            MessageBox.Show(string.Concat(new object[]
			{
				num3,
				" Drawing's name changed \n",		
                num4,
                " Drawings kept the existing name"
            }), Variables.title);
		}
		private void Close_tool_Click(object sender, EventArgs e)
		{
			Application.Exit();
		}
	}
}
