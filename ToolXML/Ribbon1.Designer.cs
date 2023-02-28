namespace ToolXML
{
	partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
	{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		public Ribbon1()
			: base(Globals.Factory.GetRibbonFactory())
		{
			InitializeComponent();
		}

		/// <summary> 
		/// Clean up any resources being used.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Component Designer generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.tab1 = this.Factory.CreateRibbonTab();
			this.group1 = this.Factory.CreateRibbonGroup();
			this.ReadSheet = this.Factory.CreateRibbonButton();
			this.SaveToXMl = this.Factory.CreateRibbonButton();
			this.ReadFromXml = this.Factory.CreateRibbonButton();
			this.tab1.SuspendLayout();
			this.group1.SuspendLayout();
			this.SuspendLayout();
			// 
			// tab1
			// 
			this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
			this.tab1.Groups.Add(this.group1);
			this.tab1.Label = "ToolXml";
			this.tab1.Name = "tab1";
			// 
			// group1
			// 
			this.group1.Items.Add(this.ReadSheet);
			this.group1.Items.Add(this.SaveToXMl);
			this.group1.Items.Add(this.ReadFromXml);
			this.group1.Label = "group1";
			this.group1.Name = "group1";
			// 
			// ReadSheet
			// 
			this.ReadSheet.Label = "s";
			this.ReadSheet.Name = "ReadSheet";
			this.ReadSheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
			// 
			// SaveToXMl
			// 
			this.SaveToXMl.Label = "SavaToXml";
			this.SaveToXMl.Name = "SaveToXMl";
			this.SaveToXMl.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SaveToXMl_Click);
			// 
			// ReadFromXml
			// 
			this.ReadFromXml.Label = "ReadFromXml";
			this.ReadFromXml.Name = "ReadFromXml";
			this.ReadFromXml.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ReadFromXml_Click);
			// 
			// Ribbon1
			// 
			this.Name = "Ribbon1";
			this.RibbonType = "Microsoft.Excel.Workbook";
			this.Tabs.Add(this.tab1);
			this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
			this.tab1.ResumeLayout(false);
			this.tab1.PerformLayout();
			this.group1.ResumeLayout(false);
			this.group1.PerformLayout();
			this.ResumeLayout(false);

		}

		#endregion

		internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
		internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton SaveToXMl;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton ReadFromXml;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton ReadSheet;
	}

	partial class ThisRibbonCollection
	{
		internal Ribbon1 Ribbon1
		{
			get { return this.GetRibbon<Ribbon1>(); }
		}
	}
}
