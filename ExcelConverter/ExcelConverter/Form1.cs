using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelConverter
{
	public partial class Form1 : Form
	{
		public Form1()
		{
			InitializeComponent();
		}		

		private void Btn_CreateDirectory_Click(object sender, EventArgs e)
		{
			string _currentPath = Environment.CurrentDirectory;

			Label_Text.Text = _currentPath;

			bool _folderExist = System.IO.Directory.Exists( _currentPath );
		}

		private void Btn_ConvertExcel_Click(object sender, EventArgs e)
		{
			Label_Text.Text = "Running";
		}
	}
}
