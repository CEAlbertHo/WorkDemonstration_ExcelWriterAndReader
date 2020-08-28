using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelConverter
{
	public partial class Form1 : Form
	{
		const string SourceFolderName		= "Excel_SourceFolder";
		const string ConvertedFolderName	= "Excel_ConvertedFolder";
		static readonly string[] AcceptedExcelExtentionArray = new string[]{ ".xlsx" };

		public Form1()
		{
			InitializeComponent();
		}

		#region Button Event

		private void Btn_CreateDirectory_Click(object sender, EventArgs e)
		{
			string _sourceFolderPath	= GetSourceFolderPath();
			string _convertedFolderPath = GetConvertedFolderPath();
			
			bool _sourceFolderExist		= Directory.Exists( _sourceFolderPath );
			bool _convertedFolderExist	= Directory.Exists( _convertedFolderPath );

			if( !_sourceFolderExist )
				Directory.CreateDirectory( _sourceFolderPath );

			if( !_convertedFolderExist )
				Directory.CreateDirectory( _convertedFolderPath );

			// Re-Check
			_sourceFolderExist		= Directory.Exists( _sourceFolderPath );
			_convertedFolderExist	= Directory.Exists( _convertedFolderPath );

			// Output Result
			if( _sourceFolderExist && _convertedFolderExist )
			{
				Label_Text.Text = "CreateDirectory Successed.";
			}
			else
			{
				Label_Text.Text = "CreateDirectory Failed.";
			}
		}

		private void Btn_ConvertExcel_Click(object sender, EventArgs e)
		{
			Label_Text.Text		= "Running";
			
			int _successCount	= 0;
			int _failCount		= 0;

			string[] _filesPath = Directory.GetFiles( GetSourceFolderPath(), "*.*", SearchOption.AllDirectories );

			for( int i=0; i < _filesPath.Length; i++ )
			{
				bool _result = TryConvertExcel( _filesPath[ i ] );
				if( _result )
					_successCount++;
				else
					_failCount++;
			}

			Label_Text.Text = string.Format( "ConvertExcel Done.\nSuccessed Num : {0}.\nFailed Num : {1}", _successCount, _failCount );
		}

		#endregion

		#region ConvertExcel

		private bool TryConvertExcel( string iFilePath )
		{
			bool _extentionCheck = CheckExtention( iFilePath );
			if( !_extentionCheck )
				return false;

			// Do Convert
			ConvertExcelToBinary( iFilePath );

			return true;
		}

		private bool CheckExtention( string iFilePath )
		{
			string _extention = Path.GetExtension( iFilePath );
			
			if( _extention == string.Empty ||
				_extention == null )
			{
				return false;
			}

			// 清單中有 = OK
			for( int i=0; i < AcceptedExcelExtentionArray.Length; i++ )
			{
				bool _result = _extention.ToLower().Equals( AcceptedExcelExtentionArray[ i ].ToLower() );
				
				if( _result )
					return true;
			}

			return false;
		}

		private bool ConvertExcelToBinary( string iFilePath )
		{

		}

		#endregion

		#region Helper

		private string GetSourceFolderPath()
		{
			string _currentPath = Environment.CurrentDirectory;
			
			return Path.Combine( _currentPath, SourceFolderName );
		}

		private string GetConvertedFolderPath()
		{
			string _currentPath = Environment.CurrentDirectory;
			
			return Path.Combine( _currentPath, ConvertedFolderName );
		}

		#endregion
	}
}
