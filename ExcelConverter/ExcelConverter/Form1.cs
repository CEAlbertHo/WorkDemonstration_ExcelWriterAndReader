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

// ExcelDataReader 和 ExcelDataReader.DataSet 用
using System.Data;
using ExcelDataReader;

namespace ExcelConverter
{
	public partial class Form1 : Form
	{
		// Define
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
			
			List<string> _successFilePath	= new List<string>();
			List<string> _failFilePath		= new List<string>();

			string[] _filesPath = Directory.GetFiles( GetSourceFolderPath(), "*.*", SearchOption.AllDirectories );

			for( int i=0; i < _filesPath.Length; i++ )
			{
				bool _result = TryConvertExcel( _filesPath[ i ] );
				if( _result )
					_successFilePath.Add( _filesPath[ i ] );
				else
					_failFilePath.Add( _filesPath[ i ] );
			}

			Label_Text.Text = string.Format( "ConvertExcel Done.\nSuccessed Num : {0}.\nFailed Num : {1}", _successFilePath.Count, _failFilePath.Count );

			
			// 輸出 Log
			string _logText = string.Empty;
			_logText += string.Format( "Successed File Count : {0}\n", _successFilePath.Count );

			for( int i=0; i < _successFilePath.Count; i++ )
			{
				_logText += string.Format( "- {0}\n", _successFilePath[ i ] );
			}

			_logText += string.Format( "\n\n---\n\n" );
			_logText += string.Format( "Failed File Count : {0}\n", _failFilePath.Count );
			for( int i=0; i < _failFilePath.Count; i++ )
			{
				_logText += string.Format( "- {0}\n", _failFilePath[ i ] );
			}

			_logText += string.Format( "\n\n---\n\n" );
			_logText += "ConvertExcel Done.";

			LogTextResult( _logText, GetConvertedFolderPath() );
		}

		#endregion

		#region ConvertExcel

		private bool TryConvertExcel( string iFilePath )
		{
			bool _extentionCheck = CheckExtention( iFilePath );
			if( !_extentionCheck )
				return false;

			// Do Convert
			bool _convertResult = ConvertExcelToBinary( iFilePath );
			if( !_convertResult )
				return false;

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
			try
			{
				FileStream _fileStream = File.Open( iFilePath, FileMode.Open, FileAccess.Read );
				IExcelDataReader excelDataReader = ExcelReaderFactory.CreateOpenXmlReader( _fileStream );

				DataSet _dataSet = excelDataReader.AsDataSet();
				DataRowCollection _dataRow = _dataSet.Tables[ 0 ].Rows;
				DataColumnCollection _dataColumn = _dataSet.Tables[ 0 ].Columns;



				return true;
			}
			catch ( Exception _exception )
			{
				Label_Text.Text		= "[Exception] Msg : " + _exception.Message;

				return false;
			}
		}

		#endregion

		#region Log

		public void LogTextResult( string iLog, string iLogFolderPath )
		{
			string _FileName = string.Format( "[{0}] Convert Excel Result.txt", DateTime.Now.ToString( "yyyy.MM.dd - HHmmss" ) );
			string _FilePath = Path.Combine( iLogFolderPath, _FileName );

			using( FileStream _fs = File.Create( _FilePath ) )
			{
				FileStream_AddText( _fs, iLog );
				_fs.Close();
			}
		}

		private void FileStream_AddText( FileStream iFileStream, string iText )
		{
			byte[] _info = new UTF8Encoding( true ).GetBytes( iText );
			iFileStream.Write( _info, 0, _info.Length );
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
