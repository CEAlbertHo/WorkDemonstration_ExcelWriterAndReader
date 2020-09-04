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
using System.Security.AccessControl;

namespace ExcelConverter
{
	public enum EExcelContentType
	{
		None,
		Single,		// 單一表格
		Multiple,	// 多重表格
	}

	public partial class Form1 : Form
	{
		// Define
		const string SourceFolderName		= "Excel_SourceFolder";
		const string ConvertedFolderName	= "Excel_ConvertedFolder";
		static readonly string[] AcceptedExcelExtentionArray = new string[]{ ".xlsx" };

		const string ExcelContentType_Single	= "Single";
		const string ExcelContentType_Multiple	= "Multiple";

		const string EExcelReadSymbol = "#";
		const string EExcelReadSymbol_EndOfSheet = "#END";
		


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

		/// <summary>
		///  實際執行 Excel 轉換
		/// </summary>
		/// <param name="iFilePath"> Excel 檔案路徑 </param>
		/// <returns></returns>
		private bool ConvertExcelToBinary( string iFilePath )
		{
			try
			{
				// 讀取整張 Excel 資料
				#region 讀取整張 Excel 資料

				FileStream _readFileStream = File.Open( iFilePath, FileMode.Open, FileAccess.Read );
				IExcelDataReader excelDataReader = ExcelReaderFactory.CreateOpenXmlReader( _readFileStream );

				DataSet _dataSet = excelDataReader.AsDataSet();
				DataRowCollection _dataRow = _dataSet.Tables[ 0 ].Rows;
				DataColumnCollection _dataColumn = _dataSet.Tables[ 0 ].Columns;


				// 讀取表格類型
				EExcelContentType _excelContentType = GetExcelContentType( _dataColumn );
				if( _excelContentType == EExcelContentType.None )
					return false;

				switch( _excelContentType )
				{
					case EExcelContentType.Single:
						// 處理 單一表格
						break;

					case EExcelContentType.Multiple:
						// 處理多重表格
						break;
				}

				_readFileStream.Close();

				#endregion

				return true;
			}
			catch ( Exception _exception )
			{
				Label_Text.Text		= "[Exception] Msg : " + _exception.Message;

				string _logFileName = Path.GetFileName( iFilePath );

				LogConvertException( _exception.Message, _logFileName, GetConvertedFolderPath() );

				return false;
			}
		}

		private EExcelContentType GetExcelContentType( DataColumnCollection iDataColumn )
		{
			switch( iDataColumn[ 0 ].ToString() )
			{
				case ExcelContentType_Single:
					return EExcelContentType.Single;

				case ExcelContentType_Multiple:
					return EExcelContentType.Multiple;
			}

			return EExcelContentType.None;
		}

		#endregion

		#region Log

		private void LogTextResult( string iLog, string iLogFolderPath )
		{
			string _FileName = string.Format( "[{0}] Convert Excel Result.txt", DateTime.Now.ToString( "yyyy.MM.dd - HHmmss" ) );

			LogTextFile( iLog, _FileName, iLogFolderPath );
		}

		private void LogConvertException( string iLog, string iLogFileName, string iLogFolderPath )
		{
			string _FileName = string.Format( "[{0}] Convert Excel Exception - {1}.txt", DateTime.Now.ToString( "yyyy.MM.dd - HHmmss" ), iLogFileName );

			LogTextFile( iLog, _FileName, iLogFolderPath );
		}

		private void LogTextFile( string iLog, string iFileName, string iLogFolderPath )
		{
			string _FilePath = Path.Combine( iLogFolderPath, iFileName );

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
