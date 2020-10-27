// Copyright ©2020 Albert Ho ( rt135792005@gmail.com ). All rights reserved.
using System;
using System.IO;
using System.Text;
using System.Diagnostics;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Windows.Forms;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;

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

	public enum EExcelReadWork
	{
		None,
		GroupIndex,		// 讀取到 GroupIndex
		DataTypeRow,	// 讀取到資料格式 Row
		DataContentRow,	// 讀取到資料內容 Row
		End,
	}

	public partial class Form1 : Form
	{
		// Define
		const string SourceFolderName		= "Excel_SourceFolder";
		const string ConvertedFolderName	= "Excel_ConvertedFolder";
		static readonly string[] AcceptedExcelExtentionArray = new string[]{ ".xlsx" };
		const string OutputExtension		= ".bytes";		// https://docs.unity3d.com/Manual/class-TextAsset.html

		// Excel 關鍵字 - 表格類型
		const string ExcelContentType_Single	= "#Single";
		const string ExcelContentType_Multiple	= "#Multiple";
		
		// Excel 關鍵字 - 特殊關鍵字
		const string EExcelSymbol = "#";
		const string EExcelReadSymbol_GroupPrefix	= "#Group";
		const string EExcelReadSymbol_EndOfLine		= "#EOL";
		const string EExcelReadSymbol_EndOfSheet	= "#END";

		// Excel 關鍵字 - 類型
		const string EExcelReadSymbol_IndexType		= "#Index";
		const string EExcelReadSymbol_StringType	= "#String";
		const string EExcelReadSymbol_IntType		= "#Int";
		const string EExcelReadSymbol_FloatType		= "#Float";

		// 加密相關
		static readonly bool	EnableStreamEncrypt		= true;
		static readonly string	AesCrypto_EncryptKey	= "xKE2e5m8OWw2rvH7DNW2lWu8ZxOPdV22U9DnoMapnbI=";  // 備註：Base64 String
		static readonly string	AesCrypto_EncryptIV		= "6WLfMD6yNbWYcnxTKlC07A==";  // 備註：Base64 String

		// 計時器
		static readonly bool EnableStopwatch = true;

		// Log Output 相關 ( 備註：會把 Excel 讀到的檔案轉成 txt. 不過會大幅影響效能. 要修正效率的話要改成 StringBuilder.Append 再做測試 )
		static readonly bool EnableLogTxt = false;



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

			TimeSpan _totalConvertTime = new TimeSpan( 0, 0, 0 );
			Dictionary<string, string> _fileConvertTimeDic = new Dictionary<string, string>();

			string[] _filesPath = Directory.GetFiles( GetSourceFolderPath(), "*.*", SearchOption.AllDirectories );

			for( int i=0; i < _filesPath.Length; i++ )
			{
				// Stopwatch
				Stopwatch _stopwatch = new Stopwatch();
				if( EnableStopwatch )
				{
					_stopwatch.Start();
				}

				// 轉換檔案
				bool _result = TryConvertExcel( _filesPath[ i ] );
				if( _result )
					_successFilePath.Add( _filesPath[ i ] );
				else
					_failFilePath.Add( _filesPath[ i ] );

				if( EnableStopwatch )
				{
					_stopwatch.Stop();

					TimeSpan _ts = _stopwatch.Elapsed;
					_totalConvertTime = _totalConvertTime.Add( _ts );
					
					string _elapsedTime = string.Format( "{0}.{1}", _ts.Seconds, _ts.Milliseconds );
					_fileConvertTimeDic.Add( _filesPath[ i ], _elapsedTime );
				}
			}

			Label_Text.Text = string.Format( "ConvertExcel Done.\nSuccessed Num : {0}.\nFailed Num : {1}", _successFilePath.Count, _failFilePath.Count );


			// 輸出轉換結果 Log
			#region 輸出轉換結果 Log

			string _logText = string.Empty;
			_logText += string.Format( "Successed File Count : {0}\n", _successFilePath.Count );

			for( int i=0; i < _successFilePath.Count; i++ )
			{
				if( EnableStopwatch )
				{
					string _convertTimeStr = "N/A";
					_fileConvertTimeDic.TryGetValue( _successFilePath[ i ], out _convertTimeStr );
					_logText += string.Format( "- {0} ({1})\n", _successFilePath[ i ], _convertTimeStr );
				}
				else
					_logText += string.Format( "- {0}\n", _successFilePath[ i ] );
			}

			_logText += string.Format( "\n\n---\n\n" );
			_logText += string.Format( "Failed File Count : {0}\n", _failFilePath.Count );
			for( int i=0; i < _failFilePath.Count; i++ )
			{
				if( EnableStopwatch )
				{
					string _convertTimeStr = "N/A";
					_fileConvertTimeDic.TryGetValue( _failFilePath[ i ], out _convertTimeStr );
					_logText += string.Format( "- {0} ({1})\n", _failFilePath[ i ], _convertTimeStr );
				}
				else
					_logText += string.Format( "- {0}\n", _failFilePath[ i ] );
			}

			_logText += string.Format( "\n\n---\n\n" );
			_logText += "ConvertExcel Done.";

			if( EnableStopwatch )
			{
				_logText += string.Format( " ( TotalTime : {0}:{1}.{2} )", _totalConvertTime.Minutes, _totalConvertTime.Seconds, _totalConvertTime.Milliseconds );
			}

			LogTextResult( _logText, GetConvertedFolderPath() );
			
			#endregion
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
				
				#endregion

				// 讀取表格類型
				#region 讀取表格類型

				DataRowCollection _dataRow = _dataSet.Tables[ 0 ].Rows;
				EExcelContentType _excelContentType = GetExcelContentType( _dataRow );

				#endregion

				// 轉換 Excel 資料
				#region 轉換 Excel 資料

				bool _convertResult = false;

				switch( _excelContentType )
				{
					// 處理 單一表格
					case EExcelContentType.Single:
						_convertResult = ConvertExcelToBinary_Single( iFilePath, _dataSet );
						break;

					// 處理多重表格
					case EExcelContentType.Multiple:
						_convertResult = ConvertExcelToBinary_Multiple( iFilePath, _dataSet );
						break;
				}

				#endregion

				// 用完記得關掉
				_readFileStream.Close();

				return _convertResult;
			}
			catch( Exception _exception )
			{
				Label_Text.Text		= "[Exception] Msg : " + _exception.Message;

				string _logFileName = Path.GetFileName( iFilePath );

				LogConvertException( _exception.Message, _logFileName, GetConvertedFolderPath() );

				return false;
			}
		}

		private EExcelContentType GetExcelContentType( DataRowCollection iDataRow )
		{
			string _zerozeroSymbol = iDataRow[ 0 ][ 0 ].ToString();

			switch( _zerozeroSymbol )
			{
				case ExcelContentType_Single:
					return EExcelContentType.Single;

				case ExcelContentType_Multiple:
					return EExcelContentType.Multiple;
			}

			return EExcelContentType.None;
		}

		private bool ConvertExcelToBinary_Single( string iFilePath, DataSet iDataSet )
		{
			// 取單一表格的 Binary Stream
			MemoryStream _outputBinaryStream;
			ConvertExcelToBinary_Single( iFilePath, iDataSet, out _outputBinaryStream );

			if( _outputBinaryStream == null )
				return false;

			// 輸出檔案
			try
			{
				// 輸出路徑
				string _newFileName = string.Format( "{0}{1}", Path.GetFileNameWithoutExtension( iFilePath ), OutputExtension );
				string _newFilePath = iFilePath.Replace( GetSourceFolderPath(), GetConvertedFolderPath() );
				_newFilePath		= _newFilePath.Replace( Path.GetFileName( iFilePath ), _newFileName );

				// 如果輸出路徑不存在，建立資料夾
				string _newFileFolderPath = _newFilePath.Replace( Path.GetFileName( _newFilePath ), string.Empty );
				if( !Directory.Exists( _newFileFolderPath ) )
				{
					Directory.CreateDirectory( _newFileFolderPath );
				}

				// 避免忘了關造成 Stream 開著，改用 using				
				using( FileStream _fileStream = new FileStream( _newFilePath, FileMode.Create ) )
				{
					if( EnableStreamEncrypt )
					{
						AesCryptoServiceProvider _aesAlg = new AesCryptoServiceProvider();
						
						// ToDO : 保留, 之後做 Key 產生器可以用
						/*_aesAlg.GenerateKey();
						_aesAlg.GenerateIV();
						string _generateKey = Convert.ToBase64String( _aesAlg.Key );
						string _generateIV	= Convert.ToBase64String( _aesAlg.IV );*/

						byte[] _bytesKey = Convert.FromBase64String( AesCrypto_EncryptKey );
						byte[] _bytesIV	 = Convert.FromBase64String( AesCrypto_EncryptIV );
						ICryptoTransform _encryptor = _aesAlg.CreateEncryptor( _bytesKey, _bytesIV );

						using( CryptoStream _cryptoStream = new CryptoStream( _fileStream, _encryptor, CryptoStreamMode.Write ) )
						{
							using( BinaryWriter _binaryWriter = new BinaryWriter( _cryptoStream, Encoding.UTF8, true ) )
							{
								byte[] _byteArr = _outputBinaryStream.ToArray();
								_binaryWriter.Write( _outputBinaryStream.ToArray() );
							}
						}
					}
					else
					{
						using( BinaryWriter _binaryWriter = new BinaryWriter( _fileStream, Encoding.UTF8, true ) )
						{
							_binaryWriter.Write( _outputBinaryStream.ToArray() );
						}
					}					
				}

				return true;
			}
			catch( Exception _exception )
			{
				string _logFileName = Path.GetFileName( iFilePath );

				LogConvertException( _exception.Message, _logFileName, GetConvertedFolderPath() );

				return false;
			}
		}

		// 待改名 & 測試
		private bool ConvertExcelToBinary_Single( string iFilePath, DataSet iDataSet, out MemoryStream _iOutputBinaryStream )
		{
			DataRowCollection _dataRowCollection = iDataSet.Tables[ 0 ].Rows;
			DataColumnCollection _dataColumnCollection = iDataSet.Tables[ 0 ].Columns;

			// ToDo : 待重新命名
			string _testLogStr = string.Empty;

			// 資料
			int _groupIndex = -1;

			// EExcelReadWork.DataTypeRow 相關資料
			List<string> _typeList = null;  // 格式對應表
			string _typeHeaderStr = string.Empty;

			// EExcelReadWork.DataContentRow 相關資料
			int _dataRowCount = 0;
			MemoryStream _dataContentStream  = new MemoryStream();		
			BinaryWriter _dataContentBinaryWriter = new BinaryWriter( _dataContentStream, Encoding.UTF8, true );

			// 備註 : 讀取順序由左到右 ( A1:Z1 )，再來由上到下 ( A1:Z1 讀完後後讀 A2:Z2 )
			for( int _rowIndex = 0; _rowIndex < _dataRowCollection.Count; _rowIndex++ )
			{
				string _firstStr = _dataRowCollection[ _rowIndex ][ 0 ].ToString();

				// 檢查是否要 跳過處理 / 處理特殊功能
				EExcelReadWork _readWorkType = Get_StrReadWorkType( _firstStr );

				// 處理結束符號
				if( _readWorkType == EExcelReadWork.End )
					break;

				switch( _readWorkType )
				{					
					// ToDo : 表格類型應該也要寫進去
					// ToDo : 可以再加一個版本欄位

					// ToDo : 判斷 Group 符號, 回傳 Index
					case EExcelReadWork.GroupIndex:
						_groupIndex = GetConvertData_GroupIndex( _firstStr );
						break;

					case EExcelReadWork.DataTypeRow:
						_typeList = GetConvertData_StrTypeList( _dataRowCollection[ _rowIndex ], _dataColumnCollection.Count );
						
						// 把 Type 也寫進去檔案裡面. 可以留個紀錄 ( 也可以透過程式解開 )
						_typeHeaderStr = string.Empty;
						for( int i=0; i < _typeList.Count; i++ )
						{
							if( i == 0 )
								_typeHeaderStr = _typeList[ i ];
							else
								_typeHeaderStr += string.Format( ",{0}", _typeList[ i ] );
						}

						break;

					case EExcelReadWork.DataContentRow:
						MemoryStream _memoryStream;
						GetConvertData_DataContentRow( _dataRowCollection[ _rowIndex ], _typeList, out _memoryStream );
						
						if( _memoryStream == null )
						{
							_iOutputBinaryStream = null;
							return false;
						}
						else
						{
							_dataContentBinaryWriter.Write( _memoryStream.ToArray() );
							_dataRowCount++;
						}

						break;
				}

				// Log Text				
				if( EnableLogTxt )
				{
					for( int _colIndex = 0; _colIndex < _dataColumnCollection.Count; _colIndex++ )
					{
						string _readStr = _dataRowCollection[ _rowIndex ][ _colIndex ].ToString();
						if( _readStr == string.Empty )
							continue;

						_testLogStr += _readStr + "  ";
					}				
				
					_testLogStr += "\n";
				}
			}

			// Log Text
			if( EnableLogTxt )
			{
				string _logFileName = Path.GetFileName( iFilePath );
				LogTextFile( _testLogStr, "測試" +_logFileName + ".txt", GetConvertedFolderPath() );
			}
			

			// 回傳資料整理
			MemoryStream _outputBinaryStream = new MemoryStream();
			BinaryWriter _outputBinaryWriter = new BinaryWriter( _outputBinaryStream, Encoding.UTF8, true );			
			
			_outputBinaryWriter.Write( _typeHeaderStr );
			_outputBinaryWriter.Write( _dataRowCount );
			_outputBinaryWriter.Write( _dataContentStream.ToArray() );

			_iOutputBinaryStream = _outputBinaryStream;


			// 關閉資料流
			_dataContentBinaryWriter.Close();
			_dataContentStream.Close();

			_outputBinaryWriter.Close();
			_outputBinaryStream.Close();

			return true;
		}

		private bool ConvertExcelToBinary_Multiple( string iFilePath, DataSet iDataSet )
		{
			//DataRowCollection _dataRow = _dataSet.Tables[ 0 ].Rows;
			//DataColumnCollection _dataColumn = _dataSet.Tables[ 0 ].Columns;

			return false;
		}

		/// <summary>
		/// 根據字串內容回傳 EExcelReadWork 類型
		/// </summary>
		private EExcelReadWork Get_StrReadWorkType( string iCheckStr )
		{
			// 檢查起始符號
			bool _startWithSymbols = iCheckStr.StartsWith( EExcelSymbol );
			if( !_startWithSymbols )
				return EExcelReadWork.DataContentRow;

			if( iCheckStr.StartsWith( EExcelReadSymbol_GroupPrefix ) )
			{
				return EExcelReadWork.GroupIndex;
			}

			switch( iCheckStr )
			{
				case EExcelReadSymbol_IndexType:
				case EExcelReadSymbol_StringType:
				case EExcelReadSymbol_IntType:
				case EExcelReadSymbol_FloatType:

				case EExcelReadSymbol_EndOfLine:
					return EExcelReadWork.DataTypeRow;

				case EExcelReadSymbol_EndOfSheet:
					return EExcelReadWork.End;

				default:
					return EExcelReadWork.None;
			}
		}

		private int GetConvertData_GroupIndex( string iString )
		{
			try
			{
				string _index = iString.Replace( EExcelReadSymbol_GroupPrefix, string.Empty );

				return Convert.ToInt32( _index );
			}
			catch
			{
				return -1;
			}
		}

		private List<string> GetConvertData_StrTypeList( DataRow iDataRow, int iDataColumnCount )
		{
			List<string> _typeList = new List<string>();

			for( int _colIndex = 0; _colIndex < iDataColumnCount; _colIndex++ )
			{
				string _readStr = iDataRow[ _colIndex ].ToString();

				// EOL 檢查
				if( _readStr == EExcelReadSymbol_EndOfLine )
					break;

				_typeList.Add( _readStr );
			}

			return _typeList;
		}

		/// <summary>
		/// 把整排的 DataRow 轉成 BinaryStream
		/// </summary>
		/// <param name="iDataRow"></param>
		/// <param name="iDataColumnCount"></param>
		/// <param name="iTypeRefList"></param>
		/// <param name="iMemoryStream"></param>
		private void GetConvertData_DataContentRow( DataRow iDataRow, List<string> iTypeRefList, out MemoryStream iMemoryStream )
		{
			if( iTypeRefList == null )
			{
				iMemoryStream = null;
				return;
			}

			try
			{
				// 準備 Binary 空間
				MemoryStream _memoryStream = new MemoryStream();
				BinaryWriter _binaryWriter = new BinaryWriter( _memoryStream, Encoding.UTF8, true );

				for( int _colIndex = 0; _colIndex < iTypeRefList.Count; _colIndex++ )
				{
					string _typeStr = iTypeRefList[ _colIndex ];
					string _readStr = iDataRow[ _colIndex ].ToString();

					if( _typeStr == string.Empty )
						continue;
				
					// 根據類型寫入 Binary
					switch( _typeStr )
					{
						// Excel 關鍵字 - 類型
						#region Excel 關鍵字 - 類型
						
						case EExcelReadSymbol_IndexType:
							int _index = Convert.ToInt32( _readStr );
							_binaryWriter.Write( _index );
							continue;

						case EExcelReadSymbol_StringType:
							_binaryWriter.Write( _readStr );
							continue;

						case EExcelReadSymbol_IntType:
							int _intValue = Convert.ToInt32( _readStr );
							_binaryWriter.Write( _intValue );
							continue;

						case EExcelReadSymbol_FloatType:
							float _floatValue = (float)Convert.ToDouble( _readStr );
							_binaryWriter.Write( _floatValue );
							continue;

						#endregion

						case EExcelReadSymbol_EndOfLine:
							break;

						default:
							iMemoryStream = null;
							return;
					}
				}

				iMemoryStream = _memoryStream;
				return;
			}
			catch
			{
				iMemoryStream = null;
				return;
			}
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
