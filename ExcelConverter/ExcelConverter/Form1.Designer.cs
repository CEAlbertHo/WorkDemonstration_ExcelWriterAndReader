namespace ExcelConverter
{
	partial class Form1
	{
		/// <summary>
		/// 設計工具所需的變數。
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary>
		/// 清除任何使用中的資源。
		/// </summary>
		/// <param name="disposing">如果應該處置受控資源則為 true，否則為 false。</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Windows Form 設計工具產生的程式碼

		/// <summary>
		/// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改
		/// 這個方法的內容。
		/// </summary>
		private void InitializeComponent()
		{
			this.Btn_CreateDirectory = new System.Windows.Forms.Button();
			this.Btn_ConvertExcel = new System.Windows.Forms.Button();
			this.Label_Text = new System.Windows.Forms.Label();
			this.SuspendLayout();
			// 
			// Btn_CreateDirectory
			// 
			this.Btn_CreateDirectory.Font = new System.Drawing.Font("Arial", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.Btn_CreateDirectory.Location = new System.Drawing.Point(12, 196);
			this.Btn_CreateDirectory.Name = "Btn_CreateDirectory";
			this.Btn_CreateDirectory.Size = new System.Drawing.Size(200, 123);
			this.Btn_CreateDirectory.TabIndex = 0;
			this.Btn_CreateDirectory.Text = "Create Directory";
			this.Btn_CreateDirectory.UseVisualStyleBackColor = true;
			this.Btn_CreateDirectory.Click += new System.EventHandler(this.Btn_CreateDirectory_Click);
			// 
			// Btn_ConvertExcel
			// 
			this.Btn_ConvertExcel.Font = new System.Drawing.Font("Arial", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.Btn_ConvertExcel.Location = new System.Drawing.Point(372, 196);
			this.Btn_ConvertExcel.Name = "Btn_ConvertExcel";
			this.Btn_ConvertExcel.Size = new System.Drawing.Size(200, 123);
			this.Btn_ConvertExcel.TabIndex = 1;
			this.Btn_ConvertExcel.Text = "Convert Excel";
			this.Btn_ConvertExcel.UseVisualStyleBackColor = true;
			this.Btn_ConvertExcel.Click += new System.EventHandler(this.Btn_ConvertExcel_Click);
			// 
			// Label_Text
			// 
			this.Label_Text.AutoSize = true;
			this.Label_Text.Font = new System.Drawing.Font("Arial", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.Label_Text.Location = new System.Drawing.Point(12, 9);
			this.Label_Text.Name = "Label_Text";
			this.Label_Text.Size = new System.Drawing.Size(0, 27);
			this.Label_Text.TabIndex = 2;
			// 
			// Form1
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(584, 331);
			this.Controls.Add(this.Label_Text);
			this.Controls.Add(this.Btn_ConvertExcel);
			this.Controls.Add(this.Btn_CreateDirectory);
			this.Name = "Form1";
			this.Text = "Excel Converter";
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Button Btn_CreateDirectory;
		private System.Windows.Forms.Button Btn_ConvertExcel;
		private System.Windows.Forms.Label Label_Text;
	}
}

