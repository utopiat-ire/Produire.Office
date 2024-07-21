namespace LibTestApp
{
	partial class Form1
	{
		/// <summary>
		/// 必要なデザイナー変数です。
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary>
		/// 使用中のリソースをすべてクリーンアップします。
		/// </summary>
		/// <param name="disposing">マネージド リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Windows フォーム デザイナーで生成されたコード

		/// <summary>
		/// デザイナー サポートに必要なメソッドです。このメソッドの内容を
		/// コード エディターで変更しないでください。
		/// </summary>
		private void InitializeComponent()
		{
			this.WordButton = new System.Windows.Forms.Button();
			this.ExcelButton = new System.Windows.Forms.Button();
			this.SuspendLayout();
			// 
			// WordButton
			// 
			this.WordButton.Location = new System.Drawing.Point(156, 69);
			this.WordButton.Name = "WordButton";
			this.WordButton.Size = new System.Drawing.Size(99, 38);
			this.WordButton.TabIndex = 0;
			this.WordButton.Text = "Word";
			this.WordButton.UseVisualStyleBackColor = true;
			this.WordButton.Click += new System.EventHandler(this.WordButton_Click);
			// 
			// ExcelButton
			// 
			this.ExcelButton.Location = new System.Drawing.Point(290, 69);
			this.ExcelButton.Name = "ExcelButton";
			this.ExcelButton.Size = new System.Drawing.Size(99, 38);
			this.ExcelButton.TabIndex = 1;
			this.ExcelButton.Text = "Excel";
			this.ExcelButton.UseVisualStyleBackColor = true;
			this.ExcelButton.Click += new System.EventHandler(this.ExcelButton_Click);
			// 
			// Form1
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(10F, 18F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(800, 450);
			this.Controls.Add(this.ExcelButton);
			this.Controls.Add(this.WordButton);
			this.Name = "Form1";
			this.Text = "日本語Officeライブラリのテスト";
			this.ResumeLayout(false);

		}

		#endregion

		private System.Windows.Forms.Button WordButton;
		private System.Windows.Forms.Button ExcelButton;
	}
}

