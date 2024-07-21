using Produire.Office.エクセル;
using Produire.Office.ワード;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SampleApp
{
	public partial class Form1 : Form
	{
		public Form1()
		{
			InitializeComponent();
		}

		private void WordButton_Click(object sender, EventArgs e)
		{
			ワード マイアプリ = new ワード();
			マイアプリ.起動();
			var マイ文章 = マイアプリ.新規文書を作成();
			マイ文章.内容 = "こんにちは";
		}

		private void ExcelButton_Click(object sender, EventArgs e)
		{
			エクセル マイアプリ = new エクセル();
			マイアプリ.起動();
			var マイブック = マイアプリ.新しいワークブックを作成();
			var マイシート = マイブック.選択シート;
			マイシート.セル("A4").内容 = "こんにちは";
		}
	}
}
