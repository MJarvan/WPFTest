using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace WPFTest
{
	/// <summary>
	/// Window1.xaml 的交互逻辑
	/// </summary>
	public partial class ReNameWindow : Window
	{/// <summary>
	 /// 原名字符
	 /// </summary>
		public string nameStr
		{
			get; set;
		}

		public ReNameWindow()
		{
			InitializeComponent();
		}

		private void Window_Loaded(object sender, RoutedEventArgs e)
		{
			textblock.Text = "该图片原名为: " + nameStr + " ,请在下方填写你想要的名字";
			textbox.Focus();
		}

		private void comfirmButton_Click(object sender, RoutedEventArgs e)
		{
			Window window = Application.Current.MainWindow;
			MainWindow mainwindow = window as MainWindow;
			mainwindow.renameStr = textbox.Text;
			this.Close();
		}

		private void cancelButton_Click(object sender, RoutedEventArgs e)
		{
			this.Close();
		}
	}
}
