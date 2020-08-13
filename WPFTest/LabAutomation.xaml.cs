using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.HSSF.Util;
using System.Windows.Controls.Primitives;
using NPOI.SS.Util;
using System.Diagnostics;
using Microsoft.CSharp;
using System.CodeDom.Compiler;
using System.Reflection;
using System.Text.RegularExpressions;

namespace WPFTest
{
    /// <summary>
    /// LabAutomation.xaml 的交互逻辑
    /// </summary>
    public partial class LabAutomation : Window
    {
		/// <summary>
		/// txt内容
		/// </summary>
		List<List<string>> txtList = new List<List<string>>();

		/// <summary>
		/// 检测化合物名称合计
		/// </summary>
		List<KeyValuePair<string,string>> compoundsNameList = new List<KeyValuePair<string,string>>();
		List<KeyValuePair<string,string>> preCompoundsNameList = new List<KeyValuePair<string,string>>();

		/// <summary>
		/// 样品名称合计
		/// </summary>
		List<string> sampleNameList = new List<string>();
		/// <summary>
		/// 添加了平行样之后的合计
		/// </summary>
		List<string> newsampleNameList = new List<string>();
		/// <summary>
		/// 添加了平行样之后的分样
		/// </summary>
		List<KeyValuePair<List<string>,List<int>>> finalsampleNameList = new List<KeyValuePair<List<string>,List<int>>>();

		/// <summary>
		/// 委托单号
		/// </summary>
		string ReportNo = string.Empty;

		//调整一个表格的样品总列数
		//int importTakeNum = 8;
		int importTakeNum = 4;
		//调整一个横表格的总列数
		int verticalSheetColumnCount = 10;
		//调整一个竖表格的总列数
		int horizontalSheetColumnCount = 8;
		/// <summary>
		/// 每个化合物的datatable
		/// </summary>
		DataSet compoundsDataSet = new DataSet();
		public LabAutomation()
        {
            InitializeComponent();
        }

		private void Window_Loaded(object sender,RoutedEventArgs e)
		{
			topScrollViewer.DragEnter += scDragEnter;
			topScrollViewer.Drop += scDrop;
			mainScrollViewer.DragEnter += scDragEnter;
			mainScrollViewer.Drop += scDrop;
			samplingquantityLabel.Tag = 0;
			dilutionratioLabel.Tag = 1;
			constantvolumeLabel.Tag = 2;
		}

		/// <summary>
		/// 拖动放下
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void scDrop(object sender,DragEventArgs e)
		{
			//foreach(string str in e.Data.GetFormats())
			//{
			//	MessageBox.Show(str);
			//}
			ScrollViewer scrollViewer = sender as ScrollViewer;

			if (e.Data.GetDataPresent(DataFormats.FileDrop))
			{
				e.Effects = DragDropEffects.Link;

				string[] paths = (string[])e.Data.GetData(DataFormats.FileDrop);
				if (int.Parse(scrollViewer.Tag.ToString()) == 0)
				{
					//导入模板
					CreateTemplate(paths[0]);
				}
				else if (int.Parse(scrollViewer.Tag.ToString()) == 1)
				{
					//创建数据结构
					CreateTxt(paths[0]);
				}
			}
			e.Handled = true;
		}

		/// <summary>
		/// 导入模板到程序里面
		/// </summary>
		/// <param name="path"></param>
		private void CreateTemplate(string path)
		{
			string symbol = "：";
			if (File.Exists(path)) 
			{
				List<string> alldata = File.ReadAllLines(path,Encoding.UTF8).ToList();
				foreach(string data in alldata)
				{
					//加载常规设置项
					string key = data.Split(symbol)[0];
					string value = data.Split(symbol)[1];
					if ((key + symbol) == samplingquantityLabel.Content.ToString())
					{
						string strTextBox = Regex.Replace(value,"[a-z]","",RegexOptions.IgnoreCase);
						string strComboBox = value.Replace(strTextBox,string.Empty);
						samplingquantityTextBox.Text = strTextBox;
						foreach (ComboBoxItem comboBoxItem in samplingquantityComboBox.Items)
						{
							if (strComboBox == comboBoxItem.Content.ToString())
							{
								comboBoxItem.IsSelected = true;
							}
						}
					}
					else if ((key + symbol) == dilutionratioLabel.Content.ToString())
					{
						dilutionratioTextBox.Text = value;
					}
					else if ((key + symbol) == constantvolumeLabel.Content.ToString())
					{
						string strTextBox = Regex.Replace(value,"[a-z]","",RegexOptions.IgnoreCase);
						string strComboBox = value.Replace(strTextBox,string.Empty);
						constantvolumeTextBox.Text = strTextBox;
						foreach (ComboBoxItem comboBoxItem in constantvolumeComboBox.Items)
						{
							if (strComboBox == comboBoxItem.Content.ToString())
							{
								comboBoxItem.IsSelected = true;
							}
						}
					}
					else if ((key + symbol) == coefficientLabel.Content.ToString())
					{
						coefficientTextBox.Text = value;
					}
					else if ((key + symbol) == TargetCompanyLabel.Content.ToString())
					{
						foreach (ComboBoxItem comboBoxItem in TargetCompanyComboBox.Items)
						{
							if (value == comboBoxItem.Content.ToString())
							{
								comboBoxItem.IsSelected = true;
							}
						}
					}
					else if ((key + symbol) == AccuracyLabel.Content.ToString())
					{
						foreach (ComboBoxItem comboBoxItem in AccuracyComboBox.Items)
						{
							if (value == comboBoxItem.Content.ToString())
							{
								comboBoxItem.IsSelected = true;
							}
						}
					}
					else if ((key + symbol) == FormulaLabel.Content.ToString())
					{
						foreach (ComboBoxItem comboBoxItem in FormulaComboBox.Items)
						{
							if (value == comboBoxItem.Content.ToString())
							{
								comboBoxItem.IsSelected = true;
							}
						}
					}
					else if (key == testZDRadioButton.Content.ToString())
					{
						testZDRadioButton.IsChecked = true;
						foreach (ComboBoxItem comboBoxItem in ZDJCCompanyComboBox.Items)
						{
							if (value == comboBoxItem.Content.ToString())
							{
								comboBoxItem.IsSelected = true;
							}
						}
					}
					else if (key == testJCRadioButton.Content.ToString())
					{
						testJCRadioButton.IsChecked = true;
						foreach (ComboBoxItem comboBoxItem in ZDJCCompanyComboBox.Items)
						{
							if (value == comboBoxItem.Content.ToString())
							{
								comboBoxItem.IsSelected = true;
							}
						}
					}
					//加载化合物项
					else
					{
						//有没有添加化合物的文档
						if (maingrid.Children.Count > 0)
						{
							TabControl tabControl = maingrid.Children[0] as TabControl;
							foreach (TabItem tabItem in tabControl.Items)
							{
								StackPanel stackPanel = tabItem.Header as StackPanel;
								Label label = stackPanel.Children[0] as Label;
								TextBox textBox = stackPanel.Children[1] as TextBox;
								if (label.Content.ToString() == key)
								{
									textBox.Text = value;
								}
							}
						}
						else
						{
							KeyValuePair<string,string> keyValuePair = new KeyValuePair<string,string>(key,value);
							preCompoundsNameList.Add(keyValuePair);
						}
					}
				}
			}
		}
		/// <summary>
		/// 通过文本创造核心内容
		/// </summary>
		/// <param name="path"></param>
		private void CreateTxt(string path)
		{
			AllClear();
			if (File.Exists(path))
			{
				List<string> alldata = File.ReadAllLines(path,Encoding.UTF8).ToList();
				List<string> data = new List<string>();
				TabControl tabControl = new TabControl();
				tabControl.Name = "tabControl";
				for (int i = 0; i < alldata.Count; i++)
				{
					//""为txt文档的分页符
					if (alldata[i] != "")
					{
						data.Add(alldata[i]);
					}
					else
					{
						List<string> adddata = data.ToList();
						txtList.Add(adddata);
						data.Clear();
					}
				}

				foreach (List<string> vs in txtList)
				{
					bool isOK = CreateDataTable(tabControl,vs);
					if (!isOK)
					{
						return;
					}
				}
				AddParallelSamplesToList();

				maingrid.Children.Add(tabControl);
				ReportNoLabel.Content = ReportNo;
			}
		}

		/// <summary>
		/// 全部清空,重置
		/// </summary>
		private void AllClear()
		{
			txtList.Clear();
			compoundsNameList.Clear();
			sampleNameList.Clear();
			newsampleNameList.Clear();
			ReportNo = string.Empty;
			ReportNoLabel.Content = ReportNo;
			compoundsDataSet.Tables.Clear();
			maingrid.Children.Clear();
			finalsampleNameList.Clear();
		}

		private bool CreateDataTable(TabControl tabControl,List<string> vs)
		{
			//前三列是datatable的属性名和表头需要单独保存
			//string[] id = vs[0].Split("\t");
			sampleNameList.Clear();
			string[] name = vs[1].Split("\t");

			////缺少自我填写的最低检测质量浓度
			//if (name.Length < 3)
			//{
			//	MessageBox.Show(name[1] + "缺少最低检测质量浓度或检出限，请补充填写！");
			//	return false;
			//}
			//KeyValuePair<string,string> keyValuePair = new KeyValuePair<string,string>(name[1],name[2]);


			List<string> tablehead = vs[2].Split("\t").ToList();
			DataTable datatable = new DataTable();
			datatable.TableName = name[1];
			for (int i = 0; i < tablehead.Count; i++)
			{
				string headname = string.IsNullOrEmpty(tablehead[i]) ? "编号" : tablehead[i];
				datatable.Columns.Add(headname);
			}
			for (int j = 3; j < vs.Count; j++)
			{
				DataRow dr = datatable.NewRow();
				dr.ItemArray = vs[j].Split("\t");
				datatable.Rows.Add(dr);
			}
			//提取委托单号,删除后缀.gcd以及BQ的行
			for (int k = datatable.Rows.Count - 1; k >= 0; k--)
			{
				string oldname = datatable.Rows[k]["数据文件名"].ToString();
				//
				if (oldname.Contains("BQ"))
				{
					datatable.Rows[k].Delete();
				}
				else
				{
					string[] newname = oldname.Replace(".gcd","").Split(" ");

					if (sampleNameList.Count != datatable.Rows.Count)
					{
						if (newname.Length > 1)
						{
							ReportNo = (ReportNo == string.Empty) ? newname[0] : ReportNo;
							sampleNameList.Insert(0,newname[1]);
							datatable.Rows[k]["数据文件名"] = newname[1];
						}
						else
						{
							datatable.Rows[k]["数据文件名"] = newname[0];
							sampleNameList.Insert(0,newname[0]);
						}
					}
				}
			}
			datatable.AcceptChanges();
			//根据有机组要求只要三列
			DataTable newdatatable = new DataTable();
			newdatatable.TableName = datatable.TableName;
			//for (int l = 0; l < datatable.Columns.Count; l++)
			//{
			//	if (datatable.Columns[l].ColumnName == "编号" || datatable.Columns[l].ColumnName == "数据文件名" || datatable.Columns[l].ColumnName == "浓度")
			//	{
			//		DataColumn dataColumn = datatable.Columns[l];
			//		newdatatable.Columns.Add(dataColumn.ColumnName,dataColumn.DataType);
			//	}
			//}
			newdatatable.Columns.Add("编号");
			newdatatable.Columns.Add("数据文件名");
			newdatatable.Columns.Add("浓度");

			for (int i = 0; i < datatable.Rows.Count; i++)
			{
				DataRow dr = newdatatable.NewRow();
				for (int j = 0; j < newdatatable.Columns.Count; j++)
				{
					string newColumnName = newdatatable.Columns[j].ColumnName;
					dr[newColumnName] = datatable.Rows[i][newColumnName];
				}
				newdatatable.Rows.Add(dr);
			}

			TabItem tabItem = new TabItem();
			//tabItem.Header = name[1] + " | " + name[2];
			StackPanel stackPanel = CreateStackPanel(name[1]);
			tabItem.Header = stackPanel;
			DataGrid dg = new DataGrid();
			dg.Name = "dataGrid";
			dg.ItemsSource = newdatatable.DefaultView;
			dg.CanUserSortColumns = true;
			dg.CanUserReorderColumns = true;
			tabItem.Content = dg;
			tabControl.Items.Add(tabItem);
			compoundsDataSet.Tables.Add(newdatatable);

			return true;
		}

		/// <summary>
		/// 创建tabheader用的stackpanel
		/// </summary>
		/// <returns></returns>
		private StackPanel CreateStackPanel(string compoundsName)
		{
			StackPanel stackPanel = new StackPanel();
			stackPanel.Orientation = Orientation.Horizontal;
			stackPanel.VerticalAlignment = System.Windows.VerticalAlignment.Center;
			stackPanel.HorizontalAlignment = System.Windows.HorizontalAlignment.Center;
			Label label = new Label();
			label.Content = compoundsName;
			label.HorizontalContentAlignment = System.Windows.HorizontalAlignment.Center;
			label.HorizontalAlignment = System.Windows.HorizontalAlignment.Center;
			TextBox textBox = new TextBox();
			textBox.Width = 50;
			textBox.HorizontalAlignment = System.Windows.HorizontalAlignment.Center;
			textBox.VerticalAlignment = System.Windows.VerticalAlignment.Center;
			if (preCompoundsNameList.Count > 0)
			{
				foreach (KeyValuePair<string,string> keyValuePair in preCompoundsNameList)
				{
					if (keyValuePair.Key == compoundsName)
					{
						textBox.Text = keyValuePair.Value;
					}
				}
			}

			stackPanel.Children.Add(label);
			stackPanel.Children.Add(textBox);

			return stackPanel;
		}

		/// <summary>
		/// 添加平行样
		/// </summary>
		private void AddParallelSamplesToList()
		{
			//先加完所有的平行样再分组比较好
			List<string> psNameList = sampleNameList.ToList();
			List<int> numList = new List<int>();
			string Ebanlance = "Dup";
			string Cbanlance = "平均";
			List<string> addNameList = psNameList.Where(x => x.Contains(Ebanlance)).ToList();
			for (int i = 0; i < addNameList.Count; i++)
			{
				int num = psNameList.IndexOf(addNameList[i]) + 1;
				numList.Add(num);
			}
			for (int j = 0; j < numList.Count; j++)
			{
				if (j == 0)
				{
					psNameList.Insert(numList[j],addNameList[j].Replace(Ebanlance,Cbanlance));
				}
				else
				{
					psNameList.Insert(numList[j] + j,addNameList[j].Replace(Ebanlance,Cbanlance));
				}
			}

			newsampleNameList = psNameList.ToList();
			newsampleNameList.Add("以下空白");

			int Count = psNameList.Count % importTakeNum > 0 ? psNameList.Count / importTakeNum + 1 : psNameList.Count / importTakeNum;

			for (int i = 0; i < Count; i++)
			{
				List<int> finalnumList = new List<int>();
				if (i == Count - 1)
				{
					List<string> cellList = psNameList.ToList();
					if (cellList.Exists(x => x.Contains(Cbanlance)))
					{
						List<string> newcellList = cellList.Where(x => x.Contains(Cbanlance)).ToList();
						foreach (string strNum in newcellList)
						{
							finalnumList.Add(cellList.IndexOf(strNum));
						}
					}
					KeyValuePair<List<string>,List<int>> keyValuePair = new KeyValuePair<List<string>,List<int>>(cellList,finalnumList);
					finalsampleNameList.Add(keyValuePair);
				}
				else
				{
					List<string> cellList = psNameList.Take(importTakeNum).ToList();
					psNameList.RemoveRange(0,importTakeNum);
					if (cellList.Exists(x => x.Contains(Cbanlance)))
					{
						List<string> newcellList = cellList.Where(x => x.Contains(Cbanlance)).ToList();
						foreach (string strNum in newcellList)
						{
							finalnumList.Add(cellList.IndexOf(strNum));
						}
					}
					KeyValuePair<List<string>,List<int>> keyValuePair = new KeyValuePair<List<string>,List<int>>(cellList,finalnumList);
					finalsampleNameList.Add(keyValuePair);
				}
			}
		}


		/// <summary>
		/// 导出生成Excel
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void importExcel_Click(object sender,RoutedEventArgs e)
		{
			if (finalsampleNameList.Count == 0 || newsampleNameList.Count == 0)
			{
				return;
			}
			//判断化合物是否大于4，从而分割成横表或者竖表
			if (compoundsNameList.Count > 2)
			{
				CreateVerticalExcel();
			}
			else
			{
				CreateHorizontalExcel();
			}

		}

		/// <summary>
		/// 添加自己填的检出限
		/// </summary>
		private void AddDetectionLimit()
		{
			compoundsNameList.Clear();
			TabControl tabControl = maingrid.Children[0] as TabControl;
			foreach (TabItem tabItem in tabControl.Items)
			{
				string compoundsName = string.Empty;
				string modelC = string.Empty;
				StackPanel stackPanel = tabItem.Header as StackPanel;
				foreach (var item in stackPanel.Children)
				{
					if (item.GetType() == typeof(Label))
					{
						compoundsName = (item as Label).Content.ToString();
					}
					else if (item.GetType() == typeof(TextBox))
					{
						if ((item as TextBox).Text != null && (item as TextBox).Text != "" && (item as TextBox).Text != string.Empty)
						{

							modelC = (item as TextBox).Text;
						}
						else
						{
							return;
						}
					}
				}
				KeyValuePair<string,string> keyValuePair = new KeyValuePair<string,string>(compoundsName,modelC);
				compoundsNameList.Add(keyValuePair);
			}

			if (compoundsNameList.Count > 4)
			{
				KeyValuePair<string,string> keyValuePair = new KeyValuePair<string,string>("以下空白",string.Empty);
				compoundsNameList.Add(keyValuePair);
			}
		}

		/// <summary>
		/// 创建竖表Excel
		/// </summary>
		private void CreateHorizontalExcel()
		{
			var workbook = new HSSFWorkbook();
			var sheet = workbook.CreateSheet("色谱分析结果汇总1-水");
			sheet.ForceFormulaRecalculation = true;
			//设置顶部大标题样式
			HSSFCellStyle cellStyle = CreateStyle(workbook);
			HSSFCellStyle bordercellStyle = CreateStyle(workbook);
			bordercellStyle.BorderLeft = BorderStyle.Thin;
			bordercellStyle.BorderTop = BorderStyle.Thin;
			bordercellStyle.BorderLeft = BorderStyle.Thin;
			bordercellStyle.BorderRight = BorderStyle.Thin;
			//int Count = 0;
			//前九行 大表头
			for (int i = 0; i < 9; i++)
			{
				//第一行最右显示委托单号
				HSSFRow row = (HSSFRow)sheet.CreateRow(i); //创建行或者获取行
				row.HeightInPoints = 20;
				switch (i)
				{
					case 0:
						{
							var reportNameCell = row.CreateCell(horizontalSheetColumnCount - 3);
							reportNameCell.CellStyle = cellStyle;
							reportNameCell.SetCellValue("委托单号：");
							var reportCell = row.CreateCell(horizontalSheetColumnCount - 2);
							reportCell.CellStyle = cellStyle;
							reportCell.SetCellValue(ReportNo);
							CellRangeAddress region = new CellRangeAddress(i,i,horizontalSheetColumnCount - 2,horizontalSheetColumnCount - 1);
							sheet.AddMergedRegion(region);
							break;
						}
					case 1:
						{
							row.HeightInPoints = 30;
							var nameCell = row.CreateCell(0);
							var cellStyleFont = (HSSFFont)workbook.CreateFont(); //创建字体
							//假如字体大小只需要是粗体的话直接使用下面该属性即可
							cellStyleFont.IsBold = true; //字体加粗
							cellStyleFont.FontHeightInPoints = 20; //字体大小
							HSSFCellStyle newcellStyle = CreateStyle(workbook);
							newcellStyle.SetFont(cellStyleFont); //将字体绑定到样式
							nameCell.CellStyle = newcellStyle;
							nameCell.SetCellValue("色谱分析结果汇总表");
							CellRangeAddress region = new CellRangeAddress(i,i,0,horizontalSheetColumnCount - 1);
							sheet.AddMergedRegion(region);
							break;
						}
					case 2:
						{
							var samplekindCell = row.CreateCell(0);
							samplekindCell.CellStyle = cellStyle;
							samplekindCell.SetCellValue("样品类别：");
							//科学计数法
							/*var testCell = row.CreateCell(1);
							testCell.CellStyle = cellStyle;
							double testNum = double.Parse(ReportNo.Replace("W",""));
							string returnvalue = ScientificCounting(testNum);
							if (returnvalue.Contains("×"))
							{
								//excel富文本
								HSSFRichTextString rts1 = new HSSFRichTextString(returnvalue);
								var cellStyleFont = (HSSFFont)workbook.CreateFont(); //创建字体
								cellStyleFont.TypeOffset = FontSuperScript.Super;//字体上标下标
								rts1.ApplyFont(returnvalue.Length - 1,returnvalue.Length,cellStyleFont);
								testCell.SetCellValue(rts1);
							}*/
							CellRangeAddress firstregion = new CellRangeAddress(i,i,1,2);
							sheet.AddMergedRegion(firstregion);
							var instrumentnumberCell = row.CreateCell(3);
							instrumentnumberCell.CellStyle = cellStyle;
							instrumentnumberCell.SetCellValue("仪器编号：");
							var analysisdateCell = row.CreateCell(5);
							analysisdateCell.CellStyle = cellStyle;
							analysisdateCell.SetCellValue("分析日期：");
							CellRangeAddress secondregion = new CellRangeAddress(i,i,horizontalSheetColumnCount - 2,horizontalSheetColumnCount - 1);
							sheet.AddMergedRegion(secondregion);
							break;
						}
					case 3:
						{
							var instrumentkindCell = row.CreateCell(0);
							instrumentkindCell.CellStyle = cellStyle;
							instrumentkindCell.SetCellValue("仪器型号：");
							CellRangeAddress firstregion = new CellRangeAddress(i,i,1,4);
							sheet.AddMergedRegion(firstregion);
							var chromatographiccolumnCell = row.CreateCell(5);
							chromatographiccolumnCell.CellStyle = cellStyle;
							chromatographiccolumnCell.SetCellValue("色谱柱：");
							CellRangeAddress secondregion = new CellRangeAddress(i,i,horizontalSheetColumnCount - 2,horizontalSheetColumnCount - 1);
							sheet.AddMergedRegion(secondregion);
							break;
						}
					case 4:
						{
							row.HeightInPoints = 30;
							var methodbasisCell = row.CreateCell(0);
							methodbasisCell.CellStyle = cellStyle;
							methodbasisCell.SetCellValue("方法依据：");
							CellRangeAddress firstregion = new CellRangeAddress(i,i,1,4);
							sheet.AddMergedRegion(firstregion);
							var chromatographiccolumnCell = row.CreateCell(5);
							chromatographiccolumnCell.CellStyle = cellStyle;
							chromatographiccolumnCell.SetCellValue("计算公式：");
							var formulacell = row.CreateCell(horizontalSheetColumnCount - 2);
							formulacell.CellStyle = cellStyle;
							CellRangeAddress region = new CellRangeAddress(i,8,horizontalSheetColumnCount - 2,horizontalSheetColumnCount - 1);
							sheet.AddMergedRegion(region);
							//要和公式那一块绑定在一起
							StringBuilder stringBuilder = new StringBuilder("计算公式：" + FormulaComboBox.Text + "\n");
							if (FormulaComboBox.Text.Contains("X"))
							{
								stringBuilder.Append("X—水样TPH的质量浓度(" + ZDJCCompanyComboBox.Text + ")\n");
							}
							else
							{
								stringBuilder.Append("C—样品中目标物的质量浓度(" + ZDJCCompanyComboBox.Text + ")\n");
							}
							//测定浓度要根据她自己填的
							stringBuilder.Append("Ci——目标物上机测定浓度(" + TargetCompanyComboBox.Text + ")\n");
							if (FormulaComboBox.Text.Contains("V1"))
							{
								stringBuilder.Append("V1——定容体积(" + constantvolumeComboBox.Text + ")\n");
							}
							stringBuilder.Append("f——稀释倍数\n");
							stringBuilder.Append("V——取样量(" + samplingquantityComboBox.Text + ")\n");
							formulacell.SetCellValue(stringBuilder.ToString());
							break;
						}
					case 5:
						{
							var marknumberCell = row.CreateCell(0);
							marknumberCell.CellStyle = cellStyle;
							marknumberCell.SetCellValue("标曲编号：");
							CellRangeAddress firstregion = new CellRangeAddress(i,i,1,4);
							sheet.AddMergedRegion(firstregion);
							break;
						}
					case 6:
						{
							var analyticalconditionsCell = row.CreateCell(0);
							analyticalconditionsCell.CellStyle = cellStyle;
							analyticalconditionsCell.SetCellValue("分析条件：");
							CellRangeAddress firstregion = new CellRangeAddress(i,i + 2,1,4);
							sheet.AddMergedRegion(firstregion);
							break;
						}
					default:
						{
							break;
						}
				}
			}

			int Count = 4;
			//正规格式第一行表头
			for (int i = 9; i < 14; i++)
			{
				HSSFRow formalrow = (HSSFRow)sheet.CreateRow(i); //创建行或者获取行
				formalrow.HeightInPoints = 20;
				for (int j = 0; j < horizontalSheetColumnCount; j++)
				{
					var cell = formalrow.CreateCell(j);
					cell.CellStyle = bordercellStyle;
					if (i == 9)
					{
						if (j < Count)
						{
							CellRangeAddress region = new CellRangeAddress(9,13,j,j);
							sheet.AddMergedRegion(region);
						}
						switch (j)
						{
							case 0:
								{
									cell.SetCellValue("样品编号");
									break;
								}
							case 1:
								{
									cell.SetCellValue("水样体积\nV(" + samplingquantityComboBox.Text + ")");
									break;
								}
							case 2:
								{
									cell.SetCellValue("稀释倍数f");
									break;
								}
							case 3:
								{
									cell.SetCellValue("试样体积\nV1(" + constantvolumeComboBox.Text + ")");
									break;
								}
							case 4:
								{
									CellRangeAddress region = new CellRangeAddress(9,9,j,horizontalSheetColumnCount - 1);
									sheet.AddMergedRegion(region);
									cell.SetCellValue("目标化合物");
									break;
								}
							default:
								{
									cell.SetCellValue(string.Empty);
									break;
								}
						}
					}
					else if (i == 11 && j == 4)
					{
						cell.SetCellValue("-");
						CellRangeAddress targetregion = new CellRangeAddress(formalrow.RowNum,formalrow.RowNum,j,j + 1);
						sheet.AddMergedRegion(targetregion);
					}
					else if (i == 11 && j == 6)
					{
						if (testZDRadioButton.IsChecked == true)
						{
							cell.SetCellValue(testZDRadioButton.Content + "(" + ZDJCCompanyComboBox.Text + ")");
						}
						else if (testJCRadioButton.IsChecked == true)
						{
							cell.SetCellValue(testJCRadioButton.Content + "(" + ZDJCCompanyComboBox.Text + ")");
						}
						CellRangeAddress outregion = new CellRangeAddress(formalrow.RowNum,formalrow.RowNum,j,horizontalSheetColumnCount - 1);
						sheet.AddMergedRegion(outregion);
					}
					else if (i == 12 && j == 4)
					{
						cell.SetCellValue("-");
						CellRangeAddress targetregion = new CellRangeAddress(formalrow.RowNum,formalrow.RowNum,j,j + 1);
						sheet.AddMergedRegion(targetregion);
					}
					else if (i == 13 && j == 4)
					{
						cell.SetCellValue("目标化合物测定值Ci(" + TargetCompanyComboBox.Text + ")");
						CellRangeAddress targetregion = new CellRangeAddress(formalrow.RowNum,formalrow.RowNum,j,j + 1);
						sheet.AddMergedRegion(targetregion);
					}
					else if (i == 13 && j == 6)
					{
						cell.SetCellValue("目标化合物浓度C(" + ZDJCCompanyComboBox.Text + ")");
						CellRangeAddress targetregion = new CellRangeAddress(formalrow.RowNum,formalrow.RowNum,j,horizontalSheetColumnCount - 1);
						sheet.AddMergedRegion(targetregion);
					}
					else
					{
						cell.SetCellValue(string.Empty);
					}
				}
			}
			foreach (KeyValuePair<string,string> keyValuePair in compoundsNameList)
			{
				CreateHorizontalSheet(sheet,bordercellStyle,keyValuePair.Key,keyValuePair.Value,Count);
				Count++;
			}

			//自动调整列距
			for (int i = 0; i < horizontalSheetColumnCount; i++)
			{
				//if (i == horizontalSheetColumnCount - 1)
				//{
				//	sheet.SetColumnWidth(i,20 * 256);
				//}
				//else
				//{
				//	sheet.AutoSizeColumn(i);
				//}
				sheet.AutoSizeColumn(i);
				int width = sheet.GetColumnWidth(i);
				if (width < 15 * 256)
				{
					sheet.SetColumnWidth(i,15 * 256);
				}
			}

			ExportToExcel(workbook);
		}



		/// <summary>
		/// 创建横表Excel
		/// </summary>
		private void CreateVerticalExcel()
		{
			var workbook = new HSSFWorkbook();
			var sheet = workbook.CreateSheet("色谱分析结果汇总1-水");
			sheet.ForceFormulaRecalculation = true;
			//设置顶部大标题样式
			HSSFCellStyle cellStyle = CreateStyle(workbook);

			int Count = 0;
			foreach (KeyValuePair<List<string>,List<int>> keyValuePair in finalsampleNameList)
			{
				CreateVerticalSheet(sheet,cellStyle,keyValuePair.Key,keyValuePair.Value,Count);
				Count++;
			}

			// 自动调整列距
			for (int i = 0; i < verticalSheetColumnCount * Count; i++)
			{
				if (i % verticalSheetColumnCount == 0)
				{
					sheet.SetColumnWidth(i,40 * 256);
				}
				else
				{
					sheet.AutoSizeColumn(i);
				}
			}

			ExportToExcel(workbook);
		}

		/// <summary>
		/// 导出到Excel
		/// </summary>
		/// <param name="workbook"></param>
		private void ExportToExcel(HSSFWorkbook workbook)
		{
			//自己选位置
			/*System.Windows.Forms.FolderBrowserDialog fbd = new System.Windows.Forms.FolderBrowserDialog();
			fbd.ShowDialog();
			if (fbd.SelectedPath != string.Empty)
			{
				string filename = sheet.SheetName + ".xls";
				string path = System.IO.Path.Combine(fbd.SelectedPath,filename);
				using (FileStream stream = new FileStream(path,FileMode.OpenOrCreate,FileAccess.ReadWrite))
				{
					workbook.Write(stream);
					stream.Flush();
				}
			}*/
			//特定位置
			try
			{
				string path = @"E:\CreateExcel\" + ReportNo + @"\";
				//创建用户临时图片文件夹或者清空临时文件夹所有文件
				if (!Directory.Exists(path))
				{
					Directory.CreateDirectory(path);
				}
				string filename = ReportNo + "-" + workbook.GetSheetAt(0).SheetName + ".xls";
				string fullpath = System.IO.Path.Combine(path,filename);
				if (File.Exists(fullpath))
				{
					File.Delete(fullpath);
				}
				using (FileStream stream = new FileStream(fullpath,FileMode.OpenOrCreate,FileAccess.ReadWrite))
				{
					workbook.Write(stream);
					stream.Flush();
				}
				Process process = new Process();
				ProcessStartInfo processStartInfo = new ProcessStartInfo(fullpath);
				processStartInfo.UseShellExecute = true;
				process.StartInfo = processStartInfo;
				process.Start();
			}
			catch (Exception e)
			{
				MessageBox.Show(e.Message);
			}
		}

		/// <summary>
		/// 创建竖表
		/// </summary>
		/// <param name="sheet"></param>
		/// <param name="cellStyle"></param>
		/// <param name="compoundName"></param>
		/// <param name="modelC"></param>
		/// <param name="Count"></param>
		private void CreateHorizontalSheet(ISheet sheet,HSSFCellStyle cellStyle,string compoundName,string modelC,int Count)
		{
			HSSFRow samplerow = (HSSFRow)sheet.GetRow(10);
			HSSFRow limitrow = (HSSFRow)sheet.GetRow(12);
			var sampleCell = samplerow.GetCell(Count);
			sampleCell.SetCellValue(compoundName);//合并单元格后，只需对第一个位置赋值即可（TODO:顶部标题）
			var anSampleCell = samplerow.GetCell(Count + 2);
			anSampleCell.SetCellValue(compoundName);//合并单元格后，只需对第一个位置赋值即可（TODO:顶部标题）
			var limitCell = limitrow.GetCell(Count + 2);
			limitCell.SetCellValue(modelC);//合并单元格后，只需对第一个位置赋值即可（TODO:顶部标题）

			foreach (string sampleName in newsampleNameList.OrderBy(x=>x.Trim()))
			{
				for (int i = 0; i < horizontalSheetColumnCount; i++)
				{
					HSSFRow row = (i == 0 && Count == 4) ? (HSSFRow)sheet.CreateRow(sheet.PhysicalNumberOfRows) : (HSSFRow)sheet.GetRow(14 + newsampleNameList.OrderBy(x => x.Trim()).ToList().IndexOf(sampleName));
					row.HeightInPoints = 20;
					int a = row.LastCellNum;
					var cell = (row.Cells.Contains(row.GetCell(i))) ? row.GetCell(i) : row.CreateCell(i);
					cell.CellStyle = cellStyle;

					if (i == 0 && Count == 4)
					{
						cell.SetCellValue(sampleName);
					}
					else if (i == 1 && Count == 4)
					{
						string setvalue = (sampleName == "以下空白") ? string.Empty : samplingquantityTextBox.Text;
						cell.SetCellValue(setvalue);
					}
					else if (i == 2 && Count == 4)
					{
						string setvalue = (sampleName == "以下空白") ? string.Empty : dilutionratioTextBox.Text;
						cell.SetCellValue(setvalue);
					}
					else if (i == 3 && Count == 4)
					{
						string setvalue = (sampleName == "以下空白") ? string.Empty : constantvolumeTextBox.Text;
						cell.SetCellValue(setvalue);
					}
					else if (i == Count)
					{

						if (sampleName.Contains("平均"))
						{
							cell.SetCellValue("-----");
						}
						else
						{
							foreach (DataTable dataTable in compoundsDataSet.Tables)
							{
								if (dataTable.TableName == compoundName)
								{
									//找到该化合物对应的样品编号和浓度数据
									foreach (DataRow dataRow in dataTable.Rows)
									{
										string dtsampleName = dataRow["数据文件名"].ToString();
										if (dtsampleName == sampleName)
										{
											string potency = dataRow["浓度"].ToString();
											cell.SetCellValue(potency);
										}
									}
								}
							}
						}
					}
					else if (i == Count + 2)
					{
						if (sampleName.Contains("平均"))
						{
							//获取平均样的位置
							int num = newsampleNameList.IndexOf(sampleName);
							//平均样的前两个加起来除以二就是平均值
							string setvalue = CompareCompoundWithFormula(compoundName,modelC,newsampleNameList[num - 1],newsampleNameList[num - 2]);
							cell.SetCellValue(setvalue);
						}
						else
						{
							string setvalue = (sampleName == "以下空白") ? string.Empty : CompareCompoundWithFormula(compoundName,modelC,sampleName);
							cell.SetCellValue(setvalue);
						}
					}
				}
            }
        }


		/// <summary>
		/// 创建横表
		/// </summary>
		/// <param name="sheet"></param>
		/// <param name="cellStyle"></param>
		/// <param name="cellList"></param>
		/// <param name="advantageNum"></param>
		/// <param name="Count"></param>
		private void CreateVerticalSheet(ISheet sheet,HSSFCellStyle cellStyle,List<string> cellList,List<int> advantageNum,int Count)
		{
			cellStyle.BorderBottom = BorderStyle.Thin;
			cellStyle.BorderRight = BorderStyle.Thin;
			cellStyle.BorderTop = BorderStyle.Thin;
			cellStyle.BorderLeft = BorderStyle.Thin;
			//前四行
			for (int i = 0; i < 4; i++)
			{
				//第一个单元格 公式
				HSSFRow row = (Count == 0) ? (HSSFRow)sheet.CreateRow(i) : (HSSFRow)sheet.GetRow(i); //创建行或者获取行
				row.HeightInPoints = 20;

				if (i == 3)
				{
					row.HeightInPoints = 50;
				}
				var formulacell = (Count == 0) ? row.CreateCell(i) : row.CreateCell(verticalSheetColumnCount * Count);
				formulacell.CellStyle = cellStyle;
				if (i == 0)
				{
					CellRangeAddress region = new CellRangeAddress(0,3,verticalSheetColumnCount * Count,verticalSheetColumnCount * Count);
					sheet.AddMergedRegion(region);
					//要和公式那一块绑定在一起
					StringBuilder stringBuilder = new StringBuilder("计算公式：" + FormulaComboBox.Text + "\n");
					if (FormulaComboBox.Text.Contains("X"))
					{
						stringBuilder.Append("X—水样TPH的质量浓度(" + ZDJCCompanyComboBox.Text + ")\n");
					}
					else
					{
						stringBuilder.Append("C—样品中目标物的质量浓度(" + ZDJCCompanyComboBox.Text + ")\n");
					}
					stringBuilder.Append("Ci——目标物上机测定浓度(" + TargetCompanyComboBox.Text + ")\n");
					if (FormulaComboBox.Text.Contains("V1"))
					{
						stringBuilder.Append("V1——定容体积(" + constantvolumeComboBox.Text + ")\n");
					}
					stringBuilder.Append("f——稀释倍数\n");
					stringBuilder.Append("V——取样量(" + samplingquantityComboBox.Text + ")\n");
					formulacell.SetCellValue(stringBuilder.ToString());//合并单元格后，只需对第一个位置赋值即可（TODO:顶部标题）

				}
				//从第二列开始
				for (int j = verticalSheetColumnCount * Count + 1; j < verticalSheetColumnCount * Count + verticalSheetColumnCount; j++)
				{
					var cell = row.CreateCell(j);
					cell.CellStyle = cellStyle;
					//按照已有的弄
					//用于最右边的那一版导出，因为有以下空白的特殊情况
					if (j - verticalSheetColumnCount * Count - 2 >= cellList.Count && j - verticalSheetColumnCount * Count - 2 < importTakeNum)
					{
						string setvalue = (i == 0) ? "以下空白" : "";
						cell.SetCellValue(setvalue);
					}
					else if (j - verticalSheetColumnCount * Count - 2 >= cellList.Count + importTakeNum && j - verticalSheetColumnCount * Count - 2 < importTakeNum * 2)
					{
						string setvalue = (i == 0) ? "以下空白" : "";
						cell.SetCellValue(setvalue);
					}
					else if (i == (int)samplingquantityLabel.Tag)
					{
						//常数
						if ((j - verticalSheetColumnCount * Count) == 1)
						{
							cell.SetCellValue(samplingquantityLabel.Content.ToString().Replace("：",string.Empty) + "(" + samplingquantityComboBox.Text + ")");
						}
						//是Ci浓度,都为-
						else if (j < verticalSheetColumnCount * Count + cellList.Count + 2)
						{
							cell.SetCellValue("-");
						}
						//是C浓度
						else if (advantageNum.Count > 0)
						{
							foreach (int num in advantageNum)
							{
								if (j - verticalSheetColumnCount * Count - 2 == num + cellList.Count)
								{
									cell.SetCellValue("-");
									break;
								}
								else
								{
									cell.SetCellValue(samplingquantityTextBox.Text);
								}
							}
						}
						else
						{
							cell.SetCellValue(samplingquantityTextBox.Text);
						}
					}
					else if (i == (int)constantvolumeLabel.Tag)
					{
						//常数
						if ((j - verticalSheetColumnCount * Count) == 1)
						{
							cell.SetCellValue(constantvolumeLabel.Content.ToString().Replace("：",string.Empty) + "(" + constantvolumeComboBox.Text + ")");
						}
						//是Ci浓度,都为-
						else if (j < verticalSheetColumnCount * Count + cellList.Count + 2)
						{
							cell.SetCellValue("-");
						}
						//是C浓度
						else if (advantageNum.Count > 0)
						{
							foreach (int num in advantageNum)
							{
								if (j - verticalSheetColumnCount * Count - 2 == num + cellList.Count)
								{
									cell.SetCellValue("-");
									break;
								}
								else
								{
									cell.SetCellValue(constantvolumeTextBox.Text);
								}
							}
						}
						else
						{
							cell.SetCellValue(constantvolumeTextBox.Text);
						}
					}
					else if (i == (int)dilutionratioLabel.Tag)
					{
						//常数
						if ((j - verticalSheetColumnCount * Count) == 1)
						{
							cell.SetCellValue(dilutionratioLabel.Content.ToString().Replace("：",string.Empty));
						}
						//是Ci浓度,都为-
						else if (j < verticalSheetColumnCount * Count + cellList.Count + 2)
						{
							cell.SetCellValue("-");
						}
						//是C浓度
						else if (advantageNum.Count > 0)
						{
							foreach (int num in advantageNum)
							{
								if (j - verticalSheetColumnCount * Count - 2 == num + cellList.Count)
								{
									cell.SetCellValue("-");
									break;
								}
								else
								{
									cell.SetCellValue(dilutionratioTextBox.Text);
								}
							}
						}
						else
						{
							cell.SetCellValue(dilutionratioTextBox.Text);
						}
					}
					//第四行
					else if ((j - verticalSheetColumnCount * Count) == 1)
					{
						cell.SetCellValue("样品编号");
					}
					else
					{
						//是Ci浓度的样品
						if (j - verticalSheetColumnCount * Count - 2 < cellList.Count)
						{
							cell.SetCellValue(cellList[j - verticalSheetColumnCount * Count - 2]);
						}
						//是C浓度
						else if (j - verticalSheetColumnCount * Count - 2 < cellList.Count * 2)
						{
							cell.SetCellValue(cellList[j - verticalSheetColumnCount * Count - 2 - cellList.Count]);
						}
						else if (j - verticalSheetColumnCount * Count - 2 >= cellList.Count * 2 && j - verticalSheetColumnCount * Count - 2 < cellList.Count + importTakeNum)
						{
							cell.SetCellValue(cellList[j - verticalSheetColumnCount * Count - 2 - importTakeNum]);
						}
					} 

				}
			}
			//第四行表头
			HSSFRow HearderRow = (Count == 0) ? (HSSFRow)sheet.CreateRow(4) : (HSSFRow)sheet.GetRow(4); //创建行或者获取行
			HearderRow.HeightInPoints = 20;
			for (int k = verticalSheetColumnCount * Count; k < verticalSheetColumnCount * (Count + 1); k++)
			{
				var cell = HearderRow.CreateCell(k);
				cell.CellStyle = cellStyle;
				//合并Ci
				if (!sheet.MergedRegions.Exists(x => x.FirstRow == 4 && x.LastRow == 4 && x.FirstColumn == 2 + verticalSheetColumnCount * Count && x.LastColumn == 5 + verticalSheetColumnCount * Count))
				{
					CellRangeAddress newregion = new CellRangeAddress(4,4,2 + verticalSheetColumnCount * Count,5 + verticalSheetColumnCount * Count);
					sheet.AddMergedRegion(newregion);
				}
				//合并C
				if (!sheet.MergedRegions.Exists(x => x.FirstRow == 4 && x.LastRow == 4 && x.FirstColumn == 6 + verticalSheetColumnCount * Count && x.LastColumn == 9 + verticalSheetColumnCount * Count))
				{
					CellRangeAddress newregion = new CellRangeAddress(4,4,6 + verticalSheetColumnCount * Count,9 + verticalSheetColumnCount * Count);
					sheet.AddMergedRegion(newregion);
				}
				if (k == 0 + verticalSheetColumnCount * Count)
				{
					cell.SetCellValue("目标化合物");//合并单元格后，只需对第一个位置赋值即可（TODO:顶部标题）
				}
				else if (k == 1 + verticalSheetColumnCount * Count)
				{
					if (testZDRadioButton.IsChecked == true)
					{
						cell.SetCellValue(testZDRadioButton.Content + "(" + ZDJCCompanyComboBox.Text + ")");
					}
					else if (testJCRadioButton.IsChecked == true)
					{
						cell.SetCellValue(testJCRadioButton.Content + "(" + ZDJCCompanyComboBox.Text + ")");
					}
				}
				else if (k == 2 + verticalSheetColumnCount * Count)
				{
					cell.SetCellValue("目标物上机测定浓度 Ci (" + TargetCompanyComboBox.Text + ")");
				}
				else if (k == 6 + verticalSheetColumnCount * Count)
				{
					cell.SetCellValue("目标化合物浓度 C (" + ZDJCCompanyComboBox.Text + ")");
				}
			}

			for (int k = 0; k < compoundsNameList.Count; k++)
			{
				HSSFRow compoundsRow = (Count == 0) ? (HSSFRow)sheet.CreateRow(4 + k + 1) : (HSSFRow)sheet.GetRow(4 + k + 1); //创建行或者获取行
				compoundsRow.HeightInPoints = 20;
				string compoundName = compoundsNameList[k].Key;
				string modelC = compoundsNameList[k].Value;
				for (int l = verticalSheetColumnCount * Count; l < verticalSheetColumnCount * Count + verticalSheetColumnCount; l++)
				{
					var compoundsCell = compoundsRow.CreateCell(l);
					compoundsCell.CellStyle = cellStyle;
					{
						//第一列是化合物名称
						if (l - verticalSheetColumnCount * Count == 0)
						{
							compoundsCell.SetCellValue(compoundName);
						}
						//第二列是判断标准
						else if (l - verticalSheetColumnCount * Count == 1)
						{
							compoundsCell.SetCellValue(modelC);
						}
						//前四列是浓度度数
						else if (l - verticalSheetColumnCount * Count - 2 < cellList.Count)
						{
							if (modelC == string.Empty)
							{
								compoundsCell.SetCellValue(string.Empty);
							}
							else
							{
								string sampleName = cellList[l - verticalSheetColumnCount * Count - 2];
								if (sampleName.Contains("平均"))
								{
									compoundsCell.SetCellValue("-----");
								}
								else
								{
									foreach (DataTable dataTable in compoundsDataSet.Tables)
									{
										if (dataTable.TableName == compoundName)
										{
											//找到该化合物对应的样品编号和浓度数据
											foreach (DataRow dataRow in dataTable.Rows)
											{
												string dtsampleName = dataRow["数据文件名"].ToString();
												if (dtsampleName == sampleName)
												{
													string potency = dataRow["浓度"].ToString();
													compoundsCell.SetCellValue(potency);
												}
											}
										}
									}
								}
							}
						}
						else if (cellList.Count == importTakeNum)
						{
							//后四列是计算判断
							if (l - verticalSheetColumnCount * Count - 2 >= cellList.Count && l - verticalSheetColumnCount * Count - 2 < cellList.Count * 2)
							{
								if (modelC == string.Empty)
								{
									compoundsCell.SetCellValue(string.Empty);
								}
								else
								{
									string sampleName = cellList[l - verticalSheetColumnCount * Count - 2 - cellList.Count];
									if (sampleName.Contains("平均"))
									{
										//获取平均样的位置
										int num = newsampleNameList.IndexOf(sampleName);
										//平均样的前两个加起来除以二就是平均值
										string setvalue = CompareCompoundWithFormula(compoundName,modelC,newsampleNameList[num - 1],newsampleNameList[num - 2]);
										compoundsCell.SetCellValue(setvalue);
									}
									else
									{
										string setvalue = CompareCompoundWithFormula(compoundName,modelC,sampleName);
										compoundsCell.SetCellValue(setvalue);
									}
								}
							} 
						}
						else if (cellList.Count < importTakeNum)
						{
							//最后一版特殊判断
							if (l - verticalSheetColumnCount * Count - 2 >= importTakeNum && l - verticalSheetColumnCount * Count - 2 < cellList.Count + importTakeNum)
							{
								if (modelC == string.Empty)
								{
									compoundsCell.SetCellValue(string.Empty);
								}
								else
								{
									string sampleName = cellList[l - verticalSheetColumnCount * Count - 2 - importTakeNum];
									if (sampleName.Contains("平均"))
									{
										//获取平均样的位置
										int num = newsampleNameList.IndexOf(sampleName);
										//平均样的前两个加起来除以二就是平均值
										string setvalue = CompareCompoundWithFormula(compoundName,modelC,newsampleNameList[num - 1],newsampleNameList[num - 2]);
										compoundsCell.SetCellValue(setvalue);
									}
									else
									{
										string setvalue = CompareCompoundWithFormula(compoundName,modelC,sampleName);
										compoundsCell.SetCellValue(setvalue);
									}
								}
							}
						}


					}
				}
			}
		}

		/// <summary>
		/// 科学计数法
		/// </summary>
		/// <param name="testNum"></param>
		/// <returns></returns>
		private string ScientificCounting(double testNum)
		{
			string returnnum = string.Empty;
			string oneNum = "1";
			if (testNum.ToString().Length > 4)
			{
				for (int i = 0; i < testNum.ToString().Length - 1; i++)
				{
					oneNum += "0";
				}

				double onenum = double.Parse(oneNum);
				returnnum = (testNum / onenum).ToString() + "×" + "10" + (testNum.ToString().Length - 1).ToString();
			}
			return returnnum;
		}

		/// <summary>
		/// 计算平行样浓度平均值
		/// </summary>
		/// <param name="compoundName"></param>
		/// <param name="modelC"></param>
		/// <param name="sampleName1"></param>
		/// <param name="sampleName2"></param>
		/// <returns></returns>
		private string CompareCompoundWithFormula(string compoundName,string modelC,string sampleName1,string sampleName2)
		{
			//计算公式C = Ci×f×V1 / V * k
			//稀释倍数
			double f = double.NaN;
			//定容体积
			double V1 = double.NaN;
			//取样量
			double V = double.NaN;
			//系数k
			double k = double.NaN;
			if (Regex.IsMatch(dilutionratioTextBox.Text,"^([0-9]{1,}[.][0-9]*)$"))
			{
				f = double.Parse(dilutionratioTextBox.Text);
			}
			if (Regex.IsMatch(constantvolumeTextBox.Text,"^([0-9]{1,}[.][0-9]*)$"))
			{
				V1 = double.Parse(constantvolumeTextBox.Text);
			}
			if (Regex.IsMatch(samplingquantityTextBox.Text,"^([0-9]{1,}[.][0-9]*)$"))
			{
				V = double.Parse(samplingquantityTextBox.Text);
			}
			if (Regex.IsMatch(coefficientTextBox.Text,"^([0-9]{1,}[.][0-9]*)$"))
			{
				k = double.Parse(coefficientTextBox.Text);
			}
			//目标物上机测定浓度
			double Ci;
			double C1 = double.NaN;
			double C2 = double.NaN;
			foreach (DataTable dataTable in compoundsDataSet.Tables)
			{
				//找到该化合物对应的datatable
				if (dataTable.TableName == compoundName)
				{
					for (int i = 0; i < dataTable.Rows.Count; i++)
					{
						for (int j = 0; j < dataTable.Columns.Count; j++)
						{
							//找到该化合物对应的样品编号和浓度数据
							string dtsampleName = dataTable.Rows[i][j].ToString();

							if (dtsampleName == sampleName1)
							{
								string potency = dataTable.Rows[i]["浓度"].ToString();
								if (!potency.Contains("-"))
								{
									Ci = double.Parse(potency);
									//公式计算
									//先用写死的，然后之后学习反射
									if (FormulaComboBox.Text.Contains("V1"))
									{
										C1 = Ci * f * V1 / V * k;
									}
									else
									{
										C1 = Ci * f / V * k;
									}
									
								}
							}
							if (dtsampleName == sampleName2)
							{
								string potency = dataTable.Rows[i]["浓度"].ToString();
								if (!potency.Contains("-"))
								{
									Ci = double.Parse(potency);
									//公式计算
									//先用写死的，然后之后学习反射
									if (FormulaComboBox.Text.Contains("V1"))
									{
										C2 = Ci * f * V1 / V * k;
									}
									else
									{
										C2 = Ci * f / V * k;
									}

								}
							}
						}
					}
				}
			}

			double C = (C1 + C2) / 2;
			if (C > double.Parse(modelC))
			{
				string realC = CalculateAccuracyC(compoundName,C.ToString());
				return realC;
			}
			return "<" + modelC;
		}


		/// <summary>
		/// 计算目标化合物浓度
		/// </summary>
		/// <param name="compoundName"></param>
		/// <param name="modelC"></param>
		/// <param name="sampleName"></param>
		/// <returns></returns>
		private string CompareCompoundWithFormula(string compoundName,string modelC,string sampleName)
		{
			//计算公式C = Ci×f×V1 / V * k
			//稀释倍数
			double f = double.NaN;
			//定容体积
			double V1 = double.NaN;
			//取样量
			double V = double.NaN;
			//系数k
			double k = double.NaN;
			if (Regex.IsMatch(dilutionratioTextBox.Text,"^([0-9]{1,}[.][0-9]*)$"))
			{
				f = double.Parse(dilutionratioTextBox.Text);
			}
			if (Regex.IsMatch(constantvolumeTextBox.Text,"^([0-9]{1,}[.][0-9]*)$"))
			{
				V1 = double.Parse(constantvolumeTextBox.Text);
			}
			if (Regex.IsMatch(samplingquantityTextBox.Text,"^([0-9]{1,}[.][0-9]*)$"))
			{
				V = double.Parse(samplingquantityTextBox.Text);
			}
			if (Regex.IsMatch(coefficientTextBox.Text,"^([0-9]{1,}[.][0-9]*)$"))
			{
				k = double.Parse(coefficientTextBox.Text);
			}

			//目标物上机测定浓度
			double Ci;
			foreach (DataTable dataTable in compoundsDataSet.Tables)
			{
				//找到该化合物对应的datatable
				if (dataTable.TableName == compoundName)
				{
					for (int i = 0; i < dataTable.Rows.Count; i++)
					{
						for (int j = 0; j < dataTable.Columns.Count; j++)
						{
							//找到该化合物对应的样品编号和浓度数据
							string dtsampleName = dataTable.Rows[i][j].ToString();
							if (dtsampleName == sampleName)
							{
								string potency = dataTable.Rows[i]["浓度"].ToString();
								if (!potency.Contains("-"))
								{
									Ci = double.Parse(potency);
									//公式计算
									//先用写死的，然后之后学习反射
									double C = double.NaN;
									double moleculeV1 = double.Parse((constantvolumeComboBox.SelectedItem as ComboBoxItem).Tag.ToString());
									double denominatorV = double.Parse((samplingquantityComboBox.SelectedItem as ComboBoxItem).Tag.ToString());
									double taggetC = double.Parse((TargetCompanyComboBox.SelectedItem as ComboBoxItem).Tag.ToString());
									double ZDJCCi = double.Parse((ZDJCCompanyComboBox.SelectedItem as ComboBoxItem).Tag.ToString());
									//单位换算
									if (FormulaComboBox.Text.Contains("V1"))
									{
										if (samplingquantityComboBox.Text != constantvolumeComboBox.Text)
										{
											//系数加上分子除以分母
											k *= moleculeV1 / denominatorV;
											k *= taggetC / ZDJCCi;
										}
										C = Ci * f * V1 / V * k;
									}
									else
									{
										//判断C和Ci/V的单位是否一致
										//比如C为mg/L
										//Ci为mg
										//V为mL
										//后台系数k就应该L/mL = 1000
										string CCi = TargetCompanyComboBox.Text.Split("/")[0]; 
										string CV = TargetCompanyComboBox.Text.Split("/")[1];
				
										foreach (ComboBoxItem comboBoxItem in samplingquantityComboBox.Items)
										{
											if (comboBoxItem.Content.ToString() == CV)
											{
												k *= double.Parse(comboBoxItem.Tag.ToString()) / denominatorV;
												k *= taggetC / ZDJCCi;
											}
										}
										C = Ci * f / V * k;
		
									}
									if (C > double.Parse(modelC))
									{
										string realC = CalculateAccuracyC(compoundName,C.ToString());
										return realC;
									}
								}
							}
						}
					}
				}
			}

			if (testZDRadioButton.IsChecked == true)
			{
				return "<" + modelC;
			}
			else
			{
				return "ND";
			}
		}

		private HSSFCellStyle CreateStyle(HSSFWorkbook workbook)
		{
			HSSFCellStyle cellStyle = (HSSFCellStyle)workbook.CreateCellStyle(); //创建列头单元格实例样式
			cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center; //水平居中
			cellStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center; //垂直居中
			cellStyle.WrapText = true;//自动换行
			//cellStyle.BorderBottom = BorderStyle.Thin;
			//cellStyle.BorderRight = BorderStyle.Thin;
			//cellStyle.BorderTop = BorderStyle.Thin;
			//cellStyle.BorderLeft = BorderStyle.Thin;
			cellStyle.TopBorderColor = HSSFColor.Black.Index;//DarkGreen(黑绿色)
			cellStyle.RightBorderColor = HSSFColor.Black.Index;
			cellStyle.BottomBorderColor = HSSFColor.Black.Index;
			cellStyle.LeftBorderColor = HSSFColor.Black.Index;

			return cellStyle;
		}
		private void ComplierCode(string expression)
		{
			CSharpCodeProvider objCSharpCodePrivoder = new CSharpCodeProvider();

			CompilerParameters objCompilerParameters = new CompilerParameters();

			//添加需要引用的dll
			objCompilerParameters.ReferencedAssemblies.Add("System.dll");
			objCompilerParameters.ReferencedAssemblies.Add("System.Windows.Forms.dll");
			//是否生成可执行文件
			objCompilerParameters.GenerateExecutable = false;
			//是否生成在内存中
			objCompilerParameters.GenerateInMemory = true;

			//编译代码
			CompilerResults cr = objCSharpCodePrivoder.CompileAssemblyFromSource(objCompilerParameters,FormulaComboBox.Text);

			if (cr.Errors.HasErrors)
			{
				var msg = string.Join(Environment.NewLine,cr.Errors.Cast<CompilerError>().Select(err => err.ErrorText));
				MessageBox.Show(msg,"编译错误");
			}
			else
			{
				Assembly objAssembly = cr.CompiledAssembly;
				object objHelloWorld = objAssembly.CreateInstance("Test");
				MethodInfo objMI = objHelloWorld.GetType().GetMethod("Hello");
				objMI.Invoke(objHelloWorld,null);
			}
		}

		/// <summary>
		/// 计算精度
		/// </summary>
		/// <param name="compoundName"></param>
		/// <param name="C"></param>
		/// <returns></returns>
		private string CalculateAccuracyC(string compoundName,string C)
		{
			double answer = double.NaN;
			string accuracy = AccuracyComboBox.SelectedItem.ToString();
			int num = 0;
			//选择默认方式
			if (accuracy == AccuracyComboBox.Items[0].ToString())
			{
				foreach (KeyValuePair<string,string> keyValuePair in compoundsNameList)
				{

					if (keyValuePair.Key == compoundName)
					{
						string[] beforeValue = keyValuePair.Value.Split(".");
						//没有小数点的
						if (beforeValue.Length < 2)
						{
							answer = Math.Round(double.Parse(C),0);
						}
						else
						{
							answer = Math.Round(double.Parse(C),beforeValue[beforeValue.Length - 1].Length);
						}
						string[] afterValue = answer.ToString().Trim().Split(".");
						num = beforeValue[beforeValue.Length - 1].Length - afterValue[afterValue.Length - 1].Length;
					}
				}
			}
			//选择其他位数
			else
			{
				string[] beforeValue = accuracy.Split(":");

				answer = Math.Round(double.Parse(C),int.Parse(beforeValue[beforeValue.Length - 1]));
				string[] afterValue = answer.ToString().Trim().Split(".");
				num = int.Parse(beforeValue[beforeValue.Length - 1]) - afterValue[afterValue.Length - 1].Length;

			}
			//计算后补零
			if (num != 0)
			{
				string newanswer = answer.ToString();
				for (int i = 0; i < num; i++)
				{
					newanswer += "0";
				}
				return newanswer;
			}
			return answer.ToString().Trim();
		}


		/// <summary>
		/// 导出模板按钮
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void importAll_Click(object sender,RoutedEventArgs e)
		{
			if (finalsampleNameList.Count == 0 || newsampleNameList.Count == 0)
			{
				return;
			}
			System.Windows.Forms.SaveFileDialog sfd = new System.Windows.Forms.SaveFileDialog();
			sfd.Filter = "文本文件(*.txt)|*.txt|所有文件|*.*";//设置文件类型
			//sfd.FileName = "保存";//设置默认文件名
			sfd.DefaultExt = "txt";//设置默认格式（可以不设）
			sfd.AddExtension = true;//设置自动在文件名中添加扩展名
			sfd.ShowDialog();
			if (sfd.FileName != string.Empty)
			{
				string fullpath = sfd.FileName;
				using (FileStream stream = new FileStream(fullpath,FileMode.Create,FileAccess.ReadWrite))
				{
					StreamWriter streamWriter = new StreamWriter(stream);
					//streamWriter.WriteLine(strReportNoLabel.Content + ReportNo);
					streamWriter.WriteLine(samplingquantityLabel.Content + samplingquantityTextBox.Text + samplingquantityComboBox.Text);
					streamWriter.WriteLine(dilutionratioLabel.Content + dilutionratioTextBox.Text);
					streamWriter.WriteLine(constantvolumeLabel.Content + constantvolumeTextBox.Text + constantvolumeComboBox.Text);
					streamWriter.WriteLine(coefficientLabel.Content + coefficientTextBox.Text);
					streamWriter.WriteLine(TargetCompanyLabel.Content + TargetCompanyComboBox.Text);
					streamWriter.WriteLine(AccuracyLabel.Content + AccuracyComboBox.Text);
					streamWriter.WriteLine(FormulaLabel.Content + FormulaComboBox.Text);
					if (testZDRadioButton.IsChecked == true)
					{
						streamWriter.WriteLine(testZDRadioButton.Content + "：" + ZDJCCompanyComboBox.Text);
					}
					else if (testJCRadioButton.IsChecked == true)
					{
						streamWriter.WriteLine(testJCRadioButton.Content + "：" + ZDJCCompanyComboBox.Text);
					}
					foreach (KeyValuePair<string,string> keyValuePair in compoundsNameList)
					{
						streamWriter.WriteLine(keyValuePair.Key + "：" + keyValuePair.Value);
					}
					streamWriter.Flush();
					stream.Flush();
				}
				Process process = new Process();
				ProcessStartInfo processStartInfo = new ProcessStartInfo(fullpath);
				processStartInfo.UseShellExecute = true;
				process.StartInfo = processStartInfo;
				process.Start();
			}
		}

		/// <summary>
		/// 生成compoundsNameList
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void importExcel_MouseMove(object sender,MouseEventArgs e)
		{
			if (finalsampleNameList.Count == 0 || newsampleNameList.Count == 0)
			{
				return;
			}
			AddDetectionLimit();
		}

		private void TextBox_TextChanged(object sender,TextChangedEventArgs e)
		{
			TextBox textbox = sender as TextBox;
			string text = textbox.Text.Trim();
			if (text != null && text != "")
			{
				Regex r = new Regex(@"^([0-9\.]*)$");
				if (r.IsMatch(textbox.Text.Trim()) == false)
				{
					textbox.Text = textbox.Text.Remove(textbox.Text.Length - 1,1);
				}

				int numDecimal = 0;
				for (int i = 0; i < textbox.Text.Length; i++)
				{
					if (textbox.Text[i].ToString() == ".")
					{
						numDecimal++;
					}
				}

				if (numDecimal > 1)
				{
					textbox.Text = textbox.Text.Remove(textbox.Text.Length - 1,1);
				}
			}
			textbox.SelectionStart = textbox.Text.Length;
		}
		private static string lastText = string.Empty;
		private void TextBox_KeyUp(object sender,KeyEventArgs e)
		{
			TextBox textBox = sender as TextBox;
			//判断按键是不是要输入的类型。
			if (textBox.Text != "" && textBox.Text != string.Empty && textBox.Text != lastText)
			{
				if (e.Key == Key.Decimal)
				{
					
				}
				//0-9
				else if (((int)e.Key < 34 || (int)e.Key > 43) && ((int)e.Key < 74 || (int)e.Key > 83))
				{
					textBox.Text = textBox.Text.Remove(textBox.Text.Length - 1,1);
				}
			}

			textBox.SelectionStart = textBox.Text.Length;
			lastText = textBox.Text;
			e.Handled = true;
		}

		/// <summary>
		/// 搜索
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void searchTextBox_TextChanged(object sender,RoutedEventArgs e)
		{
			string searchText = searchTextBox.Text;
			TabControl tabControl = GetVisualChild<TabControl>(maingrid);
			if (tabControl != null)
			{
				foreach (TabItem tabItem in tabControl.Items)
				{
					if (tabItem.IsSelected)
					{
						string header = tabItem.Header.ToString();
						DataGrid dataGrid = tabItem.Content as DataGrid;
						if (searchText != null && searchText != "")
						{
							for (int i = 0; i < dataGrid.ItemContainerGenerator.Items.Count - 1; i++)
							{
								dataGrid.ScrollIntoView(dataGrid.Items[i]);
								DataGridRow dgv = (DataGridRow)dataGrid.ItemContainerGenerator.ContainerFromIndex(i);
								if (dgv == null)
								{
									dataGrid.UpdateLayout();
									dataGrid.ScrollIntoView(dataGrid.Items[i]);
									dgv = (DataGridRow)dataGrid.ItemContainerGenerator.ContainerFromIndex(i);
								}
								bool showdgv = false;
								DataRow dr = (dgv.Item as DataRowView).Row;
								for (int j = 0; j < dr.ItemArray.Length; j++)
								{
									dgv.UpdateLayout();
									DataGridCellsPresenter presenter = GetVisualChild<DataGridCellsPresenter>(dgv);
									DataGridCell cell = (DataGridCell)presenter.ItemContainerGenerator.ContainerFromIndex(j);
									string cellcontent = dr[j].ToString().Trim();
									if (cellcontent.ToLower().Contains(searchText.ToLower()))
									{
										cell.Background = new SolidColorBrush(Colors.Orange);
										showdgv = true;
									}
									else
									{
										cell.Background = null;
									}
								}
								if (showdgv)
								{
									dgv.Visibility = Visibility.Visible;
								}
								else
								{
									dgv.Visibility = Visibility.Collapsed;
								}
							}
						}
						else
						{
							for (int i = 0; i < dataGrid.ItemContainerGenerator.Items.Count - 1; i++)
							{
								DataGridRow dgv = (DataGridRow)dataGrid.ItemContainerGenerator.ContainerFromIndex(i);
								if (dgv == null)
								{
									dataGrid.UpdateLayout();
									dataGrid.ScrollIntoView(dataGrid.Items[i]);
									dgv = (DataGridRow)dataGrid.ItemContainerGenerator.ContainerFromIndex(i);
								}
								dgv.Visibility = Visibility.Visible;
								DataRow dr = (dgv.Item as DataRowView).Row;
								for (int j = 0; j < dr.ItemArray.Length; j++)
								{
									dgv.UpdateLayout();
									DataGridCellsPresenter presenter = GetVisualChild<DataGridCellsPresenter>(dgv);
									DataGridCell cell = (DataGridCell)presenter.ItemContainerGenerator.ContainerFromIndex(j);
									cell.Background = null;
								}
							}
						}
					}
				}
			}
		}

		/// <summary>
		/// 拖动进入
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void scDragEnter(object sender,DragEventArgs e)
		{
			if (e.Data.GetDataPresent(DataFormats.FileDrop))
			{
				e.Effects = DragDropEffects.Link;
			}
			else
			{
				e.Effects = DragDropEffects.None;
			}
		}

		#region 辅助函数

		/// <summary>
		/// 获取父可视对象中第一个指定类型的子可视对象
		/// </summary>
		/// <typeparam name="T">可视对象类型</typeparam>
		/// <param name="parent">父可视对象</param>
		/// <returns>第一个指定类型的子可视对象</returns>
		public static T GetVisualChild<T>(Visual parent) where T : Visual
		{
			T child = default(T);
			int numVisuals = VisualTreeHelper.GetChildrenCount(parent);
			for (int i = 0; i < numVisuals; i++)
			{
				Visual v = (Visual)VisualTreeHelper.GetChild(parent,i);
				child = v as T;
				if (child == null)
				{
					child = GetVisualChild<T>(v);
				}
				if (child != null)
				{
					break;
				}
			}
			return child;
		}

		/// <summary>
		/// 父控件+控件名找到子控件
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="obj"></param>
		/// <param name="name"></param>
		/// <returns></returns>
		public T GetChildObject<T>(DependencyObject obj,string name) where T : FrameworkElement
		{
			DependencyObject child = null;
			T grandChild = null;
			for (int i = 0; i <= VisualTreeHelper.GetChildrenCount(obj) - 1; i++)
			{
				child = VisualTreeHelper.GetChild(obj,i);
				if (child is T && (((T)child).Name == name || string.IsNullOrEmpty(name)))
				{
					return (T)child;
				}
				else
				{
					grandChild = GetChildObject<T>(child,name);
					if (grandChild != null)
						return grandChild;
				}
			}
			return null;
		}


		#endregion

	}
}
