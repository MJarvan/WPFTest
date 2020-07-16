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

		/// <summary>
		/// 样品名称合计
		/// </summary>
		List<string> sampleNameList = new List<string>();

		/// <summary>
		/// 添加了平行样之后的合计
		/// </summary>
		List<KeyValuePair<List<string>,int>> newsampleNameList = new List<KeyValuePair<List<string>,int>>();

		/// <summary>
		/// 委托单号
		/// </summary>
		string ReportNo = string.Empty;

		//调整一个表格的样品总列数
		int importTakeNum = 8;
		//调整一个表格的总列数
		int sheetColumnCount = 10;
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
			scrollviewer.DragEnter += scDragEnter;
			scrollviewer.Drop += scDrop;
			samplingquantityLabel.Tag = 0;
			dilutionratioLabel.Tag = 1;
			constantvolumeLabel.Tag = 2;
		}

		/// <summary>
		/// 选择路径
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void choiceButton_Click(object sender,RoutedEventArgs e)
		{
			maingrid.Children.Clear();
			System.Windows.Forms.OpenFileDialog ofd = new System.Windows.Forms.OpenFileDialog();
			ofd.ShowDialog();
			if (ofd.FileName != string.Empty)
			{
				CreateTxt(ofd.FileName);
			}
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

			if (e.Data.GetDataPresent(DataFormats.FileDrop))
			{
				e.Effects = DragDropEffects.Link;

				string[] paths = (string[])e.Data.GetData(DataFormats.FileDrop);
				foreach (string path in paths)
				{
					CreateTxt(path);
				}
			}
			e.Handled = true;
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
				bool isOK = false;

				foreach (List<string> vs in txtList)
				{
					isOK = CreateDataTable(tabControl,vs);
					if (!isOK)
					{
						return;
					}
				}
				AddParallelSamplesToList();
				KeyValuePair<string,string> keyValuePair = new KeyValuePair<string,string>("以下空白",string.Empty);
				compoundsNameList.Add(keyValuePair);
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
			ReportNo = string.Empty;
			compoundsDataSet.Tables.Clear();
			maingrid.Children.Clear();
		}

		private bool CreateDataTable(TabControl tabControl,List<string> vs)
		{
			//前三列是datatable的属性名和表头需要单独保存
			//string[] id = vs[0].Split("\t");
			sampleNameList.Clear();
			string[] name = vs[1].Split("\t");

			//缺少自我填写的最低检测质量浓度
			if (name.Length < 3)
			{
				MessageBox.Show(name[1] + "缺少最低检测质量浓度，请补充填写！");
				return false;
			}
			KeyValuePair<string,string> keyValuePair = new KeyValuePair<string,string>(name[1],name[2]);
			compoundsNameList.Add(keyValuePair);

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
			//提取委托单号,删除后缀.gcd
			for (int k = 0; k < datatable.Rows.Count; k++)
			{
				string oldname = datatable.Rows[k]["数据文件名"].ToString();
				string[] newname = oldname.Replace(".gcd","").Split(" ");

				if (sampleNameList.Count != datatable.Rows.Count)
				{
					if (newname.Length > 1)
					{
						ReportNo = newname[0];
						sampleNameList.Add(newname[1]);
						datatable.Rows[k]["数据文件名"] = newname[1];
					}
					else
					{
						datatable.Rows[k]["数据文件名"] = newname[0];
						sampleNameList.Add(newname[0]);
					}
				}
			}
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
			tabItem.Header = name[1] + " | " + name[2];
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
		/// 添加平行样
		/// </summary>
		private void AddParallelSamplesToList()
		{
			//先加完所有的平行样再分组比较好
			List<string> psNameList = sampleNameList.ToList();
			List<int> numList = new List<int>();
			string banlance = "Dup";
			List<string> addNameList = psNameList.Where(x => x.Contains(banlance)).ToList();
			for (int i = 0; i < addNameList.Count; i++)
			{
				int num = psNameList.IndexOf(addNameList[i]) + 1;
				numList.Add(num);
			}
			for (int j = 0; j < numList.Count; j++)
			{
				if (j == 0)
				{
					psNameList.Insert(numList[j],psNameList.Where(x => x.Contains(banlance)).FirstOrDefault().Replace(banlance,"平均"));
				}
				else
				{
					psNameList.Insert(numList[j] + j,psNameList.Where(x => x.Contains(banlance)).FirstOrDefault().Replace(banlance,"平均"));

				}
			}

			int Count = psNameList.Count % importTakeNum > 0 ? psNameList.Count / importTakeNum + 1 : psNameList.Count / importTakeNum;

			for (int i = 0; i < Count; i++)
			{
				if (i == Count - 1)
				{
					List<string> cellList = psNameList.ToList();
					int num = 0;
					if (cellList.Exists(x => x.Contains(banlance)))
					{
						num = cellList.IndexOf(cellList.Where(x => x.Contains(banlance)).FirstOrDefault()) + 1;
					}
					KeyValuePair<List<string>,int> keyValuePair = new KeyValuePair<List<string>,int>(cellList,num);
					newsampleNameList.Add(keyValuePair);
				}
				else
				{
					List<string> cellList = psNameList.Take(importTakeNum).ToList();
					psNameList.RemoveRange(0,importTakeNum);
					int num = 0;
					if (cellList.Exists(x => x.Contains(banlance)))
					{
						num = cellList.IndexOf(cellList.Where(x => x.Contains(banlance)).FirstOrDefault()) + 1;
					}
					KeyValuePair<List<string>,int> keyValuePair = new KeyValuePair<List<string>,int>(cellList,num);
					newsampleNameList.Add(keyValuePair);
				}
			}
		}


		/// <summary>
		/// 导出生成Excel
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void importAll_Click(object sender,RoutedEventArgs e)
		{
			if (compoundsNameList.Count == 0)
			{
				return;
			}
			var workbook = new HSSFWorkbook();
			var sheet = workbook.CreateSheet("色谱分析结果汇总1-水");
			sheet.ForceFormulaRecalculation = true;
			//设置顶部大标题样式
			HSSFCellStyle cellStyle = CreateStyle(workbook);

			int Count = 0;
			foreach (KeyValuePair<List<string>,int> keyValuePair in newsampleNameList)
			{
				CreateSheet(sheet,cellStyle,keyValuePair.Key,keyValuePair.Value,Count);
				Count++;
			}

			//自动调整列距
			for (int i = 0; i < sheetColumnCount * Count; i++)
			{
				if (i % sheetColumnCount == 0)
				{
					sheet.SetColumnWidth(i,30 * 256);
				}
				else
				{
					sheet.AutoSizeColumn(i);
				}
			}

			ExportToExcel(workbook);
		}

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
				string path = @"D:\CreateExcel\" + ReportNo + @"\";
				//创建用户临时图片文件夹或者清空临时文件夹所有文件
				if (!Directory.Exists(path))
				{
					Directory.CreateDirectory(path);
				}
				string filename = workbook.GetSheetAt(0).SheetName + ".xls";
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

		private void CreateSheet(ISheet sheet,HSSFCellStyle cellStyle,List<string> cellList,int advantageNum,int Count)
		{
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
				var formulacell = (Count == 0) ? row.CreateCell(i) : row.CreateCell(sheetColumnCount * Count);
				formulacell.CellStyle = cellStyle;
				if (i == 0)
				{
					CellRangeAddress region = new CellRangeAddress(0,3,sheetColumnCount * Count,sheetColumnCount * Count);
					sheet.AddMergedRegion(region);
					StringBuilder stringBuilder = new StringBuilder("计算公式：C = Ci×f×V1 / V\n");
					stringBuilder.Append("式中：C——样品中目标物浓度\n");
					stringBuilder.Append("Ci——目标物上机测定浓度\n");
					stringBuilder.Append("V1——定容体积\n");
					stringBuilder.Append("f——稀释倍数\n");
					stringBuilder.Append("V——取样量\n");
					formulacell.SetCellValue(stringBuilder.ToString());//合并单元格后，只需对第一个位置赋值即可（TODO:顶部标题）
				}
				//从第二列开始
				for (int j = sheetColumnCount * Count + 1; j < sheetColumnCount * Count + sheetColumnCount; j++)
				{
					var cell = row.CreateCell(j);
					cell.CellStyle = cellStyle;
					//按照已有的弄
					if ((j - sheetColumnCount * Count) - 2 >= cellList.Count)
					{
						string setvalue = (i == 0) ? "以下空白" : "";
						cell.SetCellValue(setvalue);
					}
					else if (i == (int)samplingquantityLabel.Tag)
					{
						if ((j - sheetColumnCount * Count) - 2 == advantageNum)
						{
							cell.SetCellValue("-");
							continue;
						}
						string setvalue = ((j - sheetColumnCount * Count) == 1) ? samplingquantityLabel.Content.ToString() : samplingquantityTextBox.Text;
						cell.SetCellValue(setvalue);
					}
					else if (i == (int)constantvolumeLabel.Tag)
					{
						if ((j - sheetColumnCount * Count) - 2 == advantageNum)
						{
							cell.SetCellValue("-");
							continue;
						}
						string setvalue = ((j - sheetColumnCount * Count) == 1) ? constantvolumeLabel.Content.ToString() : constantvolumeTextBox.Text;
						cell.SetCellValue(setvalue);
					}
					else if (i == (int)dilutionratioLabel.Tag)
					{
						if ((j - sheetColumnCount * Count) - 2 == advantageNum)
						{
							cell.SetCellValue("-");
							continue;
						}
						string setvalue = ((j - sheetColumnCount * Count) == 1) ? dilutionratioLabel.Content.ToString() : dilutionratioTextBox.Text;
						cell.SetCellValue(setvalue);
					}
					else if ((j - sheetColumnCount * Count) == 1)
					{
						cell.SetCellValue("样品编号");
					}
					else if((j - sheetColumnCount * Count) - 2 < cellList.Count)
					{
						cell.SetCellValue(cellList[(j - sheetColumnCount * Count) - 2]);
					}
				}
			}
			//第四行表头
			HSSFRow HearderRow = (Count == 0) ? (HSSFRow)sheet.CreateRow(4) : (HSSFRow)sheet.GetRow(4); //创建行或者获取行
			HearderRow.HeightInPoints = 20;
			for (int k = sheetColumnCount * Count; k < sheetColumnCount * (Count + 1); k++)
			{
				var cell = HearderRow.CreateCell(k);
				cell.CellStyle = cellStyle;
				if (!sheet.MergedRegions.Exists(x => x.FirstRow == 4 && x.LastRow == 4 && x.FirstColumn == 2 + sheetColumnCount * Count && x.LastColumn == 9 + sheetColumnCount * Count))
				{
					CellRangeAddress newregion = new CellRangeAddress(4,4,2 + sheetColumnCount * Count,9 + sheetColumnCount * Count);
					sheet.AddMergedRegion(newregion);
				}
				if (k == 0 + sheetColumnCount * Count)
				{
					cell.SetCellValue("目标化合物");//合并单元格后，只需对第一个位置赋值即可（TODO:顶部标题）
				}
				else if (k == 1 + sheetColumnCount * Count)
				{
					cell.SetCellValue("最低检测质量浓度(mg/L)");
				}
				else if (k == 2 + sheetColumnCount * Count)
				{
					cell.SetCellValue("目标化合物浓度 C (mg/L)");
				}
			}

			for (int k = 0; k < compoundsNameList.Count; k++)
			{
				HSSFRow compoundsRow = (Count == 0) ? (HSSFRow)sheet.CreateRow(4 + k + 1) : (HSSFRow)sheet.GetRow(4 + k + 1); //创建行或者获取行
				compoundsRow.HeightInPoints = 20;
				//var headerCell = compoundsRow.CreateCell(0);
				string compoundName = compoundsNameList[k].Key;
				//headerCell.SetCellValue(compoundName);
				//headerCell.CellStyle = cellStyle;
				string modelC = compoundsNameList[k].Value;
				for (int l = sheetColumnCount * Count; l < sheetColumnCount * Count + sheetColumnCount; l++)
				{
					var compoundsCell = compoundsRow.CreateCell(l);
					compoundsCell.CellStyle = cellStyle;
					{
						//第一列是化合物名称
						if (l - sheetColumnCount * Count == 0)
						{
							compoundsCell.SetCellValue(compoundName);
						}
						//第二列是判断标准
						else if (l - sheetColumnCount * Count == 1)
						{
							compoundsCell.SetCellValue(modelC);
						}
						else if (l - sheetColumnCount * Count - 2 < cellList.Count)
						{
							if (modelC == string.Empty)
							{
								compoundsCell.SetCellValue(string.Empty);
							}
							else
							{
								string sampleName = cellList[l - sheetColumnCount * Count - 2];
								string setvalue = CompareCompoundWithFormula(compoundName,modelC,sampleName);
								compoundsCell.SetCellValue(setvalue);
							}
						}
					}
				}
			}
		}

		private HSSFCellStyle CreateStyle(HSSFWorkbook workbook)
		{
			HSSFCellStyle cellStyle = (HSSFCellStyle)workbook.CreateCellStyle(); //创建列头单元格实例样式
			cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center; //水平居中
			cellStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center; //垂直居中
			cellStyle.WrapText = true;//自动换行
			cellStyle.BorderBottom = BorderStyle.Thin;
			cellStyle.BorderRight = BorderStyle.Thin;
			cellStyle.BorderTop = BorderStyle.Thin;
			cellStyle.BorderLeft = BorderStyle.Thin;
			cellStyle.TopBorderColor = HSSFColor.Black.Index;//DarkGreen(黑绿色)
			cellStyle.RightBorderColor = HSSFColor.Black.Index;
			cellStyle.BottomBorderColor = HSSFColor.Black.Index;
			cellStyle.LeftBorderColor = HSSFColor.Black.Index;

			return cellStyle;
		}

		private string CompareCompoundWithFormula(string compoundName,string modelC,string sampleName)
		{
			//计算公式C = Ci×f×V1 / V
			//稀释倍数
			double f = double.Parse(dilutionratioTextBox.Text);
			//定容体积
			double V1 = double.Parse(constantvolumeTextBox.Text);
			//取样量
			double V = double.Parse(samplingquantityTextBox.Text);
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
									double C = Ci * f * V1 / V;
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
			return "<" + modelC;
		}

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
						answer = Math.Round(double.Parse(C),beforeValue[beforeValue.Length - 1].Length);
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
					newanswer = newanswer + "0";
				}
				return newanswer;
			}
			return answer.ToString().Trim();
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
