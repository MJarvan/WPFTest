﻿using System;
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

		private void CreateTxt(string path)
		{
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
					CreateDataTable(tabControl,vs);
					KeyValuePair<string,string> keyValuePair = new KeyValuePair<string,string>("以下空白","0");
					compoundsNameList.Add(keyValuePair);

				}
				maingrid.Children.Add(tabControl);
			}
		}

		private void CreateDataTable(TabControl tabControl,List<string> vs)
		{
			//前三列是datatable的属性名和表头需要单独保存
			//string[] id = vs[0].Split("\t");
			string[] name = vs[1].Split("\t");
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
				//List<string> data = vs[j].Split("\t").ToList();
				DataRow dr = datatable.NewRow();
				dr.ItemArray = vs[j].Split("\t");
				datatable.Rows.Add(dr);
			}
			//删除后缀.gcd
			for (int k = 0; k < datatable.Rows.Count; k++)
			{
				string oldname = datatable.Rows[k]["数据文件名"].ToString();
				string newname = oldname.Replace(".gcd","");
				if (sampleNameList.Count != datatable.Rows.Count)
				{
					sampleNameList.Add(newname);
				}
				datatable.Rows[k]["数据文件名"] = newname;
			}
			//根据有机组要求只要三列
			DataTable newdatatable = new DataTable();
			newdatatable.TableName = datatable.TableName;
			newdatatable.Columns.Add("编号");
			newdatatable.Columns.Add("数据文件名");
			newdatatable.Columns.Add("浓度");
			//for (int l = 0; l < datatable.Columns.Count; l++)
			//{
			//	if (datatable.Columns[l].ColumnName == "编号" || datatable.Columns[l].ColumnName == "数据文件名" || datatable.Columns[l].ColumnName == "浓度")
			//	{
			//		DataColumn dataColumn = datatable.Columns[l];
			//		newdatatable.Columns.Add(dataColumn.ColumnName,dataColumn.DataType);
			//	}
			//}
			for (int i = 0; i < datatable.Rows.Count; i++)
			{
				DataRow dr = newdatatable.NewRow();
				for (int j = 0; j < datatable.Columns.Count; j++)
				{
					string ColumnName = datatable.Columns[j].ColumnName;
					for (int k = 0; k < newdatatable.Columns.Count; k++)
					{
						string newColumnName = newdatatable.Columns[k].ColumnName;

						if (newColumnName == ColumnName)
						{
							dr[ColumnName] = datatable.Rows[i][j];
						}
					}
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
			
			List<string> cellList = sampleNameList.Take(7).ToList();
			int num = 0;
			string banlance = "Dup";
			if (cellList.Exists(x => x.Contains(banlance)))
			{
				cellList.RemoveAt(cellList.Count - 1);
				num = cellList.IndexOf(cellList.Where(x => x.Contains(banlance)).FirstOrDefault()) + 1;
				cellList.Insert(num,cellList.Where(x => x.Contains(banlance)).FirstOrDefault().Replace(banlance,"平均"));
			}
			//前四行
			for (int i = 0; i < 4; i++)
			{
				//第一个单元格 公式
				HSSFRow row = (HSSFRow)sheet.CreateRow(i); //创建行
				if (i == 0)
				{
					CellRangeAddress region = new CellRangeAddress(0,3,0,0);
					sheet.AddMergedRegion(region);
					var formulacell = row.CreateCell(0);
					StringBuilder stringBuilder = new StringBuilder("计算公式：C = Ci×f×V1 / V\n");
					stringBuilder.Append("式中：C——样品中目标物浓度\n");
					stringBuilder.Append("Ci——目标物上机测定浓度\n");
					stringBuilder.Append("V1——定容体积\n");
					stringBuilder.Append("f——稀释倍数\n");
					stringBuilder.Append("V——取样量\n");
					formulacell.SetCellValue(stringBuilder.ToString());//合并单元格后，只需对第一个位置赋值即可（TODO:顶部标题）
					formulacell.CellStyle = cellStyle;
				}
				else
				{
					//从第二列开始
					for (int j = 1; j < 10; j++)
					{
						var cell = row.CreateCell(j);
						if (j == 9)
						{
							string setvalue = (i == 1) ? "" : "以下空白";
							cell.SetCellValue(setvalue);
						}
						else if (i == (int)samplingquantityLabel.Tag)
						{
							if (j == num)
							{
								cell.SetCellValue("-");
								break;
							}
							string setvalue = (j == 1) ? samplingquantityTextBox.Text : samplingquantityLabel.Name;
							cell.SetCellValue(setvalue);
						}
						else if (i == (int)constantvolumeLabel.Tag)
						{
							if (j == num)
							{
								cell.SetCellValue("-");
								break;
							}
							string setvalue = (j == 1) ? constantvolumeTextBox.Text : constantvolumeLabel.Name;
							cell.SetCellValue(setvalue);
						}
						else if (i == (int)dilutionratioLabel.Tag)
						{
							if (j == num)
							{
								cell.SetCellValue("-");
								break;
							}
							string setvalue = (j == 1) ? dilutionratioTextBox.Text : dilutionratioLabel.Name;
							cell.SetCellValue(setvalue);
						}
						else if (j == 1)
						{
							cell.SetCellValue("样品编号");
						}
						else
						{
							cell.SetCellValue(cellList[j - 2]);
						}
						cell.CellStyle = cellStyle;
					}
				}
			}
			//第四行表头
			/*HSSFRow HearderRow = (HSSFRow)sheet.CreateRow(4); //创建行
			var firstheaderCell = HearderRow.CreateCell(0);
			firstheaderCell.SetCellValue("目标化合物");//合并单元格后，只需对第一个位置赋值即可（TODO:顶部标题）
			firstheaderCell.CellStyle = cellStyle;
			var secondheaderCell = HearderRow.CreateCell(1);
			secondheaderCell.SetCellValue("最低检测质量浓度(mg/L)");
			secondheaderCell.CellStyle = cellStyle;
			var thirdheaderCell = HearderRow.CreateCell(2);
			thirdheaderCell.SetCellValue("目标化合物浓度 C (mg/L)");
			thirdheaderCell.CellStyle = cellStyle;
			CellRangeAddress newregion = new CellRangeAddress(4,4,2,9);
			sheet.AddMergedRegion(newregion);


			for (int k = 0; k < compoundsNameList.Count; k++)
			{
				HSSFRow compoundsRow = (HSSFRow)sheet.CreateRow(4 + k + 1); //创建行
				var headerCell = compoundsRow.CreateCell(0);
				string compoundName = compoundsNameList[k].Key;
				headerCell.SetCellValue(compoundName);
				float modelC = float.Parse(compoundsNameList[k].Value);
				for (int l = 1; l < 9; l++)
				{
					var compoundsCell = compoundsRow.CreateCell(l);
					//第一列是判断标准
					if (l == 1)
					{
						compoundsCell.SetCellValue(modelC);
					}
					else if (l != 8)
					{
						string sampleName = cellList[l - 1];
						string setvalue = CompareCompoundWithFormula(compoundName,modelC,sampleName);
						compoundsCell.SetCellValue(setvalue);
					}
					else
					{
						compoundsCell.SetCellValue(string.Empty);
					}
				}
			}*/

			//导出Excel
			for (int i = 1; i < 10; i++)
			{
				sheet.AutoSizeColumn(i);
			}
				
			System.Windows.Forms.FolderBrowserDialog fbd = new System.Windows.Forms.FolderBrowserDialog();
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

		private string CompareCompoundWithFormula(string compoundName,float modelC,string sampleName)
		{
			//计算公式C = Ci×f×V1 / V
			//稀释倍数
			float f = float.Parse(dilutionratioTextBox.Text);
			//定容体积
			float V1 = float.Parse(constantvolumeTextBox.Text);
			//取样量
			float V = float.Parse(samplingquantityTextBox.Text);
			//目标物上机测定浓度
			float Ci;
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
									Ci = float.Parse(potency);
									//公式计算
									float C = Ci * f * V1 / V;
									if (C < modelC)
									{
										return "<" + modelC;
									}
									else
									{
										return potency;
									}
								}
							}
						}
					}
				}
			}
			return string.Empty;
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