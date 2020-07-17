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
using Microsoft.CSharp;
using System.CodeDom.Compiler;
using System.Reflection;

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
			/*textblock.Text = "该图片原名为: " + nameStr + " ,请在下方填写你想要的名字";
			textbox.Focus();*/
            this.txtExpression.Focus();
        }

        private void comfirmButton_Click(object sender, RoutedEventArgs e)
		{
			/*Window window = Application.Current.MainWindow;
			MainWindow mainwindow = window as MainWindow;
			mainwindow.renameStr = textbox.Text;
			this.Close();*/
		}

		private void cancelButton_Click(object sender, RoutedEventArgs e)
		{
			this.Close();
		}

        private void btnCalculate_Click(object sender,RoutedEventArgs e)
        {
            try
            {
                string expression = this.txtExpression.Text.Trim();
                this.txtResult.Text = this.ComplierCode(expression).ToString();
            }
            catch (Exception ex)
            {
                this.txtResult.Text = ex.Message;
            }
        }

        private object ComplierCode(string expression)
        {
            string code = WrapExpression(expression);

            CSharpCodeProvider csharpCodeProvider = new CSharpCodeProvider();

            //编译的参数
            CompilerParameters compilerParameters = new CompilerParameters();
            //compilerParameters.ReferencedAssemblies.AddRange();
            compilerParameters.CompilerOptions = "/t:library";
            compilerParameters.GenerateInMemory = true;
            //开始编译
            CompilerResults compilerResults = csharpCodeProvider.CompileAssemblyFromSource(compilerParameters,code);
            if (compilerResults.Errors.Count > 0)
                throw new Exception("编译出错！");

            Assembly assembly = compilerResults.CompiledAssembly;
            Type type = assembly.GetType("ExpressionCalculate");
            MethodInfo method = type.GetMethod("Calculate");
            return method.Invoke(null,null);
        }

        private string WrapExpression(string expression)
        {
            string code = @"
                using System;

                class ExpressionCalculate
                {
                    public static DateTime start_dt = Convert.ToDateTime(""{start_dt}"");
                    public static DateTime end_dt = Convert.ToDateTime(""{end_dt}"");
                    public static DateTime current_dt = DateTime.Now;

                    public static object Calculate()
                    {
                        return {0};
                    }
                }
            ";

            return code.Replace("{0}",expression);
        }
    }
}
