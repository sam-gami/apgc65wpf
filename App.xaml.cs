using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Diagnostics;
using AtgPubCs;

namespace apgc65wpf
{
  /// <summary>
  /// App.xaml の相互作用ロジック
  /// </summary>
  public partial class App : Application
  {
    [STAThread]
    public static void Main()
    {
      Process pro = null;
      Process pros = null;
      int i = 0;
      string para = null;
      pro = Process.GetCurrentProcess();

      if (pro.ProcessName.IndexOf(".") < 0)
      {
        para = pro.ProcessName + ".exe";
      }
      else
      {
        para = z.Left(pro.ProcessName, pro.ProcessName.IndexOf(".")) + ".exe";
      }
      string clfl = @"C:\RPP\CLIENT.FIL";
      if (System.IO.File.Exists(clfl))
      {
        App app = new App();
        app.StartupUri = new Uri("MainWindow.xaml", UriKind.Relative);
        app.InitializeComponent();
        app.Run();
      }
      else
      {
        System.Diagnostics.Process.Start("AtgLogin.exe", para);
      }
    }
  }
}
