using System;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Text.RegularExpressions;
using AtgPubCs;
using System.Linq;
using C1.WPF.Excel;
using C1.WPF.FlexGrid;
using C1.WPF;
using System.Windows.Threading;
using System.Windows.Automation.Peers;
using System.Windows.Automation.Provider;
using System.Windows.Controls;


namespace apgc65wpf
{
  public static class ButtonExtensions//PerformClick用拡張
  {
    public static void PerformClick(this Button button)
    {
      if (button == null)
        throw new ArgumentNullException("button");

      var provider = new ButtonAutomationPeer(button) as IInvokeProvider;
      provider.Invoke();
    }
  }
  public partial class MainWindow : Window
  {
    public static string gsvid;
    public bool newmode = true;
    private const int COLUMN_MARGIN = 2;
    private const int ITEM_COUNT = 7;
    public OleDbConnection cn = new OleDbConnection();
    public static DataTable svcombodtb = new DataTable();
    public static DataTable oscombodtb = new DataTable();
    public DataTable osedcombodtb = new DataTable();
    public DataTable vdcombodtb = new DataTable();
    public DataTable nicloadorgdtb = new DataTable();
    public DataTable nicloadeditdtb = new DataTable();
    public DataTable nicvdcombodtb = new DataTable();
    public DataTable nicdtldtb = new DataTable();
    public DataTable cpuvdcombodtb = new DataTable();
    public static DataTable usetype1dtb = new DataTable();
    public DataTable usetype2dtb = new DataTable();
    public DataTable usetype3dtb = new DataTable();
    public DataTable editcpmstdtb = new DataTable();
    public DataTable editcpmstorgdtb = new DataTable();
    public DataTable placedtb = new DataTable();
    public DataTable vmsvcombodtb = new DataTable();
    public DataTable svallgriddtb = new DataTable();
    public DataTable hpvvmalldtb = new DataTable();
    public int nUserId;
    public string nUserName;

    public MainWindow()
    {
      InitializeComponent();
      nUserId = Clf.GetUserID();
      nUserName = Clf.GetUserName();

      // Enter キーでフォーカス移動する
      this.KeyDown += (sender, e) =>
      {
        if (e.Key != Key.Enter) { return; }
        var direction = Keyboard.Modifiers == ModifierKeys.Shift ? FocusNavigationDirection.Previous : FocusNavigationDirection.Next;
        (FocusManager.GetFocusedElement(this) as FrameworkElement)?.MoveFocus(new TraversalRequest(direction));
      };

      // textBox という TextBox では何も入力されていない場合は、フォーカス移動しない
      this.txtSvName.KeyDown += (sender, e) =>
      {
        if (e.Key != Key.Enter) { return; }
        if (!string.IsNullOrEmpty(((TextBox)sender).Text)) { return; }
        e.Handled = true;   // Hanbled を true にすると Window の KeyDown イベントは発生しない
      };
    }
    private void MainGrid_Initialized(object sender, EventArgs e)
    {
      MakeNicDtldtb();//←この分は、Load時に設定するとエラーになるため、Initializedに入れる。
    }
    private void MainGrid_Loaded(object sender, RoutedEventArgs e)//メイングリッドのロード時
    {
      //grNic.ColumnHeaders.Height = 60;

      string pn = Process.GetCurrentProcess().ProcessName;

      if (Process.GetProcessesByName(pn).GetUpperBound(0) > 0)
      {
        MessageBox.Show("二重起動はできません！");
        this.Close();
        return;
      } 


      setEditMode(1);
      newmode = true;
      txtUser.Content = nUserId + ":" + nUserName;
      txtMachine.Text = Environment.MachineName;
      txtYMD.Content = DateTime.Now.ToString("yyyy年MM月dd日(ddd)");
      setAllCombo();
      getActSV();
      getDelSV();
      getRealSVCount();
      getVMSVCount();
      Keyboard.Focus(txtSvName);
      //grnicinit();
      var ch = grNic.ColumnHeaders;
      ch.Rows[0].Height = 35;
      showsvallgrid();
      hpvvmallload();
    }
    //コントロールのイベント begin
    //***************************************************************************************************
    private void textBoxPrice_PreviewExecuted(object sender, ExecutedRoutedEventArgs e)
    {
      // 貼り付けを許可しない
      if (e.Command == ApplicationCommands.Paste)
      {
        e.Handled = true;
      }
    }
    private void Window_PreviewKeyDown(object sender, KeyEventArgs e) //ファンクションキー押下
    {
      string pushedKey = "";
      switch (e.Key)
      {
        case Key.F1:
        case Key.F2:
          var win = new Whlpsvid();
          win.ShowDialog();
          txtSvId.Text = gsvid;
          if (!string.IsNullOrEmpty(txtSvId.Text))
          {
            edcpdtload();
            setEditMode(2);
          }
          break;
        case Key.F3:
        case Key.F4:
        case Key.F5:
        case Key.F6:
        case Key.F7:
        case Key.F8:
        case Key.F9:
        case Key.F11:
        case Key.F12:
          pushedKey = e.Key.ToString();
          break;
        case Key.System:
          if (e.SystemKey == Key.F10)
          {
            pushedKey = "F10";
          }
          break;
      }
    }
    private void btCancel_Click(object sender, RoutedEventArgs e) //取り消しボタンクリック
    {
      if (MessageBox.Show("画面を初期化します。よろしいですか？", "画面の初期化", MessageBoxButton.YesNo,
     MessageBoxImage.Information) == MessageBoxResult.No)
      { return; }
      else
      {
        WindowInit();
      }
    }
    private void btSVIDHELP_Click(object sender, RoutedEventArgs e) //サーバーヘルプボタンクリック
    {
      var win = new Whlpsvid();
      win.ShowDialog();
      txtSvId.Text = gsvid;
      if (!string.IsNullOrEmpty(txtSvId.Text))
      {
        nicDataClear();
        edcpdtload();
        setEditMode(2);
      }
    }
    private void cmbOS_DropDownClosed(object sender, EventArgs e) //OSコンボ閉じたとき
    {
      int id = cmbOS.SelectedIndex;
      if (id == -1)
      {
        txtOSID.Text = "";
      }
      else
      {
        txtOSID.Text = cmbOS.SelectedValue.ToString();
      }

    }
    private void cmbED_DropDownClosed(object sender, EventArgs e) //エディションコンボ閉じたとき
    {
      int id = cmbED.SelectedIndex;
      if (id == -1)
      {
        txtEDID.Text = "";
      }
      else
      {
        txtEDID.Text = cmbED.SelectedValue.ToString();
      }
    }
    private void cmbVD_DropDownClosed(object sender, EventArgs e) //本体ベンダーコンボ閉じたとき
    {
      int id = cmbVD.SelectedIndex;
      if (id == -1)
      {
        txtVDID.Text = "";
      }
      else
      {
        txtVDID.Text = cmbVD.SelectedValue.ToString();
      }
    }
    private void cmbUSETYPE1_DropDownClosed(object sender, EventArgs e) //用途１コンボ閉じたとき
    {
      int id = cmbUSETYPE1.SelectedIndex;
      if (id == -1)
      {
        txtUSETYPE1ID.Text = "";
      }
      else
      {
        txtUSETYPE1ID.Text = cmbUSETYPE1.SelectedValue.ToString();
      }
    }
    private void cmbUSETYPE2_DropDownClosed(object sender, EventArgs e) //用途２コンボ閉じたとき
    {
      int id = cmbUSETYPE2.SelectedIndex;
      if (id == -1)
      {
        txtUSETYPE2ID.Text = "";
      }
      else
      {
        txtUSETYPE2ID.Text = cmbUSETYPE2.SelectedValue.ToString();
      }
    }
    private void cmbUSETYPE3_DropDownClosed(object sender, EventArgs e) //用途３コンボ閉じたとき
    {
      int id = cmbUSETYPE3.SelectedIndex;
      if (id == -1)
      {
        txtUSETYPE3ID.Text = "";
      }
      else
      {
        txtUSETYPE3ID.Text = cmbUSETYPE3.SelectedValue.ToString();
      }
    }
    private void cmbCPUVD_DropDownClosed(object sender, EventArgs e) //CPUベンダーコンボ閉じたとき
    {
      int id = cmbCPUVD.SelectedIndex;
      if (id == -1)
      {
        txtCPUVDID.Text = "";
      }
      else
      {
        txtCPUVDID.Text = cmbCPUVD.SelectedValue.ToString();
      }
    }
    private void cmbPlace_DropDownClosed(object sender, EventArgs e) //設置場所コンボ閉じたとき
    {
      int id = cmbPlace.SelectedIndex;
      if (id == -1)
      {
        txtPlaceID.Text = "";
      }
      else
      {
        txtPlaceID.Text = cmbPlace.SelectedValue.ToString();
      }
    }
    private void cmbHPVS_DropDownClosed(object sender, EventArgs e) //仮想サーバーコンボ閉じたとき
    {
      int id = cmbHPVS.SelectedIndex;
      if (id == -1)
      {
        txtHPVSID.Text = "";
      }
      else
      {
        txtHPVSID.Text = cmbHPVS.SelectedValue.ToString();
      }
    }
    private void cmbNicVendor_DropDownClosed(object sender, EventArgs e)//Nicベンダーコンボ閉じたとき
    {
      int id = cmbNicVendor.SelectedIndex;
      if (id == -1)
      {
        txtNICVDID.Text = "";
      }
      else
      {
        txtNICVDID.Text = cmbNicVendor.SelectedValue.ToString();
      }
    }
    private void btNicAdd_Click(object sender, RoutedEventArgs e) //btNicAddボタンでNicデータ行を追加 
    {
      addNicToGr();
    }
    private void btNicDelete_Click(object sender, RoutedEventArgs e) //btNicDeleteボタンでgrNicの該当行を削除
    {
      int rc = grNic.Rows.Count();
      int r = grNic.Selection.Row;
      DataTable dt;
      if (newmode == true) { dt = nicdtldtb; } else { dt = nicloadeditdtb; }
      if (rc != 0)
      {
        if (MessageBox.Show("行 [ " + (r + 1).ToString() + " ] を削除します。よろしいですか？", "Information", MessageBoxButton.YesNo,
       MessageBoxImage.Information) == MessageBoxResult.No)
        { return; }
        else
        {
          DTf.DtRowRemove(r, dt); btgrNicEditInit.PerformClick(); gnicnmbset();
          nicDataClear();
          //tr = dt.Rows[r];dt.Rows.Remove(tr) ; btgrNicEditInit.PerformClick(); gnicnmbset();
        }
      }
      else { MessageBox.Show("行が存在しません！"); }
    }
    private void btNicRenew_Click(object sender, RoutedEventArgs e)
    {
      int r = grNic.Selection.Row;
      if (MessageBox.Show("[ " + (r + 1).ToString() + " ] を更新します。よろしいですか？",
        "Information", MessageBoxButton.YesNo, MessageBoxImage.Information) == MessageBoxResult.No)
      { return; }
      else { nicRowdataRenew(r); }
    }

    private void btNicClear_Click(object sender, RoutedEventArgs e)
    {
      if (MessageBox.Show("Nic情報を初期化します。よろしいですか？", "Information", MessageBoxButton.YesNo,
     MessageBoxImage.Information) == MessageBoxResult.No)
      { return; }
      else
      {
        if (newmode == true) { MakeNicDtldtb(); nicDataClear(); } else { edcpdtload(); setEditMode(2); }

      }

    }

    private void txtSvName_GotFocus(object sender, RoutedEventArgs e)//txtSvNameフォーカス時
    {
    }
    private void txtSvId_LostFocus(object sender, RoutedEventArgs e) //txtSvIdフォーカス離脱時 
    {
      if (!string.IsNullOrEmpty(txtSvId.Text))
      {
        string sql = "";
        cn = Dcn.newcfgcn2("svmente", 5);
        sql = " select count(cpid) from cpmst where cpid =" + txtSvId.Text;
        //MessageBox.Show(sql);
        int r = Dfn.DbCountChk(cn, sql, "txtSvId_LostFocus", "");
        if (r == 0)
        {
          MessageBox.Show("その番号は存在しません！");
          txtSvName.Focus();
        }
        else if (r == 1) { nicDataClear(); edcpdtload(); }
      }
      //DbCountChk
    }

    private void txtSvName_LostFocus(object sender, RoutedEventArgs e)
    {
      if (newmode == true)
      {
        if (chkSvName() == true)
        {
          MessageBox.Show("そのサーバー名は既に登録されています！", "登録済み", MessageBoxButton.OK, MessageBoxImage.Warning);
          txtSvName.Focus();
          txtSvName.Text = null; txtSvName.Focus();
        }
      }

    }

    private void btgrNicInit_Click(object sender, RoutedEventArgs e)
    {
      grnicinit();
    }
    private void btgrNicEditInit_Click(object sender, RoutedEventArgs e)
    {
      DataTable t = nicloadeditdtb; grNic.DataContext = t; grnicinit(); SetgrNicData();
    }

    private void btSave_Click(object sender, RoutedEventArgs e) //更新ボタンクリック
    {
      if (string.IsNullOrEmpty(txtSvName.Text)) { MessageBox.Show("サーバー名は必須です！", "サーバー名の入力", MessageBoxButton.OK, MessageBoxImage.Warning); txtSvName.Focus(); return; }
      if (newmode == true)
      {
        if (MessageBox.Show("データを追加します。よろしいですか？",
          "Information", MessageBoxButton.YesNo, MessageBoxImage.Information) == MessageBoxResult.No)
        { return; }
        else { AddDataToDB(); }
      }
      else if (newmode == false)
      {
        if (MessageBox.Show("データを更新します。よろしいですか？",
    "Information", MessageBoxButton.YesNo, MessageBoxImage.Information) == MessageBoxResult.No)
        { return; }
        else { RenewEditDataToDB(); }
      }
    }

    private void btDelete_Click(object sender, RoutedEventArgs e) //削除ボタンクリック
    {
      if (newmode == false)
      {
        if (MessageBox.Show("データを削除します。よろしいですか？",
    "Information", MessageBoxButton.YesNo, MessageBoxImage.Information) == MessageBoxResult.No)
        { return; }
        else { DelSvToDB(); }
      }
    }


    private void btExit_Click(object sender, RoutedEventArgs e)
    {
      if (MessageBox.Show("プログラムを終了します。よろしいですか？",
  "Information", MessageBoxButton.YesNo, MessageBoxImage.Information) == MessageBoxResult.No)
      { txtSvName.Focus(); return; }
      else { this.Close(); }
    }

    private void btShowSvAll_Click(object sender, RoutedEventArgs e)
    {
      svallshowunit();
    }

    private void btVMHPVEXP_Click(object sender, RoutedEventArgs e) //VM ハイパーバイザー一覧エクスポート
    {
      var dlg = new Microsoft.Win32.SaveFileDialog();
      dlg.DefaultExt = "xlsx";
      dlg.Filter = "Excel Workbook (*.xlsx)|*.xlsx|"
        + "HTML File(*.htm; *.html)| *.htm; *.html | "
        + "Comma Separated Values(*.csv) | *.csv | " + "Text File(*.txt) | *.txt";
      if (dlg.ShowDialog() == true)
      {
        var ext = System.IO.Path.GetExtension(dlg.SafeFileName).ToLower();
        ext = ext == ".htm" ? "ehtm" : ext == ".html" ? "ehtm" : ext;
        switch (ext)
        {
          case "ehtm":
            {
              grVMHPVALL.Save(dlg.FileName,
              C1.WPF.FlexGrid.FileFormat.Html, SaveOptions.Formatted);
              break;
            }
          case ".csv":
            {
              grVMHPVALL.Save(dlg.FileName,
              C1.WPF.FlexGrid.FileFormat.Csv, SaveOptions.Formatted);
              break;
            }
          case ".txt":
            {
              grVMHPVALL.Save(dlg.FileName,
              C1.WPF.FlexGrid.FileFormat.Text, SaveOptions.Formatted);
              break;
            }
          default:
            {
              XlsSave(dlg.FileName, grVMHPVALL);
              break;
            }
        }
      }

    }

    private void btXlsExport_Click(object sender, RoutedEventArgs e) //サーバー一覧のエクスポート
    {
      var dlg = new Microsoft.Win32.SaveFileDialog();
      dlg.DefaultExt = "xlsx";
      dlg.Filter = "Excel Workbook (*.xlsx)|*.xlsx|"
        + "HTML File(*.htm; *.html)| *.htm; *.html | "
        + "Comma Separated Values(*.csv) | *.csv | " + "Text File(*.txt) | *.txt";
      if (dlg.ShowDialog() == true)
      {
        var ext = System.IO.Path.GetExtension(dlg.SafeFileName).ToLower();
        ext = ext == ".htm" ? "ehtm" : ext == ".html" ? "ehtm" : ext;
        switch (ext)
        {
          case "ehtm":
            {
              grSVALL.Save(dlg.FileName,
              C1.WPF.FlexGrid.FileFormat.Html, SaveOptions.Formatted);
              break;
            }
          case ".csv":
            {
              grSVALL.Save(dlg.FileName,
              C1.WPF.FlexGrid.FileFormat.Csv, SaveOptions.Formatted);
              break;
            }
          case ".txt":
            {
              grSVALL.Save(dlg.FileName,
              C1.WPF.FlexGrid.FileFormat.Text, SaveOptions.Formatted);
              break;
            }
          default:
            {
              XlsSave(dlg.FileName, grSVALL);
              break;
            }
        }
      }

    }


    private void button5_Click(object sender, RoutedEventArgs e)
    {
      int rc1 = nicloadeditdtb.Rows.Count;
      int rc2 = nicloadorgdtb.Rows.Count;
      MessageBox.Show(rc1.ToString());
      MessageBox.Show(rc2.ToString());
    }
    private void grNic_Click(object sender, MouseButtonEventArgs e)
    {
      SetgrNicData();
    }

    //***************************************************************************************************
    //コントロールのイベント end


    //作成モジュール begin
    //***************************************************************************************************
    private void setEditMode(int m) //編集モード表示の設定
    {
      if (m == 1)
      {
        txtEditMode.Background = new SolidColorBrush(Color.FromArgb(0xFF, 0x21, 0x07, 0x78));
        txtEditMode.Foreground = new SolidColorBrush(Colors.White);
        txtEditMode.Text = "新 規";
        btDelete.Visibility = Visibility.Hidden;
      }
      if (m == 2)
      {
        txtEditMode.Background = new SolidColorBrush(Colors.Yellow);
        txtEditMode.Foreground = new SolidColorBrush(Colors.Black);
        txtEditMode.Text = "編 集";
        btDelete.Visibility = Visibility.Visible;
      }
    }
    private void txtSvId_PreviewTextInput(object sender, TextCompositionEventArgs e) //数値のみ入力できる
    {
      // 0-9のみ
      e.Handled = !new Regex("[0-9]").IsMatch(e.Text);
    }
    private void setAllCombo() //コンボボックス一括設定
    {
      cn = Dcn.newcfgcn2("svmente", 5);
      string sql = "";
      //OSコンボ
      cn = Dcn.newcfgcn2("svmente", 5);
      sql = " select 0 osid,null osname union select osid,osname from osmst order by osid ";
      Dfn.MkDbFromDtb(cn, oscombodtb, sql, "OSコンボ", "");
      cmbOS.DataContext = oscombodtb;
      cmbOS.SelectedValuePath = "osid";
      //OSEDコンボ
      cn = Dcn.newcfgcn2("svmente", 5);
      sql = " select 0 osedid,null osedname union select osedid,osedname from osedmst order by osedid ";
      Dfn.MkDbFromDtb(cn, osedcombodtb, sql, "OSコンボ", "");
      cmbED.DataContext = osedcombodtb;
      cmbED.SelectedValuePath = "osedid";
      //本体ベンダー
      cn = Dcn.newcfgcn2("svmente", 5);
      sql = " select 0 vdid,null vdname union select vdid,vdname from vendormst order by vdid ";
      Dfn.MkDbFromDtb(cn, vdcombodtb, sql, "本体ベンダー", "");
      cmbVD.DataContext = vdcombodtb;
      cmbVD.SelectedValuePath = "vdid";
      //CPUベンダー
      cn = Dcn.newcfgcn2("svmente", 5);
      sql = " select 0 cpuvdid,null cpuvdname union select cpuvdid,cpuvdname from cpuvdmst order by cpuvdid ";
      Dfn.MkDbFromDtb(cn, cpuvdcombodtb, sql, "CPUベンダー", "");
      cmbCPUVD.DataContext = cpuvdcombodtb;
      cmbCPUVD.SelectedValuePath = "cpuvdid";
      //用途コンボ１
      cn = Dcn.newcfgcn2("svmente", 5);
      sql = " select 0 usetypeid,null usetypename union select usetypeid,usetypename from usetypemst order by usetypeid ";
      Dfn.MkDbFromDtb(cn, usetype1dtb, sql, "用途コンボ１", "");
      cmbUSETYPE1.DataContext = usetype1dtb;
      cmbUSETYPE1.SelectedValuePath = "usetypeid";
      //用途コンボ２
      cn = Dcn.newcfgcn2("svmente", 5);
      usetype2dtb = usetype1dtb.Copy();
      cmbUSETYPE2.DataContext = usetype2dtb;
      cmbUSETYPE2.SelectedValuePath = "usetypeid";
      //用途コンボ３
      usetype3dtb = usetype1dtb.Copy();
      cmbUSETYPE3.DataContext = usetype3dtb;
      cmbUSETYPE3.SelectedValuePath = "usetypeid";
      //設置場所コンボ
      cn = Dcn.newcfgcn2("svmente", 5);
      sql = "  select 0 placeid,null plname union select placeid,plname from placemst order by placeid ";
      Dfn.MkDbFromDtb(cn, placedtb, sql, "設置場所コンボ", "");
      cmbPlace.DataContext = placedtb;
      cmbPlace.SelectedValuePath = "placeid";
      //仮想サーバーコンボ
      cn = Dcn.newcfgcn2("svmente", 5);
      sql = " select 0 cpid,null cpname union select cpid,cpname \r\n";
      sql = sql + " from cpmst \r\n";
      sql = sql + " where osid between 50 and 64 or osid in (15,16,17,18,19,20,21,22,23,33,34) order by cpid \r\n";
      Dfn.MkDbFromDtb(cn, vmsvcombodtb, sql, "仮想サーバーコンボ", "");
      cmbHPVS.DataContext = vmsvcombodtb;
      cmbHPVS.SelectedValuePath = "cpid";
      //Nicベンダーコンボ
      cn = Dcn.newcfgcn2("svmente", 5);
      sql = " select nicvdid,nicvdname from nicvdmst order by nicvdid ";
      Dfn.MkDbFromDtb(cn, nicvdcombodtb, sql, "Nicベンダーコンボ", "");
      cmbNicVendor.DataContext = nicvdcombodtb;
      cmbNicVendor.SelectedValuePath = "nicvdid";
    }
    private int getCmbIdx(DataTable dt, string fld, int vl) //DataTableのインデックス値を検索 dt:目的のﾃﾞｰﾀﾃｰﾌﾞﾙ fld:列名 vl:探す値
    {
      string flvl;
      flvl = fld + " = " + vl.ToString();
      var r = dt.Select(flvl);
      int i;
      i = dt.Rows.IndexOf(r[0]);
      return i;
    }
    private int getDtbIdx(DataTable dt, string fld, int vl) //DataTableのインデックス値を検索 dt:目的のﾃﾞｰﾀﾃｰﾌﾞﾙ fld:列名 vl:探す値
    {
      string flvl;
      flvl = fld + " = " + vl.ToString();
      var r = dt.Select(flvl);
      int i;
      i = dt.Rows.IndexOf(r[0]);
      return i;
    }
    private void WindowInit() //メインウィンドウの初期化
    {
      newmode = true;
      txtSvId.Text = null;
      txtSvName.Text = null;
      txtOSID.Text = null;
      cmbOS.SelectedIndex = -1;
      txtVer.Text = null;
      txtEDID.Text = null;
      cmbED.SelectedIndex = -1;
      txtVDID.Text = null;
      cmbVD.SelectedIndex = -1;
      txtTYPE.Text = null;
      txtCPUVDID.Text = null;
      cmbCPUVD.SelectedIndex = -1;
      txtCPUTYPE.Text = null;
      txtRAM.Value = 0;
      txtUSETYPE1ID.Text = null;
      cmbUSETYPE1.SelectedIndex = -1;
      txtUSETYPE2ID.Text = null;
      cmbUSETYPE2.SelectedIndex = -1;
      txtUSETYPE3ID.Text = null;
      cmbUSETYPE3.SelectedIndex = -1;
      txtPlaceID.Text = null;
      cmbPlace.SelectedIndex = -1;
      chkVM.IsChecked = false;
      txtHPVSID.Text = null;
      cmbHPVS.SelectedIndex = -1;
      txtAdminID.Text = null;
      txtAdminPASS.Text = null;
      txtNote.Text = null;
      txtMainMac.Text = null;
      txtMainIP.Text = null;
      txtMainSubnet.Text = null;
      txtMainGW.Text = null;
      MakeNicDtldtb();
      MakeNicDtldtb();
      grnicinit();
      nicDataClear();
      setEditMode(1);
      showsvallgrid();
      getActSV();
      getDelSV();
      getRealSVCount();
      getVMSVCount();
      hpvvmallload();
      txtSvName.Focus();
    }
    private void SetGrid()
    {
      cn = Dcn.newcfgcn2("Connection1", 5);
      var dt = new DataTable();
      string sql = "select p.pdctcode '品番',p.pdctname '品名',c.cstmcode '顧客番号',c.cstmabbr '顧客名' from product p left join customer c on p.cstmcode=c.cstmcode ";
      Dfn.MkDbFromDtb(cn, dt, sql, "", "");
    }
    private void SetSvCombo()
    {
      cn = Dcn.newcfgcn2("svmente", 5);
      string sql = "";
      sql = sql + " select c.cpid,c.cpname,n.ip from cpmst c " + "\r\n";
      sql = sql + "  left join ( " + "\r\n";
      sql = sql + " select cpid, " + "\r\n";
      sql = sql + "  substring(mainip,1,3)+'.'+substring(mainip,4,3)+'.'+substring(mainip,7,3)+'.'+substring(mainip,10,3) ip " + "\r\n";
      sql = sql + " from nicmst where manage=1 " + "\r\n";
      sql = sql + " ) n on c.cpid=n.cpid " + "\r\n";
      Dfn.MkDbFromDtb(cn, svcombodtb, sql, "", "");

      //cmbSV.DataContext = svcombodtb;
      //cmbSV.ItemsSource = svcombodtb.DefaultView;
      //cmbSV.DisplayMemberPath = "cpname";
      //cmbSV.SelectedValuePath = "cpid";

      //comboBox.DataContext = svcombodtb;
      //comboBox.SelectedValuePath = "cpid";

      //cmbSV.DataContext = svcombodtb;
      //cmbSV.SelectedValuePath = "cpid";

    }
    private void getActSV() //稼働中サーバーのカウント表示
    {
      int cnt = 0;
      string sql = "select count(cpname) from cpmst where isnull(updtsgmt,0)=0 ";
      cn = Dcn.newcfgcn2("svmente", 5);
      cnt = Dfn.DbCountChk(cn, sql, "getActSV", "");
      txtActSv.Content = cnt.ToString();
    }
    private void getDelSV() //撤去サーバーのカウント表示
    {
      int cnt = 0;
      string sql = "select count(cpname) from cpmst where isnull(updtsgmt,0)=9 ";
      cn = Dcn.newcfgcn2("svmente", 5);
      cnt = Dfn.DbCountChk(cn, sql, "getDelSV", "");
      txtDelSv.Content = cnt.ToString();
    }
    private void getRealSVCount() //稼働中実機サーバーのカウント表示
    {
      int cnt = 0;
      string sql = "select count(cpid) from cpmst where updtsgmt<>9 and vm=0 ";
      cn = Dcn.newcfgcn2("svmente", 5);
      cnt = Dfn.DbCountChk(cn, sql, "getRealSVCount", "");
      txtRealSv.Content = cnt.ToString();
    }
    private void getVMSVCount() //稼働中仮想サーバーのカウント表示
    {
      int cnt = 0;
      string sql = "select count(cpid) from cpmst where updtsgmt<>9 and vm=1 ";
      cn = Dcn.newcfgcn2("svmente", 5);
      cnt = Dfn.DbCountChk(cn, sql, "getVMCount", "");
      txtVmSv.Content = cnt.ToString();
    }

    private bool chkSvName() //サーバーが登録されているかどうかの確認モジュール
    {
      bool rtn = false;
      int cnt = 9999;
      string sql = "select count(rtrim(ltrim(cpname))) cnt from cpmst where rtrim(ltrim(cpname)) ='" + txtSvName.Text + "'";
      cn = Dcn.newcfgcn2("svmente", 5);
      cnt = Dfn.DbCountChk(cn, sql, "chkSvName", "");
      if (cnt == 1) { rtn = true; } else { rtn = false; }
      return rtn;
    }
    private void hpvvmallload() //ハイパーバイザーと仮想マシンの一覧
    {
      string sql = "";
      sql = "" + "\r\n";
      sql = sql + " select  " + "\r\n";
      sql = sql + "  c.cpname 'VM名',c.admin 'ログインID',c.adminpass 'ログインPass',c2.cpname 'ハイパーバイザー', " + "\r\n";
      sql = sql + "  nm.macadrs 'Mac Adress',nm.ip 'IP Adress',nm.pingip 'Ping IP',nm.subnet 'SubNet',nm.gw 'Gateway'   " + "\r\n";
      sql = sql + " from cpmst c  " + "\r\n";
      sql = sql + "  left join  " + "\r\n";
      sql = sql + "  ( " + "\r\n";
      sql = sql + " select cpid, " + "\r\n";
      sql = sql + "  substring(macadrs,1,2)+':'+substring(macadrs,3,2)+':'+substring(macadrs,5,2)+':' " + "\r\n";
      sql = sql + "  +substring(macadrs,7,2)+':'+substring(macadrs,9,2)+':'+substring(macadrs,11,2) macadrs, " + "\r\n";
      sql = sql + "  substring(mainip,1,3)+'.'+substring(mainip,4,3)+'.'+substring(mainip,7,3) " + "\r\n";
      sql = sql + "  +'.'+substring(mainip,10,3) ip, " + "\r\n";
      sql = sql + "  convert(varchar,convert(int,substring(mainip,1,3)))+'.'+ " + "\r\n";
      sql = sql + "  convert(varchar,convert(int,substring(mainip,4,3)))+'.'+ " + "\r\n";
      sql = sql + "  convert(varchar,convert(int,substring(mainip,7,3)))+'.'+ " + "\r\n";
      sql = sql + "  convert(varchar,convert(int,substring(mainip,10,3))) pingip, " + "\r\n";
      sql = sql + "  substring(mainsubnet,1,3)+'.'+substring(mainsubnet,4,3)+'.'+substring(mainsubnet,7,3) " + "\r\n";
      sql = sql + "  +'.'+substring(mainsubnet,10,3) subnet,  " + "\r\n";
      sql = sql + "  substring(maingw,1,3)+'.'+substring(maingw,4,3)+'.'+substring(maingw,7,3) " + "\r\n";
      sql = sql + "  +'.'+substring(maingw,10,3) gw  " + "\r\n";
      sql = sql + " from nicmst  " + "\r\n";
      sql = sql + " where manage=1 and updtsgmt<>9  " + "\r\n";
      sql = sql + "  ) nm on c.cpid=nm.cpid  " + "\r\n";
      sql = sql + "  left join  " + "\r\n";
      sql = sql + "  (select cpid,cpname from cpmst) c2 on c.hpvsid=c2.cpid   " + "\r\n";
      sql = sql + "  where c.vm=1  " + "\r\n";
      sql = sql + "  order by nm.ip  " + "\r\n";
      cn = Dcn.newcfgcn2("svmente", 5);
      Dfn.MkDbFromDtb(cn, hpvvmalldtb, sql, "VM HPVS一覧出力時", "");
      grVMHPVALL.DataContext = hpvvmalldtb;

    }

    private void edcpdtload() //編集データのロード
    {
      setEditMode(2);
      newmode = false;
      cn = Dcn.newcfgcn2("svmente", 5);
      string sql = "";
      sql = sql + " select \r\n";
      sql = sql + "  c.cpid,c.cpname,c.osid,c.osverid,c.ver,c.osedid, \r\n";
      sql = sql + "  c.vdid,c.[type],c.cpuvdid,c.cputype,c.ram, \r\n";
      sql = sql + "  c.usetype1id,c.usetype2id,c.usetype3id, \r\n";
      sql = sql + "  c.placeid,c.baseid,c.vm,c.hpvsid,c.mngip,c.[admin],c.adminpass,c.note, \r\n";
      sql = sql + "  c.regsdate,c.regsopr,c.altrdate,c.altropr \r\n";
      sql = sql + " from cpmst c \r\n";
      sql = sql + " where c.cpid= " + txtSvId.Text + " \r\n";
      Dfn.MkDbFromDtb(cn, editcpmstdtb, sql, "edcpdtload", "");
      DataTable t = editcpmstdtb;
      editcpmstorgdtb = t.Copy();
      if (String.IsNullOrEmpty(t.Rows[0]["cpname"].ToString())) { txtSvName.Text = ""; }
      else { txtSvName.Text = t.Rows[0]["cpname"].ToString(); }
      if (String.IsNullOrEmpty(t.Rows[0]["osid"].ToString())) { txtOSID.Text = ""; cmbOS.SelectedIndex = -1; }
      else
      {
        txtOSID.Text = t.Rows[0]["osid"].ToString();
        cmbOS.SelectedIndex = getCmbIdx(oscombodtb, "osid", Int32.Parse(t.Rows[0]["osid"].ToString()));
      }
      if (String.IsNullOrEmpty(t.Rows[0]["ver"].ToString())) { txtVer.Text = ""; }
      else { txtVer.Text = t.Rows[0]["ver"].ToString(); }
      if (String.IsNullOrEmpty(t.Rows[0]["osedid"].ToString())) { txtEDID.Text = ""; cmbED.SelectedIndex = -1; }
      else
      {
        txtEDID.Text = t.Rows[0]["osedid"].ToString();
        cmbED.SelectedIndex = getCmbIdx(osedcombodtb, "osedid", Int32.Parse(t.Rows[0]["osedid"].ToString()));
      }
      if (String.IsNullOrEmpty(t.Rows[0]["vdid"].ToString())) { txtVDID.Text = ""; cmbVD.SelectedIndex = -1; }
      else
      {
        txtVDID.Text = t.Rows[0]["vdid"].ToString();
        cmbVD.SelectedIndex = getCmbIdx(vdcombodtb, "vdid", Int32.Parse(t.Rows[0]["vdid"].ToString()));
      }
      if (String.IsNullOrEmpty(t.Rows[0]["type"].ToString())) { txtTYPE.Text = ""; }
      else { txtTYPE.Text = t.Rows[0]["type"].ToString(); }
      if (String.IsNullOrEmpty(t.Rows[0]["cpuvdid"].ToString())) { txtCPUVDID.Text = ""; cmbCPUVD.SelectedIndex = -1; }
      else
      {
        txtCPUVDID.Text = t.Rows[0]["cpuvdid"].ToString();
        cmbCPUVD.SelectedIndex = getCmbIdx(cpuvdcombodtb, "cpuvdid", Int32.Parse(t.Rows[0]["cpuvdid"].ToString()));
      }
      if (String.IsNullOrEmpty(t.Rows[0]["cputype"].ToString())) { txtCPUTYPE.Text = ""; }
      else { txtCPUTYPE.Text = t.Rows[0]["cputype"].ToString(); }
      if (String.IsNullOrEmpty(t.Rows[0]["ram"].ToString())) { txtRAM.Value = 0; }
      else { txtRAM.Value = Int32.Parse(t.Rows[0]["ram"].ToString()); }
      if (String.IsNullOrEmpty(t.Rows[0]["usetype1id"].ToString()))
      { txtUSETYPE1ID.Text = ""; cmbUSETYPE1.SelectedIndex = -1; }
      else
      {
        txtUSETYPE1ID.Text = t.Rows[0]["usetype1id"].ToString();
        cmbUSETYPE1.SelectedIndex = getCmbIdx(usetype1dtb, "usetypeid", Int32.Parse(t.Rows[0]["usetype1id"].ToString()));
      }
      if (String.IsNullOrEmpty(t.Rows[0]["usetype2id"].ToString()))
      { txtUSETYPE2ID.Text = ""; cmbUSETYPE2.SelectedIndex = -1; }
      else
      {
        txtUSETYPE2ID.Text = t.Rows[0]["usetype2id"].ToString();
        cmbUSETYPE2.SelectedIndex = getCmbIdx(usetype2dtb, "usetypeid", Int32.Parse(t.Rows[0]["usetype2id"].ToString()));
      }
      if (String.IsNullOrEmpty(t.Rows[0]["usetype3id"].ToString()))
      { txtUSETYPE3ID.Text = ""; cmbUSETYPE3.SelectedIndex = -1; }
      else
      {
        txtUSETYPE3ID.Text = t.Rows[0]["usetype3id"].ToString();
        cmbUSETYPE3.SelectedIndex = getCmbIdx(usetype3dtb, "usetypeid", Int32.Parse(t.Rows[0]["usetype3id"].ToString()));
      }
      if (String.IsNullOrEmpty(t.Rows[0]["placeid"].ToString()))
      { txtPlaceID.Text = ""; cmbPlace.SelectedIndex = -1; }
      else
      {
        txtPlaceID.Text = t.Rows[0]["placeid"].ToString();
        cmbPlace.SelectedIndex = getCmbIdx(placedtb, "placeid", Int32.Parse(t.Rows[0]["placeid"].ToString()));
      }

      if (String.IsNullOrEmpty(t.Rows[0]["vm"].ToString())) { chkVM.IsChecked = false; }
      else
      {
        if (bool.Parse(t.Rows[0]["vm"].ToString()) == true) { chkVM.IsChecked = true; } else { chkVM.IsChecked = false; }
      }

      if (String.IsNullOrEmpty(t.Rows[0]["hpvsid"].ToString()))
      { txtHPVSID.Text = ""; cmbHPVS.SelectedIndex = -1; }
      else
      {
        txtHPVSID.Text = t.Rows[0]["hpvsid"].ToString();
        cmbHPVS.SelectedIndex = getCmbIdx(vmsvcombodtb, "cpid", Int32.Parse(t.Rows[0]["hpvsid"].ToString()));
      }


      if (String.IsNullOrEmpty(t.Rows[0]["note"].ToString())) { txtNote.Text = ""; }
      else { txtNote.Text = t.Rows[0]["note"].ToString(); }
      if (String.IsNullOrEmpty(t.Rows[0]["admin"].ToString())) { txtAdminID.Text = ""; }
      else { txtAdminID.Text = t.Rows[0]["admin"].ToString(); }
      if (String.IsNullOrEmpty(t.Rows[0]["adminpass"].ToString())) { txtAdminPASS.Text = ""; }
      else { txtAdminPASS.Text = t.Rows[0]["adminpass"].ToString(); }
      loadNicToGr();

    }
    private void loadNicToGr() //編集用データのNicデータをグリッドにセット
    {
      cn = Dcn.newcfgcn2("svmente", 5);
      string sql = "";
      sql = sql + " select  " + "\r\n";
      sql = sql + "  nicid,macadrs,mainip,mainsubnet,maingw, " + "\r\n";
      sql = sql + "  subip1,subip2,ifdevname, " + "\r\n";
      sql = sql + "  cpid,secno,onboard,manage,nicvdid,typeno,note   " + "\r\n";
      sql = sql + " from nicmst  " + "\r\n";
      sql = sql + " where cpid = " + txtSvId.Text + " \r\n";
      sql = sql + "  and updtsgmt<>9 " + " \r\n";
      sql = sql + " order by secno " + "\r\n";
      textBoxTest.Text = sql;
      Dfn.MkDbFromDtb(cn, nicloadorgdtb, sql, "Nicデータロード時初期値", "");
      nicloadeditdtb = nicloadorgdtb.Copy();
      btgrNicEditInit.PerformClick();
    }
    private void MakeNicDtldtb() //Nicデータ追加用データテーブルのイニシャライズ
    {
      grNic.DataContext = null;
      DataTable t = nicdtldtb;
      t.Rows.Clear();
      t.Columns.Clear();
      t.Columns.Add("nicid", Type.GetType("System.Int32"));
      t.Columns.Add("macadrs", Type.GetType("System.String"));
      t.Columns.Add("mainip", Type.GetType("System.String"));
      t.Columns.Add("mainsubnet", Type.GetType("System.String"));
      t.Columns.Add("maingw", Type.GetType("System.String"));
      t.Columns.Add("subip1", Type.GetType("System.String"));
      t.Columns.Add("subip2", Type.GetType("System.String"));
      t.Columns.Add("ifdevname", Type.GetType("System.String"));
      t.Columns.Add("cpid", Type.GetType("System.Int32"));
      t.Columns.Add("secno", Type.GetType("System.Int32"));
      t.Columns.Add("onboard", Type.GetType("System.Boolean"));
      t.Columns.Add("manage", Type.GetType("System.Boolean"));
      t.Columns.Add("nicvdid", Type.GetType("System.Int32"));
      t.Columns.Add("typeno", Type.GetType("System.String"));
      t.Columns.Add("note", Type.GetType("System.String"));
      grNic.DataContext = t;
      btgrNicInit.PerformClick();

    }
    private void MkEdNewNicDtldtb()
    {
      grNic.DataContext = null;
      DataTable t = nicloadeditdtb;
      t.Rows.Clear();
      t.Columns.Clear();
      t.Columns.Add("nicid", Type.GetType("System.Int32"));
      t.Columns.Add("macadrs", Type.GetType("System.String"));
      t.Columns.Add("mainip", Type.GetType("System.String"));
      t.Columns.Add("mainsubnet", Type.GetType("System.String"));
      t.Columns.Add("maingw", Type.GetType("System.String"));
      t.Columns.Add("subip1", Type.GetType("System.String"));
      t.Columns.Add("subip2", Type.GetType("System.String"));
      t.Columns.Add("ifdevname", Type.GetType("System.String"));
      t.Columns.Add("cpid", Type.GetType("System.Int32"));
      t.Columns.Add("secno", Type.GetType("System.Int32"));
      t.Columns.Add("onboard", Type.GetType("System.Boolean"));
      t.Columns.Add("manage", Type.GetType("System.Boolean"));
      t.Columns.Add("nicvdid", Type.GetType("System.Int32"));
      t.Columns.Add("typeno", Type.GetType("System.String"));
      t.Columns.Add("note", Type.GetType("System.String"));
      grNic.DataContext = t;
      btgrNicInit.PerformClick();
    }
    private void grnicinit() //ネットワークグリッドgrNicのイニシャライズ
    {
      C1FlexGrid g = grNic;
      int rc = g.Rows.Count();
      var ch = g.ColumnHeaders;
      ch.Rows[0].Height = 35;
      ch[0, 0] = "ID";
      ch[0, 1] = "MAC";
      ch[0, 2] = "主IP";
      ch[0, 3] = "主SubNet";
      ch[0, 4] = "主GW";
      ch[0, 5] = "副IP1";
      ch[0, 6] = "副IP2";
      ch[0, 7] = "Nic\r\nIF名";
      ch[0, 8] = "CPID";
      ch[0, 9] = "並順";
      ch[0, 10] = "板上";
      ch[0, 11] = "管理";
      ch[0, 12] = "ﾍﾞﾝﾀﾞ\r\n番号";
      ch[0, 13] = "型式";
      ch[0, 14] = "備考";
      for (int i = 0; i < 15; ++i)
      {
        g.Columns[i].HeaderHorizontalAlignment = HorizontalAlignment.Center;
        g.Columns[i].HeaderFontSize = 13;
        g.Columns[i].HeaderForeground = new SolidColorBrush(Colors.White);
        g.Columns[i].HeaderBackground = new SolidColorBrush(Colors.RoyalBlue);
        //g.Background = new SolidColorBrush(Colors.Pink);
      }
      g.Columns[0].Width = new GridLength(38);
      g.Columns[1].Width = new GridLength(90);
      for (int i = 2; i < 7; ++i) { g.Columns[i].Width = new GridLength(86); }
      for (int i = 0; i < 15; ++i) { g.Columns[i].HorizontalAlignment = HorizontalAlignment.Center; }
      g.Columns[7].HorizontalAlignment = HorizontalAlignment.Left;
      g.Columns[7].Width = new GridLength(70);
      g.Columns[7].FontSize = 11;
      g.Columns[7].TextWrapping = true;
      g.Columns[7].HeaderTextWrapping = true;
      g.Columns[8].Width = new GridLength(40);
      g.Columns[9].Width = new GridLength(45);
      g.Columns[10].Width = new GridLength(38);
      g.Columns[11].Width = new GridLength(38);
      g.Columns[12].Width = new GridLength(55);
      g.Columns[13].Width = new GridLength(130);
      g.Columns[13].HorizontalAlignment = HorizontalAlignment.Left;
      g.Columns[13].FontSize = 11;
      g.Columns[14].Width = new GridLength(130);
      g.Columns[14].HorizontalAlignment = HorizontalAlignment.Left;

      if (rc > 0)
      {
        for (int j = 0; j < g.Rows.Count; ++j)
        {
          g.RowHeaders[j, 0] = (j + 1).ToString();
          g.Rows[j].HeaderHorizontalAlignment = HorizontalAlignment.Center;
        }
      }

    }
    private void addNicToGr() //Nic情報を行に追加
    {
      DataRow row;
      if (newmode == true) { row = nicdtldtb.NewRow(); }
      else { row = nicloadeditdtb.NewRow(); }
      row["nicid"] = DBNull.Value;
      row["macadrs"] = txtMainMac.Value;
      if (txtMainIP.Value == "999999999999") { row["mainip"] = DBNull.Value; }
      else { row["mainip"] = txtMainIP.Value; }
      if (txtMainSubnet.Value == "999999999999") { row["mainsubnet"] = DBNull.Value; }
      else { row["mainsubnet"] = txtMainSubnet.Value; }
      if (txtMainGW.Value == "999999999999") { row["maingw"] = DBNull.Value; }
      else { row["maingw"] = txtMainGW.Value; }
      if (txtSUBIP1.Value == "999999999999") { row["subip1"] = DBNull.Value; }
      else { row["subip1"] = txtSUBIP1.Value; }
      if (txtSUBIP2.Value == "999999999999") { row["subip2"] = DBNull.Value; }
      else { row["subip2"] = txtSUBIP2.Value; }
      row["ifdevname"] = txtIFDEVNAME.Text;
      if (newmode == true) { row["cpid"] = DBNull.Value; } else { row["cpid"] = Int32.Parse(txtSvId.Text); }
      row["secno"] = nicdtldtb.Rows.Count + 1;
      if (newmode == true) { row["secno"] = nicdtldtb.Rows.Count + 1; }
      else { row["secno"] = nicloadeditdtb.Rows.Count + 1; }
      row["onboard"] = chkONBOARD.IsChecked;
      row["manage"] = chkNICMANAGE.IsChecked;
      if (cmbNicVendor.SelectedIndex == -1) { row["nicvdid"] = DBNull.Value; }
      else { row["nicvdid"] = cmbNicVendor.SelectedValue; }
      if (string.IsNullOrEmpty(txtNICTYPE.Text)) { row["typeno"] = ""; }
      else { row["typeno"] = txtNICTYPE.Text; }
      row["note"] = txtNicNote.Text;
      if (newmode == true) { nicdtldtb.Rows.Add(row); }
      else { nicloadeditdtb.Rows.Add(row); }
      grnicinit();
      gnicnmbset();
      nicDataClear();
      txtMainMac.Focus();
    }
    private void nicDataClear() //Nic情報をクリア
    {
      txtNICID.Text = null; txtMainMac.Value = null; txtMainIP.Value = null; txtMainSubnet.Value = null;
      txtMainGW.Value = null; txtSUBIP1.Value = null; txtSUBIP2.Value = null; txtIFDEVNAME.Text = null;
      txtCPID.Text = null; txtSECNO.Text = null; chkONBOARD.IsChecked = false; chkNICMANAGE.IsChecked = false;
      txtNICVDID.Text = null; cmbNicVendor.SelectedIndex = -1; txtNICTYPE.Text = null; txtNicNote.Text = null;
    }
    private void grNicRenew()
    {
      int c = nicdtldtb.Rows.Count;
      for (int i = 0; i < c; ++i)
      {
        nicdtldtb.Rows[i]["secno"] = i + 1;
      }
    }
    private void gnicnmbset()
    {
      DataTable t;
      if (newmode == true) { t = nicdtldtb; } else { t = nicloadeditdtb; }
      int trc = t.Rows.Count;
      if (trc > 0)
      {
        for (int i = 0; i < trc; ++i)
        {
          t.Rows[i]["secno"] = i + 1;
        }
      }
      C1FlexGrid g = grNic;
      int rc = g.Rows.Count();
      if (rc > 0)
      {
        for (int i = 0; i < g.Rows.Count; ++i)
        {
          g.RowHeaders[i, 0] = (i + 1).ToString();
          g.Rows[i].HeaderHorizontalAlignment = HorizontalAlignment.Center;
        }
      }
    }
    private void SetgrNicData() // grNicをクリックした時、該当行のデータを画面にセット 
    {
      DataTable t;
      if (newmode == true) { t = nicdtldtb; } else { t = nicloadeditdtb; }
      int rc = grNic.Rows.Count;
      if (rc > 0)
      {
        int r = grNic.Selection.Row;
        if (t.Rows[r]["nicid"] != DBNull.Value) { txtNICID.Text = t.Rows[r]["nicid"].ToString(); } else { txtNICID.Text = null; }
        if (t.Rows[r]["macadrs"] != DBNull.Value) { txtMainMac.Value = t.Rows[r]["macadrs"].ToString(); } else { txtMainMac.Value = ""; }
        if (t.Rows[r]["mainip"] != DBNull.Value) { txtMainIP.Value = t.Rows[r]["mainip"].ToString(); } else { txtMainIP.Value = ""; }
        if (t.Rows[r]["mainsubnet"] != DBNull.Value) { txtMainSubnet.Value = t.Rows[r]["mainsubnet"].ToString(); }
        else { txtMainSubnet.Value = ""; }
        if (t.Rows[r]["maingw"] != DBNull.Value) { txtMainGW.Value = t.Rows[r]["maingw"].ToString(); } else { txtMainGW.Value = ""; }
        if (t.Rows[r]["subip1"] != DBNull.Value) { txtSUBIP1.Value = t.Rows[r]["subip1"].ToString(); } else { txtSUBIP1.Value = ""; }
        if (t.Rows[r]["subip2"] != DBNull.Value) { txtSUBIP2.Value = t.Rows[r]["subip2"].ToString(); } else { txtSUBIP2.Value = ""; }
        if (t.Rows[r]["ifdevname"] != DBNull.Value) { txtIFDEVNAME.Text = t.Rows[r]["ifdevname"].ToString(); }
        else { txtIFDEVNAME.Text = null; }
        if (t.Rows[r]["cpid"] != DBNull.Value) { txtCPID.Text = t.Rows[r]["cpid"].ToString(); } else { txtCPID.Text = null; }
        if (t.Rows[r]["secno"] != DBNull.Value) { txtSECNO.Text = t.Rows[r]["secno"].ToString(); } else { txtSECNO.Text = null; }
        if (t.Rows[r]["onboard"] != DBNull.Value) { chkONBOARD.IsChecked = bool.Parse(t.Rows[r]["onboard"].ToString()); }
        else { chkONBOARD.IsChecked = false; }
        if (t.Rows[r]["manage"] != DBNull.Value) { chkNICMANAGE.IsChecked = bool.Parse(t.Rows[r]["manage"].ToString()); }
        else { chkNICMANAGE.IsChecked = false; }
        if (t.Rows[r]["nicvdid"] != DBNull.Value)
        {
          txtNICVDID.Text = t.Rows[r]["nicvdid"].ToString();
          cmbNicVendor.SelectedIndex = getCmbIdx(nicvdcombodtb, "nicvdid", Int32.Parse(t.Rows[r]["nicvdid"].ToString()));
        }
        else { txtNICVDID.Text = null; cmbNicVendor.SelectedIndex = -1; }
        if (t.Rows[r]["typeno"] != DBNull.Value) { txtNICTYPE.Text = t.Rows[r]["typeno"].ToString(); } else { txtNICTYPE.Text = null; }
        if (t.Rows[r]["note"] != DBNull.Value) { txtNicNote.Text = t.Rows[r]["note"].ToString(); } else { txtNicNote.Text = null; }
      }
    }

    private void AddDataToDB() //情報のＤＢへの新規追加 
    {
      if (chkSvName() == true)
      {
        MessageBox.Show("そのサーバー名は既に登録されています！", "サーバー名重複確認", MessageBoxButton.OK, MessageBoxImage.Error);
        return;
      }
      string sql = "";
      sql = sql + " insert into cpmst \r\n";
      sql = sql + "  ( \r\n";
      sql = sql + " cpname,osid,osverid,ver,osedid, \r\n";
      sql = sql + " vdid,[type],cpuvdid,cputype,ram, \r\n";
      sql = sql + " usetype1id,usetype2id,usetype3id, " + "\r\n";
      sql = sql + " placeid,baseid,vm,hpvsid,mngip,[admin],adminpass,note,updtsgmt, " + "\r\n";
      sql = sql + " regsdate,regsopr,altrdate,altropr " + "\r\n";
      sql = sql + " ) " + "\r\n";
      sql = sql + "  values  " + "\r\n";
      sql = sql + " ( " + "\r\n";
      sql = sql + "'" + txtSvName.Text + "',";
      if (string.IsNullOrEmpty(txtOSID.Text)) { sql = sql + "null,null,"; }
      else { sql = sql + cmbOS.SelectedValue.ToString() + ",null,"; }
      if (string.IsNullOrEmpty(txtVer.Text)) { sql = sql + "'',"; }
      else { sql = sql + "'" + txtVer.Text + "',"; }
      if (string.IsNullOrEmpty(txtEDID.Text)) { sql = sql + "null,\r\n"; }
      else { sql = sql + cmbED.SelectedValue.ToString() + ",\r\n"; }
      if (string.IsNullOrEmpty(txtVDID.Text)) { sql = sql + "null,"; }
      else { sql = sql + cmbVD.SelectedValue.ToString() + ","; }
      if (string.IsNullOrEmpty(txtTYPE.Text)) { sql = sql + "'',"; }
      else { sql = sql + "'" + txtTYPE.Text + "',"; }
      if (string.IsNullOrEmpty(txtCPUVDID.Text)) { sql = sql + "null,"; }
      else { sql = sql + cmbCPUVD.SelectedValue.ToString() + ","; }
      if (string.IsNullOrEmpty(txtCPUTYPE.Text)) { sql = sql + "null,"; }
      else { sql = sql + "'" + txtCPUTYPE.Text.Replace("'", "''") + "',"; }
      sql = sql + txtRAM.Value + ", \r\n";
      if (string.IsNullOrEmpty(txtUSETYPE1ID.Text)) { sql = sql + "null,"; }
      else { sql = sql + cmbUSETYPE1.SelectedValue.ToString() + ","; }
      if (string.IsNullOrEmpty(txtUSETYPE2ID.Text)) { sql = sql + "null,"; }
      else { sql = sql + cmbUSETYPE2.SelectedValue.ToString() + ","; }
      if (string.IsNullOrEmpty(txtUSETYPE3ID.Text)) { sql = sql + "null, \r\n"; }
      else { sql = sql + cmbUSETYPE3.SelectedValue.ToString() + ", \r\n"; }
      if (string.IsNullOrEmpty(txtPlaceID.Text)) { sql = sql + "null,null,"; }
      else { sql = sql + cmbPlace.SelectedValue.ToString() + ",null,"; }
      if (chkVM.IsChecked == true) { sql = sql + "1,"; } else { sql = sql + "0,"; }
      if (string.IsNullOrEmpty(txtHPVSID.Text)) { sql = sql + "null,null,"; }
      else { sql = sql + cmbHPVS.SelectedValue.ToString() + ",null,"; }
      if (string.IsNullOrEmpty(txtAdminID.Text)) { sql = sql + "'',"; }
      else { sql = sql + "'" + txtAdminID.Text + "',"; }
      if (string.IsNullOrEmpty(txtAdminPASS.Text)) { sql = sql + "'',"; }
      else { sql = sql + "'" + txtAdminPASS.Text.Replace("'", "''") + "',"; }
      if (string.IsNullOrEmpty(txtNote.Text)) { sql = sql + "'',0, \r\n"; }
      else { sql = sql + "'" + txtNote.Text.Replace("'", "''") + "',0, \r\n"; }
      sql = sql + " getdate()," + nUserId.ToString() + ",getdate()," + nUserId.ToString() + " \r\n";
      sql = sql + " ); " + "\r\n";
      cn = Dcn.newcfgcn2("svmente", 5);
      Int32 vcpid = Dfn.DbGetNextTblID(cn, "cpmst", "", "");
      string sqlsum = "";
      sqlsum = sql + "\r\n" + GetAddNewNicToDBSql(vcpid);
      textBoxTest.Text = sqlsum;
      cn = Dcn.newcfgcn2("svmente", 5);
      Dfn.Dbprocess(cn, sqlsum, "AddMainDataToDB", "");
      MessageBox.Show("新規追加が完了しました！", "新規追加の完了", MessageBoxButton.OK, MessageBoxImage.Information);
      WindowInit();
    }
    private string GetAddNewNicToDBSql(int vcpid) //新規追加用Nic情報SQL作成
    {
      string rtn = "";
      if (grNic.Rows.Count() > 0)
      {
        DataTable t = nicdtldtb;
        string allsql = "";
        for (int i = 0; i < t.Rows.Count; ++i)
        {
          string sql = "";
          sql = sql + "insert into nicmst \r\n";
          sql = sql + " (macadrs,mainip,mainsubnet,maingw,";
          sql = sql + "subip1,subip2,ifdevname,cpid,secno, \r\n";
          sql = sql + " onboard,manage,nicvdid,typeno,note,updtsgmt,";
          sql = sql + "regsdate,regsopr,altrdate,altropr) \r\n";
          sql = sql + " values \r\n";
          sql = sql + " ('" + t.Rows[i]["macadrs"] + "','" + t.Rows[i]["mainip"] + "','"
            + t.Rows[i]["mainsubnet"] + "','" + t.Rows[i]["maingw"] + "', \r\n";
          sql = sql + " '" + t.Rows[i]["subip1"] + "','" + t.Rows[i]["subip2"] + "',";
          if (string.IsNullOrEmpty(t.Rows[i]["ifdevname"].ToString())) { sql = sql + " '',"; }
          else { sql = sql + " '" + t.Rows[i]["ifdevname"].ToString().Replace("'", "''") + "',"; }
          sql = sql + vcpid.ToString() + ",";
          sql = sql + t.Rows[i]["secno"].ToString() + ", \r\n";
          sql = sql + "'" + t.Rows[i]["onboard"] + "',";
          sql = sql + "'" + t.Rows[i]["manage"] + "',";
          if (string.IsNullOrEmpty(t.Rows[i]["nicvdid"].ToString())) { sql = sql + " null,"; }
          else { sql = sql + t.Rows[i]["nicvdid"].ToString().Replace("'", "''") + ","; }
          if (string.IsNullOrEmpty(t.Rows[i]["typeno"].ToString())) { sql = sql + " '',"; }
          else { sql = sql + "'" + t.Rows[i]["typeno"] + "',"; }
          if (string.IsNullOrEmpty(t.Rows[i]["note"].ToString())) { sql = sql + " '',0,"; }
          else { sql = sql + "'" + t.Rows[i]["note"].ToString().Replace("'", "''") + "',0, \r\n"; }
          sql = sql + " getdate()," + nUserId.ToString() + ",getdate()," + nUserId.ToString() + " \r\n";
          sql = sql + ");\r\n";
          allsql = allsql + sql + "\r\n";
        }
        //textBoxTest.Text = allsql;
        rtn = allsql;
      }
      return rtn;
    }
    private void RenewEditDataToDB() //変更情報のＤＢへの書き込み
    {
      string sql = "";
      sql = sql + "update cpmst set \r\n cpname='" + txtSvName.Text + "' \r\n";
      if (!string.IsNullOrEmpty(txtOSID.Text)) { sql = sql + ",osid=" + cmbOS.SelectedValue.ToString(); }
      if (!string.IsNullOrEmpty(txtVer.Text)) { sql = sql + ",ver='" + txtVer.Text + "'"; }
      if (!string.IsNullOrEmpty(txtEDID.Text)) { sql = sql + ",osedid=" + cmbED.SelectedValue.ToString(); }
      if (!string.IsNullOrEmpty(txtVDID.Text)) { sql = sql + ",vdid=" + cmbVD.SelectedValue.ToString(); }
      if (!string.IsNullOrEmpty(txtTYPE.Text)) { sql = sql + ",[type]='" + txtTYPE.Text + "' \r\n"; }
      if (!string.IsNullOrEmpty(txtCPUVDID.Text)) { sql = sql + ",cpuvdid=" + cmbCPUVD.SelectedValue.ToString(); }
      if (!string.IsNullOrEmpty(txtCPUTYPE.Text)) { sql = sql + ",cputype='" + txtCPUTYPE.Text + "'"; }
      sql = sql + ",ram=" + txtRAM.Value.ToString();
      if (!string.IsNullOrEmpty(txtUSETYPE1ID.Text))
      { sql = sql + ",usetype1id=" + cmbUSETYPE1.SelectedValue.ToString(); }
      if (!string.IsNullOrEmpty(txtUSETYPE2ID.Text))
      { sql = sql + ",usetype2id=" + cmbUSETYPE2.SelectedValue.ToString(); }
      if (!string.IsNullOrEmpty(txtUSETYPE3ID.Text))
      { sql = sql + ",usetype3id=" + cmbUSETYPE3.SelectedValue.ToString() + " \r\n"; }
      if (!string.IsNullOrEmpty(txtPlaceID.Text))
      { sql = sql + ",placeid=" + cmbPlace.SelectedValue.ToString(); }
      sql = sql + ",vm='" + chkVM.IsChecked + "' \r\n";
      if (!string.IsNullOrEmpty(txtHPVSID.Text))
      { sql = sql + ",hpvsid=" + cmbHPVS.SelectedValue.ToString(); }
      if (!string.IsNullOrEmpty(txtAdminID.Text)) { sql = sql + ",[admin]='" + txtAdminID.Text + "'"; }
      if (!string.IsNullOrEmpty(txtAdminPASS.Text)) { sql = sql + ",adminpass='" + txtAdminPASS.Text + "'"; }
      if (!string.IsNullOrEmpty(txtNote.Text)) { sql = sql + ",note='" + txtNote.Text + "'"; }
      sql = sql + ",altrdate=getdate(),altropr=" + nUserId.ToString() + "\r\n";
      sql = sql + "where cpid=" + txtSvId.Text + " ;\r\n";
      sql = sql + GetEditNicDataSql(Int32.Parse(txtSvId.Text.ToString()));
      textBoxTest2.Text = sql;
      cn = Dcn.newcfgcn2("svmente", 5);
      Dfn.Dbprocess(cn, sql, "RenewEditDataToDB", "");
      MessageBox.Show("更新が完了しました！", "更新の完了", MessageBoxButton.OK, MessageBoxImage.Information);
      WindowInit();

    }

    private string GetEditNicDataSql(int cpid)
    {
      DataTable o = nicloadorgdtb; //nicデータ編集前の内容のDatatable
      DataTable t = nicloadeditdtb; //nicデータ編集後の内容のDatatable
      string allsql = "";
      string sql = "";
      int orcnt = o.Rows.Count;
      int trcnt = t.Rows.Count;

      if (orcnt > 0) //元のnicデータがあった場合
      {
        if (trcnt == 0) //編集後のnicデータが全くない場合、既存のnicデータをすべて削除して終了
        {
          //sql = "delete from nicmst where cpid =" + txtSvId.Text + " ;\r\n";
          sql = "update nicmst set updtsgmt=9,deldate=getdate() where cpid =" + txtSvId.Text + " ;\r\n";
          return sql;
        }

        else //編集後のnicデータが存在する場合
        {
          for (int i = 0; i < orcnt; i++) //オリジナルデータ行の数まで
          {
            //オリジナルのnicidが編集テーブルに存在するか
            int j = t.Select("nicid=" + o.Rows[i]["nicid"]).Length;

            if (j == 0) //編集テーブルに存在しない場合
            {
              //sql = "delete from nicmst where nicid=" + o.Rows[i]["nicid"].ToString() + " ;\r\n";
              sql = "update nicmst set updtsgmt=9,deldate=getdate() where nicid=" + o.Rows[i]["nicid"].ToString() + " ;\r\n";
              allsql = allsql + sql;
            }

            //既存のnicidの更新処理
            else //編集テーブルに存在する場合
            {
              int r = getDtbIdx(t, "nicid", Int32.Parse(o.Rows[i]["nicid"].ToString()));
              sql = "update nicmst set macadrs='" + t.Rows[r]["macadrs"] + "',"
                + "mainip='" + t.Rows[r]["mainip"] + "',"
                + "mainsubnet='" + t.Rows[r]["mainsubnet"] + "',\r\n"
                + "maingw='" + t.Rows[r]["maingw"] + "',"
                + "subip1='" + t.Rows[r]["subip1"] + "',"
                + "subip2='" + t.Rows[r]["subip2"] + "',\r\n";
              if (string.IsNullOrEmpty(t.Rows[r]["ifdevname"].ToString())) { sql = sql + "ifdevname='',"; }
              else { sql = sql + "ifdevname='" + t.Rows[r]["ifdevname"] + "',"; }
              sql = sql + "secno=" + t.Rows[r]["secno"] + ","
                + "onboard='" + t.Rows[r]["onboard"] + "'," + "manage='" + t.Rows[r]["manage"] + "',\r\n";
              if (string.IsNullOrEmpty(t.Rows[r]["nicvdid"].ToString())) { sql = sql + "nicvdid=null,"; }
              else { sql = sql + "nicvdid=" + t.Rows[r]["nicvdid"] + ","; }
              if (string.IsNullOrEmpty(t.Rows[r]["typeno"].ToString())) { sql = sql + "typeno='',"; }
              else { sql = sql + "typeno='" + t.Rows[r]["typeno"] + "',"; }
              if (string.IsNullOrEmpty(t.Rows[r]["note"].ToString())) { sql = sql + "note='',\r\n"; }
              else { sql = sql + "note='" + t.Rows[r]["note"] + "',\r\n"; }
              sql = sql + "altrdate=getdate(),altropr=" + nUserId + "\r\n"
                + "where nicid=" + t.Rows[r]["nicid"] + ";\r\n";
              allsql = allsql + sql; //update内容をallsqlにsql内容を追加
            }
          }
          //元のnicデータがなかった場合
          //↓編集テーブルでの追加行処理
          for (int i = 0; i < trcnt; i++) //編集テーブルの行の数まで
          {
            if (string.IsNullOrEmpty(t.Rows[i]["nicid"].ToString())) //nicidが無い場合(追加行の場合)
            {
              sql = "insert into nicmst \r\n(\r\nmacadrs,mainip,mainsubnet,maingw,subip1,subip2,ifdevname,\r\n"
                + "cpid,secno,onboard,manage,nicvdid,typeno,note,updtsgmt,regsdate,regsopr,altrdate,altropr\r\n)\r\n"
                + " values \r\n(\r\n";
              sql = sql + "'" + t.Rows[i]["macadrs"] + "','" + t.Rows[i]["mainip"] + "','"
                + t.Rows[i]["mainsubnet"] + "',\r\n'" + t.Rows[i]["maingw"] + "','"
                + t.Rows[i]["subip1"] + "','" + t.Rows[i]["subip2"] + "',\r\n'"
                + t.Rows[i]["ifdevname"] + "',\r\n" + txtSvId.Text + "," + t.Rows[i]["secno"]
                + ",'" + t.Rows[i]["onboard"] + "','" + t.Rows[i]["manage"] + "',";
              if (string.IsNullOrEmpty(t.Rows[i]["nicvdid"].ToString()))
              { sql = sql + "null,'"; }
              else { sql = sql + t.Rows[i]["nicvdid"] + ",'"; }
              sql = sql + t.Rows[i]["typeno"] + "',\r\n'" + t.Rows[i]["note"] + "',0,\r\n"
                + "getdate()," + nUserId + ",getdate()," + nUserId + ");\r\n";
              allsql = allsql + sql;
            }
          }

        }
        return allsql;
      }

      else
      { //元のオリジナルテーブルに行がなかった場合

        if (trcnt == 0) //編集後のnicデータが全くない場合sqlは空
        {
          sql = " ;\r\n";
          return sql;
        }

        else //編集後の行がある場合
        {
          //↓編集テーブルでの追加行処理
          for (int i = 0; i < trcnt; i++) //編集テーブルの行の数まで
          {
            if (string.IsNullOrEmpty(t.Rows[i]["nicid"].ToString())) //nicidが無い場合(追加行の場合)
            {
              sql = "insert into nicmst \r\n(\r\nmacadrs,mainip,mainsubnet,maingw,subip1,subip2,ifdevname,\r\n"
                + "cpid,secno,onboard,manage,nicvdid,typeno,note,updtsgmt,regsdate,regsopr,altrdate,altropr\r\n)\r\n"
                + " values \r\n(\r\n";
              sql = sql + "'" + t.Rows[i]["macadrs"] + "','" + t.Rows[i]["mainip"] + "','"
                + t.Rows[i]["mainsubnet"] + "',\r\n'" + t.Rows[i]["maingw"] + "','"
                + t.Rows[i]["subip1"] + "','" + t.Rows[i]["subip2"] + "',\r\n'"
                + t.Rows[i]["ifdevname"] + "',\r\n" + txtSvId.Text + "," + t.Rows[i]["secno"]
                + ",'" + t.Rows[i]["onboard"] + "','" + t.Rows[i]["manage"] + "',";
              if (string.IsNullOrEmpty(t.Rows[i]["nicvdid"].ToString()))
              { sql = sql + "null,'"; }
              else { sql = sql + t.Rows[i]["nicvdid"] + ",'"; }
              sql = sql + t.Rows[i]["typeno"] + "',\r\n'" + t.Rows[i]["note"] + "',0,\r\n"
                + "getdate()," + nUserId + ",getdate()," + nUserId + ");\r\n";
              allsql = allsql + sql;
            }
          }
        }
        return allsql;
      }
    }

    private void nicRowdataRenew(int r) //nicデータ行の編集更新
    {
      DataTable t;
      if (newmode == true) { t = nicdtldtb; } else { t = nicloadeditdtb; }
      t.Rows[r]["macadrs"] = txtMainMac.Value;
      if (txtMainIP.Value == "999999999999") { t.Rows[r]["mainip"] = DBNull.Value; }
      else { t.Rows[r]["mainip"] = txtMainIP.Value; }
      if (txtMainSubnet.Value == "999999999999") { t.Rows[r]["mainsubnet"] = DBNull.Value; }
      else { t.Rows[r]["mainsubnet"] = txtMainSubnet.Value; }
      if (txtMainGW.Value == "999999999999") { t.Rows[r]["maingw"] = DBNull.Value; }
      else { t.Rows[r]["maingw"] = txtMainGW.Value; }
      if (txtSUBIP1.Value == "999999999999") { t.Rows[r]["subip1"] = DBNull.Value; }
      else { t.Rows[r]["subip1"] = txtSUBIP1.Value; }
      if (txtSUBIP2.Value == "999999999999") { t.Rows[r]["subip2"] = DBNull.Value; }
      else { t.Rows[r]["subip2"] = txtSUBIP2.Value; }
      if (string.IsNullOrEmpty(txtIFDEVNAME.Text)) { t.Rows[r]["ifdevname"] = DBNull.Value; }
      else { t.Rows[r]["ifdevname"] = txtIFDEVNAME.Text; }
      t.Rows[r]["onboard"] = chkONBOARD.IsChecked;
      t.Rows[r]["manage"] = chkNICMANAGE.IsChecked;
      if (cmbNicVendor.SelectedIndex == -1) { t.Rows[r]["nicvdid"] = DBNull.Value; }
      else { t.Rows[r]["nicvdid"] = cmbNicVendor.SelectedValue; }
      if (string.IsNullOrEmpty(txtNICTYPE.Text)) { t.Rows[r]["typeno"] = DBNull.Value; }
      else { t.Rows[r]["typeno"] = txtNICTYPE.Text; }
      if (string.IsNullOrEmpty(txtNicNote.Text)) { t.Rows[r]["note"] = DBNull.Value; }
      else { t.Rows[r]["note"] = txtNicNote.Text; }
    }

    private void DelSvToDB() //サーバー削除処理
    {
      string sql = "";
      //sql = "delete from nicmst where cpid=" + txtSvId.Text + "\r\n"; //付属のNICから先に削除
      //sql = sql + "delete from cpmst where cpid=" + txtSvId.Text + "\r\n"; //本体はそのあと

      sql = "update nicmst set updtsgmt=9,deldate=getdate() where cpid=" + txtSvId.Text + "\r\n"; //付属のNICから先に削除
      sql = sql + "update cpmst set updtsgmt=9,deldate=getdate(),altrdate=getdate() where cpid=" + txtSvId.Text + "\r\n"; //本体はそのあと
      cn = Dcn.newcfgcn2("svmente", 5);
      Dfn.Dbprocess(cn, sql, "DelSvToDB", "");
      MessageBox.Show("削除が完了しました！", "削除の完了", MessageBoxButton.OK, MessageBoxImage.Information);
      WindowInit();
    }

    private void showsvallgrid() //サーバー一覧の表示
    {
      string sql = "";
      sql = sql + " select c.cpid,c.cpname,t1.usetypename, " + "\r\n";
      sql = sql + "   substring(n.mainip,1,3)+'.'+substring(n.mainip,4,3) " + "\r\n";
      sql = sql + "   +'.'+substring(n.mainip,7,3)+'.'+substring(n.mainip,10,3) ip, " + "\r\n";
      sql = sql + "  convert(varchar,convert(int,substring(n.mainip,1,3)))+'.'+ " + "\r\n";
      sql = sql + "  convert(varchar,convert(int,substring(n.mainip,4,3)))+'.'+ " + "\r\n";
      sql = sql + "  convert(varchar,convert(int,substring(n.mainip,7,3)))+'.'+ " + "\r\n";
      sql = sql + "  convert(varchar,convert(int,substring(n.mainip,10,3))) cmdip, " + "\r\n";
      sql = sql + "   substring(n.macadrs,1,2)+':'+substring(n.macadrs,3,2)+':'+ " + "\r\n";
      sql = sql + "   substring(n.macadrs,5,2)+':'+substring(n.macadrs,7,2)+':'+ " + "\r\n";
      sql = sql + "   substring(n.macadrs,9,2)+':'+substring(n.macadrs,11,2) mac, " + "\r\n";
      sql = sql + "   o.osname,c.ver,p.plname, " + "\r\n";
      sql = sql + "   case when updtsgmt=0 then '' else '削除' end updtsgmt," + "\r\n";
      sql = sql + "   convert(varchar,c.regsdate,111) regsdate, " + "\r\n";
      sql = sql + "   convert(varchar,c.altrdate,111) altrdate,c.altropr " + "\r\n";
      sql = sql + "  from cpmst c " + "\r\n";
      sql = sql + "   left join " + "\r\n";
      sql = sql + "   ( " + "\r\n";
      sql = sql + "  select nicid,cpid,mainip,macadrs from nicmst where manage=1 and updtsgmt<>9 " + "\r\n";
      sql = sql + "  ) n on c.cpid=n.cpid " + "\r\n";
      sql = sql + "   left join usetypemst t1 on c.usetype1id=t1.usetypeid " + "\r\n";
      sql = sql + "   left join placemst p on c.placeid=p.placeid " + "\r\n";
      sql = sql + "   left join osmst o on c.osid=o.osid " + "\r\n";
      sql = sql + "  order by c.cpid " + "\r\n";
      textBoxTest2.Text = sql;
      cn = Dcn.newcfgcn2("svmente", 5);
      Dfn.MkDbFromDtb(cn, svallgriddtb, sql, "サーバ一覧セット時", "");
      grSVALL.DataContext = svallgriddtb;
      btShowSvAll.PerformClick();
    }

    private void svallshowunit()
    {
      C1FlexGrid g = grSVALL;
      var ch = g.ColumnHeaders;
      ch.Rows[0].Height = 35;
      ch[0, 0] = "ID";
      ch[0, 1] = "サーバー名";
      ch[0, 2] = "用途";
      ch[0, 3] = "IP";
      ch[0, 4] = "Ping IP";
      ch[0, 5] = "MAC";
      ch[0, 6] = "OS";
      ch[0, 7] = "Ver";
      ch[0, 8] = "設置場所";
      ch[0, 9] = "削除";
      ch[0, 10] = "登録日";
      ch[0, 11] = "更新日";
      ch[0, 12] = "更新者";

      for (int i = 0; i < 13; ++i)
      {
        g.Columns[i].HeaderHorizontalAlignment = HorizontalAlignment.Center;
        g.Columns[i].HeaderFontSize = 13;
        g.Columns[i].HeaderForeground = new SolidColorBrush(Colors.White);
        g.Columns[i].HeaderBackground = new SolidColorBrush(Colors.RoyalBlue);
      }

      g.Columns[0].Width = new GridLength(42);
      g.Columns[0].DataType = typeof(string);
      g.Columns[1].Width = new GridLength(110);
      g.Columns[2].Width = new GridLength(110);
      g.Columns[3].Width = new GridLength(100);
      g.Columns[4].Width = new GridLength(100);
      g.Columns[5].Width = new GridLength(110);
      g.Columns[6].Width = new GridLength(110);
      g.Columns[7].Width = new GridLength(35);
      g.Columns[8].Width = new GridLength(150);
      g.Columns[9].Width = new GridLength(40);
      //g.Columns[10].DataType = typeof(Boolean);
      g.Columns[10].Width = new GridLength(85);
      g.Columns[11].Width = new GridLength(85);
      g.Columns[12].Width = new GridLength(50);
      int rc = g.Rows.Count();

      if (rc > 0)
      {
        for (int j = 0; j < g.Rows.Count; ++j)
        {
          g.RowHeaders[j, 0] = (j + 1).ToString();
          g.Rows[j].HeaderHorizontalAlignment = HorizontalAlignment.Center;
        }
      }

    }

    public void XlsSave(string filename, C1FlexGrid flexgrid) //エクセル保存用モジュール
    {
      // 保存するエクセルブックを作成します
      var book = new C1XLBook();
      book.Sheets.Clear();
      var xlSheet = book.Sheets.Add("Sheet1");
      ExcelFilter.Save(flexgrid, xlSheet);
      // エクセルブックを保存します
      book.Save(filename, C1.WPF.Excel.FileFormat.OpenXml);
    }

    //***************************************************************************************************
    //作成モジュール end

    private void button1_Click(object sender, RoutedEventArgs e)
    {
      string a = "2.5' 3.5' インチ 3.5インチ '";
      a = a.Replace("'", "''");
      MessageBox.Show(a.ToString());
      //textBoxTest.Text = txtMainMac.Value;
    }

    private void cmbOS_GotFocus(object sender, RoutedEventArgs e)
    {

    }

    private void btShowVmHpvAll_Click(object sender, RoutedEventArgs e)
    {

    }



    //***************************************************************************************************
    //作成モジュール end


  }
}
