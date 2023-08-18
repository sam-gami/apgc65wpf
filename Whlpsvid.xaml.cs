using System;
using System.Data;
using System.Data.OleDb;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using apgc65wpf;
using AtgPubCs;
using C1.WPF.FlexGrid;

namespace apgc65wpf
{
  /// <summary>
  /// Whlpsvid.xaml の相互作用ロジック
  /// </summary>
  public partial class Whlpsvid : Window
  {
    public OleDbConnection cn = new OleDbConnection();
    public DataTable initdtb = new DataTable();
    public DataTable ossrchcombodtb = new DataTable();
    public DataTable usetypesrchdtb = new DataTable();
    public DataTable svsrchgriddtb = new DataTable();

    public Whlpsvid()
    {
      InitializeComponent();
      // Enter キーでフォーカス移動する
      this.KeyDown += (sender, e) =>
      {
        if (e.Key != Key.Enter) { return; }
        var direction = Keyboard.Modifiers == ModifierKeys.Shift ? FocusNavigationDirection.Previous : FocusNavigationDirection.Next;
        (FocusManager.GetFocusedElement(this) as FrameworkElement)?.MoveFocus(new TraversalRequest(direction));
      };
    }
    private void Window_Loaded(object sender, RoutedEventArgs e)//Windowsロード時
    {
      ossrchcombodtb = MainWindow.oscombodtb.Copy();
      usetypesrchdtb = MainWindow.usetype1dtb.Copy();
      setCombo();
      grsvsrchinit();
    }

    private void HelpWindow_Initialized(object sender, EventArgs e)
    {
      initdtbinit();

    }

    //作成モジュール begin
    //***************************************************************************************************

    private void initdtbinit()
    {
      DataTable t = initdtb;
      t.Rows.Clear();
      t.Columns.Clear();
      t.Columns.Add("cpid", Type.GetType("System.Int32"));
      t.Columns.Add("cpname", Type.GetType("System.String"));
      t.Columns.Add("osname", Type.GetType("System.String"));
      t.Columns.Add("ver", Type.GetType("System.String"));
      t.Columns.Add("usetypename", Type.GetType("System.String"));
      t.Columns.Add("plname", Type.GetType("System.String"));
      grSVSRCH.DataContext = t;

    }

    private void setCombo()//コンボ設定
    {
      cmbOS.DataContext = ossrchcombodtb;
      cmbOS.SelectedValuePath = "osid";
      cmbUSETYPE1.DataContext = usetypesrchdtb;
      cmbUSETYPE1.SelectedValuePath = "usetypeid";
    }
    private void setgrSvSrch()//検索結果グリッドの表示
    {
      string sql = "";
      sql = sql + " select c.cpid,c.cpname,o.osname,c.ver,u.usetypename,p.plname " + "\r\n";
      sql = sql + " from cpmst c " + "\r\n";
      sql = sql + "  left join osmst o on c.osid=o.osid " + "\r\n";
      sql = sql + "  left join osedmst oe on c.osedid=oe.osedid " + "\r\n";
      sql = sql + "  left join usetypemst u on c.usetype1id=u.usetypeid " + "\r\n";
      sql = sql + "  left join placemst p on c.placeid=p.placeid " + "\r\n";
      sql = sql + " where c.cpid is not null " + "\r\n";
      if (!string.IsNullOrEmpty(txtOSID.Text))
      { sql = sql + "  and c.osid = " + txtOSID.Text + "\r\n"; }
      if (!string.IsNullOrEmpty(txtUSETYPE1ID.Text))
      { sql = sql + "  and c.usetype1id = " + txtUSETYPE1ID.Text + "\r\n"; }
      if (!string.IsNullOrEmpty(txtSvName.Text))
      { sql = sql + "  and c.cpname like '%" + txtSvName.Text + "%'" + "\r\n"; }
      if (chkVM.IsChecked==true)
      { sql = sql + "  and c.vm = 1 \r\n"; }
      sql = sql + " order by c.cpid " + "\r\n";
      textBox.Text = sql;
      cn = Dcn.newcfgcn2("svmente", 5);
      Dfn.MkDbFromDtb(cn, svsrchgriddtb, sql, "setgrSvSrch", "");
      grSVSRCH.DataContext = svsrchgriddtb;
      grsvsrchinit();
    }
    private void grsvsrchinit()//検索結果グリッドのイニシャライズ
    {
      C1FlexGrid g = grSVSRCH;
      var ch = g.ColumnHeaders;
      ch.Rows[0].Height = 35;
      g.SelectionMode = C1.WPF.FlexGrid.SelectionMode.Row;
      g.Columns[0].Header = "ID";
      g.Columns[1].Header = "サーバー名";
      g.Columns[2].Header = "ＯＳ";
      g.Columns[3].Header = "Ver";
      g.Columns[4].Header = "用途";
      g.Columns[5].Header = "設置場所";
      for (int i = 0; i < g.Columns.Count(); ++i)
      {
        g.Columns[i].HeaderHorizontalAlignment = HorizontalAlignment.Center;
        g.Columns[i].HeaderFontSize = 14;
        g.Columns[i].HeaderBackground = new SolidColorBrush(Colors.PapayaWhip);
      }
      g.Columns[0].Width = new GridLength(60);
      g.Columns[1].Width = new GridLength(140);
      g.Columns[2].Width = new GridLength(160);
      g.Columns[3].Width = new GridLength(50);
      g.Columns[4].Width = new GridLength(150);
      g.Columns[5].Width = new GridLength(150);
      //int c = 60;
      for (int j = 0; j < g.Rows.Count; ++j)
      {
        g.RowHeaders[j, 0] = (j + 1).ToString();
        g.Rows[j].HeaderHorizontalAlignment = HorizontalAlignment.Center;
      }
    }




    //***************************************************************************************************
    //作成モジュール end

    //コントロールのイベント begin
    //***************************************************************************************************
    private void cmbOS_DropDownClosed(object sender, EventArgs e)//cmbOS閉じるとき
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
    private void cmbUSETYPE1_DropDownClosed_1(object sender, EventArgs e)//cmbUSETYPE1閉じるとき
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
    private void btSRCH_Click(object sender, RoutedEventArgs e)//btSRCHクリック
    {
      setgrSvSrch();
    }
    private void grSVSRCH_MouseDoubleClick(object sender, MouseButtonEventArgs e)//検索グリッドgrSVSRCHnoのダブルクリック
    {
      getSetrowData(sender, e);
    }
    private void getSetrowData(object sender, MouseButtonEventArgs e)//検索グリッド行データの取得
    {
      C1FlexGrid g = grSVSRCH;
      var ht = g.HitTest(e);
      int r = ht.Row;
      MainWindow.gsvid = g[r, 0].ToString();
      this.Close();
    }
    private void btExit_Click(object sender, RoutedEventArgs e)//終了ボタンbtExitクリック
    {
      this.Close();
    }
    private void btCancel_Click(object sender, RoutedEventArgs e)//取り消しボタンbtCancelクリック
    {
      txtOSID.Text = null;
      cmbOS.SelectedIndex = -1;
      txtUSETYPE1ID.Text = null;
      cmbUSETYPE1.SelectedIndex = -1;
      txtSvName.Text = null;
      grSVSRCH.DataContext = null;
    }


    //***************************************************************************************************
    //コントロールのイベント end


  }
}
