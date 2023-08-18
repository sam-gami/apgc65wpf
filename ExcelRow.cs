using Microsoft.VisualBasic;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using C1.WPF.FlexGrid;
using C1.WPF.Excel;
using System.Globalization;
using System.Windows;

namespace apgc65wpf
{
  /// <summary>
  /// 編集可能、ツリー ノードとして使用でき、各列と関連するセルの
  /// スタイルのコレクションを保守するグリッド行
  /// </summary>
  public class ExcelRow : GroupRow
  {
    // ** フィールド
    private Dictionary<Column, CellStyle> _cellStyles;
    // 既定の有効桁数を6桁に設定します
    private const string DEFAULT_FORMAT = "#,##0.######";
// ** ctor
public ExcelRow(ExcelRow styleRow)
    {
      IsReadOnly = false;
      if (styleRow != null && styleRow.Grid != null)
      {
        foreach (var c in styleRow.Grid.Columns)
        {
          dynamic cs = styleRow.GetCellStyle(c);
          if (cs != null)
          {
            this.SetCellStyle(c, cs.Clone());
          }
        }
      }
    }
    public ExcelRow()
    : this(null)
    {
    }
    // ** オブジェクト・モデル
    /// <summary>
    /// データを取得する場合、書式を適用するためにオーバーライドされる
    /// </summary>
    public override string GetDataFormatted(Column col)
    {
      // データを取得します
      dynamic data = GetDataRaw(col);
      // 書式を適用します
      dynamic ifmt = data as IFormattable;
      if (ifmt != null)
      {
        // セルの書式を取得します
        dynamic s = GetCellStyle(col) as ExcelCellStyle;
        dynamic fmt = s != null && (!string.IsNullOrEmpty(s.Format)) ?
        s.Format : DEFAULT_FORMAT;
        data = ifmt.ToString(fmt, CultureInfo.CurrentUICulture);
      }
      // 完了
      return data != null ? data.ToString() : string.Empty;
    }
    // ** オブジェクト・モデル
    /// <summary>
    /// この行では、セルにスタイルを適用します
    /// </summary>
    public void SetCellStyle(Column col, CellStyle style)
    {
      if (!object.ReferenceEquals(style, GetCellStyle(col)))
      {
        if (_cellStyles == null)
          {
          _cellStyles = new Dictionary<Column, CellStyle>();
        }
        _cellStyles[col] = style;
        if (Grid != null)
        {
          Grid.Invalidate(new CellRange(this.Index, col.Index));
        }
      }
    }
    /// <summary>
    /// この行では、セルに適用したスタイルを取得します
    /// </summary>
    public CellStyle GetCellStyle(Column col)
    {
      CellStyle s = null;
      if (_cellStyles != null)
      {
        _cellStyles.TryGetValue(col, out s);
      }
      return s;
    }
  }
}
