using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Net;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Media;
using System.ComponentModel;
using C1.WPF.FlexGrid;

namespace apgc65wpf
{
  /// <summary>
  /// Excelスタイルのセル境界線と書式指定文字列を提供するためのCellStyleクラスを拡張します
  /// </summary>
  public class ExcelCellStyle : CellStyle
  {
    // ** フィールド
    private string _format;
    private Thickness _bdrThickness;
    private Brush _bdrLeft;
    private Brush _bdrTop;
    private Brush _bdrRight;
    private Brush _bdrBottom;
    private static Thickness _thicknessEmpty = new Thickness(0);
    // ** オブジェクト・モデル
    public string Format
    {
      get { return _format; }
      set
      {
        if (value != _format)
        {
          _format = value;
          OnPropertyChanged(new PropertyChangedEventArgs("Format"));
        }
      }
    }
    public Thickness CellBorderThickness
    {
      get { return _bdrThickness; }
      set
      {
        if (value != _bdrThickness)
        {
          _bdrThickness = value;
          OnPropertyChanged(new
          PropertyChangedEventArgs("BorderThickness"));
        }
      }
    }
    public Brush CellBorderBrushLeft
    {
      get { return _bdrLeft; }
      set
      {
        if (!object.ReferenceEquals(value, _bdrLeft))
        {
          _bdrLeft = value;
          OnPropertyChanged(new
          PropertyChangedEventArgs("BorderColorLeft"));
        }
      }
    }
public Brush CellBorderBrushTop
    {
      get { return _bdrTop; }
      set
      {
        if (!object.ReferenceEquals(value, _bdrTop))
        {
          _bdrTop = value;
          OnPropertyChanged(new
          PropertyChangedEventArgs("BorderColorTop"));
        }
      }
    }
    public Brush CellBorderBrushRight
    {
      get { return _bdrRight; }
      set
      {
        if (!object.ReferenceEquals(value, _bdrRight))
        {
          _bdrRight = value;
          OnPropertyChanged(new
          PropertyChangedEventArgs("BorderColorRight"));
        }
      }
    }
    public Brush CellBorderBrushBottom
    {
      get { return _bdrBottom; }
      set
      {
        if (!object.ReferenceEquals(value, _bdrBottom))
        {
          _bdrBottom = value;
          OnPropertyChanged(new
          PropertyChangedEventArgs("BorderColorBottom"));
        }
      }
    }
    // ** オーバーライド
    public override void Apply(Border bdr, SelectedState selState)
    {
      base.Apply(bdr, selState);
      ApplyBorder(bdr, _bdrLeft, new Thickness(_bdrThickness.Left, 0, 0, 0));
      ApplyBorder(bdr, _bdrTop, new Thickness(0, _bdrThickness.Top, 0, 0));
      ApplyBorder(bdr, _bdrRight, new Thickness(0, 0, _bdrThickness.Right, 0));
      ApplyBorder(bdr, _bdrBottom, new Thickness(0, 0, 0,
      _bdrThickness.Bottom));
    }
    private void ApplyBorder(Border bdr, Brush br, Thickness t)
    {
      if (br != null && t != _thicknessEmpty)
        {
          // 内部の境界線を作成します
          dynamic inner = new Border();
          inner.BorderThickness = t;
          inner.BorderBrush = br;
          // コンテンツに拡張します
          dynamic content = bdr.Child;
          bdr.Child = inner;
          inner.Child = content;
          // パディングします
          inner.Padding = bdr.Padding;
          bdr.Padding = _thicknessEmpty;
        }
    }
  }
}