   M S B A M L   `   `   `  ���� D  @apgc65wpf, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null%clr-namespace:apgc65wpf	apgc65wpf      apgc65wpf.Whlpsvid   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     )localclr-namespace:apgc65wpf     ^ ZC1.WPF.FlexGrid.4, Version=4.0.20162.524, Culture=neutral, PublicKeyToken=2aa4ec5576d6c3ceU QC1.WPF.4, Version=4.0.20162.524, Culture=neutral, PublicKeyToken=2aa4ec5576d6c3ceZ VC1.WPF.Tile.4, Version=4.0.20162.524, Culture=neutral, PublicKeyToken=2aa4ec5576d6c3ce^ ZC1.WPF.DataGrid.4, Version=4.0.20162.524, Culture=neutral, PublicKeyToken=2aa4ec5576d6c3cee	 aC1.WPF.DateTimeEditors.4, Version=4.0.20162.524, Culture=neutral, PublicKeyToken=2aa4ec5576d6c3cea
 ]C1.WPF.RichTextBox.4, Version=4.0.20162.524, Culture=neutral, PublicKeyToken=2aa4ec5576d6c3ceBc1/http://schemas.componentone.com/winfx/2006/xaml     	 
 -     �Name$  
HelpWindow�� � Title$ サーバー検索��$	��510��$	��821�� � Icon$ Resources/titlelogo.ico��$��5q� � 
ResizeMode$ NoResize=�.��� ���	#FFA0E0D1$��-6,-7,-8,-8q�.+��� -    ��Name$ tbSB2�� �� TextAlignment$ Right=�$
��Left=�$��
180,10,0,0q�$
��Wrap=�$��OS：��$	��Top=�$��22��$��70��$��Yu Gothic UI Semibold)�$
��Bold"��� -    ��Name$ txtOSID��$
��Left=�$��Center=�$��22��$��
255,10,0,0q� �� TextWrapping$
 Wrap=�$	��Top=�$��32�����	#FFF4D3FD �� 
IsReadOnly . -  'System.Windows.Input.KeyboardNavigation	   TabNavigation$
	 None=���.  .��
 J� Effect
 3  -System.Windows.Media.Effects.DropShadowEffect  .���� -    ��Name$ cmbOS��$
��Left=�$��
292,10,0,0q�$	��Top=�$	��260������ 
   ���� .R���  �� Orientation$ 
Horizontal=�.+��� �� ��� *osid+ �� StringFormat$ d4���� $	� | ���� ��� *osname+�� -   $ tbSB11��$ Right=�$
��Left=�$��
180,37,0,0q�$
��Wrap=�$��用途１：��$	��Top=�$��22��$��70��$��Yu Gothic UI Semibold)�$
��Bold"��� -   $ txtUSETYPE1ID��$
��Left=�$��Center=�$��22��$��
255,37,0,0q�$
 Wrap=�$	��Top=�$��32�����	#FFF4D3FD . $
	 None=���.  .��
   .���� -   $ cmbUSETYPE1��$
��Left=�$��
292,37,0,0q�$	��Top=�$	��260��$��22������ 
   ���� .R��� $ 
Horizontal=�.+��� �� ��� *	usetypeid+$ d4���� $	� | ���� ��� *usetypename+�� -   $ tbSB1��$ Right=�$
��Left=�$��
180,64,0,0q�$
��Wrap=�$��サーバー名：��$	��Top=�$��22��$��70��$��Yu Gothic UI Semibold)�$
��Bold"��� -	   $ 	txtSvName��$
��Left=�$��0��$��22��$��
255,64,0,0q�$
 Wrap=�$	��Top=�$	��297�� �� PreferredImeState$	 Off=�$��	MS Gothic)�$��4,4,2,1q�.��
   .��   C1.WPF.FlexGrid.C1FlexGrid  -
     Name$ grSVSRCH��   AllowSorting .  $
��Left=�$��	10,99,0,0q�$	��Top=�$	��366��$	��777��   ShowMarquee .    RowBackground
 ����    AlternatingRowBackground
 ����    CursorBackground
 ����    EditorBackground
 ����    TopLeftCellBackground �	#FFC9C9C9   RowHeaderBackground �	#FFEDEDED"   RowHeaderSelectedBackground �	#FFFFD58D   GridLinesBrush �	#FFD0D7E5   HeaderGridLinesBrush �	#FFB6B6B6   SelectionBackground �	#FFEAECF5   GridLinesVisibility$	 All=�   
IsReadOnly .    ItemsSource �� 
   %   ColumnHeaderSelectedBackground ��   �� EndPoint$	  0,1&�.��� ! � Color$! 	#FFF8D79B��" � Offset$" 0Q�� $! 	#FFF1C15F��$" 1Q�#   ColumnHeaderBackground# �� $	  0,1&�.��� $! 	#FFF6F6F6��$" 0Q�� $! 	#FFDEDEDE��$" 1Q��� -   $ ��Name$$ btSRCH��$
��Left=�$��	10,54,0,0q�$	��Top=�$��68��$��40��$��16(����	#FFE67B7B
   .���� -   % ��Name$% Image24_png��$��35��$��Image24.png��$
��None=�$��63���� -   $$ btExit��$
��Left=�$��	10,10,0,0q�$	��Top=�$��68��$��40�����	#FF2CAE4A
   .���� -   $% Image23_png��$��35��$��Image23.png��$
��None=�$��63���� -   $$ btCancel��$
��Left=�$��	95,32,0,0q�$	��Top=�$��68��$��40�����	#FF6888DC
   .���� -   $% Image25_png��$��35��$��Image25.png��$
��Fill=�$��63���� -   $ textBox��$
��Left=�$��32��$��	587,9,0,0q�$
 Wrap=�$��TextBox��$	��Top=�$	��200��$y�Hidden=��� -   & ��Name$& chkVM��$��仮想マシン��$
��Left=�$��
614,42,0,0q�$	��Top=�$
��Bold"�