   M S B A M L   `   `   `  ���� D  @apgc65wpf, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null%clr-namespace:apgc65wpf	apgc65wpf      apgc65wpf.MainWindow   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     )localclr-namespace:apgc65wpf     ^ ZC1.WPF.FlexGrid.4, Version=4.0.20162.524, Culture=neutral, PublicKeyToken=2aa4ec5576d6c3ceU QC1.WPF.4, Version=4.0.20162.524, Culture=neutral, PublicKeyToken=2aa4ec5576d6c3ceZ VC1.WPF.Tile.4, Version=4.0.20162.524, Culture=neutral, PublicKeyToken=2aa4ec5576d6c3ce^ ZC1.WPF.DataGrid.4, Version=4.0.20162.524, Culture=neutral, PublicKeyToken=2aa4ec5576d6c3cee	 aC1.WPF.DateTimeEditors.4, Version=4.0.20162.524, Culture=neutral, PublicKeyToken=2aa4ec5576d6c3cea
 ]C1.WPF.RichTextBox.4, Version=4.0.20162.524, Culture=neutral, PublicKeyToken=2aa4ec5576d6c3ceBc1/http://schemas.componentone.com/winfx/2006/xaml     	 
 -     � Title$$  サーバー管理システム��$	��980��$
��1260�� � Icon$ Resources/titlelogo.ico�����	#FF9FBCF0$
��Left=�$	��Top=� � WindowStartupLocation$ Manual=�
 � Top$ 0�� � Left$ 0��
���   �$��	Meiryo UI)�.��� -    �Name$ MainGrid��$��	0,0,0,-52q����	#FF9FBCF0$
��Left=�$	��Top=�$	��941��$
��1240��.+�� -   $ grSUBJ��$��52��$	��Top=�$��
1,1,1027,0q�$	��205��$
��Left=� J� Effect 3  -System.Windows.Media.Effects.DropShadowEffect  .+��� -    ��Name$ Image11_png��$��2,2,0,0q�$��Image11.png��$
��None=�$
��Left=�$��47��$	��Top=��� -    ��Name$ 	textBlock��	 �� TextAlignment$	 Center=�$
��Bold"�$
��Left=�$��	211,5,0,0q�$
��Wrap=�$��操作モード��$	��Top=�$��88�����	#FF68FB23   �� -   
 ��Name$
 
bdEditMode�����	#FF787474$��2,2,3,2q�$
��Left=�$��28��$��
211,25,0,0q�$	��Top=�$��88��   ._��� -   $ txtEditMode��$
��Wrap=�$	 Center=� �� Padding$ 0,0,0,0q�$��	Meiryo UI)�$��16(����	#FF210778���	#FFFDFDF9�� -   $ tbSB0��$	 Right=�$
��Left=�$��	10,58,0,0q�$
��Wrap=�$��サーバーID：��$	��Top=�$��22��$��70��$
��Bold"�$��	Meiryo UI)�$
��11.5(��� -	    ��Name$ txtSvId�� �� TextAlignment$ Center=�$
��Left=�$��0��$��22��$��	85,58,0,0q� �� TextWrapping$
 Wrap=�$	��Top=�$��68�� �� IsInputMethodEnabled .   �� PreferredImeState$	 Off=�$��	MS Gothic)�$��4,4,2,1q�.��   .���� -
   $ tbSB1��$	 Right=�$
��Left=�$��
200,58,0,0q�$
��Wrap=�$��サーバー名：��$	��Top=�$��22��$��65��$��	Meiryo UI)�$
��Bold"�$
��11.5(��� -   $ 	txtSvName��$
��Left=�$��0��$��22��$��
270,58,0,0q�$
 Wrap=�$	��Top=�$	��195�� .  $	 Off=�$��	MS Gothic)�$��4,4,2,1q�.��   .���� -   $ tbSB2��$	 Right=�$
��Left=�$��
470,58,0,0q�$
��Wrap=�$��OS：��$	��Top=�$��22��$��33��$��	Meiryo UI)�$
��Bold"�$
��11.5(��� -   $ txtOSID��$
��Left=�$��Center=�$��22��$��
508,58,0,0q�$
 Wrap=�$	��Top=�$��32�����	#FFF4D3FD �� 
IsReadOnly . -  'System.Windows.Input.KeyboardNavigation   TabNavigation$
 None=���.  $��	Meiryo UI)�.��   .���� -    ��Name$ cmbOS��$
��Left=�$��
545,58,0,0q�$	��Top=�$	��269��$��	MS Gothic)�$��22������    ���� .R���  �� Orientation$ 
Horizontal=�.+��� �� ��� *osid+ �� StringFormat$ d4�� | �� ��� *osname+�� -   $ tbSB3��$	 Right=�$
��Left=�$��
825,58,0,0q�$
��Wrap=�$��Ver：��$	��Top=�$��22��$��38��$��	Meiryo UI)�$
��Bold"�$
��11.5(��� -   $ txtVer��$
��Left=�$��22��$��
868,58,0,0q�$
 Wrap=�$	��Top=�$��60��$��Center=�$	 Off=�$��	MS Gothic)�$��0,4,2,0q�$ Center=�.��   .���� -   $ tbSB4��$	 Right=�$
��Left=�$��
932,58,0,0q�$
��Wrap=�$��ED：��$	��Top=�$��22��$��30��$��	Meiryo UI)�$
��Bold"�$
��11.5(��� -   $ txtEDID��$
��Left=�$��Center=�$��22��$��
967,58,0,0q�$
 Wrap=�$	��Top=�$��32�����	#FFF4D3FD . $
 None=���.  $��	Meiryo UI)�.��   .���� -   $ cmbED��$
��Left=�$��1004,58,0,0q�$	��Top=�$	��226��$��22��$��	MS Gothic)�����    ���� .R��� $ 
Horizontal=�.+��� �� ��� *osedid+$ d4���� $	� | ���� ��� *
osedname+�� -   $ tbSB5��$	 Right=�$
��Left=�$��	10,85,0,0q�$
��Wrap=�$#��PC・マザーベンダー：��$	��Top=�$��22��$	��106��$��	Meiryo UI)�$
��Bold"�$
��11.5(��� -   $ txtVDID��$
��Left=�$��Center=�$��22��$��
121,85,0,0q�$
 Wrap=�$	��Top=�$��32�����	#FFF4D3FD . $
 None=���.  $��	Meiryo UI)�.��   .���� -   $ cmbVD��$
��Left=�$��
158,85,0,0q�$	��Top=�$	��165��$��22��$��	MS Gothic)�����    ���� .R��� $ 
Horizontal=�.+��� �� ��� *vdid+$ d4���� $	� | ���� ��� *vdname+�� -   $ tbSB6��$	 Right=�$
��Left=�$��
328,85,0,0q�$
��Wrap=�$��	型式：��$	��Top=�$��22��$��48��$��	Meiryo UI)�$
��Bold"�$
��11.5(��� -   $ txtTYPE��$
��Left=�$��22��$��
381,85,0,0q�$
 Wrap=�$	��Top=�$	��179��$��Center=�$	 Off=�$��	MS Gothic)�$��4,4,2,1q�$
 Left=�.��   .���� -   $ tbSB7��$	 Right=�$
��Left=�$��
575,86,0,0q�$
��Wrap=�$��CPUベンダー：��$	��Top=�$��22��$��80��$��	Meiryo UI)�$
��Bold"�$
��11.5(��� -   $ 
txtCPUVDID��$
��Left=�$��Center=�$��22��$��
660,85,0,0q�$
 Wrap=�$	��Top=�$��32�����	#FFF4D3FD . $
 None=���.  $��	Meiryo UI)�.��   .���� -   $ cmbCPUVD��$
��Left=�$��
697,85,0,0q�$	��Top=�$	��117��$��22��$��	MS Gothic)�����    ���� .R��� $ 
Horizontal=�.+��� �� ��� *	cpuvdid+$ d4���� $	� | ���� ��� *	cpuvdname+�� -   $ tbSB8��$	 Right=�$
��Left=�$��
825,86,0,0q�$
��Wrap=�$��CPU型式：��$	��Top=�$��22��$��59��$��	Meiryo UI)�$
��Bold"�$��11(��� -   $ 
txtCPUTYPE��$
��Left=�$��22��$��
889,85,0,0q�$
 Wrap=�$	��Top=�$	��190��$��Center=�$	 Off=�$��	MS Gothic)�$��4,4,2,1q�$
 Left=�.��   .���� -   $ tbSB9��$	 Right=�$
��Left=�$��1085,85,0,0q�$
��Wrap=�$��RAM：��$	��Top=�$��22��$��42��$��	Meiryo UI)�$
��Bold"�$
��11.5(�  C1.WPF.C1NumericBox  -     Name$ txtRAM��   TextAlignment$ Center=�$
��Left=�$��1132,85,0,0q�$	��Top=�$��68��   Minimum$ 0Q�   Culture$ ja-JP��$��	MS Gothic)�   �� -    $ tbSB10��$	 Center=�$
��Left=�$��1205,86,0,0q�$
��Wrap=�$��GB��$	��Top=�$��22��$��25��$��Yu Gothic UI Semibold)�$
��Bold"��� -!   $ tbSB11��$	 Right=�$
��Left=�$��
10,112,0,0q�$
��Wrap=�$��用途１：��$	��Top=�$��22��$��57��$��	Meiryo UI)�$
��Bold"�$
��11.5(��� -"   $ txtUSETYPE1ID��$
��Left=�$��Center=�$��22��$��
72,112,0,0q�$
 Wrap=�$	��Top=�$��32�����	#FFF4D3FD . $
 None=���.  $��	Meiryo UI)�.��   .���� -#   $ cmbUSETYPE1��$
��Left=�$��109,112,0,0q�$	��Top=�$	��165��$��22��$��
MS PGothic)�����    ���� .R��� $ 
Horizontal=�.+��� �� ��� *	usetypeid+$ d4���� $	� | ���� ��� *usetypename+�� -$   $ tbSB12��$	 Right=�$
��Left=�$��286,112,0,0q�$
��Wrap=�$��用途２：��$	��Top=�$��22��$��53��$��	Meiryo UI)�$
��Bold"�$
��11.5(��� -%   $ txtUSETYPE2ID��$
��Left=�$��Center=�$��22��$��344,112,0,0q�$
 Wrap=�$	��Top=�$��32�����	#FFF4D3FD . $
 None=���.  $��	Meiryo UI)�.��   .���� -&   $ cmbUSETYPE2��$
��Left=�$��381,112,0,0q�$	��Top=�$	��165��$��22��$��
MS PGothic)�����    ���� .R��� $ 
Horizontal=�.+��� �� ��� *	usetypeid+$ d4���� $	� | ���� ��� *usetypename+�� -'   $ tbSB13��$	 Right=�$
��Left=�$��554,112,0,0q�$
��Wrap=�$��用途３：��$	��Top=�$��22��$��53��$��	Meiryo UI)�$
��Bold"�$
��11.5(��� -(   $ txtUSETYPE3ID��$
��Left=�$��Center=�$��22��$��612,112,0,0q�$
 Wrap=�$	��Top=�$��32�����	#FFF4D3FD . $
 None=���.  $��	Meiryo UI)�.��   .���� -)   $ cmbUSETYPE3��$
��Left=�$��649,112,0,0q�$	��Top=�$	��165��$��22��$��
MS PGothic)�����    ���� .R��� $ 
Horizontal=�.+��� �� ��� *	usetypeid+$ d4���� $	� | ���� ��� *usetypename+�� -*   $ tbSB14��$	 Right=�$
��Left=�$��819,112,0,0q�$
��Wrap=�$��設置場所：��$	��Top=�$��22��$��65��$��	Meiryo UI)�$
��Bold"�$
��11.5(��� -+   $ 
txtPlaceID��$
��Left=�$��Center=�$��22��$��889,112,0,0q�$
 Wrap=�$	��Top=�$��42�����	#FFF4D3FD . $
 None=���.  $��	Meiryo UI)�.��   .���� -,   $ cmbPlace��$
��Left=�$��936,112,0,0q�$	��Top=�$	��230��$��22��$��	Meiryo UI)�����    ���� .R��� $ 
Horizontal=�.+��� �� ��� *	placeid+$ d4���� $	� | ���� ��� *plname+�� --   $ tbSB15��$	 Right=�$
��Left=�$��
10,139,0,0q�$
��Wrap=�$��	仮想：��$	��Top=�$��22��$��57��$��	Meiryo UI)�$
��Bold"�$
��11.5(��� -.    ��Name$ chkVM��$�� ��$
��Left=�$��
72,140,0,0q�$	��Top=��� -/   $ tbSB16��$	 Right=�$
��Left=�$��
97,139,0,0q�$
��Wrap=�$��HPVS：��$	��Top=�$��22��$��60��$��	Meiryo UI)�$
��Bold"� J� RenderTransformOrigin$ 0.528,1.045&�$
��11.5(��� -0   $ 	txtHPVSID��$
��Left=�$��Center=�$��22��$��162,139,0,0q�$
 Wrap=�$	��Top=�$��32�����	#FFF4D3FD . $
 None=���.  $��	Meiryo UI)�.��   .���� -1   $ cmbHPVS��$
��Left=�$��199,139,0,0q�$	��Top=�$	��165��$��22��$��	MS Gothic)�����    ���� .R��� $ 
Horizontal=�.+��� �� ��� *cpid+$ d4���� $	� | ���� ��� *cpname+�� -2   $ tbSB17��$	 Right=�$
��Left=�$��369,139,0,0q�$
��Wrap=�$��
AdminID：��$	��Top=�$��22��$��66��$��	Meiryo UI)�$
��Bold"�$
��11.5(��� -3   $ 
txtAdminID��$
��Left=�$��22��$��440,139,0,0q�$
 Wrap=�$	��Top=�$	��106��$��Center=�$	 Off=�$��MS UI Gothic)�$��0,4,2,0q�$ Center=�.��   .���� -4   $ tbSB18��$	 Right=�$
��Left=�$��558,139,0,0q�$
��Wrap=�$��AdminPass：��$	��Top=�$��22��$��86��$��	Meiryo UI)�$
��Bold"�$
��11.5(��� -5   $ txtAdminPASS��$
��Left=�$��22��$��649,139,0,0q�$
 Wrap=�$	��Top=�$	��165��$��Center=�$	 Off=�$��MS UI Gothic)�$��0,4,2,0q�$ Center=�.��   .���� -6   $ tbSB19��$
��Left=�$��	8,166,0,0q�$
��Wrap=�$��	備考：��$	��Top=�$��22��$��59��$	 Right=�$��	Meiryo UI)�$
��Bold"�$
��11.5(��� -7   $ txtNote��$
��Left=�$��21��$��
72,166,0,0q�$
 Wrap=�$	��Top=�$
��1158��$ On=�! �� PreferredImeConversionMode$ FullShape,Native=�$��	Meiryo UI)�.��   .��� $
��Left=�$	��351��$��
10,203,0,0q�$	��Top=�$��38��$ 0.212,0.417&����	#FF166428   .+��� -8   $ Image12_png��$��Image12.png��$
��None=�$��
3,10,0,-12q�$
��Left=�$	��Top=�$	��296��� -9   $ grNIC��$
��Left=�$��350.5��$��
48,203,0,0q�$	��Top=�$
��1182�����	#FFD7E2C3    � RowDefinitions �� $	��41* ��� .+��� -:   $ tbSBGRNIC00��$	 Right=�$
��Left=�$��0,10,0,0q�$
��Wrap=�$��NicID：��$	��Top=�$��22��$��	Meiryo UI)�$
��Bold"�$��74��$
��11.5(��� -;   $ txtNICID��$ Center=�$��0,0,0,0q�$
��Left=�$��Center=�$��22��$��	79,10,0,0q�$
 Wrap=�$	��Top=�$��61�����	#FFF4D3FD . $
 None=���.  $��	Meiryo UI)�.��   .���� -<   $ 
tbSBGRNIC1��$
��Left=�$��
145,10,0,0q�$
��Wrap=�$��MAC：��$	��Top=�$��22��$��50��$	 Right=�$��	Meiryo UI)�$
��Bold"�$
��11.5(�  C1.WPF.C1MaskedTextBox  -=     Name$ 
txtMainMac��$ Center=�$
��Left=�$��
200,10,0,0q�$
 Wrap=�$�� ��$	��Top=�$��22��$	��150��   
PromptChar$ A��    Mask$  &&:&&:&&:&&:&&:&&��$	 Off=�$��	Meiryo UI)�$��13(�.��   .���� ->   $ 
tbSBGRNIC2��$
��Left=�$��
365,10,0,0q�$
��Wrap=�$��主IP：��$	��Top=�$��22��$��59��$	 Right=�$��	Meiryo UI)�$
��Bold"�$
��11.5(�  -?   $ 	txtMainIP��$ Center=�$
��Left=�$��
429,10,0,0q�$
 Wrap=�$�� ��$	��Top=�$��22��$	��150��$ 9��$  ###.###.###.###��$��13(�$	 Off=�$��	Meiryo UI)�.��   .���� -@   $ 
tbSBGRNIC3��$
��Left=�$��
592,10,0,0q�$
��Wrap=�$��主SubNet：��$	��Top=�$��22��$��86��$	 Right=�$��	Meiryo UI)�$
��Bold"�$
��11.5(�  -A   $ txtMainSubnet��$ Center=�$
��Left=�$��
683,10,0,0q�$
 Wrap=�$�� ��$	��Top=�$��22��$	��150��$ 9��$  ###.###.###.###��$��13(�$	 Off=�$��	Meiryo UI)�.��   .���� -B   $ 
tbSBGRNIC4��$
��Left=�$��
838,10,0,0q�$
��Wrap=�$��主GW：��$	��Top=�$��22��$��76��$	 Right=�$��	Meiryo UI)�$
��Bold"�$
��11.5(�  -C   $ 	txtMainGW��$ Center=�$
��Left=�$��
919,10,0,0q�$
 Wrap=�$�� ��$	��Top=�$��22��$	��150��$ 9��$  ###.###.###.###��$��13(�$	 Off=�$��	Meiryo UI)�.��   .���� -D   $ tbSBGRNIC5b��$
��Left=�$��0,37,0,0q�$
��Wrap=�$��	副IP1：��$	��Top=�$��22��$��74��$	 Right=�$��	Meiryo UI)�$
��Bold"�$
��11.5(�  -E   $ 	txtSUBIP1��$ Center=�$
��Left=�$��	81,37,0,0q�$
 Wrap=�$�� ��$	��Top=�$��22��$	��150��$ 9��$  ###.###.###.###��$��13(�$	 Off=�$��	Meiryo UI)�.��   .���� -F   $ tbSBGRNIC6b��$
��Left=�$��
253,37,0,0q�$
��Wrap=�$��	副IP2：��$	��Top=�$��22��$��59��$	 Right=�$��	Meiryo UI)�$
��Bold"�$
��11.5(�  -G   $ 	txtSUBIP2��$ Center=�$
��Left=�$��
317,37,0,0q�$
 Wrap=�$�� ��$	��Top=�$��22��$	��150��$ 9��$  ###.###.###.###��$��13(�$	 Off=�$��	Meiryo UI)�.��   .���� -H   $ tbSBGRNIC7b��$
��Left=�$��
493,37,0,0q�$
��Wrap=�$��Nicデバイス名：��$	��Top=�$��22��$��94��$	 Right=�$��	Meiryo UI)�$
��Bold"�$
��11.5(��� -I   $ txtIFDEVNAME��$
��Left=�$��22��$��
592,37,0,0q�$
 Wrap=�$	��Top=�$	��100��$	 Off=�$��	Meiryo UI)�.��   .���� -J   $ tbSBGRNIC7c��$
��Left=�$��
697,37,0,0q�$
��Wrap=�$��(eth0 enp0s1など)��$	��Top=�$��22��$	��119��$
	 Left=�$��	Meiryo UI)�$
��Bold"�$
��11.5(��� -K   $ tbSBGRNIC8b��$	 Right=�$
��Left=�$��
821,37,0,0q�$
��Wrap=�$��CPID：��$	��Top=�$��22��$��	Meiryo UI)�$
��Bold"�$��55��$
��11.5(��� -L   $ txtCPID��$
��Left=�$��Center=�$��22��$��
881,37,0,0q�$
 Wrap=�$	��Top=�$��61�����	#FFF4D3FD . $
 None=���.  .��   .���� -M   $ tbSBGRNIC9b��$	 Right=�$
��Left=�$��
947,37,0,0q�$
��Wrap=�$��	並番：��$	��Top=�$��22��$��	Meiryo UI)�$
��Bold"�$��75��$
��11.5(��� -N   $ txtSECNO��$
��Left=�$��Center=�$��22��$��1027,37,0,0q�$
 Wrap=�$	��Top=�$��42�����	#FFF4D3FD . $
 None=���.  .��   .���� -O   $ tbSBGRNIC10b��$
��Left=�$��0,66,0,0q�$
��Wrap=�$��オンボード：��$	��Top=�$��22��$��74��$	 Right=�$��	Meiryo UI)�$
��Bold"�$
��11.5(��� -P   $ 
chkONBOARD��$�� ��$
��Left=�$��	79,66,0,0q�$	��Top=�$��15��$��20���� -Q   $ tbSBGRNIC11b��$
��Left=�$��
104,66,0,0q�$
��Wrap=�$��管理用：��$	��Top=�$��22��$��78��$	 Right=�$��	Meiryo UI)�$
��Bold"�$
��11.5(��� -R   $ chkNICMANAGE��$�� ��$
��Left=�$��
187,66,0,0q�$	��Top=�$��15��$��20���� -S   $ tbSBGRNIC12b��$	 Right=�$
��Left=�$��
239,64,0,0q�$
��Wrap=�$��Nicベンダー：��$	��Top=�$��22��$��	Meiryo UI)�$
��Bold"�$��74��$
��11.5(��� -T   $ 
txtNICVDID��$
��Left=�$��Center=�$��22��$��
318,64,0,0q�$
 Wrap=�$	��Top=�$��32�����	#FFF4D3FD . $
 None=���.  $��	Meiryo UI)�.��   .���� -U   $ cmbNicVendor��$
��Left=�$��
355,64,0,0q�$	��Top=�$	��155��$��22��$��	Meiryo UI)�����    ���� .R��� $ 
Horizontal=�.+��� �� ��� *	nicvdid+$ d4���� $	� | ���� ��� *	nicvdname+�� -V   $ tbSBGRNIC13b��$
��Left=�$��
528,64,0,0q�$
��Wrap=�$��	型番：��$	��Top=�$��22��$��59��$	 Right=�$��	Meiryo UI)�$
��Bold"�$
��11.5(��� -W   $ 
txtNICTYPE��$
��Left=�$��22��$��
592,64,0,0q�$
 Wrap=�$	��Top=�$	��284��$	 Off=�.��   .���� -X   $ tbSBGRNIC14b��$
��Left=�$��	31,93,0,0q�$
��Wrap=�$��	備考：��$	��Top=�$��22��$��43��$	 Right=�$��	Meiryo UI)�$
��Bold"�$
��11.5(��� -Y   $ 
txtNicNote��$
��Left=�$��22��$��	81,91,0,0q�$
 Wrap=�$	��Top=�$	��795��$ On=�$ FullShape,Native=�.��   .���� -Z   ! ��Name$! btNicAdd��$��	行追加��$
��Left=�$��1102,10,0,0q�$	��Top=�$��70��$��25��$��14(�$
��Bold"����	#FF6AEC9F$��	Meiryo UI)�   �� -[   $! btNicDelete��$��	行削除��$
��Left=�$��1102,40,0,0q�$	��Top=�$��70��$��25��$��14(�$
��Bold"����	#FFEAA9B5$��	Meiryo UI)�   �� -\   $! 
btNicRenew��$��更新��$
��Left=�$��1102,70,0,0q�$	��Top=�$��70��$��25��$��14(�$
��Bold"����	#FF94CFDC$��	Meiryo UI)�      C1.WPF.FlexGrid.C1FlexGrid  -]   "  Name$" grNic��#   
IsReadOnly# . $   AllowSorting$ .  $
��Left=�$��
10,139,0,0q�$	��Top=�$	��205��$
��1162��%   ShowMarquee% . &   RowBackground& �	#FFBDD0F9'   AlternatingRowBackground
' ���� (   CursorBackground
( ���� )   EditorBackground
) ���� *   TopLeftCellBackground* �	#FFC9C9C9+   RowHeaderBackground+ �	#FFEDEDED",   RowHeaderSelectedBackground, �	#FFFFD58D-   GridLinesBrush- �	#FF4274D6.   HeaderGridLinesBrush. �	#FFB6B6B6/   SelectionBackground/ �	#FFEAECF5$��MS UI Gothic)�0   ItemsSource0 �� %1   ColumnHeaderSelectedBackground1 �� 2 �� EndPoint$	2 0,1&�.��� 3 � Color$3 	#FFF8D79B��4 � Offset$4 0Q�� $3 	#FFF1C15F��$4 1Q�5   ColumnHeaderBackground5 �� $	2 0,1&�.��� $3 	#FFF6F6F6��$4 0Q�� $3 	#FFDEDEDE��$4 1Q��� -^   $! 
btNicClear��$��	初期化��$
��Left=�$��1102,100,0,0q�$	��Top=�$��70��$��25��$��14(�$
��Bold"����	#FFF5F7C2$��	Meiryo UI)�   � $
��Left=�$��19��$��	858,5,0,0q�$	��Top=�$	��382�����	#FF707AE0   .+��� -_   $ 
txtMachine��$	 Center=�$
��Left=�$��0,1,0,0q�$
��Wrap=�$	��Top=�$	��118�����	#FFFFF15A$��16��$��	Meiryo UI)�$��11(�$ 0,1,0,0q�$
��Bold"��� -`   6 ��Name$6 txtUser��$�� ��$
��Left=�$	��5,0q�$��	242,0,0,0q�$	��Top=�$��1q�7 J� OpacityMask
7 �   �$	��140��$��18�����	#FFFFF15A$��	Meiryo UI)�$
��Bold"�$��11(�&   System.Windows.Media.TextOptions8   TextFormattingMode$8 Display=��� -a   $6 txtYMD��$�� ��$
��Left=�$	��5,0q�$��	120,0,0,0q�$	��Top=�$��1q�
7 �   �$	��122��$��18�����	#FFEEEBCB$��	Meiryo UI)�$
��Bold"�$��11(�$8 Display=�� $
��Left=�$��24��$��859,29,-1,0q�$	��Top=�$	��382��.+��� -b   $6 lbActSv��$��	存在中��$
��Left=�$	��Top=�$��40��$��Center=����	#FFADF7CB���	#FF0C0A01$��22��$
��Bold"�$��0,3,0,0q�   �� -c   $6 txtActSv��$�� ��$
��Left=�$	��Top=�$��35��$��Center=����	#FFF7F4AD���	#FF0C0A01$��22��$
��Bold"�$��0,3,0,0q�$��44,0,0,0q��� -d   $6 
lbActSvDai��$	��台��$
��Left=�$��80,1,0,0q�$	��Top=�$��16��$��22��$
��Bold"�$��0,3,0,0q��� -e   $6 lbDelSv��$��	撤去済��$
��Left=�$	��Top=�$��40��$��Center=����	#FFADF7CB���	#FF0C0A01$��22��$
��Bold"�$��0,3,0,0q�$��97,0,0,0q�   �� -f   $6 txtDelSv��$�� ��$
��Left=�$	��Top=�$��35��$��Center=����	#FFF7F4AD���	#FF0C0A01$��22��$
��Bold"�$��0,3,0,0q�$��	141,0,0,0q��� -g   $6 lbActSvDai2��$	��台��$
��Left=�$��	177,1,0,0q�$	��Top=�$��16��$��22��$
��Bold"�$��0,3,0,0q��� -h   $6 lbRealSv��$��実機��$
��Left=�$	��Top=�$��35��$��Center=����	#FFF7ADCB���	#FF0C0A01$��22��$
��Bold"�$��0,3,0,0q�$��	196,0,0,0q�   �� -i   $6 	txtRealSv��$�� ��$
��Left=�$	��Top=�$��35��$��Center=����	#FFF7F4AD���	#FF0C0A01$��22��$
��Bold"�$��0,3,0,0q�$��	235,0,0,0q��� -j   $6 lbRealSVDai��$	��台��$
��Left=�$��	271,1,0,0q�$	��Top=�$��16��$��22��$
��Bold"�$��0,3,0,0q��� -k   $6 lbVmSv��$��仮想��$
��Left=�$	��Top=�$��35��$��Center=����	#FFF7ADCB���	#FF0C0A01$��22��$
��Bold"�$��0,3,0,0q�$��	290,0,0,0q�   �� -l   $6 txtVmSv��$�� ��$
��Left=�$	��Top=�$��35��$��Center=����	#FFF7F4AD���	#FF0C0A01$��22��$
��Bold"�$��0,3,0,0q�$��	329,0,0,0q��� -m   $6 lbVmDai��$	��台��$
��Left=�$��	364,1,0,0q�$	��Top=�$��16��$��22��$
��Bold"�$��0,3,0,0q��� -n   $! button1��$��button1��$
��Left=�$��739,876,0,0q�$	��Top=�$��75���� -o   $ textBoxTest��$
��Left=�$��56��$��356,876,0,0q�$
 Wrap=�$��TextBox��$	��Top=�$	��374���� -p   $! btExit��$
��Left=�$��	318,5,0,0q�$	��Top=�$��85��$��45�����	#FF54CDE8   .���� -q   $ Image28_png��$��42��$��Image28.png��$
��Fill=�$��83���� -r   $! btSave��$
��Left=�$��	421,5,0,0q�$	��Top=�$��85��$��45�����	#FF1EB879   .���� -s   $ Image16_png��$��42��$��Image16.png��$
��Fill=�$��83���� -t   $! btCancel��$
��Left=�$��	526,5,0,0q�$	��Top=�$��85��$��45�����	#FFEC6363   .���� -u   $ Image15_png��$��Image15.png��$
��Fill=�$��42��$��83���� -v   $! btDelete��$
��Left=�$��	632,5,0,0q�$	��Top=�$��85��$��45�����	#FFA29797   .���� -w   $ Image27_png��$��42��$��Image27.png��$
��Fill=�$��83���� -x   $! 
btSVIDHELP��$
��Left=�$��
153,58,0,0q�$	��Top=�$��37��$��22�����	#FF7286E2   .���� -y   $ Image26_png��$��18��$��Image26.png��$
��None=�$��34���� -z   $! 
btVMHPVEXP��$
��Left=�$��	739,5,0,0q�$	��Top=�$��85��$��45�����#FF6633   .���� -{   $ Image30_png��$��42��$��Image30.png��$
��Fill=�$��83���� -|   $! btgrNicInit��$��brgrNicInit(消すな)��$
��Left=�$��825,876,0,0q�$	��Top=�$	��109��$��33���� -}   $! btgrNicEditInit��$��btgrNicEditInit��$
��Left=�$��1004,33,0,0q�$	��Top=�$��75��$y�	Collapsed=��� -~   $! button5��$��button5��$
��Left=�$��739,900,0,0q�$	��Top=�$��75���� -   $ textBoxTest2��$
��Left=�$��56��$��
48,876,0,0q�$
 Wrap=�$��TextBox��$	��Top=�$	��303��  -�   $" grSVALL��# . $
��Left=�$��
48,561,0,0q�$	��Top=�$	��294��$
��1182��$��MS UI Gothic)�9   ClipboardCopyMode$9 ExcludeHeader=�0 ��    �� -�   $! btShowSvAll��$��btShowSvAll(消すな)��$
��Left=�$��939,876,0,0q�$	��Top=�$	��121��$��33���� -�   $! btXlsExport��$
��Left=�$��	9,561,0,0q�$	��Top=�$��38��$	��294�����	#FFC5AFAF   .���� -�   $ Image29_png��$	��272��$��Image29.png��$
��None=�$��33��  -�   $" 
grVMHPVALL��# . $
��Left=�$��
48,937,0,0q�$	��Top=�$��85��$	��626��$��MS UI Gothic)�0 �� �� -�   $! btShowVmHpvAll��$��Button��$
��Left=�$��1065,876,0,0q�$	��Top=�$	��125��$��33��