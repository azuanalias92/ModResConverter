   M S B A M L   `   `   `  ����@  <SortD, Version=1.1.8.0, Culture=neutral, PublicKeyToken=nullclr-namespace:SortDSortD  5   	       SortD.MainWindow   5   	   W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   %localclr-namespace:SortD     5   	     � Title$  
MainWindow��5   	   $	��600��6   $
��1080��6)   .��5
      � 
����d�$��0,0,0,0q�6'   .+�5   
   ]� .+�5      o� $	��Top=�6   .@�5      m� $��_File��6   .@�5      m� -   5        $��_Open��5      6.   m� 5      -   5        $��_Export��5      60   5      m� 5      -   5        $��	_Settings��5      5      5      �� 5      6   5      �� 5   
   -   5         ��Name$ path1��5      $
��Left=�6!   $	��125��6<   $��	50,50,0,0q�6I   $	��Top=�6^   $	��800��6v   �� 5   
   -   5         ��Name$ comboX��5      $
��Left=�6$   $��900,50,50,0q�6?   $	��Top=�6T   $	��115��6l   6x   �� 5   
   $��	File Path��6   $
��Left=�6$   $��	50,20,0,0q�6?   $	��Top=�6R   $	��100��6j   
�������6v   �� 5   
   $��
X-location��6   $
��Left=�6%   $��
900,20,0,0q�6@   $	��Top=�6T   
�������6m   �� 5   
   $��
Y-location��6   $
��Left=�6%   $��
900,75,0,0q�6@   $	��Top=�6T   
�������6m   �� 5   
   -   5        $ comboY��5      $
��Left=�6$   $��900,100,50,0q�6?   $	��Top=�6U   $	��115��6m   6y   �� 5   
   $
��Data��6   $
��Left=�6   $��
50,200,0,0q�6:   $	��Top=�6N   $	��100��6f   
�������6r   &   System.Windows.Controls.DataGrid  5    
   -   5          Name$ 	dataGrid1��5          AutoGenerateColumns . $
��Left=�6�   $	��250��6�   $��
50,225,0,0q�6�   $	��Top=�6�   $	��965��6�      
IsReadOnly . "   VerticalScrollBarVisibility$ Visible=�6  $   HorizontalScrollBarVisibility$ Visible=�62  $��0q�6Z     	CellStyle 5!      �� 5"      	 �� 
TargetType*  $System.Windows.Controls.DataGridCell"	  6   .�5#      �� 
 �� Property
 � �� �� Value$ 10q�60   �� 5$      
 � �� 5%      �� 5&       �� 
TargetType"  6.   .R�5'   "   ��  �� Padding# z��6)   #��z��6M   #��z��6y   #��z��6�    J� SnapsToDevicePixels . ._�5(   &   �� # z 67   5)   #   5*      5,      5-      5/      �� 50   
   $��Spacing��6   $
��Left=�6"   $��900,125,0,0q�6=   $	��Top=�6R   
�������6k   �� 51   
   -   5        $ 
comboSpace��51      $
��Left=�6(   $��900,150,50,0q�6C   $	��Top=�6Y   $	��115��6q   6}   �� 52   
   -	   5         ��Name$ clearBtn��52      $��Reset��6#   $
��Left=�63   $��945,515,0,0q�6N   $	��Top=�6c   $��70��6{   
�������$��16(�6�   6�   �� 53   
   $��	SP Mode :��6   $
��Left=�6$   $��
50,510,0,0q�6?   $	��Top=�6S   
�������$��75��6   �� 54   
   -
   5        $ 
toggle_btn��54      $	��OFF��6&   $
��Left=�64   $��120,515,0,0q�6O   $	��Top=�6d   $��75��6|   
����p��6�   55      56      