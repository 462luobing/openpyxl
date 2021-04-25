# openpyxl
操作工作表

导入语句：import openpyxl

创建工作表：自定=openpyxl.Worksheet('路径')

保存工作表：自定.save('路径')[如路径是与原来的路径一样则会进改变原来的工作表后进行保存；如果是一个新路径则原来的工作表不会改变，在新的路径下创建一个新的工作表]

打开工作表：自定2=openpyxl.load_worksheet('路径',data_only=True(在运用公式算法时。此语句是否加入决定了输出时，是显示结果还是公式））[此语句执行时，工作表必须存在]

切换工作表中的指定表：自定3=自定2['工作表中的指定表']

显示工作表中全部的指定表：自定=自定2.worksheets[数字(代表切到第几个工作表，从0开始)][返回一个列表，可以使用for进行遍历，在i.title中可详细输出指定表的名称]

移除工作表中的指定表：自定2.remove(切换到工作表中的指定表，如：自定2['sheet1’])

新建工作表中的指定表：自定2.create_sheet('名称')

复制工作表中的指定表：自定=自定2.copy_worksheet(切换到工作表中的指定表，如：自定2['sheet1’])，为复制的指定表修改名称：自定.title='名称'

查询工作表中全部内容共占的行数：自定=自定2.max_row
查询工作表中全部内容共占的列数：自定=自定2.max_column

确定内容所在的行\列：自定=自定2['内容'].row\column

插入行：自定2.insert_rows(idx=代表从第几行插入，amount=插入几行)
插入列：自定2.insert_cols(idx=代表从第几列插入，amount=插入几列)

删除行：自定2.delete_rows(idx=代表从第几行删除，amount=删除几行)
删除列：自定2.delete_cols(idx=代表从第几列删除，amount=删除几列)

移动某块区域的内容到其他单元格：自定2.move_range('单元格区域',row=行,cols=列（正向下，负向上））

冻结单元格：自定2.free_panes='单元格位置'(列：A2（代表从此位置（不包括此位置）之前的全部窗口处于冻结状态）

合并单元格：自定2.merge_cells('单元格区域',start_row=起始行号，start_column=起始列号,end_row=结束行号，end_column=结束列号)
取消单元格合并:自定2.unmerge_cells('单元格区域',start_row=起始行号，start_column=起始列号,end_row=结束行号，end_column=结束列号)

行分组：row_dimensions.group('数字'第几行开始,'数字'第几行结束,hidden=True(True:进行分组后为隐藏，False:进行分组后不隐藏))
列分组：column_dimensions.group('数字'第几列开始,'数字'第几列结束,hidden=True(True:进行分组后为隐藏，False:进行分组后不隐藏)

设置行高：row_dimensions['单元格位置(数字)'].height=数字
设置列宽:column_dimensions['单元格位置(字母)'].width=数字

读取单元格内容：单格：自定=自定2['单元格(A1,B5)'].value
								区域：自定=自定2['A1:B5](字母后带有数字代表按行读取，字母后无数字代表按列读取）
											自定=自定2['单元格位置'].value
								利用iter_rows和iter_cols进行指定区域的读取：按行：自定=自定2.iter_rows(min_row=,max_row=,min_col=,max_col=)
																												  	按列：自定=自定2.iter_cols(min_row=,max_row=,min_col=,max_col=)
                                                            读取整张工作表:自定=list(自定2.values)
                                                            
 向单元格内写入内容：自定2['单元格位置']='内容'
                    自定2.cell(row=,column=,value='内容')
                    直接向工作表中添加一行内容：自定2.append（list）
           
 设置单元格内批注：自定=openpyxl.comment.Comments('内容','作者')
 									自定2['单元格位置'].comment=自定
                  
设置边框：自定j=openpuxl.styles.Side(style(线样式)='thin/double/thick',color(颜色)='十六进制')
					自定=openpyxl.styles.Border(left=自定j,right=自定j,top=自定j,bottom=自定j)
					自定2['单元格位置']border=自定

设置插入图片：自定j=openpyxl.drawing.image.Image('路径')
							设置宽高:自定j.height=数字
              自定j.width=数字
              自定2.add_image(自定j,'单元格位置')

设置字体:自定=openpyxxl.styles.Font(namd=u'字体名称'(中文时须加u在前),size=大小,bold(是否加粗)=True/False,italic(是否倾斜)=True/False,strick(是否显示删除线)=True/False,color='十六								进制',underline(下划线)='None'(默认)/single(单下划线)/double(双下划线)/singleAccounting(会计用单下划线)/doubleAccounting(会计用双下划线),vertAlign(上下标)='None'(默								认)/superscript(上标)/subscript(下标)
													         				自定2['单元格位置'].font=自定
                
设置字体对齐：自定=openpyxl.styles.Alignment(horizontal='general(常规)'/justify(两端对齐)/right(靠右)/left(靠左)/center(居中)/centerContinuous(跨列居中)/distributed(分散对  												齐)/fill(填充),vertical='center(垂直居中)'/bottom(靠下)/justify(两端对齐)/distributed(分散对齐),text_rotation=指定文本旋转角度,wrap_text(是否自动换																行)=True/False，shrink_to_fit(是否缩小字体填充),indent:指定缩进）
             							自定2['单元格位置'].alignment=自定

制作柱状图：
				        新建一个柱状图：
           自定j=openpy.chart.Barchart()
        设定数据的范围：
           自定p=openpyxl.chart.Reference(自定2，min_row,max_row,min_col,max_col)
        向柱状图中添加数据:
           自定j.add_data(自定p,titles_from_data=True/False(如果自定p中的min_row等包括x轴显示的则为True,否则为False))
        设置x轴显示:
           自定k=openpyyxl.chart.Reference(自定2，min_row,max_row,min_col,max_col(四个因x轴需要而来选择定))
        向柱状图中添加X轴:
           自定j.set_categories(自定k)
        设置柱状图长/高:自定j.height=,自定j.width=，
        导入到表格:自定2.add_chart(自定j,'单元格位置')


制作折线图：
		        新建一个折线图：
           自定j=openpy.chart.Linechart()
        设定数据的范围：
           自定p=openpyxl.chart.Reference(自定2，min_row,max_row,min_col,max_col)
        向柱状图中添加数据:
           自定j.add_data(自定p,titles_from_data=True/False(如果自定p中的min_row等包括x轴显示的则为True,否则为False),from_rows=True)
        设置x轴显示:
           自定k=openpyyxl.chart.Reference(自定2，min_row,max_row,min_col,max_col(四个因x轴需要而来选择定))
        向柱状图中添加X轴:
           自定j.set_categories(自定k)
        设置柱状图长/高:自定j.height=,自定j.width=，
        导入到表格:自定2.add_chart(自定j,'单元格位置')

数字转字母:openpyxl.utils.get_column_letter(数字)
字母转数字:openpyxl.utils.column_index_from_string('字母')











