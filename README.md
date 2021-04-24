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





















