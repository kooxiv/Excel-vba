# Excel-vba

Excel 宏
=========

- 批量替换公式 示例为 批量添加 IFError(原公式,0)

       Sub addIFError0()

        '示例为选中的单元格批量添加 IFError

        For Each Item In Selection.Cells '循环 选中的单元格
        Dim str As String

        str = Item.Formula '获取原公式 如 =1+2
        str = Replace(str, "=", "") '去掉原公式的=号 转为: 1+2

        Item.FormulaR1C1 = "=IFError(" + str + ",0)" '将单元格赋值为新公式 示例为 IFError(原公式,0) 同样可以换成其它公式
        Next

      End Sub

>
