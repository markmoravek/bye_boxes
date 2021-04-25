Attribute VB_Name = "Module4"
Sub bye_boxes()
Attribute bye_boxes.VB_ProcData.VB_Invoke_Func = " \n14"

Dim Shp As Object

For Each Shp In ActiveSheet.Shapes
Shp.Delete
Next Shp

End Sub
