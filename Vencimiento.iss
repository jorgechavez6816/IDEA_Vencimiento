Sub Main
	Call Aging()	'Ejemplo-Detalle de ventas.IMD
End Sub


' Análisis: Vencimiento
Function Aging
	Set db = Client.OpenDatabase("Ejemplo-Detalle de ventas.IMD")
	Set task = db.Aging
	task.Info "2015/12/31", "FECHA_FACT", "TOTAL"
	task.IntervalTypeIndex = 0
	task.Intervals 30, 60, 90, 120, 150, 180
	dbName = "Vencimiento_01.IMD"
	task.CreateAgeDB dbName, ""
	task.AddFieldToInc "FECHA_FACT"
	task.AddFieldToInc "TOTAL"
	task.CreateVirtualDatabase = False
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function