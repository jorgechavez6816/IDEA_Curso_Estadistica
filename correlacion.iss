Sub Main
	IgnoreWarning(True)
	Call Summarization()	'ED Ventas-2010-L4-Ventas diarias por producto.IMD
	Call Summarization1()	'ED Ventas-2010-L4-Ventas diarias por producto.IMD
	Call Summarization2()	'ED Ventas-2010-L4-Ventas diarias por producto.IMD
	Call AppendDatabase()	'Ventas con cupon 2010-L4.IMD
	Call Correlation()	'Completas - Ventas con cupon 2010.IMD
	Client.RefreshFileExplorer
End Sub


' Análisis: Resumen
Function Summarization
	Set db = Client.OpenDatabase("ED Ventas-2010-L4-Ventas diarias por producto.IMD")
	Set task = db.Summarization
	task.AddFieldToSummarize "ID_LOCAL"
	task.AddFieldToSummarize "AÑO"
	task.AddFieldToSummarize "MES"
	task.AddFieldToTotal "SUMA_CON_IMP"
	task.AddFieldToTotal "VENTAS_CUPON"
	dbName = "Ventas con cupon 2010-L4.IMD"
	task.OutputDBName = dbName
	task.CreatePercentField = FALSE
	task.StatisticsToInclude = SM_SUM
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function
' Análisis: Resumen
Function Summarization1
	Set db = Client.OpenDatabase("ED Ventas-2010-L1-Ventas diarias por producto.IMD")
	Set task = db.Summarization
	task.AddFieldToSummarize "ID_LOCAL"
	task.AddFieldToSummarize "AÑO"
	task.AddFieldToSummarize "MES"
	task.AddFieldToTotal "SUMA_CON_IMP"
	task.AddFieldToTotal "VENTAS_CUPON"
	dbName = "Ventas con cupon 2010-L1.IMD"
	task.OutputDBName = dbName
	task.CreatePercentField = FALSE
	task.StatisticsToInclude = SM_SUM
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function
' Análisis: Resumen
Function Summarization2
	Set db = Client.OpenDatabase("ED Ventas-2010-L3-Ventas diarias por producto.IMD")
	Set task = db.Summarization
	task.AddFieldToSummarize "ID_LOCAL"
	task.AddFieldToSummarize "AÑO"
	task.AddFieldToSummarize "MES"
	task.AddFieldToTotal "SUMA_CON_IMP"
	task.AddFieldToTotal "VENTAS_CUPON"
	dbName = "Ventas con cupon 2010-L3.IMD"
	task.OutputDBName = dbName
	task.CreatePercentField = FALSE
	task.StatisticsToInclude = SM_SUM
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Archivo: Anexar bases de datos
Function AppendDatabase
	Set db = Client.OpenDatabase("Ventas con cupon 2010-L1.IMD")
	Set task = db.AppendDatabase
	task.AddDatabase "Ventas con cupon 2010-L3.IMD"
	task.AddDatabase "Ventas con cupon 2010-L4.IMD"
	task.Criteria = " MES  >= 9"
	dbName = "Completas - Ventas con cupon 2010.IMD"
	task.PerformTask dbName, ""
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Análisis: Correlación
Function Correlation
	Set db = Client.OpenDatabase("Completas - Ventas con cupon 2010.IMD")
	Set task = db.Correlation
	task.AddFieldForCorrelation "SUMA_CON_IMP_SUMA"
	task.AddFieldForCorrelation "VENTAS_CUPON_SUMA"
	task.AuditUnitField = "ID_LOCAL"
	resultName = db.UniqueResultName("Correlación de ventas y cupones 2010")
	task.ResultName = resultName
	dbName = "Correlación de ventas y cupones 2010.IMD"
	task.OutputDBName = dbName
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase(dbName)
End Function




