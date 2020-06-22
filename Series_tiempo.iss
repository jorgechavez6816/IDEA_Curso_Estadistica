Sub Main
	IgnoreWarning(True)
	Call Summarization()	'ED Ventas-2010-L4-Ventas diarias por producto.IMD
	Call Summarization1()	'ED Ventas-2010-L4-Ventas diarias por producto.IMD
	Call Summarization2()	'ED Ventas-2010-L4-Ventas diarias por producto.IMD
	Call Summarization3()	'ED Ventas-2010-L4-Ventas diarias por producto.IMD
	Call Summarization4()	'ED Ventas-2010-L4-Ventas diarias por producto.IMD
	Call Summarization5()	'ED Ventas-2010-L4-Ventas diarias por producto.IMD
	Call AppendDatabase()	'Series de tiempo-2011-L3.IMD
	Call TimeSeries()	'Completas - Series de tiempo.IMD
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
	dbName = "Series de tiempo-2010-L4.IMD"
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
	Set db = Client.OpenDatabase("ED Ventas-2011-L4-Ventas diarias por producto.IMD")
	Set task = db.Summarization
	task.AddFieldToSummarize "ID_LOCAL"
	task.AddFieldToSummarize "AÑO"
	task.AddFieldToSummarize "MES"
	task.AddFieldToTotal "SUMA_CON_IMP"
	dbName = "Series de tiempo-2011-L4.IMD"
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
	Set db = Client.OpenDatabase("ED Ventas-2010-L1-Ventas diarias por producto.IMD")
	Set task = db.Summarization
	task.AddFieldToSummarize "ID_LOCAL"
	task.AddFieldToSummarize "AÑO"
	task.AddFieldToSummarize "MES"
	task.AddFieldToTotal "SUMA_CON_IMP"
	dbName = "Series de tiempo-2010-L1.IMD"
	task.OutputDBName = dbName
	task.CreatePercentField = FALSE
	task.StatisticsToInclude = SM_SUM
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Análisis: Resumen
Function Summarization3
	Set db = Client.OpenDatabase("ED Ventas-2011-L1-Ventas diarias por producto.IMD")
	Set task = db.Summarization
	task.AddFieldToSummarize "ID_LOCAL"
	task.AddFieldToSummarize "AÑO"
	task.AddFieldToSummarize "MES"
	task.AddFieldToTotal "SUMA_CON_IMP"
	dbName = "Series de tiempo-2011-L1.IMD"
	task.OutputDBName = dbName
	task.CreatePercentField = FALSE
	task.StatisticsToInclude = SM_SUM
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Análisis: Resumen
Function Summarization4
	Set db = Client.OpenDatabase("ED Ventas-2010-L3-Ventas diarias por producto.IMD")
	Set task = db.Summarization
	task.AddFieldToSummarize "ID_LOCAL"
	task.AddFieldToSummarize "AÑO"
	task.AddFieldToSummarize "MES"
	task.AddFieldToTotal "SUMA_CON_IMP"
	dbName = "Series de tiempo-2010-L3.IMD"
	task.OutputDBName = dbName
	task.CreatePercentField = FALSE
	task.StatisticsToInclude = SM_SUM
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Análisis: Resumen
Function Summarization5
	Set db = Client.OpenDatabase("ED Ventas-2011-L3-Ventas diarias por producto.IMD")
	Set task = db.Summarization
	task.AddFieldToSummarize "ID_LOCAL"
	task.AddFieldToSummarize "AÑO"
	task.AddFieldToSummarize "MES"
	task.AddFieldToTotal "SUMA_CON_IMP"
	dbName = "Series de tiempo-2011-L3.IMD"
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
	Set db = Client.OpenDatabase("Series de tiempo-2010-L1.IMD")
	Set task = db.AppendDatabase
	task.AddDatabase "Series de tiempo-2010-L3.IMD"
	task.AddDatabase "Series de tiempo-2010-L4.IMD"
	task.AddDatabase "Series de tiempo-2011-L1.IMD"
	task.AddDatabase "Series de tiempo-2011-L3.IMD"
	task.AddDatabase "Series de tiempo-2011-L4.IMD"
	dbName = "Completas - Series de tiempo.IMD"
	task.PerformTask dbName, ""
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Análisis: Series de tiempo
Function TimeSeries
	Set db = Client.OpenDatabase("Completas - Series de tiempo.IMD")
	Set task = db.TimeSeriesAnalysisTask
	task.SeasonalLength = 12
	task.TimeSeriesField "SUMA_CON_IMP_SUMA" 
	task.AuditUnitField = "ID_LOCAL"
	task.GenerateForecasts TRUE, 8
	task.TimeScale = 1 
	task.CalendarValue = 1 
	task.TimeScaleStartAndIncrement 1, 1 
	resultName = db.UniqueResultName("Series de tiempo - Ventas de 2010 y 2011")
	task.ResultName = resultName
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
End Function