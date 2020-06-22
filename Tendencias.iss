Sub Main
	IgnoreWarning(True)
	Call Summarization()	'ED Ventas-2010-L4-Ventas diarias por producto.IMD
	Call Summarization1()	'ED Ventas-2011-L4-Ventas diarias por producto.IMD
	Call Summarization2()	'ED Ventas-2010-L1-Ventas diarias por producto.IMD
	Call Summarization3()	'ED Ventas-2010-L4-Ventas diarias por producto.IMD
	Call JoinDatabase()	'Ventas mensuales-2010-L4.IMD
	Call JoinDatabase1()	'Ventas mensuales-2010-L4.IMD
	Call ModifyField()	'Ventas mensuales 2010-L4 y L1.IMD
	Call ModifyField1()	'Ventas mensuales 2010-L4 y L1.IMD
	Call AppendDatabase()	'Ventas mensuales 2011-L4 y L1.IMD
	Call TrendAnalysis()	'Ventas mensuales-L4 y L1 (2010 y 2011).IMD
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
	dbName = "Ventas mensuales-2010-L4.IMD"
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
	dbName = "Ventas mensuales-2011-L4.IMD"
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
	dbName = "Ventas mensuales-2010-L1.IMD"
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
	dbName = "Ventas mensuales-2011-L1.IMD"
	task.OutputDBName = dbName
	task.CreatePercentField = FALSE
	task.StatisticsToInclude = SM_SUM
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Archivo: Unir bases de datos
Function JoinDatabase
	Set db = Client.OpenDatabase("Ventas mensuales-2010-L4.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "Ventas mensuales-2010-L1.IMD"
	task.IncludeAllPFields
	task.IncludeAllSFields
	task.AddMatchKey "AÑO", "AÑO", "A"
	task.AddMatchKey "MES", "MES", "A"
	task.CreateVirtualDatabase = False
	dbName = "Ventas mensuales 2010-L4 y L1.IMD"
	task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Archivo: Unir bases de datos
Function JoinDatabase1
	Set db = Client.OpenDatabase("Ventas mensuales-2011-L4.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "Ventas mensuales-2011-L1.IMD"
	task.IncludeAllPFields
	task.IncludeAllSFields
	task.AddMatchKey "AÑO", "AÑO", "A"
	task.AddMatchKey "MES", "MES", "A"
	task.CreateVirtualDatabase = False
	dbName = "Ventas mensuales 2011-L4 y L1.IMD"
	task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Modificar campo
Function ModifyField
	Set db = Client.OpenDatabase("Ventas mensuales 2010-L4 y L1.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "LOCAL1_BRUTO"
	field.Description = ""
	field.Type = WI_NUM_FIELD
	field.Equation = ""
	field.Decimals = 4
	task.ReplaceField "SUMA_CON_IMP_SUMA1", field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Modificar campo
Function ModifyField1
	Set db = Client.OpenDatabase("Ventas mensuales 2011-L4 y L1.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "LOCAL1_BRUTO"
	field.Description = ""
	field.Type = WI_NUM_FIELD
	field.Equation = ""
	field.Decimals = 4
	task.ReplaceField "SUMA_CON_IMP_SUMA1", field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Archivo: Anexar bases de datos
Function AppendDatabase
	Set db = Client.OpenDatabase("Ventas mensuales 2010-L4 y L1.IMD")
	Set task = db.AppendDatabase
	task.AddDatabase "Ventas mensuales 2011-L4 y L1.IMD"
	dbName = "Ventas mensuales-L4 y L1 (2010 y 2011).IMD"
	task.PerformTask dbName, ""
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Análisis: Análisis de tendencias
Function TrendAnalysis
	Set db = Client.OpenDatabase("Ventas mensuales-L4 y L1 (2010 y 2011).IMD")
	Set task = db.TrendAnalysis
	task.TrendField "SUMA_CON_IMP_SUMA" 
	task.RefField "LOCAL1_BRUTO"
	task.AuditUnitField = "ID_LOCAL"
	task.GenerateForecasts TRUE, 24, 8
	task.TimeScale = 1 
	task.CalendarValue = 1 
	task.TimeScaleStartAndIncrement 1, 1 
	resultName = db.UniqueResultName("Análisis de tendencias de L4 y L1")
	task.ResultName = resultName
	
	task.CreateMAPEDB = TRUE 
	task.OutputDBName = "Análisis de tendencias.IMD"
	
	task.CreateForecastDB = TRUE 
	task.OutputDBName = "Análisis de tendencias.IMD"
	
	task.CreateDB = TRUE 
	task.OutputDBName = "Análisis de tendencias.IMD"
	
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase("Análisis de tendencias PEAP.IMD")
End Function



	
