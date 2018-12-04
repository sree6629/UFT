'Create QTP object
Set QTP = CreateObject("QuickTest.Application")
QTP.Launch
QTP.Visible = TRUE
 
'Open QTP Test
QTP.Open "QTP Test Path", TRUE 'Set the QTP test path
 
'Set Result location
Set qtpResultsOpt = CreateObject("QuickTest.RunResultsOptions")
qtpResultsOpt.ResultsLocation = "Result path" 'Set the results location
 
'Run QTP test
QTP.Test.Run qtpResultsOpt
 
'Close QTP
QTP.Test.Close
QTP.Quit
