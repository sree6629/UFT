Public ReportUtil: Set ReportUtil = New clsReportUtil

Class clsReportUtil

	Private path
	private FSO
	private report
	

	Private Sub Class_Initialize
		Set FSO = CreateObject("Scripting.FileSystemObject")
		path = gBasePath & "\output\TestResult.html"
		Set report = FSO.OpenTextFile(path,8,True)
		createReportHeader()
	End Sub
	
	Private Sub createReportHeader()
		report.WriteLine("<table style='width: 930px;margin: 0;padding: 0;table-layout: fixed;border-collapse: collapse;font: 11px/1.4 Trebuchet MS;'>")
		report.WriteLine("<thead style='margin: 0;padding: 0;'>")
		report.WriteLine("<tr style='margin: 0;padding: 0;'>")
		report.WriteLine("<th style='width: 100px;margin: 0;padding: 6px;background: #333;color: white;font-weight: bold;border: 1px solid #ccc;text-align: auto;'>TCID</th>")
		report.WriteLine("<th style='width: 100px;margin: 0;padding: 6px;background: #333;color: white;font-weight: bold;border: 1px solid #ccc;text-align: auto;'>TIME</th>")
		report.WriteLine("<th style='width: 400px;margin: 0;padding: 6px;background: #333;color: white;font-weight: bold;border: 1px solid #ccc;text-align: auto;'>TestCase Description</th>")
		report.WriteLine("<th style='width: 100px;margin: 0;padding: 6px;background: #333;color: white;font-weight: bold;border: 1px solid #ccc;text-align: auto;'>Status</th>")
		report.WriteLine("</tr>")
		report.WriteLine("</thead><tbody style='margin: 0;padding: 0;'>")
	End Sub
	
	Public Sub UpdateTestCaseInfo(ByVal TCID, ByVal TestDescription)
		report.WriteLine("<tr style='margin: 0;padding: 0;'>")
		report.WriteLine("<td style='margin: 0;padding: 6px;border: 1px solid #ccc;text-align: left;background: #FFFFFF;'><b>" & TCID & "</b></td>")
		report.WriteLine("<td style='margin: 0;padding: 6px;border: 1px solid #ccc;text-align: left;background: #FFFFFF;'><b>" & now & "</b></td>")
		report.WriteLine("<td style='margin: 0;padding: 6px;border: 1px solid #ccc;text-align: left;background: #FFFFFF;'>" & TestDescription & "</td>")
	End Sub
	
	Public Sub UpdateTestCaseResult(ByVal TestResult)
		If TestResult = "PASS" Then
			report.WriteLine("<td style='margin: 0;padding: 6px;border: 1px solid #ccc;text-align: left;background: #99FF99;'>PASS</td>")
		Else
			report.WriteLine("<td style='margin: 0;padding: 6px;border: 1px solid #ccc;text-align: left;background: #FFB2B2;'>FAIL</td>")
		End If
		
		report.WriteLine("</tr>")
	End Sub
	
	Public Sub ReportEvent(ByVal TCID, ByVal TestDescription, ByVal TestStatus)
		UpdateTestCaseInfo TCID, TestDescription
		UpdateTestCaseResult(TestResult)
	End Sub

	Private Sub Class_Terminate(  )
		report.WriteLine("</tbody></thead></table>")
	End Sub


End Class
