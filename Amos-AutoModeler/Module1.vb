Imports System
Imports Microsoft.VisualBasic
Imports Amos
Imports AmosEngineLib
Imports AmosEngineLib.AmosEngine.TMatrixID
Imports MiscAmosTypes
Imports MiscAmosTypes.cDatabaseFormat
Imports System.Xml
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text.RegularExpressions

<System.ComponentModel.Composition.Export(GetType(Amos.IPlugin))>
Public Class CustomCode
    Implements IPlugin

    Public Function Name() As String Implements IPlugin.Name
        Return "Auto Modeler"
    End Function

    Public Function Description() As String Implements IPlugin.Description
        Return "Structural model modification indices-driven model maker"
    End Function

    Structure Estimates
        'An estimates struct holds the estimates
        Public Cmin As Double
        Public Df As Double
        Public CD As Double
        Public CFI As Double
        Public SRMR As Double
        Public Rmsea As Double
        Public Pclose As Double
    End Structure

    Structure DualEstimates
        Public Cmin As Double
        Public Df As Double
        Public CFI As Double
        Public CD As Double
        Public Rmsea As Double
        Public Pclose As Double
    End Structure

    Public Shared bMid As Boolean = False
    Public Shared bBad As Boolean = False
    Public Shared bConstraint As Boolean = False

    Public Function Mainsub() As Integer Implements IPlugin.MainSub
        pd.GetCheckBox("AnalysisPropertiesForm", "ModsCheck").Checked = True
        Dim pattern As String = "\be[0-9]+\b"
        Dim variable As PDElement
        Dim observedVariables As New Collections.ArrayList
        Dim z As Integer = 1

        If pd.PDElements.Count = 0 Then
            MsgBox("No variables detected, add observed variables to the model to use this plugin.")
            Exit Function
        End If

        For Each variable In pd.PDElements
            If Not variable.IsObservedVariable Then
                MsgBox("The plugin only works with observed variables, remove any other elements first.")
                Exit Function
            End If
            observedVariables.Add(variable)
        Next

        MsgBox("Amos will prompt you every time the model runs, please select 'Proceed with the analysis'")

        For x = 1 To observedVariables.Count - 1
            pd.AnalyzeCalculateEstimates()
            Dim tableMI As XmlElement = GetXML("body/ div / div[@ntype='models']/div[@ntype='model'][position() = 1]/div[@ntype='group'][position() = 1]/div[@ntype='modificationindices']/div[@nodecaption='Covariances:']/table/tbody")
            Dim numMI As Integer = GetNodeCount(tableMI)
            Dim dictMI As New Dictionary(Of Double, Integer)
            For y = 1 To numMI
                Dim m As Match = Regex.Match(MatrixName(tableMI, y, 0), pattern)
                Dim n As Match = Regex.Match(MatrixName(tableMI, y, 2), pattern)
                If Not m.Success And Not n.Success Then
                    dictMI.Add(MatrixElement(tableMI, y, 3), y)
                End If
            Next
            Dim keys As New List(Of Double)

            keys = dictMI.Keys.ToList
            keys.Sort()
            keys.Reverse()

            pd.DiagramDrawPath(MatrixName(tableMI, dictMI(keys(0)), 2), MatrixName(tableMI, dictMI(keys(0)), 0))
            pd.DiagramDrawUniqueVariable(MatrixName(tableMI, dictMI(keys(0)), 0))
            For Each E In pd.PDElements
                If E.NameOrCaption = "" Then
                    E.NameOrCaption = "e" & z
                    z += 1
                End If
            Next
        Next

        pd.AnalyzeCalculateEstimates()

        Dim tableRW As XmlElement = GetXML("body/div/div[@ntype='models']/div[@ntype='model'][position() = 1]/div[@ntype='group'][position() = 1]/div[@ntype='modificationindices']/div[@nodecaption='Regression Weights:']/table/tbody")
        Dim numRW As Integer = GetNodeCount(tableRW)

        Do Until numRW = 0
            Dim dictRW As New Dictionary(Of Double, Integer)
            For y = 1 To numRW
                dictRW.Add(MatrixElement(tableRW, y, 3), y)
            Next

            Dim keys As New List(Of Double)

            keys = dictRW.Keys.ToList
            Keys.Sort()
            keys.Reverse()
            pd.DiagramDrawPath(MatrixName(tableRW, dictRW(keys(0)), 2), MatrixName(tableRW, dictRW(keys(0)), 0))
            pd.DiagramDrawUniqueVariable(MatrixName(tableRW, dictRW(keys(0)), 0))
            For Each E In pd.PDElements
                If E.NameOrCaption = "" Then
                    E.NameOrCaption = "e" & z
                    z += 1
                End If
            Next

            pd.AnalyzeCalculateEstimates()

            tableRW = GetXML("body/div/div[@ntype='models']/div[@ntype='model'][position() = 1]/div[@ntype='group'][position() = 1]/div[@ntype='modificationindices']/div[@nodecaption='Regression Weights:']/table/tbody")
            numRW = GetNodeCount(tableRW)
        Loop

        Dim intButton As Integer
        intButton = MsgBox("Does your data file have missing data?", 3, "Missing Data Check")

        If intButton = 6 Then
            'Uncheck Mods and check means and intercepts
            Amos.pd.GetCheckBox("AnalysisPropertiesForm", "ModsCheck").Checked = False
            Amos.pd.GetCheckBox("AnalysisPropertiesForm", "MeansInterceptsCheck").Checked = True

            'Fits the specified model.
            Amos.pd.AnalyzeCalculateEstimates()

            'Produce the output
            LessHTML()

        ElseIf intButton = 7 Then
            'Ensure Mods and ResidualMom are checked
            Amos.pd.GetCheckBox("AnalysisPropertiesForm", "ModsCheck").Checked = True
            Amos.pd.GetCheckBox("AnalysisPropertiesForm", "ResidualMomCheck").Checked = True

            'Fits the specified model.
            Amos.pd.AnalyzeCalculateEstimates()

            'Produce the output
            CreateHTML()
        Else
            Exit Function
        End If

    End Function

    Sub CreateHTML()
        If (System.IO.File.Exists("AutoModeler.html")) Then
            System.IO.File.Delete("AutoModeler.html")
        End If

        Dim estimates As Estimates = GetEstimates()

        Dim listValues As List(Of varSummed) = GetLowestIndicator()

        'Set up the listener To output the debugs
        Dim debug As New AmosDebug.AmosDebug
        Dim resultWriter As New TextWriterTraceListener("AutoModeler.html")
        Trace.Listeners.Add(resultWriter)

        'Write the beginning Of the document
        debug.PrintX("<html><body><h1>Model Fit Measures</h1><hr/>")
        debug.PrintX("<p>NOTE: The model that has been produced is atheoretical and is based solely on modification indices. Some arrows may be better represented in the opposite direction and some regression arrows may be more appropriate as covariances.</p>")

        'Populate model fit measures in data table
        debug.PrintX("<table><tr><th>Measure</th><th>Estimate</th><th>Threshold</th><th>Interpretation</th></tr>")

        debug.PrintX("<tr><td>CMIN</td><td>" + estimates.Cmin.ToString("#0.000") + "</td><td>--</td><td>--</td></tr>")
        debug.PrintX("<tr><td>DF</td><td>" + estimates.Df.ToString("#0.000") + "</td><td>--</td><td>--</td></tr>")
        debug.PrintX("<tr><td>CMIN/DF</td><td>" + estimates.CD.ToString("#0.000") + "</td><td>Between 1 and 3</td><td>")

        If estimates.CD < 1 Then
            debug.PrintX("Need more DF</td></tr>")
        ElseIf estimates.CD >= 1 And estimates.CD <= 3 Then
            debug.PrintX("Excellent</td></tr>")
        ElseIf estimates.CD <= 5 Then
            debug.PrintX("Acceptable</td></tr>")
        ElseIf estimates.CD = Nothing Then
            debug.PrintX("Not Estimated</td></tr>")
        Else
            debug.PrintX("Terrible</td></tr>")
            bBad = True
        End If

        debug.PrintX("<tr><td>CFI</td><td>" + estimates.CFI.ToString("#0.000") + "</td><td>>0.95</td><td>")

        If estimates.CFI > 0.95 Then
            debug.PrintX("Excellent</td></tr>")
        ElseIf estimates.CFI > 0.9 Then
            debug.PrintX("Acceptable</td></tr>")
        ElseIf estimates.CFI = Nothing Then
            debug.PrintX("Not Estimated</td></tr>")
        Else
            debug.PrintX("Terrible</td></tr>")
            bBad = True
        End If

        debug.PrintX("<tr><td>SRMR</td><td>" + estimates.SRMR.ToString("#0.000") + "</td><td><0.08</td><td>")

        If estimates.SRMR < 0.08 Then
            debug.PrintX("Excellent</td></tr>")
        ElseIf estimates.SRMR < 0.1 Then
            debug.PrintX("Acceptable</td></tr>")
        ElseIf estimates.SRMR = Nothing Then
            debug.PrintX("Not Estimated</td></tr>")
        Else
            debug.PrintX("Terrible</td></tr>")
            bBad = True
        End If

        debug.PrintX("<tr><td>RMSEA</td><td>" + estimates.Rmsea.ToString("#0.000") + "</td><td><0.06</td><td>")

        If estimates.Rmsea < 0.06 Then
            debug.PrintX("Excellent</td></tr>")
        ElseIf estimates.Rmsea < 0.08 Then
            debug.PrintX("Acceptable</td></tr>")
        ElseIf estimates.Rmsea = Nothing Then
            debug.PrintX("Not Estimated</td></tr>")
        Else
            debug.PrintX("Terrible</td></tr>")
            bBad = True
        End If

        debug.PrintX("<tr><td>PClose</td><td>" + estimates.Pclose.ToString("#0.000") + "</td><td>>0.05</td><td>")

        If estimates.Pclose > 0.05 Then
            debug.PrintX("Excellent</td></tr>")
        ElseIf estimates.Pclose > 0.01 Then
            debug.PrintX("Acceptable</td></tr>")
        ElseIf estimates.Pclose = Nothing Then
            debug.PrintX("Not Estimated</td></tr>")
        Else
            debug.PrintX("Terrible</td></tr>")
            bBad = True
        End If

        debug.PrintX("</table><br>")

        'Check standardized residual covariances table for the indicator with largest sum of absolute values.
        If bMid = False And bBad = False Then
            debug.PrintX("Congratulations, your model fit is excellent!")
        ElseIf bMid = True And bBad = False Then
            debug.PrintX("Congratulations, your model fit is acceptable.")
        Else
            debug.PrintX("Your model fit could improve. Based on the standardized residual covariances, we recommend removing " + listValues.First.Name + ".")
        End If

        If bConstraint = True Then
            debug.PrintX("<br>This indicator has a path constraint. You will need to change the constraint after removing " + listValues.First.Name + ".")
        End If

        'Write reference table and credits
        debug.PrintX("<hr/><h3> Cutoff Criteria*</h3><table><tr><th>Measure</th><th>Terrible</th><th>Acceptable</th><th>Excellent</th></tr>")
        debug.PrintX("<tr><td>CMIN/DF</td><td>> 5</td><td>> 3</td><td>> 1</td></tr>")
        debug.PrintX("</td></tr><tr><td>CFI</td><td><0.90</td><td><0.95</td><td>>0.95</td></tr>")
        debug.PrintX("</td></tr><tr><td>SRMR</td><td>>0.10</td><td>>0.08</td><td><0.08</td></tr>")
        debug.PrintX("</td></tr><tr><td>RMSEA</td><td>>0.08</td><td>>0.06</td><td><0.06</td></tr>")
        debug.PrintX("</td></tr><tr><td>PClose</td><td><0.01</td><td><0.05</td><td>>0.05</td></tr></table>")
        debug.PrintX("<p>*Note: Hu and Bentler (1999, ""Cutoff Criteria for Fit Indexes in Covariance Structure Analysis: Conventional Criteria Versus New Alternatives"") recommend combinations of measures. Personally, I prefer a combination of CFI>0.95 and SRMR<0.08. To further solidify evidence, add the RMSEA<0.06.</p>")
        debug.PrintX("<p>**If you would like to cite this tool directly, please use the following:")
        debug.PrintX("Gaskin, J. & Lim, J. (2016), ""Model Fit Measures"", AMOS Plugin. <a href=\""http://statwiki.kolobkreations.com"">Gaskination's StatWiki</a>.</p>")

        'Write Style And close
        debug.PrintX("<style>h1{margin-left:60px;}table{border:1px solid black;border-collapse:collapse;}td{border:1px solid black;text-align:center;padding:5px;}th{text-weight:bold;padding:10px;border: 1px solid black;}</style>")
        debug.PrintX("</body></html>")

        'Take down our debugging, release file, open html
        Trace.Flush()
        Trace.Listeners.Remove(resultWriter)
        resultWriter.Close()
        resultWriter.Dispose()
        Process.Start("AutoModeler.html")
    End Sub

    Sub LessHTML()
        If (System.IO.File.Exists("AutoModeler.html")) Then
            System.IO.File.Delete("AutoModeler.html")
        End If

        Dim estimates As Estimates = GetEstimates()

        Dim listValues As List(Of varSummed) = GetLowestIndicator()

        'Set up the listener To output the debugs
        Dim debug As New AmosDebug.AmosDebug
        Dim resultWriter As New TextWriterTraceListener("ModelFit.html")
        Trace.Listeners.Add(resultWriter)

        'Write the beginning Of the document
        debug.PrintX("<html><body><h1>Model Fit Measures</h1><hr/>")
        debug.PrintX("<p>NOTE: The model that has been produced is atheoretical and is based solely on modification indices. Some arrows may be better represented in the opposite direction and some regression arrows may be more appropriate as covariances.</p>")

        'Populate model fit measures in data table
        debug.PrintX("<table><tr><th>Measure</th><th>Estimate</th><th>Threshold</th><th>Interpretation</th></tr>")

        debug.PrintX("<tr><td>CMIN</td><td>" + estimates.Cmin.ToString("#0.000") + "</td><td>--</td><td>--</td></tr>")
        debug.PrintX("<tr><td>DF</td><td>" + estimates.Df.ToString("#0.000") + "</td><td>--</td><td>--</td></tr>")
        debug.PrintX("<tr><td>CMIN/DF</td><td>" + estimates.CD.ToString("#0.000") + "</td><td>Between 1 and 3</td><td>")

        If estimates.CD < 1 Then
            debug.PrintX("Need more DF</td></tr>")
        ElseIf estimates.CD >= 1 And estimates.CD <= 3 Then
            debug.PrintX("Excellent</td></tr>")
        ElseIf estimates.CD <= 5 Then
            debug.PrintX("Acceptable</td></tr>")
        ElseIf estimates.CD = Nothing Then
            debug.PrintX("Not Estimated</td></tr>")
        Else
            debug.PrintX("Terrible</td></tr>")
            bBad = True
        End If

        debug.PrintX("<tr><td>CFI</td><td>" + estimates.CFI.ToString("#0.000") + "</td><td>>0.95</td><td>")

        If estimates.CFI > 0.95 Then
            debug.PrintX("Excellent</td></tr>")
        ElseIf estimates.CFI > 0.9 Then
            debug.PrintX("Acceptable</td></tr>")
        ElseIf estimates.CFI = Nothing Then
            debug.PrintX("Not Estimated</td></tr>")
        Else
            debug.PrintX("Terrible</td></tr>")
            bBad = True
        End If

        debug.PrintX("<tr><td>RMSEA</td><td>" + estimates.Rmsea.ToString("#0.000") + "</td><td><0.06</td><td>")

        If estimates.Rmsea < 0.06 Then
            debug.PrintX("Excellent</td></tr>")
        ElseIf estimates.Rmsea < 0.08 Then
            debug.PrintX("Acceptable</td></tr>")
        ElseIf estimates.Rmsea = Nothing Then
            debug.PrintX("Not Estimated</td></tr>")
        Else
            debug.PrintX("Terrible</td></tr>")
            bBad = True
        End If

        debug.PrintX("<tr><td>PClose</td><td>" + estimates.Pclose.ToString("#0.000") + "</td><td>>0.05</td><td>")

        If estimates.Pclose > 0.05 Then
            debug.PrintX("Excellent</td></tr>")
        ElseIf estimates.Pclose > 0.01 Then
            debug.PrintX("Acceptable</td></tr>")
        ElseIf estimates.Pclose = Nothing Then
            debug.PrintX("Not Estimated</td></tr>")
        Else
            debug.PrintX("Terrible</td></tr>")
            bBad = True
        End If

        debug.PrintX("</table><br>")

        'Check standardized residual covariances table for the indicator with largest sum of absolute values.
        If bMid = False And bBad = False Then
            debug.PrintX("Congratulations, your model fit is excellent!")
        ElseIf bMid = True And bBad = False Then
            debug.PrintX("Congratulations, your model fit is acceptable.")
        Else
            debug.PrintX("Your model fit could improve. Based on the standardized residual covariances, we recommend removing " + listValues.First.Name + ".")
        End If

        If bConstraint = True Then
            debug.PrintX("<br>This indicator has a path constraint. You will need to change the constraint after removing " + listValues.First.Name + ".")
        End If

        'Write reference table and credits
        debug.PrintX("<hr/><h3> Cutoff Criteria*</h3><table><tr><th>Measure</th><th>Terrible</th><th>Acceptable</th><th>Excellent</th></tr>")
        debug.PrintX("<tr><td>CMIN/DF</td><td>> 5</td><td>> 3</td><td>> 1</td></tr>")
        debug.PrintX("</td></tr><tr><td>CFI</td><td><0.90</td><td><0.95</td><td>>0.95</td></tr>")
        debug.PrintX("</td></tr><tr><td>RMSEA</td><td>>0.08</td><td>>0.06</td><td><0.06</td></tr>")
        debug.PrintX("</td></tr><tr><td>PClose</td><td><0.01</td><td><0.05</td><td>>0.05</td></tr></table>")
        debug.PrintX("<p>*Note: Hu and Bentler (1999, ""Cutoff Criteria for Fit Indexes in Covariance Structure Analysis: Conventional Criteria Versus New Alternatives"") recommend combinations of measures. Personally, I prefer a combination of CFI>0.95 and SRMR<0.08. To further solidify evidence, add the RMSEA<0.06.</p>")
        debug.PrintX("<p>**If you would like to cite this tool directly, please use the following:")
        debug.PrintX("Gaskin, J. & Lim, J. (2016), ""Model Fit Measures"", AMOS Plugin. <a href=\""http://statwiki.kolobkreations.com"">Gaskination's StatWiki</a>.</p>")

        'Write Style And close
        debug.PrintX("<style>h1{margin-left:60px;}table{border:1px solid black;border-collapse:collapse;}td{border:1px solid black;text-align:center;padding:5px;}th{text-weight:bold;padding:10px;border: 1px solid black;}</style>")
        debug.PrintX("</body></html>")

        'Take down our debugging, release file, open html
        Trace.Flush()
        Trace.Listeners.Remove(resultWriter)
        resultWriter.Close()
        resultWriter.Dispose()
        Process.Start("AutoModeler.html")
    End Sub

    Function GetEstimates() As Estimates

        'Array to hold estimates
        Dim estimates As Estimates

        'Get CFI from Baseline Comparisions table
        Dim CFI As XmlElement = GetXML("body/div/div[@ntype='modelfit']/div[@nodecaption='Baseline Comparisons']/table/tbody/tr[position() = 1]/td[position() = 6]")

        'Specify and fit the object to the model
        Dim Sem As New AmosEngineLib.AmosEngine
        Sem.NeedEstimates(SampleCorrelations)
        Sem.NeedEstimates(ImpliedCorrelations)
        Amos.pd.SpecifyModel(Sem)
        Sem.FitModel()

        'Calculate SRMR
        Dim N As Integer
        Dim i As Integer
        Dim j As Integer
        Dim SRMR As Double
        Dim Sample(,) As Double
        Dim Implied(,) As Double

        Sem.GetEstimates(SampleCorrelations, Sample)
        Sem.GetEstimates(ImpliedCorrelations, Implied)
        N = UBound(Sample, 1) + 1
        SRMR = 0
        For i = 1 To N - 1
            For j = 0 To i - 1
                SRMR = SRMR + (Sample(i, j) - Implied(i, j)) ^ 2
            Next
        Next
        SRMR = System.Math.Sqrt(SRMR / (N * (N - 1) / 2))

        estimates.Cmin = Sem.Cmin
        estimates.Df = Sem.Df
        estimates.CD = Sem.Cmin / Sem.Df
        estimates.CFI = CFI.InnerText
        estimates.SRMR = SRMR
        estimates.Rmsea = Sem.Rmsea
        estimates.Pclose = Sem.Pclose

        Sem.Dispose()

        Return estimates

    End Function

    Function DualData() As DualEstimates
        'Array to hold estimates
        Dim estimates As DualEstimates

        'Get CFI from Baseline Comparisions table
        Dim CFI As XmlElement = GetXML("body/div/div[@ntype='modelfit']/div[@nodecaption='Baseline Comparisons']/table/tbody/tr[position() = 1]/td[position() = 6]")

        'Specify and fit the object to the model
        Dim Sem As New AmosEngineLib.AmosEngine
        Amos.pd.SpecifyModel(Sem)
        Sem.FitModel()

        estimates.Cmin = Sem.Cmin
        estimates.Df = Sem.Df
        estimates.CFI = CFI.InnerText
        estimates.Rmsea = Sem.Rmsea
        estimates.Pclose = Sem.Pclose

        Sem.Dispose()

        Return estimates
    End Function

    Function GetLowestIndicator() As List(Of varSummed)
        Dim tableSRC As XmlElement = GetXML("body/div/div[@ntype='models']/div[@ntype='model'][position() = 1]/div[@ntype='group'][position() = 1]/div[@ntype='estimates']/div[@ntype='matrices']/div[@ntype='ppml'][position() = 2]/table/tbody")
        Dim headSRC As XmlElement = GetXML("body/div/div[@ntype='models']/div[@ntype='model'][position() = 1]/div[@ntype='group'][position() = 1]/div[@ntype='estimates']/div[@ntype='matrices']/div[@ntype='ppml'][position() = 2]/table/thead")
        Dim tableRW As XmlElement = GetXML("body/div/div[@ntype='models']/div[@ntype='model'][position() = 1]/div[@ntype='group'][position() = 1]/div[@ntype='estimates']/div[@ntype='scalars']/div[@nodecaption='Regression Weights:']/table/tbody")

        Dim observedVariables As New ArrayList
        For Each variable As PDElement In pd.PDElements
            If variable.IsObservedVariable Then
                observedVariables.Add(variable)
            End If
        Next

        Dim numObserved As Integer = observedVariables.Count

        'The following section checks if there are at least two unobserved variables connected to the latent variable
        Dim unavailableLatent As New ArrayList

        For Each variable As PDElement In pd.PDElements
            Dim numRW As Integer = 0
            If variable.IsLatentVariable Then
                For i = 1 To numObserved
                    If MatrixName(tableRW, i, 2) = variable.NameOrCaption Then
                        numRW += 1
                    End If
                Next
                'If there is less than three, the program will not recommend removing those latent variables.
                If numRW < 3 Then
                    unavailableLatent.Add(variable.NameOrCaption)
                End If
            End If
        Next

        'The list of latent variables with too few observed variables.
        Dim unavailableObserved As New ArrayList
        For Each latent As String In unavailableLatent
            For d = 1 To observedVariables.Count
                If MatrixName(tableRW, d, 2) = latent Then
                    unavailableObserved.Add(MatrixName(tableRW, d, 0))
                End If
            Next
        Next

        'These counters are used to process the standardized residual covariances table
        Dim column As Integer = 0
        Dim rowOffset As Integer = 0
        Dim row As Integer = 1
        Dim dSum As Double 'Stores the sum of values in the SRC
        Dim listValues As New List(Of varSummed) 'A list of objects that will hold a string and value
        For i = 0 To numObserved - 1 'For the number of observed variables
            Dim varName As String = MatrixName(headSRC, 1, (i + 1))
            For b = 1 To column 'Add the column
                dSum = dSum + Math.Abs(MatrixElement(tableSRC, (b + rowOffset), (i + 1)))
            Next
            For c = 1 To row 'Add the row
                dSum = dSum + Math.Abs(MatrixElement(tableSRC, (i + 1), c))
            Next
            column -= 1
            rowOffset += 1
            row += 1
            Dim oValues As New varSummed(varName, dSum) 'Assign the name of the variable and the summed value to an object
            If Not unavailableObserved.Contains(varName) Then
                listValues.Add(oValues) 'Add object to list
            End If
            dSum = 0
        Next
        listValues = listValues.OrderBy(Function(x) x.Total).ToList() 'Sort the list of values

        For d = 1 To numObserved
            If MatrixName(tableRW, d, 0) = listValues.First.Name And MatrixElement(tableRW, d, 3) = 1 Then
                bConstraint = True
            End If
        Next

        Return listValues

    End Function

    '
    Function interpret(good As Double, mid As Double, bad As Double, estimate As Double, dfCheck As Boolean) As String
        Dim interpretation As String
        'IF CMIN/DF < 1 Then needs more DF
        Select Case estimate
            Case Is > good
                interpretation = "Excellent"
            Case Is > mid
                interpretation = "Acceptable"
                bMid = True
            Case Is < 1 And dfCheck = True
                interpretation = "Need more DF"
                bBad = True
            Case Is = Nothing
                interpretation = "Not Estimated"
            Case Else
                interpretation = "Terrible"
                bBad = True
        End Select

        interpret = interpretation
    End Function

    'Get the number of rows in an xml table.
    Function GetNodeCount(table As XmlElement) As Integer

        Dim nodeCount As Integer = 0

        'Handles a model with zero correlations
        Try
            nodeCount = table.ChildNodes.Count
        Catch ex As NullReferenceException
            nodeCount = 0
        End Try

        GetNodeCount = nodeCount

    End Function

    'Use an output table path to get the xml version of the table.
    Public Function GetXML(path As String) As XmlElement

        'Gets the xpath expression for an output table.
        Dim doc As Xml.XmlDocument = New Xml.XmlDocument()
        doc.Load(pd.ProjectName & ".AmosOutput")
        Dim nsmgr As XmlNamespaceManager = New XmlNamespaceManager(doc.NameTable)
        Dim eRoot As Xml.XmlElement = doc.DocumentElement

        Return eRoot.SelectSingleNode(path, nsmgr)

    End Function

    'Get a string element from an xml table.
    Function MatrixName(eTableBody As XmlElement, row As Long, column As Long) As String

        Dim e As XmlElement

        Try
            e = eTableBody.ChildNodes(row - 1).ChildNodes(column) 'This means that the rows are not 0 based.
            MatrixName = e.InnerText
        Catch ex As Exception
            MatrixName = ""
        End Try

    End Function

    'Get a number from an xml table
    Function MatrixElement(eTableBody As XmlElement, row As Long, column As Long) As Double

        Dim e As XmlElement

        Try
            e = eTableBody.ChildNodes(row - 1).ChildNodes(column) 'This means that the rows are not 0 based.
            MatrixElement = CDbl(e.GetAttribute("x"))
        Catch ex As Exception
            MatrixElement = 0
        End Try

    End Function

End Class

Public Class varSummed
    Public Name As String
    Public Total As Double

    Public Sub New(ByVal sName As String, ByVal dTotal As Double)
        'constructor
        Name = sName
        Total = dTotal
        'storing the values in constructor
    End Sub

End Class