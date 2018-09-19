Imports System
Imports Microsoft.VisualBasic
Imports Amos
Imports AmosEngineLib
Imports AmosEngineLib.AmosEngine.TMatrixID
Imports MiscAmosTypes
Imports MiscAmosTypes.cDatabaseFormat
Imports System.Xml
Imports System.Text.RegularExpressions

<System.ComponentModel.Composition.Export(GetType(Amos.IPlugin))>
Public Class CustomCode
    Implements IPlugin

    Public Function Name() As String Implements IPlugin.Name
        Return "Specific bias test"
    End Function

    Public Function Description() As String Implements IPlugin.Description
        Return "Common method bias test for your dataset."
    End Function

    Public Function Mainsub() As Integer Implements IPlugin.MainSub

        'Properties
        Dim unobservedVariables As New Collections.ArrayList 'The selected unobserved variable in the model
        Dim observedVariables As New Collections.ArrayList 'An array for all observed variables
        Dim selectedVariables As New ArrayList 'An array of your selected unobserved variables
        Dim connectedVariables As New ArrayList 'An array of the observed variables connected to your selected variables
        Dim covariedVariables As New ArrayList
        Dim pValue1 As New Double
        Dim pValue2 As New Double
        Dim variable As PDElement
        'These arrays hold the chi square and df results from a test
        Dim estimates() As Double
        Dim estimates2() As Double
        Dim estimates3() As Double
        Dim conclusion As String = ""

        'Arraylist that will be used to check if the selected variable is covaried.
        For Each variable In pd.PDElements
            If variable.IsCovariance Then
                If Not covariedVariables.Contains(variable.Variable1.NameOrCaption) Then
                    covariedVariables.Add(variable.Variable1.NameOrCaption)
                End If
                If Not covariedVariables.Contains(variable.Variable2.NameOrCaption) Then
                    covariedVariables.Add(variable.Variable2.NameOrCaption)
                End If
            End If
        Next

        For Each variable In pd.PDElements
            'Get the selected unobserved variable.
            If variable.IsSelected And variable.IsUnobservedVariable Then
                If covariedVariables.Contains(variable.NameOrCaption) Then
                    MsgBox("Please only select the specific bias latent factor(s).")
                    Exit Function
                End If
                unobservedVariables.Add(variable)
                variable.Value1 = 1
            End If
        Next

        'Checks if there are no unobserved variables selected.
        If unobservedVariables.Count = 0 Then
            MsgBox("Please select specific bias latent factor(s).")
            Exit Function
        Else
            'An array of the selected variables.
            For Each variable In unobservedVariables
                selectedVariables.Add(variable.NameOrCaption)
            Next
        End If

        'Make an array of all variables connected to your selected variables, we will not motify these constraints.
        For Each variable In pd.PDElements
            If variable.IsPath Then
                If variable.Variable1.IsExogenousVariable And selectedVariables.Contains(variable.Variable1.NameOrCaption) Then
                    connectedVariables.Add(variable.Variable2.NameOrCaption)
                End If
            End If
        Next

        'Make an array of all observed variables
        For Each variable In pd.PDElements
            If variable.IsObservedVariable Then
                If Not connectedVariables.Contains(variable.NameOrCaption) Then
                    observedVariables.Add(variable)
                End If
            End If
        Next

        'Draw the paths to all observed variables
        For Each variable In observedVariables
            For Each unobserved In unobservedVariables
                pd.DiagramDrawPath(unobserved, variable)
            Next
        Next

        'Fits the specified model.
        MsgBox(“This plugin will run multiple tests to determine if there is bias and if it is evenly distributed. Whenever prompted by AMOS, please click "“Proceed with the analysis”".”)

        'Function to get the chi-square and df for the unconstrained model
        estimates = GetEstimates()

        If Not GetXML("body/ div / div[@ntype='models']/div[@ntype='model'][position() = 1]/div[@ntype='modelnotes']/div[@ntype='result']").InnerText.Contains("Minimum was achieved") Then
            MsgBox("Solution could not be generated, try running the model without the model bias plugin to troubleshoot. This is not a plugin problem, but a model problem. Adding a common latent factor often creates instability in measurement models.")
            erasePaths(connectedVariables, selectedVariables)
            Exit Function
        End If

        'Zero out the constraints of the paths
        erasePaths(connectedVariables, selectedVariables)
        zeroConstraints(unobservedVariables, observedVariables)

        'Function to get the chi-square and df for the zero constrained model
        estimates2 = getEstimates()

        'The chi-squared difference between the unconstrained and zero constrained models
        pValue1 = AmosEngine.ChiSquareProbability(Math.Abs(estimates(0) - estimates2(0)), Math.Abs(estimates(1) - estimates2(1)))

        'Significance test for the chi-squared difference after zero constrained test
        If pValue1 > 0.05 Then
            'touchUp(unobservedVariables)
            conclusion = "The null hypothesis cannot be rejected (i.e., the constrained And unconstrained models are the same Or ""invariant"").
                            You have demonstrated that you were unable To detect any specific response bias affecting your model. Therefore no bias distribution test was made (Of equal constraints).
                            You can move on to causal modeling, but make sure to retain the Specific Bias construct(s) to include as control in the causal model. "
            printHtml1(estimates, estimates2, pValue1, conclusion)
            Exit Function
        Else
            'Set the regression weights to 'a'
            erasePaths(connectedVariables, selectedVariables)
            equalConstraints(unobservedVariables, observedVariables)
        End If

        unconstrainedPaths(unobservedVariables, observedVariables)

        'Function to get the chi-square and df for the equal constrained model
        estimates3 = getEstimates()
        erasePaths(connectedVariables, selectedVariables)
        'The chi-squared difference between the unconstrained and the equal constrained models
        pValue2 = AmosEngine.ChiSquareProbability(Math.Abs(estimates(0) - estimates3(0)), Math.Abs(estimates(1) - estimates3(1)))

        'Significance test for the chi-squared difference after equal constrained test
        If pValue2 > 0.05 Then
            'touchUp(unobservedVariable)
            conclusion = "The chi-square test for the zero constrained model was significant (i.e., measurable bias was detected). Therefore a bias distribution test was made (of equal constraints).
                        The chi-square difference test between the constrained (to be equal) and unconstrained models indicates invariance 
                        (i.e., fail to reject null - that they are equal), the bias is equally distributed. Make note of this in your report. e.g., 
                        ""A test of equal specific bias demonstrated evenly distributed bias."" Move on to causal modeling with the SB constructs retained (keep them)."
            printHtml2(estimates, estimates2, estimates3, pValue1, pValue2, conclusion)
            Exit Function
        Else
            'touchUp(unobservedVariable)
            conclusion = "The chi-square test for the zero constrained model was significant (i.e., measurable bias was detected). Therefore a bias distribution test was made (of equal constraints).
                        The chi-square test is significant on this test as well (i.e., unevenly distributed bias), you should retain the SB construct for subsequent causal analyses. 
                        Make note of this in your report. e.g., ""A test of equal specific bias demonstrated unevenly distributed bias."""
            printHtml2(estimates, estimates2, estimates3, pValue1, pValue2, conclusion)
            Exit Function
        End If

    End Function

    'Gets the cmin And the df For the given model condition.
    Function GetEstimates() As Double()

        pd.AnalyzeCalculateEstimates()

        'Properties
        Dim modelNotes As XmlElement = GetXML("body/div/div[@ntype='models']/div[@ntype='model'][position() = 1]/div[@ntype='modelnotes']/div[@ntype='result']")

        'Use regex to extract the chi-square and df from the result
        Dim result As String = modelNotes.InnerText
        Dim myRegex As Match = Regex.Match(result, "\d+(\.\d{1,3})?", RegexOptions.IgnoreCase)
        If myRegex.Success Then
            Dim baseEstimates() As Double = {Convert.ToDouble(Convert.ToString(myRegex.Value)), Convert.ToDouble(Convert.ToString(myRegex.NextMatch))}
            GetEstimates = baseEstimates
        Else
            MsgBox(modelNotes.InnerText)
            Exit Function
        End If

    End Function

    Private Sub erasePaths(connectedVariables As ArrayList, selectedVariables As ArrayList)
        Dim variable As PDElement
        Dim pathFromUnobserved As New Collections.ArrayList 'An array for all observed variables

        'Check for newly created paths.
        For Each variable In pd.PDElements
            If variable.IsPath Then
                If selectedVariables.Contains(variable.Variable1.NameOrCaption) Then
                    If Not connectedVariables.Contains(variable.Variable2.NameOrCaption) Then
                        pathFromUnobserved.Add(variable)
                    End If
                End If
            End If
        Next

        'Clear out the paths.
        For Each variable In pathFromUnobserved
            pd.EditErase(variable)
        Next
    End Sub

    'Use an output table path to get the xml version of the table.
    Public Function GetXML(path As String) As XmlElement

        'Gets the xpath expression for an output table.
        Dim doc As Xml.XmlDocument = New Xml.XmlDocument()
        doc.Load(Amos.pd.ProjectName & ".AmosOutput")
        Dim nsmgr As XmlNamespaceManager = New XmlNamespaceManager(doc.NameTable)
        Dim eRoot As Xml.XmlElement = doc.DocumentElement

        Return eRoot.SelectSingleNode(path, nsmgr)

    End Function

    Private Sub zeroConstraints(unobservedVariables As Collections.ArrayList, observedVariables As Collections.ArrayList)

        Dim thisPath As PDElement 'A path type amos object.
        Dim variable As PDElement

        'Redraw those paths with a regression weight of 0.
        For Each variable In observedVariables
            For Each unobserved In unobservedVariables
                thisPath = pd.DiagramDrawPath(unobserved, variable)
                thisPath.Value1 = 0
            Next
        Next
    End Sub

    Private Sub equalConstraints(unobservedVariables As Collections.ArrayList, observedVariables As Collections.ArrayList)
        Dim thisPath As PDElement 'A path type amos object.
        Dim variable As PDElement

        'Redraw those paths with a regression weight of 0.
        For Each variable In observedVariables
            For Each unobserved In unobservedVariables
                thisPath = pd.DiagramDrawPath(unobserved, variable)
                thisPath.Value1 = "a"
            Next

        Next
    End Sub

    Private Sub unconstrainedPaths(unobservedVariables As Collections.ArrayList, observedVariables As Collections.ArrayList)
        Dim variable As PDElement

        'Redraw paths
        For Each variable In observedVariables
            For Each unobserved In unobservedVariables
                pd.DiagramDrawPath(unobserved, variable)
            Next
        Next
    End Sub

    Private Sub touchUp(unobservedVariable As PDElement)
        pd.EditTouchUp(unobservedVariable)
        pd.EditSelect(unobservedVariable)
    End Sub

    Private Sub printHtml1(estimates() As Double, estimates2() As Double, pValue1 As Double, conclusion As String)
        'Remove the old table files
        If (System.IO.File.Exists("ModelBias.html")) Then
            System.IO.File.Delete("ModelBias.html")
        End If

        'Start the Amos debugger to print the table
        Dim debug As New AmosDebug.AmosDebug

        'Set up the listener To output the debugs
        Dim resultWriter As New TextWriterTraceListener("ModelBias.html")
        Trace.Listeners.Add(resultWriter)

        'Write the beginning Of the document and the table header
        debug.PrintX("<html><body><h1>Specific Bias Tests</h1><hr/>")

        debug.PrintX("<h2>Zero Constraints Test (is there specific bias?)<h2>")
        debug.PrintX("<table><tr><td></td><th>X<sup>2</sup></th><th>DF</th><th>Delta</th><th>p-value</th></tr>")
        debug.PrintX("<tr><th>Unconstrained Model</th><td>" & estimates(0).ToString("#0.000") & "</td><td>" & estimates(1).ToString & "</td><td rowspan=""2"">X<sup>2</sup>=" & FormatNumber(CDbl(Math.Abs((estimates(0) - estimates2(0))).ToString), 3) & "<br>DF=" & Math.Abs((estimates(1) - estimates2(1))).ToString & "</td><td rowspan=""2"">" & FormatNumber(CDbl(pValue1.ToString), 3) & "</td></tr>")
        debug.PrintX("<tr><th>Zero Constrained Model</th><td>" + estimates2(0).ToString("#0.000") & "</td><td>" & estimates2(1).ToString & "</td></tr>")
        debug.PrintX("</table><hr/><h3>Conclusion</h3>")
        debug.PrintX("<p>" & conclusion & "</p><hr/>")

        debug.PrintX("<p>--If you would like to cite this tool directly, please use the following:")
        debug.PrintX("Gaskin, J. & Lim, J. (2017), ""CFA Tool"", AMOS Plugin. <a href=\""http://statwiki.kolobkreations.com"">Gaskination's StatWiki</a>.</p>")

        'Write Style And close
        debug.PrintX("<style>table{border:1px solid black;border-collapse:collapse;}td{border:1px solid black;text-align:center;padding:5px;}th{text-weight:bold;padding:10px;border: 1px solid black;}</style>")
        debug.PrintX("</body></html>")

        'Take down our debugging, release file, open html
        Trace.Flush()
        Trace.Listeners.Remove(resultWriter)
        resultWriter.Close()
        resultWriter.Dispose()
        Process.Start("ModelBias.html")
    End Sub

    Private Sub printHtml2(estimates() As Double, estimates2() As Double, estimates3() As Double, pValue1 As Double, pValue2 As Double, conclusion As String)
        'Remove the old table files
        If (System.IO.File.Exists("ModelBias.html")) Then
            System.IO.File.Delete("ModelBias.html")
        End If

        'Start the Amos debugger to print the table
        Dim debug As New AmosDebug.AmosDebug

        'Set up the listener To output the debugs
        Dim resultWriter As New TextWriterTraceListener("ModelBias.html")
        Trace.Listeners.Add(resultWriter)

        'Write the beginning Of the document and the table header
        debug.PrintX("<html><body><h1>Specific Bias Tests</h1><hr/>")

        debug.PrintX("<h2>Zero Constraints Test (is there specific bias?)<h2>")
        debug.PrintX("<table><tr><td></td><th>X<sup>2</sup></th><th>DF</th><th>Delta</th><th>p-value</th></tr>")
        debug.PrintX("<tr><th>Unconstrained Model</th><td>" & estimates(0).ToString("#0.000") & "</td><td>" & estimates(1).ToString & "</td><td rowspan=""2"">X<sup>2</sup>=" & FormatNumber(CDbl(Math.Abs((estimates(0) - estimates2(0))).ToString), 3) & "<br>DF=" & Math.Abs((estimates(1) - estimates2(1))).ToString & "</td><td rowspan=""2"">" & FormatNumber(CDbl(pValue1.ToString), 3) & "</td></tr>")
        debug.PrintX("<tr><th>Zero Constrained Model</th><td>" + estimates2(0).ToString("#0.000") & "</td><td>" & estimates2(1).ToString & "</td></tr></table><hr/>")

        debug.PrintX("<h2>Equal Constraints Test (is bias evenly distributed?)<h2>")
        debug.PrintX("<table><tr><td></td><th>X<sup>2</sup></th><th>DF</th><th>Delta</th><th>p-value</th></tr>")
        debug.PrintX("<tr><th>Unconstrained Model</th><td>" & estimates(0).ToString("#0.000") & "</td><td>" & estimates(1).ToString & "</td><td rowspan=""2"">X<sup>2</sup>=" & FormatNumber(CDbl(Math.Abs((estimates(0) - estimates3(0))).ToString), 3) & "<br>DF=" & Math.Abs((estimates(1) - estimates3(1))).ToString & "</td><td rowspan=""2"">" & FormatNumber(CDbl(pValue2.ToString), 3) & "</td></tr>")
        debug.PrintX("<tr><th>Equal Constrained Model</th><td>" + estimates3(0).ToString("#0.000") & "</td><td>" & estimates3(1).ToString & "</td></tr></table><hr/>")

        'Interpretation
        debug.PrintX("<h3>Conclusion</h3><p>" & conclusion & "</p><hr/>")

        'References
        debug.PrintX("<p>--If you would like to cite this tool directly, please use the following:")
        debug.PrintX("Gaskin, J. & Lim, J. (2017), ""CFA Tool"", AMOS Plugin. <a href=\""http://statwiki.kolobkreations.com"">Gaskination's StatWiki</a>.</p>")

        'Write style And close
        debug.PrintX("<style>table{border:1px solid black;border-collapse:collapse;}td{border:1px solid black;text-align:center;padding:5px;}th{text-weight:bold;padding:10px;border: 1px solid black;}</style>")
        debug.PrintX("</body></html>")

        'Take down our debugging, release file, open html
        Trace.Flush()
        Trace.Listeners.Remove(resultWriter)
        resultWriter.Close()
        resultWriter.Dispose()
        Process.Start("ModelBias.html")
    End Sub



End Class
