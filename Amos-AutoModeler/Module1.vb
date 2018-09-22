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

    Public Function Mainsub() As Integer Implements IPlugin.MainSub
        Amos.pd.GetCheckBox("AnalysisPropertiesForm", "ModsCheck").Checked = True
        Dim pattern As String = "\be.\b"
        Dim variable As PDElement
        Dim observedVariables As New Collections.ArrayList

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
            Amos.pd.AnalyzeCalculateEstimates()
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
                    E.NameOrCaption = "e" & x
                End If
            Next

        Next

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
        doc.Load(Amos.pd.ProjectName & ".AmosOutput")
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
