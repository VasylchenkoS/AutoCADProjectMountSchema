Imports AutoCADProjectMountSchema.com.vasilchenko.Classes
Imports AutoCADProjectMountSchema.com.vasilchenko.Modules
Imports Autodesk.AutoCAD.ApplicationServices

Namespace com.vasilchenko.DatabaseConnection
    Module DatabaseDataAccessObject

        Private ReadOnly DatabaseTableList As New List(Of String)
        Friend Function GetLocations() As List(Of String)
            Dim locationsList As New List(Of String)

            Dim strSqlQuery = "SELECT DISTINCT [LOC] FROM [COMP] 
WHERE ([INST] LIKE '7[00-99]%' OR [INST] IS NULL) AND [LOC] IS NOT NULL 
UNION 
SELECT [LOC] FROM [TERMS] 
WHERE ([INST] LIKE '7[00-99]%' OR [INST] IS NULL) AND [LOC] IS NOT NULL 
ORDER BY [LOC]"

            Dim dbOleDataTable As Data.DataTable = DatabaseConnections.GetOleDataTable(strSqlQuery)
            If Not IsNothing(dbOleDataTable) Then
                For Each dbRow In dbOleDataTable.Rows
                    locationsList.Add(dbRow.Item("LOC"))
                Next
            End If

            Return locationsList
        End Function
        Friend Function GetElementsInLocation(strLocation As String) As SortedList(Of String, ElementClass)
            Dim acElementsList As New SortedList(Of String, ElementClass)
            Dim sqlQuery = String.Format("SELECT DISTINCT [TAGNAME], [FAMILY], [MFG], [CAT], [INST], [PAR1_CHLD2], [DESC1], [DESC2], [DESC3] FROM [COMP] WHERE [LOC] = '{0}' AND [TAGNAME] IS NOT NULL ORDER BY [PAR1_CHLD2], [TAGNAME]", strLocation)

            Dim objDataTable As Data.DataTable = GetOleDataTable(sqlQuery)
            If Not IsNothing(objDataTable) Then
                For Each dbRow In objDataTable.Rows
                    If Not acElementsList.ContainsKey((dbRow.item("TAGNAME"))) Then
                        Dim acElement As New ElementClass With {
                    .Location = strLocation
                }
                        If Not IsDBNull(dbRow.item("TAGNAME")) Then acElement.Tag = dbRow.item("TAGNAME")
                        If Not IsDBNull(dbRow.item("FAMILY")) Then acElement.Family = dbRow.item("FAMILY") Else acElement.Family = ""
                        If Not IsDBNull(dbRow.item("INST")) Then acElement.Instance = dbRow.item("INST") Else acElement.Instance = ""
                        If Not IsDBNull(dbRow.item("MFG")) Then acElement.Manufacture = dbRow.item("MFG") Else acElement.Manufacture = "IGNORED"
                        If Not IsDBNull(dbRow.item("CAT")) Then acElement.CatalogName = dbRow.item("CAT") Else acElement.CatalogName = "IGNORED"
                        If Not IsDBNull(dbRow.item("DESC1")) Then acElement.Desc3 = dbRow.item("DESC1") Else acElement.Desc1 = ""
                        If Not IsDBNull(dbRow.item("DESC2")) Then acElement.Desc3 = dbRow.item("DESC2") Else acElement.Desc2 = ""
                        If Not IsDBNull(dbRow.item("DESC3")) Then acElement.Desc3 = dbRow.item("DESC3") Else acElement.Desc3 = ""

                        FillConnection(acElement)


                        If Not IsNothing(acElement.Connections) Then
                            FillBlockPath(acElement)

                            If acElement.Manufacture.Equals("IGNORED") Then
                                acElement.Connections.Sort(Function(x, y) AdditionalFunctions.GetFirstNumericFromString(x.ConnectionPin.Pin).CompareTo(AdditionalFunctions.GetFirstNumericFromString(y.ConnectionPin.Pin)))
                                For i = 1 To acElement.Connections.Count
                                    acElement.Connections(i - 1).BlockTagNumber = Format(i, "0#")
                                Next
                            End If

                            acElementsList.Add(acElement.Tag, acElement)
                        End If

                    End If
                Next
            End If
            Return acElementsList
        End Function

        Private Sub FillConnection(acElement As ElementClass)
            Dim sqlQuery = String.Format("SELECT DISTINCT [WIRENO], [NAM1], [PIN1], [LOC1], [NAM2], [PIN2], [LOC2], [CBL] FROM [WFRM2ALL] WHERE ([NAM1] IS NOT NULL AND [NAM2] IS NOT NULL) AND ([NAM1] = '{1}' OR [NAM2] = '{1}')
AND ([NAM1] IN (SELECT [TAGNAME] FROM [COMP] WHERE [LOC] = '{0}') OR [NAM2] IN (SELECT [TAGNAME] FROM [COMP] WHERE [LOC] = '{0}')) AND ([LOC1]='{0}' OR [LOC2]='{0}')", acElement.Location, acElement.Tag)

            Dim dbOleDataTable As Data.DataTable = GetOleDataTable(sqlQuery)
            If Not (IsNothing(dbOleDataTable) Or dbOleDataTable.Rows.Count = 0) Then
                acElement.Connections = New List(Of ConnectionClass)
                For Each dbRow In dbOleDataTable.Rows
                    Dim acPinCur As New PinClass
                    Dim acPinTo As New PinClass
                    If dbRow.Item("NAM1").Equals(acElement.Tag) Then
                        acPinCur.TagName = dbRow.Item("NAM1").ToString
                        acPinCur.Pin = dbRow.Item("PIN1").ToString
                        acPinTo.TagName = dbRow.Item("NAM2").ToString
                        acPinTo.Pin = dbRow.Item("PIN2").ToString
                    Else
                        acPinCur.TagName = dbRow.Item("NAM2").ToString
                        acPinCur.Pin = dbRow.Item("PIN2").ToString
                        acPinTo.TagName = dbRow.Item("NAM1").ToString
                        acPinTo.Pin = dbRow.Item("PIN1").ToString
                    End If

                    If Not IsDBNull(dbRow.Item("WIRENO")) Then
                        acPinCur.Wireno = dbRow.Item("WIRENO")
                        acPinTo.Wireno = dbRow.Item("WIRENO")
                    Else
                        Core.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("[WARNING!]:Для вывода {0}:{1} не назначено номера провода {2}", acPinCur.TagName, acPinCur.Pin, vbCrLf)
                        Exit Sub
                    End If

                    If Not IsDBNull(dbRow.Item("CBL")) Then
                        acPinCur.Wireno = dbRow.CableName("CBL")
                        acPinTo.CableName = dbRow.Item("CBL")
                    End If

                    If IsNothing(acElement.Connections.Find(Function(x) x.ConnectionPin.TagName = acPinCur.TagName And x.ConnectionPin.Pin = acPinCur.Pin)) Then
                        Dim acConnection As New ConnectionClass With {
                    .ConnectionPin = acPinCur,
                    .ConnectionTo = New List(Of PinClass)({acPinTo})
                    }
                        acElement.Connections.Add(acConnection)
                    Else
                        acElement.Connections.Find(Function(x) x.ConnectionPin.TagName = acPinCur.TagName And x.ConnectionPin.Pin = acPinCur.Pin).ConnectionTo.Add(acPinTo)
                    End If
                Next
            End If
        End Sub

        Private Sub FillBlockPath(ByRef acElement As ElementClass)
            If acElement.CatalogName.Equals("IGNORED") Then
                acElement.BlockPath = String.Format("{0}MSH_04_{1}.dwg", My.Resources.MountGraphicPathString, Format(acElement.Connections.Count, "0#"))
            Else
                Dim strCatName = acElement.CatalogName.Split(New Char() {" ", "-", "/", "."})(0)
                If strCatName.Length < 4 Then strCatName = String.Concat(strCatName, "_", acElement.CatalogName.Split(New Char() {" ", "-", "/", "."})(1))
                acElement.BlockPath = String.Format("{0}MSH_04_{1}.dwg", My.Resources.MountGraphicPathString, strCatName)
            End If

            If Not IO.File.Exists(acElement.BlockPath) Then
                Throw New ArgumentNullException(String.Format("[WARNING!]:Для елемента {0} {1}  не сущесвует компоновочного образа{2}{3}", acElement.Manufacture, acElement.CatalogName, IIf(acElement.CatalogName.Equals("IGNORED"), "Количество соединений - " & acElement.Connections.Count, ""), vbCrLf))
            End If

        End Sub

    End Module
End Namespace
