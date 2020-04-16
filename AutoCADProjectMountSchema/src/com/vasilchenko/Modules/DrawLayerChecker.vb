Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Colors
Imports Autodesk.AutoCAD.DatabaseServices

Namespace com.vasilchenko.Modules
    Module DrawLayerChecker

        Public Sub CheckLayers(acDocument As Document, acDatabase As Autodesk.AutoCAD.DatabaseServices.Database)
            Dim objLayerList As New List(Of String) _
                ({"PASSY", "PCAT", "PDESC", "PGRP", "PITEM", "PLOC", "PMFG", "PMISC", "PMNT", "PRTG", "PSYMS", "PTAG", "PTERM", "PWIRE", "LINK", "JUMPER", "Cable", "_MULTI_WIRE", "Монтажные отверстия", "Оси"})

            Using acTransanction As Transaction = acDatabase.TransactionManager.StartTransaction()
                Dim acLayerTbl = CType(acTransanction.GetObject(acDatabase.LayerTableId, OpenMode.ForRead), LayerTable)

                For Each strLayerName As String In objLayerList
                    If acLayerTbl.Has(strLayerName) = False Then
                        Using acLayerTblRec As LayerTableRecord = New LayerTableRecord()
                            acLayerTblRec.Name = strLayerName
                            Select Case strLayerName
                                Case "PASSY"
                                    acLayerTblRec.Color = Color.FromColorIndex(ColorMethod.ByAci, 134)
                                Case "PCAT"
                                    acLayerTblRec.Color = Color.FromColorIndex(ColorMethod.ByAci, 9)
                                Case "PDESC"
                                    acLayerTblRec.Color = Color.FromColorIndex(ColorMethod.ByAci, 7)
                                Case "PGRP"
                                    acLayerTblRec.Color = Color.FromColorIndex(ColorMethod.ByAci, 8)
                                Case "PITEM"
                                    acLayerTblRec.Color = Color.FromColorIndex(ColorMethod.ByAci, 20)
                                Case "PLOC"
                                    acLayerTblRec.Color = Color.FromColorIndex(ColorMethod.ByAci, 8)
                                Case "PMFG"
                                    acLayerTblRec.Color = Color.FromColorIndex(ColorMethod.ByAci, 9)
                                Case "PMISC"
                                    acLayerTblRec.Color = Color.FromColorIndex(ColorMethod.ByAci, 2)
                                Case "PMNT"
                                    acLayerTblRec.Color = Color.FromColorIndex(ColorMethod.ByAci, 8)
                                Case "PRTG"
                                    acLayerTblRec.Color = Color.FromColorIndex(ColorMethod.ByAci, 130)
                                Case "PSYMS"
                                    acLayerTblRec.Color = Color.FromColorIndex(ColorMethod.ByAci, 7)
                                Case "PTAG"
                                    acLayerTblRec.Color = Color.FromColorIndex(ColorMethod.ByAci, 51)
                                Case "PWIRE"
                                    acLayerTblRec.Color = Color.FromColorIndex(ColorMethod.ByAci, 6)
                                Case "PTERM"
                                    acLayerTblRec.Color = Color.FromColorIndex(ColorMethod.ByAci, 7)
                                Case "Cable"
                                    acLayerTblRec.Color = Color.FromColorIndex(ColorMethod.ByAci, 3)
                                    acLayerTblRec.LineWeight = LineWeight.LineWeight030
                                Case "Монтажные отверстия"
                                    acLayerTblRec.Color = Color.FromColorIndex(ColorMethod.ByAci, 1)
                                    acLayerTblRec.LineWeight = LineWeight.LineWeight050
                                Case "Оси"
                                    acLayerTblRec.Color = Color.FromColorIndex(ColorMethod.ByAci, 1)
                                    acLayerTblRec.LineWeight = LineWeight.LineWeight005
                            End Select

                            acLayerTbl.UpgradeOpen()
                            acLayerTbl.Add(acLayerTblRec)

                            acTransanction.AddNewlyCreatedDBObject(acLayerTblRec, True)
                        End Using
                    Else
                        Dim acCurLayer As LayerTableRecord = acTransanction.GetObject(acLayerTbl.Item(strLayerName), OpenMode.ForWrite)
                        acLayerTbl.UpgradeOpen()
                        Select Case strLayerName
                            Case "_MULTI_WIRE"
                                acCurLayer.Color = Color.FromColorIndex(ColorMethod.ByAci, 3)
                                acCurLayer.LineWeight = LineWeight.LineWeight030
                            Case "LINK"
                                acCurLayer.Color = Color.FromColorIndex(ColorMethod.ByAci, 11)
                            Case "JUMPER"
                                acCurLayer.Color = Color.FromColorIndex(ColorMethod.ByAci, 250)
                                acCurLayer.LineWeight = LineWeight.LineWeight005
                        End Select
                    End If
                Next
                acTransanction.Commit()
            End Using
        End Sub

    End Module
End Namespace
