Imports AutoCADProjectMountSchema.com.vasilchenko.Classes
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry

Namespace com.vasilchenko.Modules
    Module ElementBuilder
        Friend Sub CreateElementList(strLocation As String)

            Dim acDocument As Document = Application.DocumentManager.MdiActiveDocument
            Dim acDatabase As Database = acDocument.Database

            Dim acElementList As SortedList(Of String, ElementClass) = DatabaseConnection.GetElementsInLocation(strLocation)

            DrawLayerChecker.CheckLayers(acDocument, acDatabase)

            Dim acGroupedSortedElementList = From acElement In acElementList
                                             Order By If(acElement.Value.CatalogName.Equals("IGNORED"), "IGNORED", acElement.Value.CatalogName)

            Dim acEditor As Editor = acDocument.Editor
            Dim acCurManufacture As String = ""
            Dim acInsertPt As Point3d

            For Each acElements In acGroupedSortedElementList

                If (acCurManufacture.Equals("") OrElse (acCurManufacture.Equals("IGNORED") And Not acElements.Value.Manufacture.Equals(acCurManufacture))) Then
                    acCurManufacture = acElements.Value.Manufacture
                    Dim strPromtDescription = IIf(acCurManufacture.Equals("IGNORED"), "Выберите точку вставки для существующих объектов" & vbCrLf, "Выберите точку вставки для новых объектов" & vbCrLf)
                    Dim acPromptPntOpt As New PromptPointOptions(strPromtDescription)
                    Dim acPromptPntResult As PromptPointResult = AcEditor.GetPoint(acPromptPntOpt)
                    If acPromptPntResult.Status <> PromptStatus.OK Then
                        AcEditor.WriteMessage("Отмена вставки" & vbCrLf)
                        Exit Sub
                    End If
                    acInsertPt = acPromptPntResult.Value
                End If

                DrawBlocks(acDatabase, acElements.Value, acInsertPt)

            Next

        End Sub

        Private Sub DrawBlocks(acDatabase As Database, acElement As ElementClass, ByRef acInsertPt As Point3d)

            Using acTransaction As Transaction = acDatabase.TransactionManager.StartTransaction
                Dim acBlockTable As BlockTable = Nothing
                Dim acInsObjectId As ObjectId = Nothing

                Dim strBlkName As String = SymbolUtilityServices.GetBlockNameFromInsertPathName(acElement.BlockPath)
                acBlockTable = acTransaction.GetObject(acDatabase.BlockTableId, OpenMode.ForRead)

                If acBlockTable.Has(strBlkName) Then
                    Dim acCurrBlkTblRcd As BlockTableRecord = acTransaction.GetObject(acBlockTable.Item(strBlkName), OpenMode.ForRead)
                    acInsObjectId = acCurrBlkTblRcd.Id
                Else
                    Try
                        Dim acNewDbDwg As New Autodesk.AutoCAD.DatabaseServices.Database(False, True)
                        acNewDbDwg.ReadDwgFile(acElement.BlockPath, FileOpenMode.OpenTryForReadShare, True, "")
                        acInsObjectId = acDatabase.Insert(strBlkName, acNewDbDwg, True)
                        acNewDbDwg.Dispose()
                    Catch ex As Exception
                        MsgBox(String.Format("Не найден графический образ с именем {0}", strBlkName))
                        Exit Sub
                    End Try
                End If

                Using acBlkRef As New BlockReference(acInsertPt, acInsObjectId)
                    acBlkRef.Layer = "PSYMS"
                    acBlkRef.ScaleFactors = New Scale3d(1.0)
                    Dim acBlockTableRecord As BlockTableRecord = acTransaction.GetObject(acBlockTable.Item(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                    acBlockTableRecord.AppendEntity(acBlkRef)
                    acTransaction.AddNewlyCreatedDBObject(acBlkRef, True)
                    Dim acBlockTableAttrRec As BlockTableRecord = acTransaction.GetObject(acInsObjectId, OpenMode.ForRead)
                    Dim acAttrObjectId As ObjectId

                    For Each acAttrObjectId In acBlockTableAttrRec
                        Dim acAttrEntity As Entity = acTransaction.GetObject(acAttrObjectId, OpenMode.ForWrite)
                        Dim acAttrDefinition As AttributeDefinition = TryCast(acAttrEntity, AttributeDefinition)
                        If (acAttrDefinition IsNot Nothing) Then
                            Dim acAttrReference As New AttributeReference()
                            acAttrReference.SetAttributeFromBlock(acAttrDefinition, acBlkRef.BlockTransform)
                            Select Case acAttrReference.Tag
                                Case "WIDTH"
                                    acElement.Width = CDbl(acAttrReference.TextString)
                                Case Else
                                    If Not acElement.Manufacture.Equals("IGNORED") And acAttrReference.Tag.StartsWith("TERM") Then
                                        If acElement.Connections.Exists(Function(x) acAttrReference.TextString.Contains(x.ConnectionPin.Pin)) Then
                                            acElement.Connections.Find(Function(x) acAttrReference.TextString.Contains(x.ConnectionPin.Pin)).BlockTagNumber = Strings.Right(acAttrReference.Tag, 2)
                                        End If
                                    End If
                            End Select
                        End If
                    Next

                    For Each acAttrObjectId In acBlockTableAttrRec
                        Dim acAttrEntity As Entity = acTransaction.GetObject(acAttrObjectId, OpenMode.ForRead)
                        Dim acAttrDefinition = TryCast(acAttrEntity, AttributeDefinition)
                        If (acAttrDefinition IsNot Nothing) Then
                            Using acAttrReference As New AttributeReference()
                                acAttrReference.SetAttributeFromBlock(acAttrDefinition, acBlkRef.BlockTransform)
                                Select Case acAttrReference.Tag
                                    Case "INST"
                                        acAttrReference.TextString = acElement.Instance
                                        acAttrReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD_IEC")
                                        acAttrReference.Layer = "PLOC"
                                    Case "LOC"
                                        acAttrReference.TextString = acElement.Location
                                        acAttrReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD_IEC")
                                        acAttrReference.Layer = "PLOC"
                                    Case "MFG"
                                        acAttrReference.TextString = acElement.Manufacture
                                        acAttrReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD_IEC")
                                        acAttrReference.Layer = "PMFG"
                                    Case "CAT"
                                        acAttrReference.TextString = acElement.CatalogName
                                        acAttrReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD_IEC")
                                        acAttrReference.Layer = "PCAT"
                                    Case "P_TAG1"
                                        acAttrReference.TextString = acElement.Tag
                                        acAttrReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD_IEC")
                                        acAttrReference.Layer = "PTAG"
                                    Case "DESC1"
                                        acAttrReference.TextString = acElement.Desc1
                                        acAttrReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD_IEC")
                                        acAttrReference.Layer = "PTAG"
                                    Case "DESC2"
                                        acAttrReference.TextString = acElement.Desc2
                                        acAttrReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD_IEC")
                                        acAttrReference.Layer = "PTAG"
                                    Case "DESC3"
                                        acAttrReference.TextString = acElement.Desc3
                                        acAttrReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD_IEC")
                                        acAttrReference.Layer = "PTAG"
                                    Case Else
                                        If acAttrReference.Tag.StartsWith("TERM") And Not acAttrReference.Tag.StartsWith("TERMDESC") Then
                                            If acElement.Connections.Exists(Function(x) acAttrReference.Tag.Equals("TERM" & x.BlockTagNumber)) Then
                                                acAttrReference.TextString = acElement.Connections.Find(Function(x) acAttrReference.Tag.Equals("TERM" & x.BlockTagNumber)).ConnectionPin.Pin
                                                acAttrReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD_IEC")
                                                acAttrReference.Layer = "PTERM"
                                            Else
                                                acAttrReference.TextString = acAttrReference.TextString.Split(New Char() {","})(0)
                                            End If
                                        ElseIf acAttrReference.Tag.StartsWith("TERMDESC") Then
                                            If acElement.Connections.Exists(Function(x) acAttrReference.Tag.Equals("TERMDESC" & x.BlockTagNumber)) Then
                                                acAttrReference.TextString = acElement.Connections.Find(Function(x) acAttrReference.Tag.Equals("TERMDESC" & x.BlockTagNumber)).ConnnectionToAsString
                                                acAttrReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD_IEC")
                                                acAttrReference.Layer = "PDESC"
                                            End If
                                        ElseIf acAttrReference.Tag.StartsWith("WIRENO") Then
                                            If acElement.Connections.Exists(Function(x) acAttrReference.Tag.Equals("WIRENO" & x.BlockTagNumber)) Then
                                                acAttrReference.TextString = acElement.Connections.Find(Function(x) acAttrReference.Tag.Equals("WIRENO" & x.BlockTagNumber)).ConnectionPin.Wireno
                                                acAttrReference.TextStyleId = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead), TextStyleTable)("WD_IEC")
                                                acAttrReference.Layer = "PWIRE"
                                            End If
                                        End If
                                End Select

                                acBlkRef.AttributeCollection.AppendAttribute(acAttrReference)
                                acTransaction.AddNewlyCreatedDBObject(acAttrReference, True)
                            End Using
                        End If
                    Next
                End Using
                acInsertPt = New Point3d(acInsertPt.X + acElement.Width + 5, acInsertPt.Y, acInsertPt.Z)
                acTransaction.Commit()
            End Using
        End Sub
    End Module
End Namespace

