Option Explicit

' 改修履歴: CHANGELOG.md 参照
' 施工指示書取込: 公開 API ファサード（実装は分割モジュールへ委譲）

Public Sub ImportConstructionDocument()
    mod_Construction_Import_Load.ImportConstructionDocument
End Sub

Public Sub RefreshSubcontractorPriceColumns(ByVal ws As Worksheet, _
                                            Optional ByVal changedRows As Collection = Nothing)
    mod_Construction_SubconPrice.RefreshSubcontractorPriceColumns ws, changedRows
End Sub

Public Sub RefreshBasicInfoConstructionTotals(Optional ByVal changedVendorIndex As Long = 0)
    mod_Construction_BasicTotals.RefreshBasicInfoConstructionTotals changedVendorIndex
End Sub

Public Sub ClearVendorAliasMapCache()
    mod_Construction_BasicTotals.ClearVendorAliasMapCache
End Sub

Public Sub UpdateBasicInfoTaxTotals(Optional ByVal wsInfo As Worksheet)
    mod_Construction_BasicTotals.UpdateBasicInfoTaxTotals wsInfo
End Sub

Public Function ResolveBasicInfoVendorInfoIndex(ByVal vendorDisplayName As String, _
                                                Optional ByVal workTypeKeyword As String = "") As Long
    ResolveBasicInfoVendorInfoIndex = mod_Construction_BasicTotals.ResolveBasicInfoVendorInfoIndex(vendorDisplayName, workTypeKeyword)
End Function

Public Sub ApplySanpaiRowRestrictions(ByVal ws As Worksheet)
    mod_Construction_OutputFormat.ApplySanpaiRowRestrictions ws
End Sub

Public Sub RefreshConstructionReferenceUnitPricesOnExistingSheets()
    mod_Construction_OutputFormat.RefreshConstructionReferenceUnitPricesOnExistingSheets
End Sub

Public Sub RefreshConstructionReferencePricesForUnitPriceChange( _
    ByVal wsUnitPrice As Worksheet, ByVal changedRange As Range)
    mod_Construction_OutputFormat.RefreshConstructionReferencePricesForUnitPriceChange wsUnitPrice, changedRange
End Sub

Public Function GetProjectMasterLineOrderRank(ByVal lineName As String) As Long
    GetProjectMasterLineOrderRank = mod_Construction_LineMapping.GetProjectMasterLineOrderRank(lineName)
End Function

Public Sub ClearProjectMasterLineOrderCache()
    mod_Construction_LineMapping.ClearProjectMasterLineOrderCache
End Sub

Public Function IsManagedConstructionImportOutputSheet(ByVal ws As Worksheet) As Boolean
    IsManagedConstructionImportOutputSheet = mod_Construction_OutputLayout.IsManagedConstructionImportOutputSheet(ws)
End Function

Public Function IsWeldingOutputSheet(ByVal ws As Worksheet) As Boolean
    IsWeldingOutputSheet = mod_Construction_OutputLayout.IsWeldingOutputSheet(ws)
End Function

Public Function OutputSheetCol(ByVal ws As Worksheet, ByVal baseCol As Long) As Long
    OutputSheetCol = mod_Construction_OutputLayout.OutputSheetCol(ws, baseCol)
End Function

Public Function OutputSheetSeiriColumn(ByVal ws As Worksheet) As Long
    OutputSheetSeiriColumn = mod_Construction_OutputLayout.OutputSheetSeiriColumn(ws)
End Function

Public Function OutputSheetSubconPriceFirstCol(ByVal ws As Worksheet) As Long
    OutputSheetSubconPriceFirstCol = mod_Construction_OutputLayout.OutputSheetSubconPriceFirstCol(ws)
End Function

Public Function OutputSheetVendorColumns(ByVal ws As Worksheet) As Collection
    Set OutputSheetVendorColumns = mod_Construction_OutputLayout.OutputSheetVendorColumns(ws)
End Function
