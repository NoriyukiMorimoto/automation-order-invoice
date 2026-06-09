Option Explicit

' UI message text (code points avoid import encoding issues)

Public Function UiMsgBasicInfoSheetNotFoundText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H57FA,&H672C,&H60C5,&H5831,&H30B7,&H30FC,&H30C8,&H304C,&H898B,&H3064,&H304B,&H308A,&H307E,&H305B,&H3093,&H3002)
    UiMsgBasicInfoSheetNotFoundText = cached
End Function

Public Function UiMsgBasicInfoSheetNotFoundCheckNameText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H57FA,&H672C,&H60C5,&H5831,&H30B7,&H30FC,&H30C8,&H304C,&H898B,&H3064,&H304B,&H308A,&H307E,&H305B,&H3093,&H3002,&H30B7,&H30FC,&H30C8,&H540D,&H3092,&H78BA,&H8A8D,&H3057,&H3066,&H304F,&H3060,&H3055,&H3044,&H3002)
    UiMsgBasicInfoSheetNotFoundCheckNameText = cached
End Function

Public Function UiMsgUnitPriceImportableSheetNotFoundText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H5358,&H4FA1,&H8868,&H30D6,&H30C3,&H30AF,&H306B,&H53D6,&H308A,&H8FBC,&H307F,&H53EF,&H80FD,&H306A,&H30B7,&H30FC,&H30C8,&H304C,&H898B,&H3064,&H304B,&H308A,&H307E,&H305B,&H3093,&H3067,&H3057,&H305F,&H3002)
    UiMsgUnitPriceImportableSheetNotFoundText = cached
End Function

Public Function UiMsgImportCompleteTitleText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H5B8C,&H4E86)
    UiMsgImportCompleteTitleText = cached
End Function

Public Function UiMsgBasicInfoYearNotFoundB4Text() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H57FA,&H672C,&H60C5,&H5831,&H30B7,&H30FC,&H30C8,&H0020,&H0042,&H0034,&H0020,&H306B,&H0034,&H6841,&H306E,&H5E74,&H5EA6,&H304C,&H898B,&H3064,&H304B,&H308A,&H307E,&H305B,&H3093,&H3002)
    UiMsgBasicInfoYearNotFoundB4Text = cached
End Function

Public Function UiMsgBasicInfoYearNotFoundB4ExampleText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H57FA,&H672C,&H60C5,&H5831,&H30B7,&H30FC,&H30C8,&H0020,&H0042,&H0034,&H0020,&H306B,&H0034,&H6841,&H306E,&H5E74,&H5EA6,&H304C,&H898B,&H3064,&H304B,&H308A,&H307E,&H305B,&H3093,&H3002,&H4F8B,&H003A,&H0020,&H0032,&H0030,&H0032,&H0036)
    UiMsgBasicInfoYearNotFoundB4ExampleText = cached
End Function

Public Function UiMsgBasicInfoBranchOfficeEmptyText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H57FA,&H672C,&H60C5,&H5831,&H30B7,&H30FC,&H30C8,&H0020,&H0042,&H0036,&H0020,&H307E,&H305F,&H306F,&H0020,&H0043,&H0036,&H0020,&H304C,&H7A7A,&H3067,&H3059,&H3002,&H652F,&H5E97,&H540D,&H30FB,&H51FA,&H5F35,&H6240,&H540D,&H3092,&H78BA,&H8A8D,&H3057,&H3066,&H304F,&H3060,&H3055,&H3044,&H3002)
    UiMsgBasicInfoBranchOfficeEmptyText = cached
End Function

Public Function UiMsgBasicInfoLineTypeEmptyC20Text() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H57FA,&H672C,&H60C5,&H5831,&H30B7,&H30FC,&H30C8,&H0020,&H0043,&H0032,&H0030,&H0020,&H306E,&H7DDA,&H533A,&H533A,&H5206,&H3092,&H9078,&H629E,&H3057,&H3066,&H304F,&H3060,&H3055,&H3044,&H3002)
    UiMsgBasicInfoLineTypeEmptyC20Text = cached
End Function

Public Function UiMsgBasicInfoProjectNameEmptyC21Text() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H57FA,&H672C,&H60C5,&H5831,&H30B7,&H30FC,&H30C8,&H0020,&H0043,&H0032,&H0031,&H0020,&H306E,&H5358,&H4FA1,&H9069,&H7528,&H5DE5,&H4E8B,&H4EF6,&H540D,&H3092,&H9078,&H629E,&H3057,&H3066,&H304F,&H3060,&H3055,&H3044,&H3002)
    UiMsgBasicInfoProjectNameEmptyC21Text = cached
End Function

Public Function UiMsgBasicInfoProjectNamePurchaseExcludedC21Text() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H57FA,&H672C,&H60C5,&H5831,&H30B7,&H30FC,&H30C8,&H0020,&H0043,&H0032,&H0031,&H0020,&H306F,&H8ECC,&H9053,&H6750,&H6599,&H8CFC,&H5165,&H5145,&H5F53,&H4EE5,&H5916,&H306E,&H5358,&H4FA1,&H9069,&H7528,&H5DE5,&H4E8B,&H4EF6,&H540D,&H3092,&H9078,&H629E,&H3057,&H3066,&H304F,&H3060,&H3055,&H3044,&H3002)
    UiMsgBasicInfoProjectNamePurchaseExcludedC21Text = cached
End Function

Public Function UiMsgPurchaseUnitPriceAutoCreateAfterC24Text() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H8CFC,&H5165,&H5145,&H5F53,&H5358,&H4FA1,&H306F,&H0020,&H0043,&H0032,&H0034,&H0020,&H78BA,&H5B9A,&H5F8C,&H306B,&H81EA,&H52D5,&H4F5C,&H6210,&H3057,&H307E,&H3059,&H3002)
    UiMsgPurchaseUnitPriceAutoCreateAfterC24Text = cached
End Function

Public Function UiMsgBasicInfoPriceKindEmptyC22Text() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H57FA,&H672C,&H60C5,&H5831,&H30B7,&H30FC,&H30C8,&H0020,&H0043,&H0032,&H0032,&H0020,&H306E,&H5358,&H4FA1,&H533A,&H5206,&H3092,&H9078,&H629E,&H3057,&H3066,&H304F,&H3060,&H3055,&H3044,&H3002)
    UiMsgBasicInfoPriceKindEmptyC22Text = cached
End Function

Public Function UiMsgUnitPriceMasterFileNotFoundText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H51FA,&H5F35,&H6240,&H5225,&H005F,&H5358,&H4FA1,&H9069,&H7528,&H7DDA,&H533A,&H002E,&H0078,&H006C,&H0073,&H0078,&H0020,&H304C,&H898B,&H3064,&H304B,&H308A,&H307E,&H305B,&H3093,&H3002)
    UiMsgUnitPriceMasterFileNotFoundText = cached
End Function

Public Function UiMsgUnitPriceMasterFileUnreadableText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H51FA,&H5F35,&H6240,&H5225,&H005F,&H5358,&H4FA1,&H9069,&H7528,&H7DDA,&H533A,&H002E,&H0078,&H006C,&H0073,&H0078,&H0020,&H3092,&H53C2,&H7167,&H3067,&H304D,&H307E,&H305B,&H3093,&H3067,&H3057,&H305F,&H3002)
    UiMsgUnitPriceMasterFileUnreadableText = cached
End Function

Public Function UiMsgUnitPriceMasterSheetNotFoundText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H5358,&H4FA1,&H9069,&H7528,&H7DDA,&H533A,&H30B7,&H30FC,&H30C8,&H304C,&H898B,&H3064,&H304B,&H308A,&H307E,&H305B,&H3093,&H3067,&H3057,&H305F,&H3002)
    UiMsgUnitPriceMasterSheetNotFoundText = cached
End Function

Public Function UiMsgUnitPriceMasterSheetNameCheckText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H30D6,&H30C3,&H30AF,&H5185,&H306E,&H30B7,&H30FC,&H30C8,&H540D,&H3092,&H78BA,&H8A8D,&H3057,&H3066,&H304F,&H3060,&H3055,&H3044,&H3002)
    UiMsgUnitPriceMasterSheetNameCheckText = cached
End Function

Public Function UiMsgUnitPriceMasterLineTypeAmbiguousText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H5358,&H4FA1,&H9069,&H7528,&H7DDA,&H533A,&H30B7,&H30FC,&H30C8,&H306B,&H652F,&H5E97,&H30FB,&H51FA,&H5F35,&H6240,&H306F,&H898B,&H3064,&H304B,&H308A,&H307E,&H3057,&H305F,&H304C,&H3001,&H7DDA,&H533A,&H533A,&H5206,&H306B,&H4E00,&H81F4,&H3059,&H308B,&H5358,&H4FA1,&H9069,&H7528,&H4FDD,&H7DDA,&H533A,&H3092,&H7279,&H5B9A,&H3067,&H304D,&H307E,&H305B,&H3093,&H3067,&H3057,&H305F,&H3002)
    UiMsgUnitPriceMasterLineTypeAmbiguousText = cached
End Function

Public Function UiMsgUnitPriceMasterBranchOfficeNotFoundText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H5358,&H4FA1,&H9069,&H7528,&H7DDA,&H533A,&H30B7,&H30FC,&H30C8,&H306B,&H8A72,&H5F53,&H3059,&H308B,&H652F,&H5E97,&H30FB,&H51FA,&H5F35,&H6240,&H304C,&H898B,&H3064,&H304B,&H308A,&H307E,&H305B,&H3093,&H3067,&H3057,&H305F,&H3002)
    UiMsgUnitPriceMasterBranchOfficeNotFoundText = cached
End Function

Public Function UiMsgBranchLabelText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H652F,&H5E97,&HFF1A)
    UiMsgBranchLabelText = cached
End Function

Public Function UiMsgBranchNameLabelText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H652F,&H5E97,&H540D,&HFF1A)
    UiMsgBranchNameLabelText = cached
End Function

Public Function UiMsgOfficeLabelText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H51FA,&H5F35,&H6240,&HFF1A)
    UiMsgOfficeLabelText = cached
End Function

Public Function UiMsgOfficeNameLabelText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H51FA,&H5F35,&H6240,&H540D,&HFF1A)
    UiMsgOfficeNameLabelText = cached
End Function

Public Function UiMsgLineTypeLabelText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H7DDA,&H533A,&H533A,&H5206,&HFF1A)
    UiMsgLineTypeLabelText = cached
End Function

Public Function UiMsgProjectNameLabelText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H5DE5,&H4E8B,&H4EF6,&H540D,&HFF1A)
    UiMsgProjectNameLabelText = cached
End Function

Public Function UiMsgUnitPriceMasterLoadFailedText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H5358,&H4FA1,&H9069,&H7528,&H7DDA,&H533A,&H30C7,&H30FC,&H30BF,&H306E,&H8AAD,&H307F,&H8FBC,&H307F,&H306B,&H5931,&H6557,&H3057,&H307E,&H3057,&H305F,&H3002)
    UiMsgUnitPriceMasterLoadFailedText = cached
End Function

Public Function UiMsgUnitPriceBookByProjectNotFoundText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H5DE5,&H4E8B,&H4EF6,&H540D,&H306B,&H4E00,&H81F4,&H3059,&H308B,&H5358,&H4FA1,&H8868,&H304C,&H898B,&H3064,&H304B,&H308A,&H307E,&H305B,&H3093,&H3002)
    UiMsgUnitPriceBookByProjectNotFoundText = cached
End Function

Public Function UiMsgUnitPriceDataFolderNotFoundText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H5358,&H4FA1,&H30C7,&H30FC,&H30BF,&H30D5,&H30A9,&H30EB,&H30C0,&H304C,&H898B,&H3064,&H304B,&H308A,&H307E,&H305B,&H3093,&H3002)
    UiMsgUnitPriceDataFolderNotFoundText = cached
End Function

Public Function UiMsgLineTypeFolderNotFoundText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H7DDA,&H533A,&H533A,&H5206,&H30D5,&H30A9,&H30EB,&H30C0,&H304C,&H898B,&H3064,&H304B,&H308A,&H307E,&H305B,&H3093,&H3002)
    UiMsgLineTypeFolderNotFoundText = cached
End Function

Public Function UiMsgBranchGroupFolderNotFoundText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H652F,&H793E,&H30D5,&H30A9,&H30EB,&H30C0,&H304C,&H898B,&H3064,&H304B,&H308A,&H307E,&H305B,&H3093,&H3002)
    UiMsgBranchGroupFolderNotFoundText = cached
End Function

Public Function UiMsgUnitPriceSectionFolderNotFoundText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H5358,&H4FA1,&H9069,&H7528,&H4FDD,&H7DDA,&H533A,&H30D5,&H30A9,&H30EB,&H30C0,&H304C,&H898B,&H3064,&H304B,&H308A,&H307E,&H305B,&H3093,&H3002)
    UiMsgUnitPriceSectionFolderNotFoundText = cached
End Function

Public Function UiMsgYearFolderNotFoundText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H5E74,&H5EA6,&H30D5,&H30A9,&H30EB,&H30C0,&H304C,&H898B,&H3064,&H304B,&H308A,&H307E,&H305B,&H3093,&H3002)
    UiMsgYearFolderNotFoundText = cached
End Function

Public Function UiMsgPriceKindFolderNotFoundText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H5358,&H4FA1,&H533A,&H5206,&H30D5,&H30A9,&H30EB,&H30C0,&H304C,&H898B,&H3064,&H304B,&H308A,&H307E,&H305B,&H3093,&H3002)
    UiMsgPriceKindFolderNotFoundText = cached
End Function

Public Function UiMsgUnitPriceBookOpenFailedText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H5358,&H4FA1,&H8868,&H30D6,&H30C3,&H30AF,&H3092,&H958B,&H3051,&H307E,&H305B,&H3093,&H3067,&H3057,&H305F,&H3002)
    UiMsgUnitPriceBookOpenFailedText = cached
End Function

Public Function UiMsgLineNameFormShowFailedText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H7A4D,&H7B97,&H7DDA,&H533A,&H9078,&H629E,&H30D5,&H30A9,&H30FC,&H30E0,&H3092,&H8868,&H793A,&H3067,&H304D,&H307E,&H305B,&H3093,&H3067,&H3057,&H305F,&H3002)
    UiMsgLineNameFormShowFailedText = cached
End Function

Public Function UiMsgUnitPriceImportFailedText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H5358,&H4FA1,&H8868,&H306E,&H53D6,&H308A,&H8FBC,&H307F,&H306B,&H5931,&H6557,&H3057,&H307E,&H3057,&H305F,&H3002)
    UiMsgUnitPriceImportFailedText = cached
End Function

Public Function UiMsgPurchaseUnitPriceImportFailedText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H8CFC,&H5165,&H5145,&H5F53,&H5358,&H4FA1,&H8868,&H306E,&H53D6,&H308A,&H8FBC,&H307F,&H306B,&H5931,&H6557,&H3057,&H307E,&H3057,&H305F,&H3002)
    UiMsgPurchaseUnitPriceImportFailedText = cached
End Function

Public Function UiMsgWeldingUnitPriceImportFailedText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H30EC,&H30FC,&H30EB,&H6EB6,&H63A5,&H5358,&H4FA1,&H8868,&H306E,&H53D6,&H308A,&H8FBC,&H307F,&H306B,&H5931,&H6557,&H3057,&H307E,&H3057,&H305F,&H3002)
    UiMsgWeldingUnitPriceImportFailedText = cached
End Function

Public Function UiMsgPurchaseReferenceKeyNotFoundText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H8CFC,&H5165,&H5145,&H5F53,&H5358,&H4FA1,&H8868,&H306E,&H53C2,&H7167,&H30AD,&H30FC,&H3092,&H53D6,&H5F97,&H3067,&H304D,&H307E,&H305B,&H3093,&H3067,&H3057,&H305F,&H3002)
    UiMsgPurchaseReferenceKeyNotFoundText = cached
End Function

Public Function UiMsgPurchaseSheetByKeyNotFoundPrefixText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H8CFC,&H5165,&H5145,&H5F53,&H5358,&H4FA1,&H8868,&H306B,&H53C2,&H7167,&H30AD,&H30FC,&H300C)
    UiMsgPurchaseSheetByKeyNotFoundPrefixText = cached
End Function

Public Function UiMsgPurchaseSheetByKeyNotFoundSuffixText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H300D,&H306B,&H4E00,&H81F4,&H3059,&H308B,&H30B7,&H30FC,&H30C8,&H304C,&H898B,&H3064,&H304B,&H308A,&H307E,&H305B,&H3093,&H3002)
    UiMsgPurchaseSheetByKeyNotFoundSuffixText = cached
End Function

Public Function UiMsgPurchaseImportSheetNameFailedText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H8CFC,&H5165,&H5145,&H5F53,&H5358,&H4FA1,&H8868,&H306E,&H53D6,&H8FBC,&H5148,&H30B7,&H30FC,&H30C8,&H540D,&H3092,&H751F,&H6210,&H3067,&H304D,&H307E,&H305B,&H3093,&H3067,&H3057,&H305F,&H3002)
    UiMsgPurchaseImportSheetNameFailedText = cached
End Function

Public Function UiMsgWeldingImportableSheetNotFoundText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H30EC,&H30FC,&H30EB,&H6EB6,&H63A5,&H5358,&H4FA1,&H8868,&H306B,&H53D6,&H308A,&H8FBC,&H307F,&H53EF,&H80FD,&H306A,&H30B7,&H30FC,&H30C8,&H304C,&H898B,&H3064,&H304B,&H308A,&H307E,&H305B,&H3093,&H3002)
    UiMsgWeldingImportableSheetNotFoundText = cached
End Function

Public Function UiMsgWeldingImportSheetNameFailedText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H30EC,&H30FC,&H30EB,&H6EB6,&H63A5,&H5358,&H4FA1,&H8868,&H306E,&H53D6,&H8FBC,&H5148,&H30B7,&H30FC,&H30C8,&H540D,&H3092,&H751F,&H6210,&H3067,&H304D,&H307E,&H305B,&H3093,&H3067,&H3057,&H305F,&H3002)
    UiMsgWeldingImportSheetNameFailedText = cached
End Function

Public Function UiMsgWeldingUnitPriceSettingMissingSuffixText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H306B,&H306F,&H6EB6,&H63A5,&H5358,&H4FA1,&H306E,&H8A2D,&H5B9A,&H304C,&H3042,&H308A,&H307E,&H305B,&H3093,&H3002)
    UiMsgWeldingUnitPriceSettingMissingSuffixText = cached
End Function

Public Function UiMsgUnitPriceClearConfirmPromptText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H5358,&H4FA1,&H60C5,&H5831,&H3092,&H30AF,&H30EA,&H30A2,&H3057,&H307E,&H3059,&H304B,&HFF1F)
    UiMsgUnitPriceClearConfirmPromptText = cached
End Function

Public Function UiMsgUnitPriceClearConfirmYesLineText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H306F,&H3044,&HFF1A,&H0043,&H0032,&H0034,&H306E,&H9078,&H629E,&H5185,&H5BB9,&H3068,&H3001,&H4F5C,&H6210,&H6E08,&H307F,&H306E,&H5358,&H4FA1,&H30B7,&H30FC,&H30C8,&H3092,&H524A,&H9664,&H3057,&H307E,&H3059,&H3002)
    UiMsgUnitPriceClearConfirmYesLineText = cached
End Function

Public Function UiMsgUnitPriceClearConfirmNoLineText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H3044,&H3044,&H3048,&HFF1A,&H5358,&H4FA1,&H60C5,&H5831,&H3092,&H6B8B,&H3057,&H307E,&H3059,&H3002)
    UiMsgUnitPriceClearConfirmNoLineText = cached
End Function

Public Function UiMsgUnitPriceClearConfirmTitleText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H5358,&H4FA1,&H60C5,&H5831,&H30AF,&H30EA,&H30A2)
    UiMsgUnitPriceClearConfirmTitleText = cached
End Function

Public Function UiMsgImportCompleteCountSuffixText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H4EF6,&H306E,&H7A4D,&H7B97,&H7DDA,&H533A,&H5358,&H4FA1,&H8868,&H3092,&H53D6,&H308A,&H8FBC,&H307F,&H307E,&H3057,&H305F,&H3002)
    UiMsgImportCompleteCountSuffixText = cached
End Function

Public Function UiMsgPurchaseSheetCreatedPrefixText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H8CFC,&H5165,&H5145,&H5F53,&H5358,&H4FA1,&H8868,&H3092,&H300C)
    UiMsgPurchaseSheetCreatedPrefixText = cached
End Function

Public Function UiMsgWeldingSheetCreatedPrefixText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H30EC,&H30FC,&H30EB,&H6EB6,&H63A5,&H5358,&H4FA1,&H8868,&H3092,&H300C)
    UiMsgWeldingSheetCreatedPrefixText = cached
End Function

Public Function UiMsgSheetCreatedSuffixText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H300D,&H30B7,&H30FC,&H30C8,&H306B,&H4F5C,&H6210,&H3057,&H307E,&H3057,&H305F,&H3002)
    UiMsgSheetCreatedSuffixText = cached
End Function

Public Function UiMsgManagerListFileUnreadableText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H51FA,&H5F35,&H6240,&H9577,&H30EA,&H30B9,&H30C8,&H30D5,&H30A1,&H30A4,&H30EB,&H3092,&H53C2,&H7167,&H3067,&H304D,&H307E,&H305B,&H3093,&H3067,&H3057,&H305F,&H3002)
    UiMsgManagerListFileUnreadableText = cached
End Function

Public Function UiMsgManagerBranchOfficeNotFoundText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H8A72,&H5F53,&H3059,&H308B,&H652F,&H5E97,&H540D,&H30FB,&H51FA,&H5F35,&H6240,&H540D,&H304C,&H898B,&H3064,&H304B,&H308A,&H307E,&H305B,&H3093,&H3067,&H3057,&H305F,&H3002)
    UiMsgManagerBranchOfficeNotFoundText = cached
End Function

Public Function UiMsgManagerListFolderNotFoundText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H51FA,&H5F35,&H6240,&H9577,&H30EA,&H30B9,&H30C8,&H30D5,&H30A9,&H30EB,&H30C0,&H304C,&H898B,&H3064,&H304B,&H308A,&H307E,&H305B,&H3093,&H3002)
    UiMsgManagerListFolderNotFoundText = cached
End Function

Public Function UiMsgManagerListFileNotFoundSuffixText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H0020,&H5E74,&H306E,&H51FA,&H5F35,&H6240,&H9577,&H30EA,&H30B9,&H30C8,&H30D5,&H30A1,&H30A4,&H30EB,&H304C,&H898B,&H3064,&H304B,&H308A,&H307E,&H305B,&H3093,&H3002,&H30D5,&H30A1,&H30A4,&H30EB,&H540D,&H306B,&H5E74,&H5EA6,&H304C,&H542B,&H307E,&H308C,&H3066,&H3044,&H308B,&H304B,&H78BA,&H8A8D,&H3057,&H3066,&H304F,&H3060,&H3055,&H3044,&H3002)
    UiMsgManagerListFileNotFoundSuffixText = cached
End Function

Public Function UiMsgDateInputFailedText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H65E5,&H4ED8,&H3092,&H5165,&H529B,&H3067,&H304D,&H307E,&H305B,&H3093,&H3067,&H3057,&H305F,&H3002)
    UiMsgDateInputFailedText = cached
End Function

Public Function UiMsgDebugLogEmptyText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H30ED,&H30B0,&H306F,&H3042,&H308A,&H307E,&H305B,&H3093,&H3002)
    UiMsgDebugLogEmptyText = cached
End Function

Public Function UiMsgDebugLogFlushSuffixText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H0020,&H4EF6,&H306E,&H30ED,&H30B0,&H3092,&H30B7,&H30FC,&H30C8,&H306B,&H51FA,&H529B,&H3057,&H307E,&H3057,&H305F,&H3002)
    UiMsgDebugLogFlushSuffixText = cached
End Function

Public Function UiMsgProjectDataNotFoundText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H30C7,&H30FC,&H30BF,&H304C,&H898B,&H3064,&H304B,&H308A,&H307E,&H305B,&H3093,&H3002)
    UiMsgProjectDataNotFoundText = cached
End Function

Public Function UiMsgProjectStatusLoadingCaptionText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H5DE5,&H4E8B,&H73FE,&H6CC1,&H8868,&H30C7,&H30FC,&H30BF,&H3092,&H8AAD,&H8FBC,&H4E2D,&H002E,&H002E,&H002E)
    UiMsgProjectStatusLoadingCaptionText = cached
End Function

Public Function UiMsgProjectNumberSelectionCaptionText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H5DE5,&H4E8B,&H756A,&H53F7,&H9078,&H629E)
    UiMsgProjectNumberSelectionCaptionText = cached
End Function

Public Function UiMsgProjectStatusLoadFailedText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H5DE5,&H4E8B,&H73FE,&H6CC1,&H8868,&H30C7,&H30FC,&H30BF,&H306E,&H8AAD,&H307F,&H8FBC,&H307F,&H306B,&H5931,&H6557,&H3057,&H307E,&H3057,&H305F,&H3002)
    UiMsgProjectStatusLoadFailedText = cached
End Function

Public Function UiMsgProjectSelectionCaptionSuffixText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H0020,&H5DE5,&H4E8B,&H9078,&H629E,&HFF08,&H76F4,&H63A5,&H5165,&H529B,&H306E,&H5834,&H5408,&H306F,&H30BB,&H30EB,&H3078,&H5165,&H529B,&H3057,&H3066,&H304F,&H3060,&H3055,&H3044,&HFF09)
    UiMsgProjectSelectionCaptionSuffixText = cached
End Function

Public Function UiMsgProjectApplyToBasicInfoFailedText() As String
    Static cached As String
    If cached = "" Then cached = CommonTextFromChars(&H5DE5,&H4E8B,&H60C5,&H5831,&H3092,&H57FA,&H672C,&H60C5,&H5831,&H30B7,&H30FC,&H30C8,&H3078,&H53CD,&H6620,&H3067,&H304D,&H307E,&H305B,&H3093,&H3067,&H3057,&H305F,&H3002)
    UiMsgProjectApplyToBasicInfoFailedText = cached
End Function

