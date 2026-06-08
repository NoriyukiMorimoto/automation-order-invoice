# -*- coding: utf-8 -*-
import os

BASE = os.path.dirname(os.path.abspath(__file__))

REPLACEMENTS = [
    ('MsgBox "\u57fa\u672c\u60c5\u5831\u30b7\u30fc\u30c8\u304c\u898b\u3064\u304b\u308a\u307e\u305b\u3093\u3002\u30b7\u30fc\u30c8\u540d\u3092\u78ba\u8a8d\u3057\u3066\u304f\u3060\u3055\u3044\u3002"', 'MsgBox UiMsgBasicInfoSheetNotFoundCheckNameText()'),
    ('MsgBox "\u57fa\u672c\u60c5\u5831\u30b7\u30fc\u30c8\u304c\u898b\u3064\u304b\u308a\u307e\u305b\u3093\u3002"', 'MsgBox UiMsgBasicInfoSheetNotFoundText()'),
    ('MsgBox "\u5358\u4fa1\u8868\u30d6\u30c3\u30af\u306b\u53d6\u308a\u8fbc\u307f\u53ef\u80fd\u306a\u30b7\u30fc\u30c8\u304c\u898b\u3064\u304b\u308a\u307e\u305b\u3093\u3067\u3057\u305f\u3002"', 'MsgBox UiMsgUnitPriceImportableSheetNotFoundText()'),
    ('vbInformation, "\u5b8c\u4e86"', 'vbInformation, UiMsgImportCompleteTitleText()'),
    ('MsgBox "\u57fa\u672c\u60c5\u5831\u30b7\u30fc\u30c8 B4 \u306b4\u6841\u306e\u5e74\u5ea6\u304c\u898b\u3064\u304b\u308a\u307e\u305b\u3093\u3002\u4f8b: 2026"', 'MsgBox UiMsgBasicInfoYearNotFoundB4ExampleText()'),
    ('MsgBox "\u57fa\u672c\u60c5\u5831\u30b7\u30fc\u30c8 B4 \u306b4\u6841\u306e\u5e74\u5ea6\u304c\u898b\u3064\u304b\u308a\u307e\u305b\u3093\u3002"', 'MsgBox UiMsgBasicInfoYearNotFoundB4Text()'),
    ('MsgBox "\u57fa\u672c\u60c5\u5831\u30b7\u30fc\u30c8 B6 \u307e\u305f\u306f C6 \u304c\u7a7a\u3067\u3059\u3002\u652f\u5e97\u540d\u30fb\u51fa\u5f35\u6240\u540d\u3092\u78ba\u8a8d\u3057\u3066\u304f\u3060\u3055\u3044\u3002"', 'MsgBox UiMsgBasicInfoBranchOfficeEmptyText()'),
    ('MsgBox "\u57fa\u672c\u60c5\u5831\u30b7\u30fc\u30c8 C20 \u306e\u7dda\u533a\u533a\u5206\u3092\u9078\u629e\u3057\u3066\u304f\u3060\u3055\u3044\u3002"', 'MsgBox UiMsgBasicInfoLineTypeEmptyC20Text()'),
    ('MsgBox "\u57fa\u672c\u60c5\u5831\u30b7\u30fc\u30c8 C21 \u306e\u5358\u4fa1\u9069\u7528\u5de5\u4e8b\u4ef6\u540d\u3092\u9078\u629e\u3057\u3066\u304f\u3060\u3055\u3044\u3002"', 'MsgBox UiMsgBasicInfoProjectNameEmptyC21Text()'),
    ('MsgBox "\u57fa\u672c\u60c5\u5831\u30b7\u30fc\u30c8 C21 \u306f\u8ecc\u9053\u6750\u6599\u8cfc\u5165\u5145\u5f53\u4ee5\u5916\u306e\u5358\u4fa1\u9069\u7528\u5de5\u4e8b\u4ef6\u540d\u3092\u9078\u629e\u3057\u3066\u304f\u3060\u3055\u3044\u3002" & vbCrLf & _\n               "\u8cfc\u5165\u5145\u5f53\u5358\u4fa1\u306f C24 \u78ba\u5b9a\u5f8c\u306b\u81ea\u52d5\u4f5c\u6210\u3057\u307e\u3059\u3002"',
     'MsgBox UiMsgBasicInfoProjectNamePurchaseExcludedC21Text() & vbCrLf & _\n               UiMsgPurchaseUnitPriceAutoCreateAfterC24Text()'),
    ('MsgBox "\u57fa\u672c\u60c5\u5831\u30b7\u30fc\u30c8 C22 \u306e\u5358\u4fa1\u533a\u5206\u3092\u9078\u629e\u3057\u3066\u304f\u3060\u3055\u3044\u3002"', 'MsgBox UiMsgBasicInfoPriceKindEmptyC22Text()'),
    ('MsgBox "\u51fa\u5f35\u6240\u5225_\u5358\u4fa1\u9069\u7528\u7dda\u533a.xlsx \u304c\u898b\u3064\u304b\u308a\u307e\u305b\u3093\u3002"', 'MsgBox UiMsgUnitPriceMasterFileNotFoundText()'),
    ('MsgBox "\u51fa\u5f35\u6240\u5225_\u5358\u4fa1\u9069\u7528\u7dda\u533a.xlsx \u3092\u53c2\u7167\u3067\u304d\u307e\u305b\u3093\u3067\u3057\u305f\u3002"', 'MsgBox UiMsgUnitPriceMasterFileUnreadableText()'),
    ('MsgBox "\u5358\u4fa1\u9069\u7528\u7dda\u533a\u30b7\u30fc\u30c8\u304c\u898b\u3064\u304b\u308a\u307e\u305b\u3093\u3067\u3057\u305f\u3002" & vbCrLf & _\n               "\u30d6\u30c3\u30af\u5185\u306e\u30b7\u30fc\u30c8\u540d\u3092\u78ba\u8a8d\u3057\u3066\u304f\u3060\u3055\u3044\u3002"',
     'MsgBox UiMsgUnitPriceMasterSheetNotFoundText() & vbCrLf & _\n               UiMsgUnitPriceMasterSheetNameCheckText()'),
    ('MsgBox "\u5358\u4fa1\u9069\u7528\u7dda\u533a\u30b7\u30fc\u30c8\u306b\u652f\u5e97\u30fb\u51fa\u5f35\u6240\u306f\u898b\u3064\u304b\u308a\u307e\u3057\u305f\u304c\u3001\u7dda\u533a\u533a\u5206\u306b\u4e00\u81f4\u3059\u308b\u5358\u4fa1\u9069\u7528\u4fdd\u7dda\u533a\u3092\u7279\u5b9a\u3067\u304d\u307e\u305b\u3093\u3067\u3057\u305f\u3002" & vbCrLf & _\n               "\u652f\u5e97\uff1a"',
     'MsgBox UiMsgUnitPriceMasterLineTypeAmbiguousText() & vbCrLf & _\n               UiMsgBranchLabelText()'),
    ('MsgBox "\u5358\u4fa1\u9069\u7528\u7dda\u533a\u30b7\u30fc\u30c8\u306b\u8a72\u5f53\u3059\u308b\u652f\u5e97\u30fb\u51fa\u5f35\u6240\u304c\u898b\u3064\u304b\u308a\u307e\u305b\u3093\u3067\u3057\u305f\u3002" & vbCrLf & _\n               "\u652f\u5e97\uff1a"',
     'MsgBox UiMsgUnitPriceMasterBranchOfficeNotFoundText() & vbCrLf & _\n               UiMsgBranchLabelText()'),
    ('"\u652f\u5e97\uff1a"', 'UiMsgBranchLabelText()'),
    ('"\u51fa\u5f35\u6240\uff1a"', 'UiMsgOfficeLabelText()'),
    ('"\u7dda\u533a\u533a\u5206\uff1a"', 'UiMsgLineTypeLabelText()'),
    ('"\u5de5\u4e8b\u4ef6\u540d\uff1a"', 'UiMsgProjectNameLabelText()'),
    ('MsgBox "\u5358\u4fa1\u9069\u7528\u7dda\u533a\u30c7\u30fc\u30bf\u306e\u8aad\u307f\u8fbc\u307f\u306b\u5931\u6557\u3057\u307e\u3057\u305f\u3002"', 'MsgBox UiMsgUnitPriceMasterLoadFailedText()'),
    ('MsgBox "\u5de5\u4e8b\u4ef6\u540d\u306b\u4e00\u81f4\u3059\u308b\u5358\u4fa1\u8868\u304c\u898b\u3064\u304b\u308a\u307e\u305b\u3093\u3002" & vbCrLf & _\n               "\u5de5\u4e8b\u4ef6\u540d\uff1a"',
     'MsgBox UiMsgUnitPriceBookByProjectNotFoundText() & vbCrLf & _\n               UiMsgProjectNameLabelText()'),
    ('MsgBox "\u5358\u4fa1\u30c7\u30fc\u30bf\u30d5\u30a9\u30eb\u30c0\u304c\u898b\u3064\u304b\u308a\u307e\u305b\u3093\u3002"', 'MsgBox UiMsgUnitPriceDataFolderNotFoundText()'),
    ('MsgBox "\u7dda\u533a\u533a\u5206\u30d5\u30a9\u30eb\u30c0\u304c\u898b\u3064\u304b\u308a\u307e\u305b\u3093\u3002"', 'MsgBox UiMsgLineTypeFolderNotFoundText()'),
    ('MsgBox "\u652f\u793e\u30d5\u30a9\u30eb\u30c0\u304c\u898b\u3064\u304b\u308a\u307e\u305b\u3093\u3002"', 'MsgBox UiMsgBranchGroupFolderNotFoundText()'),
    ('MsgBox "\u5358\u4fa1\u9069\u7528\u4fdd\u7dda\u533a\u30d5\u30a9\u30eb\u30c0\u304c\u898b\u3064\u304b\u308a\u307e\u305b\u3093\u3002"', 'MsgBox UiMsgUnitPriceSectionFolderNotFoundText()'),
    ('MsgBox "\u5e74\u5ea6\u30d5\u30a9\u30eb\u30c0\u304c\u898b\u3064\u304b\u308a\u307e\u305b\u3093\u3002"', 'MsgBox UiMsgYearFolderNotFoundText()'),
    ('MsgBox "\u5358\u4fa1\u533a\u5206\u30d5\u30a9\u30eb\u30c0\u304c\u898b\u3064\u304b\u308a\u307e\u305b\u3093\u3002"', 'MsgBox UiMsgPriceKindFolderNotFoundText()'),
    ('MsgBox "\u5358\u4fa1\u8868\u30d6\u30c3\u30af\u3092\u958b\u3051\u307e\u305b\u3093\u3067\u3057\u305f\u3002"', 'MsgBox UiMsgUnitPriceBookOpenFailedText()'),
    ('MsgBox "\u7a4d\u7b97\u7dda\u533a\u9078\u629e\u30d5\u30a9\u30fc\u30e0\u3092\u8868\u793a\u3067\u304d\u307e\u305b\u3093\u3067\u3057\u305f\u3002"', 'MsgBox UiMsgLineNameFormShowFailedText()'),
    ('MsgBox "\u5358\u4fa1\u8868\u306e\u53d6\u308a\u8fbc\u307f\u306b\u5931\u6557\u3057\u307e\u3057\u305f\u3002"', 'MsgBox UiMsgUnitPriceImportFailedText()'),
    ('MsgBox "\u8cfc\u5165\u5145\u5f53\u5358\u4fa1\u8868\u306e\u53d6\u308a\u8fbc\u307f\u306b\u5931\u6557\u3057\u307e\u3057\u305f\u3002"', 'MsgBox UiMsgPurchaseUnitPriceImportFailedText()'),
    ('MsgBox "\u30ec\u30fc\u30eb\u6eb6\u63a5\u5358\u4fa1\u8868\u306e\u53d6\u308a\u8fbc\u307f\u306b\u5931\u6557\u3057\u307e\u3057\u305f\u3002"', 'MsgBox UiMsgWeldingUnitPriceImportFailedText()'),
    ('MsgBox "\u8cfc\u5165\u5145\u5f53\u5358\u4fa1\u8868\u306e\u53c2\u7167\u30ad\u30fc\u3092\u53d6\u5f97\u3067\u304d\u307e\u305b\u3093\u3067\u3057\u305f\u3002"', 'MsgBox UiMsgPurchaseReferenceKeyNotFoundText()'),
    ('MsgBox "\u8cfc\u5165\u5145\u5f53\u5358\u4fa1\u8868\u306b\u53c2\u7167\u30ad\u30fc\u300c" & referenceKey & "\u300d\u306b\u4e00\u81f4\u3059\u308b\u30b7\u30fc\u30c8\u304c\u898b\u3064\u304b\u308a\u307e\u305b\u3093\u3002"',
     'MsgBox UiMsgPurchaseSheetByKeyNotFoundPrefixText() & referenceKey & UiMsgPurchaseSheetByKeyNotFoundSuffixText()'),
    ('MsgBox "\u8cfc\u5165\u5145\u5f53\u5358\u4fa1\u8868\u306e\u53d6\u8fbc\u5148\u30b7\u30fc\u30c8\u540d\u3092\u751f\u6210\u3067\u304d\u307e\u305b\u3093\u3067\u3057\u305f\u3002"', 'MsgBox UiMsgPurchaseImportSheetNameFailedText()'),
    ('MsgBox "\u30ec\u30fc\u30eb\u6eb6\u63a5\u5358\u4fa1\u8868\u306b\u53d6\u308a\u8fbc\u307f\u53ef\u80fd\u306a\u30b7\u30fc\u30c8\u304c\u898b\u3064\u304b\u308a\u307e\u305b\u3093\u3002"', 'MsgBox UiMsgWeldingImportableSheetNotFoundText()'),
    ('MsgBox "\u30ec\u30fc\u30eb\u6eb6\u63a5\u5358\u4fa1\u8868\u306e\u53d6\u8fbc\u5148\u30b7\u30fc\u30c8\u540d\u3092\u751f\u6210\u3067\u304d\u307e\u305b\u3093\u3067\u3057\u305f\u3002"', 'MsgBox UiMsgWeldingImportSheetNameFailedText()'),
    ('BuildMissingWeldingUnitPriceMessage = sectionName & "\u306b\u306f\u6eb6\u63a5\u5358\u4fa1\u306e\u8a2d\u5b9a\u304c\u3042\u308a\u307e\u305b\u3093\u3002"',
     'BuildMissingWeldingUnitPriceMessage = sectionName & UiMsgWeldingUnitPriceSettingMissingSuffixText()'),
    ('If MsgBox("\u5358\u4fa1\u60c5\u5831\u3092\u30af\u30ea\u30a2\u3057\u307e\u3059\u304b\uff1f" & vbCrLf & _\n              "\u306f\u3044\uff1aC24\u306e\u9078\u629e\u5185\u5bb9\u3068\u3001\u4f5c\u6210\u6e08\u307f\u306e\u5358\u4fa1\u30b7\u30fc\u30c8\u3092\u524a\u9664\u3057\u307e\u3059\u3002" & vbCrLf & _\n              "\u3044\u3044\u3048\uff1a\u5358\u4fa1\u60c5\u5831\u3092\u6b8b\u3057\u307e\u3059\u3002", vbQuestion + vbYesNo, "\u5358\u4fa1\u60c5\u5831\u30af\u30ea\u30a2")',
     'If MsgBox(UiMsgUnitPriceClearConfirmPromptText() & vbCrLf & _\n              UiMsgUnitPriceClearConfirmYesLineText() & vbCrLf & _\n              UiMsgUnitPriceClearConfirmNoLineText(), vbQuestion + vbYesNo, UiMsgUnitPriceClearConfirmTitleText()'),
    ('BuildImportCompleteMessage = CStr(selectedSheetNames.Count) & "\u4ef6\u306e\u7a4d\u7b97\u7dda\u533a\u5358\u4fa1\u8868\u3092\u53d6\u308a\u8fbc\u307f\u307e\u3057\u305f\u3002"',
     'BuildImportCompleteMessage = CStr(selectedSheetNames.Count) & UiMsgImportCompleteCountSuffixText()'),
    ('BuildImportCompleteMessage = BuildImportCompleteMessage & vbCrLf & "\u8cfc\u5165\u5145\u5f53\u5358\u4fa1\u8868\u3092\u300c" & purchaseSheetName & "\u300d\u30b7\u30fc\u30c8\u306b\u4f5c\u6210\u3057\u307e\u3057\u305f\u3002"',
     'BuildImportCompleteMessage = BuildImportCompleteMessage & vbCrLf & UiMsgPurchaseSheetCreatedPrefixText() & purchaseSheetName & UiMsgSheetCreatedSuffixText()'),
    ('BuildImportCompleteMessage = BuildImportCompleteMessage & vbCrLf & "\u30ec\u30fc\u30eb\u6eb6\u63a5\u5358\u4fa1\u8868\u3092\u300c" & weldingSheetName & "\u300d\u30b7\u30fc\u30c8\u306b\u4f5c\u6210\u3057\u307e\u3057\u305f\u3002"',
     'BuildImportCompleteMessage = BuildImportCompleteMessage & vbCrLf & UiMsgWeldingSheetCreatedPrefixText() & weldingSheetName & UiMsgSheetCreatedSuffixText()'),
    ('MsgBox "\u51fa\u5f35\u6240\u9577\u30ea\u30b9\u30c8\u30d5\u30a1\u30a4\u30eb\u3092\u53c2\u7167\u3067\u304d\u307e\u305b\u3093\u3067\u3057\u305f\u3002"', 'MsgBox UiMsgManagerListFileUnreadableText()'),
    ('MsgBox "\u8a72\u5f53\u3059\u308b\u652f\u5e97\u540d\u30fb\u51fa\u5f35\u6240\u540d\u304c\u898b\u3064\u304b\u308a\u307e\u305b\u3093\u3067\u3057\u305f\u3002" & vbCrLf & _\n               "\u652f\u5e97\u540d\uff1a" & BranchName & vbCrLf & _\n               "\u51fa\u5f35\u6240\u540d\uff1a" & OfficeName, vbExclamation',
     'MsgBox UiMsgManagerBranchOfficeNotFoundText() & vbCrLf & _\n               UiMsgBranchNameLabelText() & BranchName & vbCrLf & _\n               UiMsgOfficeNameLabelText() & OfficeName, vbExclamation'),
    ('MsgBox "\u51fa\u5f35\u6240\u9577\u30ea\u30b9\u30c8\u30d5\u30a9\u30eb\u30c0\u304c\u898b\u3064\u304b\u308a\u307e\u305b\u3093\u3002"', 'MsgBox UiMsgManagerListFolderNotFoundText()'),
    ('MsgBox yearText & " \u5e74\u306e\u51fa\u5f35\u6240\u9577\u30ea\u30b9\u30c8\u30d5\u30a1\u30a4\u30eb\u304c\u898b\u3064\u304b\u308a\u307e\u305b\u3093\u3002\u30d5\u30a1\u30a4\u30eb\u540d\u306b\u5e74\u5ea6\u304c\u542b\u307e\u308c\u3066\u3044\u308b\u304b\u78ba\u8a8d\u3057\u3066\u304f\u3060\u3055\u3044\u3002"',
     'MsgBox yearText & UiMsgManagerListFileNotFoundSuffixText()'),
    ('MsgBox "\u65e5\u4ed8\u3092\u5165\u529b\u3067\u304d\u307e\u305b\u3093\u3067\u3057\u305f\u3002"', 'MsgBox UiMsgDateInputFailedText()'),
    ('MsgBox "\u30ed\u30b0\u306f\u3042\u308a\u307e\u305b\u3093\u3002"', 'MsgBox UiMsgDebugLogEmptyText()'),
    ('MsgBox mLogCount & " \u4ef6\u306e\u30ed\u30b0\u3092\u30b7\u30fc\u30c8\u306b\u51fa\u529b\u3057\u307e\u3057\u305f\u3002"',
     'MsgBox mLogCount & UiMsgDebugLogFlushSuffixText()'),
    ('MsgBox "\u30c7\u30fc\u30bf\u304c\u898b\u3064\u304b\u308a\u307e\u305b\u3093\u3002"', 'MsgBox UiMsgProjectDataNotFoundText()'),
    ('Me.Caption = "\u5de5\u4e8b\u73fe\u6cc1\u8868\u30c7\u30fc\u30bf\u3092\u8aad\u8fbc\u4e2d..."', 'Me.Caption = UiMsgProjectStatusLoadingCaptionText()'),
    ('Me.Caption = "\u5de5\u4e8b\u756a\u53f7\u9078\u629e"', 'Me.Caption = UiMsgProjectNumberSelectionCaptionText()'),
    ('MsgBox "\u5de5\u4e8b\u73fe\u6cc1\u8868\u30c7\u30fc\u30bf\u306e\u8aad\u307f\u8fbc\u307f\u306b\u5931\u6557\u3057\u307e\u3057\u305f\u3002"', 'MsgBox UiMsgProjectStatusLoadFailedText()'),
    ('Me.Caption = targetBranch & " " & targetOffice & " \u5de5\u4e8b\u9078\u629e\uff08\u76f4\u63a5\u5165\u529b\u306e\u5834\u5408\u306f\u30bb\u30eb\u3078\u5165\u529b\u3057\u3066\u304f\u3060\u3055\u3044\uff09"',
     'Me.Caption = targetBranch & " " & targetOffice & UiMsgProjectSelectionCaptionSuffixText()'),
    ('MsgBox "\u5de5\u4e8b\u60c5\u5831\u3092\u57fa\u672c\u60c5\u5831\u30b7\u30fc\u30c8\u3078\u53cd\u6620\u3067\u304d\u307e\u305b\u3093\u3067\u3057\u305f\u3002"', 'MsgBox UiMsgProjectApplyToBasicInfoFailedText()'),
]

TARGETS = [
    'mod_MaterialPriceImport.bas',
    'mod_FillManagerName.bas',
    'mod_BasicInfoCalendar.bas',
    'mod_DebugLog.bas',
    'Project_Number_Selection.frm',
]

def read_cp932(path):
    with open(path, 'rb') as f:
        return f.read().decode('cp932')


def write_cp932(path, text):
    with open(path, 'wb') as f:
        f.write(text.encode('cp932'))


def normalize_pattern(pattern):
    return pattern.replace('\n', '\r\n')


def apply_replacements(dry_run=False):
    ordered_replacements = sorted(
        REPLACEMENTS,
        key=lambda pair: len(normalize_pattern(pair[0])),
        reverse=True,
    )
    for fn in TARGETS:
        path = os.path.join(BASE, fn)
        text = read_cp932(path)
        count = 0
        for old, new in ordered_replacements:
            old_norm = normalize_pattern(old)
            new_norm = normalize_pattern(new)
            if old_norm in text:
                text = text.replace(old_norm, new_norm)
                count += 1
        if dry_run:
            print(fn, count, 'replacements (dry-run)')
        else:
            write_cp932(path, text)
            print(fn, count, 'replacements')


if __name__ == '__main__':
    import sys
    apply_replacements(dry_run='--dry-run' in sys.argv)
