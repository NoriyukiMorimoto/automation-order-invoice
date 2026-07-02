# -*- coding: utf-8 -*-
"""Split oversized VBA modules while preserving mod_* facade wrappers."""
from __future__ import annotations

import re
import shutil
from pathlib import Path

ENC = "cp932"
ROOT = Path(__file__).resolve().parent

# ---------------------------------------------------------------------------
# Phase 1: mod_VendorMaster
# ---------------------------------------------------------------------------

VENDOR_BLOCK_LAYOUT = {
    "InitVendorBlockCountFromSheet",
    "SyncVendorBlocksFromCount",
    "EnsureVendorCountInputValidation",
    "VendorCountValidationTitleText",
    "VendorCountValidationInputText",
    "VendorCountValidationErrorText",
    "CleanupLegacyVendorListDebrisInColumnAD",
    "CountExistingVendorBlocks",
    "GetVendorBlockCount",
    "VendorNameCellByIndex",
    "VendorLabelColumnByIndex",
    "VendorValueColumnByIndex",
    "VendorSpacerColumnByIndex",
    "ApplyVendorRow10ValueCellFormat",
    "VendorBlockNeedsPresentationRestore",
    "EnsureVendorBlockFromTemplate",
    "CopyVendorBlockFromTemplate",
    "CopyVendorBlockMergeAreasFromTemplate",
    "VendorBlockRangeByIndex",
    "RestoreVendorBlockPresentationFromTemplate",
    "RestoreVendorBlockLabelTextsFromTemplate",
    "CopyRangeFormats",
    "RestoreVendorBlockValueColumnRightBorders",
    "CopyCellBorderEdge",
    "CopyRangeBorders",
    "CopyVendorBlockFormatsFromTemplate",
    "CopyVendorBlockTotalRowsFromTemplate",
    "ClearVendorWorkTypeWhenCompanyEmpty",
    "ApplyVendorBlockColumnWidths",
    "ClearUnusedVendorBlocks",
    "ClearVendorBlockColumns",
    "VendorInfoHeaderText",
    "VendorInfoHeaderPrefixText",
    "SafeUnmergeRange",
    "IsSyncVendorBlocksInProgressImpl",
    "ResetVendorBlockSyncStateImpl",
}

VENDOR_UNIT_PRICE = {
    "GetVendorUnitPriceMonitorRange",
    "ApplyImportedUnitPriceJrHeadersForBasicInfo",
    "HandleVendorUnitPriceMonitorChange",
    "CollectMonitorChangedValueColumns",
    "ChangedRangeIncludesVendorWorkTypeRow",
    "AddUniqueLongToCollection",
    "IsLongInCollection",
    "CollectOutsourceRatioOnlyChangedValueColumns",
    "GetPreferredWeldingRatioColumnFromChange",
    "ApplyConstructionUnitPriceImportedRowDecorations",
    "ApplyConstructionUnitPriceImportedRowDecorationsFast",
    "RefreshConstructionUnitPriceSheetDataDecorations",
    "RefreshAllConstructionUnitPriceSheetDataDecorations",
    "HandleConstructionUnitPriceSheetChange",
    "RefreshAllVendorUnitPricesForBasicInfo",
    "SyncVendorUnitPriceBlocksAfterCountChange",
    "RefreshVendorUnitPriceSheetForSyncSafely",
    "ApplyVendorUnitPriceAddedBlocksToSheetSafely",
    "ClearVendorUnitPriceRemovedBlocksOnSheetSafely",
    "RefreshVendorUnitPriceForValueColumn",
    "RefreshVendorUnitPriceOutsourceRatioOnlyForValueColumn",
    "IsVendorUnitPriceBlockAlreadyBuilt",
    "RefreshVendorUnitPriceBlocksOnSheet",
    "ShouldApplyVendorUnitPriceBlock",
    "HasVendorName",
    "ApplyVendorUnitPriceBlockToSheet",
    "BuildVendorUnitPriceNameMap",
    "ResolveVendorUnitPriceName",
    "ClearVendorUnitPriceBlockOnSheet",
    "ApplyVendorUnitPriceJrHeader",
    "BuildVendorUnitPriceJrHeaderText",
    "ApplyVendorUnitPriceColumnWidths",
    "ApplyVendorUnitPriceOutsourceRatioRow",
    "ApplyVendorUnitPriceOutsourceRatioRowFont",
    "ClearVendorUnitPriceOutsourceRatioRow",
    "GetVendorOutsourceRatioPercentValue",
    "GetVendorOutsourceRatioNumericValue",
    "ApplyVendorUnitPriceMergedHeader",
    "ApplyVendorUnitPriceMergedVendorName",
    "ApplyVendorUnitPriceDataRows",
    "ApplyVendorUnitPriceDataColumn",
    "ReadVendorUnitPriceColumnValues",
    "RowNeedsFormulaFromArrays",
    "ApplyVendorUnitPriceGreyFillRange",
    "VendorUnitPriceRowNeedsFormulaForSource",
    "ApplyVendorUnitPriceFormulaSegment",
    "ApplyVendorUnitPriceBaseRowBorders",
    "ApplyVendorUnitPriceNewRowFill",
    "GetVendorUnitPriceInitialFillLastColumn",
    "ResetClearedVendorUnitPriceRows",
    "RefreshVendorUnitPriceBordersForSheet",
    "HandleVendorUnitPriceSourceChanges",
    "EnsureVendorUnitPriceNewRowFillForSourceRows",
    "ApplyVendorUnitPriceSourceRowsForRange",
    "ApplyVendorUnitPriceSourceRowsForRangeFast",
    "ApplyVendorUnitPriceSourceRowIfNeededFromValue",
    "IsNumericSourceValue",
    "IsBlankSourceValue",
    "ApplyVendorUnitPriceSourceEfDecorationsFast",
    "ApplyVendorUnitPriceSourceCellDecoration",
    "ApplyVendorUnitPriceSourceColumnsNumberFormatFast",
    "ApplyVendorUnitPriceSourceRowIfNeeded",
    "HasNumericVendorUnitPriceSource",
    "IsVendorUnitPriceSourceCellBlank",
    "ApplyVendorUnitPriceSourceGreyFill",
    "ApplyVendorUnitPriceSourceColumnsNumberFormat",
    "ApplyVendorUnitPriceCellsForSourceRow",
    "ApplyVendorUnitPriceCell",
    "ApplyVendorUnitPriceGreyFill",
    "ApplyVendorUnitPriceFont",
    "ApplyVendorUnitPriceBorders",
    "EnsureApplicationCalculationAutomatic",
    "IsVendorUnitPriceSourceBlank",
    "BuildVendorUnitPriceFormula",
    "BuildVendorUnitPriceFormulaR1C1",
    "BuildVendorUnitPriceHeaderText",
    "GetVendorOutsourceRatioAddress",
    "GetVendorUnitPriceLastDataRow",
    "VendorUnitPriceDayColumnByValueColumn",
    "GetVendorIndexFromValueColumn",
    "IsRailConstructionVendorBlock",
    "HasVendorOutsourceRatio",
    "VendorRailConstructionText",
    "VendorUnitPriceOutsourceLabelText",
    "VendorUnitPriceOutsourceRatioLabelText",
    "VendorJrUnitPriceLabelText",
    "VendorUnitPriceFontNameText",
    "VendorUnitPriceDayLabelText",
    "VendorUnitPriceNightLabelText",
    "VendorWasteDisposalKeywordText",
}

VENDOR_MASTER_PUBLIC = {
    "IsSyncVendorBlocksInProgress",
    "ResetVendorBlockSyncState",
    "RefreshVendorListForBasicInfo",
    "FillVendorInfoToBasicInfo",
    "ShowAllVendorSelection",
    "ApplyVendorSelection",
    "NotifyVendorBasicInfoBlockChanged",
    "GetVendorIndexFromValueColumnPublic",
    "ResolveVendorNameChangeCell",
    "GetAllVendorSelectionData",
    "ScheduleVendorListDropdown",
    "CancelScheduledVendorListDropdown",
    "HideVendorComboBox",
    "PromptVendorListDropdown",
    "CommitVendorComboBoxSelection",
    "GetVendorNameRange",
    "InitVendorBlockCountFromSheet",
    "SyncVendorBlocksFromCount",
    "GetVendorUnitPriceMonitorRange",
    "ApplyImportedUnitPriceJrHeadersForBasicInfo",
    "HandleVendorUnitPriceMonitorChange",
    "ApplyConstructionUnitPriceImportedRowDecorations",
    "RefreshConstructionUnitPriceSheetDataDecorations",
    "RefreshAllConstructionUnitPriceSheetDataDecorations",
    "HandleConstructionUnitPriceSheetChange",
    "RefreshAllVendorUnitPricesForBasicInfo",
    "BuildVendorUnitPriceNameMap",
    "EnsureVendorCountInputValidation",
    "CleanupLegacyVendorListDebrisInColumnAD",
    "ClearVendorRowsCache",
}

PROC_RE = re.compile(r"^(Public|Private)\s+(Sub|Function)\s+(\w+)", re.M)
CONST_RE = re.compile(r"^Private Const\s+(\w+)", re.M)
DIM_RE = re.compile(r"^Private\s+(m\w+)\s+As", re.M)


def read_text(path: Path) -> str:
    return path.read_bytes().decode(ENC)


def write_text(path: Path, text: str) -> None:
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = "\r\n".join(text.split("\n"))
    if not text.endswith("\r\n"):
        text += "\r\n"
    path.write_bytes(text.encode(ENC))


def parse_procedures(source: str) -> dict[str, str]:
    matches = list(PROC_RE.finditer(source))
    procs: dict[str, str] = {}
    for i, m in enumerate(matches):
        start = m.start()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(source)
        name = m.group(3)
        procs[name] = source[start:end].rstrip() + "\n"
    return procs


def extract_header(source: str) -> str:
    first_proc = PROC_RE.search(source)
    if not first_proc:
        return "Option Explicit\r\n\r\n"
    return source[: first_proc.start()].rstrip() + "\n\n"


def consts_in_text(header: str, body: str) -> list[str]:
    names = CONST_RE.findall(header)
    used = []
    for line in header.splitlines():
        m = CONST_RE.match(line.strip())
        if not m:
            continue
        name = m.group(1)
        if re.search(r"\b" + re.escape(name) + r"\b", body):
            used.append(line.rstrip())
    return used


def module_vars_for(module: str, header: str) -> list[str]:
    vars_needed: dict[str, set[str]] = {
        "master": {"mVendorPromptTime", "mVendorTargetAddress", "mVendorRowsCache", "mVendorNameIndexCache"},
        "block": {"mLastVendorBlockCount", "mSyncVendorBlocksInProgress"},
        "unitprice": set(),
    }
    lines = []
    for line in header.splitlines():
        m = DIM_RE.match(line.strip())
        if m and m.group(1) in vars_needed[module]:
            lines.append(line.rstrip())
    return lines


def rename_impl(procs: dict[str, str]) -> None:
    if "IsSyncVendorBlocksInProgress" in procs:
        procs["IsSyncVendorBlocksInProgressImpl"] = procs.pop("IsSyncVendorBlocksInProgress").replace(
            "IsSyncVendorBlocksInProgress()", "IsSyncVendorBlocksInProgressImpl()", 1
        )
    if "ResetVendorBlockSyncState" in procs:
        procs["ResetVendorBlockSyncStateImpl"] = procs.pop("ResetVendorBlockSyncState").replace(
            "ResetVendorBlockSyncState", "ResetVendorBlockSyncStateImpl", 1
        )


def patch_calls(text: str, module_name: str) -> str:
    block_helpers = {
        "GetVendorBlockCount",
        "CountExistingVendorBlocks",
        "VendorNameCellByIndex",
        "VendorLabelColumnByIndex",
        "VendorValueColumnByIndex",
        "VendorSpacerColumnByIndex",
        "ClearVendorWorkTypeWhenCompanyEmpty",
        "RestoreVendorBlockValueColumnRightBorders",
        "ApplyVendorBlockColumnWidths",
        "ApplyVendorRow10ValueCellFormat",
        "VendorInfoHeaderText",
        "VendorInfoHeaderPrefixText",
        "EnsureVendorBlockFromTemplate",
        "VendorBlockNeedsPresentationRestore",
        "ClearUnusedVendorBlocks",
        "SafeUnmergeRange",
        "InitVendorBlockCountFromSheet",
        "SyncVendorBlocksFromCount",
        "EnsureVendorCountInputValidation",
        "CleanupLegacyVendorListDebrisInColumnAD",
    }
    unit_helpers = {
        "GetVendorUnitPriceMonitorRange",
        "ApplyImportedUnitPriceJrHeadersForBasicInfo",
        "HandleVendorUnitPriceMonitorChange",
        "ApplyConstructionUnitPriceImportedRowDecorations",
        "RefreshConstructionUnitPriceSheetDataDecorations",
        "RefreshAllConstructionUnitPriceSheetDataDecorations",
        "HandleConstructionUnitPriceSheetChange",
        "RefreshAllVendorUnitPricesForBasicInfo",
        "SyncVendorUnitPriceBlocksAfterCountChange",
        "RefreshVendorUnitPriceForValueColumn",
        "RefreshVendorUnitPriceOutsourceRatioOnlyForValueColumn",
        "GetVendorIndexFromValueColumn",
        "GetVendorUnitPriceLastDataRow",
        "BuildVendorUnitPriceNameMap",
        "ApplyConstructionUnitPriceImportedRowDecorationsFast",
    }
    master_helpers = {
        "RefreshVendorListForBasicInfo",
        "ClearVendorInfoBlock",
        "GetVendorNameRange",
        "GetVendorTargetCell",
        "ResolveVendorNameChangeCell",
        "LoadVendorRows",
        "HasVendorRows",
        "GetVendorNameIndex",
        "LoadAllVendorRows",
        "VendorWritableValueCell",
    }

    def prefix(name: str, mod: str) -> None:
        nonlocal text
        if mod == module_name:
            return
        text = re.sub(r"\b" + re.escape(name) + r"\b", mod + "." + name, text)

    for n in sorted(block_helpers, key=len, reverse=True):
        prefix(n, "mod_VendorBlockLayout")
    for n in sorted(unit_helpers, key=len, reverse=True):
        prefix(n, "mod_VendorUnitPrice")
    for n in sorted(master_helpers, key=len, reverse=True):
        prefix(n, "mod_VendorMaster")

    for mod in ("mod_VendorMaster", "mod_VendorBlockLayout", "mod_VendorUnitPrice"):
        text = text.replace(mod + "." + mod + ".", mod + ".")
    return text


def make_public(text: str) -> str:
    return re.sub(r"^Private (Sub|Function)", r"Public \1", text, count=1, flags=re.M)


def build_module(name: str, header: str, proc_names: set[str], procs: dict[str, str], public_all: bool) -> str:
    bodies = []
    combined = ""
    for pname in sorted(proc_names, key=lambda n: procs[n][:80]):
        if pname not in procs:
            raise KeyError(f"{name}: missing procedure {pname}")
        body = procs[pname]
        if public_all:
            body = make_public(body)
        bodies.append(body)
        combined += body
    const_lines = consts_in_text(header, combined)
    if name == "mod_VendorBlockLayout":
        var_lines = module_vars_for("block", header)
    elif name == "mod_VendorUnitPrice":
        var_lines = module_vars_for("unitprice", header)
    else:
        var_lines = module_vars_for("master", header)

    out = "Option Explicit\n\n"
    if const_lines:
        out += "\n".join(const_lines) + "\n\n"
    if var_lines:
        out += "\n".join(var_lines) + "\n\n"
    out += "\n".join(bodies)
    return patch_calls(out, name)


def extract_signature(orig: str) -> str:
    lines = orig.splitlines()
    sig = lines[0].strip()
    idx = 0
    while ")" not in sig and idx + 1 < len(lines):
        idx += 1
        sig = sig + " " + lines[idx].strip()
    return sig


def parse_vba_param_names(params_str: str) -> str:
    parts: list[str] = []
    depth = 0
    buf: list[str] = []
    for ch in params_str:
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        elif ch == "," and depth == 0:
            parts.append("".join(buf).strip())
            buf = []
            continue
        buf.append(ch)
    tail = "".join(buf).strip()
    if tail:
        parts.append(tail)

    names: list[str] = []
    for part in parts:
        token = part.strip()
        token = re.sub(r"^Optional\s+", "", token)
        token = re.sub(r"^(ByVal|ByRef|ParamArray)\s+", "", token)
        token = re.sub(r"^_\s*", "", token)
        if "=" in token:
            token = token.split("=", 1)[0].strip()
        if " As " in token:
            token = token.split(" As ", 1)[0].strip()
        if token and token != "_":
            names.append(token)
    return ", ".join(names)


def build_facade_wrapper(public_name: str, target_mod: str, orig: str) -> str:
    sig = extract_signature(orig).replace("Private ", "Public ")
    sig = re.sub(r"\s*_\s*", " ", sig)
    sig_one_line = re.sub(r"\s+", " ", sig).strip()
    if "(" not in sig_one_line:
        raise ValueError(f"Cannot parse signature for {public_name}")
    params = sig_one_line.split("(", 1)[1].rsplit(")", 1)[0]
    call_args = parse_vba_param_names(params)

    if "Function" in sig_one_line:
        ret_match = re.search(r"\)\s+As\s+(\w+)", sig_one_line)
        ret_type = ret_match.group(1) if ret_match else ""
        decl = sig_one_line
        if ret_type in {"Range", "Object", "Worksheet"}:
            assign = f"Set {public_name} = {target_mod}.{public_name}({call_args})"
        else:
            assign = f"{public_name} = {target_mod}.{public_name}({call_args})"
        return f"{decl}\n    {assign}\nEnd Function\n"

    decl = sig_one_line
    if call_args:
        call_line = f"Call {target_mod}.{public_name}({call_args})"
    else:
        call_line = f"Call {target_mod}.{public_name}()"
    return f"{decl}\n    {call_line}\nEnd Sub\n"


def split_vendor_master() -> None:
    src_path = ROOT / "mod_VendorMaster.bas.pre_split"
    if not src_path.exists():
        src_path = ROOT / "mod_VendorMaster.bas"
        shutil.copy2(src_path, ROOT / "mod_VendorMaster.bas.pre_split")
        src_path = ROOT / "mod_VendorMaster.bas.pre_split"

    source = read_text(src_path)
    header = extract_header(source)
    procs = parse_procedures(source)
    rename_impl(procs)

    master_names = set(procs) - VENDOR_BLOCK_LAYOUT - VENDOR_UNIT_PRICE
    master_names -= {"IsSyncVendorBlocksInProgressImpl", "ResetVendorBlockSyncStateImpl"}

    block = build_module("mod_VendorBlockLayout", header, VENDOR_BLOCK_LAYOUT, procs, True)
    unit = build_module("mod_VendorUnitPrice", header, VENDOR_UNIT_PRICE, procs, True)

    # facade wrappers in master for moved public APIs
    wrappers = []
    moved_public = (VENDOR_BLOCK_LAYOUT | VENDOR_UNIT_PRICE) & VENDOR_MASTER_PUBLIC
    moved_public |= {"IsSyncVendorBlocksInProgress", "ResetVendorBlockSyncState"}
    for pname in sorted(moved_public):
        if pname == "IsSyncVendorBlocksInProgress":
            wrappers.append(
                "Public Function IsSyncVendorBlocksInProgress() As Boolean\n"
                "    IsSyncVendorBlocksInProgress = mod_VendorBlockLayout.IsSyncVendorBlocksInProgressImpl()\n"
                "End Function\n"
            )
            continue
        if pname == "ResetVendorBlockSyncState":
            wrappers.append(
                "Public Sub ResetVendorBlockSyncState()\n"
                "    mod_VendorBlockLayout.ResetVendorBlockSyncStateImpl\n"
                "End Sub\n"
            )
            continue
        target = "mod_VendorBlockLayout" if pname in VENDOR_BLOCK_LAYOUT else "mod_VendorUnitPrice"
        orig = procs.get(pname, "")
        if not orig:
            continue
        wrappers.append(build_facade_wrapper(pname, target, orig))

    for pname in moved_public:
        master_names.discard(pname)

    master = build_module("mod_VendorMaster", header, master_names, procs, False)
    # make key helpers public for cross-module
    for helper in [
        "ClearVendorInfoBlock",
        "GetVendorTargetCell",
        "VendorWritableValueCell",
        "LoadVendorRows",
        "HasVendorRows",
        "GetVendorNameIndex",
        "LoadAllVendorRows",
    ]:
        master = master.replace(f"Private Sub {helper}", f"Public Sub {helper}", 1)
        master = master.replace(f"Private Function {helper}", f"Public Function {helper}", 1)

    master = master.rstrip() + "\n\n" + "\n".join(wrappers)

    write_text(ROOT / "mod_VendorBlockLayout.bas", block)
    write_text(ROOT / "mod_VendorUnitPrice.bas", unit)
    write_text(ROOT / "mod_VendorMaster.bas", master)
    print("Phase 1 split complete.")


def split_vba_module(
    backup_file: str,
    output_modules: dict[str, set[str]],
    facade_file: str,
    facade_public: set[str],
    public_helpers: list[str] | None = None,
    patch_fn=None,
) -> None:
    src_path = ROOT / backup_file
    if not src_path.exists():
        live = ROOT / backup_file.replace(".pre_split", ".bas")
        shutil.copy2(live, src_path)

    source = read_text(src_path)
    header = extract_header(source)
    procs = parse_procedures(source)

    all_assigned: set[str] = set()
    for mod_name, names in output_modules.items():
        if mod_name == facade_file:
            continue
        overlap = all_assigned & names
        if overlap:
            raise ValueError(f"Duplicate assignment: {overlap}")
        all_assigned |= names
    extra = all_assigned - set(procs)
    if extra:
        raise ValueError(f"{backup_file}: unknown procedures: {sorted(extra)}")

    built: dict[str, str] = {}
    for mod_name, proc_names in output_modules.items():
        if mod_name == facade_file:
            continue
        body = build_module(mod_name, header, proc_names, procs, True)
        if patch_fn:
            body = patch_fn(body, mod_name)
        built[mod_name] = body

    others = set()
    for mod, names in output_modules.items():
        if mod != facade_file:
            others |= names
    facade_names = set(procs) - others

    moved_public = others & facade_public
    wrappers: list[str] = []
    for pname in sorted(moved_public):
        target_mod = next(m for m, n in output_modules.items() if m != facade_file and pname in n)
        orig = procs.get(pname, "")
        if orig:
            wrappers.append(build_facade_wrapper(pname, target_mod, orig))

    facade_names -= moved_public
    facade = build_module(facade_file, header, facade_names, procs, False)
    if patch_fn:
        facade = patch_fn(facade, facade_file)
    if public_helpers:
        for helper in public_helpers:
            facade = facade.replace(f"Private Sub {helper}", f"Public Sub {helper}", 1)
            facade = facade.replace(f"Private Function {helper}", f"Public Function {helper}", 1)
    facade = facade.rstrip() + "\n\n" + "\n".join(wrappers)

    for mod_name, body in built.items():
        if mod_name != facade_file:
            write_text(ROOT / f"{mod_name}.bas", body)
    write_text(ROOT / f"{facade_file}.bas", facade)


def patch_material_price_calls(text: str, module_name: str) -> str:
    main_helpers = {
        "LogUP", "LogUPB", "JoinCollectionText", "IsImportingUnitPriceData",
        "IsClearingImportedLineNames", "GetCurrentUnitPriceImportBatchId",
    }
    lookup_helpers = {
        "TryReadUnitPriceRequest", "TryLoadUnitPriceMasterRow", "ResolveUnitPriceSourceFilePaths",
        "ResolveUnitPricePriceFolderPath", "LoadWorksheetNameCandidatesFromWorkbooks",
        "LoadUnitPriceProjectNamesForBasicInfo", "GetMasterFilePath", "GetUnitPriceDataRootPath",
        "FindUnitPriceWorkbooks", "FindPurchaseUnitPriceWorkbook", "FindWeldingUnitPriceWorkbook",
        "NormalizeMatchText", "FirstExistingFilePath", "CollectionContainsText",
    }
    import_helpers = {
        "ImportSelectedUnitPriceSheets", "ImportAndMergePurchaseUnitPriceSheets",
        "ImportAndMergeWeldingUnitPriceSheets", "PromptLineNameSelection",
        "OrderUnitPriceSheetNamesByProjectMasterFColumn", "GenerateImportBatchId",
        "SetCurrentUnitPriceImportBatchId", "MarkImportedUnitPriceSheet",
        "IsCurrentImportBatchUnitPriceSheet", "DeleteImportedUnitPriceSheets",
        "ClearUnitPriceSheets", "MakeUniqueWorksheetName", "WriteSelectedLineNames",
        "IsImportedUnitPriceSheet", "IsConstructionUnitPriceSheet",
    }
    format_helpers = {
        "ApplyImportedUnitPriceSheetFormat", "FormatImportedLineNamesCell",
        "FillBlankUnitPriceEFCells", "FillBlankWeldingUnitPriceCells",
        "WriteUnitPriceProjectNameValidation", "ResetUnitPriceValidation",
    }

    def prefix(name: str, mod: str) -> None:
        nonlocal text
        if mod == module_name:
            return
        text = re.sub(r"\b" + re.escape(name) + r"\b", mod + "." + name, text)

    for n in sorted(format_helpers, key=len, reverse=True):
        prefix(n, "mod_UnitPriceSheetFormat")
    for n in sorted(import_helpers, key=len, reverse=True):
        prefix(n, "mod_UnitPriceSheetImport")
    for n in sorted(lookup_helpers, key=len, reverse=True):
        prefix(n, "mod_UnitPriceMasterLookup")
    for n in sorted(main_helpers, key=len, reverse=True):
        prefix(n, "mod_MaterialPriceImport")

    for mod in ("mod_MaterialPriceImport", "mod_UnitPriceMasterLookup", "mod_UnitPriceSheetImport", "mod_UnitPriceSheetFormat"):
        text = text.replace(mod + "." + mod + ".", mod + ".")
    return text


UP_MASTER_LOOKUP = {
    "TryReadUnitPriceRequest", "TryLoadUnitPriceMasterRow", "TryLoadUnitPriceMasterRowFromWorkbook",
    "FindWorksheetByName", "FindAdoWorksheetName", "BuildAdoSheetTableName",
    "FillUnitPriceMasterRowFromRecordset", "FillUnitPriceMasterRowFromWorksheet",
    "GetUnitPriceMasterRowScore", "UnitPriceMasterRowFolderExistsForLineType",
    "UnitPriceMasterRowTextMatchesLineType", "ResolveUnitPriceSourceFilePaths",
    "ResolveUnitPricePriceFolderPath", "ResolveUnitPriceFolderName",
    "NormalizeWorkbookPathForCompare", "FindOpenWorkbookByPath",
    "LoadWorksheetNamesFromWorkbookByAdo", "LoadWorksheetNamesFromWorkbookByExcel",
    "LoadWorksheetNamesFromWorkbook", "LoadWorksheetNameCandidatesFromWorkbooks",
    "LookupProjectNameByWorkName", "LookupLineTypeByWorkName", "GetMasterFilePath",
    "LoadUnitPriceProjectNamesForBasicInfo", "AddUnitPriceProjectNames", "LoadUnitPriceProjectNames",
    "ResolveUnitPriceProjectNameMasterConfig", "LoadUnitPriceProjectNamesByAdo",
    "LoadUnitPriceProjectNamesByAdo2", "BuildAdoSheetRangeName", "CollectionContainsText",
    "FindUnitPriceWorkbooks", "FindPurchaseUnitPriceWorkbook", "FindWeldingUnitPriceWorkbook",
    "FindWorkbookByKeyword", "RemoveFileExtension", "IsPurchaseUnitPriceProjectName",
    "FindChildFolderByKey", "FolderTextMatches", "MasterTextMatches", "NormalizeMatchText",
    "NormalizeFolderName", "GetUnitPriceDataRootPath", "FirstExistingFilePath",
    "OrderInvoiceDocumentFolderText", "MasterDataFolderText", "UnitPriceMasterFolderText",
    "BuildImportCompleteMessage",
}

UP_SHEET_IMPORT = {
    "BuildUniqueUnitPriceSheetDisplayName", "PromptLineNameSelection", "ImportSelectedUnitPriceSheets",
    "ImportAndMergePurchaseUnitPriceSheets", "ImportAndMergeWeldingUnitPriceSheets",
    "AppendSheetDataExcludingHeader", "FindLastUsedRow", "ImportPurchaseUnitPriceSheetsByReference",
    "ImportWeldingUnitPriceSheetsIfRequired", "BuildMissingWeldingUnitPriceMessage",
    "IsWeldingUnitPriceRequired", "LoadWeldingSheetNames", "LoadPurchaseSheetNamesByReference",
    "AddPurchaseSheetNamesByReferenceKey", "BuildPurchaseFallbackReferenceKey",
    "TrimPurchaseReferenceSuffix", "PurchaseSheetNameMatchesReferenceKey", "BuildPurchaseReferenceKey",
    "ExtractLeadingDigits", "GetPathBaseName", "BuildPurchaseSheetName", "BuildWeldingSheetName",
    "TrimLeadingDigitsAndSeparators", "OrderUnitPriceSheetNamesByProjectMasterFColumn",
    "WriteSelectedLineNames", "GenerateImportBatchId", "SetCurrentUnitPriceImportBatchId",
    "GetCurrentUnitPriceImportBatchId", "ClearCurrentUnitPriceImportBatchId",
    "IsCurrentImportBatchUnitPriceSheet", "MarkImportedUnitPriceSheet", "ClearUnitPriceSheets",
    "DeleteImportedUnitPriceSheets", "DeleteImportedUnitPriceSheetsExcept", "DeleteStagedWorksheets",
    "WorksheetIsInCollection", "DeleteWorksheetIfExists", "IsImportedUnitPriceSheet",
    "IsProtectedSystemWorksheet", "IsImportedUnitPriceSheetByNameSuffix", "TextEndsWith",
    "IsImportedUnitPriceSheetByMarker", "IsImportedUnitPriceSheetByTabColor",
    "MakeUniqueWorksheetName", "MakeSafeWorksheetName", "WorksheetExists",
}

UP_SHEET_FORMAT = {
    "ApplyImportedUnitPriceSheetFormat", "ApplyWeldingUnitPriceSheetSectionFormat",
    "FormatImportedLineNamesCell", "CalculateImportedLineNameFontSize",
    "NormalizeImportedLineNameText", "CountImportedLineNameLines", "FillBlankUnitPriceEFCells",
    "FillBlankWeldingUnitPriceCells", "IsWeldingYearHeaderRow", "WriteUnitPriceLineTypeValidation",
    "WriteUnitPriceKindValidation", "WriteUnitPriceProjectNameValidation", "ResetUnitPriceValidation",
    "ClearUnitPriceProjectNameValidation", "IsImportedLineNamesMonitorRangeChanged",
    "IsImportedLineNamesCellEmpty", "WeldingProjectNameLabelKeywordText",
    "WeldingLineSectionLabelKeywordText", "IsUnitPriceSheetDataSeiriRow",
    "ImportedUnitPriceSheetFontNameText", "IsConstructionUnitPriceSheet",
    "IsSpecialImportedUnitPriceSheet",
}

MATERIAL_PRICE_PUBLIC = {
    "GetMasterFilePath", "AutoFillLineTypeFromWorkName", "AutoFillUnitPriceFieldsFromWorkName",
    "AutoFillProjectNameFromWorkName", "RefreshUnitPriceProjectNameValidation",
    "ImportConstructionUnitPriceForBasicInfo", "ClearAndImportUnitPriceForBasicInfo",
    "IsImportingUnitPriceData", "ConfirmAndClearUnitPriceForBasicInfo",
    "SilentClearUnitPriceForBasicInfo", "HandleImportedLineNamesCellChange",
    "IsConstructionUnitPriceSheet", "IsCurrentImportBatchUnitPriceSheet",
    "FormatImportedLineNamesCell", "GetCurrentUnitPriceImportBatchId",
}


def split_material_price_import() -> None:
    backup = "mod_MaterialPriceImport.bas.pre_split"
    if not (ROOT / backup).exists():
        shutil.copy2(ROOT / "mod_MaterialPriceImport.bas", ROOT / backup)
    split_vba_module(
        backup,
        {
            "mod_UnitPriceMasterLookup": UP_MASTER_LOOKUP,
            "mod_UnitPriceSheetImport": UP_SHEET_IMPORT,
            "mod_UnitPriceSheetFormat": UP_SHEET_FORMAT,
        },
        "mod_MaterialPriceImport",
        MATERIAL_PRICE_PUBLIC | UP_MASTER_LOOKUP | UP_SHEET_IMPORT | UP_SHEET_FORMAT,
        public_helpers=["LogUP", "LogUPB", "JoinCollectionText", "ImportUnitPriceData"],
        patch_fn=patch_material_price_calls,
    )
    print("Phase 2 split complete.")


WELDING_TEMOTO = {
    "ResolveTemotoMasterFilePath", "FindAdoSheetNameWUP", "TemotoMasterSheetNameText",
    "TemotoMasterFilePatternText", "MasterDataFolderText", "OrderInvoiceDocumentFolderTextWUP",
}


def patch_welding_calls(text: str, module_name: str) -> str:
    temoto_helpers = {
        "ResolveTemotoMasterFilePath", "FindAdoSheetNameWUP", "TemotoMasterSheetNameText",
        "TemotoMasterFilePatternText", "MasterDataFolderText", "OrderInvoiceDocumentFolderTextWUP",
    }
    for n in sorted(temoto_helpers, key=len, reverse=True):
        if module_name != "mod_WeldingTemotoMaster":
            text = re.sub(r"\b" + re.escape(n) + r"\b", "mod_WeldingTemotoMaster." + n, text)
    text = text.replace("mod_WeldingTemotoMaster.mod_WeldingTemotoMaster.", "mod_WeldingTemotoMaster.")
    return text


def split_welding_unit_price() -> None:
    backup = ROOT / "mod_WeldingUnitPrice.bas.pre_split"
    if not backup.exists():
        shutil.copy2(ROOT / "mod_WeldingUnitPrice.bas", backup)
    split_vba_module(
        "mod_WeldingUnitPrice.bas.pre_split",
        {"mod_WeldingTemotoMaster": WELDING_TEMOTO},
        "mod_WeldingUnitPrice",
        set(),
        patch_fn=patch_welding_calls,
    )
    print("Phase 3 split complete.")


BASIC_GUIDE_TEXTS = {
    "GetC9CommentText", "GetC22CommentText", "GetC23CommentText", "GetC24CommentText",
    "GetC13CommentText", "GetF9CommentText", "GetF11CommentText", "GetF27CommentText",
    "GetF29CommentText", "GetF29KidoCommentText", "GetF31CommentText", "GetF31YosetsuCommentText",
    "GetF31KidoTemotoCommentText", "GetOutsourceRatioLabelText", "GetPatternOutsourceRatioText",
    "GetF30RailPatternCommentText", "GetKidoKojiText", "GetYosetsuKojiText",
    "GetDefaultRow30LabelText", "GetWeldingRow30LabelText", "GetRailPatternRow30LabelText",
    "VendorRowLabelPrefixText",
}


def patch_guide_calls(text: str, module_name: str) -> str:
    for n in sorted(BASIC_GUIDE_TEXTS, key=len, reverse=True):
        if module_name != "mod_BasicInfoGuideTexts":
            text = re.sub(r"\b" + re.escape(n) + r"\b", "mod_BasicInfoGuideTexts." + n, text)
    text = text.replace("mod_BasicInfoGuideTexts.mod_BasicInfoGuideTexts.", "mod_BasicInfoGuideTexts.")
    return text


def split_basic_info_guide() -> None:
    backup = "mod_BasicInfoGuide.bas.pre_split"
    if not (ROOT / backup).exists():
        shutil.copy2(ROOT / "mod_BasicInfoGuide.bas", ROOT / backup)
    split_vba_module(
        backup,
        {"mod_BasicInfoGuideTexts": BASIC_GUIDE_TEXTS},
        "mod_BasicInfoGuide",
        set(),
        patch_fn=patch_guide_calls,
    )
    print("Phase 4 split complete.")


if __name__ == "__main__":
    import sys
    phase = sys.argv[1] if len(sys.argv) > 1 else "all"
    if phase in ("1", "all"):
        split_vendor_master()
    if phase in ("2", "all"):
        split_material_price_import()
    if phase in ("3", "all"):
        split_welding_unit_price()
    if phase in ("4", "all"):
        split_basic_info_guide()
