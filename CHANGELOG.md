# CHANGELOG

改修履歴は�? `.bas` ファイルのヘッダーコメントから本ファイルへ移管しました�?
�?モジュールのヘッダーには `改修履歴: CHANGELOG.md 参�?�` と記載して�?ます�?

---

## mod_Construction_Order_Import.bas?��施工�?示書・施工通知書取込モジュール?�?

### #1
- 工事件名別マスタの F列（積算線区?��と G列（施工�?示書記載線区名）を読み込み�?
  施工�?示書側の契�?線区名から括弧書きを除去して単価シート名へ変換する処�?を追�?�?
- 変換後�?�単価シートで整�?番号を�?�合し、昼は E列、夜�?� F列�?�単価を取込シー�? O列へ転記�?
- P列へ `O列×F列` の金額式、Q列へ単価一致?��不一致の比�?結果と�?定色・罫線�?�中央�?えを設定�?

### #2
- 工事件名別マスタの探索を�?数ルート対応にし、C21 とファイル名�?�空白・中黒�?�表記差を吸収�?
- C20/C21、探索フォルダ、解決したマスタパスを施工�?示書取込ログへ追�?�?

### #3
- Q列が「単価不一致」�?�行につ�?て、R列へ対応する単価表シート名を含む独自工種入力案�??を表示�?
- 施工�?示書等�?�出力表の罫線�?�列�?調整�?囲をR列まで拡張�?

### #4
- 購入�?当通知シートでは、施工業�?(A�?)・管�?室(H�?)・外注単価(K�?)・外注金�?(L�?)�?
  適用対象外として非表示にするよう変更�?

### #5
- 施工�?示書・施工通知書・購入�?当通知シート�?�N列を罫線対象外に変更�?
- 施工�?示書・施工通知書シート�?�R列を罫線対象外に変更�?
- O列�?�縮小表示を解除し、�?��?をP列と同じ�?に変更�?

### #6
- 施工�?示書・施工通知書シート�?�Q列右罫線を表示するよう変更�?
- 単価不一致時�?�R列案�??へ、基本�?報シートから�?�取込みする手�??�?2行目として追�?�?
- 単価不一致行�?�高さ�?2行�??に変更�?

### #7
- 購入�?当単価の照合�?�を、空�?のA列（施工業�??��からB列（整�?番号?��へ修正�?
- 購入�?当単価転記ログへ照合対象件数と未一致件数を追�?�?

### #8
- 購入�?当通知では施工業�?列を削除し、整�?番号をA列へ移動�?
- 数量F列以降�?�共通�?�位置を維持するため、B列を空白の区�?り�?�として追�?�?
- 購入�?当通知の単価照合をA列で行い、P列へO列×F列、Q列へ単価比�?を設定�?
- 施工�?示書・施工通知書・購入�?当通知のO1を縮小表示に変更�?

### #9
- 単価不一致時�?�R列�?�取込み案�??につ�?て、取込元A3の値から
  施工通知書また�?�施工�?示書を判定して表示するよう変更�?

### #10
- 施工�?示書・施工通知書シート�?�R列�?�、単価不一致の案�??がある�?�合に列�?59.0で表示し�?
  案�??がな�?場合�?�非表示にするよう変更�?

### #11
- �{�H�w����(�H��)�ŕې���t�B���^�[��Ɏ{�H��Ђ�ݒ肵���ہAI��́uJR���v�v���x����
  J��̍��v�����č쐬����Ȃ��s����C���B�{�H��ЕʒP����X�V���� JR ���v�s���ĕ`�悷��B

### #12
- �{�H�w����(�H��)�E�{�H�w����(�n��)�E�{�H�ʒm���ŁA�t�B���^�[�����ɍ��v�s���f�[�^�r����
  �c��s����C���B�����ԍ���������瑖�����čŏI�f�[�^�s����肵�A���v�s����ɂ��̒�����
  �Ĕz�u����B�{�H��Еʍ��v�����ꏈ���ōĕ`�悷��B

### #13
- �{�H�w�������̎捞�݊������ɕ\�����Ă����������b�Z�[�W�{�b�N�X���\���ɕύX�B
  �捞�݌��ʂ̓��O�o�݂͂̂Ƃ���B�G���[���̃��b�Z�[�W�͏]���ǂ���\������B

### #14
- �{�H�w����(�H��)�E�{�H�w����(�n��)�E�w���[���V�[�g�̎捞�ݍ쐬���ɁA
  �E�B���h�E�g�̌Œ��1�s�ځi�w�b�_�[�s�j���Œ肷��悤�ύX�B

### #15
- �{�H��БI����̒P���E���z��X�V���������B��\�����ς��Ȃ��ꍇ�͗�̍폜�E�}����
  �r���̑S�ʍĕ`����ȗ����A�ύX�s�̂ݕ����X�V����B�P���E���z�͔z��ꊇ�������݂ɕύX�B

### #16
- �{�H��Еʂ̕����X�V��ɋ��z��(L��)�̌v�Z�����ݒ肳�ꂸ�A�P���Z���̊D�F�h��Ԃ���
  �c��s����C���B���z���̍ēK�p�ƒP���̗L���ɉ������h��Ԃ��X�V����Ɏ��s����B

### #17
- �{�H��БI�����Ɂu�I�u�W�F�N�g�ϐ��܂��� With �u���b�N�ϐ����ݒ肳��Ă��܂���v(Err 91)
  �ƌ�\�������s����C���B�}�X�^�Ǎ����Ŏc���� Err �𐳏�I�����ɏE��Ȃ��悤
  �G���[�����𕪗����A��}����̎{�H��З񐔂��V�[�g��̍H�핪�ޗ�ʒu����Ď擾����B

### #18
- �{�H��ЕʒP���̎Q�Ƃ�P���ꗗ�\�̃��[���ɍ��킹�ďC���B�Y�p�����s�����O���A�{�H��З�
  �i����̂݁EStep 2�j�ƋƎ҃}�X�^�ʖ��Ő�����������A���於����P���V�[�g���𒼐ڏƍ�
  ����t�H�[���o�b�N��ǉ��B���l�̂Ȃ��Z���͎捞�o�͂֔��f���Ȃ��B

### #19
- �Ǝ��H��̎���͒P�����{�H��З�֔��f�ł���悤�AJR �Q��(E/F)����ł��{�H��З��
  ���l������΍̗p����B�{�H��ЍđI�����̓V�[�g�S�̂��ďƍ����A�����X�V���s���͈ꊇ�X�V
  �փt�H�[���o�b�N�B�P���V�[�g�����s�̑����E�����Z����̋ƎҖ��E����z��̈��S�Q�Ƃ�ǉ��B

### #20
- �{�H��ЕύX���� Setup �i�K Err 91 �������BSetup ��i�K�� Resume Next �ɕ������A
  �{�H��З�J�n�ʒu�� JR ���z��̉E�ׂ��瓮�I�����B����}�b�v�E�Ǝҕʖ��L���b�V����
  Err.Clear / Nothing �K�[�h�A�P���L���b�V���̏d���o�^�h�~��ǉ��B

### #21
- �{�H�w����(�H��)��A��E�{�H�w����(�n��)��A/B��Ŏ{�H��БI�����A��{���̋Ǝҏ��F�ŃZ����h��B
  ��Ж��̓ˍ��͋Ǝ҃}�X�^�ʖ��E�P���\�\�����̊������[���ɏ]���B

### #22
- �{�H�w����(�n��)�V�[�g�쐬���� F��EI��ER��𒆉������ɂ���B

### #23
- ��{���V�[�g�� C22/C24 �ύX���A�����̎{�H�w�������V�[�g�̎Q�ƒP����
  �i�H�� M ��E�n�� N ��j���ēǂݍ��݂��� `RefreshConstructionReferenceUnitPricesOnExistingSheets` ��ǉ��B
- �Q�ƒP���������ł��Ȃ��s�� M/N ����N���A���A�P����r�E�ē�����Čv�Z����B

### #24
- �{�H�w����(�H��)�̌_����於�ɕt�� `(�O��)` / `�i�O���j` ����於�ƍ����ɏ������A
  �P���V�[�g���E�H�������ʃ}�X�^�Ƃ̓ˍ������P�B
- �]���� S\| �ȗ��ƍ��i���ʓ������ׂď����j�� `(��)(��)(��)` �܂ŏ�����ƍ����邽�ߔp�~�B
  `(�O��)` �̂ݏ������A�������ʂ͕ێ����ďƍ�����B

### #27
- �{�H�w����(�n��)�̎Q�ƒP��(N��)�̓��[���n�ڒP���V�[�g�{�n�ڗp�����ԍ��L�[
  �iBuildWeldingLookupKey�j�ōďƍ�����悤�C���B�H���P���V�[�g�s�ʉ����͗n�ڂɂ͕s�K�B
- �P���s��v���̈ē������n�ڃV�[�g�Ő��ʗ����Q�Ƃ��Ă����s����C���B

### #28
- ���[���n�ڒP���V�[�g�ɕ������悪�c�ɘA�Ȃ�\���ɑΉ��BA��u�ώZ����F�v�s��B�񂩂�
  ���於���擾���A�{�H�w����(�n��)�̌_�����{�����ԍ��̕����L�[��N��Q�ƒP�����ƍ�����B
- �{�H��ВP��(�n��)�̍s���������������L�[�ɓ���B

### #29
- �{�H�w��(�ʒm)��(�H��)�捞�ݎ��AM��(�Q�ƒP��)����Ő����ԍ����P���}�X�^�ɖ����s�ɂ��āA
  G��̑Ώې���((�O��)���O)�V�[�g�̍ŉ����� B/C/E ��̒l�� B/C/D ��Ƃ��ĒǋL����B
  �ǋL�s�� E/F ��(���P��/��P��)�͋�̊ԉ��F�A���͌�͏����t�������ŉ����B�r���� mod_VendorMaster �o�R�œK�p�B

### #30
- G��� `(�O��)` ���t���s�������ԍ����o�^�Ȃ�P���V�[�g�֒ǋL����悤�C���B
  `(�O��)` �̓V�[�g���ƍ����̂ݏ������A�ǋL�Ώۂ��珜�O���Ȃ��i�O�������H���ŒǋL����Ȃ��s��������j�B
- �ǋL��␔�E�V�[�g�����o�E�d���X�L�b�v�E�G���[�����O�o�́B

### #31
- �����ԍ����o�^����P���V�[�g�֒ǋL����ہAA��Ɂu�Ǝ��H��v����͂���B

### #32
- �{�H�w����(�n��)�̎Q�ƒP���ƍ��ŁA�H�������ʃ}�X�^ F��(�ώZ����)�� G��(�{�H�w�����L�ڐ��於)��
  �\�L��(��: �֐� vs �֐��{��)�ɂ�蕡���L�[����v���Ȃ��s����C���B
  ���[���n�ڒP���V�[�g�̐ώZ���於�� G��F �G�C���A�X�ϊ����Ă��琮���ԍ��Əƍ�����B

### #33
- �H�������ʃ}�X�^ F/G ���於�G�C���A�X���Q�Əƍ��S�̂֓���K�p�B
  �H���V�[�g�̒P���V�[�g������(F����v����G�E�G�C���A�X�ł��T��)�AFindImportedUnitPriceSheetName�A
  ResolveUnitPriceSheetName �̃G�C���A�X�t�H�[���o�b�N�A�n�ڒP���L���b�V���ւ̋t�����L�[�o�^�A
  �{�H��ВP��(�n��)�ƍ��� FindWeldingPriceRecordKey ����ǉ��B

### #34
- ���於�G�C���A�X�Ǎ��� C21 �w���1�t�@�C�������łȂ��AC20 �t�H���_(�ݗ����y�ѐV����)����
  �S�H�������ʃ}�X�^ .xlsx(�@�O���������A�C��ڃ����O���[�������H�� ��)����W�񂷂�悤�g���B

### #35
- �{�H��ВP���̐���ƍ��������B�_����於�̐��K���ɑS�p�����E���ʂ̓����ǉ����A
  �H�������ʃ}�X�^ F��(�ώZ����)���� G�񖢋L�ڂł�����₷������(�{�����{�A�擪�Z�O�����g�̐��ȗ���)��
  �����G�C���A�X�Ƃ��ēo�^�B�P���V�[�g�������͌��L�[�����Ɏ��s����悤�ύX�B
- �{�H��З�̍Čv�Z���ɐ��於�G�C���A�X�L���b�V����j�����A�ŐV�}�X�^���Q�Ƃ���悤�C���B

### #36
- �{�H��З�̒P���Čv�Z�������Ȃ��ɂȂ�s����C���B�Čv�Z�̂��тɃ}�X�^�S����
  ADO �ēǍ����Ă�����������߁A����}�b�v�\�z���͎����ƍ���D�悵�ăV�[�g�S�����������B

### #37
- `mod_Construction_Order_Import.bas`�i��5600�s�j��Ӗ��ʂ�7���W���[���֕����B
  �O���Ăяo���͏]���ǂ��� `mod_Construction_Order_Import` �̌��J API ����Ϗ�����\���ɕύX�B

### #38
- �{�H��ВP���̐���ƍ����ċ����B���於���P���V�[�g�Ή��̓o�^�����L�[�S���֊g�����A
  �H���p�̋t�����G�C���A�X(G�񗪏̂̏W��)���ƍ����ɒǉ��B�P���V�[�g���̏d������T�t�B�b�N�X
  `�i�t�@�C�����j` ���������ƍ��ɂ��Ή��B���斢�������̓��O�o�͂�ǉ��B

### #39
- ���W���[��������̌��J API �d���ɂ��u���O���K�؂ł͂���܂���v�������B
  �������W���[�����̎葱������ `Core` �T�t�B�b�N�X�t���ɕύX���A
  `mod_Construction_Order_Import` �݂̂��]���̌��J����񋟂���悤�����B

---

## mod_MaterialPriceImport.bas?��工事単価インポ�?�トモジュール?�?

### #17
- ResetUnitPriceValidation の AlertStyle �? xlValidAlertStop �? xlValidAlertInformation に変更�?
  非表示列をリスト�??に使用する場合、xlValidAlertStop では Excel がリスト外値と判定して入力をブロ�?クするため�?

### #18
- ResetUnitPriceValidation �? MergeArea 対応に修正�?
  C22 等がセル結合されて�?る�?�合、targetCell が結合副セル?�?D22 等）を�?してしま�?
  入力規則が誤ったセルに設定される問題を解消�?MergeArea.Cells(1,1) で代表セルを確実に取得する�?

### #19
- 購入�?当�?�レール溶接単価シート取込後に基本�?報シートがアク�?ィブにならな�?問題を修正�?
  ImportUnitPriceData の全取込処�?完�?後�?MsgBox 直前）に wsInfo.Activate を追�?し�?
  常に基本�?報シートへ戻るよ�?にした�?

### #20
- 新幹線区�?で購入�?当単価の参�?�キーが取得できな�?問題を修正�?
  BuildPurchaseReferenceKey は保線区フォルダ名�?�先�?�数字を参�?�キーとして使って�?たが�?
  新幹線�?�保線区フォルダ名�?�「��沢新幹線保線区」�?�ように数字で始まらな�?ため "" を返し
  「参照キーを取得できませんでした」エラーになって�?た�?
  先�?�数字が取れた�?�合�?�そ�?�まま?��在来線�?�既存動作を維�??���?
  取れなかった�?�合�?�フォルダ名�?�体を正規化して参�?�キーとする�?
  あわせて PurchaseSheetNameMatchesReferenceKey にキーがシート名に含まれる部�?一致パターンを追�?し�?
  漢字キーにも対応した�?

### #21
- 同じ支店�?�出張所で在来線／新幹線�?�単価適用保線区が�?数ある場合、C20 の線区区�?に合うマスタ行を
  優先して選択するよ�?修正。福井�?�張所の新幹線単価適用時に在来線�?�の「敦賀地域鉄道部」を先に拾ってしま�?�?
  新幹線フォルダ配下で見つからな�?問題を解消�?

### #22
- 工事件名に一致する単価表ブックが�?数ある場合、最初�?�1ブック�?けでなく�?�ブックのシートを
  対応積算線区選択フォー�?の候補に出すよ�?修正�?
  新幹線�?�軌道整備他で「基地線」ブ�?ク�?けが候補になり、「本線」ブ�?クが選択できな�?問題を解消�?

### #23
- 単価表ブック検索で同じファイルパスが�?数回返った�?�合�?重�?して候補に追�?しな�?よう修正�?
  対応積算線区選択フォー�?に同じ本線�?�基地線が2回表示される問題を解消�?

### #24
- 新幹線�?�購入�?当単価表で、シート名が「福井新幹線保線区」ではな�?
  �?2026単価一覧表(材料)福井」�?�ような短縮名称の場合でも取り込めるよう修正�?
  フルキー一致、短縮キー一致、単一シート採用の�?で判定する�?

### #25
- 購入�?当単価シート�?�タブ色�? #FFCC99 から #99FFCC に変更�?

### #26
- 工事単価シート�?�タブ色�? #DDEBF7 から #A5ABE5 に変更�?

### #27
- 溶接単価ブックが無�?保線区で、C23 が溶接ありの場合に「○○には溶接単価の設定がありません。」と警告するよ�?修正�?
- C23 の溶接工事有無判定を MergeArea と表記ゆれに対応�?

### #28
- C24 単価シート作�?�ロジ�?クに [UnitPrice] タグ付き�?バッグログを追�?�?
  mod_DebugLog 経由でイミディエイト＋ログシートに記録する�?
- 追跡対象は ImportUnitPriceData の全フェーズ?�?
  TryReadUnitPriceRequest、TryLoadUnitPriceMasterRow、ResolveUnitPricePriceFolderPath�?
  ResolveUnitPriceSourceFilePaths、ImportSelectedUnitPriceSheets�?
  ImportPurchaseUnitPriceSheetsByReference、ImportWeldingUnitPriceSheetsIfRequired、WriteSelectedLineNames�?
- 既存�?� [FillMgr] ログと同形式�?�同タグ運用�?

### #29
- �H���P���捞�������Ɋ����{�H�w�������V�[�g�̎Q�ƒP�����ēǂݍ��݂���悤�A�g�B
- �捞�������� `IsImportingUnitPriceData` �� C24 �ύX�C�x���g����̓�d�X�V��}�~�B

### #30
- ���[���n�ڒP���V�[�g�捞���AB�񂪐��l(�����ԍ�)�łȂ��s��E/F��͋󔒓h��Ԃ��ΏۊO�Ƃ���B
- A��u�H�������F�v�u�ώZ����F�v�s�̉E��(B��)�����l�ߕ\�����鏑��������ǉ��B

### #31
- �w���[���P���捞�݂ō쐬����V�[�g��A��Z������13�ɐݒ�B

### #32
- �H�������E����敪�E�P���K�p�敪�̓��͋K�����X�g�������ݗ�� AE/AF/AG ���� AM/AN/AO �ֈړ��B
  �{�H��Ѓu���b�N�ő�10�ЁiE?AH��j�Ƃ̗�Փ˂�����B

---

## mod_FillManagerName.bas?���?�張所長�? 自動�?�力／支店�?�出張所バリ�?ーション再構築モジュール?�?

### #7, #9, #10, #15, #16, #17
- �?種改修済み?��詳細はコード�??コメント参照?���?

### #19
- RefreshBranchOfficeValidation の戻り値?�?Boolean?��で制御するよう変更�?

### #22
- B6変更時に C6 コンボが開かな�?問題を修正�?

### #23
- B6変更時コンボ表示フロー追跡用ログを挿入?�?mod_DebugLog 使用?���?
  �?バッグ確認後�?� mod_DebugLog の呼び出しを削除すること�?

### #24
- �x�X�E�o�����̓��͋K�����X�g�������ݗ�� AA/AB ���� AK/AL �ֈړ��i�{�H���10�Ѓu���b�N�Ƃ̏Փˉ���j�B
- ���e���v���[�g T:AG �s2?9 �̃��X�g�c�[���� `CleanupLegacyBasicInfoValidationListDebris` ��ǉ��B

---

## mod_VendorInfoColors.bas (vendor info colors)

### #1
- Fixed background/foreground palette for vendor blocks 1-10 (�Ǝҏ��F).
- Apply to basic-info row 10 (3-column blocks) based on F9 subcontractor count.
- Public `GetVendorInfoColorBackground` / `GetVendorInfoColorForeground` for future sheet tab colors.
- Called from `mod_VendorMaster.InitVendorBlockCountFromSheet` and `SyncVendorBlocksFromCount`.

### #2
- Row 10 coloring applies to label/value columns only (E/F per block); spacer column (G etc.) restores default dark fill.

### #3
- `ApplyOutputSheetVendorCellColor` colors construction output sheet vendor cells (A or welding A/B) by vendor info index.

---

## mod_VendorMaster.bas?��業�?マスタ参�?�?��業�?選択モジュール?�?

### #7
- WriteVendorValidationList の Dictionary→セル単位ループを
  Variant 2次�?配�?�＋Range 一括代入へ置換�?

### #8
- LoadVendorComboBoxItems のセル単位走査を�?��?�読込ループへ変更�?

### #9
- NormalizeText / CommonGetBasicInfoWorksheet / 日本語名生�?? / ADO 接続生成�?�
  mod_Common に�?�?。重�?定義を撤去�?

### #18
- 業�?�?報8行目に2行挿入のため、F列セル参�?��? +2行シフト�?
  BASIC_INFO_VENDOR_NAME_CELL: F9 �? F11
  BASIC_INFO_VENDOR_CLEAR_RANGES: F9:F14,F16:F21 �? F11:F16,F18:F23
  ApplyVendorRowToBasicInfo: F9→F23 �?囲を対応更新
  FitVendorComboBoxToF9 �? FitVendorComboBoxToF11

### #29
- 適用積算線区の単価シートに表示する施工会社名を、基本�?報 B6 の支店シートにある
  業�?マスタ B列�?�会社名で照合し、対応す�? A列�?�値へ変換するよう変更�?

### #30
- 施工�?示書・施工通知書の施工会社候補も、業�?マスタB列で照合したA列値へ変更�?
- A列で使用中の施工会社ごとに、J列�?�右へ「会社名＋単価」「会社名＋��額」�?�2列を追�?�?
- 契�?線区名、整�?番号、昼夜別、単価表の会社名�?�ッダーを�?�合して会社別単価を転記し�?
  金額�?�へ会社別単価×数量�?�数式を設定�?

### #31
- �O���H���u���b�N30�s��(�n�ڎ茳�P���p�^�[��)��P���Ď��Ώۂɒǉ��B�ύX���ɗn�ڒP�����ēW�J����B

---

## mod_WeldingUnitPrice.bas�i���[���n�ڒP�� �{�H��ЕʒP���W�J���W���[���j

### #1
- �O���H���u���b�N31�s�ڂւ�60.1%��������(WriteRailOutsourceRatioToBasicInfo)��p�~�B
- �O����ВP�������̃p�^�[���Q�Ƃ����[���n�ڒP���V�[�gF3�����{���30�s��(F30/I30�c)�֕ύX�B
- �p�^�[���I�������u�O�N�x�P���K�p�p�^�[���v�u�O���䗦�K�p�p�^�[���v�u�����w���K�p�p�^�[���v�ɓ���B

### #2
- ��{���Ŏ{�H��Ж���ύX�����ہA���[���n�ڒP���V�[�g�Ɋ�{���ɖ����{�H��З񂪎c��s����C���B
- �H���P���V�[�g�Ɠ��l�A��Ж��ύX���ɗ]�����N���A����{���̃u���b�N�\���֓�������B

---

## mod_BasicInfoGuide.bas�i��{���V�[�g ���̓K�C�h�Ǘ����W���[���j

### #1
- �����̓K�C�h�i���F�h��j�̃Z�������F�����ɕύX���A���͊m���i#06111D�j�̕����F�𔒂ɐݒ�B
- ���͕s�Z���i�ΐ��~�j����щ�А��O�Z���̃N���A������������K�p�B

### #2
- �H���敪���O���H���̂Ƃ��A����30�s�ڃ��x�����u�n�ڎ茳�P���p�^�[���v�ɐؑցB
- �H���敪��30�s�ڂ�3���h���b�v�_�E��(�O�N�x/�O���䗦/�����w�� �e�u�K�p�p�^�[���v)��F11�����̓��̓K�C�h��K�p�B

---

## mod_BasicInfoUpdate.bas?��基本�?報シー�? 期間?��請求回数更新モジュール?�?

### #9
- NormalizeText / GetBasicInfoWorksheet / 日本語シート名生�?��?� mod_Common に�?�?済み�?
  重�?定義を撤去し、�?�通関数経由で参�?��?

### #11
- SilentClearBasicInfo を追�?�?B6/C6 変更時に確認メ�?セージなしで基本�?報と単価シートをクリア�?

### #15
- エラーハンドラ�?で Err.Number / Err.Description を退避し�?
  そ�?�後�?� Application 設定復帰を確実に行う構�?へ整�?�?

### #18
- 業�?�?報8行目に2行挿入のため、BASIC_INFO_CLEAR_RANGES の
  F列�?囲�? F9:F14,F16:F21 �? F11:F16,F18:F23 に更新�?

### #21
- SilentClearBasicInfo / ClearBasicInfo で EnableEvents を上書きしな�?よう保存�?�復�?するように変更�?
  親呼び出し�???�?Worksheet_Change など?��が EnableEvents=False 中にこれらを呼び出して�?
  EnableEvents 状態を損なわな�?ようにする�?

### #22
- BASIC_INFO_CLEAR_RANGES �Ɋe�Ǝ҃u���b�N30�s��(F30,I30,�c)��ǉ��B

### #23
- ClearBasicInfo: �{�H�w�����V�[�g�폜�� SyncVendorBlocksFromCount ���O�Ɏ��s���A
  ���v�Čv�Z�� C31:C35 ��������������C���B
- ClearBasicInfoYenTotalCells ��ǉ��i�����Z���Ή��j�BSync ��ɂ��ď����B

---

## mod_common.bas?���?�通ユー�?ィリ�?ィモジュール?�?

- �?モジュールに散ら�?�って�?た汎用関数?��文字�?�正規化、シート取得、日本語名生�?��?
  ADO 接続生成、年抽出�?長整数配�?�ソートなど?��を�?�?�?
- ChrW で構築する日本語名は Static 変数でキャ�?シュし、�?�構築�?�コストを排除する�?
- 全モジュールから本モジュールの関数を呼ぶ前提とする�?

---

## mod_OrderTpl_Shared.bas / mod_OrderTpl_Generate.bas / mod_OrderTpl_Header.bas / mod_OrderTpl_Detail.bas（注文書テンプレート取込モジュール）

### #1（新規作成 2026-07-08）
- 基本情報シートの施工会社（11行目、F列始まり3列おき）確定時に、`マスタデータ/注文書テンプレート.xlsx` の
  5シート（内訳明細・受注者用・注文請書・支店控・別紙Ⅲ）を購入充当指示（通知）シートの右側へ会社ごとに挿入する。
- シート名末尾には業者マスタ(全社版)のP列（略称、例 `(泉州)`）を付帯。照合キーは基本情報11行目の会社名と業者マスタB列（請求者氏名）。
- 2社目以降は直前ブロックの別紙Ⅲ(略称)の右へ順に追加。同一会社の再確定時は既存5シートを削除して再作成。
  会社名クリア時は対応する5シートを削除（孤児シート後始末）。
- 内訳明細ヘッダー転記: C2=部店コード（出張所別_単価適用線区の単価適用線区シートB/C列照合→G列）、
  C3=注文番号（ブロック27行目）、B6=基本情報C9、C6=C10、K5/K6=C15/C16（和暦）、N2=C2（和暦）、N3=F6、
  O5=ブロック11行目、O6=ブロック16行目。結合セルは上下左右中央揃え。入力文字はすべて BIZ UDゴシック。
- 内訳明細明細転記（11行目〜小計行直前）:
  - 軌道工事: 施工指示書(工事)/施工通知書(工事)（存在する方を自動使用）をA列施工業者で抽出→
    2行空けて施工指示書(溶接)/施工通知書(溶接)をB列軌道手元会社で抽出。
  - 溶接工事: 溶接シートをA列溶接会社で抽出のみ。
  - 業者名の突合は既存の名寄せ（`mod_Construction_BasicTotals.GetVendorAliasMap`/`ResolveVendorCanonicalKey`）を使用し、
    業者名（略式）・請求者氏名（正規名）どちらの表記でも一致させる。
  - セクション見出しは「契約線区名_管理室」（工事は`(軌道)`除去、溶接は`(溶接指示書用)`除去+`_レール溶接`付加）。
    セクション間に空白1行。セクション内は昼→夜、整理番号昇順。
  - 列マッピング: A=整理番号(中央)、B=工事種類(左詰め縮小)、C=昼夜別(中央)、D=単位、E=数量、F=JR単価(整数)、
    G==E*F(General)、N/O/P=E/F/G と同内容。数量書式は単位により整数(口/穴/回/個/箇所/本/組/式/枚)・
    小数3桁(m/m3/M/m²/t)、溶接は桁切りなし。
  - 行不足時は小計行の直前へ行挿入し、G:P列の数式・書式をテンプレート行から引き継ぐ。7〜10行目を印刷タイトル行に設定。
- 施工指示書等が未取込の場合はヘッダーのみ転記（ログ記録）。取込後に `RefreshAllVendorOrderDetails`
  （Alt+F8のマクロ一覧から実行可）で全確定会社の内訳明細を再転記できる。
- 業者マスタ未一致・P列略称未入力・同一会社の重複入力・テンプレート未検出時はメッセージを表示して中止。

## Sheet1.cls（基本情報シートのイベント）

### 追記（2026-07-08）
- 施工会社セル（11行目）変更処理の末尾に `mod_OrderTpl_Generate.HandleVendorNameCellChange` の呼び出しを追加。
- F9（下請負会社数）変更処理（`SyncVendorBlocksFromCount` 直後）に `RemoveOrphanGeneratedSheets` を追加。会社数を減らした時点で対象外会社の注文書テンプレート5シートを削除する。

## mod_Construction_OutputLayout.bas / mod_BasicInfoUpdate.bas（出力シート名の表記統一）

### 追記（2026-07-08）
- 施工指示書・施工通知書の取込時に生成する出力シート名を「施行」表記へ統一
  （施工指示書(工事)→施行指示書(工事)、施工通知書(溶接)→施行通知書(溶接) 等）。
  `BuildConstructionSheetName` で取込元A3由来の文書名を変換する。
- `NormalizeOutputSheetNameKey` は施行/施工の表記ゆれを同一キーへ正規化し、
  旧表記で作成済みのシートも既知テンプレートシートとして認識する(後方互換)。

## mod_OrderTpl_Generate.bas / ThisWorkbook.cls / mod_subcontractorselector.bas（施工会社割当ての自動再転記）

### 追記（2026-07-08）
- 施行指示書・施行通知書シートの施工会社列（工事:A列、溶接:A・B列）が変更されると、
  各社の内訳明細を自動で再転記するようにした。
  - 手入力・ドロップダウン変更: `Workbook_SheetChange` から `ScheduleOrderDetailRefresh` を予約
    （1秒の遅延実行で連続変更を1回にまとめる。`Application.OnTime` 使用）。
  - 選択フォームからの一括割当て: `SelectSubcontractorForSelection` の書込完了後に
    `RefreshAllVendorOrderDetailsSilent` を即時実行。
- `RefreshAllVendorOrderDetails` を Core 化し、完了メッセージなしの
  `RefreshAllVendorOrderDetailsSilent` と OnTime 用 `RunScheduledOrderDetailRefresh` を追加。
- `Workbook_BeforeClose` に遅延予約のキャンセル(`CancelScheduledOrderDetailRefresh`)を追加。
- 併せて、シートコピー後のリネーム参照を `Worksheets(...)` から `Sheets(...)` へ修正
  （アンカー `Index` との整合。再生成時の巻き戻りを修正）。

## mod_OrderTpl_Detail.bas（集計ブロックの動的配置）

### #2（2026-07-08）
- 小計行の扱いを「下へずらす」方式から「最終データ行から2行空けた位置へ配置」する方式へ変更
  （`PositionSubtotalRow`: 不足時は行挿入・余剰時は行削除で小計行位置を調整）。
- 小計行の下に 値引・計・消費税・合計・罫線用空白行 を生成する `BuildSummaryBlock` を追加。
  - ラベルは A:C 結合・上下左右中央揃え・BIZ UDゴシック。
  - G/J/M/P 列: 小計=SUM(11:最終データ行)、値引=-MOD(小計,1000)（百の位以下をマイナス表示）、
    計=小計+値引、消費税=ROUNDDOWN(計×税率,0)（税率は基本情報B34のカッコ内を
    `ResolveBasicInfoTaxRate` で解析）、合計=計+消費税。表示形式は #,##0（桁区切り・小数なし）。
  - 罫線: 小計行の上罫線=二重線(A:P)、ブロック内横罫線=細線、罫線用空白行の下罫線=中線(A:P)。
    縦罫線は行挿入時に上方セルから継承。
- 再転記時は既生成の集計ブロック(値引〜罫線用空白行)を削除してから再構築する(`ResetDetailArea` 拡張)。

## mod_Construction_LineMapping.bas（(LR化)等の線区名接尾辞対応）

### 追記（2026-07-09）
- 契約線区名の末尾に工事種別マーカー（例: `(LR化)`）が付くと工事件名マスタG列
  （施工指示書記載線区名）と一致せず、単価シート名が解決できない不具合を修正。
  年初単価（M列）が未転記になり単価比較（O列）が常に「単価不一致」となっていた。
- `CollectLineLookupCandidates` に「末尾の括弧書きを除いた表記」を最終候補として追加
  （`RemoveTrailingParenGroup` 新設）。既存候補が優先されるため、阪和線(南)・関西空港線(東)・
  側線(昼間拡大) 等の正規名称の照合には影響しない。(LR化)以外の未知のマーカーにも対応。
- 独自工種ガイダンス文言のフォールバック表示（単価シート名未解決時）も末尾の括弧書きを
  除去した線区名を表示するよう変更（「…(LR化)シートに入力してください」→「…シートに入力してください」）。

## mod_Construction_LineMapping.bas / mod_OrderTpl_Shared.bas（(溶接通知書用)マーカー対応）

### 追記（2026-07-09）
- 施工通知書取込では溶接側の契約線区名の接尾辞が「(溶接通知書用)」となるが、
  `RemoveWeldingInstructionMarker` は「(溶接指示書用)」しか除去せず、溶接単価の行照合キーが
  一致しないため年初単価（N列）が未転記・単価比較が常に「単価不一致」となる不具合を修正。
  同関数で「(溶接通知書用)」（半角/全角括弧とも）も除去するようにした。
- mod_OrderTpl_Shared の `OrderTplStripLineSuffix`（内訳明細のセクション見出し生成）も
  自前の接尾辞除去をやめ、`RemoveWeldingInstructionMarker` / `RemoveTrackDesignationMarker` へ
  委譲する形に変更（通知書表記にも自動対応）。

## mod_Construction_LineMapping.bas（(LR化)マーカーの除去）

### 追記（2026-07-09）
- `RemoveTrackDesignationMarker` で「(LR化)」（半角/全角括弧、半角/全角LRの計4表記）も
  除去するようにした。内訳明細のセクション見出し（`OrderTplStripLineSuffix` 経由）と
  工事側の単価シート名照合の両方で (LR化) が除去される。

## mod_VendorUnitPrice.bas（単価シート塗りつぶし範囲の修正）

### 追記（2026-07-09）
- 単価シートの追記行（独自工種行など）の塗りつぶし範囲下限が J列固定
  （`VENDOR_UNIT_PRICE_INITIAL_FILL_LAST_COL = 10`）だったため、確定業者が1社のみでも
  I・J列（2社目ブロック）が塗られる不具合を修正。下限を F列（=6、単価元E/Fの右端）へ変更し、
  `GetVendorUnitPriceInitialFillLastColumn` が確定業者のブロック数に応じて拡張する
  （1社=E〜H、2社=E〜J …）。`mod_VendorMaster.bas.pre_split` の同宣言も追随更新。

## mod_OrderTpl_Header.bas / mod_OrderTpl_Generate.bas / Sheet1.cls（ヘッダーのライブ反映）

### 追記（2026-07-09）
- テンプレート取込後に基本情報シートを入力・変更しても内訳明細のヘッダー
  （C3注文番号・N2作成日・N3所長・O5外注会社名・O6業者コード等）へ反映されない問題を修正。
  従来は会社確定時の一括転記のみだったため、後から入力した項目が空欄のままだった。
- `HandleBasicInfoHeaderSourceChange` を新設し、Sheet1(基本情報)の `Worksheet_Change` から
  転記元セル（C2/C9/C10/C15:C16/F6、各ブロックの16行目=業者コード・27行目=注文番号）の
  変更を検知して全確定会社のヘッダーを再転記する。監視ゲートに C2・F6・各ブロック16行目を追加。
- `ApplyVendorSheetHeaders`（ディスパッチャ）を新設。内訳明細に加え、受注者用・注文請書・
  支店控・別紙Ⅲの転記スタブを用意（転記仕様が確定したら各スタブへ実装すれば、
  生成時・基本情報変更時の両方で自動的に反映される設計）。
  生成/再転記処理はディスパッチャ経由へ変更。

## mod_VendorBlockLayout.bas / mod_OrderTpl_Header.bas / Sheet1.cls（F9同期直後の反映・タブ色）

### 追記（2026-07-09）
- F9（下請負会社数）変更でブロックを同期した直後、11行目で会社名を選択しても
  業者情報・注文書シートへの転記が始まらない不具合を修正（基本情報シートを
  一度離れて戻るまで反映されなかった）。原因は `SyncVendorBlocksFromCount` の
  正常終了パスが `Exit Sub` で同期中フラグ（`mSyncVendorBlocksInProgress`）を
  解除しないまま抜けており、以後の Worksheet_Change / Workbook_SheetChange が
  ガードで無視されていたため。正常終了時も ExitHandler へ落としてフラグ解除・
  状態復元を必ず行うようにした。
- 生成される5シート（内訳明細〜別紙Ⅲ）のシート見出し（タブ）色に、対応する
  工事区分列（基本情報10行目）セルの塗りつぶし色を適用（`ApplyVendorSheetTabColor`）。
  工事区分の変更時も追従する（ヘッダー転記元の監視範囲へ10行目を追加）。
  Sheet1 のヘッダー反映フックはガイド更新（10行目の色付け）後に実行する順序へ変更。

## mod_OrderTpl_Detail.bas（明細部の書式・罫線・セクション見出しの改善）

### #3（2026-07-09）
- 明細部(11行目以降)の表示形式を一新:
  - E〜P列すべてに桁区切り書式（ゼロ値は非表示: `#,##0;-#,##0;`）。
  - 数量列（E/H/K/N）は単位により整数（口/穴/回/個/箇所/本/組/式/枚）・
    小数3桁（m/m3/M/㎡/t: `#,##0.000;-#,##0.000;`）。変更(H:I)・今回(K:L)の手入力にも書式適用済み。
  - 11行目以降（集計ブロック含む）のフォントサイズを10ポイントへ統一。
- 罫線:
  - 明細部（11行目〜小計行の直前）の横罫線を転記の最後にすべて細線で引き直す。
  - 集計ブロック（小計行〜合計行の1行下）の縦罫線: D/E間・G/H間・J/K間・M/N間=中線、
    E〜G間・H〜J間・K〜M間・N〜P間=細線。
- セクション見出し行（契約線区名・管理室名）: A:C を結合し、左詰め・縮小表示でセル内に収める。
  再転記時は結合を解除してから明細をクリアする。
- 施行通知書には管理室名が無いため、管理室名が空の場合は見出し末尾の「_」を付けない
  （施行指示書は従来どおり「線区名_管理室名」）。

## mod_OrderTpl_Generate.bas / ThisWorkbook.cls（コントロールイベント内シートコピー失敗の修正）

### 追記（2026-07-09）
- 会社名をActiveXコンボ等で確定した際に「注文書シートの作成中にエラーが発生しました。
  このシートをコピーできませんでした。」(実行時エラー1004)となる問題を修正。
  コントロールイベントの実行コンテキスト内（コントロールがフォーカス保持中）では
  `Worksheets.Copy` が失敗するというExcelの制限が原因。
- シート生成を `Application.OnTime` による遅延実行へ変更（`ScheduleVendorSheetGeneration` /
  `RunScheduledVendorSheetGeneration`）。イベント完了・フォーカス正常化後に生成されるため、
  入力経路（直接入力・コンボ・選択フォーム）によらず安定して動作する。
  複数ブロックの連続確定は予約をまとめて1回で実行。`Workbook_BeforeClose` で予約をキャンセル。

## Sheet1.cls / mod_OrderTpl_Header.bas / mod_OrderTpl_Generate.bas（欠損復旧・未実装関数の補完）

### 追記（2026-07-09）
- Sheet1.cls が末尾欠損（`GetBasicInfoChangeGateRange` の途中で切断）していたため復旧。
  変更監視ゲートのリファクタ（`IsIgnoredBasicInfoChange` 方式）を完成させた:
  `GetBasicInfoChangeGateRange`（ヘッダー転記元+ガイド監視+固定セル+業者ブロック10社分の
  10/11/16/27/29/31行目）、`IntersectsBasicInfoMonitorTarget`、`GetTargetValueText` を実装。
- 参照されていた未実装関数を実装:
  - `mod_OrderTpl_Header.GetBasicInfoHeaderSourceMonitorRange`（監視範囲の公開ラッパー）
  - `mod_OrderTpl_Generate.RemoveAllGeneratedOrderTemplateSheets`（生成済みテンプレートシートの
    全削除。基本情報クリアの `DeleteConstructionImportSheets` 冒頭から呼ばれる）

## mod_OrderTpl_Detail.bas / mod_OrderTpl_Header.bas（テンプレート列追加への対応）

### 追記（2026-07-09）
- 注文書テンプレート(内訳明細)のB列とC列の間に1列追加された新レイアウトへ対応
  （C列以降が1列右へシフト: 昼夜別=D、単位=E、数量=F/I/L/O、単価=G/J/M/P、金額=H/K/N/Q）。
- 新テンプレートには金額数式が入っていないため、金額列(H/K/N/Q)へ
  `=数量×単価`(R1C1: =RC[-2]*RC[-1]) の数式をコード側で明細部全行に設定するようにした。
- 数量列(F/I/L/O)の単位別書式（整数/小数3桁・桁区切り・ゼロ非表示）を新列位置へ適用。
- セクション見出し・集計ブロックのラベル結合を A:D へ変更(テンプレートの小計 A33:D33 に合わせる)。
- 罫線・フォント適用範囲を A〜Q へ拡張。縦罫線は E/F・H/I・K/L・N/O間=中線、
  F〜H・I〜K・L〜N・O〜Q間=細線へシフト。
- ヘッダー転記先を新位置へ変更: 工事名称=D6、工期=L5/L6、作成日=O2、所長=O3、
  外注会社名=P5、業者コード=P6（部店コードC2・注文番号C3・工事番号B6は結合拡張のみで位置不変）。

## mod_BasicInfoOrderNumberKeypad.bas / frmNumericKeypad.frm / clsKeypadBtn.cls / Sheet1.cls（注文番号セルの数値入力補助フォーム追加）

### 追記（2026-07-10）
- 基本情報シートの施工会社ブロック27行目(注文番号、`mod_OrderTpl_Shared.ORDER_TPL_BLOCK_ORDER_NO_ROW`)
  セル(F27/I27/L27...、F9の施工会社数に応じて3列おきに最大10社分)をダブルクリックすると、
  テンキー形式の数値入力補助フォーム`frmNumericKeypad`を表示するようにした。
- フォームは実行時に`Controls.Add`で数字ボタン(0-9)・クリア・バックスペース・OK・キャンセルを
  動的生成し、ボタンのクリックは`clsKeypadBtn`(WithEvents)で受け止める。
- ダブルクリックイベント中の同期モーダル表示によるハングを避けるため、
  `mod_PrefectureSelector`と同様に`Application.OnTime`で1秒後に遅延起動する方式とした。
- 確定時は入力値をセルへ書き戻した上で`mod_BasicInfoGuide.OnCellChanged`と
  `mod_OrderTpl_Header.HandleBasicInfoHeaderSourceChange`を呼び出し、注文書テンプレート
  各シートのヘッダーへ即時反映されるようにした。
- `Sheet1.cls`の`Worksheet_BeforeDoubleClick`へ判定・起動呼び出しを追加。
- frmNumericKeypad.frm の末尾切れによるコンパイルエラー（BackspaceKeyText が As Str で未完、InvalidNumberText 欠落）を修正。
- frmNumericKeypad のキャンセルボタン幅を拡大し、「キャンセル」が見切れないよう調整。

## mod_OrderTpl_Shared/Detail.bas / mod_Construction_BasicTotals.bas / ThisWorkbook.cls（単価転記・罫線・消費税行）

### 追記（2026-07-10）
- 内訳明細の単価列（G/J/M/P）の取得元を「JR単価」から施行指示書(工事/溶接)・施行通知書(工事/溶接)
  シートの施工会社別単価列（ヘッダー「会社名+単価」例: 藤本軌道㈱単価）へ変更。
  会社名は業者マスタの名寄せ（正規名/略称）で照合し、該当行（整理番号一致=同一行）の数値を転記する。
  軌道工事=工事シート(+レール溶接セクションは溶接シート)、溶接工事=溶接シートから取得。
  単価列未展開・未割当時は空欄としログに記録。
- J列/M列の単価は、I列/L列（前回迄/今回の数量）に数値が入力された時のみ自動入力する
  （`HandleBreakdownQuantityCellChange` を `Workbook_SheetChange` から起動。数量を消すと単価もクリア）。
- 11行目〜小計行の直前のB:C列を行ごとに結合（セクション見出し行はA:D結合のため除外）。
- 縦罫線を再定義: D/E列間=11行目〜小計行の直前のみ（小計行以下は消去）、
  G/H・J/K・M/N・P/Q列間=11行目〜合計行の1行下まで。集計ブロック内の他の縦罫線は消去。
- 基本情報シートの各施工会社ブロックに34/35行目を追加生成:
  ラベル列へB34:B35（消費税(10%)：/税込み金額：）のデータ・罫線をコピーし、
  34行目=33行目(契約金額税抜)×税率（B34カッコ内、C34と同じ切り捨て）、35行目=33+34行目。
  書式は33行目を模倣。`RefreshBasicInfoConstructionTotalsCore` の末尾で毎回更新するため、
  F9の会社数増減・金額変動に追従し、対象外ブロックの34/35行目は消去する
  （`RefreshVendorBlockTaxRows` 新設）。
- RefreshVendorBlockTaxRows の対象外ブロック消去を ClearContents+背景復元に変更し、WorkbookOptimize が施工会社34/35行目の UsedRange 伸長で AH 以降の塗りつぶしを消さないよう列別判定に修正。
- 内訳明細: O2/L5/L6を14pt、集計枠罫線(A:Q)拡張、金額列ゼロ非表示(I12/L12連動含む)を調整。

## mod_BasicInfoLaborInsurance.bas / frmNumericKeypad.frm / clsRosaiCheck.cls（労災保険 加入負担 選択）

### #1
- 基本情報シート施工会社ブロック **39行目（労災保険 加入負担）** の会社列セル（F/I/L/O/R/U/X/AA/AD/AG）をダブルクリックすると、`frmNumericKeypad` を**労災モード**で再利用し、甲/乙のチェックボックスを上下に並べた選択フォームを表示する（Caption「労災保険 加入負担選択」、`Application.OnTime` で遅延起動）。
- どちらかにチェックが入った段階でフォームを閉じ（甲/乙は排他）、対象セルへチェック字形を横並びで書き込む（例: OFF甲 / ON乙 → セル中央揃え）。
- 甲の状態を対応する **受注者用（略称）シートの H30**、乙の状態を **J30** へチェック字形単体で上下左右中央揃え表示（受注者用シート未生成時はセル書込のみ）。
- `frmNumericKeypad`: New 直後に `ConfigureLaborInsuranceMode` を呼ぶと、Initialize が生成したテンキーを撤去して甲/乙チェックボックスへ組み替える（既存のテンキー機能は非破壊）。チェック字形は CP932 外のため `ChrW$(&H2610/&H2611)` で構築。
- `clsRosaiCheck.cls`: 動的生成チェックボックスの Click を WithEvents で受け、フォームの `HandleRosaiCheck` へ中継。
- `Sheet1.cls`: `Worksheet_BeforeDoubleClick` に `IsLaborInsuranceTarget` 判定と`RequestLaborInsuranceSelection` の分岐を追加。

