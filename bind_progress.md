# Binding 진행 현황 — SDK method 중 Rust safe binding이 없는 것

기준일: 2026-10-09 (commit `0dda932`)

## 산정 기준

- **SDK 쪽**: `hwp_sdk_doc/HwpObject.tlb.IDL`의 각 `interface`에서 `propget`/`propput`이 아닌 멤버(= method)만 대상으로 했다.
  property는 이 문서의 범위가 아니다.
- **Rust 쪽**: `crates/hwp_core/src/`의 `method/`(SDK method 층)와 `pset/`에서 해당 interface에 대응하는 타입의 `impl` 블록이
  그 COM 멤버 이름을 실제로 호출하면 "binding 있음"으로 보았다. `HwpObject::call_with("Name", ..)` 같은
  범용 문자열 호출로만 쓸 수 있는 것은 "binding 없음"이다.
- `actions/*.rs`의 Action 래퍼(`hwp.act()`, `Run("ActionId")` 기반)와 `ext/`의 편의 API는 IDL method가 아니므로 여기서 다루지 않는다. Action 쪽 계획은 `plan.md` 참고.
- `HwpCtrl.tlb.IDL`(HwpCtrl ActiveX)과 ParameterSet interface(`HCharShape` 등 — method 없이 property만 있음)는 제외했다.
- `※` 표시는 IDL에는 있지만 `HwpAutomation_2504.pdf` 본문에서 이름이 검색되지 않는 method다(문서화 안 됨, 동작은 실환경 확인 필요).
- 시그니처 표기: `[in]` 인자는 `이름: 타입`, `[out]` 인자는 `out 이름: 타입`, `[optional]`은 `[이름: 타입]`, `[out, retval]`은 `-> 타입`.

## 요약

| Interface | Rust 타입 | method 수 | binding 있음 | binding 없음 |
|---|---|---:|---:|---:|
| `IHwpObject` | `HwpObject` | 174 | 16 | 158 |
| `IDHwpAction` | (없음) | 5 | 0 | 5 |
| `IDHwpParameterSet` | (없음) | 11 | 0 | 11 |
| `IDHwpParameterArray` | (없음) | 4 | 0 | 4 |
| `HAction` | `HAction` | 4 | 4 | 0 |
| `HSet` | (없음) | 1 | 0 | 1 |
| `IDHwpCtrlCode` | (없음) | 1 | 0 | 1 |
| `IXHwpDocuments` | `XHwpDocuments` | 5 | 2 | 3 |
| `IXHwpDocument` | `XHwpDocument` | 9 | 1 | 8 |
| `IXHwpWindows` | `XHwpWindows` | 2 | 0 | 2 |
| `IXHwpWindow` | `XHwpWindow` | 2 | 0 | 2 |
| `IXHwpToolbarLayout` | `XHwpToolbarLayout` | 10 | 9 | 1 |
| `IXHwpToolbar` | `XHwpToolbar` | 4 | 3 | 1 |
| `IXHwpToolbarButton` | `XHwpToolbarButton` (method 없음) | 7 | 0 | 7 |
| `IXHwpToolbarComboBox` | (없음) | 16 | 0 | 16 |
| `IXHwpToolbarMenuButton` | `XHwpToolbarMenuButton` | 3 | 1 | 2 |
| `IXHwpTabs` | (없음) | 2 | 0 | 2 |
| `IXHwpTab` | (없음) | 1 | 0 | 1 |
| `IXHwpPrint` | (없음) | 5 | 0 | 5 |
| `IXHwpRange` | (없음) | 36 | 0 | 36 |
| `IXHwpFind` | (없음) | 6 | 0 | 6 |
| `IXHwpSelection` | (없음) | 23 | 0 | 23 |
| `IXHwpFormComboBox` | (없음) | 4 | 0 | 4 |
| `IXHwpCharacterShape` | (없음) | 10 | 0 | 10 |
| `IXHwpParagraphShape` | (없음) | 19 | 0 | 19 |
| `IXHwpMessageBox` | (없음) | 1 | 0 | 1 |
| `IXHwpSendMail` | (없음) | 1 | 0 | 1 |
| **합계** | | **366** | **36** | **330** |

이와 별도로 이벤트 interface `IHwpObjectEvents`의 콜백 16개는 전부 미구현이다(맨 아래 참고).

## `IHwpObject` — 미구현 158 / 174

binding 있음: `SaveAs`, `SelectText`, `MovePos`, `InitScan`, `ReleaseScan`, `GetText`, `GetPos`, `KeyIndicator`, `GetTextFile`, `SetTextFile`, `Run`, `Open`, `Clear`, `Quit`, `SetMessageBoxMode`, `RegisterModule`

### 파일 · 문서 (13)

- [ ] `Save(save_if_dirty: VARIANT) -> VARIANT_BOOL`
- [ ] `Insert(Path: BSTR, Format: VARIANT, arg: VARIANT)`
- [ ] `CreatePageImage(Path: BSTR, pgno: VARIANT, resolution: VARIANT, depth: VARIANT, Format: VARIANT) -> VARIANT_BOOL`
- [ ] `GetPageText(pgno: long, option: VARIANT) -> BSTR`
- [ ] `GetFileInfo(filename: BSTR) -> IDispatch`
- [ ] `GetBinDataPath(binid: long) -> BSTR`
- [ ] `SetTitleName(Title: BSTR)`
- [ ] `GetHeadingString() -> BSTR`
- [ ] `SetBarCodeImage(lpImagePath: BSTR, pgno: long, index: long, X: long, Y: long, Width: long, Height: long) -> VARIANT_BOOL`
- [ ] `FileTranslate(curLang: BSTR, transLang: BSTR) -> VARIANT_BOOL`
- [ ] `GetTranslateLangList(curLang: BSTR) -> BSTR`
- [ ] `SetDRMAuthority(authority: long) -> VARIANT_BOOL`
- [ ] `ConvertPUAHangulToUnicode(Text: VARIANT) -> long` ※

### 필드 (누름틀) (11)

- [ ] `CreateField(Direction: BSTR, memo: VARIANT, name: VARIANT) -> VARIANT_BOOL`
- [ ] `MoveToField(Field: BSTR, Text: VARIANT, start: VARIANT, select: VARIANT) -> VARIANT_BOOL`
- [ ] `FieldExist(Field: BSTR) -> VARIANT_BOOL`
- [ ] `GetFieldText(Field: BSTR) -> BSTR`
- [ ] `PutFieldText(Field: BSTR, Text: BSTR)`
- [ ] `RenameField(oldname: BSTR, newname: BSTR)`
- [ ] `GetCurFieldName(option: VARIANT) -> BSTR`
- [ ] `SetCurFieldName(Field: BSTR, option: VARIANT, Direction: BSTR, memo: BSTR) -> VARIANT_BOOL`
- [ ] `ModifyFieldProperties(Field: BSTR, remove: long, Add: long) -> long`
- [ ] `SetFieldViewOption(option: long) -> long`
- [ ] `GetFieldList(Number: VARIANT, option: VARIANT) -> BSTR`

### 메타태그 (9)

- [ ] `SetCurMetatagName(tag: BSTR) -> VARIANT_BOOL`
- [ ] `GetCurMetatagName() -> BSTR`
- [ ] `RenameMetatag(oldtag: BSTR, newtag: BSTR)`
- [ ] `MetatagExist(tag: BSTR) -> VARIANT_BOOL`
- [ ] `MoveToMetatag(tag: BSTR, Text: VARIANT, start: VARIANT, select: VARIANT) -> VARIANT_BOOL`
- [ ] `GetMetatagNameText(tag: BSTR) -> BSTR`
- [ ] `PutMetatagNameText(tag: BSTR, Text: BSTR)`
- [ ] `GetMetatagList(Number: VARIANT, option: VARIANT) -> BSTR`
- [ ] `ModifyMetatagProperties(tag: BSTR, remove: long, Add: long) -> long`

### 캐럿 위치 · 선택 (6)

- [ ] `SetPos(List: long, Para: long, pos: long) -> VARIANT_BOOL`
- [ ] `GetPosBySet() -> IDispatch`
- [ ] `SetPosBySet(dispVal: IDispatch*) -> VARIANT_BOOL`
- [ ] `GetSelectedPos(out slist: long*, out spara: long*, out spos: long*, out elist: long*, out epara: long*, out epos: long*) -> VARIANT_BOOL`
- [ ] `GetSelectedPosBySet(sset: IDispatch*, eset: IDispatch*) -> VARIANT_BOOL`
- [ ] `GetMousePos(XRelTo: long, YRelTo: long) -> IDispatch`

### Action · ParameterSet (9)

- [ ] `CreateAction(actidstr: BSTR) -> IDispatch`
- [ ] `CreateSet(setidstr: BSTR) -> IDispatch`
- [ ] `ReleaseAction(action: IDispatch*)`
- [ ] `InitHParameterSet()`
- [ ] `LockCommand(ActID: BSTR, isLock: VARIANT_BOOL)`
- [ ] `IsCommandLock(actionID: BSTR) -> VARIANT_BOOL`
- [ ] `IsActionEnable(actionID: BSTR) -> VARIANT_BOOL`
- [ ] `ReplaceAction(OldActionID: BSTR, NewActionID: BSTR) -> VARIANT_BOOL`
- [ ] `CheckXObject(bstring: BSTR) -> IDispatch` ※

### 컨트롤 · 그림 (6)

- [ ] `InsertCtrl(CtrlID: BSTR, [initparam: VARIANT]) -> IDispatch`
- [ ] `DeleteCtrl(ctrl: IDispatch*) -> VARIANT_BOOL`
- [ ] `FindCtrl() -> BSTR`
- [ ] `UnSelectCtrl()`
- [ ] `InsertPicture(Path: BSTR, Embedded: VARIANT, sizeoption: VARIANT, Reverse: VARIANT, watermark: VARIANT, Effect: VARIANT, Width: VARIANT, Height: VARIANT) -> IDispatch`
- [ ] `InsertBackgroundPicture(BorderType: BSTR, Path: BSTR, Embedded: VARIANT, filloption: VARIANT, watermark: VARIANT, Effect: VARIANT, Brightness: VARIANT, Contrast: VARIANT) -> VARIANT_BOOL`

### 스타일 · 글꼴 (5)

- [ ] `ExportStyle(param: IDispatch*) -> VARIANT_BOOL`
- [ ] `ImportStyle(param: IDispatch*) -> VARIANT_BOOL`
- [ ] `GetFontList([langid: VARIANT]) -> BSTR`
- [ ] `ReplaceFont(langid: long, desFontName: BSTR, desFontType: long, newFontName: BSTR, newFontType: long) -> VARIANT_BOOL`
- [ ] `ScanFont() -> VARIANT_BOOL`

### 스크립트 매크로 · 메시지 박스 · 사용자 정보 (5)

- [ ] `GetScriptSource(filename: BSTR) -> BSTR`
- [ ] `RunScriptMacro(FunctionName: BSTR, uMacroType: long, uScriptType: long) -> VARIANT_BOOL`
- [ ] `GetMessageBoxMode() -> long`
- [ ] `SetUserInfo(userInfoId: long, Value: BSTR)`
- [ ] `GetUserInfo(userInfoId: long) -> BSTR`

### 개인 정보 보호 (4)

- [ ] `SetPrivateInfoPassword(Password: BSTR) -> VARIANT_BOOL`
- [ ] `RegisterPrivateInfoPattern(PrivateType: long, PrivatePattern: BSTR) -> VARIANT_BOOL`
- [ ] `FindPrivateInfo(PrivateType: long, [PrivateString: VARIANT]) -> long`
- [ ] `ProtectPrivateInfo(PotectingChar: BSTR, [PrivatePatternType: VARIANT]) -> VARIANT_BOOL`

### 단위 · 색 · 날짜 변환 (7)

- [ ] `MiliToHwpUnit(mili: double) -> long` ※
- [ ] `PointToHwpUnit(Point: double) -> long`
- [ ] `RGBColor(red: unsigned char, green: unsigned char, blue: unsigned char) -> long`
- [ ] `SolarToLunar(sYear: long, sMonth: long, sDay: long, out lYear: long*, out lMonth: long*, out lDay: long*, out lLeap: VARIANT_BOOL*) -> VARIANT_BOOL`
- [ ] `SolarToLunarBySet(sYear: long, sMonth: long, sDay: long) -> IDispatch`
- [ ] `LunarToSolar(lYear: long, lMonth: long, lDay: long, lLeap: VARIANT_BOOL, out sYear: long*, out sMonth: long*, out sDay: long*) -> VARIANT_BOOL`
- [ ] `LunarToSolarBySet(lYear: long, lMonth: long, lDay: long, lLeap: VARIANT_BOOL) -> IDispatch`

### 문자열 → 열거값 변환 헬퍼 (83)

모두 `Name(문자열: BSTR) -> long` 형태다. ParameterSet 항목에 넣을 정수 값을 SDK 문자열 이름으로 얻는 용도.

- [ ] `HwpLineWidth(LineWidth: BSTR) -> long` ※
- [ ] `HwpLineType(LineType: BSTR) -> long` ※
- [ ] `PicEffect(PicEffect: BSTR) -> long` ※
- [ ] `EndStyle(EndStyle: BSTR) -> long` ※
- [ ] `EndSize(EndSize: BSTR) -> long` ※
- [ ] `HwpOutlineStyle(HwpOutlineStyle: BSTR) -> long` ※
- [ ] `DrawFillImage(fillimage: BSTR) -> long` ※
- [ ] `DrawShadowType(ShadowType: BSTR) -> long` ※
- [ ] `Gradation(Gradation: BSTR) -> long` ※
- [ ] `HatchStyle(HatchStyle: BSTR) -> long` ※
- [ ] `FillAreaType(FillArea: BSTR) -> long` ※
- [ ] `FontType(FontType: BSTR) -> long`
- [ ] `HwpUnderlineType(HwpUnderlineType: BSTR) -> long` ※
- [ ] `HwpOutlineType(HwpOutlineType: BSTR) -> long` ※
- [ ] `CharShadowType(ShadowType: BSTR) -> long` ※
- [ ] `StrikeOut(StrikeOutType: BSTR) -> long` ※
- [ ] `DSMark(DiacSymMark: BSTR) -> long` ※
- [ ] `ParaHeadAlign(ParaHeadAlign: BSTR) -> long` ※
- [ ] `NumberFormat(NumFormat: BSTR) -> long` ※
- [ ] `PlacementType(Restart: BSTR) -> long` ※
- [ ] `GutterMethod(GutterType: BSTR) -> long` ※
- [ ] `TextDir(TextDirection: BSTR) -> long` ※
- [ ] `PageType(PageType: BSTR) -> long` ※
- [ ] `ColDefType(ColDefType: BSTR) -> long` ※
- [ ] `ColLayoutType(ColLayoutType: BSTR) -> long` ※
- [ ] `Handler(Handler: BSTR) -> long` ※
- [ ] `CreateMode(CreationMode: BSTR) -> long` ※
- [ ] `CreateID(CreationID: BSTR) -> long` ※
- [ ] `HAlign(HAlign: BSTR) -> long` ※
- [ ] `LineWrapType(LineWrap: BSTR) -> long` ※
- [ ] `VAlign(VAlign: BSTR) -> long` ※
- [ ] `VertRel(VertRel: BSTR) -> long` ※
- [ ] `HorzRel(HorzRel: BSTR) -> long` ※
- [ ] `WidthRel(WidthRel: BSTR) -> long` ※
- [ ] `HeightRel(HeightRel: BSTR) -> long` ※
- [ ] `TextWrapType(TextWrap: BSTR) -> long` ※
- [ ] `TextFlowType(TextFlow: BSTR) -> long` ※
- [ ] `TextArtAlign(TextArtAlign: BSTR) -> long` ※
- [ ] `SideType(SideType: BSTR) -> long` ※
- [ ] `BorderShape(BorderType: BSTR) -> long` ※
- [ ] `TableBreak(PageBreak: BSTR) -> long` ※
- [ ] `HwpZoomType(ZoomType: BSTR) -> long` ※
- [ ] `PrintRange(PrintRange: BSTR) -> long` ※
- [ ] `PrintType(PrintMethod: BSTR) -> long` ※
- [ ] `PrintDevice(PrintDevice: BSTR) -> long` ※
- [ ] `LineSpacingMethod(LineSpacing: BSTR) -> long` ※
- [ ] `BreakWordLatin(BreakLatinWord: BSTR) -> long` ※
- [ ] `TextAlign(TextAlign: BSTR) -> long` ※
- [ ] `HeadType(HeadingType: BSTR) -> long` ※
- [ ] `AutoNumType(autonum: BSTR) -> long` ※
- [ ] `PageNumPosition(pagenumpos: BSTR) -> long` ※
- [ ] `GridMethod(GridMethod: BSTR) -> long` ※
- [ ] `GridViewLine(GridViewLine: BSTR) -> long` ※
- [ ] `DrawAspect(DrawAspect: BSTR) -> long` ※
- [ ] `CellApply(CellApply: BSTR) -> long` ※
- [ ] `SubtPos(SubtPos: BSTR) -> long` ※
- [ ] `ImageFormat(ImageFormat: BSTR) -> long` ※
- [ ] `MailType(MailType: BSTR) -> long` ※
- [ ] `Delimiter(Delimiter: BSTR) -> long` ※
- [ ] `DbfCodeType(DbfCode: BSTR) -> long` ※
- [ ] `FindDir(FindDir: BSTR) -> long` ※
- [ ] `TableSwapType(tableswap: BSTR) -> long` ※
- [ ] `SortDelimiter(SortDelimiter: BSTR) -> long` ※
- [ ] `Revision(Revision: BSTR) -> long` ※
- [ ] `Canonical(Canonical: BSTR) -> long` ※
- [ ] `Signature(Signature: BSTR) -> long` ※
- [ ] `Encrypt(Encrypt: BSTR) -> long` ※
- [ ] `Hash(Hash: BSTR) -> long` ※
- [ ] `ArcType(ArcType: BSTR) -> long` ※
- [ ] `Numbering(Numbering: BSTR) -> long` ※
- [ ] `PrintPaper(PrintPaper: BSTR) -> long` ※
- [ ] `BrushType(BrushType: BSTR) -> long` ※
- [ ] `Hiding(Hiding: BSTR) -> long` ※
- [ ] `MacroState(MacroState: BSTR) -> long` ※
- [ ] `TableFormat(TableFormat: BSTR) -> long` ※
- [ ] `TableTarget(TableTarget: BSTR) -> long` ※
- [ ] `ViewFlag(ViewFlag: BSTR) -> long` ※
- [ ] `Slash(Slash: BSTR) -> long` ※
- [ ] `CrookedSlash(CrookedSlash: BSTR) -> long` ※
- [ ] `WatermarkBrush(WatermarkBrush: BSTR) -> long` ※
- [ ] `StyleType(StyleType: BSTR) -> long` ※
- [ ] `PresentEffect(prsnteffect: BSTR) -> long` ※
- [ ] `HwpUnderlineShape(HwpUnderlineShape: BSTR) -> long` ※

## `IDHwpAction` — 미구현 5 / 5

대응하는 Rust 타입 자체가 없다.

- [ ] `GetDefault(param: IDispatch*) -> long`
- [ ] `CreateSet() -> IDispatch`
- [ ] `Execute(param: IDispatch*) -> VARIANT_BOOL`
- [ ] `PopupDialog(param: IDispatch*) -> long`
- [ ] `Run() -> long`

## `IDHwpParameterSet` — 미구현 11 / 11

대응하는 Rust 타입 자체가 없다.

- [ ] `Clone() -> IDispatch`
- [ ] `CreateItemArray(itemid: BSTR, Count: long) -> IDispatch`
- [ ] `CreateItemSet(itemid: BSTR, SetID: BSTR) -> IDispatch`
- [ ] `GetIntersection(srcset: IDispatch*)`
- [ ] `IsEquivalent(srcset: IDispatch*) -> VARIANT_BOOL`
- [ ] `Item(itemid: BSTR) -> VARIANT`
- [ ] `ItemExist(itemid: BSTR) -> VARIANT_BOOL`
- [ ] `Merge(srcset: IDispatch*) -> VARIANT_BOOL`
- [ ] `RemoveAll(SetID: BSTR)`
- [ ] `RemoveItem(itemid: BSTR)`
- [ ] `SetItem(itemid: BSTR, newVal: VARIANT)`

## `IDHwpParameterArray` — 미구현 4 / 4

대응하는 Rust 타입 자체가 없다.

- [ ] `Clone() -> IDispatch`
- [ ] `Copy(srcarray: IDispatch*)`
- [ ] `Item(index: long) -> VARIANT`
- [ ] `SetItem(index: long, newVal: VARIANT)`

## `HAction` — 미구현 0 / 4

binding 있음: `GetDefault`, `Execute`, `PopupDialog`, `Run`

미구현 method 없음.

## `HSet` — 미구현 1 / 1

전용 Rust 타입 없이 `DispObj`로만 반환된다(`HCharShape::h_set` 등).

- [ ] `SetItem(itemid: BSTR, newVal: VARIANT)`

## `IDHwpCtrlCode` — 미구현 1 / 1

대응하는 Rust 타입 자체가 없다.

- [ ] `GetAnchorPos(type: long) -> IDispatch`

## `IXHwpDocuments` — 미구현 3 / 5

binding 있음: `Add`, `FindItem`

- [ ] `Close(isDirty: VARIANT_BOOL)`
- [ ] `_AddDocument(frame: VARIANT, docment: VARIANT)` ※
- [ ] `_RemoveDocument(docment: VARIANT)` ※

## `IXHwpDocument` — 미구현 8 / 9

binding 있음: `SetActive_XHwpDocument`

- [ ] `Close(isDirty: VARIANT_BOOL) -> VARIANT_BOOL`
- [ ] `Save(save_if_dirty: VARIANT) -> VARIANT_BOOL`
- [ ] `SaveAs(Path: BSTR, Format: VARIANT, arg: VARIANT) -> VARIANT_BOOL`
- [ ] `Undo(Count: long) -> VARIANT_BOOL`
- [ ] `Redo(Count: long) -> VARIANT_BOOL`
- [ ] `Open(filename: BSTR, Format: VARIANT, arg: VARIANT) -> VARIANT_BOOL`
- [ ] `SendBrowser() -> VARIANT_BOOL`
- [ ] `Clear(option: VARIANT)`

## `IXHwpWindows` — 미구현 2 / 2

- [ ] `Add() -> IDispatch`
- [ ] `Close(isDirty: VARIANT_BOOL) -> VARIANT_BOOL`

## `IXHwpWindow` — 미구현 2 / 2

- [ ] `Close(isDirty: VARIANT_BOOL)`
- [ ] `ChangeScreenMode(bFullMode: VARIANT_BOOL)` ※

## `IXHwpToolbarLayout` — 미구현 1 / 10

binding 있음: `CreateToolbar`, `DeleteToolbar`, `GetToolbar`, `ShowToolbar`, `CreateToolbarButton`, `CreateMenuButton`, `ChangeSerializePath`, `IsNewSerializePath`, `GetToolBoxToolBar`

- [ ] `CreateToolbarButtonEx(name: BSTR, aid: BSTR, type: long, style: long) -> IDispatch` ※

## `IXHwpToolbar` — 미구현 1 / 4

binding 있음: `InsertToolbarButton`, `DeleteToolbarButton`, `GetToolbarButton`

- [ ] `InsertToolbarButtonEx(ToolBarButton: IDispatch*, pos: long, type: long)` ※

## `IXHwpToolbarButton` — 미구현 7 / 7

Rust 타입 `XHwpToolbarButton`은 선언돼 있지만 method가 하나도 없다. `XHwpToolbarMenuButton`도 IDL상 이 interface를 상속하므로 아래 method가 함께 빠져 있다.

- [ ] `GetHandle() -> VARIANT` ※
- [ ] `GetActionName() -> BSTR` ※
- [ ] `SetText(name: BSTR)` ※
- [ ] `GetText() -> BSTR`
- [ ] `GetStyle() -> long` ※
- [ ] `ModifyStyle(remove: long, Add: long)` ※
- [ ] `GetType() -> long` ※

## `IXHwpToolbarComboBox` — 미구현 16 / 16

대응하는 Rust 타입 자체가 없다.

- [ ] `GetHandle() -> VARIANT` ※
- [ ] `GetActionName() -> BSTR` ※
- [ ] `AddString(strAddString: BSTR)` ※
- [ ] `SelectString(lStartAfter: long, strSelectString: BSTR)` ※
- [ ] `FindString(lStartAfter: long, strFindString: BSTR) -> long` ※
- [ ] `GetLBText(lIndex: long) -> BSTR` ※
- [ ] `InsertItem(lIndex: long, strInsertString: BSTR)` ※
- [ ] `DeleteItem(lIndex: long)` ※
- [ ] `ResetContent()`
- [ ] `ShowDropDown()` ※
- [ ] `GetStyle() -> long` ※
- [ ] `ModifyStyle(remove: long, Add: long)` ※
- [ ] `LimitText(lLimitChars: long)` ※
- [ ] `GetType() -> long` ※
- [ ] `GetWindowText() -> BSTR` ※
- [ ] `SetWindowText(strText: BSTR)` ※

## `IXHwpToolbarMenuButton` — 미구현 2 / 3

binding 있음: `InsertMenuButton`

- [ ] `DeleteMenuButton(pos: long) -> VARIANT_BOOL` ※
- [ ] `GetMenuButton(pos: long) -> IDispatch` ※

## `IXHwpTabs` — 미구현 2 / 2

대응하는 Rust 타입 자체가 없다.

- [ ] `Add() -> IDispatch`
- [ ] `Close(isDirty: VARIANT_BOOL) -> VARIANT_BOOL`

## `IXHwpTab` — 미구현 1 / 1

대응하는 Rust 타입 자체가 없다.

- [ ] `Close(isDirty: VARIANT_BOOL)`

## `IXHwpPrint` — 미구현 5 / 5

대응하는 Rust 타입 자체가 없다.

- [ ] `Run() -> VARIANT_BOOL`
- [ ] `RunToFax() -> VARIANT_BOOL` ※
- [ ] `RunToImage(Format: unsigned int, filename: BSTR, colordepth: unsigned int, resolution: unsigned long) -> VARIANT_BOOL` ※
- [ ] `RunToPDF() -> VARIANT_BOOL` ※
- [ ] `RunFilePreview() -> VARIANT_BOOL` ※

## `IXHwpRange` — 미구현 36 / 36

대응하는 Rust 타입 자체가 없다.

- [ ] `MovePrevChar() -> VARIANT_BOOL` ※
- [ ] `MoveNextChar() -> VARIANT_BOOL` ※
- [ ] `MovePrevPos() -> VARIANT_BOOL` ※
- [ ] `MoveNextPos() -> VARIANT_BOOL` ※
- [ ] `MovePrevPosEx() -> VARIANT_BOOL` ※
- [ ] `MoveNextPosEx() -> VARIANT_BOOL` ※
- [ ] `MovePrevWord() -> VARIANT_BOOL` ※
- [ ] `MoveNextWord() -> VARIANT_BOOL` ※
- [ ] `MoveWordBegin() -> VARIANT_BOOL` ※
- [ ] `MoveWordEnd() -> VARIANT_BOOL` ※
- [ ] `MoveDocumentBegin() -> VARIANT_BOOL` ※
- [ ] `MoveDocumentEnd() -> VARIANT_BOOL` ※
- [ ] `MoveLineBegin() -> VARIANT_BOOL` ※
- [ ] `MoveLineEnd() -> VARIANT_BOOL` ※
- [ ] `MoveParagraphBegin() -> VARIANT_BOOL` ※
- [ ] `MoveParagraphEnd() -> VARIANT_BOOL` ※
- [ ] `MovePageUp() -> VARIANT_BOOL` ※
- [ ] `MovePageDown() -> VARIANT_BOOL` ※
- [ ] `MoveListBegin() -> VARIANT_BOOL` ※
- [ ] `MoveListEnd() -> VARIANT_BOOL` ※
- [ ] `MoveTopLevelBegin() -> VARIANT_BOOL` ※
- [ ] `MoveTopLevelEnd() -> VARIANT_BOOL` ※
- [ ] `MoveNextParagraphBegin() -> VARIANT_BOOL` ※
- [ ] `MovePrevParagraphEnd() -> VARIANT_BOOL` ※
- [ ] `MovePrevParagraphBegin() -> VARIANT_BOOL` ※
- [ ] `MoveSectionUp() -> VARIANT_BOOL` ※
- [ ] `MoveSectionDown() -> VARIANT_BOOL` ※
- [ ] `MoveParentList() -> VARIANT_BOOL` ※
- [ ] `MoveTopLevelList() -> VARIANT_BOOL` ※
- [ ] `MoveRootList() -> VARIANT_BOOL` ※
- [ ] `Close() -> VARIANT_BOOL`
- [ ] `CloseEx() -> VARIANT_BOOL` ※
- [ ] `Cut() -> VARIANT_BOOL`
- [ ] `Copy() -> VARIANT_BOOL`
- [ ] `Erase() -> VARIANT_BOOL` ※
- [ ] `Paste() -> VARIANT_BOOL` ※

## `IXHwpFind` — 미구현 6 / 6

대응하는 Rust 타입 자체가 없다.

- [ ] `RunRepeat() -> VARIANT_BOOL` ※
- [ ] `RunReverse() -> VARIANT_BOOL` ※
- [ ] `RunForward() -> VARIANT_BOOL` ※
- [ ] `RunBackward() -> VARIANT_BOOL` ※
- [ ] `RunReplace() -> VARIANT_BOOL` ※
- [ ] `RunAllReplace() -> VARIANT_BOOL` ※

## `IXHwpSelection` — 미구현 23 / 23

대응하는 Rust 타입 자체가 없다.

- [ ] `MovePrevChar() -> VARIANT_BOOL` ※
- [ ] `MoveNextChar() -> VARIANT_BOOL` ※
- [ ] `MovePrevPos() -> VARIANT_BOOL` ※
- [ ] `MoveNextPos() -> VARIANT_BOOL` ※
- [ ] `MovePrevWord() -> VARIANT_BOOL` ※
- [ ] `MoveNextWord() -> VARIANT_BOOL` ※
- [ ] `MoveWordBegin() -> VARIANT_BOOL` ※
- [ ] `MoveWordEnd() -> VARIANT_BOOL` ※
- [ ] `MoveDocumentBegin() -> VARIANT_BOOL` ※
- [ ] `MoveDocumentEnd() -> VARIANT_BOOL` ※
- [ ] `MoveLineBegin() -> VARIANT_BOOL` ※
- [ ] `MoveLineEnd() -> VARIANT_BOOL` ※
- [ ] `MoveParagraphBegin() -> VARIANT_BOOL` ※
- [ ] `MoveParagraphEnd() -> VARIANT_BOOL` ※
- [ ] `MovePageUp() -> VARIANT_BOOL` ※
- [ ] `MovePageDown() -> VARIANT_BOOL` ※
- [ ] `MoveListBegin() -> VARIANT_BOOL` ※
- [ ] `MoveListEnd() -> VARIANT_BOOL` ※
- [ ] `MoveTopLevelBegin() -> VARIANT_BOOL` ※
- [ ] `MoveTopLevelEnd() -> VARIANT_BOOL` ※
- [ ] `MovePrevParagraphEnd() -> VARIANT_BOOL` ※
- [ ] `MoveNextParagraphBegin() -> VARIANT_BOOL` ※
- [ ] `MovePrevParagraphBegin() -> VARIANT_BOOL` ※

## `IXHwpFormComboBox` — 미구현 4 / 4

대응하는 Rust 타입 자체가 없다.

- [ ] `InsertString(newVal: BSTR, index: long)`
- [ ] `DeleteString(newVal: unsigned long)`
- [ ] `FindStringExact(newVal: long, string: BSTR) -> long`
- [ ] `ResetContent()`

## `IXHwpCharacterShape` — 미구현 10 / 10

대응하는 Rust 타입 자체가 없다.

- [ ] `ShapeNormal()` ※
- [ ] `HeightIncrease()` ※
- [ ] `HeightDecrease()` ※
- [ ] `WidthIncrease()` ※
- [ ] `WidthDecrease()` ※
- [ ] `SpacingIncrease()` ※
- [ ] `SpacingDecrease()` ※
- [ ] `NextFaceName()` ※
- [ ] `PrevFaceName()` ※
- [ ] `SuperSubScript()` ※

## `IXHwpParagraphShape` — 미구현 19 / 19

대응하는 Rust 타입 자체가 없다.

- [ ] `AlignJustify()` ※
- [ ] `AlignLeft()` ※
- [ ] `AlignCenter()` ※
- [ ] `AlignRight()` ※
- [ ] `AlignFit()` ※
- [ ] `AlignDiv()` ※
- [ ] `WithNext()` ※
- [ ] `Protect()` ※
- [ ] `IncreaseLeftMargin()` ※
- [ ] `DecreaseLeftMargin()` ※
- [ ] `IncreaseRightMargin()` ※
- [ ] `DecreaseRightMargin()` ※
- [ ] `IndentAtCaret()` ※
- [ ] `IndentNegative()` ※
- [ ] `IndentPositive()` ※
- [ ] `DecreaseLineSpacing()` ※
- [ ] `IncreaseLineSpacing()` ※
- [ ] `IncreaseMargin()` ※
- [ ] `DecreaseMargin()` ※

## `IXHwpMessageBox` — 미구현 1 / 1

대응하는 Rust 타입 자체가 없다.

- [ ] `DoModal()`

## `IXHwpSendMail` — 미구현 1 / 1

대응하는 Rust 타입 자체가 없다.

- [ ] `SendMail()` ※

## `IHwpObjectEvents` (이벤트 sink) — 미구현 16 / 16

호출하는 method가 아니라 한글이 호출해 주는 콜백이다. `IConnectionPoint::Advise` 기반 sink 구현이 코드에 없다.
SDK 참고: `한글오토메이션EventHandler추가_2504.pdf`.

- [ ] `Quit()`
- [ ] `CreateXHwpWindow()`
- [ ] `CloseXHwpWindow()`
- [ ] `NewDocument(newVal: long)`
- [ ] `DocumentBeforeClose(newVal: long)`
- [ ] `DocumentBeforeOpen(newVal: long)`
- [ ] `DocumentAfterOpen(newVal: long)`
- [ ] `DocumentBeforeSave(newVal: long)`
- [ ] `DocumentAfterSave(newVal: long)`
- [ ] `DocumentAfterClose(newVal: long)`
- [ ] `DocumentChange(newVal: long)`
- [ ] `DocumentBeforePrint(newVal: long)`
- [ ] `DocumentAfterPrint(newVal: long)`
- [ ] `DocumentClickedHyperlink(linkType: long, szAddress: BSTR, szShowText: BSTR) -> VARIANT_BOOL`
- [ ] `DocumentModifiedHyperlink(linkType: long, szAddress: BSTR, szShowText: BSTR) -> VARIANT_BOOL`
- [ ] `BeforeQuit(newVal: long)`
