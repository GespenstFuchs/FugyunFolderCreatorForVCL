unit TMemoHistoryUnit;

{$MODE Delphi}

interface

uses
  Vcl.StdCtrls, Winapi.Windows, Winapi.Messages, Generics.Collections;

type
  /// <summary>
  /// メモ履歴情報クラス
  /// </summary>
  MemoHistoryInfo = class
  private
    SetText: string;
    SetCaretPosition: Integer;
    SetVScrollInfo: TScrollInfo;
    SetHScrollInfo: TScrollInfo;
    constructor Create(Text: String; CaretPosition: Integer; VScrollInfo, HScrollInfo: TScrollInfo);
  end;

  /// <summary>
  /// メモ履歴クラス
  /// </summary>
  MemoHistory = class
  class var
    /// <summary>
    /// 対象メモ
    /// </summary>
    TargetMemo: TMemo;

    /// <summary>
    /// 最大履歴数
    /// </summary>
    MaxHistoryCount: Integer;

    /// <summary>
    /// 履歴情報リスト
    /// </summary>
    HistoryInfoList: TList<MemoHistoryInfo>;

    /// <summary>
    /// カレント履歴位置
    /// （０ＢＡＳＥ）
    /// </summary>
    CurrentHistoryIndex: Integer;

    /// <summary>
    /// 履歴追加フラグ
    /// （True；追加可能・False：追加不可）
    /// </summary>
    HistoryAddFlg: Boolean;

  private
    function GetTMemoScrollInfo(fnBar: Integer): TScrollInfo;
    procedure SetTMemoScrollInfo(HistoryInfo: MemoHistoryInfo);

  public
    constructor Create(TMemo: TMemo; MaxHistoryCount: Integer = 1000);
    procedure HistoryAdd;
    procedure Undo;
    procedure Redo;
    function IsUndoFlg: Boolean;
    function IsRedoFlg: Boolean;
    destructor Destroy; override;
  end;

implementation

/// <summary>
/// コンストラクタ
/// </summary>
/// <param name="Text">テキスト</param>
/// <param name="CaretPosition">カーソル位置</param>
/// <param name="VScrollInfo">垂直スクロール位置</param>
/// <param name="HScrollInfo">水平スクロール位置</param>
constructor MemoHistoryInfo.Create(Text: String; CaretPosition: Integer; VScrollInfo, HScrollInfo: TScrollInfo);
begin
  SetText := Text;
  SetCaretPosition := CaretPosition;
  SetVScrollInfo := VScrollInfo;
  SetHScrollInfo := HScrollInfo;
end;

/// <summary>
/// コンストラクタ
/// </summary>
/// <param name="TMemo">メモ（初期設定済み）</param>
/// <param name="MaxHistoryCount">履歴数（デフォルト：１０００）</param>
constructor MemoHistory.Create(TMemo: TMemo; MaxHistoryCount: Integer = 1000);
var
  NewHistoryInfo: MemoHistoryInfo;
begin
  TargetMemo := TMemo;
  MemoHistory.MaxHistoryCount := MaxHistoryCount;
  HistoryInfoList := TList<MemoHistoryInfo>.Create;
  HistoryAddFlg := True;

  // 新しい履歴情報を作成して追加する。
  NewHistoryInfo := MemoHistoryInfo.Create(
    TargetMemo.Text,
    TargetMemo.SelStart,
    GetTMemoScrollInfo(SB_VERT),
    GetTMemoScrollInfo(SB_HORZ));
  HistoryInfoList.Add(NewHistoryInfo);

  // カレント履歴位置を設定する。
  CurrentHistoryIndex := HistoryInfoList.Count - 1;
end;

/// <summary>
/// 履歴追加処理
/// </summary>
procedure MemoHistory.HistoryAdd;
var
  NewHistoryInfo: MemoHistoryInfo;
begin
  // 履歴追加フラグを判定する。
  if HistoryAddFlg then
  begin
    // 追加可能の場合

    // カレント位置を判定する。
    if HistoryInfoList.Count <> CurrentHistoryIndex + 1 then
    begin
      // 最新以外の場合

      // 履歴情報リストを更新する。
      while HistoryInfoList.Count > CurrentHistoryIndex + 1 do
      begin
        HistoryInfoList.Delete(HistoryInfoList.Count - 1);
      end;
    end;

    // 情報リストの項目数と、最大履歴数を判定する。
    if HistoryInfoList.Count = MaxHistoryCount then
    begin
      // 一致する場合、最初の要素を削除する。
      HistoryInfoList.Delete(0);
    end;

    // 新しい履歴情報を作成して追加する。
    NewHistoryInfo := MemoHistoryInfo.Create(
      TargetMemo.Text,
      TargetMemo.SelStart,
      GetTMemoScrollInfo(SB_VERT),
      GetTMemoScrollInfo(SB_HORZ));
    HistoryInfoList.Add(NewHistoryInfo);

    // カレント履歴位置を設定する。
    CurrentHistoryIndex := HistoryInfoList.Count - 1;
  end
  else
  begin
    // 追加不可の場合
    HistoryAddFlg := True;
  end;
end;

/// <summary>
/// 元に戻す処理
/// </summary>
procedure MemoHistory.Undo;
var
  TargetIndex: Integer;
  HistoryInfo: MemoHistoryInfo;
begin
  // 使用する履歴情報位置を判定する。
  TargetIndex := CurrentHistoryIndex - 1;
  if 0 <= TargetIndex then
  begin
    // ０以上の場合

    // 描画を停止する。
    TargetMemo.Lines.BeginUpdate;

    // 履歴追加フラグを設定する。
    HistoryAddFlg := False;

    TargetMemo.SetFocus;

    HistoryInfo := HistoryInfoList[TargetIndex];
    TargetMemo.Text := HistoryInfo.SetText;
    TargetMemo.SelStart := HistoryInfo.SetCaretPosition;
    SetTMemoScrollInfo(HistoryInfo);

    // デクリメントした値を、カレント履歴位置として設定する。
    CurrentHistoryIndex := TargetIndex;

    // 描画を行う。
    TargetMemo.Lines.EndUpdate;
  end;
end;

/// <summary>
/// やり直し処理
/// </summary>
procedure MemoHistory.Redo;
var
  TargetIndex: Integer;
  HistoryInfo: MemoHistoryInfo;
begin
  TargetIndex := CurrentHistoryIndex + 1;

  // 項目数・対象位置を判定する。
  if HistoryInfoList.Count > TargetIndex then
  begin
    // 項目数の方が多い場合

    // 描画を停止する。
    TargetMemo.Lines.BeginUpdate;

    // 履歴情報を取得・保持する。
    HistoryInfo := HistoryInfoList[TargetIndex];

    // 履歴追加フラグを設定する。
    HistoryAddFlg := False;

    TargetMemo.SetFocus;

    TargetMemo.Text := HistoryInfo.SetText;
    TargetMemo.SelStart := HistoryInfo.SetCaretPosition;
    SetTMemoScrollInfo(HistoryInfo);

    // カレント履歴位置を設定する。
    CurrentHistoryIndex := TargetIndex;

    // 描画を行う。
    TargetMemo.Lines.EndUpdate;
  end;
end;

/// <summary>
/// 元に戻す判定処理
/// </summary>
/// <returns>チェック結果（True：可能・False：不可）</returns>
function MemoHistory.IsUndoFlg: Boolean;
begin
  Result := CurrentHistoryIndex > 0;
end;

/// <summary>
/// やり直し判定判定処理
/// </summary>
/// <returns>チェック結果（True：可能・False：不可）</returns>
function MemoHistory.IsRedoFlg: Boolean;
var
  TargetIndex: Integer;
  HistoryInfo: MemoHistoryInfo;
begin
  TargetIndex := CurrentHistoryIndex + 1;

  // 履歴情報リストから履歴情報が取得可能か判定する。
  if  (0 <= TargetIndex) and (TargetIndex < HistoryInfoList.Count) then
  begin
    // 取得可能の場合
    HistoryInfo := HistoryInfoList[TargetIndex];
  end
  else
  begin
    // 取得不可の場合
    HistoryInfo := nil;
  end;

  Result :=  HistoryInfo <> nil;
end;

/// <summary>
/// デストラクタ
/// </summary>
destructor MemoHistory.Destroy;
begin
  HistoryInfoList.Free;
  inherited Destroy;
end;

{$region 'プライベートメソッド'}

/// <summary>
/// メモ・スクロール情報取得処理
/// </summary>
/// <param name="fnBar">対象スクロールバー</param>
/// <returns>テキストボックス・スクロール情報</returns>
function MemoHistory.GetTMemoScrollInfo(fnBar: Integer): TScrollInfo;
var
  ScrollInfo: TScrollInfo;
begin
  // ScrollInfo構造体を初期化する。
  ScrollInfo.cbSize := SizeOf(TScrollInfo);
  ScrollInfo.fMask := SIF_POS;
  ScrollInfo.nMin := 0;
  ScrollInfo.nMax := 0;
  ScrollInfo.nPage := 0;
  ScrollInfo.nPos := 0;
  ScrollInfo.nTrackPos := 0;

  // スクロール情報の取得し、返却する。
  GetScrollInfo(MemoHistory.TargetMemo.Handle, fnBar, ScrollInfo);
  Result := ScrollInfo;
end;

/// <summary>
/// メモ・スクロール情報設定処理
/// </summary>
/// <param name="HistoryInfo">履歴情報</param>
procedure MemoHistory.SetTMemoScrollInfo(HistoryInfo: MemoHistoryInfo);
var
  VScrollInfo, HScrollInfo: TScrollInfo;
begin
  VScrollInfo := HistoryInfo.SetVScrollInfo;
  HScrollInfo := HistoryInfo.SetHScrollInfo;

  // スクロール位置を設定する。
  SetScrollInfo(MemoHistory.TargetMemo.Handle, SB_VERT, VScrollInfo, True);
  SetScrollInfo(MemoHistory.TargetMemo.Handle, SB_HORZ, HScrollInfo, True);

  // スクロール位置を反映させる。
  SendMessage(MemoHistory.TargetMemo.Handle, WM_VSCROLL, MakeWParam(SB_THUMBPOSITION, VScrollInfo.nPos), 0);
  SendMessage(MemoHistory.TargetMemo.Handle, WM_HSCROLL, MakeWParam(SB_THUMBPOSITION, HScrollInfo.nPos), 0);
end;

{$endregion}

end.

