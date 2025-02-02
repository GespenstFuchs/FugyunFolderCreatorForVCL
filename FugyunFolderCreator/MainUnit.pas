unit MainUnit;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Classes, System.Actions,
  Vcl.Forms, Vcl.Controls, Vcl.StdCtrls, Vcl.Buttons, Vcl.ActnList,
  Vcl.PlatformDefaultStyleActnCtrls, Vcl.ActnMan, Vcl.Menus,
  ShellAPI, System.IOUtils, Generics.Collections, System.StrUtils, System.Character,
  TMemoHistoryUnit;

type
  TMainForm = class(TForm)
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    CreateBitBtn: TBitBtn;
    DeleteBitBtn: TBitBtn;
    AllDeleteBitBtn: TBitBtn;
    ClearBitBtn: TBitBtn;
    PathMemo: TMemo;
    ActionManager1: TActionManager;
    CreateAction: TAction;
    DeleteAction: TAction;
    AllDeleteAction: TAction;
    ClearAction: TAction;
    PopupMenu1: TPopupMenu;
    NUndo: TMenuItem;
    NCut: TMenuItem;
    NCopy: TMenuItem;
    N5: TMenuItem;
    NPaste: TMenuItem;
    NDelete: TMenuItem;
    NSelectRow: TMenuItem;
    N9: TMenuItem;
    NSelectAll: TMenuItem;
    N11: TMenuItem;
    NSelectRowFolderOpen: TMenuItem;
    NRedo: TMenuItem;
    procedure FormCreate(Sender: TObject);
    procedure CreateBitBtnClick(Sender: TObject);
    procedure DeleteBitBtnClick(Sender: TObject);
    procedure AllDeleteBitBtnClick(Sender: TObject);
    procedure ClearBitBtnClick(Sender: TObject);
    procedure ClearActionExecute(Sender: TObject);
    procedure CreateActionExecute(Sender: TObject);
    procedure DeleteActionExecute(Sender: TObject);
    procedure AllDeleteActionExecute(Sender: TObject);
    procedure NUndoClick(Sender: TObject);
    procedure NCutClick(Sender: TObject);
    procedure NCopyClick(Sender: TObject);
    procedure NPasteClick(Sender: TObject);
    procedure NDeleteClick(Sender: TObject);
    procedure NSelectRowClick(Sender: TObject);
    procedure NSelectAllClick(Sender: TObject);
    procedure NSelectRowFolderOpenClick(Sender: TObject);
    procedure NRedoClick(Sender: TObject);
    procedure PathMemoChange(Sender: TObject);
    procedure PopupMenu1Popup(Sender: TObject);
  private
    procedure FolderCreateTran(PathMemo: TMemo);
    procedure FolderDeleteTran(PathMemo: TMemo; AllDeleteFlg: Boolean);
    function CheckTran(
      Path: string;
      DriveLetterList: TList<string>;
      RowIndex: Integer): Boolean;
    function IsSurrogateOrCombiningChar(TargetPath: string): Boolean;
    function GetNormalizedStringLength(const TargetStr: string): Integer;
    function IsExistsDirectory(TargetPath: string): Boolean;
    function GetDriveLetterList: TList<string>;
    function ConvertNumberWide(NarrowNumber: Integer): string;
  protected
    procedure FilesDropped(var Msg1: TWMDropFiles); message WM_DROPFILES;
  public
  end;

var
  MainForm: TMainForm;
  NarrowNumberDict: TDictionary<string, string>;
  History: MemoHistory;

const

{$region '記号'}

  /// <summary>
  /// 記号：\
  /// </summary>
  Yen: string = '\';

  /// <summary>
  /// 記号：\\
  /// </summary>
  const DoubleYen: string = '\\';

{$endregion}

{$region '配列'}

  /// <summary>
  /// 使用不可文字配列
  /// </summary>
  InvalidCharAr: array[0..7] of string = ('/', ':', '*', '?', '"', '<', '>', '|');

{$endregion}

{$region 'エラーメッセージ'}

  /// <summary>
  /// エラーメッセージ：タイトル
  /// </summary>
  ErTitle: PWideChar = 'エラー';

{$endregion}

implementation

{$R *.dfm}

/// <summary>
/// 初期表示処理
/// </summary>
/// <param name="Sender">オブジェクト</param>
procedure TMainForm.FormCreate(Sender: TObject);
begin
  // 各ボタン名を設定する。
  CreateBitBtn.Caption := 'フォルダ作成'#13#10'（Ｆ１）';
  DeleteBitBtn.Caption := 'フォルダ削除'#13#10'（Ｆ２）';
  AllDeleteBitBtn.Caption := 'フォルダ全削除'#13#10'（Ｆ３）';
  ClearBitBtn.Caption := '入力したパスをクリア'#13#10'（Ｆ４）';

  // コンテキストメニューのフォントを設定する。
  Screen.MenuFont.Name := 'Yu Gothic UI';
  Screen.MenuFont.Size := 11;

  // メモ履歴クラスのインスタンスを生成する。
  History := MemoHistory.Create(PathMemo);

  // ドロップを許可する。
  DragAcceptFiles(Self.Handle, True);

  // 半角数値連想配列を初期化する。
  NarrowNumberDict := TDictionary<string, string>.Create;
  NarrowNumberDict.Add('0', '０');
  NarrowNumberDict.Add('1', '１');
  NarrowNumberDict.Add('2', '２');
  NarrowNumberDict.Add('3', '３');
  NarrowNumberDict.Add('4', '４');
  NarrowNumberDict.Add('5', '５');
  NarrowNumberDict.Add('6', '６');
  NarrowNumberDict.Add('7', '７');
  NarrowNumberDict.Add('8', '８');
  NarrowNumberDict.Add('9', '９');
end;

/// <summary>
/// フォルダ作成ボタン押下処理
/// </summary>
/// <param name="Sender">オブジェクト</param>
procedure TMainForm.CreateBitBtnClick(Sender: TObject);
begin
  FolderCreateTran(PathMemo);
end;

/// <summary>
/// フォルダ削除ボタン押下処理
/// </summary>
/// <param name="Sender">オブジェクト</param>
procedure TMainForm.DeleteBitBtnClick(Sender: TObject);
begin
  FolderDeleteTran(PathMemo, False);
end;

/// <summary>
/// フォルダ全削除ボタン押下処理
/// </summary>
/// <param name="Sender">オブジェクト</param>
procedure TMainForm.AllDeleteBitBtnClick(Sender: TObject);
begin
  FolderDeleteTran(PathMemo, True);
end;

/// <summary>
/// 入力したパスをクリアボタン押下処理
/// </summary>
/// <param name="Sender">オブジェクト</param>
procedure TMainForm.ClearBitBtnClick(Sender: TObject);
begin
  PathMemo.Clear;
  PathMemo.SetFocus;
end;

/// <summary>
/// パスメモ・テキスト変更処理
/// </summary>
/// <param name="Sender">オブジェクト</param>
procedure TMainForm.PathMemoChange(Sender: TObject);
begin
  // 標準の履歴情報を削除する。
  PathMemo.ClearUndo;

  // 履歴情報を追加する。
  History.HistoryAdd;
end;

/// <summary>
/// ドロップ処理
/// </summary>
/// <param name="MSG1">ドロップファイル</param>
procedure TMainForm.FilesDropped(var MSG1: TWMDropFiles);
var
  FNameSize, FileCount: UINT;
  FileName: array[0..255] of Char;
  FileList: TStringList;
  Index: integer;
  SetText: string;
begin
  FileList := TStringList.Create;

  try
    // ドロップされた全ファイルに処理を行う。
    FileCount := DragQueryFile(Msg1.Drop, $FFFFFFFF, nil, 0);
    for Index := 0 to FileCount - 1 do
    begin
      // ファイル名の長さを取得・保持する。
      FNameSize := DragQueryFile(Msg1.Drop, 0, nil, 0) + 1;
      // ファイル名を取得する。
      DragQueryFile(Msg1.Drop, Index, FileName, FNameSize);
      // フォルダのみ限定するため、フォルダ存在チェック処理を行う。
      if IsExistsDirectory(FileName) = True then
      begin
        // ファイル名を追加する。
        FileList.Add(FileName);
      end;
    end;

    SetText := '';

    // リストの項目の有無を判定する。
    if FileList.Count > 0 then
    begin
      // 挿入開始位置を指定するため、カーソル位置を末尾に移動する。
      PathMemo.SelStart := Length(PathMemo.Text);

      // 全リスト項目に処理を行う。
      for Index := 0 to FileList.Count - 1 do
      begin
        if SetText = '' then
        begin
          SetText := FileList[Index];
        end
        else
        begin
          SetText := SetText + #13#10 + FileList[Index];
        end;
      end;

      // パスメモの入力の有無を判定する。
      if PathMemo.Lines.Text = '' then
      begin
        // 未入力の場合
        PathMemo.Perform(EM_REPLACESEL, 1, LPARAM(PChar(SetText)));
      end
      else
      begin
        // 入力されている場合
        PathMemo.Perform(EM_REPLACESEL, 1, LPARAM(PChar(#13#10 + SetText)));
      end;
    end;
  finally
    // メモリを解放する。
    FileList.Free;
    DragFinish(Msg1.Drop);
  end;
end;

{$region 'アクション'}

/// <summary>
/// フォルダ作成アクション処理
/// </summary>
/// <param name="Sender">オブジェクト</param>
procedure TMainForm.CreateActionExecute(Sender: TObject);
begin
  FolderCreateTran(PathMemo);
end;

/// <summary>
/// フォルダ削除アクション処理
/// </summary>
/// <param name="Sender">オブジェクト</param>
procedure TMainForm.DeleteActionExecute(Sender: TObject);
begin
  FolderDeleteTran(PathMemo, False);
end;

/// <summary>
/// フォルダ全削除アクション処理
/// </summary>
/// <param name="Sender">オブジェクト</param>
procedure TMainForm.AllDeleteActionExecute(Sender: TObject);
begin
  FolderDeleteTran(PathMemo, True);
end;

/// <summary>
/// 入力したパスをクリアアクション処理
/// </summary>
/// <param name="Sender">オブジェクト</param>
procedure TMainForm.ClearActionExecute(Sender: TObject);
begin
  PathMemo.Clear;
  PathMemo.SetFocus;
end;

{$endregion}

{$region 'コンテキストメニュー'}

/// <summary>
/// 表示処理
/// </summary>
/// <param name="Sender">オブジェクト</param>
procedure TMainForm.PopupMenu1Popup(Sender: TObject);
begin
  // 元に戻す判定処理の結果を判定する。
  if History.IsUndoFlg then
  begin
    // 可能の場合
    NUndo.Enabled := True;
  end
  else
  begin
    // 不可の場合
    NUndo.Enabled := False;
  end;

  // やり直し判定処理の結果を判定する。
  if History.IsRedoFlg then
  begin
    // 可能の場合
    NRedo.Enabled := True;
  end
  else
  begin
    // 不可の場合
    NRedo.Enabled := False;
  end;
end;

/// <summary>
/// 元に戻す
/// </summary>
/// <param name="Sender">オブジェクト</param>
procedure TMainForm.NUndoClick(Sender: TObject);
begin
  History.Undo;
end;

/// <summary>
/// やり直し
/// </summary>
/// <param name="Sender">オブジェクト</param>
procedure TMainForm.NRedoClick(Sender: TObject);
begin
  History.Redo;
end;

/// <summary>
/// 切り取り
/// </summary>
/// <param name="Sender">オブジェクト</param>
procedure TMainForm.NCutClick(Sender: TObject);
begin
  PathMemo.CutToClipboard;
end;

/// <summary>
/// コピー
/// </summary>
/// <param name="Sender">オブジェクト</param>
procedure TMainForm.NCopyClick(Sender: TObject);
begin
  PathMemo.CopyToClipboard;
end;

/// <summary>
/// 貼り付け
/// </summary>
/// <param name="Sender">オブジェクト</param>
procedure TMainForm.NPasteClick(Sender: TObject);
begin
  PathMemo.PasteFromClipboard;
end;

/// <summary>
/// 削除
/// </summary>
/// <param name="Sender">オブジェクト</param>
procedure TMainForm.NDeleteClick(Sender: TObject);
begin
  var DeleteText: string;
  var SelectionStart: Integer := PathMemo.SelStart;

    // 選択範囲を判定する。
  if PathMemo.SelLength = 0 then
  begin
    // カーソル位置と文字数を判定する。
    if SelectionStart < Length(PathMemo.Text) then
    begin
      // カーソル位置より、文字数の方が多い場合

      // カーソル開始位置に存在する文字を判定する。
      if PathMemo.Text[SelectionStart + 1] = #13 then
      begin
        // 改行コードを判定する。
        if (SelectionStart + 2 < Length(PathMemo.Text)) and (PathMemo.Text[SelectionStart + 2] = #10) then
        begin
          // 【\r\n】の場合
          DeleteText := PathMemo.Text[SelectionStart + 2];
        end
        else
        begin
          // 上記以外の場合
          DeleteText := PathMemo.Text[SelectionStart + 1];
        end;
      end
      else
      begin
         // 改行コード以外の場合
         DeleteText := PathMemo.Text[SelectionStart + 1];
      end
    end;

    // 範囲選択を行い、選択された範囲をブランクで置換（削除）する。
    PathMemo.SelStart := SelectionStart;
    PathMemo.SelLength := Length(DeleteText);
    PathMemo.Perform(EM_REPLACESEL, 1, LPARAM(PChar('')));
  end
  else
  begin
    // 存在する場合
    PathMemo.Perform(EM_REPLACESEL, 1, LPARAM(PChar('')));
  end;

  // カーソルを設定する。
  PathMemo.SelStart := SelectionStart;
end;

/// <summary>
/// 行選択
/// </summary>
/// <param name="Sender">オブジェクト</param>
procedure TMainForm.NSelectRowClick(Sender: TObject);
var
  Index, SelStart, SelEnd, StartLine, EndLine: Integer;
begin
  // 選択開始位置・選択終了位置を取得し、そこから更に開始行位置・終了行位置を取得・保持する。
  PathMemo.Perform(EM_GETSEL, WPARAM(@SelStart), LPARAM(@SelEnd));
  StartLine := PathMemo.Perform(EM_LINEFROMCHAR, SelStart, 0);
  EndLine := PathMemo.Perform(EM_LINEFROMCHAR, SelEnd, 0);

  // 選択開始行位置・選択終了行位置を判定する。
  if StartLine = EndLine then
  begin
    // 同一行の場合
    PathMemo.SelStart := PathMemo.Perform(EM_LINEINDEX, StartLine, 0);
    PathMemo.SelLength := PathMemo.Perform(EM_LINELENGTH, PathMemo.SelStart, 0);
  end
  else
  begin
    // 不一致の場合

    // 選択文字数を計算する。
    var SelLength: Integer := 0;
    for Index := StartLine to EndLine do
    begin
      // １行分桁数・改行コード桁数を加算する。
      SelLength := SelLength + PathMemo.Lines[Index].Length;
      SelLength := SelLength + PathMemo.Lines.LineBreak.Length;
    end;

    // 選択する。
    PathMemo.SelStart := PathMemo.Perform(EM_LINEINDEX, StartLine, 0);
    PathMemo.SelLength := SelLength;
  end
end;

/// <summary>
/// 全選択
/// </summary>
/// <param name="Sender">オブジェクト</param>
procedure TMainForm.NSelectAllClick(Sender: TObject);
begin
  PathMemo.SelectAll;
end;

/// <summary>
/// 選択行のフォルダを表示
/// </summary>
/// <param name="Sender">オブジェクト</param>
procedure TMainForm.NSelectRowFolderOpenClick(Sender: TObject);
var
  Index, SelStart, SelEnd, StartLine, EndLine: Integer;
  PathList: TStrings;
begin
  // 選択開始位置・選択終了位置を取得し、そこから更に開始行位置・終了行位置を取得・保持する。
  PathMemo.Perform(EM_GETSEL, WPARAM(@SelStart), LPARAM(@SelEnd));
  StartLine := PathMemo.Perform(EM_LINEFROMCHAR, SelStart, 0);
  EndLine := PathMemo.Perform(EM_LINEFROMCHAR, SelEnd, 0);

  PathList := TStringList.Create;
  var TempPath: string;

  for Index := StartLine to EndLine do
  begin
    // 入力されたパスの有無を判定する。
    if (PathMemo.Lines[Index] <> '') and (PathList.IndexOf(PathMemo.Lines[Index]) = -1) then
      // ドライブレターを大文字にし、パスを設定する。
      TempPath := UpperCase(PathMemo.Lines[Index].Substring(0, 1)) + PathMemo.Lines[Index].Substring(1);

      // フォルダ存在チェック処理を行い、結果を判定する。
      if IsExistsDirectory(TempPath) = True then
      begin
        // 存在している場合
        PathList.Add(TempPath);
      end;
  end;

  for Index := 0 to PathList.Count - 1 do
  begin
    // フォルダを表示する。
    ShellExecute(Handle, 'explore', PWideChar(WideString(PathList[Index])), nil, nil, SW_SHOWNORMAL);
  end;

  // リストを破棄する。
  PathList.Free;
end;

{$endregion}

{$region '内部処理'}

/// <summary>
/// フォルダ作成処理
/// </summary>
/// <param name="PathMemo">パスメモ</param>
procedure TMainForm.FolderCreateTran(PathMemo: TMemo);
var
  ReturnValue, Index: Integer;
  PathMemoText, Path, TempPath: string;
  PathAr: TStrings;
  DriveLetterList, CreatePathList: TList<string>;
begin
  DriveLetterList := TList<string>.Create;
  CreatePathList := TList<string>.Create;

  try
    ReturnValue := Application.MessageBox(
      '入力されたパスでフォルダを作成してよろしいですか？',
      'フォルダ作成確認',
      MB_YESNO or MB_DEFBUTTON2 or MB_ICONQUESTION);
    if ReturnValue = IDNO then
    begin
      Exit;
    end;

    // 未入力チェック
    PathMemoText := PathMemo.Text;
    if Trim(PathMemoText.Replace('　', '')) = '' then
    begin
      Application.MessageBox(
        'フォルダパスが未入力です。',
        ErTitle,
        MB_OK or MB_DEFBUTTON1 or MB_ICONERROR);
      Exit;
    end;

    // ドライブ文字リストを取得・保持する。
    DriveLetterList := GetDriveLetterList;

    PathAr := PathMemo.Lines;

    // 入力されたパスを基に処理を行う。
    for Index:= 0 to PathAr.Count - 1 do
    begin
      Path := PathAr[Index];

      // パスの有無を判定する。
      if Trim(Path.Replace('　', '')) = '' then
      begin
        // 存在しない場合
        CreatePathList.Add('');
      end
      else
      begin
        // 存在する場合

        // チェック処理を行い、結果を判定する。
        if CheckTran(Path, DriveLetterList, Index + 1) then
        begin
          // 正常の場合

          // パスを生成し、作成パスリストに、設定する。
          CreatePathList.Add(UpperCase(Path.Substring(0, 1)) + Path.Substring(1));
        end
        else
        begin
          // エラーの場合
          Exit;
        end;
      end;
    end;

    // 予約語チェック
    Index := 0;

    try
      // 仮フォルダパスを取得・保持する。
      while True do
      begin
        TempPath := DriveLetterList[0] + Yen + Index.ToString;

        // フォルダ検索処理を呼び出し、結果の有無を判定する。
        if IsExistsDirectory(TempPath) then
        begin
          // 存在する場合
          Index := Index + 1;
        end
        else
        begin
          // 存在しない場合
          Break;
        end;
      end;

      for Index:= 0 to CreatePathList.Count - 1 do
      begin
        // 仮フォルダを作成し、作成の有無を判定する。
        // （予約語を使用しても、DelphiだとExceptionが発生しないため、フォルダの有無で判定する。）
        ForceDirectories(TempPath + CreatePathList[Index].Substring(2));
        if not IsExistsDirectory(TempPath + CreatePathList[Index].Substring(2)) then
        begin
          // フォルダが存在しない場合

          // メッセージボックスを表示した状態で、仮フォルダが存在する状態になるため、仮フォルダを削除する。
          TDirectory.Delete(TempPath, True);
          TempPath := '';

          Application.MessageBox(
            PChar(ConvertNumberWide(Index + 1) + '行目：フォルダ名に使用出来ない文字列（予約語）が含まれています。'),
            ErTitle,
            MB_OK or MB_DEFBUTTON1 or MB_ICONERROR);
          Exit;
        end;
      end;

      // メッセージボックスを表示した状態で、仮フォルダが存在する状態になるため、仮フォルダを削除する。
      TDirectory.Delete(TempPath, True);
      TempPath := '';

      // フォルダを作成する。
      for Index:= 0 to CreatePathList.Count - 1 do
      begin
        if CreatePathList[Index] <> '' Then
        begin
          ForceDirectories(CreatePathList[Index]);
        end;
      end;

      Application.MessageBox(
        PChar('フォルダが作成されました。'),
        'フォルダ作成完了',
        MB_OK or MB_DEFBUTTON1 or MB_ICONINFORMATION);
    except
      on Ex: Exception do
      begin
        Application.MessageBox(
          PWideChar(Ex.Message),
          ErTitle,
          MB_OK or MB_DEFBUTTON1 or MB_ICONERROR);
      end;
    end;
  finally
    // 仮フォルダのパスの保持の有無を判定する。
    if TempPath <> '' then
    begin
      // 保持している場合、仮フォルダを削除する。
      TDirectory.Delete(TempPath, True);
    end;

    DriveLetterList.Free;
    CreatePathList.Free;
    PathMemo.SetFocus;
  end;
end;

/// <summary>
/// フォルダ削除処理
/// </summary>
/// <param name="PathMemo">パスメモ</param>
/// <param name="AllDeleteFlg">全削除フラグ（True：全削除・False：削除）</param>
procedure TMainForm.FolderDeleteTran(PathMemo: TMemo; AllDeleteFlg: Boolean);
var
  Title, Text, PathMemoText, Path, TempPath: string;
  ReturnValue, Index: Integer;
  PathAr: TStrings;
  FolderNameAr: TArray<string>;
  DriveLetterList, DeletePathList: TList<string>;
begin
  DriveLetterList := TList<string>.Create;
  DeletePathList := TList<string>.Create;

  try
    // 全削除フラグを判定する。
    if AllDeleteFlg then
    begin
      // 全削除の場合
      Title := 'フォルダ全削除確認';
      Text := '入力されたパスの全フォルダを削除してよろしいですか？'#13#10'（フォルダ内にファイルが存在する場合、ファイルごと削除します。）';
    end
    else
    begin
      // 削除の場合
      Title := 'フォルダ削除確認';
      Text := '入力されたパスのフォルダを削除してよろしいですか？'#13#10'（フォルダ内にファイルが存在する場合、ファイルごと削除します。）';
    end;

    ReturnValue := Application.MessageBox(
      PChar(Text),
      PChar(Title),
      MB_YESNO or MB_DEFBUTTON2 or MB_ICONQUESTION);
    if ReturnValue = IDNO then
    begin
      Exit;
    end;

    // 未入力チェック
    PathMemoText := PathMemo.Text;
    if Trim(PathMemoText.Replace('　', '')) = '' then
    begin
      Application.MessageBox(
        'フォルダパスが未入力です。',
        ErTitle,
        MB_OK or MB_DEFBUTTON1 or MB_ICONERROR);
      Exit;
    end;

    // ドライブ文字リストを取得・保持する。
    DriveLetterList := GetDriveLetterList;

    PathAr := PathMemo.Lines;

    // 入力されたパスを基に処理を行う。
    for Index:= 0 to PathAr.Count - 1 do
    begin
      Path := PathAr[Index];

      // パスの有無を判定する。
      if Trim(Path.Replace('　', '')) <> '' then
      begin
        // パスが存在する場合

        // チェック処理を行い、結果を判定する。
        if CheckTran(Path, DriveLetterList, Index + 1) = False then
        begin
          // エラーの場合、処理を終了する。
          Exit;
        end;

        // パスを生成する。
        TempPath := UpperCase(Path.Substring(0, 1)) + Path.Substring(1);
        if IsExistsDirectory(TempPath) then
        begin
          // 存在する場合

          // パスが既に保持されているか判定する。
          if not DeletePathList.Contains(TempPath) then
          begin
            // 存在していない場合、削除パスリストに設定する。
            DeletePathList.Add(TempPath);
          end;
        end
        else
        begin
          // 存在しない場合
          Application.MessageBox(
            PChar(ConvertNumberWide(Index + 1) + '行目：入力されたフォルダが存在しません。'),
            ErTitle,
            MB_OK or MB_DEFBUTTON1 or MB_ICONERROR);
          Exit;
        end;
      end;
    end;

    try
      // 削除パスリストの全項目に処理を行う。
      for Index:= 0 to DeletePathList.Count - 1 do
      begin
        // 全削除フラグを判定する。
        if AllDeleteFlg then
        begin
          // 全削除の場合
          FolderNameAr := SplitString(DeletePathList[Index], Yen);
          Path := FolderNameAr[0] + Yen + FolderNameAr[1]
        end
        else
        begin
          // 削除の場合
          Path := DeletePathList[Index];
        end;

        // フォルダの有無を判定する。
        // （フォルダパスが重複する場合、存在しないフォルダに削除を行い、エラーが発生するため。）
        if IsExistsDirectory(Path) then
        begin
          // 存在する場合、フォルダを削除する。
        TDirectory.Delete(Path, True)
        end;
      end;

      // 全削除フラグを判定する。
      if AllDeleteFlg then
      begin
        // 全削除の場合
        Title := 'フォルダ全削除完了';
        Text := 'フォルダが全削除されました。';
      end
      else
      begin
        // 削除の場合
        Title := 'フォルダ削除完了';
        Text := 'フォルダが削除されました。';
      end;

      Application.MessageBox(
        PChar(Text),
        PChar(Title),
        MB_OK or MB_DEFBUTTON1 or MB_ICONINFORMATION);
    except
      on Ex: Exception do
      begin
        Application.MessageBox(
          PWideChar(Ex.Message),
          ErTitle,
          MB_OK or MB_DEFBUTTON1 or MB_ICONERROR);
      end;
    end;

  finally
    DeletePathList.Free;
    DriveLetterList.Free;
  end;
end;

/// <summary>
/// チェック処理
/// </summary>
/// <param name="Path">パス</param>
/// <param name="DriveLetterList">ドライブ文字リスト</param>
/// <param name="RowIndex">行数</param>
/// <returns>チェック結果（True：正常・False：エラー）</returns>
function TMainForm.CheckTran(
  Path: string;
  DriveLetterList: TList<string>;
  RowIndex: Integer): Boolean;
var
  TempPath, FolderName, DriveLetter: string;
  FolderNameAr: TArray<string>;
  Index: Integer;
begin
  // 文字数チェック
  if 240 < Path.Length then
  begin
    // 文字数が２４０文字以上の場合、エラーとする。
    Application.MessageBox(
      PChar(ConvertNumberWide(RowIndex) + '行目：パスの入力可能文字数は、２４０文字までです。'#13#10'入力文字数：' + ConvertNumberWide(Path.Length) + '文字'),
      ErTitle,
      MB_OK or MB_DEFBUTTON1 or MB_ICONERROR);
    Exit(False);
  end;

  // パスを配列化する。
  FolderNameAr := SplitString(Path, Yen);

  // 要素数・末端要素を判定する。
  if (Length(FolderNameAr) > 0) and (FolderNameAr[High(FolderNameAr)] = '') then
  begin
    // 要素が存在し、末端要素が空要素の場合、削除する。
    //（【StringSplitOptions.RemoveEmptyEntries】の処理）
    SetLength(FolderNameAr, Length(FolderNameAr) - 1);
  end;

  // 要素数チェック
  if 2 > Length(FolderNameAr) then
  begin
    // 要素数が２個以下の場合、エラーとする。
    Application.MessageBox(
      PChar(ConvertNumberWide(RowIndex) + '行目：フォルダ名が入力されていません。'),
      ErTitle,
      MB_OK or MB_DEFBUTTON1 or MB_ICONERROR);
    Exit(False);
  end;

  // サロゲートペア・結合文字チェック
  if not IsSurrogateOrCombiningChar(Path) then
  begin
    // エラーの場合
    Application.MessageBox(
      PChar(ConvertNumberWide(RowIndex) + '行目：フォルダ名に使用出来ない文字（サロゲートペア・結合文字）が含まれています。'),
      ErTitle,
      MB_OK or MB_DEFBUTTON1 or MB_ICONERROR);
    Exit(False);
  end;

  // ドライブ文字を除いたパスを設定する。（【\】無し）
  TempPath := string.Join(Yen, FolderNameAr, 1, Length(FolderNameAr)-1);

  // パス使用文字チェック
  for Index := 0 to High(InvalidCharAr) do
  begin
    if Pos(InvalidCharAr[Index], TempPath) <> 0 then
    begin
      Application.MessageBox(
        PChar(ConvertNumberWide(RowIndex) + '行目：フォルダ名に使用出来ない文字が含まれています。'),
        ErTitle,
        MB_OK or MB_DEFBUTTON1 or MB_ICONERROR);
      Exit(False);
    end;
  end;

  // 無名フォルダチェック
  if (Pos(DoubleYen, Path) <> 0) or Path.EndsWith(Yen) Then
  begin
    Application.MessageBox(
        PChar(ConvertNumberWide(RowIndex) + '行目：フォルダ名が指定されていない箇所があります。'),
        ErTitle,
        MB_OK or MB_DEFBUTTON1 or MB_ICONERROR);
      Exit(False);
  end;

  // スペースチェック
  for FolderName in FolderNameAr do
  begin
    // 全角スペースを置換し、値を判定する。
    if Trim(FolderName.Replace('　', '')) = '' then
    begin
      Application.MessageBox(
        PChar(ConvertNumberWide(RowIndex) + '行目：スペース（全角・半角問わず）のみのフォルダは、処理出来ません。'),
        ErTitle,
        MB_OK or MB_DEFBUTTON1 or MB_ICONERROR);
      Exit(False);
    end;
  end;

  // ドライブ文字を、大文字に変換し、保持する。
  DriveLetter := UpperCase(FolderNameAr[0]);

  // ドライブチェック
  if not DriveLetterList.Contains(DriveLetter) then
  begin
    Application.MessageBox(
      PChar(ConvertNumberWide(RowIndex) + '行目：フォルダが作成出来ないドライブ文字（C:やD:）が入力されています。'),
      ErTitle,
      MB_OK or MB_DEFBUTTON1 or MB_ICONERROR);
    Exit(False);
  end;

  Result := True;
end;

/// <summary>
/// サロゲートペア・結合文字チェック処理
/// </summary>
/// <param name="TargetPath">対象パス</param>
/// <returns>チェック結果（True：正常・False：エラー）</returns>
function TMainForm.IsSurrogateOrCombiningChar(TargetPath: string): Boolean;
var
  Index: Integer;
begin
  Result:= True;

  // 改行コードを置換する。
  TargetPath := StringReplace(TargetPath, #13#10, '', [rfReplaceAll]);

  // 対象文字列の有無を判定する。
  if TargetPath <> '' then
  begin
    // 全文字列に処理を行う。
    for Index:= 1 to Length(TargetPath) do
    begin
      // サロゲートか判定する。
      if TargetPath[Index].IsSurrogate then
      begin
        // サロゲートの場合
        Exit(False);
      end
    end;

    // 文字数・正規後文字数を判定する。
    if Length(TargetPath) <> GetNormalizedStringLength(TargetPath) then
    begin
      // 不一致の場合
      Exit(False);
    end;
  end;
end;

/// <summary>
/// 正規後文字数取得処理
/// </summary>
/// <param name="TargetStr">対象文字列</param>
/// <returns>正規化した後文字数</returns>
function TMainForm.GetNormalizedStringLength(const TargetStr: string): Integer;
var
  Buf: string;
begin
  if not IsNormalizedString(NormalizationC, PChar(TargetStr), -1) then
  begin
    SetLength(Buf, NormalizeString(NormalizationC, PChar(TargetStr), Length(TargetStr), nil, 0));
    Result := NormalizeString(NormalizationC, PChar(TargetStr), Length(TargetStr), PChar(Buf), Length(Buf));
  end
  else
  begin
    Result := Length(TargetStr);
  end;
end;

/// <summary>
/// フォルダ存在チェック処理
/// </summary>
/// <param name="TargetStr">対象パス</param>
/// <returns>チェック結果（true：正常・false：エラー）</returns>
function TMainForm.IsExistsDirectory(TargetPath: string): Boolean;
var
  ExpandedPath: string;
begin
  if TargetPath = '' then
  begin
    Result := False;
  end
  else
  begin
    // パスを正規化し、検索結果を返却する。
    ExpandedPath := TPath.GetFullPath(TargetPath);
    Result := TDirectory.Exists(ExpandedPath, False);
  end;
end;

/// <summary>
/// ドライブレターリスト取得処理
/// </summary>
/// <returns>ドライブレターリスト</returns>
function TMainForm.GetDriveLetterList: TList<string>;
var
  Drive: Char;
  DriveType: Integer;
  VolumeName: array[0..MAX_PATH - 1] of Char;
  MaxComponentLength, FileSystemFlags: DWORD;
  EditFlg: Boolean;
  DriveLetterList: TList<string>;
begin
  DriveLetterList := TList<string>.Create;

  for Drive := 'A' to 'Z' do
  begin
    // ドライブ種類を取得・保持する。
    DriveType := GetDriveType(PChar(Drive + ':\'));

    // ドライブ種類を判定する。
    if DRIVE_FIXED = DriveType then
    begin
      // 固定ドライブの場合

      // 編集可能か判定する。
      EditFlg := GetVolumeInformation(
        PChar(Drive + ':\'),
        VolumeName,
        SizeOf(VolumeName),
        nil,
        MaxComponentLength,
        FileSystemFlags,
        nil,
        0);
      if EditFlg = True then
      begin
        // 編集可能の場合
        DriveLetterList.Add(Drive + ':');
      end;
    end;
  end;

  // ドライブレターリストを返却する。
  Result := DriveLetterList;
end;

/// <summary>
/// 半角数値→全角数値変換処理
/// </summary>
/// <param name="NarrowNumber">半角数値</param>
/// <returns>変換した全角数値</returns>
function TMainForm.ConvertNumberWide(NarrowNumber: Integer): string;
var
  NarrowNumberStr, ConvertNumber, Value: string;
  Number: Char;
begin
  NarrowNumberStr := NarrowNumber.ToString;
  ConvertNumber := '';
  for Number in NarrowNumberStr do
  begin
    if NarrowNumberDict.TryGetValue(Number, Value) then
    begin
      ConvertNumber := ConvertNumber + Value;
    end;
  end;
  Result := ConvertNumber;
end;

{$endregion}

end.

