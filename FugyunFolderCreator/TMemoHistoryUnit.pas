unit TMemoHistoryUnit;

{$MODE Delphi}

interface

uses
  Vcl.StdCtrls, Winapi.Windows, Winapi.Messages, Generics.Collections;

type
  /// <summary>
  /// �����������N���X
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
  /// ���������N���X
  /// </summary>
  MemoHistory = class
  class var
    /// <summary>
    /// �Ώۃ���
    /// </summary>
    TargetMemo: TMemo;

    /// <summary>
    /// �ő嗚��
    /// </summary>
    MaxHistoryCount: Integer;

    /// <summary>
    /// ������񃊃X�g
    /// </summary>
    HistoryInfoList: TList<MemoHistoryInfo>;

    /// <summary>
    /// �J�����g�����ʒu
    /// �i�O�a�`�r�d�j
    /// </summary>
    CurrentHistoryIndex: Integer;

    /// <summary>
    /// ����ǉ��t���O
    /// �iTrue�G�ǉ��\�EFalse�F�ǉ��s�j
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
/// �R���X�g���N�^
/// </summary>
/// <param name="Text">�e�L�X�g</param>
/// <param name="CaretPosition">�J�[�\���ʒu</param>
/// <param name="VScrollInfo">�����X�N���[���ʒu</param>
/// <param name="HScrollInfo">�����X�N���[���ʒu</param>
constructor MemoHistoryInfo.Create(Text: String; CaretPosition: Integer; VScrollInfo, HScrollInfo: TScrollInfo);
begin
  SetText := Text;
  SetCaretPosition := CaretPosition;
  SetVScrollInfo := VScrollInfo;
  SetHScrollInfo := HScrollInfo;
end;

/// <summary>
/// �R���X�g���N�^
/// </summary>
/// <param name="TMemo">�����i�����ݒ�ς݁j</param>
/// <param name="MaxHistoryCount">���𐔁i�f�t�H���g�F�P�O�O�O�j</param>
constructor MemoHistory.Create(TMemo: TMemo; MaxHistoryCount: Integer = 1000);
var
  NewHistoryInfo: MemoHistoryInfo;
begin
  TargetMemo := TMemo;
  MemoHistory.MaxHistoryCount := MaxHistoryCount;
  HistoryInfoList := TList<MemoHistoryInfo>.Create;
  HistoryAddFlg := True;

  // �V�������������쐬���Ēǉ�����B
  NewHistoryInfo := MemoHistoryInfo.Create(
    TargetMemo.Text,
    TargetMemo.SelStart,
    GetTMemoScrollInfo(SB_VERT),
    GetTMemoScrollInfo(SB_HORZ));
  HistoryInfoList.Add(NewHistoryInfo);

  // �J�����g�����ʒu��ݒ肷��B
  CurrentHistoryIndex := HistoryInfoList.Count - 1;
end;

/// <summary>
/// ����ǉ�����
/// </summary>
procedure MemoHistory.HistoryAdd;
var
  NewHistoryInfo: MemoHistoryInfo;
begin
  // ����ǉ��t���O�𔻒肷��B
  if HistoryAddFlg then
  begin
    // �ǉ��\�̏ꍇ

    // �J�����g�ʒu�𔻒肷��B
    if HistoryInfoList.Count <> CurrentHistoryIndex + 1 then
    begin
      // �ŐV�ȊO�̏ꍇ

      // ������񃊃X�g���X�V����B
      while HistoryInfoList.Count > CurrentHistoryIndex + 1 do
      begin
        HistoryInfoList.Delete(HistoryInfoList.Count - 1);
      end;
    end;

    // ��񃊃X�g�̍��ڐ��ƁA�ő嗚�𐔂𔻒肷��B
    if HistoryInfoList.Count = MaxHistoryCount then
    begin
      // ��v����ꍇ�A�ŏ��̗v�f���폜����B
      HistoryInfoList.Delete(0);
    end;

    // �V�������������쐬���Ēǉ�����B
    NewHistoryInfo := MemoHistoryInfo.Create(
      TargetMemo.Text,
      TargetMemo.SelStart,
      GetTMemoScrollInfo(SB_VERT),
      GetTMemoScrollInfo(SB_HORZ));
    HistoryInfoList.Add(NewHistoryInfo);

    // �J�����g�����ʒu��ݒ肷��B
    CurrentHistoryIndex := HistoryInfoList.Count - 1;
  end
  else
  begin
    // �ǉ��s�̏ꍇ
    HistoryAddFlg := True;
  end;
end;

/// <summary>
/// ���ɖ߂�����
/// </summary>
procedure MemoHistory.Undo;
var
  TargetIndex: Integer;
  HistoryInfo: MemoHistoryInfo;
begin
  // �g�p���闚�����ʒu�𔻒肷��B
  TargetIndex := CurrentHistoryIndex - 1;
  if 0 <= TargetIndex then
  begin
    // �O�ȏ�̏ꍇ

    // �`����~����B
    TargetMemo.Lines.BeginUpdate;

    // ����ǉ��t���O��ݒ肷��B
    HistoryAddFlg := False;

    TargetMemo.SetFocus;

    HistoryInfo := HistoryInfoList[TargetIndex];
    TargetMemo.Text := HistoryInfo.SetText;
    TargetMemo.SelStart := HistoryInfo.SetCaretPosition;
    SetTMemoScrollInfo(HistoryInfo);

    // �f�N�������g�����l���A�J�����g�����ʒu�Ƃ��Đݒ肷��B
    CurrentHistoryIndex := TargetIndex;

    // �`����s���B
    TargetMemo.Lines.EndUpdate;
  end;
end;

/// <summary>
/// ��蒼������
/// </summary>
procedure MemoHistory.Redo;
var
  TargetIndex: Integer;
  HistoryInfo: MemoHistoryInfo;
begin
  TargetIndex := CurrentHistoryIndex + 1;

  // ���ڐ��E�Ώۈʒu�𔻒肷��B
  if HistoryInfoList.Count > TargetIndex then
  begin
    // ���ڐ��̕��������ꍇ

    // �`����~����B
    TargetMemo.Lines.BeginUpdate;

    // ���������擾�E�ێ�����B
    HistoryInfo := HistoryInfoList[TargetIndex];

    // ����ǉ��t���O��ݒ肷��B
    HistoryAddFlg := False;

    TargetMemo.SetFocus;

    TargetMemo.Text := HistoryInfo.SetText;
    TargetMemo.SelStart := HistoryInfo.SetCaretPosition;
    SetTMemoScrollInfo(HistoryInfo);

    // �J�����g�����ʒu��ݒ肷��B
    CurrentHistoryIndex := TargetIndex;

    // �`����s���B
    TargetMemo.Lines.EndUpdate;
  end;
end;

/// <summary>
/// ���ɖ߂����菈��
/// </summary>
/// <returns>�`�F�b�N���ʁiTrue�F�\�EFalse�F�s�j</returns>
function MemoHistory.IsUndoFlg: Boolean;
begin
  Result := CurrentHistoryIndex > 0;
end;

/// <summary>
/// ��蒼�����蔻�菈��
/// </summary>
/// <returns>�`�F�b�N���ʁiTrue�F�\�EFalse�F�s�j</returns>
function MemoHistory.IsRedoFlg: Boolean;
var
  TargetIndex: Integer;
  HistoryInfo: MemoHistoryInfo;
begin
  TargetIndex := CurrentHistoryIndex + 1;

  // ������񃊃X�g���痚����񂪎擾�\�����肷��B
  if  (0 <= TargetIndex) and (TargetIndex < HistoryInfoList.Count) then
  begin
    // �擾�\�̏ꍇ
    HistoryInfo := HistoryInfoList[TargetIndex];
  end
  else
  begin
    // �擾�s�̏ꍇ
    HistoryInfo := nil;
  end;

  Result :=  HistoryInfo <> nil;
end;

/// <summary>
/// �f�X�g���N�^
/// </summary>
destructor MemoHistory.Destroy;
begin
  HistoryInfoList.Free;
  inherited Destroy;
end;

{$region '�v���C�x�[�g���\�b�h'}

/// <summary>
/// �����E�X�N���[�����擾����
/// </summary>
/// <param name="fnBar">�ΏۃX�N���[���o�[</param>
/// <returns>�e�L�X�g�{�b�N�X�E�X�N���[�����</returns>
function MemoHistory.GetTMemoScrollInfo(fnBar: Integer): TScrollInfo;
var
  ScrollInfo: TScrollInfo;
begin
  // ScrollInfo�\���̂�����������B
  ScrollInfo.cbSize := SizeOf(TScrollInfo);
  ScrollInfo.fMask := SIF_POS;
  ScrollInfo.nMin := 0;
  ScrollInfo.nMax := 0;
  ScrollInfo.nPage := 0;
  ScrollInfo.nPos := 0;
  ScrollInfo.nTrackPos := 0;

  // �X�N���[�����̎擾���A�ԋp����B
  GetScrollInfo(MemoHistory.TargetMemo.Handle, fnBar, ScrollInfo);
  Result := ScrollInfo;
end;

/// <summary>
/// �����E�X�N���[�����ݒ菈��
/// </summary>
/// <param name="HistoryInfo">�������</param>
procedure MemoHistory.SetTMemoScrollInfo(HistoryInfo: MemoHistoryInfo);
var
  VScrollInfo, HScrollInfo: TScrollInfo;
begin
  VScrollInfo := HistoryInfo.SetVScrollInfo;
  HScrollInfo := HistoryInfo.SetHScrollInfo;

  // �X�N���[���ʒu��ݒ肷��B
  SetScrollInfo(MemoHistory.TargetMemo.Handle, SB_VERT, VScrollInfo, True);
  SetScrollInfo(MemoHistory.TargetMemo.Handle, SB_HORZ, HScrollInfo, True);

  // �X�N���[���ʒu�𔽉f������B
  SendMessage(MemoHistory.TargetMemo.Handle, WM_VSCROLL, MakeWParam(SB_THUMBPOSITION, VScrollInfo.nPos), 0);
  SendMessage(MemoHistory.TargetMemo.Handle, WM_HSCROLL, MakeWParam(SB_THUMBPOSITION, HScrollInfo.nPos), 0);
end;

{$endregion}

end.

