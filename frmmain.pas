unit frmMain;

(*
 aViewer is an accessibility API object inspection tool.

Copyright (C) 2014 The Paciello Group

This program is free software; you can redistribute it and/or
modify it under the terms of the GNU General Public License
as published by the Free Software Foundation; either version 2
of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program; if not, write to the Free Software
Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.
*)
{$mode objfpc}{$H+}

interface

uses
  Windows, Classes, SysUtils, FileUtil, activexcontainer, Forms, Controls, strUtils,
  Graphics, Dialogs, StdCtrls, ComCtrls, ExtCtrls, ActnList, Menus, Math,
  SHDocVw_1_1_TLB, MSHTML_TLB, iAccessible2Lib_tlb, UIAutomationClient_TLB,
  ISimpleDOM, UIA_TLB, IntList, IniFiles, oleacc, Variants, regexpr,
  ActiveX, FocusRectWnd;

{$EXTERNALSYM VariantInit}
procedure VariantInit(var varg: olevariant); stdcall;
  external 'oleaut32.dll' Name 'VariantInit';




type
  HWINEVENTHOOK = HANDLE;
   {$EXTERNALSYM TFNWinEventProc}
  TFNWinEventProc = procedure(hWinEventHook: THandle; event: DWORD;
    hwnd: HWND; idObject, idChild: longint; idEventThread, dwmsEventTime: DWORD); stdcall;

  IStylesProvider = interface;

  IStylesProvider = interface(IUnknown)
    ['{19b6b649-f5d7-4a6d-bdcb-129252be588a}']
    function get_StyleId(out retVal: SYSINT): HResult; stdcall;
    function get_StyleName(out retVal: WideString): HResult; stdcall;
    function get_FillColor(out retVal: SYSINT): HResult; stdcall;
    function get_FillPatternStyle(out retVal: WideString): HResult; stdcall;
    function get_Shape(out retVal: WideString): HResult; stdcall;
    function get_FillPatternColor(out retVal: SYSINT): HResult; stdcall;
    function get_ExtendedProperties(out retVal: WideString): HResult; stdcall;
  end;

  PSCData = ^TSCData;

  TSCData = record
    Name: string;
    SCKey: TShortCut;
    actName: string;
  end;
  PNodeData = ^TNodeData;

  TNodeData = record
    Value1: string;
    Value2: string;
    Acc: IAccessible;
    iID: integer;

  end;
  PTreeData = ^TTreeData;

  TTreeData = record
    Acc: IAccessible;
    uiEle: IUIAUTOMATIONELEMENT;
    iID: integer;
    dummy: boolean;
  end;
   {PIE = ^FIE;
    FIE = record
        Iweb: IWebBrowser2;
    end; }

  PCP = ^TCP;

  TCP = record
    CP: IConnectionPoint;
    Cookie: integer;
  end;
  PUIA = ^TUIA;

  TUIA = record
    ID: integer;
    PName: string;
    Value: olevariant;
  end;

  { TwndMain }

  TwndMain = class(TForm)
    acFocus: TAction;
    acCursor: TAction;
    acRect: TAction;
    acCopy: TAction;
    acOnlyFocus: TAction;
    acParent: TAction;
    acChild: TAction;
    acPrevS: TAction;
    acNextS: TAction;
    acHelp: TAction;
    acShowTip: TAction;
    acSetting: TAction;
    ActionList1: TActionList;
    AXC1: TActiveXContainer;
    ImageList1: TImageList;
    ImageList2: TImageList;
    MainMenu1: TMainMenu;
    mnuAll: TMenuItem;
    mnuARIA: TMenuItem;
    mnuBln: TMenuItem;
    mnublnCode: TMenuItem;
    mnublnIA2: TMenuItem;
    mnublnMSAA: TMenuItem;
    mnuHTML: TMenuItem;
    mnuIA2: TMenuItem;
    mnuLang: TMenuItem;
    mnuMSAA: TMenuItem;
    mnuOAll: TMenuItem;
    mnuOpenB: TMenuItem;
    mnuOSel: TMenuItem;
    mnuSAll: TMenuItem;
    mnuSave: TMenuItem;
    mnuSelD: TMenuItem;
    mnuSelMode: TMenuItem;
    mnuSSel: TMenuItem;
    mnuTarget: TMenuItem;
    mnuTVCont: TMenuItem;
    mnuTVOAll: TMenuItem;
    mnuTVOpen: TMenuItem;
    mnuTVOSel: TMenuItem;
    mnuTVSAll: TMenuItem;
    mnuTVSave: TMenuItem;
    mnuTVSSel: TMenuItem;
    mnuUIA: TMenuItem;
    mnuView: TMenuItem;
    N1: TMenuItem;
    PageControl1: TPageControl;
    PopupMenu2: TPopupMenu;
    Splitter1: TSplitter;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    dlgTask: TTaskDialog;
    Timer1: TTimer;
    Timer2: TTimer;
    ToolBar1: TToolBar;
    tbFocus: TToolButton;
    tbCursor: TToolButton;
    tbParent: TToolButton;
    tbChild: TToolButton;
    ToolButton1: TToolButton;
    tbHelp: TToolButton;
    ToolButton3: TToolButton;
    tbRectAngle: TToolButton;
    tbShowTip: TToolButton;
    tbPrevS: TToolButton;
    tbNextS: TToolButton;
    tbSetting: TToolButton;
    tbCopy: TToolButton;
    ToolButton6: TToolButton;
    tbOnlyFocus: TToolButton;
    ToolButton8: TToolButton;
    tvMSAA: TTreeView;
    tvUIA: TTreeView;
    procedure acCursorExecute(Sender: TObject);
    procedure acFocusExecute(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure tvMSAAAddition(Sender: TObject; Node: TTreeNode);
    procedure tvMSAADeletion(Sender: TObject; Node: TTreeNode);
    procedure tvMSAAExpanding(Sender: TObject; Node: TTreeNode;
      var AllowExpansion: boolean);
  private
    WB1: TEvsWebBrowser;
    bTer: boolean;
    HTMLsFF: array [0..2, 0..1] of string;
    ARIAs: array [0..1, 0..1] of string;
    IA2Sts: array [0..17] of string;
    QSFailed: string;
    sHTML, sTxt, sTypeIE, sTypeFF, sARIA, NodeTxt, Err_Inter: string;
    bSelMode: boolean;
    iFocus, ShowSrcLen, iPID, iProg: integer;
    hHook: THandle;
    LangList, ClsNames: TStringList;
    rType, rTarg: string;
    arPT: array [0..2] of TPoint;

    Roles: array [0..42] of string;
    StyleID: array [0..16] of string;
    oldPT: TPoint;
    WndFocus, WndLabel, WndDesc, WndTarg: TwndFocusRect;
    //WndTip: TfrmTipWnd;
    hRgn1, hRgn2, hRgn3: hRgn;

    Created: boolean;
    CEle: IHTMLElement;
    SDom: ISImpleDOMNode;
    UIEle: IUIAutomationElement;
    //TreeTH: TreeThread;
    //UIATH: TreeThread4UIA;
    //thMSEx: MSAAExTh;
    //thUIAEx: UIAExTh;
    ActTV: TTreeview;
    Treemode, bPFunc, bAllSave: boolean;
    //UIA: IUIAutomation;
    P2W: integer;
    cDPI: integer;
    bFirstTime, bTabEvt: boolean;
    TipText, TipTextIA2, sFilter: string;
    sNode: TTreeNode;
    pNode: TTreeNode;
    sRC: TRect;
    hWndTip: THandle;
    iDefIndex: integer;
    ti: TOOLINFO;
    TransPath, SPath, TransDir: string;
    UIAuto: IUIAutomation;
    iAcc, accRoot, DocAcc: IAccessible;
    VarParent: variant;
    ExTip, nSelected: boolean;
    DefFont, DefW, DefH: integer;
    ScaleX, ScaleY, DefX, DefY: double;
    sMSAAtxt, sHTMLtxt, sUIATxt, sARIATxt, sIA2Txt: string;
    dEventTime: dword;
    ConvErr, HelpURL, DllPath, APPDir, sTrue, sFalse: string;

    procedure OnStatusTextChange(Sender: TObject; Text_: WideString);
    procedure ExecCmdLine;
    procedure Load;
    procedure mnuLangChildClick(Sender: TObject);
    function LoadLang: string;
    function MSAAText(pAcc: IAccessible = nil; TextOnly: boolean = False): string;
    function HTMLText: string;
    function HTMLText4FF: string;
    function UIAText(tEle: IUIAutomationElement = nil; HTMLout: boolean = False;
      tab: string = ''): string;
    function SetIA2Text(pAcc: IAccessible = nil; SetTL: boolean = True): string;
    function GetIA2State(iRole: integer): string;
    function GetIA2Role(iRole: integer): string;
    procedure GetNaviState(AllFalse: boolean = False);
    procedure ShowRectWnd(bkClr: TColor);
    procedure ShowRectWnd2(bkClr: TColor; RC: TRect);
    procedure CreateMSAATree;
    procedure GetTreeStructure;
    function Get_RoleText(Acc: IAccessible; Child: integer): string;
    function ARIAText: string;
    procedure WriteHTML;
  public

  end;


var
  wndMain: TwndMain;

  ac: IAccessible;
  DMode: boolean;
  sBodyTxt, sCodeTxt: string;
  refNode, LoopNode, sNode: TTreeNode;
  TBList, uTBList, DList, uDList: TIntegerList;
  lMSAA: array[0..11] of string;
  lIA2: array [0..15] of string;
  lUIA: array [0..61] of string;
  HTMLs: array [0..2, 0..1] of string;
  flgMSAA, flgIA2, flgUIA, flgUIA2: integer;
  scList: TList;
  None: string;




implementation

//function SetWinEventHook(EventMin, EventMax: UINT; HMod: HMODULE;
//  EventProc: WINEVENTPROC; IDProcess, IDThread: DWORD; Flags: UINT): HWINEVENTHOOK;
//  external 'User32.dll';

//function UnhookWinEvent(EventHook: HWINEVENTHOOK): longbool; external 'User32.dll';

function SetWinEventHook(eventMin, eventMax: DWORD; hmodWinEventProc: HMODULE;
  pfnWinEventProc: TFNWinEventProc; idProcess, idThread, dwFlags: DWORD): THandle;
  stdcall; external 'user32.dll';
{$EXTERNALSYM SetWinEventHook}

function UnhookWinEvent(hWinEventHook: THandle): BOOL; stdcall; external 'user32.dll';
{$EXTERNALSYM UnhookWinEvent}

{$R *.lfm}

{ TwndMain }

function GetTemp(var S: string): boolean;
var
  Len: integer;
begin
  Len := Windows.GetTempPath(0, nil);
  if Len > 0 then
  begin
    SetLength(S, Len);
    Len := Windows.GetTempPath(Len, PChar(S));
    SetLength(S, Len);
    S := SysUtils.IncludeTrailingPathDelimiter(S);
    Result := Len > 0;
  end
  else
    Result := False;
end;

procedure ShowErr(Msg: string);
begin
  if Dmode then
    MessageDlg(Msg, mtError, [mbOK], 0);
end;

procedure WinEvent(hWinEventHook: THandle; event: DWORD; HWND: HWND;
  idObject, idChild: longint; idEventThread, dwmsEventTime: DWORD); stdcall;
var

  vChild: variant;
  pAcc: IAccessible;
  iDis: IDispatch;
  hr: HResult;
const
  EVENT_OBJECT_FOCUS = $8005;

begin
  if wndMain.bPFunc then
    Exit;
  if event = EVENT_OBJECT_FOCUS then
  begin
    hr := AccessibleObjectFromEvent(HWND, idObject, idChild, @pAcc, vChild);
    if (hr = 0) and (Assigned(pAcc)) and (wndMain.dEventTime < dwmsEventTime) then
    begin
      wndMain.UIAuto.ElementFromiAccessible(pAcc, vChild, wndMain.uiEle);
      try
        wndMain.dEventTime := dwmsEventTime;

        if (wndMain.acOnlyFocus.Checked) then
        begin
          wndMain.VarParent := vChild;
          wndMain.iAcc := pAcc;
          pAcc.Get_accParent(iDis);
          wndMain.accRoot := iDis as IAccessible;
          wndMain.ShowRectWnd(clYellow);
        end
        else if (wndMain.acFocus.Checked) then
        begin
          wndMain.VarParent := vChild;
          wndMain.iAcc := pAcc;
          pAcc.Get_accParent(iDis);
          wndMain.accRoot := iDis as IAccessible;
          wndMain.Timer2.Enabled := False;
          wndMain.Timer2.Enabled := True;
        end;
      except
        on E: Exception do
          ShowErr(E.Message);
      end;
    end;

  end;

end;

function GetMultiState(State: cardinal): WideString;
var
  PC: PWideChar;
  List: TStringList;
  i: integer;
begin
  List := TStringList.Create;
  try
    if (State and STATE_SYSTEM_ALERT_HIGH) <> 0 then
    begin
      PC := WideStrAlloc(255);
      GetStateTextW(STATE_SYSTEM_ALERT_HIGH, PC, StrBufSize(PC));
      List.Add(PC);
      StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_ALERT_MEDIUM) <> 0 then
    begin
      PC := WideStrAlloc(255);
      GetStateTextW(STATE_SYSTEM_ALERT_MEDIUM, PC, StrBufSize(PC));
      List.Add(PC);
      StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_ALERT_LOW) <> 0 then
    begin
      PC := WideStrAlloc(255);
      GetStateTextW(STATE_SYSTEM_ALERT_LOW, PC, StrBufSize(PC));
      List.Add(PC);
      StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_ANIMATED) <> 0 then
    begin
      PC := WideStrAlloc(255);
      GetStateTextW(STATE_SYSTEM_ANIMATED, PC, StrBufSize(PC));
      List.Add(PC);
      StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_BUSY) <> 0 then
    begin
      PC := WideStrAlloc(255);
      GetStateTextW(STATE_SYSTEM_BUSY, PC, StrBufSize(PC));
      List.Add(PC);
      StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_CHECKED) <> 0 then
    begin
      PC := WideStrAlloc(255);
      GetStateTextW(STATE_SYSTEM_CHECKED, PC, StrBufSize(PC));
      List.Add(PC);
      StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_COLLAPSED) <> 0 then
    begin
      PC := WideStrAlloc(255);
      GetStateTextW(STATE_SYSTEM_COLLAPSED, PC, StrBufSize(PC));
      List.Add(PC);
      StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_DEFAULT) <> 0 then
    begin
      PC := WideStrAlloc(255);
      GetStateTextW(STATE_SYSTEM_DEFAULT, PC, StrBufSize(PC));
      List.Add(PC);
      StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_EXPANDED) <> 0 then
    begin
      PC := WideStrAlloc(255);
      GetStateTextW(STATE_SYSTEM_EXPANDED, PC, StrBufSize(PC));
      List.Add(PC);
      StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_EXTSELECTABLE) <> 0 then
    begin
      PC := WideStrAlloc(255);
      GetStateTextW(STATE_SYSTEM_EXTSELECTABLE, PC, StrBufSize(PC));
      List.Add(PC);
      StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_FLOATING) <> 0 then
    begin
      PC := WideStrAlloc(255);
      GetStateTextW(STATE_SYSTEM_FLOATING, PC, StrBufSize(PC));
      List.Add(PC);
      StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_FOCUSABLE) <> 0 then
    begin
      PC := WideStrAlloc(255);
      GetStateTextW(STATE_SYSTEM_FOCUSABLE, PC, StrBufSize(PC));
      List.Add(PC);
      StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_FOCUSED) <> 0 then
    begin
      PC := WideStrAlloc(255);
      GetStateTextW(STATE_SYSTEM_FOCUSED, PC, StrBufSize(PC));
      List.Add(PC);
      StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_HASPOPUP) <> 0 then
    begin
      PC := WideStrAlloc(255);
      GetStateTextW(STATE_SYSTEM_HASPOPUP, PC, StrBufSize(PC));
      List.Add(PC);
      StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_HOTTRACKED) <> 0 then
    begin
      PC := WideStrAlloc(255);
      GetStateTextW(STATE_SYSTEM_HOTTRACKED, PC, StrBufSize(PC));
      List.Add(PC);
      StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_INVISIBLE) <> 0 then
    begin
      PC := WideStrAlloc(255);
      GetStateTextW(STATE_SYSTEM_INVISIBLE, PC, StrBufSize(PC));
      List.Add(PC);
      StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_LINKED) <> 0 then
    begin
      PC := WideStrAlloc(255);
      GetStateTextW(STATE_SYSTEM_LINKED, PC, StrBufSize(PC));
      List.Add(PC);
      StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_MARQUEED) <> 0 then
    begin
      PC := WideStrAlloc(255);
      GetStateTextW(STATE_SYSTEM_MARQUEED, PC, StrBufSize(PC));
      List.Add(PC);
      StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_MIXED) <> 0 then
    begin
      PC := WideStrAlloc(255);
      GetStateTextW(STATE_SYSTEM_MIXED, PC, StrBufSize(PC));
      List.Add(PC);
      StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_MOVEABLE) <> 0 then
    begin
      PC := WideStrAlloc(255);
      GetStateTextW(STATE_SYSTEM_MOVEABLE, PC, StrBufSize(PC));
      List.Add(PC);
      StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_MULTISELECTABLE) <> 0 then
    begin
      PC := WideStrAlloc(255);
      GetStateTextW(STATE_SYSTEM_MULTISELECTABLE, PC, StrBufSize(PC));
      List.Add(PC);
      StrDispose(PC);
    end;
    //if (State and STATE_SYSTEM_NORMAL) <> 0 then
    if State = 0 then
    begin
      PC := WideStrAlloc(255);
      GetStateTextW(STATE_SYSTEM_NORMAL, PC, StrBufSize(PC));
      List.Add(PC);
      StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_OFFSCREEN) <> 0 then
    begin
      PC := WideStrAlloc(255);
      GetStateTextW(STATE_SYSTEM_OFFSCREEN, PC, StrBufSize(PC));
      List.Add(PC);
      StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_PRESSED) <> 0 then
    begin
      PC := WideStrAlloc(255);
      GetStateTextW(STATE_SYSTEM_PRESSED, PC, StrBufSize(PC));
      List.Add(PC);
      StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_PROTECTED) <> 0 then
    begin
      PC := WideStrAlloc(255);
      GetStateTextW(STATE_SYSTEM_PROTECTED, PC, StrBufSize(PC));
      List.Add(PC);
      StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_READONLY) <> 0 then
    begin
      PC := WideStrAlloc(255);
      GetStateTextW(STATE_SYSTEM_READONLY, PC, StrBufSize(PC));
      List.Add(PC);
      StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_SELECTABLE) <> 0 then
    begin
      PC := WideStrAlloc(255);
      GetStateTextW(STATE_SYSTEM_SELECTABLE, PC, StrBufSize(PC));
      List.Add(PC);
      StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_SELECTED) <> 0 then
    begin
      PC := WideStrAlloc(255);
      GetStateTextW(STATE_SYSTEM_SELECTED, PC, StrBufSize(PC));
      List.Add(PC);
      StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_SELFVOICING) <> 0 then
    begin
      PC := WideStrAlloc(255);
      GetStateTextW(STATE_SYSTEM_SELFVOICING, PC, StrBufSize(PC));
      List.Add(PC);
      StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_SIZEABLE) <> 0 then
    begin
      PC := WideStrAlloc(255);
      GetStateTextW(STATE_SYSTEM_SIZEABLE, PC, StrBufSize(PC));
      List.Add(PC);
      StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_TRAVERSED) <> 0 then
    begin
      PC := WideStrAlloc(255);
      GetStateTextW(STATE_SYSTEM_TRAVERSED, PC, StrBufSize(PC));
      List.Add(PC);
      StrDispose(PC);
    end;
    if (State and STATE_SYSTEM_UNAVAILABLE) <> 0 then
    begin
      PC := WideStrAlloc(255);
      GetStateTextW(STATE_SYSTEM_UNAVAILABLE, PC, StrBufSize(PC));
      List.Add(PC);
      StrDispose(PC);
    end;
    for I := 0 to List.Count - 1 do
    begin
      if i = 0 then
        Result := List.Strings[i]
      else
        Result := Result + ' , ' + List.Strings[i];
    end;
  finally
    List.Free;
  end;

end;

function DoubleToInt(d: double): integer;
begin
  SetRoundMode(rmUP);
  Result := Trunc(SimpleRoundTo(d));
end;

function GetLastErrorStr(ErrorCode: integer): string;
const
  MAX_MES = 512;
var
  Buf: PChar;
begin
  Buf := AllocMem(MAX_MES);
  try
    FormatMessage(Format_Message_From_System, nil, ErrorCode,
      (SubLang_Default shl 10) + Lang_Neutral,
      Buf, MAX_MES, nil);
  finally
    Result := Buf;
    FreeMem(Buf);
  end;
end;

function IntToBoolStr(i: integer): string;
begin
  Result := 'True';
  if i = 0 then
    Result := 'False';
end;

function TruncPow(i, t: integer): integer;
begin
  Result := Trunc(IntPower(i, t));
end;

function IsWinVista: boolean;
var
  VI: TOSVersionInfo;
begin
  Result := False;
  FillChar(VI, SizeOf(VI), 0);
  VI.dwOSVersionInfoSize := SizeOf(VI);
  if GetVersionEx(VI) then
  begin
    if (VI.dwMajorVersion >= 6) then
      Result := True;
  end;
end;

function IsLimited: boolean;
var
  hProcess, hToken: THandle;
  pt: TTokenElevationType;
  dwLength: DWORD;
begin
  Result := True;
  if not IsWinVista then
  begin
    Result := False;
    Exit;
  end;
  hProcess := GetCurrentProcess();
  if OpenProcessToken(hProcess, TOKEN_QUERY {or TOKEN_QUERY_SOURCE}, hToken) then
  begin
    if GetTokenInformation(hToken, Windows.TTokenInformationClass(
      TokenElevationType), @pt, sizeOf(@pt), dwLength) then
    begin
      if pt <> TokenElevationTypeLimited then
        Result := False;
      CloseHandle(hToken);

    end;
  end;
end;

function DeleteCRLF(d: WideString): WideString;
var
  s: WideString;
begin
  s := StringReplace(d, #10, ' ', [rfReplaceAll, rfIgnoreCase]);
  s := StringReplace(s, #13, '', [rfReplaceAll, rfIgnoreCase]);
  Result := s;
end;

function VarHaveValue(v: variant): boolean;
begin
  Result := True;
  if VarType(v) = varEmpty then
    Result := False;
  if VarIsEmpty(v) then
    Result := False;
  if VarIsClear(v) then
    Result := False;

end;

function Execute(FileName, Param: string; RunAdmin: boolean = False): cardinal;
var
  op: string;
begin
  if RunAdmin then
    op := 'runas'
  else
    op := 'open';
  Result := ShellExecute(0, PChar(op), PChar(FileName), PChar(Param),
    nil, SW_SHOWNORMAL);
end;

function GetMyDocPath: string;
var
  IIDList: PItemIDList;
  buffer: array [0..MAX_PATH - 1] of char;
const
  CSIDL_PERSONAL = $5;
begin
  IIDList := nil;
  Result := '';
  if SUCCEEDED(SHGetSpecialFolderLocation(Application.Handle,
    CSIDL_PERSONAL, IIDList)) then
  begin
    if not SHGetPathFromIDList(IIDList, buffer) then
    begin
      raise Exception.Create('A virtual diractory cannot be acquired.');
    end
    else
      Result := StrPas(Buffer);
  end;
end;

procedure TwndMain.ExecCmdLine;
var
  i: integer;
  d: string;
begin
  if ParamCount > 0 then
  begin
    DMode := False;
    for i := 1 to ParamCount do
    begin
      d := LowerCase(ParamStr(i));
      if (d = '-fronly') or (d = '/fronly') then
      begin

        acOnlyFocus.Checked := True;
        tbOnlyFocus.Down := True;
        //ExecOnlyFocus;
      end
      else if (d = '-d') or (d = '/d') then
        DMode := True;
    end;

  end;
end;

function GetFileVersionString(filename: string = ''): string;
var
  rcFI: record
    Dummy: DWord;
    VerInfo: Pointer;
    VerInfoSize: DWord;
    VerValueSize: DWord;
    VerValue: PVSFixedFileInfo;
  end;
begin
  Result := '';
  rcFI.Dummy := 0;
  if filename = '' then
    filename := ParamStr(0);
  rcFI.VerInfoSize := GetFileVersionInfoSize(PChar(filename), rcFI.Dummy);
  if rcFI.VerInfoSize = 0 then
    exit;
  GetMem(rcFI.VerInfo, rcFI.VerInfoSize);
  try
    GetFileVersionInfo(PChar(filename), 0, rcFI.VerInfoSize, rcFI.VerInfo);
    VerQueryValue(rcFI.VerInfo, '\', Pointer(rcFI.VerValue), rcFI.VerValueSize);
    with rcFI.VerValue^ do
    begin
      Result := IntToStr(dwFileVersionMS shr 16);
      Result := Result + '.' + IntToStr(dwFileVersionMS and $FFFF);
      Result := Result + '.' + IntToStr(dwFileVersionLS shr 16);
      Result := Result + '.' + IntToStr(dwFileVersionLS and $FFFF);
    end;

  finally
    FreeMem(rcFI.VerInfo, rcFI.VerInfoSize);
  end;

end;

procedure TwndMain.mnuLangChildClick(Sender: TObject);
var
  i: integer;
  bChk: boolean;
  lf: string;
  ini: TMemIniFile;
begin
  if not (Sender is TMenuitem) then
    Exit;
  (Sender as TMenuitem).Checked := not (Sender as TMenuitem).Checked;
  if mnuLang.Count > 0 then
  begin
    bChk := False;
    for i := 0 to mnuLang.Count - 1 do
    begin
      if i > LangList.Count - 1 then
        Break;
      if mnuLang.Items[i].Checked then
      begin
        bChk := True;
        lf := LangList[i];
        Break;
      end;
    end;
    if not bChk then
    begin
      if (Sender is TMenuitem) then
      begin
        if (Sender as TMenuitem).MenuIndex > LangList.Count - 1 then
        begin
          mnuLang.Items[0].Checked := True;
          lf := LangList[0];
        end
        else
        begin
          (Sender as TMenuitem).Checked := True;
          lf := LangList[(Sender as TMenuitem).MenuIndex];
        end;
      end;
    end;
    if LowerCase(TransPath) <> LowerCase(TransDir + lf) then
    begin
      TransPath := TransDir + lf;
      //showmessage(Transpath);
      LoadLang;
      ini := TMemIniFile.Create(SPath, TEncoding.Unicode);
      try
        Ini.WriteString('Settings', 'FontName', Font.Name);
        ini.WriteInteger('Settings', 'FontSize', Font.Size);
        Ini.WriteInteger('Settings', 'Charset', Font.Charset);
        Ini.WriteString('Settings', 'LangFile', lf);
        ini.UpdateFile;
      finally
        ini.Free;
      end;
    end;
  end;
end;

function TwndMain.LoadLang: string;
var
  ini: TMemIniFile;
  d, Msg, UIA_fail: string;
begin
  Result := 'CoCreateInstance failed.';
  ini := TMemIniFile.Create(TransPath, TEncoding.UTF8);
  ClsNames := TStringList.Create;
  try

    lMSAA[0] := ini.ReadString('MSAA', 'MSAA', 'MS Active Accessibility');
    lMSAA[1] := ini.ReadString('MSAA', 'accName', 'accName');
    lMSAA[2] := ini.ReadString('MSAA', 'accRole', 'accRole');
    lMSAA[3] := ini.ReadString('MSAA', 'accState', 'accState');
    lMSAA[4] := ini.ReadString('MSAA', 'accDescription', 'accDescription');
    lMSAA[5] := ini.ReadString('MSAA', 'accDefaultAction', 'accDefaultAction');
    lMSAA[6] := ini.ReadString('MSAA', 'accValue', 'accValue');
    lMSAA[7] := ini.ReadString('MSAA', 'accParent', 'accParent');
    lMSAA[8] := ini.ReadString('MSAA', 'accChildCount', 'accChildCount');
    lMSAA[9] := ini.ReadString('MSAA', 'accHelp', 'accHelp');
    lMSAA[10] := ini.ReadString('MSAA', 'accHelpTopic', 'accHelpTopic');
    lMSAA[11] := ini.ReadString('MSAA', 'accKeyboardShortcut',
      'accKeyboardShortcut');


    lIA2[0] := ini.ReadString('IA2', 'IA2', 'IAccessible2');
    lIA2[1] := ini.ReadString('IA2', 'Name', 'Name');
    lIA2[2] := ini.ReadString('IA2', 'Role', 'Role');
    lIA2[3] := ini.ReadString('IA2', 'States', 'States');
    lIA2[4] := ini.ReadString('IA2', 'Description', 'Description');
    lIA2[5] := ini.ReadString('IA2', 'Relations', 'Relations');
    //lIA2[6] := LoadTranslation('IA2', 'RelationTargets','Relation Targets');
    lIA2[6] := ini.ReadString('IA2', 'Attributes', 'Object Attributes');
    lIA2[7] := ini.ReadString('IA2', 'Value', 'Value');
    lIA2[8] := ini.ReadString('IA2', 'LocalizedExtendedRole',
      'LocalizedExtendedRole');
    lIA2[9] := ini.ReadString('IA2', 'LocalizedExtendedStates',
      'LocalizedExtendedStates');
    lIA2[10] := ini.ReadString('IA2', 'RangeValue', 'RangeValue');
    lIA2[11] := ini.ReadString('IA2', 'RV_Value', 'Value');
    lIA2[12] := ini.ReadString('IA2', 'RV_Minimum', 'Minimum');
    lIA2[13] := ini.ReadString('IA2', 'RV_Maximum', 'Maximum');
    lIA2[14] := ini.ReadString('IA2', 'Textattributes', 'Textattributes');
    lIA2[15] := ini.ReadString('IA2', 'uniqueID', 'uniqueID');


    lUIA[0] := ini.ReadString('UIA', 'UIA', 'UIAutomation');
    lUIA[1] := ini.ReadString('UIA', 'CurrentAcceleratorKey',
      'CurrentAcceleratorKey');
    lUIA[2] := ini.ReadString('UIA', 'CurrentAccessKey', 'CurrentAccessKey');
    lUIA[3] := ini.ReadString('UIA', 'CurrentAriaProperties',
      'CurrentAriaProperties');
    lUIA[4] := ini.ReadString('UIA', 'CurrentAriaRole', 'CurrentAriaRole');
    lUIA[5] := ini.ReadString('UIA', 'CurrentAutomationId', 'CurrentAutomationId');
    lUIA[6] := ini.ReadString('UIA', 'CurrentBoundingRectangle',
      'CurrentBoundingRectangle');
    lUIA[7] := ini.ReadString('UIA', 'CurrentClassName', 'CurrentClassName');
    lUIA[8] := ini.ReadString('UIA', 'CurrentControlType', 'CurrentControlType');
    lUIA[9] := ini.ReadString('UIA', 'CurrentCulture', 'CurrentCulture');
    lUIA[10] := ini.ReadString('UIA', 'CurrentFrameworkId', 'CurrentFrameworkId');
    lUIA[11] := ini.ReadString('UIA', 'CurrentHasKeyboardFocus',
      'CurrentHasKeyboardFocus');
    lUIA[12] := ini.ReadString('UIA', 'CurrentHelpText', 'CurrentHelpText');
    lUIA[13] := ini.ReadString('UIA', 'CurrentIsControlElement',
      'CurrentIsControlElement');
    lUIA[14] := ini.ReadString('UIA', 'CurrentIsContentElement',
      'CurrentIsContentElement');
    lUIA[15] := ini.ReadString('UIA', 'CurrentIsDataValidForForm',
      'CurrentIsDataValidForForm');
    lUIA[16] := ini.ReadString('UIA', 'CurrentIsEnabled', 'CurrentIsEnabled');
    lUIA[17] := ini.ReadString('UIA', 'CurrentIsKeyboardFocusable',
      'CurrentIsKeyboardFocusable');
    lUIA[18] := ini.ReadString('UIA', 'CurrentIsOffscreen', 'CurrentIsOffscreen');
    lUIA[19] := ini.ReadString('UIA', 'CurrentIsPassword', 'CurrentIsPassword');
    lUIA[20] := ini.ReadString('UIA', 'CurrentIsRequiredForForm',
      'CurrentIsRequiredForForm');
    lUIA[21] := ini.ReadString('UIA', 'CurrentItemStatus', 'CurrentItemStatus');
    lUIA[22] := ini.ReadString('UIA', 'CurrentItemType', 'CurrentItemType');
    lUIA[23] := ini.ReadString('UIA', 'CurrentLocalizedControlType',
      'CurrentLocalizedControlType');
    lUIA[24] := ini.ReadString('UIA', 'CurrentName', 'CurrentName');
    lUIA[25] := ini.ReadString('UIA', 'CurrentNativeWindowHandle',
      'CurrentNativeWindowHandle');
    lUIA[26] := ini.ReadString('UIA', 'CurrentOrientation', 'CurrentOrientation');
    lUIA[27] := ini.ReadString('UIA', 'CurrentProcessId', 'CurrentProcessId');
    lUIA[28] := ini.ReadString('UIA', 'CurrentProviderDescription',
      'CurrentProviderDescription');
    lUIA[29] := ini.ReadString('UIA', 'CurrentControllerFor',
      'CurrentControllerFor');
    lUIA[30] := ini.ReadString('UIA', 'CurrentDescribedBy', 'CurrentDescribedBy');
    lUIA[31] := ini.ReadString('UIA', 'CurrentFlowsTo', 'CurrentFlowsTo');
    lUIA[32] := ini.ReadString('UIA', 'CurrentLabeledBy', 'CurrentLabeledBy');
    lUIA[33] := ini.ReadString('UIA', 'CurrentLiveSetting', 'CurrentLiveSetting');
    //Added 2014/06/23
    lUIA[34] := ini.ReadString('UIA', 'IsDockPatternAvailable',
      'IsDockPatternAvailable');
    lUIA[35] := ini.ReadString('UIA', 'IsExpandCollapsePatternAvailable',
      'IsExpandCollapsePatternAvailable');
    lUIA[36] := ini.ReadString('UIA', 'IsGridItemPatternAvailable',
      'IsGridItemPatternAvailable');
    lUIA[37] := ini.ReadString('UIA', 'IsGridPatternAvailable',
      'IsGridPatternAvailable');
    lUIA[38] := ini.ReadString('UIA', 'IsInvokePatternAvailable',
      'IsInvokePatternAvailable');
    lUIA[39] := ini.ReadString('UIA', 'IsMultipleViewPatternAvailable',
      'IsMultipleViewPatternAvailable');
    lUIA[40] := ini.ReadString('UIA', 'IsRangeValuePatternAvailable',
      'IsRangeValuePatternAvailable');
    lUIA[41] := ini.ReadString('UIA', 'IsScrollPatternAvailable',
      'IsScrollPatternAvailable');
    lUIA[42] := ini.ReadString('UIA', 'IsScrollItemPatternAvailable',
      'IsScrollItemPatternAvailable');
    lUIA[43] := ini.ReadString('UIA', 'IsSelectionItemPatternAvailable',
      'IsSelectionItemPatternAvailable');
    lUIA[44] := ini.ReadString('UIA', 'IsSelectionPatternAvailable',
      'IsSelectionPatternAvailable');
    lUIA[45] := ini.ReadString('UIA', 'IsTablePatternAvailable',
      'IsTablePatternAvailable');
    lUIA[46] := ini.ReadString('UIA', 'IsTableItemPatternAvailable',
      'IsTableItemPatternAvailable');
    lUIA[47] := ini.ReadString('UIA', 'IsTextPatternAvailable',
      'IsTextPatternAvailable');
    lUIA[48] := ini.ReadString('UIA', 'IsTogglePatternAvailable',
      'IsTogglePatternAvailable');
    lUIA[49] := ini.ReadString('UIA', 'IsTransformPatternAvailable',
      'IsTransformPatternAvailable');
    lUIA[50] := ini.ReadString('UIA', 'IsValuePatternAvailable',
      'IsValuePatternAvailable');
    lUIA[51] := ini.ReadString('UIA', 'IsWindowPatternAvailable',
      'IsWindowPatternAvailable');
    lUIA[52] := ini.ReadString('UIA', 'IsItemContainerPatternAvailable',
      'IsItemContainerPatternAvailable');
    lUIA[53] := ini.ReadString('UIA', 'IsVirtualizedItemPatternAvailable',
      'IsVirtualizedItemPatternAvailable');
    //Added 2014/06/25
    lUIA[54] := ini.ReadString('UIA', 'RangeValue', 'RangeValue');
    lUIA[55] := ini.ReadString('UIA', 'RV_Value', 'Value');
    lUIA[56] := ini.ReadString('UIA', 'RV_IsReadOnly', 'IsReadOnly');
    lUIA[57] := ini.ReadString('UIA', 'RV_Minimum', 'Minimum');
    lUIA[58] := ini.ReadString('UIA', 'RV_Maximum', 'Maximum');
    lUIA[59] := ini.ReadString('UIA', 'RV_LargeChange', 'LargeChange');
    lUIA[60] := ini.ReadString('UIA', 'RV_SmallChange', 'SmallChange');
    lUIA[61] := ini.ReadString('UIA', 'StyleID', 'Style');

    StyleID[0] := ini.ReadString('UIA', 'StyleId_Custom', 'A custom style.');
    StyleID[1] := ini.ReadString('UIA', 'StyleId_Heading1',
      'A first level heading.');
    StyleID[2] := ini.ReadString('UIA', 'StyleId_Heading2',
      'A second level heading.');
    StyleID[3] := ini.ReadString('UIA', 'StyleId_Heading3',
      'A third level heading.');
    StyleID[4] := ini.ReadString('UIA', 'StyleId_Heading4',
      'A fourth level heading.');
    StyleID[5] := ini.ReadString('UIA', 'StyleId_Heading5',
      'A fifth level heading.');
    StyleID[6] := ini.ReadString('UIA', 'StyleId_Heading6',
      'A sixth level heading.');
    StyleID[7] := ini.ReadString('UIA', 'StyleId_Heading7',
      'A seventh level heading.');
    StyleID[8] := ini.ReadString('UIA', 'StyleId_Heading8',
      'An eighth level heading.');
    StyleID[9] := ini.ReadString('UIA', 'StyleId_Heading9',
      'A ninth level heading.');
    StyleID[10] := ini.ReadString('UIA', 'StyleId_Title', 'A title.');
    StyleID[11] := ini.ReadString('UIA', 'StyleId_Subtitle', 'A subtitle.');
    StyleID[12] := ini.ReadString('UIA', 'StyleId_Normal', 'Normal style.');
    StyleID[13] := ini.ReadString('UIA', 'StyleId_Emphasis',
      'Text that is emphasized.');
    StyleID[14] := ini.ReadString('UIA', 'StyleId_Quote', 'A quotation.');
    StyleID[15] := ini.ReadString('UIA', 'StyleId_BulletedList',
      'A list with bulleted items.');
    StyleID[16] := ini.ReadString('UIA', 'StyleId_NumberedList',
      'A list with numbered items.');

    HelpURL := ini.ReadString('General', 'Help_URL', 'http://www.google.com/');
    Caption := ini.ReadString('General', 'MSAA_Caption', 'Accessibility Viewer');
    Caption := Caption + ' - ' + GetFileVersionString(application.ExeName);
    Font.Name := ini.ReadString('General', 'FontName', 'Arial');
    Font.Size := ini.ReadInteger('General', 'FontSize', 9);
    d := ini.ReadString('General', 'Charset', 'ASCII_CHARSET');
    d := UpperCase(d);
    if d = 'ANSI_CHARSET' then
      Font.Charset := 0
    else if d = 'DEFAULT_CHARSET' then
      Font.Charset := 1
    else if d = 'SYMBOL_CHARSET' then
      Font.Charset := 2
    else if d = 'MAC_CHARSET' then
      Font.Charset := 77
    else if d = 'SHIFTJIS_CHARSET' then
      Font.Charset := 128
    else if d = 'HANGEUL_CHARSET' then
      Font.Charset := 129
    else if d = 'JOHAB_CHARSET' then
      Font.Charset := 130
    else if d = 'GB2312_CHARSET' then
      Font.Charset := 134
    else if d = 'CHINESEBIG5_CHARSET' then
      Font.Charset := 136
    else if d = 'KGREEK_CHARSET' then
      Font.Charset := 161
    else if d = 'TURKISH_CHARSET' then
      Font.Charset := 162
    else if d = 'VIETNAMESE_CHARSET' then
      Font.Charset := 163
    else if d = 'HEBREW_CHARSET' then
      Font.Charset := 177
    else if d = 'ARABIC_CHARSET' then
      Font.Charset := 178
    else if d = 'BALTIC_CHARSET' then
      Font.Charset := 186
    else if d = 'RUSSIAN_CHARSET' then
      Font.Charset := 204
    else if d = 'THAI_CHARSET' then
      Font.Charset := 222
    else if d = 'EASTEUROPE_CHARSET' then
      Font.Charset := 238
    else if d = 'OEM_CHARSET' then
      Font.Charset := 255
    else
      Font.Charset := 0;
    DefFont := Font.Size;
    mnuSelD.Caption :=
      ini.ReadString('General', 'MSAA_gbSelDisplay', 'Select Display');
    tbFocus.Hint := ini.ReadString('General', 'MSAA_tbFocusHint', 'Watch Focus');
    tbCursor.Hint := ini.ReadString('General', 'MSAA_tbCursorHint', 'Watch Cursor');
    tbRectAngle.Hint :=
      ini.ReadString('General', 'MSAA_tbRectAngleHint', 'Show Highlight Rectangle');

    tbShowtip.Hint := ini.ReadString('General', 'MSAA_tbBalloonHint',
      'Show Balloon tip');
    acShowTip.Hint := tbShowtip.Hint;
    tbCopy.Hint := ini.ReadString('General', 'MSAA_tbCopyHint',
      'Copy Text to Clipborad');
    tbOnlyFocus.Hint :=
      ini.ReadString('General', 'MSAA_tbFocusOnly', 'Focus rectangle only');

    tbParent.Hint := ini.ReadString('General', 'MSAA_tbParentHint',
      'Navigates to parent object');
    tbChild.Hint := ini.ReadString('General', 'MSAA_tbChildHint',
      'Navigates to first child object');
    tbPrevS.Hint := ini.ReadString('General', 'MSAA_tbPrevSHint',
      'Navigates to previous sibling object');
    tbNextS.Hint := ini.ReadString('General', 'MSAA_tbNextSHint',
      'Navigates to next sibling object');
    tbHelp.Hint := ini.ReadString('General', 'MSAA_tbHelpHint', 'Show online help');
    //tbHelp.Hint := tbHelp.Hint + '(' +HelpURL + ')';

    //tbMSAAMode.Hint := ini.ReadString('General', 'MSAA_tbMSAAModeHint', 'UIA mode');
    tbSetting.Hint :=
      ini.ReadString('General', 'MSAA_tbMSAASetHint', 'Show Setting Dialog');
    mnuLang.Caption := ini.ReadString('General', 'mnuLang', '&Language');
    mnuView.Caption := ini.ReadString('General', 'mnuView', '&View');
    mnuMSAA.Caption := ini.ReadString('General', 'MSAA_cbMSAA', 'MSAA');
    mnuARIA.Caption := ini.ReadString('General', 'MSAA_cbARIA', 'ARIA');
    mnuHTML.Caption := ini.ReadString('General', 'MSAA_cbHTML', 'HTML');
    mnuIA2.Caption := ini.ReadString('General', 'MSAA_cbIA2', 'IA2');
    mnuUIA.Caption := ini.ReadString('General', 'MSAA_cbUIA', 'UIAutomation');
    mnuMSAA.Hint := ini.ReadString('General', 'MSAA_cbMSAAHint', '');
    mnuARIA.Hint := ini.ReadString('General', 'MSAA_cbARIAHint', '');
    mnuHTML.Hint := ini.ReadString('General', 'MSAA_cbHTMLHint', '');
    mnuIA2.Hint := ini.ReadString('General', 'MSAA_cbIA2Hint', '');
    mnuUIA.Hint := ini.ReadString('General', 'MSAA_cbUIAHint', '');
    None := ini.ReadString('General', 'none', '(none)');
    ConvErr := ini.ReadString('General', 'MSAA_ConvertError',
      'Format is invalid: %s');
    Msg := ini.ReadString('General', 'MSAA_HookIsFailed',
      'SetWinEventHook is Failed!!');

    sTrue := ini.ReadString('General', 'true', 'True');
    sFalse := ini.ReadString('General', 'false', 'False');
    mnuTVSave.Caption := ini.ReadString('General', 'mnuSave', '&Save');
    mnuTVSAll.Caption := ini.ReadString('General', 'mnuSaveAll', '&All items');
    mnuTVSSel.Caption :=
      ini.ReadString('General', 'mnuSaveSelect', '&Selected items');

    mnuTVOpen.Caption :=
      ini.ReadString('General', 'mnuOpen', '&Open in Browser');
    mnuTVOAll.Caption := ini.ReadString('General', 'mnuOpenAll', '&All items');
    mnuTVOSel.Caption :=
      ini.ReadString('General', 'mnuOpenSelect', '&Selected items');

    mnuSave.Caption := ini.ReadString('General', 'mnuSave', '&Save');
    mnuSAll.Caption := ini.ReadString('General', 'mnuSaveAll', '&All items');
    mnuSSel.Caption :=
      ini.ReadString('General', 'mnuSaveSelect', '&Selected items');

    mnuOpenB.Caption := ini.ReadString('General', 'mnuOpen', '&Open in Browser');
    mnuOAll.Caption := ini.ReadString('General', 'mnuOpenAll', '&All items');
    mnuOSel.Caption :=
      ini.ReadString('General', 'mnuOpenSelect', '&Selected items');

    mnuSelMode.Caption :=
      ini.ReadString('General', 'mnuSelMode', 'S&elect Mode');

    tbFocus.Caption := ini.ReadString('General', 'MSAA_tbFocusName', 'Focus');
    tbCursor.Caption := ini.ReadString('General', 'MSAA_tbCursorName', 'Cursor');
    tbRectAngle.Caption :=
      ini.ReadString('General', 'MSAA_tbRectAngleName', 'Highlight Rectangle');
    tbShowtip.Caption :=
      ini.ReadString('General', 'MSAA_tbBalloonName', 'Balloon tip');
    tbCopy.Caption := ini.ReadString('General', 'MSAA_tbCopyName', 'Copy');
    tbOnlyFocus.Caption :=
      ini.ReadString('General', 'MSAA_tbFocusOnlyName', 'Focus only');
    tbParent.Caption := ini.ReadString('General', 'MSAA_tbParentName', 'Parent');
    tbChild.Caption := ini.ReadString('General', 'MSAA_tbChildName', 'Child');
    tbPrevS.Caption := ini.ReadString('General', 'MSAA_tbPrevSName', 'Previous');
    tbNextS.Caption := ini.ReadString('General', 'MSAA_tbNextSName', 'Next');
    tbHelp.Caption := ini.ReadString('General', 'MSAA_tbHelpName', 'Help');
    {tbMSAAMode.Caption :=
      ini.ReadString('General', 'MSAA_tbMSAAModeName', 'MSAA');
    tbRegister.Caption :=
      ini.ReadString('General', 'MSAA_tbMSAASetName', 'Settings');

    mnuReg.Caption := ini.ReadString('General', 'MSAA_mnuReg', '&Register');
    mnuUnReg.Caption :=
      ini.ReadString('General', 'MSAA_mnuUnreg', '&Unregister');
    mnuReg.Hint := ini.ReadString('General', 'MSAA_mnuRegHint',
      'Register IAccessible2Proxy.dll');
    mnuUnReg.Hint := ini.ReadString('General', 'MSAA_mnuUnregHint',
      'Unregister IAccessible2Proxy.dll');


    //TreeList1.AccDesc := GetLongHint(TreeList1.Hint);
    TreeView1.AccName :=
      PChar(ini.ReadString('General', 'MSAA_TVName', 'Accessibility Tree'));
    TreeView1.Hint := ini.ReadString('General', 'MSAA_TVHint', '');
    TreeView1.AccDesc := GetLongHint(TreeView1.Hint);
    UIA_fail := ini.ReadString('General', 'UIA_CreationError',
      'CoCreateInstance failed. ');

    PB1.Hint := ini.ReadString('General', 'Collapse_TV_Hint',
      'Collapse(or Expand) Button for TreeView');

    PB1.AccDesc := PB1.Hint;
    acTVCol.Caption := ini.ReadString('General', 'mnuTreeView', '&TreeView');
    acTLCol.Caption := ini.ReadString('General', 'mnuTreeList', 'Tree&List');
    acMMCol.Caption := ini.ReadString('General', 'mnuCodeEdit', '&CodeEdit');
    mnuColl.Caption := ini.ReadString('General', 'mnuCollapse', '&Collapse'); }

    mnuTVcont.Caption :=
      ini.ReadString('General', 'mnuTVContent', '&Treeview contents');
    mnuTarget.Caption := ini.ReadString('General', 'mnuTarget', '&Target only');
    mnuAll.Caption := ini.ReadString('General', 'mnuAll', '&All related objects');

    rType := ini.ReadString('IA2', 'RelationType', 'Relation Type');
    rTarg := ini.ReadString('IA2', 'RelationTargets', 'Relation Targets');
    mnuBln.Caption := ini.ReadString('General', 'mnubln', 'Balloon tip(&B)');
    mnublnMSAA.Caption :=
      ini.ReadString('General', 'mnublnMSAA', '&MS Active Accessibility');
    mnublnIA2.Caption :=
      ini.ReadString('General', 'mnublnIA2', '&IAccessible2');
    mnublnCode.Caption :=
      ini.ReadString('General', 'mnublnCode', '&Source code');

    ShowSrcLen := ini.ReadInteger('General', 'ShowSrcLen', 2000);
    ClsNames.CommaText :=
      ini.ReadString('General', 'ClassNames',
      '"mozillawindowclass","chrome_renderwidgethosthwnd","mozillawindowclass","chrome_widgetwin_0","chrome_widgetwin_1"');



    sFilter := ini.ReadString('General', 'SaveDLGFilter',
      'HTML File|*.htm*|Text File|*.txt|All|*.*');

    sHTML := ini.ReadString('HTML', 'HTML', 'HTML');
    HTMLs[0, 0] := ini.ReadString('HTML', 'Element_name', 'Element Name');
    HTMLs[1, 0] := ini.ReadString('HTML', 'Attributes', 'Attributes');
    HTMLs[2, 0] := ini.ReadString('HTML', 'Code', 'Code');
    sTypeIE := '(' + ini.ReadString('HTML', 'outerHTML', 'outerHTML') + ')';


    HTMLsFF[0, 0] := HTMLs[0, 0];
    HTMLsFF[1, 0] := HTMLs[1, 0];
    HTMLsFF[2, 0] := HTMLs[2, 0];
    HTMLs[2, 0] := HTMLs[2, 0] + StypeIE;
    sTxt := ini.ReadString('HTML', 'Text', 'Text');
    sTypeFF := '(' + ini.ReadString('HTML', 'innerHTML', 'innerHTML') + ')';

    sAria := ini.ReadString('ARIA', 'ARIA', 'ARIA');
    ARIAs[0, 0] := ini.ReadString('ARIA', 'Role', 'Role');
    ARIAs[1, 0] := ini.ReadString('ARIA', 'Attributes', 'Attributes');

    Roles[0] := ini.ReadString('IA2', 'IA2_ROLE_CANVAS', 'canvas');
    Roles[1] := ini.ReadString('IA2', 'IA2_ROLE_CAPTION', 'Caption');
    Roles[2] := ini.ReadString('IA2', 'IA2_ROLE_CHECK_MENU_ITEM',
      'Check menu item');
    Roles[3] := ini.ReadString('IA2', 'IA2_ROLE_COLOR_CHOOSER', 'Color chooser');
    Roles[4] := ini.ReadString('IA2', 'IA2_ROLE_DATE_EDITOR', 'Date editor');
    Roles[5] := ini.ReadString('IA2', 'IA2_ROLE_DESKTOP_ICON', 'Desktop icon');
    Roles[6] := ini.ReadString('IA2', 'IA2_ROLE_DESKTOP_PANE', 'Desktop pane');
    Roles[7] := ini.ReadString('IA2', 'IA2_ROLE_DIRECTORY_PANE', 'Directory pane');
    Roles[8] := ini.ReadString('IA2', 'IA2_ROLE_EDITBAR', 'Editbar');
    Roles[9] := ini.ReadString('IA2', 'IA2_ROLE_EMBEDDED_OBJECT',
      'Embedded object');
    Roles[10] := ini.ReadString('IA2', 'IA2_ROLE_ENDNOTE', 'Endnote');
    Roles[11] := ini.ReadString('IA2', 'IA2_ROLE_FILE_CHOOSER', 'File chooser');
    Roles[12] := ini.ReadString('IA2', 'IA2_ROLE_FONT_CHOOSER', 'Font chooser');
    Roles[13] := ini.ReadString('IA2', 'IA2_ROLE_FOOTER', 'Footer');
    Roles[14] := ini.ReadString('IA2', 'IA2_ROLE_FOOTNOTE', 'Footnote');
    Roles[15] := ini.ReadString('IA2', 'IA2_ROLE_FORM', 'Form');
    Roles[16] := ini.ReadString('IA2', 'IA2_ROLE_FRAME', 'Frame');
    Roles[17] := ini.ReadString('IA2', 'IA2_ROLE_GLASS_PANE', 'Glass pane');
    Roles[18] := ini.ReadString('IA2', 'IA2_ROLE_HEADER', 'Header');
    Roles[19] := ini.ReadString('IA2', 'IA2_ROLE_HEADING', 'Heading');
    Roles[20] := ini.ReadString('IA2', 'IA2_ROLE_ICON', 'Icon');
    Roles[21] := ini.ReadString('IA2', 'IA2_ROLE_IMAGE_MAP', 'Image map');
    Roles[22] := ini.ReadString('IA2', 'IA2_ROLE_INPUT_METHOD_WINDOW',
      'Input method window');
    Roles[23] := ini.ReadString('IA2', 'IA2_ROLE_INTERNAL_FRAME', 'Internal frame');
    Roles[24] := ini.ReadString('IA2', 'IA2_ROLE_LABEL', 'Label');
    Roles[25] := ini.ReadString('IA2', 'IA2_ROLE_LAYERED_PANE', 'Layered pane');
    Roles[26] := ini.ReadString('IA2', 'IA2_ROLE_NOTE', 'Note');
    Roles[27] := ini.ReadString('IA2', 'IA2_ROLE_OPTION_PANE', 'Option pane');
    Roles[28] := ini.ReadString('IA2', 'IA2_ROLE_PAGE', 'Role pane');
    Roles[29] := ini.ReadString('IA2', 'IA2_ROLE_PARAGRAPH', 'Paragraph');
    Roles[30] := ini.ReadString('IA2', 'IA2_ROLE_RADIO_MENU_ITEM',
      'Radio menu item');
    Roles[31] := ini.ReadString('IA2', 'IA2_ROLE_REDUNDANT_OBJECT',
      'Redundant object');
    Roles[32] := ini.ReadString('IA2', 'IA2_ROLE_ROOT_PANE', 'Root pane');
    Roles[33] := ini.ReadString('IA2', 'IA2_ROLE_RULER', 'Ruler');
    Roles[34] := ini.ReadString('IA2', 'IA2_ROLE_SCROLL_PANE', 'Scroll pane');
    Roles[35] := ini.ReadString('IA2', 'IA2_ROLE_SECTION', 'Section');
    Roles[36] := ini.ReadString('IA2', 'IA2_ROLE_SHAPE', 'Shape');
    Roles[37] := ini.ReadString('IA2', 'IA2_ROLE_SPLIT_PANE', 'Split pane');
    Roles[38] := ini.ReadString('IA2', 'IA2_ROLE_TEAR_OFF_MENU', 'Tear off menu');
    Roles[39] := ini.ReadString('IA2', 'IA2_ROLE_TERMINAL', 'Terminal');
    Roles[40] := ini.ReadString('IA2', 'IA2_ROLE_TEXT_FRAME', 'Text frame');
    Roles[41] := ini.ReadString('IA2', 'IA2_ROLE_TOGGLE_BUTTON', 'Toggle button');
    Roles[42] := ini.ReadString('IA2', 'IA2_ROLE_VIEW_PORT', 'View port');

    IA2Sts[0] := ini.ReadString('IA2', 'IA2_STATE_ACTIVE', 'Active');
    IA2Sts[1] := ini.ReadString('IA2', 'IA2_STATE_ARMED', 'Armed');
    IA2Sts[2] := ini.ReadString('IA2', 'IA2_STATE_DEFUNCT', 'Defunct');
    IA2Sts[3] := ini.ReadString('IA2', 'IA2_STATE_EDITABLE', 'Editable');
    IA2Sts[4] := ini.ReadString('IA2', 'IA2_STATE_HORIZONTAL', 'Horizontal');
    IA2Sts[5] := ini.ReadString('IA2', 'IA2_STATE_ICONIFIED', 'Iconified');
    IA2Sts[6] := ini.ReadString('IA2', 'IA2_STATE_INVALID_ENTRY', 'Invalid entry');
    IA2Sts[7] := ini.ReadString('IA2', 'IA2_STATE_MANAGES_DESCENDANTS',
      'Manages descendants');
    IA2Sts[8] := ini.ReadString('IA2', 'IA2_STATE_MODAL', 'Modal');
    IA2Sts[9] := ini.ReadString('IA2', 'IA2_STATE_MULTI_LINE', 'Multi line');
    IA2Sts[10] := ini.ReadString('IA2', 'IA2_STATE_OPAQUE', 'Opaque');
    IA2Sts[11] := ini.ReadString('IA2', 'IA2_STATE_REQUIRED', 'Required');
    IA2Sts[12] := ini.ReadString('IA2', 'IA2_STATE_SELECTABLE_TEXT',
      'Selectable text');
    IA2Sts[13] := ini.ReadString('IA2', 'IA2_STATE_SINGLE_LINE', 'Single line');
    IA2Sts[14] := ini.ReadString('IA2', 'IA2_STATE_STALE', 'Stale');
    IA2Sts[15] := ini.ReadString('IA2', 'IA2_STATE_SUPPORTS_AUTOCOMPLETION',
      'Supports autocompletion');
    IA2Sts[16] := ini.ReadString('IA2', 'IA2_STATE_TRANSIENT', 'Transient');
    IA2Sts[17] := ini.ReadString('IA2', 'IA2_STATE_VERTICAL', 'Vertical');


    QSFailed := ini.ReadString('IA2', 'QS_Failed', 'Query Service Failed');
    Err_Inter := ini.ReadString('General', 'Error_Interface',
      'Interface not available');

    dlgTask.Title := ini.ReadString('TaskDLG', 'title',
      'This function may take some time, do you want to continue?');
    dlgTask.Text := ini.ReadString('TaskDLG', 'text',
      'Click "OK" to acquires tree view items that are not exposed');
    dlgTask.VerificationText :=
      ini.ReadString('TaskDLG', 'showagainmsg', 'Don''t show this message again');
  finally
    Result := UIA_fail;
    ini.Free;
  end;
end;

procedure TwndMain.Load;
var
  ini: TMemIniFile;
  d, Msg, UIA_fail: string;
  b: boolean;
  ic: TIcon;
  hLib: Thandle;
  i: integer;
  mItem: TMenuItem;
  hr: HResult;
  sList: TStringList;
const
  EVENT_MIN = $00000001;
  EVENT_MAX = $7FFFFFFF;
  WINEVENT_OUTOFCONTEXT = 0;
  WINEVENT_SKIPOWNPROCESS = $0002;
begin
  hHook := SetWinEventHook(EVENT_MIN, EVENT_MAX, 0, @WinEvent, 0, 0,
    WINEVENT_OUTOFCONTEXT or WINEVENT_SKIPOWNPROCESS);
  Msg := 'SetWinEventHook is Failed!!';
  flgMSAA := 125;
  flgIA2 := 119;
  flgUIA := 50332680;
  flgUIA2 := 0;
  iDefIndex := -1;
  if fileexists(sPath) then
  begin
    sList := TStringList.Create;
    try
      sList.LoadFromFile(sPath);
      if sList.Encoding <> TEncoding.Unicode then
        sList.SaveToFile(sPath, TEncoding.Unicode);
    finally
      sList.Free;
    end;
  end;

  //iAccS := nil;
  ini := TMemIniFile.Create(SPath, TEncoding.Unicode);
  try
    if mnuLang.Visible then
    begin
      d := Ini.ReadString('Settings', 'LangFile', 'Default.ini');
      TransPath := TransDir + d;
    end;

    LoadLang;
    Width := ini.ReadInteger('Settings', 'Width', 335);
    Height := ini.ReadInteger('Settings', 'Height', 500);

    //P2W, P4H: integer;
    PageControl1.Width := ini.ReadInteger('Settings', 'P2W', 350);


    flgMSAA := ini.ReadInteger('Settings', 'flgMSAA', flgMSAA);
    flgIA2 := ini.ReadInteger('Settings', 'flgIA2', flgIA2);
    flgUIA := ini.ReadInteger('Settings', 'flgUIA', flgUIA);
    flgUIA2 := ini.ReadInteger('Settings', 'flgUIA2', flgUIA2);
    if (flgMSAA = 0) and (flgIA2 = 0) and (flgUIA = 0) and (flgUIA2 = 0) then
      flgMSAA := 1;
    ExTip := ini.ReadBool('Settings', 'ExTooltip', False);
    Top := ini.ReadInteger('Settings', 'Top', (Screen.Height - Height) div 2);
    Left := ini.ReadInteger('Settings', 'Left', (Screen.Width - Width) div 2);
    Font.Name := Ini.ReadString('Settings', 'FontName', Font.Name);
    Font.Size := ini.ReadInteger('Settings', 'FontSize', Font.Size);
    Font.Charset := ini.ReadInteger('Settings', 'Charset', 0);


    mnuMSAA.Checked := ini.ReadBool('Settings', 'vMSAA', True);
    mnuARIA.Checked := ini.ReadBool('Settings', 'vARIA', True);
    mnuHTML.Checked := ini.ReadBool('Settings', 'vHTML', True);
    mnuIA2.Checked := ini.ReadBool('Settings', 'vIA2', True);
    mnuUIA.Checked := ini.ReadBool('Settings', 'vUIA', True);

    mnublnMSAA.Checked := ini.ReadBool('Settings', 'bMSAA', True);
    mnublnIA2.Checked := ini.ReadBool('Settings', 'bIA2', True);
    mnublnCode.Checked := ini.ReadBool('Settings', 'bCode', True);

    acFocus.ShortCut := word(ini.ReadInteger('Shortcut', acFocus.Name, 115));
    acCursor.ShortCut := word(ini.ReadInteger('Shortcut', acCursor.Name, 116));
    acRect.ShortCut := word(ini.ReadInteger('Shortcut', acRect.Name, 117));
    acShowtip.ShortCut := word(ini.ReadInteger('Shortcut', acShowtip.Name, 114));
    acCopy.ShortCut := word(ini.ReadInteger('Shortcut', acCopy.Name, 118));
    acOnlyfocus.ShortCut := word(ini.ReadInteger('Shortcut', acOnlyfocus.Name, 119));
    acParent.ShortCut := word(ini.ReadInteger('Shortcut', acParent.Name, 120));
    acChild.ShortCut := word(ini.ReadInteger('Shortcut', acChild.Name, 121));
    acPrevS.ShortCut := word(ini.ReadInteger('Shortcut', acPrevS.Name, 122));
    acNextS.ShortCut := word(ini.ReadInteger('Shortcut', acNextS.Name, 123));
    acHelp.ShortCut := word(ini.ReadInteger('Shortcut', acHelp.Name, 112));
    acSetting.ShortCut := word(ini.ReadInteger('Shortcut', acSetting.Name, 8309));

    {acMSAAMode.ShortCut := word(ini.ReadInteger('Shortcut', acMSAAMode.Name, 8308));
    acTreeFocus.ShortCut :=
      word(ini.ReadInteger('Shortcut', acTreeFocus.Name, 8304));
    acListFocus.ShortCut :=
      word(ini.ReadInteger('Shortcut', acListFocus.Name, 8305));
    acMemoFocus.ShortCut :=
      word(ini.ReadInteger('Shortcut', acMemoFocus.Name, 8306));
    ac3ctrls.ShortCut := word(ini.ReadInteger('Shortcut', ac3ctrls.Name, 8307)); }

    b := ini.ReadBool('Settings', 'TVAll', False);
    mnuAll.Checked := b;
    mnuTarget.Checked := not b;
    if (not mnuMSAA.Checked) and (not mnuARIA.Checked) and
      (not mnuHTML.Checked) and (not mnuIA2.Checked) and (not mnuUIA.Checked) then
      mnuMSAA.Checked := True;
  finally
    ini.Free;
  end;
  if mnuLang.Visible then
  begin
    if LangList.Count = 0 then
      mnuLang.Enabled := False;
    for i := 0 to LangList.Count - 1 do
    begin
      if FileExists(TransDir + LangList[i]) then
      begin
        ini := TMemIniFile.Create(TransDir + LangList[i], TEncoding.UTF8);
        try

          mItem := TMenuItem.Create(Self);
          mItem.Caption := ini.ReadString('General', 'Language', 'English');
          mItem.RadioItem := True;
          mItem.GroupIndex := 1;
          //mItem.AutoCheck := True;
          mItem.OnClick := @mnuLangChildClick;
          if LowerCase(TransPath) = LowerCase(TransDir + LangList[i]) then
            mItem.Checked := True
          else
            mItem.Checked := False;
          mnuLang.Add(mItem);
        finally
          ini.Free;
        end;
      end;
    end;
    if LangList.Count = 1 then
      mnuLang.Items[0].Checked := True;
  end;

  CoInitialize(nil);
  hr := CoCreateInstance(CLASS_CUIAutomation, nil, CLSCTX_INPROC_SERVER,
    IID_IUIAutomation, UIAuto);
  if (UIAuto = nil) or (hr <> S_OK) then
  begin
    ShowMessage(UIA_fail);

    UIAuto := nil;
    mnuView.Enabled := False;
    Toolbar1.Enabled := False;
    PageControl1.Enabled := False;
    axc1.Enabled := False;

  end
  else
  begin
    //UIAuto.AddFocusChangedEventHandler(nil, self);
    hr := UIAuto.ElementFromHandle(pointer(handle), uiEle);
    if (hr = 0) and (Assigned(uiEle)) then
      uiEle.Get_CurrentProcessId(iPID);
    uiEle := nil;
  end;
  if hHook = 0 then
  begin
    ShowMessage(Msg);
    Toolbar1.Enabled := False;
    mnuSelD.Enabled := False;
  end;
end;

procedure TwndMain.FormCreate(Sender: TObject);
var
  url, onull: olevariant;
  i: integer;
  Rec: TSearchRec;
begin
  WB1 := TEvsWebBrowser.Create(Self);
  AXC1.ComServer := WB1.ComServer;
  AXC1.Active := True;
  WB1.OnStatusTextChange := @OnStatusTextChange;
  WB1.ComServer.Set_Silent(True);
  onull := NULL;
  url := 'about:blank';
  WB1.ComServer.Navigate2(url, onull, onull, onull, onull);
  APPDir := IncludeTrailingPathDelimiter(ExtractFileDir(Application.ExeName));
  TransDir := IncludeTrailingPathDelimiter(AppDir + 'Languages');
  if not DirectoryExists(TransDir) then
    TransDir := IncludeTrailingPathDelimiter(AppDir + 'Lang');

  PageControl1.ActivePageIndex := 0;
  ActTV := tvMSAA;
  mnuLang.Visible := FileExists(TransDir + 'Default.ini');
  if mnuLang.Visible then
    Transpath := TransDir + 'Default.ini'
  else
    Transpath := APPDir + ChangeFileExt(ExtractFileName(
      Application.ExeName), '.ini');
  DllPath := APPDir + 'IAccessible2Proxy.dll';
  arPT[0] := Point(0, 0);
  arPT[1] := Point(0, 0);
  arPT[2] := Point(0, 0);
  DMode := False;

  LangList := TStringList.Create;
  Created := True;
  SPath := IncludeTrailingPathDelimiter(GetMyDocPath) + 'MSAAV.ini';
  acParent.Enabled := False;
  acChild.Enabled := False;
  acPrevS.Enabled := False;
  acNextS.Enabled := False;
  bSelMode := False;
  mnuSelMode.Checked := bSelMode;
  if mnuLang.Visible then
  begin
    if (FindFirst(TransDir + '*.ini', faAnyFile, Rec) = 0) then
    begin
      repeat
        if ((Rec.Name <> '.') and (Rec.Name <> '..')) then
        begin
          if ((Rec.Attr and faDirectory) = 0) then
          begin
            LangList.Add(Rec.Name);
          end;
        end;
      until (FindNext(Rec) <> 0);
    end;
  end;
  Load;
  ExecCmdLine;
end;

procedure TwndMain.acFocusExecute(Sender: TObject);
begin
  acFocus.Checked := not acFocus.Checked;
end;

procedure TwndMain.acCursorExecute(Sender: TObject);
begin
  acCursor.Checked := not acCursor.Checked;
end;

procedure TwndMain.FormDestroy(Sender: TObject);
begin
  WB1.Free;
end;

procedure TwndMain.Timer1Timer(Sender: TObject);
var
  Wnd: HWND;
  pAcc: IAccessible;
  v: variant;
  iDis: IDispatch;
  hr: HResult;
  icPID: integer;
  tagPT: tagPoint;
begin
  if (not acCursor.Checked) or bPFunc then
    Exit;


  arPT[2] := arPT[1];
  arPT[1] := arPT[0];
  getcursorpos(arPT[0]);

  if (arPT[0].X = oldPT.X) and (arPT[0].Y = oldPT.Y) then
    Exit;
  try
    if (arPT[0].X = arPT[1].X) and (arPT[0].X = arPT[2].X) and
      (arPT[2].X = arPT[1].X) and (arPT[0].Y = arPT[1].Y) and
      (arPT[0].Y = arPT[2].Y) and (arPT[2].Y = arPT[1].Y) then
    begin
      hr := AccessibleObjectFrompoint(arPT[0], @pAcc, v);
      if (hr = 0) and (Assigned(pAcc)) then
      begin
        if ((not acOnlyFocus.Checked) and (acCursor.Checked)) then
        begin
          if SUCCEEDED(WindowFromAccessibleObject(pAcc, Wnd)) then
          begin

            tagPT.X := arPT[0].X;
            tagPT.Y := arPT[0].Y;
            hr := UIAuto.ElementFromPoint(tagPT, uiEle);
            if (hr = 0) and (Assigned(uiEle)) then
            begin
              hr := uiEle.Get_CurrentProcessId(icPID);
              if (hr = S_OK) and (icPID = iPID) then
                exit;
            end;
            Treemode := False;
            iAcc := pAcc;
            VarParent := v;

            if iAcc <> nil then
            begin
              oldPT := arPT[0];
              arPT[0] := Point(0, 0);
              arPT[1] := Point(0, 0);
              arPT[2] := Point(0, 0);
            end;

            hr := pAcc.Get_accParent(iDis);
            if (hr = 0) and (Assigned(iDis)) then
            begin
              hr := iDis.QueryInterface(IID_IACCESSIBLE, accRoot);
              if (hr = 0) and (Assigned(accRoot)) then
                accRoot := iDis as IAccessible;
            end;
            bTabEvt := False;
            if PageCOntrol1.ActivePageIndex = 0 then
            begin
              //ShowMSAAText;
              GetTreeStructure;

            end
            else
            begin
            end;
            {if not mnutvUIA.Checked then
            begin
              PageControl1.ActivePageIndex := 0;
              TabSheet1.TabVisible := True;
              TabSheet2.TabVisible := False;
              acShowTip.Enabled := True;
              ShowMSAAText;
            end
            else
            begin
              PageControl1.ActivePageIndex := 1;
              TabSheet1.TabVisible := False;
              TabSheet2.TabVisible := True;
              acShowTip.Enabled := True;
              ShowText4UIA;
            end;}
          end
          else
            iAcc := nil;
        end;
      end;
    end;
  except

  end;

end;

procedure TwndMain.tvMSAAAddition(Sender: TObject; Node: TTreeNode);
var
  Role, nTxt: string;
  ovChild, ovRole: olevariant;
  ws: WideString;
begin
  if not Assigned(Node.Data) then
    Exit;
  {if (TTreeData(Node.Data^).dummy)  then
    Exit;}

  //mnuTVSAll.Enabled := True;
  //mnuTVOAll.Enabled := True;

  Node.ImageIndex := 7;
  //Node.ExpandedImageIndex := 7;
  Node.SelectedIndex := 7;

  TTreeData(Node.Data^).Acc.Get_accName(TTreeData(Node.Data^).iID, ws);
  if ws = '' then
    nTxt := None
  else
    nTxt := string(ws);

  Role := Get_RoleText(TTreeData(Node.Data^).Acc, TTreeData(Node.Data^).iID);

  ovChild := TTreeData(Node.Data^).iID;
  nTxt := StringReplace(nTxt, #13, ' ', [rfReplaceAll]);
  nTxt := StringReplace(nTxt, #10, ' ', [rfReplaceAll]);
  Node.Text := nTxt + ' - ' + Role;
    {if (mnuAll.Checked) and (NodeTxt <> '') and (NodeTxt = Node.Text) and (not nSelected) then
    begin
      Node.Expanded := True;
      Node.Selected := True;
      nSelected := true;
    end;   }
  if SUCCEEDED(TTreeData(Node.Data^).Acc.Get_accRole(ovChild, ovRole)) then
  begin
    if VarHaveValue(ovRole) then
    begin
      if VarIsType(ovRole, VT_I4) and (TVarData(ovRole).VInteger <= 61) then
      begin
        Node.ImageIndex := TVarData(ovRole).VInteger - 1;
        //Node.ExpandedImageIndex := Node.ImageIndex;
        Node.SelectedIndex := Node.ImageIndex;
        // showmessage(inttostr(rNode.ImageIndex));
      end;
    end;
  end;

end;

procedure TwndMain.tvMSAADeletion(Sender: TObject; Node: TTreeNode);
begin
  if Assigned(Node.Data) then
    Dispose(PTreeData(Node.Data));
end;

procedure TwndMain.tvMSAAExpanding(Sender: TObject; Node: TTreeNode;
  var AllowExpansion: boolean);
var
  i, iChild, dChild: integer;
  iObtain: plongint;
  aChildren: array of TVariantArg;
  hr, hr_SCld: hResult;
  tAcc: iAccessible;
  TD: PTreeData;
  eNode: TTreeNode;
begin

  try
    if Assigned(Node.Data) and (TTreeData(Node.Data^).dummy) then
    begin
      if (Node.HasChildren) and (Node.Items[0].Text = 'avwr_dummy') then
      begin
        Node.DeleteChildren;
        TTreeData(Node.Data^).dummy := False;
        TTreeData(Node.Data^).Acc.Get_accChildCount(iChild);
        SetLength(aChildren, iChild);
        for i := 0 to iChild - 1 do
        begin
          VariantInit(olevariant(aChildren[i]));
        end;
        hr_SCld := AccessibleChildren(TTreeData(Node.Data^).Acc, 0,
          iChild, @aChildren[0], iObtain);
        if hr_SCld = 0 then
        begin
          for i := 0 to integer(iObtain) - 1 do
          begin
            if aChildren[i].vt = VT_DISPATCH then
            begin
              hr := IDispatch(aChildren[i].pdispVal).QueryInterface(
                IID_IACCESSIBLE, tAcc);
              if Assigned(tAcc) and (hr = S_OK) then
              begin
                New(TD);
                TD^.Acc := tAcc;
                TD^.UIEle := nil;
                TD^.iID := 0;
                TD^.dummy := True;

                eNode := tvMSAA.Items.AddChildObject(Node, '', Pointer(TD));
                //TBList.Add(integer(eNode.ItemId));
                tAcc.Get_accChildCount(dChild);
                if (dChild > 0) then
                begin
                  //DList.Add(integer(eNode.ItemId));
                  tvMSAA.Items.AddChild(eNode, 'avwr_dummy');
                end;
              end;
            end
            else
            begin
              New(TD);
              TD^.Acc := TTreeData(Node.Data^).Acc;
              TD^.UIEle := nil;
              TD^.iID := aChildren[i].lVal;
              TD^.dummy := False;
              eNode := tvMSAA.Items.AddChildObject(Node, '', Pointer(TD));
              //TBList.Add(integer(eNode.ItemId));
            end;
            Application.ProcessMessages;
          end;
        end;
      end;
    end;

  finally
    AllowExpansion := True;
  end;

end;

function TwndMain.MSAAText(pAcc: IAccessible = nil; TextOnly: boolean = False): string;
var
  PC: PWideChar;
  ovValue, ovChild: olevariant;
  MSAAs: array [0..10] of string;
  ws: WideString;
  i: integer;
  pDis: IDispatch;
  pa: IAccessible;
begin
  if not Assigned(iAcc) then
    Exit;
  if not mnuMSAA.Checked then
    Exit;
  //if flgMSAA = 0 then Exit;
  if pAcc = nil then
    pAcc := iAcc;
  for i := 0 to 10 do
  begin
    MSAAs[i] := none;
  end;


  ovChild := varParent;

  if (flgMSAA and 1) <> 0 then
  begin
    try
      if SUCCEEDED(pAcc.Get_accName(ovChild, ws)) then
        MSAAs[0] := string(ws);
    except
      on E: Exception do
        MSAAs[0] := E.Message;
    end;
  end;
  if (flgMSAA and 2) <> 0 then
  begin
    try
      if SUCCEEDED(pAcc.Get_accRole(ovChild, ovValue)) then
      begin
        if VarHaveValue(ovValue) then
        begin
          if VarIsNumeric(ovValue) then
          begin
            PC := WideStrAlloc(255);
            GetRoleTextW(ovValue, PC, StrBufSize(PC));
            MSAAs[1] := PC;
            StrDispose(PC);
          end
          else if VarIsStr(ovValue) then
          begin
            MSAAs[1] := VarToStr(ovValue);
          end;
        end;
      end;
    except
      on E: Exception do
        MSAAs[1] := E.Message;
    end;
  end;
  if (flgMSAA and 4) <> 0 then
  begin
    try
      if SUCCEEDED(pAcc.Get_accState(ovChild, ovValue)) then
      begin
        if VarHaveValue(ovValue) then
        begin
          if VarIsNumeric(ovValue) then
          begin
            MSAAs[2] := GetMultiState(cardinal(ovValue));
          end
          else if VarIsStr(ovValue) then
          begin
            MSAAs[2] := VarToStr(ovValue);
          end;
        end;
      end;
    except
      on E: Exception do
        MSAAs[2] := E.Message;
    end;
  end;
  if (flgMSAA and 8) <> 0 then
  begin
    try
      if SUCCEEDED(pAcc.Get_accDescription(ovChild, ws)) then
        MSAAs[3] := string(ws);
    except
      on E: Exception do
        MSAAs[3] := E.Message;
    end;
  end;

  if (flgMSAA and 16) <> 0 then
  begin
    try
      if SUCCEEDED(pAcc.Get_accDefaultAction(ovChild, ws)) then
        MSAAs[4] := string(ws);
    except
      on E: Exception do
        MSAAs[4] := E.Message;
    end;
  end;

  if (flgMSAA and 32) <> 0 then
  begin
    try
      if SUCCEEDED(pAcc.Get_accValue(ovChild, ws)) then
        MSAAs[5] := string(ws);
    except
      on E: Exception do
        MSAAs[5] := E.Message;
    end;
  end;

  if (flgMSAA and 64) <> 0 then
  begin
    try

      if SUCCEEDED(pAcc.Get_accParent(pDis)) then
      begin
        if Assigned(pDis) then
        begin
          if SUCCEEDED(pDis.QueryInterface(IID_IACCESSIBLE, pa)) then
          begin
            pa.Get_accName(CHILDID_SELF, ws);
            MSAAs[6] := string(ws);
          end;
        end;
      end;
    except
      on E: Exception do
        MSAAs[6] := E.Message;
    end;
  end;

  if (flgMSAA and 128) <> 0 then
  begin
    try
      if SUCCEEDED(pAcc.Get_accChildCount(i)) then
        MSAAs[7] := IntToStr(i);
    except
      on E: Exception do
        MSAAs[7] := E.Message;
    end;
  end;

  if (flgMSAA and 256) <> 0 then
  begin
    try
      if SUCCEEDED(pAcc.Get_accHelp(ovChild, ws)) then
        MSAAs[8] := string(ws);
    except

      on E: Exception do
        MSAAs[8] := E.Message;
    end;
  end;
  if (flgMSAA and 512) <> 0 then
  begin
    try
      if SUCCEEDED(pAcc.Get_accHelpTopic(ws, ovChild, i)) then
        if ws <> '' then
        begin
          MSAAs[9] := string(ws) + ' , ' + IntToStr(i);
        end;
    except
      on E: Exception do
        MSAAs[9] := E.Message;
    end;
  end;
  if (flgMSAA and 1024) <> 0 then
  begin
    try
      if SUCCEEDED(pAcc.Get_accKeyboardShortcut(ovChild, ws)) then
        MSAAs[10] := string(ws);
    except

      on E: Exception do
        MSAAs[10] := E.Message;
    end;
  end;

  Result := lMSAA[0] + #13#10;
  if not TextOnly then
  begin

    sBodyTxt := sBodytxt + '<h1>' + lMSAA[0] + '</h1>' + #13#10 +
      '<table><tbody>' + #13#10;
    for i := 0 to 10 do
    begin
      if i = 4 then
      begin
        if (MSAAs[i] = none) or (MSAAs[i] = '') then
          sBodyTxt := sBodytxt + '<tr><td class="name">' +
            lMSAA[i + 1] + '</td><td class="value">' + MSAAs[i] + '</td></tr>' + #13#10
        else
          sBodyTxt := sBodytxt + '<tr><td class="name">' +
            lMSAA[i + 1] +
            '<button type="button" id="exe_da">execute</button></td><td class="value">' +
            MSAAs[i] + '</td></tr>' + #13#10;
      end
      else
        sBodyTxt := sBodytxt + '<tr><td class="name">' + lMSAA[i + 1] +
          '</td><td class="value">' + MSAAs[i] + '</td></tr>' + #13#10;
    end;
    sBodyTxt := sBodytxt + '</tbody></table>';
  end;

  for i := 0 to 10 do
  begin
    if (flgMSAA and TruncPow(2, i)) <> 0 then
    begin
      Result := Result + lMSAA[i + 1] + ':' + #9 + MSAAs[i] + #13#10;
    end;
  end;
  TipText := Result + #13#10;

end;

function TwndMain.HTMLText: string;
var
  i: integer;
  regEx: TRegExpr;
  iPos: integer;
  atname, atValue, q, hAttrs, mtext: string;
begin
  if not Assigned(CEle) then
  begin
    //GetNaviState(True);

    Exit;
  end;
  for i := 0 to 2 do
    HTMLs[i, 1] := '';
  HTMLs[0, 1] := CEle.tagName;
  sBodyTxt := sBodyTxt + #13#10 + '<h1>' + sHTML + '</h1><table><tbody>';
  sBodyTxt := sBodyTxt + #13#10 + '<tr><td class="name">' + HTMLs[0, 0] +
    '</td><td class="value">' + HTMLs[0, 1] + '</td></tr>';
  HTMLs[2, 1] := CEle.outerHTML;
  regEx := TRegExpr.Create('<("[^"]*"|''[^'']*''|[^''">])*>');
  if regEx.Exec(HTMLs[2, 1]) then
  begin
    mtext := regEx.Match[0];

    {regEx.Expression := '(\S+)=[""'']?((?:.(?![""'']?\s+(?:\S+)=|[>""'']))+.)[""'']?';
    if regEx.Exec(mtext) then
    begin
      //matches := TRegEx.matches(match.Value, '(\S+)=[""'']?((?:.(?![""'']?\s+(?:\S+)=|[>""'']))+.)[""'']?');
      sBodyTxt := sBodyTxt + #13#10 + '<tr><th class="col2" colspan="2">' +
        HTMLs[1, 0] + '</th></tr><tr><td colspan="2"><table><tbody>';
      repeat


        iPos := Pos('=', regEx.Match[1]);
        atname := Copy(regEx.Match[1], 1, iPos - 1);
        atValue := Copy(regEx.Match[1], iPos + 1, Length(regEx.Match[1]));
        q := Copy(atValue, 1, 1);
        if (q = '''') or (q = '"') then
        begin
          atValue := Copy(atValue, 2, Length(atValue));
          atValue := Copy(atValue, 1, Length(atValue) - 1);
        end;
        sBodyTxt := sBodyTxt + #13#10 + '<tr><td class="name">' +
          atname + '</td><td class="value">' + atValue + '</td></tr>';

        hAttrs := hAttrs + atname + '=' + atValue + #13#10;

      until not regEx.ExecNext;
      sBodyTxt := sBodyTxt + #13#10 + '</tbody></table></td></tr>';
    end;}
  end
  else
    sBodyTxt := sBodyTxt + #13#10 + '<tr><td class="name">' + HTMLs[1, 0] +
      '</td><td class="value">' + none + '</td></tr>';

  sBodyTxt := sBodyTxt + #13#10 + '</tbody></table>';

  HTMLs[1, 1] := hAttrs;

  sCodeTxt := StringReplace(HTMLs[2, 1], '&', '&amp;', [rfReplaceAll, rfIgnoreCase]);
  sCodeTxt := StringReplace(sCodeTxt, '<', '&lttavwr;', [rfReplaceAll, rfIgnoreCase]);
  sCodeTxt := StringReplace(sCodeTxt, '>', '&gttavwr;', [rfReplaceAll, rfIgnoreCase]);
  sCodeTxt := StringReplace(sCodeTxt, '"', '&quot;', [rfReplaceAll, rfIgnoreCase]);
  sCodeTxt := StringReplace(sCodeTxt, '&lttavwr;', '<span class="tagclr">&lt;</span>',
    [rfReplaceAll, rfIgnoreCase]);
  sCodeTxt := StringReplace(sCodeTxt, '&gttavwr;', '<span class="tagclr">&gt;</span>',
    [rfReplaceAll, rfIgnoreCase]);

  sCodeTxt := '<h1>' + HTMLs[2, 0] + '</h1>' + #13#10 + '<pre><code>' +
    #13#10 + sCodeTxt + #13#10 + '</code></pre>';


  Result := sHTML + #13#10 + HTMLs[0, 0] + ':' + #9 + HTMLs[0, 1] +
    #13#10 + HTMLs[1, 0] + ':' + #9 + HTMLs[1, 1] + #13#10 + HTMLs[2, 0] +
    ':' + #9 + HTMLs[2, 1] + #13#10#13#10;

end;

function TwndMain.HTMLText4FF: string;
var

  hAttrs, Path, d: string;
  i: integer;
  aPC, aPC2: array [0..64] of pWidechar;
  aSI: array [0..64] of smallint;
  PC, PC2: PWideChar;
  SI: smallint;
  PU, PU2: PUINT;
  WD, tWD: word;
  iText: ISimpleDOMText;
  iSP: iServiceProvider;
begin

  if not Assigned(SDOM) then
  begin
    //GetNaviState(True);
    Exit;
  end;
  for i := 0 to 2 do
    HTMLsFF[i, 1] := '';
  try
    Path := IncludeTrailingPathDelimiter(ExtractFileDir(Application.ExeName));
    SDOM.get_nodeInfo(PC, SI, PC2, PU, PU2, tWD);

    if tWD <> 3 then
      HTMLsFF[0, 1] := PC
    else
      HTMLsFF[0, 1] := sTxt;
    sBodyTxt := sBodyTxt + #13#10 + '<h1>' + sHTML + '</h1><table><tbody>';
    sBodyTxt := sBodyTxt + #13#10 + '<tr><td class="name">' +
      HTMLsFF[0, 0] + '</td><td class="value">' + HTMLsFF[0, 1] + '</td></tr>';

    SDOM.get_attributes(65, aPC[0], aSI[0], aPC2[0], WD);
    if WD > 0 then
    begin
      sBodyTxt := sBodyTxt + #13#10 + '<tr><th class="col2" colspan="2">' +
        HTMLsFF[1, 0] + '</th></tr><td colspan="2"><table><tbody>';
      for i := 0 to WD - 1 do
      begin
        sBodyTxt := sBodyTxt + #13#10 + '<tr><td class="name">' +
          aPC[i] + '</td><td class="value">' + aPC2[i] + '</td></tr>';
        hattrs := hattrs + WideString(aPC[i]) + '","' + WideString(aPC2[i]) + #13#10;
      end;
      sBodyTxt := sBodyTxt + #13#10 + '</tbody></table></td></tr>';
    end
    else
    begin
      sBodyTxt := sBodyTxt + #13#10 + '<tr><td class="name">' +
        HTMLsFF[1, 0] + '</td><td class="value">' + none + '</td></tr>';
    end;
    sBodyTxt := sBodyTxt + #13#10 + '</tbody></table>';

    if hattrs = '' then
      HTMLsFF[1, 1] := none
    else
      HTMLsFF[1, 1] := hattrs;
    PC := '';
    if tWD <> 3 then
    begin
      SDOM.get_innerHTML(PC);
      d := HTMLsFF[2, 0] + StypeFF;
    end
    else
    begin
      iSP := SDOM as IServiceProvider;
      if SUCCEEDED(iSP.QueryInterface(IID_ISIMPLEDOMTEXT, iText)) then
      begin
        iText.get_domText(PC);
        d := HTMLsFF[2, 0] + '(' + sTxt + ')';
      end;
    end;

    sCodeTxt := StringReplace(PC, '&', '&amp;', [rfReplaceAll, rfIgnoreCase]);
    sCodeTxt := StringReplace(sCodeTxt, '<', '&lttavwr;', [rfReplaceAll, rfIgnoreCase]);
    sCodeTxt := StringReplace(sCodeTxt, '>', '&gttavwr;', [rfReplaceAll, rfIgnoreCase]);
    sCodeTxt := StringReplace(sCodeTxt, '"', '&quot;', [rfReplaceAll, rfIgnoreCase]);
    sCodeTxt := StringReplace(sCodeTxt, '&lttavwr;',
      '<span class="tagclr">&lt;</span>', [rfReplaceAll, rfIgnoreCase]);
    sCodeTxt := StringReplace(sCodeTxt, '&gttavwr;',
      '<span class="tagclr">&gt;</span>', [rfReplaceAll, rfIgnoreCase]);

    sCodeTxt := '<h1>' + d + '</h1>' + #13#10 + '<pre><code>' +
      #13#10 + sCodeTxt + #13#10 + '</code></pre>';

    HTMLsFF[2, 1] := PC;

    Result := sHTML + #13#10 + HTMLsFF[0, 0] + ':' + #9 + HTMLsFF[0, 1] +
      #13#10 + HTMLsFF[1, 0] + ':' + #9 + HTMLsFF[1, 1] + #13#10 +
      d + ':' + #9 + HTMLsFF[2, 1] + #13#10#13#10;

  except

  end;
end;

function TwndMain.UIAText(tEle: IUIAutomationElement = nil;
  HTMLout: boolean = False; tab: string = ''): string;
var
  iRes, i, t, iLen, iBool, iStyleID: integer;
  dblRV: double;
  ini: TMemInifile;
  ResEle: IUIAUTOMATIONELEMENT;
  UIArray: IUIAutomationElementArray;
  iaTarg: IAccessible;
  oVal: olevariant;
  ws: WideString;
  RC: UIAutomationClient_TLB.tagRECT;
  p: Pointer;
  OT: OrientationType;
  iInt: IInterface;
  UIAs: array [0 .. 59] of string;
  iLeg: IUIAutomationLegacyIAccessiblePattern;
  iRV: IUIAutomationRangeValuePattern;
  iSID: IUIAutomationStylesPattern;
  iSP: IStylesProvider;
  iIface: IInterface;
  hr: HResult;
  ND: PNodeData;

  tagPT: UIAutomationClient_TLB.tagPoint;
  cRC: TRect;
const
  IsProps: array [0 .. 19] of integer = (30027, 30028, 30029, 30030, 30031,
    30032, 30033, 30034, 30035, 30036, 30037, 30038, 30039, 30040, 30041, 30042,
    30043, 30044, 30108, 30109);

  function GetMSAA: IAccessible;
  var
    rIacc: IAccessible;
  begin
    Result := nil;
    if SUCCEEDED(ResEle.GetCurrentPattern(10018, iInt)) then
    begin
      if SUCCEEDED(iInt.QueryInterface(IID_IUIAutomationLegacyIAccessiblePattern,
        iLeg)) then
      begin
        if SUCCEEDED(iLeg.GetIAccessible(rIacc)) then
          Result := rIacc;
      end;
    end;

  end;

  function PropertyPatternIS(iID: integer): string;
  begin

    case iID of
      30001:
        Result := 'rect';
      30002, 30012:
        Result := 'int';
      30003:
        Result := 'controltypeid';
      30008, 30009, 30010, 30016, 30017, 30019, 30022, 30025, 30103:
        Result := 'bool';
      // 30018: Result := 'iuiautomationelement';
      30020:
        Result := 'int'; // 'uia_hwnd';
      // 30104, 30105, 30106: Result := 'iuiautomationelementarray';
      30023:
        Result := 'orientationtype';
      else
        Result := 'str';
    end;
  end;

  function GetCID(iID: integer): string;
  begin
    ini := TMemInifile.Create(TransPath, TEncoding.UTF8);
    try
      Result := ini.ReadString('UIA', IntToStr(iID), None);
    finally
      ini.Free;
    end;
  end;

  function GetOT: string;
  begin
    ini := TMemInifile.Create(TransPath, TEncoding.UTF8);
    try
      if OT = OrientationType_Horizontal then
        Result := ini.ReadString('UIA', 'OrientationType_Horizontal', 'Horizontal')
      else if OT = OrientationType_Vertical then
        Result := ini.ReadString('UIA', 'OrientationType_Vertical', 'Vertical')
      else
        Result := ini.ReadString('UIA', 'OrientationType_None', 'None');
    finally
      ini.Free;
    end;
  end;

begin
  if not Assigned(UIAuto) then
    Exit;
  if not Assigned(UIEle) then
    Exit;

  if tEle = nil then
    tEle := uiEle;
  if (not Assigned(tEle)) then
    Exit;

  try
    if (Assigned(uiEle)) then
    begin
      for i := 0 to 32 do
        UIAs[i] := None;

      if (flgUIA and TruncPow(2, 0)) <> 0 then
      begin
        try
          if SUCCEEDED(tEle.Get_CurrentAcceleratorKey(ws)) then
          begin
            UIAs[0] := ws;
          end;
        except
          on E: Exception do
            UIAs[0] := E.Message;
        end;
      end;
      if (flgUIA and TruncPow(2, 1)) <> 0 then
      begin
        try
          if SUCCEEDED(tEle.Get_CurrentAccessKey(ws)) then
          begin
            UIAs[1] := ws;
          end;
        except
          on E: Exception do
            UIAs[1] := E.Message;
        end;
      end;
      if (flgUIA and TruncPow(2, 2)) <> 0 then
      begin
        try

          if SUCCEEDED(tEle.Get_CurrentAriaProperties(ws)) then
          begin
            UIAs[2] := ws;
          end;

        except
          on E: Exception do
            UIAs[2] := E.Message;
        end;
      end;
      if (flgUIA and TruncPow(2, 3)) <> 0 then
      begin
        try
          VarClear(oVal);
          TVarData(oVal).VType := varString;
          tEle.Get_CurrentAriaRole(ws);
          if ws <> '' then
            UIAs[3] := ws;
        except
          on E: Exception do
            UIAs[3] := E.Message;
        end;
      end;
      if (flgUIA and TruncPow(2, 4)) <> 0 then
      begin
        try
          if SUCCEEDED(tEle.Get_CurrentAutomationId(ws)) then
          begin
            UIAs[4] := ws;
          end;
        except
          on E: Exception do
            UIAs[4] := E.Message;
        end;
      end;
      if (flgUIA and TruncPow(2, 5)) <> 0 then
      begin
        try
          if SUCCEEDED(tEle.Get_CurrentBoundingRectangle(RC)) then
          begin
            UIAs[5] := IntToStr(RC.left) + ',' + IntToStr(RC.top) +
              ',' + IntToStr(RC.right) + ',' + IntToStr(RC.bottom);
          end;
        except
          on E: Exception do
            UIAs[5] := E.Message;
        end;
      end;
      if (flgUIA and TruncPow(2, 6)) <> 0 then
      begin
        try
          if SUCCEEDED(tEle.Get_CurrentClassName(ws)) then
          begin
            UIAs[6] := ws;
          end;
        except
          on E: Exception do
            UIAs[6] := E.Message;
        end;
      end;
      if (flgUIA and TruncPow(2, 7)) <> 0 then
      begin
        try
          if SUCCEEDED(tEle.Get_CurrentControlType(i)) then
          begin
            UIAs[7] := GetCID(i);
          end;
        except
          on E: Exception do
            UIAs[7] := E.Message;
        end;
      end;
      if (flgUIA and TruncPow(2, 8)) <> 0 then
      begin
        try
          if SUCCEEDED(tEle.Get_CurrentCulture(i)) then
          begin
            UIAs[8] := IntToStr(i);
          end;
        except
          on E: Exception do
            UIAs[8] := E.Message;
        end;
      end;
      if (flgUIA and TruncPow(2, 9)) <> 0 then
      begin
        try
          if SUCCEEDED(tEle.Get_CurrentFrameWorkID(ws)) then
          begin
            UIAs[9] := ws;
          end;
        except
          on E: Exception do
            UIAs[9] := E.Message;
        end;
      end;
      if (flgUIA and TruncPow(2, 10)) <> 0 then
      begin
        try
          if SUCCEEDED(tEle.Get_CurrentHasKeyboardFocus(i)) then
          begin
            UIAs[10] := IntToBoolStr(i);
          end;
        except
          on E: Exception do
            UIAs[10] := E.Message;
        end;
      end;
      if (flgUIA and TruncPow(2, 11)) <> 0 then
      begin
        try
          if SUCCEEDED(tEle.Get_CurrentHelpText(ws)) then
          begin
            UIAs[11] := ws;
          end;
        except
          on E: Exception do
            UIAs[11] := E.Message;
        end;
      end;
      if (flgUIA and TruncPow(2, 12)) <> 0 then
      begin
        try
          if SUCCEEDED(tEle.Get_CurrentIsControlElement(i)) then
          begin
            UIAs[12] := IntToBoolStr(i);
          end;
        except
          on E: Exception do
            UIAs[12] := E.Message;
        end;
      end;
      if (flgUIA and TruncPow(2, 13)) <> 0 then
      begin
        try
          if SUCCEEDED(tEle.Get_CurrentIsContentElement(i)) then
          begin
            UIAs[13] := IntToBoolStr(i);
          end;
        except
          on E: Exception do
            UIAs[13] := E.Message;
        end;
      end;
      if (flgUIA and TruncPow(2, 14)) <> 0 then
      begin
        try
          if SUCCEEDED(tEle.Get_CurrentIsDataValidForForm(i)) then
          begin
            UIAs[14] := IntToBoolStr(i);
          end;
        except
          on E: Exception do
            UIAs[14] := E.Message;
        end;
      end;
      if (flgUIA and TruncPow(2, 15)) <> 0 then
      begin
        try
          if SUCCEEDED(tEle.Get_CurrentIsEnabled(i)) then
          begin
            UIAs[15] := IntToBoolStr(i);
          end;
        except
          on E: Exception do
            UIAs[15] := E.Message;
        end;
      end;
      if (flgUIA and TruncPow(2, 16)) <> 0 then
      begin
        try
          if SUCCEEDED(tEle.Get_CurrentIsKeyboardFocusable(i)) then
          begin
            UIAs[16] := IntToBoolStr(i);
          end;
        except
          on E: Exception do
            UIAs[16] := E.Message;
        end;
      end;
      if (flgUIA and TruncPow(2, 17)) <> 0 then
      begin
        try
          if SUCCEEDED(tEle.Get_CurrentIsOffscreen(i)) then
          begin
            UIAs[17] := IntToBoolStr(i);
          end;
        except
          on E: Exception do
            UIAs[17] := E.Message;
        end;
      end;
      if (flgUIA and TruncPow(2, 18)) <> 0 then
      begin
        try
          if SUCCEEDED(tEle.Get_CurrentIsPassword(i)) then
          begin
            UIAs[18] := IntToBoolStr(i);
          end;
        except
          on E: Exception do
            UIAs[18] := E.Message;
        end;
      end;
      if (flgUIA and TruncPow(2, 19)) <> 0 then
      begin
        try
          if SUCCEEDED(tEle.Get_CurrentIsRequiredForForm(i)) then
          begin
            UIAs[19] := IntToBoolStr(i);
          end;
        except
          on E: Exception do
            UIAs[19] := E.Message;
        end;
      end;
      if (flgUIA and TruncPow(2, 20)) <> 0 then
      begin
        try
          if SUCCEEDED(tEle.Get_CurrentItemStatus(ws)) then
          begin
            UIAs[20] := ws;
          end;
        except
          on E: Exception do
            UIAs[20] := E.Message;
        end;
      end;
      if (flgUIA and TruncPow(2, 21)) <> 0 then
      begin
        try
          if SUCCEEDED(tEle.Get_CurrentItemType(ws)) then
          begin
            UIAs[21] := ws;
          end;
        except
          on E: Exception do
            UIAs[21] := E.Message;
        end;
      end;
      if (flgUIA and TruncPow(2, 22)) <> 0 then
      begin
        try
          if SUCCEEDED(tEle.Get_CurrentLocalizedControlType(ws)) then
          begin
            UIAs[22] := ws;
          end;
        except
          on E: Exception do
            UIAs[22] := E.Message;
        end;
      end;
      if (flgUIA and TruncPow(2, 23)) <> 0 then
      begin
        try
          if SUCCEEDED(tEle.Get_CurrentName(ws)) then
          begin
            UIAs[23] := ws;
          end;
        except
          on E: Exception do
            UIAs[23] := E.Message;
        end;
      end;
      if (flgUIA and TruncPow(2, 24)) <> 0 then
      begin
        try
          if SUCCEEDED(tEle.Get_CurrentNativeWindowHandle(p)) then
          begin
            UIAs[24] := IntToStr(integer(p));
          end;
        except
          on E: Exception do
            UIAs[24] := E.Message;
        end;
      end;
      if (flgUIA and TruncPow(2, 25)) <> 0 then
      begin
        try
          if SUCCEEDED(tEle.Get_CurrentOrientation(OT)) then
          begin

            UIAs[25] := GetOT;
          end;
        except
          on E: Exception do
            UIAs[25] := E.Message;
        end;
      end;
      if (flgUIA and TruncPow(2, 26)) <> 0 then
      begin
        try
          if SUCCEEDED(tEle.Get_CurrentProcessId(i)) then
          begin
            UIAs[26] := IntToStr(i);
          end;
        except
          on E: Exception do
            UIAs[26] := E.Message;
        end;
      end;
      if (flgUIA and TruncPow(2, 27)) <> 0 then
      begin
        try
          if SUCCEEDED(tEle.Get_CurrentProviderDescription(ws)) then
          begin
            UIAs[27] := ws;
          end;
        except
          on E: Exception do
            UIAs[27] := E.Message;
        end;
      end;



      if (flgUIA2 and TruncPow(2, 1)) <> 0 then
      begin
        try // LiveSetting
          if SUCCEEDED(tEle.GetCurrentPropertyValue(30135, oVal)) then
          begin
            if VarHaveValue(oVal) then
            begin
              if VarIsStr(oVal) or VarIsNumeric(oVal) then
                UIAs[32] := oVal;
            end;
          end;
        except
          on E: Exception do
            UIAs[32] := E.Message;
        end;
      end;

      for i := 0 to 19 do
      begin
        if (flgUIA2 and TruncPow(2, i + 2)) <> 0 then
        begin
          try // LiveSetting
            if SUCCEEDED(tEle.GetCurrentPropertyValue(IsProps[i], oVal)) then
            begin
              if VarHaveValue(oVal) then
              begin
                if VarIsType(oVal, varBoolean) then
                begin
                  iBool := oVal;
                  UIAs[33 + i] := strUtils.IfThen(iBool <> 0, sTrue, sFalse);
                end;
              end;
            end;
          except
            on E: Exception do
              UIAs[33 + i] := E.Message;
          end;
        end;
      end;
      if (flgUIA2 and TruncPow(2, 22)) <> 0 then
      begin
        // #30047~30052
        // GetCurrentPatternAs http://msdn.microsoft.com/en-us/library/windows/desktop/ee696039(v=vs.85).aspx
        // UIA_RangeValuePatternId 10003
        try
          hr := tEle.GetCurrentPattern(UIA_RangeValuePatternId, iIface);
          if hr = S_OK then
          begin
            if Assigned(iIface) then
            begin
              if SUCCEEDED(iIface.QueryInterface(
                IID_IUIAutomationRangeValuePattern, iRV)) then
              begin
                if SUCCEEDED(iRV.Get_currentValue(dblRV)) then
                  UIAs[53] := Floattostr(dblRV);
                if SUCCEEDED(iRV.Get_CurrentIsReadOnly(iBool)) then
                  UIAs[54] := strUtils.IfThen(iBool <> 0, sTrue, sFalse);
                if SUCCEEDED(iRV.Get_CurrentMinimum(dblRV)) then
                  UIAs[55] := Floattostr(dblRV);
                if SUCCEEDED(iRV.Get_CurrentMaximum(dblRV)) then
                  UIAs[56] := Floattostr(dblRV);
                if SUCCEEDED(iRV.Get_CurrentLargeChange(dblRV)) then
                  UIAs[57] := Floattostr(dblRV);
                if SUCCEEDED(iRV.Get_CurrentSmallChange(dblRV)) then
                  UIAs[58] := Floattostr(dblRV);
              end;
            end;
          end;
        except
          on E: Exception do
            UIAs[53] := E.Message;
        end;
      end;
      if (flgUIA2 and TruncPow(2, 23)) <> 0 then
      begin
        try
          hr := tEle.GetCurrentPattern(UIA_StylesPatternId, iIface);
          if (hr = S_OK) and Assigned(iIface) then
          begin
            hr := iIface.QueryInterface(IID_IUIAutomationStylesPattern, iSID);
            if (hr = S_OK) and Assigned(iSP) then
              // hr := iiFace.QueryInterface(IID_IStylesProvider, iSP);
              // if (hr = S_OK) and Assigned(iSP) then
            begin
              hr := iSID.Get_CurrentStyleId(iStyleID);
              // outputdebugstring(PWideChar('current style id is ' + inttostr(istyleid)));
              // hr := iSP.get_StyleId(iStyleID);
              if (hr = S_OK)
              { and ((iStyleID >= 70000) and (iStyleID <= 70016)) } then
              begin
                // UIAs[59] := StyleID[iStyleID - 70000];
                UIAs[59] := IntToStr(iStyleID);
              end;
            end;
          end;
        except
          on E: Exception do
            UIAs[59] := E.Message;
        end;
      end;
      if HTMLout then
        Result := Result + #13#10 + tab + #9 + '<strong>' + lUIA[0] +
          '</strong>' + #13#10#9 + tab + '<ul>'
      else
      begin
        Result := lUIA[0] + #13#10;
        sBodyTxt := sBodyTxt + #13#10 + '<h1>' + lUIA[0] + '</h1><table><tbody>';
      end;

      for i := 0 to 54 do
      begin

        if i < 31 then
        begin
          if (flgUIA and TruncPow(2, i)) <> 0 then
          begin
            if i >= 28 then
            begin
              if (not HTMLout) then
                sBodyTxt := sBodyTxt + #13#10 + '<tr><th colspan="2">' +
                  lUIA[i + 1] + '</th></tr><tr><td colspan="2"><table><tbody>';
              UIArray := nil;
              iRes := E_FAIL;
              if i = 28 then
              begin
                iRes := tEle.Get_CurrentControllerFor(UIArray);
              end
              else if i = 29 then
              begin
                iRes := tEle.Get_CurrentDescribedBy(UIArray);
              end
              else if i = 30 then
              begin
                iRes := tEle.Get_CurrentFlowsTo(UIArray);
              end;
              if Assigned(UIArray) and SUCCEEDED(iRes) then
              begin
                if SUCCEEDED(UIArray.Get_Length(iLen)) then
                begin
                  for t := 0 to iLen - 1 do
                  begin
                    // cNode := nodes.AddChild(Node, inttostr(t));

                    ResEle := nil;
                    if SUCCEEDED(UIArray.GetElement(t, ResEle)) then
                    begin
                      if Assigned(ResEle) then
                      begin
                        ws := None;
                        if (Pagecontrol1.ActivePageIndex = 0) and (not HTMLOut) then
                        begin
                          iaTarg := GetMSAA;
                          if Assigned(iaTarg) then
                          begin
                            iaTarg.Get_accName(0, ws);

                          end;
                        end
                        else
                        begin
                          // UIA_LegacyIAccessibleNamePropertyId = 30092
                          ResEle.GetCurrentPropertyValue(30092, oVal);
                          ws := string(oVal);
                        end;
                        if not HTMLout then
                        begin
                          Result := Result + lUIA[i + 1] + ':' + #9 + ws + #13#10;
                          sBodyTxt :=
                            sBodyTxt + #13#10 + '<tr><td class="name">' +
                            IntToStr(t) + '</td><td class="value">' + ws + '</td></tr>';
                        end
                        else
                        begin

                          Result :=
                            Result + #13#10#9#9 + tab + '<li>' +
                            lUIA[i + 1] + ': ' + ws + ';</li>';
                        end;
                      end;
                    end;
                  end;

                end;
              end
              else
              begin
                if (not HTMLout) then
                begin
                  sBodyTxt := sBodyTxt + #13#10 +
                    '<tr><td class="name"></td><td class="value">' + none + '</td></tr>';
                end;
              end;
              sBodyTxt := sBodyTxt + #13#10 + '</tbody></table></td></tr>';
            end
            else
            begin
              if (not HTMLout) then
              begin
                sBodyTxt := sBodyTxt + #13#10 + '<tr><td class="name">' +
                  lUIA[i + 1] + '</td><td class="value">' + UIAs[i] + '</td></tr>';
                Result := Result + lUIA[i + 1] + ':' + #9 + UIAs[i] + #13#10;
              end
              else
              begin
                Result := Result + #13#10#9#9 + tab + '<li>' +
                  lUIA[i + 1] + ': ' + UIAs[i] + ';</li>';
              end;
            end;
          end;
        end
        else
        begin
          if (flgUIA2 and TruncPow(2, i - 31)) <> 0 then
          begin
            if i = 31 then
            begin
              try
                iRes := tEle.Get_CurrentLabeledBy(ResEle);
                if SUCCEEDED(iRes) then
                begin

                  if not SUCCEEDED(ResEle.Get_CurrentName(ws)) then
                    ws := None;

                  if (not HTMLout) then
                  begin
                    sBodyTxt :=
                      sBodyTxt + #13#10 + '<tr><td class="name">' +
                      lUIA[i + 1] + '</td><td class="value">' + ws + '</td></tr>';
                    Result := Result + lUIA[i + 1] + ':' + #9 + ws + #13#10;
                  end
                  else
                  begin
                    Result := Result + #13#10#9#9 + tab + '<li>' +
                      lUIA[i + 1] + ': ' + ws + ';</li>';
                  end;
                end
                else
                if (not HTMLout) then
                begin
                  sBodyTxt :=
                    sBodyTxt + #13#10 + '<tr><td class="name">' +
                    lUIA[i + 1] + '</td><td class="value">' + none + '</td></tr>';
                end;
              except
                on E: Exception do
                  // UIAs[31] := E.Message;
              end;
            end
            else if i = 53 then
            begin
              // node := nodes.AddChild(rNode, lUIA[54]);
              if (not HTMLout) then
              begin
                sBodyTxt := sBodyTxt + #13#10 + '<tr><th colspan="2">' +
                  lUIA[54] + '</th></tr><tr><td colspan="2"><table><tbody>';
                Result := Result + lUIA[54];
              end
              else
              begin
                Result := Result + #13#10#9#9 + tab + '<li>' + lUIA[54] +
                  ';' + #13#10#9#9#9 + tab + '<ul>';
              end;
              try
                for t := 0 to 5 do
                begin
                  if (not HTMLout) then
                  begin
                    sBodyTxt :=
                      sBodyTxt + #13#10 + '<tr><td class="name">' +
                      lUIA[55 + t] + '</td><td class="value">' +
                      UIAs[53 + t] + '</td></tr>';
                    Result := Result + #9 + lUIA[55 + t] + ':' + UIAs[53 + t] + #13#10;
                  end
                  else
                  begin
                    Result := Result + #13#10#9#9#9 + tab + '<li>' +
                      lUIA[55 + t] + ': ' + UIAs[53 + t] + ';</li>';
                  end;
                end;
                if HTMLout then
                  Result := Result + #13#10#9#9#9 + tab + '</ul>' +
                    #13#10#9#9 + tab + '</li>'
                else
                  sBodyTxt := sBodyTxt + #13#10 + '</tbody></table></td></tr>';
              except
                on E: Exception do
                  // UIAs[31] := E.Message;
              end;
            end
            else if i = 54 then
            begin
              if (not HTMLout) then
              begin
                sBodyTxt := sBodyTxt + #13#10 + '<tr><td class="name">' +
                  lUIA[61] + '</td><td class="value">' + UIAs[59] + '</td></tr>';
                Result := Result + lUIA[61] + ':' + #9 + UIAs[59] + #13#10;
              end
              else
              begin
                Result := Result + #13#10#9#9 + tab + '<li>' + lUIA[61] +
                  ': ' + UIAs[59] + ';</li>';
              end;
            end
            else
            begin
              if (not HTMLout) then
              begin
                sBodyTxt := sBodyTxt + #13#10 + '<tr><td class="name">' +
                  lUIA[i + 1] + '</td><td class="value">' + UIAs[i] + '</td></tr>';
                Result := Result + lUIA[i + 1] + ':' + #9 + UIAs[i] + #13#10;
              end
              else
              begin
                Result := Result + #13#10#9#9 + tab + '<li>' +
                  lUIA[i + 1] + ': ' + UIAs[i] + ';</li>';
              end;
            end;
          end;
        end;
      end;
      if not HTMLout then
        sBodyTxt := sBodyTxt + #13#10 + '</tbody></table>';
    end;
  finally
  end;

end;

function TwndMain.GetIA2Role(iRole: integer): string;
var
  PC: PWideChar;
  ovValue: olevariant;

begin

  case iRole of
    1025..1067: Result := Roles[iRole - 1025];
    else
    begin
      PC := WideStrAlloc(255);
      ovValue := iROle;
      GetRoleTextW(ovValue, PC, StrBufSize(PC));
      Result := PC;
      StrDispose(PC);
    end;
  end;
end;

function TwndMain.GetIA2State(iRole: integer): string;
begin
  Result := '';

  if (iRole and IA2_STATE_ACTIVE) <> 0 then
  begin
    if Result <> '' then
      Result := Result + ', ';
    Result := Result + IA2Sts[0];
  end;

  if (iRole and IA2_STATE_ARMED) <> 0 then
  begin
    if Result <> '' then
      Result := Result + ', ';
    Result := Result + IA2Sts[1];
  end;

  if (iRole and IA2_STATE_DEFUNCT) <> 0 then
  begin
    if Result <> '' then
      Result := Result + ', ';
    Result := Result + IA2Sts[2];
  end;

  if (iRole and IA2_STATE_EDITABLE) <> 0 then
  begin
    if Result <> '' then
      Result := Result + ', ';
    Result := Result + IA2Sts[3];
  end;

  if (iRole and IA2_STATE_HORIZONTAL) <> 0 then
  begin
    if Result <> '' then
      Result := Result + ', ';
    Result := Result + IA2Sts[4];
  end;

  if (iRole and IA2_STATE_ICONIFIED) <> 0 then
  begin
    if Result <> '' then
      Result := Result + ', ';
    Result := Result + IA2Sts[5];
  end;

  if (iRole and IA2_STATE_INVALID_ENTRY) <> 0 then
  begin
    if Result <> '' then
      Result := Result + ', ';
    Result := Result + IA2Sts[6];
  end;

  if (iRole and IA2_STATE_MANAGES_DESCENDANTS) <> 0 then
  begin
    if Result <> '' then
      Result := Result + ', ';
    Result := Result + IA2Sts[7];
  end;

  if (iRole and IA2_STATE_MODAL) <> 0 then
  begin
    if Result <> '' then
      Result := Result + ', ';
    Result := Result + IA2Sts[8];
  end;

  if (iRole and IA2_STATE_MULTI_LINE) <> 0 then
  begin
    if Result <> '' then
      Result := Result + ', ';
    Result := Result + IA2Sts[9];
  end;

  if (iRole and IA2_STATE_OPAQUE) <> 0 then
  begin
    if Result <> '' then
      Result := Result + ', ';
    Result := Result + IA2Sts[10];
  end;

  if (iRole and IA2_STATE_REQUIRED) <> 0 then
  begin
    if Result <> '' then
      Result := Result + ', ';
    Result := Result + IA2Sts[11];
  end;

  if (iRole and IA2_STATE_SELECTABLE_TEXT) <> 0 then
  begin
    if Result <> '' then
      Result := Result + ', ';
    Result := Result + IA2Sts[12];
  end;

  if (iRole and IA2_STATE_SINGLE_LINE) <> 0 then
  begin
    if Result <> '' then
      Result := Result + ', ';
    Result := Result + IA2Sts[13];
  end;

  if (iRole and IA2_STATE_STALE) <> 0 then
  begin
    if Result <> '' then
      Result := Result + ', ';
    Result := Result + IA2Sts[14];
  end;

  if (iRole and IA2_STATE_SUPPORTS_AUTOCOMPLETION) <> 0 then
  begin
    if Result <> '' then
      Result := Result + ', ';
    Result := Result + IA2Sts[15];
  end;

  if (iRole and IA2_STATE_TRANSIENT) <> 0 then
  begin
    if Result <> '' then
      Result := Result + ', ';
    Result := Result + IA2Sts[16];
  end;

  if (iRole and IA2_STATE_VERTICAL) <> 0 then
  begin
    if Result <> '' then
      Result := Result + ', ';
    Result := Result + IA2Sts[17];
  end;

end;

function TwndMain.SetIA2Text(pAcc: IAccessible = nil; SetTL: boolean = True): string;
var
  iSP: IServiceProvider;
  iInter: IInterface;
  ia2, ia2Targ: IAccessible2;
  iAV: IAccessibleValue;
  iaTarg: IAccessible;

  hr, i, iRole, t, p, iTarg, iPos, iUID: integer;
  MSAAs: array [0 .. 13] of WideString;
  ovChild, ovValue: olevariant;
  iAL: IAccessibleRelation;
  s, ws: WideString;
  oList, cList: TStringList;
  iSOffset, IEOffset: integer;
  ia2Txt: IAccessibleText;
begin
  Result := '';
  if not Assigned(iAcc) then
    Exit;
  if pAcc = nil then
    pAcc := iAcc;
  if (not mnuMSAA.Checked) then
    Exit;

  // if flgIA2 = 0 then Exit;
  oList := TStringList.Create;
  try

    hr := pAcc.QueryInterface(IID_IServiceProvider, iSP);
    if SUCCEEDED(hr) and (Assigned(iSP)) then
    begin
      try
        hr := iSP.QueryService(IID_IACCESSIBLE, IID_IACCESSIBLE2, ia2);

      except
        hr := E_FAIL;
      end;
      if SUCCEEDED(hr) and Assigned(ia2) then
      begin
        for i := 0 to 8 do
        begin
          MSAAs[i] := None;
        end;

        ovChild := VarParent;
        if (flgIA2 and TruncPow(2, 0)) <> 0 then
        begin
          try
            if SUCCEEDED(ia2.Get_accName(ovChild, ws)) then
            begin
              MSAAs[0] := ws;
            end;
          except
            on E: Exception do
            begin
              MSAAs[0] := E.Message;

            end;
          end;
        end;

        if (flgIA2 and TruncPow(2, 1)) <> 0 then
        begin
          try
            if SUCCEEDED(ia2.Role(iRole)) then
            begin
              MSAAs[1] := GetIA2Role(iRole);
              ;
            end;
          except
            on E: Exception do
              MSAAs[1] := E.Message;
          end;
        end;
        if (flgIA2 and TruncPow(2, 2)) <> 0 then
        begin
          try
            if SUCCEEDED(ia2.Get_states(iRole)) then
            begin
              MSAAs[2] := GetIA2State(iRole);
            end;
          except
            on E: Exception do
              MSAAs[2] := E.Message;
          end;
        end;
        if (flgIA2 and TruncPow(2, 3)) <> 0 then
        begin
          try
            if SUCCEEDED(ia2.Get_accDescription(ovChild, ws)) then
            begin
              MSAAs[3] := ws;
            end;
          except
            on E: Exception do
              MSAAs[3] := E.Message;
          end;
        end;

        if (flgIA2 and TruncPow(2, 5)) <> 0 then
        begin

          cList := TStringList.Create;
          try

            try
              if SUCCEEDED(ia2.Get_attributes(s)) then
              begin
                cList.Delimiter := ';';
                cList.DelimitedText := s;
                for i := 0 to cList.Count - 1 do
                begin
                  if cList[i] = '' then
                    continue;

                  if i = 0 then
                    MSAAs[5] := cList[i]
                  else
                    MSAAs[5] := MSAAs[5] + ', ' + cList[i];
                  oList.Add(cList[i]);

                end;
              end;
            except
              on E: Exception do
              begin
                MSAAs[5] := MSAAs[5] + E.Message;
                oList.Add(E.Message);
              end;
            end;
          finally
            cList.Free;
          end;

        end;

        if (flgIA2 and TruncPow(2, 6)) <> 0 then
        begin
          try
            if SUCCEEDED(ia2.Get_accValue(ovChild, ws)) then
            begin
              MSAAs[6] := ws;
            end;
          except
            on E: Exception do
              MSAAs[6] := E.Message;
          end;
        end;

        if (flgIA2 and TruncPow(2, 7)) <> 0 then
        begin
          try
            if SUCCEEDED(ia2.Get_localizedExtendedRole(ws)) then
            begin
              MSAAs[7] := ws;
            end;
          except
            on E: Exception do
              MSAAs[7] := E.Message;
          end;
        end;
        if (flgIA2 and TruncPow(2, 8)) <> 0 then
        begin
          try
            if SUCCEEDED(ia2.Get_localizedExtendedStates(10, PWideString1(ws),
              iRole)) then
            begin
              MSAAs[8] := ws;
            end;
          except
            on E: Exception do
              MSAAs[8] := E.Message;
          end;
        end;

        if (flgIA2 and TruncPow(2, 9)) <> 0 then
        begin
          try
            if SUCCEEDED(iSP.QueryService(iid_iaccessiblevalue,
              iid_iaccessiblevalue, iAV)) then
            begin
              if SUCCEEDED(iAV.Get_currentValue(ovValue)) then
                MSAAs[9] := ovValue;
              if SUCCEEDED(iAV.Get_minimumValue(ovValue)) then
                MSAAs[10] := ovValue;
              if SUCCEEDED(iAV.Get_maximumValue(ovValue)) then
                MSAAs[11] := ovValue;
            end;
          except
            on E: Exception do
              MSAAs[9] := E.Message;
          end;
        end;

        if (flgIA2 and TruncPow(2, 10)) <> 0 then
        begin
          try
            if SUCCEEDED(ia2.QueryInterface(IID_IAccessibleText, ia2Txt)) then
            begin

              ia2Txt.Get_attributes(1, iSOffset, IEOffset, ws);
              MSAAs[12] := SysUtils.StringReplace(ws, '\,', ',',
                [rfReplaceAll, rfIgnoreCase]);

            end;
          except
            on E: Exception do
              MSAAs[12] := E.Message;
          end;
        end;

        if (flgIA2 and TruncPow(2, 11)) <> 0 then
        begin
          try
            if SUCCEEDED(ia2.Get_uniqueID(iUID)) then
            begin
              MSAAs[13] := IntToStr(iUID);

            end;
          except
            on E: Exception do
              MSAAs[13] := E.Message;
          end;
        end;
        Result := lIA2[0] + #13#10;
        // nodes := TreeList1.Items;
        if SetTL then
        begin
          sBodyTxt := sBodyTxt + #13#10 + '<h1>' + lIA2[0] + '</h1><table><tbody>';
        end;
        for i := 0 to 11 do
        begin
          if (flgIA2 and TruncPow(2, i)) <> 0 then
          begin
            if i = 5 then
            begin
              Result := Result + lIA2[i + 1] + ':' + #9 + MSAAs[i] + '(';
              // + #13#10;
              if SetTL then
              begin
                sBodyTxt := sBodyTxt + #13#10 + '<tr><th class="col2" colspan="2">' +
                  lIA2[i + 1] + '</th></tr><tr><td colspan="2"><table><tbody>';

                for p := 0 to oList.Count - 1 do
                begin
                  t := pos(':', oList[p]);
                  sBodyTxt := sBodyTxt + #13#10 + '<tr><td class="name">' +
                    copy(oList[p], 0, t - 1) + '</td><td class="value">' +
                    copy(oList[p], t + 1, Length(oList[p])) + '</td></tr>';

                end;
                sBodyTxt := sBodyTxt + #13#10 + '</tbody></table></td></tr>';
              end
              else
              begin
                for p := 0 to oList.Count - 1 do
                begin
                  Result := Result + oList[p];
                end;
                Result := Result + ')' + #13#10;
              end;
            end
            else if (i = 4) { and ((flgIA2 and 16) <> 0) } then
            begin
              try
                if SUCCEEDED(ia2.Get_nRelations(iRole)) then
                begin
                  if SetTL then
                  begin
                    sBodyTxt :=
                      sBodyTxt + #13#10 + '<tr><th class="col2" colspan="2">' +
                      lIA2[i + 1] + '</th></tr><tr><td colspan="2"><table><tbody>';
                    sBodyTxt :=
                      sBodyTxt + #13#10 +
                      '<tr><th class="center">Type</th><th class="center">Targets</th></tr>';
                  end;
                  for p := 0 to iRole - 1 do
                  begin
                    Result := Result + #13#10 + #9 + rType + ':';
                    ia2.Get_relation(p, iAL);
                    iAL.Get_RelationType(s);


                    if SetTL then
                    begin
                      sBodyTxt :=
                        sBodyTxt + #13#10 + '<tr><td class="name">' +
                        s + '</td><td class="value"><ol>';

                    end;
                    Result := Result + s;
                    Result := Result + #13#10 + #9 + rTarg + ':';

                    iAL.Get_nTargets(iTarg);
                    for t := 0 to iTarg do
                    begin
                      iInter := nil;
                      hr := iAL.Get_target(t, iInter);
                      if SUCCEEDED(hr) then
                      begin
                        hr := iInter.QueryInterface(IID_IACCESSIBLE2, ia2Targ);
                        if SUCCEEDED(hr) then
                        begin
                          hr := ia2Targ.QueryInterface(IID_IServiceProvider, iSP);
                          if SUCCEEDED(hr) then
                          begin
                            hr :=
                              iSP.QueryService(IID_IACCESSIBLE, IID_IACCESSIBLE, iaTarg);
                            if SUCCEEDED(hr) then
                            begin

                              // s := iaTarg.accName[0];
                              iaTarg.Get_accName(0, s);
                              if s = '' then
                                s := None;
                              // iaTarg.Get_accValue(0, s);
                              ia2Targ.Role(iRole);
                              ws := GetIA2Role(iRole);
                              if ws = '' then
                                ws := None;
                              // iaTarg.Get_accValue(0, ws);
                              Result := Result + s + ' - ' + ws + #13#10 + #9;
                              if SetTL then
                              begin
                                sBodyTxt :=
                                  sBodyTxt + #13#10 + '<li>' + s + ' - ' + ws + '</li>';
                              end;
                            end;
                          end;
                        end;
                      end;
                    end;
                    if SetTL then
                      sBodyTxt := sBodyTxt + #13#10 + '</ol>';
                  end;
                  if SetTL then
                    sBodyTxt :=
                      sBodyTxt + #13#10 + '</td></tr></tbody></table></td></tr>';
                end;
              except
                on E: Exception do
                  ShowErr(E.Message);
              end;
            end
            else if i = 9 then
            begin
              if SetTL then
              begin
                sBodyTxt := sBodyTxt + #13#10 + '<tr><th class="col2" colspan="2">' +
                  lIA2[10] + '</th></tr><tr><td class="col2" colspan="2"><table><tbody>';

              end;
              for t := 0 to 2 do
              begin
                if SetTL then
                begin
                  sBodyTxt := sBodyTxt + #13#10 + '<tr><td class="name">' +
                    lIA2[11 + t] + '</td><td class="value">' +
                    MSAAs[t + 9] + '</td></tr>';

                end;
                Result := Result + lIA2[11 + t] + ':' + #9 + MSAAs[t + 9] + #13#10;
              end;
              sBodyTxt := sBodyTxt + #13#10 + '</tbody></table></td></tr>';
            end
            else if i = 10 then
            begin

              if SetTL then
              begin
                // Node := nodes.AddChild(rNode, lIA2[14]);

                if MSAAs[12] = '' then
                begin
                  sBodyTxt := sBodyTxt + #13#10 + '<tr><td class="name">' +
                    lIA2[14] + '</td><td class="value">' + none + '</td></tr>';

                end
                else
                begin
                  sBodyTxt := sBodyTxt + #13#10 + '<tr><th class="col2" colspan="2">' +
                    lIA2[14] +
                    '</th></tr><tr><td class="col2" colspan="2"><table><tbody>';

                  cList := TStringList.Create;
                  try
                    cList.Delimiter := ';';

                    cList.StrictDelimiter := True;
                    cList.DelimitedText := MSAAs[12];
                    for p := 0 to cList.Count - 1 do
                    begin
                      if cList[p] <> '' then
                      begin
                        iPos := pos(':', cList[p]);
                        sBodyTxt :=
                          sBodyTxt + #13#10 + '<tr><td class="name">' +
                          copy(cList[p], 0, iPos - 1) + '</td><td class="value">' +
                          copy(cList[p], iPos + 1, Length(cList[p])) + '</td></tr>';

                      end;
                    end;
                    sBodyTxt := sBodyTxt + #13#10 + '</tbody></table></td></tr>';
                  finally
                    cList.Free;
                  end;
                end;
              end;
              Result := Result + lIA2[14] + ':' + #9 + MSAAs[12] + #13#10;
            end
            else if i = 11 then
            begin
              if SetTL then
              begin
                sBodyTxt := sBodyTxt + #13#10 + '<tr><td class="name">' +
                  lIA2[15] + '</td><td class="value">' + MSAAs[13] + '</td></tr>';
              end;
              Result := Result + lIA2[15] + ':' + #9 + MSAAs[13] + #13#10;
            end
            else
            begin
              if SetTL then
              begin
                sBodyTxt := sBodyTxt + #13#10 + '<tr><td class="name">' +
                  lIA2[i + 1] + '</td><td class="value">' + MSAAs[i] + '</td></tr>';
              end;

              Result := Result + lIA2[i + 1] + ':' + #9 + MSAAs[i] + #13#10;
            end;
          end;
        end;
        sBodyTxt := sBodyTxt + #13#10 + '</tbody></table>';

      end
      else // IAccessible2 failed
      begin
        if SetTL then
        begin
          sBodyTxt := sBodyTxt + #13#10 + '<h1>' + lIA2[0] + '(' + Err_Inter + ')</h1>';

        end;
      end;
    end;
    TipTextIA2 := Result + #13#10 + #13#10;
  finally
    oList.Free;
  end;

end;

procedure TwndMain.GetNaviState(AllFalse: boolean = False);
begin
  if AllFalse then
  begin
    tbParent.Enabled := False;
    tbChild.Enabled := False;
    tbPrevS.Enabled := False;
    tbNextS.Enabled := False;
    acParent.Enabled := tbParent.Enabled;
    acChild.Enabled := tbChild.Enabled;
    acPrevS.Enabled := tbPrevS.Enabled;
    acNextS.Enabled := tbNextS.Enabled;
  end
  else
  begin
    GetNaviState(True);
    if (tvMSAA.Items.Count > 0) and (Pagecontrol1.ActivePageIndex = 0) then
    begin
      if (tvMSAA.SelectionCount > 0) then
      begin
        if tvMSAA.Selected.Parent <> nil then
        begin
          tbParent.Enabled := True;
        end;
        if tvMSAA.Selected.HasChildren then
        begin
          tbChild.Enabled := True;
        end;
        if tvMSAA.Selected.getNextSibling <> nil then
        begin
          tbNextS.Enabled := True;
        end;
        if tvMSAA.Selected.getPrevSibling <> nil then
        begin
          tbPrevS.Enabled := True;
        end;
      end;

    end;
    if (tvUIA.Items.Count > 0) and (Pagecontrol1.ActivePageIndex = 1) then
    begin
      if (tvUIA.SelectionCount > 0) then
      begin
        if tvUIA.Selected.Parent <> nil then
        begin
          tbParent.Enabled := True;
        end;
        if tvUIA.Selected.HasChildren then
        begin
          tbChild.Enabled := True;
        end;
        if tvUIA.Selected.getNextSibling <> nil then
        begin
          tbNextS.Enabled := True;
        end;
        if tvUIA.Selected.getPrevSibling <> nil then
        begin
          tbPrevS.Enabled := True;
        end;
      end;

    end;
    acParent.Enabled := tbParent.Enabled;
    acChild.Enabled := tbChild.Enabled;
    acPrevS.Enabled := tbPrevS.Enabled;
    acNextS.Enabled := tbNextS.Enabled;

  end;
end;

procedure TwndMain.ShowRectWnd(bkClr: TColor);
var
  RC: TRect;
  vChild: variant;
begin
  if not Assigned(iAcc) then
    Exit;

  try
    if not Assigned(WndFocus) then
      WndFocus := TWndFocusRect.Create(self);
    WndFocus.Shape1.Brush.Color := bkClr;
    vChild := varParent;//CHILDID_SELF;
    iAcc.accLocation(RC.Left, RC.Top, RC.Right, RC.Bottom, vChild);
    SetWindowPos(WndFocus.Handle, HWND_TOPMOST, RC.Left - 5,
      RC.Top - 5, RC.Right + 10, RC.Bottom + 10, SWP_SHOWWINDOW or SWP_NOACTIVATE);
    if hRgn1 <> 0 then
    begin
      DeleteObject(hRgn1);
      DeleteObject(hRgn2);
      DeleteObject(hRgn3);
      hRgn1 := CreateRectRgn(0, 0, WndFocus.Width, WndFocus.Height);
      SetWindowRgn(WndFocus.Handle, hRgn1, False);
      InvalidateRect(WndFocus.Handle, nil, False);
    end;
    hRgn1 := CreateRectRgn(0, 0, 1, 1);
    hRgn2 := CreateRectRgn(0, 0, WndFocus.Width, WndFocus.Height);
    hRgn3 := CreateRectRgn(5, 5, WndFocus.Width - 5, WndFocus.Height - 5);
    CombineRgn(hRgn1, hRgn2, hRgn3, RGN_DIFF);
    SetWindowRgn(WndFocus.Handle, hRgn1, True);
    WndFocus.Visible := True;
  except

  end;
end;

procedure TwndMain.ShowRectWnd2(bkClr: TColor; RC: TRect);
begin

  try
    if not Assigned(WndFocus) then
      WndFocus := TWndFocusRect.Create(self);
    WndFocus.Shape1.Brush.Color := bkClr;

    SetWindowPos(WndFocus.Handle, HWND_TOPMOST, RC.Left - 5,
      RC.Top - 5, RC.Right + 10, RC.Bottom + 10, SWP_SHOWWINDOW or SWP_NOACTIVATE);
    if hRgn1 <> 0 then
    begin
      DeleteObject(hRgn1);
      DeleteObject(hRgn2);
      DeleteObject(hRgn3);
      hRgn1 := CreateRectRgn(0, 0, WndFocus.Width, WndFocus.Height);
      SetWindowRgn(WndFocus.Handle, hRgn1, False);
      InvalidateRect(WndFocus.Handle, nil, False);
    end;
    hRgn1 := CreateRectRgn(0, 0, 1, 1);
    hRgn2 := CreateRectRgn(0, 0, WndFocus.Width, WndFocus.Height);
    hRgn3 := CreateRectRgn(5, 5, WndFocus.Width - 5, WndFocus.Height - 5);
    CombineRgn(hRgn1, hRgn2, hRgn3, RGN_DIFF);
    SetWindowRgn(WndFocus.Handle, hRgn1, True);
    WndFocus.Visible := True;
  except

  end;
end;

function AddTreeItem(tv: TTreeView; AccItem: iAccessible;
  UIAItem: IUIAutomationElement; AccID: integer; ParentItem: TTreeNode;
  bDummy: boolean = False): TTreeNode;
var
  TD: PTreeData;
  dTxt: string;
begin
  New(TD);
  TD^.Acc := AccItem;
  TD^.UIEle := UIAItem;
  TD^.iID := AccID;
  TD^.dummy := bDummy;
  Result := tv.Items.AddChildObject(ParentItem, '', Pointer(TD));
end;

procedure TwndMain.CreateMSAATree;
var
  Selnode: TTreeNode;
  cAcc, tAcc, SyncAcc, pAcc, compAcc: iAccessible;
  iChild, iCH, i, itarg, iComp, dChild: integer;
  iObtain: plongint;
  cNode, dNode, SSelNode: TTreeNode;
  TD: PTreeData;

  hr, hr_SCld: HResult;
  aChildren: array of TVariantArg;
  bFSame, bAddDL: boolean;
  iDis: IDispatch;
begin

  Selnode := AddTreeItem(tvMSAA, iAcc, nil, Varparent, nil);
  Selnode.Expand(True);
  Selnode.Selected := True;

  if VarParent = 0 then
  begin
    iAcc.Get_accChildCount(iChild);
    SetLength(aChildren, iChild);
    for i := 0 to iChild - 1 do
    begin
      VariantInit(olevariant(aChildren[i]));
    end;
    hr_SCld := AccessibleChildren(iAcc, 0, iChild, @aChildren[0], iObtain);

    if hr_SCld = 0 then
    begin
      for i := 0 to integer(iObtain) - 1 do
      begin
        Application.ProcessMessages;

        if aChildren[i].vt = VT_DISPATCH then
        begin
          itarg := i;
          tAcc := nil;
          hr := IDispatch(aChildren[itarg].pdispVal).QueryInterface(
            IID_IACCESSIBLE, tAcc);
          if Assigned(tAcc) and (hr = S_OK) then
          begin
            cNode := AddTreeItem(tvMSAA, tAcc, nil, 0, SelNode, True);
            tvMSAA.Items.AddChild(cNode, 'avwr_dummy');
          end;
        end
        else
        begin
          iCH := aChildren[i].lVal;
          AddTreeItem(tvMSAA, iAcc, nil, iCH, SelNode);
        end;

      end;
    end;
  end;
  if mnuAll.Checked then
  begin
    pAcc := nil;
    if iCH = 0 then
    begin
      hr := cAcc.Get_accParent(iDis);
      if (hr = S_OK) and (Assigned(iDis)) then
        hr := iDis.QueryInterface(IID_IACCESSIBLE, pAcc);
    end
    else
    begin
      hr := S_OK;
      pAcc := cAcc;
    end;
  end;
end;

function TwndMain.Get_RoleText(Acc: IAccessible; Child: integer): string;
var
  PC: PWideChar;
  ovValue, ovChild: olevariant;
begin
  ovChild := Child;
  //ovValue := Acc.accRole[ovChild];
  Acc.Get_accRole(ovChild, ovValue);
  Result := None;
  try
    //if (IDispatch(ovValue) = nil) then

    if VarHaveValue(ovValue) then
    begin
      //GetMem(PC,255);
      if VarIsNumeric(ovValue) then
      begin
        PC := WideStrAlloc(255);
        GetRoleTextW(ovValue, PC, StrBufSize(PC));
        Result := PC;
        StrDispose(PC);
      end
      else if VarIsStr(ovValue) then
      begin
        Result := VarToStr(ovValue);
      end;
    end;
  except
    on E: Exception do
      ShowErr(E.Message);

  end;
end;

function TwndMain.ARIAText: string;
var
  iEle5: IHTMLElement5;
  List, List2: TStringList;
  i: integer;
  aAttrs, LowerS: string;
  aPC, aPC2: array [0 .. 64] of PWideChar;
  aSI: array [0 .. 64] of smallint;
  WD: word;
  hr: HResult;
begin

  if Assigned(WndLabel) then
  begin
    FreeAndNil(WndLabel);
    WndLabel := nil;
  end;
  if Assigned(WndDesc) then
  begin
    FreeAndNil(WndDesc);
    WndDesc := nil;
  end;
  if Assigned(WndTarg) then
  begin
    FreeAndNil(WndTarg);
    WndTarg := nil;
  end;
  aAttrs := '';
  if (Assigned(CEle)) then
  begin
    hr := CEle.QueryInterface(IID_IHTMLELEMENT5, iEle5);
    if (SUCCEEDED(hr)) and (Assigned(Iele5)) then
    begin
      if iEle5.hasAttribute('role') then
        aAttrs := aAttrs + '"role","' + iEle5.role + '"' + #13#10;
      if iEle5.hasAttribute('aria-Activedescendant') then
        aAttrs := aAttrs + '"aria-Activedescendant","' +
          iEle5.ariaActivedescendant + '"' + #13#10;
      if iEle5.hasAttribute('aria-Busy') then
        aAttrs := aAttrs + '"aria-Busy","' + iEle5.ariaBusy + '"' + #13#10;
      if iEle5.hasAttribute('aria-Checked') then
        aAttrs := aAttrs + '"aria-Checked","' + iEle5.ariaChecked + '"' + #13#10;
      if iEle5.hasAttribute('aria-Controls') then
        aAttrs := aAttrs + '"aria-Controls","' + iEle5.ariaControls + '"' + #13#10;
      if iEle5.hasAttribute('aria-Describedby') then
        aAttrs := aAttrs + '"aria-Describedby","' + iEle5.ariaDescribedby + '"' + #13#10;
      if iEle5.hasAttribute('aria-Disabled') then
        aAttrs := aAttrs + '"aria-Disabled","' + iEle5.ariaDisabled + '"' + #13#10;
      if iEle5.hasAttribute('aria-Expanded') then
        aAttrs := aAttrs + '"aria-Expanded","' + iEle5.ariaExpanded + '"' + #13#10;
      if iEle5.hasAttribute('aria-Flowto') then
        aAttrs := aAttrs + '"aria-Flowto","' + iEle5.ariaFlowto + '"' + #13#10;
      if iEle5.hasAttribute('aria-Haspopup') then
        aAttrs := aAttrs + '"aria-Haspopup","' + iEle5.ariaHaspopup + '"' + #13#10;
      if iEle5.hasAttribute('aria-Hidden') then
        aAttrs := aAttrs + '"aria-Hidden","' + iEle5.ariaHidden + '"' + #13#10;
      if iEle5.hasAttribute('aria-Invalid') then
        aAttrs := aAttrs + '"aria-Invalid","' + iEle5.ariaInvalid + '"' + #13#10;
      if iEle5.hasAttribute('aria-Labelledby') then
        aAttrs := aAttrs + '"aria-Labelledby","' + iEle5.ariaLabelledby + '"' + #13#10;
      if iEle5.hasAttribute('aria-Level') then
        aAttrs := aAttrs + '"aria-Level","' + IntToStr(iEle5.ariaLevel) + '"' + #13#10;
      if iEle5.hasAttribute('aria-Live') then
        aAttrs := aAttrs + '"aria-Live","' + iEle5.ariaLive + '"' + #13#10;
      if iEle5.hasAttribute('aria-Multiselectable') then
        aAttrs := aAttrs + '"aria-Multiselectable","' + iEle5.ariaMultiselectable +
          '"' + #13#10;
      if iEle5.hasAttribute('aria-Owns') then
        aAttrs := aAttrs + '"aria-Owns","' + iEle5.ariaOwns + '"' + #13#10;
      if iEle5.hasAttribute('aria-Posinset') then
        aAttrs := aAttrs + '"aria-Posinset","' + IntToStr(iEle5.ariaPosinset) +
          '"' + #13#10;
      if iEle5.hasAttribute('aria-Pressed') then
        aAttrs := aAttrs + '"aria-Pressed","' + iEle5.ariaPressed + '"' + #13#10;
      if iEle5.hasAttribute('aria-Readonly') then
        aAttrs := aAttrs + '"aria-Readonly","' + iEle5.ariaReadonly + '"' + #13#10;
      if iEle5.hasAttribute('aria-Relevant') then
        aAttrs := aAttrs + '"aria-Relevant","' + iEle5.ariaRelevant + '"' + #13#10;
      if iEle5.hasAttribute('aria-Required') then
        aAttrs := aAttrs + '"aria-Required","' + iEle5.ariaRequired + '"' + #13#10;
      if iEle5.hasAttribute('aria-Secret') then
        aAttrs := aAttrs + '"aria-Secret","' + iEle5.ariaSecret + '"' + #13#10;
      if iEle5.hasAttribute('aria-Selected') then
        aAttrs := aAttrs + '"aria-Selected","' + iEle5.ariaSelected + '"' + #13#10;
      if iEle5.hasAttribute('aria-Setsize') then
        aAttrs := aAttrs + '"aria-Setsize","' + IntToStr(iEle5.ariaSetsize) +
          '"' + #13#10;
      if iEle5.hasAttribute('aria-Valuemin') then
        aAttrs := aAttrs + '"aria-Valuemin","' + iEle5.ariaValuemin + '"' + #13#10;
      if iEle5.hasAttribute('aria-Valuemax') then
        aAttrs := aAttrs + '"aria-Valuemax","' + iEle5.ariaValuemax + '"' + #13#10;
      if iEle5.hasAttribute('aria-Valuenow') then
        aAttrs := aAttrs + '"aria-Valuenow","' + iEle5.ariaValuenow + '"' + #13#10;

    end;

  end
  else if (Assigned(SDom)) then
  begin
    // function get_attributes(maxAttribs: WORD; out attribNames: TBSTR; out nameSpaceID: Smallint; out attribValues: TBSTR; out numAttribs: WORD): HRESULT; stdcall;
    // function get_attributesForNames(numAttribs: WORD; out attribNames: TBSTR; out nameSpaceID: Smallint; out attribValues: TBSTR): HRESULT; stdcall;
    SDom.Get_attributes(65, aPC[0], aSI[0], aPC2[0], WD);
    for i := 0 to WD - 1 do
    begin
      LowerS := LowerCase(aPC[i]);
      if LowerS = 'role' then
      begin
        aAttrs := aAttrs + '"' + ARIAs[0, 1] + '","' + aPC2[i] + '"' + #13#10;
      end
      else if copy(LowerS, 1, 4) = 'aria' then
      begin

        aAttrs := aAttrs + '"' + aPC[i] + '","' + aPC2[i] + '"' + #13#10;

      end;
    end;
  end;

  if aAttrs = '' then
    ARIAs[1, 1] := None
  else
    ARIAs[1, 1] := aAttrs;

  sBodyTxt := sBodyTxt + #13#10 + '<h1>' + sARIA + '</h1><table><tbody>';
  if aAttrs <> '' then
  begin

    List := TStringList.Create;
    List2 := TStringList.Create;
    try
      List.Text := aAttrs;
      for i := 0 to List.Count - 1 do
      begin
        List2.Clear;
        List2.CommaText := List[i];

        if List2.Count >= 1 then
        begin
          if List2.Count >= 2 then
          begin
            sBodyTxt := sBodyTxt + #13#10 + '<tr><td class="name">' +
              List2[0] + '</td><td class="value">' + List2[1] + '</td></tr>';
          end
          else
            sBodyTxt := sBodyTxt + #13#10 + '<tr><td class="name">' +
              List2[0] + '</td><td class="value"></td></tr>';
        end;
      end;
    finally
      List.Free;
      List2.Free;
    end;
  end
  else
  begin
    sBodyTxt := sBodyTxt + #13#10 + '<tr><td class="name">' + ARIAs[1, 0] +
      '</td><td class="value">' + ARIAs[1, 1] + '</td></tr>';
  end;
  sBodyTxt := sBodyTxt + #13#10 + '</tbody></table>';


  Result := sARIA + #13#10 + ARIAs[1, 1] + #13#10#13#10;
end;

procedure TwndMain.GetTreeStructure;
var
  iSP: IServiceProvider;
  iEle, paEle: IHTMLElement;
  Path, cAccRole: string;
  ws: WideString;
  iSEle: ISimpleDOMNode;
  SI: smallint;
  PU, PU2: PUINT;
  WD: word;
  PC, PC2: PChar;
  pAcc: IAccessible;
  iRes: integer;
  iDoc: IHTMLDocument2;
  pEle, TempEle: ISimpleDOMNode;
  i: integer;
  AWnd: HWND;
  bSame: boolean;
  gtcStart, gtcStop: cardinal;

  procedure GetAccTxt;
  begin
    iAcc.Get_accName(VarParent, ws);
    NodeTxt := string(ws);
    if NodeTxt = '' then
      NodeTxt := None;
    NodeTxt := StringReplace(NodeTxt, #13, ' ', [rfReplaceAll]);
    NodeTxt := StringReplace(NodeTxt, #10, ' ', [rfReplaceAll]);
    cAccRole := Get_RoleText(iAcc, VarParent);
    NodeTxt := NodeTxt + ' - ' + cAccRole;

  end;

const
  Class_ISimpleDOMNode: TGUID = '{0D68D6D0-D93D-4D08-A30D-F00DD1F45B24}';
begin
  if not Assigned(iAcc) then
    Exit;
  sMSAAtxt := '';
  sHTMLtxt := '';
  sUIATxt := '';
  sARIATxt := '';
  sIA2Txt := '';
  nSelected := False;
  sBodyTxt := '';
  sCodeTxt := '';

  GetNaviState(True);
  CEle := nil;
  SDom := nil;
  Path := IncludeTrailingPathDelimiter(ExtractFileDir(Application.ExeName));
  if acOnlyFocus.Checked then
    ShowRectWnd(clYellow)
  else if (acRect.Checked) then
  begin
    ShowRectWnd(clRed);
  end;

  try
    if (PageControl1.ActivePageIndex = 0) then
      sMSAAtxt := MSAAText;
    iRes := iAcc.QueryInterface(IID_IServiceProvider, iSP);
    if SUCCEEDED(iRes) and Assigned(iSP) then
    begin
      iRes := iSP.QueryService(IID_IHTMLElement, IID_IHTMLElement, iEle);
      if SUCCEEDED(iRes) and Assigned(iEle) then
      begin
        CEle := iEle;
        if (PageControl1.ActivePageIndex = 0) or (TreeMode) then
        begin
          if mnuARIA.Checked then
          begin
            sARIATxt := ARIAText;
          end;
          if mnuHTML.Checked then
          begin
            HTMLText;
          end;
        end;

        if (not Treemode) then
        begin
          iDoc := iEle.Document as IHTMLDocument2;
          paEle := iEle;
          if mnuAll.Checked then
          begin
            paEle := iDoc.body;

          end;
          if Assigned(paEle) then
          begin
            iRes := paEle.QueryInterface(IID_IServiceProvider, iSP);
            if SUCCEEDED(iRes) and Assigned(iSP) then
            begin
              iRes := iSP.QueryService(IID_IACCESSIBLE, IID_IACCESSIBLE, pAcc);
              if SUCCEEDED(iRes) and Assigned(pAcc) then
              begin
                CreateMSAATree;
                WriteHTML;
              end;
            end;
          end;
        end;
      end
      else  //IHTMLElement is not support
      begin
        iRes := iSP.QueryService(IID_ISIMPLEDOMNODE, IID_ISIMPLEDOMNODE, isEle);
        if SUCCEEDED(iRes) and Assigned(isEle) then
        begin

        end
        else
        begin

        end;
      end;
    end;
  finally
  end;
end;

procedure TwndMain.WriteHTML;
var
  List: TStringList;
  RS: TResourceStream;
  hr: hresult;
  WBDoc: IHTMLDocument2;
  arOle: variant;
  url, onull: olevariant;
  sTemp: string;
begin
  if sBodyTxt = '' then
    exit;

  if GetTemp(sTemp) then
  begin
    sTemp := sTemp + 'avwr_temp.html';
    url := sTemp;

    List := TStringList.Create;
    RS := TResourceStream.Create(hInstance, 'VLIST', PChar('TEXT'));
    try
      try

        List.LoadFromStream(RS);
        List.Text := StringReplace(List.Text, '%body%', sBodyTxt,
          [rfReplaceAll, rfIgnoreCase]);
        List.Text := StringReplace(List.Text, '%code%', sCodeTxt,
          [rfReplaceAll, rfIgnoreCase]);
        List.Text := StringReplace(List.Text, '%fs%', IntToStr(Font.Size) +
          'pt', [rfReplaceAll, rfIgnoreCase]);
        List.Text := StringReplace(List.Text, '%ff%', Font.Name,
          [rfReplaceAll, rfIgnoreCase]);
        List.SaveToFile(sTemp);
        WB1.ComServer.Navigate2(url, onull, onull, onull, onull);
        {hr := wb1.ComServer.Document.QueryInterface(IID_IHTMLDOCUMENT2, WBDoc);
        if (SUCCEEDED(hr)) and (Assigned(WBDoc)) then
        begin
          arOle := VarArrayCreate([0, 0], VarVariant);
          try
            VarArrayLock(arOle);
            arOle[0] := List.Text;
            WBDoc.Writeln(PSafeArray(System.TVarData(arOle).VArray));
            WBDoc.Close;
            //WBDoc.onclick := (TWBEvent.Create(WBOnClick) as IDispatch);
          finally
            VarArrayUnLock(arOle);
            VarClear(arOle);
          end;
        end;}
      except
        on E: Exception do
        begin
          ShowMessage(E.Message);
          Exit;
        end;
      end;
    finally
      RS.Free;
      List.Free;
    end;

  end;
end;

procedure TwndMain.OnStatusTextChange(Sender: TObject; Text_: WideString);
begin
  //Label1.Caption:='Browser: '+UTF8Encode(Text_);
end;

end.
