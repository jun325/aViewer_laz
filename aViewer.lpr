program aViewer;

{$mode objfpc}{$H+}

uses
  {$IFDEF UNIX}{$IFDEF UseCThreads}
  cthreads,
  {$ENDIF}{$ENDIF}
  Interfaces, // this includes the LCL widgetset
  Forms, Windows, Messages, Dialogs, SysUtils, frmMain, Registry,
  MSHTML_4_0_TLB;

{$R *.res}
{$R resource.rc}
procedure RegFBE;
var
   Reg: TRegistry;
   sFN: string;
   cType: Cardinal;
   lFlag: LongBool;
const
  SCS_64BIT_BINARY = 6;
begin
  lFlag := GetBinaryType(PChar(application.ExeName), cType);
  if lFlag then
  begin
    if cType = SCS_64BIT_BINARY then
    	sFN := 'aViewer64bit.exe'
    else
    	sFN := 'aViewer32bit.exe';

  end
  else
    sFN := 'aViewer32bit.exe';
  Reg := TRegistry.Create(KEY_ALL_ACCESS );
  try
  	Reg.RootKey := HKEY_CURRENT_USER;
    if Reg.OpenKey('SOFTWARE\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION\', false) then
    begin

      if (not Reg.ValueExists(sFN)) then
      	Reg.WriteInteger(sFN, $00002af8);
      if (not Reg.ValueExists('aViewer.exe')) then
      	Reg.WriteInteger('aViewer.exe', $00002af8);

      Reg.CloseKey;
    end;
    {if Reg.OpenKey('SOFTWARE\WOW6432Node\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION\', False) then
    begin
			if (not Reg.ValueExists(sFN)) then
      	Reg.WriteInteger(sFN, $00002af8);
      Reg.CloseKey;
    end;}

  finally
  	Reg.Free;
  end;
end;

begin
  RegFBE;
  RequireDerivedFormResource:=True;
  Application.Scaled:=True;
  Application.Initialize;
  Application.CreateForm(TwndMain, wndMain);
  Application.Run;
end.

