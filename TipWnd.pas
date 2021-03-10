unit TipWnd;
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
interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics,
  Controls, Forms, Dialogs, StdCtrls, ExtCtrls, LCLType;

type

  { TfrmTipWnd }

  TfrmTipWnd = class(TForm)
    Panel1: TPanel;
    Label1: TLabel;
    Tipinfo: TMemo;
    procedure FormClose(Sender: TObject; var caAction: TCloseAction);
    procedure FormCloseQuery(Sender: TObject; var CanClose: boolean);
  private
    { Private �錾 }
  protected
    procedure CreateParams(var Params: TCreateParams); override;
  public
    { Public �錾 }
    hParent: hWnd;
  end;

var
  frmTipWnd: TfrmTipWnd;

implementation

{$R *.dfm}

procedure TfrmTipWnd.FormClose(Sender: TObject; var caAction: TCloseAction);
begin
    caAction := caFree;
    frmTipWnd := nil;
end;

procedure TfrmTipWnd.FormCloseQuery(Sender: TObject; var CanClose: boolean);
begin

end;

procedure TfrmTipWnd.CreateParams(var Params: TCreateParams);
begin
  inherited CreateParams(Params);

  //Params.ExStyle := Params.ExStyle or WS_EX_NOACTIVATE;
  //Params.Style := Params.Style or WS_POPUP;
end;

end.
