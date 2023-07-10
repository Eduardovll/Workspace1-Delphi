unit UProgresso;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls, Buttons, Data.DB, Datasnap.DBClient, Vcl.Grids,
  Vcl.DBGrids, UClasses;

type
  TTipoLog = (tlAviso, tlErro);
  TFrmProgresso = class(TForm)
    BtnCancelar: TButton;
    LblQtdErros: TLabel;
    LblRegistroAtual: TLabel;
    PgbGeracao: TProgressBar;
    LblAcaoAtual: TLabel;
    spdRestaurar: TSpeedButton;
    DsTable: TDataSource;
    pgb: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    DBGrid1: TDBGrid;
    MemProgresso: TMemo;
    procedure FormCreate(Sender: TObject);
    procedure BtnCancelarClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormResize(Sender: TObject);
    procedure spdRestaurarClick(Sender: TObject);
  private
    { Private declarations }
  public
    ExibirGrade: Boolean;
    ExibirFormatado: Boolean;
    procedure AdicionarLog(Registro: Integer; Tipo: Char; Mensagem: String); overload; deprecated;
    procedure AdicionarLog(Tipo: TTipoLog; Ident, Mensagem: String); overload;
    procedure AdicionarLog(Tipo: TTipoLog; NumReg : Integer; Mensagem: String); overload;
    procedure ExibirProgresso(Atual, Total: Integer);
    procedure ExibirAcaoAtual(Descricao: String);
    procedure Registrar(Tipo: Char; Id, Mensagem: String);
    procedure CriarDS(Layout: TLayout);
    procedure Init;
  end;

var
  FrmProgresso: TFrmProgresso;
  QtdErros: Integer;
  Cancelar, Encerrado: Boolean;
  UltAtualizTela: Int64; //ultima atualização da tela

implementation

{$R *.dfm}

{ TFrmProgresso }

procedure TFrmProgresso.AdicionarLog(Registro: Integer; Tipo: Char;
  Mensagem: String);
begin
  if LowerCase(Tipo) = 'e' then
  begin
    Inc(QtdErros);
    LblQtdErros.Caption := IntToStr(QtdErros) + ' Erros';
    MemProgresso.Lines.Add('Erro no registro nº' + IntToStr(Registro) + ': ' + Mensagem);
  end
  else if LowerCase(Tipo) = 'a' then
    MemProgresso.Lines.Add('Aviso no registro nº' + IntToStr(Registro) + ': ' + Mensagem);
end;

procedure TFrmProgresso.Registrar(Tipo: Char; Id, Mensagem: String);
begin
  if UpperCase(Tipo) = 'E' then
  begin
    Inc(QtdErros);
    LblQtdErros.Caption := IntToStr(QtdErros) + ' Erros';
    MemProgresso.Lines.Add('Erro, ' + Id + ': ' + Mensagem);
  end
  else if UpperCase(Tipo) = 'A' then
    MemProgresso.Lines.Add('Aviso, ' + Id + ': ' + Mensagem);
end;

procedure TFrmProgresso.spdRestaurarClick(Sender: TObject);
begin
  WindowState := wsNormal;
end;

procedure TFrmProgresso.ExibirProgresso(Atual, Total: Integer);
begin
  if (GetTickCount - UltAtualizTela) > 200 then
  begin
    UltAtualizTela := GetTickCount;
    PgbGeracao.Position := Round((Atual * 100) / Total);
    LblRegistroAtual.Caption := 'Registro: ' + IntToStr(Atual) + '/' + IntToStr(Total);
    Application.ProcessMessages;
  end;
end;

procedure TFrmProgresso.FormCreate(Sender: TObject);
begin
  LblRegistroAtual.Caption := '';
  LblQtdErros.Caption := '';
  UltAtualizTela := 0;
end;

procedure TFrmProgresso.FormResize(Sender: TObject);
begin
  spdRestaurar.Visible := (WindowState = wsMaximized);
end;

procedure TFrmProgresso.Init;
begin
  if not ExibirGrade then
  begin
    pgb.ActivePage := TabSheet2;
    TabSheet1.Hide;
  end
  else
  begin
    TabSheet1.Show;
    pgb.ActivePage := TabSheet1;
  end;
end;

procedure TFrmProgresso.ExibirAcaoAtual(Descricao: String);
begin
  FrmProgresso.PgbGeracao.Position := 0;
  FrmProgresso.LblAcaoAtual.Caption := Descricao;
  MemProgresso.Lines.Add(Descricao + '...');
  Application.ProcessMessages;
end;

procedure TFrmProgresso.AdicionarLog(Tipo: TTipoLog; Ident, Mensagem: String);
begin
  if Tipo = tlErro then
  begin
    Inc(QtdErros);
    LblQtdErros.Caption := IntToStr(QtdErros) + ' Erros';
    MemProgresso.Lines.Add('Erro, ' + Ident + ': ' + Mensagem);
  end
  else if Tipo = tlAviso then
    MemProgresso.Lines.Add('Aviso, ' + Ident + ': ' + Mensagem);
end;

procedure TFrmProgresso.AdicionarLog(Tipo: TTipoLog; NumReg: Integer;
  Mensagem: String);
begin
  if Tipo = tlErro then
  begin
    Inc(QtdErros);
    LblQtdErros.Caption := IntToStr(QtdErros) + ' Erros';
    MemProgresso.Lines.Add('Erro, registro nº' + IntToStr(NumReg) + ': ' + Mensagem);
  end
  else if Tipo = tlAviso then
    MemProgresso.Lines.Add('Aviso, registro nº' + IntToStr(NumReg) + ': ' + Mensagem);
end;

procedure TFrmProgresso.BtnCancelarClick(Sender: TObject);
begin
  Cancelar := True;
end;

procedure TFrmProgresso.CriarDS(Layout: TLayout);
var Client : TClientDataSet;
begin
  Client := TClientDataSet.Create(FrmProgresso);

  Layout.CreateDataSet(Client);

  dsTable.DataSet := Client;
end;

procedure TFrmProgresso.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  if not Encerrado then
    Cancelar := True
  else
    Sender.Free;
end;

end.

