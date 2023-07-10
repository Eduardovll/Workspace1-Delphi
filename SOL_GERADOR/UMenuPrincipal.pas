unit UMenuPrincipal;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, ToolWin, Menus, ActnList, StdCtrls, pngimage, ExtCtrls, IniFiles, WideStrings, DB, SqlExpr, ADODB,
  System.Actions;


type
  TfrmMenuPrincipal = class(TForm)
    TreeViewMenu: TTreeView;
    ActionList1: TActionList;
    Image1: TImage;
    StatusBar1: TStatusBar;
    CkbAutomatico: TCheckBox;
    ActSmNossoPao: TAction;
    ActSmDucidoGestora: TAction;
    ActSmBelaVistaGsMarket: TAction;
    ActSmUlianGestora: TAction;
    procedure FormCreate(Sender: TObject);
    procedure TreeViewMenuClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure CkbAutomaticoClick(Sender: TObject);
    procedure ActAriusNagaiExecute(Sender: TObject);
    procedure ActGetwayMamaExecute(Sender: TObject);
    procedure ActByteCortezExecute(Sender: TObject);
    procedure ActMyCommercePrediletoExecute(Sender: TObject);
    procedure ActHiperLojaoExecute(Sender: TObject);
    procedure ActAlhoSupExecute(Sender: TObject);
    procedure ActSmPortalExecute(Sender: TObject);
    procedure ActSmDouradoExecute(Sender: TObject);
    procedure ActSmJardimLagoExecute(Sender: TObject);
    procedure ActSmPegorinExecute(Sender: TObject);
    procedure ActSmAnaMaraExecute(Sender: TObject);
    procedure ActSmMerkoFruitExecute(Sender: TObject);
    procedure ActSmBeneExecute(Sender: TObject);
    procedure ActSmNovaOpcaoExecute(Sender: TObject);
    procedure ActSmMultimarketingExecute(Sender: TObject);
    procedure ActSmSaoFranciscoExecute(Sender: TObject);
    procedure ActSmRioBrancoExecute(Sender: TObject);
    procedure ActSmRecantoDoPaoExecute(Sender: TObject);
    procedure ActSmBonfimExecute(Sender: TObject);
    procedure ActSmPandaExecute(Sender: TObject);
    procedure ActSmLaticinioMelloExecute(Sender: TObject);
    procedure ActSmTrioExecute(Sender: TObject);
    procedure ActSmKuzziExecute(Sender: TObject);
    procedure ActRonyMGExecute(Sender: TObject);
    procedure ActSmViaExecute(Sender: TObject);
    procedure ActSmParaisoExecute(Sender: TObject);
    procedure ActSmSuperComprasGestorExecute(Sender: TObject);
    procedure ActSmNossoPaoExecute(Sender: TObject);
    procedure ActSmDucidoGestoraExecute(Sender: TObject);
    procedure ActSmDaTerraFacileteExecute(Sender: TObject);
    procedure ActSmBelaVistaGsMarketExecute(Sender: TObject);
    procedure ActSmUlianGestoraExecute(Sender: TObject);
    procedure ActSmBomPrecoOurinhosExecute(Sender: TObject);


    
  private
    { Private declarations }
    function EncontrarNo(const aNome: string): TTreeNode;
  public
    { Public declarations }
    procedure CriarTelaSistema(Form: TForm; FormClass :TFormClass);
    procedure ConstruirArvore;
  end;

var
  frmMenuPrincipal: TfrmMenuPrincipal;

implementation

uses UFrmModelo, UUTilidades, UFrmAriusNagai, UFrmGetWaymama, UFrmByteCortez, UFrmMyCommercePredileto,
  UFrmHiperLojao, UFrmAlhoSupermercados, UFrmPortal, UFrmSmDourado,
  UFrmJardimLago, UFrmSmPegorin, UFrmSmAnaMara, UFrmMerkoFruit, UFrmCampoGrande,
  UFrmSmNovaOpcao, UFrmSmMultimarketing, UFrmSmSaoFrancisco, UFrmSmRioBranco,
  UFrmSmRecantoDoPao, UFrmSmBonfim, UFrmSmPanda, UFrmSmLaticinioMello,
  UFrmAriusTrio, UFrmSmKuzzi, UFrmSmRonyMG, UFrmSmVia, UFrmSmParaiso,
  UFrmSmSuperComprasGestor, UFrmSmNossoPao, UFrmSmPinkCosmGiga, UFrmSmDucido,
  UFrmSmDaTerraFacilete, UFrmSmBelaVistaGsMarket, UFrmSmUlianGestora, UFrmSmBomPrecoOurinhosGestora;



{$R *.dfm}



procedure TfrmMenuPrincipal.ActAlhoSupExecute(Sender: TObject);
begin
  CriarTelaSistema(FrmAlhoSupermercados, TFrmAlhoSupermercados);
end;

procedure TfrmMenuPrincipal.ActAriusNagaiExecute(Sender: TObject);
begin
  CriarTelaSistema(FrmAriusNagai, TFrmAriusNagai);
end;

procedure TfrmMenuPrincipal.ActByteCortezExecute(Sender: TObject);
begin
   CriarTelaSistema(FrmByteCortez, TFrmByteCortez);
end;

procedure TfrmMenuPrincipal.ActGetwayMamaExecute(Sender: TObject);
begin
   CriarTelaSistema(FrmGetWayMama, TFrmGetWayMama);
end;

procedure TfrmMenuPrincipal.ActHiperLojaoExecute(Sender: TObject);
begin
  CriarTelaSistema(FrmHiperLojao, TFrmHiperLojao);
end;

procedure TfrmMenuPrincipal.ActMyCommercePrediletoExecute(Sender: TObject);
begin
   CriarTelaSistema(FrmMyCommercePredileto, TFrmMyCommercePredileto);
end;

procedure TfrmMenuPrincipal.ActRonyMGExecute(Sender: TObject);
begin
  CriarTelaSistema(FrmSmRonyMG, TFrmSmRonyMG);
end;

procedure TfrmMenuPrincipal.ActSmAnaMaraExecute(Sender: TObject);
begin
  CriarTelaSistema(FrmSmAnaMara, TFrmSmAnaMara);
end;

procedure TfrmMenuPrincipal.ActSmBeneExecute(Sender: TObject);
begin
  CriarTelaSistema(FrmCampoGrande, TFrmCampoGrande);
end;

procedure TfrmMenuPrincipal.ActSmBomPrecoOurinhosExecute(Sender: TObject);
begin
 CriarTelaSistema(FrmSmBomPrecoOurinhosGestora, TFrmSmBomPrecoOurinhosGestora);
end;

procedure TfrmMenuPrincipal.ActSmBonfimExecute(Sender: TObject);
begin
  CriarTelaSistema(FrmSmBonfim, TFrmSmBonfim);
end;

procedure TfrmMenuPrincipal.ActSmDaTerraFacileteExecute(Sender: TObject);
begin
  CriarTelaSistema(FrmSmDaTerraFacilete, TFrmSmDaTerraFacilete);
end;

procedure TfrmMenuPrincipal.ActSmBelaVistaGsMarketExecute(Sender: TObject);
begin
  CriarTelaSistema(FrmSmBelaVistaGsMarket, TFrmSmBelaVistaGsMarket);
end;

procedure TfrmMenuPrincipal.ActSmUlianGestoraExecute(Sender: TObject);
begin
  CriarTelaSistema(FrmSmUlianGestora, TFrmSmUlianGestora);
end;

procedure TfrmMenuPrincipal.ActSmDouradoExecute(Sender: TObject);
begin
  CriarTelaSistema(FrmSmDourado, TFrmSmDourado);
end;

procedure TfrmMenuPrincipal.ActSmDucidoGestoraExecute(Sender: TObject);
begin
  CriarTelaSistema(FrmSmDucido, TFrmSmDucido);
end;

procedure TfrmMenuPrincipal.ActSmJardimLagoExecute(Sender: TObject);
begin
  CriarTelaSistema(FrmJardimLago, TFrmJardimLago);
end;

procedure TfrmMenuPrincipal.ActSmKuzziExecute(Sender: TObject);
begin
  CriarTelaSistema(FrmSmKuzzi, TFrmSmKuzzi);
end;


procedure TfrmMenuPrincipal.ActSmLaticinioMelloExecute(Sender: TObject);
begin
  CriarTelaSistema(FrmSmLaticinioMello, TFrmSmLaticinioMello);
end;

procedure TfrmMenuPrincipal.ActSmMerkoFruitExecute(Sender: TObject);
begin
  CriarTelaSistema(FrmMerkoFruit, TFrmMerkoFruit);
end;

procedure TfrmMenuPrincipal.ActSmMultimarketingExecute(Sender: TObject);
begin
  CriarTelaSistema(FrmSmMultimarketing, TFrmSmMultimarketing);
end;

procedure TfrmMenuPrincipal.ActSmNovaOpcaoExecute(Sender: TObject);
begin
  CriarTelaSistema(FrmSmNovaOpcao, TFrmSmNovaOpcao);
end;

procedure TfrmMenuPrincipal.ActSmPandaExecute(Sender: TObject);
begin
  CriarTelaSistema(FrmSmPanda, TFrmSmPanda);
end;

procedure TfrmMenuPrincipal.ActSmParaisoExecute(Sender: TObject);
begin
  CriarTelaSistema(FrmSmParaiso, TFrmSmParaiso);
end;

procedure TfrmMenuPrincipal.ActSmPegorinExecute(Sender: TObject);
begin
  CriarTelaSistema(FrmSmPegorin, TFrmSmPegorin);
end;

procedure TfrmMenuPrincipal.ActSmPortalExecute(Sender: TObject);
begin
  CriarTelaSistema(FrmPortal, TFrmPortal);
end;

procedure TfrmMenuPrincipal.ActSmRecantoDoPaoExecute(Sender: TObject);
begin
  CriarTelaSistema(FrmSmRecantoDoPao, TFrmSmRecantoDoPao);
end;

procedure TfrmMenuPrincipal.ActSmRioBrancoExecute(Sender: TObject);
begin
  CriarTelaSistema(FrmSmRioBranco, TFrmSmRioBranco);
end;

procedure TfrmMenuPrincipal.ActSmSaoFranciscoExecute(Sender: TObject);
begin
  CriarTelaSistema(FrmSmSaoFrancisco, TFrmSmSaoFrancisco);
end;

procedure TfrmMenuPrincipal.ActSmSuperComprasGestorExecute(Sender: TObject);
begin
  CriarTelaSistema(FrmSmSuperComprasGestor, TFrmSmSuperComprasGestor);
end;

procedure TfrmMenuPrincipal.ActSmNossoPaoExecute(Sender: TObject);
begin
  CriarTelaSistema(FrmSmNossoPao, TFrmSmNossoPao);
end;

procedure TfrmMenuPrincipal.ActSmTrioExecute(Sender: TObject);
begin
  CriarTelaSistema(FrmAriusTrio, TFrmAriusTrio);
end;

procedure TfrmMenuPrincipal.ActSmViaExecute(Sender: TObject);
begin
  CriarTelaSistema(FrmSmVia, TFrmSmVia);
end;

procedure TfrmMenuPrincipal.CkbAutomaticoClick(Sender: TObject);
var x: integer;
begin
   ArqConf.WriteBool('CheckBox', 'Abrir Tela', CkbAutomatico.Checked);
   if CkbAutomatico.Checked then
   begin
     if TreeViewMenu.Selected.AbsoluteIndex > 0 then
     begin
       for x := 0 to ActionList1.ActionCount - 1 do
        if TAction(ActionList1.Actions[x]).Caption = Tela then
        begin
          ActionList1.Actions[x].Execute;
        end;
     end;
   end;
end;

procedure TfrmMenuPrincipal.ConstruirArvore;
var
  i: Integer;
  no: TTreeNode;
  ac: TAction;
begin
  for i := 0 to Pred(ActionList1.ActionCount) do
  begin
    ac := TAction(ActionList1.Actions[i]);
    with TreeViewMenu.Items do
    begin
      no := EncontrarNo(ac.Category);
      if no=nil then
        no := Add(GetFirstNode, ac.Category);
        TreeViewMenu.Selected := AddChild(no, ac.Caption);
    end;
  end;
end;

procedure TfrmMenuPrincipal.CriarTelaSistema(Form: TForm; FormClass :TFormClass);
begin
   if not Verdade then
   begin
    Form := FormClass.Create(Self);
    Verdade := True;
   end;
end;

function TfrmMenuPrincipal.EncontrarNo(const aNome: string): TTreeNode;
var
  i: Integer;
begin
  Result := nil;
  with TreeViewMenu.Items do
  begin
    for i := 0 to Pred(Count) do
      if Item[i].Text= aNome then
      begin
        Result := Item[i];
        Break;
      end;
  end;
end;

procedure TfrmMenuPrincipal.FormCreate(Sender: TObject);
begin
  ArqConf := TIniFile.Create(ExtractFilePath(Application.ExeName) + 'gerador.ini');
  Verdade := False;
  ConstruirArvore;
  CkbAutomatico.Checked := ArqConf.ReadBool('CheckBox', 'Abrir Tela', False);
end;

procedure TfrmMenuPrincipal.FormShow(Sender: TObject);
var x :integer;
begin
  Tela := ArqConf.ReadString('TreeView', 'Ultima Tela', '');
 if CkbAutomatico.Checked then
 begin
   if TreeViewMenu.Selected.AbsoluteIndex > 0 then
   begin
     for x := 0 to ActionList1.ActionCount - 1 do
      if TAction(ActionList1.Actions[x]).Caption = Tela then
      begin
        ActionList1.Actions[x].Execute;
      end;
   end;
 end;
end;

procedure TfrmMenuPrincipal.TreeViewMenuClick(Sender: TObject);
var
  x : Integer;
  Indice, CaptionAct, NomeAct : String;
begin
  if TreeViewMenu.Selected.AbsoluteIndex > 0 then
  begin
    for x := 0 to ActionList1.ActionCount - 1 do
     if TAction(ActionList1.Actions[x]).Caption = TreeViewMenu.Selected.Text then
       ActionList1.Actions[x].Execute;
       ArqConf.WriteString('TreeView', 'Ultima Tela', TreeViewMenu.Selected.Text);
  end;

end;

end.
