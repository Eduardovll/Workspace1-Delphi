unit UFrmModelo;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs,
  Vcl.StdCtrls, UClasses,
  XPMan, ComCtrls, DB, SqlExpr, FMTBcd, ADODB,
  ExtCtrls, DateUtils, jpeg, TabNotBk, ShellAPI, IniFiles, WideStrings,
  DBXFirebird, DBXMsSQL, Filectrl, Math, Provider, DBClient, pngimage,
  Vcl.ImgList, Vcl.Buttons, Vcl.Menus, Data.DBXOracle, UProgresso, UUtilidades, xProc,
  UFrmSetParceiro, UFrmSetFiscal, UFrmSetProdutos ,UFrmSetOutros;//, dxGDIPlusClasses;

type
  TFrmModeloSis = class(TForm)
    PctArquivos: TPageControl;
    AbaParceiros: TTabSheet;
    CkbFornecedor: TCheckBox;
    CkbCondPagForn: TCheckBox;
    CkbDivisaoForn: TCheckBox;
    CkbTransportadora: TCheckBox;
    CkbStatusPdv: TCheckBox;
    CkbCliente: TCheckBox;
    CkbCondPagCli: TCheckBox;
    AbaProdutos: TTabSheet;
    CkbProdSimilar: TCheckBox;
    CkbProduto: TCheckBox;
    CkbCodigoBarras: TCheckBox;
    CkbProdLoja: TCheckBox;
    CkbProdForn: TCheckBox;
    CkbComposicao: TCheckBox;
    AbaFiscal: TTabSheet;
    CkbNFFornec: TCheckBox;
    CkbOutrasNFs: TCheckBox;
    CkbNFTransf: TCheckBox;
    CkbNFClientes: TCheckBox;
    CkbTributacao: TCheckBox;
    AbaOutros: TTabSheet;
    CkbAjuste: TCheckBox;
    Financeiro: TTabSheet;
    CkbFinanceiro: TCheckBox;
    CkbFinanceiroPagar: TCheckBox;
    CkbFinanceiroReceber: TCheckBox;
    CkbVenda: TCheckBox;
    CkbPlContas: TCheckBox;
    CkbSeGruSub: TCheckBox;
    CkbMapaResumo: TCheckBox;
    GbxData: TGroupBox;
    Label2: TLabel;
    Label3: TLabel;
    DtpInicial: TDateTimePicker;
    DtpFinal: TDateTimePicker;
    GroupBox: TGroupBox;
    GrpCamBanco: TGroupBox;
    EdtCamBanco: TEdit;
    GbxCamArquivo: TGroupBox;
    EdtCamArquivo: TEdit;
    PCBancoDados: TPageControl;
    TabOracle: TTabSheet;
    edtSchema: TEdit;
    Label5: TLabel;
    edtSenhaOracle: TEdit;
    Label4: TLabel;
    TabMySql: TTabSheet;
    edtHostMSql: TEdit;
    edtNomeBDMSql: TEdit;
    edtUserMSql: TEdit;
    edtPortaMSql: TEdit;
    edtSenhaMSql: TEdit;
    StaticText1: TStaticText;
    StaticText2: TStaticText;
    StaticText3: TStaticText;
    StaticText4: TStaticText;
    StaticText5: TStaticText;
    edtInst: TEdit;
    Label1: TLabel;
    CkbReceitas: TCheckBox;
    CkbInfoNutricionais: TCheckBox;
    CkbNf: TCheckBox;
    edtIpOra: TEdit;
    Label6: TLabel;
    CkbProdComprador: TCheckBox;
    TabSqlServer: TTabSheet;
    EdtIPSQLServer: TEdit;
    EdtUserSQLServer: TEdit;
    EdtBancoSQLServer: TEdit;
    EdtSenhaSQLServer: TEdit;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    CkbEnderecoCliente: TCheckBox;
    CkbFinanceiroPagarEmAbertos: TCheckBox;
    CkbFinanceiroReceberEmAberto: TCheckBox;
    CkbFinanceiroReceberCartoes: TCheckBox;
    BtnGerar: TBitBtn;
    BitSair: TBitBtn;
    GrpCamBancoSoliduss: TGroupBox;
    EdtCamBancoSoliduss: TEdit;
    CkbNcm: TCheckBox;
    BtnAbPastaArq: TSpeedButton;
    BtnAltCamBanco: TSpeedButton;
    BtnAltCamBancoSoliduss: TSpeedButton;
    BtnAltCamArq: TSpeedButton;
    CkbDecomposicao: TCheckBox;
    CkbFinanceiroReceberBoleto: TCheckBox;
    MainMenu1: TMainMenu;
    CkbFinanceiroReceberCheque: TCheckBox;
    CkbProdLocalizacao: TCheckBox;
    CkbProdProducao: TCheckBox;
    ckbGrade: TCheckBox;
    CkbFormatado: TCheckBox;
    rdgExtensao: TRadioGroup;
    CkbCest: TCheckBox;
    LblVesao: TLabel;
    ImgLogo: TImage;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BtnAltCamArqClick(Sender: TObject);
    procedure BtnAbPastaArqClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure PorData;
    procedure EdtCamArquivoChange(Sender: TObject);
    procedure DtpInicialChange(Sender: TObject);
    procedure GerarArquivos;
    procedure BtnGerar12Click(Sender: TObject);
    procedure CkbNFFornecClck(Sender: TObject);
    procedure edtSchemaChange(Sender: TObject);
    procedure edtSenhaOracleChange(Sender: TObject);
    procedure edtHostMSqlChange(Sender: TObject);
    procedure edtPortaMSqlChange(Sender: TObject);
    procedure edtNomeBDMSqlChange(Sender: TObject);
    procedure edtUserMSqlChange(Sender: TObject);
    procedure edtSenhaMSqlChange(Sender: TObject);
    procedure edtInstChange(Sender: TObject);
    procedure edtIpOraChange(Sender: TObject);
    procedure EdtIPSQLServerChange(Sender: TObject);
    procedure EdtUserSQLServerChange(Sender: TObject);
    procedure EdtBancoSQLServerChange(Sender: TObject);
    procedure EdtSenhaSQLServerChange(Sender: TObject);
    procedure CkbPlContasClick(Sender: TObject);
    procedure BtnGerarClick(Sender: TObject);
    procedure BitSairClick(Sender: TObject);
    procedure CkbClienteSensattaClick(Sender: TObject);
    procedure EdtCamBancoSolidussChange(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure BtnAltCamBancoSolidussClick(Sender: TObject);
    procedure CkbVendaClick(Sender: TObject);
    procedure CkbNFClientesClick(Sender: TObject);
    procedure CkbNFTransfClick(Sender: TObject);
    procedure CkbFinanceiroPagarEmAbertosClick(Sender: TObject);
    procedure CkbFinanceiroReceberEmAbertoClick(Sender: TObject);
    procedure CkbFinanceiroReceberChequeClick(Sender: TObject);
    procedure EdtCamBancoExit(Sender: TObject);
    procedure BtnAltCamBancoClick(Sender: TObject);
  private
    procedure SetGerarCodPagCli;
    procedure SetEntraNF( pLayout :TLayout );
    procedure SetEmissaoNF( pLayout :TLayout );
    procedure SetTransferenciaNF( pLayout :TLayout );
    procedure SetOutrasNF( pLayout :TLayout );
    procedure SetFileExtensionInLayout(Layout : TLayout);

    { Private declarations }
  public
    { Public declarations }
    procedure GerarCliente; Virtual; Abstract;
    procedure GerarFornecedor; Virtual; Abstract;

    procedure GerarSecao; Virtual; Abstract;
    procedure GerarGrupo; Virtual; Abstract;
    procedure GerarSubGrupo; Virtual; Abstract;

    procedure GerarInfoNutricionais; Virtual; Abstract;
    procedure GerarReceitas; Virtual; Abstract;
    procedure GerarProduto; Virtual; Abstract;
    procedure GerarProdLoja; Virtual; Abstract;
    procedure GerarCodigoBarras; Virtual; Abstract;
    procedure GerarProdSimilar; Virtual; Abstract;
    procedure GerarStatusPdv; Virtual; Abstract;
    procedure GerarProdForn; Virtual; Abstract;
    procedure GerarFinanceiro( Tipo, Situacao :Integer ); Virtual; Abstract;
    procedure GerarDecomposicao; virtual; abstract;
    procedure GerarComposicao; Virtual; Abstract;
    procedure GerarCEST; Virtual; Abstract;

    procedure GerarReceber; virtual; Abstract;
    procedure GerarFinanceiroPagar(Aberto:String); Virtual; Abstract;
    procedure GerarFinanceiroReceber(Aberto:String); Virtual; Abstract;
    procedure GerarFinanceiroReceberBoleto; Virtual; Abstract;
    procedure GerarFinanceiroReceberCheque; Virtual; Abstract;
    procedure GerarFinanceiroReceberCartao; Virtual; Abstract;

    procedure GerarNFClientes; Virtual; Abstract;
    procedure GerarNFitensClientes; Virtual; Abstract;

    procedure GerarNFFornec; Virtual; Abstract;
    procedure GerarNFitensFornec; Virtual; Abstract;

    procedure GerarVenda; Virtual; Abstract;
    procedure GerarTributacao; Virtual; Abstract;

    procedure GerarNCM; Virtual; Abstract;
    procedure GerarNCMUF; Virtual; Abstract;

    procedure GerarMapaResumo; Virtual; Abstract;
    procedure GerarDivisaoForn; Virtual; Abstract;
    procedure GerarCondPagForn; Virtual; Abstract;
    procedure GerarTransportadora; Virtual; Abstract;
    procedure GerarCondPagCli; Virtual; Abstract;


    procedure GerarNFTransf; Virtual; Abstract;
    procedure GerarNFItensTransf; Virtual; Abstract;

    procedure GerarOutrasNFs; Virtual; Abstract;
    procedure GerarItensOutrasNFs; Virtual; Abstract;

    procedure GerarNf; Virtual; Abstract;
    procedure GerarItensNf; Virtual; Abstract;

    procedure GerarAjuste; Virtual; Abstract;
    procedure GerarPlanoContas; Virtual;  Abstract;
    procedure GerarEnderecocobrancacliente; virtual; Abstract;
    procedure GerarProdComprador; Virtual;  Abstract;
    procedure GerarFinanceiroReceberCartoes; Virtual;  Abstract;
    procedure GeraLocalizacao; Virtual; Abstract;
    procedure GerarProducao; Virtual; Abstract;

    var Layout     : TLayout;
        Parceiro   : TParceiro;
        Fiscal     : TFiscal;
        Produto    : TProdutos;
        Outros     : TOutros;
        Financeiros: TFinanceiro;
  end;

  function SetCountTotal(Qry :string ; dataini :string = '19/09/1985'; datafim :String = '19/09/1985' ; dataini2 :string = '19/09/1985'; datafim2 :String = '19/09/1985'):Integer; overload;
  function SetCountTotal(Qry: String; Cnx : TADOConnection) : Integer; overload;



  procedure AdicionarCaminhoDlls;

var
  FrmModeloSis: TFrmModeloSis;
  TipoBD: Integer;
  Nome, Tela: String;
  Linha: String;
  ArqConf, ArqConfSolidus : TIniFile;

  QryPrincipal :TSQLQuery;
  ScnBanco, OrclBanco, SqlServerBanco, MySqlBanco, FBCon: TSQLConnection;
  AcoBanco: TADOConnection;

  v_SQL      : TSQLQuery;
  v_PosSQL   : TADOQuery;

  Fornecedores, DivFornecedor, CondPagForn, Transportadoras, StatusPDV ,Clientes, EndCobrancaCliente, CondPagcliente,

  Secao, Grupo, SubGrupo, ProdutosSimilar, Produtos, CodBarras, NCM, NCMuf, ProdutosLoja, ComposicaoProduto, DecompProduto, Receitas, InfoNutricional,
  ProdutoComprador, ProdutoLocalizacao, ProdutoFornecedor, Producao,

  Tributacoes, NFfornecedores, NFClientes, NFTransferencia, OutrasNF, NFDadosFiscais,
  NFfornItens, NFCliItens, NFTransfItens, OutrasNFItens ,
  TituloFinanceiro, Cest,

  MapaResumo, Ajuste, Vendas, PlanoContas :TLayout;
  Verdade: Boolean;


implementation

{$R *.dfm}

procedure TFrmModeloSis.BitSairClick(Sender: TObject);
begin
  close;
end;

procedure TFrmModeloSis.BtnAbPastaArqClick(Sender: TObject);
begin
  ShellExecute(Handle, 'OPEN', PChar(EdtCamArquivo.Text), nil, nil, SW_NORMAL);
end;

procedure TFrmModeloSis.BtnAltCamArqClick(Sender: TObject);
var
  Dir: String;
begin
  Dir := EdtCamArquivo.Text;
  if SelectDirectory(
    'Selecione o local aonde serão gravados os arquivos de texto:', '', Dir,
    [sdNewUI, sdNewFolder]) then
    EdtCamArquivo.Text := Dir;
end;

procedure TFrmModeloSis.BtnAltCamBancoClick(Sender: TObject);
begin
  ShellExecute(Handle, 'OPEN', PChar(EdtCamBanco.Text), nil, nil, SW_NORMAL);
end;

procedure TFrmModeloSis.BtnAltCamBancoSolidussClick(Sender: TObject);
begin
  ShellExecute(Handle, 'OPEN', PChar(EdtCamArquivo.Text), nil, nil, SW_NORMAL);
end;

procedure TFrmModeloSis.BtnGerar12Click(Sender: TObject);
begin
  GerarArquivos;
end;

procedure TFrmModeloSis.BtnGerarClick(Sender: TObject);

begin
  GerarArquivos;
  Layout.SetDirectory := EdtCamArquivo.Text;
end;

procedure TFrmModeloSis.CkbClienteSensattaClick(Sender: TObject);
begin
  GerarArquivos;
end;

procedure TFrmModeloSis.CkbFinanceiroPagarEmAbertosClick(Sender: TObject);
begin
  PorData;
end;

procedure TFrmModeloSis.CkbFinanceiroReceberChequeClick(Sender: TObject);
begin
 PorData;
end;

procedure TFrmModeloSis.CkbFinanceiroReceberEmAbertoClick(Sender: TObject);
begin
 PorData;
end;

procedure TFrmModeloSis.CkbNFClientesClick(Sender: TObject);
begin
  PorData;
end;

procedure TFrmModeloSis.CkbNFFornecClck(Sender: TObject);
begin
 PorData;
end;

procedure TFrmModeloSis.CkbNFTransfClick(Sender: TObject);
begin
 PorData;
end;

procedure TFrmModeloSis.CkbPlContasClick(Sender: TObject);
begin
 PorData;
end;

procedure TFrmModeloSis.CkbVendaClick(Sender: TObject);
begin
 PorData;
end;

procedure TFrmModeloSis.DtpInicialChange(Sender: TObject);
begin
  ArqConf.WriteDate(Nome, 'Data Inicial', DtpInicial.Date);
  ArqConf.WriteDate(Nome, 'Data Final', DtpFinal.Date);
end;

procedure TFrmModeloSis.EdtBancoSQLServerChange(Sender: TObject);
begin
  ArqConf.WriteString(Nome, 'BDSqlServerBD', EdtBancoSQLServer.Text);
end;

procedure TFrmModeloSis.EdtCamArquivoChange(Sender: TObject);
begin
  ArqConf.WriteString(Nome, 'Pasta Textos', EdtCamArquivo.Text);
end;

procedure TFrmModeloSis.EdtCamBancoExit(Sender: TObject);
begin
  ArqConf.WriteString(Nome, 'Banco', EdtCamBanco.Text);
end;

procedure TFrmModeloSis.EdtCamBancoSolidussChange(Sender: TObject);
begin
  ArqConf.WriteString(Nome, EdtCamBancoSoliduss.Name, EdtCamBancoSoliduss.Text);
end;

procedure TFrmModeloSis.edtHostMSqlChange(Sender: TObject);
begin
  ArqConf.WriteString(Nome, 'BDMSqlHost', edtHostMSql.Text);
end;

procedure TFrmModeloSis.edtInstChange(Sender: TObject);
begin
  ArqConf.WriteString(Nome, 'BDOraInst', edtInst.Text);
end;

procedure TFrmModeloSis.edtIpOraChange(Sender: TObject);
begin
  ArqConf.WriteString(Nome, 'BDOraIp', edtIpOra.Text);
end;

procedure TFrmModeloSis.EdtIPSQLServerChange(Sender: TObject);
begin
  ArqConf.WriteString(Nome, 'BDSqlServerIp', EdtIPSQLServer.Text);
end;

procedure TFrmModeloSis.edtNomeBDMSqlChange(Sender: TObject);
begin
  ArqConf.WriteString(Nome, 'BDMsqlNomeBD', edtNomeBDMSql.Text);
end;

procedure TFrmModeloSis.edtPortaMSqlChange(Sender: TObject);
begin
  ArqConf.WriteString(Nome, 'BDMSqlPorta', edtPortaMSql.Text);
end;

procedure TFrmModeloSis.edtSchemaChange(Sender: TObject);
begin
  ArqConf.WriteString(Nome, 'BDOraSchema', edtSchema.Text);
end;

procedure TFrmModeloSis.edtSenhaMSqlChange(Sender: TObject);
begin
  ArqConf.WriteString(Nome, 'BDMSqlSenha', edtSenhaMSql.Text);
end;

procedure TFrmModeloSis.edtSenhaOracleChange(Sender: TObject);
begin
  ArqConf.WriteString(Nome, 'BDOraSenha', edtSenhaOracle.Text);
end;

procedure TFrmModeloSis.EdtSenhaSQLServerChange(Sender: TObject);
begin
  ArqConf.WriteString(Nome, 'BDSqlServerSenha', EdtSenhaSQLServer.Text);
end;

procedure TFrmModeloSis.edtUserMSqlChange(Sender: TObject);
begin
  ArqConf.WriteString(Nome, 'BDMSqlUser', edtUserMSql.Text);
end;

procedure TFrmModeloSis.EdtUserSQLServerChange(Sender: TObject);
begin
  ArqConf.WriteString(Nome, 'BDSqlServerUser', EdtUserSQLServer.Text);
end;

procedure TFrmModeloSis.SetEmissaoNF( pLayout :TLayout );
begin

end;

procedure TFrmModeloSis.SetEntraNF( pLayout :TLayout );
begin

end;

procedure TFrmModeloSis.SetGerarCodPagCli;
begin
//
end;

procedure TFrmModeloSis.SetOutrasNF(pLayout: TLayout);
begin
  //
end;

procedure TFrmModeloSis.SetFileExtensionInLayout(Layout : TLayout);
begin
  Layout.SetFileExtension(iif(rdgExtensao.ItemIndex = 0, fxtTxt, fxtCsv));
end;

procedure TFrmModeloSis.FormActivate(Sender: TObject);
begin
//if GbxData.Visible = False then
//     Image1.Left := 431
//  else
//     Image1.Left := 489;
end;

procedure TFrmModeloSis.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Verdade := False;
  Action := caFree;
  TForm(Nome) := nil;
end;

procedure TFrmModeloSis.FormCreate(Sender: TObject);
begin


  Nome := Name;
  // ArqConf := TIniFile.Create(ExtractFilePath(Application.ExeName) + 'gerador.ini');
  EdtCamArquivo.Text       := ArqConf.ReadString(Nome, 'Pasta Textos', '');
  EdtCamBanco.Text         := ArqConf.ReadString(Nome, 'Banco', '');
  EdtCamBancoSoliduss.Text := ArqConf.ReadString(Nome, EdtCamBancoSoliduss.Name, '');
  DtpInicial.Date          := ArqConf.ReadDate(Nome, 'Data Inicial', Date());
  DtpFinal.Date            := ArqConf.ReadDate(Nome, 'Data Final', Date());

  // Tipo Banco de Dados
  TipoBD := ArqConf.ReadInteger(Nome, 'BD', -1);
  // Oracle
  edtSchema.Text      := ArqConf.ReadString(Nome, 'BDOraSchema', '');
  edtSenhaOracle.Text := ArqConf.ReadString(Nome, 'BDOraSenha', '');
  edtInst.Text        := ArqConf.ReadString(Nome, 'BDOraInst', '');
  edtIpOra.Text       := ArqConf.ReadString(Nome, 'BDOraIp', '');

  // MySql
  edtHostMSql.Text   := ArqConf.ReadString(Nome, 'BDMSqlHost', '');
  edtPortaMSql.Text  := ArqConf.ReadString(Nome, 'BDMSqlPorta', '');
  edtNomeBDMSql.Text := ArqConf.ReadString(Nome, 'BDMsqlNomeBD', '');
  edtUserMSql.Text   := ArqConf.ReadString(Nome, 'BDMSqlUser', '');
  edtSenhaMSql.Text  := ArqConf.ReadString(Nome, 'BDMSqlSenha', '');

  // Sql Server
  EdtIPSQLServer.Text    := ArqConf.ReadString(Nome, 'BDSqlServerIp', '');
  EdtUserSQLServer.Text  := ArqConf.ReadString(Nome, 'BDSqlServerUser', '');
  EdtBancoSQLServer.Text := ArqConf.ReadString(Nome, 'BDSqlServerBD', '');
  EdtSenhaSQLServer.Text := ArqConf.ReadString(Nome, 'BDSqlServerSenha', '');

  Layout                 := TLayout.Create('Layout','Layout');
  Layout.SetDirectory    := EdtCamArquivo.Text;

  Fornecedores       := Parceiro.GetFieldFornecedor;
  DivFornecedor      := Parceiro.GetFieldDivisaoFornecedor;
  CondPagForn        := Parceiro.GetFieldCondFornecedor;
  Transportadoras    := Parceiro.GetFieldTransportadora;
  StatusPDV          := Parceiro.GetFieldStatusClientePDV;
  Clientes           := Parceiro.GetFieldClientes;
  EndCobrancaCliente := Parceiro.GetFieldEndCobCliente;
  CondPagcliente     := Parceiro.GetFieldPagamentoCliente;

  Secao    := Produto.GetFieldSecao;
  Grupo    := Produto.GetFieldGrupo;
  SubGrupo := Produto.GetFieldSubGrupo;

  ProdutosSimilar := Produto.GetFieldProdutosSimilares;
  Produtos := Produto.GetFieldProdutos;
  CodBarras := Produto.GetFieldCodigoBarra;

  NCM := Produto.GetFieldNCM;
  NCMuf := Produto.GetFieldNCMUF;
  Cest := Produto.GetFieldCest;

  ProdutosLoja := Produto.GetFieldProdutoLoja;
  ProdutoFornecedor := Produto.GetFieldProdutoFornedores;
  ComposicaoProduto := Produto.GetFieldComposicoesProdutos;
  DecompProduto := Produto.GetFieldDecomposicoesProdutos;
  Receitas := Produto.GetFieldReceitas;
  InfoNutricional := Produto.GetFieldInfoNutricionais;
  ProdutoComprador := Produto.GetFieldProdutoComprador;
  ProdutoLocalizacao := Produto.GetFieldProdutoLocalizacao;

  Tributacoes := Fiscal.GetFieldTributacoes;
  NFfornecedores := Fiscal.GetFieldNFornecedores;
  NFfornItens := Fiscal.GetFieldItensNFornecedores;

  NFClientes := Fiscal.GetFieldINFClientes;
  NFCliItens := Fiscal.GetFieldItensNFClientes;

  NFTransferencia := Fiscal.GetFieldTransferencia;
  NFTransfItens := Fiscal.GetFieldItensTransferencia;

  OutrasNF := Fiscal.GetFieldOutrasNFs;
  OutrasNFItens := Fiscal.GetFieldItensOutrasNFs;

  TituloFinanceiro := Financeiros.GetFieldFinanceiro;

  MapaResumo       := Outros.GetFieldMapaResumo;
  Ajuste           := Outros.GetFieldAjuste;
  Vendas           := Outros.GetFieldVendas;
  PlanoContas      := Outros.GetFieldPlanoContas;
  Producao         := Produto.GetFieldProducao;

  v_SQL            := TSQLQuery.Create(Self);
  v_PosSQL         := TADOQuery.Create(Self);

  case TipoBD of
    0 :v_SQL.SQLConnection := ScnBanco;
    3 :v_SQL.SQLConnection := OrclBanco;
  end;

end;

procedure TFrmModeloSis.GerarArquivos;
var
  Erro: String;
  QtdRegProg: TStringList;
  Completo: Boolean;
begin
  if not Assigned(QryPrincipal) then
     QryPrincipal := TSQLQuery.Create(Self);

  PathDir := EdtCamArquivo.Text;

  Erro := '';

  if EdtCamArquivo.Text = '' then
  begin
    MessageDlg(
      'Você não selecionou o caminho em que será salvo o arquivo de texto.',
      mtError, [mbOK], 0);
    Exit;
  end;

  if (DtpInicial.Date > DtpFinal.Date) then
  begin
    MessageDlg('A data inicial não pode ser maior do que a data final.',
      mtError, [mbOK], 0);
    Exit;
  end;

  Cursor := crHourGlass;
  PctArquivos.Enabled := False;
  BtnGerar.Enabled := False;

  FrmProgresso := TFrmProgresso.Create(Self);
  FrmProgresso.ExibirGrade := CkbGrade.Checked;
  FrmProgresso.ExibirFormatado := CkbFormatado.Checked;
  FrmProgresso.Init;
  FrmProgresso.Show;

  QtdErros := 0;
  FrmProgresso.LblQtdErros.Caption := '';

  Completo := False;
  if (edtSenhaOracle.Text <> '') and (edtSchema.Text <> '') and
    (edtInst.Text <> '') and (edtIpOra.Text <> '') then
    TipoBD := 3
  else if (edtSenhaMSql.Text <> '') and (edtHostMSql.Text <> '') and
    (edtNomeBDMSql.Text <> '') and (edtUserMSql.Text <> '') and
    (edtPortaMSql.Text <> '') then
    TipoBD := 4
  else if (EdtSenhaSQLServer.Text <> '') and (EdtIPSQLServer.Text <> '') and
     (EdtBancoSQLServer.Text <> '') and (EdtUserSQLServer.Text <> '') then
    TipoBD := 5;
  case TipoBD of
    1:
      begin
        if EdtCamBanco.Text = '' then
          AbrirBancoFB(Self, EdtCamBanco);
        CriarFB(EdtCamBanco);
        QryPrincipal.Active := False;
        QryPrincipal.SQLConnection := ScnBanco;
      end;
    2:
      begin
        if EdtCamBanco.Text = '' then
          AbrirBancoAccess(Self, EdtCamBanco);
        CriarAccess(EdtCamBanco);
      end;
    3: begin
        CriarOracle(edtSenhaOracle, edtSchema, edtInst, edtIpOra);
        QryPrincipal.SQLConnection := OrclBanco;
      end;
    4: begin
        CriarMySql(edtHostMSql, edtPortaMSql, edtNomeBDMSql, edtUserMSql,
        edtSenhaMSql);
        QryPrincipal.SQLConnection := MySqlBanco;
       end;
    5: begin
      CriarSQLServer(EdtSenhaSQLServer, EdtUserSQLServer, EdtIPSQLServer,
        EdtBancoSQLServer);
      QryPrincipal.Active := False;
      QryPrincipal.SQLConnection := SqlServerBanco;
    end;
    7:
      AbrirBancoDBF(Self, EdtCamBanco);
  end;

  While (not Cancelar) and (not Completo) do
  begin
    Linha := '';
    try
      if CkbFornecedor.Checked then
      begin
        CkbFornecedor.Checked := False;
        SetFileExtensionInLayout(Fornecedores);
        Fornecedores.Start;
        Layout := Fornecedores;
        GerarFornecedor;
        Fornecedores.Finish;
      end
      else if CkbCest.checked then
      begin
        CkbCest.Checked := False;
        SetFileExtensionInLayout(Cest);
        Cest.Start;
        Layout := Cest;
        GerarCEST;
        Cest.Finish;
      end
      else if CkbDivisaoForn.Checked then
      begin
        CkbDivisaoForn.Checked := False;
        SetFileExtensionInLayout(DivFornecedor);
        DivFornecedor.Start;
        Layout := Divfornecedor;
        GerarDivisaoForn;
        DivFornecedor.Finish;
      end
      else if CkbCondPagForn.Checked then
      begin
        CkbCondPagForn.Checked := False;
        SetFileExtensionInLayout(CondPagForn);
        CondPagForn.Start;
        Layout := CondPagForn;
        GerarCondPagForn;
        CondPagForn.Finish;
      end
      else if CkbTransportadora.Checked then
      begin
        CkbTransportadora.Checked := False;
        SetFileExtensionInLayout(Transportadoras);
        Transportadoras.Start;
        Layout := Transportadoras;
        GerarTransportadora;
        Transportadoras.Finish;
      end
      else if CkbStatusPdv.Checked then
      begin
        CkbStatusPdv.Checked := False;
        SetFileExtensionInLayout(StatusPDV);
        StatusPDV.Start;
        Layout := StatusPDV;
        GerarStatusPdv;
        StatusPDV.Finish;
      end
      else if CkbCliente.Checked then
      begin
        CkbCliente.Checked := False;
        SetFileExtensionInLayout(Clientes);
        Clientes.Start;
        Layout := Clientes;
        GerarCliente;
        Clientes.Finish;
      end
      else if CkbEnderecoCliente.Checked then
      begin
        CkbEnderecoCliente.Checked := False;
        GerarEnderecocobrancacliente;
      end


      else if CkbCondPagCli.Checked then
      begin
        CkbCondPagCli.Checked := False;
        SetFileExtensionInLayout(CondPagcliente);
        CondPagcliente.Start;
        Layout := CondPagcliente;
        GerarCondPagCli;
        CondPagcliente.Finish;
      end
      else if CkbSeGruSub.Checked then
      begin
        CkbSeGruSub.Checked := False;

        SetFileExtensionInLayout(Secao);
        Secao.Start;
        Layout := Secao;
        GerarSecao;
        Secao.Finish;


        SetFileExtensionInLayout(Grupo);
        Grupo.Start;
        Layout := Grupo;
        GerarGrupo;
        Grupo.Finish;

        SetFileExtensionInLayout(SubGrupo);
        SubGrupo.Start;
        Layout := SubGrupo;
        GerarSubGrupo;
        SubGrupo.Finish;
      end
      else if CkbInfoNutricionais.Checked then
      begin
        CkbInfoNutricionais.Checked := False;
        SetFileExtensionInLayout(InfoNutricional);
        InfoNutricional.Start;
        Layout := InfoNutricional;
        GerarInfoNutricionais;
        InfoNutricional.Finish;
      end
      else if CkbReceitas.Checked then
      begin
        CkbReceitas.Checked := False;
        SetFileExtensionInLayout(Receitas);
        Receitas.Start;
        Layout := Receitas;
        GerarReceitas;
        Receitas.Finish;
      end
      else if CkbProdSimilar.Checked then
      begin
        CkbProdSimilar.Checked := False;
        SetFileExtensionInLayout(ProdutosSimilar);
        ProdutosSimilar.Start;
        Layout := ProdutosSimilar;
        GerarProdSimilar;
        ProdutosSimilar.Finish;
      end
      else if CkbProduto.Checked then
      begin
        CkbProduto.Checked := False;
        SetFileExtensionInLayout(Produtos);
        Produtos.Start;
        Layout := Produtos;
        GerarProduto;
        Produtos.Finish;
      end
      else if CkbCodigoBarras.Checked then
      begin
        CkbCodigoBarras.Checked := False;
        SetFileExtensionInLayout(CodBarras);
        CodBarras.Start;
        Layout := CodBarras;
        GerarCodigoBarras;
        CodBarras.Finish;
      end
      else if CkbNcm.Checked then
      begin
        CkbNcm.Checked := False;

        SetFileExtensionInLayout(NCM);
        NCM.Start;
        Layout := NCM;
        GerarNCM;
        NCM.Finish;

        SetFileExtensionInLayout(NCMuf);
        NCMuf.Start;
        Layout := NCMuf;
        GerarNCMUF;
        NCMuf.Finish;
      end

      else if CkbProdLoja.Checked then
      begin
        CkbProdLoja.Checked := False;
        SetFileExtensionInLayout(ProdutosLoja);
        ProdutosLoja.Start;
        Layout := ProdutosLoja;
        GerarProdLoja;
        ProdutosLoja.Finish;
      end
      else if CkbProdForn.Checked then
      begin
        CkbProdForn.Checked := False;
        SetFileExtensionInLayout(DivFornecedor);
        ProdutoFornecedor.Start;
        Layout := ProdutoFornecedor;
        GerarProdForn;
        ProdutoFornecedor.Finish;
      end
      else if CkbComposicao.Checked then
      begin
        CkbComposicao.Checked := False;
        SetFileExtensionInLayout(DivFornecedor);
        ComposicaoProduto.Start;
        Layout := ComposicaoProduto;
        GerarComposicao;
        ComposicaoProduto.Finish;
      end
      else if CkbDecomposicao.Checked then
      begin
        CkbDecomposicao.Checked := False;
        SetFileExtensionInLayout(DivFornecedor);
        DecompProduto.Start;
        Layout := DecompProduto;
        GerarDecomposicao;
        DecompProduto.Finish;
      end
      else if CkbTributacao.Checked then
      begin
        CkbTributacao.Checked := False;
        SetFileExtensionInLayout(DivFornecedor);
        Tributacoes.Start;
        Layout := Tributacoes;
        GerarTributacao;
        Tributacoes.Finish;
      end
      else if CkbNFFornec.Checked then
      begin
        CkbNFFornec.Checked := False;

        SetFileExtensionInLayout(NFfornecedores);
        NFfornecedores.Start;
        Layout := NFfornecedores;
        GerarNFFornec;
        NFfornecedores.Finish;

        SetFileExtensionInLayout(NFfornItens);
        NFfornItens.Start;
        Layout := NFfornItens;
        GerarNFitensFornec;
        NFfornItens.Finish;

      end
      else if CkbNFClientes.Checked then
      begin
        CkbNFClientes.Checked := False;

        SetFileExtensionInLayout(NFClientes);
        NFClientes.Start;
        Layout := NFClientes;
        GerarNFClientes;
        NFClientes.Finish;

        SetFileExtensionInLayout(NFCliItens);
        NFCliItens.Start;
        Layout := NFCliItens;
        GerarNFitensClientes;
        NFCliItens.Finish;
      end
      else if CkbNFTransf.Checked then
      begin
        CkbNFTransf.Checked := False;

        SetFileExtensionInLayout(NFTransferencia);
        NFTransferencia.Start;
        Layout := NFTransferencia;
        GerarNFTransf;
        NFTransferencia.Finish;

        SetFileExtensionInLayout(NFTransfItens);
        NFTransfItens.Start;
        Layout := NFTransfItens;
        GerarNFItensTransf;
        NFTransfItens.Finish;
      end
      else if CkbOutrasNFs.Checked then
      begin
        CkbOutrasNFs.Checked := False;
        SetFileExtensionInLayout(OutrasNF);
        OutrasNF.Start;
        Layout := OutrasNF;
        GerarOutrasNFs;
        OutrasNF.Finish;
      end
      else if CkbNf.Checked then
      begin
        CkbNf.Checked := False;
        GerarNf;
      end
      else if CkbFinanceiro.Checked then
      begin
        CkbFinanceiro.Checked := False;
        SetFileExtensionInLayout(TituloFinanceiro);
        TituloFinanceiro.Start('Financeiro', 'FINANCEIRO');
        Layout := TituloFinanceiro;
        GerarFinanceiro( 0,0 );
        TituloFinanceiro.Finish;
      end
      else if CkbFinanceiroPagar.Checked then
      begin
        CkbFinanceiroPagar.Checked := False;
        SetFileExtensionInLayout(TituloFinanceiro);
        TituloFinanceiro.Start('Financeiro à Pagar', 'FINANCEIRO_PAGAR');
        Layout := TituloFinanceiro;
        GerarFinanceiro( 1,2 );
        TituloFinanceiro.Finish;
      end
      else if CkbFinanceiroReceber.Checked then
      begin
        CkbFinanceiroReceber.Checked := False;
        SetFileExtensionInLayout(TituloFinanceiro);
        TituloFinanceiro.Start('Financeiro à Receber', 'FINANCEIRO_RECEBER');
        Layout := TituloFinanceiro;
        GerarFinanceiro( 2,2 );
        TituloFinanceiro.Finish;
      end
      else if CkbFinanceiroReceberBoleto.Checked then
      begin
        CkbFinanceiroReceberBoleto.Checked := False;
        SetFileExtensionInLayout(TituloFinanceiro);
        TituloFinanceiro.Start('Financeiro Boleto', 'FINANCEIRO_BOLETO');
        Layout := TituloFinanceiro;
        GerarFinanceiroReceberBoleto;
        TituloFinanceiro.Finish;
      end
      else if CkbFinanceiroReceberCheque.Checked then
      begin
        CkbFinanceiroReceberCheque.Checked := False;
        SetFileExtensionInLayout(TituloFinanceiro);
        TituloFinanceiro.Start('Financeiro Cheque', 'FINANCEIRO_CHEQUE');
        Layout := TituloFinanceiro;
        GerarFinanceiroReceberCheque;
        TituloFinanceiro.Finish;
      end
      else if CkbFinanceiroReceberCartoes.Checked then
      begin
        CkbFinanceiroReceberCartoes.Checked := False;
        SetFileExtensionInLayout(TituloFinanceiro);
        TituloFinanceiro.Start('Financeiro à Receber ( Cartão )', 'FINANCEIRO_CARTAO');
        Layout := TituloFinanceiro;
        GerarFinanceiro( 3,1 );
        TituloFinanceiro.Finish;
      end
      else if CkbFinanceiroPagarEmAbertos.Checked then
      begin
        CkbFinanceiroPagarEmAbertos.Checked := False;
        SetFileExtensionInLayout(TituloFinanceiro);
        TituloFinanceiro.Start('Financeiro á Pagar ( Títulos em Abertos )', 'FINANCEIRO_PAGAR_ABERTO');
        Layout := TituloFinanceiro;
        GerarFinanceiro( 1,1 );
        TituloFinanceiro.Finish;
      end
      else if CkbFinanceiroReceberEmAberto.Checked then
      begin
        CkbFinanceiroReceberEmAberto.Checked := False;
        SetFileExtensionInLayout(TituloFinanceiro);
        TituloFinanceiro.Start('Financeiro á Receber ( Títulos em Aberto )', 'FINANCEIRO_RECEBER_ABERTO');
        Layout := TituloFinanceiro;
        GerarFinanceiro( 2,1 );
        TituloFinanceiro.Finish;
      end
      else if CkbFinanceiroPagar.Checked then
      begin
        SetFileExtensionInLayout(TituloFinanceiro);
        TituloFinanceiro.Start('Financeiro á Pagar', 'FINANCEIRO_PAGAR');
        Layout := TituloFinanceiro;
        CkbFinanceiroPagar.Checked := False;
        GerarFinanceiro( 1,2 );
      end
      else if CkbMapaResumo.Checked then
      begin
        CkbMapaResumo.Checked := False;
        SetFileExtensionInLayout(MapaResumo);
        MapaResumo.Start;
        Layout := MapaResumo;
        GerarMapaResumo;
        MapaResumo.Finish;
      end
      else if CkbAjuste.Checked then
      begin
        CkbAjuste.Checked := False;
        SetFileExtensionInLayout(Ajuste);
        Ajuste.Start;
        Layout := Ajuste;
        GerarAjuste;
        Ajuste.Finish;
      end
      else if CkbVenda.Checked then
      begin
        CkbVenda.Checked := False;
        SetFileExtensionInLayout(Vendas);
        Vendas.Start;
        Layout := Vendas;
        GerarVenda;
        Vendas.Finish;
      end
      else if CkbPlContas.Checked then
      begin
        CkbPlContas.Checked := False;
        SetFileExtensionInLayout(PlanoContas);
        PlanoContas.Start;
        Layout := PlanoContas;
        GerarPlanoContas;
        PlanoContas.Finish;
      end
      else if CkbProdComprador.Checked then
      begin
        CkbProdComprador.Checked := False;
        SetFileExtensionInLayout(ProdutoComprador);
        ProdutoComprador.Start;
        Layout := ProdutoComprador;
        GerarProdComprador;
        ProdutoComprador.Finish;
      end
      else if CkbProdLocalizacao.Checked then
      begin
        CkbProdLocalizacao.Checked := False;
        SetFileExtensionInLayout(ProdutoLocalizacao);
        ProdutoLocalizacao.Start;
        Layout := ProdutoLocalizacao;
        GeraLocalizacao;
        ProdutoLocalizacao.Finish;
      end
      else if CkbProdProducao.Checked then
      begin
        CkbProdProducao.Checked := False;
        SetFileExtensionInLayout(Producao);
        Producao.Start;
        Layout := Producao;
        GerarProducao;
        Producao.Finish;
      end
      else
        Completo := True;

    except
      On E: Exception do
      begin
        //MessageDlg('Ocorreu o erro "' + E.Message + '" ao gerar o arquivo.',
        //  mtError, [mbOK], 0);
        //CloseFile(Arquivo);
        Erro := E.Message;
      end;
    end;

    FrmProgresso.LblAcaoAtual.Caption := '';
  end; // while

  if Erro <> '' then
  begin
    MessageDlg('A geração de texto falhou devido ao erro:' + #10#13 + Erro,
      mtError, [mbOK], 0);
    FrmProgresso.Free;
  end
  else if Cancelar then
  begin
    MessageDlg('A geração de arquivos de texto foi cancelada.', mtInformation,
      [mbOK], 0);
    Cancelar  := False;
    FrmProgresso.Free;
  end
  else
  begin
    FrmProgresso.PgbGeracao.Position := 100;
    QtdRegProg := StrSplit(FrmProgresso.LblRegistroAtual.Caption, '/');
    if QtdRegProg.Count = 2 then
      FrmProgresso.LblRegistroAtual.Caption := QtdRegProg[1] + '/' + QtdRegProg
        [1];
    // MessageBox(0, 'A geração de arquivos de texto está completa.', 'Information', MB_ICONINFORMATION or MB_OK);
    MessageDlg('A geração de arquivos de texto está completa.', mtInformation,
      [mbOK], 0);
    FrmProgresso.Free;
  end;

  // FrmProgresso.Hide;
  //FrmProgresso.Free;
  Encerrado := True;

  BtnGerar.Enabled := True;
  PctArquivos.Enabled := True;
  Cursor := crDefault;
end;


procedure TFrmModeloSis.PorData;
begin
  if CkbNFFornec.Checked or CkbNFClientes.Checked or CkbNFTransf.Checked or
    CkbOutrasNFs.Checked or CkbFinanceiro.Checked or CkbFinanceiroReceber.
    Checked or CkbFinanceiroPagar.Checked or CkbFinanceiroPagarEmAbertos.Checked or
    CkbFinanceiroReceberEmAberto.Checked or CkbVenda.Checked or CkbMapaResumo.Checked or
    CkbNf.Checked or CkbFinanceiroReceberBoleto.Checked or CkbFinanceiroReceberCheque.Checked or
    CkbFinanceiroReceberCartoes.Checked
    then
  Begin
//    Image1.Left := 580;
//    GbxData.Enabled := True;
      DtpInicial.Enabled := True;
      DtpFinal.Enabled := True;
  End
  else
  Begin
//    Image1.Left := 431;
//    GbxData.Enabled := False;
      DtpInicial.Enabled := False;
      DtpFinal.Enabled := False;
  End;
end;

procedure TFrmModeloSis.SetTransferenciaNF( pLayout :TLayout );
begin

end;

function SetCountTotal(Qry :string ; dataini :string = '19/09/1985'; datafim :String = '19/09/1985' ; dataini2 :string = '19/09/1985'; datafim2 :String = '19/09/1985' ): Integer;
begin
   with v_SQL do begin
    if SQLConnection = nil then
    begin
      case TipoBD of
       0,1: SQLConnection := ScnBanco;
       3: SQLConnection := OrclBanco;
      end;
    end;

    Close;
    SQL.Clear;

//    ShowMessage(qry);

    SQL.Add('SELECT COUNT(*) TOTAL FROM(' );
    SQL.Add(Qry);
    SQL.Add(')');

    if Assigned(Params.FindParam('INI')) and
       Assigned(Params.FindParam('FIM')) and
       Assigned(Params.FindParam('INI2')) and
       Assigned(Params.FindParam('FIM2')) then
    begin
     ParamByName('INI').AsDate := StrToDateTime(DataIni);
     ParamByName('FIM').AsDate := StrToDateTime(DataFim);
     ParamByName('INI2').AsDate := StrToDateTime(DataIni2);
     ParamByName('FIM2').AsDate := StrToDateTime(DataFim2);
    end
    else if Assigned(Params.FindParam('INI')) and
            Assigned(Params.FindParam('FIM')) then
    begin
     ParamByName('INI').AsDate := StrToDateTime(DataIni);
     ParamByName('FIM').AsDate := StrToDateTime(DataFim);
    end;


    Open;
    Result := FieldByName('TOTAL').AsInteger;
   end;
end;

function SetCountTotal(Qry: String; Cnx : TADOConnection): Integer;
begin
   v_PosSQL.Connection := Cnx;
   with v_PosSQL do begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT COUNT(*) TOTAL FROM(' );
    SQL.Add(Qry);
    SQL.Add(') AS TAB_TOTAL');

    Open;
    Result := FieldByName('TOTAL').AsInteger;
   end;
end;




procedure AdicionarCaminhoDlls;
begin
  SetDllDirectory(PChar(ExtractFilePath(Application.ExeName) + '\bin'));
  { DefinirVariavelDeSistema('PATH', ObterVariavelDeSistema('PATH') + ';'
    + ExtractFilePath(Application.ExeName) + '\bin'); }
end;

end.
