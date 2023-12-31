unit UFrmHiperLojao;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, UFrmModelo, Data.DBXOracle, Data.DB,
  Data.SqlExpr, Vcl.Menus, Vcl.StdCtrls, Vcl.Buttons, Vcl.ComCtrls,
  Vcl.Imaging.pngimage, Vcl.ExtCtrls, Data.DBXFirebird, Data.Win.ADODB,
  Vcl.Grids, Vcl.DBGrids, Datasnap.Provider, Datasnap.DBClient,
  //dxGDIPlusClasses,
  Math, OraProvider;

type
  TFrmHiperLojao = class(TFrmModeloSis)
    CbxLoja: TComboBox;
    lblLoja: TLabel;
    btnScriptDeposito: TButton;
    Label11: TLabel;
    btnCustoCheio: TButton;
    btnAtualizaEstoque: TButton;
    Label12: TLabel;
    btnEstoqueTroca: TButton;
    Label13: TLabel;
    Label14: TLabel;
    btnFidelidade: TButton;
    Label15: TLabel;
    btnFornPrefe: TButton;
    Label16: TLabel;
    procedure EdtCamBancoExit(Sender: TObject);
    procedure BtnGerarClick(Sender: TObject);
    procedure btnScriptDepositoClick(Sender: TObject);
    procedure CkbProdutoClick(Sender: TObject);
    procedure btnCustoCheioClick(Sender: TObject);
    procedure btnAtualizaEstoqueClick(Sender: TObject);
    procedure CkbProdLojaClick(Sender: TObject);
    procedure btnEstoqueTrocaClick(Sender: TObject);
    procedure CkbProdFornClick(Sender: TObject);
    procedure btnFidelidadeClick(Sender: TObject);
    procedure btnFornPrefeClick(Sender: TObject);
  private

    { Private declarations }
  public
    { Public declarations }
    procedure GerarCliente;           Override;
    procedure GerarFornecedor;        Override;
    procedure GerarCondPagForn;       Override;
    procedure GerarDivisaoForn;      Override;
    procedure GerarCondPagCli;       Override;
    procedure GerarTransportadora;      Override;
    procedure GerarCest; Override;

    procedure GerarSecao;           Override;
    procedure GerarGrupo;           Override;
    procedure GerarSubGrupo;        Override;

    procedure GerarProduto;           Override;
    procedure GerarCodigoBarras;      Override;
    procedure GerarProdLoja;          Override;
    procedure GerarNCM;               Override;
    procedure GerarNCMUF;                                 Override;
    procedure GerarProdSimilar;                           Override;
    procedure GerarProdForn;                              Override;
    procedure GerarInfoNutricionais;                      Override;
    procedure GerarDecomposicao;                          Override;
    procedure GerarComposicao;                            Override;
    procedure GerarProducao;                              Override;

    procedure GerarNFFornec;                              Override;
    procedure GerarNFitensFornec;                         Override;
    procedure GerarNFClientes;                            Override;
    procedure GerarNFitensClientes;                       Override;
    procedure GerarVenda;                                 Override;

    procedure GerarFinanceiro( Tipo, Situacao :Integer ); Override;
    procedure GerarFinanceiroReceber(Aberto:String);      Override;
    procedure GerarFinanceiroReceberCartao;               Override;
    procedure GerarFinanceiroPagar(Aberto:String);        Override;



    procedure GerarEstDeposito;
    procedure GerarCustoCheio;
    procedure AtualizaEstoque;
    procedure EstoqueTroca;
    procedure PercentualFidelidade;
    procedure FornecedorPreferencial;

  end;

var
  FrmHiperLojao: TFrmHiperLojao;
  ListNCM    : TStringList;
  TotalCont  : Integer;
  NumLinha : Integer;
  Arquivo: TextFile;
  FlgGeraDados : Boolean = false;
  FlgGeraCest : Boolean = false;
  FlgGeraAmarrarCest : Boolean = false;
  FlgGeraEstoqueDeposito : Boolean = False;
  FlgGeraCustoCheio : Boolean = False;
  FlgAtualizaEstoque : Boolean = False;
  FlgEstoqueTroca : Boolean = False;
  FlgPercentualFidelidade : Boolean = False;
  FlgFornecedorPreferencial : Boolean = False;


implementation

{$R *.dfm}

uses xProc, UUtilidades, UProgresso;


procedure TFrmHiperLojao.GerarProducao;
begin
  inherited;
  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
    SQL.Add('    PRODUTOS_COMPOSICAO.PRODUTO_BASE AS COD_PRODUTO,');
    SQL.Add('    PRODUTOS_COMPOSICAO.PRODUTO AS COD_PRODUTO_PRODUCAO,');
    SQL.Add('    COMPOSICAO.FATOR_PRODUCAO AS QTD_PRODUCAO,');
    SQL.Add('    PRODUTOS.UNIDADE_VENDA AS DES_UNIDADE,');
    SQL.Add('    PRODUTOS_COMPOSICAO.QTDE AS QTD_RECEITA,');
    SQL.Add('    COMPOSICAO.RENDIMENTO AS QTD_RENDIMENTO');
    SQL.Add('FROM');
    SQL.Add('    PRODUTOS');
    SQL.Add('LEFT JOIN');
    SQL.Add('    PRODUTOS_COMPOSICAO');
    SQL.Add('ON');
    SQL.Add('    PRODUTOS.ID = PRODUTOS_COMPOSICAO.PRODUTO_BASE ');
    SQL.Add('LEFT JOIN');
    SQL.Add('     COMPOSICAO');
    SQL.Add('ON');
    SQL.Add('     PRODUTOS_COMPOSICAO.PRODUTO_BASE = COMPOSICAO.PRODUTO_BASE');
    SQL.Add('WHERE');
    SQL.Add('    PRODUTOS.COMPOSTO = 2');
    SQL.Add('AND');
    SQL.Add('    PRODUTOS_COMPOSICAO.PRODUTO_BASE IS NOT NULL');


    Open;

    First;
    NumLinha := 0;

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);

      Layout.SetValues(QryPrincipal, NumLinha, RecordCount);

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;
end;

procedure TFrmHiperLojao.GerarProduto;
var
 cod_produto, codbarras : string;
 TotalCount, count, NovoCodigo : Integer;
 QrySalvaCodProd : TSQLQuery;

begin
  inherited;

  QrySalvaCodProd := TSQLQuery.Create(FrmProgresso);
  with QrySalvaCodProd do
  begin
    SQLConnection := ScnBanco;

      SQL.Clear;
      SQL.Add('ALTER TABLE PRODUTOS ');
      SQL.Add('ADD COD_PRODUTO VARCHAR(13) DEFAULT NULL ;');

      try
        ExecSQL;
      except
      end;

      SQL.Clear;
      SQL.Add('UPDATE PRODUTOS');
      SQL.Add('SET COD_PRODUTO = :COD_PRODUTO');
      SQL.Add('WHERE PRO_BARRA = :COD_BARRA_PRINCIPAL');
      SQL.Add('AND CHAR_LENGTH(PRO_REFERENCIA) = 9');


      try
        ExecSQL;
      except
      end;

  end;


  if FlgGeraEstoqueDeposito then
  begin
    GerarEstDeposito;
    Exit;
  end;

  if FlgGeraCustoCheio then
  begin
    GerarCustoCheio;
    Exit;
  end;

  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

     SQL.Add('   SELECT DISTINCT      ');
     SQL.Add('      ');
     SQL.Add('       CASE   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''25578,'' THEN CAST(25578 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''28687.'' THEN CAST(28687 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''29397.'' THEN CAST(29397 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''29*504'' THEN CAST(29504 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''296,5'' THEN CAST(2965 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''29878.'' THEN CAST(29878 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''29870,'' THEN CAST(29870 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''2995,'' THEN CAST(2995 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''30023Q'' THEN CAST(30023 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''30045,'' THEN CAST(30045 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''30304,'' THEN CAST(30304 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN CHAR_LENGTH(PRODUTOS.PRO_REFERENCIA) = 9 THEN PRODUTOS.COD_PRODUTO   ');
     SQL.Add('           ELSE PRODUTOS.PRO_REFERENCIA        ');
     SQL.Add('       END AS COD_PRODUTO,    ');
     SQL.Add('          ');
     SQL.Add('       CASE   ');
     SQL.Add('           WHEN LINHAS.LIN_CODIGO = 1 AND GRUPOS.GRU_CODIGO = 23 AND SUB_GRUPOS.SGR_CODIGO = 9999 THEN PRODUTOS.PRO_REFERENCIA   ');
     SQL.Add('           WHEN LINHAS.LIN_CODIGO = 1 AND GRUPOS.GRU_CODIGO = 23 AND SUB_GRUPOS.SGR_CODIGO = 9996 THEN PRODUTOS.PRO_REFERENCIA   ');
     SQL.Add('           ELSE PRODUTOS.PRO_BARRA    ');
     SQL.Add('       END AS COD_BARRA_PRINCIPAL,    ');
     SQL.Add('            ');
     SQL.Add('       TRIM(COALESCE(PRODUTOS.PRO_DESCRICAO, ''A DEFINIR'')) AS DES_REDUZIDA,      ');
     SQL.Add('       TRIM(COALESCE(PRODUTOS.PRO_DESCRICAO, ''A DEFINIR'')) AS DES_PRODUTO,      ');
     SQL.Add('       CASE   ');
     SQL.Add('          WHEN PRO_UNIDCOMPRA = ''UN'' THEN COALESCE(PRODUTOS.PRO_FATORCOMPRA, 1)   ');
     SQL.Add('          ELSE COALESCE(PRODUTOS.PRO_FTUCUV, 1)   ');
     SQL.Add('       END AS QTD_EMBALAGEM_COMPRA,   ');
     SQL.Add('          ');
     SQL.Add('       CASE   ');
     SQL.Add('           WHEN LINHAS.LIN_CODIGO = 1 AND GRUPOS.GRU_CODIGO = 23 AND SUB_GRUPOS.SGR_CODIGO = 9999 THEN ''KG''   ');
     SQL.Add('           WHEN LINHAS.LIN_CODIGO = 1 AND GRUPOS.GRU_CODIGO = 23 AND SUB_GRUPOS.SGR_CODIGO = 9996 THEN ''KG''   ');
     SQL.Add('           ELSE SUBSTRING(COALESCE(PRODUTOS.PRO_UNIDCOMPRA, ''UN'') FROM 1 FOR 2)    ');
     SQL.Add('       END AS DES_UNIDADE_COMPRA,      ');
     SQL.Add('          ');
     SQL.Add('       1 AS QTD_EMBALAGEM_VENDA,      ');
     SQL.Add('          ');
     SQL.Add('       CASE   ');
     SQL.Add('           WHEN LINHAS.LIN_CODIGO = 1 AND GRUPOS.GRU_CODIGO = 23 AND SUB_GRUPOS.SGR_CODIGO = 9999 THEN ''KG''   ');
     SQL.Add('           WHEN LINHAS.LIN_CODIGO = 1 AND GRUPOS.GRU_CODIGO = 23 AND SUB_GRUPOS.SGR_CODIGO = 9996 THEN ''KG''   ');
     SQL.Add('           ELSE ''UN''    ');
     SQL.Add('       END AS DES_UNIDADE_VENDA,      ');
     SQL.Add('          ');
     SQL.Add('       0 AS TIPO_IPI,      ');
     SQL.Add('       0 AS VAL_IPI,      ');
     SQL.Add('       COALESCE(LINHAS.LIN_CODIGO, 999) AS COD_SECAO,      ');
     SQL.Add('       COALESCE(GRUPOS.GRU_CODIGO, 999) AS COD_GRUPO,      ');
     SQL.Add('       COALESCE(SUB_GRUPOS.SGR_CODIGO, 999) AS COD_SUB_GRUPO,      ');
     SQL.Add('       0 AS COD_PRODUTO_SIMILAR,      ');
     SQL.Add('          ');
     SQL.Add('       CASE   ');
     SQL.Add('           WHEN LINHAS.LIN_CODIGO = 1 AND GRUPOS.GRU_CODIGO = 23 AND SUB_GRUPOS.SGR_CODIGO = 9999 THEN ''S''   ');
     SQL.Add('           WHEN LINHAS.LIN_CODIGO = 1 AND GRUPOS.GRU_CODIGO = 23 AND SUB_GRUPOS.SGR_CODIGO = 9996 THEN ''S''   ');
     SQL.Add('           ELSE ''N''   ');
     SQL.Add('       END AS IPV,      ');
     SQL.Add('          ');
     SQL.Add('          ');
     SQL.Add('       COALESCE(PRODUTOS.PRO_NUM_VALIDADE, 0) AS DIAS_VALIDADE,      ');
     SQL.Add('       0 AS TIPO_PRODUTO,      ');
     SQL.Add('       ''N'' AS FLG_NAO_PIS_COFINS,      ');
     SQL.Add('          ');
     SQL.Add('       CASE   ');
     SQL.Add('           WHEN LINHAS.LIN_CODIGO = 1 AND GRUPOS.GRU_CODIGO = 23 AND SUB_GRUPOS.SGR_CODIGO = 9999 THEN ''S''   ');
     SQL.Add('           WHEN LINHAS.LIN_CODIGO = 1 AND GRUPOS.GRU_CODIGO = 23 AND SUB_GRUPOS.SGR_CODIGO = 9996 THEN ''S''   ');
     SQL.Add('           ELSE ''N''   ');
     SQL.Add('       END AS FLG_ENVIA_BALANCA,      ');
     SQL.Add('          ');
     SQL.Add('       1 AS TIPO_NAO_PIS_COFINS,      ');
     SQL.Add('       0 AS TIPO_EVENTO,      ');
     SQL.Add('       0 AS COD_ASSOCIADO,      ');
     SQL.Add('       '''' AS DES_OBSERVACAO,      ');
     SQL.Add('       0 AS COD_INFO_NUTRICIONAL,      ');
     //SQL.Add('       0 AS COD_INFO_RECEITA,      ');
     SQL.Add('       999 AS COD_TAB_SPED,      ');
     SQL.Add('       ''N'' AS FLG_ALCOOLICO,      ');
     SQL.Add('       0 AS TIPO_ESPECIE,      ');
     SQL.Add('       0 AS COD_CLASSIF,      ');
     SQL.Add('       1 AS VAL_VDA_PESO_BRUTO,      ');
     SQL.Add('       1 AS VAL_PESO_EMB,      ');
     SQL.Add('       0 AS TIPO_EXPLOSAO_COMPRA,      ');
     SQL.Add('       '''' AS DTA_INI_OPER,      ');
     SQL.Add('       '''' AS DES_PLAQUETA,      ');
     SQL.Add('       '''' AS MES_ANO_INI_DEPREC,      ');
     SQL.Add('       0 AS TIPO_BEM,      ');
     SQL.Add('       COALESCE(PRODUTOS.PRO_FORNECEDOR, 0) AS COD_FORNECEDOR,      ');
     SQL.Add('       0 AS NUM_NF,      ');
     SQL.Add('       '''' AS DTA_ENTRADA,      ');
     SQL.Add('       0 AS COD_NAT_BEM,      ');
     SQL.Add('       0 AS VAL_ORIG_BEM,      ');
     SQL.Add('       TRIM(COALESCE(PRODUTOS.PRO_DESCRICAO, ''A DEFINIR'')) AS DES_PRODUTO_ANT ');
     SQL.Add('   FROM      ');
     SQL.Add('       PRODUTOS      ');
     SQL.Add('   LEFT JOIN LINHAS ON PRODUTOS.PRO_LINHA = LINHAS.LIN_CODIGO      ');
     SQL.Add('   LEFT JOIN GRUPOS ON PRODUTOS.PRO_GRUPO = GRUPOS.GRU_CODIGO      ');
     SQL.Add('   LEFT JOIN SUB_GRUPOS ON PRODUTOS.PRO_SUBGRUPO = SUB_GRUPOS.SGR_CODIGO      ');
     SQL.Add('   LEFT JOIN PRODUTOS_EMPRESA ON PRODUTOS.PRO_CODIGO = PRODUTOS_EMPRESA.PRE_PRODUTO   ');
     SQL.Add('   WHERE PRODUTOS_EMPRESA.PRE_EMPRESA = '+CbxLoja.Text+'    ');
     SQL.Add('   ORDER BY(      ');
     SQL.Add('       CASE WHEN PRODUTOS.PRO_LEGENDA IN (''N'',''P'') THEN PRODUTOS.PRO_LEGENDA END      ');
     SQL.Add('   ) DESC,      ');
     SQL.Add('   (      ');
     SQL.Add('       CASE WHEN PRODUTOS.PRO_LEGENDA NOT IN (''N'', ''P'') THEN PRODUTOS.PRO_LEGENDA END      ');
     SQL.Add('   ), PRODUTOS.PRO_CODIGO ASC   ');

     //SQL.Add('   LEFT JOIN PRODUTOS_CLAS_TRIB_FEDERAL AS TRIBUTACAO ON PRODUTOS.PRO_CODIGO = TRIBUTACAO.PTF_PRODUTO   ');
     //SQL.Add('   LEFT JOIN PRODUTOS_CLAS_TRIBUTACAO ON PRODUTOS.PRO_CODIGO = PRODUTOS_CLAS_TRIBUTACAO.PST_PRODUTO   ');
     //SQL.Add('   LEFT JOIN NAT_RECEITA AS SPED ON PRODUTOS_CLAS_TRIBUTACAO.PST_TPPED = SPED.NRC_SEQUENCIA   ');



    Open;
    First;
    NumLinha := 0;
    count := 100000;
    NovoCodigo := 400000;

    TotalCount := SetCountTotal(SQL.Text);





    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);
      Layout.SetValues(QryPrincipal, NumLinha, TotalCount);





      Layout.FieldByName('DES_OBSERVACAO').AsString := StrReplace(StrLBReplace(FieldByName('DES_OBSERVACAO').AsString), '\n', '');
      Layout.FieldByName('DES_REDUZIDA').AsString := StrReplace(StrLBReplace(FieldByName('DES_REDUZIDA').AsString), '\n', '');
      Layout.FieldByName('DES_PRODUTO').AsString := StrReplace(StrLBReplace(FieldByName('DES_PRODUTO').AsString), '\n', '');
      Layout.FieldByName('DES_PRODUTO_ANT').AsString := StrReplace(StrLBReplace(FieldByName('DES_PRODUTO').AsString), '\n', '');

      if Layout.FieldByName('COD_PRODUTO').AsInteger = 0 then
      begin
        with QrySalvaCodProd do
        begin
          Inc(NovoCodigo);
          ParamByName('COD_PRODUTO').Value := NovoCodigo;
          ParamByName('COD_BARRA_PRINCIPAL').Value := Layout.FieldByName('COD_BARRA_PRINCIPAL').AsString;
          Layout.FieldByName('COD_PRODUTO').AsInteger := ParamByName('COD_PRODUTO').Value;
          ExecSQL();
        end;
      end;


      Layout.FieldByName('COD_PRODUTO').AsString := GerarPLU( Layout.FieldByName('COD_PRODUTO').AsString );

//      codbarras := StrRetNums(Layout.FieldByName('COD_BARRA_PRINCIPAL').AsString);
//
//      if( (codbarras = '') or (StrToFloat(codbarras) = 0)) then
//         codbarras := ''
//      else if ( Length(TiraZerosEsquerda(codbarras)) < 8 ) then
//         codbarras := GerarPLU(codbarras)
//      else
//         if(not CodBarrasValido(codbarras)) then
//            codbarras := '';
//
//      Layout.FieldByName('COD_BARRA_PRINCIPAL').AsString := codbarras;

//      if Layout.FieldByName('COD_PRODUTO').AsInteger = 0 then
//      begin
//        Inc(NovoCodigo);
//        Layout.FieldByName('COD_PRODUTO').AsString := GerarPLU(IntToStr(NovoCodigo));
//      end;



      if Length(StrLBReplace(Trim(StrRetNums( FieldByName('COD_BARRA_PRINCIPAL').AsString) ))) < 8 then
       Layout.FieldByName('COD_BARRA_PRINCIPAL').AsString := GerarPLU(FieldByName('COD_BARRA_PRINCIPAL').AsString);

      if not CodBarrasValido(Layout.FieldByName('COD_BARRA_PRINCIPAL').AsString) then
       Layout.FieldByName('COD_BARRA_PRINCIPAL').AsString := '';

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal.RecNo, 'E', E.Message);
    end;
    Next;
    end;

    Close
  end;
end;





procedure TFrmHiperLojao.GerarSecao;
var
   TotalCount : integer;
begin
  inherited;

  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

     SQL.Add('   SELECT DISTINCT   ');
     SQL.Add('       COALESCE(LINHAS.LIN_CODIGO, 999) AS COD_SECAO,   ');
     SQL.Add('       COALESCE(LINHAS.LIN_DESCRICAO, ''A DEFINIR'') AS DES_SECAO,   ');
     SQL.Add('       0 AS VAL_META   ');
     SQL.Add('   FROM   ');
     SQL.Add('       PRODUTOS   ');
     SQL.Add('   LEFT JOIN LINHAS ON PRODUTOS.PRO_LINHA = LINHAS.LIN_CODIGO   ');
     SQL.Add('   LEFT JOIN PRODUTOS_EMPRESA ON PRODUTOS.PRO_CODIGO = PRODUTOS_EMPRESA.PRE_PRODUTO');
     SQL.Add('   WHERE PRODUTOS_EMPRESA.PRE_EMPRESA = '+CbxLoja.Text+' ');



    Open;

    First;
    NumLinha := 0;
    TotalCount := SetCountTotal(SQL.Text);

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);

      Layout.SetValues(QryPrincipal, NumLinha, TotalCount);

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;

end;

procedure TFrmHiperLojao.GerarSubGrupo;
var
   TotalCount : Integer;
begin
  inherited;

  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

     SQL.Add('   SELECT DISTINCT   ');
     SQL.Add('       COALESCE(LINHAS.LIN_CODIGO, 999) AS COD_SECAO,   ');
     SQL.Add('       COALESCE(GRUPOS.GRU_CODIGO, 999) AS COD_GRUPO,   ');
     SQL.Add('       COALESCE(SUB_GRUPOS.SGR_CODIGO, 999) AS COD_SUB_GRUPO,   ');
     SQL.Add('       COALESCE(SUB_GRUPOS.SGR_NOME, ''A DEFINIR'') AS DES_SUB_GRUPO,   ');
     SQL.Add('       0 AS VAL_META,   ');
     SQL.Add('       0 AS VAL_MARGEM_REF,   ');
     SQL.Add('       0 AS QTD_DIA_SEGURANCA,   ');
     SQL.Add('       ''N'' AS FLG_ALCOOLICO   ');
     SQL.Add('   FROM   ');
     SQL.Add('       PRODUTOS   ');
     SQL.Add('   LEFT JOIN LINHAS ON PRODUTOS.PRO_LINHA = LINHAS.LIN_CODIGO   ');
     SQL.Add('   LEFT JOIN GRUPOS ON PRODUTOS.PRO_GRUPO = GRUPOS.GRU_CODIGO   ');
     SQL.Add('   LEFT JOIN SUB_GRUPOS ON PRODUTOS.PRO_SUBGRUPO = SUB_GRUPOS.SGR_CODIGO   ');
     SQL.Add('   LEFT JOIN PRODUTOS_EMPRESA ON PRODUTOS.PRO_CODIGO = PRODUTOS_EMPRESA.PRE_PRODUTO');
     SQL.Add('   WHERE PRODUTOS_EMPRESA.PRE_EMPRESA = '+CbxLoja.Text+' ');



    Open;

    First;
    NumLinha := 0;

    TotalCount := SetCountTotal(SQL.Text);

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);

      Layout.SetValues(QryPrincipal, NumLinha, TotalCount);

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;

end;

procedure TFrmHiperLojao.GerarTransportadora;
var
  TotalCount, count : Integer;
begin
  inherited;
  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

     SQL.Add('   SELECT   ');
     SQL.Add('       TRANSPORTADORA.FOR_CODIGO AS COD_TRANSPORTADORA,   ');
     SQL.Add('       TRANSPORTADORA.FOR_RAZAO AS DES_TRANSPORTADORA,   ');
     SQL.Add('       REPLACE(REPLACE(REPLACE(TRANSPORTADORA.FOR_CGC, ''.'', ''''), ''-'', ''''), ''/'', '''') AS NUM_CGC,   ');
     SQL.Add('          ');
     SQL.Add('       CASE TRANSPORTADORA.FOR_IE      ');
     SQL.Add('           WHEN ''0000'' THEN ''ISENTO''      ');
     SQL.Add('           WHEN ''00000'' THEN ''ISENTO''      ');
     SQL.Add('           WHEN ''000000'' THEN ''ISENTO''      ');
     SQL.Add('           WHEN ''0000000'' THEN ''ISENTO''      ');
     SQL.Add('           WHEN ''isento'' THEN ''ISENTO''      ');
     SQL.Add('           WHEN '''' THEN ''ISENTO''      ');
     SQL.Add('           WHEN ''ISENTO'' THEN ''ISENTO''      ');
     SQL.Add('           ELSE REPLACE(REPLACE(REPLACE(TRANSPORTADORA.FOR_IE, ''.'', ''''), ''-'', ''''), ''/'', '''')     ');
     SQL.Add('       END AS NUM_INSC_EST,     ');
     SQL.Add('          ');
     SQL.Add('       COALESCE(TRANSPORTADORA.FOR_ENDERECO, ''A DEFINIR'') AS DES_ENDERECO,   ');
     SQL.Add('       COALESCE(TRANSPORTADORA.FOR_BAIRRO, ''A DEFINIR'') AS DES_BAIRRO,   ');
     SQL.Add('       TRANSPORTADORA.FOR_CIDADE AS DES_CIDADE,   ');
     SQL.Add('       COALESCE(TRANSPORTADORA.FOR_UF, ''SP'') AS DES_SIGLA,   ');
     SQL.Add('       COALESCE(TRANSPORTADORA.FOR_CEP, ''13560642'') AS NUM_CEP,   ');
     SQL.Add('       COALESCE(FORCONTATOS.FCT_FONE, '''') AS NUM_FONE,   ');
     SQL.Add('       '''' AS NUM_FAX,   ');
     SQL.Add('       COALESCE(FORCONTATOS.FCT_CONTATO, '''') AS DES_CONTATO,   ');
     SQL.Add('       2 AS COD_CONDICAO,   ');
     SQL.Add('       30 AS NUM_CONDICAO,   ');
     SQL.Add('       COALESCE(TRANSPORTADORA.FOR_NUMERO, ''S/N'') AS NUM_ENDERECO,   ');
     SQL.Add('       COALESCE(CAST(TRANSPORTADORA.FOR_OBS AS VARCHAR(300)), '''') AS DES_OBSERVACAO,   ');
     SQL.Add('       8 AS COD_ENTIDADE,   ');
     SQL.Add('       COALESCE(TRANSPORTADORA.FOR_EMAIL, '''') AS DES_EMAIL,   ');
     SQL.Add('       COALESCE(TRANSPORTADORA.FOR_HOMEPAGE, '''') AS DES_WEB_SITE   ');
     SQL.Add('   FROM   ');
     SQL.Add('       FORNECEDORES AS TRANSPORTADORA   ');
     SQL.Add('   LEFT JOIN FORCONTATOS ON TRANSPORTADORA.FOR_CODIGO = FORCONTATOS.FCT_FORNECEDOR   ');
     SQL.Add('   WHERE TRANSPORTADORA.FOR_TRANSPORTADORA = ''S''   ');
     SQL.Add('   ORDER BY TRANSPORTADORA.FOR_CODIGO DESC   ');



    Open;

    First;

    NumLinha := 0;

    count := 100000;

    TotalCount := SetCountTotal(SQL.Text);

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);

      Layout.SetValues(QryPrincipal, NumLinha, TotalCount);

      if Length(Layout.FieldByName('NUM_CGC').AsString) > 11 then
      begin
        if not ValidaCGC(Layout.FieldByName('NUM_CGC').AsString) then
          Layout.FieldByName('NUM_CGC').AsString := '';
      end
      else
        if not ValidaCPF(Layout.FieldByName('NUM_CGC').AsString) then
          Layout.FieldByName('NUM_CGC').AsString := '';

      Layout.FieldByName('DES_OBSERVACAO').AsString := StrReplace(StrLBReplace(FieldByName('DES_OBSERVACAO').AsString), '\n', '');
      Layout.FieldByName('DES_EMAIL').AsString := StrReplace(StrLBReplace(FieldByName('DES_EMAIL').AsString), '\n', '');

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;
end;

procedure TFrmHiperLojao.GerarVenda;
var
   TotalCount : Integer;
begin
  inherited;

  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

     SQL.Add('   SELECT DISTINCT   ');
     SQL.Add('      ');
     SQL.Add('       CASE      ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''25578,'' THEN CAST(25578 + 50000 AS INTEGER)      ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''28687.'' THEN CAST(28687 + 50000 AS INTEGER)      ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''29397.'' THEN CAST(29397 + 50000 AS INTEGER)      ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''29*504'' THEN CAST(29504 + 50000 AS INTEGER)      ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''296,5'' THEN CAST(2965 + 50000 AS INTEGER)      ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''29878.'' THEN CAST(29878 + 50000 AS INTEGER)      ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''29870,'' THEN CAST(29870 + 50000 AS INTEGER)      ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''2995,'' THEN CAST(2995 + 50000 AS INTEGER)      ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''30023Q'' THEN CAST(30023 + 50000 AS INTEGER)      ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''30045,'' THEN CAST(30045 + 50000 AS INTEGER)      ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''30304,'' THEN CAST(30304 + 50000 AS INTEGER)      ');
     SQL.Add('           WHEN CHAR_LENGTH(PRODUTOS.PRO_REFERENCIA) = 9 THEN PRODUTOS.COD_PRODUTO      ');
     SQL.Add('           ELSE PRODUTOS.PRO_REFERENCIA      ');
     SQL.Add('       END  AS COD_PRODUTO,   ');
     SQL.Add('      ');
     SQL.Add('       '+CbxLoja.Text+' AS COD_LOJA,   ');
     SQL.Add('       0 AS IND_TIPO,   ');
     SQL.Add('       1 AS NUM_PDV,   ');
     SQL.Add('       VENDAS.CB_PEDIDOS AS QTD_TOTAL_PRODUTO,   ');
     SQL.Add('       VENDAS.CB_VALOR_LIQUIDO AS VAL_TOTAL_PRODUTO,   ');
     SQL.Add('       COALESCE(PRECOS.PRE_PRECOUN, 1) AS VAL_PRECO_VENDA,   ');
     SQL.Add('       VENDAS.CB_VALOR_CUSTO AS VAL_CUSTO_REP,   ');
     SQL.Add('       ''13/07/2020'' AS DTA_SAIDA,   ');
     SQL.Add('       ''132020'' AS DTA_MENSAL,   ');
     SQL.Add('       1 AS NUM_IDENT,   ');
     SQL.Add('      ');
     SQL.Add('       CASE   ');
     SQL.Add('           WHEN LINHAS.LIN_CODIGO = 1 AND GRUPOS.GRU_CODIGO = 23 AND SUB_GRUPOS.SGR_CODIGO = 9999 THEN PRODUTOS.PRO_REFERENCIA   ');
     SQL.Add('           WHEN LINHAS.LIN_CODIGO = 1 AND GRUPOS.GRU_CODIGO = 23 AND SUB_GRUPOS.SGR_CODIGO = 9996 THEN PRODUTOS.PRO_REFERENCIA   ');
     SQL.Add('           ELSE PRODUTOS.PRO_BARRA    ');
     SQL.Add('       END AS COD_EAN,   ');
     SQL.Add('      ');
     SQL.Add('       ''0000'' AS DES_HORA,   ');
     SQL.Add('       CLIENTES.CLI_CODIGO AS COD_CLIENTE,   ');
     SQL.Add('       1 AS COD_ENTIDADE,   ');
     SQL.Add('       0 AS VAL_BASE_ICMS,   ');
     SQL.Add('       '''' AS DES_SITUACAO_TRIB,   ');
     SQL.Add('       0 AS VAL_ICMS,   ');
     SQL.Add('       VENDAS.CB_DESCRICAO AS NUM_CUPOM_FISCAL,   ');
     SQL.Add('       VENDAS.CB_VALOR AS VAL_VENDA_PDV,   ');
     SQL.Add('       1 AS COD_TRIBUTACAO,   ');
     SQL.Add('       ''N'' AS FLG_CUPOM_CANCELADO,   ');
     SQL.Add('       00000001 AS NUM_NCM,   ');
     SQL.Add('       999 AS COD_TAB_SPED,   ');
     SQL.Add('       ''N'' AS FLG_NAO_PIS_COFINS,   ');
     SQL.Add('       1 AS TIPO_NAO_PIS_COFINS,   ');
     SQL.Add('       ''N'' AS FLG_ONLINE,   ');
     SQL.Add('       ''N'' AS FLG_OFERTA,   ');
     SQL.Add('       0 AS COD_ASSOCIADO   ');
     SQL.Add('   FROM   ');
     SQL.Add('       CUBO_RESUMO AS VENDAS   ');
     SQL.Add('   LEFT JOIN PRODUTOS ON PRODUTOS.PRO_CODIGO = VENDAS.CB_ID   ');
     SQL.Add('   LEFT JOIN LINHAS ON PRODUTOS.PRO_LINHA = LINHAS.LIN_CODIGO      ');
     SQL.Add('   LEFT JOIN GRUPOS ON PRODUTOS.PRO_GRUPO = GRUPOS.GRU_CODIGO      ');
     SQL.Add('   LEFT JOIN SUB_GRUPOS ON PRODUTOS.PRO_SUBGRUPO = SUB_GRUPOS.SGR_CODIGO   ');
     SQL.Add('   LEFT JOIN CUBO_RESUMO_CLIENTES ON CUBO_RESUMO_CLIENTES.CBC_SECAO = VENDAS.CB_SECAO   ');
     SQL.Add('   LEFT JOIN CLIENTES ON CLIENTES.CLI_CODIGO = CUBO_RESUMO_CLIENTES.CBC_CLIENTE   ');
     SQL.Add('   LEFT JOIN PRECOS ON PRODUTOS.PRO_CODIGO = VENDAS.CB_ID   ');
     SQL.Add('   WHERE VENDAS.CB_AGRUPADOR = ''PR''   ');

   // SQL.Add('WHERE VENDAS.DATA >= :INI');
   // SQL.Add('AND VENDAS.DATA <= :FIM');
   // SQL.Add('AND CAIXA.SITUACAO = ''A''');
   // SQL.Add('AND VENDAS.BAIXADO = ''N''');


   // ParamByName('INI').AsDate := DtpInicial.Date;
   // ParamByName('FIM').AsDate := DtpFinal.Date;


    Open;

    First;

    TotalCount := SetCountTotal(SQL.Text, ParamByName('INI').AsString, ParamByName('FIM').AsString );

    NumLinha := 0;


    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);

      Layout.SetValues(QryPrincipal, NumLinha, TotalCount);
      Layout.FieldByName('COD_PRODUTO').AsString := GerarPLU(QryPrincipal.FieldByName('COD_PRODUTO').AsString);
      Layout.FieldByName('DTA_SAIDA').AsDateTime := QryPrincipal.FieldByName('DTA_SAIDA').AsDateTime;

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;
end;



procedure TFrmHiperLojao.PercentualFidelidade;
var
  COD_PRODUTO : string;
begin
  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

     SQL.Add('   SELECT DISTINCT   ');
     SQL.Add('       CASE   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''25578,'' THEN CAST(25578 + 50000 AS INTEGER)      ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''28687.'' THEN CAST(28687 + 50000 AS INTEGER)      ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''29397.'' THEN CAST(29397 + 50000 AS INTEGER)      ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''29*504'' THEN CAST(29504 + 50000 AS INTEGER)      ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''296,5'' THEN CAST(2965 + 50000 AS INTEGER)      ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''29878.'' THEN CAST(29878 + 50000 AS INTEGER)      ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''29870,'' THEN CAST(29870 + 50000 AS INTEGER)      ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''2995,'' THEN CAST(2995 + 50000 AS INTEGER)      ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''30023Q'' THEN CAST(30023 + 50000 AS INTEGER)      ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''30045,'' THEN CAST(30045 + 50000 AS INTEGER)      ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''30304,'' THEN CAST(30304 + 50000 AS INTEGER)      ');
     SQL.Add('           WHEN CHAR_LENGTH(PRODUTOS.PRO_REFERENCIA) = 9 THEN PRODUTOS.COD_PRODUTO      ');
     SQL.Add('           ELSE PRODUTOS.PRO_REFERENCIA           ');
     SQL.Add('       END AS COD_PRODUTO,   ');
     SQL.Add('      ');
     SQL.Add('       CAST(PRF_PERC01 AS INTEGER) AS PER_FIDELIDADE   ');
     SQL.Add('   FROM   ');
     SQL.Add('       PRODUTOS_FORNECEDORES   ');
     SQL.Add('   LEFT JOIN PRODUTOS ON PRODUTOS.PRO_CODIGO = PRODUTOS_FORNECEDORES.PRF_CODIGO   ');
     SQL.Add('   WHERE   ');
     SQL.Add('       PRF_PERC01 IS NOT NULL   ');
     SQL.Add('   AND   ');
     SQL.Add('       PRF_EMPRESA = '+CbxLoja.Text+'  ');
     SQL.Add('   AND   ');
     SQL.Add('       PRF_PADRAO = ''S''   ');
     SQL.Add('   AND   ');
     SQL.Add('       PRF_PERC01 <> 0   ');
     SQL.Add('   ORDER BY (   ');
     SQL.Add('       CASE WHEN PRODUTOS.PRO_LEGENDA IN (''N'',''P'') THEN PRODUTOS.PRO_LEGENDA END   ');
     SQL.Add('   ) DESC,   ');
     SQL.Add('   (   ');
     SQL.Add('       CASE WHEN PRODUTOS.PRO_LEGENDA NOT IN (''N'', ''P'') THEN PRODUTOS.PRO_LEGENDA END   ');
     SQL.Add('   ), PRODUTOS.PRO_CODIGO ASC   ');
     SQL.Add('      ');

     Open;
     First;

     NumLinha := 0;

     while not Eof do
     begin
       try
        if Cancelar then
        Break;

        Inc(NumLinha);

        COD_PRODUTO := GerarPLU(QryPrincipal.FieldByName('COD_PRODUTO').AsString);

        Writeln(Arquivo, 'UPDATE TAB_PRODUTO_LOJA SET PER_FIDELIDADE = '+QryPrincipal.FieldByName('PER_FIDELIDADE').AsString+' WHERE COD_PRODUTO = '+COD_PRODUTO+'; ');

        if NumLinha = 500 then
          Writeln(Arquivo, 'COMMIT WORK;');
        if NumLinha = 1000 then
            Writeln(Arquivo, 'COMMIT WORK;');
        if NumLinha = 1500 then
            Writeln(Arquivo, 'COMMIT WORK;');
        if NumLinha = 2000 then
            Writeln(Arquivo, 'COMMIT WORK;');
        if NumLinha = 2500 then
            Writeln(Arquivo, 'COMMIT WORK;');
        if NumLinha = 3000 then
            Writeln(Arquivo, 'COMMIT WORK;');
        if NumLinha = 3500 then
            Writeln(Arquivo, 'COMMIT WORK;');
        if NumLinha = 4000 then
            Writeln(Arquivo, 'COMMIT WORK;');
        if NumLinha = 4500 then
            Writeln(Arquivo, 'COMMIT WORK;');
        if NumLinha = 5000 then
            Writeln(Arquivo, 'COMMIT WORK;');
        if NumLinha = 5500 then
            Writeln(Arquivo, 'COMMIT WORK;');
        if NumLinha = 6000 then
            Writeln(Arquivo, 'COMMIT WORK;');
        if NumLinha = 6500 then
            Writeln(Arquivo, 'COMMIT WORK;');
        if NumLinha = 7000 then
            Writeln(Arquivo, 'COMMIT WORK;');
        if NumLinha = 7500 then
            Writeln(Arquivo, 'COMMIT WORK;');
        if NumLinha = 8000 then
            Writeln(Arquivo, 'COMMIT WORK;');
        if NumLinha = 8500 then
            Writeln(Arquivo, 'COMMIT WORK;');
        if NumLinha = 9000 then
            Writeln(Arquivo, 'COMMIT WORK;');
        if NumLinha = 9500 then
            Writeln(Arquivo, 'COMMIT WORK;');
        if NumLinha = 10000 then
            Writeln(Arquivo, 'COMMIT WORK;');
        if NumLinha = 10500 then
            Writeln(Arquivo, 'COMMIT WORK;');
        if NumLinha = 20500 then
            Writeln(Arquivo, 'COMMIT WORK;');
        if NumLinha = 30500 then
            Writeln(Arquivo, 'COMMIT WORK;');
        if NumLinha = 40500 then
            Writeln(Arquivo, 'COMMIT WORK;');
        if NumLinha = 50500 then
            Writeln(Arquivo, 'COMMIT WORK;');
        if NumLinha = 60500 then
            Writeln(Arquivo, 'COMMIT WORK;');
        if NumLinha = 70000 then
            Writeln(Arquivo, 'COMMIT WORK;');
        if NumLinha = 70500 then
            Writeln(Arquivo, 'COMMIT WORK;');
        if NumLinha = 80000 then
            Writeln(Arquivo, 'COMMIT WORK;');


       except on E: Exception do
         FrmProgresso.AdicionarLog(QryPrincipal.RecNo, 'E', E.Message);
       end;
        Next;
     end;
      Writeln(Arquivo, 'COMMIT WORK;');
      Close;
  end;
end;

procedure TFrmHiperLojao.AtualizaEstoque;
var
  NovoCodigo : Integer;
  COD_PRODUTO : string;
begin
  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

     SQL.Add('   SELECT DISTINCT   ');
     SQL.Add('      ');
     SQL.Add('       CASE PRODUTOS_ESTOQUES.PST_LOCAL   ');
     SQL.Add('           WHEN 2 THEN 1      ');
     SQL.Add('           WHEN 4 THEN 6      ');
     SQL.Add('           WHEN 5 THEN 3      ');
     SQL.Add('       END AS COD_LOJA,   ');
     SQL.Add('      ');
     SQL.Add('       CASE          ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''25578,'' THEN CAST(25578 + 50000 AS INTEGER)      ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''28687.'' THEN CAST(28687 + 50000 AS INTEGER)      ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''29397.'' THEN CAST(29397 + 50000 AS INTEGER)      ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''29*504'' THEN CAST(29504 + 50000 AS INTEGER)      ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''296,5'' THEN CAST(2965 + 50000 AS INTEGER)      ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''29878.'' THEN CAST(29878 + 50000 AS INTEGER)      ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''29870,'' THEN CAST(29870 + 50000 AS INTEGER)      ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''2995,'' THEN CAST(2995 + 50000 AS INTEGER)      ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''30023Q'' THEN CAST(30023 + 50000 AS INTEGER)      ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''30045,'' THEN CAST(30045 + 50000 AS INTEGER)      ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''30304,'' THEN CAST(30304 + 50000 AS INTEGER)      ');
     SQL.Add('           WHEN CHAR_LENGTH(PRODUTOS.PRO_REFERENCIA) = 9 THEN PRODUTOS.COD_PRODUTO      ');
     SQL.Add('           ELSE PRODUTOS.PRO_REFERENCIA           ');
     SQL.Add('       END AS COD_PRODUTO,    ');
     SQL.Add('      ');
     SQL.Add('       CASE PRODUTOS_ESTOQUES.PST_LOCAL   ');
     SQL.Add('           WHEN 1 THEN COALESCE(ESTOQUE.PST_FISICO, 1)   ');
     SQL.Add('           ELSE COALESCE(PRODUTOS_ESTOQUES.PST_FISICO, 1)   ');
     SQL.Add('       END AS QTD_EST_ATUAL,   ');
     SQL.Add('      ');
     SQL.Add('       CASE PRODUTOS_ESTOQUES.PST_LOCAL   ');
     SQL.Add('           WHEN 2 THEN COALESCE(DEPOSITO.PST_FISICO, 1)   ');
     SQL.Add('           ELSE COALESCE(PRODUTOS_ESTOQUES.PST_FISICO, 1)   ');
     SQL.Add('       END AS QTD_EST_DEP   ');
     SQL.Add('   FROM   ');
     SQL.Add('       PRODUTOS   ');
     SQL.Add('   LEFT JOIN PRODUTOS_ESTOQUES ON PRODUTOS.PRO_CODIGO = PRODUTOS_ESTOQUES.PST_PRODUTO   ');
     SQL.Add('   LEFT JOIN (      ');
     SQL.Add('       SELECT      ');
     SQL.Add('           PST_LOCAL,      ');
     SQL.Add('           PST_PRODUTO,      ');
     SQL.Add('           PST_FISICO      ');
     SQL.Add('      FROM      ');
     SQL.Add('           PRODUTOS_ESTOQUES      ');
     SQL.Add('       WHERE PST_LOCAL = 2   ');
     SQL.Add('       GROUP BY PST_LOCAL, PST_PRODUTO, PST_FISICO      ');
     SQL.Add('   ) AS ESTOQUE      ');
     SQL.Add('   ON      ');
     SQL.Add('       PRODUTOS_ESTOQUES.PST_PRODUTO = ESTOQUE.PST_PRODUTO   ');
     SQL.Add('   LEFT JOIN (   ');
     SQL.Add('       SELECT   ');
     SQL.Add('           PST_LOCAL,      ');
     SQL.Add('           PST_PRODUTO,      ');
     SQL.Add('           PST_FISICO      ');
     SQL.Add('      FROM      ');
     SQL.Add('           PRODUTOS_ESTOQUES      ');
     SQL.Add('       WHERE PST_LOCAL = 1   ');
     SQL.Add('       GROUP BY PST_LOCAL, PST_PRODUTO, PST_FISICO      ');
     SQL.Add('   ) AS DEPOSITO   ');
     SQL.Add('   ON      ');
     SQL.Add('       PRODUTOS_ESTOQUES.PST_PRODUTO = DEPOSITO.PST_PRODUTO   ');
     SQL.Add('   WHERE (      ');
     SQL.Add('       CASE PRODUTOS_ESTOQUES.PST_LOCAL      ');
     SQL.Add('           WHEN 2 THEN 1      ');
     SQL.Add('           WHEN 4 THEN 6      ');
     SQL.Add('           WHEN 5 THEN 3      ');
     SQL.Add('       END      ');
     SQL.Add('   ) = '+CbxLoja.Text+'   ');
     SQL.Add('   ORDER BY (         ');
     SQL.Add('       CASE WHEN PRODUTOS.PRO_LEGENDA IN (''N'',''P'') THEN PRODUTOS.PRO_LEGENDA END         ');
     SQL.Add('   ) DESC,         ');
     SQL.Add('   (         ');
     SQL.Add('       CASE WHEN PRODUTOS.PRO_LEGENDA NOT IN (''N'', ''P'') THEN PRODUTOS.PRO_LEGENDA END         ');
     SQL.Add('   ), PRODUTOS.PRO_CODIGO ASC   ');


    Open;
    First;

    NumLinha := 0;
    //NovoCodigo := 4000000;

    while not Eof do
    begin
      try
        if Cancelar then
        Break;

        Inc(NumLinha);

        COD_PRODUTO := GerarPLU(QryPrincipal.FieldByName('COD_PRODUTO').AsString);



        if CbxLoja.Text = '1' then
        begin
          Writeln(Arquivo, 'UPDATE TAB_PRODUTO_LOJA SET QTD_EST_ATUAL = '+QryPrincipal.FieldByName('QTD_EST_ATUAL').AsString+', QTD_EST_DEP = '+QryPrincipal.FieldByName('QTD_EST_DEP').AsString+'  WHERE COD_PRODUTO = '+COD_PRODUTO+' AND COD_LOJA = 1; ');
        end;

        if CbxLoja.Text = '3' then
        begin
          Writeln(Arquivo, 'UPDATE TAB_PRODUTO_LOJA SET QTD_EST_ATUAL = '+QryPrincipal.FieldByName('QTD_EST_ATUAL').AsString+'  WHERE COD_PRODUTO = '+COD_PRODUTO+' AND COD_LOJA = 3; ');
        end;

        if CbxLoja.Text = '6' then
        begin
          Writeln(Arquivo, 'UPDATE TAB_PRODUTO_LOJA SET QTD_EST_ATUAL = '+QryPrincipal.FieldByName('QTD_EST_ATUAL').AsString+'  WHERE COD_PRODUTO = '+COD_PRODUTO+' AND COD_LOJA = 2; ');
        end;



          if NumLinha = 500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 1000 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 1500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 2000 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 2500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 3000 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 3500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 4000 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 4500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 5000 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 5500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 6000 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 6500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 7000 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 7500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 8000 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 8500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 9000 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 9500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 10000 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 10500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 20500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 30500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 40500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 50500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 60500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 70000 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 70500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 80000 then
            Writeln(Arquivo, 'COMMIT WORK;');

      except on E: Exception do
        FrmProgresso.AdicionarLog(QryPrincipal.RecNo, 'E', E.Message);
      end;
        Next;
    end;
      Writeln(Arquivo, 'COMMIT WORK;');
      Close;
  end;
end;

procedure TFrmHiperLojao.btnAtualizaEstoqueClick(Sender: TObject);
begin
  inherited;
  FlgAtualizaEstoque := True;
  BtnGerar.Click;
  FlgAtualizaEstoque := False;
end;

procedure TFrmHiperLojao.btnCustoCheioClick(Sender: TObject);
begin
  inherited;
  FlgGeraCustoCheio := True;
  BtnGerar.Click;
  FlgGeraCustoCheio := False;
end;

procedure TFrmHiperLojao.btnEstoqueTrocaClick(Sender: TObject);
begin
  inherited;
  FlgEstoqueTroca := True;
  BtnGerar.Click;
  FlgEstoqueTroca := False;

end;

procedure TFrmHiperLojao.btnFidelidadeClick(Sender: TObject);
begin
  inherited;
  FlgPercentualFidelidade := True;
  BtnGerar.Click;
  FlgPercentualFidelidade :=  False;
end;

procedure TFrmHiperLojao.btnFornPrefeClick(Sender: TObject);
begin
  inherited;
  FlgFornecedorPreferencial := True;
  BtnGerar.Click;
  FlgFornecedorPreferencial := False;
end;

procedure TFrmHiperLojao.BtnGerarClick(Sender: TObject);
begin

  if FlgGeraEstoqueDeposito then
  begin
    AssignFile(Arquivo, EdtCamArquivo.Text + '\SCRIPT_ESTOQUE_DEPOSITO.TXT' );
    Rewrite(Arquivo);
    CkbProduto.Checked := True;
  end;

  if FlgGeraCustoCheio then
  begin
    AssignFile(Arquivo, EdtCamArquivo.Text + '\SCRIPT_CUSTO_CHEIO.TXT' );
    Rewrite(Arquivo);
    CkbProduto.Checked := True;
  end;

  if FlgAtualizaEstoque then
  begin
    AssignFile(Arquivo, EdtCamArquivo.Text + '\SCRIPT_ATUALIZA_ESTOQUE.TXT' );
    Rewrite(Arquivo);
    CkbProdLoja.Checked := True;
  end;

  if FlgEstoqueTroca then
  begin
    AssignFile(Arquivo, EdtCamArquivo.Text + '\SCRIPT_ESTOQUE_TROCA.TXT' );
    Rewrite(Arquivo);
    CkbProdForn.Checked := True;
  end;


  if FlgPercentualFidelidade then
  begin
    AssignFile(Arquivo, EdtCamArquivo.Text + '\SCRIPT_PERCENTUAL_FIDELIDADE.TXT' );
    Rewrite(Arquivo);
    CkbProdLoja.Checked := True;
  end;

  if FlgFornecedorPreferencial then
  begin
    AssignFile(Arquivo, EdtCamArquivo.Text + '\SCRIPT_FORNECEDOR_PREFERENCIAL.TXT' );
    Rewrite(Arquivo);
    CkbProdForn.Checked := True;
  end;

  inherited;

  if FlgGeraEstoqueDeposito then
    CloseFile(Arquivo);

  if FlgGeraCustoCheio then
    CloseFile(Arquivo);

  if FlgAtualizaEstoque then
    CloseFile(Arquivo);

  if FlgEstoqueTroca then
    CloseFile(Arquivo);

  if FlgPercentualFidelidade then
    CloseFile(Arquivo);

  if FlgFornecedorPreferencial then
    CloseFile(Arquivo);


end;

procedure TFrmHiperLojao.btnScriptDepositoClick(Sender: TObject);
begin
  inherited;
  FlgGeraEstoqueDeposito := True;
  BtnGerar.Click;
  FlgGeraEstoqueDeposito := False;
end;

procedure TFrmHiperLojao.CkbProdFornClick(Sender: TObject);
begin
  inherited;
  btnEstoqueTroca.Enabled := True;
  btnFornPrefe.Enabled := True;

  if CkbProdForn.Checked = False then
  begin
    btnEstoqueTroca.Enabled := False;
    btnFornPrefe.Enabled := False;
  end;
end;

procedure TFrmHiperLojao.CkbProdLojaClick(Sender: TObject);
begin
  inherited;
  btnAtualizaEstoque.Enabled := True;
  btnFidelidade.Enabled := True;

  if CkbProdLoja.Checked = False then
  begin
    btnAtualizaEstoque.Enabled := False;
    btnFidelidade.Enabled := False;
  end;

end;

procedure TFrmHiperLojao.CkbProdutoClick(Sender: TObject);
begin
  inherited;
  btnScriptDeposito.Enabled := True;
  btnCustoCheio.Enabled := True;
  
  if CkbProduto.Checked = False then
    begin
      btnScriptDeposito.Enabled := False;
      btnCustoCheio.Enabled := False;
    end;
    
end;

procedure TFrmHiperLojao.EdtCamBancoExit(Sender: TObject);
begin
  inherited;
  CriarFB(EdtCamBanco);
end;

procedure TFrmHiperLojao.EstoqueTroca;
var
  COD_PRODUTO : string;
begin
  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

     SQL.Add('   SELECT DISTINCT   ');
     SQL.Add('      ');
     SQL.Add('       CASE PRODUTOS_ESTOQUES.PST_LOCAL   ');
     SQL.Add('           WHEN 1 THEN 1   ');
     SQL.Add('           WHEN 3 THEN 3   ');
     SQL.Add('           WHEN 6 THEN 2   ');
     SQL.Add('       END AS COD_LOJA,   ');
     SQL.Add('      ');
     SQL.Add('       CASE   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''25578,'' THEN CAST(25578 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''28687.'' THEN CAST(28687 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''29397.'' THEN CAST(29397 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''29*504'' THEN CAST(29504 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''296,5'' THEN CAST(2965 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''29878.'' THEN CAST(29878 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''29870,'' THEN CAST(29870 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''2995,'' THEN CAST(2995 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''30023Q'' THEN CAST(30023 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''30045,'' THEN CAST(30045 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''30304,'' THEN CAST(30304 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN CHAR_LENGTH(PRODUTOS.PRO_REFERENCIA) = 9 THEN PRODUTOS.COD_PRODUTO   ');
     SQL.Add('           ELSE PRODUTOS.PRO_REFERENCIA   ');
     SQL.Add('       END AS COD_PRODUTO,   ');
     SQL.Add('      ');
     SQL.Add('       PRODUTOS_FORNECEDORES.PRF_FORNECEDOR AS COD_FORNECEDOR,   ');
     SQL.Add('       PRODUTOS_ESTOQUES.PST_FISICO AS QTD_EST_TROCA');
     SQL.Add('      ');
     SQL.Add('   FROM   ');
     SQL.Add('       PRODUTOS_ESTOQUES   ');
     SQL.Add('   LEFT JOIN PRODUTOS_FORNECEDORES ON PRODUTOS_FORNECEDORES.PRF_CODIGO = PRODUTOS_ESTOQUES.PST_PRODUTO   ');
     SQL.Add('   LEFT JOIN PRODUTOS ON PRODUTOS.PRO_CODIGO = PRODUTOS_ESTOQUES.PST_PRODUTO   ');
     SQL.Add('   WHERE PRODUTOS_ESTOQUES.PST_LOCAL = 3   ');
     SQL.Add('   AND PRODUTOS_FORNECEDORES.PRF_FORNECEDOR IS NOT NULL   ');
     SQL.Add('   AND PRODUTOS_ESTOQUES.PST_FISICO <> 0   ');
     SQL.Add('   ORDER BY(         ');
     SQL.Add('       CASE WHEN PRODUTOS.PRO_LEGENDA IN (''N'',''P'') THEN PRODUTOS.PRO_LEGENDA END         ');
     SQL.Add('   ) DESC,         ');
     SQL.Add('   (         ');
     SQL.Add('       CASE WHEN PRODUTOS.PRO_LEGENDA NOT IN (''N'', ''P'') THEN PRODUTOS.PRO_LEGENDA END         ');
     SQL.Add('   ), PRODUTOS.PRO_CODIGO ASC         ');
     SQL.Add('      ');


    Open;
    First;

    NumLinha := 0;

    while not Eof do
    begin
      try
        if Cancelar then
          Break;

          Inc(NumLinha);

          COD_PRODUTO := GerarPLU(QryPrincipal.FieldByName('COD_PRODUTO').AsString);

          Writeln(Arquivo, 'UPDATE TAB_PRODUTO_FORNECEDOR SET QTD_EST_TROCA = '+QryPrincipal.FieldByName('QTD_EST_TROCA').AsString+'  WHERE COD_PRODUTO = '+COD_PRODUTO+' AND COD_FORNECEDOR = '+QryPrincipal.FieldByName('COD_FORNECEDOR').AsString+'; ');

          if NumLinha = 500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 1000 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 1500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 2000 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 2500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 3000 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 3500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 4000 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 4500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 5000 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 5500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 6000 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 6500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 7000 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 7500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 8000 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 8500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 9000 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 9500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 10000 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 10500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 20500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 30500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 40500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 50500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 60500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 70000 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 70500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 80000 then
            Writeln(Arquivo, 'COMMIT WORK;');


      except on E: Exception do
        FrmProgresso.AdicionarLog(QryPrincipal.RecNo, 'E', E.Message);
      end;
        Next;
    end;
      Writeln(Arquivo, 'COMMIT WORK;');
      Close;
  end;
end;

procedure TFrmHiperLojao.FornecedorPreferencial;
var
  COD_PRODUTO : string;
begin
  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

     SQL.Add('   SELECT DISTINCT   ');
     SQL.Add('       CASE     ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''25578,'' THEN CAST(25578 + 50000 AS INTEGER)      ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''28687.'' THEN CAST(28687 + 50000 AS INTEGER)      ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''29397.'' THEN CAST(29397 + 50000 AS INTEGER)      ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''29*504'' THEN CAST(29504 + 50000 AS INTEGER)      ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''296,5'' THEN CAST(2965 + 50000 AS INTEGER)      ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''29878.'' THEN CAST(29878 + 50000 AS INTEGER)      ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''29870,'' THEN CAST(29870 + 50000 AS INTEGER)      ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''2995,'' THEN CAST(2995 + 50000 AS INTEGER)      ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''30023Q'' THEN CAST(30023 + 50000 AS INTEGER)      ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''30045,'' THEN CAST(30045 + 50000 AS INTEGER)      ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''30304,'' THEN CAST(30304 + 50000 AS INTEGER)      ');
     SQL.Add('           WHEN CHAR_LENGTH(PRODUTOS.PRO_REFERENCIA) = 9 THEN PRODUTOS.COD_PRODUTO      ');
     SQL.Add('           ELSE PRODUTOS.PRO_REFERENCIA           ');
     SQL.Add('       END AS COD_PRODUTO,   ');
     SQL.Add('       PRODUTOS_FORNECEDORES.PRF_PADRAO AS FLG_PREFERENCIAL,   ');
     SQL.Add('       FORNECEDORES.FOR_CODIGO AS COD_FORNECEDOR ');
     SQL.Add('   FROM PRODUTOS_FORNECEDORES   ');
     SQL.Add('   INNER JOIN PRODUTOS ON PRODUTOS_FORNECEDORES.PRF_CODIGO = PRODUTOS.PRO_CODIGO      ');
     SQL.Add('   INNER JOIN FORNECEDORES ON PRODUTOS_FORNECEDORES.PRF_FORNECEDOR = FORNECEDORES.FOR_CODIGO      ');
     SQL.Add('   LEFT JOIN PRODUTOS_EMPRESA ON PRODUTOS.PRO_CODIGO = PRODUTOS_EMPRESA.PRE_PRODUTO   ');
     SQL.Add('   WHERE PRODUTOS_EMPRESA.PRE_EMPRESA = '+CbxLoja.Text+'   ');
     SQL.Add('   AND PRODUTOS_FORNECEDORES.PRF_PADRAO = ''S''   ');
     SQL.Add('   ORDER BY(      ');
     SQL.Add('       CASE WHEN PRODUTOS.PRO_LEGENDA IN (''N'',''P'') THEN PRODUTOS.PRO_LEGENDA END      ');
     SQL.Add('   ) DESC,      ');
     SQL.Add('   (      ');
     SQL.Add('       CASE WHEN PRODUTOS.PRO_LEGENDA NOT IN (''N'', ''P'') THEN PRODUTOS.PRO_LEGENDA END      ');
     SQL.Add('   ), PRODUTOS.PRO_CODIGO ASC   ');

     Open;
     First;
     NumLinha := 0;

     while not Eof do
     begin
       try
        if Cancelar then
          Break;

          Inc(NumLinha);

          COD_PRODUTO := GerarPLU(QryPrincipal.FieldByName('COD_PRODUTO').AsString);

          Writeln(Arquivo, 'UPDATE TAB_PRODUTO_FORNECEDOR SET FLG_PREFERENCIAL = '''+QryPrincipal.FieldByName('FLG_PREFERENCIAL').AsString+'''  WHERE COD_PRODUTO = '+COD_PRODUTO+' AND COD_FORNECEDOR = '+QryPrincipal.FieldByName('COD_FORNECEDOR').AsString+'; ');

          if NumLinha = 500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 1000 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 1500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 2000 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 2500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 3000 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 3500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 4000 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 4500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 5000 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 5500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 6000 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 6500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 7000 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 7500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 8000 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 8500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 9000 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 9500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 10000 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 10500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 20500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 30500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 40500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 50500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 60500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 70000 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 70500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 80000 then
            Writeln(Arquivo, 'COMMIT WORK;');
          
       except on E: Exception do
         FrmProgresso.AdicionarLog(QryPrincipal.RecNo, 'E', E.Message);
       end;
       Next;
     end;
     Writeln(Arquivo, 'COMMIT WORK;');
     Close;
  end;
end;

procedure TFrmHiperLojao.GerarCest;
var
   TotalCount : integer;
   count : integer;
begin
  inherited;
  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT DISTINCT');
    SQL.Add('    COD_CEST,');
    SQL.Add('    NUM_CEST,');
    SQL.Add('    DES_CEST');
    SQL.Add('FROM');
    SQL.Add('(');
    SQL.Add('    SELECT ');
    SQL.Add('        0 AS COD_CEST,');
    SQL.Add('        CASE');
    SQL.Add('            WHEN COALESCE(PRODUTO.CEST, '''') = '''' THEN ''9999999''');
    SQL.Add('            ELSE PRODUTO.CEST');
    SQL.Add('        END AS NUM_CEST,');
    SQL.Add('        ''A DEFINIR'' AS DES_CEST');
    SQL.Add('    FROM PRODUTOS AS PRODUTO');
    SQL.Add('');
    SQL.Add('    UNION ALL');
    SQL.Add('');
    SQL.Add('    SELECT');
    SQL.Add('        0 AS COD_CEST,');
    SQL.Add('        CASE');
    SQL.Add('            WHEN COALESCE(PRODUTO.CEST, '''') = '''' THEN ''9999999''');
    SQL.Add('            WHEN PRODUTO.CEST = ''FF'' THEN ''9999999''');
    SQL.Add('            ELSE PRODUTO.CEST');
    SQL.Add('        END AS NUM_CEST,');
    SQL.Add('        ''A DEFINIR'' AS DES_CEST');
    SQL.Add('    FROM TAB_PRODUTO_AUX AS PRODUTO');
    SQL.Add(')');



    Open;
    First;

    count := 0;
    NumLinha := 0;

    TotalCount := SetCountTotal(SQL.Text);

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);
      Inc(count);
      Layout.SetValues(QryPrincipal, NumLinha, TotalCount);

      Layout.FieldByName('COD_CEST').AsInteger := count;
      Layout.FieldByName('NUM_CEST').AsString := StrRetNums( Layout.FieldByName('NUM_CEST').AsString );

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;
end;

procedure TFrmHiperLojao.GerarCliente;
begin
  inherited;

  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

     SQL.Add('   SELECT   ');
     SQL.Add('       CLIENTES.CLI_CODIGO AS COD_CLIENTE,   ');
     SQL.Add('       CLIENTES.CLI_RAZAO AS DES_CLIENTE,   ');
     SQL.Add('       COALESCE(REPLACE(REPLACE(REPLACE(CLIENTES.CLI_CGC, ''.'', ''''), ''-'', ''''), ''/'', ''''), '''') AS NUM_CGC,   ');
     SQL.Add('       CASE CLIENTES.CLI_IE   ');
     SQL.Add('           WHEN ''01'' THEN ''ISENTO''   ');
     SQL.Add('           WHEN ''1'' THEN ''ISENTO''   ');
     SQL.Add('           WHEN NULL THEN ''ISENTO''   ');
     SQL.Add('           WHEN ''isento''  THEN ''ISENTO'' ');
     SQL.Add('           ELSE COALESCE(REPLACE(REPLACE(REPLACE(CLIENTES.CLI_IE, ''.'', ''''), ''-'', ''''), ''/'', ''''), ''ISENTO'')    ');
     SQL.Add('       END AS NUM_INSC_EST,   ');
     SQL.Add('       CLIENTES.CLI_ENDFAT AS DES_ENDERECO,   ');
     SQL.Add('       CLIENTES.CLI_BAIFAT AS DES_BAIRRO,   ');
     SQL.Add('       CLIENTES.CLI_CIDFAT AS DES_CIDADE,   ');
     SQL.Add('       CLIENTES.CLI_UFFAT AS DES_SIGLA,   ');
     SQL.Add('       CLIENTES.CLI_CEPFAT AS NUM_CEP,   ');
     SQL.Add('       CLIENTES.CLI_FONE AS NUM_FONE,   ');
     SQL.Add('       CLIENTES.CLI_FAX AS NUM_FAX,   ');
     SQL.Add('       '''' AS DES_CONTATO,   ');
     SQL.Add('       0 AS FLG_SEXO,   ');
     SQL.Add('       0 AS VAL_LIMITE_CRETID,   ');
     SQL.Add('       CLIENTES.CLI_LIMITE AS VAL_LIMITE_CONV,   ');
     SQL.Add('       0 AS VAL_DEBITO,   ');
     SQL.Add('       0 AS VAL_RENDA,   ');
     SQL.Add('       0 AS COD_CONVENIO,   ');
     SQL.Add('       0 AS COD_STATUS_PDV,   ');
     SQL.Add('      ');
     SQL.Add('       CASE CLIENTES.CLI_PESSOA   ');
     SQL.Add('           WHEN ''J'' THEN ''S''   ');
     SQL.Add('           ELSE ''N''   ');
     SQL.Add('       END AS FLG_EMPRESA,   ');
     SQL.Add('      ');
     SQL.Add('       ''N'' AS FLG_CONVENIO,   ');
     SQL.Add('       ''N'' AS MICRO_EMPRESA,   ');
     SQL.Add('       CLIENTES.CLI_CADASTRO AS DTA_CADASTRO,   ');
     SQL.Add('       COALESCE(CLIENTES.CLI_NUMFAT, ''S/N'') AS NUM_ENDERECO,   ');
     SQL.Add('       '''' AS NUM_RG,   ');
     SQL.Add('       0 AS FLG_EST_CIVIL,   ');
     SQL.Add('       CLIENTES.CLI_FONE1 AS NUM_CELULAR,   ');
     SQL.Add('       '''' AS DTA_ALTERACAO,   ');
     SQL.Add('       CAST(COALESCE(CLIENTES.CLI_OBS, '''') AS VARCHAR(300)) AS DES_OBSERVACAO,   ');
     SQL.Add('       '''' AS DES_COMPLEMENTO,   ');
     SQL.Add('       COALESCE(CLIENTES.CLI_EMAIL, '''') AS DES_EMAIL,   ');
     SQL.Add('       CLIENTES.CLI_FANTASIA AS DES_FANTASIA,   ');
     SQL.Add('       '''' AS DTA_NASCIMENTO,   ');
     SQL.Add('       '''' AS DES_PAI,   ');
     SQL.Add('       '''' AS DES_MAE,   ');
     SQL.Add('       '''' AS DES_CONJUGE,   ');
     SQL.Add('       '''' AS NUM_CPF_CONJUGE,   ');
     SQL.Add('       0 AS VAL_DEB_CONV,   ');
     SQL.Add('       ''N'' AS INATIVO,   ');
     SQL.Add('       '''' AS DES_MATRICULA,   ');
     SQL.Add('       ''N'' AS NUM_CGC_ASSOCIADO,   ');
     SQL.Add('       ''N'' AS FLG_PROD_RURAL,   ');
     SQL.Add('       0 AS COD_STATUS_PDV_CONV,   ');
     SQL.Add('       ''S'' AS FLG_ENVIA_CODIGO,   ');
     SQL.Add('       '''' AS DTA_NASC_CONJUGE,   ');
     SQL.Add('       0 AS COD_CLASSIF   ');
     SQL.Add('   FROM   ');
     SQL.Add('       CLIENTES   ');




    Open;
    First;
    NumLinha := 0;

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);

      Layout.SetValues(QryPrincipal, NumLinha, RecordCount);

      Layout.FieldByName('NUM_CGC').AsString := StrRetNums(Layout.FieldByName('NUM_CGC').AsString);

//      if StrRetNums(Layout.FieldByName('NUM_RG').AsString) = '' then
//        Layout.FieldByName('NUM_RG').AsString := ''
//      else
//        Layout.FieldByName('NUM_RG').AsString := StrRetNums(Layout.FieldByName('NUM_RG').AsString);

      Layout.FieldByName('NUM_CEP').AsString := StrRetNums(Layout.FieldByName('NUM_CEP').AsString);

//      if QryPrincipal.FieldByName('NUM_INSC_EST').AsString <> 'ISENTO' then
//         Layout.FieldByName('NUM_INSC_EST').AsString := StrRetNums(Layout.FieldByName('NUM_INSC_EST').AsString);

//      Layout.FieldByName('DTA_NASCIMENTO').AsDateTime := FieldByName('DTA_NASCIMENTO').AsDateTime;
      Layout.FieldByName('DTA_CADASTRO').AsDateTime := FieldByName('DTA_CADASTRO').AsDateTime;

      Layout.FieldByName('NUM_FONE').AsString := StrRetNums( FieldByName('NUM_FONE').AsString );




      if Layout.FieldByName('FLG_EMPRESA').AsString = 'S' then
      begin
        if not ValidaCGC(Layout.FieldByName('NUM_CGC').AsString) then
          Layout.FieldByName('NUM_CGC').AsString := '';
      end
      else
        if not ValidaCpf(Layout.FieldByName('NUM_CGC').AsString) then
          Layout.FieldByName('NUM_CGC').AsString := '';



      Layout.FieldByName('DES_OBSERVACAO').AsString := StrLBReplace(FieldByName('DES_OBSERVACAO').AsString);
      Layout.FieldByName('DES_EMAIL').AsString := StrReplace(StrLBReplace(FieldByName('DES_EMAIL').AsString), '\n', '');

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;

end;

procedure TFrmHiperLojao.GerarCodigoBarras;
var
 count, NovoCodigo : Integer;
 cod_antigo, codbarras : string;
begin
  inherited;
  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

     SQL.Add('   SELECT    ');
     SQL.Add('       CASE      ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''25578,'' THEN CAST(25578 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''28687.'' THEN CAST(28687 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''29397.'' THEN CAST(29397 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''29*504'' THEN CAST(29504 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''296,5'' THEN CAST(2965 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''29878.'' THEN CAST(29878 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''29870,'' THEN CAST(29870 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''2995,'' THEN CAST(2995 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''30023Q'' THEN CAST(30023 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''30045,'' THEN CAST(30045 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''30304,'' THEN CAST(30304 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN CHAR_LENGTH(PRODUTOS.PRO_REFERENCIA) = 9 THEN PRODUTOS.COD_PRODUTO   ');
     SQL.Add('           ELSE PRODUTOS.PRO_REFERENCIA        ');
     SQL.Add('       END AS COD_PRODUTO,   ');
     SQL.Add('       CASE   ');
     SQL.Add('           WHEN LINHAS.LIN_CODIGO = 1 AND GRUPOS.GRU_CODIGO = 23 AND SUB_GRUPOS.SGR_CODIGO = 9999 THEN PRODUTOS.PRO_REFERENCIA   ');
     SQL.Add('           WHEN LINHAS.LIN_CODIGO = 1 AND GRUPOS.GRU_CODIGO = 23 AND SUB_GRUPOS.SGR_CODIGO = 9996 THEN PRODUTOS.PRO_REFERENCIA   ');
     SQL.Add('           ELSE PRODUTOS.PRO_BARRA   ');
     SQL.Add('        END AS COD_EAN   ');
     SQL.Add('   FROM      ');
     SQL.Add('       PRODUTOS   ');
     SQL.Add('   LEFT JOIN LINHAS ON PRODUTOS.PRO_LINHA = LINHAS.LIN_CODIGO      ');
     SQL.Add('   LEFT JOIN GRUPOS ON PRODUTOS.PRO_GRUPO = GRUPOS.GRU_CODIGO      ');
     SQL.Add('   LEFT JOIN SUB_GRUPOS ON PRODUTOS.PRO_SUBGRUPO = SUB_GRUPOS.SGR_CODIGO     ');
     SQL.Add('   ORDER BY(      ');
     SQL.Add('       CASE WHEN PRODUTOS.PRO_LEGENDA IN (''N'',''P'') THEN PRODUTOS.PRO_LEGENDA END      ');
     SQL.Add('   ) DESC,      ');
     SQL.Add('   (      ');
     SQL.Add('       CASE WHEN PRODUTOS.PRO_LEGENDA NOT IN (''N'', ''P'') THEN PRODUTOS.PRO_LEGENDA END      ');
     SQL.Add('   ), PRODUTOS.PRO_CODIGO ASC      ');

     //SQL.Add('   LEFT JOIN PRODUTOS_EMPRESA ON PRODUTOS.PRO_CODIGO = PRODUTOS_EMPRESA.PRE_PRODUTO');
     //SQL.Add('   WHERE PRODUTOS_EMPRESA.PRE_EMPRESA = '+CbxLoja.Text+' ');


    Open;
    First;
    NumLinha := 0;
    NovoCodigo := 400000;

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);

      Layout.SetValues(QryPrincipal, NumLinha, RecordCount);



      Layout.FieldByName('COD_PRODUTO').AsString := GerarPLU( Layout.FieldByName('COD_PRODUTO').AsString );

//      Layout.FieldByName('COD_PRODUTO').AsString := StrRetNums(Layout.FieldByName('COD_PRODUTO').AsString);
//
//      codbarras := StrRetNums(Layout.FieldByName('COD_EAN').AsString);
//
//      if( (codbarras = '') or (StrToFloat(codbarras) = 0)) then
//         codbarras := ''
//      else if ( Length(TiraZerosEsquerda(codbarras)) < 8 ) then
//         codbarras := GerarPLU(codbarras)
//      else
//         if(not CodBarrasValido(codbarras)) then
//            codbarras := '';
//
//      Layout.FieldByName('COD_EAN').AsString := codbarras;

//      if Layout.FieldByName('COD_PRODUTO').AsInteger = 0 then
//      begin
//        Inc(NovoCodigo);
//        Layout.FieldByName('COD_PRODUTO').AsString := GerarPLU(IntToStr(NovoCodigo));
//      end;


//
      Layout.FieldByName('COD_EAN').AsString := StrRetNums(Layout.FieldByName('COD_EAN').AsString);

      if Length(StrLBReplace(Trim(StrRetNums( FieldByName('COD_EAN').AsString) ))) < 8 then
        Layout.FieldByName('COD_EAN').AsString := GerarPLU(FieldByName('COD_EAN').AsString);

      if not CodBarrasValido(Layout.FieldByName('COD_EAN').AsString) then
        Layout.FieldByName('COD_EAN').AsString := '';



      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;
end;

procedure TFrmHiperLojao.GerarComposicao;
begin
  inherited;
  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
    SQL.Add('    COMPOSICAO.PRODUTO_BASE AS COD_PRODUTO,');
    SQL.Add('    COMPOSICAO.PRODUTO AS COD_PRODUTO_COMP,');
    SQL.Add('    COMPOSICAO.QTDE AS QTD_PRODUTO,');
    SQL.Add('    0 AS VAL_VENDA,');
    SQL.Add('    0 AS PER_RATEIO,');
    SQL.Add('    0 AS VAL_DIF');
    SQL.Add('FROM');
    SQL.Add('    PRODUTOS');
    SQL.Add('LEFT JOIN');
    SQL.Add('    PRODUTOS_COMPOSICAO COMPOSICAO');
    SQL.Add('ON');
    SQL.Add('    PRODUTOS.ID = COMPOSICAO.PRODUTO_BASE    ');
    SQL.Add('WHERE');
    SQL.Add('    PRODUTOS.COMPOSTO = 1');
    SQL.Add('AND');
    SQL.Add('    COMPOSICAO.PRODUTO_BASE IS NOT NULL');

    Open;

    First;
    NumLinha := 0;

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);

      Layout.SetValues(QryPrincipal, NumLinha, RecordCount);

//      Layout.FieldByName('COD_PRODUTO').AsString := GerarPLU( Layout.FieldByName('COD_PRODUTO').AsString );
//
//      Layout.FieldByName('COD_PRODUTO_COMP').AsString := GerarPLU( Layout.FieldByName('COD_PRODUTO_COMP').AsString );

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;
end;

procedure TFrmHiperLojao.GerarCondPagCli;
begin
  inherited;

  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

     SQL.Add('   SELECT   ');
     SQL.Add('       CLIENTES.CLI_CODIGO AS COD_CLIENTE,   ');
     SQL.Add('       30 AS NUM_CONDICAO,   ');
     SQL.Add('       2 AS COD_CONDICAO,   ');
     SQL.Add('       1 AS COD_ENTIDADE   ');
     SQL.Add('   FROM   ');
     SQL.Add('       CLIENTES   ');



    Open;

    First;
    NumLinha := 0;

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);

      Layout.SetValues(QryPrincipal, NumLinha, RecordCount);

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;

end;

procedure TFrmHiperLojao.GerarCondPagForn;
begin
  inherited;

  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

     SQL.Add('   SELECT   ');
     SQL.Add('       FORNECEDORES.FOR_CODIGO AS COD_FORNECEDOR,   ');
     SQL.Add('       30 AS NUM_CONDICAO,   ');
     SQL.Add('       2 AS COD_CONDICAO,   ');
     SQL.Add('       8 AS COD_ENTIDADE,   ');
     SQL.Add('       REPLACE(REPLACE(REPLACE(FORNECEDORES.FOR_CGC, ''.'', ''''), ''-'', ''''), ''/'', '''') AS NUM_CGC   ');
     SQL.Add('   FROM   ');
     SQL.Add('       FORNECEDORES   ');



    Open;

    First;

    NumLinha := 0;

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);

      Layout.SetValues(QryPrincipal, NumLinha, RecordCount);

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;

end;

procedure TFrmHiperLojao.GerarCustoCheio;
var
  COD_PRODUTO : string;
  NovoCodigo : Integer;
begin
  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

     SQL.Add('   SELECT   ');
     SQL.Add('       PRODUTOS_ESTOQUES.PST_LOCAL AS COD_LOJA,      ');
     SQL.Add('       CASE     ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''25578,'' THEN CAST(25578 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''28687.'' THEN CAST(28687 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''29397.'' THEN CAST(29397 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''29*504'' THEN CAST(29504 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''296,5'' THEN CAST(2965 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''29878.'' THEN CAST(29878 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''29870,'' THEN CAST(29870 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''2995,'' THEN CAST(2995 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''30023Q'' THEN CAST(30023 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''30045,'' THEN CAST(30045 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''30304,'' THEN CAST(30304 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN CHAR_LENGTH(PRODUTOS.PRO_REFERENCIA) = 9 THEN PRODUTOS.COD_PRODUTO   ');
     SQL.Add('           ELSE PRODUTOS.PRO_REFERENCIA        ');
     SQL.Add('       END AS COD_PRODUTO,      ');
     SQL.Add('       CASE   ');
     SQL.Add('           WHEN PRODUTOS.PRO_UNIDCOMPRA <> ''UN'' THEN SUBSTRING((PRODUTOS.PRO_VRLIQ_1 / PRODUTOS.PRO_FTUCUV) FROM 1 FOR 5)   ');
     SQL.Add('           ELSE SUBSTRING(REPLACE(COALESCE(PRODUTOS.PRO_VRLIQ_1, 0), '','', ''.'')FROM 1 FOR 5)   ');
     SQL.Add('       END AS VAL_CUSTO_CHEIO   ');
     SQL.Add('   FROM      ');
     SQL.Add('       PRODUTOS      ');
     SQL.Add('   LEFT JOIN PRODUTOS_ESTOQUES ON PRODUTOS_ESTOQUES.PST_PRODUTO = PRODUTOS.PRO_CODIGO      ');
     SQL.Add('   WHERE PRODUTOS_ESTOQUES.PST_LOCAL = '+CbxLoja.Text+'   ');
     SQL.Add('   AND PRODUTOS.PRO_VRLIQ_1 <> 0   ');
     SQL.Add('   AND PRODUTOS.PRO_LEGENDA IN (''N'', ''P'') ');
     SQL.Add('   ORDER BY(      ');
     SQL.Add('       CASE WHEN PRODUTOS.PRO_LEGENDA IN (''N'',''P'') THEN PRODUTOS.PRO_LEGENDA END      ');
     SQL.Add('   ) DESC,      ');
     SQL.Add('   (      ');
     SQL.Add('       CASE WHEN PRODUTOS.PRO_LEGENDA NOT IN (''N'', ''P'') THEN PRODUTOS.PRO_LEGENDA END      ');
     SQL.Add('   ), PRODUTOS.PRO_CODIGO ASC      ');
     SQL.Add('      ');

     Open;
     First;

     NumLinha := 0;
     NovoCodigo := 400000;

     while not Eof do
     begin
       try
          if Cancelar then
          Break;

          Inc(NumLinha);

//        if QryPrincipal.FieldByName('COD_PRODUTO').AsInteger = 0 then
//        begin
//          Inc(NovoCodigo);
//          COD_PRODUTO := GerarPLU(IntToStr(NovoCodigo));
//        end;

          COD_PRODUTO := GerarPLU(QryPrincipal.FieldByName('COD_PRODUTO').AsString);

          Writeln(Arquivo, 'UPDATE TAB_PRODUTO_LOJA SET VAL_CUSTO_CHEIO = '+QryPrincipal.FieldByName('VAL_CUSTO_CHEIO').AsString+'  WHERE COD_PRODUTO = '+COD_PRODUTO+' ; ');


          if NumLinha = 500 then        
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 1000 then        
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 1500 then        
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 2000 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 2500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 3000 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 3500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 4000 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 4500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 5000 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 5500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 6000 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 6500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 7000 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 7500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 8000 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 8500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 9000 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 9500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 10000 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 10500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 20500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 30500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 40500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 50500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 60500 then
            Writeln(Arquivo, 'COMMIT WORK;');
          if NumLinha = 70000 then
            Writeln(Arquivo, 'COMMIT WORK;');

//          if Layout.FieldByName('COD_PRODUTO').AsInteger = 0 then
//          begin
//            Inc(NovoCodigo);
//            Layout.FieldByName('COD_PRODUTO').AsString := GerarPLU(IntToStr(NovoCodigo));
//          end;

       except on E: Exception do
         FrmProgresso.AdicionarLog(QryPrincipal.RecNo, 'E', E.Message);
       end;
        Next;
     end;
      Writeln(Arquivo, 'COMMIT WORK;');
      Close;
  end;
end;

procedure TFrmHiperLojao.GerarDecomposicao;
begin
  inherited;

  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
    SQL.Add('    DECOMPOSICAO.PRODUTO_BASE AS COD_PRODUTO,');
    SQL.Add('    DECOMPOSICAO.PRODUTO AS COD_PRODUTO_DECOM,');
    SQL.Add('    DECOMPOSICAO.QTDE * 100 AS QTD_DECOMP,');
    SQL.Add('    PRODUTOS.UNIDADE_COMPRA AS DES_UNIDADE');
    SQL.Add('FROM');
    SQL.Add('    PRODUTOS');
    SQL.Add('LEFT JOIN');
    SQL.Add('    PRODUTOS_COMPOSICAO DECOMPOSICAO');
    SQL.Add('ON');
    SQL.Add('    PRODUTOS.ID = DECOMPOSICAO.PRODUTO_BASE');
    SQL.Add('WHERE');
    SQL.Add('    PRODUTOS.COMPOSTO = 4');



    Open;

    First;
    NumLinha := 0;

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);

      Layout.SetValues(QryPrincipal, NumLinha, RecordCount);

//      Layout.FieldByName('COD_PRODUTO').AsString := GerarPLU( Layout.FieldByName('COD_PRODUTO').AsString );
//
//      Layout.FieldByName('COD_PRODUTO_DECOM').AsString := GerarPLU( Layout.FieldByName('COD_PRODUTO_DECOM').AsString );

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;

end;

procedure TFrmHiperLojao.GerarDivisaoForn;
begin
  inherited;
    with QryPrincipal do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
    SQL.Add('    DIVISAO.FORNECEDOR AS COD_FORNECEDOR,');
    SQL.Add('    DIVISAO.ID AS COD_DIVISAO,');
    SQL.Add('    DIVISAO.DESCRITIVO AS DES_DIVISAO,');
    SQL.Add('    FORNECEDORES.LOGRADOURO || '' '' || FORNECEDORES.ENDERECO AS DES_ENDERECO,');
    SQL.Add('    FORNECEDORES.BAIRRO AS DES_BAIRRO,');
    SQL.Add('    FORNECEDORES.CEP AS NUM_CEP,');
    SQL.Add('    FORNECEDORES.CIDADE AS DES_CIDADE,');
    SQL.Add('    FORNECEDORES.ESTADO AS DES_SIGLA,');
    SQL.Add('    FORNECEDORES.TELEFONE1 AS NUM_FONE,');
    SQL.Add('    '''' AS DES_CONTATO,');
    SQL.Add('    FORNECEDORES.EMAIL AS DES_EMAIL,');
    SQL.Add('    FORNECEDORES.OBSERVACAO AS DES_OBSERVACAO');
    SQL.Add('FROM');
    SQL.Add('    FORNECEDORES_LINHAS DIVISAO');
    SQL.Add('LEFT JOIN');
    SQL.Add('    FORNECEDORES');
    SQL.Add('ON');
    SQL.Add('    DIVISAO.FORNECEDOR = FORNECEDORES.ID');

    Open;

    First;
    NumLinha := 0;

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);

      Layout.SetValues(QryPrincipal, NumLinha, RecordCount);

      Layout.FieldByName('DES_OBSERVACAO').AsString := StrReplace(StrLBReplace(FieldByName('DES_OBSERVACAO').AsString), '\n', '');
      Layout.FieldByName('NUM_CEP').AsString := StrRetNums(Layout.FieldByName('NUM_CEP').AsString);

      Layout.FieldByName('DES_EMAIL').AsString := StrReplace(StrLBReplace(FieldByName('DES_EMAIL').AsString), '\n', '');

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;
end;


procedure TFrmHiperLojao.GerarEstDeposito;
var
  COD_PRODUTO : string;
  NovoCodigo : Integer;
begin
   with QryPrincipal do
   begin
     Close;
     SQL.Clear;

     SQL.Add('   SELECT   ');
     SQL.Add('       PRODUTOS_ESTOQUES.PST_LOCAL AS COD_LOJA,   ');
     SQL.Add('       CASE   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''25578,'' THEN CAST(25578 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''28687.'' THEN CAST(28687 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''29397.'' THEN CAST(29397 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''29*504'' THEN CAST(29504 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''296,5'' THEN CAST(2965 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''29878.'' THEN CAST(29878 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''29870,'' THEN CAST(29870 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''2995,'' THEN CAST(2995 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''30023Q'' THEN CAST(30023 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''30045,'' THEN CAST(30045 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''30304,'' THEN CAST(30304 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN CHAR_LENGTH(PRODUTOS.PRO_REFERENCIA) = 9 THEN PRODUTOS.COD_PRODUTO   ');
     SQL.Add('           ELSE PRODUTOS.PRO_REFERENCIA        ');
     SQL.Add('       END AS COD_PRODUTO,   ');
     SQL.Add('       PRODUTOS_ESTOQUES.PST_FISICO AS QTD_EST_DEP   ');
     SQL.Add('   FROM   ');
     SQL.Add('       PRODUTOS   ');
     SQL.Add('   LEFT JOIN PRODUTOS_ESTOQUES ON PRODUTOS_ESTOQUES.PST_PRODUTO = PRODUTOS.PRO_CODIGO   ');
     SQL.Add('   WHERE PRODUTOS_ESTOQUES.PST_LOCAL = 1   ');
     SQL.Add('   AND PRODUTOS_ESTOQUES.PST_FISICO <> 0   ');
     SQL.Add('   ORDER BY(   ');
     SQL.Add('       CASE WHEN PRODUTOS.PRO_LEGENDA IN (''N'',''P'') THEN PRODUTOS.PRO_LEGENDA END   ');
     SQL.Add('   ) DESC,   ');
     SQL.Add('   (   ');
     SQL.Add('       CASE WHEN PRODUTOS.PRO_LEGENDA NOT IN (''N'', ''P'') THEN PRODUTOS.PRO_LEGENDA END   ');
     SQL.Add('   ), PRODUTOS.PRO_CODIGO ASC   ');

     Open;
     First;

     NumLinha := 0;
     NovoCodigo := 400000;

     while not Eof do
     begin
       try
        if Cancelar then
        Break;

        Inc(NumLinha);

//        if QryPrincipal.FieldByName('COD_PRODUTO').AsInteger = 0 then
//        begin
//          Inc(NovoCodigo);
//          COD_PRODUTO := GerarPLU(IntToStr(NovoCodigo));
//          Writeln(Arquivo, 'UPDATE TAB_PRODUTO_LOJA SET QTD_EST_DEP = '+QryPrincipal.FieldByName('QTD_EST_DEP').AsString+' WHERE COD_PRODUTO = '+COD_PRODUTO+' ; ');
//        end;

//        if QryPrincipal.FieldByName('COD_PRODUTO').AsInteger <> 0 then
//        begin
          COD_PRODUTO := GerarPLU(QryPrincipal.FieldByName('COD_PRODUTO').AsString);
          Writeln(Arquivo, 'UPDATE TAB_PRODUTO_LOJA SET QTD_EST_DEP = '+QryPrincipal.FieldByName('QTD_EST_DEP').AsString+' WHERE COD_PRODUTO = '+COD_PRODUTO+' ; ');
//        end;

        if NumLinha = 500 then
          Writeln(Arquivo, 'COMMIT WORK;');
        if NumLinha = 1000 then        
          Writeln(Arquivo, 'COMMIT WORK;');
        if NumLinha = 1500 then        
          Writeln(Arquivo, 'COMMIT WORK;');
        if NumLinha = 2000 then
          Writeln(Arquivo, 'COMMIT WORK;');
        if NumLinha = 2500 then
          Writeln(Arquivo, 'COMMIT WORK;');
        if NumLinha = 3000 then
          Writeln(Arquivo, 'COMMIT WORK;');
        if NumLinha = 3500 then
          Writeln(Arquivo, 'COMMIT WORK;');
        if NumLinha = 4000 then
          Writeln(Arquivo, 'COMMIT WORK;');
        if NumLinha = 4400 then
          Writeln(Arquivo, 'COMMIT WORK;');

          
       except on E: Exception do
        FrmProgresso.AdicionarLog(QryPrincipal.RecNo, 'E', E.Message);
       end;
       Next;
     end;
      Writeln(Arquivo, 'COMMIT WORK;');
      Close;
   end;
end;

procedure TFrmHiperLojao.GerarFinanceiro(Tipo, Situacao: Integer);
begin
  inherited;
  if Tipo = 1 then
    GerarFinanceiroPagar(IntToStr(Situacao));

  if Tipo = 2 then
    GerarFinanceiroReceber(IntToStr(Situacao));

  if Tipo = 3 then
    GerarFinanceiroReceberCartao;

end;

procedure TFrmHiperLojao.GerarFinanceiroPagar(Aberto: String);
var
   TotalCount : Integer;
   cgc: string;
begin
  inherited;
  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

     SQL.Add('   SELECT   ');
     SQL.Add('       1 AS TIPO_PARCEIRO,   ');
     SQL.Add('       PAGAR.PAG_FORNECEDOR AS COD_PARCEIRO,   ');
     SQL.Add('       0 AS TIPO_CONTA,   ');
     SQL.Add('      ');
     SQL.Add('       CASE PAGAR.PAG_TIPO    ');
     SQL.Add('           WHEN 1 THEN 20   ');
     SQL.Add('           WHEN 2 THEN 20   ');
     SQL.Add('           WHEN 3 THEN 20   ');
     SQL.Add('           WHEN 4 THEN 21   ');
     SQL.Add('           WHEN 5 THEN 21   ');
     SQL.Add('           WHEN 6 THEN 22   ');
     SQL.Add('           WHEN 7 THEN 23   ');
     SQL.Add('           WHEN 8 THEN 23   ');
     SQL.Add('           WHEN 9 THEN 20   ');
     SQL.Add('           WHEN 10 THEN 24   ');
     SQL.Add('           WHEN 11 THEN 24   ');
     SQL.Add('           WHEN 12 THEN 25   ');
     SQL.Add('           WHEN 13 THEN 25   ');
     SQL.Add('           WHEN 14 THEN 27   ');
     SQL.Add('           WHEN 15 THEN 28   ');
     SQL.Add('           WHEN 16 THEN 29   ');
     SQL.Add('           WHEN 17 THEN 30   ');
     SQL.Add('           WHEN 18 THEN 31   ');
     SQL.Add('           WHEN 19 THEN 32   ');
     SQL.Add('           WHEN 20 THEN 32   ');
     SQL.Add('           WHEN 21 THEN 33   ');
     SQL.Add('           WHEN 22 THEN 34   ');
     SQL.Add('           WHEN 23 THEN 35   ');
     SQL.Add('           WHEN 24 THEN 36   ');
     SQL.Add('           WHEN 25 THEN 37   ');
     SQL.Add('       END AS COD_ENTIDADE,   ');
     SQL.Add('      ');
     SQL.Add('       PAGAR.PAG_TITULO AS NUM_DOCTO,   ');
     SQL.Add('       999 AS COD_BANCO,   ');
     SQL.Add('       '''' AS DES_BANCO,   ');
     SQL.Add('       PAGAR.PAG_EMISSAO AS DTA_EMISSAO,   ');
     SQL.Add('       PAGAR.PAG_VENCIMENTO AS DTA_VENCIMENTO,   ');
     SQL.Add('       PAGAR.PAG_VRORI AS VAL_PARCELA,   ');
     SQL.Add('       PAGAR.PAG_JUROSDIA AS VAL_JUROS,   ');
     SQL.Add('       0 AS VAL_DESCONTO,   ');
     SQL.Add('       ''N'' AS FLG_QUITADO,   ');
     SQL.Add('       '''' AS DTA_QUITADA,   ');
     SQL.Add('       998 AS COD_CATEGORIA,   ');
     SQL.Add('       998 AS COD_SUBCATEGORIA,   ');
     SQL.Add('       PAGAR.PAG_PARCELA AS NUM_PARCELA,   ');
     SQL.Add('       PAGAR.PAG_QTDE_PARCELA AS QTD_PARCELA,   ');
     SQL.Add('       PAGAR.PAG_EMPRESA AS COD_LOJA,   ');
     SQL.Add('       REPLACE(REPLACE(REPLACE(FORNECEDORES.FOR_CGC, ''.'', ''''), ''-'', ''''), ''/'', '''') AS NUM_CGC,   ');
     SQL.Add('       0 AS NUM_BORDERO,   ');
     SQL.Add('       PAGAR.PAG_NF AS NUM_NF,   ');
     SQL.Add('       '''' AS NUM_SERIE_NF,   ');
     SQL.Add('       PAGAR.PAG_VRTITULO AS VAL_TOTAL_NF,   ');
     SQL.Add('       CAST(COALESCE(PAGAR.PAG_OBS, '''') AS VARCHAR(400)) AS DES_OBSERVACAO,   ');
     SQL.Add('       0 AS NUM_PDV,   ');
     SQL.Add('       0 AS NUM_CUPOM_FISCAL,   ');
     SQL.Add('       0 AS COD_MOTIVO,   ');
     SQL.Add('       0 AS COD_CONVENIO,   ');
     SQL.Add('       0 AS COD_BIN,   ');
     SQL.Add('       '''' AS DES_BANDEIRA,   ');
     SQL.Add('       '''' AS DES_REDE_TEF,   ');
     SQL.Add('       0 AS VAL_RETENCAO,   ');
     SQL.Add('       0 AS COD_CONDICAO,   ');
     SQL.Add('       '''' AS DTA_PAGTO,   ');
     SQL.Add('       '''' AS DTA_ENTRADA,   ');
     SQL.Add('       '''' AS NUM_NOSSO_NUMERO,   ');
     SQL.Add('       '''' AS COD_BARRA,   ');
     SQL.Add('       ''N'' AS FLG_BOLETO_EMIT,   ');
     SQL.Add('       '''' AS NUM_CGC_CPF_TITULAR,   ');
     SQL.Add('       '''' AS DES_TITULAR,   ');
     SQL.Add('       30 AS NUM_CONDICAO,   ');
     SQL.Add('       0 AS VAL_CREDITO,   ');
     SQL.Add('       ''999'' AS COD_BANCO_PGTO,   ');
     SQL.Add('       ''PAGTO-1'' AS DES_CC,   ');
     SQL.Add('       0 AS COD_BANDEIRA,   ');
     SQL.Add('       '''' AS DTA_PRORROGACAO,   ');
     SQL.Add('       1 AS NUM_SEQ_FIN,   ');
     SQL.Add('       0 AS COD_COBRANCA,   ');
     SQL.Add('       '''' AS DTA_COBRANCA,   ');
     SQL.Add('       ''N'' AS FLG_ACEITE,   ');
     SQL.Add('       0 AS TIPO_ACEITE   ');
     SQL.Add('   FROM   ');
     SQL.Add('       FN_PAGAR AS PAGAR   ');
     SQL.Add('   LEFT JOIN FORNECEDORES ON PAGAR.PAG_FORNECEDOR = FORNECEDORES.FOR_CODIGO   ');
     SQL.Add('   WHERE PAGAR.PAG_EMPRESA = '+CbxLoja.Text+'   ');
     SQL.Add('   AND PAGAR.PAG_QUITADO = ''N''   ');
     SQL.Add('   --AND PAGAR.PAG_TITULO = 10224   ');
     SQL.Add('   AND PAGAR.PAG_EMISSAO BETWEEN :INI AND :FIM   ');
     SQL.Add('   ORDER BY COD_PARCEIRO, NUM_DOCTO, DTA_VENCIMENTO   ');

     ParamByName('INI').AsDate := DtpInicial.Date;
     ParamByName('FIM').AsDate := DtpFinal.Date;


    Open;
    First;
//
//    if( Aberto = '1' ) then
//      TotalCount := SetCountTotal(SQL.Text)
//    else
//      TotalCount := SetCountTotal(SQL.Text, ParamByName('INI').AsString, ParamByName('FIM').AsString );
    NumLinha := 0;
    TotalCount := SetCountTotal(SQL.Text, ParamByName('INI').AsString, ParamByName('FIM').AsString );

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);

      Layout.SetValues(QryPrincipal, NumLinha, TotalCount);

        Layout.FieldByName('DES_OBSERVACAO').AsString := StrLBReplace(QryPrincipal.FieldByName('DES_OBSERVACAO').AsString);
        Layout.FieldByName('NUM_NF').AsString := StrRetNums(QryPrincipal.FieldByName('NUM_NF').AsString);
        Layout.FieldByName('NUM_CUPOM_FISCAL').AsString := StrRetNums(QryPrincipal.FieldByName('NUM_CUPOM_FISCAL').AsString);


        if QryPrincipal.FieldByName('DTA_QUITADA').AsString <> '' then
          Layout.FieldByName('DTA_QUITADA').AsString:= FormatDateTime('dd/mm/yyyy',QryPrincipal.FieldByName('DTA_QUITADA').AsDateTime);

        if QryPrincipal.FieldByName('DTA_PAGTO').AsString <> '' then
          Layout.FieldByName('DTA_PAGTO').AsString:= FormatDateTime('dd/mm/yyyy',QryPrincipal.FieldByName('DTA_PAGTO').AsDateTime);

        if QryPrincipal.FieldByName('DTA_EMISSAO').AsString <> '' then
          Layout.FieldByName('DTA_EMISSAO').AsString := FormatDateTime('dd/mm/yyyy',QryPrincipal.FieldByName('DTA_EMISSAO').AsDateTime);

        if QryPrincipal.FieldByName('DTA_VENCIMENTO').AsString <> '' then
          Layout.FieldByName('DTA_VENCIMENTO').AsString:= FormatDateTime('dd/mm/yyyy',QryPrincipal.FieldByName('DTA_VENCIMENTO').AsDateTime);

        if QryPrincipal.FieldByName('DTA_ENTRADA').AsString <> '' then
          Layout.FieldByName('DTA_ENTRADA').AsString:= FormatDateTime('dd/mm/yyyy',QryPrincipal.FieldByName('DTA_ENTRADA').AsDateTime);


      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;
end;

procedure TFrmHiperLojao.GerarFinanceiroReceber(Aberto: String);
var
   TotalCount : Integer;
   cgc: string;
begin
  inherited;
  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

     SQL.Add('   SELECT   ');
     SQL.Add('       0 AS TIPO_PARCEIRO,   ');
     SQL.Add('       RECEBER.REC_CLIENTE AS COD_PARCEIRO,   ');
     SQL.Add('       1 AS TIPO_CONTA,   ');
     SQL.Add('       CASE RECEBER.REC_TIPO   ');
     SQL.Add('           WHEN 14 THEN 1   ');
     SQL.Add('           WHEN 16 THEN 2   ');
     SQL.Add('           WHEN 20 THEN 8   ');
     SQL.Add('           WHEN 22 THEN 2   ');
     SQL.Add('           WHEN 24 THEN 1   ');
     SQL.Add('           WHEN 28 THEN 13   ');
     SQL.Add('           WHEN 30 THEN 14   ');
     SQL.Add('           WHEN 34 THEN 6   ');
     SQL.Add('           WHEN 35 THEN 6   ');
     SQL.Add('           WHEN 36 THEN 7   ');
     SQL.Add('           WHEN 37 THEN 7   ');
     SQL.Add('           WHEN 38 THEN 15   ');
     SQL.Add('           WHEN 39 THEN 15   ');
     SQL.Add('           WHEN 40 THEN 16   ');
     SQL.Add('           WHEN 41 THEN 16   ');
     SQL.Add('           WHEN 43 THEN 17   ');
     SQL.Add('           WHEN 44 THEN 17   ');
     SQL.Add('           WHEN 46 THEN 1   ');
     SQL.Add('           WHEN 55 THEN 18   ');
     SQL.Add('           WHEN 56 THEN 18   ');
     SQL.Add('           WHEN 59 THEN 1   ');
     SQL.Add('           WHEN 61 THEN 1   ');
     SQL.Add('           WHEN 63 THEN 7   ');
     SQL.Add('           WHEN 64 THEN 7   ');
     SQL.Add('           WHEN 65 THEN 6   ');
     SQL.Add('           WHEN 66 THEN 6   ');
     SQL.Add('           WHEN 67 THEN 15   ');
     SQL.Add('           WHEN 68 THEN 15   ');
     SQL.Add('           WHEN 75 THEN 19   ');
     SQL.Add('           WHEN 76 THEN 19   ');
     SQL.Add('       END AS COD_ENTIDADE,   ');
     SQL.Add('       RECEBER.REC_TITULO AS NUM_DOCTO,   ');
     SQL.Add('       999 AS COD_BANCO,   ');
     SQL.Add('       '''' AS DES_BANCO,   ');
     SQL.Add('       RECEBER.REC_EMISSAO AS DTA_EMISSAO,   ');
     SQL.Add('       RECEBER.REC_VENCIMENTO AS DTA_VENCIMENTO,   ');
     SQL.Add('       RECEBER.REC_VRORI AS VAL_PARCELA,   ');
     SQL.Add('       COALESCE(RECEBER.REC_JUROSDIA, 0) AS VAL_JUROS,   ');
     SQL.Add('       0 AS VAL_DESCONTO,   ');
     SQL.Add('       ''N'' AS FLG_QUITADO,   ');
     SQL.Add('       '''' AS DTA_QUITADA,   ');
     SQL.Add('       998 AS COD_CATEGORIA,   ');
     SQL.Add('       998 AS COD_SUBCATEGORIA,   ');
     SQL.Add('       RECEBER.REC_PARCELA AS NUM_PARCELA,   ');
     SQL.Add('       QUANTIDADE.QTD_PARCELA AS QTD_PARCELA,   ');
     SQL.Add('       RECEBER.REC_EMPRESA AS COD_LOJA,   ');
     SQL.Add('       COALESCE(REPLACE(REPLACE(REPLACE(CLIENTES.CLI_CGC, ''.'', ''''), ''-'', ''''), ''/'', ''''), '''') AS NUM_CGC,   ');
     SQL.Add('       0 AS NUM_BORDERO,   ');
     SQL.Add('       RECEBER.REC_NF AS NUM_NF,   ');
     SQL.Add('       RECEBER.REC_SERIE AS NUM_SERIE_NF,   ');
     SQL.Add('       SOMA.VAL_TOTAL AS VAL_TOTAL_NF,   ');
     SQL.Add('       CAST(COALESCE(RECEBER.REC_OBS, '''') AS VARCHAR(300)) AS DES_OBSERVACAO,   ');
     SQL.Add('       COALESCE(RECEBER.REC_TERMINAL, 1) AS NUM_PDV,   ');
     SQL.Add('       0 AS NUM_CUPOM_FISCAL,   ');
     SQL.Add('       0 AS COD_MOTIVO,   ');
     SQL.Add('       0 AS COD_CONVENIO,   ');
     SQL.Add('       0 AS COD_BIN,   ');
     SQL.Add('       '''' AS DES_BANDEIRA,   ');
     SQL.Add('       '''' AS DES_REDE_TEF,   ');
     SQL.Add('       0 AS VAL_RETENCAO,   ');
     SQL.Add('       0 AS COD_CONDICAO,   ');
     SQL.Add('       RECEBER.REC_PAGAMENTO AS DTA_PAGTO,   ');
     SQL.Add('       '''' AS DTA_ENTRADA,   ');
     SQL.Add('       '''' AS NUM_NOSSO_NUMERO,   ');
     SQL.Add('       '''' AS COD_BARRA,   ');
     SQL.Add('       ''N'' AS FLG_BOLETO_EMIT,   ');
     SQL.Add('       '''' AS NUM_CGC_CPF_TITULAR,   ');
     SQL.Add('       '''' AS DES_TITULAR,   ');
     SQL.Add('       30 AS NUM_CONDICAO,   ');
     SQL.Add('       0 AS VAL_CREDITO,   ');
     SQL.Add('       ''999'' AS COD_BANCO_PGTO,   ');
     SQL.Add('       ''RECEBTO-1'' AS DES_CC,   ');
     SQL.Add('       0 AS COD_BANDEIRA,   ');
     SQL.Add('       '''' AS DTA_PRORROGACAO,   ');
     SQL.Add('       1 AS NUM_SEQ_FIN,   ');
     SQL.Add('       0 AS COD_COBRANCA,   ');
     SQL.Add('       '''' AS DTA_COBRANCA,   ');
     SQL.Add('       ''N'' AS FLG_ACEITE,   ');
     SQL.Add('       0 AS TIPO_ACEITE   ');
     SQL.Add('   FROM   ');
     SQL.Add('       RECEBER   ');
     SQL.Add('   LEFT JOIN CLIENTES ON RECEBER.REC_CLIENTE = CLIENTES.CLI_CODIGO   ');
     SQL.Add('   LEFT JOIN (   ');
     SQL.Add('       SELECT   ');
     SQL.Add('           REC_CLIENTE,   ');
     SQL.Add('           REC_TITULO,   ');
     SQL.Add('           COUNT(REC_TITULO) AS QTD_PARCELA   ');
     SQL.Add('       FROM   ');
     SQL.Add('           RECEBER   ');
     SQL.Add('       -- WHERE REC_QUITADO = ''N''   ');
     SQL.Add('       GROUP BY   ');
     SQL.Add('           REC_CLIENTE, REC_TITULO   ');
     SQL.Add('   ) AS QUANTIDADE   ');
     SQL.Add('   ON   ');
     SQL.Add('       RECEBER.REC_CLIENTE = QUANTIDADE.REC_CLIENTE   ');
     SQL.Add('   AND   ');
     SQL.Add('       RECEBER.REC_TITULO = QUANTIDADE.REC_TITULO   ');
     SQL.Add('   LEFT JOIN (   ');
     SQL.Add('       SELECT   ');
     SQL.Add('           REC_CLIENTE,   ');
     SQL.Add('           REC_TITULO,   ');
     SQL.Add('           SUM(REC_VRORI) AS VAL_TOTAL   ');
     SQL.Add('       FROM   ');
     SQL.Add('           RECEBER   ');
     SQL.Add('       -- WHERE REC_QUITADO = ''N''   ');
     SQL.Add('       GROUP BY   ');
     SQL.Add('           REC_CLIENTE, REC_TITULO   ');
     SQL.Add('   ) AS SOMA   ');
     SQL.Add('   ON   ');
     SQL.Add('       RECEBER.REC_CLIENTE = SOMA.REC_CLIENTE   ');
     SQL.Add('   AND   ');
     SQL.Add('       RECEBER.REC_TITULO = SOMA.REC_TITULO   ');
     SQL.Add('   WHERE RECEBER.REC_QUITADO = ''N''   ');
     SQL.Add('   AND RECEBER.REC_EMPRESA = '+CbxLoja.Text+'   ');
     SQL.Add('AND');
     SQL.Add(' RECEBER.REC_EMISSAO BETWEEN :INI AND :FIM ');
     SQL.Add('   ORDER BY COD_PARCEIRO, NUM_DOCTO, DTA_VENCIMENTO   ');
  //end;

    ParamByName('INI').AsDate := DtpInicial.Date;
    ParamByName('FIM').AsDate := DtpFinal.Date;

    Open;

    First;
    NumLinha := 0;
//    codParceiro := 0;
//    numDocto := '';
//    count := 0;
    TotalCount := SetCountTotal(SQL.Text, ParamByName('INI').AsString, ParamByName('FIM').AsString );


    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);

      Layout.SetValues(QryPrincipal, NumLinha, TotalCount);

//      if( CbxLoja.Text = '2' ) then
//      begin
//         cgc := StrRetNums(Layout.FieldByName('NUM_CGC').AsString);
//         if( Length(cgc) > 11 ) then begin
//           if( not CNPJEValido(cgc) ) then
//            Layout.FieldByName('COD_PARCEIRO').AsInteger := Layout.FieldByName('COD_PARCEIRO').AsInteger + 2000
//           else
//            Layout.FieldByName('COD_PARCEIRO').AsInteger := 0;
//         end
//         else
//         begin
//            if( not CPFEValido(cgc) ) then
//               Layout.FieldByName('COD_PARCEIRO').AsInteger := Layout.FieldByName('COD_PARCEIRO').AsInteger + 2000
//            else
//               Layout.FieldByName('COD_PARCEIRO').AsInteger := 0;
//         end;
//      end;
      Layout.FieldByName('DES_OBSERVACAO').AsString := StrLBReplace(QryPrincipal.FieldByName('DES_OBSERVACAO').AsString);
      Layout.FieldByName('NUM_NF').AsString := StrRetNums(QryPrincipal.FieldByName('NUM_NF').AsString);
      Layout.FieldByName('NUM_CUPOM_FISCAL').AsString := StrRetNums(QryPrincipal.FieldByName('NUM_CUPOM_FISCAL').AsString);


        if QryPrincipal.FieldByName('DTA_QUITADA').AsString <> '' then
          Layout.FieldByName('DTA_QUITADA').AsString:= FormatDateTime('dd/mm/yyyy',QryPrincipal.FieldByName('DTA_QUITADA').AsDateTime);

        if QryPrincipal.FieldByName('DTA_PAGTO').AsString <> '' then
          Layout.FieldByName('DTA_PAGTO').AsString:= FormatDateTime('dd/mm/yyyy',QryPrincipal.FieldByName('DTA_PAGTO').AsDateTime);

        if QryPrincipal.FieldByName('DTA_EMISSAO').AsString <> '' then
          Layout.FieldByName('DTA_EMISSAO').AsString := FormatDateTime('dd/mm/yyyy',QryPrincipal.FieldByName('DTA_EMISSAO').AsDateTime);

        if QryPrincipal.FieldByName('DTA_VENCIMENTO').AsString <> '' then
          Layout.FieldByName('DTA_VENCIMENTO').AsString:= FormatDateTime('dd/mm/yyyy',QryPrincipal.FieldByName('DTA_VENCIMENTO').AsDateTime);

        if QryPrincipal.FieldByName('DTA_ENTRADA').AsString <> '' then
          Layout.FieldByName('DTA_ENTRADA').AsString:= FormatDateTime('dd/mm/yyyy',QryPrincipal.FieldByName('DTA_ENTRADA').AsDateTime);

//      if Aberto = '1' then
//      begin
//        Layout.FieldByName('DTA_QUITADA').AsString := '';
//        Layout.FieldByName('DTA_PAGTO').AsString := '';
//      end
//      else
//      begin
//        Layout.FieldByName('DTA_QUITADA').AsString := FormatDateTime('dd/mm/yyyy',QryPrincipal.FieldByName('DTA_QUITADA').AsDateTime);
//        Layout.FieldByName('DTA_PAGTO').AsString := FormatDateTime('dd/mm/yyyy',QryPrincipal.FieldByName('DTA_PAGTO').AsDateTime);
//      end;

      //Layout.FieldByName('DES_OBSERVACAO').AsString := StrReplace(StrLBReplace(FieldByName('DES_OBSERVACAO').AsString), '\n', '');

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;
end;

procedure TFrmHiperLojao.GerarFinanceiroReceberCartao;
begin
  inherited;
  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

//    SQL.Add('SELECT');
//    SQL.Add('');
//    SQL.Add('    CASE RECEBER.TIPO_CADASTRO');
//    SQL.Add('        WHEN 0 THEN 0');
//    SQL.Add('        WHEN 1 THEN 3');
//    SQL.Add('        WHEN 4 THEN 4');
//    SQL.Add('        WHEN 5 THEN 0');
//    SQL.Add('    END AS TIPO_PARCEIRO, -- TIPO_CADASTRO');
//    SQL.Add('');
//    SQL.Add('     CASE');
//    SQL.Add('          WHEN RECEBER.TIPO_CADASTRO = 5 THEN 2400 + RECEBER.ID_CADASTRO ');
//    SQL.Add('          WHEN RECEBER.TIPO_CADASTRO = 5 AND COALESCE(RECEBER.ID_CADASTRO, 0) = 0 THEN 6');
//    SQL.Add('          WHEN RECEBER.TIPO_CADASTRO = 4 THEN 99');
//    SQL.Add('          ELSE CASE WHEN COALESCE(RECEBER.ID_CADASTRO, 0) = 0 THEN 99999 ELSE RECEBER.ID_CADASTRO END');
//    SQL.Add('     END AS COD_PARCEIRO,  ');
//    SQL.Add('');
//    SQL.Add('    1 AS TIPO_CONTA,');
//    SQL.Add('');
//    SQL.Add('    CASE RECEBER.FORMA_PAGTO');
//    SQL.Add('        WHEN 1 THEN 1');
//    SQL.Add('        WHEN 2 THEN 2');
//    SQL.Add('        WHEN 3 THEN 4');
//    SQL.Add('        WHEN 4 THEN 10');
//    SQL.Add('        WHEN 5 THEN 11');
//    SQL.Add('        WHEN 6 THEN 6');
//    SQL.Add('        WHEN 7 THEN 12');
//    SQL.Add('        WHEN 8 THEN 3');
//    SQL.Add('        WHEN 9 THEN 13');
//    SQL.Add('        WHEN 10 THEN 5');
//    SQL.Add('        WHEN 11 THEN 7');
//    SQL.Add('        WHEN 12 THEN 14');
//    SQL.Add('        WHEN 13 THEN 15');
//    SQL.Add('        WHEN 14 THEN 16');
//    SQL.Add('        WHEN 15 THEN 17');
//    SQL.Add('        WHEN 16 THEN 18');
//    SQL.Add('        WHEN 17 THEN 19');
//    SQL.Add('        WHEN 18 THEN 20');
//    SQL.Add('        WHEN 19 THEN 21');
//    SQL.Add('        WHEN 20 THEN 22');
//    SQL.Add('        WHEN 21 THEN 23');
//    SQL.Add('        WHEN 22 THEN 24');
//    SQL.Add('        WHEN 23 THEN 25');
//    SQL.Add('        WHEN 24 THEN 26');
//    SQL.Add('        WHEN 25 THEN 27');
//    SQL.Add('        ELSE 1');
//    SQL.Add('    END AS COD_ENTIDADE,');
//    SQL.Add('');
//    SQL.Add('    RECEBER.ARQUIVO AS NUM_DOCTO,');
//    SQL.Add('    999 AS COD_BANCO,');
//    SQL.Add('    '''' AS DES_BANCO,');
//    SQL.Add('    RECEBER.EMISSAO AS DTA_EMISSAO,');
//    SQL.Add('    RECEBER.VENCIMENTO AS DTA_VENCIMENTO,');
//    SQL.Add('    RECEBER.VALOR AS VAL_PARCELA,');
//    SQL.Add('    RECEBER.ACRESCIMO + RECEBER.CARTORIO + COALESCE(RECEBER.CREDITO, 0) AS VAL_JUROS,');
//    SQL.Add('    RECEBER.DESCONTO AS VAL_DESCONTO,');
//    SQL.Add('');
//    SQL.Add('    CASE ');
//    SQL.Add('        WHEN RECEBER.PAGAMENTO IS NULL THEN ''N''');
//    SQL.Add('        ELSE ''S''');
//    SQL.Add('    END AS FLG_QUITADO,');
//    SQL.Add('');
//    SQL.Add('    CASE ');
//    SQL.Add('        WHEN RECEBER.PAGAMENTO IS NULL THEN NULL');
//    SQL.Add('        ELSE RECEBER.PAGAMENTO');
//    SQL.Add('    END AS DTA_QUITADA,');
//    SQL.Add('');
//    SQL.Add('    ');
//    SQL.Add('    CASE RECEBER.CAIXA');
//    SQL.Add('        WHEN 2 THEN ''001''');
//    SQL.Add('        ELSE ''997''');
//    SQL.Add('    END AS COD_CATEGORIA,');
//    SQL.Add('');
//    SQL.Add('    CASE RECEBER.CAIXA');
//    SQL.Add('        WHEN 2 THEN ''032''');
//    SQL.Add('        ELSE ''997''');
//    SQL.Add('    END AS COD_SUBCATEGORIA,');
//    SQL.Add('');
//    SQL.Add('    RECEBER.PARCELA AS NUM_PARCELA,');
//    SQL.Add('    RECEBER.TOTAL_PARCELA AS QTD_PARCELA,');
//    SQL.Add('    RECEBER.EMPRESA AS COD_LOJA,');
//    SQL.Add('    RECEBER.CPF_CNPJ AS NUM_CGC,');
//    SQL.Add('    COALESCE(RECEBER.BORDERO, 0) AS NUM_BORDERO,');
//    SQL.Add('    RECEBER.NF AS NUM_NF,');
//    SQL.Add('    '''' AS NUM_SERIE_NF,');
//    SQL.Add('    CASE WHEN NF.VAL_TOTAL_NF = 0 THEN RECEBER.VALOR ELSE NF.VAL_TOTAL_NF END AS VAL_TOTAL_NF, -- EFETUAR A SOMA');
//    SQL.Add('    ''COBRAN�A: '' || RECEBER.DATACOB || '' | 1 DEVOL: '' || RECEBER.DEVOLUCAOA || '' | 2 DEVOL : '' || RECEBER.DEVOLUCAOB || '' | ''  || RECEBER.OBSERVACAO AS DES_OBSERVACAO,');
//    SQL.Add('    COALESCE(RECEBER.PDV, 0) AS NUM_PDV,');
//    SQL.Add('    RECEBER.NOTA AS NUM_CUPOM_FISCAL,');
//    SQL.Add('    0 AS COD_MOTIVO,');
//    SQL.Add('');
//    SQL.Add('    CASE RECEBER.FORMA_PAGTO');
//    SQL.Add('        WHEN 14 THEN (SELECT COALESCE(24000 + CLIENTES.EMPRESA_CONVENIO, 0) FROM CLIENTES WHERE CLIENTES.ID = RECEBER.ID_CADASTRO)');
//    SQL.Add('        ELSE 0');
//    SQL.Add('    END AS COD_CONVENIO,');
//    SQL.Add('');
//    SQL.Add('    0 AS COD_BIN,');
//    SQL.Add('    '''' AS DES_BANDEIRA,');
//    SQL.Add('    '''' AS DES_REDE_TEF,');
//    SQL.Add('    0 AS VAL_RETENCAO,');
//    SQL.Add('    0 AS COD_CONDICAO,');
//    SQL.Add('');
//    SQL.Add('    CASE ');
//    SQL.Add('        WHEN RECEBER.PAGAMENTO IS NULL THEN NULL');
//    SQL.Add('        ELSE RECEBER.PAGAMENTO');
//    SQL.Add('    END AS DTA_PAGTO,');
//    SQL.Add('');
//    SQL.Add('    RECEBER.DATAHORA_CADASTRO AS DTA_ENTRADA,');
//    SQL.Add('');
//    SQL.Add('    '''' AS NUM_NOSSO_NUMERO,');
//    SQL.Add('    COALESCE(RECEBER.CODBARRAS, '''') AS COD_BARRA,');
//    SQL.Add('    ''N'' AS FLG_BOLETO_EMIT,');
//    SQL.Add('    '''' AS NUM_CGC_CPF_TITULAR,');
//    SQL.Add('    '''' AS DES_TITULAR,');
//    SQL.Add('    CASE RECEBER.FORMA_PAGTO');
//    SQL.Add('        WHEN 11 THEN 0');
//    SQL.Add('        ELSE 30');
//    SQL.Add('    END AS NUM_CONDICAO,');
//    SQL.Add('    0 AS VAL_CREDITO,');
//    SQL.Add('    ''999'' AS COD_BANCO_PGTO,');
//    SQL.Add('    ''RECEBTO-1'' AS DES_CC,');
//
//    SQL.Add('    CASE ');
//    SQL.Add('        WHEN RECEBER.TIPO_CADASTRO = 4 THEN CASE WHEN RECEBER.EMPRESA = 1 THEN 9999 ELSE 999 END');
//    SQL.Add('        ELSE 0');
//    SQL.Add('        END AS COD_BANDEIRA,');
//
//
//    SQL.Add('    '''' AS DTA_PRORROGACAO,');
//    SQL.Add('    1 AS NUM_SEQ_FIN,');
//    SQL.Add('    CASE RECEBER.COBRADOR');
//    SQL.Add('        WHEN 1 THEN 3405');
//    SQL.Add('        WHEN 2 THEN 3403');
//    SQL.Add('        WHEN 3 THEN 3404');
//    SQL.Add('        ELSE 0');
//    SQL.Add('    END AS COD_COBRANCA,');
//    SQL.Add('    RECEBER.DATACOB AS DTA_COBRANCA,');
//    SQL.Add('    CASE');
//    SQL.Add('        WHEN LENGTH(RECEBER.CODBARRAS) > 0 THEN ''S''');
//    SQL.Add('        ELSE ''N''');
//    SQL.Add('    END AS FLG_ACEITE,');
//    SQL.Add('    CASE');
//    SQL.Add('        WHEN LENGTH(RECEBER.CODBARRAS) = 34 THEN 4 ');
//    SQL.Add('        WHEN LENGTH(RECEBER.CODBARRAS) > 34 THEN 1 ');
//    SQL.Add('        ELSE 0');
//    SQL.Add('    END AS TIPO_ACEITE');
//
//    SQL.Add('FROM');
//    SQL.Add('    CONTAS RECEBER');

    SQL.Add('SELECT');
    SQL.Add('');
    SQL.Add('CASE RECEBER.TIPO_CADASTRO');
    SQL.Add('    WHEN 0 THEN 0');
    SQL.Add('    WHEN 1 THEN 3');
    SQL.Add('    WHEN 4 THEN 4');
    SQL.Add('    WHEN 5 THEN 0');
    SQL.Add('END AS TIPO_PARCEIRO, -- TIPO_CADASTRO');
    SQL.Add('');
    SQL.Add('    CASE');
    SQL.Add('        WHEN RECEBER.TIPO_CADASTRO = 5 THEN 2400 + RECEBER.ID_CADASTRO ');
    SQL.Add('        WHEN RECEBER.TIPO_CADASTRO = 5 AND COALESCE(RECEBER.ID_CADASTRO, 0) = 0 THEN 6');
    SQL.Add('        WHEN RECEBER.TIPO_CADASTRO = 4 THEN 99');
    SQL.Add('        ELSE CASE WHEN COALESCE(RECEBER.ID_CADASTRO, 0) = 0 THEN 99999 ELSE RECEBER.ID_CADASTRO END');
    SQL.Add('    END AS COD_PARCEIRO,  ');
    SQL.Add('');
    SQL.Add('1 AS TIPO_CONTA,');
    SQL.Add('');
    SQL.Add('CASE RECEBER.FORMA_PAGTO');
    SQL.Add('    WHEN 1 THEN 1');
    SQL.Add('    WHEN 2 THEN 2');
    SQL.Add('    WHEN 3 THEN 4');
    SQL.Add('    WHEN 4 THEN 10');
    SQL.Add('    WHEN 5 THEN 11');
    SQL.Add('    WHEN 6 THEN 6');
    SQL.Add('    WHEN 7 THEN 12');
    SQL.Add('    WHEN 8 THEN 3');
    SQL.Add('    WHEN 9 THEN 13');
    SQL.Add('    WHEN 10 THEN 5');
    SQL.Add('    WHEN 11 THEN 7');
    SQL.Add('    WHEN 12 THEN 14');
    SQL.Add('    WHEN 13 THEN 15');
    SQL.Add('    WHEN 14 THEN 16');
    SQL.Add('    WHEN 15 THEN 17');
    SQL.Add('    WHEN 16 THEN 18');
    SQL.Add('    WHEN 17 THEN 19');
    SQL.Add('    WHEN 18 THEN 20');
    SQL.Add('    WHEN 19 THEN 21');
    SQL.Add('    WHEN 20 THEN 22');
    SQL.Add('    WHEN 21 THEN 23');
    SQL.Add('    WHEN 22 THEN 24');
    SQL.Add('    WHEN 23 THEN 25');
    SQL.Add('    WHEN 24 THEN 26');
    SQL.Add('    WHEN 25 THEN 27');
    SQL.Add('    ELSE 1');
    SQL.Add('END AS COD_ENTIDADE,');
    SQL.Add('');
    SQL.Add('RECEBER.ARQUIVO AS NUM_DOCTO,');
    SQL.Add('999 AS COD_BANCO,');
    SQL.Add(''''' AS DES_BANCO,');
    SQL.Add('RECEBER.EMISSAO AS DTA_EMISSAO,');
    SQL.Add('RECEBER.VENCIMENTO AS DTA_VENCIMENTO,');
    SQL.Add('RECEBER.VALOR AS VAL_PARCELA,');
    SQL.Add('RECEBER.ACRESCIMO + RECEBER.CARTORIO + COALESCE(RECEBER.CREDITO, 0) AS VAL_JUROS,');
    SQL.Add('RECEBER.DESCONTO AS VAL_DESCONTO,');
    SQL.Add('');
    SQL.Add('CASE ');
    SQL.Add('    WHEN RECEBER.PAGAMENTO IS NULL THEN ''N''');
    SQL.Add('    ELSE ''S''');
    SQL.Add('END AS FLG_QUITADO,');
    SQL.Add('');
    SQL.Add('CASE ');
    SQL.Add('    WHEN RECEBER.PAGAMENTO IS NULL THEN NULL');
    SQL.Add('    ELSE RECEBER.PAGAMENTO');
    SQL.Add('END AS DTA_QUITADA,');
    SQL.Add('');
    SQL.Add('');
    SQL.Add('CASE RECEBER.CAIXA');
    SQL.Add('    WHEN 2 THEN ''001''');
    SQL.Add('    ELSE ''997''');
    SQL.Add('END AS COD_CATEGORIA,');
    SQL.Add('');
    SQL.Add('CASE RECEBER.CAIXA');
    SQL.Add('    WHEN 2 THEN ''032''');
    SQL.Add('    ELSE ''997''');
    SQL.Add('END AS COD_SUBCATEGORIA,');
    SQL.Add('');
    SQL.Add('RECEBER.PARCELA AS NUM_PARCELA,');
    SQL.Add('RECEBER.TOTAL_PARCELA AS QTD_PARCELA,');
    SQL.Add('RECEBER.EMPRESA AS COD_LOJA,');
    SQL.Add('RECEBER.CPF_CNPJ AS NUM_CGC,');
    SQL.Add('COALESCE(RECEBER.BORDERO, 0) AS NUM_BORDERO,');
    SQL.Add('RECEBER.NF AS NUM_NF,');
    SQL.Add(''''' AS NUM_SERIE_NF,');
    SQL.Add('CASE WHEN NF.VAL_TOTAL_NF = 0 THEN RECEBER.VALOR ELSE NF.VAL_TOTAL_NF END AS VAL_TOTAL_NF, -- EFETUAR A SOMA');
    SQL.Add('''COBRAN�A: '' || RECEBER.DATACOB || '' | 1 DEVOL: '' || RECEBER.DEVOLUCAOA || '' | 2 DEVOL : '' || RECEBER.DEVOLUCAOB || '' | ''  || RECEBER.OBSERVACAO AS DES_OBSERVACAO,');
    SQL.Add('COALESCE(RECEBER.PDV, 0) AS NUM_PDV,');
    SQL.Add('RECEBER.NOTA AS NUM_CUPOM_FISCAL,');
    SQL.Add('0 AS COD_MOTIVO,');
    SQL.Add('');
    SQL.Add('CASE RECEBER.FORMA_PAGTO');
    SQL.Add('    WHEN 14 THEN (SELECT COALESCE(24000 + CLIENTES.EMPRESA_CONVENIO, 0) FROM CLIENTES WHERE CLIENTES.ID = RECEBER.ID_CADASTRO)');
    SQL.Add('    ELSE 0');
    SQL.Add('END AS COD_CONVENIO,');
    SQL.Add('');
    SQL.Add('0 AS COD_BIN,');
//    SQL.Add('ADM_CARTOES.DESCRITIVO AS DES_BANDEIRA,');
    SQL.Add(' '''' AS DES_BANDEIRA,');
    SQL.Add(''''' AS DES_REDE_TEF,');
    SQL.Add('0 AS VAL_RETENCAO,');
    SQL.Add('0 AS COD_CONDICAO,');
    SQL.Add('');
    SQL.Add('CASE ');
    SQL.Add('    WHEN RECEBER.PAGAMENTO IS NULL THEN NULL');
    SQL.Add('    ELSE RECEBER.PAGAMENTO');
    SQL.Add('END AS DTA_PAGTO,');
    SQL.Add('');
    SQL.Add('RECEBER.DATAHORA_CADASTRO AS DTA_ENTRADA,');
    SQL.Add('');
    SQL.Add(''''' AS NUM_NOSSO_NUMERO,');
    SQL.Add('COALESCE(RECEBER.CODBARRAS, '''') AS COD_BARRA,');
    SQL.Add('''N'' AS FLG_BOLETO_EMIT,');
    SQL.Add(''''' AS NUM_CGC_CPF_TITULAR,');
    SQL.Add(''''' AS DES_TITULAR,');
    SQL.Add('CASE RECEBER.FORMA_PAGTO');
    SQL.Add('    WHEN 11 THEN 0');
    SQL.Add('    ELSE 30');
    SQL.Add('END AS NUM_CONDICAO,');
    SQL.Add('0 AS VAL_CREDITO,');
    SQL.Add('''999'' AS COD_BANCO_PGTO,');
    SQL.Add('''RECEBTO-1'' AS DES_CC,');
    SQL.Add('');
    SQL.Add(' 10000 + RECEBER.ID_CADASTRO AS COD_BANDEIRA,');
    SQL.Add('');
    SQL.Add('');
    SQL.Add(''''' AS DTA_PRORROGACAO,');
    SQL.Add('1 AS NUM_SEQ_FIN,');
    SQL.Add('CASE RECEBER.COBRADOR');
    SQL.Add('    WHEN 1 THEN 3405');
    SQL.Add('    WHEN 2 THEN 3403');
    SQL.Add('    WHEN 3 THEN 3404');
    SQL.Add('    ELSE 0');
    SQL.Add('END AS COD_COBRANCA,');
    SQL.Add('RECEBER.DATACOB AS DTA_COBRANCA,');
    SQL.Add('CASE');
    SQL.Add('    WHEN LENGTH(RECEBER.CODBARRAS) > 0 THEN ''S''');
    SQL.Add('    ELSE ''N''');
    SQL.Add('END AS FLG_ACEITE,');
    SQL.Add('CASE');
    SQL.Add('    WHEN LENGTH(RECEBER.CODBARRAS) = 34 THEN 4 ');
    SQL.Add('    WHEN LENGTH(RECEBER.CODBARRAS) > 34 THEN 1 ');
    SQL.Add('    ELSE 0');
    SQL.Add('END AS TIPO_ACEITE');
    SQL.Add('');
    SQL.Add('FROM');
    SQL.Add('CONTAS RECEBER');
    SQL.Add('LEFT JOIN');
    SQL.Add('ADM_CARTOES  ');
    SQL.Add('ON');
    SQL.Add('RECEBER.ID_CADASTRO = ADM_CARTOES.ID');
    SQL.Add('LEFT JOIN');
    SQL.Add('    (');
    SQL.Add('        SELECT ');
    SQL.Add('            NF,');
    SQL.Add('            TIPO_CADASTRO,');
    SQL.Add('            ID_CADASTRO,');
    SQL.Add('            SUM(VALOR - DESCONTO + ACRESCIMO + CARTORIO + COALESCE(CREDITO, 0)) AS VAL_TOTAL_NF');
    SQL.Add('        FROM ');
    SQL.Add('            CONTAS  ');
    SQL.Add('        WHERE');
    SQL.Add('            CONTAS.TIPO_CONTA = 1');
    SQL.Add('        AND');
    SQL.Add('            CONTAS.EMPRESA = '+ CbxLoja.Text +'');
    SQL.Add('        AND');
    SQL.Add('            CONTAS.TIPO_CADASTRO IN (4) -- Adicionar o filtro de cartoes');
    SQL.Add('        AND');
    SQL.Add('            CONTAS.PARCELA > 0');
    SQL.Add('        AND');
    SQL.Add('            CONTAS.VALOR > 0');
    SQL.Add('        GROUP BY');
    SQL.Add('            NF,');
    SQL.Add('            TIPO_CADASTRO,');
    SQL.Add('            ID_CADASTRO');
    SQL.Add('    ) NF');
    SQL.Add('ON');
    SQL.Add('    RECEBER.NF = NF.NF');
    SQL.Add('AND');
    SQL.Add('    RECEBER.TIPO_CADASTRO = NF.TIPO_CADASTRO');
    SQL.Add('AND');
    SQL.Add('    RECEBER.ID_CADASTRO = NF.ID_CADASTRO        ');
    SQL.Add('WHERE');
    SQL.Add('    RECEBER.TIPO_CONTA = 1');
    SQL.Add('AND');
    SQL.Add('    RECEBER.TIPO_CADASTRO IN (4) -- Adicionar o filtro de cartoes');
    SQL.Add('AND');
    SQL.Add('    RECEBER.PARCELA > 0');

    SQL.Add('AND');
    SQL.Add('    RECEBER.VALOR > 0');


    SQL.Add('AND');
    SQL.Add('    RECEBER.EMPRESA = '+ CbxLoja.Text +' ');

    SQL.Add('AND');
    SQL.Add('    RECEBER.EMISSAO >= '''+FormatDateTime('dd/mm/yyyy',DtpInicial.Date)+''' ');
    SQL.Add('AND');
    SQL.Add('    RECEBER.EMISSAO <= '''+FormatDateTime('dd/mm/yyyy',DtpFinal.Date)+''' ');

    SQL.Add('ORDER BY');
    SQL.Add('    NUM_DOCTO, COD_PARCEIRO');

    Open;

    First;
    NumLinha := 0;
//    codParceiro := 0;
//    numDocto := '';
//    count := 0;

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);

      Layout.SetValues(QryPrincipal, NumLinha, RecordCount);

//      if( (codParceiro = QryPrincipal.FieldByName('COD_PARCEIRO').AsInteger) and (numDocto = QryPrincipal.FieldByName('NUM_DOCTO').AsString) ) then
//      begin
//         inc(count);
//         if( numDocto <> '' ) then
//            Layout.FieldByName('NUM_DOCTO').AsString := numDocto + ' - ' + IntToStr(count)
//         else
//            Layout.FieldByName('NUM_DOCTO').AsString := IntToStr(count);
//      end
//      else
//      begin
//         count := 0;
//         numDocto := QryPrincipal.FieldByName('NUM_DOCTO').AsString;
//         codParceiro := QryPrincipal.FieldByName('COD_PARCEIRO').AsInteger;
//      end;

      Layout.FieldByName('DTA_ENTRADA').AsString:= FormatDateTime('dd/mm/yyyy',QryPrincipal.FieldByName('DTA_ENTRADA').AsDateTime);
      Layout.FieldByName('DTA_EMISSAO').AsString:= FormatDateTime('dd/mm/yyyy',QryPrincipal.FieldByName('DTA_EMISSAO').AsDateTime);
      Layout.FieldByName('DTA_VENCIMENTO').AsString:= FormatDateTime('dd/mm/yyyy',QryPrincipal.FieldByName('DTA_VENCIMENTO').AsDateTime);

//      if Aberto = '1' then
//      begin
//        Layout.FieldByName('DTA_QUITADA').AsString := '';
//        Layout.FieldByName('DTA_PAGTO').AsString := '';
//      end
//      else
//      begin
        Layout.FieldByName('DTA_QUITADA').AsString:= FormatDateTime('dd/mm/yyyy',QryPrincipal.FieldByName('DTA_QUITADA').AsDateTime);
        Layout.FieldByName('DTA_PAGTO').AsString:= FormatDateTime('dd/mm/yyyy',QryPrincipal.FieldByName('DTA_PAGTO').AsDateTime);
//      end;

      Layout.FieldByName('DTA_COBRANCA').AsDateTime:= QryPrincipal.FieldByName('DTA_COBRANCA').AsDateTime;

      Layout.FieldByName('DES_OBSERVACAO').AsString := StrReplace(StrLBReplace(FieldByName('DES_OBSERVACAO').AsString), '\n', '');

      Layout.FieldByName('COD_BARRA').AsString := StrRetNums(Layout.FieldByName('COD_BARRA').AsString);

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;
end;

procedure TFrmHiperLojao.GerarFornecedor;
var
   observacao, email : string;
begin
  inherited;

  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

     SQL.Add('   SELECT   ');
     SQL.Add('       FORNECEDORES.FOR_CODIGO AS COD_FORNECEDOR,   ');
     SQL.Add('       FORNECEDORES.FOR_RAZAO AS DES_FORNECEDOR,   ');
     SQL.Add('       FORNECEDORES.FOR_FANTASIA AS DES_FANTASIA,   ');
     SQL.Add('       REPLACE(REPLACE(REPLACE(FORNECEDORES.FOR_CGC, ''.'', ''''), ''-'', ''''), ''/'', '''') AS NUM_CGC,   ');
     SQL.Add('       CASE FORNECEDORES.FOR_IE   ');
     SQL.Add('           WHEN ''0000'' THEN ''ISENTO''   ');
     SQL.Add('           WHEN ''00000'' THEN ''ISENTO''   ');
     SQL.Add('           WHEN ''000000'' THEN ''ISENTO''   ');
     SQL.Add('           WHEN ''0000000'' THEN ''ISENTO''   ');
     SQL.Add('           WHEN ''isento'' THEN ''ISENTO''   ');
     SQL.Add('           WHEN '''' THEN ''ISENTO''   ');
     SQL.Add('           WHEN ''ISENTO'' THEN ''ISENTO''   ');
     SQL.Add('           ELSE REPLACE(REPLACE(REPLACE(FORNECEDORES.FOR_IE, ''.'', ''''), ''-'', ''''), ''/'', '''')  ');
     SQL.Add('       END AS NUM_INSC_EST,   ');
     SQL.Add('       COALESCE(FORNECEDORES.FOR_ENDERECO, ''A DEFINIR'') AS DES_ENDERECO,   ');
     SQL.Add('       COALESCE(FORNECEDORES.FOR_BAIRRO, ''A DEFINIR'') AS DES_BAIRRO,   ');
     SQL.Add('       FORNECEDORES.FOR_CIDADE AS DES_CIDADE,   ');
     SQL.Add('       COALESCE(FORNECEDORES.FOR_UF, '''') AS DES_SIGLA,   ');
     SQL.Add('       COALESCE(FORNECEDORES.FOR_CEP, ''13560642'') AS NUM_CEP,   ');
     SQL.Add('       COALESCE(FORCONTATOS.FCT_FONE, '''') AS NUM_FONE,   ');
     SQL.Add('       '''' AS NUM_FAX,   ');
     SQL.Add('       COALESCE(FORCONTATOS.FCT_CONTATO, '''') AS DES_CONTATO,   ');
     SQL.Add('       0 AS QTD_DIA_CARENCIA,   ');
     SQL.Add('       0 AS NUM_FREQ_VISITA,   ');
     SQL.Add('       0 AS VAL_DESCONTO,   ');
     SQL.Add('       COALESCE(FORNECEDORES.FOR_PRAZO01, 0) AS NUM_PRAZO,   ');
     SQL.Add('       ''N'' AS ACEITA_DEVOL_MER,   ');
     SQL.Add('       ''N'' AS CAL_IPI_VAL_BRUTO,   ');
     SQL.Add('       ''N'' AS CAL_ICMS_ENC_FIN,   ');
     SQL.Add('       ''N'' AS CAL_ICMS_VAL_IPI,   ');
     SQL.Add('       ''N'' AS MICRO_EMPRESA,   ');
     SQL.Add('       FORNECEDORES.FOR_CODIGO AS COD_FORNECEDOR_ANT,   ');
     SQL.Add('       COALESCE(FORNECEDORES.FOR_NUMERO, ''S/N'') AS NUM_ENDERECO,   ');
     SQL.Add('       COALESCE(CAST(FORNECEDORES.FOR_OBS AS VARCHAR(300)), '''') AS DES_OBSERVACAO,   ');
     SQL.Add('       COALESCE(FORNECEDORES.FOR_EMAIL, '''') AS DES_EMAIL,   ');
     SQL.Add('       COALESCE(FORNECEDORES.FOR_HOMEPAGE, '''') AS DES_WEB_SITE,    ');
     SQL.Add('       ''N'' AS FABRICANTE,   ');
     SQL.Add('       ''N'' AS FLG_PRODUTOR_RURAL,   ');
     SQL.Add('       0 AS TIPO_FRETE,   ');
     SQL.Add('       ''N'' AS FLG_SIMPLES,   ');
     SQL.Add('       ''N'' AS FLG_SUBSTITUTO_TRIB,   ');
     SQL.Add('       0 AS COD_CONTACCFORN,   ');
     SQL.Add('       ''N'' AS INATIVO,   ');
     SQL.Add('       0 AS COD_CLASSIF,   ');
     SQL.Add('       FORNECEDORES.FOR_CADASTRO_DATA AS DTA_CADASTRO,   ');
     SQL.Add('       0 AS VAL_CREDITO,   ');
     SQL.Add('       0 AS VAL_DEBITO,   ');
     SQL.Add('       1 AS PED_MIN_VAL,   ');
     SQL.Add('       '''' AS DES_EMAIL_VEND,   ');
     SQL.Add('       '''' AS SENHA_COTACAO,   ');
     SQL.Add('       -1 AS TIPO_PRODUTOR,   ');
     SQL.Add('       COALESCE(FORCONTATOS.FCT_CELULAR, '''') AS NUM_CELULAR   ');
     SQL.Add('   FROM   ');
     SQL.Add('       FORNECEDORES   ');
     SQL.Add('   LEFT JOIN FORCONTATOS ON FORNECEDORES.FOR_CODIGO = FORCONTATOS.FCT_FORNECEDOR ');


    Open;

    First;
    NumLinha := 0;

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);

      Layout.SetValues(QryPrincipal, NumLinha, RecordCount);

      Layout.FieldByName('NUM_CGC').AsString := StrRetNums(Layout.FieldByName('NUM_CGC').AsString);
      Layout.FieldByName('NUM_CEP').AsString := StrRetNums(Layout.FieldByName('NUM_CEP').AsString);
      Layout.FieldByName('DTA_CADASTRO').AsDateTime := FieldByName('DTA_CADASTRO').AsDateTime;
      //Layout.FieldByName('NUM_INSC_EST').AsString := StrRetNums(Layout.FieldByName('NUM_INSC_EST').AsString);
//


      if (QryPrincipal.FieldByName('NUM_INSC_EST').AsString = '') OR
      (QryPrincipal.FieldByName('NUM_INSC_EST').AsString = 'ISENTO') then
      begin
         Layout.FieldByName('NUM_INSC_EST').AsString := 'ISENTO';
      end;

      if (Layout.FieldByName('NUM_INSC_EST').AsString = 'ISENTO') OR
      (Layout.FieldByName('NUM_INSC_EST').AsString = 'isento') then
      begin
        Layout.FieldByName('NUM_INSC_EST').AsString := 'ISENTO';
      end;



      if QryPrincipal.FieldByName('DES_ENDERECO').AsString = '' then
        Layout.FieldByName('DES_ENDERECO').AsString := 'A DEFINIR';


      if Length(Layout.FieldByName('NUM_CGC').AsString) > 11 then
      begin
        if not ValidaCGC(Layout.FieldByName('NUM_CGC').AsString) then
          Layout.FieldByName('NUM_CGC').AsString := '';
      end
      else
        if not ValidaCPF(Layout.FieldByName('NUM_CGC').AsString) then
          Layout.FieldByName('NUM_CGC').AsString := '';

      Layout.FieldByName('NUM_FONE').AsString := StrRetNums( FieldByName('NUM_FONE').AsString );

      Layout.FieldByName('DES_OBSERVACAO').AsString := StrLBReplace(FieldByName('DES_OBSERVACAO').AsString);
      Layout.FieldByName('DES_EMAIL').AsString := StrReplace(StrLBReplace(FieldByName('DES_EMAIL').AsString), '\n', '');


//      if Layout.FieldByName('FLG_PRODUTOR_RURAL').AsString = 'S' then
//      begin
//        if StrRetNums(Layout.FieldByName('NUM_INSC_EST').AsString) = '' then
//            Layout.FieldByName('TIPO_PRODUTOR').AsInteger := 0
//        else
//            Layout.FieldByName('TIPO_PRODUTOR').AsInteger := 1
//      end;

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal.RecNo, 'E', E.Message);
    end;
    Next;
    end;
    Close;
  end;
end;

procedure TFrmHiperLojao.GerarGrupo;
var
   TotalCount : Integer;
begin
  inherited;

  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

   SQL.Add('   SELECT DISTINCT   ');
   SQL.Add('       COALESCE(LINHAS.LIN_CODIGO, 999) AS COD_SECAO,   ');
   SQL.Add('       COALESCE(GRUPOS.GRU_CODIGO, 999) AS COD_GRUPO,   ');
   SQL.Add('       COALESCE(GRUPOS.GRU_NOME, ''A DEFINIR'') AS DES_GRUPO,   ');
   SQL.Add('       0 AS VAL_META   ');
   SQL.Add('   FROM   ');
   SQL.Add('       PRODUTOS   ');
   SQL.Add('   LEFT JOIN LINHAS ON PRODUTOS.PRO_LINHA = LINHAS.LIN_CODIGO   ');
   SQL.Add('   LEFT JOIN GRUPOS ON PRODUTOS.PRO_GRUPO = GRUPOS.GRU_CODIGO   ');
   SQL.Add('   LEFT JOIN PRODUTOS_EMPRESA ON PRODUTOS.PRO_CODIGO = PRODUTOS_EMPRESA.PRE_PRODUTO');
   SQL.Add('   WHERE PRODUTOS_EMPRESA.PRE_EMPRESA = '+CbxLoja.Text+' ');


    Open;

    First;
    NumLinha := 0;

    TotalCount := SetCountTotal(SQL.Text);

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);

      Layout.SetValues(QryPrincipal, NumLinha, TotalCount);

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;

end;

procedure TFrmHiperLojao.GerarInfoNutricionais;
begin
  inherited;

  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
    SQL.Add('    NUTRICIONAL.ID AS COD_INFO_NUTRICIONAL,');
    SQL.Add('    NUTRICIONAL.DESCRITIVO AS DES_INFO_NUTRICIONAL,');
    SQL.Add('    NUTRICIONAL.QUANTIDADE AS PORCAO,');
    SQL.Add('    NUTRICIONAL.VALOR_CALORICO AS VALOR_CALORICO,');
    SQL.Add('    NUTRICIONAL.CARBOIDRATOS AS CARBOIDRATO,');
    SQL.Add('    NUTRICIONAL.PROTEINA AS PROTEINA,');
    SQL.Add('    NUTRICIONAL.GORDURAS AS GORDURA_TOTAL,');
    SQL.Add('    NUTRICIONAL.GORDURAS_SATURADA AS GORDURA_SATURADA,');
    SQL.Add('    NUTRICIONAL.COLESTEROL AS COLESTEROL,');
    SQL.Add('    NUTRICIONAL.FIBRA_ALIMENTAR AS FIBRA_ALIMENTAR,');
    SQL.Add('    NUTRICIONAL.CALCIO AS CALCIO,');
    SQL.Add('    NUTRICIONAL.FERRO AS FERRO,');
    SQL.Add('    NUTRICIONAL.SODIO AS SODIO,');
    SQL.Add('    (NUTRICIONAL.VALOR_CALORICO * 100) / 2000 AS VD_VALOR_CALORICO,');
    SQL.Add('    (NUTRICIONAL.CARBOIDRATOS * 100) / 300 AS VD_CARBOIDRATO,');
    SQL.Add('    (NUTRICIONAL.PROTEINA * 100) / 75 AS VD_PROTEINA,');
    SQL.Add('    (NUTRICIONAL.GORDURAS * 100) / 55 AS VD_GORDURA_TOTAL,');
    SQL.Add('    (NUTRICIONAL.GORDURAS_SATURADA * 100) / 22 AS VD_GORDURA_SATURADA,');
    SQL.Add('    (NUTRICIONAL.COLESTEROL * 100) / 300 AS VD_COLESTEROL,');
    SQL.Add('    (NUTRICIONAL.FIBRA_ALIMENTAR * 100) / 25 AS VD_FIBRA_ALIMENTAR,');
    SQL.Add('    (NUTRICIONAL.CALCIO * 100) / 1000 AS VD_CALCIO,');
    SQL.Add('    (NUTRICIONAL.FERRO * 100) / 14 AS VD_FERRO,');
    SQL.Add('    (NUTRICIONAL.SODIO * 100) / 2400 AS VD_SODIO,');
    SQL.Add('    NUTRICIONAL.GORDURATRANS AS GORDURA_TRANS,');
    SQL.Add('    0 AS VD_GORDURA_TRANS,');
    SQL.Add('');
    SQL.Add('    CASE NUTRICIONAL.UNIDADE');
    SQL.Add('        WHEN 0 THEN ''G''');
    SQL.Add('        WHEN 1 THEN ''ML''');
    SQL.Add('        WHEN 2 THEN ''UN''');
    SQL.Add('        ELSE ''KG''');
    SQL.Add('    END AS UNIDADE_PORCAO,');
    SQL.Add('');
    SQL.Add('    CASE MED_CASEIRA');
    SQL.Add('        WHEN 25 THEN NUTRICIONAL.MEDIDAI || '' '' || ''PITADA(S)''');
    SQL.Add('        WHEN 6 THEN NUTRICIONAL.MEDIDAI || '' '' || ''PACOTE(S)''');
    SQL.Add('        WHEN 21 THEN NUTRICIONAL.MEDIDAI || '' '' || ''FIL�(S)''');
    SQL.Add('        WHEN 20 THEN NUTRICIONAL.MEDIDAI || '' '' || ''BIFE(S)''');
    SQL.Add('        WHEN 2 THEN NUTRICIONAL.MEDIDAI || '' '' || ''COLHER(ES) DE CH�''');
    SQL.Add('        WHEN 5 THEN NUTRICIONAL.MEDIDAI || '' '' || ''UNIDADE''');
    SQL.Add('        WHEN 24 THEN NUTRICIONAL.MEDIDAI || '' '' || ''PRATO(S) FUNDO(S)''');
    SQL.Add('        WHEN 4 THEN NUTRICIONAL.MEDIDAI || '' '' || ''DE X�CARA(S)''');
    SQL.Add('        WHEN 8 THEN NUTRICIONAL.MEDIDAI || '' '' || ''FATIA(S) FINA(S)''');
    SQL.Add('        WHEN 7 THEN NUTRICIONAL.MEDIDAI || '' '' || ''FATIA(S)''');
    SQL.Add('        WHEN 3 THEN NUTRICIONAL.MEDIDAI || '' '' || ''X�CARA(S)''');
    SQL.Add('        WHEN 15 THEN NUTRICIONAL.MEDIDAI || '' '' || ''COPO(S)''');
    SQL.Add('        WHEN 0 THEN NUTRICIONAL.MEDIDAI || '' '' || ''COLHER(ES) DE SOPA''');
    SQL.Add('        WHEN 16 THEN NUTRICIONAL.MEDIDAI || '' '' || ''POR��O(�ES)''');
    SQL.Add('        WHEN 9 THEN NUTRICIONAL.MEDIDAI || '' '' || ''PEDA�O(S)''');
    SQL.Add('    END AS DES_PORCAO,');
    SQL.Add('    -- '''' AS DES_PORCAO,');
    SQL.Add('');
    SQL.Add('    NUTRICIONAL.MEDIDAI AS PARTE_INTEIRA_MED_CASEIRA,');
    SQL.Add('    MED_CASEIRA AS MED_CASEIRA_UTILIZADA');
    SQL.Add('FROM');
    SQL.Add('    NUTRICIONAL');
    SQL.Add('INNER JOIN');
    SQL.Add('    VALORES_NUTRI VD');
    SQL.Add('ON');
    SQL.Add('    NUTRICIONAL.REFVD = VD.ID');

    Open;

    First;

    NumLinha := 0;

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);

      Layout.SetValues(QryPrincipal, NumLinha, RecordCount);

//      Layout.FieldByName('COD_INFO_NUTRICIONAL').AsString := GerarPLU( Layout.FieldByName('COD_INFO_NUTRICIONAL').AsString );

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;

end;

procedure TFrmHiperLojao.GerarNCM;
var
 count, TotalCount : Integer;
begin
  inherited;
  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

     SQL.Add('   SELECT DISTINCT   ');
     SQL.Add('       1 AS COD_NCM,   ');
     SQL.Add('       ''A DEFINIR'' AS DES_NCM,   ');
     SQL.Add('       84099999 AS NUM_NCM,   ');
     SQL.Add('       ''N'' AS FLG_NAO_PIS_COFINS,   ');
     SQL.Add('       1 AS TIPO_NAO_PIS_COFINS,   ');
     SQL.Add('       999 AS COD_TAB_SPED,   ');
     SQL.Add('       '''' AS NUM_CEST,   ');
     SQL.Add('       ''SP'' AS DES_SIGLA,   ');
     SQL.Add('       1 AS COD_TRIB_ENTRADA,   ');
     SQL.Add('       1 AS COD_TRIB_SAIDA,   ');
     SQL.Add('       0 AS PER_IVA,   ');
     SQL.Add('       0 AS PER_FCP_ST   ');
     SQL.Add('   FROM   ');
     SQL.Add('       PRODUTOS   ');
     SQL.Add('   LEFT JOIN PRODUTOS_EMPRESA ON PRODUTOS.PRO_CODIGO = PRODUTOS_EMPRESA.PRE_PRODUTO');
     SQL.Add('   WHERE PRODUTOS_EMPRESA.PRE_EMPRESA = '+CbxLoja.Text+' ');

    Open;
    First;

    count := 0;


    NumLinha := 0;

    TotalCount := SetCountTotal(SQL.Text);

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);
      Inc(count);
      Layout.SetValues(QryPrincipal, NumLinha, TotalCount);

      Layout.FieldByName('NUM_NCM').AsString := StrRetNums(Layout.FieldByName('NUM_NCM').AsString);
      Layout.FieldByName('NUM_CEST').AsString := StrRetNums( Layout.FieldByName('NUM_CEST').AsString );

      //Layout.FieldByName('COD_NCM').AsInteger := count;

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;
end;

procedure TFrmHiperLojao.GerarNCMUF;
var
 count, TotalCount : Integer;
begin
  inherited;
  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

     SQL.Add('   SELECT DISTINCT   ');
     SQL.Add('       1 AS COD_NCM,   ');
     SQL.Add('       ''A DEFINIR'' AS DES_NCM,   ');
     SQL.Add('       84099999 AS NUM_NCM,   ');
     SQL.Add('       ''N'' AS FLG_NAO_PIS_COFINS,   ');
     SQL.Add('       1 AS TIPO_NAO_PIS_COFINS,   ');
     SQL.Add('       999 AS COD_TAB_SPED,   ');
     SQL.Add('       '''' AS NUM_CEST,   ');
     SQL.Add('       ''SP'' AS DES_SIGLA,   ');
     SQL.Add('       1 AS COD_TRIB_ENTRADA,   ');
     SQL.Add('       1 AS COD_TRIB_SAIDA,   ');
     SQL.Add('       0 AS PER_IVA,   ');
     SQL.Add('       0 AS PER_FCP_ST   ');
     SQL.Add('   FROM   ');
     SQL.Add('       PRODUTOS   ');
     SQL.Add('   LEFT JOIN PRODUTOS_EMPRESA ON PRODUTOS.PRO_CODIGO = PRODUTOS_EMPRESA.PRE_PRODUTO');
     SQL.Add('   WHERE PRODUTOS_EMPRESA.PRE_EMPRESA = '+CbxLoja.Text+' ');

    Open;
    First;

    count := 0;


    NumLinha := 0;

    TotalCount := SetCountTotal(SQL.Text);

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);
      Inc(count);
      Layout.SetValues(QryPrincipal, NumLinha, TotalCount);

      //Layout.FieldByName('COD_NCM').AsInteger := count;

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;
end;

procedure TFrmHiperLojao.GerarNFClientes;
begin
  inherited;
  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

    Open;

    First;
    NumLinha := 0;

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);

      Layout.SetValues(QryPrincipal, NumLinha, RecordCount);

      Layout.FieldByName('DTA_EMISSAO').AsDateTime := QryPrincipal.FieldByName('DTA_EMISSAO').AsDateTime;
      Layout.FieldByName('DTA_ENTRADA').AsDateTime := QryPrincipal.FieldByName('DTA_ENTRADA').AsDateTime;

      Layout.FieldByName('DES_OBSERVACAO').AsString := StrLBReplace(FieldByName('DES_OBSERVACAO').AsString);

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;
end;

procedure TFrmHiperLojao.GerarNFFornec;
begin
  inherited;
  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

     SQL.Add('   SELECT   ');
     SQL.Add('       CAPA.NFC_FORNECEDOR AS COD_FORNECEDOR,   ');
     SQL.Add('       CAPA.NFC_NUMERO AS NUM_NF_FORN,   ');
     SQL.Add('       CAPA.NFC_SERIE AS NUM_SERIE_NF,   ');
     SQL.Add('       '''' AS NUM_SUBSERIE_NF,   ');
     SQL.Add('       '''' AS CFOP,   ');
     SQL.Add('       0 AS TIPO_NF,   ');
     SQL.Add('       CAPA.NFC_MODELO AS DES_ESPECIE,   ');
     SQL.Add('       CAPA.NFC_VRTOTAL AS VAL_TOTAL_NF,   ');
     SQL.Add('       CAPA.NFC_DTEMISSAO AS DTA_EMISSAO,   ');
     SQL.Add('       CAPA.NFC_DTENTSAI AS DTA_ENTRADA,   ');
     SQL.Add('       COALESCE(CAPA.NFC_IPI_VR, 0) AS VAL_TOTAL_IPI,   ');
     SQL.Add('       0 AS VAL_VENDA_VAREJO,   ');
     SQL.Add('       0 AS VAL_FRETE,   ');
     SQL.Add('       0 AS VAL_ACRESCIMO,   ');
     SQL.Add('       COALESCE(CAPA.NFC_DESCONTO, 0) AS VAL_DESCONTO,   ');
     SQL.Add('       REPLACE(REPLACE(REPLACE(FORNECEDORES.FOR_CGC, ''.'', ''''), ''-'', ''''), ''/'', '''') AS NUM_CGC,   ');
     SQL.Add('       CAPA.NFC_ICMS_VRBASE AS VAL_TOTAL_BC,   ');
     SQL.Add('       CAPA.NFC_ICMS_VR AS VAL_TOTAL_ICMS,   ');
     SQL.Add('       CAPA.NFC_ICMS_VRBASEST AS VAL_BC_SUBST,   ');
     SQL.Add('       CAPA.NFC_ICMS_VRST AS VAL_ICMS_SUBST,   ');
     SQL.Add('       0 AS VAL_FUNRURAL,   ');
     SQL.Add('       1 AS COD_PERFIL,   ');
     SQL.Add('       0 AS VAL_DESP_ACESS,   ');
     SQL.Add('       COALESCE(CAPA.NFC_CANCELADA, ''N'') AS FLG_CANCELADO,   ');
     SQL.Add('       COALESCE(CAPA.NFC_OBS_01, '''') AS DES_OBSERVACAO,   ');
     SQL.Add('       CAPA.NFC_NFE_CHAVE_IDENTIFICACAO AS NUM_CHAVE_ACESSO   ');
     SQL.Add('   FROM   ');
     SQL.Add('       NF_COMPRA AS CAPA   ');
     SQL.Add('   INNER JOIN FORNECEDORES ON CAPA.NFC_FORNECEDOR = FORNECEDORES.FOR_CODIGO   ');
     SQL.Add('   WHERE CAPA.NFC_FORCLI = ''F''   ');
     SQL.Add('AND');
     SQL.Add(' CAPA.NFC_EMPRESA = '+CbxLoja.Text+' ');
     SQL.Add(' AND ');
     SQL.Add(' CAPA.NFC_DTEMISSAO BETWEEN :INI AND :FIM ');
     SQL.Add(' ORDER BY CAPA.NFC_NUMERO, CAPA.NFC_FORNECEDOR, CAPA.NFC_SERIE   ');

    ParamByName('INI').AsDate := DtpInicial.Date;
    ParamByName('FIM').AsDate := DtpFinal.Date;



    Open;

    First;
    //count := 100000;

    //TotalCount := SetCountTotal(SQL.Text);
    TotalCont := SetCountTotal(SQL.Text,ParamByName('INI').AsString,ParamByName('FIM').AsString);
    NumLinha := 0;

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);

      Layout.SetValues(QryPrincipal, NumLinha, TotalCont);

      Layout.FieldByName('DTA_EMISSAO').AsDateTime := QryPrincipal.FieldByName('DTA_EMISSAO').AsDateTime;
      Layout.FieldByName('DTA_ENTRADA').AsDateTime := QryPrincipal.FieldByName('DTA_ENTRADA').AsDateTime;

      Layout.FieldByName('DES_OBSERVACAO').AsString := StrLBReplace(FieldByName('DES_OBSERVACAO').AsString);

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;
end;

procedure TFrmHiperLojao.GerarNFitensClientes;
begin
  inherited;
  with QryPrincipal do
  begin
    Close;
    SQL.Clear;


    Open;

    First;
    NumLinha := 0;

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);

      Layout.SetValues(QryPrincipal, NumLinha, RecordCount);



      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;
end;

procedure TFrmHiperLojao.GerarNFitensFornec;
var
   fornecedor, nota, serie : string;
   count, TotalCount, NovoCodigo : integer;

begin
  inherited;

  with QryPrincipal do
  begin
    Close;
    SQL.Clear;
     SQL.Add('   SELECT      ');
     SQL.Add('       ITENS.NIC_FORNECEDOR AS COD_FORNECEDOR,      ');
     SQL.Add('       ITENS.NIC_NUMERO AS NUM_NF_FORN,      ');
     SQL.Add('       ITENS.NIC_SERIE AS NUM_SERIE_NF,      ');
     SQL.Add('       CASE      ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''25578,'' THEN CAST(25578 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''28687.'' THEN CAST(28687 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''29397.'' THEN CAST(29397 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''29*504'' THEN CAST(29504 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''296,5'' THEN CAST(2965 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''29878.'' THEN CAST(29878 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''29870,'' THEN CAST(29870 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''2995,'' THEN CAST(2995 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''30023Q'' THEN CAST(30023 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''30045,'' THEN CAST(30045 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''30304,'' THEN CAST(30304 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN CHAR_LENGTH(PRODUTOS.PRO_REFERENCIA) = 9 THEN PRODUTOS.COD_PRODUTO   ');
     SQL.Add('           ELSE PRODUTOS.PRO_REFERENCIA        ');
     SQL.Add('       END AS COD_PRODUTO,      ');
     SQL.Add('       1 AS COD_TRIBUTACAO,      ');
//     SQL.Add('       CASE   ');
//     SQL.Add('           WHEN SUBSTRING(REPLACE(COALESCE(ITENS.NIC_EMBALAGEM, 1), ''/'', '''') FROM 4 FOR 5) = '''' THEN 1   ');
//     SQL.Add('           WHEN SUBSTRING(REPLACE(COALESCE(ITENS.NIC_EMBALAGEM, 1), ''/'', '''') FROM 4 FOR 5) = 0 THEN 1   ');
//     SQL.Add('           ELSE SUBSTRING(REPLACE(COALESCE(ITENS.NIC_EMBALAGEM, 1), ''/'', '''') FROM 4 FOR 5)    ');
//     SQL.Add('       END AS QTD_EMBALAGEM,   ');
     SQL.Add('              CASE      ');
     SQL.Add('                   WHEN SUBSTRING(REPLACE(COALESCE(ITENS.NIC_EMBALAGEM, 1), ''/'', '''') FROM 4 FOR 5) = '''' THEN 1   ');
     SQL.Add('                   WHEN SUBSTRING(REPLACE(COALESCE(ITENS.NIC_EMBALAGEM, 1), ''/'', '''') FROM 4 FOR 5) = '',5'' THEN 1   ');
     SQL.Add('                   WHEN SUBSTRING(REPLACE(COALESCE(ITENS.NIC_EMBALAGEM, 1), ''/'', '''') FROM 4 FOR 5) = ''04'' THEN 4   ');
     SQL.Add('                   WHEN SUBSTRING(REPLACE(COALESCE(ITENS.NIC_EMBALAGEM, 1), ''/'', '''') FROM 4 FOR 5) = ''05'' THEN 5   ');
     SQL.Add('                   WHEN SUBSTRING(REPLACE(COALESCE(ITENS.NIC_EMBALAGEM, 1), ''/'', '''') FROM 4 FOR 5) = ''06'' THEN 6   ');
     SQL.Add('                   WHEN SUBSTRING(REPLACE(COALESCE(ITENS.NIC_EMBALAGEM, 1), ''/'', '''') FROM 4 FOR 5) = 0 THEN 1   ');
     SQL.Add('                   ELSE SUBSTRING(REPLACE(COALESCE(ITENS.NIC_EMBALAGEM, 1), ''/'', '''') FROM 4 FOR 5)       ');
     SQL.Add('               END AS QTD_EMBALAGEM,   ');
     SQL.Add('       COALESCE(ITENS.NIC_QUANTIDADE_RECEBIDA, 1) AS QTD_ENTRADA,      ');
     SQL.Add('       SUBSTRING(COALESCE(ITENS.NIC_EMBALAGEM, ''UN'')FROM 1 FOR 2) AS DES_UNIDADE,   ');
     //SQL.Add('       ITENS.NIC_VRLIQ AS VAL_TABELA,      ');
     SQL.Add('       ITENS.NIC_VRTOTAL / COALESCE(ITENS.NIC_QUANTIDADE_RECEBIDA, 1) AS VAL_TABELA,      ');
     SQL.Add('       COALESCE(ITENS.NIC_DESCONTO_VALOR, 0) AS VAL_DESCONTO_ITEM,      ');
     SQL.Add('       0 AS VAL_ACRESCIMO_ITEM,      ');
     SQL.Add('       (ITENS.NIC_IPI_VR / COALESCE(ITENS.NIC_QUANTIDADE_RECEBIDA, 1)) AS VAL_IPI_ITEM,      ');
     SQL.Add('       0 AS VAL_SUBST_ITEM,      ');
     SQL.Add('       ITENS.NIC_FRETE_VR AS VAL_FRETE_ITEM,      ');
     SQL.Add('       ITENS.NIC_ICMS_VR AS VAL_CREDITO_ICMS,      ');
     SQL.Add('       0 AS VAL_VENDA_VAREJO,      ');
     SQL.Add('       ITENS.NIC_VRTOTAL AS VAL_TABELA_LIQ,      ');
     SQL.Add('       REPLACE(REPLACE(REPLACE(FORNECEDORES.FOR_CGC, ''.'', ''''), ''-'', ''''), ''/'', '''') AS NUM_CGC,      ');
     SQL.Add('       COALESCE(ITENS.NIC_ICMS_VRBASE, 0) AS VAL_TOT_BC_ICMS,      ');
     SQL.Add('       COALESCE(ITENS.NIC_ICMS_VROUTROS, 0) AS VAL_TOT_OUTROS_ICMS,      ');
     SQL.Add('       REPLACE(ITENS.NIC_CFOP_CFOP, ''.'', '''') AS CFOP,      ');
     SQL.Add('       COALESCE(ITENS.NIC_ICMS_VRISENTO, 0) AS VAL_TOT_ISENTO,      ');
     SQL.Add('       COALESCE(ITENS.NIC_ICMS_VRBASEST, 0) AS VAL_TOT_BC_ST,      ');
     SQL.Add('       COALESCE(ITENS.NIC_ICMS_VRST, 0) AS VAL_TOT_ST,      ');
     SQL.Add('       ITENS.NIC_SEQUENCIA AS NUM_ITEM,      ');
     SQL.Add('       0 AS TIPO_IPI,      ');
     SQL.Add('       REPLACE(ITENS.NIC_CLASSIFICACAO, ''.'', '''') AS NUM_NCM,      ');
     SQL.Add('       COALESCE(ITENS.NIC_REFERENCIA, '''') AS DES_REFERENCIA      ');
     SQL.Add('   FROM   ');
     SQL.Add('       NF_COMPRA_ITENS AS ITENS   ');
     SQL.Add('   INNER JOIN FORNECEDORES ON ITENS.NIC_FORNECEDOR = FORNECEDORES.FOR_CODIGO   ');
     SQL.Add('   LEFT JOIN PRODUTOS ON ITENS.NIC_PRO_CODIGO = PRODUTOS.PRO_CODIGO   ');
     SQL.Add('   INNER JOIN NF_COMPRA AS CAPA ON ITENS.NIC_NUMERO = CAPA.NFC_NUMERO AND ITENS.NIC_FORNECEDOR = CAPA.NFC_FORNECEDOR   ');
     SQL.Add('   WHERE ITENS.NIC_FORCLI = ''F''   ');
     SQL.Add('   AND ITENS.NIC_EMPRESA = '+CbxLoja.Text+' ');
     SQL.Add('   AND');
     SQL.Add('   CAPA.NFC_DTEMISSAO BETWEEN :INI AND :FIM');
     SQL.Add('   ORDER BY ITENS.NIC_NUMERO, ITENS.NIC_FORNECEDOR, ITENS.NIC_SERIE   ');

    ParamByName('INI').AsDate := DtpInicial.Date;
    ParamByName('FIM').AsDate := DtpFinal.Date;

    Open;

    First;
//
//    count := 100000;
//    TotalCount := SetCountTotal(SQL.Text);
    TotalCont := SetCountTotal(SQL.Text,ParamByName('INI').AsString,ParamByName('FIM').AsString);

    NumLinha := 0;
    NovoCodigo := 400000;

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);

      Layout.SetValues(QryPrincipal, NumLinha, TotalCont);

//      if( (Layout.FieldByName('COD_FORNECEDOR').AsString = fornecedor) and
//          (Layout.FieldByName('NUM_NF_FORN').AsString = nota) and
//          (Layout.FieldByName('NUM_SERIE_NF').AsString = serie) ) then
//      begin
//          inc(count);
//      end
//      else
//      begin
//        fornecedor := Layout.FieldByName('COD_FORNECEDOR').AsString;
//        nota := Layout.FieldByName('NUM_NF_FORN').AsString;
//        serie := Layout.FieldByName('NUM_SERIE_NF').AsString;
//        count := 1;
//      end;
//
//      Layout.FieldByName('NUM_ITEM').AsInteger := count;

//      if Layout.FieldByName('COD_PRODUTO').AsInteger = 0 then
//      begin
//        Inc(NovoCodigo);
//        Layout.FieldByName('COD_PRODUTO').AsString := GerarPLU(IntToStr(NovoCodigo));
//      end;
////

      Layout.FieldByName('COD_PRODUTO').AsString := GerarPLU( Layout.FieldByName('COD_PRODUTO').AsString ) ;


      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;

end;

procedure TFrmHiperLojao.GerarProdForn;
var
   TotalCount, NovoCodigo : Integer;
begin
  inherited;

  if FlgEstoqueTroca then
  begin
    EstoqueTroca;
    Exit;
  end;

  if FlgFornecedorPreferencial then
  begin
    FornecedorPreferencial;
    Exit;
  end;

  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

     SQL.Add('   SELECT DISTINCT  ');
     SQL.Add('       CASE  ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''25578,'' THEN CAST(25578 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''28687.'' THEN CAST(28687 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''29397.'' THEN CAST(29397 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''29*504'' THEN CAST(29504 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''296,5'' THEN CAST(2965 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''29878.'' THEN CAST(29878 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''29870,'' THEN CAST(29870 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''2995,'' THEN CAST(2995 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''30023Q'' THEN CAST(30023 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''30045,'' THEN CAST(30045 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''30304,'' THEN CAST(30304 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN CHAR_LENGTH(PRODUTOS.PRO_REFERENCIA) = 9 THEN PRODUTOS.COD_PRODUTO   ');
     SQL.Add('           ELSE PRODUTOS.PRO_REFERENCIA        ');
     SQL.Add('       END AS COD_PRODUTO,   ');
     SQL.Add('       PRODUTOS_FORNECEDORES.PRF_FORNECEDOR AS COD_FORNECEDOR,   ');
     SQL.Add('       COALESCE(PRODUTOS.PRO_REFERENCIAFOR, '''') AS DES_REFERENCIA,   ');
     SQL.Add('       REPLACE(REPLACE(REPLACE(FORNECEDORES.FOR_CGC, ''.'', ''''), ''-'', ''''), ''/'', '''') AS NUM_CGC,   ');
     SQL.Add('       0 AS COD_DIVISAO,   ');
     SQL.Add('       COALESCE(SUBSTRING(PRODUTOS.PRO_UNIDCOMPRA FROM 1 FOR 2), ''UN'') AS DES_UNIDADE_COMPRA,   ');
     SQL.Add('       1 AS QTD_EMBALAGEM_COMPRA   ');
     SQL.Add('   FROM   ');
     SQL.Add('       PRODUTOS_FORNECEDORES   ');
     SQL.Add('   INNER JOIN PRODUTOS ON PRODUTOS_FORNECEDORES.PRF_CODIGO = PRODUTOS.PRO_CODIGO   ');
     SQL.Add('   INNER JOIN FORNECEDORES ON PRODUTOS_FORNECEDORES.PRF_FORNECEDOR = FORNECEDORES.FOR_CODIGO   ');
     SQL.Add('   LEFT JOIN PRODUTOS_EMPRESA ON PRODUTOS.PRO_CODIGO = PRODUTOS_EMPRESA.PRE_PRODUTO');
     SQL.Add('   WHERE PRODUTOS_EMPRESA.PRE_EMPRESA = '+CbxLoja.Text+' ');
     //SQL.Add('   AND CHAR_LENGTH(PRODUTOS.PRO_REFERENCIA) <= 8');
     SQL.Add('   ORDER BY(   ');
     SQL.Add('       CASE WHEN PRODUTOS.PRO_LEGENDA IN (''N'',''P'') THEN PRODUTOS.PRO_LEGENDA END   ');
     SQL.Add('   ) DESC,   ');
     SQL.Add('   (   ');
     SQL.Add('       CASE WHEN PRODUTOS.PRO_LEGENDA NOT IN (''N'', ''P'') THEN PRODUTOS.PRO_LEGENDA END   ');
     SQL.Add('   ), PRODUTOS.PRO_CODIGO ASC   ');



    Open;

    First;

    NumLinha := 0;
    NovoCodigo := 400000;

    TotalCount := SetCountTotal(SQL.Text);

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);

      Layout.SetValues(QryPrincipal, NumLinha, TotalCount);


      Layout.FieldByName('COD_PRODUTO').AsString := GerarPLU( Layout.FieldByName('COD_PRODUTO').AsString );
     

      Layout.FieldByName('NUM_CGC').AsString := StrRetNums(Layout.FieldByName('NUM_CGC').AsString);

//      if Layout.FieldByName('COD_PRODUTO').AsInteger = 0 then
//      begin
//        Inc(NovoCodigo);
//        Layout.FieldByName('COD_PRODUTO').AsString := GerarPLU(IntToStr(NovoCodigo));
//      end;

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;

end;

procedure TFrmHiperLojao.GerarProdLoja;
var
   TotalCount, NovoCodigo : integer;
begin
  inherited;

  if FlgAtualizaEstoque then
  begin
    AtualizaEstoque;
    Exit;
  end;

  if FlgPercentualFidelidade then
  begin
    PercentualFidelidade;
    Exit;
  end;

  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

     SQL.Add('   SELECT DISTINCT   ');
     SQL.Add('       CASE       ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''25578,'' THEN CAST(25578 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''28687.'' THEN CAST(28687 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''29397.'' THEN CAST(29397 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''29*504'' THEN CAST(29504 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''296,5'' THEN CAST(2965 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''29878.'' THEN CAST(29878 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''29870,'' THEN CAST(29870 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''2995,'' THEN CAST(2995 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''30023Q'' THEN CAST(30023 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''30045,'' THEN CAST(30045 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''30304,'' THEN CAST(30304 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN CHAR_LENGTH(PRODUTOS.PRO_REFERENCIA) = 9 THEN PRODUTOS.COD_PRODUTO   ');
     SQL.Add('           ELSE PRODUTOS.PRO_REFERENCIA        ');
     SQL.Add('       END AS COD_PRODUTO,      ');
     SQL.Add('       CASE   ');
     SQL.Add('           WHEN PRODUTOS.PRO_UNIDCOMPRA <> ''UN'' THEN (PRODUTOS.PRO_VRLIQ / PRODUTOS.PRO_FTUCUV)   ');
     SQL.Add('           ELSE COALESCE(FIDELIDADE.PRF_VRBRU, 1)   ');
     SQL.Add('       END AS VAL_CUSTO_REP,   ');
     SQL.Add('       COALESCE(PRECOS.PRE_PRECOUN, 1) AS VAL_VENDA,      ');
     SQL.Add('       0 AS VAL_OFERTA,   ');
     SQL.Add('                    ');
     SQL.Add('       CASE PRODUTOS_ESTOQUES.PST_LOCAL   ');
     SQL.Add('           WHEN 1 THEN COALESCE(ESTOQUE.PST_FISICO, 1)   ');
     SQL.Add('           ELSE COALESCE(PRODUTOS_ESTOQUES.PST_FISICO, 1)   ');
     SQL.Add('       END AS QTD_EST_VDA,   ');
     SQL.Add('      ');
     SQL.Add('       '''' AS TECLA_BALANCA,      ');
     SQL.Add('       1 AS COD_TRIBUTACAO,      ');
     SQL.Add('       0 AS VAL_MARGEM,      ');
     SQL.Add('       1 AS QTD_ETIQUETA,      ');
     SQL.Add('       1 AS COD_TRIB_ENTRADA,      ');
     SQL.Add('                 ');
     SQL.Add('       CASE PRODUTOS.PRO_LEGENDA      ');
     SQL.Add('           WHEN ''C'' THEN ''S''      ');
     SQL.Add('           WHEN ''F'' THEN ''S''      ');
     SQL.Add('           WHEN ''L'' THEN ''S''      ');
     SQL.Add('           WHEN ''N'' THEN ''N''      ');
     SQL.Add('           WHEN ''P'' THEN ''N''      ');
     SQL.Add('           WHEN ''R'' THEN ''S''      ');
     SQL.Add('       END AS FLG_INATIVO,      ');
     SQL.Add('                 ');
     SQL.Add('       CASE       ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''25578,'' THEN CAST(25578 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''28687.'' THEN CAST(28687 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''29397.'' THEN CAST(29397 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''29*504'' THEN CAST(29504 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''296,5'' THEN CAST(2965 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''29878.'' THEN CAST(29878 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''29870,'' THEN CAST(29870 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''2995,'' THEN CAST(2995 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''30023Q'' THEN CAST(30023 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''30045,'' THEN CAST(30045 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN PRODUTOS.PRO_REFERENCIA = ''30304,'' THEN CAST(30304 + 50000 AS INTEGER)   ');
     SQL.Add('           WHEN CHAR_LENGTH(PRODUTOS.PRO_REFERENCIA) = 9 THEN PRODUTOS.COD_PRODUTO   ');
     SQL.Add('           ELSE PRODUTOS.PRO_REFERENCIA        ');
     SQL.Add('       END AS COD_PRODUTO_ANT,      ');
     SQL.Add('      ');
     SQL.Add('       00000001 AS NUM_NCM,      ');
     SQL.Add('       1 AS TIPO_NCM,      ');
     SQL.Add('       0 AS VAL_VENDA_2,      ');
     SQL.Add('       '''' AS DTA_VALIDA_OFERTA,      ');
     SQL.Add('       1 AS QTD_EST_MINIMO,      ');
     SQL.Add('       NULL AS COD_VASILHAME,      ');
     SQL.Add('                 ');
     SQL.Add('       CASE PRODUTOS.PRO_LEGENDA      ');
     SQL.Add('           WHEN ''P'' THEN ''S''      ');
     SQL.Add('           ELSE ''N''      ');
     SQL.Add('       END AS FORA_LINHA,      ');
     SQL.Add('                 ');
     SQL.Add('       0 AS QTD_PRECO_DIF,      ');
     SQL.Add('       0 AS VAL_FORCA_VDA,      ');
     SQL.Add('       '''' AS NUM_CEST,      ');
     SQL.Add('       0 AS PER_IVA,      ');
     SQL.Add('       0 AS PER_FCP_ST,      ');
     SQL.Add('       COALESCE(FIDELIDADE.PRF_PERC01, 0.00) AS PER_FIDELIDADE ');
     SQL.Add('   FROM      ');
     SQL.Add('       PRODUTOS      ');
     SQL.Add('       LEFT JOIN PRECOS ON PRODUTOS.PRO_CODIGO = PRECOS.PRE_CODIGO      ');
     SQL.Add('       LEFT JOIN PRODUTOS_ESTOQUES ON PRODUTOS.PRO_CODIGO = PRODUTOS_ESTOQUES.PST_PRODUTO   ');
     SQL.Add('       LEFT JOIN (   ');
     SQL.Add('           SELECT   ');
     SQL.Add('               PST_LOCAL,   ');
     SQL.Add('               PST_PRODUTO,   ');
     SQL.Add('               PST_FISICO   ');
     SQL.Add('          FROM   ');
     SQL.Add('               PRODUTOS_ESTOQUES   ');
     SQL.Add('           WHERE PST_LOCAL = 2   ');
     SQL.Add('           --AND PST_PRODUTO = 50785   ');
     SQL.Add('           GROUP BY PST_LOCAL, PST_PRODUTO, PST_FISICO   ');
     SQL.Add('       ) AS ESTOQUE   ');
     SQL.Add('       ON   ');
     SQL.Add('           PRODUTOS_ESTOQUES.PST_PRODUTO = ESTOQUE.PST_PRODUTO   ');
     SQL.Add('       LEFT JOIN (   ');
     SQL.Add('           SELECT DISTINCT   ');
     SQL.Add('               PRF_CODIGO,   ');
     SQL.Add('               PRF_PERC01,   ');
     SQL.Add('               PRF_VRBRU ');
     SQL.Add('           FROM   ');
     SQL.Add('               PRODUTOS_FORNECEDORES   ');
     SQL.Add('           WHERE   ');
     SQL.Add('               PRF_PERC01 IS NOT NULL   ');
     SQL.Add('           AND ');
     SQL.Add('              PRF_EMPRESA = '+CbxLoja.Text+' ');
     SQL.Add('           AND ');
     SQL.Add('              PRF_PADRAO = ''S'' ');
     SQL.Add('       ) AS FIDELIDADE   ');
     SQL.Add('       ON   ');
     SQL.Add('           PRODUTOS.PRO_CODIGO = FIDELIDADE.PRF_CODIGO   ');
     SQL.Add('       WHERE (   ');
     SQL.Add('           CASE PRODUTOS_ESTOQUES.PST_LOCAL   ');
     SQL.Add('               WHEN 2 THEN 1   ');
     SQL.Add('               WHEN 4 THEN 6   ');
     SQL.Add('               WHEN 5 THEN 3   ');
     SQL.Add('           END   ');
     SQL.Add('       ) = '+CbxLoja.Text+'   ');
     SQL.Add('       ORDER BY (      ');
     SQL.Add('           CASE WHEN PRODUTOS.PRO_LEGENDA IN (''N'',''P'') THEN PRODUTOS.PRO_LEGENDA END      ');
     SQL.Add('       ) DESC,      ');
     SQL.Add('       (      ');
     SQL.Add('           CASE WHEN PRODUTOS.PRO_LEGENDA NOT IN (''N'', ''P'') THEN PRODUTOS.PRO_LEGENDA END      ');
     SQL.Add('       ), PRODUTOS.PRO_CODIGO ASC   ');


     //ParamByName('COD_LOJA').AsString := CbxLoja.Text;




    Open;
    First;
    NumLinha := 0;
    //NovoCodigo := 400000;
    //NewCodNcm := 0;

    TotalCount := SetCountTotal(SQL.Text);

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);
      Layout.SetValues(QryPrincipal, NumLinha, TotalCount);


        Layout.FieldByName('COD_PRODUTO').AsString := GerarPLU( Layout.FieldByName('COD_PRODUTO').AsString ) ;

        Layout.FieldByName('COD_PRODUTO_ANT').AsString := GerarPLU( Layout.FieldByName('COD_PRODUTO_ANT').AsString ) ;


      //Layout.FieldByName('COD_PRODUTO_ANT').AsString :=  Layout.FieldByName('COD_PRODUTO_ANT').AsString;

//      if Layout.FieldByName('COD_PRODUTO').AsInteger = 0 then
//      begin
//        Inc(NovoCodigo);
//        Layout.FieldByName('COD_PRODUTO').AsString := GerarPLU(IntToStr(NovoCodigo));
//      end;
//
//      if Layout.FieldByName('COD_PRODUTO_ANT').AsInteger = 0 then
//      begin
//        Inc(NovoCodigo);
//        Layout.FieldByName('COD_PRODUTO_ANT').AsString := GerarPLU(IntToStr(NovoCodigo));
//      end;


      //Inc(NewCodNcm);
      //QryPrincipal.FieldByName('COD_NCM').AsInteger :=  NewCodNcm;

//      with QryCriaCodNcm do
//      begin
//        //ParamByName('COD_NCM').Value := 1;
////        ShowMessage(QryCriaCodNcm.Text);
//        ExecSQL;
//      end;


      //ShowMessage(Layout.FieldByName('TIPO_NCM').AsString);

      //Layout.FieldByName('NUM_NCM').AsString := StrRetNums(Layout.FieldByName('NUM_NCM').AsString);



//      Layout.FieldByName('DTA_VALIDA_OFERTA').AsDateTime := FieldByName('DTA_VALIDA_OFERTA').AsDateTime;

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal.RecNo, 'E', E.Message);
    end;
    Next;
    end;
    Close;
  end;
end;

procedure TFrmHiperLojao.GerarProdSimilar;
begin
  inherited;
  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
    SQL.Add('     FAMILIAS.ID AS COD_PRODUTO_SIMILAR,');
    SQL.Add('     FAMILIAS.DESCRITIVO AS DES_PRODUTO_SIMILAR,');
    SQL.Add('     0 AS VAL_META');
    SQL.Add('FROM');
    SQL.Add('     FAMILIAS');


    Open;    
    
    First;
    NumLinha := 0;

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);

      Layout.SetValues(QryPrincipal, NumLinha, RecordCount);

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;
end;

end.
