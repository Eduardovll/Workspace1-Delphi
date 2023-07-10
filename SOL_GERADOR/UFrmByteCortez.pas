unit UFrmByteCortez;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, UFrmModelo, Data.DBXOracle, Data.DB,
  Data.SqlExpr, Vcl.Menus, Vcl.StdCtrls, Vcl.Buttons, Vcl.ComCtrls,
  Vcl.Imaging.pngimage, Vcl.ExtCtrls, Data.DBXFirebird, Data.Win.ADODB,
  Vcl.Grids, Vcl.DBGrids, Datasnap.Provider, Datasnap.DBClient,
  //dxGDIPlusClasses,
  Math;

type
  TFrmByteCortez = class(TFrmModeloSis)
    btnGeraCest: TButton;
    BtnAmarrarCest: TButton;
    CbxLoja: TComboBox;
    lblLoja: TLabel;
    procedure btnGeraCestClick(Sender: TObject);
    procedure BtnAmarrarCestClick(Sender: TObject);
    procedure EdtCamBancoExit(Sender: TObject);
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

    procedure GerarScriptCEST;
    procedure GerarScriptAmarrarCEST;

  end;

var
  FrmByteCortez: TFrmByteCortez;
  ListNCM    : TStringList;
  TotalCont  : Integer;
  NumLinha : Integer;
  Arquivo: TextFile;
  FlgGeraDados : Boolean = false;
  FlgGeraCest : Boolean = false;
  FlgGeraAmarrarCest : Boolean = false;

implementation

{$R *.dfm}

uses xProc, UUtilidades, UProgresso;


procedure TFrmByteCortez.GerarProducao;
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

procedure TFrmByteCortez.GerarProduto;
var
 cod_produto, codbarras : string;
 TotalCount, count : Integer;
begin
  inherited;
  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT     ');
    SQL.Add('    PRODUTO.COD_PROD AS COD_PRODUTO,');
    SQL.Add('    PRODUTO.COD_BARRA AS COD_BARRA_PRINCIPAL,');
    SQL.Add('    PRODUTO.NOME_PROD AS DES_REDUZIDA,');
    SQL.Add('    PRODUTO.NOME_PROD AS DES_PRODUTO,');
    SQL.Add('    COALESCE(PRODUTO.QTD_EMBALAGEM_COMPRA, 1) AS QTD_EMBALAGEM_COMPRA,');
    SQL.Add('    PRODUTO.UNIDADE AS DES_UNIDADE_COMPRA,');
    SQL.Add('    1 AS QTD_EMBALAGEM_VENDA,');
    SQL.Add('    PRODUTO.UNIDADE AS DES_UNIDADE_VENDA,');
    SQL.Add('    0 AS TIPO_IPI,');
    SQL.Add('    0 AS VAL_IPI,');
    SQL.Add('    PRODUTO.COD_MARCA AS COD_SECAO,');
    SQL.Add('    999 AS COD_GRUPO,');
    SQL.Add('    999 AS COD_SUB_GRUPO,');
    SQL.Add('    0 AS COD_PRODUTO_SIMILAR, --');
    SQL.Add('    ');
    SQL.Add('    CASE');
    SQL.Add('        WHEN COALESCE(PRODUTO.BALANCA, ''N'') = ''S'' AND UPPER(PRODUTO.UNIDADE) = ''KG'' THEN ''S''');
    SQL.Add('        ELSE ''N''');
    SQL.Add('    END AS IPV,');
    SQL.Add('');
    SQL.Add('    0 AS DIAS_VALIDADE, --');
    SQL.Add('    0 AS TIPO_PRODUTO,');
    SQL.Add('');
    SQL.Add('    CASE COALESCE(PRODUTO.PIS_CST, 0)');
    SQL.Add('        WHEN 0 THEN ''N''');
    SQL.Add('        WHEN 1 THEN ''N''');
    SQL.Add('        WHEN 4 THEN ''S''');
    SQL.Add('        WHEN 6 THEN ''S''');
    SQL.Add('        WHEN 7 THEN ''N''');
    SQL.Add('        WHEN 8 THEN ''S''');
    SQL.Add('    END AS FLG_NAO_PIS_COFINS,');
    SQL.Add('');
    SQL.Add('    COALESCE(PRODUTO.BALANCA, ''N'') AS FLG_ENVIA_BALANCA,');
    SQL.Add('');
    SQL.Add('    CASE COALESCE(PRODUTO.PIS_CST, 0)');
    SQL.Add('        WHEN 0 THEN -1');
    SQL.Add('        WHEN 1 THEN -1');
    SQL.Add('        WHEN 4 THEN 1');
    SQL.Add('        WHEN 6 THEN 0');
    SQL.Add('        WHEN 7 THEN -1');
    SQL.Add('        WHEN 8 THEN 3');
    SQL.Add('    END AS TIPO_NAO_PIS_COFINS, --');
    SQL.Add('');
    SQL.Add('    0 AS TIPO_EVENTO,');
    SQL.Add('    0 AS COD_ASSOCIADO,');
    SQL.Add('    PRODUTO.DETALHE AS DES_OBSERVACAO,');
    SQL.Add('    0 AS COD_INFO_NUTRICIONAL,');
    SQL.Add('    0 AS COD_INFO_RECEITA,');
    SQL.Add('    999 AS COD_TAB_SPED, --');
    SQL.Add('    ''N'' AS FLG_ALCOOLICO,');
    SQL.Add('    0 AS TIPO_ESPECIE,');
    SQL.Add('    0 AS COD_CLASSIF,');
    SQL.Add('    CASE WHEN COALESCE(PRODUTO.PESO, 0) = 0 THEN 1 ELSE PRODUTO.PESO END AS VAL_VDA_PESO_BRUTO,');
    SQL.Add('    1 AS VAL_PESO_EMB,');
    SQL.Add('    0 AS TIPO_EXPLOSAO_COMPRA,');
    SQL.Add('    '''' AS DTA_INI_OPER,');
    SQL.Add('    '''' AS DES_PLAQUETA,');
    SQL.Add('    '''' AS MES_ANO_INI_DEPREC,');
    SQL.Add('    0 AS TIPO_BEM,');
    SQL.Add('    0 AS COD_FORNECEDOR,');
    SQL.Add('    0 AS NUM_NF,');
    SQL.Add('    '''' AS DTA_ENTRADA,');
    SQL.Add('    0 AS COD_NAT_BEM,');
    SQL.Add('    0 AS VAL_ORIG_BEM');
    SQL.Add('FROM');
    SQL.Add('    PRODUTOS AS PRODUTO');
    SQL.Add('');
    SQL.Add('UNION ALL');
    SQL.Add('');
    SQL.Add('SELECT     ');
    SQL.Add('    100000 + PRODUTO.COD_PROD AS COD_PRODUTO,');
    SQL.Add('    PRODUTO.COD_BARRA AS COD_BARRA_PRINCIPAL,');
    SQL.Add('    PRODUTO.NOME_PROD AS DES_REDUZIDA,');
    SQL.Add('    PRODUTO.NOME_PROD AS DES_PRODUTO,');
    SQL.Add('    COALESCE(PRODUTO.QTD_EMBALAGEM_COMPRA, 1) AS QTD_EMBALAGEM_COMPRA,');
    SQL.Add('    PRODUTO.UNIDADE AS DES_UNIDADE_COMPRA,');
    SQL.Add('    1 AS QTD_EMBALAGEM_VENDA,');
    SQL.Add('    PRODUTO.UNIDADE AS DES_UNIDADE_VENDA,');
    SQL.Add('    0 AS TIPO_IPI,');
    SQL.Add('    0 AS VAL_IPI,');
    SQL.Add('    200 + PRODUTO.COD_MARCA AS COD_SECAO,');
    SQL.Add('    999 AS COD_GRUPO,');
    SQL.Add('    999 AS COD_SUB_GRUPO,');
    SQL.Add('    0 AS COD_PRODUTO_SIMILAR, --');
    SQL.Add('    ');
    SQL.Add('    CASE');
    SQL.Add('        WHEN COALESCE(PRODUTO.BALANCA, ''N'') = ''S'' AND UPPER(PRODUTO.UNIDADE) = ''KG'' THEN ''S''');
    SQL.Add('        ELSE ''N''');
    SQL.Add('    END AS IPV,');
    SQL.Add('');
    SQL.Add('    0 AS DIAS_VALIDADE, --');
    SQL.Add('    0 AS TIPO_PRODUTO,');
    SQL.Add('');
    SQL.Add('    CASE COALESCE(PRODUTO.PIS_CST, 0)');
    SQL.Add('        WHEN 0 THEN ''N''');
    SQL.Add('        WHEN 1 THEN ''N''');
    SQL.Add('        WHEN 4 THEN ''S''');
    SQL.Add('        WHEN 5 THEN ''S''');
    SQL.Add('        WHEN 6 THEN ''S''');
    SQL.Add('        WHEN 7 THEN ''N''');
    SQL.Add('        WHEN 8 THEN ''S''');
    SQL.Add('    END AS FLG_NAO_PIS_COFINS,');
    SQL.Add('');
    SQL.Add('    CASE');
    SQL.Add('        WHEN COALESCE(PRODUTO.BALANCA, ''N'') = ''N'' THEN ''N''');
    SQL.Add('        WHEN COALESCE(PRODUTO.BALANCA, ''N'') = '''' THEN ''N''');
    SQL.Add('        ELSE ''S''');
    SQL.Add('    END AS FLG_ENVIA_BALANCA,');
    SQL.Add('');
    SQL.Add('    CASE COALESCE(PRODUTO.PIS_CST, 0)');
    SQL.Add('        WHEN 0 THEN -1');
    SQL.Add('        WHEN 1 THEN -1');
    SQL.Add('        WHEN 4 THEN 1');
    SQL.Add('        WHEN 5 THEN 2');
    SQL.Add('        WHEN 6 THEN 0');
    SQL.Add('        WHEN 7 THEN -1');
    SQL.Add('        WHEN 8 THEN 3');
    SQL.Add('    END AS TIPO_NAO_PIS_COFINS, --');
    SQL.Add('');
    SQL.Add('    0 AS TIPO_EVENTO,');
    SQL.Add('    0 AS COD_ASSOCIADO,');
    SQL.Add('    PRODUTO.DETALHE AS DES_OBSERVACAO,');
    SQL.Add('    0 AS COD_INFO_NUTRICIONAL,');
    SQL.Add('    0 AS COD_INFO_RECEITA,');
    SQL.Add('    999 AS COD_TAB_SPED, --');
    SQL.Add('    ''N'' AS FLG_ALCOOLICO,');
    SQL.Add('    0 AS TIPO_ESPECIE,');
    SQL.Add('    0 AS COD_CLASSIF,');
    SQL.Add('    CASE WHEN COALESCE(PRODUTO.PESO, 0) = 0 THEN 1 ELSE PRODUTO.PESO END AS VAL_VDA_PESO_BRUTO,');
    SQL.Add('    1 AS VAL_PESO_EMB,');
    SQL.Add('    0 AS TIPO_EXPLOSAO_COMPRA,');
    SQL.Add('    '''' AS DTA_INI_OPER,');
    SQL.Add('    '''' AS DES_PLAQUETA,');
    SQL.Add('    '''' AS MES_ANO_INI_DEPREC,');
    SQL.Add('    0 AS TIPO_BEM,');
    SQL.Add('    0 AS COD_FORNECEDOR,');
    SQL.Add('    0 AS NUM_NF,');
    SQL.Add('    '''' AS DTA_ENTRADA,');
    SQL.Add('    0 AS COD_NAT_BEM,');
    SQL.Add('    0 AS VAL_ORIG_BEM');
    SQL.Add('FROM');
    SQL.Add('    TAB_PRODUTO_AUX AS PRODUTO');
    SQL.Add('LEFT JOIN PRODUTOS');
    SQL.Add('ON PRODUTO.COD_BARRA = PRODUTOS.COD_BARRA');
    SQL.Add('WHERE PRODUTOS.COD_PROD IS NULL');


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

      Layout.FieldByName('COD_PRODUTO').AsString := GerarPLU( Layout.FieldByName('COD_PRODUTO').AsString );

      Layout.FieldByName('DES_OBSERVACAO').AsString := StrReplace(StrLBReplace(FieldByName('DES_OBSERVACAO').AsString), '\n', '');
      Layout.FieldByName('DES_REDUZIDA').AsString := StrReplace(StrLBReplace(FieldByName('DES_REDUZIDA').AsString), '\n', '');
      Layout.FieldByName('DES_PRODUTO').AsString := StrReplace(StrLBReplace(FieldByName('DES_PRODUTO').AsString), '\n', '');

      codbarras := StrRetNums(Layout.FieldByName('COD_BARRA_PRINCIPAL').AsString);

      if( (codbarras = '') or (StrToFloat(codbarras) = 0)) then
         codbarras := ''
      else if ( Length(TiraZerosEsquerda(codbarras)) < 8 ) then
         codbarras := GerarPLU(codbarras)
      else
         if(not CodBarrasValido(codbarras)) then
            codbarras := '';

      Layout.FieldByName('COD_BARRA_PRINCIPAL').AsString := codbarras;

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

procedure TFrmByteCortez.GerarScriptAmarrarCEST;
begin
  with QryPrincipal do
  begin
    Close;
    Sql.Clear;

    SQL.Add('SELECT');
    SQL.Add('	NOME,');
    SQL.Add('	CEST');
    SQL.Add('FROM');
    SQL.Add('	CLASSIFICACAO');
    SQL.Add('WHERE');
    SQL.Add('  CEST IS NOT NULL');


    Open;
    First;

    while not Eof do
    begin
      try
        if Cancelar then
        Break;

        Writeln(Arquivo,'UPDATE TAB_NCM SET COD_CEST =  (SELECT COD_CEST FROM TAB_CEST WHERE NUM_CEST = '+QryPrincipal.FieldByName('CEST').AsString+' ) WHERE NUM_NCM = '+QryPrincipal.FieldByName('NOME').AsString+' ;');

      except
        On E: Exception do
        FrmProgresso.AdicionarLog(QryPrincipal.RecNo, 'E', E.Message);
      end;

    Next;
    end;
    WriteLn(Arquivo, 'COMMIT WORK;');
    Close;
  end;
end;

procedure TFrmByteCortez.GerarScriptCEST;
var
  codigo : integer;
begin

  with QryPrincipal do
  begin
    Close;
    Sql.Clear;

    SQL.Add('SELECT');
    SQL.Add('	0 AS COD_CEST,');
    SQL.Add('	CEST.CODIGO AS NUM_CEST,');
    SQL.Add('	CAST(CEST.DESCRICAO AS VARCHAR2(50)) AS DES_CEST');
    SQL.Add('FROM');
    SQL.Add('	CEST');
    SQL.Add('ORDER BY');
    SQL.Add('  NUM_CEST ASC');

    codigo := 1000;

    Open;
    First;

    while not Eof do
    begin
      try
        if Cancelar then
        Break;

        inc(codigo);
        Writeln(Arquivo,'INSERT INTO TAB_CEST(COD_CEST, NUM_CEST, DES_CEST) VALUES ( '+ IntToStr(codigo) +', '+QryPrincipal.FieldByName('NUM_CEST').AsString+', '''+QryPrincipal.FieldByName('DES_CEST').AsString+''' ) ;');

      except
        On E: Exception do
        FrmProgresso.AdicionarLog(QryPrincipal.RecNo, 'E', E.Message);
      end;

    Next;
    end;
    WriteLn(Arquivo, 'COMMIT WORK;');
    Close;
  end;

end;

procedure TFrmByteCortez.GerarSecao;
var
   TotalCount : integer;
begin
  inherited;

  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT DISTINCT');
    SQL.Add('    MARCAS.CODIGO AS COD_SECAO,');
    SQL.Add('    MARCAS.DESCRICAO AS DES_SECAO,');
    SQL.Add('    0 AS VAL_META');
    SQL.Add('FROM PRODUTOS');
    SQL.Add('LEFT JOIN MARCAS');
    SQL.Add('ON PRODUTOS.COD_MARCA = MARCAS.CODIGO    ');
    SQL.Add('');
    SQL.Add('UNION ALL');
    SQL.Add('');
    SQL.Add('SELECT DISTINCT');
    SQL.Add('    200 + MARCAS.CODIGO AS COD_SECAO,');
    SQL.Add('    MARCAS.DESCRICAO AS DES_SECAO,');
    SQL.Add('    0 AS VAL_META');
    SQL.Add('FROM TAB_PRODUTO_AUX AS PRODUTOS');
    SQL.Add('LEFT JOIN TAB_MARCAS_AUX AS MARCAS');
    SQL.Add('ON PRODUTOS.COD_MARCA = MARCAS.CODIGO');


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

procedure TFrmByteCortez.GerarSubGrupo;
var
   TotalCount : Integer;
begin
  inherited;

  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT DISTINCT');
    SQL.Add('    MARCAS.CODIGO AS COD_SECAO,');
    SQL.Add('    999 AS COD_GRUPO,');
    SQL.Add('    999 AS COD_SUB_GRUPO,');
    SQL.Add('    ''A DEFINIR'' AS DES_SUB_GRUPO,');
    SQL.Add('    0 AS VAL_META,');
    SQL.Add('    0 AS VAL_MARGEM_REF,');
    SQL.Add('    0 AS QTD_DIA_SEGURANCA,');
    SQL.Add('    ''N'' AS FLG_ALCOOLICO');
    SQL.Add('FROM PRODUTOS');
    SQL.Add('LEFT JOIN MARCAS');
    SQL.Add('ON PRODUTOS.COD_MARCA = MARCAS.CODIGO');
    SQL.Add('');
    SQL.Add('UNION ALL');
    SQL.Add('');
    SQL.Add('SELECT DISTINCT');
    SQL.Add('    200 + MARCAS.CODIGO AS COD_SECAO,');
    SQL.Add('    999 AS COD_GRUPO,');
    SQL.Add('    999 AS COD_SUB_GRUPO,');
    SQL.Add('    ''A DEFINIR'' AS DES_SUB_GRUPO,');
    SQL.Add('    0 AS VAL_META,');
    SQL.Add('    0 AS VAL_MARGEM_REF,');
    SQL.Add('    0 AS QTD_DIA_SEGURANCA,');
    SQL.Add('    ''N'' AS FLG_ALCOOLICO');
    SQL.Add('FROM TAB_PRODUTO_AUX AS PRODUTOS');
    SQL.Add('LEFT JOIN TAB_MARCAS_AUX AS MARCAS');
    SQL.Add('ON PRODUTOS.COD_MARCA = MARCAS.CODIGO');


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

procedure TFrmByteCortez.GerarTransportadora;
begin
  inherited;
  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
    SQL.Add('    TRANSPORTADORAS.ID AS COD_TRANSPORTADORA,');
    SQL.Add('    TRANSPORTADORAS.DESCRITIVO AS DES_TRANSPORTADORA,');
    SQL.Add('    TRANSPORTADORAS.CNPJ_CPF AS NUM_CGC,');
    SQL.Add('    TRANSPORTADORAS.INSCRICAO_RG AS NUM_INSC_EST,');
    SQL.Add('    TRANSPORTADORAS.LOGRADOURO || TRANSPORTADORAS.ENDERECO AS DES_ENDERECO,');
    SQL.Add('    TRANSPORTADORAS.BAIRRO AS DES_BAIRRO,');
    SQL.Add('    TRANSPORTADORAS.CIDADE AS DES_CIDADE,');
    SQL.Add('    TRANSPORTADORAS.ESTADO AS DES_SIGLA,');
    SQL.Add('    TRANSPORTADORAS.CEP AS NUM_CEP,');
    SQL.Add('    TRANSPORTADORAS.TELEFONE1 AS NUM_FONE,');
    SQL.Add('    TRANSPORTADORAS.FAX AS NUM_FAX,');
    SQL.Add('    '''' AS DES_CONTATO,');
    SQL.Add('    2 AS COD_CONDICAO,');
    SQL.Add('    30 AS NUM_CONDICAO,');
    SQL.Add('    TRANSPORTADORAS.NUMERO AS NUM_ENDERECO,');
    SQL.Add('    TRANSPORTADORAS.OBSERVACAO AS DES_OBSERVACAO,');
    SQL.Add('    8 AS COD_ENTIDADE, --');
    SQL.Add('    TRANSPORTADORAS.EMAIL AS DES_EMAIL,');
    SQL.Add('    TRANSPORTADORAS.SITE AS DES_WEB_SITE');
    SQL.Add('FROM');
    SQL.Add('    TRANSPORTADORAS');
    SQL.Add('ORDER BY');
    SQL.Add('    TRANSPORTADORAS.ID DESC');


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

procedure TFrmByteCortez.GerarVenda;
var
   TotalCount : Integer;
begin
  inherited;

  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
    SQL.Add('    VENDAS.COD_PROD AS COD_PRODUTO,');
    SQL.Add('    1 AS COD_LOJA,');
    SQL.Add('    0 AS IND_TIPO,');
    SQL.Add('    CAIXA.COD_CAIXA AS NUM_PDV,');
    SQL.Add('    VENDAS.QUANTIDADE AS QTD_TOTAL_PRODUTO,');
    SQL.Add('    ROUND(VENDAS.VALOR, 2) AS VAL_TOTAL_PRODUTO,');
    SQL.Add('    ROUND(VENDAS.PRECO_VEND, 2) AS VAL_PRECO_VENDA,');
    SQL.Add('    ROUND(VENDAS.PRECO_CUST, 2) AS VAL_CUSTO_REP,');
//    SQL.Add('    (VENDAS.PRECO_CUST * VENDAS.QUANTIDADE) AS VAL_CUSTO_REP,');
    SQL.Add('    VENDAS.DATA AS DTA_SAIDA,');
    SQL.Add('    LPAD(EXTRACT(MONTH FROM VENDAS.DATA), 2, ''00'') || EXTRACT(YEAR FROM VENDAS.DATA) AS DTA_MENSAL,');
    SQL.Add('    VENDAS.NUM_OPER AS NUM_IDENT, --');
    SQL.Add('    PRODUTO.COD_BARRA AS COD_EAN,');
    SQL.Add('    ''0000'' AS DES_HORA,');
    SQL.Add('    CAIXA.COD_CLI  AS COD_CLIENTE,');
    SQL.Add('    1 AS COD_ENTIDADE, --');
    SQL.Add('    0 AS VAL_BASE_ICMS,');
    SQL.Add('');
    SQL.Add('    CASE PRODUTO.ICMS');
    SQL.Add('        WHEN ''4,5'' THEN ''T''');
    SQL.Add('        WHEN ''F1'' THEN ''F''');
    SQL.Add('        WHEN ''FF'' THEN ''F''');
    SQL.Add('        WHEN ''II'' THEN ''I''');
    SQL.Add('        WHEN ''T01'' THEN ''T''');
    SQL.Add('        WHEN ''T02'' THEN ''T''');
    SQL.Add('        WHEN ''T03'' THEN ''T''');
    SQL.Add('        WHEN ''T04'' THEN ''T''');
    SQL.Add('    END AS DES_SITUACAO_TRIB, --');
    SQL.Add('');
    SQL.Add('    0 AS VAL_ICMS, --');
    SQL.Add('    VENDAS.NUM_OPER AS NUM_CUPOM_FISCAL, --');
    SQL.Add('    VENDAS.PRECO_VEND AS VAL_VENDA_PDV,');
    SQL.Add('');
    SQL.Add('    CASE PRODUTO.ICMS');
    SQL.Add('        WHEN ''4,5'' THEN 38');
    SQL.Add('        WHEN ''F1'' THEN 13');
    SQL.Add('        WHEN ''FF'' THEN 13');
    SQL.Add('        WHEN ''II'' THEN 1');
    SQL.Add('        WHEN ''T01'' THEN 4');
    SQL.Add('        WHEN ''T02'' THEN 2');
    SQL.Add('        WHEN ''T03'' THEN 3');
    SQL.Add('        WHEN ''T04'' THEN 5');
    SQL.Add('    END AS COD_TRIBUTACAO,');
    SQL.Add('');
    SQL.Add('    ''N'' AS FLG_CUPOM_CANCELADO, --');
    SQL.Add('    PRODUTO.CLASFISCAL AS NUM_NCM,');
    SQL.Add('    999 AS COD_TAB_SPED, --');
    SQL.Add('');
    SQL.Add('    CASE COALESCE(PRODUTO.PIS_CST, 0)');
    SQL.Add('        WHEN 0 THEN ''N''');
    SQL.Add('        WHEN 1 THEN ''N''');
    SQL.Add('        WHEN 4 THEN ''S''');
    SQL.Add('        WHEN 6 THEN ''S''');
    SQL.Add('        WHEN 7 THEN ''N''');
    SQL.Add('        WHEN 8 THEN ''S''');
    SQL.Add('    END AS FLG_NAO_PIS_COFINS,');
    SQL.Add('');
    SQL.Add('    CASE COALESCE(PRODUTO.PIS_CST, 0)');
    SQL.Add('        WHEN 0 THEN -1');
    SQL.Add('        WHEN 1 THEN -1');
    SQL.Add('        WHEN 4 THEN 1');
    SQL.Add('        WHEN 6 THEN 0');
    SQL.Add('        WHEN 7 THEN -1');
    SQL.Add('        WHEN 8 THEN 3');
    SQL.Add('    END AS TIPO_NAO_PIS_COFINS,');
    SQL.Add('');
    SQL.Add('    ''N'' AS FLG_ONLINE,');
    SQL.Add('    ''N'' AS FLG_OFERTA, --');
    SQL.Add('    0 AS COD_ASSOCIADO');
    SQL.Add('FROM VENDA AS VENDAS ');
    SQL.Add('INNER JOIN PRODUTOS AS PRODUTO');
    SQL.Add('ON VENDAS.COD_PROD = PRODUTO.COD_PROD');
    SQL.Add('INNER JOIN CAIXA');
    SQL.Add('ON VENDAS.NUM_OPER = CAIXA.NUM_OPER ');
    SQL.Add('WHERE VENDAS.DATA >= :INI');
    SQL.Add('AND VENDAS.DATA <= :FIM');
    SQL.Add('AND CAIXA.SITUACAO = ''A''');
    SQL.Add('AND VENDAS.BAIXADO = ''N''');


    ParamByName('INI').AsDate := DtpInicial.Date;
    ParamByName('FIM').AsDate := DtpFinal.Date;


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

procedure TFrmByteCortez.BtnAmarrarCestClick(Sender: TObject);
begin
  inherited;
    inherited;
  FlgGeraAmarrarCest := True;
  BtnGerar.Click;
  FlgGeraAmarrarCest := False;
end;

procedure TFrmByteCortez.btnGeraCestClick(Sender: TObject);
begin
  inherited;
  FlgGeraCest := True;
  BtnGerar.Click;
  FlgGeraCest := False;
end;

procedure TFrmByteCortez.EdtCamBancoExit(Sender: TObject);
begin
  inherited;
  CriarFB(EdtCamBanco);
end;

procedure TFrmByteCortez.GerarCest;
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

procedure TFrmByteCortez.GerarCliente;
begin
  inherited;

  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
    SQL.Add('    CLIENTE.CODIGO AS COD_CLIENTE,');
    SQL.Add('    CLIENTE.NOME AS DES_CLIENTE,');
    SQL.Add('    CLIENTE.CGC AS NUM_CGC,');
    SQL.Add('    ');
    SQL.Add('    CASE CLIENTE.TIPO');
    SQL.Add('        WHEN ''J'' THEN');
    SQL.Add('            CASE ');
    SQL.Add('                WHEN COALESCE(CLIENTE.IE, '''') = '''' THEN ''ISENTO''');
    SQL.Add('                ELSE CLIENTE.IE');
    SQL.Add('            END');
    SQL.Add('        ELSE CLIENTE.IE');
    SQL.Add('    END AS NUM_INSC_EST,');
    SQL.Add('');
    SQL.Add('    CASE');
    SQL.Add('        WHEN COALESCE(CLIENTE.ENDERECO, '''') = '''' THEN ''A CONFIRMAR''');
    SQL.Add('        ELSE CLIENTE.ENDERECO');
    SQL.Add('    END AS DES_ENDERECO,');
    SQL.Add('');
    SQL.Add('    CASE');
    SQL.Add('        WHEN COALESCE(CLIENTE.BAIRRO, '''') = '''' THEN ''A CONFIRMAR''');
    SQL.Add('        ELSE CLIENTE.BAIRRO');
    SQL.Add('    END AS DES_BAIRRO,');
    SQL.Add('');
    SQL.Add('    CLIENTE.CIDADE AS DES_CIDADE,');
    SQL.Add('    CLIENTE.UF AS DES_SIGLA,');
    SQL.Add('    CLIENTE.CEP AS NUM_CEP,');
    SQL.Add('');
    SQL.Add('    CASE');
    SQL.Add('        WHEN COALESCE(CLIENTE.FONE, '''') <> '''' THEN CLIENTE.DDD || CLIENTE.FONE');
    SQL.Add('        ELSE ''''');
    SQL.Add('    END AS NUM_FONE,');
    SQL.Add('');
    SQL.Add('    '''' AS NUM_FAX,');
    SQL.Add('    CLIENTE.NOME AS DES_CONTATO,');
    SQL.Add('    0 AS FLG_SEXO,--');
    SQL.Add('    0 AS VAL_LIMITE_CRETID,');
    SQL.Add('    CLIENTE.LIMITE AS VAL_LIMITE_CONV,');
    SQL.Add('    0 AS VAL_DEBITO,');
    SQL.Add('    0 AS VAL_RENDA,');
    SQL.Add('');
    SQL.Add('    CASE');
    SQL.Add('        WHEN CLIENTE.COD_CONV = 1 THEN 99999');
    SQL.Add('        ELSE 0');
    SQL.Add('    END AS COD_CONVENIO, --');
    SQL.Add('');
    SQL.Add('    CASE CLIENTE.OBS');
    SQL.Add('        WHEN ''BLQ'' THEN 8');
    SQL.Add('        ELSE 0');
    SQL.Add('    END AS COD_STATUS_PDV, --');
    SQL.Add('    ');
    SQL.Add('    CASE CLIENTE.TIPO');
    SQL.Add('        WHEN ''J'' THEN ''S''');
    SQL.Add('        WHEN ''E'' THEN ''S''');
    SQL.Add('        ELSE ''N''');
    SQL.Add('    END AS FLG_EMPRESA,');
    SQL.Add('');
    SQL.Add('    ''N'' AS FLG_CONVENIO,');
    SQL.Add('    ''N'' AS MICRO_EMPRESA,');
    SQL.Add('    CLIENTE.DT_INC AS DTA_CADASTRO,');
    SQL.Add('');
    SQL.Add('    CASE');
    SQL.Add('        WHEN COALESCE(CLIENTE.NUMERO, '''') = '''' THEN ''1234''');
    SQL.Add('        ELSE CLIENTE.NUMERO');
    SQL.Add('    END AS NUM_ENDERECO,');
    SQL.Add('');
    SQL.Add('    '''' AS NUM_RG, --');
    SQL.Add('');
    SQL.Add('    CASE UPPER(TRIM(CLIENTE.CIVIL))');
    SQL.Add('        WHEN ''AMASIADO'' THEN 4 ');
    SQL.Add('        WHEN ''CASADO'' THEN 1');
    SQL.Add('        WHEN ''CASADA'' THEN 1');
    SQL.Add('        WHEN ''DIVORCIADO'' THEN 3');
    SQL.Add('        WHEN ''SOLTEIRA'' THEN 0');
    SQL.Add('        WHEN ''SOLTEIRO'' THEN 0');
    SQL.Add('        WHEN ''VIUVA'' THEN 2');
    SQL.Add('        ELSE 0');
    SQL.Add('    END AS FLG_EST_CIVIL,');
    SQL.Add('');
    SQL.Add('    CASE');
    SQL.Add('        WHEN COALESCE(CLIENTE.FONE1, '''') <> '''' THEN CLIENTE.DDD || CLIENTE.FONE1');
    SQL.Add('        ELSE ''''');
    SQL.Add('    END AS NUM_CELULAR,');
    SQL.Add('');
    SQL.Add('    CLIENTE.DTATUALIZACAO AS DTA_ALTERACAO,');
    SQL.Add('    CLIENTE.LEMBRETE || '' / '' || COALESCE(REF1, ''''),');
    SQL.Add('    COALESCE(CLIENTE.LEMBRETE, '''') || '' / '' || COALESCE(CLIENTE.REF1, '''') || '' / '' || COALESCE(CLIENTE.REF2, '''') || '' / '' || COALESCE(CLIENTE.REF3, '''') || '' /EMPRESA: '' || COALESCE(CLIENTE.EMPRESA, '''') || '' /LEMBRETE'' || COALESCE');
    SQL.Add('    (CLIENTE.LEMBRETE, '''') AS DES_OBSERVACAO, --');
    SQL.Add('    CLIENTE.END_COMPLEMENTO AS DES_COMPLEMENTO,');
    SQL.Add('    CLIENTE.EMAIL AS DES_EMAIL,');
    SQL.Add('    CLIENTE.NOME AS DES_FANTASIA, --');
    SQL.Add('    CLIENTE.NASC AS DTA_NASCIMENTO,');
    SQL.Add('    CLIENTE.PAI AS DES_PAI,');
    SQL.Add('    CLIENTE.MAE AS DES_MAE,');
    SQL.Add('    CLIENTE.CONJUGE AS DES_CONJUGE,');
    SQL.Add('    CLIENTE.CPF_CO AS NUM_CPF_CONJUGE,');
    SQL.Add('    0 AS VAL_DEB_CONV,');
    SQL.Add('');
    SQL.Add('    CASE CLIENTE.OBS');
    SQL.Add('        WHEN ''INT'' THEN ''S''');
    SQL.Add('        ELSE ''N''');
    SQL.Add('    END AS INATIVO, --');
    SQL.Add('');
    SQL.Add('    '''' AS DES_MATRICULA,');
    SQL.Add('    ''N'' AS NUM_CGC_ASSOCIADO,');
    SQL.Add('    ''N'' AS FLG_PROD_RURAL,');
    SQL.Add('');
    SQL.Add('    CASE CLIENTE.OBS');
    SQL.Add('        WHEN ''BLQ'' THEN 8');
    SQL.Add('        ELSE 0');
    SQL.Add('    END AS COD_STATUS_PDV_CONV, ');
    SQL.Add('');
    SQL.Add('    ''S'' AS FLG_ENVIA_CODIGO,');
    SQL.Add('    CLIENTE.NAS_CO AS DTA_NASC_CONJUGE,');
    SQL.Add('    0 AS COD_CLASSIF');
    SQL.Add('FROM CADCLI AS CLIENTE');
    SQL.Add('');
    SQL.Add('');
    SQL.Add('UNION ALL');
    SQL.Add('');
    SQL.Add('SELECT');
    SQL.Add('    2000 + CLIENTE.CODIGO AS COD_CLIENTE,');
    SQL.Add('    CLIENTE.NOME AS DES_CLIENTE,');
    SQL.Add('    CLIENTE.CGC AS NUM_CGC,');
    SQL.Add('    ');
    SQL.Add('    CASE CLIENTE.TIPO');
    SQL.Add('        WHEN ''J'' THEN');
    SQL.Add('            CASE ');
    SQL.Add('                WHEN COALESCE(CLIENTE.IE, '''') = '''' THEN ''ISENTO''');
    SQL.Add('                ELSE CLIENTE.IE');
    SQL.Add('            END');
    SQL.Add('        ELSE CLIENTE.IE');
    SQL.Add('    END AS NUM_INSC_EST,');
    SQL.Add('');
    SQL.Add('    CASE');
    SQL.Add('        WHEN COALESCE(CLIENTE.ENDERECO, '''') = '''' THEN ''A CONFIRMAR''');
    SQL.Add('        ELSE CLIENTE.ENDERECO');
    SQL.Add('    END AS DES_ENDERECO,');
    SQL.Add('');
    SQL.Add('    CASE');
    SQL.Add('        WHEN COALESCE(CLIENTE.BAIRRO, '''') = '''' THEN ''A CONFIRMAR''');
    SQL.Add('        ELSE CLIENTE.BAIRRO');
    SQL.Add('    END AS DES_BAIRRO,');
    SQL.Add('');
    SQL.Add('    CLIENTE.CIDADE AS DES_CIDADE,');
    SQL.Add('    CLIENTE.UF AS DES_SIGLA,');
    SQL.Add('    CLIENTE.CEP AS NUM_CEP,');
    SQL.Add('');
    SQL.Add('    CASE');
    SQL.Add('        WHEN COALESCE(CLIENTE.FONE, '''') <> '''' THEN CLIENTE.DDD || CLIENTE.FONE');
    SQL.Add('        ELSE ''''');
    SQL.Add('    END AS NUM_FONE,');
    SQL.Add('');
    SQL.Add('    '''' AS NUM_FAX,');
    SQL.Add('    CLIENTE.NOME AS DES_CONTATO,');
    SQL.Add('    0 AS FLG_SEXO,--');
    SQL.Add('    0 AS VAL_LIMITE_CRETID,');
    SQL.Add('    CLIENTE.LIMITE AS VAL_LIMITE_CONV,');
    SQL.Add('    0 AS VAL_DEBITO,');
    SQL.Add('    0 AS VAL_RENDA,');
    SQL.Add('');
    SQL.Add('    CASE');
    SQL.Add('        WHEN CLIENTE.COD_CONV = 1 THEN 99999');
    SQL.Add('        ELSE 0');
    SQL.Add('    END AS COD_CONVENIO, --');
    SQL.Add('');
    SQL.Add('    CASE CLIENTE.OBS');
    SQL.Add('        WHEN ''BLQ'' THEN 8');
    SQL.Add('        ELSE 0');
    SQL.Add('    END AS COD_STATUS_PDV, --');
    SQL.Add('    ');
    SQL.Add('    CASE CLIENTE.TIPO');
    SQL.Add('        WHEN ''J'' THEN ''S''');
    SQL.Add('        WHEN ''E'' THEN ''S''');
    SQL.Add('        ELSE ''N''');
    SQL.Add('    END AS FLG_EMPRESA,');
    SQL.Add('');
    SQL.Add('    ''N'' AS FLG_CONVENIO,');
    SQL.Add('    ''N'' AS MICRO_EMPRESA,');
    SQL.Add('    CLIENTE.DT_INC AS DTA_CADASTRO,');
    SQL.Add('');
    SQL.Add('    CASE');
    SQL.Add('        WHEN COALESCE(CLIENTE.NUMERO, '''') = '''' THEN ''1234''');
    SQL.Add('        ELSE CLIENTE.NUMERO');
    SQL.Add('    END AS NUM_ENDERECO,');
    SQL.Add('');
    SQL.Add('    '''' AS NUM_RG, --');
    SQL.Add('');
    SQL.Add('    CASE UPPER(TRIM(CLIENTE.CIVIL))');
    SQL.Add('        WHEN ''AMASIADO'' THEN 4 ');
    SQL.Add('        WHEN ''CASADO'' THEN 1');
    SQL.Add('        WHEN ''CASADA'' THEN 1');
    SQL.Add('        WHEN ''DIVORCIADO'' THEN 3');
    SQL.Add('        WHEN ''SOLTEIRA'' THEN 0');
    SQL.Add('        WHEN ''SOLTEIRO'' THEN 0');
    SQL.Add('        WHEN ''VIUVA'' THEN 2');
    SQL.Add('        ELSE 0');
    SQL.Add('    END AS FLG_EST_CIVIL,');
    SQL.Add('');
    SQL.Add('    CASE');
    SQL.Add('        WHEN COALESCE(CLIENTE.FONE1, '''') <> '''' THEN CLIENTE.DDD || CLIENTE.FONE1');
    SQL.Add('        ELSE ''''');
    SQL.Add('    END AS NUM_CELULAR,');
    SQL.Add('');
    SQL.Add('    CLIENTE.DTATUALIZACAO AS DTA_ALTERACAO,');
    SQL.Add('    CLIENTE.LEMBRETE || '' / '' || COALESCE(CLIENTE.REF1, ''''),');
    SQL.Add('    COALESCE(CLIENTE.LEMBRETE, '''') || '' / '' || COALESCE(CLIENTE.REF1, '''') || '' / '' || COALESCE(CLIENTE.REF2, '''') || '' / '' || COALESCE(CLIENTE.REF3, '''') || '' /EMPRESA: '' || COALESCE(CLIENTE.EMPRESA, '''') || '' /LEMBRETE'' || COALESCE');
    SQL.Add('    (CLIENTE.LEMBRETE, '''') AS DES_OBSERVACAO, --');
    SQL.Add('    CLIENTE.END_COMPLEMENTO AS DES_COMPLEMENTO,');
    SQL.Add('    CLIENTE.EMAIL AS DES_EMAIL,');
    SQL.Add('    CLIENTE.NOME AS DES_FANTASIA, --');
    SQL.Add('    CLIENTE.NASC AS DTA_NASCIMENTO,');
    SQL.Add('    CLIENTE.PAI AS DES_PAI,');
    SQL.Add('    CLIENTE.MAE AS DES_MAE,');
    SQL.Add('    CLIENTE.CONJUGE AS DES_CONJUGE,');
    SQL.Add('    CLIENTE.CPF_CO AS NUM_CPF_CONJUGE,');
    SQL.Add('    0 AS VAL_DEB_CONV,');
    SQL.Add('');
    SQL.Add('    CASE CLIENTE.OBS');
    SQL.Add('        WHEN ''INT'' THEN ''S''');
    SQL.Add('        ELSE ''N''');
    SQL.Add('    END AS INATIVO, --');
    SQL.Add('');
    SQL.Add('    '''' AS DES_MATRICULA,');
    SQL.Add('    ''N'' AS NUM_CGC_ASSOCIADO,');
    SQL.Add('    ''N'' AS FLG_PROD_RURAL,');
    SQL.Add('');
    SQL.Add('    CASE CLIENTE.OBS');
    SQL.Add('        WHEN ''BLQ'' THEN 8');
    SQL.Add('        ELSE 0');
    SQL.Add('    END AS COD_STATUS_PDV_CONV, ');
    SQL.Add('');
    SQL.Add('    ''S'' AS FLG_ENVIA_CODIGO,');
    SQL.Add('    CLIENTE.NAS_CO AS DTA_NASC_CONJUGE,');
    SQL.Add('    0 AS COD_CLASSIF');
    SQL.Add('FROM TAB_CLIENTE_AUX AS CLIENTE');
    SQL.Add('LEFT JOIN CADCLI');
    SQL.Add('ON COALESCE(CLIENTE.CGC, '''') <> ''''');
    SQL.Add('AND REPLACE(REPLACE(REPLACE(COALESCE(CLIENTE.CGC, ''''), ''.'', ''''), ''-'', ''''), ''/'', '''') = REPLACE(REPLACE(REPLACE(COALESCE(CADCLI.CGC, ''''), ''.'', ''''), ''-'', ''''), ''/'', '''')');
    SQL.Add('WHERE CADCLI.CODIGO IS NULL');


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

      if StrRetNums(Layout.FieldByName('NUM_RG').AsString) = '' then
        Layout.FieldByName('NUM_RG').AsString := ''
      else
        Layout.FieldByName('NUM_RG').AsString := StrRetNums(Layout.FieldByName('NUM_RG').AsString);

      Layout.FieldByName('NUM_CEP').AsString := StrRetNums(Layout.FieldByName('NUM_CEP').AsString);

      if QryPrincipal.FieldByName('NUM_INSC_EST').AsString <> 'ISENTO' then
         Layout.FieldByName('NUM_INSC_EST').AsString := StrRetNums(Layout.FieldByName('NUM_INSC_EST').AsString);

      Layout.FieldByName('DTA_NASCIMENTO').AsDateTime := FieldByName('DTA_NASCIMENTO').AsDateTime;
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

procedure TFrmByteCortez.GerarCodigoBarras;
var
 count : Integer;
 cod_antigo, codbarras : string;
begin
  inherited;
  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
    SQL.Add('    PRODUTO.COD_PROD AS COD_PRODUTO,');
    SQL.Add('    PRODUTO.COD_BARRA AS COD_EAN');
    SQL.Add('FROM');
    SQL.Add('    PRODUTOS AS PRODUTO');
    SQL.Add('');
    SQL.Add('UNION ALL');
    SQL.Add('');
    SQL.Add('SELECT');
    SQL.Add('    100000 + PRODUTO.COD_PROD AS COD_PRODUTO,');
    SQL.Add('    PRODUTO.COD_BARRA AS COD_EAN');
    SQL.Add('FROM');
    SQL.Add('    TAB_PRODUTO_AUX AS PRODUTO');
    SQL.Add('LEFT JOIN');
    SQL.Add('    PRODUTOS');
    SQL.Add('ON PRODUTO.COD_BARRA = PRODUTOS.COD_BARRA');
    SQL.Add('WHERE PRODUTOS.COD_PROD IS NULL');

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

      Layout.FieldByName('COD_PRODUTO').AsString := GerarPLU( Layout.FieldByName('COD_PRODUTO').AsString );

      codbarras := StrRetNums(Layout.FieldByName('COD_EAN').AsString);

      if( (codbarras = '') or (StrToFloat(codbarras) = 0)) then
         codbarras := ''
      else if ( Length(TiraZerosEsquerda(codbarras)) < 8 ) then
         codbarras := GerarPLU(codbarras)
      else
         if(not CodBarrasValido(codbarras)) then
            codbarras := '';

      Layout.FieldByName('COD_EAN').AsString := codbarras;

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;
end;

procedure TFrmByteCortez.GerarComposicao;
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

procedure TFrmByteCortez.GerarCondPagCli;
begin
  inherited;

  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
    SQL.Add('    CLIENTE.CODIGO AS COD_CLIENTE,');
    SQL.Add('    CLIENTE.DATAPAGTO AS NUM_CONDICAO,');
    SQL.Add('    2 AS COD_CONDICAO,');
    SQL.Add('    1 AS COD_ENTIDADE');
    SQL.Add('FROM');
    SQL.Add('    CADCLI AS CLIENTE');
    SQL.Add('');
    SQL.Add('UNION ALL');
    SQL.Add('');
    SQL.Add('SELECT');
    SQL.Add('    2000 + CLIENTE.CODIGO AS COD_CLIENTE,');
    SQL.Add('    CLIENTE.DATAPAGTO AS NUM_CONDICAO,');
    SQL.Add('    2 AS COD_CONDICAO,');
    SQL.Add('    1 AS COD_ENTIDADE');
    SQL.Add('FROM TAB_CLIENTE_AUX AS CLIENTE');
    SQL.Add('LEFT JOIN CADCLI');
    SQL.Add('ON COALESCE(CLIENTE.CGC, '''') <> ''''');
    SQL.Add('AND REPLACE(REPLACE(REPLACE(COALESCE(CLIENTE.CGC, ''''), ''.'', ''''), ''-'', ''''), ''/'', '''') = REPLACE(REPLACE(REPLACE(COALESCE(CADCLI.CGC, ''''), ''.'', ''''), ''-'', ''''), ''/'', '''')');
    SQL.Add('WHERE CADCLI.CODIGO IS NULL');


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

procedure TFrmByteCortez.GerarCondPagForn;
begin
  inherited;

  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
    SQL.Add('    FORNECEDOR.CODIGO AS COD_FORNECEDOR,');
    SQL.Add('    30 AS NUM_CONDICAO,');
    SQL.Add('    2 AS COD_CONDICAO,');
    SQL.Add('    1 AS COD_ENTIDADE,');
    SQL.Add('    '''' AS NUM_CGC');
    SQL.Add('FROM');
    SQL.Add('    FORNEC AS FORNECEDOR');
    SQL.Add('');
    SQL.Add('UNION ALL');
    SQL.Add('');
    SQL.Add('SELECT');
    SQL.Add('    1000 + FORNECEDOR.CODIGO AS COD_FORNECEDOR,');
    SQL.Add('    30 AS NUM_CONDICAO,');
    SQL.Add('    2 AS COD_CONDICAO,');
    SQL.Add('    1 AS COD_ENTIDADE,');
    SQL.Add('    '''' AS NUM_CGC');
    SQL.Add('FROM TAB_FORNECEDOR_AUX AS FORNECEDOR');
    SQL.Add('LEFT JOIN FORNEC');
    SQL.Add('ON COALESCE(FORNEC.CGC, '''') <> ''''');
    SQL.Add('AND REPLACE(REPLACE(REPLACE(COALESCE(FORNECEDOR.CGC, ''''), ''.'', ''''), ''-'', ''''), ''/'', '''') = REPLACE(REPLACE(REPLACE(COALESCE(FORNEC.CGC, ''''), ''.'', ''''), ''-'', ''''), ''/'', '''')');
    SQL.Add('WHERE FORNEC.CODIGO IS NULL');


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

procedure TFrmByteCortez.GerarDecomposicao;
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

procedure TFrmByteCortez.GerarDivisaoForn;
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

procedure TFrmByteCortez.GerarFinanceiro(Tipo, Situacao: Integer);
begin
  inherited;
  if Tipo = 1 then
    GerarFinanceiroPagar(IntToStr(Situacao));

  if Tipo = 2 then
    GerarFinanceiroReceber(IntToStr(Situacao));

  if Tipo = 3 then
    GerarFinanceiroReceberCartao;

end;

procedure TFrmByteCortez.GerarFinanceiroPagar(Aberto: String);
var
   TotalCount : Integer;
   cgc: string;
begin
  inherited;
  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
    SQL.Add('    1 AS TIPO_PARCEIRO,');
    SQL.Add('    PAGAR.COD_FORNEC AS COD_PARCEIRO,');
    SQL.Add('    0 AS TIPO_CONTA,');
    SQL.Add('    8 AS COD_ENTIDADE,');
    SQL.Add('    PAGAR.NR_DOCUMENTO AS NUM_DOCTO,');
    SQL.Add('    999 AS COD_BANCO,');
    SQL.Add('    0 AS DES_BANCO,');
    SQL.Add('    PAGAR.EMISSAO AS DTA_EMISSAO,');
    SQL.Add('    PAGAR.VENCIMENTO AS DTA_VENCIMENTO,');
    SQL.Add('    PAGAR.VALOR AS VAL_PARCELA,');
    SQL.Add('    COALESCE(PAGAR.JUROS, 0) AS VAL_JUROS,');
    SQL.Add('    COALESCE(PAGAR.DESCONTO, 0) AS VAL_DESCONTO,');
    SQL.Add('');
    SQL.Add('    CASE ');
    SQL.Add('        WHEN PAGAR.BAIXADO = ''S'' THEN ''S''');
    SQL.Add('        ELSE ''N''');
    SQL.Add('    END AS FLG_QUITADO,');
    SQL.Add('');
    SQL.Add('    CASE');
    SQL.Add('        WHEN PAGAR.BAIXADO = ''S'' THEN LPAD(EXTRACT(DAY FROM PAGAR.DT_PGTO), 2, ''0'') || ''/'' || LPAD(EXTRACT(MONTH FROM PAGAR.DT_PGTO), 2, ''0'')|| ''/'' || LPAD(EXTRACT(YEAR FROM PAGAR.DT_PGTO), 4, ''0'')');
    SQL.Add('        ELSE ''''');
    SQL.Add('    END AS DTA_QUITADA,    ');
    SQL.Add('');
    SQL.Add('    998 AS COD_CATEGORIA,');
    SQL.Add('    998 AS COD_SUBCATEGORIA,');
    SQL.Add('    ');
    SQL.Add('    SUBSTRING(PAGAR.PRESTACAO FROM 1 FOR (POSITION(''/'' IN PAGAR.PRESTACAO) - 1)) AS NUM_PARCELA,');
    SQL.Add('');
    SQL.Add('    PARCELAS.QTD_PARCELA AS QTD_PARCELA,');
    SQL.Add('    1 AS COD_LOJA,');
    SQL.Add('    FORNEC.CGC AS NUM_CGC,');
    SQL.Add('    0 AS NUM_BORDERO,');
    SQL.Add('    PAGAR.NR_DOCUMENTO AS NUM_NF,');
    SQL.Add('    '''' AS NUM_SERIE_NF,');
    SQL.Add('    PARCELAS.VAL_TOTAL_NF AS VAL_TOTAL_NF,');
    SQL.Add('    ''HISTORICO: '' || PAGAR.HISTORICO || '' OBSERVACAO:'' || PAGAR.OBSERVACAO AS DES_OBSERVACAO,');
    SQL.Add('    0 AS NUM_PDV,');
    SQL.Add('    PAGAR.NUM_OPER AS NUM_CUPOM_FISCAL,');
    SQL.Add('    0 AS COD_MOTIVO,');
    SQL.Add('    0 AS COD_CONVENIO,');
    SQL.Add('    0 AS COD_BIN,');
    SQL.Add('    '''' AS DES_BANDEIRA,');
    SQL.Add('    '''' AS DES_REDE_TEF,');
    SQL.Add('    0 AS VAL_RETENCAO,');
    SQL.Add('    0 AS COD_CONDICAO,');
    SQL.Add('');
    SQL.Add('    CASE');
    SQL.Add('        WHEN PAGAR.BAIXADO = ''S'' THEN LPAD(EXTRACT(DAY FROM PAGAR.DT_PGTO), 2, ''0'') || ''/'' || LPAD(EXTRACT(MONTH FROM PAGAR.DT_PGTO), 2, ''0'')|| ''/'' || LPAD(EXTRACT(YEAR FROM PAGAR.DT_PGTO), 4, ''0'')');
    SQL.Add('        ELSE ''''');
    SQL.Add('    END AS DTA_PAGTO,');
    SQL.Add('');
    SQL.Add('    PAGAR.EMISSAO AS DTA_ENTRADA,    ');
    SQL.Add('    '''' AS NUM_NOSSO_NUMERO,');
    SQL.Add('    '''' AS COD_BARRA,');
    SQL.Add('    ''N'' AS FLG_BOLETO_EMIT,');
    SQL.Add('    '''' AS NUM_CGC_CPF_TITULAR,');
    SQL.Add('    '''' AS DES_TITULAR,');
    SQL.Add('    30 AS NUM_CONDICAO,');
    SQL.Add('    0 AS VAL_CREDITO,');
    SQL.Add('    ''999'' AS COD_BANCO_PGTO,');
    SQL.Add('    ''PAGTO-1'' AS DES_CC,');
    SQL.Add('    0 AS COD_BANDEIRA,');
    SQL.Add('    '''' AS DTA_PRORROGACAO,');
    SQL.Add('    1 AS NUM_SEQ_FIN,');
    SQL.Add('    0 AS COD_COBRANCA,');
    SQL.Add('    '''' AS DTA_COBRANCA,');
    SQL.Add('    ''N'' AS FLG_ACEITE,');
    SQL.Add('    0 AS TIPO_ACEITE');
    SQL.Add('FROM     ');
    SQL.Add('    PAGAR ');
    SQL.Add('LEFT JOIN');
    SQL.Add('    (');
    SQL.Add('        SELECT');
    SQL.Add('            NR_DOCUMENTO,');
    SQL.Add('            COD_FORNEC,');
    SQL.Add('            COUNT(*) AS QTD_PARCELA,');
    SQL.Add('            SUM(PAGAR.VALOR) AS VAL_TOTAL_NF');
    SQL.Add('        FROM');
    SQL.Add('            PAGAR');
    SQL.Add('        WHERE');
    SQL.Add('            COALESCE(PAGAR.NR_DOCUMENTO, '''') <> ''''');
    SQL.Add('        GROUP by');
    SQL.Add('            NR_DOCUMENTO,');
    SQL.Add('            COD_FORNEC');
    SQL.Add('    ) AS PARCELAS');
    SQL.Add('ON');
    SQL.Add('    PAGAR.NR_DOCUMENTO = PARCELAS.NR_DOCUMENTO');
    SQL.Add('AND');
    SQL.Add('    PAGAR.COD_FORNEC = PARCELAS.COD_FORNEC');
    SQL.Add('LEFT JOIN');
    SQL.Add('    FORNEC');
    SQL.Add('ON');
    SQL.Add('    PAGAR.COD_FORNEC = FORNEC.CODIGO');
    SQL.Add('WHERE');

    if Aberto = '1' then
    begin
        SQL.Add('    PAGAR.BAIXADO <> ''S''');
    end
    else
    begin
        SQL.Add('    PAGAR.BAIXADO = ''S''');
        SQL.Add('AND');
        SQL.Add('    PAGAR.DT_PGTO >= :INI ');
        SQL.Add('AND');
        SQL.Add('    PAGAR.DT_PGTO <= :FIM ');
        ParamByName('INI').AsDate := DtpInicial.Date;
        ParamByName('FIM').AsDate := DtpFinal.Date;
    end;

    SQL.Add('ORDER BY');
    SQL.Add('    PAGAR.NR_DOCUMENTO,');
    SQL.Add('    PAGAR.COD_FORNEC,');
    SQL.Add('    PAGAR.EMISSAO     ');

    Open;
    First;

    if( Aberto = '1' ) then
      TotalCount := SetCountTotal(SQL.Text)
    else
      TotalCount := SetCountTotal(SQL.Text, ParamByName('INI').AsString, ParamByName('FIM').AsString );


    NumLinha := 0;

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);

      Layout.SetValues(QryPrincipal, NumLinha, TotalCount);

      if( CbxLoja.Text = '2' ) then
      begin
         cgc := StrRetNums(Layout.FieldByName('NUM_CGC').AsString);
         if( Length(cgc) > 11 ) then begin
           if( not CNPJEValido(cgc) ) then
            Layout.FieldByName('COD_PARCEIRO').AsInteger := Layout.FieldByName('COD_PARCEIRO').AsInteger + 1000
           else
            Layout.FieldByName('COD_PARCEIRO').AsInteger := 0;
         end
         else
         begin
            if( not CPFEValido(cgc) ) then
               Layout.FieldByName('COD_PARCEIRO').AsInteger := Layout.FieldByName('COD_PARCEIRO').AsInteger + 1000
            else
               Layout.FieldByName('COD_PARCEIRO').AsInteger := 0;
         end;
      end;

      Layout.FieldByName('NUM_CGC').AsString := StrRetNums(Layout.FieldByName('NUM_CGC').AsString);

      Layout.FieldByName('DTA_ENTRADA').AsString:= FormatDateTime('dd/mm/yyyy',QryPrincipal.FieldByName('DTA_ENTRADA').AsDateTime);
      Layout.FieldByName('DTA_EMISSAO').AsString:= FormatDateTime('dd/mm/yyyy',QryPrincipal.FieldByName('DTA_EMISSAO').AsDateTime);
      Layout.FieldByName('DTA_VENCIMENTO').AsString:= FormatDateTime('dd/mm/yyyy',QryPrincipal.FieldByName('DTA_VENCIMENTO').AsDateTime);

      Layout.FieldByName('NUM_NF').AsString := StrRetNums(Layout.FieldByName('NUM_NF').AsString);

      if Aberto = '1' then
      begin
        Layout.FieldByName('DTA_QUITADA').AsString := '';
        Layout.FieldByName('DTA_PAGTO').AsString := '';
      end
      else
      begin
        Layout.FieldByName('DTA_QUITADA').AsString:= FormatDateTime('dd/mm/yyyy',QryPrincipal.FieldByName('DTA_QUITADA').AsDateTime);
        Layout.FieldByName('DTA_PAGTO').AsString:= FormatDateTime('dd/mm/yyyy',QryPrincipal.FieldByName('DTA_PAGTO').AsDateTime);
      end;

      Layout.FieldByName('DES_OBSERVACAO').AsString := StrReplace(StrLBReplace(FieldByName('DES_OBSERVACAO').AsString), '\n', '');

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;
end;

procedure TFrmByteCortez.GerarFinanceiroReceber(Aberto: String);
var
   TotalCount : Integer;
   cgc : string;
begin
  inherited;
  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
    SQL.Add('    0 AS TIPO_PARCEIRO,');
    SQL.Add('    RECEBER.COD_CLI AS COD_PARCEIRO,');
    SQL.Add('    1 AS TIPO_CONTA,');
    SQL.Add('    1 AS COD_ENTIDADE,');
    SQL.Add('    RECEBER.SEQ AS NUM_DOCTO,');
    SQL.Add('    999 AS COD_BANCO,');
    SQL.Add('    '''' AS DES_BANCO,');
    SQL.Add('    RECEBER.DT_VENDA AS DTA_EMISSAO,');
    SQL.Add('    RECEBER.VENCIMENTO AS DTA_VENCIMENTO,');
    SQL.Add('    CASE');
    SQL.Add('        WHEN RECEBER.BAIXADO = ''S'' THEN RECEBER.VALOR');
    SQL.Add('        ELSE RECEBER.VR_ATUAL');
    SQL.Add('    END AS VAL_PARCELA,');
    SQL.Add('    0 AS VAL_JUROS,');
    SQL.Add('    0 AS VAL_DESCONTO,');
    SQL.Add('');
    SQL.Add('    CASE ');
    SQL.Add('        WHEN RECEBER.BAIXADO = ''S'' THEN ''S''');
    SQL.Add('        ELSE ''N''');
    SQL.Add('    END AS FLG_QUITADO,');
    SQL.Add('    ');
    SQL.Add('    CASE ');
    SQL.Add('        WHEN RECEBER.BAIXADO = ''S'' THEN LPAD(EXTRACT(DAY FROM RECEBER.DT_PAGAMENTO), 2, ''0'') || ''/'' || LPAD(EXTRACT(MONTH FROM RECEBER.DT_PAGAMENTO), 2, ''0'')|| ''/'' || LPAD(EXTRACT(YEAR FROM RECEBER.DT_PAGAMENTO), 4, ''0'')');
    SQL.Add('        ELSE ''''');
    SQL.Add('    END AS DTA_QUITADA,');
    SQL.Add('');
    SQL.Add('    ''997'' AS COD_CATEGORIA,');
    SQL.Add('    ''997'' AS COD_SUBCATEGORIA,');
    SQL.Add('    1 AS NUM_PARCELA,');
    SQL.Add('    1 AS QTD_PARCELA,');
    SQL.Add('    1 AS COD_LOJA,');
    SQL.Add('    CLIENTE.CGC AS NUM_CGC,');
    SQL.Add('    0 AS NUM_BORDERO,');
    SQL.Add('    RECEBER.N_FISCAL AS NUM_NF,');
    SQL.Add('    '''' AS NUM_SERIE_NF,');
    SQL.Add('    RECEBER.VALOR AS VAL_TOTAL_NF,');
    SQL.Add('    '''' AS DES_OBSERVACAO,');
    SQL.Add('    COALESCE(CAIXA.COD_CAIXA, 1) AS NUM_PDV,');
    SQL.Add('    RECEBER.N_FISCAL AS NUM_CUPOM_FISCAL,');
    SQL.Add('    0 AS COD_MOTIVO,');
    SQL.Add('    0 AS COD_CONVENIO,');
    SQL.Add('    0 AS COD_BIN,');
    SQL.Add('    '''' AS DES_BANDEIRA,');
    SQL.Add('    '''' AS DES_REDE_TEF,');
    SQL.Add('    0 AS VAL_RETENCAO,');
    SQL.Add('    0 AS COD_CONDICAO,');
    SQL.Add('');
    SQL.Add('    CASE ');
    SQL.Add('        WHEN RECEBER.BAIXADO = ''S'' THEN LPAD(EXTRACT(DAY FROM RECEBER.DT_PAGAMENTO), 2, ''0'') || ''/'' || LPAD(EXTRACT(MONTH FROM RECEBER.DT_PAGAMENTO), 2, ''0'')|| ''/'' || LPAD(EXTRACT(YEAR FROM RECEBER.DT_PAGAMENTO), 4, ''0'')');
    SQL.Add('        ELSE ''''');
    SQL.Add('    END AS DTA_PAGTO,');
    SQL.Add('');
    SQL.Add('    RECEBER.DT_VENDA AS DTA_ENTRADA,');
    SQL.Add('    '''' AS NUM_NOSSO_NUMERO,');
    SQL.Add('    '''' AS COD_BARRA,');
    SQL.Add('    ''N'' AS FLG_BOLETO_EMIT,');
    SQL.Add('    '''' AS NUM_CGC_CPF_TITULAR,');
    SQL.Add('    '''' AS DES_TITULAR,');
    SQL.Add('    30 AS NUM_CONDICAO,');
    SQL.Add('    0 AS VAL_CREDITO,');
    SQL.Add('    999 AS COD_BANCO_PGTO,');
    SQL.Add('    ''RECEBTO-1'' AS DES_CC,');
    SQL.Add('    0 AS COD_BANDEIRA,');
    SQL.Add('    '''' AS DTA_PRORROGACAO,');
    SQL.Add('    1 AS NUM_SEQ_FIN,');
    SQL.Add('    0 AS COD_COBRANCA,');
    SQL.Add('    '''' AS DTA_COBRANCA,');
    SQL.Add('    ''N'' AS FLG_ACEITE,');
    SQL.Add('    0 AS TIPO_ACEITE');
    SQL.Add('FROM RECEBER');
    SQL.Add('LEFT JOIN CAIXA');
    SQL.Add('ON RECEBER.NUM_OPER = CAIXA.NUM_OPER');
    SQL.Add('LEFT JOIN CADCLI AS CLIENTE');
    SQL.Add('ON RECEBER.COD_CLI = CLIENTE.CODIGO');
    SQL.Add('WHERE RECEBER.VALOR > 0');

    if Aberto = '1' then
    begin
      SQL.Add('AND RECEBER.BAIXADO <> ''S''');
    end
    else
    begin
      SQL.Add('AND RECEBER.DT_PAGAMENTO >= :INI ');
      SQL.Add('AND RECEBER.DT_PAGAMENTO <= :FIM ');
      SQL.Add('AND RECEBER.BAIXADO = ''S'' ');

      ParamByName('INI').AsDate := DtpInicial.Date;
      ParamByName('FIM').AsDate := DtpFinal.Date;
    end;

    Open;

    First;

    if( Aberto = '1' ) then
      TotalCount := SetCountTotal(SQL.Text)
    else
      TotalCount := SetCountTotal(SQL.Text, ParamByName('INI').AsString, ParamByName('FIM').AsString );

    Open;

    First;

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);

      Layout.SetValues(QryPrincipal, NumLinha, TotalCount);

      if( CbxLoja.Text = '2' ) then
      begin
         cgc := StrRetNums(Layout.FieldByName('NUM_CGC').AsString);
         if( Length(cgc) > 11 ) then begin
           if( not CNPJEValido(cgc) ) then
            Layout.FieldByName('COD_PARCEIRO').AsInteger := Layout.FieldByName('COD_PARCEIRO').AsInteger + 2000
           else
            Layout.FieldByName('COD_PARCEIRO').AsInteger := 0;
         end
         else
         begin
            if( not CPFEValido(cgc) ) then
               Layout.FieldByName('COD_PARCEIRO').AsInteger := Layout.FieldByName('COD_PARCEIRO').AsInteger + 2000
            else
               Layout.FieldByName('COD_PARCEIRO').AsInteger := 0;
         end;
      end;

      Layout.FieldByName('NUM_CGC').AsString := StrRetNums(Layout.FieldByName('NUM_CGC').AsString);

      Layout.FieldByName('DTA_ENTRADA').AsString:= FormatDateTime('dd/mm/yyyy',QryPrincipal.FieldByName('DTA_ENTRADA').AsDateTime);
      Layout.FieldByName('DTA_EMISSAO').AsString:= FormatDateTime('dd/mm/yyyy',QryPrincipal.FieldByName('DTA_EMISSAO').AsDateTime);
      Layout.FieldByName('DTA_VENCIMENTO').AsString:= FormatDateTime('dd/mm/yyyy',QryPrincipal.FieldByName('DTA_VENCIMENTO').AsDateTime);

      if Aberto = '1' then
      begin
        Layout.FieldByName('DTA_QUITADA').AsString := '';
        Layout.FieldByName('DTA_PAGTO').AsString := '';
      end
      else
      begin
        Layout.FieldByName('DTA_QUITADA').AsString := FormatDateTime('dd/mm/yyyy',QryPrincipal.FieldByName('DTA_QUITADA').AsDateTime);
        Layout.FieldByName('DTA_PAGTO').AsString := FormatDateTime('dd/mm/yyyy',QryPrincipal.FieldByName('DTA_PAGTO').AsDateTime);
      end;

      Layout.FieldByName('DES_OBSERVACAO').AsString := StrReplace(StrLBReplace(FieldByName('DES_OBSERVACAO').AsString), '\n', '');

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;
end;

procedure TFrmByteCortez.GerarFinanceiroReceberCartao;
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
//    SQL.Add('    ''COBRANÇA: '' || RECEBER.DATACOB || '' | 1 DEVOL: '' || RECEBER.DEVOLUCAOA || '' | 2 DEVOL : '' || RECEBER.DEVOLUCAOB || '' | ''  || RECEBER.OBSERVACAO AS DES_OBSERVACAO,');
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
    SQL.Add('''COBRANÇA: '' || RECEBER.DATACOB || '' | 1 DEVOL: '' || RECEBER.DEVOLUCAOA || '' | 2 DEVOL : '' || RECEBER.DEVOLUCAOB || '' | ''  || RECEBER.OBSERVACAO AS DES_OBSERVACAO,');
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

procedure TFrmByteCortez.GerarFornecedor;
var
   observacao, email : string;
begin
  inherited;

  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
    SQL.Add('    FORNECEDOR.CODIGO AS COD_FORNECEDOR,');
    SQL.Add('    FORNECEDOR.NOME AS DES_FORNECEDOR,');
    SQL.Add('');
    SQL.Add('    CASE');
    SQL.Add('        WHEN COALESCE(FORNECEDOR.FANTASIA, '''') = '''' THEN FORNECEDOR.NOME');
    SQL.Add('        ELSE FORNECEDOR.FANTASIA');
    SQL.Add('    END AS DES_FANTASIA,');
    SQL.Add('');
    SQL.Add('    FORNECEDOR.CGC AS NUM_CGC,');
    SQL.Add('');
    SQL.Add('    CASE');
    SQL.Add('        WHEN COALESCE(FORNECEDOR.IE, '''') = '''' THEN ''ISENTO''');
    SQL.Add('        ELSE FORNECEDOR.IE');
    SQL.Add('    END AS NUM_INSC_EST,');
    SQL.Add('');
    SQL.Add('    CASE');
    SQL.Add('        WHEN COALESCE(FORNECEDOR.ENDERECO, '''') = '''' THEN ''A CONFIRMAR''');
    SQL.Add('        ELSE FORNECEDOR.ENDERECO');
    SQL.Add('    END AS DES_ENDERECO,');
    SQL.Add('');
    SQL.Add('    CASE');
    SQL.Add('        WHEN COALESCE(FORNECEDOR.BAIRRO, '''') = '''' THEN ''A CONFIRMAR''');
    SQL.Add('        ELSE FORNECEDOR.BAIRRO');
    SQL.Add('    END AS DES_BAIRRO,');
    SQL.Add('');
    SQL.Add('    FORNECEDOR.CIDADE AS DES_CIDADE,');
    SQL.Add('    FORNECEDOR.ESTADO AS DES_SIGLA,');
    SQL.Add('');
    SQL.Add('    CASE');
    SQL.Add('        WHEN COALESCE(FORNECEDOR.CEP, '''') = '''' THEN ''11111111111''');
    SQL.Add('        ELSE FORNECEDOR.CEP');
    SQL.Add('    END AS NUM_CEP,');
    SQL.Add('');
    SQL.Add('    FORNECEDOR.FONE AS NUM_FONE,');
    SQL.Add('    FORNECEDOR.FAX AS NUM_FAX,');
    SQL.Add('    FORNECEDOR.CONTATO AS DES_CONTATO,');
    SQL.Add('    0 AS QTD_DIA_CARENCIA,');
    SQL.Add('    0 AS NUM_FREQ_VISITA,');
    SQL.Add('    0 AS VAL_DESCONTO,');
    SQL.Add('    0 AS NUM_PRAZO,');
    SQL.Add('    ''N'' AS ACEITA_DEVOL_MER,');
    SQL.Add('    ''N'' AS CAL_IPI_VAL_BRUTO,');
    SQL.Add('    ''N'' AS CAL_ICMS_ENC_FIN,');
    SQL.Add('    ''N'' AS CAL_ICMS_VAL_IPI,');
    SQL.Add('    ''N'' AS MICRO_EMPRESA,');
    SQL.Add('    0 AS COD_FORNECEDOR_ANT,');
    SQL.Add('');
    SQL.Add('    CASE');
    SQL.Add('        WHEN COALESCE(FORNECEDOR.NUMERO, '''') = '''' THEN ''1234''');
    SQL.Add('        ELSE FORNECEDOR.NUMERO');
    SQL.Add('    END AS NUM_ENDERECO,');
    SQL.Add('');
    SQL.Add('    -- COALESCE(FORNECEDOR.FONE2, '''') || '' / '' || COALESCE(FORNECEDOR.FAX2, '''') || '' REPRESENTANTE: '' || COALESCE(FORNECEDOR.REPR, '''') || '' DETALHE: '' || COALESCE(CAST(FORNECEDOR.DETALHE AS VARCHAR(200)), '''') AS DES_OBSERVACAO,');
    SQL.Add('    COALESCE(FORNECEDOR.FONE2, '''') || '' / '' || COALESCE(FORNECEDOR.FAX2, '''') || '' REPRESENTANTE: '' || COALESCE(FORNECEDOR.REPR, '''') AS DES_OBSERVACAO,');
    SQL.Add('');
    SQL.Add('    FORNECEDOR.EMAIL AS DES_EMAIL,');
    SQL.Add('    FORNECEDOR.SITE AS DES_WEB_SITE,');
    SQL.Add('    ''N'' AS FABRICANTE,');
    SQL.Add('    ''N'' AS FLG_PRODUTOR_RURAL,');
    SQL.Add('    0 AS TIPO_FRETE,');
    SQL.Add('    ''N'' AS FLG_SIMPLES,');
    SQL.Add('    ''N'' AS FLG_SUBSTITUTO_TRIB,');
    SQL.Add('    0 AS COD_CONTACCFORN,');
    SQL.Add('    CASE FORNECEDOR.STATUS');
    SQL.Add('        WHEN ''A'' THEN ''N''');
    SQL.Add('        ELSE ''S''');
    SQL.Add('    END AS INATIVO,');
    SQL.Add('    21 AS COD_CLASSIF,');
    SQL.Add('    FORNECEDOR.DTCADASTRO AS DTA_CADASTRO,');
    SQL.Add('    0 AS VAL_CREDITO,');
    SQL.Add('    0 AS VAL_DEBITO,');
    SQL.Add('    0 AS PED_MIN_VAL,');
    SQL.Add('    '''' AS DES_EMAIL_VEND,');
    SQL.Add('    '''' AS SENHA_COTACAO,');
    SQL.Add('    -1 AS TIPO_PRODUTOR,');
    SQL.Add('    FORNECEDOR.FONE1 AS NUM_CELULAR');
    SQL.Add('FROM FORNEC AS FORNECEDOR');
    SQL.Add('');
    SQL.Add('UNION ALL');
    SQL.Add('');
    SQL.Add('SELECT');
    SQL.Add('    1000 + FORNECEDOR.CODIGO AS COD_FORNECEDOR,');
    SQL.Add('    FORNECEDOR.NOME AS DES_FORNECEDOR,');
    SQL.Add('');
    SQL.Add('    CASE');
    SQL.Add('        WHEN COALESCE(FORNECEDOR.FANTASIA, '''') = '''' THEN FORNECEDOR.NOME');
    SQL.Add('        ELSE FORNECEDOR.FANTASIA');
    SQL.Add('    END AS DES_FANTASIA,');
    SQL.Add('');
    SQL.Add('    FORNECEDOR.CGC AS NUM_CGC,');
    SQL.Add('');
    SQL.Add('    CASE');
    SQL.Add('        WHEN COALESCE(FORNECEDOR.IE, '''') = '''' THEN ''ISENTO''');
    SQL.Add('        ELSE FORNECEDOR.IE');
    SQL.Add('    END AS NUM_INSC_EST,');
    SQL.Add('');
    SQL.Add('    CASE');
    SQL.Add('        WHEN COALESCE(FORNECEDOR.ENDERECO, '''') = '''' THEN ''A CONFIRMAR''');
    SQL.Add('        ELSE FORNECEDOR.ENDERECO');
    SQL.Add('    END AS DES_ENDERECO,');
    SQL.Add('');
    SQL.Add('    CASE');
    SQL.Add('        WHEN COALESCE(FORNECEDOR.BAIRRO, '''') = '''' THEN ''A CONFIRMAR''');
    SQL.Add('        ELSE FORNECEDOR.BAIRRO');
    SQL.Add('    END AS DES_BAIRRO,');
    SQL.Add('');
    SQL.Add('    FORNECEDOR.CIDADE AS DES_CIDADE,');
    SQL.Add('    FORNECEDOR.ESTADO AS DES_SIGLA,');
    SQL.Add('');
    SQL.Add('    CASE');
    SQL.Add('        WHEN COALESCE(FORNECEDOR.CEP, '''') = '''' THEN ''11111111111''');
    SQL.Add('        ELSE FORNECEDOR.CEP');
    SQL.Add('    END AS NUM_CEP,');
    SQL.Add('');
    SQL.Add('    FORNECEDOR.FONE AS NUM_FONE,');
    SQL.Add('    FORNECEDOR.FAX AS NUM_FAX,');
    SQL.Add('    FORNECEDOR.CONTATO AS DES_CONTATO,');
    SQL.Add('    0 AS QTD_DIA_CARENCIA,');
    SQL.Add('    0 AS NUM_FREQ_VISITA,');
    SQL.Add('    0 AS VAL_DESCONTO,');
    SQL.Add('    0 AS NUM_PRAZO,');
    SQL.Add('    ''N'' AS ACEITA_DEVOL_MER,');
    SQL.Add('    ''N'' AS CAL_IPI_VAL_BRUTO,');
    SQL.Add('    ''N'' AS CAL_ICMS_ENC_FIN,');
    SQL.Add('    ''N'' AS CAL_ICMS_VAL_IPI,');
    SQL.Add('    ''N'' AS MICRO_EMPRESA,');
    SQL.Add('    0 AS COD_FORNECEDOR_ANT,');
    SQL.Add('');
    SQL.Add('    CASE');
    SQL.Add('        WHEN COALESCE(FORNECEDOR.NUMERO, '''') = '''' THEN ''1234''');
    SQL.Add('        ELSE FORNECEDOR.NUMERO');
    SQL.Add('    END AS NUM_ENDERECO,');
    SQL.Add('');
    SQL.Add('    -- COALESCE(FORNECEDOR.FONE2, '''') || '' / '' || COALESCE(FORNECEDOR.FAX2, '''') || '' REPRESENTANTE: '' || COALESCE(FORNECEDOR.REPR, '''') || '' DETALHE: '' || COALESCE(CAST(FORNECEDOR.DETALHE AS VARCHAR(200)), '''') AS DES_OBSERVACAO,');
    SQL.Add('    COALESCE(FORNECEDOR.FONE2, '''') || '' / '' || COALESCE(FORNECEDOR.FAX2, '''') || '' REPRESENTANTE: '' || COALESCE(FORNECEDOR.REPR, '''') AS DES_OBSERVACAO,');
    SQL.Add('');
    SQL.Add('    FORNECEDOR.EMAIL AS DES_EMAIL,');
    SQL.Add('    FORNECEDOR.SITE AS DES_WEB_SITE,');
    SQL.Add('    ''N'' AS FABRICANTE,');
    SQL.Add('    ''N'' AS FLG_PRODUTOR_RURAL,');
    SQL.Add('    0 AS TIPO_FRETE,');
    SQL.Add('    ''N'' AS FLG_SIMPLES,');
    SQL.Add('    ''N'' AS FLG_SUBSTITUTO_TRIB,');
    SQL.Add('    0 AS COD_CONTACCFORN,');
    SQL.Add('    CASE FORNECEDOR.STATUS');
    SQL.Add('        WHEN ''A'' THEN ''N''');
    SQL.Add('        ELSE ''S''');
    SQL.Add('    END AS INATIVO,');
    SQL.Add('    21 AS COD_CLASSIF,');
    SQL.Add('    FORNECEDOR.DTCADASTRO AS DTA_CADASTRO,');
    SQL.Add('    0 AS VAL_CREDITO,');
    SQL.Add('    0 AS VAL_DEBITO,');
    SQL.Add('    0 AS PED_MIN_VAL,');
    SQL.Add('    '''' AS DES_EMAIL_VEND,');
    SQL.Add('    '''' AS SENHA_COTACAO,');
    SQL.Add('    -1 AS TIPO_PRODUTOR,');
    SQL.Add('    FORNECEDOR.FONE1 AS NUM_CELULAR');
    SQL.Add('FROM TAB_FORNECEDOR_AUX AS FORNECEDOR');
    SQL.Add('LEFT JOIN  FORNEC');
    SQL.Add('ON COALESCE(FORNEC.CGC, '''') <> ''''');
    SQL.Add('AND REPLACE(REPLACE(REPLACE(COALESCE(FORNECEDOR.CGC, ''''), ''.'', ''''), ''-'', ''''), ''/'', '''') = REPLACE(REPLACE(REPLACE(COALESCE(FORNEC.CGC, ''''), ''.'', ''''), ''-'', ''''), ''/'', '''')');
    SQL.Add('WHERE FORNEC.CODIGO IS NULL');

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
//      Layout.FieldByName('NUM_INSC_EST').AsString := StrRetNums(Layout.FieldByName('NUM_INSC_EST').AsString);
      Layout.FieldByName('NUM_CEP').AsString := StrRetNums(Layout.FieldByName('NUM_CEP').AsString);

      if QryPrincipal.FieldByName('NUM_INSC_EST').AsString = '0' then
         Layout.FieldByName('NUM_INSC_EST').AsString := 'ISENTO';

      if QryPrincipal.FieldByName('NUM_INSC_EST').AsString <> 'ISENTO' then
         Layout.FieldByName('NUM_INSC_EST').AsString := StrRetNums(Layout.FieldByName('NUM_INSC_EST').AsString);

      if Length(Layout.FieldByName('NUM_CGC').AsString) > 11 then
      begin
        if not ValidaCGC(Layout.FieldByName('NUM_CGC').AsString) then
          Layout.FieldByName('NUM_CGC').AsString := '';
      end
      else
        if not ValidaCPF(Layout.FieldByName('NUM_CGC').AsString) then
          Layout.FieldByName('NUM_CGC').AsString := '';

      Layout.FieldByName('NUM_FONE').AsString := StrRetNums( FieldByName('NUM_FONE').AsString );

      observacao := StrReplace(StrLBReplace(FieldByName('DES_OBSERVACAO').AsString), '\n', '');
      email := StrReplace(StrLBReplace(FieldByName('DES_EMAIL').AsString), '\n', '');


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

procedure TFrmByteCortez.GerarGrupo;
var
   TotalCount : Integer;
begin
  inherited;

  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT DISTINCT');
    SQL.Add('    MARCAS.CODIGO AS COD_SECAO,');
    SQL.Add('    999 AS COD_GRUPO,');
    SQL.Add('    ''A DEFINIR'' AS DES_GRUPO,');
    SQL.Add('    0 AS VAL_META');
    SQL.Add('FROM PRODUTOS');
    SQL.Add('LEFT JOIN MARCAS');
    SQL.Add('ON PRODUTOS.COD_MARCA = MARCAS.CODIGO');
    SQL.Add('');
    SQL.Add('UNION ALL');
    SQL.Add('');
    SQL.Add('SELECT DISTINCT');
    SQL.Add('    200 + MARCAS.CODIGO AS COD_SECAO,');
    SQL.Add('    999 AS COD_GRUPO,');
    SQL.Add('    ''A DEFINIR'' AS DES_GRUPO,');
    SQL.Add('    0 AS VAL_META');
    SQL.Add('FROM TAB_PRODUTO_AUX AS PRODUTOS');
    SQL.Add('LEFT JOIN TAB_MARCAS_AUX AS MARCAS');
    SQL.Add('ON PRODUTOS.COD_MARCA = MARCAS.CODIGO');

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

procedure TFrmByteCortez.GerarInfoNutricionais;
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
    SQL.Add('        WHEN 21 THEN NUTRICIONAL.MEDIDAI || '' '' || ''FILÉ(S)''');
    SQL.Add('        WHEN 20 THEN NUTRICIONAL.MEDIDAI || '' '' || ''BIFE(S)''');
    SQL.Add('        WHEN 2 THEN NUTRICIONAL.MEDIDAI || '' '' || ''COLHER(ES) DE CHÁ''');
    SQL.Add('        WHEN 5 THEN NUTRICIONAL.MEDIDAI || '' '' || ''UNIDADE''');
    SQL.Add('        WHEN 24 THEN NUTRICIONAL.MEDIDAI || '' '' || ''PRATO(S) FUNDO(S)''');
    SQL.Add('        WHEN 4 THEN NUTRICIONAL.MEDIDAI || '' '' || ''DE XÍCARA(S)''');
    SQL.Add('        WHEN 8 THEN NUTRICIONAL.MEDIDAI || '' '' || ''FATIA(S) FINA(S)''');
    SQL.Add('        WHEN 7 THEN NUTRICIONAL.MEDIDAI || '' '' || ''FATIA(S)''');
    SQL.Add('        WHEN 3 THEN NUTRICIONAL.MEDIDAI || '' '' || ''XÍCARA(S)''');
    SQL.Add('        WHEN 15 THEN NUTRICIONAL.MEDIDAI || '' '' || ''COPO(S)''');
    SQL.Add('        WHEN 0 THEN NUTRICIONAL.MEDIDAI || '' '' || ''COLHER(ES) DE SOPA''');
    SQL.Add('        WHEN 16 THEN NUTRICIONAL.MEDIDAI || '' '' || ''PORÇÃO(ÕES)''');
    SQL.Add('        WHEN 9 THEN NUTRICIONAL.MEDIDAI || '' '' || ''PEDAÇO(S)''');
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

procedure TFrmByteCortez.GerarNCM;
var
 count, TotalCount : Integer;
begin
  inherited;
  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT * FROM (');
    SQL.Add('    SELECT DISTINCT');
    SQL.Add('        0 AS COD_NCM,');
    SQL.Add('        COALESCE(NCM.DESCRICAO, ''A DEFINIR'') AS DES_NCM,');
    SQL.Add('');
    SQL.Add('        CASE ');
    SQL.Add('            WHEN CAST(COALESCE(PRODUTO.CLASFISCAL, 0) AS INTEGER) = 0 THEN ''99999999''');
    SQL.Add('            ELSE TRIM(PRODUTO.CLASFISCAL)');
    SQL.Add('        END AS NUM_NCM,');
    SQL.Add('');
    SQL.Add('        PRODUTO.PIS_CST, ');
    SQL.Add('');
    SQL.Add('        CASE COALESCE(PRODUTO.PIS_CST, 0)');
    SQL.Add('            WHEN 0 THEN ''N''');
    SQL.Add('            WHEN 1 THEN ''N''');
    SQL.Add('            WHEN 4 THEN ''S''');
    SQL.Add('            WHEN 6 THEN ''S''');
    SQL.Add('            WHEN 7 THEN ''N''');
    SQL.Add('            WHEN 8 THEN ''S''');
    SQL.Add('        END AS FLG_NAO_PIS_COFINS,');
    SQL.Add('');
    SQL.Add('        CASE COALESCE(PRODUTO.PIS_CST, 0)');
    SQL.Add('            WHEN 0 THEN -1');
    SQL.Add('            WHEN 1 THEN -1');
    SQL.Add('            WHEN 4 THEN 1');
    SQL.Add('            WHEN 6 THEN 0');
    SQL.Add('            WHEN 7 THEN -1');
    SQL.Add('            WHEN 8 THEN 3');
    SQL.Add('        END AS TIPO_NAO_PIS_COFINS,');
    SQL.Add('');
    SQL.Add('        999 AS COD_TAB_SPED,');
    SQL.Add('');
    SQL.Add('        CASE');
    SQL.Add('            WHEN COALESCE(PRODUTO.CEST, '''') = '''' THEN ''9999999''');
    SQL.Add('            ELSE PRODUTO.CEST');
    SQL.Add('        END AS NUM_CEST,');
    SQL.Add('');
    SQL.Add('        ''SP'' AS DES_SIGLA,');
    SQL.Add('');
    SQL.Add('        CASE PRODUTO.ICMS');
    SQL.Add('            WHEN ''4,5'' THEN 38');
    SQL.Add('            WHEN ''F1'' THEN 13');
    SQL.Add('            WHEN ''FF'' THEN 13');
    SQL.Add('            WHEN ''II'' THEN 1');
    SQL.Add('            WHEN ''T01'' THEN 4');
    SQL.Add('            WHEN ''T02'' THEN 2');
    SQL.Add('            WHEN ''T03'' THEN 3');
    SQL.Add('            WHEN ''T04'' THEN 5');
    SQL.Add('        END AS COD_TRIB_ENTRADA,');
    SQL.Add('');
    SQL.Add('        CASE PRODUTO.ICMS');
    SQL.Add('            WHEN ''4,5'' THEN 38');
    SQL.Add('            WHEN ''F1'' THEN 13');
    SQL.Add('            WHEN ''FF'' THEN 13');
    SQL.Add('            WHEN ''II'' THEN 1');
    SQL.Add('            WHEN ''T01'' THEN 4');
    SQL.Add('            WHEN ''T02'' THEN 2');
    SQL.Add('            WHEN ''T03'' THEN 3');
    SQL.Add('            WHEN ''T04'' THEN 5');
    SQL.Add('        END AS COD_TRIB_SAIDA,');
    SQL.Add('');
    SQL.Add('        0 AS PER_IVA,');
    SQL.Add('        0 AS PER_FCP_ST');
    SQL.Add('    FROM PRODUTOS AS PRODUTO');
    SQL.Add('    LEFT JOIN NCM');
    SQL.Add('    ON PRODUTO.CLASFISCAL = NCM.NCM');
    SQL.Add('');
    SQL.Add('    UNION ALL');
    SQL.Add('');
    SQL.Add('    SELECT DISTINCT');
    SQL.Add('        0 AS COD_NCM,');
    SQL.Add('        ''A DEFINIR'' AS DES_NCM,');
    SQL.Add('');
    SQL.Add('        CASE ');
    SQL.Add('            WHEN CAST(COALESCE(PRODUTO.CLASFISCAL, 0) AS INTEGER) = 0 THEN ''99999999''');
    SQL.Add('            ELSE TRIM(PRODUTO.CLASFISCAL)');
    SQL.Add('        END AS NUM_NCM,');
    SQL.Add('');
    SQL.Add('        PRODUTO.PIS_CST, ');
    SQL.Add('');
    SQL.Add('        CASE COALESCE(PRODUTO.PIS_CST, 0)');
    SQL.Add('            WHEN 0 THEN ''N''');
    SQL.Add('            WHEN 1 THEN ''N''');
    SQL.Add('            WHEN 4 THEN ''S''');
    SQL.Add('            WHEN 5 THEN ''S''');
    SQL.Add('            WHEN 6 THEN ''S''');
    SQL.Add('            WHEN 7 THEN ''N''');
    SQL.Add('            WHEN 8 THEN ''S''');
    SQL.Add('        END AS FLG_NAO_PIS_COFINS,');
    SQL.Add('');
    SQL.Add('        CASE COALESCE(PRODUTO.PIS_CST, 0)');
    SQL.Add('            WHEN 0 THEN -1');
    SQL.Add('            WHEN 1 THEN -1');
    SQL.Add('            WHEN 4 THEN 1');
    SQL.Add('            WHEN 5 THEN 2');
    SQL.Add('            WHEN 6 THEN 0');
    SQL.Add('            WHEN 7 THEN -1');
    SQL.Add('            WHEN 8 THEN 3');
    SQL.Add('        END AS TIPO_NAO_PIS_COFINS,');
    SQL.Add('');
    SQL.Add('        999 AS COD_TAB_SPED,');
    SQL.Add('');
    SQL.Add('        CASE');
    SQL.Add('            WHEN COALESCE(PRODUTO.CEST, '''') = '''' THEN ''9999999''');
    SQL.Add('            WHEN PRODUTO.CEST = ''FF'' THEN ''9999999''');
    SQL.Add('            ELSE PRODUTO.CEST');
    SQL.Add('        END AS NUM_CEST,');
    SQL.Add('');
    SQL.Add('        ''SP'' AS DES_SIGLA,');
    SQL.Add('');
    SQL.Add('        CASE PRODUTO.ICMS');
    SQL.Add('            WHEN ''4,5'' THEN 38');
    SQL.Add('            WHEN ''F1'' THEN 13');
    SQL.Add('            WHEN ''FF'' THEN 13');
    SQL.Add('            WHEN ''II'' THEN 1');
    SQL.Add('            WHEN ''T03'' THEN 4');
    SQL.Add('            WHEN ''T04'' THEN 2');
    SQL.Add('            WHEN ''T05'' THEN 3');
    SQL.Add('            WHEN ''T06'' THEN 5');
    SQL.Add('            WHEN ''T07'' THEN 25');
    SQL.Add('        END AS COD_TRIB_ENTRADA,');
    SQL.Add('');
    SQL.Add('        CASE PRODUTO.ICMS');
    SQL.Add('            WHEN ''4,5'' THEN 38');
    SQL.Add('            WHEN ''F1'' THEN 13');
    SQL.Add('            WHEN ''FF'' THEN 13');
    SQL.Add('            WHEN ''II'' THEN 1');
    SQL.Add('            WHEN ''T03'' THEN 4');
    SQL.Add('            WHEN ''T04'' THEN 2');
    SQL.Add('            WHEN ''T05'' THEN 3');
    SQL.Add('            WHEN ''T06'' THEN 5');
    SQL.Add('            WHEN ''T07'' THEN 25');
    SQL.Add('        END AS COD_TRIB_SAIDA,');
    SQL.Add('');
    SQL.Add('        0 AS PER_IVA,');
    SQL.Add('        0 AS PER_FCP_ST');
    SQL.Add('    FROM TAB_PRODUTO_AUX AS PRODUTO');
    SQL.Add('    LEFT JOIN PRODUTOS');
    SQL.Add('    ON PRODUTO.COD_BARRA = PRODUTOS.COD_BARRA');
    SQL.Add('    WHERE PRODUTOS.COD_PROD IS NULL');
    SQL.Add(')');
    SQL.Add('ORDER BY NUM_NCM');

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

      Layout.FieldByName('COD_NCM').AsInteger := count;

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;
end;

procedure TFrmByteCortez.GerarNCMUF;
var
 count, TotalCount : Integer;
begin
  inherited;
  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT * FROM (');
    SQL.Add('    SELECT DISTINCT');
    SQL.Add('        0 AS COD_NCM,');
    SQL.Add('        COALESCE(NCM.DESCRICAO, ''A DEFINIR'') AS DES_NCM,');
    SQL.Add('');
    SQL.Add('        CASE ');
    SQL.Add('            WHEN CAST(COALESCE(PRODUTO.CLASFISCAL, 0) AS INTEGER) = 0 THEN ''99999999''');
    SQL.Add('            ELSE TRIM(PRODUTO.CLASFISCAL)');
    SQL.Add('        END AS NUM_NCM,');
    SQL.Add('');
    SQL.Add('        PRODUTO.PIS_CST, ');
    SQL.Add('');
    SQL.Add('        CASE COALESCE(PRODUTO.PIS_CST, 0)');
    SQL.Add('            WHEN 0 THEN ''N''');
    SQL.Add('            WHEN 1 THEN ''N''');
    SQL.Add('            WHEN 4 THEN ''S''');
    SQL.Add('            WHEN 6 THEN ''S''');
    SQL.Add('            WHEN 7 THEN ''N''');
    SQL.Add('            WHEN 8 THEN ''S''');
    SQL.Add('        END AS FLG_NAO_PIS_COFINS,');
    SQL.Add('');
    SQL.Add('        CASE COALESCE(PRODUTO.PIS_CST, 0)');
    SQL.Add('            WHEN 0 THEN -1');
    SQL.Add('            WHEN 1 THEN -1');
    SQL.Add('            WHEN 4 THEN 1');
    SQL.Add('            WHEN 6 THEN 0');
    SQL.Add('            WHEN 7 THEN -1');
    SQL.Add('            WHEN 8 THEN 3');
    SQL.Add('        END AS TIPO_NAO_PIS_COFINS,');
    SQL.Add('');
    SQL.Add('        999 AS COD_TAB_SPED,');
    SQL.Add('');
    SQL.Add('        CASE');
    SQL.Add('            WHEN COALESCE(PRODUTO.CEST, '''') = '''' THEN ''9999999''');
    SQL.Add('            ELSE PRODUTO.CEST');
    SQL.Add('        END AS NUM_CEST,');
    SQL.Add('');
    SQL.Add('        ''SP'' AS DES_SIGLA,');
    SQL.Add('');
    SQL.Add('        CASE PRODUTO.ICMS');
    SQL.Add('            WHEN ''4,5'' THEN 38');
    SQL.Add('            WHEN ''F1'' THEN 13');
    SQL.Add('            WHEN ''FF'' THEN 13');
    SQL.Add('            WHEN ''II'' THEN 1');
    SQL.Add('            WHEN ''T01'' THEN 4');
    SQL.Add('            WHEN ''T02'' THEN 2');
    SQL.Add('            WHEN ''T03'' THEN 3');
    SQL.Add('            WHEN ''T04'' THEN 5');
    SQL.Add('        END AS COD_TRIB_ENTRADA,');
    SQL.Add('');
    SQL.Add('        CASE PRODUTO.ICMS');
    SQL.Add('            WHEN ''4,5'' THEN 38');
    SQL.Add('            WHEN ''F1'' THEN 13');
    SQL.Add('            WHEN ''FF'' THEN 13');
    SQL.Add('            WHEN ''II'' THEN 1');
    SQL.Add('            WHEN ''T01'' THEN 4');
    SQL.Add('            WHEN ''T02'' THEN 2');
    SQL.Add('            WHEN ''T03'' THEN 3');
    SQL.Add('            WHEN ''T04'' THEN 5');
    SQL.Add('        END AS COD_TRIB_SAIDA,');
    SQL.Add('');
    SQL.Add('        0 AS PER_IVA,');
    SQL.Add('        0 AS PER_FCP_ST');
    SQL.Add('    FROM PRODUTOS AS PRODUTO');
    SQL.Add('    LEFT JOIN NCM');
    SQL.Add('    ON PRODUTO.CLASFISCAL = NCM.NCM');
    SQL.Add('');
    SQL.Add('    UNION ALL');
    SQL.Add('');
    SQL.Add('    SELECT DISTINCT');
    SQL.Add('        0 AS COD_NCM,');
    SQL.Add('        ''A DEFINIR'' AS DES_NCM,');
    SQL.Add('');
    SQL.Add('        CASE ');
    SQL.Add('            WHEN CAST(COALESCE(PRODUTO.CLASFISCAL, 0) AS INTEGER) = 0 THEN ''99999999''');
    SQL.Add('            ELSE TRIM(PRODUTO.CLASFISCAL)');
    SQL.Add('        END AS NUM_NCM,');
    SQL.Add('');
    SQL.Add('        PRODUTO.PIS_CST, ');
    SQL.Add('');
    SQL.Add('        CASE COALESCE(PRODUTO.PIS_CST, 0)');
    SQL.Add('            WHEN 0 THEN ''N''');
    SQL.Add('            WHEN 1 THEN ''N''');
    SQL.Add('            WHEN 4 THEN ''S''');
    SQL.Add('            WHEN 5 THEN ''S''');
    SQL.Add('            WHEN 6 THEN ''S''');
    SQL.Add('            WHEN 7 THEN ''N''');
    SQL.Add('            WHEN 8 THEN ''S''');
    SQL.Add('        END AS FLG_NAO_PIS_COFINS,');
    SQL.Add('');
    SQL.Add('        CASE COALESCE(PRODUTO.PIS_CST, 0)');
    SQL.Add('            WHEN 0 THEN -1');
    SQL.Add('            WHEN 1 THEN -1');
    SQL.Add('            WHEN 4 THEN 1');
    SQL.Add('            WHEN 5 THEN 2');
    SQL.Add('            WHEN 6 THEN 0');
    SQL.Add('            WHEN 7 THEN -1');
    SQL.Add('            WHEN 8 THEN 3');
    SQL.Add('        END AS TIPO_NAO_PIS_COFINS,');
    SQL.Add('');
    SQL.Add('        999 AS COD_TAB_SPED,');
    SQL.Add('');
    SQL.Add('        CASE');
    SQL.Add('            WHEN COALESCE(PRODUTO.CEST, '''') = '''' THEN ''9999999''');
    SQL.Add('            WHEN PRODUTO.CEST = ''FF'' THEN ''9999999''');
    SQL.Add('            ELSE PRODUTO.CEST');
    SQL.Add('        END AS NUM_CEST,');
    SQL.Add('');
    SQL.Add('        ''SP'' AS DES_SIGLA,');
    SQL.Add('');
    SQL.Add('        CASE PRODUTO.ICMS');
    SQL.Add('            WHEN ''4,5'' THEN 38');
    SQL.Add('            WHEN ''F1'' THEN 13');
    SQL.Add('            WHEN ''FF'' THEN 13');
    SQL.Add('            WHEN ''II'' THEN 1');
    SQL.Add('            WHEN ''T03'' THEN 4');
    SQL.Add('            WHEN ''T04'' THEN 2');
    SQL.Add('            WHEN ''T05'' THEN 3');
    SQL.Add('            WHEN ''T06'' THEN 5');
    SQL.Add('            WHEN ''T07'' THEN 25');
    SQL.Add('        END AS COD_TRIB_ENTRADA,');
    SQL.Add('');
    SQL.Add('        CASE PRODUTO.ICMS');
    SQL.Add('            WHEN ''4,5'' THEN 38');
    SQL.Add('            WHEN ''F1'' THEN 13');
    SQL.Add('            WHEN ''FF'' THEN 13');
    SQL.Add('            WHEN ''II'' THEN 1');
    SQL.Add('            WHEN ''T03'' THEN 4');
    SQL.Add('            WHEN ''T04'' THEN 2');
    SQL.Add('            WHEN ''T05'' THEN 3');
    SQL.Add('            WHEN ''T06'' THEN 5');
    SQL.Add('            WHEN ''T07'' THEN 25');
    SQL.Add('        END AS COD_TRIB_SAIDA,');
    SQL.Add('');
    SQL.Add('        0 AS PER_IVA,');
    SQL.Add('        0 AS PER_FCP_ST');
    SQL.Add('    FROM TAB_PRODUTO_AUX AS PRODUTO');
    SQL.Add('    LEFT JOIN PRODUTOS');
    SQL.Add('    ON PRODUTO.COD_BARRA = PRODUTOS.COD_BARRA');
    SQL.Add('    WHERE PRODUTOS.COD_PROD IS NULL');
    SQL.Add(')');
    SQL.Add('ORDER BY NUM_NCM');

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

      Layout.FieldByName('COD_NCM').AsInteger := count;

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;
end;

procedure TFrmByteCortez.GerarNFClientes;
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

procedure TFrmByteCortez.GerarNFFornec;
var
   TotalCount : integer;
begin
  inherited;
  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
    SQL.Add('    CAPA.COD_FORNEC AS COD_FORNECEDOR,');
    SQL.Add('    CAPA.DOCUMENTO AS NUM_NF_FORN,');
    SQL.Add('    COALESCE(CAPA.SERIE, ''0'') AS NUM_SERIE_NF,');
    SQL.Add('    '''' AS NUM_SUBSERIE_NF,');
    SQL.Add('    CAPA.CFOP AS CFOP,');
    SQL.Add('    0 AS TIPO_NF,');
    SQL.Add('    ');
    SQL.Add('    CASE ');
    SQL.Add('        WHEN CAPA.NFECHAVE IS NOT NULL THEN''NFE''');
    SQL.Add('        ELSE ''NF''');
    SQL.Add('    END AS DES_ESPECIE,');
    SQL.Add('');
    SQL.Add('    CAPA.VALORTOTALNOTA AS VAL_TOTAL_NF,');
    SQL.Add('    CAPA.DATA AS DTA_EMISSAO,');
    SQL.Add('    CAPA.DATA AS DTA_ENTRADA,');
    SQL.Add('    CAPA.VALORIPI AS VAL_TOTAL_IPI,');
    SQL.Add('    CAPA.VALORTOTALPRODUTOS AS VAL_VENDA_VAREJO,');
    SQL.Add('    COALESCE(CAPA.VALORFRETE, 0) AS VAL_FRETE,');
    SQL.Add('    COALESCE(CAPA.VALORSEGURO, 0) AS VAL_ACRESCIMO,');
    SQL.Add('    COALESCE(CAPA.DESCONTO, 0) AS VAL_DESCONTO,');
    SQL.Add('    '''' AS NUM_CGC,');
    SQL.Add('    CAPA.BASECALCICMS AS VAL_TOTAL_BC,');
    SQL.Add('    CAPA.VALORICMS AS VAL_TOTAL_ICMS,');
    SQL.Add('    CAPA.BASECALCICMSST AS VAL_BC_SUBST,');
    SQL.Add('    CAPA.VALORICMSST AS VAL_ICMS_SUBST,');
    SQL.Add('    0 AS VAL_FUNRURAL,');
    SQL.Add('    1 AS COD_PERFIL,');
    SQL.Add('    COALESCE(CAPA.VALOROUTRASDESPESAS, 0) AS VAL_DESP_ACESS,');
    SQL.Add('    ''N'' AS FLG_CANCELADO,');
    SQL.Add('    '''' AS DES_OBSERVACAO,');
    SQL.Add('    CAPA.NFECHAVE AS NUM_CHAVE_ACESSO');
    SQL.Add('FROM ESTOQUEMESTRE AS CAPA');
    SQL.Add('WHERE CAPA.MODELO = ''55''');
    SQL.Add('AND CAPA.DATA >= :INI');
    SQL.Add('AND CAPA.DATA <= :FIM');

    ParamByName('INI').AsDate := DtpInicial.Date;
    ParamByName('FIM').AsDate := DtpFinal.Date;


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

procedure TFrmByteCortez.GerarNFitensClientes;
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

procedure TFrmByteCortez.GerarNFitensFornec;
var
   fornecedor, nota, serie : string;
   count, TotalCount : integer;

begin
  inherited;

  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
    SQL.Add('    ITENS.COD_FORNEC AS COD_FORNECEDOR,');
    SQL.Add('    CAPA.DOCUMENTO AS NUM_NF_FORN,');
    SQL.Add('    CAPA.SERIE AS NUM_SERIE_NF,');
    SQL.Add('    ITENS.COD_PROD AS COD_PRODUTO,');
    SQL.Add('');
    SQL.Add('    CASE');
    SQL.Add('        WHEN ITENS.CSTNFENTRADA = ''000'' AND ITENS.ICMSNFENTRADA = 0 THEN 13 ');
    SQL.Add('        WHEN ITENS.CSTNFENTRADA = ''000'' AND ITENS.ICMSNFENTRADA = 4 THEN 27 ');
    SQL.Add('        WHEN ITENS.CSTNFENTRADA = ''000'' AND ITENS.ICMSNFENTRADA = 4.5 THEN 38 ');
    SQL.Add('        WHEN ITENS.CSTNFENTRADA = ''000'' AND ITENS.ICMSNFENTRADA = 7 THEN 2 ');
    SQL.Add('        WHEN ITENS.CSTNFENTRADA = ''000'' AND ITENS.ICMSNFENTRADA = 12 THEN 3 ');
    SQL.Add('        WHEN ITENS.CSTNFENTRADA = ''000'' AND ITENS.ICMSNFENTRADA = 18 THEN 4 ');
    SQL.Add('        WHEN ITENS.CSTNFENTRADA = ''000'' AND ITENS.ICMSNFENTRADA = 25 THEN 5 ');
    SQL.Add('        WHEN ITENS.CSTNFENTRADA = ''020'' AND ITENS.ICMSNFENTRADA = 0 THEN 7 ');
    SQL.Add('        WHEN ITENS.CSTNFENTRADA = ''020'' AND ITENS.ICMSNFENTRADA = 7 THEN 7 ');
    SQL.Add('        WHEN ITENS.CSTNFENTRADA = ''020'' AND ITENS.ICMSNFENTRADA = 11 THEN 7 ');
    SQL.Add('        WHEN ITENS.CSTNFENTRADA = ''020'' AND ITENS.ICMSNFENTRADA = 12 THEN 6 ');
    SQL.Add('        WHEN ITENS.CSTNFENTRADA = ''020'' AND ITENS.ICMSNFENTRADA = 18 THEN 7 ');
    SQL.Add('        WHEN ITENS.CSTNFENTRADA = ''040'' AND ITENS.ICMSNFENTRADA = 0 THEN 1 ');
    SQL.Add('        WHEN ITENS.CSTNFENTRADA = ''041'' AND ITENS.ICMSNFENTRADA = 0 THEN 23 ');
    SQL.Add('        WHEN ITENS.CSTNFENTRADA = ''051'' AND ITENS.ICMSNFENTRADA = 0 THEN 20 ');
    SQL.Add('        WHEN ITENS.CSTNFENTRADA = ''060'' AND ITENS.ICMSNFENTRADA = 0 THEN 13 ');
    SQL.Add('        WHEN ITENS.CSTNFENTRADA = ''060'' AND ITENS.ICMSNFENTRADA = 18 THEN 13 ');
    SQL.Add('    END AS COD_TRIBUTACAO,');
    SQL.Add('');
    SQL.Add('    1 AS QTD_EMBALAGEM,');
    SQL.Add('');
    SQL.Add('    ITENS.QUANT_ENT AS QTD_ENTRADA,');
    SQL.Add('');
    SQL.Add('    ''UN'' AS DES_UNIDADE,');
    SQL.Add('    ITENS.CUSTO_ENT AS VAL_TABELA,');
    SQL.Add('    COALESCE(ITENS.DESCONTO, 0) AS VAL_DESCONTO_ITEM,');
    SQL.Add('    0 AS VAL_ACRESCIMO_ITEM,');
    SQL.Add('    ITENS.IPI AS VAL_IPI_ITEM,');
    SQL.Add('    ITENS.SUBSTITUICAO AS VAL_SUBST_ITEM,');
    SQL.Add('    ITENS.FRETE AS VAL_FRETE_ITEM,');
    SQL.Add('    ITENS.ICMS_VALOR AS VAL_CREDITO_ICMS,');
    SQL.Add('    ITENS.CUSTO_ENT AS VAL_VENDA_VAREJO,');
    SQL.Add('    ((ROUND(ITENS.CUSTO_ENT, 2) - ROUND(COALESCE(ITENS.IPI, 0), 2) - ROUND(COALESCE(ITENS.DESCONTO, 0), 2)) * ITENS.QUANT_ENT) AS VAL_TABELA_LIQ,');
    SQL.Add('    '''' AS NUM_CGC,');
    SQL.Add('    COALESCE(ITENS.ICMS_BC, 0) AS VAL_TOT_BC_ICMS,');
    SQL.Add('    ITENS.EMB_OUTROS AS VAL_TOT_OUTROS_ICMS,');
    SQL.Add('    ITENS.CFOPNFENTRADA AS CFOP,');
    SQL.Add('    0 AS VAL_TOT_ISENTO,');
    SQL.Add('    ITENS.ICMSST_BC AS VAL_TOT_BC_ST,');
    SQL.Add('    COALESCE(ROUND(ITENS.SUBSTITUICAO, 2), 0) * ITENS.QUANT_ENT AS VAL_TOT_ST,');
    SQL.Add('    1 AS NUM_ITEM,');
    SQL.Add('    0 AS TIPO_IPI,');
    SQL.Add('');
    SQL.Add('    CASE ');
    SQL.Add('        WHEN CAST(COALESCE(PRODUTO.CLASFISCAL, 0) AS INTEGER) = 0 THEN ''99999999''');
    SQL.Add('        ELSE TRIM(PRODUTO.CLASFISCAL)');
    SQL.Add('    END AS NUM_NCM,');
    SQL.Add('');
    SQL.Add('    '''' AS DES_REFERENCIA');
    SQL.Add('FROM ESTOQUE AS ITENS');
    SQL.Add('LEFT JOIN ESTOQUEMESTRE AS CAPA');
    SQL.Add('ON ITENS.SEQ_MESTRE = CAPA.SEQ');
    SQL.Add('AND CAPA.MODELO = ''55''');
    SQL.Add('LEFT JOIN PRODUTOS AS PRODUTO');
    SQL.Add('ON ITENS.COD_PROD = PRODUTO.COD_PROD');
    SQL.Add('WHERE CAPA.DATA >= :INI');
    SQL.Add('AND CAPA.DATA <= :FIM');
    SQL.Add('AND ITENS.VENDA_ENT > 0');
    SQL.Add('ORDER BY CAPA.SEQ');


    ParamByName('INI').AsDate := DtpInicial.Date;
    ParamByName('FIM').AsDate := DtpFinal.Date;


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

      if( (Layout.FieldByName('COD_FORNECEDOR').AsString = fornecedor) and
          (Layout.FieldByName('NUM_NF_FORN').AsString = nota) and
          (Layout.FieldByName('NUM_SERIE_NF').AsString = serie) ) then
      begin
          inc(count);
      end
      else
      begin
        fornecedor := Layout.FieldByName('COD_FORNECEDOR').AsString;
        nota := Layout.FieldByName('NUM_NF_FORN').AsString;
        serie := Layout.FieldByName('NUM_SERIE_NF').AsString;
        count := 1;
      end;
//
      Layout.FieldByName('NUM_ITEM').AsInteger := count;
//
      Layout.FieldByName('COD_PRODUTO').AsString := GerarPLU(QryPrincipal.FieldByName('COD_PRODUTO').AsString);

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;

end;

procedure TFrmByteCortez.GerarProdForn;
var
   TotalCount : Integer;
begin
  inherited;

  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
    SQL.Add('    PRODUTOS_FORNECEDOR.COD_PROD AS COD_PRODUTO,');
    SQL.Add('    PRODUTOS_FORNECEDOR.COD_FORNEC AS COD_FORNECEDOR,');
    SQL.Add('    PRODUTOS_FORNECEDOR.REFERENCIA AS DES_REFERENCIA,');
    SQL.Add('    '''' AS NUM_CGC,');
    SQL.Add('    0 AS COD_DIVISAO,');
    SQL.Add('    PRODUTO.UNIDADE AS DES_UNIDADE_COMPRA,');
    SQL.Add('    COALESCE(QTD_EMBALAGEM_COMPRA, 1) AS QTD_EMBALAGEM_COMPRA');
    SQL.Add('FROM PRODUTOS_FORNECEDOR');
    SQL.Add('LEFT JOIN PRODUTOS AS PRODUTO');
    SQL.Add('ON PRODUTOS_FORNECEDOR.COD_PROD = PRODUTO.COD_PROD');
    SQL.Add('');
    SQL.Add('UNION ALL');
    SQL.Add('');
    SQL.Add('SELECT');
    SQL.Add('    PRODUTO_FORNECEDOR.COD_PROD AS COD_PRODUTO,');
    SQL.Add('    0 AS COD_FORNECEDOR,');
    SQL.Add('    PRODUTO_FORNECEDOR.REFERENCIA AS DES_REFERENCIA,');
    SQL.Add('    PRODUTO_FORNECEDOR.CGC AS NUM_CGC,');
    SQL.Add('    0 AS COD_DIVISAO,');
    SQL.Add('    COALESCE(PRODUTO_FORNECEDOR.UNIDADE, ''UN'') AS DES_UNIDADE_COMPRA,');
    SQL.Add('    COALESCE(PRODUTO_FORNECEDOR.QTD_EMBALAGEM_COMPRA, 1) AS QTD_EMBALAGEM_COMPRA');
    SQL.Add('FROM TAB_PRODUTOS_FORNECEDOR_AUX AS PRODUTO_FORNECEDOR');
    SQL.Add('LEFT JOIN PRODUTOS');
    SQL.Add('ON PRODUTO_FORNECEDOR.COD_BARRA = PRODUTOS.COD_BARRA');
    SQL.Add('WHERE PRODUTOS.COD_PROD IS NULL');


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
      Layout.FieldByName('COD_PRODUTO').AsString := GerarPLU( Layout.FieldByName('COD_PRODUTO').AsString );

      Layout.FieldByName('NUM_CGC').AsString := StrRetNums(Layout.FieldByName('NUM_CGC').AsString);

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;

end;

procedure TFrmByteCortez.GerarProdLoja;
var
   TotalCount : integer;
begin
  inherited;
  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
    SQL.Add('    PRODUTO.COD_PROD AS COD_PRODUTO,');
    SQL.Add('    PRODUTO.PRECO_CUST AS VAL_CUSTO_REP,');
    SQL.Add('    PRODUTO.PRECO_VEND AS VAL_VENDA,');
    SQL.Add('    0 AS VAL_OFERTA,');
    SQL.Add('    COALESCE(ESTOQUE.QUANTIDADE, 0) AS QTD_EST_VDA,');
    SQL.Add('    '''' AS TECLA_BALANCA,');
    SQL.Add('');
    SQL.Add('    CASE PRODUTO.ICMS');
    SQL.Add('        WHEN ''4,5'' THEN 38');
    SQL.Add('        WHEN ''F1'' THEN 13');
    SQL.Add('        WHEN ''FF'' THEN 13');
    SQL.Add('        WHEN ''II'' THEN 1');
    SQL.Add('        WHEN ''T01'' THEN 4');
    SQL.Add('        WHEN ''T02'' THEN 2');
    SQL.Add('        WHEN ''T03'' THEN 3');
    SQL.Add('        WHEN ''T04'' THEN 5');
    SQL.Add('    END AS COD_TRIBUTACAO,');
    SQL.Add('');
    SQL.Add('    PRODUTO.MARGEM AS VAL_MARGEM,');
    SQL.Add('    1 AS QTD_ETIQUETA,');
    SQL.Add('');
    SQL.Add('    CASE PRODUTO.ICMS');
    SQL.Add('        WHEN ''4,5'' THEN 38');
    SQL.Add('        WHEN ''F1'' THEN 13');
    SQL.Add('        WHEN ''FF'' THEN 13');
    SQL.Add('        WHEN ''II'' THEN 1');
    SQL.Add('        WHEN ''T01'' THEN 4');
    SQL.Add('        WHEN ''T02'' THEN 2');
    SQL.Add('        WHEN ''T03'' THEN 3');
    SQL.Add('        WHEN ''T04'' THEN 5');
    SQL.Add('    END AS COD_TRIB_ENTRADA,');
    SQL.Add('');
    SQL.Add('    CASE PRODUTO.STATUS');
    SQL.Add('        WHEN ''A'' THEN ''N''');
    SQL.Add('        ELSE ''S''');
    SQL.Add('    END AS FLG_INATIVO,');
    SQL.Add('');
    SQL.Add('    PRODUTO.COD_PROD AS COD_PRODUTO_ANT,');
    SQL.Add('    ');
    SQL.Add('    CASE ');
    SQL.Add('        WHEN CAST(COALESCE(PRODUTO.CLASFISCAL, 0) AS INTEGER) = 0 THEN ''99999999''');
    SQL.Add('        ELSE TRIM(PRODUTO.CLASFISCAL)');
    SQL.Add('    END AS NUM_NCM,');
    SQL.Add('');
    SQL.Add('    0 AS TIPO_NCM,');
    SQL.Add('    0 AS VAL_VENDA_2,');
    SQL.Add('    '''' AS DTA_VALIDA_OFERTA,');
    SQL.Add('    PRODUTO.QUANT_MIN AS QTD_EST_MINIMO,');
    SQL.Add('    NULL AS COD_VASILHAME,');
    SQL.Add('    ''N'' AS FORA_LINHA,');
    SQL.Add('    0 AS QTD_PRECO_DIF,');
    SQL.Add('    0 AS VAL_FORCA_VDA,');
    SQL.Add('');
    SQL.Add('    CASE');
    SQL.Add('        WHEN COALESCE(PRODUTO.CEST, '''') = '''' THEN ''9999999''');
    SQL.Add('        ELSE PRODUTO.CEST');
    SQL.Add('    END AS NUM_CEST,');
    SQL.Add('');
    SQL.Add('    0 AS PER_IVA,');
    SQL.Add('    0 AS PER_FCP_ST');
    SQL.Add('FROM PRODUTOS AS PRODUTO');
    SQL.Add('LEFT JOIN ESTOQUEFILIAL AS ESTOQUE');
    SQL.Add('ON PRODUTO.COD_PROD = ESTOQUE.COD_PROD');
    SQL.Add('');
    SQL.Add('UNION ALL');
    SQL.Add('');
    SQL.Add('SELECT');
    SQL.Add('    100000 + PRODUTO.COD_PROD AS COD_PRODUTO,');
    SQL.Add('    PRODUTO.PRECO_CUST AS VAL_CUSTO_REP,');
    SQL.Add('    PRODUTO.PRECO_VEND AS VAL_VENDA,');
    SQL.Add('    0 AS VAL_OFERTA,');
    SQL.Add('    PRODUTO.QTD_EST_ATUAL AS QTD_EST_VDA,');
    SQL.Add('    '''' AS TECLA_BALANCA,');
    SQL.Add('');
    SQL.Add('    CASE PRODUTO.ICMS');
    SQL.Add('        WHEN ''4,5'' THEN 38');
    SQL.Add('        WHEN ''F1'' THEN 13');
    SQL.Add('        WHEN ''FF'' THEN 13');
    SQL.Add('        WHEN ''II'' THEN 1');
    SQL.Add('        WHEN ''T03'' THEN 4');
    SQL.Add('        WHEN ''T04'' THEN 2');
    SQL.Add('        WHEN ''T05'' THEN 3');
    SQL.Add('        WHEN ''T06'' THEN 5');
    SQL.Add('        WHEN ''T07'' THEN 25');
    SQL.Add('    END AS COD_TRIBUTACAO,');
    SQL.Add('');
    SQL.Add('    PRODUTO.MARGEM AS VAL_MARGEM,');
    SQL.Add('    1 AS QTD_ETIQUETA,');
    SQL.Add('');
    SQL.Add('    CASE PRODUTO.ICMS');
    SQL.Add('        WHEN ''4,5'' THEN 38');
    SQL.Add('        WHEN ''F1'' THEN 13');
    SQL.Add('        WHEN ''FF'' THEN 13');
    SQL.Add('        WHEN ''II'' THEN 1');
    SQL.Add('        WHEN ''T03'' THEN 4');
    SQL.Add('        WHEN ''T04'' THEN 2');
    SQL.Add('        WHEN ''T05'' THEN 3');
    SQL.Add('        WHEN ''T06'' THEN 5');
    SQL.Add('        WHEN ''T07'' THEN 25');
    SQL.Add('    END AS COD_TRIB_ENTRADA,');
    SQL.Add('');
    SQL.Add('    CASE PRODUTO.STATUS');
    SQL.Add('        WHEN ''A'' THEN ''N''');
    SQL.Add('        ELSE ''S''');
    SQL.Add('    END AS FLG_INATIVO,');
    SQL.Add('');
    SQL.Add('    PRODUTO.COD_PROD AS COD_PRODUTO_ANT,');
    SQL.Add('    ');
    SQL.Add('    CASE ');
    SQL.Add('        WHEN CAST(COALESCE(PRODUTO.CLASFISCAL, 0) AS INTEGER) = 0 THEN ''99999999''');
    SQL.Add('        ELSE TRIM(PRODUTO.CLASFISCAL)');
    SQL.Add('    END AS NUM_NCM,');
    SQL.Add('');
    SQL.Add('    0 AS TIPO_NCM,');
    SQL.Add('    0 AS VAL_VENDA_2,');
    SQL.Add('    '''' AS DTA_VALIDA_OFERTA,');
    SQL.Add('    PRODUTO.QUANT_MIN AS QTD_EST_MINIMO,');
    SQL.Add('    NULL AS COD_VASILHAME,');
    SQL.Add('    ''N'' AS FORA_LINHA,');
    SQL.Add('    0 AS QTD_PRECO_DIF,');
    SQL.Add('    0 AS VAL_FORCA_VDA,');
    SQL.Add('');
    SQL.Add('    CASE');
    SQL.Add('        WHEN COALESCE(PRODUTO.CEST, '''') = '''' THEN ''9999999''');
    SQL.Add('        WHEN PRODUTO.CEST = ''FF'' THEN ''9999999''');
    SQL.Add('        ELSE PRODUTO.CEST');
    SQL.Add('    END AS NUM_CEST,');
    SQL.Add('');
    SQL.Add('    0 AS PER_IVA,');
    SQL.Add('    0 AS PER_FCP_ST');
    SQL.Add('FROM TAB_PRODUTO_AUX AS PRODUTO');
    SQL.Add('LEFT JOIN PRODUTOS');
    SQL.Add('ON PRODUTO.COD_BARRA = PRODUTOS.COD_BARRA');
    SQL.Add('WHERE PRODUTOS.COD_PROD IS NULL');

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

      Layout.FieldByName('COD_PRODUTO').AsString := GerarPLU( Layout.FieldByName('COD_PRODUTO').AsString );
      Layout.FieldByName('COD_PRODUTO_ANT').AsString := GerarPLU(QryPrincipal.FieldByName('COD_PRODUTO_ANT').AsString);

      Layout.FieldByName('NUM_NCM').AsString := StrRetNums(Layout.FieldByName('NUM_NCM').AsString);

      Layout.FieldByName('NUM_CEST').AsString := StrRetNums( Layout.FieldByName('NUM_CEST').AsString );

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

procedure TFrmByteCortez.GerarProdSimilar;
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
