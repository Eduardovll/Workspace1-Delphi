unit UFrmVRAlbano;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, UFrmModelo, Data.DBXOracle, Data.DB,
  Data.SqlExpr, Vcl.Menus, Vcl.StdCtrls, Vcl.Buttons, Vcl.ComCtrls,
  Vcl.Imaging.pngimage, Vcl.ExtCtrls, Data.DBXFirebird, Data.Win.ADODB,
  Vcl.Grids, Vcl.DBGrids, Datasnap.Provider, Datasnap.DBClient, dxGDIPlusClasses,
  Math;

type
  TFrmVRAlbano = class(TFrmModeloSis)
    btnGeraCest: TButton;
    BtnAmarrarCest: TButton;
    ADOPostgre: TADOConnection;
    QryPrincipal2: TADOQuery;
    procedure btnGeraCestClick(Sender: TObject);
    procedure BtnAmarrarCestClick(Sender: TObject);
    procedure BtnGerarClick(Sender: TObject);
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
  FrmVRAlbano: TFrmVRAlbano;
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


procedure TFrmVRAlbano.GerarProducao;
begin
  inherited;
  with QryPrincipal2 do
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

      Layout.SetValues(QryPrincipal2, NumLinha, RecordCount);

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;
end;

procedure TFrmVRAlbano.GerarProduto;
var
 cod_produto : Integer;
 saldo : currency;
 QryComponente : TADOQuery;
 CompAcum, CompOrig, CompAqui : Currency;
 CodBem : Integer;
begin
  inherited;

  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
SQL.Add('  PRODUTO.ID AS COD_PRODUTO,');
SQL.Add('  BARRAS.EAN AS COD_BARRA_PRINCIPAL,');
SQL.Add('  PRODUTO.DESCRICAOREDUZIDA AS DES_REDUZIDA,');
SQL.Add('  PRODUTO.DESCRICAOCOMPLETA AS DES_PRODUTO,');
SQL.Add('  PRODUTO.QTDEMBALAGEM AS QTD_EMBALAGEM_COMPRA,');
SQL.Add('  TIPOEMBALAGEM.DESCRICAO AS DES_UNIDADE_COMPRA,');
SQL.Add('  PRODUTO.QTDEMBALAGEM AS QTD_EMBALAGEM_VENDA,');
SQL.Add('  TIPOEMBALAGEM.DESCRICAO AS DES_UNIDADE_VENDA,');
SQL.Add('  0 AS TIPO_IPI,');
SQL.Add('  0 AS VAL_IPI,');
SQL.Add('  PRODUTO.MERCADOLOGICO1 AS COD_SECAO,');
SQL.Add('  PRODUTO.MERCADOLOGICO2 AS COD_GRUPO,');
SQL.Add('  PRODUTO.MERCADOLOGICO3 AS COD_SUB_GRUPO,');
SQL.Add('  PRODUTO.ID_FAMILIAPRODUTO AS COD_PRODUTO_SIMILAR,');
SQL.Add('');
SQL.Add('  CASE');
SQL.Add('      WHEN COALESCE(PRODUTO.PESAVEL, ''false'') = ''true'' AND UPPER(TIPOEMBALAGEM.DESCRICAO) = ''KG'' THEN ''S''');
SQL.Add('      ELSE ''N''');
SQL.Add('  END AS IPV,');
SQL.Add('');
SQL.Add('  PRODUTO.VALIDADE AS DIAS_VALIDADE,');
SQL.Add('  0 AS TIPO_PRODUTO,');
SQL.Add('');
SQL.Add('  CASE PRODUTO.ID_TIPOPISCOFINS');
SQL.Add('    WHEN ''0'' THEN ''N''');
SQL.Add('    ELSE ''S''');
SQL.Add('  END AS FLG_NAO_PIS_COFINS,');
SQL.Add('');
SQL.Add('  CASE PRODUTO.PESAVEL');
SQL.Add('    WHEN ''true'' THEN ''S''');
SQL.Add('    ELSE ''N''');
SQL.Add('  END AS FLG_ENVIA_BALANCA,');
SQL.Add('');
SQL.Add('  CASE PRODUTO.ID_TIPOPISCOFINS');
SQL.Add('    WHEN ''0'' THEN -1');
SQL.Add('    WHEN ''2'' THEN 2');
SQL.Add('    WHEN ''3'' THEN 1');
SQL.Add('    WHEN ''4'' THEN 4');
SQL.Add('    WHEN ''7'' THEN 0');
SQL.Add('    ELSE ''0''');
SQL.Add('  END AS TIPO_NAO_PIS_COFINS,');
SQL.Add('  ');
SQL.Add('  0 AS TIPO_EVENTO,');
SQL.Add('  0 AS COD_ASSOCIADO,');
SQL.Add('  '''' AS DES_OBSERVACAO,');
SQL.Add('  COALESCE(NUTRICIONAL.ID_NUTRICIONALTOLEDO, 0) AS COD_INFO_NUTRICIONAL,');
SQL.Add('  0 AS COD_INFO_RECEITA,');
SQL.Add('COALESCE(LPAD(CAST(PRODUTO.TIPONATUREZARECEITA AS VARCHAR(3)), 3, ''000''), ''999'') AS COD_TAB_SPED,');
SQL.Add('  ''N'' AS FLG_ALCOOLICO,');
SQL.Add('  0 AS TIPO_ESPECIE,');
SQL.Add('  0 AS COD_CLASSIF,');
SQL.Add('  1 AS VAL_VDA_PESO_BRUTO,');
SQL.Add('  1 AS VAL_PESO_EMB,');
SQL.Add('  0 AS TIPO_EXPLOSAO_COMPRA,');
SQL.Add('  '''' AS DTA_INI_OPER,');
SQL.Add('  '''' AS DES_PLAQUETA,');
SQL.Add('  '''' AS MES_ANO_INI_DEPREC,');
SQL.Add('  0 AS TIPO_BEM,');
SQL.Add('  0 AS COD_FORNECEDOR,');
SQL.Add('  0 AS NUM_NF,');
SQL.Add('  '''' AS DTA_ENTRADA,');
SQL.Add('  0 AS COD_NAT_BEM,');
SQL.Add('  0 AS VAL_ORIG_BEM');
SQL.Add('FROM PUBLIC.PRODUTO');
SQL.Add('LEFT JOIN PUBLIC.TIPOEMBALAGEM');
SQL.Add('ON PRODUTO.ID_TIPOEMBALAGEM = TIPOEMBALAGEM.ID');
SQL.Add('LEFT JOIN PUBLIC.NUTRICIONALTOLEDOITEM AS NUTRICIONAL');
SQL.Add('ON PRODUTO.ID = NUTRICIONAL.ID_PRODUTO');
SQL.Add('LEFT JOIN (');
SQL.Add('  SELECT ');
SQL.Add('    ID_PRODUTO,');
SQL.Add('    MAX(CODIGOBARRAS) AS EAN');
SQL.Add('  FROM PUBLIC.PRODUTOAUTOMACAO ');
SQL.Add('  GROUP BY ID_PRODUTO');
SQL.Add(') AS BARRAS');
SQL.Add('ON PRODUTO.ID = BARRAS.ID_PRODUTO');

    Open;
    First;
    NumLinha := 0;


    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);
      Layout.SetValues(QryPrincipal2, NumLinha, RecordCount);

//      inc(cod_produto);

      Layout.FieldByName('COD_PRODUTO').AsString := GerarPLU( Layout.FieldByName('COD_PRODUTO').AsString );
      Layout.FieldByName('COD_PRODUTO_SIMILAR').AsString := GerarPLU( Layout.FieldByName('COD_PRODUTO_SIMILAR').AsString );

      if ( Length(TiraZerosEsquerda(Layout.FieldByName('COD_BARRA_PRINCIPAL').AsString)) < 8 ) then
        Layout.FieldByName('COD_BARRA_PRINCIPAL').AsString := GerarPLU( Layout.FieldByName('COD_BARRA_PRINCIPAL').AsString );

      if( not CodBarrasValido(Layout.FieldByName('COD_BARRA_PRINCIPAL').AsString) ) then
        Layout.FieldByName('COD_BARRA_PRINCIPAL').AsString := '';

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
    end;
    Next;
    end;
    Close
  end;
end;

procedure TFrmVRAlbano.GerarScriptAmarrarCEST;
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

procedure TFrmVRAlbano.GerarScriptCEST;
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

procedure TFrmVRAlbano.GerarSecao;
begin
  inherited;

  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
    SQL.Add('  MERCADOLOGICO1 AS COD_SECAO,');
    SQL.Add('  DESCRICAO AS DES_SECAO,');
    SQL.Add('  0 AS VAL_META');
    SQL.Add('FROM PUBLIC.MERCADOLOGICO');
    SQL.Add('WHERE NIVEL = 1');

    Open;

    First;
    NumLinha := 0;

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);

      Layout.SetValues(QryPrincipal2, NumLinha, RecordCount);

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;

end;

procedure TFrmVRAlbano.GerarSubGrupo;
begin
  inherited;

  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
    SQL.Add('  MERCADOLOGICO1 AS COD_SECAO,');
    SQL.Add('  MERCADOLOGICO2 AS COD_GRUPO,');
    SQL.Add('  MERCADOLOGICO3 AS COD_SUB_GRUPO,');
    SQL.Add('  DESCRICAO AS DES_SUB_GRUPO,');
    SQL.Add('  0 AS VAL_META,');
    SQL.Add('  0 AS VAL_MARGEM_REF,');
    SQL.Add('  0 AS QTD_DIA_SEGURANCA,');
    SQL.Add('  ''N'' AS FLG_ALCOOLICO');
    SQL.Add('FROM PUBLIC.MERCADOLOGICO');
    SQL.Add('WHERE NIVEL = 3');

    Open;

    First;
    NumLinha := 0;

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);

      Layout.SetValues(QryPrincipal2, NumLinha, RecordCount);

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;

end;

procedure TFrmVRAlbano.GerarTransportadora;
begin
  inherited;
  with QryPrincipal2 do
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

      Layout.SetValues(QryPrincipal2, NumLinha, RecordCount);

      Layout.FieldByName('DES_OBSERVACAO').AsString := StrReplace(StrLBReplace(FieldByName('DES_OBSERVACAO').AsString), '\n', '');
      Layout.FieldByName('DES_EMAIL').AsString := StrReplace(StrLBReplace(FieldByName('DES_EMAIL').AsString), '\n', '');

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;
end;

procedure TFrmVRAlbano.GerarVenda;
begin
  inherited;

  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
    SQL.Add('  VENDAITEM.ID_PRODUTO AS COD_PRODUTO,');
    SQL.Add('  VENDA.ID_LOJA AS COD_LOJA,');
    SQL.Add('  0 AS IND_TIPO,');
    SQL.Add('  VENDA.ECF AS NUM_PDV,');
    SQL.Add('  VENDAITEM.QUANTIDADE AS QTD_TOTAL_PRODUTO,');
    SQL.Add('  VENDAITEM.VALORTOTAL - VENDAITEM.VALORDESCONTO + VENDAITEM.VALORACRESCIMO AS VAL_TOTAL_PRODUTO,');
    SQL.Add('  VENDAITEM.PRECOVENDA AS VAL_PRECO_VENDA,');
    SQL.Add('  VENDAITEM.CUSTOCOMIMPOSTO AS VAL_CUSTO_REP,');
    SQL.Add('  VENDA.DATA AS DTA_SAIDA,');
    SQL.Add('  CONCAT( LPAD(CAST(EXTRACT(MONTH FROM VENDA.DATA) AS VARCHAR(2)), 2, ''0''), CAST(EXTRACT(YEAR FROM VENDA.DATA) AS VARCHAR(4)) ) AS DTA_MENSAL,');
    SQL.Add('  VENDA.NUMEROCUPOM AS NUM_IDENT,');
    SQL.Add('  VENDAITEM.CODIGOBARRAS AS COD_EAN,');
    SQL.Add('  CONCAT( LPAD(CAST(EXTRACT(HOUR FROM VENDA.HORAINICIO) AS VARCHAR(2)), 2, ''0''), LPAD(CAST(EXTRACT(MINUTE FROM VENDA.HORAINICIO) AS VARCHAR(2)), 2, ''0'') ) AS DES_HORA,');
    SQL.Add('  COALESCE(VENDA.ID_CLIENTEPREFERENCIAL, 0) AS COD_CLIENTE,');
    SQL.Add('  1 AS COD_ENTIDADE,');
    SQL.Add('  -- VENDAITEM.VALORTOTAL AS VAL_BASE_ICMS,');
    SQL.Add('  0 AS VAL_BASE_ICMS,');
    SQL.Add('  ');
    SQL.Add('  CASE PRODUTOALIQUOTA.ID_ALIQUOTADEBITO');
    SQL.Add('    WHEN ''0'' THEN ''T''');
    SQL.Add('    WHEN ''2'' THEN ''T''');
    SQL.Add('    WHEN ''4'' THEN ''T''');
    SQL.Add('    WHEN ''5'' THEN ''T''');
    SQL.Add('    WHEN ''6'' THEN ''I''');
    SQL.Add('    WHEN ''7'' THEN ''F''');
    SQL.Add('    WHEN ''21'' THEN ''T'' ');
    SQL.Add('    WHEN ''29'' THEN ''T'' ');
    SQL.Add('    WHEN ''30'' THEN ''T''');
    SQL.Add('    WHEN ''38'' THEN ''T''');
    SQL.Add('    WHEN ''40'' THEN ''T''');
    SQL.Add('    WHEN ''75'' THEN ''I''');
    SQL.Add('    WHEN ''99'' THEN ''I''');
    SQL.Add('  END AS DES_SITUACAO_TRIB,');
    SQL.Add('');
    SQL.Add('  0 AS VAL_ICMS,');
    SQL.Add('  -- COALESCE(VENDAITEM.VALORICMSDESONERADO, 0) AS VAL_ICMS,');
    SQL.Add('  VENDA.NUMEROCUPOM AS NUM_CUPOM_FISCAL,');
    SQL.Add('  VENDAITEM.PRECOVENDA AS VAL_VENDA_PDV,');
    SQL.Add('');
    SQL.Add('  CASE PRODUTOALIQUOTA.ID_ALIQUOTADEBITO');
    SQL.Add('    WHEN ''0'' THEN 2');
    SQL.Add('    WHEN ''2'' THEN 4');
    SQL.Add('    WHEN ''4'' THEN 8');
    SQL.Add('    WHEN ''5'' THEN 6');
    SQL.Add('    WHEN ''6'' THEN 1');
    SQL.Add('    WHEN ''7'' THEN 25');
    SQL.Add('    WHEN ''21'' THEN 39 ');
    SQL.Add('    WHEN ''29'' THEN 39 ');
    SQL.Add('    WHEN ''30'' THEN 40');
    SQL.Add('    WHEN ''38'' THEN 3');
    SQL.Add('    WHEN ''40'' THEN 41');
    SQL.Add('    WHEN ''75'' THEN 1');
    SQL.Add('    WHEN ''99'' THEN 1');
    SQL.Add('  END AS COD_TRIBUTACAO,');
    SQL.Add('');
    SQL.Add('  CASE ');
    SQL.Add('    WHEN VENDA.CANCELADO OR VENDAITEM.CANCELADO = TRUE THEN ''S''    ');
    SQL.Add('    ELSE ''N''');
    SQL.Add('  END AS FLG_CUPOM_CANCELADO,');
    SQL.Add('');
    SQL.Add('  CONCAT( LPAD(CAST(PRODUTO.NCM1 AS VARCHAR(4)), 4, ''0''), LPAD(CAST(PRODUTO.NCM2 AS VARCHAR(2)), 2, ''0''), LPAD(CAST(PRODUTO.NCM3 AS VARCHAR(2)), 2, ''0'') ) AS NUM_NCM,');
    SQL.Add('COALESCE(LPAD(CAST(PRODUTO.TIPONATUREZARECEITA AS VARCHAR(3)), 3, ''000''), ''999'') AS COD_TAB_SPED,');
    SQL.Add('');
    SQL.Add('  CASE PRODUTO.ID_TIPOPISCOFINS');
    SQL.Add('    WHEN ''0'' THEN ''N''');
    SQL.Add('    ELSE ''S''');
    SQL.Add('  END AS FLG_NAO_PIS_COFINS,');
    SQL.Add('');
    SQL.Add('  CASE PRODUTO.ID_TIPOPISCOFINS');
    SQL.Add('    WHEN ''0'' THEN -1');
    SQL.Add('    WHEN ''2'' THEN 2');
    SQL.Add('    WHEN ''3'' THEN 1');
    SQL.Add('    WHEN ''4'' THEN 4');
    SQL.Add('    WHEN ''7'' THEN 0');
    SQL.Add('    ELSE ''0''');
    SQL.Add('  END AS TIPO_NAO_PIS_COFINS,');
    SQL.Add('  ''N'' AS FLG_ONLINE,');
    SQL.Add('');
    SQL.Add('  CASE VENDAITEM.OFERTA');
    SQL.Add('    WHEN true THEN ''S''');
    SQL.Add('    ELSE ''N''');
    SQL.Add('  END AS FLG_OFERTA,');
    SQL.Add('');
    SQL.Add('  0 AS COD_ASSOCIADO');
    SQL.Add('FROM PDV.VENDAITEM AS VENDAITEM');
    SQL.Add('');
    SQL.Add('LEFT JOIN PUBLIC.PRODUTO AS PRODUTO');
    SQL.Add('ON VENDAITEM.ID_PRODUTO = PRODUTO.ID');
    SQL.Add('');
    SQL.Add('LEFT JOIN PUBLIC.PRODUTOALIQUOTA AS PRODUTOALIQUOTA');
    SQL.Add('ON PRODUTO.ID = PRODUTOALIQUOTA.ID_PRODUTO');
    SQL.Add('');
    SQL.Add('LEFT JOIN PDV.VENDA AS VENDA');
    SQL.Add('ON VENDAITEM.ID_VENDA = VENDA.ID');
    SQL.Add('WHERE VENDA.DATA >= '''+FormatDateTime('yyyy-mm-dd',DtpInicial.Date)+''' ');
    SQL.Add('AND VENDA.DATA <= '''+FormatDateTime('yyyy-mm-dd',DtpFinal.Date)+''' ');


    Open;

    First;

    NumLinha := 0;


    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);

      Layout.SetValues(QryPrincipal2, NumLinha, RecordCount);
      Layout.FieldByName('COD_PRODUTO').AsString := GerarPLU(QryPrincipal2.FieldByName('COD_PRODUTO').AsString);

      Layout.FieldByName('DTA_SAIDA').AsDateTime := QryPrincipal2.FieldByName('DTA_SAIDA').AsDateTime;

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;
end;

procedure TFrmVRAlbano.BtnAmarrarCestClick(Sender: TObject);
begin
  inherited;
    inherited;
  FlgGeraAmarrarCest := True;
  BtnGerar.Click;
  FlgGeraAmarrarCest := False;
end;

procedure TFrmVRAlbano.btnGeraCestClick(Sender: TObject);
begin
  inherited;
  FlgGeraCest := True;
  BtnGerar.Click;
  FlgGeraCest := False;
end;

procedure TFrmVRAlbano.BtnGerarClick(Sender: TObject);
begin
    ADOPostgre.Connected := false;
    ADOPostgre.ConnectionString := 'Provider=MSDASQL.1;Password='+ edtSenhaOracle.Text +';Persist Security Info=True;User ID='+ edtSchema.Text +';Data Source='+ edtInst.Text +';Initial Catalog='+ edtInst.Text +';';

    ADOPostgre.Connected := true;
  inherited;
    ADOPostgre.Connected := false;
end;

procedure TFrmVRAlbano.GerarCest;
var
   count : integer;
begin
  inherited;
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT DISTINCT');
    SQL.Add('  0 AS COD_CEST,');
    SQL.Add('  CASE');
    SQL.Add('    WHEN CEST.ID IS NULL THEN ''9999999''');
    SQL.Add('    WHEN CONCAT(LPAD(CAST(CEST1 AS VARCHAR(2)), 2, ''0''), LPAD(CAST(CEST2 AS VARCHAR(3)), 3,''0''), LPAD(CAST(CEST3 AS VARCHAR(2)), 2, ''0'')) = ''0000000'' THEN ''9999999''');
    SQL.Add('    ELSE CONCAT(LPAD(CAST(CEST1 AS VARCHAR(2)), 2, ''0''), LPAD(CAST(CEST2 AS VARCHAR(3)), 3,''0''), LPAD(CAST(CEST3 AS VARCHAR(2)), 2, ''0''))');
    SQL.Add('  END AS NUM_CEST,');
    SQL.Add('  COALESCE(CEST.DESCRICAO, ''A DEFINIR'') AS DES_CEST');
    SQL.Add('FROM PUBLIC.PRODUTO AS PRODUTO');
    SQL.Add('LEFT JOIN PUBLIC.CEST AS CEST');
    SQL.Add('ON PRODUTO.ID_CEST = CEST.ID');


    Open;
    First;

    count := 0;


    NumLinha := 0;

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);
      Inc(count);
      Layout.SetValues(QryPrincipal2, NumLinha, RecordCount);

      Layout.FieldByName('COD_CEST').AsInteger := count;

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;
end;

procedure TFrmVRAlbano.GerarCliente;
begin
  inherited;

  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
    SQL.Add('  CLIENTE.ID AS COD_CLIENTE,');
    SQL.Add('  CLIENTE.NOME AS DES_CLIENTE,');
    SQL.Add('');
    SQL.Add('  CASE CLIENTE.ID_TIPOINSCRICAO');
    SQL.Add('    WHEN 0 THEN LPAD(CAST(CLIENTE.CNPJ AS VARCHAR(14)), 14, ''0'')');
    SQL.Add('    ELSE LPAD(CAST(CLIENTE.CNPJ AS VARCHAR(11)), 11, ''0'')');
    SQL.Add('  END AS NUM_CGC,');
    SQL.Add('');
    SQL.Add('  CASE CLIENTE.ID_TIPOINSCRICAO');
    SQL.Add('    WHEN 0 THEN CLIENTE.INSCRICAOESTADUAL');
    SQL.Add('    ELSE ''''');
    SQL.Add('  END AS NUM_INSC_EST,');
    SQL.Add('');
    SQL.Add('  CLIENTE.ENDERECO AS DES_ENDERECO,');
    SQL.Add('  CLIENTE.BAIRRO AS DES_BAIRRO,');
    SQL.Add('  MUNICIPIO.DESCRICAO AS DES_CIDADE,');
    SQL.Add('  ESTADO.SIGLA AS DES_SIGLA,');
    SQL.Add('  CLIENTE.CEP AS NUM_CEP,');
    SQL.Add('  CLIENTE.TELEFONE AS NUM_FONE,');
    SQL.Add('  '''' AS NUM_FAX,');
    SQL.Add('  '''' AS DES_CONTATO,');
    SQL.Add('');
    SQL.Add('  CASE CLIENTE.SEXO');
    SQL.Add('    WHEN 1 THEN 0');
    SQL.Add('    ELSE 1');
    SQL.Add('  END AS FLG_SEXO,');
    SQL.Add('');
    SQL.Add('  0 AS VAL_LIMITE_CRETID, --');
    SQL.Add('  CLIENTE.VALORLIMITE AS VAL_LIMITE_CONV,');
    SQL.Add('  0 AS VAL_DEBITO,');
    SQL.Add('  CLIENTE.SALARIO + CLIENTE.OUTRARENDA AS VAL_RENDA,');
    SQL.Add('  0 AS COD_CONVENIO,');
    SQL.Add('  0 AS COD_STATUS_PDV,');
    SQL.Add('');
    SQL.Add('  CASE CLIENTE.ID_TIPOINSCRICAO');
    SQL.Add('    WHEN 0 THEN ''S''');
    SQL.Add('    ELSE ''N''');
    SQL.Add('  END AS FLG_EMPRESA,');
    SQL.Add('');
    SQL.Add('  ''N'' AS FLG_CONVENIO,');
    SQL.Add('  ''N'' AS MICRO_EMPRESA,');
    SQL.Add('  CLIENTE.DATACADASTRO AS DTA_CADASTRO,');
    SQL.Add('  CLIENTE.NUMERO AS NUM_ENDERECO,');
    SQL.Add('  CLIENTE.INSCRICAOESTADUAL AS NUM_RG,');
    SQL.Add('');
    SQL.Add('  CASE CLIENTE.ID_TIPOESTADOCIVIL');
    SQL.Add('    WHEN 2 THEN 1');
    SQL.Add('    WHEN 3 THEN 2');
    SQL.Add('    WHEN 4 THEN 4');
    SQL.Add('    WHEN 6 THEN 3');
    SQL.Add('    ELSE 0');
    SQL.Add('  END AS FLG_EST_CIVIL,');
    SQL.Add('');
    SQL.Add('  CLIENTE.CELULAR AS NUM_CELULAR,');
    SQL.Add('  '''' AS DTA_ALTERACAO,');
    SQL.Add('  CLIENTE.OBSERVACAO2 AS DES_OBSERVACAO,');
    SQL.Add('  CLIENTE.COMPLEMENTO AS DES_COMPLEMENTO,');
    SQL.Add('  CLIENTE.EMAIL AS DES_EMAIL,');
    SQL.Add('  CLIENTE.NOME AS DES_FANTASIA,');
    SQL.Add('  CLIENTE.DATANASCIMENTO AS DTA_NASCIMENTO,');
    SQL.Add('  CLIENTE.NOMEPAI AS DES_PAI,');
    SQL.Add('  CLIENTE.NOMEMAE AS DES_MAE,');
    SQL.Add('  CLIENTE.NOMECONJUGE AS DES_CONJUGE,');
    SQL.Add('  CLIENTE.CPFCONJUGE AS NUM_CPF_CONJUGE,');
    SQL.Add('  0 AS VAL_DEB_CONV,');
    SQL.Add('  ''N'' AS INATIVO,');
    SQL.Add('  '''' AS DES_MATRICULA,');
    SQL.Add('  ''N'' AS NUM_CGC_ASSOCIADO,');
    SQL.Add('  ''N'' AS FLG_PROD_RURAL,');
    SQL.Add('  CLIENTE.ID_TIPORESTRICAOCLIENTE AS COD_STATUS_PDV_CONV,');
    SQL.Add('  ''N'' AS FLG_ENVIA_CODIGO,');
    SQL.Add('  CLIENTE.DATANASCIMENTOCONJUGE AS DTA_NASC_CONJUGE,');
    SQL.Add('  0 AS COD_CLASSIF');
    SQL.Add('FROM PUBLIC.CLIENTEPREFERENCIAL AS CLIENTE');
    SQL.Add('LEFT JOIN PUBLIC.ESTADO AS ESTADO');
    SQL.Add('ON CLIENTE.ID_ESTADO = ESTADO.ID');
    SQL.Add('LEFT JOIN PUBLIC.MUNICIPIO AS MUNICIPIO');
    SQL.Add('ON CLIENTE.ID_MUNICIPIO = MUNICIPIO.ID');
    SQL.Add('');
    SQL.Add('UNION ALL');
    SQL.Add('');
    SQL.Add('SELECT');
    SQL.Add('  3700 + CLIENTE.ID AS COD_CLIENTE,');
    SQL.Add('  CLIENTE.NOME AS DES_CLIENTE,');
    SQL.Add('');
//    SQL.Add('  CASE CLIENTE.ID_TIPOINSCRICAO');
//    SQL.Add('    WHEN 0 THEN LPAD(CAST(CLIENTE.CNPJ AS VARCHAR(14)), 14, ''0'')');
//    SQL.Add('    ELSE LPAD(CAST(CLIENTE.CNPJ AS VARCHAR(11)), 11, ''0'')');
//    SQL.Add('  END AS NUM_CGC,');
    SQL.Add('  '''' AS NUM_CGC,');
    SQL.Add('');
    SQL.Add('  CASE CLIENTE.ID_TIPOINSCRICAO');
    SQL.Add('    WHEN 0 THEN CLIENTE.INSCRICAOESTADUAL');
    SQL.Add('    ELSE ''''');
    SQL.Add('  END AS NUM_INSC_EST,');
    SQL.Add('');
    SQL.Add('  CLIENTE.ENDERECO AS DES_ENDERECO,');
    SQL.Add('  CLIENTE.BAIRRO AS DES_BAIRRO,');
    SQL.Add('  MUNICIPIO.DESCRICAO AS DES_CIDADE,');
    SQL.Add('  ESTADO.SIGLA AS DES_SIGLA,');
    SQL.Add('  CLIENTE.CEP AS NUM_CEP,');
    SQL.Add('  CLIENTE.TELEFONE AS NUM_FONE,');
    SQL.Add('  '''' AS NUM_FAX,');
    SQL.Add('  '''' AS DES_CONTATO,');
    SQL.Add('');
    SQL.Add('  0 AS FLG_SEXO,');
    SQL.Add('');
    SQL.Add('  0 AS VAL_LIMITE_CRETID, --');
    SQL.Add('  0 AS VAL_LIMITE_CONV,');
    SQL.Add('  0 AS VAL_DEBITO,');
    SQL.Add('  0 AS VAL_RENDA,');
    SQL.Add('  0 AS COD_CONVENIO,');
    SQL.Add('  0 AS COD_STATUS_PDV,');
    SQL.Add('');
    SQL.Add('  CASE CLIENTE.ID_TIPOINSCRICAO');
    SQL.Add('    WHEN 0 THEN ''S''');
    SQL.Add('    ELSE ''N''');
    SQL.Add('  END AS FLG_EMPRESA,');
    SQL.Add('');
    SQL.Add('  ''N'' AS FLG_CONVENIO,');
    SQL.Add('  ''N'' AS MICRO_EMPRESA,');
    SQL.Add('  CLIENTE.DATACADASTRO AS DTA_CADASTRO,');
    SQL.Add('  CLIENTE.NUMERO AS NUM_ENDERECO,');
    SQL.Add('  CLIENTE.INSCRICAOESTADUAL AS NUM_RG,');
    SQL.Add('');
    SQL.Add('  0 AS FLG_EST_CIVIL,');
    SQL.Add('');
    SQL.Add('  '''' AS NUM_CELULAR,');
    SQL.Add('  '''' AS DTA_ALTERACAO,');
    SQL.Add('  '''' AS DES_OBSERVACAO,');
    SQL.Add('  CLIENTE.COMPLEMENTO AS DES_COMPLEMENTO,');
    SQL.Add('  '''' AS DES_EMAIL,');
    SQL.Add('  CLIENTE.NOME AS DES_FANTASIA,');
    SQL.Add('  NULL AS DTA_NASCIMENTO,');
    SQL.Add('  '''' AS DES_PAI,');
    SQL.Add('  '''' AS DES_MAE,');
    SQL.Add('  '''' AS DES_CONJUGE,');
    SQL.Add('  0 AS NUM_CPF_CONJUGE,');
    SQL.Add('  0 AS VAL_DEB_CONV,');
    SQL.Add('  ''N'' AS INATIVO,');
    SQL.Add('  '''' AS DES_MATRICULA,');
    SQL.Add('  ''N'' AS NUM_CGC_ASSOCIADO,');
    SQL.Add('  ''N'' AS FLG_PROD_RURAL,');
    SQL.Add('  0 AS COD_STATUS_PDV_CONV,');
    SQL.Add('  ''N'' AS FLG_ENVIA_CODIGO,');
    SQL.Add('  NULL AS DTA_NASC_CONJUGE,');
    SQL.Add('  0 AS COD_CLASSIF');
    SQL.Add('FROM PUBLIC.CLIENTEEVENTUAL AS CLIENTE');
    SQL.Add('LEFT JOIN PUBLIC.ESTADO AS ESTADO');
    SQL.Add('ON CLIENTE.ID_ESTADO = ESTADO.ID');
    SQL.Add('LEFT JOIN PUBLIC.MUNICIPIO AS MUNICIPIO');
    SQL.Add('ON CLIENTE.ID_MUNICIPIO = MUNICIPIO.ID');



    Open;
    First;

//    TotalCont := SetCountTotal(SQL.Text);
    NumLinha := 0;

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);


//      Layout.SetValues(QryPrincipal2, NumLinha, TotalCont);
      Layout.SetValues(QryPrincipal2, NumLinha, RecordCount);

      Layout.FieldByName('NUM_CGC').AsString := StrRetNums(Layout.FieldByName('NUM_CGC').AsString);

      if StrRetNums(Layout.FieldByName('NUM_RG').AsString) = '' then
        Layout.FieldByName('NUM_RG').AsString := ''
      else
        Layout.FieldByName('NUM_RG').AsString := StrRetNums(Layout.FieldByName('NUM_RG').AsString);

      Layout.FieldByName('NUM_CEP').AsString := StrRetNums(Layout.FieldByName('NUM_CEP').AsString);

//      if Layout.FieldByName('DTA_NASCIMENTO').AsString <> '' then
        Layout.FieldByName('DTA_NASCIMENTO').AsDateTime := FieldByName('DTA_NASCIMENTO').AsDateTime;

//      if Layout.FieldByName('DTA_CADASTRO').AsString <> '' then
        Layout.FieldByName('DTA_CADASTRO').AsDateTime := FieldByName('DTA_CADASTRO').AsDateTime;

      if Layout.FieldByName('DTA_ALTERACAO').AsString <> '' then
        Layout.FieldByName('DTA_ALTERACAO').AsDateTime := FieldByName('DTA_ALTERACAO').AsDateTime;

      Layout.FieldByName('NUM_FONE').AsString := StrRetNums( FieldByName('NUM_FONE').AsString );

      if Layout.FieldByName('FLG_EMPRESA').AsString = 'S' then
      begin
        if not ValidaCGC(Layout.FieldByName('NUM_CGC').AsString) then
          Layout.FieldByName('NUM_CGC').AsString := '';
      end
      else
        if not ValidaCpf(Layout.FieldByName('NUM_CGC').AsString) then
          Layout.FieldByName('NUM_CGC').AsString := '';

      if Layout.FieldByName('NUM_ENDERECO').AsString = '' then
        Layout.FieldByName('NUM_ENDERECO').AsString := '0';

      Layout.FieldByName('DES_OBSERVACAO').AsString := StrReplace(StrLBReplace(FieldByName('DES_OBSERVACAO').AsString), '\n', '');
      Layout.FieldByName('DES_EMAIL').AsString := StrReplace(StrLBReplace(FieldByName('DES_EMAIL').AsString), '\n', '');

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;

end;

procedure TFrmVRAlbano.GerarCodigoBarras;
var
 count : Integer;
 cod_antigo : string;
begin
  inherited;
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
    SQL.Add('  PRODUTOAUTOMACAO.ID_PRODUTO AS COD_PRODUTO,');
    SQL.Add('  PRODUTOAUTOMACAO.CODIGOBARRAS AS COD_EAN');
    SQL.Add('FROM PUBLIC.PRODUTOAUTOMACAO');
    Open;

    
    First;

//    TotalCont := SetCountTotal(SQL.Text);
    NumLinha := 0;

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);

      Layout.SetValues(QryPrincipal2, NumLinha, RecordCount);


      Layout.FieldByName('COD_PRODUTO').AsString := GerarPLU(FieldByName('COD_PRODUTO').AsString);
//
      if ( Length(TiraZerosEsquerda(Layout.FieldByName('COD_EAN').AsString)) < 8 ) then
        Layout.FieldByName('COD_EAN').AsString := GerarPLU( Layout.FieldByName('COD_EAN').AsString );
//
      if( not CodBarrasValido(Layout.FieldByName('COD_EAN').AsString) ) then
        Layout.FieldByName('COD_EAN').AsString := '';

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;
end;

procedure TFrmVRAlbano.GerarComposicao;
begin
  inherited;
  with QryPrincipal2 do
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

      Layout.SetValues(QryPrincipal2, NumLinha, RecordCount);

//      Layout.FieldByName('COD_PRODUTO').AsString := GerarPLU( Layout.FieldByName('COD_PRODUTO').AsString );
//
//      Layout.FieldByName('COD_PRODUTO_COMP').AsString := GerarPLU( Layout.FieldByName('COD_PRODUTO_COMP').AsString );

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;
end;

procedure TFrmVRAlbano.GerarCondPagCli;
begin
  inherited;

  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
    SQL.Add('  CLIENTE.ID AS COD_CLIENTE,');
    SQL.Add('  CASE ');
    SQL.Add('    WHEN CLIENTE.VENCIMENTOCREDITOROTATIVO > 0 THEN CLIENTE.VENCIMENTOCREDITOROTATIVO ');
    SQL.Add('    ELSE 30');
    SQL.Add('  END AS NUM_CONDICAO,');
    SQL.Add('  CASE ');
    SQL.Add('    WHEN CLIENTE.VENCIMENTOCREDITOROTATIVO > 0 THEN 13');
    SQL.Add('    ELSE 2');
    SQL.Add('  END AS COD_CONDICAO,');
    SQL.Add('  CASE ');
    SQL.Add('    WHEN CLIENTE.VENCIMENTOCREDITOROTATIVO > 0 THEN 4');
    SQL.Add('    ELSE 1');
    SQL.Add('  END AS COD_ENTIDADE');
    SQL.Add('FROM PUBLIC.CLIENTEPREFERENCIAL AS CLIENTE');
    SQL.Add('');
    SQL.Add('UNION ALL');
    SQL.Add('');
    SQL.Add('SELECT');
    SQL.Add('  3700 + CLIENTE.ID AS COD_CLIENTE,');
    SQL.Add('  30 AS NUM_CONDICAO,');
    SQL.Add('  2 AS COD_CONDICAO,');
    SQL.Add('  1 AS COD_ENTIDADE');
    SQL.Add('FROM PUBLIC.CLIENTEEVENTUAL AS CLIENTE');




    Open;

    First;
    NumLinha := 0;

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);

      Layout.SetValues(QryPrincipal2, NumLinha, RecordCount);

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;

end;

procedure TFrmVRAlbano.GerarCondPagForn;
begin
  inherited;

  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('select');
    SQL.Add('  FORNECEDOR.ID AS COD_FORNECEDOR,');
    SQL.Add('  30 AS NUM_CONDICAO,');
    SQL.Add('  2 AS COD_CONDICAO,');
    SQL.Add('  8 AS COD_ENTIDADE,');
    SQL.Add('  '''' AS NUM_CGC');
    SQL.Add('FROM PUBLIC.FORNECEDOR AS FORNECEDOR   ');

    Open;

    First;

    NumLinha := 0;

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);

      Layout.SetValues(QryPrincipal2, NumLinha, RecordCount);

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;

end;

procedure TFrmVRAlbano.GerarDecomposicao;
begin
  inherited;

  with QryPrincipal2 do
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

      Layout.SetValues(QryPrincipal2, NumLinha, RecordCount);

//      Layout.FieldByName('COD_PRODUTO').AsString := GerarPLU( Layout.FieldByName('COD_PRODUTO').AsString );
//
//      Layout.FieldByName('COD_PRODUTO_DECOM').AsString := GerarPLU( Layout.FieldByName('COD_PRODUTO_DECOM').AsString );

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;

end;

procedure TFrmVRAlbano.GerarDivisaoForn;
begin
  inherited;
    with QryPrincipal2 do
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

      Layout.SetValues(QryPrincipal2, NumLinha, RecordCount);

      Layout.FieldByName('DES_OBSERVACAO').AsString := StrReplace(StrLBReplace(FieldByName('DES_OBSERVACAO').AsString), '\n', '');
      Layout.FieldByName('NUM_CEP').AsString := StrRetNums(Layout.FieldByName('NUM_CEP').AsString);

      Layout.FieldByName('DES_EMAIL').AsString := StrReplace(StrLBReplace(FieldByName('DES_EMAIL').AsString), '\n', '');

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;
end;

procedure TFrmVRAlbano.GerarFinanceiro(Tipo, Situacao: Integer);
begin
  inherited;
  if Tipo = 1 then
    GerarFinanceiroPagar(IntToStr(Situacao));

  if Tipo = 2 then
    GerarFinanceiroReceber(IntToStr(Situacao));

  if Tipo = 3 then
    GerarFinanceiroReceberCartao;

end;

procedure TFrmVRAlbano.GerarFinanceiroPagar(Aberto: String);
begin
  inherited;
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
    SQL.Add('  1 AS TIPO_PARCEIRO,');
    SQL.Add('  DOCUMENTO.ID_FORNECEDOR AS COD_PARCEIRO,');
    SQL.Add('  0 AS TIPO_CONTA,');
    SQL.Add('  CASE PARCELA.ID_TIPOPAGAMENTO');
    SQL.Add('    WHEN 6 THEN 8');
    SQL.Add('    WHEN 1 THEN 8');
    SQL.Add('    WHEN 5 THEN 1');
    SQL.Add('    WHEN 0 THEN 7');
    SQL.Add('    WHEN 4 THEN 8');
    SQL.Add('  END AS COD_ENTIDADE,');
    SQL.Add('  CASE DOCUMENTO.ID_PAGAROUTRASDESPESAS');
    SQL.Add('    WHEN NULL THEN CAST(DOCUMENTO.NUMERODOCUMENTO AS VARCHAR)');
    SQL.Add('    ELSE CONCAT(DOCUMENTO.ID_PAGAROUTRASDESPESAS, DOCUMENTO.NUMERODOCUMENTO)');
    SQL.Add('  END AS NUM_DOCTO,');
    SQL.Add('  0 AS COD_BANCO,');
    SQL.Add('  '''' AS DES_BANCO,');
    SQL.Add('  DOCUMENTO.DATAEMISSAO AS DTA_EMISSAO,');
    SQL.Add('  PARCELA.DATAVENCIMENTO AS DTA_VENCIMENTO,');
    SQL.Add('  PARCELA.VALOR AS VAL_PARCELA,');
    SQL.Add('  PARCELA.VALORACRESCIMO AS VAL_JUROS,');
    SQL.Add('  0 AS VAL_DESCONTO,');
    SQL.Add('');
    SQL.Add('  CASE PARCELA.ID_SITUACAOPAGARFORNECEDORPARCELA');
    SQL.Add('    WHEN 0 THEN ''N''');
    SQL.Add('    ELSE ''S''');
    SQL.Add('  END AS FLG_QUITADO,');
    SQL.Add('');
    SQL.Add('  '''' AS DTA_QUITADA,');
    SQL.Add('  998 AS COD_CATEGORIA,');
    SQL.Add('  998 AS COD_SUBCATEGORIA,');
    SQL.Add('  PARCELA.NUMEROPARCELA AS NUM_PARCELA,');
    SQL.Add('  QTD_PARCELA.QTD_PARCELA AS QTD_PARCELA,');
    SQL.Add('  1 AS COD_LOJA,');
    SQL.Add('  '''' AS NUM_CGC,');
    SQL.Add('  0 AS NUM_BORDERO,');
    SQL.Add('  COALESCE(NOTAENTRADA.NUMERONOTA, 0) AS NUM_NF,');
    SQL.Add('  COALESCE(NOTAENTRADA.SERIE, '''') AS NUM_SERIE_NF,');
    SQL.Add('  DOCUMENTO.VALOR AS VAL_TOTAL_NF,');
    SQL.Add('  PARCELA.OBSERVACAO AS DES_OBSERVACAO,');
    SQL.Add('  0 AS NUM_PDV,');
    SQL.Add('  0 AS NUM_CUPOM_FISCAL,');
    SQL.Add('  0 AS COD_MOTIVO,');
    SQL.Add('  0 AS COD_CONVENIO,');
    SQL.Add('  0 AS COD_BIN,');
    SQL.Add('  '''' AS DES_BANDEIRA,');
    SQL.Add('  '''' AS DES_REDE_TEF,');
    SQL.Add('  0 AS VAL_RETENCAO,');
    SQL.Add('  0 AS COD_CONDICAO,');
    SQL.Add('  PARCELA.DATAPAGAMENTO AS DTA_PAGTO,');
    SQL.Add('  DOCUMENTO.DATAENTRADA AS DTA_ENTRADA,');
    SQL.Add('  '''' AS NUM_NOSSO_NUMERO,');
    SQL.Add('  '''' AS COD_BARRA,');
    SQL.Add('  ''F'' AS FLG_BOLETO_EMIT,');
    SQL.Add('  '''' AS NUM_CGC_CPF_TITULAR,');
    SQL.Add('  '''' AS DES_TITULAR,');
    SQL.Add('  30 AS NUM_CONDICAO,');
    SQL.Add('  0 AS VAL_CREDITO,');
    SQL.Add('  0 AS COD_BANCO_PGTO,');
    SQL.Add('  '''' AS DES_CC,');
    SQL.Add('  0 AS COD_BANDEIRA,');
    SQL.Add('  '''' AS DTA_PRORROGACAO,');
    SQL.Add('  1 AS NUM_SEQ_FIN,');
    SQL.Add('  0 AS COD_COBRANCA,');
    SQL.Add('  '''' AS DTA_COBRANCA,');
    SQL.Add('  ''N'' AS FLG_ACEITE,');
    SQL.Add('  0 AS TIPO_ACEITE');
    SQL.Add('FROM PUBLIC.PAGARFORNECEDORPARCELA AS PARCELA');
    SQL.Add('LEFT JOIN PUBLIC.PAGARFORNECEDOR AS DOCUMENTO');
    SQL.Add('ON PARCELA.ID_PAGARFORNECEDOR = DOCUMENTO.ID');
    SQL.Add('LEFT JOIN (');
    SQL.Add('  SELECT');
    SQL.Add('    ID_PAGARFORNECEDOR,');
    SQL.Add('    COUNT(*) AS QTD_PARCELA');
    SQL.Add('  FROM PUBLIC.PAGARFORNECEDORPARCELA');
    SQL.Add('  GROUP BY ID_PAGARFORNECEDOR');
    SQL.Add(') AS QTD_PARCELA');
    SQL.Add('ON PARCELA.ID_PAGARFORNECEDOR = QTD_PARCELA.ID_PAGARFORNECEDOR');
    SQL.Add('LEFT JOIN PUBLIC.NOTAENTRADA AS NOTAENTRADA');
    SQL.Add('ON DOCUMENTO.ID_NOTAENTRADA = NOTAENTRADA.ID');
    SQL.Add('WHERE ID_SITUACAOPAGARFORNECEDORPARCELA = 0');


    Open;

    First;
    NumLinha := 0;

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);

      Layout.SetValues(QryPrincipal2, NumLinha, RecordCount);

      Layout.FieldByName('DTA_ENTRADA').AsString:= FormatDateTime('dd/mm/yyyy',QryPrincipal2.FieldByName('DTA_ENTRADA').AsDateTime);
      Layout.FieldByName('DTA_EMISSAO').AsString:= FormatDateTime('dd/mm/yyyy',QryPrincipal2.FieldByName('DTA_EMISSAO').AsDateTime);
      Layout.FieldByName('DTA_VENCIMENTO').AsString:= FormatDateTime('dd/mm/yyyy',QryPrincipal2.FieldByName('DTA_VENCIMENTO').AsDateTime);
      Layout.FieldByName('DTA_QUITADA').AsString := '';
        Layout.FieldByName('DTA_PAGTO').AsString := '';

//      if Aberto = '1' then
//      begin
//        Layout.FieldByName('DTA_QUITADA').AsString := '';
//        Layout.FieldByName('DTA_PAGTO').AsString := '';
//      end
//      else
//      begin
//        Layout.FieldByName('DTA_QUITADA').AsString:= FormatDateTime('dd/mm/yyyy',QryPrincipal2.FieldByName('DTA_QUITADA').AsDateTime);
//        Layout.FieldByName('DTA_PAGTO').AsString:= FormatDateTime('dd/mm/yyyy',QryPrincipal2.FieldByName('DTA_PAGTO').AsDateTime);
//      end;

      Layout.FieldByName('DES_OBSERVACAO').AsString := StrReplace(StrLBReplace(FieldByName('DES_OBSERVACAO').AsString), '\n', '');

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;
end;

procedure TFrmVRAlbano.GerarFinanceiroReceber(Aberto: String);
var
   codParceiro : Integer;
   numDocto : String;
   idTitulo, count : integer;
begin
  inherited;
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    if Aberto = '1' then
    begin
      SQL.Add('SELECT');
      SQL.Add('  0 AS TIPO_PARCEIRO,');
      SQL.Add('  TITULO.ID_CLIENTEPREFERENCIAL AS COD_PARCEIRO,');
      SQL.Add('  1 AS TIPO_CONTA,');
      SQL.Add('  4 AS COD_ENTIDADE,');
      SQL.Add('CONCAT(''P'', TITULO.NUMEROCUPOM) AS NUM_DOCTO,');
      SQL.Add('  999 AS COD_BANCO,');
      SQL.Add('  '''' AS DES_BANCO,');
      SQL.Add('  TITULO.DATAEMISSAO AS DTA_EMISSAO,');
      SQL.Add('  TITULO.DATAVENCIMENTO AS DTA_VENCIMENTO,');
      SQL.Add('  TITULO.VALOR,');
      SQL.Add('  (TITULO.VALOR - COALESCE(SOMA.RESTANTE, 0)) AS VAL_PARCELA,');
      SQL.Add('  0 AS VAL_JUROS,');
      SQL.Add('  0 AS VAL_DESCONTO,');
      SQL.Add('  CASE TITULO.ID_SITUACAORECEBERCREDITOROTATIVO');
      SQL.Add('    WHEN 1 THEN ''S''');
      SQL.Add('    ELSE ''N''');
      SQL.Add('  END AS FLG_QUITADO,');
      SQL.Add('  '''' AS DTA_QUITADA,');
      SQL.Add('  ''997'' AS COD_CATEGORIA,');
      SQL.Add('  ''997'' AS COD_SUBCATEGORIA,');
      SQL.Add('  1 AS NUM_PARCELA,');
      SQL.Add('  1 AS QTD_PARCELA,');
      SQL.Add('  1 AS COD_LOJA,');
      SQL.Add('  '''' AS NUM_CGC,');
      SQL.Add('  0 AS NUM_BORDERO,');
      SQL.Add('  0 AS NUM_NF,');
      SQL.Add('  '''' AS NUM_SERIE_NF,');
      SQL.Add('  (TITULO.VALOR - COALESCE(SOMA.RESTANTE, 0)) AS VAL_TOTAL_NF,');
      SQL.Add('  TITULO.OBSERVACAO AS DES_OBSERVACAO,');
      SQL.Add('  TITULO.ECF AS NUM_PDV,');
      SQL.Add('  TITULO.NUMEROCUPOM AS NUM_CUPOM_FISCAL,');
      SQL.Add('  0 AS COD_MOTIVO,');
      SQL.Add('  0 AS COD_CONVENIO,');
      SQL.Add('  0 AS COD_BIN,');
      SQL.Add('  '''' AS DES_BANDEIRA,');
      SQL.Add('  '''' AS DES_REDE_TEF,');
      SQL.Add('  0 AS VAL_RETENCAO,');
      SQL.Add('  0 AS COD_CONDICAO,');
      SQL.Add('  NULL AS DTA_PAGTO,');
      SQL.Add('  TITULO.DATAEMISSAO AS DTA_ENTRADA,');
      SQL.Add('  '''' AS NUM_NOSSO_NUMERO,');
      SQL.Add('  '''' AS COD_BARRA,');
      SQL.Add('  ''N'' AS FLG_BOLETO_EMIT,');
      SQL.Add('  '''' AS NUM_CGC_CPF_TITULAR,');
      SQL.Add('  '''' AS DES_TITULAR,');
      SQL.Add('  30 AS NUM_CONDICAO,');
      SQL.Add('  0 AS VAL_CREDITO,');
      SQL.Add('  999 AS COD_BANCO_PGTO,');
      SQL.Add('  ''RECEBTO'' AS DES_CC,');
      SQL.Add('  0 AS COD_BANDEIRA,');
      SQL.Add('  '''' AS DTA_PRORROGACAO,');
      SQL.Add('  1 AS NUM_SEQ_FIN,');
      SQL.Add('  0 AS COD_COBRANCA,');
      SQL.Add('  '''' AS DTA_COBRANCA,');
      SQL.Add('  ''N'' AS FLG_ACEITE,');
      SQL.Add('  0 AS TIPO_ACEITE,');
      SQL.Add('  TITULO.ID AS ID');
      SQL.Add('FROM PUBLIC.RECEBERCREDITOROTATIVO AS TITULO');
      SQL.Add('LEFT JOIN (');
      SQL.Add('	SELECT');
      SQL.Add('		ID_RECEBERCREDITOROTATIVO,');
      SQL.Add('		SUM(VALOR) AS RESTANTE');
      SQL.Add('	FROM PUBLIC.RECEBERCREDITOROTATIVOITEM');
      SQL.Add('	GROUP BY ID_RECEBERCREDITOROTATIVO');
      SQL.Add(') AS SOMA');
      SQL.Add('ON TITULO.ID = SOMA.ID_RECEBERCREDITOROTATIVO');
      SQL.Add('WHERE TITULO.ID_SITUACAORECEBERCREDITOROTATIVO = 0');
      SQL.Add('');
      SQL.Add('UNION ALL');
      SQL.Add('');
      SQL.Add('SELECT');
      SQL.Add('  0 AS TIPO_PARCEIRO,');
      SQL.Add('  3700 + ID_CLIENTEEVENTUAL AS COD_PARCEIRO,');
      SQL.Add('  1 AS TIPO_CONTA,');
      SQL.Add('  11 AS COD_ENTIDADE,');
      SQL.Add('  '''' AS NUM_DOCTO,');
      SQL.Add('  999 AS COD_BANCO,');
      SQL.Add('  '''' AS DES_BANCO,');
      SQL.Add('  TITULO.DATAEMISSAO AS DTA_EMISSAO,');
      SQL.Add('  TITULO.DATAVENCIMENTO AS DTA_VENCIMENTO,');
      SQL.Add('  TITULO.VALOR,');
      SQL.Add('  (TITULO.VALOR - COALESCE(SOMA.RESTANTE, 0)) AS VAL_PARCELA,');
      SQL.Add('  0 AS VAL_JUROS,');
      SQL.Add('  0 AS VAL_DESCONTO,');
      SQL.Add('  CASE TITULO.ID_SITUACAORECEBERVENDAPRAZO');
      SQL.Add('    WHEN 1 THEN ''S''');
      SQL.Add('    ELSE ''N''');
      SQL.Add('  END AS FLG_QUITADO,');
      SQL.Add('  '''' AS DTA_QUITADA,');
      SQL.Add('  ''997'' AS COD_CATEGORIA,');
      SQL.Add('  ''997'' AS COD_SUBCATEGORIA,');
      SQL.Add('  1 AS NUM_PARCELA,');
      SQL.Add('  1 AS QTD_PARCELA,');
      SQL.Add('  1 AS COD_LOJA,');
      SQL.Add('  '''' AS NUM_CGC,');
      SQL.Add('  0 AS NUM_BORDERO,');
      SQL.Add('  TITULO.NUMERONOTA AS NUM_NF,');
      SQL.Add('  '''' AS NUM_SERIE_NF,');
      SQL.Add('  (TITULO.VALOR - COALESCE(SOMA.RESTANTE, 0)) AS VAL_TOTAL_NF,');
      SQL.Add('  TITULO.OBSERVACAO AS DES_OBSERVACAO,');
      SQL.Add('  1 AS NUM_PDV,');
      SQL.Add('  0 AS NUM_CUPOM_FISCAL,');
      SQL.Add('  0 AS COD_MOTIVO,');
      SQL.Add('  0 AS COD_CONVENIO,');
      SQL.Add('  0 AS COD_BIN,');
      SQL.Add('  '''' AS DES_BANDEIRA,');
      SQL.Add('  '''' AS DES_REDE_TEF,');
      SQL.Add('  0 AS VAL_RETENCAO,');
      SQL.Add('  0 AS COD_CONDICAO,');
      SQL.Add('  NULL AS DTA_PAGTO,');
      SQL.Add('  TITULO.DATAEMISSAO AS DTA_ENTRADA,');
      SQL.Add('  '''' AS NUM_NOSSO_NUMERO,');
      SQL.Add('  '''' AS COD_BARRA,');
      SQL.Add('  ''N'' AS FLG_BOLETO_EMIT,');
      SQL.Add('  '''' AS NUM_CGC_CPF_TITULAR,');
      SQL.Add('  '''' AS DES_TITULAR,');
      SQL.Add('  30 AS NUM_CONDICAO,');
      SQL.Add('  0 AS VAL_CREDITO,');
      SQL.Add('  999 AS COD_BANCO_PGTO,');
      SQL.Add('  ''RECEBTO'' AS DES_CC,');
      SQL.Add('  0 AS COD_BANDEIRA,');
      SQL.Add('  '''' AS DTA_PRORROGACAO,');
      SQL.Add('  1 AS NUM_SEQ_FIN,');
      SQL.Add('  0 AS COD_COBRANCA,');
      SQL.Add('  '''' AS DTA_COBRANCA,');
      SQL.Add('  ''N'' AS FLG_ACEITE,');
      SQL.Add('  0 AS TIPO_ACEITE,');
      SQL.Add('  TITULO.ID AS ID');
      SQL.Add('FROM PUBLIC.RECEBERVENDAPRAZO AS TITULO');
      SQL.Add('LEFT JOIN (');
      SQL.Add('	SELECT');
      SQL.Add('		ID_RECEBERVENDAPRAZO,');
      SQL.Add('		SUM(VALOR) AS RESTANTE');
      SQL.Add('	FROM PUBLIC.RECEBERVENDAPRAZOITEM');
      SQL.Add('	GROUP BY ID_RECEBERVENDAPRAZO');
      SQL.Add(') AS SOMA');
      SQL.Add('ON TITULO.ID = SOMA.ID_RECEBERVENDAPRAZO');
      SQL.Add('WHERE TITULO.ID_SITUACAORECEBERVENDAPRAZO = 0');

    end
    else begin
      SQL.Add('SELECT');
      SQL.Add('  0 AS TIPO_PARCEIRO,');
      SQL.Add('  TITULO.ID_CLIENTEPREFERENCIAL AS COD_PARCEIRO,');
      SQL.Add('  1 AS TIPO_CONTA,');
      SQL.Add('');
      SQL.Add('  CASE ITEM.ID_TIPORECEBIMENTO ');
      SQL.Add('    WHEN 1 THEN 4');
      SQL.Add('    WHEN 5 THEN 8');
      SQL.Add('    WHEN 2 THEN 1');
      SQL.Add('    WHEN 0 THEN 10');
      SQL.Add('    WHEN 6 THEN 6');
      SQL.Add('    WHEN 3 THEN 2');
      SQL.Add('    WHEN 8 THEN 4');
      SQL.Add('  END AS COD_ENTIDADE,');
      SQL.Add('');
      SQL.Add('  TITULO.NUMEROCUPOM AS NUM_DOCTO,');
      SQL.Add('  999 AS COD_BANCO,');
      SQL.Add('  '''' AS DES_BANCO,');
      SQL.Add('  TITULO.DATAEMISSAO AS DTA_EMISSAO,');
      SQL.Add('  TITULO.DATAVENCIMENTO AS DTA_VENCIMENTO,');
      SQL.Add('  ITEM.VALOR AS VAL_PARCELA,');
      SQL.Add('  ITEM.VALORMULTA AS VAL_JUROS,');
      SQL.Add('  ITEM.VALORDESCONTO AS VAL_DESCONTO,');
      SQL.Add('  ''S'' AS FLG_QUITADO,');
      SQL.Add('  ITEM.DATAPAGAMENTO AS DTA_QUITADA,');
      SQL.Add('  ''997'' AS COD_CATEGORIA,');
      SQL.Add('  ''997'' AS COD_SUBCATEGORIA,');
      SQL.Add('  1 AS NUM_PARCELA,');
      SQL.Add('  QTD_PARCELA.QTD_PARCELA AS QTD_PARCELA,');
      SQL.Add('  1 AS COD_LOJA,');
      SQL.Add('  '''' AS NUM_CGC,');
      SQL.Add('  0 AS NUM_BORDERO,');
      SQL.Add('  '''' AS NUM_NF,');
      SQL.Add('  '''' AS NUM_SERIE_NF,');
      SQL.Add('  TITULO.VALOR AS VAL_TOTAL_NF,');
      SQL.Add('  TITULO.OBSERVACAO AS DES_OBSERVACAO,');
      SQL.Add('  TITULO.ECF AS NUM_PDV,');
      SQL.Add('  TITULO.NUMEROCUPOM AS NUM_CUPOM_FISCAL,');
      SQL.Add('  0 AS COD_MOTIVO,');
      SQL.Add('  0 AS COD_CONVENIO,');
      SQL.Add('  0 AS COD_BIN,');
      SQL.Add('  '''' AS DES_BANDEIRA,');
      SQL.Add('  '''' AS DES_REDE_TEF,');
      SQL.Add('  0 AS VAL_RETENCAO,');
      SQL.Add('  0 AS COD_CONDICAO,');
      SQL.Add('  ITEM.DATAPAGAMENTO AS DTA_PAGTO,');
      SQL.Add('  TITULO.DATAEMISSAO AS DTA_ENTRADA,');
      SQL.Add('  '''' AS NUM_NOSSO_NUMERO,');
      SQL.Add('  '''' AS COD_BARRA,');
      SQL.Add('  ''N'' AS FLG_BOLETO_EMIT,');
      SQL.Add('  '''' AS NUM_CGC_CPF_TITULAR,');
      SQL.Add('  '''' AS DES_TITULAR,');
      SQL.Add('  30 AS NUM_CONDICAO,');
      SQL.Add('  0 AS VAL_CREDITO,');
      SQL.Add('  999 AS COD_BANCO_PGTO,');
      SQL.Add('  ''RECEBTO'' AS DES_CC,');
      SQL.Add('  0 AS COD_BANDEIRA,');
      SQL.Add('  '''' AS DTA_PRORROGACAO,');
      SQL.Add('  1 AS NUM_SEQ_FIN,');
      SQL.Add('  0 AS COD_COBRANCA,');
      SQL.Add('  '''' AS DTA_COBRANCA,');
      SQL.Add('  ''N'' AS FLG_ACEITE,');
      SQL.Add('  0 AS TIPO_ACEITE,');
      SQL.Add('  TITULO.ID AS ID');
      SQL.Add('FROM PUBLIC.RECEBERCREDITOROTATIVOITEM AS ITEM');
      SQL.Add('LEFT JOIN PUBLIC.RECEBERCREDITOROTATIVO AS TITULO');
      SQL.Add('ON ITEM.ID_RECEBERCREDITOROTATIVO = TITULO.ID');
      SQL.Add('LEFT JOIN (');
      SQL.Add('  SELECT');
      SQL.Add('    ID_RECEBERCREDITOROTATIVO,');
      SQL.Add('    COUNT(*) AS QTD_PARCELA');
      SQL.Add('  FROM PUBLIC.RECEBERCREDITOROTATIVOITEM');
      SQL.Add('  GROUP BY ID_RECEBERCREDITOROTATIVO');
      SQL.Add(') AS QTD_PARCELA');
      SQL.Add('ON TITULO.ID = QTD_PARCELA.ID_RECEBERCREDITOROTATIVO');
      SQL.Add('WHERE TITULO.ID_SITUACAORECEBERCREDITOROTATIVO <> 2');
      SQL.Add('AND ITEM.DATAPAGAMENTO >= ''2020-01-01''');
  //    SQL.Add('AND TITULO.NUMEROCUPOM = 13285');
      SQL.Add('ORDER BY TITULO.ID, ITEM.ID');
    end;

    Open;

    First;
    NumLinha := 0;
    codParceiro := 0;
    numDocto := '';
    count := 1;

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);

      Layout.SetValues(QryPrincipal2, NumLinha, RecordCount);

//      if( (codParceiro = QryPrincipal2.FieldByName('COD_PARCEIRO').AsInteger) and (numDocto = QryPrincipal2.FieldByName('NUM_DOCTO').AsString) ) then
      if( (idTitulo = QryPrincipal2.FieldByName('ID').AsInteger)) then
      begin
         inc(count);
         Layout.FieldByName('NUM_PARCELA').AsInteger := count;
      end
      else
      begin
         count := 1;
         idTitulo := QryPrincipal2.FieldByName('ID').AsInteger;

      end;

      Layout.FieldByName('DTA_ENTRADA').AsString:= FormatDateTime('dd/mm/yyyy',QryPrincipal2.FieldByName('DTA_ENTRADA').AsDateTime);
      Layout.FieldByName('DTA_EMISSAO').AsString:= FormatDateTime('dd/mm/yyyy',QryPrincipal2.FieldByName('DTA_EMISSAO').AsDateTime);
      Layout.FieldByName('DTA_VENCIMENTO').AsString:= FormatDateTime('dd/mm/yyyy',QryPrincipal2.FieldByName('DTA_VENCIMENTO').AsDateTime);
//      Layout.FieldByName('DTA_QUITADA').AsString:= FormatDateTime('dd/mm/yyyy',QryPrincipal2.FieldByName('DTA_QUITADA').AsDateTime);
//      Layout.FieldByName('DTA_PAGTO').AsString:= FormatDateTime('dd/mm/yyyy',QryPrincipal2.FieldByName('DTA_PAGTO').AsDateTime);

      if Aberto = '1' then
      begin
        Layout.FieldByName('DTA_QUITADA').AsString := '';
        Layout.FieldByName('DTA_PAGTO').AsString := '';
      end
      else
      begin
        Layout.FieldByName('DTA_QUITADA').AsString:= FormatDateTime('dd/mm/yyyy',QryPrincipal2.FieldByName('DTA_QUITADA').AsDateTime);
        Layout.FieldByName('DTA_PAGTO').AsString:= FormatDateTime('dd/mm/yyyy',QryPrincipal2.FieldByName('DTA_PAGTO').AsDateTime);
      end;

//      Layout.FieldByName('DTA_COBRANCA').AsDateTime:= QryPrincipal2.FieldByName('DTA_COBRANCA').AsDateTime;

      Layout.FieldByName('DES_OBSERVACAO').AsString := StrReplace(StrLBReplace(FieldByName('DES_OBSERVACAO').AsString), '\n', '');

      Layout.FieldByName('COD_BARRA').AsString := StrRetNums(Layout.FieldByName('COD_BARRA').AsString);

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;
end;

procedure TFrmVRAlbano.GerarFinanceiroReceberCartao;
begin
  inherited;
  with QryPrincipal2 do
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

      Layout.SetValues(QryPrincipal2, NumLinha, RecordCount);

//      if( (codParceiro = QryPrincipal2.FieldByName('COD_PARCEIRO').AsInteger) and (numDocto = QryPrincipal2.FieldByName('NUM_DOCTO').AsString) ) then
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
//         numDocto := QryPrincipal2.FieldByName('NUM_DOCTO').AsString;
//         codParceiro := QryPrincipal2.FieldByName('COD_PARCEIRO').AsInteger;
//      end;

      Layout.FieldByName('DTA_ENTRADA').AsString:= FormatDateTime('dd/mm/yyyy',QryPrincipal2.FieldByName('DTA_ENTRADA').AsDateTime);
      Layout.FieldByName('DTA_EMISSAO').AsString:= FormatDateTime('dd/mm/yyyy',QryPrincipal2.FieldByName('DTA_EMISSAO').AsDateTime);
      Layout.FieldByName('DTA_VENCIMENTO').AsString:= FormatDateTime('dd/mm/yyyy',QryPrincipal2.FieldByName('DTA_VENCIMENTO').AsDateTime);

//      if Aberto = '1' then
//      begin
//        Layout.FieldByName('DTA_QUITADA').AsString := '';
//        Layout.FieldByName('DTA_PAGTO').AsString := '';
//      end
//      else
//      begin
        Layout.FieldByName('DTA_QUITADA').AsString:= FormatDateTime('dd/mm/yyyy',QryPrincipal2.FieldByName('DTA_QUITADA').AsDateTime);
        Layout.FieldByName('DTA_PAGTO').AsString:= FormatDateTime('dd/mm/yyyy',QryPrincipal2.FieldByName('DTA_PAGTO').AsDateTime);
//      end;

      Layout.FieldByName('DTA_COBRANCA').AsDateTime:= QryPrincipal2.FieldByName('DTA_COBRANCA').AsDateTime;

      Layout.FieldByName('DES_OBSERVACAO').AsString := StrReplace(StrLBReplace(FieldByName('DES_OBSERVACAO').AsString), '\n', '');

      Layout.FieldByName('COD_BARRA').AsString := StrRetNums(Layout.FieldByName('COD_BARRA').AsString);

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;
end;

procedure TFrmVRAlbano.GerarFornecedor;
var
   observacao, email : string;
begin
  inherited;

  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
    SQL.Add('  FORNECEDOR.ID AS COD_FORNECEDOR,');
    SQL.Add('  FORNECEDOR.RAZAOSOCIAL AS DES_FORNECEDOR,');
    SQL.Add('  FORNECEDOR.NOMEFANTASIA AS DES_FANTASIA,');
    SQL.Add('');
    SQL.Add('  CASE FORNECEDOR.ID_TIPOINSCRICAO');
    SQL.Add('    WHEN 0 THEN LPAD(CAST(FORNECEDOR.CNPJ AS VARCHAR(14)), 14, ''0'')');
    SQL.Add('    ELSE LPAD(CAST(FORNECEDOR.CNPJ AS VARCHAR(11)), 11, ''0'')');
    SQL.Add('  END AS NUM_CGC,');
    SQL.Add('');
    SQL.Add('  CASE FORNECEDOR.ID_TIPOINSCRICAO');
    SQL.Add('    WHEN 1 THEN ''ISENTO''');
    SQL.Add('    ELSE FORNECEDOR.INSCRICAOESTADUAL');
    SQL.Add('  END AS NUM_INSC_EST,');
    SQL.Add('');
    SQL.Add('  FORNECEDOR.ENDERECO AS DES_ENDERECO,');
    SQL.Add('  FORNECEDOR.BAIRRO AS DES_BAIRRO,');
    SQL.Add('  MUNICIPIO.DESCRICAO AS DES_CIDADE,');
    SQL.Add('  ESTADO.SIGLA AS DES_SIGLA,');
    SQL.Add('');
    SQL.Add('  CASE');
    SQL.Add('    WHEN LENGTH(CAST(FORNECEDOR.CEP AS VARCHAR(8))) < 8 THEN ''23017390''');
    SQL.Add('    ELSE FORNECEDOR.CEP');
    SQL.Add('  END AS NUM_CEP,');
    SQL.Add('');
    SQL.Add('  FORNECEDOR.TELEFONE AS NUM_FONE,');
    SQL.Add('  '''' AS NUM_FAX, --');
    SQL.Add('  CONTATO.NOME AS DES_CONTATO, --');
    SQL.Add('  0 AS QTD_DIA_CARENCIA, --');
    SQL.Add('  0 AS NUM_FREQ_VISITA, --');
    SQL.Add('  0 AS VAL_DESCONTO, --');
    SQL.Add('  0 AS NUM_PRAZO, --');
    SQL.Add('  ''N'' AS ACEITA_DEVOL_MER, --');
    SQL.Add('  ''N'' AS CAL_IPI_VAL_BRUTO, --');
    SQL.Add('  ''N'' AS CAL_ICMS_ENC_FIN, --');
    SQL.Add('  ''N'' AS CAL_ICMS_VAL_IPI, --');
    SQL.Add('  ''N'' AS MICRO_EMPRESA, --');
    SQL.Add('  0 AS COD_FORNECEDOR_ANT,');
    SQL.Add('  FORNECEDOR.NUMERO AS NUM_ENDERECO,');
    SQL.Add('');
    SQL.Add('  CASE FORNECEDOR.ID_TIPOINSCRICAO');
    SQL.Add('    WHEN 1 THEN ''INSCRICAO: '' || FORNECEDOR.INSCRICAOESTADUAL || '' COMPLEMENTO: '' || FORNECEDOR.COMPLEMENTO || '' '' ||FORNECEDOR.OBSERVACAO');
    SQL.Add('    ELSE ''COMPLEMENTO: '' || FORNECEDOR.COMPLEMENTO || '' '' ||FORNECEDOR.OBSERVACAO ');
    SQL.Add('  END AS DES_OBSERVACAO,');
    SQL.Add('');
    SQL.Add('  '''' AS DES_EMAIL, --');
    SQL.Add('  '''' AS DES_WEB_SITE, --');
    SQL.Add('  ''N'' AS FABRICANTE, --');
    SQL.Add('  ''N'' AS FLG_PRODUTOR_RURAL, --');
    SQL.Add('  0 AS TIPO_FRETE, --');
    SQL.Add('  ''N'' AS FLG_SIMPLES, --');
    SQL.Add('  ''N'' AS FLG_SUBSTITUTO_TRIB, --');
    SQL.Add('  0 AS COD_CONTACCFORN, --');
    SQL.Add('');
    SQL.Add('  CASE FORNECEDOR.ID_SITUACAOCADASTRO');
    SQL.Add('    WHEN 0 THEN ''S''');
    SQL.Add('    ELSE ''N''');
    SQL.Add('  END AS INATIVO, --');
    SQL.Add('');
    SQL.Add('  0 AS COD_CLASSIF, --');
    SQL.Add('  FORNECEDOR.DATACADASTRO AS DTA_CADASTRO,');
    SQL.Add('  0 AS VAL_CREDITO, --');
    SQL.Add('  0 AS VAL_DEBITO, --');
    SQL.Add('  1 AS PED_MIN_VAL, --');
    SQL.Add('  COALESCE(CONTATO.EMAIL, '''') AS DES_EMAIL_VEND, --');
    SQL.Add('  '''' AS SENHA_COTACAO, --');
    SQL.Add('  -1 AS TIPO_PRODUTOR, --');
    SQL.Add('  '''' AS NUM_CELULAR --');
    SQL.Add('FROM PUBLIC.FORNECEDOR AS FORNECEDOR');
    SQL.Add('LEFT JOIN PUBLIC.MUNICIPIO AS MUNICIPIO');
    SQL.Add('ON FORNECEDOR.ID_MUNICIPIO = MUNICIPIO.ID ');
    SQL.Add('LEFT JOIN PUBLIC.ESTADO AS ESTADO');
    SQL.Add('ON FORNECEDOR.ID_ESTADO = ESTADO.ID');
    SQL.Add('LEFT JOIN (');
    SQL.Add('  SELECT');
    SQL.Add('    ID_FORNECEDOR,');
    SQL.Add('    EMAIL,');
    SQL.Add('    NOME');
    SQL.Add('  FROM PUBLIC.FORNECEDORCONTATO');
    SQL.Add('  WHERE ID_FORNECEDOR IN (');
    SQL.Add('    SELECT');
    SQL.Add('      ID_FORNECEDOR');
    SQL.Add('    FROM PUBLIC.FORNECEDORCONTATO');
    SQL.Add('    GROUP BY ID_FORNECEDOR');
    SQL.Add('    HAVING COUNT(*) = 1');
    SQL.Add('  )');
    SQL.Add(') AS CONTATO');
    SQL.Add('ON FORNECEDOR.ID = CONTATO.ID_FORNECEDOR');
//    SQL.Add('WHERE FORNECEDOR.ID = 54');

    Open;

    First;
    NumLinha := 0;

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);

      Layout.SetValues(QryPrincipal2, NumLinha, RecordCount);

      Layout.FieldByName('NUM_CGC').AsString := StrRetNums(Layout.FieldByName('NUM_CGC').AsString);
      Layout.FieldByName('NUM_INSC_EST').AsString := StrRetNums(Layout.FieldByName('NUM_INSC_EST').AsString);
      Layout.FieldByName('NUM_CEP').AsString := StrRetNums(Layout.FieldByName('NUM_CEP').AsString);

//      if Length(Layout.FieldByName('NUM_CGC').AsString) > 11 then
//      begin
//        if not ValidaCGC(Layout.FieldByName('NUM_CGC').AsString) then
//          Layout.FieldByName('NUM_CGC').AsString := '';
//      end
//      else
//        if not ValidaCPF(Layout.FieldByName('NUM_CGC').AsString) then
//          Layout.FieldByName('NUM_CGC').AsString := '';

      if Layout.FieldByName('NUM_INSC_EST').AsString = '' then
        Layout.FieldByName('NUM_INSC_EST').AsString := 'ISENTO';

      if Layout.FieldByName('NUM_ENDERECO').AsString = '' then
        Layout.FieldByName('NUM_ENDERECO').AsString := '0';

      Layout.FieldByName('NUM_FONE').AsString := StrRetNums( FieldByName('NUM_FONE').AsString );


      if Layout.FieldByName('FLG_PRODUTOR_RURAL').AsString = 'S' then
      begin
        if StrRetNums(Layout.FieldByName('NUM_INSC_EST').AsString) = '' then
            Layout.FieldByName('TIPO_PRODUTOR').AsInteger := 0
        else
            Layout.FieldByName('TIPO_PRODUTOR').AsInteger := 1
      end;

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
    end;
    Next;
    end;
    Close;
  end;
end;

procedure TFrmVRAlbano.GerarGrupo;
begin
  inherited;
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
    SQL.Add('  MERCADOLOGICO1 AS COD_SECAO,');
    SQL.Add('  MERCADOLOGICO2 AS COD_GRUPO,');
    SQL.Add('  DESCRICAO AS DES_GRUPO,');
    SQL.Add('  0 AS VAL_META');
    SQL.Add('FROM PUBLIC.MERCADOLOGICO');
    SQL.Add('WHERE NIVEL = 2');

    Open;

    First;

    NumLinha := 0;

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);

      Layout.SetValues(QryPrincipal2, NumLinha, RecordCount);

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;

end;

procedure TFrmVRAlbano.GerarInfoNutricionais;
begin
  inherited;

  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
    SQL.Add('  NUTRICIONAL.ID AS COD_INFO_NUTRICIONAL,');
    SQL.Add('  NUTRICIONAL.DESCRICAO AS DES_INFO_NUTRICIONAL,');
    SQL.Add('  NUTRICIONAL.QUANTIDADE AS PORCAO,');
    SQL.Add('  NUTRICIONAL.CALORIA AS VALOR_CALORICO,');
    SQL.Add('  NUTRICIONAL.CARBOIDRATO AS CARBOIDRATO,');
    SQL.Add('  NUTRICIONAL.PROTEINA AS PROTEINA,');
    SQL.Add('  NUTRICIONAL.GORDURA AS GORDURA_TOTAL,');
    SQL.Add('  NUTRICIONAL.GORDURASATURADA AS GORDURA_SATURADA,');
    SQL.Add('  0 AS COLESTEROL,');
    SQL.Add('  NUTRICIONAL.FIBRA AS FIBRA_ALIMENTAR,');
    SQL.Add('  NUTRICIONAL.CALCIO AS CALCIO,');
    SQL.Add('  NUTRICIONAL.FERRO AS FERRO,');
    SQL.Add('  NUTRICIONAL.SODIO AS SODIO,');
    SQL.Add('  NUTRICIONAL.PERCENTUALCALORIA AS VD_VALOR_CALORICO,');
    SQL.Add('  NUTRICIONAL.PERCENTUALCARBOIDRATO AS VD_CARBOIDRATO,');
    SQL.Add('  NUTRICIONAL.PERCENTUALPROTEINA AS VD_PROTEINA,');
    SQL.Add('  NUTRICIONAL.PERCENTUALGORDURA AS VD_GORDURA_TOTAL,');
    SQL.Add('  NUTRICIONAL.PERCENTUALGORDURASATURADA AS VD_GORDURA_SATURADA,');
    SQL.Add('  0 AS VD_COLESTEROL,');
    SQL.Add('  NUTRICIONAL.PERCENTUALFIBRA AS VD_FIBRA_ALIMENTAR,');
    SQL.Add('  NUTRICIONAL.PERCENTUALCALCIO AS VD_CALCIO,');
    SQL.Add('  NUTRICIONAL.PERCENTUALFERRO AS VD_FERRO,');
    SQL.Add('  NUTRICIONAL.PERCENTUALSODIO AS VD_SODIO,');
    SQL.Add('  NUTRICIONAL.GORDURATRANS AS GORDURA_TRANS,');
    SQL.Add('  0 AS VD_GORDURA_TRANS,');
    SQL.Add('');
    SQL.Add('  CASE NUTRICIONAL.ID_TIPOUNIDADEPORCAO');
    SQL.Add('    WHEN 0 THEN ''G''');
    SQL.Add('    WHEN 1 THEN ''ML''');
    SQL.Add('    WHEN 2 THEN ''UN''');
    SQL.Add('  END AS UNIDADE_PORCAO,');
    SQL.Add('');
    SQL.Add('  CASE NUTRICIONAL.ID_TIPOUNIDADEPORCAO');
    SQL.Add('    WHEN 0 THEN ''GRAMA''');
    SQL.Add('    WHEN 1 THEN ''MILILITRO''');
    SQL.Add('    WHEN 2 THEN ''UNIDADE''');
    SQL.Add('  END AS DES_PORCAO,');
    SQL.Add('');
    SQL.Add('  NUTRICIONAL.MEDIDAINTEIRA AS PARTE_INTEIRA_MED_CASEIRA,');
    SQL.Add('  TIPOMEDIDA.ID AS MED_CASEIRA_UTILIZADA');
    SQL.Add('FROM PUBLIC.NUTRICIONALTOLEDO AS NUTRICIONAL');
    SQL.Add('LEFT JOIN PUBLIC.TIPOMEDIDA AS TIPOMEDIDA');
    SQL.Add('ON NUTRICIONAL.ID_TIPOMEDIDA = TIPOMEDIDA.ID');
    SQL.Add('WHERE NUTRICIONAL.ID_SITUACAOCADASTRO = 1');

    Open;

    First;

    NumLinha := 0;

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);

      Layout.SetValues(QryPrincipal2, NumLinha, RecordCount);

//      Layout.FieldByName('COD_INFO_NUTRICIONAL').AsString := GerarPLU( Layout.FieldByName('COD_INFO_NUTRICIONAL').AsString );

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;

end;

procedure TFrmVRAlbano.GerarNCM;
var
 count : Integer;
begin
  inherited;
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT DISTINCT');
SQL.Add('  0 AS COD_NCM,');
SQL.Add('  NCM.DESCRICAO AS DES_NCM,');
SQL.Add('');
SQL.Add('  CONCAT( LPAD(CAST(PRODUTO.NCM1 AS VARCHAR(4)), 4, ''0''), LPAD(CAST(PRODUTO.NCM2 AS VARCHAR(2)), 2, ''0''), LPAD(CAST(PRODUTO.NCM3 AS VARCHAR(2)), 2, ''0'') ) AS NUM_NCM,');
SQL.Add('');
SQL.Add('  CASE PRODUTO.ID_TIPOPISCOFINS');
SQL.Add('    WHEN ''0'' THEN ''N''');
SQL.Add('    ELSE ''S''');
SQL.Add('  END AS FLG_NAO_PIS_COFINS,');
SQL.Add('');
SQL.Add('  CASE PRODUTO.ID_TIPOPISCOFINS');
SQL.Add('    WHEN ''0'' THEN -1');
SQL.Add('    WHEN ''2'' THEN 2');
SQL.Add('    WHEN ''3'' THEN 1');
SQL.Add('    WHEN ''4'' THEN 4');
SQL.Add('    WHEN ''7'' THEN 0');
SQL.Add('    ELSE ''0''');
SQL.Add('  END AS TIPO_NAO_PIS_COFINS,');
SQL.Add('');
SQL.Add('COALESCE(LPAD(CAST(PRODUTO.TIPONATUREZARECEITA AS VARCHAR(3)), 3, ''000''), ''999'') AS COD_TAB_SPED,');
SQL.Add('');
SQL.Add('  CASE');
SQL.Add('    WHEN CEST.ID IS NULL THEN ''9999999''');
SQL.Add('    WHEN CONCAT(LPAD(CAST(CEST1 AS VARCHAR(2)), 2, ''0''), LPAD(CAST(CEST2 AS VARCHAR(3)), 3,''0''), LPAD(CAST(CEST3 AS VARCHAR(2)), 2, ''0'')) = ''0000000'' THEN ''9999999''');
SQL.Add('    ELSE CONCAT(LPAD(CAST(CEST1 AS VARCHAR(2)), 2, ''0''), LPAD(CAST(CEST2 AS VARCHAR(3)), 3,''0''), LPAD(CAST(CEST3 AS VARCHAR(2)), 2, ''0''))');
SQL.Add('  END AS NUM_CEST,');
SQL.Add('');
SQL.Add('  ''RJ'' AS DES_SIGLA,');
SQL.Add('');
SQL.Add('  CASE PRODUTOALIQUOTA.ID_ALIQUOTACREDITO');
SQL.Add('    WHEN ''0'' THEN 2');
SQL.Add('    WHEN ''2'' THEN 4');
SQL.Add('    WHEN ''4'' THEN 8');
SQL.Add('    WHEN ''5'' THEN 6');
SQL.Add('    WHEN ''6'' THEN 1');
SQL.Add('    WHEN ''7'' THEN 25');
SQL.Add('    WHEN ''21'' THEN 39 ');
SQL.Add('    WHEN ''29'' THEN 39 ');
SQL.Add('    WHEN ''30'' THEN 40');
SQL.Add('    WHEN ''38'' THEN 3');
SQL.Add('    WHEN ''40'' THEN 41');
SQL.Add('    WHEN ''75'' THEN 1');
SQL.Add('    WHEN ''99'' THEN 1');
SQL.Add('  END AS COD_TRIB_ENTRADA,');
SQL.Add('');
SQL.Add('  CASE PRODUTOALIQUOTA.ID_ALIQUOTADEBITO');
SQL.Add('    WHEN ''0'' THEN 2');
SQL.Add('    WHEN ''2'' THEN 4');
SQL.Add('    WHEN ''4'' THEN 8');
SQL.Add('    WHEN ''5'' THEN 6');
SQL.Add('    WHEN ''6'' THEN 1');
SQL.Add('    WHEN ''7'' THEN 25');
SQL.Add('    WHEN ''21'' THEN 39 ');
SQL.Add('    WHEN ''29'' THEN 39 ');
SQL.Add('    WHEN ''30'' THEN 40');
SQL.Add('    WHEN ''38'' THEN 3');
SQL.Add('    WHEN ''40'' THEN 41');
SQL.Add('    WHEN ''75'' THEN 1');
SQL.Add('    WHEN ''99'' THEN 1');
SQL.Add('  END AS COD_TRIB_SAIDA,');
SQL.Add('');
SQL.Add('  0 AS PER_IVA,');
SQL.Add('  0 AS PER_FCP_ST,');
SQL.Add('  COALESCE(CODIGOBENEFICIOCST.CODIGO, '''') AS COD_BENEF_FISCAL');
SQL.Add('FROM PUBLIC.PRODUTO AS PRODUTO');
SQL.Add('');
SQL.Add('LEFT JOIN PUBLIC.PRODUTOALIQUOTA AS PRODUTOALIQUOTA');
SQL.Add('ON PRODUTO.ID = PRODUTOALIQUOTA.ID_PRODUTO');
SQL.Add('');
SQL.Add('LEFT JOIN PUBLIC.PRODUTOALIQUOTABENEFICIO AS PRODUTOALIQUOTABENEFICIO');
SQL.Add('ON PRODUTOALIQUOTA.ID = PRODUTOALIQUOTABENEFICIO.ID_PRODUTOALIQUOTA');
SQL.Add('');
SQL.Add('LEFT JOIN PUBLIC.CODIGOBENEFICIOCST AS CODIGOBENEFICIOCST');
SQL.Add('ON PRODUTOALIQUOTABENEFICIO.ID_BENEFICIO = CODIGOBENEFICIOCST.ID');
SQL.Add('');
SQL.Add('LEFT JOIN PUBLIC.NCM AS NCM');
SQL.Add('ON PRODUTO.NCM1 = NCM.NCM1');
SQL.Add('AND PRODUTO.NCM2 = NCM.NCM2');
SQL.Add('AND PRODUTO.NCM3 = NCM.NCM3');
SQL.Add('');
SQL.Add('LEFT JOIN PUBLIC.CEST AS CEST');
SQL.Add('ON PRODUTO.ID_CEST = CEST.ID');
SQL.Add('');
SQL.Add('ORDER BY ');
SQL.Add('DES_NCM,');
SQL.Add('NUM_NCM,');
SQL.Add('FLG_NAO_PIS_COFINS,');
SQL.Add('TIPO_NAO_PIS_COFINS,');
SQL.Add('COD_TAB_SPED,');
SQL.Add('NUM_CEST,');
SQL.Add('DES_SIGLA,');
SQL.Add('COD_TRIB_ENTRADA,');
SQL.Add('COD_TRIB_SAIDA,');
SQL.Add('PER_IVA,');
SQL.Add('PER_FCP_ST,');
SQL.Add('COD_BENEF_FISCAL');

    Open;
    First;

    count := 0;


    NumLinha := 0;

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);
      Inc(count);
      Layout.SetValues(QryPrincipal2, NumLinha, RecordCount);

      Layout.FieldByName('COD_NCM').AsInteger := count;

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;
end;

procedure TFrmVRAlbano.GerarNCMUF;
var
 count : Integer;
begin
  inherited;
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

   SQL.Add('SELECT DISTINCT');
SQL.Add('  0 AS COD_NCM,');
SQL.Add('  NCM.DESCRICAO AS DES_NCM,');
SQL.Add('');
SQL.Add('  CONCAT( LPAD(CAST(PRODUTO.NCM1 AS VARCHAR(4)), 4, ''0''), LPAD(CAST(PRODUTO.NCM2 AS VARCHAR(2)), 2, ''0''), LPAD(CAST(PRODUTO.NCM3 AS VARCHAR(2)), 2, ''0'') ) AS NUM_NCM,');
SQL.Add('');
SQL.Add('  CASE PRODUTO.ID_TIPOPISCOFINS');
SQL.Add('    WHEN ''0'' THEN ''N''');
SQL.Add('    ELSE ''S''');
SQL.Add('  END AS FLG_NAO_PIS_COFINS,');
SQL.Add('');
SQL.Add('  CASE PRODUTO.ID_TIPOPISCOFINS');
SQL.Add('    WHEN ''0'' THEN -1');
SQL.Add('    WHEN ''2'' THEN 2');
SQL.Add('    WHEN ''3'' THEN 1');
SQL.Add('    WHEN ''4'' THEN 4');
SQL.Add('    WHEN ''7'' THEN 0');
SQL.Add('    ELSE ''0''');
SQL.Add('  END AS TIPO_NAO_PIS_COFINS,');
SQL.Add('');
SQL.Add('COALESCE(LPAD(CAST(PRODUTO.TIPONATUREZARECEITA AS VARCHAR(3)), 3, ''000''), ''999'') AS COD_TAB_SPED,');
SQL.Add('');
SQL.Add('  CASE');
SQL.Add('    WHEN CEST.ID IS NULL THEN ''9999999''');
SQL.Add('    WHEN CONCAT(LPAD(CAST(CEST1 AS VARCHAR(2)), 2, ''0''), LPAD(CAST(CEST2 AS VARCHAR(3)), 3,''0''), LPAD(CAST(CEST3 AS VARCHAR(2)), 2, ''0'')) = ''0000000'' THEN ''9999999''');
SQL.Add('    ELSE CONCAT(LPAD(CAST(CEST1 AS VARCHAR(2)), 2, ''0''), LPAD(CAST(CEST2 AS VARCHAR(3)), 3,''0''), LPAD(CAST(CEST3 AS VARCHAR(2)), 2, ''0''))');
SQL.Add('  END AS NUM_CEST,');
SQL.Add('');
SQL.Add('  ''RJ'' AS DES_SIGLA,');
SQL.Add('');
SQL.Add('  CASE PRODUTOALIQUOTA.ID_ALIQUOTACREDITO');
SQL.Add('    WHEN ''0'' THEN 2');
SQL.Add('    WHEN ''2'' THEN 4');
SQL.Add('    WHEN ''4'' THEN 8');
SQL.Add('    WHEN ''5'' THEN 6');
SQL.Add('    WHEN ''6'' THEN 1');
SQL.Add('    WHEN ''7'' THEN 25');
SQL.Add('    WHEN ''21'' THEN 39 ');
SQL.Add('    WHEN ''29'' THEN 39 ');
SQL.Add('    WHEN ''30'' THEN 40');
SQL.Add('    WHEN ''38'' THEN 3');
SQL.Add('    WHEN ''40'' THEN 41');
SQL.Add('    WHEN ''75'' THEN 1');
SQL.Add('    WHEN ''99'' THEN 1');
SQL.Add('  END AS COD_TRIB_ENTRADA,');
SQL.Add('');
SQL.Add('  CASE PRODUTOALIQUOTA.ID_ALIQUOTADEBITO');
SQL.Add('    WHEN ''0'' THEN 2');
SQL.Add('    WHEN ''2'' THEN 4');
SQL.Add('    WHEN ''4'' THEN 8');
SQL.Add('    WHEN ''5'' THEN 6');
SQL.Add('    WHEN ''6'' THEN 1');
SQL.Add('    WHEN ''7'' THEN 25');
SQL.Add('    WHEN ''21'' THEN 39 ');
SQL.Add('    WHEN ''29'' THEN 39 ');
SQL.Add('    WHEN ''30'' THEN 40');
SQL.Add('    WHEN ''38'' THEN 3');
SQL.Add('    WHEN ''40'' THEN 41');
SQL.Add('    WHEN ''75'' THEN 1');
SQL.Add('    WHEN ''99'' THEN 1');
SQL.Add('  END AS COD_TRIB_SAIDA,');
SQL.Add('');
SQL.Add('  0 AS PER_IVA,');
SQL.Add('  0 AS PER_FCP_ST,');
SQL.Add('  COALESCE(CODIGOBENEFICIOCST.CODIGO, '''') AS COD_BENEF_FISCAL');
SQL.Add('FROM PUBLIC.PRODUTO AS PRODUTO');
SQL.Add('');
SQL.Add('LEFT JOIN PUBLIC.PRODUTOALIQUOTA AS PRODUTOALIQUOTA');
SQL.Add('ON PRODUTO.ID = PRODUTOALIQUOTA.ID_PRODUTO');
SQL.Add('');
SQL.Add('LEFT JOIN PUBLIC.PRODUTOALIQUOTABENEFICIO AS PRODUTOALIQUOTABENEFICIO');
SQL.Add('ON PRODUTOALIQUOTA.ID = PRODUTOALIQUOTABENEFICIO.ID_PRODUTOALIQUOTA');
SQL.Add('');
SQL.Add('LEFT JOIN PUBLIC.CODIGOBENEFICIOCST AS CODIGOBENEFICIOCST');
SQL.Add('ON PRODUTOALIQUOTABENEFICIO.ID_BENEFICIO = CODIGOBENEFICIOCST.ID');
SQL.Add('');
SQL.Add('LEFT JOIN PUBLIC.NCM AS NCM');
SQL.Add('ON PRODUTO.NCM1 = NCM.NCM1');
SQL.Add('AND PRODUTO.NCM2 = NCM.NCM2');
SQL.Add('AND PRODUTO.NCM3 = NCM.NCM3');
SQL.Add('');
SQL.Add('LEFT JOIN PUBLIC.CEST AS CEST');
SQL.Add('ON PRODUTO.ID_CEST = CEST.ID');
SQL.Add('');
SQL.Add('ORDER BY ');
SQL.Add('DES_NCM,');
SQL.Add('NUM_NCM,');
SQL.Add('FLG_NAO_PIS_COFINS,');
SQL.Add('TIPO_NAO_PIS_COFINS,');
SQL.Add('COD_TAB_SPED,');
SQL.Add('NUM_CEST,');
SQL.Add('DES_SIGLA,');
SQL.Add('COD_TRIB_ENTRADA,');
SQL.Add('COD_TRIB_SAIDA,');
SQL.Add('PER_IVA,');
SQL.Add('PER_FCP_ST,');
SQL.Add('COD_BENEF_FISCAL');


    Open;
    First;

    count := 0;


    NumLinha := 0;

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);
      Inc(count);
      Layout.SetValues(QryPrincipal2, NumLinha, RecordCount);

      Layout.FieldByName('COD_NCM').AsInteger := count;

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;
end;

procedure TFrmVRAlbano.GerarNFClientes;
begin
  inherited;
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
    SQL.Add('    SUBSTR(CAPA.COD_PART,2, 60) AS COD_CLIENTE,');
    SQL.Add('    CAPA.NUM_DOC AS NUM_NF_CLI,');
    SQL.Add('    CAPA.SER AS NUM_SERIE_NF,');
    SQL.Add('');
    SQL.Add('    CFOP.ID_CFOP AS CFOP,');
    SQL.Add('');
    SQL.Add('    CASE CAPA.ID_COI_A');
    SQL.Add('        WHEN 1303 THEN 0');
    SQL.Add('        WHEN 30 THEN 0');
    SQL.Add('        WHEN 201 THEN 2');
    SQL.Add('        WHEN 41 THEN 2');
    SQL.Add('    END AS TIPO_NF,');
    SQL.Add('');
    SQL.Add('    CASE ');
    SQL.Add('        WHEN CAPA.CHV_NFE IS NULL THEN ''NF''');
    SQL.Add('        ELSE ''NFE''');
    SQL.Add('    END AS DES_ESPECIE,');
    SQL.Add('');
    SQL.Add('    CAPA.VL_DOC AS VAL_TOTAL_NF,');
    SQL.Add('    CAPA.DT_DOC AS DTA_EMISSAO,');
    SQL.Add('    CAPA.DT_E_S AS DTA_ENTRADA,');
    SQL.Add('    CAPA.VL_IPI AS VAL_TOTAL_IPI,');
    SQL.Add('    CAPA.VL_FRT AS VAL_FRETE,');
    SQL.Add('    0 AS VAL_ENC_FINANC, --');
    SQL.Add('    0 AS VAL_DESC_FINANC, --');
    SQL.Add('    '''' AS NUM_CGC,');
    SQL.Add('    '''' AS DES_NATUREZA,');
    SQL.Add('    CAPA.OBSERVACAO_A AS DES_OBSERVACAO,');
    SQL.Add('    ''N'' AS FLG_CANCELADA,');
    SQL.Add('    CAPA.CHV_NFE AS NUM_CHAVE_ACESSO');
    SQL.Add('FROM');
    SQL.Add('    FIS_T_C100 CAPA');
    SQL.Add('LEFT JOIN ');
    SQL.Add('    BAS_T_COI CFOP ');
    SQL.Add('ON ');
    SQL.Add('    CAPA.ID_COI_A = CFOP.ID_COI     ');
    SQL.Add('WHERE ');
    SQL.Add('    CAPA.ID_COI_A IN (1303, 30)');
    SQL.Add('AND ');
    SQL.Add('    CAPA.COD_SIT IN (6, 0)');
    SQL.Add('AND');
    SQL.Add('    CAPA.NUM_DOC IS NOT NULL');
    SQL.Add('AND');
    SQL.Add('    CAPA.STATUS_A = 2  ');

    SQL.Add('AND    ');
    SQL.Add('    CAPA.DT_DOC >= :INI ');
    SQL.Add('AND    ');
    SQL.Add('    CAPA.DT_DOC <= :FIM ');
    SQL.Add('AND    ');

    Parameters.ParamByName('INI').Value := FormatDateTime('dd/mm/yyyy', DtpInicial.Date);
    Parameters.ParamByName('FIM').Value := FormatDateTime('dd/mm/yyyy', DtpFinal.Date);

    Open;

    First;
    NumLinha := 0;

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);

      Layout.SetValues(QryPrincipal2, NumLinha, RecordCount);

      Layout.FieldByName('DTA_EMISSAO').AsDateTime := QryPrincipal2.FieldByName('DTA_EMISSAO').AsDateTime;
      Layout.FieldByName('DTA_ENTRADA').AsDateTime := QryPrincipal2.FieldByName('DTA_ENTRADA').AsDateTime;

      Layout.FieldByName('DES_OBSERVACAO').AsString := StrLBReplace(FieldByName('DES_OBSERVACAO').AsString);

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;
end;

procedure TFrmVRAlbano.GerarNFFornec;
begin
  inherited;
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

 SQL.Add('SELECT');
SQL.Add('  ENTRADA.ID_FORNECEDOR AS COD_FORNECEDOR,');
SQL.Add('  ENTRADA.NUMERONOTA AS NUM_NF_FORN,');
SQL.Add('  ENTRADA.SERIE AS NUM_SERIE_NF,');
SQL.Add('  '''' AS NUM_SUBSERIE_NF,');
SQL.Add('  -- CFOP.CFOPNOTA AS CFOP,');
SQL.Add('  0 AS TIPO_NF,');
SQL.Add('  ''NFE'' AS DES_ESPECIE,');
SQL.Add('  ENTRADA.VALORTOTAL AS VAL_TOTAL_NF,');
SQL.Add('  ENTRADA.DATAEMISSAO AS DTA_EMISSAO,');
SQL.Add('  ENTRADA.DATAENTRADA AS DTA_ENTRADA,');
SQL.Add('  ENTRADA.VALORIPI AS VAL_TOTAL_IPI,');
SQL.Add('  ENTRADA.VALORMERCADORIA AS VAL_VENDA_VAREJO,');
SQL.Add('  ENTRADA.VALORFRETE AS VAL_FRETE,');
SQL.Add('  0 AS VAL_ACRESCIMO,');
SQL.Add('  ENTRADA.VALORDESCONTO AS VAL_DESCONTO,');
SQL.Add('  '''' AS NUM_CGC,');
SQL.Add('  ENTRADA.VALORBASECALCULO AS VAL_TOTAL_BC,');
SQL.Add('  ENTRADA.VALORICMS AS VAL_TOTAL_ICMS,');
SQL.Add('  ENTRADA.VALORBASESUBSTITUICAO AS VAL_BC_SUBST,');
SQL.Add('  ENTRADA.VALORICMSSUBSTITUICAO AS VAL_ICMS_SUBST,');
SQL.Add('  ENTRADA.VALORFUNRURAL AS VAL_FUNRURAL,');
SQL.Add('  1 AS COD_PERFIL,');
SQL.Add('  0 AS VAL_DESP_ACESS,');
SQL.Add('  ''N'' AS FLG_CANCELADO,');
SQL.Add('  ENTRADA.INFORMACAOCOMPLEMENTAR AS DES_OBSERVACAO,');
SQL.Add('  ENTRADA.CHAVENFE AS NUM_CHAVE_ACESSO,');
SQL.Add('  ENTRADA.VALORFCPST AS VAL_TOT_ST_FCP');
SQL.Add('FROM PUBLIC.NOTAENTRADA AS ENTRADA');
SQL.Add('-- LEFT JOIN (');
SQL.Add('--   SELECT DISTINCT ');
SQL.Add('-- 	ID_NOTAENTRADA, ');
SQL.Add('-- 	CFOPNOTA  ');
SQL.Add('-- FROM PUBLIC.NOTAENTRADAITEM');
SQL.Add('-- ) CFOP');
SQL.Add('-- ON ENTRADA.ID = CFOP.ID_NOTAENTRADA');
SQL.Add('WHERE ENTRADA.DATAEMISSAO >= ''2019-09-01''');
SQL.Add('AND ENTRADA.ID_SITUACAONOTAENTRADA = 1');
//SQL.Add('AND ENTRADA.NUMERONOTA = 168436');
SQL.Add('ORDER BY NUM_NF_FORN');


    Open;

    First;
    NumLinha := 0;

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);

      Layout.SetValues(QryPrincipal2, NumLinha, RecordCount);

      Layout.FieldByName('DTA_EMISSAO').AsDateTime := QryPrincipal2.FieldByName('DTA_EMISSAO').AsDateTime;
      Layout.FieldByName('DTA_ENTRADA').AsDateTime := QryPrincipal2.FieldByName('DTA_ENTRADA').AsDateTime;

      Layout.FieldByName('DES_OBSERVACAO').AsString := StrLBReplace(FieldByName('DES_OBSERVACAO').AsString);

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;
end;

procedure TFrmVRAlbano.GerarNFitensClientes;
begin
  inherited;
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
    SQL.Add('    SUBSTR(CAPA.COD_PART,2, 60) AS COD_CLIENTE,');
    SQL.Add('    CAPA.NUM_DOC AS NUM_NF_CLI,');
    SQL.Add('    CAPA.SER AS NUM_SERIE_NF,');
    SQL.Add('    ITEM.COD_ITEM AS COD_PRODUTO,');
    SQL.Add('');
    SQL.Add('    CASE');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 00 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 4 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 27');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 00 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 4.5 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 38');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 00 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 7 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 2');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 00 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 11 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 35');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 00 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 12 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 3');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 00 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 4');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 00 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 25 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 5');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 10 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 4 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 28');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 10 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 12 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 12');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 10 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 13');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 10 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 33.33 THEN 16');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 10 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 20 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 29');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 10 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 25 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 14');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 10 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 30 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 31');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 7 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 41.67 THEN 6');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 12 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 10.49 THEN 47');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 12 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 30 THEN 46');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 12 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 41.66 THEN 6');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 12 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 41.6667 THEN 6');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 12 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 41.667 THEN 6');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 12 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 41.6698 THEN 6');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 12 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 41.67 THEN 6');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 12 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 41.6702 THEN 6');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 12 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 41.68 THEN 6');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 12 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 50 THEN 9');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 12 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 58.33 THEN 9');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 12 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 60 THEN 9');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 33.33 THEN 7');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 33.333 THEN 7');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 51.11 THEN 9');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 61.11 THEN 8');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 30 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 0 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 4');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 40 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 0 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 1');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 41 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 0 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 23');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 51 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 0 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 20');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 51 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 7 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 20');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 51 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 12 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 20');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 51 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 20');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 0 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 13');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 0 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 33.33 THEN 16');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 12 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 12');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 12 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 41.67 THEN 15');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 13');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 33.3 THEN 16');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 33.33 THEN 16');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 41.67 THEN 15');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 61.11 THEN 17');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 20 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 29');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 22 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 30');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 25 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 14');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 25 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 51.99 THEN 18');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 25 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 52 THEN 18');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 28 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 13');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 30 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 31');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 70 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 12 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 33.33 THEN 16');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 70 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 12 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 41.66 THEN 15');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 70 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 12 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 41.67 THEN 15');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 70 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 12 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 41.6707 THEN 15');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 70 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 12 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 41.671 THEN 15');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 70 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 33 THEN 16');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 70 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 33.3 THEN 16');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 70 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 33.32 THEN 16');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 70 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 33.33 THEN 16');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 70 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 33.3333 THEN 16');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 70 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 33.34 THEN 16');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 70 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 61.11 THEN 17');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 70 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 25 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 51.99 THEN 18');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 70 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 25 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 52 THEN 18');
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 90 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 12 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 3');
    SQL.Add('    END AS COD_TRIBUTACAO, --');
    SQL.Add('');
    SQL.Add('    ITEM.EMBALAGEM_A AS QTD_EMBALAGEM,');
    SQL.Add('    ITEM.QTD AS QTD_ENTRADA,');
    SQL.Add('    ITEM.UNID AS DES_UNIDADE,');
    SQL.Add('    ITEM.VALOR_UNITARIO_A AS VAL_TABELA,');
    SQL.Add('    ITEM.VL_DESC AS VAL_DESCONTO_ITEM,');
    SQL.Add('    ITEM.SEGURO_RATEIO_A AS VAL_ACRESCIMO_ITEM, --');
    SQL.Add('    ITEM.VL_IPI AS VAL_IPI_ITEM,');
    SQL.Add('    ITEM.VL_ICMS AS VAL_CREDITO_ICMS,');
    SQL.Add('    ITEM.VL_ITEM AS VAL_TABELA_LIQ,');
    SQL.Add('    0 AS VAL_CUSTO_REP, --');
    SQL.Add('    '''' AS NUM_CGC,');
    SQL.Add('    ITEM.VL_BC_ICMS AS VAL_TOT_BC_ICMS,');
    SQL.Add('    0 AS VAL_TOT_OUTROS_ICMS, --');
    SQL.Add('    ITEM.CFOP AS COD_FISCAL,');
    SQL.Add('    ITEM.NUM_ITEM AS NUM_ITEM,');
    SQL.Add('    0 AS TIPO_IPI');
    SQL.Add('FROM');
    SQL.Add('    FIS_T_C170 ITEM');
    SQL.Add('INNER JOIN');
    SQL.Add('    FIS_T_C100 CAPA');
    SQL.Add('ON');
    SQL.Add('    ITEM.ID_C100 = CAPA.ID_C100');
    SQL.Add('LEFT JOIN ');
    SQL.Add('    BAS_T_COI CFOP ');
    SQL.Add('ON ');
    SQL.Add('    CAPA.ID_COI_A = CFOP.ID_COI  ');
    SQL.Add('INNER JOIN');
    SQL.Add('    PRODUTOS');
    SQL.Add('ON');
    SQL.Add('    ITEM.COD_ITEM = PRODUTOS.ID       ');
    SQL.Add('WHERE ');
    SQL.Add('    CAPA.ID_COI_A IN (1303, 30)');
    SQL.Add('AND ');
    SQL.Add('    CAPA.COD_SIT IN (6, 0)     ');
    SQL.Add('AND');
    SQL.Add('    CAPA.NUM_DOC IS NOT NULL');
    SQL.Add('AND');
    SQL.Add('    CAPA.STATUS_A = 2  ');


    SQL.Add('AND    ');
    SQL.Add('    CAPA.DT_DOC >= :INI ');
    SQL.Add('AND    ');
    SQL.Add('    CAPA.DT_DOC <= :FIM ');
    SQL.Add('AND    ');

    Parameters.ParamByName('INI').Value := FormatDateTime('dd/mm/yyyy', DtpInicial.Date);
    Parameters.ParamByName('FIM').Value := FormatDateTime('dd/mm/yyyy', DtpFinal.Date);


    Open;

    First;
    NumLinha := 0;

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);

      Layout.SetValues(QryPrincipal2, NumLinha, RecordCount);



      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;
end;

procedure TFrmVRAlbano.GerarNFitensFornec;
var
  fornecedor, nota, serie : string;
  count : integer;
begin
  inherited;

  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

SQL.Add('SELECT');
SQL.Add('  ENTRADA.ID_FORNECEDOR AS COD_FORNECEDOR,');
SQL.Add('  ENTRADA.NUMERONOTA AS NUM_NF_FORN,');
SQL.Add('  ENTRADA.SERIE AS NUM_SERIE_NF,');
SQL.Add('  ITEM.ID_PRODUTO AS COD_PRODUTO,');
SQL.Add('  ITEM.QTDEMBALAGEM AS QTD_EMBALAGEM,');
SQL.Add('  ITEM.QUANTIDADE AS QTD_ENTRADA,');
SQL.Add('  TIPOEMBALAGEM.DESCRICAO AS DES_UNIDADE,');
SQL.Add('');
SQL.Add('  CASE ITEM.ID_ALIQUOTA');
SQL.Add('    WHEN ''0'' THEN 2');
SQL.Add('    WHEN ''1'' THEN 3');
SQL.Add('    WHEN ''2'' THEN 4');
SQL.Add('    WHEN ''4'' THEN 8');
SQL.Add('    WHEN ''5'' THEN 6');
SQL.Add('    WHEN ''6'' THEN 1');
SQL.Add('    WHEN ''7'' THEN 25');
SQL.Add('    WHEN ''8'' THEN 22');
SQL.Add('    WHEN ''16'' THEN 20');
SQL.Add('	  WHEN ''20'' THEN 27');
SQL.Add('    WHEN ''21'' THEN 39');
SQL.Add('    WHEN ''23'' THEN 42');
SQL.Add('    WHEN ''29'' THEN 39');
SQL.Add('    WHEN ''37'' THEN 43');
SQL.Add('	  WHEN ''38'' THEN 3');
SQL.Add('    WHEN ''40'' THEN 41');
SQL.Add('    WHEN ''46'' THEN 44 ');
SQL.Add('    WHEN ''58'' THEN 25');
SQL.Add('    WHEN ''59'' THEN 1');
SQL.Add('	  WHEN ''60'' THEN 1');
SQL.Add('    WHEN ''61'' THEN 8');
SQL.Add('    WHEN ''62'' THEN 1');
SQL.Add('    WHEN ''69'' THEN 44');
SQL.Add('	  WHEN ''70'' THEN 6');
SQL.Add('	  WHEN ''75'' THEN 1');
SQL.Add('    WHEN ''78'' THEN 45');
SQL.Add('	  WHEN ''80'' THEN 45');
SQL.Add('	  WHEN ''81'' THEN 20');
SQL.Add('    WHEN ''82'' THEN 46');
SQL.Add('	  WHEN ''85'' THEN 1');
SQL.Add('	  WHEN ''86'' THEN 47');
SQL.Add('    WHEN ''87'' THEN 47');
SQL.Add('	  WHEN ''89'' THEN 17');
SQL.Add('	  WHEN ''91'' THEN 1');
SQL.Add('    WHEN ''93'' THEN 15');
SQL.Add('    WHEN ''94'' THEN 45');
SQL.Add('	  WHEN ''95'' THEN 49');
SQL.Add('	  WHEN ''97'' THEN 6');
SQL.Add('	  WHEN ''99'' THEN 1');
SQL.Add('	  WHEN ''100'' THEN 50');
SQL.Add('	  WHEN ''104'' THEN 17');
SQL.Add('  END AS COD_TRIBUTACAO,');
SQL.Add('  -- (ITEM.QUANTIDADE * ITEM.QTDEMBALAGEM) QTTOTAL,');
SQL.Add('  ITEM.VALOR AS VAL_TABELA,');
SQL.Add('  CASE ');
SQL.Add('    WHEN ITEM.VALORDESCONTO > 0 THEN ITEM.VALORDESCONTO / (ITEM.QUANTIDADE * ITEM.QTDEMBALAGEM)');
SQL.Add('    ELSE 0');
SQL.Add('  END AS VAL_DESCONTO_ITEM,');
SQL.Add('  0 AS VAL_ACRESCIMO_ITEM, --');
SQL.Add('  CASE');
SQL.Add('   WHEN ITEM.VALORIPI > 0 THEN ITEM.VALORIPI / ITEM.QUANTIDADE');
SQL.Add('   ELSE 0');
SQL.Add('  END AS VAL_IPI_ITEM,');
SQL.Add('  0 AS VAL_IPI_PER,');
SQL.Add('  CASE ');
SQL.Add('    WHEN ITEM.VALORICMSSUBSTITUICAO > 0 THEN ITEM.VALORICMSSUBSTITUICAO / (ITEM.QUANTIDADE * ITEM.QTDEMBALAGEM)');
SQL.Add('    ELSE 0');
SQL.Add('  END AS VAL_SUBST_ITEM,');
SQL.Add('  0 AS VAL_FRETE_ITEM,');
SQL.Add('  ITEM.VALORICMS AS VAL_CREDITO_ICMS,');
SQL.Add('  0 AS VAL_VENDA_VAREJO,');
SQL.Add('  ITEM.VALORTOTAL AS VAL_TABELA_LIQ,');
SQL.Add('  '''' AS NUM_CGC,');
SQL.Add('  ITEM.VALORBASECALCULO AS VAL_TOT_BC_ICMS,');
SQL.Add('  ITEM.VALOROUTRAS AS VAL_TOT_OUTROS_ICMS,');
SQL.Add('  CAST(REPLACE(ITEM.CFOP, ''.'', '''') AS INTEGER) AS CFOP,');
SQL.Add('  ITEM.VALORISENTO AS VAL_TOT_ISENTO,');
SQL.Add('  ITEM.VALORBASESUBSTITUICAO AS VAL_TOT_BC_ST,');
SQL.Add('  ITEM.VALORICMSSUBSTITUICAO AS VAL_TOT_ST,');
SQL.Add('  1 AS NUM_ITEM,');
SQL.Add('  0 AS TIPO_IPI,');
SQL.Add('  CONCAT( LPAD(CAST(PRODUTO.NCM1 AS VARCHAR(4)), 4, ''0''), LPAD(CAST(PRODUTO.NCM2 AS VARCHAR(2)), 2, ''0''), LPAD(CAST(PRODUTO.NCM3 AS VARCHAR(2)), 2, ''0'') ) AS NUM_NCM,');
SQL.Add('  '''' AS DES_REFERENCIA,');
SQL.Add('  ITEM.VALORFCP AS VAL_TOT_ICMS_FCP,');
SQL.Add('  ITEM.VALORFCPST AS VAL_TOT_ST_FCP');
SQL.Add('FROM PUBLIC.NOTAENTRADAITEM AS ITEM');
SQL.Add('LEFT JOIN PUBLIC.NOTAENTRADA AS ENTRADA');
SQL.Add('ON ITEM.ID_NOTAENTRADA = ENTRADA.ID');
SQL.Add('LEFT JOIN PUBLIC.PRODUTO AS PRODUTO');
SQL.Add('ON ITEM.ID_PRODUTO = PRODUTO.ID');
SQL.Add('LEFT JOIN PUBLIC.TIPOEMBALAGEM');
    SQL.Add('ON PRODUTO.ID_TIPOEMBALAGEM = TIPOEMBALAGEM.ID');
    SQL.Add('WHERE ENTRADA.DATAEMISSAO >= ''2019-09-01''');
    SQL.Add('AND ENTRADA.ID_SITUACAONOTAENTRADA = 1');
//    SQL.Add('AND ENTRADA.NUMERONOTA = 168436');

    Open;

    First;
    NumLinha := 0;
    count := 0;

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);

      Layout.SetValues(QryPrincipal2, NumLinha, RecordCount);

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

//      Layout.FieldByName('VAL_TABELA').AsCurrency := Layout.FieldByName('VAL_TABELA').AsCurrency / Layout.FieldByName('QTTOTAL').AsCurrency;
//
      Layout.FieldByName('COD_PRODUTO').AsString := GerarPLU(QryPrincipal2.FieldByName('COD_PRODUTO').AsString);

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;

end;

procedure TFrmVRAlbano.GerarProdForn;
begin
  inherited;

  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
    SQL.Add('  PRODUTOFORNECEDOR.ID_PRODUTO AS COD_PRODUTO,');
    SQL.Add('  PRODUTOFORNECEDOR.ID_FORNECEDOR AS COD_FORNECEDOR,');
    SQL.Add('  PRODUTOFORNECEDOR.CODIGOEXTERNO AS DES_REFERENCIA,');
    SQL.Add('  '''' AS NUM_CGC,');
    SQL.Add('  0 AS COD_DIVISAO,');
    SQL.Add('  TIPOEMBALAGEM.DESCRICAO AS DES_UNIDADE_COMPRA,');
    SQL.Add('  COALESCE(PRODUTOFORNECEDOR.QTDEMBALAGEM, 1) AS QTD_EMBALAGEM_COMPRA,');
    SQL.Add('  0 AS QTD_TROCA,');
    SQL.Add('  ''N'' AS FLG_PREFERENCIAL');
    SQL.Add('FROM PUBLIC.PRODUTOFORNECEDOR');
    SQL.Add('LEFT JOIN PUBLIC.PRODUTO   ');
    SQL.Add('ON PRODUTOFORNECEDOR.ID_PRODUTO = PRODUTO.ID');
    SQL.Add('LEFT JOIN PUBLIC.TIPOEMBALAGEM');
    SQL.Add('ON PRODUTO.ID_TIPOEMBALAGEM = TIPOEMBALAGEM.ID');



    Open;

    First;

    NumLinha := 0;

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);

      Layout.SetValues(QryPrincipal2, NumLinha, RecordCount);

      Layout.FieldByName('COD_PRODUTO').AsString := GerarPLU(QryPrincipal2.FieldByName('COD_PRODUTO').AsString);

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;

end;

procedure TFrmVRAlbano.GerarProdLoja;
var
   cod_produto : Integer;
begin
  inherited;
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
    SQL.Add('  PRODUTO.ID AS COD_PRODUTO,');
    SQL.Add('  PRODUTOCOMPLEMENTO.CUSTOCOMIMPOSTO AS VAL_CUSTO_REP,');
    SQL.Add('  PRODUTOCOMPLEMENTO.PRECOVENDA AS VAL_VENDA,');
    SQL.Add('  0 AS VAL_OFERTA,');
    SQL.Add('  PRODUTOCOMPLEMENTO.ESTOQUE AS QTD_EST_VDA,');
    SQL.Add('  '''' AS TECLA_BALANCA,');
    SQL.Add('');
    SQL.Add('  CASE PRODUTOALIQUOTA.ID_ALIQUOTADEBITO');
    SQL.Add('    WHEN ''0'' THEN 2');
    SQL.Add('    WHEN ''2'' THEN 4');
    SQL.Add('    WHEN ''4'' THEN 8');
    SQL.Add('    WHEN ''5'' THEN 6');
    SQL.Add('    WHEN ''6'' THEN 1');
    SQL.Add('    WHEN ''7'' THEN 25');
    SQL.Add('    WHEN ''21'' THEN 39 ');
    SQL.Add('    WHEN ''29'' THEN 39 ');
    SQL.Add('    WHEN ''30'' THEN 40');
    SQL.Add('    WHEN ''38'' THEN 3');
    SQL.Add('    WHEN ''40'' THEN 41');
    SQL.Add('    WHEN ''75'' THEN 1');
    SQL.Add('    WHEN ''99'' THEN 1');
    SQL.Add('  END AS COD_TRIBUTACAO,');
    SQL.Add('');
    SQL.Add('  PRODUTOCOMPLEMENTO.MARGEM AS VAL_MARGEM,');
    SQL.Add('  1 AS QTD_ETIQUETA,');
    SQL.Add('');
    SQL.Add('  CASE PRODUTOALIQUOTA.ID_ALIQUOTACREDITO');
    SQL.Add('    WHEN ''0'' THEN 2');
    SQL.Add('    WHEN ''2'' THEN 4');
    SQL.Add('    WHEN ''4'' THEN 8');
    SQL.Add('    WHEN ''5'' THEN 6');
    SQL.Add('    WHEN ''6'' THEN 1');
    SQL.Add('    WHEN ''7'' THEN 25');
    SQL.Add('    WHEN ''21'' THEN 39 ');
    SQL.Add('    WHEN ''29'' THEN 39 ');
    SQL.Add('    WHEN ''30'' THEN 40');
    SQL.Add('    WHEN ''38'' THEN 3');
    SQL.Add('    WHEN ''40'' THEN 41');
    SQL.Add('    WHEN ''75'' THEN 1');
    SQL.Add('    WHEN ''99'' THEN 1');
    SQL.Add('  END AS COD_TRIB_ENTRADA,');
    SQL.Add('');
    SQL.Add('  CASE PRODUTOCOMPLEMENTO.ID_SITUACAOCADASTRO');
    SQL.Add('    WHEN 1 THEN ''N''');
    SQL.Add('    ELSE ''S''');
    SQL.Add('  END AS FLG_INATIVO,');
    SQL.Add('');
    SQL.Add('  PRODUTO.ID AS COD_PRODUTO_ANT,');
    SQL.Add('  CONCAT( LPAD(CAST(PRODUTO.NCM1 AS VARCHAR(4)), 4, ''0''), LPAD(CAST(PRODUTO.NCM2 AS VARCHAR(2)), 2, ''0''), LPAD(CAST(PRODUTO.NCM3 AS VARCHAR(2)), 2, ''0'') ) AS NUM_NCM,');
    SQL.Add('  0 AS TIPO_NCM,');
    SQL.Add('  0 AS VAL_VENDA_2,');
    SQL.Add('  '''' AS DTA_VALIDA_OFERTA,');
    SQL.Add('  PRODUTOCOMPLEMENTO.ESTOQUEMINIMO AS QTD_EST_MINIMO,');
    SQL.Add('  NULL AS COD_VASILHAME,');
    SQL.Add('  ''N'' AS FORA_LINHA,');
    SQL.Add('  0 AS QTD_PRECO_DIF,');
    SQL.Add('  0 AS VAL_FORCA_VDA,');
    SQL.Add('');
    SQL.Add('  CASE');
    SQL.Add('    WHEN CEST.ID IS NULL THEN ''9999999''');
    SQL.Add('    WHEN CONCAT(LPAD(CAST(CEST1 AS VARCHAR(2)), 2, ''0''), LPAD(CAST(CEST2 AS VARCHAR(3)), 3,''0''), LPAD(CAST(CEST3 AS VARCHAR(2)), 2, ''0'')) = ''0000000'' THEN ''9999999''');
    SQL.Add('    ELSE CONCAT(LPAD(CAST(CEST1 AS VARCHAR(2)), 2, ''0''), LPAD(CAST(CEST2 AS VARCHAR(3)), 3,''0''), LPAD(CAST(CEST3 AS VARCHAR(2)), 2, ''0''))');
    SQL.Add('  END AS NUM_CEST,');
    SQL.Add('');
    SQL.Add('  0 AS PER_IVA,');
    SQL.Add('  0 AS PER_FCP_ST,');
    SQL.Add('  COALESCE(CODIGOBENEFICIOCST.CODIGO, '''') AS COD_BENEF_FISCAL');
    SQL.Add('FROM PUBLIC.PRODUTO AS PRODUTO');
    SQL.Add('');
    SQL.Add('LEFT JOIN PUBLIC.PRODUTOALIQUOTA AS PRODUTOALIQUOTA');
    SQL.Add('ON PRODUTO.ID = PRODUTOALIQUOTA.ID_PRODUTO');
    SQL.Add('');
    SQL.Add('LEFT JOIN PUBLIC.PRODUTOALIQUOTABENEFICIO AS PRODUTOALIQUOTABENEFICIO');
    SQL.Add('ON PRODUTOALIQUOTA.ID = PRODUTOALIQUOTABENEFICIO.ID_PRODUTOALIQUOTA');
    SQL.Add('');
    SQL.Add('LEFT JOIN PUBLIC.CODIGOBENEFICIOCST AS CODIGOBENEFICIOCST');
    SQL.Add('ON PRODUTOALIQUOTABENEFICIO.ID_BENEFICIO = CODIGOBENEFICIOCST.ID');
    SQL.Add('');
    SQL.Add('LEFT JOIN PUBLIC.PRODUTOCOMPLEMENTO AS PRODUTOCOMPLEMENTO');
    SQL.Add('ON PRODUTO.ID = PRODUTOCOMPLEMENTO.ID_PRODUTO');
    SQL.Add('');
    SQL.Add('LEFT JOIN PUBLIC.CEST AS CEST');
    SQL.Add('ON PRODUTO.ID_CEST = CEST.ID');





    Open;
    First;
    NumLinha := 0;


    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);
      Layout.SetValues(QryPrincipal2, NumLinha, RecordCount);
//      inc(cod_produto);

//      Layout.FieldByName('COD_PRODUTO').AsInteger := cod_produto;

      Layout.FieldByName('COD_PRODUTO').AsString := GerarPLU( Layout.FieldByName('COD_PRODUTO').AsString );
      Layout.FieldByName('COD_PRODUTO_ANT').AsString := GerarPLU( Layout.FieldByName('COD_PRODUTO').AsString );

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
    end;
    Next;
    end;
    Close;
  end;
end;

procedure TFrmVRAlbano.GerarProdSimilar;
begin
  inherited;
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
    SQL.Add('  PRODSIMILAR.ID AS COD_PRODUTO_SIMILAR,');
    SQL.Add('  PRODSIMILAR.DESCRICAO AS DES_PRODUTO_SIMILAR,');
    SQL.Add('  0 AS VAL_META');
    SQL.Add('FROM FAMILIAPRODUTO AS PRODSIMILAR');
    SQL.Add('WHERE ID_SITUACAOCADASTRO = 1');


    Open;    
    
    First;
    NumLinha := 0;

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);

      Layout.SetValues(QryPrincipal2, NumLinha, RecordCount);

      Layout.FieldByName('COD_PRODUTO_SIMILAR').AsString := GerarPLU( Layout.FieldByName('COD_PRODUTO_SIMILAR').AsString );

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;
end;

end.
