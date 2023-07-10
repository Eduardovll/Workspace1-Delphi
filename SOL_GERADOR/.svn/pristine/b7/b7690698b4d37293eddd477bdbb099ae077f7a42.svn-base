unit UFrmGetWayMama;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes,
  Vcl.Graphics, ComObj,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, UFrmModelo, Data.DBXOracle, Data.DB,
  Data.SqlExpr, Vcl.Menus, Vcl.StdCtrls, Vcl.Buttons, Vcl.ComCtrls,
  Vcl.Imaging.pngimage, Vcl.ExtCtrls, Data.DBXFirebird, Data.Win.ADODB,
  Vcl.Grids, Vcl.DBGrids, Datasnap.Provider, Datasnap.DBClient,
  //dxGDIPlusClasses,
  Math;

type
  TFrmGetWayMama = class(TFrmModeloSis)
    ADOSQLServer: TADOConnection;
    QryPrincipal2: TADOQuery;
    CbxLoja: TComboBox;
    lblLoja: TLabel;
    procedure btnGeraCestClick(Sender: TObject);
    procedure BtnAmarrarCestClick(Sender: TObject);
    procedure BtnGerarClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
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

    procedure GerarFinanceiro( Tipo, Situacao :Integer ); Override;
    procedure GerarFinanceiroReceber(Aberto:String);      Override;
    procedure GerarFinanceiroReceberCartao;               Override;
    procedure GerarFinanceiroPagar(Aberto:String);        Override;

    procedure GerarScriptCEST;
    procedure GerarScriptAmarrarCEST;

  end;

var
  FrmGetWayMama: TFrmGetWayMama;
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


procedure TFrmGetWayMama.GerarProducao;
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

procedure TFrmGetWayMama.GerarProduto;
begin
  inherited;
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
    SQL.Add('    PRODUTOS.CODPROD_SAMBANET AS COD_PRODUTO,');
    SQL.Add('    PRODUTOS.BARRA AS COD_BARRA_PRINCIPAL,');
    SQL.Add('    PRODUTOS.DESC_PDV AS DES_REDUZIDA,');
    SQL.Add('    PRODUTOS.DESCRICAO AS DES_PRODUTO,');
    SQL.Add('    1 AS QTD_EMBALAGEM_COMPRA, --');
    SQL.Add('    PRODUTOS.UNIDADE AS DES_UNIDADE_COMPRA, --');
    SQL.Add('    1 AS QTD_EMBALAGEM_VENDA, -- possui a informação no banco');
    SQL.Add('    PRODUTOS.UNIDADE AS DES_UNIDADE_VENDA, --');
    SQL.Add('    0 AS TIPO_IPI, --');
    SQL.Add('    0 AS VAL_IPI, --');
    SQL.Add('    PRODUTOS.CODCRECEITA AS COD_SECAO, --');
    SQL.Add('    PRODUTOS.CODGRUPO AS COD_GRUPO, --');
    SQL.Add('    PRODUTOS.CODCATEGORIA AS COD_SUB_GRUPO, --');
    SQL.Add('    0 AS COD_PRODUTO_SIMILAR,');
    SQL.Add('');
    SQL.Add('    CASE ');
    SQL.Add('        WHEN PRODUTOS.CODSETOR IS NOT NULL AND UPPER(PRODUTOS.UNIDADE) = ''KG'' THEN ''S''');
    SQL.Add('        ELSE ''N''');
    SQL.Add('    END AS IPV, -- SETOR');
    SQL.Add('');
    SQL.Add('    0 AS DIAS_VALIDADE, -- possui a informação no banco');
    SQL.Add('    0 AS TIPO_PRODUTO,');
    SQL.Add('');
    SQL.Add('    CASE COALESCE(CST_PISSAIDA, ''01'')');
    SQL.Add('        WHEN ''01'' THEN ''S''');
    SQL.Add('        WHEN ''04'' THEN ''N''');
    SQL.Add('        WHEN ''05'' THEN ''N''');
    SQL.Add('        WHEN ''06'' THEN ''N''');
    SQL.Add('        WHEN ''07'' THEN ''S''');
    SQL.Add('        WHEN ''09'' THEN ''N''');
    SQL.Add('        WHEN ''49'' THEN ''S''');
    SQL.Add('    END AS FLG_NAO_PIS_COFINS, ');
    SQL.Add('');
    SQL.Add('    CASE ');
    SQL.Add('        WHEN PRODUTOS.CODSETOR IS NOT NULL THEN ''S''');
    SQL.Add('        ELSE ''N''');
    SQL.Add('    END AS FLG_ENVIA_BALANCA, -- SETOR');
    SQL.Add('');
    SQL.Add('    CASE COALESCE(CST_PISSAIDA, ''01'')');
    SQL.Add('        WHEN ''01'' THEN -1');
    SQL.Add('        WHEN ''04'' THEN 1');
    SQL.Add('        WHEN ''05'' THEN 2');
    SQL.Add('        WHEN ''06'' THEN 0');
    SQL.Add('        WHEN ''07'' THEN -1');
    SQL.Add('        WHEN ''09'' THEN 4');
    SQL.Add('        WHEN ''49'' THEN -1');
    SQL.Add('    END AS TIPO_NAO_PIS_COFINS,');
    SQL.Add('');
    SQL.Add('    0 AS TIPO_EVENTO,');
    SQL.Add('    0 AS COD_ASSOCIADO,');
    SQL.Add('    '''' AS DES_OBSERVACAO,');
    SQL.Add('    0 AS COD_INFO_NUTRICIONAL,');
    SQL.Add('    0 AS COD_INFO_RECEITA,');
    SQL.Add('    999 AS COD_TAB_SPED, -- ');
    SQL.Add('    PRODUTOS.ALCOOLICO AS FLG_ALCOOLICO,');
    SQL.Add('    0 AS TIPO_ESPECIE,');
    SQL.Add('    0 AS COD_CLASSIF,');
    SQL.Add('    1 AS VAL_VDA_PESO_BRUTO, --');
    SQL.Add('    1 AS VAL_PESO_EMB, --');
    SQL.Add('    0 AS TIPO_EXPLOSAO_COMPRA,');
    SQL.Add('    NULL AS DTA_INI_OPER,');
    SQL.Add('    '''' AS DES_PLAQUETA,');
    SQL.Add('    '''' AS MES_ANO_INI_DEPREC,');
    SQL.Add('    0 AS TIPO_BEM,');
    SQL.Add('    0 AS COD_FORNECEDOR,');
    SQL.Add('    0 AS NUM_NF,');
    SQL.Add('    NULL AS DTA_ENTRADA,');
    SQL.Add('    0 AS COD_NAT_BEM,');
    SQL.Add('    0 AS VAL_ORIG_BEM');
    SQL.Add('FROM ');
    SQL.Add('    PRODUTOS');


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

      Layout.FieldByName('COD_PRODUTO').AsString := GerarPLU( Layout.FieldByName('COD_PRODUTO').AsString );


      Layout.FieldByName('DES_OBSERVACAO').AsString := StrReplace(StrLBReplace(FieldByName('DES_OBSERVACAO').AsString), '\n', '');
      Layout.FieldByName('DES_REDUZIDA').AsString := StrReplace(StrLBReplace(FieldByName('DES_REDUZIDA').AsString), '\n', '');
      Layout.FieldByName('DES_PRODUTO').AsString := StrReplace(StrLBReplace(FieldByName('DES_PRODUTO').AsString), '\n', '');

      if( Layout.FieldByName('COD_BARRA_PRINCIPAL').AsCurrency = 0 ) then
      begin
         Layout.FieldByName('COD_BARRA_PRINCIPAL').AsString := '';
      end
      else
      begin
         if ( Length(TiraZerosEsquerda(Layout.FieldByName('COD_BARRA_PRINCIPAL').AsString)) <= 8 ) then
            Layout.FieldByName('COD_BARRA_PRINCIPAL').AsString := GerarPLU( Layout.FieldByName('COD_BARRA_PRINCIPAL').AsString );
      end;

//      if( not CodBarrasValido(Layout.FieldByName('COD_BARRA_PRINCIPAL').AsString) ) then
//         Layout.FieldByName('COD_BARRA_PRINCIPAL').AsString := '';



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

procedure TFrmGetWayMama.GerarScriptAmarrarCEST;
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

procedure TFrmGetWayMama.GerarScriptCEST;
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

procedure TFrmGetWayMama.GerarSecao;
begin
  inherited;

  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT DISTINCT');
    SQL.Add('    CODCRECEITA AS COD_SECAO,');
    SQL.Add('    ''A DEFINIR'' AS DES_SECAO,');
    SQL.Add('    0 AS VAL_META');
    SQL.Add('FROM');
    SQL.Add('    PRODUTOS');



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

procedure TFrmGetWayMama.GerarSubGrupo;
begin
  inherited;

  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT DISTINCT');
    SQL.Add('    CODCRECEITA AS COD_SECAO,');
    SQL.Add('    CODGRUPO AS COD_GRUPO,');
    SQL.Add('    CODCATEGORIA AS COD_SUB_GRUPO,');
    SQL.Add('    ''A DEFINIR'' AS DES_SUB_GRUPO,');
    SQL.Add('    0 AS VAL_META,');
    SQL.Add('    0 AS VAL_MARGEM_REF,');
    SQL.Add('    0 AS QTD_DIA_SEGURANCA,');
    SQL.Add('    ''N'' AS FLG_ALCOOLICO');
    SQL.Add('FROM');
    SQL.Add('    PRODUTOS');



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

procedure TFrmGetWayMama.GerarTransportadora;
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


procedure TFrmGetWayMama.BtnAmarrarCestClick(Sender: TObject);
begin
  inherited;
    inherited;
  FlgGeraAmarrarCest := True;
  BtnGerar.Click;
  FlgGeraAmarrarCest := False;
end;

procedure TFrmGetWayMama.btnGeraCestClick(Sender: TObject);
begin
  inherited;
  FlgGeraCest := True;
  BtnGerar.Click;
  FlgGeraCest := False;
end;

procedure TFrmGetWayMama.BtnGerarClick(Sender: TObject);
begin
   ADOSQLServer.Connected := false;
//   ADOSQLServer.ConnectionString := 'Provider=MSDASQL.1;Persist Security Info=False;Data Source='+edtSchema.Text+';User ID='+edtInst.Text+';Password'+edtSenhaOracle.Text+'';
   ADOSQLServer.ConnectionString := 'Provider=MSDASQL.1;Persist Security Info=False;Data Source='+edtSchema.Text+';User ID='+edtInst.Text+';Password='+edtSenhaOracle.Text+'';

   ADOSQLServer.Connected := true;
  inherited;

   ADOSQLServer.Connected := false;
end;

procedure TFrmGetWayMama.FormCreate(Sender: TObject);
begin
  inherited;
//  Left:=(Screen.Width-Width)  div 2;
//  Top:=(Screen.Height-Height) div 2;
end;

procedure TFrmGetWayMama.GerarCest;
var
   count : integer;
begin
  inherited;
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT DISTINCT');
    SQL.Add('    0 AS COD_CEST,');
    SQL.Add('    CASE WHEN COALESCE(PRODUTOS.CEST, ''0000000'') = ''0000000'' THEN ''9999999'' ELSE PRODUTOS.CEST END AS NUM_CEST,');
    SQL.Add('    ''A DEFINIR'' AS DES_CEST');
    SQL.Add('FROM    ');
    SQL.Add('    PRODUTOS');

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

procedure TFrmGetWayMama.GerarCliente;
begin

   inherited;
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
    SQL.Add('    CLIENTES.CODCLIE AS COD_CLIENTE,');
    SQL.Add('    CLIENTES.RAZAO AS DES_CLIENTE,');
    SQL.Add('    CLIENTES.CNPJ_CPF AS NUM_CGC,');
    SQL.Add('    CLIENTES.IE AS NUM_INSC_EST,');
    SQL.Add('    CLIENTES.ENDERECO AS DES_ENDERECO,');
    SQL.Add('    CLIENTES.BAIRRO AS DES_BAIRRO,');
    SQL.Add('    CLIENTES.CIDADE AS DES_CIDADE,');
    SQL.Add('    CLIENTES.ESTADO AS DES_SIGLA,');
    SQL.Add('    CLIENTES.CEP AS NUM_CEP,');
    SQL.Add('    CLIENTES.TELEFONE AS NUM_FONE,');
    SQL.Add('    CLIENTES.FAX AS NUM_FAX,');
    SQL.Add('    CLIENTES.CONTATO AS DES_CONTATO,');
    SQL.Add('    0 AS FLG_SEXO,');
    SQL.Add('    0 AS VAL_LIMITE_CRETID,');
    SQL.Add('    0 AS VAL_LIMITE_CONV,');
    SQL.Add('    0 AS VAL_DEBITO,');
    SQL.Add('    0 AS VAL_RENDA,');
    SQL.Add('    0 AS COD_CONVENIO,');
    SQL.Add('    0 AS COD_STATUS_PDV,');
    SQL.Add('');
    SQL.Add('    CASE ');
    SQL.Add('        WHEN LEN(CLIENTES.CNPJ_CPF) > 11 THEN ''S''');
    SQL.Add('        ELSE ''N'' ');
    SQL.Add('    END AS FLG_EMPRESA,');
    SQL.Add('');
    SQL.Add('    ''N'' AS FLG_CONVENIO,');
    SQL.Add('    CLIENTES.ME AS MICRO_EMPRESA,');
    SQL.Add('    CLIENTES.DTCAD AS DTA_CADASTRO,');
    SQL.Add('    CLIENTES.NUMERO AS NUM_ENDERECO,');
    SQL.Add('    CLIENTES.RG AS NUM_RG,');
    SQL.Add('    0 AS FLG_EST_CIVIL,');
    SQL.Add('    CLIENTES.CELULAR AS NUM_CELULAR,');
    SQL.Add('    '''' AS DTA_ALTERACAO,');
    SQL.Add('    '''' AS DES_OBSERVACAO,');
    SQL.Add('    CLIENTES.COMPLEMENTO AS DES_COMPLEMENTO,');
    SQL.Add('    CLIENTES.EMAIL AS DES_EMAIL,');
    SQL.Add('    CLIENTES.FANTASIA AS DES_FANTASIA,');
    SQL.Add('    CLIENTES.DTANIVER AS DTA_NASCIMENTO,');
    SQL.Add('    '''' AS DES_PAI,');
    SQL.Add('    '''' AS DES_MAE,');
    SQL.Add('    '''' AS DES_CONJUGE,');
    SQL.Add('    '''' AS NUM_CPF_CONJUGE,');
    SQL.Add('    0 AS VAL_DEB_CONV,');
    SQL.Add('');
    SQL.Add('    CASE ');
    SQL.Add('        WHEN CLIENTES.ATIVO = ''S'' THEN ''N''');
    SQL.Add('        ELSE ''S'' ');
    SQL.Add('    END AS INATIVO,');
    SQL.Add('    ');
    SQL.Add('    '''' AS DES_MATRICULA,');
    SQL.Add('    ''N'' AS NUM_CGC_ASSOCIADO,');
    SQL.Add('    ''N'' AS FLG_PROD_RURAL,');
    SQL.Add('    0 AS COD_STATUS_PDV_CONV,');
    SQL.Add('    ''S'' AS FLG_ENVIA_CODIGO,');
    SQL.Add('    '''' AS DTA_NASC_CONJUGE,');
    SQL.Add('    0 AS COD_CLASSIF');
    SQL.Add('FROM');
    SQL.Add('    CLIENTES');

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

      Layout.FieldByName('DTA_NASCIMENTO').AsDateTime := FieldByName('DTA_NASCIMENTO').AsDateTime;
      Layout.FieldByName('DTA_CADASTRO').AsDateTime := FieldByName('DTA_CADASTRO').AsDateTime;

      Layout.FieldByName('NUM_FONE').AsString := StrRetNums( FieldByName('NUM_FONE').AsString );
//
      if Layout.FieldByName('FLG_EMPRESA').AsString = 'S' then
      begin
        if not ValidaCGC(Layout.FieldByName('NUM_CGC').AsString) then
          Layout.FieldByName('NUM_CGC').AsString := '';
      end
      else
        if not ValidaCpf(Layout.FieldByName('NUM_CGC').AsString) then
          Layout.FieldByName('NUM_CGC').AsString := '';
//
//      Layout.FieldByName('DES_OBSERVACAO').AsString := StrReplace(StrLBReplace(FieldByName('DES_OBSERVACAO').AsString), '\n', '');
//      Layout.FieldByName('DES_EMAIL').AsString := StrReplace(StrLBReplace(FieldByName('DES_EMAIL').AsString), '\n', '');

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;

end;

procedure TFrmGetWayMama.GerarCodigoBarras;
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
    SQL.Add('     PRODUTOS.CODPROD_SAMBANET AS COD_PRODUTO,');
    SQL.Add('     PRODUTOS.BARRA AS COD_EAN');
    SQL.Add('FROM');
    SQL.Add('     PRODUTOS');
    SQL.Add('');
    SQL.Add('UNION ALL');
    SQL.Add('');
    SQL.Add('SELECT');
    SQL.Add('     PRODUTOS.CODPROD_SAMBANET AS COD_PRUDUTO,');
    SQL.Add('     ALTERNATIVO.BARRA AS COD_EAN');
    SQL.Add('FROM');
    SQL.Add('     ALTERNATIVO');
    SQL.Add('LEFT JOIN');
    SQL.Add('     PRODUTOS ON ALTERNATIVO.CODPROD = PRODUTOS.CODPROD');


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

      Layout.FieldByName('COD_PRODUTO').AsString := GerarPLU( Layout.FieldByName('COD_PRODUTO').AsString );
      Layout.FieldByName('COD_EAN').AsString := StrRetNums( Layout.FieldByName('COD_EAN').AsString );


      if( Layout.FieldByName('COD_EAN').AsCurrency = 0 ) then
      begin
         Layout.FieldByName('COD_EAN').AsString := '';
      end
      else
      begin
         if ( Length(TiraZerosEsquerda(Layout.FieldByName('COD_EAN').AsString)) <= 8 ) then
            Layout.FieldByName('COD_EAN').AsString := GerarPLU( Layout.FieldByName('COD_EAN').AsString );
      end;
//
//      if( not CodBarrasValido(Layout.FieldByName('COD_EAN').AsString) ) then
//        Layout.FieldByName('COD_EAN').AsString := '';

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;
end;

procedure TFrmGetWayMama.GerarComposicao;
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

procedure TFrmGetWayMama.GerarCondPagCli;
begin
  inherited;

  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
    SQL.Add('    CLIENTES.ID AS COD_CLIENTE,');
    SQL.Add('    30 AS NUM_CONDICAO,');
    SQL.Add('    2 AS COD_CONDICAO,');
    SQL.Add('    1 AS COD_ENTIDADE');
    SQL.Add('FROM');
    SQL.Add('    CLIENTES');


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

procedure TFrmGetWayMama.GerarCondPagForn;
begin
  inherited;

  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
    SQL.Add('    FORNECEDORES.ID AS COD_FORNECEDOR,');
    SQL.Add('    COALESCE(FORNECEDORES.CONDPAGTO, ''30'') AS NUM_CONDICAO,');
    SQL.Add('    2 AS COD_CONDICAO,');
    SQL.Add('    1 AS COD_ENTIDADE,');
    SQL.Add('    '''' AS NUM_CGC');
    SQL.Add('FROM');
    SQL.Add('    FORNECEDORES');
    SQL.Add('WHERE');
    SQL.Add('(    ');
    SQL.Add('    INSTR(CONDPAGTO, ''/'') = 0');
    SQL.Add('OR');
    SQL.Add('    INSTR(CONDPAGTO, ''/'') IS NULL');
    SQL.Add(') ');


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

procedure TFrmGetWayMama.GerarDecomposicao;
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

procedure TFrmGetWayMama.GerarDivisaoForn;
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

procedure TFrmGetWayMama.GerarFinanceiro(Tipo, Situacao: Integer);
begin
  inherited;
  if Tipo = 1 then
    GerarFinanceiroPagar(IntToStr(Situacao));

  if Tipo = 2 then
    GerarFinanceiroReceber(IntToStr(Situacao));

  if Tipo = 3 then
    GerarFinanceiroReceberCartao;

end;

procedure TFrmGetWayMama.GerarFinanceiroPagar(Aberto: String);
begin
  inherited;
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
    SQL.Add('    CASE PAGAR.TIPO_CADASTRO');
    SQL.Add('        WHEN 3 THEN 2');
    SQL.Add('        WHEN 2 THEN 1');
    SQL.Add('    END AS TIPO_PARCEIRO, -- TIPO_CADASTRO');
    SQL.Add('    PAGAR.ID_CADASTRO AS COD_PARCEIRO,');
    SQL.Add('    0 AS TIPO_CONTA,');
    SQL.Add('');
    SQL.Add('    CASE PAGAR.FORMA_PAGTO');
    SQL.Add('        WHEN 0 THEN 1');
    SQL.Add('        WHEN 1 THEN 8');
    SQL.Add('        WHEN 2 THEN 28');
    SQL.Add('        WHEN 3 THEN 29');
    SQL.Add('        WHEN 4 THEN 3');
    SQL.Add('        WHEN 5 THEN 2       ');
    SQL.Add('        ELSE 1       ');
    SQL.Add('    END AS COD_ENTIDADE,');
    SQL.Add('');
    SQL.Add('    PAGAR.ARQUIVO AS NUM_DOCTO,');
    SQL.Add('    999 AS COD_BANCO,');
    SQL.Add('    '''' AS DES_BANCO,');
    SQL.Add('    PAGAR.EMISSAO AS DTA_EMISSAO,');
    SQL.Add('    PAGAR.VENCIMENTO AS DTA_VENCIMENTO,');
    SQL.Add('    PAGAR.VALOR AS VAL_PARCELA,');
    SQL.Add('    PAGAR.ACRESCIMO + PAGAR.CARTORIO + COALESCE(PAGAR.CREDITO, 0) AS VAL_JUROS,');
    SQL.Add('    PAGAR.DESCONTO AS VAL_DESCONTO,');
    SQL.Add('');
    SQL.Add('    CASE ');
    SQL.Add('        WHEN PAGAR.PAGAMENTO IS NULL THEN ''N''');
    SQL.Add('        ELSE ''S''');
    SQL.Add('    END AS FLG_QUITADO,');
    SQL.Add('');
    SQL.Add('    CASE ');
    SQL.Add('        WHEN PAGAR.PAGAMENTO IS NULL THEN NULL');
    SQL.Add('        ELSE PAGAR.PAGAMENTO');
    SQL.Add('    END AS DTA_QUITADA,');
    SQL.Add('');
    SQL.Add('    998 AS COD_CATEGORIA,');
    SQL.Add('    998 AS COD_SUBCATEGORIA,');
    SQL.Add('    PAGAR.PARCELA AS NUM_PARCELA,');
    SQL.Add('    PAGAR.TOTAL_PARCELA AS QTD_PARCELA,');
    SQL.Add('    PAGAR.EMPRESA AS COD_LOJA,');
    SQL.Add('    PAGAR.CPF_CNPJ AS NUM_CGC,');
    SQL.Add('    COALESCE(PAGAR.BORDERO, 0) AS NUM_BORDERO,');
    SQL.Add('    PAGAR.NF AS NUM_NF,');
    SQL.Add('    '''' AS NUM_SERIE_NF,');
    SQL.Add('    NF.VAL_TOTAL_NF AS VAL_TOTAL_NF, -- EFETUAR A SOMA');
    SQL.Add('    PAGAR.OBSERVACAO AS DES_OBSERVACAO,');
    SQL.Add('    COALESCE(PAGAR.PDV, 0) AS NUM_PDV,');
    SQL.Add('    PAGAR.NOTA AS NUM_CUPOM_FISCAL,');
    SQL.Add('    0 AS COD_MOTIVO,');
    SQL.Add('    0 AS COD_CONVENIO,');
    SQL.Add('    0 AS COD_BIN,');
    SQL.Add('    '''' AS DES_BANDEIRA,');
    SQL.Add('    '''' AS DES_REDE_TEF,');
    SQL.Add('    0 AS VAL_RETENCAO,');
    SQL.Add('    0 AS COD_CONDICAO,');
    SQL.Add('');
    SQL.Add('    CASE ');
    SQL.Add('        WHEN PAGAR.PAGAMENTO IS NULL THEN NULL');
    SQL.Add('        ELSE PAGAR.PAGAMENTO');
    SQL.Add('    END AS DTA_PAGTO,');
    SQL.Add('');
    SQL.Add('    PAGAR.DATAHORA_CADASTRO AS DTA_ENTRADA,');
    SQL.Add('');
    SQL.Add('    '''' AS NUM_NOSSO_NUMERO,');
    SQL.Add('    COALESCE(PAGAR.CODBARRAS, '''') AS COD_BARRA,');
    SQL.Add('    ''N'' AS FLG_BOLETO_EMIT,');
    SQL.Add('    '''' AS NUM_CGC_CPF_TITULAR,');
    SQL.Add('    '''' AS DES_TITULAR,');
    SQL.Add('    30 AS NUM_CONDICAO,');
    SQL.Add('    0 AS VAL_CREDITO,');
    SQL.Add('    ''999'' AS COD_BANCO_PGTO,');
    SQL.Add('    ''PAGTO-1'' AS DES_CC,');
    SQL.Add('    0 AS COD_BANDEIRA,');
    SQL.Add('    '''' AS DTA_PRORROGACAO,');
    SQL.Add('    1 AS NUM_SEQ_FIN');
    SQL.Add('FROM');
    SQL.Add('    CONTAS PAGAR');
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
    SQL.Add('            CONTAS.TIPO_CONTA = 0');
    SQL.Add('        AND');
    SQL.Add('            CONTAS.EMPRESA = '+ CbxLoja.Text +'');
    SQL.Add('        AND');
    SQL.Add('            CONTAS.PARCELA > 0');
    SQL.Add('        AND');
    SQL.Add('            CONTAS.TIPO_CADASTRO IN (2, 3)');
    SQL.Add('        GROUP BY');
    SQL.Add('            NF,');
    SQL.Add('            TIPO_CADASTRO,');
    SQL.Add('            ID_CADASTRO');
    SQL.Add('    ) NF');
    SQL.Add('ON');
    SQL.Add('    PAGAR.NF = NF.NF');
    SQL.Add('AND');
    SQL.Add('    PAGAR.TIPO_CADASTRO = NF.TIPO_CADASTRO');
    SQL.Add('AND');
    SQL.Add('    PAGAR.ID_CADASTRO = NF.ID_CADASTRO        ');
    SQL.Add('WHERE');
    SQL.Add('    PAGAR.TIPO_CONTA = 0');
    SQL.Add('AND');
    SQL.Add('    PAGAR.TIPO_CADASTRO IN (2, 3)');
    SQL.Add('AND');
    SQL.Add('    PAGAR.PARCELA > 0 ');



    SQL.Add('AND');
    SQL.Add('    PAGAR.EMPRESA =  '+ CbxLoja.Text +' ');


    if Aberto = '1' then
    begin
        SQL.Add('AND');
        SQL.Add('    PAGAR.PAGAMENTO IS NULL');
    end
    else
    begin
        SQL.Add('AND');
        SQL.Add('    PAGAR.PAGAMENTO IS NOT NULL');
        SQL.Add('AND');
        SQL.Add('    PAGAR.PAGAMENTO >= '''+FormatDateTime('dd/mm/yyyy',DtpInicial.Date)+''' ');
        SQL.Add('AND');
        SQL.Add('    PAGAR.PAGAMENTO <= '''+FormatDateTime('dd/mm/yyyy',DtpFinal.Date)+''' ');
    end;

//    ShowMessage(sql.Text);

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

procedure TFrmGetWayMama.GerarFinanceiroReceber(Aberto: String);
var
   codParceiro : Integer;
   numDocto : String;
   count : integer;
begin
  inherited;
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
    SQL.Add('');
    SQL.Add('    CASE RECEBER.TIPO_CADASTRO');
    SQL.Add('        WHEN 0 THEN 0');
    SQL.Add('        WHEN 1 THEN 3');
    SQL.Add('        WHEN 2 THEN 1');
    SQL.Add('        WHEN 4 THEN 4');
    SQL.Add('        WHEN 5 THEN 0');
    SQL.Add('    END AS TIPO_PARCEIRO, -- TIPO_CADASTRO');
    SQL.Add('');
    SQL.Add('     CASE');
    SQL.Add('          WHEN RECEBER.TIPO_CADASTRO = 5 THEN 2400 + RECEBER.ID_CADASTRO ');
    SQL.Add('          WHEN RECEBER.TIPO_CADASTRO = 5 AND COALESCE(RECEBER.ID_CADASTRO, 0) = 0 THEN 6');
    SQL.Add('          WHEN RECEBER.TIPO_CADASTRO = 4 THEN 99');
    SQL.Add('          ELSE CASE WHEN COALESCE(RECEBER.ID_CADASTRO, 0) = 0 THEN 99999 ELSE RECEBER.ID_CADASTRO END');
    SQL.Add('     END AS COD_PARCEIRO,  ');
    SQL.Add('');
    SQL.Add('    1 AS TIPO_CONTA,');
    SQL.Add('');
    SQL.Add('    CASE RECEBER.FORMA_PAGTO');
    SQL.Add('        WHEN 0 THEN 1');
    SQL.Add('        WHEN 1 THEN 1');
    SQL.Add('        WHEN 2 THEN 2');
    SQL.Add('        WHEN 3 THEN 4');
    SQL.Add('        WHEN 4 THEN 10');
    SQL.Add('        WHEN 5 THEN 11');
    SQL.Add('        WHEN 6 THEN 6');
    SQL.Add('        WHEN 7 THEN 12');
    SQL.Add('        WHEN 8 THEN 3');
    SQL.Add('        WHEN 9 THEN 13');
    SQL.Add('        WHEN 10 THEN 5');
    SQL.Add('        WHEN 11 THEN 7');
    SQL.Add('        WHEN 12 THEN 14');
    SQL.Add('        WHEN 13 THEN 15');
    SQL.Add('        WHEN 14 THEN 16');
    SQL.Add('        WHEN 15 THEN 17');
    SQL.Add('        WHEN 16 THEN 18');
    SQL.Add('        WHEN 17 THEN 19');
    SQL.Add('        WHEN 18 THEN 20');
    SQL.Add('        WHEN 19 THEN 21');
    SQL.Add('        WHEN 20 THEN 22');
    SQL.Add('        WHEN 21 THEN 23');
    SQL.Add('        WHEN 22 THEN 24');
    SQL.Add('        WHEN 23 THEN 25');
    SQL.Add('        WHEN 24 THEN 26');
    SQL.Add('        WHEN 25 THEN 27      ');
    SQL.Add('        ELSE 1      ');
    SQL.Add('    END AS COD_ENTIDADE,');
    SQL.Add('');
    SQL.Add('    RECEBER.ARQUIVO AS NUM_DOCTO,');
    SQL.Add('    999 AS COD_BANCO,');
    SQL.Add('    '''' AS DES_BANCO,');
    SQL.Add('    RECEBER.EMISSAO AS DTA_EMISSAO,');
    SQL.Add('    RECEBER.VENCIMENTO AS DTA_VENCIMENTO,');
    SQL.Add('    RECEBER.VALOR AS VAL_PARCELA,');
    SQL.Add('    RECEBER.ACRESCIMO + RECEBER.CARTORIO + COALESCE(RECEBER.CREDITO, 0) AS VAL_JUROS,');
    SQL.Add('    RECEBER.DESCONTO AS VAL_DESCONTO,');
    SQL.Add('');
    SQL.Add('    CASE ');
    SQL.Add('        WHEN RECEBER.PAGAMENTO IS NULL THEN ''N''');
    SQL.Add('        ELSE ''S''');
    SQL.Add('    END AS FLG_QUITADO,');
    SQL.Add('');
    SQL.Add('    CASE ');
    SQL.Add('        WHEN RECEBER.PAGAMENTO IS NULL THEN NULL');
    SQL.Add('        ELSE RECEBER.PAGAMENTO');
    SQL.Add('    END AS DTA_QUITADA,');
    SQL.Add('');
    SQL.Add('    ');
    SQL.Add('    CASE RECEBER.CAIXA');
    SQL.Add('        WHEN 2 THEN ''001''');
    SQL.Add('        ELSE ''997''');
    SQL.Add('    END AS COD_CATEGORIA,');
    SQL.Add('');
    SQL.Add('    CASE RECEBER.CAIXA');
    SQL.Add('        WHEN 2 THEN ''032''');
    SQL.Add('        ELSE ''997''');
    SQL.Add('    END AS COD_SUBCATEGORIA,');
    SQL.Add('');
    SQL.Add('    RECEBER.PARCELA AS NUM_PARCELA,');
    SQL.Add('    RECEBER.TOTAL_PARCELA AS QTD_PARCELA,');
    SQL.Add('    RECEBER.EMPRESA AS COD_LOJA,');
    SQL.Add('    RECEBER.CPF_CNPJ AS NUM_CGC,');
    SQL.Add('    COALESCE(RECEBER.BORDERO, 0) AS NUM_BORDERO,');
    SQL.Add('    RECEBER.NF AS NUM_NF,');
    SQL.Add('    '''' AS NUM_SERIE_NF,');
    SQL.Add('    CASE WHEN NF.VAL_TOTAL_NF = 0 THEN RECEBER.VALOR ELSE NF.VAL_TOTAL_NF END AS VAL_TOTAL_NF, -- EFETUAR A SOMA');
    SQL.Add('    ''COBRANÇA: '' || RECEBER.DATACOB || '' | 1 DEVOL: '' || RECEBER.DEVOLUCAOA || '' | 2 DEVOL : '' || RECEBER.DEVOLUCAOB || '' | ''  || RECEBER.OBSERVACAO AS DES_OBSERVACAO,');
    SQL.Add('    COALESCE(RECEBER.PDV, 0) AS NUM_PDV,');
    SQL.Add('    RECEBER.NOTA AS NUM_CUPOM_FISCAL,');
    SQL.Add('    0 AS COD_MOTIVO,');
    SQL.Add('');
    SQL.Add('    CASE RECEBER.FORMA_PAGTO');
    SQL.Add('        WHEN 14 THEN (SELECT COALESCE(24000 + CLIENTES.EMPRESA_CONVENIO, 0) FROM CLIENTES WHERE CLIENTES.ID = RECEBER.ID_CADASTRO)');
    SQL.Add('        ELSE 0');
    SQL.Add('    END AS COD_CONVENIO,');
    SQL.Add('');
    SQL.Add('    0 AS COD_BIN,');
    SQL.Add('    '''' AS DES_BANDEIRA,');
    SQL.Add('    '''' AS DES_REDE_TEF,');
    SQL.Add('    0 AS VAL_RETENCAO,');
    SQL.Add('    0 AS COD_CONDICAO,');
    SQL.Add('');
    SQL.Add('    CASE ');
    SQL.Add('        WHEN RECEBER.PAGAMENTO IS NULL THEN NULL');
    SQL.Add('        ELSE RECEBER.PAGAMENTO');
    SQL.Add('    END AS DTA_PAGTO,');
    SQL.Add('');
    SQL.Add('    RECEBER.DATAHORA_CADASTRO AS DTA_ENTRADA,');
    SQL.Add('');
    SQL.Add('    '''' AS NUM_NOSSO_NUMERO,');
    SQL.Add('    COALESCE(RECEBER.CODBARRAS, '''') AS COD_BARRA,');
    SQL.Add('    ''N'' AS FLG_BOLETO_EMIT,');
    SQL.Add('    '''' AS NUM_CGC_CPF_TITULAR,');
    SQL.Add('    '''' AS DES_TITULAR,');
    SQL.Add('    30 AS NUM_CONDICAO,');
    SQL.Add('    0 AS VAL_CREDITO,');
    SQL.Add('    ''999'' AS COD_BANCO_PGTO,');
    SQL.Add('    ''RECEBTO-1'' AS DES_CC,');

    SQL.Add('    CASE ');
    SQL.Add('        WHEN RECEBER.TIPO_CADASTRO = 4 THEN CASE WHEN RECEBER.EMPRESA = 1 THEN 9999 ELSE 999 END');
    SQL.Add('        ELSE 0');
    SQL.Add('        END AS COD_BANDEIRA,');


    SQL.Add('    '''' AS DTA_PRORROGACAO,');
    SQL.Add('    1 AS NUM_SEQ_FIN,');
    SQL.Add('    CASE RECEBER.COBRADOR');
    SQL.Add('        WHEN 1 THEN 9001');
    SQL.Add('        WHEN 2 THEN 9002');
    SQL.Add('        WHEN 3 THEN 9003');
    SQL.Add('        ELSE 0');
    SQL.Add('    END AS COD_COBRANCA,');
    SQL.Add('    RECEBER.DATACOB AS DTA_COBRANCA,');
    SQL.Add('    CASE');
    SQL.Add('        WHEN LENGTH(RECEBER.CODBARRAS) > 0 THEN ''S''');
    SQL.Add('        ELSE ''N''');
    SQL.Add('    END AS FLG_ACEITE,');
    SQL.Add('    CASE');
    SQL.Add('        WHEN LENGTH(RECEBER.CODBARRAS) = 34 THEN 4 ');
    SQL.Add('        WHEN LENGTH(RECEBER.CODBARRAS) > 34 THEN 1 ');
    SQL.Add('        ELSE 0');
    SQL.Add('    END AS TIPO_ACEITE');

    SQL.Add('FROM');
    SQL.Add('    CONTAS RECEBER');
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
    SQL.Add('            CONTAS.TIPO_CADASTRO IN (0, 1, 2, 5) -- Adicionar o filtro de cartoes');
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
    SQL.Add('    RECEBER.ID_CADASTRO = NF.ID_CADASTRO');
    SQL.Add('WHERE');
    SQL.Add('    RECEBER.TIPO_CONTA = 1');
    SQL.Add('AND');
    SQL.Add('    RECEBER.TIPO_CADASTRO IN (0, 1, 2, 5) -- Adicionar o filtro de cartoes');
    SQL.Add('AND');
    SQL.Add('    RECEBER.PARCELA > 0');
    SQL.Add('AND');
    SQL.Add('    RECEBER.VALOR > 0');


    SQL.Add('AND');
    SQL.Add('    RECEBER.EMPRESA = '+ CbxLoja.Text +' ');

    if Aberto = '1' then
    begin
        SQL.Add('AND');
        SQL.Add('    RECEBER.PAGAMENTO IS NULL');
    end
    else
    begin
        SQL.Add('AND');
        SQL.Add('    RECEBER.PAGAMENTO IS NOT NULL');
        SQL.Add('AND');
        SQL.Add('    RECEBER.PAGAMENTO >= '''+FormatDateTime('dd/mm/yyyy',DtpInicial.Date)+''' ');
        SQL.Add('AND');
        SQL.Add('    RECEBER.PAGAMENTO <= '''+FormatDateTime('dd/mm/yyyy',DtpFinal.Date)+''' ');
    end;

    SQL.Add('ORDER BY');
    SQL.Add('    NUM_DOCTO, COD_PARCEIRO');

    Open;

    First;
    NumLinha := 0;
    codParceiro := 0;
    numDocto := '';
    count := 0;

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);

      Layout.SetValues(QryPrincipal2, NumLinha, RecordCount);

      if( (codParceiro = QryPrincipal2.FieldByName('COD_PARCEIRO').AsInteger) and (numDocto = QryPrincipal2.FieldByName('NUM_DOCTO').AsString) ) then
      begin
         inc(count);
         if( numDocto <> '' ) then
            Layout.FieldByName('NUM_DOCTO').AsString := numDocto + ' - ' + IntToStr(count)
         else
            Layout.FieldByName('NUM_DOCTO').AsString := IntToStr(count);
      end
      else
      begin
         count := 0;
         numDocto := QryPrincipal2.FieldByName('NUM_DOCTO').AsString;
         codParceiro := QryPrincipal2.FieldByName('COD_PARCEIRO').AsInteger;
      end;

      Layout.FieldByName('DTA_ENTRADA').AsString:= FormatDateTime('dd/mm/yyyy',QryPrincipal2.FieldByName('DTA_ENTRADA').AsDateTime);
      Layout.FieldByName('DTA_EMISSAO').AsString:= FormatDateTime('dd/mm/yyyy',QryPrincipal2.FieldByName('DTA_EMISSAO').AsDateTime);
      Layout.FieldByName('DTA_VENCIMENTO').AsString:= FormatDateTime('dd/mm/yyyy',QryPrincipal2.FieldByName('DTA_VENCIMENTO').AsDateTime);

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

procedure TFrmGetWayMama.GerarFinanceiroReceberCartao;
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

procedure TFrmGetWayMama.GerarFornecedor;
var
   observacao, email : string;
begin
  inherited;

  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
    SQL.Add('     FORNECEDORES.CODFORNEC AS COD_FORNECEDOR,');
    SQL.Add('     FORNECEDORES.RAZAO AS DES_FORNECEDOR,');
    SQL.Add('     FORNECEDORES.FANTASIA AS DES_FANTASIA,');
    SQL.Add('     FORNECEDORES.CNPJ_CPF AS NUM_CGC,');
    SQL.Add('     FORNECEDORES.IE AS NUM_INSC_EST,');
    SQL.Add('     FORNECEDORES.ENDERECO AS DES_ENDERECO,');
    SQL.Add('     FORNECEDORES.BAIRRO AS DES_BAIRRO,');
    SQL.Add('     FORNECEDORES.CIDADE AS DES_CIDADE,');
    SQL.Add('     FORNECEDORES.ESTADO AS DES_SIGLA,');
    SQL.Add('     FORNECEDORES.CEP AS NUM_CEP,');
    SQL.Add('     FORNECEDORES.TELEFONE AS NUM_FONE,');
    SQL.Add('     FORNECEDORES.FAX AS NUM_FAX,');
    SQL.Add('     FORNECEDORES.CONTATO AS DES_CONTATO,');
    SQL.Add('     0 AS QTD_DIA_CARENCIA,');
    SQL.Add('     FORNECEDORES.PVISITA AS NUM_FREQ_VISITA,');
    SQL.Add('     0 AS VAL_DESCONTO,');
    SQL.Add('     FORNECEDORES.PENTREGA AS NUM_PRAZO,');
    SQL.Add('     ''N'' AS ACEITA_DEVOL_MER,');
    SQL.Add('     ''N'' AS CAL_IPI_VAL_BRUTO,');
    SQL.Add('     ''N'' AS CAL_ICMS_ENC_FIN,');
    SQL.Add('     ''N'' AS CAL_ICMS_VAL_IPI,');
    SQL.Add('     ''N'' AS MICRO_EMPRESA,');
    SQL.Add('     0 AS COD_FORNECEDOR_ANT,');
    SQL.Add('     FORNECEDORES.NUMERO AS NUM_ENDERECO,');
    SQL.Add('     FORNECEDORES.OBS AS DES_OBSERVACAO,');
    SQL.Add('     FORNECEDORES.EMAIL AS DES_EMAIL,');
    SQL.Add('     '''' AS DES_WEB_SITE,');
    SQL.Add('     ''N'' AS FABRICANTE,');
    SQL.Add('     ''N'' AS FLG_PRODUTOR_RURAL,');
    SQL.Add('     0 AS TIPO_FRETE,');
    SQL.Add('     FORNECEDORES.SIMPLES AS FLG_SIMPLES,');
    SQL.Add('     ''N'' AS FLG_SUBSTITUTO_TRIB,');
    SQL.Add('     0 AS COD_CONTACCFORN,');
    SQL.Add('');
    SQL.Add('     CASE');
    SQL.Add('          WHEN FORNECEDORES.ATIVO = ''N'' THEN ''S''');
    SQL.Add('          ELSE ''N''');
    SQL.Add('     END AS INATIVO,');
    SQL.Add('');
    SQL.Add('     0 AS COD_CLASSIF,');
    SQL.Add('     '''' AS DTA_CADASTRO,');
    SQL.Add('     0 AS VAL_CREDITO,');
    SQL.Add('     0 AS VAL_DEBITO,');
    SQL.Add('     1 AS PED_MIN_VAL,');
    SQL.Add('     '''' AS DES_EMAIL_VEND,');
    SQL.Add('     '''' AS SENHA_COTACAO,');
    SQL.Add('     -1 AS TIPO_PRODUTOR,');
    SQL.Add('     FORNECEDORES.CELULAR AS NUM_CELULAR');
    SQL.Add('FROM');
    SQL.Add('     FORNECEDORES');


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
//      Layout.FieldByName('NUM_INSC_EST').AsString := StrRetNums(Layout.FieldByName('NUM_INSC_EST').AsString);
      Layout.FieldByName('NUM_CEP').AsString := StrRetNums(Layout.FieldByName('NUM_CEP').AsString);

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

procedure TFrmGetWayMama.GerarGrupo;
begin
  inherited;
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT DISTINCT');
    SQL.Add('    CODCRECEITA AS COD_SECAO,');
    SQL.Add('    CODGRUPO AS COD_GRUPO,');
    SQL.Add('    ''A DEFINIR'' AS DES_GRUPO,');
    SQL.Add('    0 AS VAL_META');
    SQL.Add('FROM');
    SQL.Add('    PRODUTOS');


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

procedure TFrmGetWayMama.GerarInfoNutricionais;
begin
  inherited;

  with QryPrincipal2 do
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

procedure TFrmGetWayMama.GerarNCM;
var
 count : Integer;
begin
  inherited;
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT DISTINCT');
    SQL.Add('    0 AS COD_NCM,');
    SQL.Add('    ''A DEFINIR'' AS DES_NCM,');
    SQL.Add('');
    SQL.Add('    CASE ');
    SQL.Add('        WHEN COALESCE(CODNCM, ''00000000'') = ''00000000'' THEN ''99999999''');
    SQL.Add('        ELSE REPLACE(CODNCM, ''.'', '''')');
    SQL.Add('    END AS NUM_NCM,');
    SQL.Add('');
    SQL.Add('    CASE COALESCE(CST_PISSAIDA, ''01'')');
    SQL.Add('        WHEN ''01'' THEN ''S''');
    SQL.Add('        WHEN ''04'' THEN ''N''');
    SQL.Add('        WHEN ''05'' THEN ''N''');
    SQL.Add('        WHEN ''06'' THEN ''N''');
    SQL.Add('        WHEN ''07'' THEN ''S''');
    SQL.Add('        WHEN ''09'' THEN ''N''');
    SQL.Add('        WHEN ''49'' THEN ''S''');
    SQL.Add('    END AS FLG_NAO_PIS_COFINS, ');
    SQL.Add('');
    SQL.Add('    CASE COALESCE(CST_PISSAIDA, ''01'')');
    SQL.Add('        WHEN ''01'' THEN -1');
    SQL.Add('        WHEN ''04'' THEN 1');
    SQL.Add('        WHEN ''05'' THEN 2');
    SQL.Add('        WHEN ''06'' THEN 0');
    SQL.Add('        WHEN ''07'' THEN -1');
    SQL.Add('        WHEN ''09'' THEN 4');
    SQL.Add('        WHEN ''49'' THEN -1');
    SQL.Add('    END AS TIPO_NAO_PIS_COFINS,');
    SQL.Add('');
    SQL.Add('    999 AS COD_TAB_SPED,');
    SQL.Add('    0 AS NUM_CEST,');
    SQL.Add('    ''SP'' AS DES_SIGLA,');
    SQL.Add('');
    SQL.Add('    CASE');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''00'' AND TRIBUTACAO.ALIQUOTA = 4.50 AND TRIBUTACAO.REDUCAO = 0.00 THEN 70  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''00'' AND TRIBUTACAO.ALIQUOTA = 7.00 AND TRIBUTACAO.REDUCAO = 0.00 THEN 2  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''00'' AND TRIBUTACAO.ALIQUOTA = 11.00 AND TRIBUTACAO.REDUCAO = 0.00 THEN 66  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''00'' AND TRIBUTACAO.ALIQUOTA = 12.00 AND TRIBUTACAO.REDUCAO = 0.00 THEN 3  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''00'' AND TRIBUTACAO.ALIQUOTA = 18.00 AND TRIBUTACAO.REDUCAO = 0.00 THEN 4  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''00'' AND TRIBUTACAO.ALIQUOTA = 18.00 AND TRIBUTACAO.REDUCAO = 33.33 THEN 7  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''00'' AND TRIBUTACAO.ALIQUOTA = 18.00 AND TRIBUTACAO.REDUCAO = 61.11 THEN 8  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''00'' AND TRIBUTACAO.ALIQUOTA = 25.00 AND TRIBUTACAO.REDUCAO = 0.00 THEN 5  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''20'' AND TRIBUTACAO.ALIQUOTA = 7.00 AND TRIBUTACAO.REDUCAO = 41.67 THEN 6  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''20'' AND TRIBUTACAO.ALIQUOTA = 7.00 AND TRIBUTACAO.REDUCAO = 61.11 THEN 8  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''20'' AND TRIBUTACAO.ALIQUOTA = 18.00 AND TRIBUTACAO.REDUCAO = 0.00 THEN 12  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''20'' AND TRIBUTACAO.ALIQUOTA = 18.00 AND TRIBUTACAO.REDUCAO = 33.33 THEN 15  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''20'' AND TRIBUTACAO.ALIQUOTA = 18.00 AND TRIBUTACAO.REDUCAO = 41.67 THEN 14  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''20'' AND TRIBUTACAO.ALIQUOTA = 18.00 AND TRIBUTACAO.REDUCAO = 61.11 THEN 16  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''40'' AND TRIBUTACAO.ALIQUOTA = 0.00 AND TRIBUTACAO.REDUCAO = 0.00 THEN 1  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''40'' AND TRIBUTACAO.ALIQUOTA = 0.00 AND TRIBUTACAO.REDUCAO = 4.00 THEN 1  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''40'' AND TRIBUTACAO.ALIQUOTA = 0.00 AND TRIBUTACAO.REDUCAO = 41.67 THEN 1  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''60'' AND TRIBUTACAO.ALIQUOTA = 0.00 AND TRIBUTACAO.REDUCAO = 0.00 THEN 12  ');
    SQL.Add('    END AS COD_TRIB_ENTRADA,');
    SQL.Add('');
    SQL.Add('    CASE');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''00'' AND TRIBUTACAO.ALIQUOTA = 4.50 AND TRIBUTACAO.REDUCAO = 0.00 THEN 70  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''00'' AND TRIBUTACAO.ALIQUOTA = 7.00 AND TRIBUTACAO.REDUCAO = 0.00 THEN 2  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''00'' AND TRIBUTACAO.ALIQUOTA = 11.00 AND TRIBUTACAO.REDUCAO = 0.00 THEN 66  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''00'' AND TRIBUTACAO.ALIQUOTA = 12.00 AND TRIBUTACAO.REDUCAO = 0.00 THEN 3  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''00'' AND TRIBUTACAO.ALIQUOTA = 18.00 AND TRIBUTACAO.REDUCAO = 0.00 THEN 4  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''00'' AND TRIBUTACAO.ALIQUOTA = 18.00 AND TRIBUTACAO.REDUCAO = 33.33 THEN 7  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''00'' AND TRIBUTACAO.ALIQUOTA = 18.00 AND TRIBUTACAO.REDUCAO = 61.11 THEN 8  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''00'' AND TRIBUTACAO.ALIQUOTA = 25.00 AND TRIBUTACAO.REDUCAO = 0.00 THEN 5  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''20'' AND TRIBUTACAO.ALIQUOTA = 7.00 AND TRIBUTACAO.REDUCAO = 41.67 THEN 6  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''20'' AND TRIBUTACAO.ALIQUOTA = 7.00 AND TRIBUTACAO.REDUCAO = 61.11 THEN 8  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''20'' AND TRIBUTACAO.ALIQUOTA = 18.00 AND TRIBUTACAO.REDUCAO = 0.00 THEN 12  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''20'' AND TRIBUTACAO.ALIQUOTA = 18.00 AND TRIBUTACAO.REDUCAO = 33.33 THEN 15  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''20'' AND TRIBUTACAO.ALIQUOTA = 18.00 AND TRIBUTACAO.REDUCAO = 41.67 THEN 14  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''20'' AND TRIBUTACAO.ALIQUOTA = 18.00 AND TRIBUTACAO.REDUCAO = 61.11 THEN 16  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''40'' AND TRIBUTACAO.ALIQUOTA = 0.00 AND TRIBUTACAO.REDUCAO = 0.00 THEN 1  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''40'' AND TRIBUTACAO.ALIQUOTA = 0.00 AND TRIBUTACAO.REDUCAO = 4.00 THEN 1  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''40'' AND TRIBUTACAO.ALIQUOTA = 0.00 AND TRIBUTACAO.REDUCAO = 41.67 THEN 1  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''60'' AND TRIBUTACAO.ALIQUOTA = 0.00 AND TRIBUTACAO.REDUCAO = 0.00 THEN 12  ');
    SQL.Add('    END AS COD_TRIB_SAIDA,');
    SQL.Add('');
    SQL.Add('    0 AS PER_IVA,');
    SQL.Add('    0 AS PER_FCP_ST   ');
    SQL.Add('FROM');
    SQL.Add('    PRODUTOS ');
    SQL.Add('LEFT JOIN');
    SQL.Add('    (');
    SQL.Add('        SELECT ');
    SQL.Add('            PRODUTOS.CODPROD,');
    SQL.Add('            SUBSTRING(CODTRIB, 2,3) AS CST, ');
    SQL.Add('            COALESCE(ALIQUOTA_ICMS.VALORTRIB, 0) AS ALIQUOTA,');
    SQL.Add('            PER_REDUC AS REDUCAO');
    SQL.Add('        FROM');
    SQL.Add('            PRODUTOS');
    SQL.Add('        LEFT JOIN ALIQUOTA_ICMS');
    SQL.Add('        ON PRODUTOS.CODALIQ = ALIQUOTA_ICMS.CODALIQ');
    SQL.Add('    ) AS TRIBUTACAO');
    SQL.Add('ON');
    SQL.Add('    PRODUTOS.CODPROD = TRIBUTACAO.CODPROD    ');
    SQL.Add('ORDER BY');
    SQL.Add('    NUM_NCM,');
    SQL.Add('    FLG_NAO_PIS_COFINS,');
    SQL.Add('    TIPO_NAO_PIS_COFINS,');
    SQL.Add('    COD_TRIB_ENTRADA,');
    SQL.Add('    COD_TRIB_SAIDA');



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

procedure TFrmGetWayMama.GerarNCMUF;
var
 count : Integer;
begin
  inherited;
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT DISTINCT');
    SQL.Add('    0 AS COD_NCM,');
    SQL.Add('    ''A DEFINIR'' AS DES_NCM,');
    SQL.Add('');
    SQL.Add('    CASE ');
    SQL.Add('        WHEN COALESCE(CODNCM, ''00000000'') = ''00000000'' THEN ''99999999''');
    SQL.Add('        ELSE REPLACE(CODNCM, ''.'', '''')');
    SQL.Add('    END AS NUM_NCM,');
    SQL.Add('');
    SQL.Add('    CASE COALESCE(CST_PISSAIDA, ''01'')');
    SQL.Add('        WHEN ''01'' THEN ''S''');
    SQL.Add('        WHEN ''04'' THEN ''N''');
    SQL.Add('        WHEN ''05'' THEN ''N''');
    SQL.Add('        WHEN ''06'' THEN ''N''');
    SQL.Add('        WHEN ''07'' THEN ''S''');
    SQL.Add('        WHEN ''09'' THEN ''N''');
    SQL.Add('        WHEN ''49'' THEN ''S''');
    SQL.Add('    END AS FLG_NAO_PIS_COFINS, ');
    SQL.Add('');
    SQL.Add('    CASE COALESCE(CST_PISSAIDA, ''01'')');
    SQL.Add('        WHEN ''01'' THEN -1');
    SQL.Add('        WHEN ''04'' THEN 1');
    SQL.Add('        WHEN ''05'' THEN 2');
    SQL.Add('        WHEN ''06'' THEN 0');
    SQL.Add('        WHEN ''07'' THEN -1');
    SQL.Add('        WHEN ''09'' THEN 4');
    SQL.Add('        WHEN ''49'' THEN -1');
    SQL.Add('    END AS TIPO_NAO_PIS_COFINS,');
    SQL.Add('');
    SQL.Add('    999 AS COD_TAB_SPED,');
    SQL.Add('    0 AS NUM_CEST,');
    SQL.Add('    ''SP'' AS DES_SIGLA,');
    SQL.Add('');
    SQL.Add('    CASE');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''00'' AND TRIBUTACAO.ALIQUOTA = 4.50 AND TRIBUTACAO.REDUCAO = 0.00 THEN 70  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''00'' AND TRIBUTACAO.ALIQUOTA = 7.00 AND TRIBUTACAO.REDUCAO = 0.00 THEN 2  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''00'' AND TRIBUTACAO.ALIQUOTA = 11.00 AND TRIBUTACAO.REDUCAO = 0.00 THEN 66  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''00'' AND TRIBUTACAO.ALIQUOTA = 12.00 AND TRIBUTACAO.REDUCAO = 0.00 THEN 3  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''00'' AND TRIBUTACAO.ALIQUOTA = 18.00 AND TRIBUTACAO.REDUCAO = 0.00 THEN 4  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''00'' AND TRIBUTACAO.ALIQUOTA = 18.00 AND TRIBUTACAO.REDUCAO = 33.33 THEN 7  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''00'' AND TRIBUTACAO.ALIQUOTA = 18.00 AND TRIBUTACAO.REDUCAO = 61.11 THEN 8  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''00'' AND TRIBUTACAO.ALIQUOTA = 25.00 AND TRIBUTACAO.REDUCAO = 0.00 THEN 5  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''20'' AND TRIBUTACAO.ALIQUOTA = 7.00 AND TRIBUTACAO.REDUCAO = 41.67 THEN 6  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''20'' AND TRIBUTACAO.ALIQUOTA = 7.00 AND TRIBUTACAO.REDUCAO = 61.11 THEN 8  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''20'' AND TRIBUTACAO.ALIQUOTA = 18.00 AND TRIBUTACAO.REDUCAO = 0.00 THEN 12  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''20'' AND TRIBUTACAO.ALIQUOTA = 18.00 AND TRIBUTACAO.REDUCAO = 33.33 THEN 15  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''20'' AND TRIBUTACAO.ALIQUOTA = 18.00 AND TRIBUTACAO.REDUCAO = 41.67 THEN 14  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''20'' AND TRIBUTACAO.ALIQUOTA = 18.00 AND TRIBUTACAO.REDUCAO = 61.11 THEN 16  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''40'' AND TRIBUTACAO.ALIQUOTA = 0.00 AND TRIBUTACAO.REDUCAO = 0.00 THEN 1  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''40'' AND TRIBUTACAO.ALIQUOTA = 0.00 AND TRIBUTACAO.REDUCAO = 4.00 THEN 1  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''40'' AND TRIBUTACAO.ALIQUOTA = 0.00 AND TRIBUTACAO.REDUCAO = 41.67 THEN 1  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''60'' AND TRIBUTACAO.ALIQUOTA = 0.00 AND TRIBUTACAO.REDUCAO = 0.00 THEN 12  ');
    SQL.Add('    END AS COD_TRIB_ENTRADA,');
    SQL.Add('');
    SQL.Add('    CASE');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''00'' AND TRIBUTACAO.ALIQUOTA = 4.50 AND TRIBUTACAO.REDUCAO = 0.00 THEN 70  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''00'' AND TRIBUTACAO.ALIQUOTA = 7.00 AND TRIBUTACAO.REDUCAO = 0.00 THEN 2  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''00'' AND TRIBUTACAO.ALIQUOTA = 11.00 AND TRIBUTACAO.REDUCAO = 0.00 THEN 66  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''00'' AND TRIBUTACAO.ALIQUOTA = 12.00 AND TRIBUTACAO.REDUCAO = 0.00 THEN 3  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''00'' AND TRIBUTACAO.ALIQUOTA = 18.00 AND TRIBUTACAO.REDUCAO = 0.00 THEN 4  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''00'' AND TRIBUTACAO.ALIQUOTA = 18.00 AND TRIBUTACAO.REDUCAO = 33.33 THEN 7  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''00'' AND TRIBUTACAO.ALIQUOTA = 18.00 AND TRIBUTACAO.REDUCAO = 61.11 THEN 8  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''00'' AND TRIBUTACAO.ALIQUOTA = 25.00 AND TRIBUTACAO.REDUCAO = 0.00 THEN 5  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''20'' AND TRIBUTACAO.ALIQUOTA = 7.00 AND TRIBUTACAO.REDUCAO = 41.67 THEN 6  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''20'' AND TRIBUTACAO.ALIQUOTA = 7.00 AND TRIBUTACAO.REDUCAO = 61.11 THEN 8  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''20'' AND TRIBUTACAO.ALIQUOTA = 18.00 AND TRIBUTACAO.REDUCAO = 0.00 THEN 12  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''20'' AND TRIBUTACAO.ALIQUOTA = 18.00 AND TRIBUTACAO.REDUCAO = 33.33 THEN 15  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''20'' AND TRIBUTACAO.ALIQUOTA = 18.00 AND TRIBUTACAO.REDUCAO = 41.67 THEN 14  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''20'' AND TRIBUTACAO.ALIQUOTA = 18.00 AND TRIBUTACAO.REDUCAO = 61.11 THEN 16  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''40'' AND TRIBUTACAO.ALIQUOTA = 0.00 AND TRIBUTACAO.REDUCAO = 0.00 THEN 1  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''40'' AND TRIBUTACAO.ALIQUOTA = 0.00 AND TRIBUTACAO.REDUCAO = 4.00 THEN 1  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''40'' AND TRIBUTACAO.ALIQUOTA = 0.00 AND TRIBUTACAO.REDUCAO = 41.67 THEN 1  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''60'' AND TRIBUTACAO.ALIQUOTA = 0.00 AND TRIBUTACAO.REDUCAO = 0.00 THEN 12  ');
    SQL.Add('    END AS COD_TRIB_SAIDA,');
    SQL.Add('');
    SQL.Add('    0 AS PER_IVA,');
    SQL.Add('    0 AS PER_FCP_ST   ');
    SQL.Add('FROM');
    SQL.Add('    PRODUTOS ');
    SQL.Add('LEFT JOIN');
    SQL.Add('    (');
    SQL.Add('        SELECT ');
    SQL.Add('            PRODUTOS.CODPROD,');
    SQL.Add('            SUBSTRING(CODTRIB, 2,3) AS CST, ');
    SQL.Add('            COALESCE(ALIQUOTA_ICMS.VALORTRIB, 0) AS ALIQUOTA,');
    SQL.Add('            PER_REDUC AS REDUCAO');
    SQL.Add('        FROM');
    SQL.Add('            PRODUTOS');
    SQL.Add('        LEFT JOIN ALIQUOTA_ICMS');
    SQL.Add('        ON PRODUTOS.CODALIQ = ALIQUOTA_ICMS.CODALIQ');
    SQL.Add('    ) AS TRIBUTACAO');
    SQL.Add('ON');
    SQL.Add('    PRODUTOS.CODPROD = TRIBUTACAO.CODPROD    ');
    SQL.Add('ORDER BY');
    SQL.Add('    NUM_NCM,');
    SQL.Add('    FLG_NAO_PIS_COFINS,');
    SQL.Add('    TIPO_NAO_PIS_COFINS,');
    SQL.Add('    COD_TRIB_ENTRADA,');
    SQL.Add('    COD_TRIB_SAIDA');

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

procedure TFrmGetWayMama.GerarNFClientes;
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
    SQL.Add('    CAPA.ID_EMPRESA_A = '+ CbxLoja.Text +'');

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

procedure TFrmGetWayMama.GerarNFFornec;
begin
  inherited;
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
    SQL.Add('    SUBSTR(CAPA.COD_PART,2, 60) AS COD_FORNECEDOR,');
    SQL.Add('    CAPA.NUM_DOC AS NUM_NF_FORN,');
    SQL.Add('    CAPA.SER AS NUM_SERIE_NF,');
    SQL.Add('    '''' AS NUM_SUBSERIE_NF,');
    SQL.Add('    CFOP.ID_CFOP AS CFOP,');
    SQL.Add('');
    SQL.Add('    CASE CAPA.ID_COI_A');
    SQL.Add('        WHEN 23 THEN 3');
    SQL.Add('        WHEN 201 THEN 2');
    SQL.Add('        WHEN 41 THEN 2');
    SQL.Add('        ELSE 0');
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
    SQL.Add('    CAPA.VL_MERC AS VAL_VENDA_VAREJO,');
    SQL.Add('    CAPA.VL_FRT AS VAL_FRETE,');
    SQL.Add('    CAPA.VL_SEG AS VAL_ACRESCIMO,');
    SQL.Add('    CAPA.VL_DESC AS VAL_DESCONTO,');
    SQL.Add('    '''' AS NUM_CGC,');
    SQL.Add('    CAPA.VL_BC_ICMS AS VAL_TOTAL_BC,');
    SQL.Add('    CAPA.VL_ICMS AS VAL_TOTAL_ICMS,');
    SQL.Add('    CAPA.VL_BC_ICMS_ST AS VAL_BC_SUBST,');
    SQL.Add('    CAPA.VL_ICMS_ST AS VAL_ICMS_SUBST,');
    SQL.Add('    0 AS VAL_FUNRURAL,');
    SQL.Add('');
    SQL.Add('    CASE CAPA.ID_COI_A');
    SQL.Add('        WHEN 25 THEN 5');
    SQL.Add('        WHEN 821 THEN 5');
    SQL.Add('        WHEN 21 THEN 1');
    SQL.Add('        WHEN 23 THEN 2');
    SQL.Add('        WHEN 1926 THEN 41');
    SQL.Add('        WHEN 32 THEN 23');
    SQL.Add('        WHEN 1 THEN 1');
    SQL.Add('    END AS COD_PERFIL,');
    SQL.Add('');
    SQL.Add('    0 AS VAL_DESP_ACESS,');
    SQL.Add('    ''N'' AS FLG_CANCELADO,');
    SQL.Add('    CAPA.OBSERVACAO_A AS DES_OBSERVACAO,');
    SQL.Add('    CAPA.CHV_NFE AS NUM_CHAVE_ACESSO');
    SQL.Add('FROM');
    SQL.Add('    FIS_T_C100 CAPA');
    SQL.Add('LEFT JOIN ');
    SQL.Add('    BAS_T_COI CFOP ');
    SQL.Add('ON ');
    SQL.Add('    CAPA.ID_COI_A = CFOP.ID_COI     ');
    SQL.Add('WHERE ');
    SQL.Add('    CAPA.ID_COI_A IN (25, 821, 21, 23, 201, 41) -- Transferencia 1926, 32 / ');
    SQL.Add('AND ');
    SQL.Add('    CAPA.COD_SIT IN (6, 0)');
    SQL.Add('AND');
    SQL.Add('    CAPA.STATUS_A = 2  ');



    SQL.Add('AND    ');
    SQL.Add('    CAPA.DT_DOC >= :INI ');
    SQL.Add('AND    ');
    SQL.Add('    CAPA.DT_DOC <= :FIM ');
    SQL.Add('AND    ');
    SQL.Add('    CAPA.ID_EMPRESA_A = '+ CbxLoja.Text +'');

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

procedure TFrmGetWayMama.GerarNFitensClientes;
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
    SQL.Add('    CAPA.ID_EMPRESA_A = '+ CbxLoja.Text +'');

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

procedure TFrmGetWayMama.GerarNFitensFornec;
begin
  inherited;

  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
    SQL.Add('    SUBSTR(CAPA.COD_PART,2, 60) AS COD_FORNECEDOR,');
    SQL.Add('    CAPA.NUM_DOC AS NUM_NF_FORN,');
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
    SQL.Add('        WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 70 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 12 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 10.49 THEN 15');
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
    SQL.Add('    ITEM.VL_DESC_PERCENTUAL_A AS VAL_DESCONTO_ITEM,');
    SQL.Add('    COALESCE(ITEM.SEGURO_RATEIO_A, 0) AS VAL_ACRESCIMO_ITEM, --');
    SQL.Add('    (ITEM.VL_IPI / ITEM.QTD) AS VAL_IPI_ITEM,');
    SQL.Add('    (ITEM.VL_ICMS_ST / ITEM.QTD) AS VAL_SUBST_ITEM,');
    SQL.Add('    COALESCE(ITEM.FRETE_RATEIO_A, 0) AS VAL_FRETE_ITEM, --');
    SQL.Add('    ITEM.VL_ICMS AS VAL_CREDITO_ICMS,');
    SQL.Add('    ITEM.VL_ITEM AS VAL_VENDA_VAREJO, --');
    SQL.Add('    ITEM.VL_ITEM AS VAL_TABELA_LIQ,');
    SQL.Add('    '''' AS NUM_CGC,');
    SQL.Add('    ITEM.VL_BC_ICMS AS VAL_TOT_BC_ICMS,');
    SQL.Add('    0 AS VAL_TOT_OUTROS_ICMS, --');
    SQL.Add('    ITEM.CFOP AS CFOP,');
    SQL.Add('    ITEM.VL_ITEM - ITEM.VL_BC_ICMS AS VAL_TOT_ISENTO, --');
    SQL.Add('    ITEM.VL_BC_ICMS_ST AS VAL_TOT_BC_ST,');
    SQL.Add('    ITEM.VL_ICMS_ST AS VAL_TOT_ST,');
    SQL.Add('    ITEM.NUM_ITEM AS NUM_ITEM,');
    SQL.Add('    0 AS TIPO_IPI,');
    SQL.Add('');
    SQL.Add('    CASE ');
    SQL.Add('        WHEN CAST(COALESCE(PRODUTOS.CLASSIFICACAO_FISCAL, ''99999999'') AS INTEGER ) = 0 THEN ''99999999''');
    SQL.Add('        ELSE COALESCE(PRODUTOS.CLASSIFICACAO_FISCAL, ''99999999'')');
    SQL.Add('    END AS NUM_NCM,');
    SQL.Add('');
    SQL.Add('    PRODUTOS_FORNECEDOR.REFERENCIA AS DES_REFERENCIA');
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
    SQL.Add('    ITEM.COD_ITEM = PRODUTOS.ID');
    SQL.Add('LEFT JOIN');
    SQL.Add('    PRODUTOS_FORNECEDOR');
    SQL.Add('ON');
    SQL.Add('    SUBSTR(CAPA.COD_PART,2, 60) = PRODUTOS_FORNECEDOR.FORNECEDOR');
    SQL.Add('AND');
    SQL.Add('    PRODUTOS_FORNECEDOR.PRODUTO = ITEM.COD_ITEM    ');
    SQL.Add('WHERE ');
    SQL.Add('    CAPA.ID_COI_A IN (25, 821, 21, 23, 201, 41) -- Transferencia 1926, 32');
    SQL.Add('AND ');
    SQL.Add('    CAPA.COD_SIT IN (6, 0)');
    SQL.Add('AND');
    SQL.Add('    CAPA.STATUS_A = 2  ');


    SQL.Add('AND    ');
    SQL.Add('    CAPA.DT_DOC >= :INI ');
    SQL.Add('AND    ');
    SQL.Add('    CAPA.DT_DOC <= :FIM ');
    SQL.Add('AND    ');
    SQL.Add('    CAPA.ID_EMPRESA_A = '+ CbxLoja.Text +'');

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

//      Layout.FieldByName('NUM_ITEM').AsInteger := count;

//      Layout.FieldByName('COD_PRODUTO').AsString := GerarPLU(QryPrincipal2.FieldByName('COD_PRODUTO').AsString);

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;

end;

procedure TFrmGetWayMama.GerarProdForn;
begin
  inherited;

  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
    SQL.Add('    PRODUTOS.CODPROD_SAMBANET AS COD_PRODUTO,');
    SQL.Add('    TAB_PRODUTO_FORNECEDOR.CODFORNEC AS COD_FORNECEDOR,');
    SQL.Add('    TAB_PRODUTO_FORNECEDOR.CODREF AS DES_REFERENCIA,');
    SQL.Add('    '''' AS NUM_CGC,');
    SQL.Add('    0 AS COD_DIVISAO,');
    SQL.Add('    PRODUTOS.UNIDADE AS DES_UNIDADE_COMPRA,');
    SQL.Add('    1 AS QTD_EMBALAGEM_COMPRA');
    SQL.Add('FROM');
    SQL.Add('    TAB_PRODUTO_FORNECEDOR');
    SQL.Add('LEFT JOIN');
    SQL.Add('    PRODUTOS');
    SQL.Add('ON');
    SQL.Add('    TAB_PRODUTO_FORNECEDOR.CODPROD = PRODUTOS.CODPROD');

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

procedure TFrmGetWayMama.GerarProdLoja;
begin
  inherited;
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
    SQL.Add('    PRODUTOS.CODPROD_SAMBANET AS COD_PRODUTO,');
    SQL.Add('    PRODUTOS.PRECO_CUST AS VAL_CUSTO_REP,');
    SQL.Add('    PRODUTOS.PRECO_UNIT AS VAL_VENDA,');
    SQL.Add('    COALESCE(PROMOCAO.PRECO_UNIT, 0) AS VAL_OFERTA,');
    SQL.Add('    PRODUTOS.ESTOQUE AS QTD_EST_VDA,');
    SQL.Add('    '''' AS TECLA_BALANCA,');
    SQL.Add('');
    SQL.Add('    CASE');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''00'' AND TRIBUTACAO.ALIQUOTA = 4.50 AND TRIBUTACAO.REDUCAO = 0.00 THEN 70  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''00'' AND TRIBUTACAO.ALIQUOTA = 7.00 AND TRIBUTACAO.REDUCAO = 0.00 THEN 2  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''00'' AND TRIBUTACAO.ALIQUOTA = 11.00 AND TRIBUTACAO.REDUCAO = 0.00 THEN 66  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''00'' AND TRIBUTACAO.ALIQUOTA = 12.00 AND TRIBUTACAO.REDUCAO = 0.00 THEN 3  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''00'' AND TRIBUTACAO.ALIQUOTA = 18.00 AND TRIBUTACAO.REDUCAO = 0.00 THEN 4  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''00'' AND TRIBUTACAO.ALIQUOTA = 18.00 AND TRIBUTACAO.REDUCAO = 33.33 THEN 7  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''00'' AND TRIBUTACAO.ALIQUOTA = 18.00 AND TRIBUTACAO.REDUCAO = 61.11 THEN 8  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''00'' AND TRIBUTACAO.ALIQUOTA = 25.00 AND TRIBUTACAO.REDUCAO = 0.00 THEN 5  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''20'' AND TRIBUTACAO.ALIQUOTA = 7.00 AND TRIBUTACAO.REDUCAO = 41.67 THEN 6  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''20'' AND TRIBUTACAO.ALIQUOTA = 7.00 AND TRIBUTACAO.REDUCAO = 61.11 THEN 8  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''20'' AND TRIBUTACAO.ALIQUOTA = 18.00 AND TRIBUTACAO.REDUCAO = 0.00 THEN 12  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''20'' AND TRIBUTACAO.ALIQUOTA = 18.00 AND TRIBUTACAO.REDUCAO = 33.33 THEN 15  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''20'' AND TRIBUTACAO.ALIQUOTA = 18.00 AND TRIBUTACAO.REDUCAO = 41.67 THEN 14  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''20'' AND TRIBUTACAO.ALIQUOTA = 18.00 AND TRIBUTACAO.REDUCAO = 61.11 THEN 16  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''40'' AND TRIBUTACAO.ALIQUOTA = 0.00 AND TRIBUTACAO.REDUCAO = 0.00 THEN 1  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''40'' AND TRIBUTACAO.ALIQUOTA = 0.00 AND TRIBUTACAO.REDUCAO = 4.00 THEN 1  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''40'' AND TRIBUTACAO.ALIQUOTA = 0.00 AND TRIBUTACAO.REDUCAO = 41.67 THEN 1  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''60'' AND TRIBUTACAO.ALIQUOTA = 0.00 AND TRIBUTACAO.REDUCAO = 0.00 THEN 12  ');
    SQL.Add('    END AS COD_TRIBUTACAO,');
    SQL.Add('');
    SQL.Add('    0 AS VAL_MARGEM,');
    SQL.Add('    1 AS QTD_ETIQUETA,');
    SQL.Add('');
    SQL.Add('    CASE');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''00'' AND TRIBUTACAO.ALIQUOTA = 4.50 AND TRIBUTACAO.REDUCAO = 0.00 THEN 70  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''00'' AND TRIBUTACAO.ALIQUOTA = 7.00 AND TRIBUTACAO.REDUCAO = 0.00 THEN 2  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''00'' AND TRIBUTACAO.ALIQUOTA = 11.00 AND TRIBUTACAO.REDUCAO = 0.00 THEN 66  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''00'' AND TRIBUTACAO.ALIQUOTA = 12.00 AND TRIBUTACAO.REDUCAO = 0.00 THEN 3  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''00'' AND TRIBUTACAO.ALIQUOTA = 18.00 AND TRIBUTACAO.REDUCAO = 0.00 THEN 4  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''00'' AND TRIBUTACAO.ALIQUOTA = 18.00 AND TRIBUTACAO.REDUCAO = 33.33 THEN 7  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''00'' AND TRIBUTACAO.ALIQUOTA = 18.00 AND TRIBUTACAO.REDUCAO = 61.11 THEN 8  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''00'' AND TRIBUTACAO.ALIQUOTA = 25.00 AND TRIBUTACAO.REDUCAO = 0.00 THEN 5  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''20'' AND TRIBUTACAO.ALIQUOTA = 7.00 AND TRIBUTACAO.REDUCAO = 41.67 THEN 6  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''20'' AND TRIBUTACAO.ALIQUOTA = 7.00 AND TRIBUTACAO.REDUCAO = 61.11 THEN 8  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''20'' AND TRIBUTACAO.ALIQUOTA = 18.00 AND TRIBUTACAO.REDUCAO = 0.00 THEN 12  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''20'' AND TRIBUTACAO.ALIQUOTA = 18.00 AND TRIBUTACAO.REDUCAO = 33.33 THEN 15  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''20'' AND TRIBUTACAO.ALIQUOTA = 18.00 AND TRIBUTACAO.REDUCAO = 41.67 THEN 14  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''20'' AND TRIBUTACAO.ALIQUOTA = 18.00 AND TRIBUTACAO.REDUCAO = 61.11 THEN 16  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''40'' AND TRIBUTACAO.ALIQUOTA = 0.00 AND TRIBUTACAO.REDUCAO = 0.00 THEN 1  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''40'' AND TRIBUTACAO.ALIQUOTA = 0.00 AND TRIBUTACAO.REDUCAO = 4.00 THEN 1  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''40'' AND TRIBUTACAO.ALIQUOTA = 0.00 AND TRIBUTACAO.REDUCAO = 41.67 THEN 1  ');
    SQL.Add('        WHEN TRIBUTACAO.CST = ''60'' AND TRIBUTACAO.ALIQUOTA = 0.00 AND TRIBUTACAO.REDUCAO = 0.00 THEN 12  ');
    SQL.Add('    END AS COD_TRIB_ENTRADA,');
    SQL.Add('');
    SQL.Add('    CASE ');
    SQL.Add('        WHEN PRODUTOS.ATIVO = ''S'' THEN ''N''');
    SQL.Add('        ELSE ''S''');
    SQL.Add('    END AS FLG_INATIVO,');
    SQL.Add('');
    SQL.Add('    PRODUTOS.CODPROD_SAMBANET AS COD_PRODUTO_ANT,');
    SQL.Add('');
    SQL.Add('    CASE ');
    SQL.Add('        WHEN COALESCE(CODNCM, ''00000000'') = ''00000000'' THEN ''99999999''');
    SQL.Add('        ELSE REPLACE(CODNCM, ''.'', '''')');
    SQL.Add('    END AS NUM_NCM,');
    SQL.Add('');
    SQL.Add('    0 AS TIPO_NCM,');
    SQL.Add('    0 AS VAL_VENDA_2, --');
    SQL.Add('');
    SQL.Add('    CASE ');
    SQL.Add('        WHEN COALESCE(PROMOCAO.DATAFIM, NULL) >= GETDATE() THEN PROMOCAO.DATAFIM');
    SQL.Add('        ELSE NULL');
    SQL.Add('    END AS DTA_VALIDA_OFERTA,');
    SQL.Add('');
    SQL.Add('    0 AS QTD_EST_MINIMO, --');
    SQL.Add('    NULL AS COD_VASILHAME,');
    SQL.Add('    ''N'' AS FORA_LINHA,');
    SQL.Add('    0 AS QTD_PRECO_DIF,');
    SQL.Add('    0 AS VAL_FORCA_VDA,');
    SQL.Add('    0 AS NUM_CEST,');
    SQL.Add('    0 AS PER_IVA,');
    SQL.Add('    0 AS PER_FCP_ST,');
    SQL.Add('    GETDATE() ');
    SQL.Add('FROM ');
    SQL.Add('    PRODUTOS');
    SQL.Add('LEFT JOIN');
    SQL.Add('    (');
    SQL.Add('        SELECT ');
    SQL.Add('            PRODUTOS.CODPROD,');
    SQL.Add('            SUBSTRING(CODTRIB, 2,3) AS CST, ');
    SQL.Add('            COALESCE(ALIQUOTA_ICMS.VALORTRIB, 0) AS ALIQUOTA,');
    SQL.Add('            PER_REDUC AS REDUCAO');
    SQL.Add('        FROM');
    SQL.Add('            PRODUTOS');
    SQL.Add('        LEFT JOIN ALIQUOTA_ICMS');
    SQL.Add('        ON PRODUTOS.CODALIQ = ALIQUOTA_ICMS.CODALIQ');
    SQL.Add('    ) AS TRIBUTACAO');
    SQL.Add('ON');
    SQL.Add('    PRODUTOS.CODPROD = TRIBUTACAO.CODPROD        ');
    SQL.Add('LEFT JOIN');
    SQL.Add('    PROMOCAO');
    SQL.Add('ON');
    SQL.Add('    PRODUTOS.CODPROD = PROMOCAO.CODPROD');
    SQL.Add('AND');
    SQL.Add('    PROMOCAO.DATAFIM >= GETDATE()    ');


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

       Layout.FieldByName('COD_PRODUTO').AsString := GerarPLU( Layout.FieldByName('COD_PRODUTO').AsString );
       Layout.FieldByName('COD_PRODUTO_ANT').AsString := GerarPLU(QryPrincipal2.FieldByName('COD_PRODUTO_ANT').AsString);

//      if( Layout.FieldByName('DTA_VALIDA_OFERTA').AsString <> '' ) then
         Layout.FieldByName('DTA_VALIDA_OFERTA').AsDateTime := FieldByName('DTA_VALIDA_OFERTA').AsDateTime;


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

procedure TFrmGetWayMama.GerarProdSimilar;
begin
  inherited;
  with QryPrincipal2 do
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

end.
