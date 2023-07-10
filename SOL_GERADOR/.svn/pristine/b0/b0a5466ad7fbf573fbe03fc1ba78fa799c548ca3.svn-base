unit UFrmSmPanda;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, UFrmModelo, Data.DBXOracle, Data.DB,
  Data.SqlExpr, Vcl.Menus, Vcl.StdCtrls, Vcl.Buttons, Vcl.ComCtrls,
  Vcl.Imaging.pngimage, Vcl.ExtCtrls, Data.DBXFirebird, Data.Win.ADODB,
  Vcl.Grids, Vcl.DBGrids, Datasnap.Provider, Datasnap.DBClient, //dxGDIPlusClasses,
  Math;

type
  TFrmSmPanda = class(TFrmModeloSis)
    btnGeraCest: TButton;
    BtnAmarrarCest: TButton;
    CbxLoja: TComboBox;
    lblLoja: TLabel;
    btnGerarEstoqueAtual: TButton;
    btnGeraCustoRep: TButton;
    btnGeraValorVenda: TButton;
    Label11: TLabel;
    procedure btnGeraCestClick(Sender: TObject);
    procedure BtnAmarrarCestClick(Sender: TObject);
    procedure EdtCamBancoExit(Sender: TObject);
    procedure btnGeraValorVendaClick(Sender: TObject);
    procedure btnGeraCustoRepClick(Sender: TObject);
    procedure btnGerarEstoqueAtualClick(Sender: TObject);
    procedure CkbProdLojaClick(Sender: TObject);
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

    procedure GerarValorVenda;
    procedure GeraCustoRep;
    procedure GeraEstoqueVenda;

  end;

var
  FrmSmPanda: TFrmSmPanda;
  ListNCM    : TStringList;
  TotalCont  : Integer;
  NumLinha : Integer;
  Arquivo: TextFile;
  FlgGeraDados : Boolean = false;
  FlgGeraCest : Boolean = false;
  FlgGeraAmarrarCest : Boolean = false;

  FlgAtualizaValVenda : Boolean = False;
  FlgAtualizaCustoRep : Boolean = False;
  FlgAtualizaEstoque  : Boolean = False;

implementation

{$R *.dfm}

uses xProc, UUtilidades, UProgresso;


procedure TFrmSmPanda.GerarProducao;
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

procedure TFrmSmPanda.GerarProduto;
var
   cod_produto, codbarras, TIPO : string;
   TotalCount, count, COD_PROD, CODIGO, NEW_CODPROD : Integer;
   QryGeraCodigoProduto : TSQLQuery;

begin
  inherited;

//  QryGeraCodigoProduto := TSQLQuery.Create(FrmProgresso);
//  with QryGeraCodigoProduto do
//  begin
//    SQLConnection := ScnBanco;
//
//    SQL.Clear;
//    SQL.Add('ALTER TABLE TAB_BARRAS_AUX ');
//    SQL.Add('ADD CODIGO_PRODUTO INT DEFAULT NULL; ');
//
//    try
//      ExecSQL;
//    except
//    end;
//
//    SQL.Clear;
//    SQL.Add('UPDATE TAB_BARRAS_AUX ');
//    SQL.Add('SET CODIGO_PRODUTO = :COD_PRODUTO  ');
//    SQL.Add('WHERE COD_BARRA_AUX = :COD_BARRA_PRINCIPAL ');
//    //SQL.Add('AND CHAR_LENGTH(COD_MATERIAL) >= 8 ');
//
//    try
//      ExecSQL;
//    except
//    end;
////
//  end;

  with QryPrincipal do
  begin
    Close;
    SQL.Clear;


     SQL.Add('   SELECT   ');
     SQL.Add('       CODIGO_PRODUTO AS COD_PRODUTO,   ');
     SQL.Add('       COD_BARRA_AUX AS COD_BARRA_PRINCIPAL,   ');
     SQL.Add('       DES_PRODUTO_AUX AS DES_REDUZIDA,   ');
     SQL.Add('       DES_PRODUTO_AUX AS DES_PRODUTO,   ');
     SQL.Add('       1 AS QTD_EMBALAGEM_COMPRA,   ');
     SQL.Add('       DES_UNIDADE_AUX AS DES_UNIDADE_COMPRA,   ');
     SQL.Add('       1 AS QTD_EMBALAGEM_VENDA,   ');
     SQL.Add('       DES_UNIDADE_AUX AS DES_UNIDADE_VENDA,   ');
     SQL.Add('       0 AS TIPO_IPI,   ');
     SQL.Add('       0 AS VAL_IPI,   ');
     SQL.Add('       999 AS COD_SECAO,   ');
     SQL.Add('       999 AS COD_GRUPO,   ');
     SQL.Add('       999 AS COD_SUB_GRUPO,   ');
     SQL.Add('       0 AS COD_PRODUTO_SIMILAR,   ');
     SQL.Add('       ''N'' AS IPV,   ');
     SQL.Add('       0 AS DIAS_VALIDADE,   ');
     SQL.Add('       0 AS TIPO_PRODUTO,   ');
     SQL.Add('       ''N'' AS FLG_NAO_PIS_COFINS,   ');
     SQL.Add('       ''N'' AS FLG_ENVIA_BALANCA,   ');
     SQL.Add('       -1 AS TIPO_NAO_PIS_COFINS,   ');
     SQL.Add('       0 AS TIPO_EVENTO,   ');
     SQL.Add('       0 AS COD_ASSOCIADO,   ');
     SQL.Add('       '''' AS DES_OBSERVACAO,   ');
     SQL.Add('       0 AS COD_INFO_NUTRICIONAL,   ');
     SQL.Add('       999 AS COD_TAB_SPED,   ');
     SQL.Add('       ''N'' AS FLG_ALCOOLICO,   ');
     SQL.Add('       0 AS TIPO_ESPECIE,   ');
     SQL.Add('       0 AS COD_CLASSIF,   ');
     SQL.Add('       1 AS VAL_VDA_PESO_BRUTO,   ');
     SQL.Add('       1 AS VAL_PESO_EMB,   ');
     SQL.Add('       0 AS TIPO_EXPLOSAO_COMPRA,   ');
     SQL.Add('       '''' AS DTA_INI_OPER,   ');
     SQL.Add('       '''' AS DES_PLAQUETA,   ');
     SQL.Add('       '''' AS MES_ANO_INI_DEPREC,   ');
     SQL.Add('       0 AS TIPO_BEM,   ');
     SQL.Add('       0 AS COD_FORNECEDOR,   ');
     SQL.Add('       0 AS NUM_NF,   ');
     SQL.Add('       '''' AS DTA_ENTRADA,   ');
     SQL.Add('       0 AS COD_NAT_BEM,   ');
     SQL.Add('       0 AS VAL_ORIG_BEM,   ');
     SQL.Add('       DES_PRODUTO_AUX AS DES_PRODUTO_ANT   ');
     SQL.Add('   FROM   ');
     SQL.Add('       TAB_BARRAS_AUX   ');
     SQL.Add('   WHERE COD_BARRA_AUX NOT IN (   ');
     SQL.Add('       SELECT DISTINCT   ');
     SQL.Add('           COD_EAN   ');
     SQL.Add('       FROM   ');
     SQL.Add('           TAB_CODIGO_BARRA   ');
     SQL.Add('   )   ');


    Open;
    First;
    NumLinha := 0;
    NEW_CODPROD := 10000;
    //count := 100000;
    //COD_PROD := 99999;
    //CODIGO := 0;

    TotalCount := SetCountTotal(SQL.Text);

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);
//      Inc(NEW_CODPROD);
      Layout.SetValues(QryPrincipal, NumLinha, TotalCount);


//
//      if Layout.FieldByName('COD_PRODUTO').AsString = '0' then
//      begin
//        with QryGeraCodigoProduto do
//        begin
//          Inc(COD_PROD);
//          Params.ParamByName('COD_PRODUTO').Value := NEW_CODPROD;
//          Params.ParamByName('COD_BARRA_PRINCIPAL').Value := Layout.FieldByName('COD_BARRA_PRINCIPAL').AsString;
//          Layout.FieldByName('COD_PRODUTO').AsInteger := Params.ParamByName('COD_PRODUTO').Value;
//          ExecSQL();
//        end;
//      end;

        //if Length(StrLBReplace(Trim(StrRetNums( FieldByName('COD_PRODUTO').AsString) ))) < 8 then


//      if Layout.FieldByName('COD_PRODUTO').AsString = '0' then
//      begin
//        Layout.FieldByName('COD_PRODUTO').AsInteger := NEW_CODPROD;
//      end;

      Layout.FieldByName('COD_PRODUTO').AsString := GerarPLU( Layout.FieldByName('COD_PRODUTO').AsString );

      if QryPrincipal.FieldByName('DTA_ENTRADA').AsString <> '' then
        Layout.FieldByName('DTA_ENTRADA').AsDateTime := FieldByName('DTA_ENTRADA').AsDateTime;



      //Layout.FieldByName('DES_OBSERVACAO').AsString := StrReplace(StrLBReplace(FieldByName('DES_OBSERVACAO').AsString), '\n', '');
      //Layout.FieldByName('DES_REDUZIDA').AsString := StrReplace(StrLBReplace(FieldByName('DES_REDUZIDA').AsString), '\n', '');
      //Layout.FieldByName('DES_PRODUTO').AsString := StrReplace(StrLBReplace(FieldByName('DES_PRODUTO').AsString), '\n', '');

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

procedure TFrmSmPanda.GerarScriptAmarrarCEST;
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

procedure TFrmSmPanda.GerarScriptCEST;
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

procedure TFrmSmPanda.GerarSecao;
var
   TotalCount : integer;
begin
  inherited;

  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

     SQL.Add('   SELECT DISTINCT   ');
     SQL.Add('       999 AS COD_SECAO,   ');
     SQL.Add('       ''A DEFINIR'' AS DES_SECAO,   ');
     SQL.Add('       0 AS VAL_META   ');
     SQL.Add('   FROM   ');
     SQL.Add('        TAB_PRODUTO   ');
     //SQL.Add('   LEFT JOIN TB_PRO_GR AS SECAO ON SECAO.CODIGO = PRODUTOS.COD_GR   ');


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

procedure TFrmSmPanda.GerarSubGrupo;
var
   TotalCount : Integer;
begin
  inherited;

  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

     SQL.Add('   SELECT DISTINCT   ');
     SQL.Add('        999 AS COD_SECAO,   ');
     SQL.Add('        999 AS COD_GRUPO,   ');
     SQL.Add('        999 AS COD_SUB_GRUPO,   ');
     SQL.Add('        ''A DEFINIR'' AS DES_SUB_GRUPO,   ');
     SQL.Add('        0 AS VAL_META,   ');
     SQL.Add('        0 AS VAL_MARGEM_REF,   ');
     SQL.Add('        0 AS QTD_DIA_SEGURANCA,   ');
     SQL.Add('        ''N'' AS FLG_ALCOOLICO   ');
     SQL.Add('   FROM   ');
     SQL.Add('        TAB_PRODUTO   ');
     //SQL.Add('   LEFT JOIN TB_PRO_GR AS SECAO ON SECAO.CODIGO = PRODUTOS.COD_GR   ');

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

procedure TFrmSmPanda.GerarTransportadora;
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

procedure TFrmSmPanda.GerarValorVenda;
var
  COD_PRODUTO : string;
begin
  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

     SQL.Add('   SELECT   ');
     SQL.Add('       TAB_PRODUTO.COD_PRODUTO AS COD_PRODUTO,   ');
     SQL.Add('       VENDA_AUX AS VAL_VENDA   ');
     SQL.Add('   FROM   ');
     SQL.Add('       TAB_EAN_AUXILIAR   ');
     SQL.Add('   LEFT JOIN TAB_PRODUTO ON TAB_PRODUTO.COD_BARRA_PRINCIPAL = TAB_EAN_AUXILIAR.EAN_AUX   ');
     SQL.Add('   WHERE TAB_PRODUTO.COD_PRODUTO IS NOT NULL   ');

    Open;
    First;

    NumLinha := 0;

    while not Eof do
    begin
      try
        if Cancelar then
        Break;

        Inc(NumLinha);

//        COD_PRODUTO := GerarPLU(QryPrincipal.FieldByName('COD_PRODUTO').AsString);
          COD_PRODUTO := QryPrincipal.FieldByName('COD_PRODUTO').AsString;


//        Writeln(Arquivo, 'UPDATE TAB_PRODUTO_LOJA SET COD_BARRA_AUX = ''G'+QryPrincipal.FieldByName('VAL_VENDA_2').AsString+''' WHERE COD_PRODUTO = '+COD_PRODUTO+' ; ');


        Writeln(Arquivo, 'UPDATE TAB_PRODUTO_LOJA SET VAL_VENDA = '''+QryPrincipal.FieldByName('VAL_VENDA').AsString+''' WHERE COD_PRODUTO = '''+QryPrincipal.FieldByName('COD_PRODUTO').AsString+''' ; ');


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

procedure TFrmSmPanda.GerarVenda;
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
     SQL.Add('           WHEN CHAR_LENGTH(PRODUTOS.COD_MATERIAL) >= 8 THEN PRODUTOS.CODIGO_PRODUTO      ');
     SQL.Add('           WHEN PRODUTOS.BALANCA = ''1'' AND PRODUTOS.COD_UNI = ''UN'' THEN PRODUTOS.COD_MATERIAL      ');
     SQL.Add('           WHEN PRODUTOS.BALANCA = ''1'' AND PRODUTOS.COD_UNI = ''KG'' THEN PRODUTOS.COD_MATERIAL      ');
     SQL.Add('           ELSE PRODUTOS.COD_MATERIAL      ');
     SQL.Add('       END AS COD_PRODUTO,   ');
     SQL.Add('      ');
     SQL.Add('       1 AS COD_LOJA,   ');
     SQL.Add('       0 AS IND_TIPO,   ');
     SQL.Add('       1 AS NUM_PDV,   ');
     SQL.Add('       VENDAS.QUANT_NORMAL AS QTD_TOTAL_PRODUTO,   ');
     SQL.Add('       VENDAS.VENDA_NORMAL AS VAL_TOTAL_PRODUTO,   ');
     SQL.Add('       VENDAS.VENDA_NORMAL AS VAL_PRECO_VENDA,   ');
     SQL.Add('       PRODUTOS.PRECO_COMPRA AS VAL_CUSTO_REP,   ');
     SQL.Add('       VENDAS.DATA AS DTA_SAIDA,   ');
     SQL.Add('       REPLACE(SUBSTRING(DATA FROM 6 FOR 2), ''-'', '''') || REPLACE(SUBSTRING(DATA FROM 1 FOR 5), ''-'', '''') AS DTA_MENSAL,   ');
     SQL.Add('       1 AS NUM_IDENT,   ');
     SQL.Add('       PRODUTOS.COD_MATERIAL AS COD_EAN,   ');
     SQL.Add('       ''0000'' AS DES_HORA,   ');
     SQL.Add('       0 AS COD_CLIENTE,   ');
     SQL.Add('       1 AS COD_ENTIDADE,   ');
     SQL.Add('       0 AS VAL_BASE_ICMS,   ');
     SQL.Add('       '''' AS DES_SITUACAO_TRIB,   ');
     SQL.Add('       0 AS VAL_ICMS,   ');
     SQL.Add('       0 AS NUM_CUPOM_FISCAL,   ');
     SQL.Add('       VENDAS.VENDA_NORMAL AS VAL_VENDA_PDV,   ');
     SQL.Add('      ');
     SQL.Add('       CASE   ');
     SQL.Add('           WHEN PRODUTOS.COD_ALI = ''3'' AND ALIQUO.DES_ALI = ''12%'' AND PRODUTOS.CST = ''000'' AND PRODUTOS.REDUCAO_BASE IS NULL THEN 3        ');
     SQL.Add('           WHEN PRODUTOS.COD_ALI = ''3'' AND ALIQUO.DES_ALI = ''12%'' AND PRODUTOS.CST = ''000'' AND PRODUTOS.REDUCAO_BASE = ''0.00'' THEN 3         ');
     SQL.Add('           WHEN PRODUTOS.COD_ALI = ''3'' AND ALIQUO.DES_ALI = ''12%'' AND PRODUTOS.CST = ''020'' AND PRODUTOS.REDUCAO_BASE IS NULL THEN 3         ');
     SQL.Add('           WHEN PRODUTOS.COD_ALI = ''3'' AND ALIQUO.DES_ALI = ''12%'' AND PRODUTOS.CST = ''020'' AND PRODUTOS.REDUCAO_BASE = ''41.67'' THEN 6         ');
     SQL.Add('           WHEN PRODUTOS.COD_ALI = ''4'' AND ALIQUO.DES_ALI = ''13%'' AND PRODUTOS.CST = ''020'' AND PRODUTOS.REDUCAO_BASE = ''0.00'' THEN 41         ');
     SQL.Add('           WHEN PRODUTOS.COD_ALI = ''6'' AND ALIQUO.DES_ALI = ''18%'' AND PRODUTOS.CST = ''000'' AND PRODUTOS.REDUCAO_BASE IS NULL THEN 4         ');
     SQL.Add('           WHEN PRODUTOS.COD_ALI = ''6'' AND ALIQUO.DES_ALI = ''18%'' AND PRODUTOS.CST = ''000'' AND PRODUTOS.REDUCAO_BASE = ''0.00'' THEN 4         ');
     SQL.Add('           WHEN PRODUTOS.COD_ALI = ''6'' AND ALIQUO.DES_ALI = ''18%'' AND PRODUTOS.CST = ''000'' AND PRODUTOS.REDUCAO_BASE = ''2.00'' THEN 4         ');
     SQL.Add('           WHEN PRODUTOS.COD_ALI = ''6'' AND ALIQUO.DES_ALI = ''18%'' AND PRODUTOS.CST = ''010'' AND PRODUTOS.REDUCAO_BASE IS NULL THEN 13         ');
     SQL.Add('           WHEN PRODUTOS.COD_ALI = ''6.00'' AND ALIQUO.DES_ALI = ''18%'' AND PRODUTOS.CST = ''020'' AND PRODUTOS.REDUCAO_BASE IS NULL THEN 4         ');
     SQL.Add('           WHEN PRODUTOS.COD_ALI = ''6.00'' AND ALIQUO.DES_ALI = ''18%'' AND PRODUTOS.CST = ''020'' AND PRODUTOS.REDUCAO_BASE = ''0.00'' THEN 4         ');
     SQL.Add('           WHEN PRODUTOS.COD_ALI = ''6.00'' AND ALIQUO.DES_ALI = ''18%'' AND PRODUTOS.CST = ''020'' AND PRODUTOS.REDUCAO_BASE = ''55.00'' THEN 4         ');
     SQL.Add('           WHEN PRODUTOS.COD_ALI = ''6'' AND ALIQUO.DES_ALI = ''18%'' AND PRODUTOS.CST = ''020'' AND PRODUTOS.REDUCAO_BASE = ''61.11'' THEN 8         ');
     SQL.Add('           WHEN PRODUTOS.COD_ALI = ''6'' AND ALIQUO.DES_ALI = ''18%'' AND PRODUTOS.CST = ''070'' AND PRODUTOS.REDUCAO_BASE IS NULL THEN 13         ');
     SQL.Add('           WHEN PRODUTOS.COD_ALI = ''12'' AND ALIQUO.DES_ALI = ''I1'' AND PRODUTOS.CST = ''040'' AND PRODUTOS.REDUCAO_BASE IS NULL THEN 1         ');
     SQL.Add('           WHEN PRODUTOS.COD_ALI = ''12'' AND ALIQUO.DES_ALI = ''I1'' AND PRODUTOS.CST = ''040'' AND PRODUTOS.REDUCAO_BASE = ''0.00'' THEN 1         ');
     SQL.Add('           WHEN PRODUTOS.COD_ALI = ''13'' AND ALIQUO.DES_ALI = ''F1'' AND PRODUTOS.CST = ''060'' AND PRODUTOS.REDUCAO_BASE IS NULL THEN 13         ');
     SQL.Add('           WHEN PRODUTOS.COD_ALI = ''13'' AND ALIQUO.DES_ALI = ''F1'' AND PRODUTOS.CST = ''060'' AND PRODUTOS.REDUCAO_BASE = ''0.00'' THEN 13         ');
     SQL.Add('           WHEN PRODUTOS.COD_ALI = ''15'' AND ALIQUO.DES_ALI = ''14%'' AND PRODUTOS.CST = ''000'' AND PRODUTOS.REDUCAO_BASE IS NULL THEN 39         ');
     SQL.Add('           WHEN PRODUTOS.COD_ALI = ''16'' AND ALIQUO.DES_ALI = ''20%'' AND PRODUTOS.CST = ''000'' AND PRODUTOS.REDUCAO_BASE IS NULL THEN 40      ');
     SQL.Add('           WHEN PRODUTOS.COD_ALI = ''16'' AND ALIQUO.DES_ALI = ''20%'' AND PRODUTOS.CST = ''000'' AND PRODUTOS.REDUCAO_BASE = ''0'' THEN 40      ');
     SQL.Add('           ELSE 1        ');
     SQL.Add('       END AS COD_TRIBUTACAO,   ');
     SQL.Add('      ');
     SQL.Add('       ''N'' AS FLG_CUPOM_CANCELADO,   ');
     SQL.Add('       PRODUTOS.NCM AS NUM_NCM,   ');
     SQL.Add('       COALESCE(PRODUTOS.PIS_COFINS_COD_CREDITO, 999) AS COD_TAB_SPED,   ');
     SQL.Add('      ');
     SQL.Add('       CASE   ');
     SQL.Add('           WHEN PRODUTOS.PIS_COFINS_CODIGO = 1 AND PIS_COFINS.CST_ENT = 50 AND PIS_COFINS.CST_SAI = 1 AND PRODUTOS.PIS_COFINS_COD_CREDITO = 0 THEN ''N''       ');
     SQL.Add('           WHEN PRODUTOS.PIS_COFINS_CODIGO = 2 AND PIS_COFINS.CST_ENT = 60 AND PIS_COFINS.CST_SAI = 1 AND PRODUTOS.PIS_COFINS_COD_CREDITO = 101 THEN ''N''       ');
     SQL.Add('           WHEN PRODUTOS.PIS_COFINS_CODIGO = 5 AND PIS_COFINS.CST_ENT = 70 AND PIS_COFINS.CST_SAI = 4 AND PRODUTOS.PIS_COFINS_COD_CREDITO = 1 THEN ''S''       ');
     SQL.Add('           WHEN PRODUTOS.PIS_COFINS_CODIGO = 5 AND PIS_COFINS.CST_ENT = 70 AND PIS_COFINS.CST_SAI = 4 AND PRODUTOS.PIS_COFINS_COD_CREDITO = 2 THEN ''S''       ');
     SQL.Add('           WHEN PRODUTOS.PIS_COFINS_CODIGO = 5 AND PIS_COFINS.CST_ENT = 70 AND PIS_COFINS.CST_SAI = 4 AND PRODUTOS.PIS_COFINS_COD_CREDITO = 202 THEN ''S''       ');
     SQL.Add('           WHEN PRODUTOS.PIS_COFINS_CODIGO = 6 AND PIS_COFINS.CST_ENT = 73 AND PIS_COFINS.CST_SAI = 6 AND PRODUTOS.PIS_COFINS_COD_CREDITO = 105 THEN ''S''       ');
     SQL.Add('           WHEN PRODUTOS.PIS_COFINS_CODIGO = 6 AND PIS_COFINS.CST_ENT = 73 AND PIS_COFINS.CST_SAI = 6 AND PRODUTOS.PIS_COFINS_COD_CREDITO = 108 THEN ''S''       ');
     SQL.Add('           WHEN PRODUTOS.PIS_COFINS_CODIGO = 6 AND PIS_COFINS.CST_ENT = 73 AND PIS_COFINS.CST_SAI = 6 AND PRODUTOS.PIS_COFINS_COD_CREDITO = 110 THEN ''S''       ');
     SQL.Add('           WHEN PRODUTOS.PIS_COFINS_CODIGO = 6 AND PIS_COFINS.CST_ENT = 73 AND PIS_COFINS.CST_SAI = 6 AND PRODUTOS.PIS_COFINS_COD_CREDITO = 111 THEN ''S''       ');
     SQL.Add('           WHEN PRODUTOS.PIS_COFINS_CODIGO = 6 AND PIS_COFINS.CST_ENT = 73 AND PIS_COFINS.CST_SAI = 6 AND PRODUTOS.PIS_COFINS_COD_CREDITO = 113 THEN ''S''       ');
     SQL.Add('           WHEN PRODUTOS.PIS_COFINS_CODIGO = 6 AND PIS_COFINS.CST_ENT = 73 AND PIS_COFINS.CST_SAI = 6 AND PRODUTOS.PIS_COFINS_COD_CREDITO = 115 THEN ''S''       ');
     SQL.Add('           WHEN PRODUTOS.PIS_COFINS_CODIGO = 6 AND PIS_COFINS.CST_ENT = 73 AND PIS_COFINS.CST_SAI = 6 AND PRODUTOS.PIS_COFINS_COD_CREDITO = 116 THEN ''S''       ');
     SQL.Add('           WHEN PRODUTOS.PIS_COFINS_CODIGO = 6 AND PIS_COFINS.CST_ENT = 73 AND PIS_COFINS.CST_SAI = 6 AND PRODUTOS.PIS_COFINS_COD_CREDITO = 117 THEN ''S''       ');
     SQL.Add('           WHEN PRODUTOS.PIS_COFINS_CODIGO = 6 AND PIS_COFINS.CST_ENT = 73 AND PIS_COFINS.CST_SAI = 6 AND PRODUTOS.PIS_COFINS_COD_CREDITO = 119 THEN ''S''       ');
     SQL.Add('           WHEN PRODUTOS.PIS_COFINS_CODIGO = 6 AND PIS_COFINS.CST_ENT = 73 AND PIS_COFINS.CST_SAI = 6 AND PRODUTOS.PIS_COFINS_COD_CREDITO = 120 THEN ''S''       ');
     SQL.Add('           WHEN PRODUTOS.PIS_COFINS_CODIGO = 6 AND PIS_COFINS.CST_ENT = 73 AND PIS_COFINS.CST_SAI = 6 AND PRODUTOS.PIS_COFINS_COD_CREDITO = 121 THEN ''S''       ');
     SQL.Add('           WHEN PRODUTOS.PIS_COFINS_CODIGO = 6 AND PIS_COFINS.CST_ENT = 73 AND PIS_COFINS.CST_SAI = 6 AND PRODUTOS.PIS_COFINS_COD_CREDITO = 122 THEN ''S''       ');
     SQL.Add('           WHEN PRODUTOS.PIS_COFINS_CODIGO = 6 AND PIS_COFINS.CST_ENT = 73 AND PIS_COFINS.CST_SAI = 6 AND PRODUTOS.PIS_COFINS_COD_CREDITO = 123 THEN ''S''       ');
     SQL.Add('           WHEN PRODUTOS.PIS_COFINS_CODIGO = 6 AND PIS_COFINS.CST_ENT = 73 AND PIS_COFINS.CST_SAI = 6 AND PRODUTOS.PIS_COFINS_COD_CREDITO = 124 THEN ''S''       ');
     SQL.Add('           WHEN PRODUTOS.PIS_COFINS_CODIGO = 6 AND PIS_COFINS.CST_ENT = 73 AND PIS_COFINS.CST_SAI = 6 AND PRODUTOS.PIS_COFINS_COD_CREDITO = 125 THEN ''S''       ');
     SQL.Add('           WHEN PRODUTOS.PIS_COFINS_CODIGO = 6 AND PIS_COFINS.CST_ENT = 73 AND PIS_COFINS.CST_SAI = 6 AND PRODUTOS.PIS_COFINS_COD_CREDITO = 126 THEN ''S''       ');
     SQL.Add('           WHEN PRODUTOS.PIS_COFINS_CODIGO = 6 AND PIS_COFINS.CST_ENT = 73 AND PIS_COFINS.CST_SAI = 6 AND PRODUTOS.PIS_COFINS_COD_CREDITO = 127 THEN ''S''       ');
     SQL.Add('           WHEN PRODUTOS.PIS_COFINS_CODIGO = 6 AND PIS_COFINS.CST_ENT = 73 AND PIS_COFINS.CST_SAI = 6 AND PRODUTOS.PIS_COFINS_COD_CREDITO = 128 THEN ''S''       ');
     SQL.Add('           WHEN PRODUTOS.PIS_COFINS_CODIGO = 6 AND PIS_COFINS.CST_ENT = 73 AND PIS_COFINS.CST_SAI = 6 AND PRODUTOS.PIS_COFINS_COD_CREDITO = 129 THEN ''S''       ');
     SQL.Add('           WHEN PRODUTOS.PIS_COFINS_CODIGO = 6 AND PIS_COFINS.CST_ENT = 73 AND PIS_COFINS.CST_SAI = 6 AND PRODUTOS.PIS_COFINS_COD_CREDITO = 130 THEN ''S''       ');
     SQL.Add('           WHEN PRODUTOS.PIS_COFINS_CODIGO = 6 AND PIS_COFINS.CST_ENT = 73 AND PIS_COFINS.CST_SAI = 6 AND PRODUTOS.PIS_COFINS_COD_CREDITO = 914 THEN ''S''       ');
     SQL.Add('           WHEN PRODUTOS.PIS_COFINS_CODIGO = 6 AND PIS_COFINS.CST_ENT = 73 AND PIS_COFINS.CST_SAI = 6 AND PRODUTOS.PIS_COFINS_COD_CREDITO = 918 THEN ''S''        ');
     SQL.Add('           ELSE ''N''       ');
     SQL.Add('       END AS FLG_NAO_PIS_COFINS,   ');
     SQL.Add('      ');
     SQL.Add('       CASE   ');
     SQL.Add('           WHEN PRODUTOS.PIS_COFINS_CODIGO = 1 AND PIS_COFINS.CST_ENT = 50 AND PIS_COFINS.CST_SAI = 1 AND PRODUTOS.PIS_COFINS_COD_CREDITO = 0 THEN -1       ');
     SQL.Add('           WHEN PRODUTOS.PIS_COFINS_CODIGO = 2 AND PIS_COFINS.CST_ENT = 60 AND PIS_COFINS.CST_SAI = 1 AND PRODUTOS.PIS_COFINS_COD_CREDITO = 101 THEN -1       ');
     SQL.Add('           WHEN PRODUTOS.PIS_COFINS_CODIGO = 5 AND PIS_COFINS.CST_ENT = 70 AND PIS_COFINS.CST_SAI = 4 AND PRODUTOS.PIS_COFINS_COD_CREDITO = 1 THEN 1       ');
     SQL.Add('           WHEN PRODUTOS.PIS_COFINS_CODIGO = 5 AND PIS_COFINS.CST_ENT = 70 AND PIS_COFINS.CST_SAI = 4 AND PRODUTOS.PIS_COFINS_COD_CREDITO = 2 THEN 1       ');
     SQL.Add('           WHEN PRODUTOS.PIS_COFINS_CODIGO = 5 AND PIS_COFINS.CST_ENT = 70 AND PIS_COFINS.CST_SAI = 4 AND PRODUTOS.PIS_COFINS_COD_CREDITO = 202 THEN 1       ');
     SQL.Add('           WHEN PRODUTOS.PIS_COFINS_CODIGO = 6 AND PIS_COFINS.CST_ENT = 73 AND PIS_COFINS.CST_SAI = 6 AND PRODUTOS.PIS_COFINS_COD_CREDITO = 105 THEN 0       ');
     SQL.Add('           WHEN PRODUTOS.PIS_COFINS_CODIGO = 6 AND PIS_COFINS.CST_ENT = 73 AND PIS_COFINS.CST_SAI = 6 AND PRODUTOS.PIS_COFINS_COD_CREDITO = 108 THEN 0       ');
     SQL.Add('           WHEN PRODUTOS.PIS_COFINS_CODIGO = 6 AND PIS_COFINS.CST_ENT = 73 AND PIS_COFINS.CST_SAI = 6 AND PRODUTOS.PIS_COFINS_COD_CREDITO = 110 THEN 0       ');
     SQL.Add('           WHEN PRODUTOS.PIS_COFINS_CODIGO = 6 AND PIS_COFINS.CST_ENT = 73 AND PIS_COFINS.CST_SAI = 6 AND PRODUTOS.PIS_COFINS_COD_CREDITO = 111 THEN 0       ');
     SQL.Add('           WHEN PRODUTOS.PIS_COFINS_CODIGO = 6 AND PIS_COFINS.CST_ENT = 73 AND PIS_COFINS.CST_SAI = 6 AND PRODUTOS.PIS_COFINS_COD_CREDITO = 113 THEN 0       ');
     SQL.Add('           WHEN PRODUTOS.PIS_COFINS_CODIGO = 6 AND PIS_COFINS.CST_ENT = 73 AND PIS_COFINS.CST_SAI = 6 AND PRODUTOS.PIS_COFINS_COD_CREDITO = 115 THEN 0       ');
     SQL.Add('           WHEN PRODUTOS.PIS_COFINS_CODIGO = 6 AND PIS_COFINS.CST_ENT = 73 AND PIS_COFINS.CST_SAI = 6 AND PRODUTOS.PIS_COFINS_COD_CREDITO = 116 THEN 0       ');
     SQL.Add('           WHEN PRODUTOS.PIS_COFINS_CODIGO = 6 AND PIS_COFINS.CST_ENT = 73 AND PIS_COFINS.CST_SAI = 6 AND PRODUTOS.PIS_COFINS_COD_CREDITO = 117 THEN 0       ');
     SQL.Add('           WHEN PRODUTOS.PIS_COFINS_CODIGO = 6 AND PIS_COFINS.CST_ENT = 73 AND PIS_COFINS.CST_SAI = 6 AND PRODUTOS.PIS_COFINS_COD_CREDITO = 119 THEN 0       ');
     SQL.Add('           WHEN PRODUTOS.PIS_COFINS_CODIGO = 6 AND PIS_COFINS.CST_ENT = 73 AND PIS_COFINS.CST_SAI = 6 AND PRODUTOS.PIS_COFINS_COD_CREDITO = 120 THEN 0       ');
     SQL.Add('           WHEN PRODUTOS.PIS_COFINS_CODIGO = 6 AND PIS_COFINS.CST_ENT = 73 AND PIS_COFINS.CST_SAI = 6 AND PRODUTOS.PIS_COFINS_COD_CREDITO = 121 THEN 0       ');
     SQL.Add('           WHEN PRODUTOS.PIS_COFINS_CODIGO = 6 AND PIS_COFINS.CST_ENT = 73 AND PIS_COFINS.CST_SAI = 6 AND PRODUTOS.PIS_COFINS_COD_CREDITO = 122 THEN 0       ');
     SQL.Add('           WHEN PRODUTOS.PIS_COFINS_CODIGO = 6 AND PIS_COFINS.CST_ENT = 73 AND PIS_COFINS.CST_SAI = 6 AND PRODUTOS.PIS_COFINS_COD_CREDITO = 123 THEN 0       ');
     SQL.Add('           WHEN PRODUTOS.PIS_COFINS_CODIGO = 6 AND PIS_COFINS.CST_ENT = 73 AND PIS_COFINS.CST_SAI = 6 AND PRODUTOS.PIS_COFINS_COD_CREDITO = 124 THEN 0       ');
     SQL.Add('           WHEN PRODUTOS.PIS_COFINS_CODIGO = 6 AND PIS_COFINS.CST_ENT = 73 AND PIS_COFINS.CST_SAI = 6 AND PRODUTOS.PIS_COFINS_COD_CREDITO = 125 THEN 0       ');
     SQL.Add('           WHEN PRODUTOS.PIS_COFINS_CODIGO = 6 AND PIS_COFINS.CST_ENT = 73 AND PIS_COFINS.CST_SAI = 6 AND PRODUTOS.PIS_COFINS_COD_CREDITO = 126 THEN 0       ');
     SQL.Add('           WHEN PRODUTOS.PIS_COFINS_CODIGO = 6 AND PIS_COFINS.CST_ENT = 73 AND PIS_COFINS.CST_SAI = 6 AND PRODUTOS.PIS_COFINS_COD_CREDITO = 127 THEN 0       ');
     SQL.Add('           WHEN PRODUTOS.PIS_COFINS_CODIGO = 6 AND PIS_COFINS.CST_ENT = 73 AND PIS_COFINS.CST_SAI = 6 AND PRODUTOS.PIS_COFINS_COD_CREDITO = 128 THEN 0       ');
     SQL.Add('           WHEN PRODUTOS.PIS_COFINS_CODIGO = 6 AND PIS_COFINS.CST_ENT = 73 AND PIS_COFINS.CST_SAI = 6 AND PRODUTOS.PIS_COFINS_COD_CREDITO = 129 THEN 0       ');
     SQL.Add('           WHEN PRODUTOS.PIS_COFINS_CODIGO = 6 AND PIS_COFINS.CST_ENT = 73 AND PIS_COFINS.CST_SAI = 6 AND PRODUTOS.PIS_COFINS_COD_CREDITO = 130 THEN 0       ');
     SQL.Add('           WHEN PRODUTOS.PIS_COFINS_CODIGO = 6 AND PIS_COFINS.CST_ENT = 73 AND PIS_COFINS.CST_SAI = 6 AND PRODUTOS.PIS_COFINS_COD_CREDITO = 914 THEN 0       ');
     SQL.Add('           WHEN PRODUTOS.PIS_COFINS_CODIGO = 6 AND PIS_COFINS.CST_ENT = 73 AND PIS_COFINS.CST_SAI = 6 AND PRODUTOS.PIS_COFINS_COD_CREDITO = 918 THEN 0        ');
     SQL.Add('           ELSE -1      ');
     SQL.Add('       END AS TIPO_NAO_PIS_COFINS,   ');
     SQL.Add('      ');
     SQL.Add('       ''N'' AS FLG_ONLINE,   ');
     SQL.Add('       ''N'' AS FLG_OFERTA,   ');
     SQL.Add('       0 AS COD_ASSOCIADO   ');
     SQL.Add('   FROM   ');
     SQL.Add('       VENDA_MERCADORIA AS VENDAS   ');
     SQL.Add('   LEFT JOIN MATERI AS PRODUTOS ON PRODUTOS.COD_MATERIAL = VENDAS.CODIGO_PRODUTO   ');
     SQL.Add('   LEFT JOIN ALIQUO ON ALIQUO.COD_ALI = PRODUTOS.COD_ALI   ');
     SQL.Add('   LEFT JOIN PIS_COFINS ON PIS_COFINS.CODIGO = PRODUTOS.PIS_COFINS_CODIGO   ');
     SQL.Add('   WHERE PRODUTOS.COD_MATERIAL IS NOT NULL    ');
     SQL.Add('   AND VENDAS.DATA >= :INI');
     SQL.Add('   AND VENDAS.DATA <= :FIM');


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
      //Layout.FieldByName('DTA_MENSAL').AsDateTime := QryPrincipal.FieldByName('DTA_MENSAL').AsDateTime;

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;
end;

procedure TFrmSmPanda.BtnAmarrarCestClick(Sender: TObject);
begin
  inherited;
    inherited;
  FlgGeraAmarrarCest := True;
  BtnGerar.Click;
  FlgGeraAmarrarCest := False;
end;

procedure TFrmSmPanda.btnGeraCestClick(Sender: TObject);
begin
  inherited;
  FlgGeraCest := True;
  BtnGerar.Click;
  FlgGeraCest := False;
end;

procedure TFrmSmPanda.btnGeraCustoRepClick(Sender: TObject);
begin
  inherited;
  FlgAtualizaCustoRep := True;
  BtnGerar.Click;
  FlgAtualizaCustoRep := False;
end;

procedure TFrmSmPanda.BtnGerarClick(Sender: TObject);
begin
//  inherited;
     if FlgAtualizaValVenda then
   begin
     AssignFile(Arquivo, EdtCamArquivo.Text + '\SCRIPT_ATUALIZA_VALOR_VENDA.TXT' );
     Rewrite(Arquivo);
     CkbProdLoja.Checked := True;
   end;

   if FlgAtualizaCustoRep then
   begin
     AssignFile(Arquivo, EdtCamArquivo.Text + '\SCRIPT_ATUALIZA_CUSTO_REP.TXT' );
     Rewrite(Arquivo);
     CkbProdLoja.Checked := True;
   end;

   if FlgAtualizaEstoque then
   begin
     AssignFile(Arquivo, EdtCamArquivo.Text + '\SCRIPT_ATUALIZA_COD_PRODUTO_ANT.TXT' );
     Rewrite(Arquivo);
     CkbProdLoja.Checked := True;
   end;

  inherited;


  if FlgAtualizaValVenda then
    CloseFile(Arquivo);

  if FlgAtualizaCustoRep then
    CloseFile(Arquivo);

  if FlgAtualizaEstoque then
    CloseFile(Arquivo);
end;

procedure TFrmSmPanda.btnGerarEstoqueAtualClick(Sender: TObject);
begin
  inherited;
  FlgAtualizaEstoque := True;
  BtnGerar.Click;
  FlgAtualizaEstoque := False;
end;

procedure TFrmSmPanda.btnGeraValorVendaClick(Sender: TObject);
begin
  inherited;
  FlgAtualizaValVenda := True;
  BtnGerar.Click;
  FlgAtualizaValVenda := False;

end;

procedure TFrmSmPanda.CkbProdLojaClick(Sender: TObject);
begin
  inherited;
  btnGeraValorVenda.Enabled := True;
  btnGeraCustoRep.Enabled := True;
  btnGerarEstoqueAtual.Enabled := True;

  if CkbProdLoja.Checked = False then
  begin
    btnGeraValorVenda.Enabled := False;
    btnGeraCustoRep.Enabled := False;
    btnGerarEstoqueAtual.Enabled := False;
  end;
end;

procedure TFrmSmPanda.EdtCamBancoExit(Sender: TObject);
begin
  inherited;
  CriarFB(EdtCamBanco);
end;

procedure TFrmSmPanda.GeraCustoRep;
var
  COD_PRODUTO : string;
begin
  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

     SQL.Add('   SELECT   ');
     SQL.Add('       TAB_PRODUTO.COD_PRODUTO AS COD_PRODUTO,   ');
     SQL.Add('       --TAB_PRODUTO.COD_BARRA_PRINCIPAL,   ');
     SQL.Add('       --EAN_AUX,   ');
     SQL.Add('       CUSTO_AUX AS VAL_CUSTO_REP   ');
     SQL.Add('   FROM   ');
     SQL.Add('       TAB_EAN_AUXILIAR   ');
     SQL.Add('   LEFT JOIN TAB_PRODUTO ON TAB_PRODUTO.COD_BARRA_PRINCIPAL = TAB_EAN_AUXILIAR.EAN_AUX   ');
     SQL.Add('   WHERE TAB_PRODUTO.COD_PRODUTO IS NOT NULL   ');



    Open;
    First;

    NumLinha := 0;

    while not Eof do
    begin
      try
        if Cancelar then
        Break;

        Inc(NumLinha);

//        COD_PRODUTO := GerarPLU(QryPrincipal.FieldByName('COD_PRODUTO').AsString);
          COD_PRODUTO := QryPrincipal.FieldByName('COD_PRODUTO').AsString;

        Writeln(Arquivo, 'UPDATE TAB_PRODUTO_LOJA SET VAL_CUSTO_REP = '''+QryPrincipal.FieldByName('VAL_CUSTO_REP').AsString+''' WHERE COD_PRODUTO = '''+COD_PRODUTO+''' ; ');

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

procedure TFrmSmPanda.GeraEstoqueVenda;
var
  COD_PRODUTO : string;
begin
  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

     SQL.Add('   SELECT DISTINCT   ');
     SQL.Add('       TAB_PRODUTO.COD_PRODUTO AS COD_PRODUTO,   ');
     SQL.Add('       TAB_EAN_AUXILIAR.COD_AUX AS COD_PRODUTO_ANT   ');
     SQL.Add('   FROM   ');
     SQL.Add('       TAB_PRODUTO   ');
     SQL.Add('   LEFT JOIN TAB_EAN_AUXILIAR ON TAB_EAN_AUXILIAR.EAN_AUX = TAB_PRODUTO.COD_BARRA_PRINCIPAL   ');
     SQL.Add('   WHERE TAB_EAN_AUXILIAR.COD_AUX IS NOT NULL   ');



    Open;
    First;

    NumLinha := 0;

    while not Eof do
    begin
      try
        if Cancelar then
        Break;

        Inc(NumLinha);

        COD_PRODUTO := QryPrincipal.FieldByName('COD_PRODUTO').AsString;

        Writeln(Arquivo, 'UPDATE TAB_PRODUTO_LOJA SET COD_PRODUTO_ANT = '''+QryPrincipal.FieldByName('COD_PRODUTO_ANT').AsString+''' WHERE COD_PRODUTO = '''+COD_PRODUTO+''' ; ');

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

procedure TFrmSmPanda.GerarCest;
var
   TotalCount : integer;
   count : integer;
begin
  inherited;
  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

     SQL.Add('   SELECT DISTINCT   ');
     SQL.Add('       0 AS COD_CEST,   ');
     SQL.Add('      ');
     SQL.Add('       CASE   ');
     SQL.Add('             WHEN PRODUTOS.CEST = '''' THEN ''99999999''   ');
     SQL.Add('             WHEN PRODUTOS.CEST IS NULL THEN ''99999999''   ');
     SQL.Add('             ELSE PRODUTOS.CEST    ');
     SQL.Add('       END AS NUM_CEST,   ');
     SQL.Add('      ');
     SQL.Add('       CEST.DESCRICAO AS DES_CEST   ');
     SQL.Add('   FROM   ');
     SQL.Add('       TB_PRO AS PRODUTOS   ');
     SQL.Add('   LEFT JOIN TB_NFE_CEST AS CEST ON CEST.CEST = PRODUTOS.CEST   ');



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

procedure TFrmSmPanda.GerarCliente;
//var
//  QryGeraCodigoCliente : TSQLQuery;
//  CODIGO_CLIENTE : Integer;
begin
  inherited;

//  QryGeraCodigoCliente := TSQLQuery.Create(FrmProgresso);
//  with QryGeraCodigoCliente do
//  begin
//    SQLConnection := ScnBanco;
//
//    SQL.Clear;
//    SQL.Add('ALTER TABLE EMD105 ');
//    SQL.Add('ADD CODIGO_CLIENTE INT DEFAULT NULL; ');
//
//    try
//      //ExecSQL;
//    except
//    end;
//
//    SQL.Clear;
//    SQL.Add('UPDATE EMD105');
//    SQL.Add('SET CODIGO_CLIENTE = :COD_CLIENTE ');
//    SQL.Add('WHERE COALESCE(REPLACE(REPLACE(REPLACE(CGC_CPF, ''.'', ''''), ''/'', ''''), ''-'', ''''), '''') = :NUM_CGC ');
//
//    try
//      //ExecSQL;
//    except
//    end;
//
//  end;

  with QryPrincipal do
  begin
    Close;
    SQL.Clear;


     SQL.Add('   SELECT   ');
     SQL.Add('       CLIENTES.CODCLI AS COD_CLIENTE,   ');
     SQL.Add('       CLIENTES.NOMECLI AS DES_CLIENTE,   ');
     SQL.Add('       COALESCE(REPLACE(REPLACE(REPLACE(CLIENTES.CNPJ, ''-'', ''''), ''.'', ''''), ''/'', ''''), COALESCE(CLIENTES.CPF, '''')) AS NUM_CGC,   ');
     SQL.Add('          ');
     SQL.Add('       CASE   ');
     SQL.Add('           WHEN CLIENTES.FLAGFISICA = ''Y'' THEN ''''   ');
     SQL.Add('           ELSE COALESCE(UPPER(CLIENTES.INSCR), '''')   ');
     SQL.Add('       END AS NUM_INSC_EST,   ');
     SQL.Add('          ');
     SQL.Add('       COALESCE(CLIENTES.ENDERECO, ''A DEFINIR'') AS DES_ENDERECO,   ');
     SQL.Add('       COALESCE(CLIENTES.BAIRRO, ''A DEFINIR'') AS DES_BAIRRO,   ');
     SQL.Add('       COALESCE(CLIENTES.CIDADE, '''') AS DES_CIDADE,   ');
     SQL.Add('       COALESCE(CLIENTES.ESTADO, '''') AS DES_SIGLA,   ');
     SQL.Add('       COALESCE(CLIENTES.CEP, '''') AS NUM_CEP,   ');
     SQL.Add('       COALESCE(CLIENTES.TELEFONE, '''') AS NUM_FONE,   ');
     SQL.Add('       COALESCE(CLIENTES.FAX, '''') AS NUM_FAX,   ');
     SQL.Add('       COALESCE(CLIENTES.CONTATO, '''') AS DES_CONTATO,   ');
     SQL.Add('          ');
     SQL.Add('       CASE   ');
     SQL.Add('           WHEN CLIENTES.SEXO = ''F'' THEN 1   ');
     SQL.Add('           ELSE 0   ');
     SQL.Add('       END AS FLG_SEXO,   ');
     SQL.Add('          ');
     SQL.Add('       COALESCE(CLIENTES.LIMITECRED, 0) AS VAL_LIMITE_CRETID,   ');
     SQL.Add('       0 AS VAL_LIMITE_CONV,   ');
     SQL.Add('       0 AS VAL_DEBITO,   ');
     SQL.Add('       COALESCE(CLIENTES.RENDA, 0) AS VAL_RENDA,   ');
     SQL.Add('       0 AS COD_CONVENIO,   ');
     SQL.Add('       0 AS COD_STATUS_PDV,   ');
     SQL.Add('          ');
     SQL.Add('       CASE   ');
     SQL.Add('           WHEN CLIENTES.FLAGFISICA = ''N'' THEN ''S''   ');
     SQL.Add('           ELSE ''N''   ');
     SQL.Add('       END AS FLG_EMPRESA,   ');
     SQL.Add('          ');
     SQL.Add('       ''N'' AS FLG_CONVENIO,   ');
     SQL.Add('       ''N'' AS MICRO_EMPRESA,   ');
     SQL.Add('       CLIENTES.DATCAD AS DTA_CADASTRO,   ');
     SQL.Add('       COALESCE(CLIENTES.NUMEROLOGRADOURO, ''S/N'') AS NUM_ENDERECO,   ');
     SQL.Add('       COALESCE(CLIENTES.IDENTIDADE, '''') AS NUM_RG,   ');
     SQL.Add('      ');
     SQL.Add('       CASE   ');
     SQL.Add('           WHEN CLIENTES.ESTADOCIVIL = ''C'' THEN 1   ');
     SQL.Add('           WHEN CLIENTES.ESTADOCIVIL = ''D'' THEN 3   ');
     SQL.Add('           WHEN CLIENTES.ESTADOCIVIL = ''S'' THEN 0   ');
     SQL.Add('           WHEN CLIENTES.ESTADOCIVIL = ''V'' THEN 2   ');
     SQL.Add('           ELSE 0   ');
     SQL.Add('       END AS FLG_EST_CIVIL,   ');
     SQL.Add('          ');
     SQL.Add('       '''' AS NUM_CELULAR,   ');
     SQL.Add('       '''' AS DTA_ALTERACAO,   ');
     //SQL.Add('       COALESCE(CAST(CLIENTES.OBS AS VARCHAR(5000)), '''') AS DES_OBSERVACAO,   ');
     SQL.Add('       '''' AS DES_OBSERVACAO,    ');
     SQL.Add('       COALESCE(CLIENTES.COMPLEMENTOLOGRADOURO, ''A DEFINIR'') AS DES_COMPLEMENTO,   ');
     SQL.Add('       COALESCE(CLIENTES.EMAIL, '''') AS DES_EMAIL,   ');
     SQL.Add('       CLIENTES.NOMECLI AS DES_FANTASIA,   ');
     SQL.Add('       CLIENTES.DATNASC AS DTA_NASCIMENTO,   ');
     SQL.Add('       COALESCE(CLIENTES.FILIACAO, '''') AS DES_PAI,   ');
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
     SQL.Add('       CLIENTE AS CLIENTES   ');



    Open;
    First;
    NumLinha := 0;
    //CODIGO_CLIENTE := 0;
    TotalCont := SetCountTotal(SQL.Text);

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);
           Layout.SetValues(QryPrincipal, NumLinha, TotalCont);
//      Layout.SetValues(QryPrincipal, NumLinha, RecordCount);


//      with QryGeraCodigoCliente do
//      begin
//        Inc(CODIGO_CLIENTE);
//        Params.ParamByName('COD_CLIENTE').Value := CODIGO_CLIENTE;
//        Params.ParamByName('NUM_CGC').Value := Layout.FieldByName('NUM_CGC').AsString;
//        Layout.FieldByName('COD_CLIENTE').AsInteger := Params.ParamByName('COD_CLIENTE').Value;
//        //ExecSQL();
//      end;

      Layout.FieldByName('NUM_CGC').AsString := StrRetNums(Layout.FieldByName('NUM_CGC').AsString);

      //if StrRetNums(Layout.FieldByName('NUM_RG').AsString) = '' then
        //Layout.FieldByName('NUM_RG').AsString := ''
      //else
        //Layout.FieldByName('NUM_RG').AsString := StrRetNums(Layout.FieldByName('NUM_RG').AsString);

      Layout.FieldByName('NUM_CEP').AsString := StrRetNums(Layout.FieldByName('NUM_CEP').AsString);

      //if QryPrincipal.FieldByName('NUM_INSC_EST').AsString <> 'ISENTO' then
         //Layout.FieldByName('NUM_INSC_EST').AsString := StrRetNums(Layout.FieldByName('NUM_INSC_EST').AsString);

      if QryPrincipal.FieldByName('DTA_CADASTRO').AsString <> '' then
        Layout.FieldByName('DTA_CADASTRO').AsString := FieldByName('DTA_CADASTRO').AsString;

      if QryPrincipal.FieldByName('DTA_ALTERACAO').AsString <> '' then
        Layout.FieldByName('DTA_ALTERACAO').AsString := FieldByName('DTA_ALTERACAO').AsString;

      if QryPrincipal.FieldByName('DTA_NASCIMENTO').AsString <> '' then
        Layout.FieldByName('DTA_NASCIMENTO').AsString := FieldByName('DTA_NASCIMENTO').AsString;



      //Layout.FieldByName('NUM_FONE').AsString := StrRetNums( FieldByName('NUM_FONE').AsString );

      if Layout.FieldByName('FLG_EMPRESA').AsString = 'S' then
      begin
        if not ValidaCGC(Layout.FieldByName('NUM_CGC').AsString) then
          Layout.FieldByName('NUM_CGC').AsString := '';
      end
      else
        if not ValidaCpf(Layout.FieldByName('NUM_CGC').AsString) then
          Layout.FieldByName('NUM_CGC').AsString := '';

      //Layout.FieldByName('DES_OBSERVACAO').AsString := StrReplace(StrLBReplace(FieldByName('DES_OBSERVACAO').AsString), '\n', '');
      //Layout.FieldByName('DES_EMAIL').AsString := StrReplace(StrLBReplace(FieldByName('DES_EMAIL').AsString), '\n', '');

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;

end;

procedure TFrmSmPanda.GerarCodigoBarras;
var
 count, NEW_CODPROD, TotalCount : Integer;
 cod_antigo, codbarras : string;
 QryGeraCodigoProduto : TSQLQuery;

begin
  inherited;

//  QryGeraCodigoProduto := TSQLQuery.Create(FrmProgresso);
//  with QryGeraCodigoProduto do
//  begin
//    SQLConnection := ScnBanco;

//    SQL.Clear;
//    SQL.Add('ALTER TABLE TAB_BARRAS_AUX ');
//    SQL.Add('ADD CODIGO_PRODUTO INT DEFAULT NULL; ');
//
//    try
//      ExecSQL;
//    except
//    end;

//    SQL.Clear;
//    SQL.Add('UPDATE TAB_BARRAS_AUX ');
//    SQL.Add('SET CODIGO_PRODUTO = :COD_PRODUTO  ');
//    SQL.Add('WHERE COD_BARRA_AUX = :COD_EAN ');
//    //SQL.Add('AND CHAR_LENGTH(COD_MATERIAL) >= 8 ');
//
//    try
//      ExecSQL;
//    except
//    end;
////
//  end;




  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

     SQL.Add('   SELECT   ');
     SQL.Add('       CODIGO_PRODUTO AS COD_PRODUTO,   ');
     SQL.Add('       COD_BARRA_AUX AS COD_EAN   ');
     SQL.Add('   FROM   ');
     SQL.Add('       TAB_BARRAS_AUX   ');
     SQL.Add('   WHERE COD_BARRA_AUX NOT IN (   ');
     SQL.Add('       SELECT DISTINCT   ');
     SQL.Add('           COD_EAN   ');
     SQL.Add('       FROM   ');
     SQL.Add('           TAB_CODIGO_BARRA   ');
     SQL.Add('   )   ');



    Open;
    First;
    NumLinha := 0;
    TotalCount := SetCountTotal(SQL.Text);
    //NEW_CODPROD := 721000;

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);
//      Inc(NEW_CODPROD);

      Layout.SetValues(QryPrincipal, NumLinha, TotalCount);



//      if Layout.FieldByName('COD_PRODUTO').AsString = '0' then
//      begin
//        with QryGeraCodigoProduto do
//        begin
//          //Inc(COD_PROD);
//          Params.ParamByName('COD_PRODUTO').Value := NEW_CODPROD;
//          Params.ParamByName('COD_EAN').Value := Layout.FieldByName('COD_EAN').AsString;
//          Layout.FieldByName('COD_PRODUTO').AsInteger := Params.ParamByName('COD_PRODUTO').Value;
//          ExecSQL();
//        end;
//      end;

//      if Layout.FieldByName('COD_PRODUTO').AsString = '0' then
//      begin
//        Layout.FieldByName('COD_PRODUTO').AsInteger := NEW_CODPROD;
//      end;


      Layout.FieldByName('COD_PRODUTO').AsString := GerarPLU( Layout.FieldByName('COD_PRODUTO').AsString );

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

procedure TFrmSmPanda.GerarComposicao;
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

procedure TFrmSmPanda.GerarCondPagCli;
begin
  inherited;

  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

     SQL.Add('   SELECT    ');
     SQL.Add('       CLIENTES.CODCLI AS COD_CLIENTE,    ');
     SQL.Add('       30 AS NUM_CONDICAO,    ');
     SQL.Add('       2 AS COD_CONDICAO,    ');
     SQL.Add('       1 AS COD_ENTIDADE    ');
     SQL.Add('   FROM    ');
     SQL.Add('       CLIENTE AS CLIENTES      ');
     SQL.Add('      ');


    Open;

    First;
    TotalCont := SetCountTotal(SQL.Text);
    NumLinha := 0;



    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);

      Layout.SetValues(QryPrincipal, NumLinha, TotalCont);
      //Layout.SetValues(QryPrincipal, NumLinha, RecordCount);

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;

end;

procedure TFrmSmPanda.GerarCondPagForn;
//var
//  COD_FORNECEDOR : Integer;
begin
  inherited;

  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

     SQL.Add('   SELECT      ');
     SQL.Add('       CAST(FORNECEDOR.CODFORN AS INTEGER) AS COD_FORNECEDOR,   ');
     SQL.Add('       30 AS NUM_CONDICAO,      ');
     SQL.Add('       2 AS COD_CONDICAO,      ');
     SQL.Add('       8 AS COD_ENTIDADE,      ');
     SQL.Add('       COALESCE(REPLACE(REPLACE(REPLACE(FORNECEDOR.CNPJ, ''-'', ''''), ''.'', ''''), ''/'', ''''), '''') AS NUM_CGC      ');
     SQL.Add('   FROM      ');
     SQL.Add('       FORNECEDOR   ');
     SQL.Add('   WHERE FORNECEDOR.CNPJ NOT IN (   ');
     SQL.Add('       SELECT DISTINCT   ');
     SQL.Add('           COALESCE(REPLACE(REPLACE(REPLACE(TAB_FORNECEDOR_AUX.NUM_CGC_AUX, ''-'', ''''), ''.'', ''''), ''/'', ''''), '''')   ');
     SQL.Add('       FROM   ');
     SQL.Add('           TAB_FORNECEDOR_AUX   ');
     SQL.Add('   )   ');







    Open;

    First;
    TotalCont := SetCountTotal(SQL.Text);
    NumLinha := 0;
//    COD_FORNECEDOR := 0;

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);

      Layout.SetValues(QryPrincipal, NumLinha, TotalCont);

//      Inc(COD_FORNECEDOR);
//      Layout.FieldByName('COD_FORNECEDOR').AsInteger := COD_FORNECEDOR;

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;

end;

procedure TFrmSmPanda.GerarDecomposicao;
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

procedure TFrmSmPanda.GerarDivisaoForn;
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

procedure TFrmSmPanda.GerarFinanceiro(Tipo, Situacao: Integer);
begin
  inherited;
  if Tipo = 1 then
    GerarFinanceiroPagar(IntToStr(Situacao));

  if Tipo = 2 then
    GerarFinanceiroReceber(IntToStr(Situacao));

  if Tipo = 3 then
    GerarFinanceiroReceberCartao;

end;

procedure TFrmSmPanda.GerarFinanceiroPagar(Aberto: String);
var
   TotalCount : Integer;
   cgc: string;
begin
  inherited;
  with QryPrincipal do
  begin
    Close;
    SQL.Clear;
    if Aberto = '1' then
    begin
        //ABERTO
       SQL.Add('   SELECT DISTINCT  ');
       SQL.Add('       1 AS TIPO_PARCEIRO,   ');
       SQL.Add('       CAST(PAGAR.CODFORN AS INTEGER) AS COD_PARCEIRO,   ');
       SQL.Add('       0 AS TIPO_CONTA,   ');
       SQL.Add('       8 AS COD_ENTIDADE,   ');
       SQL.Add('       PAGAR.NUMDOC AS NUM_DOCTO,   ');
       SQL.Add('       999 AS COD_BANCO,   ');
       SQL.Add('       '''' AS DES_BANCO,   ');
       SQL.Add('       PAGAR.DATENTR AS DTA_EMISSAO,   ');
       SQL.Add('       PAGAR.DATVENC AS DTA_VENCIMENTO,   ');
       SQL.Add('       PAGAR.VALOR AS VAL_PARCELA,   ');
       SQL.Add('       0 AS VAL_JUROS,   ');
       SQL.Add('       0 AS VAL_DESCONTO,   ');
       SQL.Add('       ''N'' AS FLG_QUITADO,   ');
       SQL.Add('       '''' AS DTA_QUITADA,   ');
       SQL.Add('       998 AS COD_CATEGORIA,   ');
       SQL.Add('       998 AS COD_SUBCATEGORIA,   ');
       SQL.Add('       SUBSTRING(PAGAR.PARCELA FROM 1 FOR 2) AS NUM_PARCELA,   ');
       SQL.Add('       PARCELAS.QTD_PARCELA AS QTD_PARCELA,   ');
       SQL.Add('       1 AS COD_LOJA,   ');
       SQL.Add('       COALESCE(REPLACE(REPLACE(REPLACE(FORNECEDOR.CNPJ, ''-'', ''''), ''.'', ''''), ''/'', ''''), '''') AS NUM_CGC,   ');
       SQL.Add('       0 AS NUM_BORDERO,   ');
       SQL.Add('       PAGAR.NUMDOC AS NUM_NF,   ');
       SQL.Add('       '''' AS NUM_SERIE_NF,   ');
       SQL.Add('       PARCELAS.VAL_TOTAL_NF AS VAL_TOTAL_NF,   ');
       SQL.Add('       '''' AS DES_OBSERVACAO,   ');
       SQL.Add('       COALESCE(CAST(PAGAR.CODCAIXAS AS INTEGER), 1) AS NUM_PDV,   ');
       SQL.Add('       '''' AS NUM_CUPOM_FISCAL,   ');
       SQL.Add('       0 AS COD_MOTIVO,   ');
       SQL.Add('       0 AS COD_CONVENIO,   ');
       SQL.Add('       0 AS COD_BIN,   ');
       SQL.Add('       '''' AS DES_BANDEIRA,   ');
       SQL.Add('       '''' AS DES_REDE_TEF,   ');
       SQL.Add('       0 AS VAL_RETENCAO,   ');
       SQL.Add('       2 AS COD_CONDICAO,   ');
       SQL.Add('       '''' AS DTA_PAGTO,   ');
       SQL.Add('       PAGAR.DATENTR AS DTA_ENTRADA,   ');
       SQL.Add('       '''' AS NUM_NOSSO_NUMERO,   ');
       SQL.Add('       '''' AS COD_BARRA,   ');
       SQL.Add('       ''N'' AS FLG_BOLETO_EMIT,   ');
       SQL.Add('       '''' AS NUM_CGC_CPF_TITULAR,   ');
       SQL.Add('       '''' AS DES_TITULAR,   ');
       SQL.Add('       30 AS NUM_CONDICAO,   ');
       SQL.Add('       0 AS VAL_CREDITO,   ');
       SQL.Add('       999 AS COD_BANCO_PGTO,   ');
       SQL.Add('       ''PAGTO'' AS DES_CC,   ');
       SQL.Add('       0 AS COD_BANDEIRA,   ');
       SQL.Add('       '''' AS DTA_PRORROGACAO,   ');
       SQL.Add('       1 AS NUM_SEQ_FIN,   ');
       SQL.Add('       0 AS COD_COBRANCA,   ');
       SQL.Add('       '''' AS DTA_COBRANCA,   ');
       SQL.Add('       ''N'' AS FLG_ACEITE,   ');
       SQL.Add('       0 AS TIPO_ACEITE   ');
       SQL.Add('   FROM   ');
       SQL.Add('       CONTAPAGAR AS PAGAR   ');
       SQL.Add('   LEFT JOIN (   ');
       SQL.Add('       SELECT   ');
       SQL.Add('           CONTAPAGAR.NUMDOC,   ');
       SQL.Add('           COUNT(*) AS QTD_PARCELA,   ');
       SQL.Add('           SUM(CONTAPAGAR.VALOR) AS VAL_TOTAL_NF   ');
       SQL.Add('       FROM   ');
       SQL.Add('           CONTAPAGAR   ');
       SQL.Add('       GROUP BY   ');
       SQL.Add('           CONTAPAGAR.NUMDOC   ');
       SQL.Add('   ) AS PARCELAS   ');
       SQL.Add('   ON PAGAR.NUMDOC = PARCELAS.NUMDOC   ');
       SQL.Add('   LEFT JOIN FORNECEDOR ON FORNECEDOR.CODFORN = PAGAR.CODFORN   ');
       SQL.Add('   WHERE PAGAR.FLAGPAGO = ''N''   ');
       SQL.Add('   AND PAGAR.NUMDOC IS NOT NULL   ');
       SQL.Add('   AND PAGAR.FLAGFORN = ''Y''   ');
       SQL.Add('   AND FORNECEDOR.CNPJ NOT IN (   ');
       SQL.Add('       SELECT DISTINCT   ');
       SQL.Add('           COALESCE(REPLACE(REPLACE(REPLACE(TAB_FORNECEDOR_AUX.NUM_CGC_AUX, ''-'', ''''), ''.'', ''''), ''/'', ''''), '''')   ');
       SQL.Add('       FROM   ');
       SQL.Add('           TAB_FORNECEDOR_AUX   ');
       SQL.Add('   )   ');
       SQL.Add('AND');
       SQL.Add('    PAGAR.DATENTR >= :INI ');
       SQL.Add('AND');
       SQL.Add('    PAGAR.DATENTR <= :FIM ');
       ParamByName('INI').AsDate := DtpInicial.Date;
       ParamByName('FIM').AsDate := DtpFinal.Date;

    end
    else
    begin
      //QUITADO
       SQL.Add('   SELECT DISTINCT  ');
       SQL.Add('       1 AS TIPO_PARCEIRO,   ');
       SQL.Add('       CAST(PAGAR.CODFORN AS INTEGER) AS COD_PARCEIRO,   ');
       SQL.Add('       0 AS TIPO_CONTA,   ');
       SQL.Add('       8 AS COD_ENTIDADE,   ');
       SQL.Add('       PAGAR.NUMDOC AS NUM_DOCTO,   ');
       SQL.Add('       999 AS COD_BANCO,   ');
       SQL.Add('       '''' AS DES_BANCO,   ');
       SQL.Add('       PAGAR.DATENTR AS DTA_EMISSAO,   ');
       SQL.Add('       PAGAR.DATVENC AS DTA_VENCIMENTO,   ');
       SQL.Add('       PAGAR.VALOR AS VAL_PARCELA,   ');
       SQL.Add('       0 AS VAL_JUROS,   ');
       SQL.Add('       0 AS VAL_DESCONTO,   ');
       SQL.Add('       ''S'' AS FLG_QUITADO,   ');
       SQL.Add('       PAGAR.DATPAG AS DTA_QUITADA,   ');
       SQL.Add('       998 AS COD_CATEGORIA,   ');
       SQL.Add('       998 AS COD_SUBCATEGORIA,   ');
       SQL.Add('       SUBSTRING(PAGAR.PARCELA FROM 1 FOR 2) AS NUM_PARCELA,   ');
       SQL.Add('       PARCELAS.QTD_PARCELA AS QTD_PARCELA,   ');
       SQL.Add('       1 AS COD_LOJA,   ');
       SQL.Add('       COALESCE(REPLACE(REPLACE(REPLACE(FORNECEDOR.CNPJ, ''-'', ''''), ''.'', ''''), ''/'', ''''), '''') AS NUM_CGC,   ');
       SQL.Add('       0 AS NUM_BORDERO,   ');
       SQL.Add('       PAGAR.NUMDOC AS NUM_NF,   ');
       SQL.Add('       '''' AS NUM_SERIE_NF,   ');
       SQL.Add('       PARCELAS.VAL_TOTAL_NF AS VAL_TOTAL_NF,   ');
       SQL.Add('       '''' AS DES_OBSERVACAO,   ');
       SQL.Add('       COALESCE(CAST(PAGAR.CODCAIXAS AS INTEGER), 1) AS NUM_PDV,   ');
       SQL.Add('       '''' AS NUM_CUPOM_FISCAL,   ');
       SQL.Add('       0 AS COD_MOTIVO,   ');
       SQL.Add('       0 AS COD_CONVENIO,   ');
       SQL.Add('       0 AS COD_BIN,   ');
       SQL.Add('       '''' AS DES_BANDEIRA,   ');
       SQL.Add('       '''' AS DES_REDE_TEF,   ');
       SQL.Add('       0 AS VAL_RETENCAO,   ');
       SQL.Add('       2 AS COD_CONDICAO,   ');
       SQL.Add('       PAGAR.DATPAG AS DTA_PAGTO,   ');
       SQL.Add('       PAGAR.DATENTR AS DTA_ENTRADA,   ');
       SQL.Add('       '''' AS NUM_NOSSO_NUMERO,   ');
       SQL.Add('       '''' AS COD_BARRA,   ');
       SQL.Add('       ''N'' AS FLG_BOLETO_EMIT,   ');
       SQL.Add('       '''' AS NUM_CGC_CPF_TITULAR,   ');
       SQL.Add('       '''' AS DES_TITULAR,   ');
       SQL.Add('       30 AS NUM_CONDICAO,   ');
       SQL.Add('       0 AS VAL_CREDITO,   ');
       SQL.Add('       999 AS COD_BANCO_PGTO,   ');
       SQL.Add('       ''PAGTO'' AS DES_CC,   ');
       SQL.Add('       0 AS COD_BANDEIRA,   ');
       SQL.Add('       '''' AS DTA_PRORROGACAO,   ');
       SQL.Add('       1 AS NUM_SEQ_FIN,   ');
       SQL.Add('       0 AS COD_COBRANCA,   ');
       SQL.Add('       '''' AS DTA_COBRANCA,   ');
       SQL.Add('       ''N'' AS FLG_ACEITE,   ');
       SQL.Add('       0 AS TIPO_ACEITE   ');
       SQL.Add('   FROM   ');
       SQL.Add('       CONTAPAGAR AS PAGAR   ');
       SQL.Add('   LEFT JOIN (   ');
       SQL.Add('       SELECT   ');
       SQL.Add('           CONTAPAGAR.NUMDOC,   ');
       SQL.Add('           COUNT(*) AS QTD_PARCELA,   ');
       SQL.Add('           SUM(CONTAPAGAR.VALOR) AS VAL_TOTAL_NF   ');
       SQL.Add('       FROM   ');
       SQL.Add('           CONTAPAGAR   ');
       SQL.Add('       GROUP BY   ');
       SQL.Add('           CONTAPAGAR.NUMDOC   ');
       SQL.Add('   ) AS PARCELAS   ');
       SQL.Add('   ON PAGAR.NUMDOC = PARCELAS.NUMDOC   ');
       SQL.Add('   LEFT JOIN FORNECEDOR ON FORNECEDOR.CODFORN = PAGAR.CODFORN   ');
       SQL.Add('   WHERE PAGAR.FLAGPAGO = ''Y''   ');
       SQL.Add('   AND PAGAR.NUMDOC IS NOT NULL   ');
       SQL.Add('   AND PAGAR.FLAGFORN = ''Y''   ');
       SQL.Add('AND');
       SQL.Add('    PAGAR.DATENTR >= :INI ');
       SQL.Add('AND');
       SQL.Add('    PAGAR.DATENTR <= :FIM ');
       ParamByName('INI').AsDate := DtpInicial.Date;
       ParamByName('FIM').AsDate := DtpFinal.Date;

    end;


    Open;
    First;

    if( Aberto = '1' ) then
      //TotalCount := SetCountTotal(SQL.Text)
      TotalCount := SetCountTotal(SQL.Text, ParamByName('INI').AsString, ParamByName('FIM').AsString )
    else
      TotalCount := SetCountTotal(SQL.Text, ParamByName('INI').AsString, ParamByName('FIM').AsString );
//    TotalCount := SetCountTotal(SQL.Text);

    NumLinha := 0;

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
//            Layout.FieldByName('COD_PARCEIRO').AsInteger := Layout.FieldByName('COD_PARCEIRO').AsInteger + 1000
//           else
//            Layout.FieldByName('COD_PARCEIRO').AsInteger := 0;
//         end
//         else
//         begin
//            if( not CPFEValido(cgc) ) then
//               Layout.FieldByName('COD_PARCEIRO').AsInteger := Layout.FieldByName('COD_PARCEIRO').AsInteger + 1000
//            else
//               Layout.FieldByName('COD_PARCEIRO').AsInteger := 0;
//         end;
//      end;

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

procedure TFrmSmPanda.GerarFinanceiroReceber(Aberto: String);
var
   TotalCount : Integer;
   cgc : string;
begin
  inherited;
  with QryPrincipal do
  begin
    Close;
    SQL.Clear;
    if Aberto = '1' then
    begin
    // ABERTO
     SQL.Add('   SELECT DISTINCT  ');
     SQL.Add('       0 AS TIPO_PARCEIRO,   ');
     SQL.Add('       RECEBER.CODCLI AS COD_PARCEIRO,   ');
     SQL.Add('       1 AS TIPO_CONTA,   ');
     SQL.Add('       8 AS COD_ENTIDADE,   ');
     SQL.Add('          ');
     SQL.Add('       CASE   ');
     SQL.Add('           WHEN RECEBER.NUMDOC LIKE ''%/%'' THEN SUBSTRING(RECEBER.NUMDOC FROM 1 FOR 5)   ');
     SQL.Add('           ELSE RECEBER.NUMDOC   ');
     SQL.Add('       END AS NUM_DOCTO,   ');
     SQL.Add('          ');
     SQL.Add('       999 AS COD_BANCO,   ');
     SQL.Add('       '''' AS DES_BANCO,   ');
     SQL.Add('       RECEBER.DATENTR AS DTA_EMISSAO,   ');
     SQL.Add('       RECEBER.DATVENC AS DTA_VENCIMENTO,   ');
     SQL.Add('       RECEBER.VALOR AS VAL_PARCELA,   ');
     SQL.Add('       COALESCE(RECEBER.JUROS, 0) AS VAL_JUROS,   ');
     SQL.Add('       0 AS VAL_DESCONTO,   ');
     SQL.Add('       ''N'' AS FLG_QUITADO,   ');
     SQL.Add('       '''' AS DTA_QUITADA,   ');
     SQL.Add('       998 AS COD_CATEGORIA,   ');
     SQL.Add('       998 AS COD_SUBCATEGORIA,   ');
     SQL.Add('       SUBSTRING(RECEBER.PARCELA FROM 1 FOR 2) AS NUM_PARCELA,   ');
     SQL.Add('       PARCELAS.QTD_PARCELA AS QTD_PARCELA,   ');
     SQL.Add('       1 AS COD_LOJA,   ');
     SQL.Add('       COALESCE(REPLACE(REPLACE(REPLACE(CLIENTE.CNPJ, ''-'', ''''), ''.'', ''''), ''/'', ''''), COALESCE(CLIENTE.CPF, '''')) AS NUM_CGC,   ');
     SQL.Add('       0 AS NUM_BORDERO,   ');
     SQL.Add('          ');
     SQL.Add('       CASE   ');
     SQL.Add('           WHEN RECEBER.NUMDOC LIKE ''%/%'' THEN SUBSTRING(RECEBER.NUMDOC FROM 1 FOR 5)   ');
     SQL.Add('           ELSE RECEBER.NUMDOC   ');
     SQL.Add('       END AS NUM_NF,   ');
     SQL.Add('          ');
     SQL.Add('       '''' AS NUM_SERIE_NF,   ');
     SQL.Add('       PARCELAS.VAL_TOTAL_NF AS VAL_TOTAL_NF,   ');
     SQL.Add('       '''' AS DES_OBSERVACAO,   ');
     SQL.Add('       1 AS NUM_PDV,   ');
     SQL.Add('       '''' AS NUM_CUPOM_FISCAL,   ');
     SQL.Add('       0 AS COD_MOTIVO,   ');
     SQL.Add('       0 AS COD_CONVENIO,   ');
     SQL.Add('       0 AS COD_BIN,   ');
     SQL.Add('       '''' AS DES_BANDEIRA,   ');
     SQL.Add('       '''' AS DES_REDE_TEF,   ');
     SQL.Add('       0 AS VAL_RETENCAO,   ');
     SQL.Add('       2 AS COD_CONDICAO,   ');
     SQL.Add('       '''' AS DTA_PAGTO,   ');
     SQL.Add('       RECEBER.DATENTR AS DTA_ENTRADA,   ');
     SQL.Add('       '''' AS NUM_NOSSO_NUMERO,   ');
     SQL.Add('       '''' AS COD_BARRA,   ');
     SQL.Add('       ''N'' AS FLG_BOLETO_EMIT,   ');
     SQL.Add('       '''' AS NUM_CGC_CPF_TITULAR,   ');
     SQL.Add('       '''' AS DES_TITULAR,   ');
     SQL.Add('       30 AS NUM_CONDICAO,   ');
     SQL.Add('       0 AS VAL_CREDITO,   ');
     SQL.Add('       ''002'' AS COD_BANCO_PGTO,   ');
     SQL.Add('       ''RECEBTO'' AS DES_CC,   ');
     SQL.Add('       0 AS COD_BANDEIRA,   ');
     SQL.Add('       '''' AS DTA_PRORROGACAO,   ');
     SQL.Add('       1 AS NUM_SEQ_FIN,   ');
     SQL.Add('       0 AS COD_COBRANCA,   ');
     SQL.Add('       '''' AS DTA_COBRANCA,   ');
     SQL.Add('       ''N'' AS FLG_ACEITE,   ');
     SQL.Add('       0 AS TIPO_ACEITE   ');
     SQL.Add('   FROM   ');
     SQL.Add('       CONTARECEBER AS RECEBER   ');
     SQL.Add('   LEFT JOIN (   ');
     SQL.Add('       SELECT   ');
     SQL.Add('           CASE   ');
     SQL.Add('               WHEN NUMDOC LIKE ''%/%'' THEN SUBSTRING(NUMDOC FROM 1 FOR 5)   ');
     SQL.Add('               ELSE NUMDOC   ');
     SQL.Add('           END AS NUM_DOC,   ');
     SQL.Add('      ');
     SQL.Add('           COUNT(*) AS QTD_PARCELA,   ');
     SQL.Add('           SUM(CONTARECEBER.VALOR) AS VAL_TOTAL_NF   ');
     SQL.Add('       FROM   ');
     SQL.Add('           CONTARECEBER   ');
     SQL.Add('       --WHERE SUBSTRING(NUMDOC FROM 1 FOR 5) = ''12642''   ');
     SQL.Add('       GROUP BY   ');
     SQL.Add('           NUM_DOC   ');
     SQL.Add('   ) AS PARCELAS   ');
     SQL.Add('   ON    ');
     SQL.Add('       CASE   ');
     SQL.Add('           WHEN RECEBER.NUMDOC LIKE ''%/%'' THEN SUBSTRING(RECEBER.NUMDOC FROM 1 FOR 5)   ');
     SQL.Add('           ELSE RECEBER.NUMDOC   ');
     SQL.Add('       END = PARCELAS.NUM_DOC   ');
     SQL.Add('   LEFT JOIN CLIENTE ON CLIENTE.CODCLI = RECEBER.CODCLI   ');
     SQL.Add('   WHERE RECEBER.FLAGPAGO = ''N''   ');
     SQL.Add('   AND RECEBER.NUMDOC IS NOT NULL   ');
     SQL.Add('   AND RECEBER.FLAGCLI = ''Y''   ');
     SQL.Add('AND RECEBER.DATPAG >= :INI ');
     SQL.Add('AND RECEBER.DATPAG <= :FIM ');

     ParamByName('INI').AsDate := DtpInicial.Date;
     ParamByName('FIM').AsDate := DtpFinal.Date;


    end
    else
    begin
      // QUITADO

       SQL.Add('   SELECT DISTINCT  ');
       SQL.Add('       0 AS TIPO_PARCEIRO,   ');
       SQL.Add('       RECEBER.CODCLI AS COD_PARCEIRO,   ');
       SQL.Add('       1 AS TIPO_CONTA,   ');
       SQL.Add('       8 AS COD_ENTIDADE,   ');
       SQL.Add('          ');
       SQL.Add('       CASE   ');
       SQL.Add('           WHEN RECEBER.NUMDOC LIKE ''%/%'' THEN SUBSTRING(RECEBER.NUMDOC FROM 1 FOR 6)   ');
       SQL.Add('           ELSE RECEBER.NUMDOC   ');
       SQL.Add('       END AS NUM_DOCTO,   ');
       SQL.Add('          ');
       SQL.Add('       999 AS COD_BANCO,   ');
       SQL.Add('       '''' AS DES_BANCO,   ');
       SQL.Add('       RECEBER.DATENTR AS DTA_EMISSAO,   ');
       SQL.Add('       RECEBER.DATVENC AS DTA_VENCIMENTO,   ');
       SQL.Add('       RECEBER.VALOR AS VAL_PARCELA,   ');
       SQL.Add('       COALESCE(RECEBER.JUROS, 0) AS VAL_JUROS,   ');
       SQL.Add('       0 AS VAL_DESCONTO,   ');
       SQL.Add('       ''S'' AS FLG_QUITADO,   ');
       SQL.Add('       RECEBER.DATPAG AS DTA_QUITADA,   ');
       SQL.Add('       998 AS COD_CATEGORIA,   ');
       SQL.Add('       998 AS COD_SUBCATEGORIA,   ');
       SQL.Add('       SUBSTRING(RECEBER.PARCELA FROM 1 FOR 2) AS NUM_PARCELA,   ');
       SQL.Add('       PARCELAS.QTD_PARCELA AS QTD_PARCELA,   ');
       SQL.Add('       1 AS COD_LOJA,   ');
       SQL.Add('       COALESCE(REPLACE(REPLACE(REPLACE(CLIENTE.CNPJ, ''-'', ''''), ''.'', ''''), ''/'', ''''), COALESCE(CLIENTE.CPF, '''')) AS NUM_CGC,   ');
       SQL.Add('       0 AS NUM_BORDERO,   ');
       SQL.Add('          ');
       SQL.Add('       CASE   ');
       SQL.Add('           WHEN RECEBER.NUMDOC LIKE ''%/%'' THEN SUBSTRING(RECEBER.NUMDOC FROM 1 FOR 6)   ');
       SQL.Add('           ELSE RECEBER.NUMDOC   ');
       SQL.Add('       END AS NUM_NF,   ');
       SQL.Add('          ');
       SQL.Add('       '''' AS NUM_SERIE_NF,   ');
       SQL.Add('       PARCELAS.VAL_TOTAL_NF AS VAL_TOTAL_NF,   ');
       SQL.Add('       '''' AS DES_OBSERVACAO,   ');
       SQL.Add('       1 AS NUM_PDV,   ');
       SQL.Add('       '''' AS NUM_CUPOM_FISCAL,   ');
       SQL.Add('       0 AS COD_MOTIVO,   ');
       SQL.Add('       0 AS COD_CONVENIO,   ');
       SQL.Add('       0 AS COD_BIN,   ');
       SQL.Add('       '''' AS DES_BANDEIRA,   ');
       SQL.Add('       '''' AS DES_REDE_TEF,   ');
       SQL.Add('       0 AS VAL_RETENCAO,   ');
       SQL.Add('       2 AS COD_CONDICAO,   ');
       SQL.Add('       RECEBER.DATPAG AS DTA_PAGTO,   ');
       SQL.Add('       RECEBER.DATENTR AS DTA_ENTRADA,   ');
       SQL.Add('       '''' AS NUM_NOSSO_NUMERO,   ');
       SQL.Add('       '''' AS COD_BARRA,   ');
       SQL.Add('       ''N'' AS FLG_BOLETO_EMIT,   ');
       SQL.Add('       '''' AS NUM_CGC_CPF_TITULAR,   ');
       SQL.Add('       '''' AS DES_TITULAR,   ');
       SQL.Add('       30 AS NUM_CONDICAO,   ');
       SQL.Add('       0 AS VAL_CREDITO,   ');
       SQL.Add('       ''002'' AS COD_BANCO_PGTO,   ');
       SQL.Add('       ''RECEBTO'' AS DES_CC,   ');
       SQL.Add('       0 AS COD_BANDEIRA,   ');
       SQL.Add('       '''' AS DTA_PRORROGACAO,   ');
       SQL.Add('       1 AS NUM_SEQ_FIN,   ');
       SQL.Add('       0 AS COD_COBRANCA,   ');
       SQL.Add('       '''' AS DTA_COBRANCA,   ');
       SQL.Add('       ''N'' AS FLG_ACEITE,   ');
       SQL.Add('       0 AS TIPO_ACEITE   ');
       SQL.Add('   FROM   ');
       SQL.Add('       CONTARECEBER AS RECEBER   ');
       SQL.Add('   LEFT JOIN (   ');
       SQL.Add('       SELECT   ');
       SQL.Add('           CASE   ');
       SQL.Add('               WHEN NUMDOC LIKE ''%/%'' THEN SUBSTRING(NUMDOC FROM 1 FOR 6)   ');
       SQL.Add('               ELSE NUMDOC   ');
       SQL.Add('           END AS NUM_DOC,   ');
       SQL.Add('      ');
       SQL.Add('           COUNT(*) AS QTD_PARCELA,   ');
       SQL.Add('           SUM(CONTARECEBER.VALOR) AS VAL_TOTAL_NF   ');
       SQL.Add('       FROM   ');
       SQL.Add('           CONTARECEBER   ');
       SQL.Add('       GROUP BY   ');
       SQL.Add('           NUM_DOC   ');
       SQL.Add('   ) AS PARCELAS   ');
       SQL.Add('   ON    ');
       SQL.Add('       CASE   ');
       SQL.Add('           WHEN RECEBER.NUMDOC LIKE ''%/%'' THEN SUBSTRING(RECEBER.NUMDOC FROM 1 FOR 6)   ');
       SQL.Add('           ELSE RECEBER.NUMDOC   ');
       SQL.Add('       END = PARCELAS.NUM_DOC   ');
       SQL.Add('   LEFT JOIN CLIENTE ON CLIENTE.CODCLI = RECEBER.CODCLI   ');
       SQL.Add('   WHERE RECEBER.FLAGPAGO = ''Y''   ');
       SQL.Add('   AND RECEBER.NUMDOC IS NOT NULL   ');
       SQL.Add('   AND RECEBER.FLAGCLI = ''Y''   ');
       SQL.Add('AND RECEBER.DATPAG >= :INI ');
       SQL.Add('AND RECEBER.DATPAG <= :FIM ');

      ParamByName('INI').AsDate := DtpInicial.Date;
      ParamByName('FIM').AsDate := DtpFinal.Date;
    end;

    Open;

    First;

    if( Aberto = '1' ) then
      //TotalCount := SetCountTotal(SQL.Text)
      TotalCount := SetCountTotal(SQL.Text, ParamByName('INI').AsString, ParamByName('FIM').AsString )
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

procedure TFrmSmPanda.GerarFinanceiroReceberCartao;
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

procedure TFrmSmPanda.GerarFornecedor;
var
   observacao, email : string;
//   COD_FORNECEDOR : Integer;
//   QryGeraCodigoFornecedor : TSQLQuery;
begin
  inherited;

//  QryGeraCodigoFornecedor := TSQLQuery.Create(FrmProgresso);
//  with QryGeraCodigoFornecedor do
//  begin
//    SQLConnection := ScnBanco;
//
//    SQL.Clear;
//    SQL.Add('ALTER TABLE EMD101 ');
//    SQL.Add('ADD CODIGO_FORNECEDOR INT DEFAULT NULL; ');
//
//    try
//      ExecSQL;
//    except
//    end;
//
//    SQL.Clear;
//    SQL.Add('UPDATE EMD101');
//    SQL.Add('SET CODIGO_FORNECEDOR = :COD_FORNECEDOR ');
//    SQL.Add('WHERE COALESCE(REPLACE(REPLACE(REPLACE(CGC_CPF, ''.'', ''''), ''/'', ''''), ''-'', ''''), '''') = :NUM_CGC ');
//    SQL.Add('AND NOME NOT LIKE ''%CONS.%''');
//    SQL.Add('AND NOME NOT LIKE ''%CONSUMIDOR%''');
//
//    try
//      ExecSQL;
//    except
//    end;
//
//  end;


  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

     SQL.Add('   SELECT   ');
     SQL.Add('       CAST(FORNECEDOR.CODFORN AS INTEGER) AS COD_FORNECEDOR,   ');
     SQL.Add('       FORNECEDOR.NOMEFORN AS DES_FORNECEDOR,   ');
     SQL.Add('       COALESCE(FORNECEDOR.RAZAOSOCIAL, FORNECEDOR.NOMEFORN) AS DES_FANTASIA,   ');
     SQL.Add('       COALESCE(REPLACE(REPLACE(REPLACE(FORNECEDOR.CNPJ, ''-'', ''''), ''.'', ''''), ''/'', ''''), '''') AS NUM_CGC,   ');
     SQL.Add('       COALESCE(FORNECEDOR.INSCR, ''ISENTO'') AS NUM_INSC_EST,   ');
     SQL.Add('       COALESCE(FORNECEDOR.ENDERECO, ''A DEFINIR'') AS DES_ENDERECO,   ');
     SQL.Add('       COALESCE(FORNECEDOR.BAIRRO, ''A DEFINIR'') AS DES_BAIRRO,   ');
     SQL.Add('       COALESCE(FORNECEDOR.CIDADE, '''') AS DES_CIDADE,   ');
     SQL.Add('       COALESCE(FORNECEDOR.ESTADO, '''') AS DES_SIGLA,   ');
     SQL.Add('       COALESCE(FORNECEDOR.CEP, '''') AS NUM_CEP,   ');
     SQL.Add('       COALESCE(FORNECEDOR.TELEFONE, '''') AS NUM_FONE,   ');
     SQL.Add('       COALESCE(FORNECEDOR.FAX, '''') AS NUM_FAX,   ');
     SQL.Add('       COALESCE(FORNECEDOR.CONTATO, '''') AS DES_CONTATO,   ');
     SQL.Add('       0 AS QTD_DIA_CARENCIA,   ');
     SQL.Add('       0 AS NUM_FREQ_VISITA,   ');
     SQL.Add('       0 AS VAL_DESCONTO,   ');
     SQL.Add('       0 AS NUM_PRAZO,   ');
     SQL.Add('       ''N'' AS ACEITA_DEVOL_MER,   ');
     SQL.Add('       ''N'' AS CAL_IPI_VAL_BRUTO,   ');
     SQL.Add('       ''N'' AS CAL_ICMS_ENC_FIN,   ');
     SQL.Add('       ''N'' AS CAL_ICMS_VAL_IPI,   ');
     SQL.Add('       ''N'' AS MICRO_EMPRESA,   ');
     SQL.Add('       FORNECEDOR.CODFORN AS COD_FORNECEDOR_ANT,   ');
     SQL.Add('       COALESCE(FORNECEDOR.NUMEROLOGRADOURO, ''S/N'') AS NUM_ENDERECO,   ');
     SQL.Add('       '''' AS DES_OBSERVACAO,   ');
     SQL.Add('       COALESCE(FORNECEDOR.EMAIL, '''') AS DES_EMAIL,   ');
     SQL.Add('       COALESCE(FORNECEDOR.WEB, '''') AS DES_WEB_SITE,   ');
     SQL.Add('       ''N'' AS FABRICANTE,   ');
     SQL.Add('       ''N'' AS FLG_PRODUTOR_RURAL,   ');
     SQL.Add('       0 AS TIPO_FRETE,   ');
     SQL.Add('       ''N'' AS FLG_SIMPLES,   ');
     SQL.Add('       ''N'' AS FLG_SUBSTITUTO_TRIB,   ');
     SQL.Add('       0 AS COD_CONTACCFORN,   ');
     SQL.Add('       ''N'' AS INATIVO,   ');
     SQL.Add('       0 AS COD_CLASSIF,   ');
     SQL.Add('       FORNECEDOR.DATCAD AS DTA_CADASTRO,   ');
     SQL.Add('       0 AS VAL_CREDITO,   ');
     SQL.Add('       0 AS VAL_DEBITO,   ');
     SQL.Add('       1 AS PED_MIN_VAL,   ');
     SQL.Add('       '''' AS DES_EMAIL_VEND,   ');
     SQL.Add('       '''' AS SENHA_COTACAO,   ');
     SQL.Add('       -1 AS TIPO_PRODUTOR,   ');
     SQL.Add('       '''' AS NUM_CELULAR   ');
     SQL.Add('   FROM   ');
     SQL.Add('       FORNECEDOR   ');
     SQL.Add('   WHERE FORNECEDOR.CNPJ NOT IN (   ');
     SQL.Add('       SELECT DISTINCT   ');
     SQL.Add('           COALESCE(REPLACE(REPLACE(REPLACE(TAB_FORNECEDOR_AUX.NUM_CGC_AUX, ''-'', ''''), ''.'', ''''), ''/'', ''''), '''')   ');
     SQL.Add('       FROM   ');
     SQL.Add('           TAB_FORNECEDOR_AUX   ');
     SQL.Add('   )   ');






    Open;

    First;
    TotalCont := SetCountTotal(SQL.Text);
    NumLinha := 0;
//    COD_FORNECEDOR := 0;

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);

      Layout.SetValues(QryPrincipal, NumLinha, TotalCont);


//      with QryGeraCodigoFornecedor do
//      begin
//        Inc(COD_FORNECEDOR);
//        Params.ParamByName('COD_FORNECEDOR').Value := COD_FORNECEDOR;
//        Params.ParamByName('NUM_CGC').Value := Layout.FieldByName('NUM_CGC').AsString;
//        Layout.FieldByName('COD_FORNECEDOR').AsInteger := Params.ParamByName('COD_FORNECEDOR').Value;
//        ExecSQL();
//      end;


       //Layout.FieldByName('DTA_CADASTRO').AsDateTime := FieldByName('DTA_CADASTRO').AsDateTime;

      //Layout.FieldByName('COD_FORNECEDOR').AsInteger := COD_FORNECEDOR;

      Layout.FieldByName('NUM_CGC').AsString := StrRetNums(Layout.FieldByName('NUM_CGC').AsString);
      //Layout.FieldByName('NUM_INSC_EST').AsString := StrRetNums(Layout.FieldByName('NUM_INSC_EST').AsString);
      Layout.FieldByName('NUM_CEP').AsString := StrRetNums(Layout.FieldByName('NUM_CEP').AsString);

//      if QryPrincipal.FieldByName('NUM_INSC_EST').AsString = '0' then
//         Layout.FieldByName('NUM_INSC_EST').AsString := 'ISENTO';
//
//      if QryPrincipal.FieldByName('NUM_INSC_EST').AsString <> 'ISENTO' then
//         Layout.FieldByName('NUM_INSC_EST').AsString := StrRetNums(Layout.FieldByName('NUM_INSC_EST').AsString);


//    if((Layout.FieldByName('COD_FORNECEDOR').AsInteger =  561 )or
//    (Layout.FieldByName('COD_FORNECEDOR').AsInteger =  623 )or
//    (Layout.FieldByName('COD_FORNECEDOR').AsInteger =  773 )or
//    (Layout.FieldByName('COD_FORNECEDOR').AsInteger =  780 )or
//    (Layout.FieldByName('COD_FORNECEDOR').AsInteger =  792 )or
//    (Layout.FieldByName('COD_FORNECEDOR').AsInteger =  794 )or
//    (Layout.FieldByName('COD_FORNECEDOR').AsInteger =  795 )or
//    (Layout.FieldByName('COD_FORNECEDOR').AsInteger =  813 )or
//    (Layout.FieldByName('COD_FORNECEDOR').AsInteger =  828 )or
//    (Layout.FieldByName('COD_FORNECEDOR').AsInteger =  843 )or
//    (Layout.FieldByName('COD_FORNECEDOR').AsInteger =  844 )or
//    (Layout.FieldByName('COD_FORNECEDOR').AsInteger =  886 )or
//    (Layout.FieldByName('COD_FORNECEDOR').AsInteger =  893 )or
//    (Layout.FieldByName('COD_FORNECEDOR').AsInteger =  910 )or
//    (Layout.FieldByName('COD_FORNECEDOR').AsInteger =  911 )or
//    (Layout.FieldByName('COD_FORNECEDOR').AsInteger =  925 )or
//    (Layout.FieldByName('COD_FORNECEDOR').AsInteger =  954 )or
//    (Layout.FieldByName('COD_FORNECEDOR').AsInteger =  1029 )or
//    (Layout.FieldByName('COD_FORNECEDOR').AsInteger =  1030 )or
//    (Layout.FieldByName('COD_FORNECEDOR').AsInteger =  1031 )or
//    (Layout.FieldByName('COD_FORNECEDOR').AsInteger =  1032 )or
//    (Layout.FieldByName('COD_FORNECEDOR').AsInteger =  1033 )or
//    (Layout.FieldByName('COD_FORNECEDOR').AsInteger =  1034 )or
//    (Layout.FieldByName('COD_FORNECEDOR').AsInteger =  1035 )or
//    (Layout.FieldByName('COD_FORNECEDOR').AsInteger =  1036 )or
//    (Layout.FieldByName('COD_FORNECEDOR').AsInteger =  1037 )or
//    (Layout.FieldByName('COD_FORNECEDOR').AsInteger =  1038 )or
//    (Layout.FieldByName('COD_FORNECEDOR').AsInteger =  1039 )or
//    (Layout.FieldByName('COD_FORNECEDOR').AsInteger =  1040 )or
//    (Layout.FieldByName('COD_FORNECEDOR').AsInteger =  1041 )or
//    (Layout.FieldByName('COD_FORNECEDOR').AsInteger =  1042 )or
//    (Layout.FieldByName('COD_FORNECEDOR').AsInteger =  1043 )or
//    (Layout.FieldByName('COD_FORNECEDOR').AsInteger =  1044 )or
//    (Layout.FieldByName('COD_FORNECEDOR').AsInteger =  1045 )or
//    (Layout.FieldByName('COD_FORNECEDOR').AsInteger =  1046 )or
//    (Layout.FieldByName('COD_FORNECEDOR').AsInteger =  1047 )or
//    (Layout.FieldByName('COD_FORNECEDOR').AsInteger =  1048 )or
//    (Layout.FieldByName('COD_FORNECEDOR').AsInteger =  1049 )or
//    (Layout.FieldByName('COD_FORNECEDOR').AsInteger =  1050 )or
//    (Layout.FieldByName('COD_FORNECEDOR').AsInteger =  1051 )or
//    (Layout.FieldByName('COD_FORNECEDOR').AsInteger =  1052 )or
//    (Layout.FieldByName('COD_FORNECEDOR').AsInteger =  1066 )or
//    (Layout.FieldByName('COD_FORNECEDOR').AsInteger =  1077 )or
//    (Layout.FieldByName('COD_FORNECEDOR').AsInteger =  1082 )or
//    (Layout.FieldByName('COD_FORNECEDOR').AsInteger =  1099 )or
//    (Layout.FieldByName('COD_FORNECEDOR').AsInteger =  1102 )or
//    (Layout.FieldByName('COD_FORNECEDOR').AsInteger =  1125 ))
//  then
//      begin
//        Layout.FieldByName('NUM')
//      end;


      if Length(Layout.FieldByName('NUM_CGC').AsString) > 11 then
      begin
        if not ValidaCGC(Layout.FieldByName('NUM_CGC').AsString) then
          Layout.FieldByName('NUM_CGC').AsString := '';
      end
      else
        if not ValidaCPF(Layout.FieldByName('NUM_CGC').AsString) then
          Layout.FieldByName('NUM_CGC').AsString := '';

      Layout.FieldByName('NUM_FONE').AsString := StrRetNums( FieldByName('NUM_FONE').AsString );

      //observacao := StrReplace(StrLBReplace(FieldByName('DES_OBSERVACAO').AsString), '\n', '');
      Layout.FieldByName('DES_EMAIL').AsString := StrReplace(StrLBReplace(FieldByName('DES_EMAIL').AsString), '\n', '');
      Layout.FieldByName('DES_EMAIL_VEND').AsString := StrReplace(StrLBReplace(FieldByName('DES_EMAIL_VEND').AsString), '\n', '');


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

procedure TFrmSmPanda.GerarGrupo;
var
   TotalCount : Integer;
begin
  inherited;

  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

     SQL.Add('   SELECT DISTINCT   ');
     SQL.Add('       999 AS COD_SECAO,   ');
     SQL.Add('       999 AS COD_GRUPO,   ');
     SQL.Add('       ''A DEFINIR'' AS DES_GRUPO,   ');
     SQL.Add('       0 AS VAL_META   ');
     SQL.Add('   FROM   ');
     SQL.Add('        TAB_PRODUTO   ');
     //SQL.Add('   LEFT JOIN TB_PRO_GR AS SECAO ON SECAO.CODIGO = PRODUTOS.COD_GR   ');

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

procedure TFrmSmPanda.GerarInfoNutricionais;
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

procedure TFrmSmPanda.GerarNCM;
var
 count, TotalCount : Integer;
begin
  inherited;
  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

     SQL.Add('   SELECT DISTINCT   ');
     SQL.Add('       0 AS COD_NCM,   ');
     SQL.Add('       ''A DEFINIR'' AS DES_NCM,   ');
     SQL.Add('       ''99999999'' AS NUM_NCM,   ');
     SQL.Add('       ''N'' AS FLG_NAO_PIS_COFINS,   ');
     SQL.Add('       -1 AS TIPO_NAO_PIS_COFINS,   ');
     SQL.Add('       999 AS COD_TAB_SPED,   ');
     SQL.Add('       ''99999999'' AS NUM_CEST,   ');
     SQL.Add('       ''SP'' AS DES_SIGLA,   ');
     SQL.Add('       12 AS COD_TRIB_ENTRADA,   ');
     SQL.Add('       12 AS COD_TRIB_SAIDA,   ');
     SQL.Add('       0 AS PER_IVA,   ');
     SQL.Add('       0 AS PER_FCP_ST   ');
     SQL.Add('   FROM   ');
     SQL.Add('       TAB_PRODUTO   ');






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

procedure TFrmSmPanda.GerarNCMUF;
var
 count, TotalCount : Integer;
begin
  inherited;
  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

     SQL.Add('   SELECT DISTINCT   ');
     SQL.Add('       0 AS COD_NCM,   ');
     SQL.Add('       ''A DEFINIR'' AS DES_NCM,   ');
     SQL.Add('       ''99999999'' AS NUM_NCM,   ');
     SQL.Add('       ''N'' AS FLG_NAO_PIS_COFINS,   ');
     SQL.Add('       -1 AS TIPO_NAO_PIS_COFINS,   ');
     SQL.Add('       999 AS COD_TAB_SPED,   ');
     SQL.Add('       ''99999999'' AS NUM_CEST,   ');
     SQL.Add('       ''SP'' AS DES_SIGLA,   ');
     SQL.Add('       12 AS COD_TRIB_ENTRADA,   ');
     SQL.Add('       12 AS COD_TRIB_SAIDA,   ');
     SQL.Add('       0 AS PER_IVA,   ');
     SQL.Add('       0 AS PER_FCP_ST   ');
     SQL.Add('   FROM   ');
     SQL.Add('       TAB_PRODUTO   ');




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

procedure TFrmSmPanda.GerarNFClientes;
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

procedure TFrmSmPanda.GerarNFFornec;
var
   TotalCount : integer;
begin
  inherited;
  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

     SQL.Add('   SELECT DISTINCT  ');
     SQL.Add('       CAST(CAPA.CODFORN AS INTEGER) AS COD_FORNECEDOR,   ');
     SQL.Add('       CAPA.NUMNOTA AS NUM_NF_FORN,   ');
     SQL.Add('       COALESCE(CAPA.SERIENOTA, '''') AS NUM_SERIE_NF,   ');
     SQL.Add('       '''' AS NUM_SUBSERIE_NF,   ');
     SQL.Add('       COALESCE(CAPA.CODCFOP, '''') AS CFOP,   ');
     SQL.Add('       0 AS TIPO_NF,   ');
     SQL.Add('       ''NFE'' AS DES_ESPECIE,   ');
     SQL.Add('       CAPA.VALORTOTALNOTA AS VAL_TOTAL_NF,   ');
     SQL.Add('       CAPA.DATAEMISSAO AS DTA_EMISSAO,   ');
     SQL.Add('       CAPA.DATA AS DTA_ENTRADA,   ');
     SQL.Add('       COALESCE(CAPA.VALORTOTALIPI, 0) AS VAL_TOTAL_IPI,   ');
     SQL.Add('       0 AS VAL_VENDA_VAREJO,   ');
     SQL.Add('       COALESCE(CAPA.VALORFRETE, 0) AS VAL_FRETE,   ');
     SQL.Add('       COALESCE(CAPA.VALORACRESCIMO, 0) AS VAL_ACRESCIMO,   ');
     SQL.Add('       COALESCE(CAPA.VALORDESCONTO, 0) AS VAL_DESCONTO,   ');
     SQL.Add('       COALESCE(REPLACE(REPLACE(REPLACE(FORNECEDOR.CNPJ, ''-'', ''''), ''.'', ''''), ''/'', ''''), '''') AS NUM_CGC,   ');
     SQL.Add('       COALESCE(CAPA.BASEICMS, 0) AS VAL_TOTAL_BC,   ');
     SQL.Add('       COALESCE(CAPA.VALORICMS, 0) AS VAL_TOTAL_ICMS,   ');
     SQL.Add('       COALESCE(CAPA.BASESUBSTTRIBUTARIA, 0) AS VAL_BC_SUBST,   ');
     SQL.Add('       COALESCE(CAPA.VALORSUBSTTRIBUTARIA, 0) AS VAL_ICMS_SUBST,   ');
     SQL.Add('       0 AS VAL_FUNRURAL,   ');
     SQL.Add('       1 AS COD_PERFIL,   ');
     SQL.Add('       0 AS VAL_DESP_ACESS,   ');
     SQL.Add('          ');
     SQL.Add('       CASE   ');
     SQL.Add('           WHEN CAPA.FLAGCANCELADA = ''N'' THEN ''N''   ');
     SQL.Add('           ELSE ''S''   ');
     SQL.Add('       END AS FLG_CANCELADO,   ');
     SQL.Add('          ');
     SQL.Add('       COALESCE(CAST(CAPA.OBS AS VARCHAR(500)), '''') AS DES_OBSERVACAO,   ');
     SQL.Add('       COALESCE(CAPA.NUMEROCHAVENFE, '''') AS NUM_CHAVE_ACESSO   ');
     SQL.Add('   FROM   ');
     SQL.Add('       MOVENTRADA AS CAPA   ');
     SQL.Add('   LEFT JOIN FORNECEDOR ON FORNECEDOR.CODFORN = CAPA.CODFORN   ');
     SQL.Add('   WHERE CAPA.CODFORN IS NOT NULL   ');
     SQL.Add('   AND CAPA.NUMNOTA IS NOT NULL   ');
     SQL.Add('   AND CAPA.DATAEMISSAO >= :INI');
     SQL.Add('   AND CAPA.DATAEMISSAO <= :FIM');
//
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

      //if Layout.FieldByName('DTA_EMISSAO').AsString <> '' then
        Layout.FieldByName('DTA_EMISSAO').AsDateTime := QryPrincipal.FieldByName('DTA_EMISSAO').AsDateTime;

      //if Layout.FieldByName('DTA_ENTRADA').AsString <> '' then
        Layout.FieldByName('DTA_ENTRADA').AsDateTime := QryPrincipal.FieldByName('DTA_ENTRADA').AsDateTime;

      Layout.FieldByName('DES_OBSERVACAO').AsString := StrLBReplace(FieldByName('DES_OBSERVACAO').AsString);

      //Layout.FieldByName('NUM_SERIE_NF').AsString =

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;
end;

procedure TFrmSmPanda.GerarNFitensClientes;
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

procedure TFrmSmPanda.GerarNFitensFornec;
var
   fornecedor, nota, serie : string;
   count, TotalCount : integer;

begin
  inherited;

  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

     SQL.Add('   SELECT DISTINCT ');
     SQL.Add('       CAST(CAPA.CODFORN AS INTEGER) AS COD_FORNECEDOR,   ');
     SQL.Add('       CAPA.NUMNOTA AS NUM_NF_FORN,   ');
     SQL.Add('       COALESCE(CAPA.SERIENOTA, '''') AS NUM_SERIE_NF,   ');
     SQL.Add('       ITENS.CODPROD AS COD_PRODUTO,   ');
     SQL.Add('       1 AS COD_TRIBUTACAO,   ');
     SQL.Add('       ITENS.QUANTIDADEEMBALAGEM AS QTD_EMBALAGEM,   ');
     SQL.Add('       ITENS.QUANTIDADE AS QTD_ENTRADA,   ');
     SQL.Add('       ''UN'' AS DES_UNIDADE,   ');
     SQL.Add('       ITENS.VALORTOTAL AS VAL_TABELA,   ');
     SQL.Add('       COALESCE(ITENS.VALORDESCONTOITEM, 0) AS VAL_DESCONTO_ITEM,   ');
     SQL.Add('       COALESCE(ITENS.VALORACRESCIMOITEM, 0) AS VAL_ACRESCIMO_ITEM,   ');
     SQL.Add('       COALESCE(ITENS.VALORIPI, 0) AS VAL_IPI_ITEM,   ');
     SQL.Add('       COALESCE(ITENS.VALORSUBSTTRIBUTARIA, 0) AS VAL_SUBST_ITEM,   ');
     SQL.Add('       0 AS VAL_FRETE_ITEM,   ');
     SQL.Add('       COALESCE(ITENS.VALORICMS, 0) AS VAL_CREDITO_ICMS,   ');
     SQL.Add('       0 AS VAL_VENDA_VAREJO,   ');
     SQL.Add('       ITENS.VALORUNITARIO AS VAL_TABELA_LIQ,   ');
     SQL.Add('       COALESCE(REPLACE(REPLACE(REPLACE(FORNECEDOR.CNPJ, ''-'', ''''), ''.'', ''''), ''/'', ''''), '''') AS NUM_CGC,   ');
     SQL.Add('       0 AS VAL_TOT_BC_ICMS,   ');
     SQL.Add('       0 AS VAL_TOT_OUTROS_ICMS,   ');
     SQL.Add('       COALESCE(CAPA.CODCFOP, '''') AS CFOP,   ');
     SQL.Add('       0 AS VAL_TOT_ISENTO,   ');
     SQL.Add('       0 AS VAL_TOT_BC_ST,   ');
     SQL.Add('       0 AS VAL_TOT_ST,   ');
     SQL.Add('       1 AS NUM_ITEM,   ');
     SQL.Add('       0 AS TIPO_IPI,   ');
     SQL.Add('       ''99999999'' AS NUM_NCM,   ');
     SQL.Add('       '''' AS DES_REFERENCIA   ');
     SQL.Add('   FROM   ');
     SQL.Add('       MOVENTRADAPROD AS ITENS   ');
     SQL.Add('   LEFT JOIN MOVENTRADA AS CAPA ON CAPA.CODMOVENTR = ITENS.CODMOVENTR   ');
     SQL.Add('   LEFT JOIN FORNECEDOR ON FORNECEDOR.CODFORN = CAPA.CODFORN   ');
     SQL.Add('   WHERE CAPA.CODFORN IS NOT NULL   ');
     SQL.Add('   AND CAPA.NUMNOTA IS NOT NULL   ');
     SQL.Add('   AND CAPA.DATAEMISSAO >= :INI  ');
     SQL.Add('   AND CAPA.DATAEMISSAO <= :FIM  ');

     //SQL.Add('   ORDER BY ITENS.ORDEM_INCLUSAO ');
//
//

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

procedure TFrmSmPanda.GerarProdForn;
var
   TotalCount, NEW_CODPROD : Integer;
begin
  inherited;

  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

     SQL.Add('   SELECT   ');
     SQL.Add('        PRODUTOS.CODIGO AS COD_PRODUTO,   ');
     SQL.Add('        PRODUTOS.COD_FORNEC AS COD_FORNECEDOR,   ');
     SQL.Add('        PROD_FORN.CODIGO_FORNEC AS DES_REFERENCIA,   ');
     SQL.Add('        COALESCE(REPLACE(REPLACE(REPLACE(FORNECEDOR.CNPJ, ''-'', ''''), ''.'', ''''), ''/'', ''''), '''') AS NUM_CGC,   ');
     SQL.Add('        0 AS COD_DIVISAO,   ');
     SQL.Add('      ');
     SQL.Add('        CASE   ');
     SQL.Add('             WHEN PRODUTOS.UNIDADE = '''' THEN ''UN''   ');
     SQL.Add('             WHEN PRODUTOS.BALANCA = ''S'' OR PRODUTOS.UNIDADE = ''KG'' THEN ''KG''   ');
     SQL.Add('             ELSE PRODUTOS.UNIDADE    ');
     SQL.Add('        END AS DES_UNIDADE_COMPRA,   ');
     SQL.Add('      ');
     SQL.Add('        PRODUTOS.UNIDADE_COMPRA AS QTD_EMBALAGEM_COMPRA,   ');
     SQL.Add('        0 AS QTD_TROCA,   ');
     SQL.Add('        ''S'' AS FLG_PREFERENCIAL   ');
     SQL.Add('   FROM   ');
     SQL.Add('        TB_PRO AS PRODUTOS   ');
     SQL.Add('   LEFT JOIN TB_PRO_COD_FOR AS PROD_FORN ON PROD_FORN.CODIGO = PRODUTOS.CODIGO   ');
     SQL.Add('   AND PROD_FORN.FORNEC = PRODUTOS.COD_FORNEC   ');
     SQL.Add('   LEFT JOIN TB_FORNEC AS FORNECEDOR ON FORNECEDOR.CODIGO = PRODUTOS.COD_FORNEC   ');
     SQL.Add('   WHERE FORNECEDOR.CODIGO > 0 ');
     SQL.Add('   AND PRODUTOS.COD_BARRA NOT LIKE ''%G%''  ');
     SQL.Add('   AND PRODUTOS.COD_BARRA NOT LIKE ''%C%'' ');

    Open;

    First;

    NumLinha := 0;

    //NEW_CODPROD := 10000;

    TotalCount := SetCountTotal(SQL.Text);

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);

      Layout.SetValues(QryPrincipal, NumLinha, TotalCount);

//      if Layout.FieldByName('COD_PRODUTO').AsString = '0' then
//      begin
//        Layout.FieldByName('COD_PRODUTO').AsInteger := NEW_CODPROD;
//      end;

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

procedure TFrmSmPanda.GerarProdLoja;
var
   TotalCount, NEW_CODPROD : integer;
begin
  inherited;

  if FlgAtualizaValVenda then
  begin
    GerarValorVenda;
    Exit;
  end;

  if FlgAtualizaCustoRep then
  begin
    GeraCustoRep;
    Exit;
  end;

  if FlgAtualizaEstoque then
  begin
    GeraEstoqueVenda;
    Exit;
  end;


  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

     SQL.Add('   SELECT   ');
     SQL.Add('       CODIGO_PRODUTO AS COD_PRODUTO,   ');
     SQL.Add('       VAL_CUSTO_REP_AUX AS VAL_CUSTO_REP,   ');
     SQL.Add('       VAL_VENDA_AUX AS VAL_VENDA,   ');
     SQL.Add('       0 AS VAL_OFERTA,   ');
     SQL.Add('       1 AS QTD_EST_VDA,   ');
     SQL.Add('       '''' AS TECLA_BALANCA,   ');
     SQL.Add('       12 AS COD_TRIBUTACAO,   ');
     SQL.Add('       0 AS VAL_MARGEM,   ');
     SQL.Add('       1 AS QTD_ETIQUETA,   ');
     SQL.Add('       12 AS COD_TRIB_ENTRADA,       ');
     SQL.Add('       ''N'' AS FLG_INATIVO,       ');
     SQL.Add('       COD_PRODUTO_AUX AS COD_PRODUTO_ANT,   ');
     SQL.Add('       00000001 AS NUM_NCM,   ');
     SQL.Add('       1 AS TIPO_NCM,   ');
     SQL.Add('       0 AS VAL_VENDA_2,   ');
     SQL.Add('       '''' AS DTA_VALIDA_OFERTA,   ');
     SQL.Add('       1 AS QTD_EST_MINIMO,   ');
     SQL.Add('       NULL AS COD_VASILHAME,   ');
     SQL.Add('       ''N'' AS FORA_LINHA,   ');
     SQL.Add('       0 AS QTD_PRECO_DIF,   ');
     SQL.Add('       0 AS VAL_FORCA_VDA,   ');
     SQL.Add('       ''9999999'' AS NUM_CEST,   ');
     SQL.Add('       0 AS PER_IVA,   ');
     SQL.Add('       0 AS PER_FCP_ST,   ');
     SQL.Add('       0 AS PER_FIDELIDADE,   ');
     SQL.Add('       '''' AS COD_INFO_RECEITA   ');
     SQL.Add('   FROM   ');
     SQL.Add('       TAB_BARRAS_AUX   ');
     SQL.Add('   WHERE COD_BARRA_AUX NOT IN (   ');
     SQL.Add('       SELECT DISTINCT   ');
     SQL.Add('           COD_EAN   ');
     SQL.Add('       FROM   ');
     SQL.Add('           TAB_CODIGO_BARRA   ');
     SQL.Add('   )   ');






    Open;
    First;
    NumLinha := 0;
    //NEW_CODPROD := 10000;

    TotalCount := SetCountTotal(SQL.Text);

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);
      //Inc(NEW_CODPROD);
      Layout.SetValues(QryPrincipal, NumLinha, TotalCount);

//      if Layout.FieldByName('COD_PRODUTO').AsString = '0' then
//      begin
//        Layout.FieldByName('COD_PRODUTO').AsInteger := NEW_CODPROD;
//      end;

//      if Layout.FieldByName('COD_PRODUTO_ANT').AsString = '0' then
//      begin
//        Layout.FieldByName('COD_PRODUTO_ANT').AsInteger := NEW_CODPROD;
//      end;

      Layout.FieldByName('COD_PRODUTO').AsString := GerarPLU( Layout.FieldByName('COD_PRODUTO').AsString );

     // Layout.FieldByName('COD_PRODUTO_ANT').AsString := GerarPLU(QryPrincipal.FieldByName('COD_PRODUTO_ANT').AsString);

//      Layout.FieldByName('NUM_NCM').AsString := StrRetNums(Layout.FieldByName('NUM_NCM').AsString);

//      if Layout.FieldByName('NUM_NCM').AsString = '00000000' then
//      begin
//        Layout.FieldByName('NUM_NCM').AsString := '00000000';
//      end
//      else
//      begin
        Layout.FieldByName('NUM_NCM').AsString := StrRetNums(Layout.FieldByName('NUM_NCM').AsString);
//      end;

      Layout.FieldByName('NUM_CEST').AsString := StrRetNums( Layout.FieldByName('NUM_CEST').AsString );

      if QryPrincipal.FieldByName('DTA_VALIDA_OFERTA').AsString <> '' then
        Layout.FieldByName('DTA_VALIDA_OFERTA').AsDateTime := FieldByName('DTA_VALIDA_OFERTA').AsDateTime;
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

procedure TFrmSmPanda.GerarProdSimilar;
begin
  inherited;
  with QryPrincipal do
  begin
    Close;
    SQL.Clear;

     SQL.Add('   SELECT    ');
     SQL.Add('       PRODUTO.COD_ART AS COD_PRODUTO_SIMILAR,   ');
     SQL.Add('       TRIM(COALESCE(PRODUTO.DES_ART, ''A DEFINIR'')) AS DES_PRODUTO_SIMILAR,   ');
     SQL.Add('       0 AS VAL_META   ');
     SQL.Add('   FROM   ');
     SQL.Add('       MATERI_GRUPO AS PRODUTO   ');
//     SQL.Add('   LEFT JOIN MATERI_AUX AS PRO_SIMILAR ON PRO_SIMILAR.COD_ART = PRODUTO.GRUPO   ');
//     SQL.Add('   WHERE PRODUTO.GRUPO IS NOT NULL   ');
//     SQL.Add('   AND PRODUTO.GRUPO > 0   ');



    Open;    

    First;
    TotalCont := SetCountTotal(SQL.Text);
    NumLinha := 0;

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);

      Layout.SetValues(QryPrincipal, NumLinha, TotalCont);

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
