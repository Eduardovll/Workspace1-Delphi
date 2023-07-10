unit UFrmSmCasteloNext;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes,
  Vcl.Graphics, ComObj,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, UFrmModelo, Data.DBXOracle, Data.DB,
  Data.SqlExpr, Vcl.Menus, Vcl.StdCtrls, Vcl.Buttons, Vcl.ComCtrls,
  Vcl.Imaging.pngimage, Vcl.ExtCtrls, Data.DBXFirebird, Data.Win.ADODB,
  Vcl.Grids, Vcl.DBGrids, Datasnap.Provider, Datasnap.DBClient, //dxGDIPlusClasses,
  Math;

type
  TFrmSmCasteloNext = class(TFrmModeloSis)
    CbxLoja: TComboBox;
    lblLoja: TLabel;
    ADOSQLServer: TADOConnection;
    QryPrincipal2: TADOQuery;
    QryAux: TADOQuery;
    procedure BtnGerarClick(Sender: TObject);
  private

    { Private declarations }
  public
    { Public declarations }
    procedure GerarCliente;           Override; (* OK *)
    procedure GerarCondPagCli;        Override; (* OK *)

    procedure GerarFornecedor;        Override; (* OK *)
    procedure GerarCondPagForn;       Override; (* OK *)

    procedure GerarSecao;             Override; (* OK *) (* UNIFICADO *)
    procedure GerarGrupo;             Override; (* OK *) (* UNIFICADO *)
    procedure GerarSubGrupo;          Override; (* OK *) (* UNIFICADO *)

    procedure GerarProdSimilar;       Override; (* OK *)

    procedure GerarProduto;           Override; (* OK *) (* UNIFICADO *)

    procedure GerarCodigoBarras;      Override; (* OK *) (* UNIFICADO *)
    procedure GerarProdLoja;          Override; (* OK *) (* UNIFICADO *)

    procedure GerarNCM;               Override; (* OK *)
    procedure GerarNCMUF;             Override; (* OK *)
    procedure GerarCest;              Override; (* OK *)

    procedure GerarProdForn;          Override; (* OK *)

    procedure GerarNFFornec;          Override; (* OK *)
    procedure GerarNFitensFornec;     Override; (* OK *)

    procedure GerarVenda;             Override; (* OK *)

    procedure GerarFinanceiro( Tipo, Situacao :Integer ); Override; (* OK *)
    procedure GerarFinanceiroReceber(Aberto:String);      Override; (* OK *)
    procedure GerarFinanceiroPagar(Aberto:String);        Override; (* OK *)
  end;

var
  FrmSmCasteloNext: TFrmSmCasteloNext;
  NumLinha : Integer;
  Arquivo: TextFile;

implementation

{$R *.dfm}

uses xProc, UUtilidades, UProgresso;

procedure TFrmSmCasteloNext.GerarProduto;
begin
  inherited;

  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT  ');
    SQL.Add('  PRODUTOS.CODPROD AS COD_PRODUTO,  ');
    SQL.Add('  BARRA AS COD_BARRA_PRINCIPAL,  ');
    SQL.Add('  DESC_PDV AS DES_REDUZIDA,  ');
    SQL.Add('  DESCRICAO AS DES_PRODUTO,  ');
    SQL.Add('  QTD_EMB AS QTD_EMBALAGEM_COMPRA,  ');
    SQL.Add('  UNIDADE_COMP AS DES_UNIDADE_COMPRA,  ');
    SQL.Add('  0 AS QTD_EMBALAGEM_VENDA,  ');
    SQL.Add('  UNIDADE AS DES_UNIDADE_VENDA,  ');
    SQL.Add('  0 AS TIPO_IPI,  ');
    SQL.Add('  0 AS VAL_IPI,  ');
    SQL.Add('  CODCRECEITA AS COD_SECAO,  ');
    SQL.Add('  CODGRUPO AS COD_GRUPO,  ');
    SQL.Add('  CODCATEGORIA AS COD_SUB_GRUPO,  ');
    SQL.Add('  PROD_FAMILIA.CODFAMILIA AS COD_PRODUTO_SIMILAR,  ');
    SQL.Add('    ');
    SQL.Add('  CASE UNIDADE   ');
    SQL.Add('    WHEN ''KG'' THEN ''S''   ');
    SQL.Add('  ELSE ''N''  ');
    SQL.Add('  END AS IPV,  ');
    SQL.Add('    ');
    SQL.Add('  VALIDADE AS DIAS_VALIDADE,  ');
    SQL.Add('  0 AS TIPO_PRODUTO,  ');
    SQL.Add('    ');
    SQL.Add('  CASE PRODUTOS.TIPO_COFINS   ');
    SQL.Add('    WHEN ''T'' THEN ''N''  ');
    SQL.Add('  ELSE ''S''   ');
    SQL.Add('  END AS FLG_NAO_PIS_COFINS,  ');
    SQL.Add('     ');
    SQL.Add('  CASE PRODUTOS.CODSETOR  ');
    SQL.Add('    WHEN 1 THEN ''S''  ');
    SQL.Add('    ELSE ''N''  ');
    SQL.Add('  END AS FLG_ENVIA_BALANCA,  ');
    SQL.Add('    ');
    SQL.Add('  CASE   ');
    SQL.Add('    WHEN ((PRODUTOS.CST_PISENTRADA = 73) AND (PRODUTOS.CST_PISSAIDA = 06))THEN 0  ');
    SQL.Add('    WHEN ((PRODUTOS.CST_PISENTRADA = 70) AND (PRODUTOS.CST_PISSAIDA = 04))THEN 1  ');
    SQL.Add('    WHEN ((PRODUTOS.CST_PISENTRADA = 75) AND (PRODUTOS.CST_PISSAIDA = 05))THEN 2  ');
    SQL.Add('    WHEN ((PRODUTOS.CST_PISENTRADA = 74) AND (PRODUTOS.CST_PISSAIDA = 09))THEN 3  ');
    SQL.Add('    WHEN ((PRODUTOS.CST_PISENTRADA = 99) AND (PRODUTOS.CST_PISSAIDA = 49))THEN 3  ');
    SQL.Add('    WHEN ((PRODUTOS.CST_PISENTRADA = 72) AND (PRODUTOS.CST_PISSAIDA = 09))THEN 4  ');
    SQL.Add('    WHEN ((PRODUTOS.CST_PISENTRADA = 50) AND (PRODUTOS.CST_PISSAIDA = 01))THEN 0  ');
    SQL.Add('  ELSE -1 END AS TIPO_NAO_PIS_COFINS,  ');
    SQL.Add('    ');
    SQL.Add('  0 AS TIPO_EVENTO,  ');
    SQL.Add('  NULL AS COD_ASSOCIADO,  ');
    SQL.Add('  OBS AS DES_OBSERVACAO,  ');
    SQL.Add('  0 AS COD_INFO_NUTRICIONAL,  ');
    SQL.Add('  COALESCE(PRODUTOS.NAT_REC, 0) AS COD_TAB_SPED,  ');
    SQL.Add('  ''N'' AS FLG_ALCOOLICO,  ');
    SQL.Add('    ');
    SQL.Add('  CASE ');
    SQL.Add('    WHEN TIPOITEM_SPED IN (0, 1, 7, 4, 99) THEN 0   ');
    SQL.Add('    WHEN TIPOITEM_SPED IN (9) THEN 1     ');
    SQL.Add('    WHEN TIPOITEM_SPED IN (8) THEN 2   ');
    SQL.Add('    WHEN TIPOITEM_SPED IN (10, 7) THEN 4 ');
    SQL.Add('  END AS TIPO_ESPECIE, ');
    SQL.Add('    ');
    SQL.Add('  0 AS  COD_CLASSIF,  ');
    SQL.Add('  1 AS VAL_VDA_PESO_BRUTO,  ');
    SQL.Add('  1 AS VAL_PESO_EMB,  ');
    SQL.Add('  0 AS  TIPO_EXPLOSAO_COMPRA,    ');
    SQL.Add('    ');
    SQL.Add('  '''' AS DTA_INI_OPER,      ');
    SQL.Add('  '''' AS DES_PLAQUETA,      ');
    SQL.Add('  '''' AS MES_ANO_INI_DEPREC,      ');
    SQL.Add('  0 AS TIPO_BEM,      ');
    SQL.Add('  0 AS COD_FORNECEDOR,      ');
    SQL.Add('  0 AS NUM_NF,      ');
    SQL.Add('  NULL AS DTA_ENTRADA,      ');
    SQL.Add('  0 AS COD_NAT_BEM,      ');
    SQL.Add('  0 AS VAL_ORIG_BEM,      ');
    SQL.Add('  COALESCE(DESCRICAO, ''A DEFINIR'') AS DES_PRODUTO_ANT        ');
    SQL.Add('FROM    ');
    SQL.Add('  PRODUTOS  ');
    SQL.Add('    LEFT JOIN PROD_FAMILIA ON (PROD_FAMILIA.CODPROD = PRODUTOS.CODPROD)  ');
    SQL.Add(' ');
    SQL.Add('WHERE  ');
    SQL.Add('  PRODUTOS.IMPORTAR_PRODUTO = ''S'' ');
    SQL.Add(' ');
    SQL.Add('UNION ALL ');
    SQL.Add(' ');
    SQL.Add('SELECT  ');
//    SQL.Add('  PROD.CODPROD AS COD_PRODUTO,  ');
    SQL.Add(' ');
    SQL.Add('  CASE WHEN ');
    SQL.Add('	   (SELECT COUNT(*) FROM PRODUTOS WHERE CODPROD = PROD.CODPROD AND DESCRICAO <> PROD.DESCRICAO) > 0 THEN ');
    SQL.Add('	     PRODUTOS.ID + ROW_NUMBER() OVER (ORDER BY PROD.CODPROD) ');
    SQL.Add('	 ELSE ');
    SQL.Add('	    PROD.CODPROD END AS COD_PRODUTO, ');
    SQL.Add(' ');
    SQL.Add('  BARRA AS COD_BARRA_PRINCIPAL,  ');
    SQL.Add('  DESC_PDV AS DES_REDUZIDA,  ');
    SQL.Add('  DESCRICAO AS DES_PRODUTO,  ');
    SQL.Add('  QTD_EMB AS QTD_EMBALAGEM_COMPRA,  ');
    SQL.Add('  UNIDADE_COMP AS DES_UNIDADE_COMPRA,  ');
    SQL.Add('  0 AS QTD_EMBALAGEM_VENDA,  ');
    SQL.Add('  UNIDADE AS DES_UNIDADE_VENDA,  ');
    SQL.Add('  0 AS TIPO_IPI,  ');
    SQL.Add('  0 AS VAL_IPI,  ');
    SQL.Add('  999 AS COD_SECAO,  ');
    SQL.Add('  999 AS COD_GRUPO,  ');
    SQL.Add('  999 AS COD_SUB_GRUPO,  ');
    SQL.Add('  0 AS COD_PRODUTO_SIMILAR,  ');
    SQL.Add('    ');
    SQL.Add('  CASE UNIDADE   ');
    SQL.Add('    WHEN ''KG'' THEN ''S''   ');
    SQL.Add('  ELSE ''N''  ');
    SQL.Add('  END AS IPV,  ');
    SQL.Add('    ');
    SQL.Add('  VALIDADE AS DIAS_VALIDADE,  ');
    SQL.Add('  0 AS TIPO_PRODUTO,  ');
    SQL.Add('    ');
    SQL.Add('  CASE PROD.TIPO_COFINS   ');
    SQL.Add('    WHEN ''T'' THEN ''N''  ');
    SQL.Add('  ELSE ''S''   ');
    SQL.Add('  END AS FLG_NAO_PIS_COFINS,  ');
    SQL.Add('     ');
    SQL.Add('  CASE PROD.CODSETOR  ');
    SQL.Add('    WHEN 1 THEN ''S''  ');
    SQL.Add('    ELSE ''N''  ');
    SQL.Add('  END AS FLG_ENVIA_BALANCA,  ');
    SQL.Add('    ');
    SQL.Add('  CASE   ');
    SQL.Add('    WHEN ((PROD.CST_PISENTRADA = 73) AND (PROD.CST_PISSAIDA = 06))THEN 0  ');
    SQL.Add('    WHEN ((PROD.CST_PISENTRADA = 70) AND (PROD.CST_PISSAIDA = 04))THEN 1  ');
    SQL.Add('    WHEN ((PROD.CST_PISENTRADA = 75) AND (PROD.CST_PISSAIDA = 05))THEN 2  ');
    SQL.Add('    WHEN ((PROD.CST_PISENTRADA = 74) AND (PROD.CST_PISSAIDA = 09))THEN 3  ');
    SQL.Add('    WHEN ((PROD.CST_PISENTRADA = 99) AND (PROD.CST_PISSAIDA = 49))THEN 3  ');
    SQL.Add('    WHEN ((PROD.CST_PISENTRADA = 72) AND (PROD.CST_PISSAIDA = 09))THEN 4  ');
    SQL.Add('    WHEN ((PROD.CST_PISENTRADA = 50) AND (PROD.CST_PISSAIDA = 01))THEN 0  ');
    SQL.Add('  ELSE -1 END AS TIPO_NAO_PIS_COFINS,  ');
    SQL.Add('    ');
    SQL.Add('  0 AS TIPO_EVENTO,  ');
    SQL.Add('  NULL AS COD_ASSOCIADO,  ');
    SQL.Add('  OBS AS DES_OBSERVACAO,  ');
    SQL.Add('  0 AS COD_INFO_NUTRICIONAL,  ');
    SQL.Add('  COALESCE(PROD.NAT_REC, 0) AS COD_TAB_SPED,  ');
    SQL.Add('  ''N'' AS FLG_ALCOOLICO,  ');
    SQL.Add('    ');
    SQL.Add('  CASE ');
    SQL.Add('    WHEN TIPOITEM_SPED IN (0, 1, 7, 4, 99) THEN 0   ');
    SQL.Add('    WHEN TIPOITEM_SPED IN (9) THEN 1     ');
    SQL.Add('    WHEN TIPOITEM_SPED IN (8) THEN 2   ');
    SQL.Add('    WHEN TIPOITEM_SPED IN (10, 7) THEN 4 ');
    SQL.Add('  END AS TIPO_ESPECIE, ');
    SQL.Add('    ');
    SQL.Add('  0 AS  COD_CLASSIF,  ');
    SQL.Add('  1 AS VAL_VDA_PESO_BRUTO,  ');
    SQL.Add('  1 AS VAL_PESO_EMB,  ');
    SQL.Add('  0 AS  TIPO_EXPLOSAO_COMPRA,    ');
    SQL.Add('    ');
    SQL.Add('  '''' AS DTA_INI_OPER,      ');
    SQL.Add('  '''' AS DES_PLAQUETA,      ');
    SQL.Add('  '''' AS MES_ANO_INI_DEPREC,      ');
    SQL.Add('  0 AS TIPO_BEM,      ');
    SQL.Add('  0 AS COD_FORNECEDOR,      ');
    SQL.Add('  0 AS NUM_NF,      ');
    SQL.Add('  NULL AS DTA_ENTRADA,      ');
    SQL.Add('  0 AS COD_NAT_BEM,      ');
    SQL.Add('  0 AS VAL_ORIG_BEM,      ');
    SQL.Add('  COALESCE(DESCRICAO, ''A DEFINIR'') AS DES_PRODUTO_ANT        ');
    SQL.Add('FROM    ');
    SQL.Add('  UNIFICACAO_PRODUTOS PROD ');
    SQL.Add('    CROSS JOIN (SELECT MAX(CODPROD) AS ID FROM PRODUTOS) AS PRODUTOS ');
    SQL.Add('WHERE ');
    SQL.Add('  LEN(SUBSTRING(BARRA, PATINDEX(''%[^0]%'', BARRA+''.''), LEN(BARRA))) >= 8  ');
    SQL.Add('  AND  ');
    SQL.Add('  PROD.IMPORTAR_PRODUTO = ''S''  ');
    SQL.Add('  AND ');
    SQL.Add('  (SELECT COUNT(*) FROM PRODUTOS WHERE BARRA = PROD.BARRA AND PROD.IMPORTAR_PRODUTO = ''N'') = 0 ');
//    SQL.Add('  AND ');
//    SQL.Add('  (SELECT COUNT(*) FROM PRODUTOS WHERE BARRA = PROD.BARRA) = 0 ');

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

        Layout.FieldByName('DES_OBSERVACAO').AsString  := StrReplace(StrLBReplace(FieldByName('DES_OBSERVACAO').AsString), '\n', '');
        {Layout.FieldByName('DES_REDUZIDA').AsString    := StrRemPont(StrReplace(StrLBReplace(FieldByName('DES_REDUZIDA').AsString), '\n', ''));
        Layout.FieldByName('DES_PRODUTO').AsString     := StrRemPont(StrReplace(StrLBReplace(FieldByName('DES_PRODUTO').AsString), '\n', ''));
        Layout.FieldByName('DES_PRODUTO_ANT').AsString := StrRemPont(StrReplace(StrLBReplace(FieldByName('DES_PRODUTO').AsString), '\n', ''));}

        if (Layout.FieldByName('COD_BARRA_PRINCIPAL').AsString = '000000000000') or
          (Layout.FieldByName('COD_BARRA_PRINCIPAL').AsString = '0000') or
          (Layout.FieldByName('COD_BARRA_PRINCIPAL').AsString = '0')  then
          Layout.FieldByName('COD_BARRA_PRINCIPAL').AsString := '';

        if QryPrincipal2.FieldByName('COD_PRODUTO').AsString = '73354' then
          Layout.FieldByName('COD_BARRA_PRINCIPAL').AsString := Layout.FieldByName('COD_PRODUTO').AsString;

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
    Close;
  end;
end;

procedure TFrmSmCasteloNext.GerarSecao;
begin
  inherited;

  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    {SQL.Add('SELECT  ');
    SQL.Add('  CODCRECEITA AS COD_SECAO, ');
    SQL.Add('  DESCRICAO AS DES_SECAO, ');
    SQL.Add('  0 AS VAL_META  ');
    SQL.Add('FROM  ');
    SQL.Add('  CRECEITA  ');}

    SQL.Add('SELECT DISTINCT  ');
    SQL.Add('  PRODUTOS.CODCRECEITA AS COD_SECAO,  ');
    SQL.Add('  SECAO.DESCRICAO AS DES_SECAO,  ');
    SQL.Add('  0 AS VAL_META  ');
    SQL.Add('FROM  ');
    SQL.Add('  PRODUTOS  ');
    SQL.Add('    INNER JOIN CRECEITA AS SECAO ON PRODUTOS.CODCRECEITA = SECAO.CODCRECEITA  ');
    SQL.Add(' ');
    SQL.Add('WHERE  ');
    SQL.Add('  PRODUTOS.IMPORTAR_PRODUTO = ''S'' ');
    SQL.Add(' ');
    SQL.Add('UNION ALL ');
    SQL.Add(' ');
    SQL.Add('SELECT DISTINCT  ');
    SQL.Add('  999 AS COD_SECAO,  ');
    SQL.Add('  ''A DEFINIR'' AS DES_SECAO,  ');
    SQL.Add('  0 AS VAL_META  ');
    SQL.Add('FROM  ');
    SQL.Add('  PRODUTOS  ');

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

procedure TFrmSmCasteloNext.GerarGrupo;
begin
  inherited;
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    {SQL.Add('SELECT  ');
    SQL.Add('  CODCRECEITA AS COD_SECAO, ');
    SQL.Add('  CODGRUPO AS COD_GRUPO, ');
    SQL.Add('  DESCRICAO AS DES_GRUPO, ');
    SQL.Add('  0 AS VAL_META  ');
    SQL.Add('FROM  ');
    SQL.Add('  GRUPO  ');}

    SQL.Add('SELECT DISTINCT   ');
    SQL.Add('  PRODUTOS.CODCRECEITA AS COD_SECAO,   ');
    SQL.Add('  PRODUTOS.CODGRUPO AS COD_GRUPO,  ');
    SQL.Add('  GRUPO.DESCRICAO AS DES_GRUPO,  ');
    SQL.Add('  0 AS VAL_META   ');
    SQL.Add('FROM   ');
    SQL.Add('  PRODUTOS   ');
    SQL.Add('    INNER JOIN GRUPO AS GRUPO ON PRODUTOS.CODGRUPO = GRUPO.CODGRUPO   ');
    SQL.Add('  ');
    SQL.Add('WHERE   ');
    SQL.Add('  PRODUTOS.IMPORTAR_PRODUTO = ''S'' ');
    SQL.Add('  ');
    SQL.Add('UNION ALL  ');
    SQL.Add('  ');
    SQL.Add('SELECT DISTINCT   ');
    SQL.Add('  999 AS COD_SECAO,   ');
    SQL.Add('  999 AS COD_GRUPO,   ');
    SQL.Add('  ''A DEFINIR'' AS DES_GRUPO, ');
    SQL.Add('  0 AS VAL_META    ');
    SQL.Add('FROM   ');
    SQL.Add('  PRODUTOS   ');

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

procedure TFrmSmCasteloNext.GerarSubGrupo;
begin
  inherited;

  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    {SQL.Add('SELECT  ');
    SQL.Add('  CODCRECEITA AS COD_SECAO,  ');
    SQL.Add('  CODGRUPO AS COD_GRUPO,  ');
    SQL.Add('  CODCATEGORIA AS COD_SUB_GRUPO,  ');
    SQL.Add('  DESCRICAO AS DES_SUB_GRUPO,  ');
    SQL.Add('  0 AS VAL_META,  ');
    SQL.Add('  0 AS VAL_MARGEM_REF,  ');
    SQL.Add('  0 AS QTD_DIA_SEGURANCA,  ');
    SQL.Add('  ''N'' AS FLG_ALCOOLICO  ');
    SQL.Add('FROM  ');
    SQL.Add('  CATEGORIA  ');}

    SQL.Add('SELECT DISTINCT  ');
    SQL.Add('  PRODUTOS.CODCRECEITA AS COD_SECAO,  ');
    SQL.Add('  PRODUTOS.CODGRUPO AS COD_GRUPO,  ');
    SQL.Add('  PRODUTOS.CODCATEGORIA AS COD_SUB_GRUPO,  ');
    SQL.Add('  SUBGRUPO.DESCRICAO AS DES_SUB_GRUPO,  ');
    SQL.Add('  0 AS VAL_META,  ');
    SQL.Add('  0 AS VAL_MARGEM_REF,  ');
    SQL.Add('  0 AS QTD_DIA_SEGURANCA,  ');
    SQL.Add('  ''N'' AS FLG_ALCOOLICO  ');
    SQL.Add('FROM  ');
    SQL.Add('  PRODUTOS  ');
    SQL.Add('    INNER JOIN CATEGORIA AS SUBGRUPO ON PRODUTOS.CODCATEGORIA = SUBGRUPO.CODCATEGORIA  ');
    SQL.Add(' ');
    SQL.Add('WHERE  ');
    SQL.Add('  PRODUTOS.IMPORTAR_PRODUTO = ''S'' ');
    SQL.Add(' ');
    SQL.Add('UNION ALL ');
    SQL.Add(' ');
    SQL.Add('SELECT DISTINCT  ');
    SQL.Add('  999 AS COD_SECAO,  ');
    SQL.Add('  999 AS COD_GRUPO,  ');
    SQL.Add('  999 AS COD_SUB_GRUPO,  ');
    SQL.Add('  ''A DEFINIR'' AS DES_SUB_GRUPO,  ');
    SQL.Add('  0 AS VAL_META,  ');
    SQL.Add('  0 AS VAL_MARGEM_REF,  ');
    SQL.Add('  0 AS QTD_DIA_SEGURANCA,  ');
    SQL.Add('  ''N'' AS FLG_ALCOOLICO  ');
    SQL.Add('FROM  ');
    SQL.Add('  PRODUTOS  ');


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

procedure TFrmSmCasteloNext.GerarVenda;
begin
  inherited;
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT ');
    SQL.Add('  CAIXAGERAL.CODPROD AS COD_PRODUTO, ');
    SQL.Add('  '+CbxLoja.Text+' AS COD_LOJA,   ');
//    SQL.Add('  1 AS COD_LOJA,   ');
    SQL.Add('  0 AS IND_TIPO, ');
    SQL.Add('  CAIXAGERAL.CODCAIXA AS NUM_PDV, ');
    SQL.Add('  CAIXAGERAL.QTD AS QTD_TOTAL_PRODUTO, ');
    SQL.Add('  CAIXAGERAL.TOTITEM AS VAL_TOTAL_PRODUTO, ');
    SQL.Add('  CAIXAGERAL.VALORUNIT AS VAL_PRECO_VENDA, ');
    SQL.Add('  CAIXAGERAL.VALORCUSTO AS VAL_CUSTO_REP, ');
    SQL.Add('  CAIXAGERAL.DATA AS DTA_SAIDA, ');
    SQL.Add('  RIGHT(''0'' + CONVERT(VARCHAR(2), MONTH(CAIXAGERAL.DATA)), 2) + CONVERT(CHAR(4), YEAR(CAIXAGERAL.DATA)) AS DTA_MENSAL, ');
    SQL.Add('  1 AS NUM_IDENT, ');
    SQL.Add('  NULL AS COD_EAN, ');
    SQL.Add('  REPLACE(CAIXAGERAL.HORA, '':'', '''') AS DES_HORA, ');
    SQL.Add('  CAIXAGERAL.CLIENTE AS COD_CLIENTE, ');
    SQL.Add('  1 AS COD_ENTIDADE, ');
    SQL.Add('  0 AS VAL_BASE_ICMS, ');
    SQL.Add('  '''' AS DES_SITUACAO_TRIB, ');
    SQL.Add('  0 AS VAL_ICMS, ');
    SQL.Add('  CAIXAGERAL.COO AS NUM_CUPOM_FISCAL, ');
    SQL.Add('  CAIXAGERAL.VALORUNIT AS VAL_VENDA_PDV, ');
    SQL.Add('   ');
    SQL.Add('  CASE ');
    SQL.Add('    WHEN ((CAIXAGERAL.CODTRIB IN (200, 20, 41))    AND (CAIXAGERAL.CODALIQ IN (''TC'', ''TB'', ''TA'', ''TG'', ''TC'', ''F'', ''I'', ''N''))    ');
    SQL.Add('	   AND (PER_REDUC IN (67.06, 0.00))) THEN 51 ');

    SQL.Add('    WHEN ((CAIXAGERAL.CODTRIB IN (200, 20))        AND (CAIXAGERAL.CODALIQ IN (''TC'', ''TA'', ''F'', ''I'', ''N''))  ');
    SQL.Add('	   AND (PER_REDUC IN (100.00))) THEN 66 ');

    SQL.Add('    WHEN ((CAIXAGERAL.CODTRIB IN (60, 260))        AND (CAIXAGERAL.CODALIQ IN (''F'', ''I'', ''N'')) AND (PER_REDUC IN (100.00, 0.00))) THEN 60 ');
    SQL.Add('    WHEN ((CAIXAGERAL.CODTRIB IN (0, 00, 000))     AND (CAIXAGERAL.CODALIQ IN (''TC'', ''TB'')) AND (PER_REDUC IN (0.00)))              THEN 3  ');
    SQL.Add('    WHEN ((CAIXAGERAL.CODTRIB IN (0, 00, 000, 20)) AND (CAIXAGERAL.CODALIQ IN (''TC'')) AND (PER_REDUC IN (58.82)))                     THEN 52 ');
    SQL.Add('    WHEN ((CAIXAGERAL.CODTRIB IN (40, 41))         AND (CAIXAGERAL.CODALIQ IN (''F'', ''I'', ''N'')) AND (PER_REDUC IN (0.00)))         THEN 1  ');
    SQL.Add('    WHEN ((CAIXAGERAL.CODTRIB IN (20))             AND (CAIXAGERAL.CODALIQ IN (''TB'')) AND (PER_REDUC IN (41.66)))                     THEN 6  ');
    SQL.Add('    WHEN ((CAIXAGERAL.CODTRIB IN (0, 00, 000))     AND (CAIXAGERAL.CODALIQ IN (''TA'')) AND (PER_REDUC IN (0.00)))                      THEN 2  ');
    SQL.Add('    WHEN ((CAIXAGERAL.CODTRIB IN (60))             AND (CAIXAGERAL.CODALIQ IN (''F'', ''I'', ''N'')) AND (PER_REDUC IN (58.82)))        THEN 56 ');
    SQL.Add('    WHEN ((CAIXAGERAL.CODTRIB IN (0, 00, 000))     AND (CAIXAGERAL.CODALIQ IN (''TD'')) AND (PER_REDUC IN (0.00)))                      THEN 5  ');
    SQL.Add('    ');
    SQL.Add('  ELSE 1 END AS COD_TRIBUTACAO, ');
    SQL.Add('   ');
    SQL.Add('  ''N'' AS FLG_CUPOM_CANCELADO, ');
    SQL.Add('  PRODUTOS.CODNCM AS NUM_NCM, ');
    SQL.Add('  COALESCE(PRODUTOS.NAT_REC, 0) AS COD_TAB_SPED, ');
    SQL.Add('   ');
    SQL.Add('  CASE PRODUTOS.TIPO_COFINS  ');
    SQL.Add('    WHEN ''T'' THEN ''N'' ');
    SQL.Add('  ELSE ''S'' ');
    SQL.Add('  END AS FLG_NAO_PIS_COFINS, ');
    SQL.Add('   ');
    SQL.Add('  CASE  ');
    SQL.Add('    WHEN ((PRODUTOS.CST_PISENTRADA = ''73'') AND (PRODUTOS.CST_PISSAIDA = ''06''))THEN 0 ');
    SQL.Add('    WHEN ((PRODUTOS.CST_PISENTRADA = ''70'') AND (PRODUTOS.CST_PISSAIDA = ''04''))THEN 1 ');
    SQL.Add('    WHEN ((PRODUTOS.CST_PISENTRADA = ''75'') AND (PRODUTOS.CST_PISSAIDA = ''05''))THEN 2 ');
    SQL.Add('    WHEN ((PRODUTOS.CST_PISENTRADA = ''74'') AND (PRODUTOS.CST_PISSAIDA = ''09''))THEN 3 ');
    SQL.Add('    WHEN ((PRODUTOS.CST_PISENTRADA = ''99'') AND (PRODUTOS.CST_PISSAIDA = ''49''))THEN 3 ');
    SQL.Add('    WHEN ((PRODUTOS.CST_PISENTRADA = ''72'') AND (PRODUTOS.CST_PISSAIDA = ''09''))THEN 4 ');
    SQL.Add('    WHEN ((PRODUTOS.CST_PISENTRADA = ''50'') AND (PRODUTOS.CST_PISSAIDA = ''01''))THEN 0 ');
    SQL.Add('  ELSE -1 END AS TIPO_NAO_PIS_COFINS, ');
    SQL.Add('   ');
    SQL.Add('  ''N'' AS FLG_ONLINE, ');
    SQL.Add('  ''N'' AS FLG_OFERTA, ');
    SQL.Add('  NULL AS COD_ASSOCIADO ');
    SQL.Add('FROM ');
    SQL.Add('  CAIXAGERAL ');
    SQL.Add('    LEFT JOIN PRODUTOS ON CAIXAGERAL.CODPROD = PRODUTOS.CODPROD ');
    SQL.Add('    LEFT JOIN ALIQUOTA_ICMS ON CAIXAGERAL.CODALIQ = ALIQUOTA_ICMS.CODALIQ ');
    SQL.Add('WHERE ');
    SQL.Add('  CAIXAGERAL.CODPROD IS NOT NULL ');
    SQL.Add('  AND ');
    SQL.Add('  CAIXAGERAL.CANCELADO = ''N'' ');
    SQL.Add('  AND ');
    SQL.Add('  CAIXAGERAL.ATUALIZADO = ''S'' ');
    SQL.Add('  AND ');
    SQL.Add('  CAST(CAIXAGERAL.DATA AS DATE) >= '''+FormatDateTime('yyyy-mm-dd',DtpInicial.Date)+''' ');
    SQL.Add('  AND');
    SQL.Add('  CAST(CAIXAGERAL.DATA AS DATE) <= '''+FormatDAteTime('yyyy-mm-dd',DtpFinal.Date)+''' ');

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

        Layout.FieldByName('NUM_NCM').AsString       := StrRetNums(Layout.FieldByName('NUM_NCM').AsString);

        Layout.FieldByName('COD_PRODUTO').AsString   := GerarPLU(QryPrincipal2.FieldByName('COD_PRODUTO').AsString);
        Layout.FieldByName('COD_ASSOCIADO').AsString := GerarPLU( Layout.FieldByName('COD_ASSOCIADO').AsString );
        if QryPrincipal2.FieldByName('DTA_SAIDA').AsString <> '' then
          Layout.FieldByName('DTA_SAIDA').AsString:= FormatDateTime('dd/mm/yyyy', QryPrincipal2.FieldByName('DTA_SAIDA').AsDateTime);

        Layout.WriteLine;
      except
        On E: Exception do
        FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
      end;
      Next;
    end;
  end;
end;

procedure TFrmSmCasteloNext.BtnGerarClick(Sender: TObject);
begin
  ADOSQLServer.Connected := False;
  ADOSQLServer.ConnectionString := 'Provider=MSDASQL.1;Password="'+edtSenhaOracle.Text+'";ID='+edtInst.Text+';Data Source='+edtSchema.Text+';Persist Security Info=False';
  ADOSQLServer.Connected := True;
  inherited;
end;

procedure TFrmSmCasteloNext.GerarCest;
var
  Count: integer;
begin
  inherited;
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT DISTINCT  ');
    SQL.Add('  0 AS COD_CEST,  ');
    SQL.Add('  ''A DEFINIR'' AS DES_CEST,  ');
    SQL.Add('  CASE WHEN COALESCE(CODCEST, '''') = '''' OR CODCEST IS NULL THEN ''99.999.99'' ELSE CODCEST END AS NUM_CEST  ');
    SQL.Add('FROM  ');
    SQL.Add('  PRODUTOS  ');
    SQL.Add(' ');
    SQL.Add('WHERE   ');
    SQL.Add('  PRODUTOS.IMPORTAR_PRODUTO = ''S'' ');
    SQL.Add('   ');
    SQL.Add('UNION ALL   ');
    SQL.Add(' ');
    SQL.Add('SELECT DISTINCT  ');
    SQL.Add('  0 AS COD_CEST,  ');
    SQL.Add('  ''A DEFINIR'' AS DES_CEST,  ');
    SQL.Add('  CASE WHEN COALESCE(CODCEST, '''') = '''' OR CODCEST IS NULL THEN ''99.999.99'' ELSE CODCEST END AS NUM_CEST  ');
    SQL.Add('FROM  ');
    SQL.Add('  UNIFICACAO_PRODUTOS PROD  ');
    SQL.Add(' ');
    SQL.Add('WHERE   ');
    SQL.Add('  LEN(SUBSTRING(BARRA, PATINDEX(''%[^0]%'', BARRA+''.''), LEN(BARRA))) >= 8   ');
    SQL.Add('  AND   ');
    SQL.Add('  PROD.IMPORTAR_PRODUTO = ''S''   ');
    SQL.Add('  AND  ');
    SQL.Add('  (SELECT COUNT(*) FROM PRODUTOS WHERE BARRA = PROD.BARRA) = 0  ');
//    SQL.Add('  AND ');
//    SQL.Add('  (SELECT COUNT(*) FROM PRODUTOS WHERE CODPROD = PROD.CODPROD) = 0 ');


    Open;
    First;

    Count    := 0;
    NumLinha := 0;

    while not EoF do
    begin
      try
        if Cancelar then
          Break;

        Inc(NumLinha);
        Inc(Count);

        Layout.SetValues(QryPrincipal2, NumLinha, RecordCount);

        Layout.FieldByName('COD_CEST').AsInteger := Count;

        Layout.FieldByName('NUM_CEST').AsString  := StrRetNums(Layout.FieldByName('NUM_CEST').AsString);

        Layout.WriteLine;
      except
        On E: Exception do
        FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
      end;
      Next;
    end;
  end;
end;

procedure TFrmSmCasteloNext.GerarCliente;
var
  Obs    : TStringList;
  QryTel : TADOQuery;
begin
  inherited;
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT ');
    SQL.Add('  CODCLIE AS COD_CLIENTE,  ');
    SQL.Add('  RAZAO AS DES_CLIENTE,  ');
    SQL.Add('  CNPJ_CPF AS NUM_CGC,  ');
    SQL.Add('  CASE WHEN COALESCE(IE, '''') = '''' OR IE IS NULL THEN ''ISENTO'' ELSE IE END AS NUM_INSC_EST,  ');
    SQL.Add('  CASE WHEN COALESCE(ENDERECO, '''') = '''' OR ENDERECO IS NULL THEN ''A DEFINIR'' ELSE ENDERECO END AS DES_ENDERECO, ');
    SQL.Add('  CASE WHEN COALESCE(BAIRRO, '''')   = '''' OR BAIRRO IS NULL   THEN ''A DEFINIR'' ELSE BAIRRO   END AS DES_BAIRRO, ');
    SQL.Add('  CASE WHEN COALESCE(CIDADE, '''')   = '''' OR CIDADE IS NULL   THEN ''VITORIA''   ELSE CIDADE   END AS DES_CIDADE, ');
    SQL.Add('  CASE WHEN COALESCE(ESTADO, '''')   = '''' OR ESTADO IS NULL   THEN ''ES''        ELSE ESTADO   END AS DES_SIGLA, ');
    SQL.Add('  CASE WHEN COALESCE(CEP, '''')      = '''' OR CEP IS NULL      THEN ''29070330''  ELSE CEP      END AS NUM_CEP, ');
    SQL.Add('  SUBSTRING(TELEFONE, PATINDEX(''%[^0]%'', TELEFONE+''.''), LEN(TELEFONE)) AS NUM_FONE, ');
    SQL.Add('  SUBSTRING(FAX, PATINDEX(''%[^0]%'', FAX+''.''), LEN(FAX)) AS NUM_FAX, ');
    SQL.Add('  CONTATO AS DES_CONTATO,  ');
    SQL.Add('                     ');
    SQL.Add('  CASE SEXO ');
    SQL.Add('    WHEN ''M'' THEN 0 ');
    SQL.Add('    WHEN ''F'' THEN 1 ');
    SQL.Add('  END AS FLG_SEXO, ');
    SQL.Add('  ');
    SQL.Add(' 0 AS VAL_LIMITE_CRETID, ');
    SQL.Add('  LIMITECRED AS VAL_LIMITE_CONV, ');
    SQL.Add('   ');
    SQL.Add('  0 AS VAL_DEBITO, ');
    SQL.Add('  RENDA AS VAL_RENDA, ');
    SQL.Add('  CASE WHEN LIMITECRED > 0.00 THEN 99999 ELSE 0 END AS COD_CONVENIO, ');
    SQL.Add('                           ');
    SQL.Add('  CASE  ');
    SQL.Add('    WHEN CODTIPOCLIE = 2 AND BLOQCARTAO = ''S'' THEN 1  ');
    SQL.Add('  ELSE 0  ');
    SQL.Add('  END AS COD_STATUS_PDV, ');
    SQL.Add('   ');
    SQL.Add('  CASE PESSOA ');
    SQL.Add('    WHEN ''J'' THEN ''S'' ');
    SQL.Add('  ELSE ''N'' ');
    SQL.Add('  END  AS FLG_EMPRESA, ');
    SQL.Add('                       ');
    SQL.Add('  ''N'' AS FLG_CONVENIO, ');
    SQL.Add('  ''N'' AS MICRO_EMPRESA, ');
    SQL.Add('  DTCAD AS DTA_CADASTRO, ');
    SQL.Add('  NUMERO AS NUM_ENDERECO, ');
    SQL.Add('  RG AS NUM_RG, ');
    SQL.Add(' ');
    SQL.Add('  CASE ESTADOCIVIL ');
    SQL.Add('    WHEN ''S'' THEN 0 ');
    SQL.Add('    WHEN ''C'' THEN 1 ');
    SQL.Add('    WHEN ''V'' THEN 2 ');
    SQL.Add('    WHEN ''O'' THEN 4   ');
    SQL.Add('  END AS FLG_EST_CIVIL, ');
    SQL.Add(' ');
    SQL.Add('  CELULAR AS NUM_CELULAR, ');
    SQL.Add('  NULL AS DTA_ALTERACAO, ');
    SQL.Add(' ');
    SQL.Add('  COALESCE(OBS1, '''') + '' '' + COALESCE(OBS2, '''')  + '' '' +  ');
    SQL.Add('  CASE WHEN COALESCE(EMPRESA, '''') = '''' THEN '''' ELSE ''Empresa: '' + EMPRESA END + '' '' +   ');
    SQL.Add('  CASE WHEN COALESCE(FONE_EMP, '''') = '''' THEN '''' ELSE ''Fone: '' + FONE_EMP END + '' '' +  ');
    SQL.Add('  ''Tempo Serviço: '' + CAST(TEMPOSERVICO AS VARCHAR) + '' '' +  ');
    SQL.Add('  CASE WHEN COALESCE(CARGO, '''') = '''' THEN '''' ELSE ''Cargo: '' + CARGO END + '' '' +  ');
    SQL.Add('  ''Renda: '' + CAST(RENDA AS VARCHAR) + '' '' AS DES_OBSERVACAO, ');
    SQL.Add(' ');
    SQL.Add('  COMPLEMENTO AS DES_COMPLEMENTO, ');
    SQL.Add(' ');
    SQL.Add('  CASE  ');
    SQL.Add('    WHEN EMAIL IS NULL     AND EMAILNFE IS NULL      THEN '''' ');
    SQL.Add('	   WHEN EMAIL IS NOT NULL AND EMAILNFE IS NULL      THEN EMAIL ');
    SQL.Add('	   WHEN EMAIL IS NULL     AND EMAILNFE IS NOT NULL  THEN EMAILNFE ');
    SQL.Add('	   WHEN EMAIL IS NOT NULL AND EMAILNFE IS NOT NULL  THEN EMAIL + '';'' + EMAILNFE ');
    SQL.Add('  END AS DES_EMAIL, ');
    SQL.Add(' ');
    SQL.Add('  CASE WHEN COALESCE(FANTASIA, '''') = '''' OR FANTASIA IS NULL THEN RAZAO ELSE FANTASIA END AS DES_FANTASIA, ');
    SQL.Add('  DTANIVER AS DTA_NASCIMENTO, ');
    SQL.Add('  NOMEPAI AS DES_PAI, ');
    SQL.Add('  NOMEMAE AS DES_MAE, ');
    SQL.Add('  NOMECONJUGE AS DES_CONJUGE, ');
    SQL.Add('  CPF_CONJUGE AS NUM_CPF_CONJUGE, ');
    SQL.Add('  0 AS VAL_DEB_CONV, ');
    SQL.Add(' ');
    SQL.Add('  CASE ATIVO ');
    SQL.Add('    WHEN ''S'' THEN ''N'' ');
    SQL.Add('  ELSE ''S'' ');
    SQL.Add('  END AS INATIVO, ');
    SQL.Add(' ');
    SQL.Add('  0 AS DES_MATRICULA, ');
    SQL.Add('  ''N'' AS NUM_CGC_ASSOCIADO, ');
    SQL.Add('   PRODRURAL AS FLG_PROD_RURAL, ');
    SQL.Add(' ');
    SQL.Add('  CASE  ');
    SQL.Add('    WHEN BLOQCARTAO = ''S'' THEN 1  ');
    SQL.Add('  ELSE 0  ');
    SQL.Add('  END AS COD_STATUS_PDV_CONV, ');
    SQL.Add('   ');
    SQL.Add('  CASE WHEN LIMITECRED > 0.00 THEN ''S'' ELSE ''N'' END AS FLG_ENVIA_CODIGO, ');
    SQL.Add('   NULL AS DTA_NASC_CONJUGE, ');
    SQL.Add('   ');
    SQL.Add('  CASE CODTIPOCLIE  ');
    SQL.Add('    WHEN 1 THEN 30 ');
    SQL.Add('	   WHEN 2 THEN 28 ');
    SQL.Add('	   WHEN 3 THEN 31 ');
    SQL.Add('	   WHEN 4 THEN 32 ');
    SQL.Add('	   WHEN 5 THEN 33 ');
    SQL.Add('	   WHEN 6 THEN 34 ');
    SQL.Add('  ELSE 0 END AS COD_CLASSIF  ');
    SQL.Add('FROM ');
    SQL.Add('  CLIENTES  ');

    Open;
    First;

    QryTel := TADOQuery.Create(FrmProgresso);
    Obs := TStringList.Create;

    with QryTel do
    begin
      Connection := ADOSQLServer;
      SQL.Clear;
      SQL.Add('SELECT CLIENTES.FONE1 AS TELEFONE1,');
      SQL.Add('CLIENTES.FONE2 AS TELEFONE2,       ');
      SQL.Add('CLIENTES.CODCLIE AS COD_CLIENTE    ');
      SQL.Add('FROM CLIENTES                    ');
      SQL.Add('WHERE CODCLIE = :COD_CLIENTE');
    end;

    NumLinha := 0;

    while not EoF do
    begin
      try
        if Cancelar then
          Break;
        Inc(NumLinha);

        Layout.SetValues(QryPrincipal2, NumLinha, RecordCount);

        Layout.FieldByName('NUM_CGC').AsString := StrRetNums(Layout.FieldByName('NUM_CGC').AsString);

        if (StrRetNums(Layout.FieldByName('NUM_ENDERECO').AsString) = '') then
           Layout.FieldByName('NUM_ENDERECO').AsString := 'S/N'
        else if strtoint(StrRetNums(Layout.FieldByName('NUM_ENDERECO').AsString)) = 0 then
           Layout.FieldByName('NUM_ENDERECO').AsString := 'S/N'
        else
           Layout.FieldByName('NUM_ENDERECO').AsString := StrRetNums(Layout.FieldByName('NUM_ENDERECO').AsString);

        if StrRetNums(Layout.FieldByName('NUM_RG').AsString) = '' then
          Layout.FieldByName('NUM_RG').AsString := ''
        else
          Layout.FieldByName('NUM_RG').AsString        := StrRetNums(Layout.FieldByName('NUM_RG').AsString);

        Layout.FieldByName('NUM_CEP').AsString         := StrRetNums(Layout.FieldByName('NUM_CEP').AsString);

        if QryPrincipal2.FieldByName('DTA_CADASTRO').AsString <> '' then
          Layout.FieldByName('DTA_CADASTRO').AsString := FormatDateTime('dd/mm/yyyy',QryPrincipal2.FieldByName('DTA_CADASTRO').AsDateTime);

        if QryPrincipal2.FieldByName('DTA_NASCIMENTO').AsString <> '' then
          Layout.FieldByName('DTA_NASCIMENTO').AsString := FormatDateTime('dd/mm/yyyy',QryPrincipal2.FieldByName('DTA_NASCIMENTO').AsDateTime);

        Layout.FieldByName('NUM_FONE').AsString        := StrRetNums( FieldByName('NUM_FONE').AsString );

        Layout.FieldByName('NUM_FAX').AsString        := StrRetNums( FieldByName('NUM_FAX').AsString );

        Layout.FieldByName('NUM_CELULAR').AsString        := StrRetNums( FieldByName('NUM_CELULAR').AsString );

        if Length(Layout.FieldByName('NUM_CGC').AsString) > 11 then
        begin
          if not ValidaCGC(Layout.FieldByName('NUM_CGC').AsString) then
            Layout.FieldByName('NUM_CGC').AsString := '';
        end
        else
          if not ValidaCpf(Layout.FieldByName('NUM_CGC').AsString) then
            Layout.FieldByName('NUM_CGC').AsString := '';

        Layout.FieldByName('DES_EMAIL').AsString      := StrReplace(StrLBReplace(FieldByName('DES_EMAIL').AsString), '\n', '');
        Layout.FieldByName('DES_ENDERECO').AsString   := StrReplace(StrLBReplace(FieldByName('DES_ENDERECO').AsString), '\n', '');
        Layout.FieldByName('DES_OBSERVACAO').AsString := StrReplace(StrLBReplace(FieldByName('DES_OBSERVACAO').AsString), '\n', '');

        if Layout.FieldByName('DES_EMAIL').AsString= ';' then
          Layout.FieldByName('DES_EMAIL').AsString := '';

        if Layout.FieldByName('FLG_EMPRESA').AsString = 'N' then
          Layout.FieldByName('NUM_INSC_EST').AsString := '';

        QryTel.Close;
        QryTel.Parameters.ParamByName('COD_CLIENTE').Value := FieldByName('COD_CLIENTE').AsInteger;
        QryTel.Open;

        if (QryTel.FieldByName('TELEFONE1').AsString <> '') then
          Obs.Add(('TEL1: ')+QryTel.FieldByName('TELEFONE1').AsString);

        if (QryTel.FieldByName('TELEFONE2').AsString <> '') then
          Obs.Add(('TEL2: ')+QryTel.FieldByName('TELEFONE2').AsString);

        Obs.Add(Layout.FieldByName('DES_OBSERVACAO').AsString);
        Layout.FieldByName('DES_OBSERVACAO').AsString := StrLBReplace(Obs.Text);
        Obs.Clear;

        Layout.WriteLine;
      except
        On E: Exception do
        FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
      end;
      Next;
    end;
  end;
end;

procedure TFrmSmCasteloNext.GerarCondPagCli;
begin
  inherited;

  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT                 ');
    SQL.Add('  CODCLIE AS COD_CLIENTE, ');
    SQL.Add('  30 AS NUM_CONDICAO, ');
    SQL.Add('  2 AS  COD_CONDICAO,      ');
    SQL.Add('  1 AS COD_ENTIDADE ');
    SQL.Add('FROM');
    SQL.Add('  CLIENTES');

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

procedure TFrmSmCasteloNext.GerarCodigoBarras;
var
  count, count1 : Integer;
  CodPrincipal : string;
begin
  inherited;
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT    ');
    SQL.Add('  CODPROD AS COD_PRODUTO,    ');
    SQL.Add('  BARRA AS COD_EAN    ');
    SQL.Add('FROM    ');
    SQL.Add('  PRODUTOS    ');
    SQL.Add('  ');
    SQL.Add('WHERE   ');
    SQL.Add('  PRODUTOS.IMPORTAR_PRODUTO = ''S''  ');
    SQL.Add('    ');
    SQL.Add('UNION ALL   ');
    SQL.Add('   ');
    SQL.Add('SELECT   ');
    SQL.Add('  CODPROD AS COD_PRODUTO,   ');
    SQL.Add('  BARRA AS COD_EAN   ');
    SQL.Add('FROM   ');
    SQL.Add('  ALTERNATIVO   ');
    SQL.Add('  ');
    SQL.Add('UNION ALL  ');
    SQL.Add('  ');
    SQL.Add('SELECT    ');
//    SQL.Add('  CODPROD AS COD_PRODUTO,    ');
    SQL.Add(' ');
    SQL.Add('  CASE WHEN ');
    SQL.Add('	   (SELECT COUNT(*) FROM PRODUTOS WHERE CODPROD = PROD.CODPROD AND DESCRICAO <> PROD.DESCRICAO) > 0 THEN ');
    SQL.Add('	     PRODUTOS.ID + ROW_NUMBER() OVER (ORDER BY PROD.CODPROD) ');
    SQL.Add('	 ELSE ');
    SQL.Add('	    PROD.CODPROD END AS COD_PRODUTO, ');
    SQL.Add(' ');
    SQL.Add('  BARRA AS COD_EAN    ');
    SQL.Add('FROM    ');
    SQL.Add('  UNIFICACAO_PRODUTOS PROD    ');
    SQL.Add('    CROSS JOIN (SELECT MAX(CODPROD) AS ID FROM PRODUTOS) AS PRODUTOS ');
    SQL.Add('  ');
    SQL.Add('WHERE  ');
    SQL.Add('  LEN(SUBSTRING(BARRA, PATINDEX(''%[^0]%'', BARRA+''.''), LEN(BARRA))) >= 8   ');
    SQL.Add('  AND   ');
    SQL.Add('  PROD.IMPORTAR_PRODUTO = ''S''   ');
    SQL.Add('  AND  ');
    SQL.Add('  (SELECT COUNT(*) FROM PRODUTOS WHERE BARRA = PROD.BARRA AND PROD.IMPORTAR_PRODUTO = ''N'') = 0  ');
//    SQL.Add('  AND ');
//    SQL.Add('  (SELECT COUNT(*) FROM PRODUTOS WHERE CODPROD = PROD.CODPROD) = 0 ');
    SQL.Add('    ');
    SQL.Add('UNION ALL   ');
    SQL.Add('   ');
    SQL.Add('SELECT   ');
    SQL.Add('  CODPROD AS COD_PRODUTO,   ');
    SQL.Add('  BARRA AS COD_EAN   ');
    SQL.Add('FROM   ');
    SQL.Add('  UNIFICACAO_PRODUTOS_ALTERNATIVO PRODALTER  ');
    SQL.Add('WHERE  ');
    SQL.Add('  LEN(SUBSTRING(BARRA, PATINDEX(''%[^0]%'', BARRA+''.''), LEN(BARRA))) >= 8   ');
    SQL.Add('  AND  ');
    SQL.Add('  (SELECT COUNT(*) FROM ALTERNATIVO WHERE BARRA = PRODALTER.BARRA) = 0  ');

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

        //CodPrincipal := FieldByName('COD_EAN').AsString;

         {if ((StrPad(FieldByName('COD_EAN').AsString,13,'0','L')) =
           (Strpad(FieldByName('COD_PRODUTO').AsString,13,'0','L')) ) then
         begin
           Layout.FieldByName('COD_EAN').AsString := GerarPLU(CodPrincipal);
         end
         else
         if (
              (
                (Length(FieldByName('COD_EAN').AsString) >= 8) and
                (Length(FieldByName('COD_EAN').AsString) <= 13) and
                (not CodBarrasValido(FieldByName('COD_EAN').AsString))
              )
              or
              (Length(FieldByName('COD_EAN').AsString) < 8)
            ) then
         begin
           Layout.FieldByName('COD_EAN').AsString := GerarPLU(CodPrincipal);
         end;}


        if Length(StrLBReplace(Trim(StrRetNums( FieldByName('COD_EAN').AsString) ))) < 8 then
         Layout.FieldByName('COD_EAN').AsString := GerarPLU(FieldByName('COD_EAN').AsString);

        if not CodBarrasValido(Layout.FieldByName('COD_EAN').AsString) then
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

procedure TFrmSmCasteloNext.GerarFinanceiro(Tipo, Situacao: Integer);
begin
  inherited;
  if Tipo = 1 then
    GerarFinanceiroPagar(IntToStr(Situacao));

  if Tipo = 2 then
    GerarFinanceiroReceber(IntToStr(Situacao));
end;

procedure TFrmSmCasteloNext.GerarFinanceiroPagar(Aberto: String);
begin
  inherited;
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT   ');
    SQL.Add('  1 AS  TIPO_PARCEIRO, ');
//    SQL.Add('  PAGAR.CODFORNEC AS COD_PARCEIRO, ');
    SQL.Add('  0 AS COD_PARCEIRO, ');
    SQL.Add('  0 AS TIPO_CONTA, ');
    SQL.Add('  CASE WHEN CODTIPODOCUMENTO = 1 THEN 8 END AS COD_ENTIDADE, ');
    SQL.Add('  CODPAGAR AS NUM_DOCTO, ');
    SQL.Add('  999 AS COD_BANCO, ');
    SQL.Add('  '''' AS DES_BANCO, ');
    SQL.Add('  DTEMISSAO AS DTA_EMISSAO, ');
    SQL.Add('  DTVENCTO AS DTA_VENCIMENTO, ');
    SQL.Add('   ');
    SQL.Add('  CASE  ');
    SQL.Add('    WHEN PAGAR.SITUACAO = ''BT'' THEN VALOR ');
    SQL.Add('  ELSE (COALESCE(VALOR, 0) - COALESCE(VALORPAGO, 0)) ');
    SQL.Add('  END AS VAL_PARCELA, ');
    SQL.Add('    ');
    SQL.Add('  VALORJUROS AS VAL_JUROS, ');
    SQL.Add('  VALORDESC AS VAL_DESCONTO, ');
    SQL.Add('   ');
    SQL.Add('  CASE SITUACAO ');
    SQL.Add('    WHEN ''BT'' THEN ''S'' ');
    SQL.Add('    WHEN ''AB'' THEN ''N'' ');
    SQL.Add('   END AS FLG_QUITADO, ');
    SQL.Add('   ');
    SQL.Add('  CASE  ');
    SQL.Add('    WHEN PAGAR.SITUACAO = ''BT'' THEN ');
    SQL.Add('      CASE WHEN PAGAR.DTPAGTO <= PAGAR.DTEMISSAO THEN PAGAR.DTEMISSAO ELSE PAGAR.DTPAGTO END ');
    SQL.Add('    WHEN PAGAR.SITUACAO = ''AB'' THEN NULL ');
    SQL.Add('  END AS DTA_QUITADA, ');
    SQL.Add('   ');
    SQL.Add('  998 AS COD_CATEGORIA, ');
    SQL.Add('  998 AS COD_SUBCATEGORIA, ');
    SQL.Add('  DESD AS NUM_PARCELA, ');
    SQL.Add('  PARCELAS.QTD_PARCELA AS QTD_PARCELA, ');
    SQL.Add('  '+CbxLoja.Text+' AS COD_LOJA,   ');
//    SQL.Add('  1 AS COD_LOJA,   ');
    SQL.Add('  FORNECEDORES.CNPJ_CPF AS NUM_CGC, ');
    SQL.Add('  0 AS NUM_BORDERO, ');
    SQL.Add('  NOTA AS NUM_NF, ');
    SQL.Add('  0 AS NUM_SERIE_NF, ');
    SQL.Add('  PARCELAS.VAL_TOTAL_NF AS VAL_TOTAL_NF, ');
    SQL.Add('  COALESCE(PAGAR.OBS, '''')+'' ''+ COALESCE(PAGAR.OBS2, '''')    AS  DES_OBSERVACAO, ');
    SQL.Add('  1 AS NUM_PDV, ');
    SQL.Add('  NULL AS NUM_CUPOM_FISCAL, ');
    SQL.Add('  0 AS COD_MOTIVO, ');
    SQL.Add('  0 AS COD_CONVENIO, ');
    SQL.Add('  0 AS COD_BIN, ');
    SQL.Add('  '''' AS DES_BANDEIRA, ');
    SQL.Add('  '''' AS DES_REDE_TEF, ');
    SQL.Add('  0 AS VAL_RETENCAO, ');
    SQL.Add('  2 AS COD_CONDICAO, ');
    SQL.Add('  DTPAGTO AS DTA_PAGTO, ');
    SQL.Add('  DTENTRADA AS DTA_ENTRADA, ');
    SQL.Add('  0 AS NUM_NOSSO_NUMERO, ');
    SQL.Add('  '''' AS COD_BARRA, ');
    SQL.Add('  ''N'' AS FLG_BOLETO_EMIT, ');
    SQL.Add('  NULL AS NUM_CGC_CPF_TITULAR, ');
    SQL.Add('  NULL AS DES_TITULAR, ');
    SQL.Add('  30 AS NUM_CONDICAO, ');
    SQL.Add('  0 AS VAL_CREDITO, ');
    SQL.Add('  999 AS COD_BANCO_PGTO,      ');
    SQL.Add('  ''PAGTO'' AS DES_CC,   ');
    SQL.Add('  0 AS COD_BANDEIRA,      ');
    SQL.Add('  '''' AS DTA_PRORROGACAO,   ');
    SQL.Add('  1 AS NUM_SEQ_FIN,      ');
    SQL.Add('  0 AS COD_COBRANCA,      ');
    SQL.Add('  '''' AS DTA_COBRANCA,   ');
    SQL.Add('  ''N'' AS FLG_ACEITE,   ');
    SQL.Add('  0 AS TIPO_ACEITE     ');
    SQL.Add('FROM   ');
    SQL.Add('  PAGAR ');
    SQL.Add('    INNER JOIN FORNECEDORES ON (FORNECEDORES.CODFORNEC = PAGAR.CODFORNEC) ');
    SQL.Add('    LEFT JOIN ( ');
    SQL.Add('      SELECT ');
    SQL.Add('       NUMTIT AS  NUM_DOCTO, ');
    SQL.Add('       SUM(VALOR) AS VAL_TOTAL_NF, ');
    SQL.Add('       MAX(DESD) AS QTD_PARCELA ');
    SQL.Add('      FROM ');
    SQL.Add('       PAGAR ');
    SQL.Add('      GROUP BY ');
    SQL.Add('       NUMTIT ');
    SQL.Add('      ) AS PARCELAS ON ');
    SQL.Add('       PAGAR.NUMTIT = PARCELAS.NUM_DOCTO ');
    SQL.Add('WHERE   ');
    SQL.Add('  SITUACAO <> ''CA'' ');
    SQL.Add('  AND  ');
    SQL.Add('  CODLOJA = 1 ');
    SQL.Add('  AND  ');
    SQL.Add('  CODTIPODOCUMENTO = 1 ');
    SQL.Add('  AND  ');
    SQL.Add('  CAST(DTENTRADA AS DATE) >= '''+FormatDateTime('yyyy-mm-dd',DtpInicial.Date)+''' ');
    SQL.Add('  AND');
    SQL.Add('  CAST(DTENTRADA AS DATE) <= '''+FormatDateTime('yyyy-mm-dd',DtpFinal.Date)+''' ');

    if Aberto = '1' then
      SQL.Add('AND PAGAR.SITUACAO = ''AB'' ')
    else
      SQL.Add('AND PAGAR.SITUACAO = ''BT'' ');

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

        Layout.FieldByName('DES_OBSERVACAO').AsString := StrLBReplace(QryPrincipal2.FieldByName('DES_OBSERVACAO').AsString);
        Layout.FieldByName('NUM_NF').AsString         := StrRetNums(QryPrincipal2.FieldByName('NUM_NF').AsString);
        Layout.FieldByName('NUM_CGC').AsString        := StrRetNums(QryPrincipal2.FieldByName('NUM_CGC').AsString);

        if QryPrincipal2.FieldByName('DTA_EMISSAO').AsString <> '' then
          Layout.FieldByName('DTA_EMISSAO').AsString := FormatDateTime('dd/mm/yyyy', QryPrincipal2.FieldByName('DTA_EMISSAO').AsDateTime);

        if QryPrincipal2.FieldByName('DTA_VENCIMENTO').AsString <> '' then
          Layout.FieldByName('DTA_VENCIMENTO').AsString:= FormatDateTime('dd/mm/yyyy', QryPrincipal2.FieldByName('DTA_VENCIMENTO').AsDateTime);

        if QryPrincipal2.FieldByName('DTA_QUITADA').AsString <> '' then
          Layout.FieldByName('DTA_QUITADA').AsString:= FormatDateTime('dd/mm/yyyy', QryPrincipal2.FieldByName('DTA_QUITADA').AsDateTime);

        if QryPrincipal2.FieldByName('DTA_PAGTO').AsString <> '' then
          Layout.FieldByName('DTA_PAGTO').AsString:= FormatDateTime('dd/mm/yyyy', QryPrincipal2.FieldByName('DTA_PAGTO').AsDateTime);

        if QryPrincipal2.FieldByName('DTA_ENTRADA').AsString <> '' then
          Layout.FieldByName('DTA_ENTRADA').AsString:= FormatDateTime('dd/mm/yyyy', QryPrincipal2.FieldByName('DTA_ENTRADA').AsDateTime);

        Layout.WriteLine;
      except
        On E: Exception do
        FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
      end;
      Next;
    end;
  end;
end;

procedure TFrmSmCasteloNext.GerarFinanceiroReceber(Aberto: String);
begin
  inherited;
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT  ');
    SQL.Add('  0 AS TIPO_PARCEIRO, ');
//    SQL.Add('  RECEBER.CODCLIE AS COD_PARCEIRO, ');
    SQL.Add('  0 AS COD_PARCEIRO, ');
    SQL.Add('  1 AS TIPO_CONTA, ');
    SQL.Add('   ');
    {SQL.Add('  CASE CODTIPODOCUMENTO  ');
    SQL.Add('      WHEN 2 THEN 2  -- CHEQUE A VISTA ');
    SQL.Add('      WHEN 3 THEN 3  -- CHEQUE A PRAZO ');
    SQL.Add('      WHEN 4 THEN 4  -- PROMISSORIA ');
    SQL.Add('      ELSE 4         -- CHEQUE DEVOLVIDO ');
    SQL.Add('  END AS COD_ENTIDADE, ');
    SQL.Add('   ');}

    SQL.Add('  4 AS COD_ENTIDADE,   ');
    SQL.Add('  COALESCE(NOTAECF, NUMTIT) AS NUM_DOCTO, --NUMTIT   ');
    SQL.Add('  999 AS COD_BANCO,   ');
    SQL.Add('  '''' AS DES_BANCO, ');
    SQL.Add('  DTEMISSAO AS DTA_EMISSAO, ');
    SQL.Add('  DTVENCTO AS DTA_VENCIMENTO, ');
    SQL.Add('   ');
    SQL.Add('  CASE WHEN RECEBER.SITUACAO = ''BT'' THEN VALOR ');
    SQL.Add('   ELSE (COALESCE(VALOR, 0) - COALESCE(VALORPAGO, 0)) ');
    SQL.Add('   END AS VAL_PARCELA, ');
    SQL.Add('    ');
    SQL.Add('  VALORJUROS AS VAL_JUROS, ');
    SQL.Add('  VALORDESC AS VAL_DESCONTO, ');
    SQL.Add('   ');
    SQL.Add('  CASE SITUACAO ');
    SQL.Add('    WHEN ''BT'' THEN ''S'' ');
    SQL.Add('    WHEN ''AB'' THEN ''N'' ');
    SQL.Add('    WHEN ''BP'' THEN ''N'' ');
    SQL.Add('  END AS FLG_QUITADO, ');
    SQL.Add('   ');
    SQL.Add('  CASE  ');
    SQL.Add('    WHEN RECEBER.SITUACAO = ''BT'' THEN ');
    SQL.Add('      CASE WHEN RECEBER.DTPAGTO <= RECEBER.DTEMISSAO THEN RECEBER.DTEMISSAO ELSE RECEBER.DTPAGTO END ');
    SQL.Add('    WHEN RECEBER.SITUACAO = ''AB'' THEN NULL ');
    SQL.Add('  END AS DTA_QUITADA, ');
    SQL.Add('   ');
    SQL.Add('  997 AS COD_CATEGORIA, ');
    SQL.Add('  997 AS COD_SUBCATEGORIA, ');
    SQL.Add('  1 AS NUM_PARCELA, ');
    SQL.Add('  1 AS QTD_PARCELA, ');
    SQL.Add('  '+CbxLoja.Text+' AS COD_LOJA,   ');
//    SQL.Add('  1 AS COD_LOJA,   ');
    SQL.Add('  CLIENTES.CNPJ_CPF AS NUM_CGC, ');
    SQL.Add('  0 AS NUM_BORDERO, ');
    SQL.Add('  0 AS NUM_NF, ');
    SQL.Add('  0 AS NUM_SERIE_NF, ');
    SQL.Add('  VALOR AS VAL_TOTAL_NF, ');
    SQL.Add('  COALESCE(OBS, '''')  + '' NUM TITULO: ''+ COALESCE(NUMTIT, '''') AS DES_OBSERVACAO, ');
    SQL.Add('  0 AS NUM_PDV, ');
    SQL.Add('  NOTAECF AS NUM_CUPOM_FISCAL, ');
    SQL.Add('  0 AS COD_MOTIVO, ');
    SQL.Add('  99999 AS COD_CONVENIO, ');
    SQL.Add('  0 AS COD_BIN, ');
    SQL.Add('  '''' AS DES_BANDEIRA, ');
    SQL.Add('  '''' AS DES_REDE_TEF, ');
    SQL.Add('  0 AS VAL_RETENCAO, ');
    SQL.Add('  2 AS COD_CONDICAO, ');
    SQL.Add('  DTPAGTO AS DTA_PAGTO, ');
    SQL.Add('  DTEMISSAO AS DTA_ENTRADA, ');
    SQL.Add('  0 AS NUM_NOSSO_NUMERO, ');
    SQL.Add('  '''' AS COD_BARRA, ');
    SQL.Add('  ''N'' AS FLG_BOLETO_EMIT, ');
    SQL.Add('  NULL AS NUM_CGC_CPF_TITULAR, ');
    SQL.Add('  NULL AS DES_TITULAR, ');
    SQL.Add('  30 AS NUM_CONDICAO, ');
    SQL.Add('  0 AS VAL_CREDITO, ');
    SQL.Add('  999 AS COD_BANCO_PGTO,      ');
    SQL.Add('  ''RECEBTO-1'' AS DES_CC, ');
    SQL.Add('  0 AS COD_BANDEIRA,      ');
    SQL.Add('  '''' AS DTA_PRORROGACAO,   ');
    SQL.Add('  1 AS NUM_SEQ_FIN,      ');
    SQL.Add('  0 AS COD_COBRANCA,      ');
    SQL.Add('  '''' AS DTA_COBRANCA,   ');
    SQL.Add('  ''N'' AS FLG_ACEITE,   ');
    SQL.Add('  0 AS TIPO_ACEITE     ');
    SQL.Add('FROM    ');
    SQL.Add('  RECEBER ');
    SQL.Add('    INNER JOIN CLIENTES ON (CLIENTES.CODCLIE = RECEBER.CODCLIE) ');
    SQL.Add('WHERE ');
//    SQL.Add('  CODTIPODOCUMENTO  IN (2,3,4,9) ');
//    SQL.Add('  AND  ');
    SQL.Add('  CODLOJA = 1 ');
    SQL.Add('  AND  ');
    SQL.Add('  SITUACAO <> ''CA'' ');
    SQL.Add('  AND  ');
    SQL.Add('  CAST(RECEBER.DTEMISSAO AS DATE) >= '''+FormatDateTime('yyyy-mm-dd',DtpInicial.Date)+''' ');
    SQL.Add('  AND');
    SQL.Add('  CAST(RECEBER.DTEMISSAO AS DATE) <= '''+FormatDateTime('yyyy-mm-dd',DtpFinal.Date)+''' ');

    if Aberto = '1' then
      SQL.Add('AND RECEBER.SITUACAO <> ''BT'' ')
    else
      SQL.Add('AND RECEBER.SITUACAO = ''BT'' ');

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

        Layout.FieldByName('DES_OBSERVACAO').AsString   := StrLBReplace(QryPrincipal2.FieldByName('DES_OBSERVACAO').AsString);
        Layout.FieldByName('NUM_NF').AsString           := StrRetNums(QryPrincipal2.FieldByName('NUM_NF').AsString);
        Layout.FieldByName('NUM_CUPOM_FISCAL').AsString := StrRetNums(QryPrincipal2.FieldByName('NUM_CUPOM_FISCAL').AsString);
        Layout.FieldByName('NUM_CGC').AsString          := StrRetNums(QryPrincipal2.FieldByName('NUM_CGC').AsString);

        if QryPrincipal2.FieldByName('DTA_QUITADA').AsString <> '' then
          Layout.FieldByName('DTA_QUITADA').AsString:= FormatDateTime('dd/mm/yyyy',QryPrincipal2.FieldByName('DTA_QUITADA').AsDateTime);

        if QryPrincipal2.FieldByName('DTA_PAGTO').AsString <> '' then
          Layout.FieldByName('DTA_PAGTO').AsString:= FormatDateTime('dd/mm/yyyy',QryPrincipal2.FieldByName('DTA_PAGTO').AsDateTime);

        if QryPrincipal2.FieldByName('DTA_EMISSAO').AsString <> '' then
          Layout.FieldByName('DTA_EMISSAO').AsString := FormatDateTime('dd/mm/yyyy',QryPrincipal2.FieldByName('DTA_EMISSAO').AsDateTime);

        if QryPrincipal2.FieldByName('DTA_VENCIMENTO').AsString <> '' then
          Layout.FieldByName('DTA_VENCIMENTO').AsString:= FormatDateTime('dd/mm/yyyy',QryPrincipal2.FieldByName('DTA_VENCIMENTO').AsDateTime);

        if QryPrincipal2.FieldByName('DTA_ENTRADA').AsString <> '' then
          Layout.FieldByName('DTA_ENTRADA').AsString:= FormatDateTime('dd/mm/yyyy',QryPrincipal2.FieldByName('DTA_ENTRADA').AsDateTime);

        Layout.WriteLine;
      except
        On E: Exception do
        FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
      end;
      Next;
    end;
  end;
end;

procedure TFrmSmCasteloNext.GerarFornecedor;
var
  Obs      : TStringList;
  QryFones : TADOQuery;
  observacao, email, inscEst : string;
begin
  inherited;

  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT ');
    SQL.Add('  CODFORNEC AS COD_FORNECEDOR, ');
    SQL.Add('  RAZAO AS DES_FORNECEDOR, ');
    SQL.Add('  CASE WHEN COALESCE(FANTASIA, '''') = '''' OR FANTASIA IS NULL THEN RAZAO ELSE FANTASIA END AS DES_FANTASIA, ');
    SQL.Add('  CNPJ_CPF AS NUM_CGC, ');
    SQL.Add('  IE AS NUM_INSC_EST, ');
    SQL.Add('  CASE WHEN COALESCE(ENDERECO, '''') = '''' OR ENDERECO IS NULL THEN ''A DEFINIR'' ELSE ENDERECO END AS DES_ENDERECO, ');
    SQL.Add('  CASE WHEN COALESCE(BAIRRO, '''')   = '''' OR BAIRRO IS NULL   THEN ''A DEFINIR'' ELSE BAIRRO   END AS DES_BAIRRO, ');
    SQL.Add('  CASE WHEN COALESCE(CIDADE, '''')   = '''' OR CIDADE IS NULL   THEN ''VITORIA''   ELSE CIDADE   END AS DES_CIDADE, ');
    SQL.Add('  CASE WHEN COALESCE(ESTADO, '''')   = '''' OR ESTADO IS NULL   THEN ''ES''        ELSE ESTADO   END AS DES_SIGLA, ');
    SQL.Add('  CASE WHEN COALESCE(CEP, '''')      = '''' OR CEP IS NULL      THEN ''29070330''  ELSE CEP      END AS NUM_CEP, ');
    SQL.Add('  TELEFONE AS NUM_FONE, ');
    SQL.Add('  FAX AS NUM_FAX, ');
    SQL.Add('  CONTATO AS DES_CONTATO, ');
    SQL.Add('  0 AS QTD_DIA_CARENCIA, ');
    SQL.Add('  PVISITA AS NUM_FREQ_VISITA, ');
    SQL.Add('  0 AS VAL_DESCONTO, ');
    SQL.Add('  PENTREGA AS NUM_PRAZO, ');
    SQL.Add('  ''N'' AS ACEITA_DEVOL_MER, ');
    SQL.Add('  ''N'' AS CAL_IPI_VAL_BRUTO, ');
    SQL.Add('  ''N'' AS CAL_ICMS_ENC_FIN, ');
    SQL.Add('  ''N'' AS CAL_ICMS_VAL_IPI, ');
    SQL.Add('   ');
    SQL.Add('  CASE MICROEPP ');
    SQL.Add('    WHEN ''S'' THEN ''S''  ');
    SQL.Add('    ELSE ''N'' ');
    SQL.Add('  END AS MICRO_EMPRESA, ');
    SQL.Add('   ');
    SQL.Add('  0 AS COD_FORNECEDOR_ANT, ');
    SQL.Add('  NUMERO AS NUM_ENDERECO, ');
    SQL.Add('  OBS AS DES_OBSERVACAO,  ');
    SQL.Add('  '''' AS DES_EMAIL, ');
    SQL.Add('  '''' AS DES_WEB_SITE, ');
    SQL.Add('  ''N'' AS FABRICANTE, ');
    SQL.Add('   ');
    SQL.Add('  CASE PRODRURAL ');
    SQL.Add('    WHEN ''S'' THEN 0  ');
    SQL.Add('    WHEN ''N'' THEN 1  ');
    SQL.Add('  END AS FLG_PRODUTOR_RURAL, ');
    SQL.Add('   ');
    SQL.Add('  0 AS TIPO_FRETE, ');
    SQL.Add('   ');
    SQL.Add('  CASE SIMPLES ');
    SQL.Add('    WHEN ''S'' THEN 0 ');
    SQL.Add('    WHEN ''N'' THEN 1   ');
    SQL.Add('  END AS FLG_SIMPLES, ');
    SQL.Add('   ');
    SQL.Add('  ''N'' AS FLG_SUBSTITUTO_TRIB,   ');
    SQL.Add('  0 AS COD_CONTACCFORN, ');
    SQL.Add('   ');
    SQL.Add('  CASE ATIVO ');
    SQL.Add('    WHEN ''S'' THEN ''N'' ');
    SQL.Add('  ELSE ''S'' ');
    SQL.Add('  END AS INATIVO, ');
    SQL.Add('    ');
    SQL.Add('  0 AS COD_CLASSIF, ');
    SQL.Add('  DTCAD AS DTA_CADASTRO, ');
    SQL.Add('  0 AS VAL_CREDITO, ');
    SQL.Add('  0 AS VAL_DEBITO, ');
    SQL.Add('  0 AS PED_MIN_VAL, ');
    SQL.Add('  EMAIL AS DES_EMAIL_VEND, ');
    SQL.Add('  '''' AS SENHA_COTACAO, ');
    SQL.Add('  -1 AS TIPO_PRODUTOR,  ');
    SQL.Add('  '''' AS NUM_CELULAR  ');
    SQL.Add('FROM  ');
    SQL.Add('  FORNECEDORES ');

    Open;
    First;

    QryFones := TADOQuery.Create(FrmProgresso);
    Obs      := TStringList.Create;

    with QryFones do
    begin
      Connection := ADOSQLServer;
      SQL.Add('SELECT');
      SQL.Add('FORNECEDORES.CODFORNEC AS COD_FORNECEDOR,');
      SQL.Add('FORNECEDORES.CELULAR,');
      SQL.Add('FORNECEDORES.FONE1');
      SQL.Add('FROM FORNECEDORES');
      SQL.Add('WHERE CODFORNEC = :COD_FORNECEDOR');
    end;

    NumLinha := 0;

    while not Eof do
    begin
      try
        if Cancelar then
        Break;
        Inc(NumLinha);

        Layout.SetValues(QryPrincipal2, NumLinha, RecordCount);

        Layout.FieldByName('DES_FORNECEDOR').AsString := StrSubstLtsAct(Layout.FieldByName('DES_FORNECEDOR').AsString);
        Layout.FieldByName('DES_FANTASIA').AsString   := StrSubstLtsAct(Layout.FieldByName('DES_FANTASIA').AsString);
        Layout.FieldByName('DES_BAIRRO').AsString     := StrSubstLtsAct(Layout.FieldByName('DES_BAIRRO').AsString);
        Layout.FieldByName('DES_ENDERECO').AsString   := StrSubstLtsAct(Layout.FieldByName('DES_ENDERECO').AsString);

        Layout.FieldByName('NUM_CGC').AsString        := StrRetNums(Layout.FieldByName('NUM_CGC').AsString);
        Layout.FieldByName('NUM_CEP').AsString        := StrRetNums(Layout.FieldByName('NUM_CEP').AsString);
        Layout.FieldByName('NUM_ENDERECO').AsString   := StrRetNums(Layout.FieldByName('NUM_ENDERECO').AsString);

        if( Layout.FieldByName('NUM_ENDERECO').AsString = '' ) then
           Layout.FieldByName('NUM_ENDERECO').AsString := 'S/N';

        if Length(Layout.FieldByName('NUM_CGC').AsString) > 11 then
        begin
          if not ValidaCGC(Layout.FieldByName('NUM_CGC').AsString) then
            Layout.FieldByName('NUM_CGC').AsString := '';
        end
        else
          if not ValidaCPF(Layout.FieldByName('NUM_CGC').AsString) then
            Layout.FieldByName('NUM_CGC').AsString := '';

        Layout.FieldByName('NUM_FONE').AsString := StrRetNums( FieldByName('NUM_FONE').AsString );
        Layout.FieldByName('NUM_FAX').AsString  := StrRetNums( FieldByName('NUM_FAX').AsString );

        observacao := StrReplace(StrLBReplace(FieldByName('DES_OBSERVACAO').AsString), '\n', '');
        email      := StrReplace(StrLBReplace(FieldByName('DES_EMAIL').AsString), '\n', '');
        inscEst    := StrRetNums(Layout.FieldByName('NUM_INSC_EST').AsString);

        if( inscEst = '' ) then
          Layout.FieldByName('NUM_INSC_EST').AsString := 'ISENTO'
        else begin
           if StrToFloat(inscEst) = 0 then
             Layout.FieldByName('NUM_INSC_EST').AsString := ''
           else
             Layout.FieldByName('NUM_INSC_EST').AsString := inscEst;
        end;

        Layout.FieldByName('DES_OBSERVACAO').AsString := observacao;
        Layout.FieldByName('DES_EMAIL').AsString      := email;

        Layout.FieldByName('DTA_CADASTRO').AsDateTime := Date;

        if Layout.FieldByName('NUM_CEP').AsString = '' then
          Layout.FieldByName('NUM_CEP').AsString := '29070330';

        if Layout.FieldByName('DES_ENDERECO').AsString = '' then
          Layout.FieldByName('DES_ENDERECO').AsString := 'A DEFINIR';

        if Layout.FieldByName('DES_BAIRRO').AsString = '' then
          Layout.FieldByName('DES_BAIRRO').AsString := 'A DEFINIR';

        if Layout.FieldByName('DES_CIDADE').AsString = '' then
          Layout.FieldByName('DES_CIDADE').AsString := 'VITORIA';

        if Layout.FieldByName('DES_SIGLA').AsString = '' then
          Layout.FieldByName('DES_SIGLA').AsString := 'ES';

        QryFones.Close;
        QryFones.Parameters.ParamByName('COD_FORNECEDOR').Value := FieldByName('COD_FORNECEDOR').AsInteger;
        QryFones.Open;

        if (QryFones.FieldByName('CELULAR').AsString <> '') then
          Obs.Add(('Celular: ')+QryFones.FieldByName('CELULAR').AsString);

        Layout.FieldByName('NUM_CELULAR').AsString := QryFones.FieldByName('CELULAR').AsString;

        if (QryFones.FieldByName('FONE1').AsString <> '') then
          Obs.Add(('TEL1: ')+QryFones.FieldByName('FONE1').AsString);

        Obs.Add(Layout.FieldByName('DES_OBSERVACAO').AsString);
        Layout.FieldByName('DES_OBSERVACAO').AsString := StrLBReplace(Obs.Text);
        Obs.Clear;

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

procedure TFrmSmCasteloNext.GerarCondPagForn;
begin
  inherited;

  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT   ');
    SQL.Add('  CODFORNEC AS COD_FORNECEDOR, ');
    SQL.Add('    ');
    SQL.Add('  CASE CONDPAGTO.DESCRICAO ');
    SQL.Add('    WHEN ''A VISTA'' THEN 0 ');
    SQL.Add('	   WHEN ''14 DIAS'' THEN 14 ');
    SQL.Add('	   WHEN ''21 DIAS'' THEN 21 ');
    SQL.Add('	   WHEN ''28 DIAS'' THEN 28 ');
    SQL.Add('	   WHEN ''30 DIAS'' THEN 30 ');
    SQL.Add('	   WHEN ''14/21/28 DIAS'' THEN 14 ');
    SQL.Add('    WHEN ''21/28/35 DIAS'' THEN 21 ');
    SQL.Add('  ELSE 30 END AS NUM_CONDICAO, ');
    SQL.Add(' ');
    SQL.Add('  2 AS COD_CONDICAO,   ');
    SQL.Add('  8 AS COD_ENTIDADE,   ');
    SQL.Add('  FORNECEDORES.CNPJ_CPF AS NUM_CGC   ');
    SQL.Add('FROM   ');
    SQL.Add('  FORNECEDORES ');
    SQL.Add('    LEFT JOIN CONDPAGTO ON FORNECEDORES.CODCONDPAGTO = CONDPAGTO.CODCONDPAGTO ');
    SQL.Add(' ');
    SQL.Add('UNION ');
    SQL.Add(' ');
    SQL.Add('SELECT   ');
    SQL.Add('  CODFORNEC AS COD_FORNECEDOR, ');
    SQL.Add('    ');
    SQL.Add('  CASE CONDPAGTO.DESCRICAO ');
    SQL.Add('    WHEN ''A VISTA'' THEN 0 ');
    SQL.Add('	   WHEN ''14 DIAS'' THEN 14 ');
    SQL.Add('	   WHEN ''21 DIAS'' THEN 21 ');
    SQL.Add('	   WHEN ''28 DIAS'' THEN 28 ');
    SQL.Add('	   WHEN ''30 DIAS'' THEN 30 ');
    SQL.Add('	   WHEN ''14/21/28 DIAS'' THEN 21 ');
    SQL.Add('    WHEN ''21/28/35 DIAS'' THEN 28 ');
    SQL.Add(' ELSE 30 END AS NUM_CONDICAO, ');
    SQL.Add(' ');
    SQL.Add('  2 AS COD_CONDICAO,   ');
    SQL.Add('  8 AS COD_ENTIDADE,   ');
    SQL.Add('  FORNECEDORES.CNPJ_CPF AS NUM_CGC   ');
    SQL.Add('FROM   ');
    SQL.Add('  FORNECEDORES ');
    SQL.Add('    LEFT JOIN CONDPAGTO ON FORNECEDORES.CODCONDPAGTO = CONDPAGTO.CODCONDPAGTO ');
    SQL.Add(' ');
    SQL.Add('UNION ');
    SQL.Add(' ');
    SQL.Add('SELECT   ');
    SQL.Add('  CODFORNEC AS COD_FORNECEDOR, ');
    SQL.Add('    ');
    SQL.Add('  CASE CONDPAGTO.DESCRICAO ');
    SQL.Add('    WHEN ''A VISTA'' THEN 0 ');
    SQL.Add('	   WHEN ''14 DIAS'' THEN 14 ');
    SQL.Add('	   WHEN ''21 DIAS'' THEN 21 ');
    SQL.Add('	   WHEN ''28 DIAS'' THEN 28 ');
    SQL.Add('	   WHEN ''30 DIAS'' THEN 30 ');
    SQL.Add('	   WHEN ''14/21/28 DIAS'' THEN 28 ');
    SQL.Add('    WHEN ''21/28/35 DIAS'' THEN 35 ');
    SQL.Add(' ELSE 30 END AS NUM_CONDICAO, ');
    SQL.Add(' ');
    SQL.Add('  2 AS COD_CONDICAO,   ');
    SQL.Add('  8 AS COD_ENTIDADE,   ');
    SQL.Add('  FORNECEDORES.CNPJ_CPF AS NUM_CGC   ');
    SQL.Add('FROM   ');
    SQL.Add('  FORNECEDORES ');
    SQL.Add('    LEFT JOIN CONDPAGTO ON FORNECEDORES.CODCONDPAGTO = CONDPAGTO.CODCONDPAGTO ');
    SQL.Add(' ');

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

        Layout.WriteLine;
      except
        On E: Exception do
        FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
      end;
      Next;
    end;
  end;
end;

procedure TFrmSmCasteloNext.GerarNCM;
var
 Count : Integer;
begin
  inherited;
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT * FROM ( ');
    //
    SQL.Add('SELECT DISTINCT   ');
    SQL.Add('  0 AS COD_NCM,  ');
    SQL.Add('  ''A DEFINIR'' AS DES_NCM,  ');
    SQL.Add('  CODNCM AS NUM_NCM,  ');
    SQL.Add('    ');
    SQL.Add('  CASE PRODUTOS.TIPO_COFINS    ');
    SQL.Add('    WHEN ''T'' THEN ''N''   ');
    SQL.Add('  ELSE ''S''     ');
    SQL.Add('  END AS FLG_NAO_PIS_COFINS,   ');
    SQL.Add('  ');
    SQL.Add('  CASE   ');
    SQL.Add('    WHEN ((PRODUTOS.CST_PISENTRADA = 73) AND (PRODUTOS.CST_PISSAIDA = 06))THEN 0  ');
    SQL.Add('    WHEN ((PRODUTOS.CST_PISENTRADA = 70) AND (PRODUTOS.CST_PISSAIDA = 04))THEN 1  ');
    SQL.Add('    WHEN ((PRODUTOS.CST_PISENTRADA = 75) AND (PRODUTOS.CST_PISSAIDA = 05))THEN 2  ');
    SQL.Add('    WHEN ((PRODUTOS.CST_PISENTRADA = 74) AND (PRODUTOS.CST_PISSAIDA = 09))THEN 3  ');
    SQL.Add('    WHEN ((PRODUTOS.CST_PISENTRADA = 99) AND (PRODUTOS.CST_PISSAIDA = 49))THEN 3  ');
    SQL.Add('    WHEN ((PRODUTOS.CST_PISENTRADA = 72) AND (PRODUTOS.CST_PISSAIDA = 09))THEN 4  ');
    SQL.Add('    WHEN ((PRODUTOS.CST_PISENTRADA = 50) AND (PRODUTOS.CST_PISSAIDA = 01))THEN 0  ');
    SQL.Add('  ELSE -1 END AS TIPO_NAO_PIS_COFINS,  ');
    SQL.Add('  ');
    SQL.Add('  COALESCE(PRODUTOS.NAT_REC, 0) AS COD_TAB_SPED,    ');
    SQL.Add(' ');
    SQL.Add('  CASE WHEN COALESCE(PRODUTOS.CODCEST, '''') = '''' OR PRODUTOS.CODCEST IS NULL THEN ''99.999.99'' ELSE PRODUTOS.CODCEST END AS NUM_CEST,    ');
    SQL.Add('    ');
    SQL.Add('  ''ES'' AS DES_SIGLA,      ');
    SQL.Add('    ');
    SQL.Add('  CASE     ');
    SQL.Add('    WHEN ((CODTRIB IN (200, 20, 41, 0))     ');
    SQL.Add('	     AND (CODALIQ_NF IN (''TC'', ''TB'', ''TA'', ''TG'', ''TC'', ''F'', ''I'', ''N''))    ');
    SQL.Add('	     AND (CODALIQ IN (''TC'', ''TB'', ''TA'', ''TG'', ''TC'', ''F'', ''I'', ''N''))    ');
    SQL.Add('	     AND (PER_REDUC IN (67.06, 0.00))) THEN 51 ');
    SQL.Add('		  ');
    SQL.Add('    WHEN ((CODTRIB IN (200, 20))          AND (CODALIQ_NF IN (''TC'', ''F'', ''I'', ''N'')) AND  ');
    SQL.Add('	  (CODALIQ IN (''TA'', ''TC'', ''TF'', ''F'', ''I'', ''N'')) AND (PER_REDUC IN (100.00))) THEN 66   ');
    SQL.Add('	 ');
    SQL.Add('	WHEN ((CODTRIB IN (60, 260))          AND (CODALIQ_NF IN (''F'', ''I'', ''N'')) AND (PER_REDUC IN (100.00, 0.00))) THEN 60     ');
    SQL.Add('    WHEN ((CODTRIB IN (0, 00, 000))       AND (CODALIQ_NF IN (''TC'', ''TB'')) AND (CODALIQ IN (''TB'')) AND (PER_REDUC IN (0.00))) THEN 3      ');
    SQL.Add('    WHEN ((CODTRIB IN (0, 00, 000, 20))   AND (CODALIQ_NF IN (''TC'')) AND (CODALIQ IN (''TA'')) AND (PER_REDUC IN (58.82))) THEN 52     ');
    SQL.Add('    WHEN ((CODTRIB IN (40, 41))           AND (CODALIQ_NF IN (''F'', ''I'', ''N'')) AND (PER_REDUC IN (0.00)))  THEN 1      ');
    SQL.Add('    WHEN ((CODTRIB IN (20))               AND (CODALIQ_NF IN (''TB'')) AND (CODALIQ IN (''TA'')) AND (PER_REDUC IN (41.66))) THEN 6      ');
    SQL.Add('    WHEN ((CODTRIB IN (0, 00, 000))       AND (CODALIQ_NF IN (''TA''))  AND (PER_REDUC IN (0.00))) THEN 2      ');
    SQL.Add('    WHEN ((CODTRIB IN (60))               AND (CODALIQ_NF IN (''F'', ''I'', ''N'')) AND (PER_REDUC IN (58.82))) THEN 56     ');
    SQL.Add('    WHEN ((CODTRIB IN (0, 00, 000))       AND (CODALIQ_NF IN (''TD'')) AND (PER_REDUC IN (0.00)))  THEN 5      ');
    SQL.Add('      ');
    SQL.Add('  ELSE 1 END AS COD_TRIB_ENTRADA, ');
    SQL.Add('            ');
    SQL.Add('  CASE     ');
    SQL.Add('    WHEN ((CODTRIB IN (200, 20, 41, 0))     ');
    SQL.Add('	     AND (CODALIQ_NF IN (''TC'', ''TB'', ''TA'', ''TG'', ''TC'', ''F'', ''I'', ''N''))    ');
    SQL.Add('	     AND (CODALIQ IN (''TC'', ''TB'', ''TA'', ''TG'', ''TC'', ''F'', ''I'', ''N''))    ');
    SQL.Add('	     AND (PER_REDUC IN (67.06, 0.00))) THEN 51 ');
    SQL.Add('		  ');
    SQL.Add('    WHEN ((CODTRIB IN (200, 20))          AND (CODALIQ_NF IN (''TC'', ''F'', ''I'', ''N'')) AND  ');
    SQL.Add('	  (CODALIQ IN (''TA'', ''TC'', ''TF'', ''F'', ''I'', ''N'')) AND (PER_REDUC IN (100.00))) THEN 66   ');
    SQL.Add('	 ');
    SQL.Add('	WHEN ((CODTRIB IN (60, 260))          AND (CODALIQ_NF IN (''F'', ''I'', ''N'')) AND (PER_REDUC IN (100.00, 0.00))) THEN 60     ');
    SQL.Add('    WHEN ((CODTRIB IN (0, 00, 000))       AND (CODALIQ_NF IN (''TC'', ''TB'')) AND (CODALIQ IN (''TB'')) AND (PER_REDUC IN (0.00))) THEN 3      ');
    SQL.Add('    WHEN ((CODTRIB IN (0, 00, 000, 20))   AND (CODALIQ_NF IN (''TC'')) AND (CODALIQ IN (''TA'')) AND (PER_REDUC IN (58.82))) THEN 52     ');
    SQL.Add('    WHEN ((CODTRIB IN (40, 41))           AND (CODALIQ_NF IN (''F'', ''I'', ''N'')) AND (PER_REDUC IN (0.00)))  THEN 1      ');
    SQL.Add('    WHEN ((CODTRIB IN (20))               AND (CODALIQ_NF IN (''TB'')) AND (CODALIQ IN (''TA'')) AND (PER_REDUC IN (41.66))) THEN 6      ');
    SQL.Add('    WHEN ((CODTRIB IN (0, 00, 000))       AND (CODALIQ_NF IN (''TA''))  AND (PER_REDUC IN (0.00))) THEN 2      ');
    SQL.Add('    WHEN ((CODTRIB IN (60))               AND (CODALIQ_NF IN (''F'', ''I'', ''N'')) AND (PER_REDUC IN (58.82))) THEN 56     ');
    SQL.Add('    WHEN ((CODTRIB IN (0, 00, 000))       AND (CODALIQ_NF IN (''TD'')) AND (PER_REDUC IN (0.00)))  THEN 5      ');
    SQL.Add('      ');
    SQL.Add('  ELSE 1 END AS COD_TRIB_SAIDA, ');
    SQL.Add('      ');
    SQL.Add('  0 AS PER_IVA,  ');
    SQL.Add('  0 AS PER_FCP_ST  ');
    SQL.Add('FROM   ');
    SQL.Add('  PRODUTOS   ');
    SQL.Add(' ');
    SQL.Add('WHERE  ');
    SQL.Add('  PRODUTOS.IMPORTAR_PRODUTO = ''S'' ');
    SQL.Add(' ');
    SQL.Add('UNION ALL ');
    SQL.Add(' ');
    SQL.Add('SELECT DISTINCT   ');
    SQL.Add('  0 AS COD_NCM,  ');
    SQL.Add('  ''A DEFINIR'' AS DES_NCM,  ');
    SQL.Add('  CODNCM AS NUM_NCM,  ');
    SQL.Add('    ');
    SQL.Add('  CASE PROD.TIPO_COFINS    ');
    SQL.Add('    WHEN ''T'' THEN ''N''   ');
    SQL.Add('  ELSE ''S''     ');
    SQL.Add('  END AS FLG_NAO_PIS_COFINS,   ');
    SQL.Add('  ');
    SQL.Add('  CASE   ');
    SQL.Add('    WHEN ((PROD.CST_PISENTRADA = 73) AND (PROD.CST_PISSAIDA = 06))THEN 0  ');
    SQL.Add('    WHEN ((PROD.CST_PISENTRADA = 70) AND (PROD.CST_PISSAIDA = 04))THEN 1  ');
    SQL.Add('    WHEN ((PROD.CST_PISENTRADA = 75) AND (PROD.CST_PISSAIDA = 05))THEN 2  ');
    SQL.Add('    WHEN ((PROD.CST_PISENTRADA = 74) AND (PROD.CST_PISSAIDA = 09))THEN 3  ');
    SQL.Add('    WHEN ((PROD.CST_PISENTRADA = 99) AND (PROD.CST_PISSAIDA = 49))THEN 3  ');
    SQL.Add('    WHEN ((PROD.CST_PISENTRADA = 72) AND (PROD.CST_PISSAIDA = 09))THEN 4  ');
    SQL.Add('    WHEN ((PROD.CST_PISENTRADA = 50) AND (PROD.CST_PISSAIDA = 01))THEN 0  ');
    SQL.Add('  ELSE -1 END AS TIPO_NAO_PIS_COFINS,  ');
    SQL.Add('  ');
    SQL.Add('  COALESCE(PROD.NAT_REC, 0) AS COD_TAB_SPED,  ');
    SQL.Add('    ');
    SQL.Add('  CASE WHEN COALESCE(PROD.CODCEST, '''') = '''' OR PROD.CODCEST IS NULL THEN ''99.999.99'' ELSE PROD.CODCEST END AS NUM_CEST,    ');
    SQL.Add('    ');
    SQL.Add('  ''ES'' AS DES_SIGLA,      ');
    SQL.Add('    ');
    SQL.Add('  CASE     ');
    SQL.Add('    WHEN ((CODTRIB IN (200, 20, 41, 0))     ');
    SQL.Add('	     AND (CODALIQ_NF IN (''TC'', ''TB'', ''TA'', ''TG'', ''TC'', ''F'', ''I'', ''N''))    ');
    SQL.Add('	     AND (CODALIQ IN (''TC'', ''TB'', ''TA'', ''TG'', ''TC'', ''F'', ''I'', ''N''))    ');
    SQL.Add('	     AND (PER_REDUC IN (67.06, 0.00))) THEN 51 ');
    SQL.Add('		  ');
    SQL.Add('    WHEN ((CODTRIB IN (200, 20))          AND (CODALIQ_NF IN (''TC'', ''F'', ''I'', ''N'')) AND  ');
    SQL.Add('	  (CODALIQ IN (''TA'', ''TC'', ''TF'', ''F'', ''I'', ''N'')) AND (PER_REDUC IN (100.00))) THEN 66   ');
    SQL.Add('	 ');
    SQL.Add('	WHEN ((CODTRIB IN (60, 260))          AND (CODALIQ_NF IN (''F'', ''I'', ''N'')) AND (PER_REDUC IN (100.00, 0.00))) THEN 60     ');
    SQL.Add('    WHEN ((CODTRIB IN (0, 00, 000))       AND (CODALIQ_NF IN (''TC'', ''TB'')) AND (CODALIQ IN (''TB'')) AND (PER_REDUC IN (0.00))) THEN 3      ');
    SQL.Add('    WHEN ((CODTRIB IN (0, 00, 000, 20))   AND (CODALIQ_NF IN (''TC'')) AND (CODALIQ IN (''TA'')) AND (PER_REDUC IN (58.82))) THEN 52     ');
    SQL.Add('    WHEN ((CODTRIB IN (40, 41))           AND (CODALIQ_NF IN (''F'', ''I'', ''N'')) AND (PER_REDUC IN (0.00)))  THEN 1      ');
    SQL.Add('    WHEN ((CODTRIB IN (20))               AND (CODALIQ_NF IN (''TB'')) AND (CODALIQ IN (''TA'')) AND (PER_REDUC IN (41.66))) THEN 6      ');
    SQL.Add('    WHEN ((CODTRIB IN (0, 00, 000))       AND (CODALIQ_NF IN (''TA''))  AND (PER_REDUC IN (0.00))) THEN 2      ');
    SQL.Add('    WHEN ((CODTRIB IN (60))               AND (CODALIQ_NF IN (''F'', ''I'', ''N'')) AND (PER_REDUC IN (58.82))) THEN 56     ');
    SQL.Add('    WHEN ((CODTRIB IN (0, 00, 000))       AND (CODALIQ_NF IN (''TD'')) AND (PER_REDUC IN (0.00)))  THEN 5      ');
    SQL.Add('      ');
    SQL.Add('  ELSE 1 END AS COD_TRIB_ENTRADA, ');
    SQL.Add('            ');
    SQL.Add('  CASE     ');
    SQL.Add('    WHEN ((CODTRIB IN (200, 20, 41, 0))     ');
    SQL.Add('	     AND (CODALIQ_NF IN (''TC'', ''TB'', ''TA'', ''TG'', ''TC'', ''F'', ''I'', ''N''))    ');
    SQL.Add('	     AND (CODALIQ IN (''TC'', ''TB'', ''TA'', ''TG'', ''TC'', ''F'', ''I'', ''N''))    ');
    SQL.Add('	     AND (PER_REDUC IN (67.06, 0.00))) THEN 51 ');
    SQL.Add('		  ');
    SQL.Add('    WHEN ((CODTRIB IN (200, 20))          AND (CODALIQ_NF IN (''TC'', ''F'', ''I'', ''N'')) AND  ');
    SQL.Add('	  (CODALIQ IN (''TA'', ''TC'', ''TF'', ''F'', ''I'', ''N'')) AND (PER_REDUC IN (100.00))) THEN 66   ');
    SQL.Add('	 ');
    SQL.Add('	WHEN ((CODTRIB IN (60, 260))          AND (CODALIQ_NF IN (''F'', ''I'', ''N'')) AND (PER_REDUC IN (100.00, 0.00))) THEN 60     ');
    SQL.Add('    WHEN ((CODTRIB IN (0, 00, 000))       AND (CODALIQ_NF IN (''TC'', ''TB'')) AND (CODALIQ IN (''TB'')) AND (PER_REDUC IN (0.00))) THEN 3      ');
    SQL.Add('    WHEN ((CODTRIB IN (0, 00, 000, 20))   AND (CODALIQ_NF IN (''TC'')) AND (CODALIQ IN (''TA'')) AND (PER_REDUC IN (58.82))) THEN 52     ');
    SQL.Add('    WHEN ((CODTRIB IN (40, 41))           AND (CODALIQ_NF IN (''F'', ''I'', ''N'')) AND (PER_REDUC IN (0.00)))  THEN 1      ');
    SQL.Add('    WHEN ((CODTRIB IN (20))               AND (CODALIQ_NF IN (''TB'')) AND (CODALIQ IN (''TA'')) AND (PER_REDUC IN (41.66))) THEN 6      ');
    SQL.Add('    WHEN ((CODTRIB IN (0, 00, 000))       AND (CODALIQ_NF IN (''TA''))  AND (PER_REDUC IN (0.00))) THEN 2      ');
    SQL.Add('    WHEN ((CODTRIB IN (60))               AND (CODALIQ_NF IN (''F'', ''I'', ''N'')) AND (PER_REDUC IN (58.82))) THEN 56     ');
    SQL.Add('    WHEN ((CODTRIB IN (0, 00, 000))       AND (CODALIQ_NF IN (''TD'')) AND (PER_REDUC IN (0.00)))  THEN 5      ');
    SQL.Add('      ');
    SQL.Add('  ELSE 1 END AS COD_TRIB_SAIDA, ');
    SQL.Add('      ');
    SQL.Add('  0 AS PER_IVA,  ');
    SQL.Add('  0 AS PER_FCP_ST  ');
    SQL.Add('FROM   ');
    SQL.Add('  UNIFICACAO_PRODUTOS PROD   ');
    SQL.Add(' ');
    SQL.Add('WHERE   ');
    SQL.Add('  LEN(SUBSTRING(BARRA, PATINDEX(''%[^0]%'', BARRA+''.''), LEN(BARRA))) >= 8  ');
    SQL.Add('  AND  ');
    SQL.Add('  PROD.IMPORTAR_PRODUTO = ''S''  ');
    SQL.Add('  AND ');
    SQL.Add('  (SELECT COUNT(*) FROM PRODUTOS WHERE BARRA = PROD.BARRA) = 0 ');
//    SQL.Add('  AND ');
//    SQL.Add('  (SELECT COUNT(*) FROM PRODUTOS WHERE CODPROD = PROD.CODPROD) = 0 ');

    //
     SQL.Add(') AS NCM ORDER BY NCM.NUM_NCM ');


    Open;
    First;

    Count := 0;

    NumLinha := 0;

    while not Eof do
    begin
      try
        if Cancelar then
          Break;

        Inc(NumLinha);
        Inc(Count);
        Layout.SetValues(QryPrincipal2, NumLinha, RecordCount);

        Layout.FieldByName('COD_NCM').AsInteger := Count;

        if (Layout.FieldByName('DES_NCM').AsString = '')  then
          Layout.FieldByName('DES_NCM').AsString := 'A DEFINIR'
        else
          Layout.FieldByName('DES_NCM').AsString := Layout.FieldByName('DES_NCM').AsString;

        Layout.FieldByName('NUM_NCM').AsString  := StrRetNums(Layout.FieldByName('NUM_NCM').AsString);
        Layout.FieldByName('NUM_CEST').AsString := StrRetNums(Layout.FieldByName('NUM_CEST').AsString);

        Layout.WriteLine;
      except
        On E: Exception do
        FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
      end;
      Next;
    end;
  end;
end;

procedure TFrmSmCasteloNext.GerarNCMUF;
var
 count : Integer;
begin
  inherited;
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;


    SQL.Add('SELECT * FROM ( ');
    //
    SQL.Add('SELECT DISTINCT   ');
    SQL.Add('  0 AS COD_NCM,  ');
    SQL.Add('  ''A DEFINIR'' AS DES_NCM,  ');
    SQL.Add('  CODNCM AS NUM_NCM,  ');
    SQL.Add('    ');
    SQL.Add('  CASE PRODUTOS.TIPO_COFINS    ');
    SQL.Add('    WHEN ''T'' THEN ''N''   ');
    SQL.Add('  ELSE ''S''     ');
    SQL.Add('  END AS FLG_NAO_PIS_COFINS,   ');
    SQL.Add('  ');
    SQL.Add('  CASE   ');
    SQL.Add('    WHEN ((PRODUTOS.CST_PISENTRADA = 73) AND (PRODUTOS.CST_PISSAIDA = 06))THEN 0  ');
    SQL.Add('    WHEN ((PRODUTOS.CST_PISENTRADA = 70) AND (PRODUTOS.CST_PISSAIDA = 04))THEN 1  ');
    SQL.Add('    WHEN ((PRODUTOS.CST_PISENTRADA = 75) AND (PRODUTOS.CST_PISSAIDA = 05))THEN 2  ');
    SQL.Add('    WHEN ((PRODUTOS.CST_PISENTRADA = 74) AND (PRODUTOS.CST_PISSAIDA = 09))THEN 3  ');
    SQL.Add('    WHEN ((PRODUTOS.CST_PISENTRADA = 99) AND (PRODUTOS.CST_PISSAIDA = 49))THEN 3  ');
    SQL.Add('    WHEN ((PRODUTOS.CST_PISENTRADA = 72) AND (PRODUTOS.CST_PISSAIDA = 09))THEN 4  ');
    SQL.Add('    WHEN ((PRODUTOS.CST_PISENTRADA = 50) AND (PRODUTOS.CST_PISSAIDA = 01))THEN 0  ');
    SQL.Add('  ELSE -1 END AS TIPO_NAO_PIS_COFINS,  ');
    SQL.Add('  ');
    SQL.Add('  COALESCE(PRODUTOS.NAT_REC, 0) AS COD_TAB_SPED,    ');
    SQL.Add(' ');
    SQL.Add('  CASE WHEN COALESCE(PRODUTOS.CODCEST, '''') = '''' OR PRODUTOS.CODCEST IS NULL THEN ''99.999.99'' ELSE PRODUTOS.CODCEST END AS NUM_CEST,    ');
    SQL.Add('    ');
    SQL.Add('  ''ES'' AS DES_SIGLA,      ');
    SQL.Add('    ');
    SQL.Add('  CASE     ');
    SQL.Add('    WHEN ((CODTRIB IN (200, 20, 41, 0))     ');
    SQL.Add('	     AND (CODALIQ_NF IN (''TC'', ''TB'', ''TA'', ''TG'', ''TC'', ''F'', ''I'', ''N''))    ');
    SQL.Add('	     AND (CODALIQ IN (''TC'', ''TB'', ''TA'', ''TG'', ''TC'', ''F'', ''I'', ''N''))    ');
    SQL.Add('	     AND (PER_REDUC IN (67.06, 0.00))) THEN 51 ');
    SQL.Add('		  ');
    SQL.Add('    WHEN ((CODTRIB IN (200, 20))          AND (CODALIQ_NF IN (''TC'', ''F'', ''I'', ''N'')) AND  ');
    SQL.Add('	  (CODALIQ IN (''TA'', ''TC'', ''TF'', ''F'', ''I'', ''N'')) AND (PER_REDUC IN (100.00))) THEN 66   ');
    SQL.Add('	 ');
    SQL.Add('	WHEN ((CODTRIB IN (60, 260))          AND (CODALIQ_NF IN (''F'', ''I'', ''N'')) AND (PER_REDUC IN (100.00, 0.00))) THEN 60     ');
    SQL.Add('    WHEN ((CODTRIB IN (0, 00, 000))       AND (CODALIQ_NF IN (''TC'', ''TB'')) AND (CODALIQ IN (''TB'')) AND (PER_REDUC IN (0.00))) THEN 3      ');
    SQL.Add('    WHEN ((CODTRIB IN (0, 00, 000, 20))   AND (CODALIQ_NF IN (''TC'')) AND (CODALIQ IN (''TA'')) AND (PER_REDUC IN (58.82))) THEN 52     ');
    SQL.Add('    WHEN ((CODTRIB IN (40, 41))           AND (CODALIQ_NF IN (''F'', ''I'', ''N'')) AND (PER_REDUC IN (0.00)))  THEN 1      ');
    SQL.Add('    WHEN ((CODTRIB IN (20))               AND (CODALIQ_NF IN (''TB'')) AND (CODALIQ IN (''TA'')) AND (PER_REDUC IN (41.66))) THEN 6      ');
    SQL.Add('    WHEN ((CODTRIB IN (0, 00, 000))       AND (CODALIQ_NF IN (''TA''))  AND (PER_REDUC IN (0.00))) THEN 2      ');
    SQL.Add('    WHEN ((CODTRIB IN (60))               AND (CODALIQ_NF IN (''F'', ''I'', ''N'')) AND (PER_REDUC IN (58.82))) THEN 56     ');
    SQL.Add('    WHEN ((CODTRIB IN (0, 00, 000))       AND (CODALIQ_NF IN (''TD'')) AND (PER_REDUC IN (0.00)))  THEN 5      ');
    SQL.Add('      ');
    SQL.Add('  ELSE 1 END AS COD_TRIB_ENTRADA, ');
    SQL.Add('            ');
    SQL.Add('  CASE     ');
    SQL.Add('    WHEN ((CODTRIB IN (200, 20, 41, 0))     ');
    SQL.Add('	     AND (CODALIQ_NF IN (''TC'', ''TB'', ''TA'', ''TG'', ''TC'', ''F'', ''I'', ''N''))    ');
    SQL.Add('	     AND (CODALIQ IN (''TC'', ''TB'', ''TA'', ''TG'', ''TC'', ''F'', ''I'', ''N''))    ');
    SQL.Add('	     AND (PER_REDUC IN (67.06, 0.00))) THEN 51 ');
    SQL.Add('		  ');
    SQL.Add('    WHEN ((CODTRIB IN (200, 20))          AND (CODALIQ_NF IN (''TC'', ''F'', ''I'', ''N'')) AND  ');
    SQL.Add('	  (CODALIQ IN (''TA'', ''TC'', ''TF'', ''F'', ''I'', ''N'')) AND (PER_REDUC IN (100.00))) THEN 66   ');
    SQL.Add('	 ');
    SQL.Add('	WHEN ((CODTRIB IN (60, 260))          AND (CODALIQ_NF IN (''F'', ''I'', ''N'')) AND (PER_REDUC IN (100.00, 0.00))) THEN 60     ');
    SQL.Add('    WHEN ((CODTRIB IN (0, 00, 000))       AND (CODALIQ_NF IN (''TC'', ''TB'')) AND (CODALIQ IN (''TB'')) AND (PER_REDUC IN (0.00))) THEN 3      ');
    SQL.Add('    WHEN ((CODTRIB IN (0, 00, 000, 20))   AND (CODALIQ_NF IN (''TC'')) AND (CODALIQ IN (''TA'')) AND (PER_REDUC IN (58.82))) THEN 52     ');
    SQL.Add('    WHEN ((CODTRIB IN (40, 41))           AND (CODALIQ_NF IN (''F'', ''I'', ''N'')) AND (PER_REDUC IN (0.00)))  THEN 1      ');
    SQL.Add('    WHEN ((CODTRIB IN (20))               AND (CODALIQ_NF IN (''TB'')) AND (CODALIQ IN (''TA'')) AND (PER_REDUC IN (41.66))) THEN 6      ');
    SQL.Add('    WHEN ((CODTRIB IN (0, 00, 000))       AND (CODALIQ_NF IN (''TA''))  AND (PER_REDUC IN (0.00))) THEN 2      ');
    SQL.Add('    WHEN ((CODTRIB IN (60))               AND (CODALIQ_NF IN (''F'', ''I'', ''N'')) AND (PER_REDUC IN (58.82))) THEN 56     ');
    SQL.Add('    WHEN ((CODTRIB IN (0, 00, 000))       AND (CODALIQ_NF IN (''TD'')) AND (PER_REDUC IN (0.00)))  THEN 5      ');
    SQL.Add('      ');
    SQL.Add('  ELSE 1 END AS COD_TRIB_SAIDA, ');
    SQL.Add('      ');
    SQL.Add('  0 AS PER_IVA,  ');
    SQL.Add('  0 AS PER_FCP_ST  ');
    SQL.Add('FROM   ');
    SQL.Add('  PRODUTOS   ');
    SQL.Add(' ');
    SQL.Add('WHERE  ');
    SQL.Add('  PRODUTOS.IMPORTAR_PRODUTO = ''S'' ');
    SQL.Add(' ');
    SQL.Add('UNION ALL ');
    SQL.Add(' ');
    SQL.Add('SELECT DISTINCT   ');
    SQL.Add('  0 AS COD_NCM,  ');
    SQL.Add('  ''A DEFINIR'' AS DES_NCM,  ');
    SQL.Add('  CODNCM AS NUM_NCM,  ');
    SQL.Add('    ');
    SQL.Add('  CASE PROD.TIPO_COFINS    ');
    SQL.Add('    WHEN ''T'' THEN ''N''   ');
    SQL.Add('  ELSE ''S''     ');
    SQL.Add('  END AS FLG_NAO_PIS_COFINS,   ');
    SQL.Add('  ');
    SQL.Add('  CASE   ');
    SQL.Add('    WHEN ((PROD.CST_PISENTRADA = 73) AND (PROD.CST_PISSAIDA = 06))THEN 0  ');
    SQL.Add('    WHEN ((PROD.CST_PISENTRADA = 70) AND (PROD.CST_PISSAIDA = 04))THEN 1  ');
    SQL.Add('    WHEN ((PROD.CST_PISENTRADA = 75) AND (PROD.CST_PISSAIDA = 05))THEN 2  ');
    SQL.Add('    WHEN ((PROD.CST_PISENTRADA = 74) AND (PROD.CST_PISSAIDA = 09))THEN 3  ');
    SQL.Add('    WHEN ((PROD.CST_PISENTRADA = 99) AND (PROD.CST_PISSAIDA = 49))THEN 3  ');
    SQL.Add('    WHEN ((PROD.CST_PISENTRADA = 72) AND (PROD.CST_PISSAIDA = 09))THEN 4  ');
    SQL.Add('    WHEN ((PROD.CST_PISENTRADA = 50) AND (PROD.CST_PISSAIDA = 01))THEN 0  ');
    SQL.Add('  ELSE -1 END AS TIPO_NAO_PIS_COFINS,  ');
    SQL.Add('  ');
    SQL.Add('  COALESCE(PROD.NAT_REC, 0) AS COD_TAB_SPED,  ');
    SQL.Add('    ');
    SQL.Add('  CASE WHEN COALESCE(PROD.CODCEST, '''') = '''' OR PROD.CODCEST IS NULL THEN ''99.999.99'' ELSE PROD.CODCEST END AS NUM_CEST,    ');
    SQL.Add('    ');
    SQL.Add('  ''ES'' AS DES_SIGLA,      ');
    SQL.Add('    ');
    SQL.Add('  CASE     ');
    SQL.Add('    WHEN ((CODTRIB IN (200, 20, 41, 0))     ');
    SQL.Add('	     AND (CODALIQ_NF IN (''TC'', ''TB'', ''TA'', ''TG'', ''TC'', ''F'', ''I'', ''N''))    ');
    SQL.Add('	     AND (CODALIQ IN (''TC'', ''TB'', ''TA'', ''TG'', ''TC'', ''F'', ''I'', ''N''))    ');
    SQL.Add('	     AND (PER_REDUC IN (67.06, 0.00))) THEN 51 ');
    SQL.Add('		  ');
    SQL.Add('    WHEN ((CODTRIB IN (200, 20))          AND (CODALIQ_NF IN (''TC'', ''F'', ''I'', ''N'')) AND  ');
    SQL.Add('	  (CODALIQ IN (''TA'', ''TC'', ''TF'', ''F'', ''I'', ''N'')) AND (PER_REDUC IN (100.00))) THEN 66   ');
    SQL.Add('	 ');
    SQL.Add('	WHEN ((CODTRIB IN (60, 260))          AND (CODALIQ_NF IN (''F'', ''I'', ''N'')) AND (PER_REDUC IN (100.00, 0.00))) THEN 60     ');
    SQL.Add('    WHEN ((CODTRIB IN (0, 00, 000))       AND (CODALIQ_NF IN (''TC'', ''TB'')) AND (CODALIQ IN (''TB'')) AND (PER_REDUC IN (0.00))) THEN 3      ');
    SQL.Add('    WHEN ((CODTRIB IN (0, 00, 000, 20))   AND (CODALIQ_NF IN (''TC'')) AND (CODALIQ IN (''TA'')) AND (PER_REDUC IN (58.82))) THEN 52     ');
    SQL.Add('    WHEN ((CODTRIB IN (40, 41))           AND (CODALIQ_NF IN (''F'', ''I'', ''N'')) AND (PER_REDUC IN (0.00)))  THEN 1      ');
    SQL.Add('    WHEN ((CODTRIB IN (20))               AND (CODALIQ_NF IN (''TB'')) AND (CODALIQ IN (''TA'')) AND (PER_REDUC IN (41.66))) THEN 6      ');
    SQL.Add('    WHEN ((CODTRIB IN (0, 00, 000))       AND (CODALIQ_NF IN (''TA''))  AND (PER_REDUC IN (0.00))) THEN 2      ');
    SQL.Add('    WHEN ((CODTRIB IN (60))               AND (CODALIQ_NF IN (''F'', ''I'', ''N'')) AND (PER_REDUC IN (58.82))) THEN 56     ');
    SQL.Add('    WHEN ((CODTRIB IN (0, 00, 000))       AND (CODALIQ_NF IN (''TD'')) AND (PER_REDUC IN (0.00)))  THEN 5      ');
    SQL.Add('      ');
    SQL.Add('  ELSE 1 END AS COD_TRIB_ENTRADA, ');
    SQL.Add('            ');
    SQL.Add('  CASE     ');
    SQL.Add('    WHEN ((CODTRIB IN (200, 20, 41, 0))     ');
    SQL.Add('	     AND (CODALIQ_NF IN (''TC'', ''TB'', ''TA'', ''TG'', ''TC'', ''F'', ''I'', ''N''))    ');
    SQL.Add('	     AND (CODALIQ IN (''TC'', ''TB'', ''TA'', ''TG'', ''TC'', ''F'', ''I'', ''N''))    ');
    SQL.Add('	     AND (PER_REDUC IN (67.06, 0.00))) THEN 51 ');
    SQL.Add('		  ');
    SQL.Add('    WHEN ((CODTRIB IN (200, 20))          AND (CODALIQ_NF IN (''TC'', ''F'', ''I'', ''N'')) AND  ');
    SQL.Add('	  (CODALIQ IN (''TA'', ''TC'', ''TF'', ''F'', ''I'', ''N'')) AND (PER_REDUC IN (100.00))) THEN 66   ');
    SQL.Add('	 ');
    SQL.Add('	WHEN ((CODTRIB IN (60, 260))          AND (CODALIQ_NF IN (''F'', ''I'', ''N'')) AND (PER_REDUC IN (100.00, 0.00))) THEN 60     ');
    SQL.Add('    WHEN ((CODTRIB IN (0, 00, 000))       AND (CODALIQ_NF IN (''TC'', ''TB'')) AND (CODALIQ IN (''TB'')) AND (PER_REDUC IN (0.00))) THEN 3      ');
    SQL.Add('    WHEN ((CODTRIB IN (0, 00, 000, 20))   AND (CODALIQ_NF IN (''TC'')) AND (CODALIQ IN (''TA'')) AND (PER_REDUC IN (58.82))) THEN 52     ');
    SQL.Add('    WHEN ((CODTRIB IN (40, 41))           AND (CODALIQ_NF IN (''F'', ''I'', ''N'')) AND (PER_REDUC IN (0.00)))  THEN 1      ');
    SQL.Add('    WHEN ((CODTRIB IN (20))               AND (CODALIQ_NF IN (''TB'')) AND (CODALIQ IN (''TA'')) AND (PER_REDUC IN (41.66))) THEN 6      ');
    SQL.Add('    WHEN ((CODTRIB IN (0, 00, 000))       AND (CODALIQ_NF IN (''TA''))  AND (PER_REDUC IN (0.00))) THEN 2      ');
    SQL.Add('    WHEN ((CODTRIB IN (60))               AND (CODALIQ_NF IN (''F'', ''I'', ''N'')) AND (PER_REDUC IN (58.82))) THEN 56     ');
    SQL.Add('    WHEN ((CODTRIB IN (0, 00, 000))       AND (CODALIQ_NF IN (''TD'')) AND (PER_REDUC IN (0.00)))  THEN 5      ');
    SQL.Add('      ');
    SQL.Add('  ELSE 1 END AS COD_TRIB_SAIDA, ');
    SQL.Add('      ');
    SQL.Add('  0 AS PER_IVA,  ');
    SQL.Add('  0 AS PER_FCP_ST  ');
    SQL.Add('FROM   ');
    SQL.Add('  UNIFICACAO_PRODUTOS PROD   ');
    SQL.Add(' ');
    SQL.Add('WHERE   ');
    SQL.Add('  LEN(SUBSTRING(BARRA, PATINDEX(''%[^0]%'', BARRA+''.''), LEN(BARRA))) >= 8  ');
    SQL.Add('  AND  ');
    SQL.Add('  PROD.IMPORTAR_PRODUTO = ''S''  ');
    SQL.Add('  AND ');
    SQL.Add('  (SELECT COUNT(*) FROM PRODUTOS WHERE BARRA = PROD.BARRA) = 0 ');
//    SQL.Add('  AND ');
//    SQL.Add('  (SELECT COUNT(*) FROM PRODUTOS WHERE CODPROD = PROD.CODPROD) = 0 ');
    //
     SQL.Add(') AS NCM ORDER BY NCM.NUM_NCM ');

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

procedure TFrmSmCasteloNext.GerarNFFornec;
begin
  inherited;
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT ');
    SQL.Add('  ENTRADANF.CODFORNEC AS COD_FORNECEDOR, ');
    SQL.Add('  ENTRADANF.NUMNOTA AS NUM_NF_FORN, ');
    SQL.Add('  ENTRADANF.SERIE AS NUM_SERIE_NF, ');
    SQL.Add('  ENTRADANF.SUBSERIE NUM_SUBSERIE_NF, ');
    SQL.Add('  ENTRADANF.CODNATOPER AS CFOP, ');
    SQL.Add('  CASE WHEN NATOPER.DESCRICAO LIKE ''DEV%'' THEN 2 ELSE 0 END AS TIPO_NF, ');
    SQL.Add('  ''NFE'' AS DES_ESPECIE, ');
    SQL.Add('  ENTRADANF.VALORNOTA AS VAL_TOTAL_NF, ');
    SQL.Add('  ENTRADANF.DTEMISSAO AS DTA_EMISSAO, ');
    SQL.Add('  ENTRADANF.DTENTRADA AS DTA_ENTRADA, ');
    SQL.Add('  ENTRADANF.VALORIPI AS VAL_TOTAL_IPI, ');
    SQL.Add('  0 AS VAL_VENDA_VAREJO, ');
    SQL.Add('  ENTRADANF.VALORFRETE AS VAL_FRETE, ');
    SQL.Add('  0 AS VAL_ACRESCIMO, ');
    SQL.Add('  ENTRADANF.VALORDESC AS VAL_DESCONTO, ');
    SQL.Add('  NULL AS NUM_CGC, ');
    SQL.Add('  ENTRADANF.BASEICMS AS VAL_TOTAL_BC, ');
    SQL.Add('  ENTRADANF.VALORICMS AS VAL_TOTAL_ICMS, ');
    SQL.Add('  ENTRADANF.BASEICMSSUBST AS VAL_BC_SUBST, ');
    SQL.Add('  ENTRADANF.VALORICMSSUBST AS VAL_ICMS_SUBST, ');
    SQL.Add('  0 AS VAL_FUNRURAL, ');
    SQL.Add('   ');
    SQL.Add('  CASE  ');
    SQL.Add('    WHEN NATOPER.DESCRICAO LIKE ''%BONIFICA%'' THEN 5 ');
    SQL.Add('  ELSE 1 END AS COD_PERFIL, ');
    SQL.Add('   ');
    SQL.Add('  0 AS VAL_DESP_ACESS, ');
    SQL.Add('  ''N'' FLG_CANCELADO, ');
    SQL.Add('  NULL AS DES_OBSERVACAO, ');
    SQL.Add('  ENTRADANF.CHAVEACESSO AS NUM_CHAVE_ACESSO ');
    SQL.Add('FROM ');
    SQL.Add('  ENTRADANF ');
    SQL.Add('    LEFT JOIN NATOPER ON ENTRADANF.CODNATOPER = NATOPER.CODNATOPER ');
    SQL.Add('WHERE ');
    SQL.Add('  CAST(ENTRADANF.DTEMISSAO AS DATE) >= '''+FormatDateTime('yyyy-mm-dd',DtpInicial.Date)+''' ');
    SQL.Add('  AND  ');
    SQL.Add('  CAST(ENTRADANF.DTEMISSAO AS DATE) <= '''+FormatDateTime('yyyy-mm-dd',DtpFinal.Date)+''' ');
    SQL.Add('   ');
    SQL.Add('  UNION ALL ');
    SQL.Add('   ');
    SQL.Add('SELECT ');
    SQL.Add('  NULL AS COD_FORNECEDOR, ');
    SQL.Add('  NF.NUMNF AS NUM_NF_FORN, ');
    SQL.Add('  NF.SERIE AS NUM_SERIE_NF, ');
    SQL.Add('  NF.SUBSERIE NUM_SUBSERIE_NF, ');
    SQL.Add('  NF.CODNATOPER AS CFOP, ');
    SQL.Add('  2 AS TIPO_NF, ');
    SQL.Add('  ''NFE'' AS DES_ESPECIE, ');
    SQL.Add('  NF.TOTAL_NF AS VAL_TOTAL_NF, ');
    SQL.Add('  NF.DTEMISSAO AS DTA_EMISSAO, ');
    SQL.Add('  NF.DTSAIDA AS DTA_ENTRADA, ');
    SQL.Add('  NF.VALOR_IPI AS VAL_TOTAL_IPI, ');
    SQL.Add('  0 AS VAL_VENDA_VAREJO, ');
    SQL.Add('  NF.FRETE AS VAL_FRETE, ');
    SQL.Add('  0 AS VAL_ACRESCIMO, ');
    SQL.Add('  NF.VALDESCONTO AS VAL_DESCONTO, ');
    SQL.Add('  CLIENTES.CNPJ_CPF AS NUM_CGC, ');
    SQL.Add('  NF.BASE_ICMS AS VAL_TOTAL_BC, ');
    SQL.Add('  NF.VALOR_ICMS AS VAL_TOTAL_ICMS, ');
    SQL.Add('  NF.BASE_SUB AS VAL_BC_SUBST, ');
    SQL.Add('  NF.VALOR_SUB AS VAL_ICMS_SUBST, ');
    SQL.Add('  0 AS VAL_FUNRURAL, ');
    SQL.Add('  0 AS COD_PERFIL, ');
    SQL.Add('  0 AS VAL_DESP_ACESS, ');
    SQL.Add('  ''N'' FLG_CANCELADO, ');
    SQL.Add('  NULL AS DES_OBSERVACAO, ');
    SQL.Add('  NF.CHAVEACESSO AS NUM_CHAVE_ACESSO ');
    SQL.Add('FROM ');
    SQL.Add('  NF  ');
    SQL.Add('    LEFT JOIN NATOPER ON NF.CODNATOPER = NATOPER.CODNATOPER ');
    SQL.Add('    INNER JOIN CLIENTES ON NF.CODCLIE = CLIENTES.CODCLIE ');
    SQL.Add('WHERE ');
    SQL.Add('  NF.CODNATOPER IN ( 5411, 5202) ');
    SQL.Add('  AND  ');
    SQL.Add('  CAST(NF.DTEMISSAO AS DATE) >= '''+FormatDateTime('yyyy-mm-dd',DtpInicial.Date)+''' ');
    SQL.Add('  AND  ');
    SQL.Add('  CAST(NF.DTEMISSAO AS DATE) <= '''+FormatDateTime('yyyy-mm-dd',DtpFinal.Date)+''' ');

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

        if QryPrincipal2.FieldByName('DTA_EMISSAO').AsString <> '' then
          Layout.FieldByName('DTA_EMISSAO').AsString:= FormatDateTime('dd/mm/yyyy',QryPrincipal2.FieldByName('DTA_EMISSAO').AsDateTime);

        if QryPrincipal2.FieldByName('DTA_ENTRADA').AsString <> '' then
          Layout.FieldByName('DTA_ENTRADA').AsString:= FormatDateTime('dd/mm/yyyy',QryPrincipal2.FieldByName('DTA_ENTRADA').AsDateTime);

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

procedure TFrmSmCasteloNext.GerarNFitensFornec;
var
   NumLinha, TotalReg, NumItem  :Integer;
   nota, serie, fornecedor, CodNf : string;
   count : integer;
begin
  inherited;

  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT ');
    SQL.Add('  ENTRADANF.CODENTRADANF, ');
    SQL.Add('  ITMENTRADANF.CODITMENTRADANF, ');
    SQL.Add('  ENTRADANF.CODFORNEC AS COD_FORNECEDOR, ');
    SQL.Add('  ENTRADANF.NUMNOTA AS NUM_NF_FORN, ');
    SQL.Add('  ENTRADANF.SERIE AS NUM_SERIE_NF, ');
    SQL.Add('  ITMENTRADANF.CODPROD AS COD_PRODUTO, ');
    SQL.Add('   ');
    SQL.Add('  CASE   ');
    SQL.Add('    WHEN ((ITMENTRADANF.CODTRIB IN (200, 20, 41))    AND (ITMENTRADANF.ICMS_CRED IN (''17'', ''12'', ''7'', ''5.6'', ''17'', ''0'', ''0'', ''0''))      ');
    SQL.Add('	   AND (ITMENTRADANF.REDUCAO IN (67.06, 0.00))) THEN 51  ');
    SQL.Add('	    ');
    SQL.Add('    WHEN ((ITMENTRADANF.CODTRIB IN (200, 20))        AND (ITMENTRADANF.ICMS_CRED IN (''17'', ''7'', ''0''))      ');
    SQL.Add('      AND (ITMENTRADANF.REDUCAO IN (100.00))) THEN 66  ');
    SQL.Add('	   ');
    SQL.Add('    WHEN ((ITMENTRADANF.CODTRIB IN (60, 260))        AND (ITMENTRADANF.ICMS_CRED IN (''0'', ''0'', ''0'')) AND (ITMENTRADANF.REDUCAO IN (100.00, 0.00))) THEN 60   ');
    SQL.Add('    WHEN ((ITMENTRADANF.CODTRIB IN (0, 00, 000))     AND (ITMENTRADANF.ICMS_CRED IN (''17'', ''12''))    AND (ITMENTRADANF.REDUCAO IN (0.00)))         THEN 3    ');
    SQL.Add('    WHEN ((ITMENTRADANF.CODTRIB IN (0, 00, 000, 20)) AND (ITMENTRADANF.ICMS_CRED IN (''17''))          AND (ITMENTRADANF.REDUCAO IN (58.82)))        THEN 52   ');
    SQL.Add('    WHEN ((ITMENTRADANF.CODTRIB IN (40,41))          AND (ITMENTRADANF.ICMS_CRED IN (''0'', ''0'', ''0'')) AND (ITMENTRADANF.REDUCAO IN (0.00)))         THEN 1    ');
    SQL.Add('    WHEN ((ITMENTRADANF.CODTRIB IN (20))             AND (ITMENTRADANF.ICMS_CRED IN (''12''))          AND (ITMENTRADANF.REDUCAO IN (41.66)))        THEN 6    ');
    SQL.Add('    WHEN ((ITMENTRADANF.CODTRIB IN (0, 00, 000))     AND (ITMENTRADANF.ICMS_CRED IN (''7''))           AND (ITMENTRADANF.REDUCAO IN (0.00)))         THEN 2    ');
    SQL.Add('    WHEN ((ITMENTRADANF.CODTRIB IN (60))             AND (ITMENTRADANF.ICMS_CRED IN (''0'', ''0'', ''0'')) AND (ITMENTRADANF.REDUCAO IN (58.82)))        THEN 56   ');
    SQL.Add('    WHEN ((ITMENTRADANF.CODTRIB IN (0, 00, 000))     AND (ITMENTRADANF.ICMS_CRED IN (''25''))          AND (ITMENTRADANF.REDUCAO IN (0.00)))         THEN 5    ');
    SQL.Add('     ');
    SQL.Add('  ELSE 1 END AS COD_TRIBUTACAO,  ');
    SQL.Add('   ');
    SQL.Add('  ITMENTRADANF.QTDEMBAL AS QTD_EMBALAGEM, ');
    SQL.Add('  ITMENTRADANF.QTDNOTA AS QTD_ENTRADA, ');
    SQL.Add('  UPPER(PRODUTOS.UNIDADE_COMP) AS DES_UNIDADE, ');
    SQL.Add('  ITMENTRADANF.PRECOUNIT AS VAL_TABELA, ');
//    SQL.Add('  ITMENTRADANF.DESCONTO AS VAL_DESCONTO_ITEM, ');

    SQL.Add('  (ITMENTRADANF.DESCONTO / ITMENTRADANF.QTDNOTA) AS VAL_DESCONTO_ITEM, ');

    SQL.Add('  0 AS VAL_ACRESCIMO_ITEM, ');
//    SQL.Add('  ITMENTRADANF.VALORIPI AS VAL_IPI_ITEM, ');

    SQL.Add('  (ITMENTRADANF.VALORIPI / ITMENTRADANF.QTDNOTA) AS VAL_IPI_ITEM, ');

    SQL.Add('  ITMENTRADANF.VALORSTMVA AS VAL_SUBST_ITEM, ');
    SQL.Add('  ITMENTRADANF.VALORFRETE AS VAL_FRETE_ITEM, ');
    SQL.Add('  ITMENTRADANF.VALORICMS AS VAL_CREDITO_ICMS, ');
    SQL.Add('  ITMENTRADANF.PRECOUNIT AS VAL_VENDA_VAREJO, ');
    SQL.Add('  ITMENTRADANF.TOTAL AS VAL_TABELA_LIQ, ');
    SQL.Add('  NULL AS NUM_CGC, ');
    SQL.Add('  ITMENTRADANF.BASEICMS AS VAL_TOT_BC_ICMS, ');
    SQL.Add('  0 AS VAL_TOT_OUTROS_ICMS, ');
    //SQL.Add('  ITMENTRADANF.CODNATOPER AS CFOP, ');

    SQL.Add('   ');
    SQL.Add('  CASE WHEN  ');
    SQL.Add('    LEN(ITMENTRADANF.CODNATOPER) = 5 THEN SUBSTRING(ITMENTRADANF.CODNATOPER, 1, LEN(ITMENTRADANF.CODNATOPER) - 1)   ');
    SQL.Add('  ELSE ');
    SQL.Add('    ITMENTRADANF.CODNATOPER END AS CFOP, ');
    SQL.Add('   ');

    SQL.Add('  0 AS VAL_TOT_ISENTO, ');
    SQL.Add('  ITMENTRADANF.BASESUBSTTRIB AS VAL_TOT_BC_ST, ');
    SQL.Add('  ITMENTRADANF.VALORSTMVA AS VAL_TOT_ST, ');
    SQL.Add('  1 AS NUM_ITEM, ');
    SQL.Add('  0 AS TIPO_IPI, ');
    SQL.Add('  PRODUTOS.CODNCM AS NUM_NCM,     ');
    SQL.Add('  NULL AS DES_REFERENCIA,  ');
    SQL.Add('  0 AS VAL_DESP_ACESS_ITEM  ');
    SQL.Add('FROM ');
    SQL.Add('  ITMENTRADANF ');
    SQL.Add('    INNER JOIN ENTRADANF ON ITMENTRADANF.CODENTRADANF = ENTRADANF.CODENTRADANF ');
    SQL.Add('    LEFT JOIN PRODUTOS ON ITMENTRADANF.CODPROD = PRODUTOS.CODPROD ');
    SQL.Add('WHERE ');
    SQL.Add('  CAST(ENTRADANF.DTEMISSAO AS DATE) >= '''+FormatDateTime('yyyy-mm-dd',DtpInicial.Date)+''' ');
    SQL.Add('  AND  ');
    SQL.Add('  CAST(ENTRADANF.DTEMISSAO AS DATE) <= '''+FormatDateTime('yyyy-mm-dd',DtpFinal.Date)+''' ');
    SQL.Add('   ');
    SQL.Add('UNION ALL ');
    SQL.Add(' ');
    SQL.Add('SELECT ');
    SQL.Add('  NF.CODNF AS CODENTRADANF,   ');
    SQL.Add('  ITMNF.CODITMNF AS CODITMENTRADANF, ');
    SQL.Add('  NULL AS COD_FORNECEDOR, ');
    SQL.Add('  NF.NUMNF AS NUM_NF_FORN, ');
    SQL.Add('  NF.SERIE AS NUM_SERIE_NF, ');
    SQL.Add('  ITMNF.CODPROD AS COD_PRODUTO, ');
    SQL.Add('   ');
    SQL.Add('  CASE   ');
    SQL.Add('    WHEN ((ITMNF.CODTRIB IN (200, 20, 41))    AND (ITMNF.ALIQ_ICMS IN (''17'', ''12'', ''7'', ''5.6'', ''17'', ''0'', ''0'', ''0''))      ');
    SQL.Add('	   AND (ITMNF.PER_REDUC IN (67.06, 0.00))) THEN 51  ');
    SQL.Add('	    ');
    SQL.Add('    WHEN ((ITMNF.CODTRIB IN (200, 20))        AND (ITMNF.ALIQ_ICMS IN (''17'', ''7'', ''0''))      ');
    SQL.Add('      AND (ITMNF.PER_REDUC IN (100.00))) THEN 66  ');
    SQL.Add('	   ');
    SQL.Add('    WHEN ((ITMNF.CODTRIB IN (60, 260))        AND (ITMNF.ALIQ_ICMS IN (''0'', ''0'', ''0'')) AND (ITMNF.PER_REDUC IN (100.00, 0.00))) THEN 60   ');
    SQL.Add('    WHEN ((ITMNF.CODTRIB IN (0, 00, 000))     AND (ITMNF.ALIQ_ICMS IN (''17'', ''12''))      AND (ITMNF.PER_REDUC IN (0.00)))         THEN 3    ');
    SQL.Add('    WHEN ((ITMNF.CODTRIB IN (0, 00, 000, 20)) AND (ITMNF.ALIQ_ICMS IN (''17''))              AND (ITMNF.PER_REDUC IN (58.82)))        THEN 52   ');
    SQL.Add('    WHEN ((ITMNF.CODTRIB IN (40,41))          AND (ITMNF.ALIQ_ICMS IN (''0'', ''0'', ''0'')) AND (ITMNF.PER_REDUC IN (0.00)))         THEN 1    ');
    SQL.Add('    WHEN ((ITMNF.CODTRIB IN (20))             AND (ITMNF.ALIQ_ICMS IN (''12''))              AND (ITMNF.PER_REDUC IN (41.66)))        THEN 6    ');
    SQL.Add('    WHEN ((ITMNF.CODTRIB IN (0, 00, 000))     AND (ITMNF.ALIQ_ICMS IN (''7''))               AND (ITMNF.PER_REDUC IN (0.00)))         THEN 2    ');
    SQL.Add('    WHEN ((ITMNF.CODTRIB IN (60))             AND (ITMNF.ALIQ_ICMS IN (''0'', ''0'', ''0'')) AND (ITMNF.PER_REDUC IN (58.82)))        THEN 56   ');
    SQL.Add('    WHEN ((ITMNF.CODTRIB IN (0, 00, 000))     AND (ITMNF.ALIQ_ICMS IN (''25''))              AND (ITMNF.PER_REDUC IN (0.00)))         THEN 5    ');
    SQL.Add('     ');
    SQL.Add('  ELSE 1 END AS COD_TRIBUTACAO, ');
    SQL.Add('   ');
    SQL.Add('  1 AS QTD_EMBALAGEM, ');
    SQL.Add('  ITMNF.QUANTIDADE AS QTD_ENTRADA, ');
    SQL.Add('  UPPER(ITMNF.UNIDADE) AS DES_UNIDADE, ');
    SQL.Add('  ITMNF.PRECO_UNIT AS VAL_TABELA, ');
    SQL.Add('  ITMNF.VALOR_DESC AS VAL_DESCONTO_ITEM, ');
    SQL.Add('  ITMNF.VALOR_ACRESCIMO AS VAL_ACRESCIMO_ITEM, ');
    SQL.Add('  ITMNF.VALOR_IPI AS VAL_IPI_ITEM, ');
    SQL.Add('  NULL AS VAL_SUBST_ITEM, ');
    SQL.Add('  NULL AS VAL_FRETE, ');
    SQL.Add('  ITMNF.VALOR_ICMS AS VAL_CREDITO_ICMS, ');
    SQL.Add('  NULL AS VAL_VENDA_VAREJO, ');
    SQL.Add('  ITMNF.VALOR_MERC AS VAL_TABELA_LIQ, ');
    SQL.Add('  CLIENTES.CNPJ_CPF AS NUM_CGC, ');
    SQL.Add('  ITMNF.BASE_ICMS AS VAL_TOTAL_BC_ICMS, ');
    SQL.Add('  0 AS VAL_TOT_OUTROS_ICMS, ');
    //SQL.Add('  ITMNF.CODNATOPER AS CFOP, ');

    SQL.Add('   ');
    SQL.Add('  CASE WHEN  ');
    SQL.Add('    LEN(ITMNF.CODNATOPER) = 5 THEN SUBSTRING(ITMNF.CODNATOPER, 1, LEN(ITMNF.CODNATOPER) - 1)   ');
    SQL.Add('  ELSE ');
    SQL.Add('    ITMNF.CODNATOPER END AS CFOP, ');
    SQL.Add('   ');

    SQL.Add('  ITMNF.ISENTOS AS VAL_TOT_ISENTO, ');
    SQL.Add('  ITMNF.BASE_SUB  AS VAL_TOT_BC_ST, ');
    SQL.Add('  ITMNF.VALOR_SUB AS VAL_TOT_ST, ');
    SQL.Add('  1 AS NUM_ITEM, ');
    SQL.Add('  0 AS TIPO_IPI, ');
    SQL.Add('  PRODUTOS.CODNCM AS NUM_NCM, ');
    SQL.Add('  NULL AS DES_REFERENCIA,  ');
    SQL.Add('  0 AS VAL_DESP_ACESS_ITEM  ');
    SQL.Add('FROM ');
    SQL.Add('  ITMNF ');
    SQL.Add('    LEFT JOIN PRODUTOS ON ITMNF.CODPROD = PRODUTOS.CODPROD ');
    SQL.Add('    INNER JOIN NF ON NF.CODNF = ITMNF.CODNF ');
    SQL.Add('    LEFT JOIN CLIENTES ON NF.CODCLIE = CLIENTES.CODCLIE ');
    SQL.Add('WHERE ');
    SQL.Add('  ITMNF.CODNATOPER IN (5411, 5202) ');
    SQL.Add('  AND  ');
    SQL.Add('  CAST(NF.DTEMISSAO AS DATE) >= '''+FormatDateTime('yyyy-mm-dd',DtpInicial.Date)+''' ');
    SQL.Add('  AND  ');
    SQL.Add('  CAST(NF.DTEMISSAO AS DATE) <= '''+FormatDateTime('yyyy-mm-dd',DtpFinal.Date)+''' ');
    SQL.Add('ORDER BY ');
    SQL.Add('1, 2 ');

    Open;

    First;
    NumLinha := 0;
    NumItem := 0;

    while not Eof do
    begin
      try
        if Cancelar then
          Break;
        Inc(NumLinha);

        Layout.SetValues(QryPrincipal2, NumLinha, RecordCount);


        if CodNf <> QryPrincipal2.FieldByName('CODENTRADANF').AsString then
        begin
          NumItem := 0;
          CodNf   := QryPrincipal2.FieldByName('CODENTRADANF').AsString;
        end;

        Inc(NumItem);

        Layout.FieldByName('NUM_ITEM').AsInteger := NumItem;

        Layout.FieldByName('COD_PRODUTO').AsString := GerarPLU(QryPrincipal2.FieldByName('COD_PRODUTO').AsString);
        Layout.FieldByName('NUM_NCM').AsString     := StrRetNums(Layout.FieldByName('NUM_NCM').AsString);

        Layout.WriteLine;
      except
        On E: Exception do
        FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
      end;
      Next;
    end;
  end;
end;

procedure TFrmSmCasteloNext.GerarProdForn;
begin
  inherited;

  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT ');
    SQL.Add('  PRODREF.CODPROD AS COD_PRODUTO, ');
    SQL.Add('  PRODREF.CODFORNEC AS COD_FORNECEDOR, ');
    SQL.Add('  PRODREF.CODREF AS DES_REFERENCIA, ');
    SQL.Add('  FORNECEDORES.CNPJ_CPF AS NUM_CGC, ');
    SQL.Add('  NULL AS COD_DIVISAO, ');
    SQL.Add('  ''UN'' AS DES_UNIDADE_COMPRA, ');
    SQL.Add('  1 AS QTD_EMBALAGEM_COMPRA, ');
    SQL.Add('  1 AS QTD_TROCA, ');
    SQL.Add('  ''N'' AS FLG_PREFERENCIAL ');
    SQL.Add('FROM  ');
    SQL.Add('  PRODREF ');
    SQL.Add('    LEFT JOIN FORNECEDORES ON(FORNECEDORES.CODFORNEC = PRODREF.CODFORNEC) ');
    SQL.Add('    LEFT JOIN PRODUTOS ON(PRODUTOS.CODPROD = PRODREF.CODPROD) ');

    SQL.Add('WHERE  ');
    SQL.Add('  PRODUTOS.IMPORTAR_PRODUTO = ''S'' ');

    SQL.Add('ORDER BY 1 ');

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
        Layout.FieldByName('NUM_CGC').AsString     := StrRetNums(Layout.FieldByName('NUM_CGC').AsString);

        Layout.WriteLine;
      except
        On E: Exception do
        FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
      end;
      Next;
    end;
  end;
end;

procedure TFrmSmCasteloNext.GerarProdLoja;
begin
  inherited;
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT   ');
    SQL.Add('  PRODUTOS.CODPROD AS COD_PRODUTO,  ');
    SQL.Add('  PROD_LOJA.PRECO_CUST AS VAL_CUSTO_REP,  ');
    SQL.Add('  PROD_LOJA.PRECO_UNIT AS VAL_VENDA,  ');
    SQL.Add('  COALESCE(PROMOCAO.PRECO_UNIT, 0) AS VAL_OFERTA,  ');
    SQL.Add('  PROD_LOJA.ESTOQUE AS QTD_EST_VDA,  ');
    SQL.Add('  '''' AS TECLA_BALANCA,  ');
    SQL.Add('    ');
    SQL.Add('  CASE     ');
    SQL.Add('    WHEN ((CODTRIB IN (200, 20, 41, 0))     ');
    SQL.Add('	     AND (CODALIQ_NF IN (''TC'', ''TB'', ''TA'', ''TG'', ''TC'', ''F'', ''I'', ''N''))    ');
    SQL.Add('	     AND (CODALIQ IN (''TC'', ''TB'', ''TA'', ''TG'', ''TC'', ''F'', ''I'', ''N''))    ');
    SQL.Add('	     AND (PER_REDUC IN (67.06, 0.00))) THEN 51 ');
    SQL.Add('		  ');
    SQL.Add('    WHEN ((CODTRIB IN (200, 20))          AND (CODALIQ_NF IN (''TC'', ''F'', ''I'', ''N'')) AND  ');
    SQL.Add('	  (CODALIQ IN (''TA'', ''TC'', ''TF'', ''F'', ''I'', ''N'')) AND (PER_REDUC IN (100.00))) THEN 66   ');
    SQL.Add('	 ');
    SQL.Add('	WHEN ((CODTRIB IN (60, 260))          AND (CODALIQ_NF IN (''F'', ''I'', ''N'')) AND (PER_REDUC IN (100.00, 0.00))) THEN 60     ');
    SQL.Add('    WHEN ((CODTRIB IN (0, 00, 000))       AND (CODALIQ_NF IN (''TC'', ''TB'')) AND (CODALIQ IN (''TB'')) AND (PER_REDUC IN (0.00))) THEN 3      ');
    SQL.Add('    WHEN ((CODTRIB IN (0, 00, 000, 20))   AND (CODALIQ_NF IN (''TC'')) AND (CODALIQ IN (''TA'')) AND (PER_REDUC IN (58.82))) THEN 52     ');
    SQL.Add('    WHEN ((CODTRIB IN (40, 41))           AND (CODALIQ_NF IN (''F'', ''I'', ''N'')) AND (PER_REDUC IN (0.00)))  THEN 1      ');
    SQL.Add('    WHEN ((CODTRIB IN (20))               AND (CODALIQ_NF IN (''TB'')) AND (CODALIQ IN (''TA'')) AND (PER_REDUC IN (41.66))) THEN 6      ');
    SQL.Add('    WHEN ((CODTRIB IN (0, 00, 000))       AND (CODALIQ_NF IN (''TA''))  AND (PER_REDUC IN (0.00))) THEN 2      ');
    SQL.Add('    WHEN ((CODTRIB IN (60))               AND (CODALIQ_NF IN (''F'', ''I'', ''N'')) AND (PER_REDUC IN (58.82))) THEN 56     ');
    SQL.Add('    WHEN ((CODTRIB IN (0, 00, 000))       AND (CODALIQ_NF IN (''TD'')) AND (PER_REDUC IN (0.00)))  THEN 5      ');
    SQL.Add('      ');
    SQL.Add('  ELSE 1 END AS COD_TRIBUTACAO, ');
    SQL.Add('  ');
    SQL.Add('  PROD_LOJA.MARGEM_PARAM AS VAL_MARGEM,  ');
    SQL.Add('  PRODUTOS.ETIQ_PROD AS QTD_ETIQUETA,  ');
    SQL.Add('    ');
    SQL.Add('  CASE     ');
    SQL.Add('    WHEN ((CODTRIB IN (200, 20, 41, 0))     ');
    SQL.Add('	     AND (CODALIQ_NF IN (''TC'', ''TB'', ''TA'', ''TG'', ''TC'', ''F'', ''I'', ''N''))    ');
    SQL.Add('	     AND (CODALIQ IN (''TC'', ''TB'', ''TA'', ''TG'', ''TC'', ''F'', ''I'', ''N''))    ');
    SQL.Add('	     AND (PER_REDUC IN (67.06, 0.00))) THEN 51 ');
    SQL.Add('		  ');
    SQL.Add('    WHEN ((CODTRIB IN (200, 20))          AND (CODALIQ_NF IN (''TC'', ''F'', ''I'', ''N'')) AND  ');
    SQL.Add('	  (CODALIQ IN (''TA'', ''TC'', ''TF'', ''F'', ''I'', ''N'')) AND (PER_REDUC IN (100.00))) THEN 66   ');
    SQL.Add('	 ');
    SQL.Add('	WHEN ((CODTRIB IN (60, 260))          AND (CODALIQ_NF IN (''F'', ''I'', ''N'')) AND (PER_REDUC IN (100.00, 0.00))) THEN 60     ');
    SQL.Add('    WHEN ((CODTRIB IN (0, 00, 000))       AND (CODALIQ_NF IN (''TC'', ''TB'')) AND (CODALIQ IN (''TB'')) AND (PER_REDUC IN (0.00))) THEN 3      ');
    SQL.Add('    WHEN ((CODTRIB IN (0, 00, 000, 20))   AND (CODALIQ_NF IN (''TC'')) AND (CODALIQ IN (''TA'')) AND (PER_REDUC IN (58.82))) THEN 52     ');
    SQL.Add('    WHEN ((CODTRIB IN (40, 41))           AND (CODALIQ_NF IN (''F'', ''I'', ''N'')) AND (PER_REDUC IN (0.00)))  THEN 1      ');
    SQL.Add('    WHEN ((CODTRIB IN (20))               AND (CODALIQ_NF IN (''TB'')) AND (CODALIQ IN (''TA'')) AND (PER_REDUC IN (41.66))) THEN 6      ');
    SQL.Add('    WHEN ((CODTRIB IN (0, 00, 000))       AND (CODALIQ_NF IN (''TA''))  AND (PER_REDUC IN (0.00))) THEN 2      ');
    SQL.Add('    WHEN ((CODTRIB IN (60))               AND (CODALIQ_NF IN (''F'', ''I'', ''N'')) AND (PER_REDUC IN (58.82))) THEN 56     ');
    SQL.Add('    WHEN ((CODTRIB IN (0, 00, 000))       AND (CODALIQ_NF IN (''TD'')) AND (PER_REDUC IN (0.00)))  THEN 5      ');
    SQL.Add('      ');
    SQL.Add('  ELSE 1 END AS COD_TRIB_ENTRADA, ');
    SQL.Add('    ');
    SQL.Add('  CASE ATIVO  ');
    SQL.Add('    WHEN ''S'' THEN ''N''  ');
    SQL.Add('  ELSE ''S''  ');
    SQL.Add('  END AS FLG_INATIVO,    ');
    SQL.Add('    ');
    SQL.Add('  PRODUTOS.CODPROD AS COD_PRODUTO_ANT,  ');
    SQL.Add('  PRODUTOS.CODNCM AS NUM_NCM,  ');
    SQL.Add('  1 AS TIPO_NCM,    ');
    SQL.Add('  COALESCE(EMBALAGENS.PRECO_UNIT, 0) AS VAL_VENDA_2,  ');
    SQL.Add('  DATAFIM AS DTA_VALIDA_OFERTA,   ');
    SQL.Add('  PRODUTOS.ESTOQUE_MIN AS QTD_EST_MINIMO,  ');
    SQL.Add('  NULL AS COD_VASILHAME,  ');
    SQL.Add('  0 AS FORA_LINHA,  ');
    SQL.Add('  0 AS QTD_PRECO_DIF,  ');
    SQL.Add('  0 AS VAL_FORCA_VDA,    ');
    SQL.Add('  CASE WHEN COALESCE(PRODUTOS.CODCEST, '''') = '''' OR PRODUTOS.CODCEST IS NULL THEN ''99.999.99'' ELSE PRODUTOS.CODCEST END AS NUM_CEST,   ');
    SQL.Add('  0 AS PER_IVA,      ');
    SQL.Add('  0 AS PER_FCP_ST,      ');
    SQL.Add('  0 AS PER_FIDELIDADE,      ');
    SQL.Add('  0 AS COD_INFO_RECEITA       ');
    SQL.Add('FROM   ');
    SQL.Add('  PRODUTOS  ');
    SQL.Add('    INNER JOIN PROD_LOJA ON PROD_LOJA.CODPROD = PRODUTOS.CODPROD  ');
    SQL.Add('    LEFT JOIN EMBALAGENS ON PRODUTOS.CODPROD  = EMBALAGENS.CODPROD  ');
    SQL.Add('    LEFT JOIN PROMOCAO ON PROMOCAO.CODPROD    = PRODUTOS.CODPROD  ');
    SQL.Add(' ');
    SQL.Add('WHERE  ');
    SQL.Add('  PRODUTOS.IMPORTAR_PRODUTO = ''S'' ');
    SQL.Add(' ');
    SQL.Add('UNION ALL ');
    SQL.Add(' ');
    SQL.Add(' ');
    SQL.Add('SELECT   ');
//    SQL.Add('  PROD.CODPROD AS COD_PRODUTO,  ');
    SQL.Add(' ');
    SQL.Add('  CASE WHEN ');
    SQL.Add('	   (SELECT COUNT(*) FROM PRODUTOS WHERE CODPROD = PROD.CODPROD AND DESCRICAO <> PROD.DESCRICAO) > 0 THEN ');
    SQL.Add('	     PRODUTOS.ID + ROW_NUMBER() OVER (ORDER BY PROD.CODPROD) ');
    SQL.Add('	 ELSE ');
    SQL.Add('	    PROD.CODPROD END AS COD_PRODUTO, ');
    SQL.Add(' ');
    SQL.Add('  0.00 AS VAL_CUSTO_REP,  ');
    SQL.Add('  0.00 AS VAL_VENDA,  ');
    SQL.Add('  0.00 AS VAL_OFERTA,  ');
    SQL.Add('  0.00 AS QTD_EST_VDA,  ');
    SQL.Add('  '''' AS TECLA_BALANCA,  ');
    SQL.Add('    ');
    SQL.Add('  CASE     ');
    SQL.Add('    WHEN ((CODTRIB IN (200, 20, 41, 0))     ');
    SQL.Add('	     AND (CODALIQ_NF IN (''TC'', ''TB'', ''TA'', ''TG'', ''TC'', ''F'', ''I'', ''N''))    ');
    SQL.Add('	     AND (CODALIQ IN (''TC'', ''TB'', ''TA'', ''TG'', ''TC'', ''F'', ''I'', ''N''))    ');
    SQL.Add('	     AND (PER_REDUC IN (67.06, 0.00))) THEN 51 ');
    SQL.Add('		  ');
    SQL.Add('    WHEN ((CODTRIB IN (200, 20))          AND (CODALIQ_NF IN (''TC'', ''F'', ''I'', ''N'')) AND  ');
    SQL.Add('	  (CODALIQ IN (''TA'', ''TC'', ''TF'', ''F'', ''I'', ''N'')) AND (PER_REDUC IN (100.00))) THEN 66   ');
    SQL.Add('	 ');
    SQL.Add('	WHEN ((CODTRIB IN (60, 260))          AND (CODALIQ_NF IN (''F'', ''I'', ''N'')) AND (PER_REDUC IN (100.00, 0.00))) THEN 60     ');
    SQL.Add('    WHEN ((CODTRIB IN (0, 00, 000))       AND (CODALIQ_NF IN (''TC'', ''TB'')) AND (CODALIQ IN (''TB'')) AND (PER_REDUC IN (0.00))) THEN 3      ');
    SQL.Add('    WHEN ((CODTRIB IN (0, 00, 000, 20))   AND (CODALIQ_NF IN (''TC'')) AND (CODALIQ IN (''TA'')) AND (PER_REDUC IN (58.82))) THEN 52     ');
    SQL.Add('    WHEN ((CODTRIB IN (40, 41))           AND (CODALIQ_NF IN (''F'', ''I'', ''N'')) AND (PER_REDUC IN (0.00)))  THEN 1      ');
    SQL.Add('    WHEN ((CODTRIB IN (20))               AND (CODALIQ_NF IN (''TB'')) AND (CODALIQ IN (''TA'')) AND (PER_REDUC IN (41.66))) THEN 6      ');
    SQL.Add('    WHEN ((CODTRIB IN (0, 00, 000))       AND (CODALIQ_NF IN (''TA''))  AND (PER_REDUC IN (0.00))) THEN 2      ');
    SQL.Add('    WHEN ((CODTRIB IN (60))               AND (CODALIQ_NF IN (''F'', ''I'', ''N'')) AND (PER_REDUC IN (58.82))) THEN 56     ');
    SQL.Add('    WHEN ((CODTRIB IN (0, 00, 000))       AND (CODALIQ_NF IN (''TD'')) AND (PER_REDUC IN (0.00)))  THEN 5      ');
    SQL.Add('      ');
    SQL.Add('  ELSE 1 END AS COD_TRIBUTACAO, ');
    SQL.Add('  ');
    SQL.Add('  0.00 AS VAL_MARGEM,  ');
    SQL.Add('  PROD.ETIQ_PROD AS QTD_ETIQUETA,  ');
    SQL.Add('    ');
    SQL.Add('  CASE     ');
    SQL.Add('    WHEN ((CODTRIB IN (200, 20, 41, 0))     ');
    SQL.Add('	     AND (CODALIQ_NF IN (''TC'', ''TB'', ''TA'', ''TG'', ''TC'', ''F'', ''I'', ''N''))    ');
    SQL.Add('	     AND (CODALIQ IN (''TC'', ''TB'', ''TA'', ''TG'', ''TC'', ''F'', ''I'', ''N''))    ');
    SQL.Add('	     AND (PER_REDUC IN (67.06, 0.00))) THEN 51 ');
    SQL.Add('		  ');
    SQL.Add('    WHEN ((CODTRIB IN (200, 20))          AND (CODALIQ_NF IN (''TC'', ''F'', ''I'', ''N'')) AND  ');
    SQL.Add('	  (CODALIQ IN (''TA'', ''TC'', ''TF'', ''F'', ''I'', ''N'')) AND (PER_REDUC IN (100.00))) THEN 66   ');
    SQL.Add('	 ');
    SQL.Add('	WHEN ((CODTRIB IN (60, 260))          AND (CODALIQ_NF IN (''F'', ''I'', ''N'')) AND (PER_REDUC IN (100.00, 0.00))) THEN 60     ');
    SQL.Add('    WHEN ((CODTRIB IN (0, 00, 000))       AND (CODALIQ_NF IN (''TC'', ''TB'')) AND (CODALIQ IN (''TB'')) AND (PER_REDUC IN (0.00))) THEN 3      ');
    SQL.Add('    WHEN ((CODTRIB IN (0, 00, 000, 20))   AND (CODALIQ_NF IN (''TC'')) AND (CODALIQ IN (''TA'')) AND (PER_REDUC IN (58.82))) THEN 52     ');
    SQL.Add('    WHEN ((CODTRIB IN (40, 41))           AND (CODALIQ_NF IN (''F'', ''I'', ''N'')) AND (PER_REDUC IN (0.00)))  THEN 1      ');
    SQL.Add('    WHEN ((CODTRIB IN (20))               AND (CODALIQ_NF IN (''TB'')) AND (CODALIQ IN (''TA'')) AND (PER_REDUC IN (41.66))) THEN 6      ');
    SQL.Add('    WHEN ((CODTRIB IN (0, 00, 000))       AND (CODALIQ_NF IN (''TA''))  AND (PER_REDUC IN (0.00))) THEN 2      ');
    SQL.Add('    WHEN ((CODTRIB IN (60))               AND (CODALIQ_NF IN (''F'', ''I'', ''N'')) AND (PER_REDUC IN (58.82))) THEN 56     ');
    SQL.Add('    WHEN ((CODTRIB IN (0, 00, 000))       AND (CODALIQ_NF IN (''TD'')) AND (PER_REDUC IN (0.00)))  THEN 5      ');
    SQL.Add('      ');
    SQL.Add('  ELSE 1 END AS COD_TRIB_ENTRADA, ');
    SQL.Add('    ');
    SQL.Add('  CASE ATIVO  ');
    SQL.Add('    WHEN ''S'' THEN ''N''  ');
    SQL.Add('  ELSE ''S''  ');
    SQL.Add('  END AS FLG_INATIVO,    ');
    SQL.Add('    ');
    SQL.Add('  PROD.CODPROD AS COD_PRODUTO_ANT,  ');
    SQL.Add('  PROD.CODNCM AS NUM_NCM,  ');
    SQL.Add('  1 AS TIPO_NCM,    ');
    SQL.Add('  COALESCE(EMBALAGENS.PRECO_UNIT, 0) AS VAL_VENDA_2,  ');
    SQL.Add('  '''' AS DTA_VALIDA_OFERTA,   ');
    SQL.Add('  PROD.ESTOQUE_MIN AS QTD_EST_MINIMO,  ');
    SQL.Add('  NULL AS COD_VASILHAME,  ');
    SQL.Add('  0 AS FORA_LINHA,  ');
    SQL.Add('  0 AS QTD_PRECO_DIF,  ');
    SQL.Add('  0 AS VAL_FORCA_VDA,    ');
    SQL.Add('  CASE WHEN COALESCE(PROD.CODCEST, '''') = '''' OR PROD.CODCEST IS NULL THEN ''99.999.99'' ELSE PROD.CODCEST END AS NUM_CEST,    ');
    SQL.Add('  0 AS PER_IVA,      ');
    SQL.Add('  0 AS PER_FCP_ST,      ');
    SQL.Add('  0 AS PER_FIDELIDADE,      ');
    SQL.Add('  0 AS COD_INFO_RECEITA       ');
    SQL.Add('FROM   ');
    SQL.Add('  UNIFICACAO_PRODUTOS PROD  ');
    SQL.Add('    CROSS JOIN (SELECT MAX(CODPROD) AS ID FROM PRODUTOS) AS PRODUTOS ');
    SQL.Add('    INNER JOIN PROD_LOJA ON PROD_LOJA.CODPROD = PROD.CODPROD  ');
    SQL.Add('    LEFT JOIN EMBALAGENS ON PROD.CODPROD = EMBALAGENS.CODPROD  ');
    SQL.Add(' ');
    SQL.Add('WHERE  ');
    SQL.Add('  LEN(SUBSTRING(BARRA, PATINDEX(''%[^0]%'', BARRA+''.''), LEN(BARRA))) >= 8  ');
    SQL.Add('  AND  ');
    SQL.Add('  PROD.IMPORTAR_PRODUTO = ''S''  ');
    SQL.Add('  AND ');
    SQL.Add('  (SELECT COUNT(*) FROM PRODUTOS WHERE BARRA = PROD.BARRA AND PROD.IMPORTAR_PRODUTO = ''N'') = 0 ');
//    SQL.Add('  AND ');
//    SQL.Add('  (SELECT COUNT(*) FROM PRODUTOS WHERE CODPROD = PROD.CODPROD) = 0 ');

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
        Layout.FieldByName('COD_PRODUTO').AsString := Layout.FieldByName('COD_PRODUTO').AsString;
        Layout.FieldByName('NUM_NCM').AsString     := StrRetNums(Layout.FieldByName('NUM_NCM').AsString);
        Layout.FieldByName('NUM_CEST').AsString    := StrRetNums(Layout.FieldByName('NUM_CEST').AsString);

        if QryPrincipal2.FieldByName('DTA_VALIDA_OFERTA').AsString <> '' then
          Layout.FieldByName('DTA_VALIDA_OFERTA').AsString:= FormatDateTime('dd/mm/yyyy', QryPrincipal2.FieldByName('DTA_VALIDA_OFERTA').AsDateTime);


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

procedure TFrmSmCasteloNext.GerarProdSimilar;
var
NumLinha :Integer;
begin
  inherited;
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;
    SQL.Add('SELECT  ');
    SQL.Add('  CODFAMILIA AS COD_PRODUTO_SIMILAR, ');
    SQL.Add('  DESCRICAO AS DES_PRODUTO_SIMILAR, ');
    SQL.Add('  0 AS VAL_META  ');
    SQL.Add('FROM  ');
    SQL.Add('  FAMILIA  ');

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
          FrmProgresso.AdicionarLog(QryPrincipal.RecNo, 'E', E.Message);
      end;
      Next;
    end;
  end;
end;

end.
