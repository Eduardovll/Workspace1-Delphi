unit UFrmSmKuzzi;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, UFrmModelo, Data.DBXOracle, Data.DB,
  Data.SqlExpr, Vcl.Menus, Vcl.StdCtrls, Vcl.Buttons, Vcl.ComCtrls,
  Vcl.Imaging.pngimage, Vcl.ExtCtrls, Data.DBXFirebird, Data.Win.ADODB,
  Vcl.Grids, Vcl.DBGrids, Datasnap.Provider, Datasnap.DBClient, //dxGDIPlusClasses,
  Math;

type
  TFrmSmKuzzi = class(TFrmModeloSis)
    btnGeraCest: TButton;
    BtnAmarrarCest: TButton;
    ADOOracle: TADOConnection;
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
    procedure GerarReceitas; Override;

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
  FrmSmKuzzi: TFrmSmKuzzi;
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


procedure TFrmSmKuzzi.GerarProducao;
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

      if not PLUValido(Layout.FieldByName('COD_PRODUTO').AsString) then
      begin
        Layout.FieldByName('COD_PRODUTO').AsString := GerarPlu(Copy(Layout.FieldByName('COD_PRODUTO').AsString, 1, Length(Layout.FieldByName('COD_PRODUTO').AsString) - 1));
      end;

      if not PLUValido(Layout.FieldByName('COD_PRODUTO_PRODUCAO').AsString) then
      begin
        Layout.FieldByName('COD_PRODUTO_PRODUCAO').AsString := GerarPlu(Copy(Layout.FieldByName('COD_PRODUTO').AsString, 1, Length(Layout.FieldByName('COD_PRODUTO').AsString) - 1));
      end;

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;
end;

procedure TFrmSmKuzzi.GerarProduto;
var
 cod_produto : string;
 count : integer;
 codigos : string;
begin
  inherited;
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
    SQL.Add('     PRODUTOS.ID AS COD_PRODUTO,');
    SQL.Add('     EAN.EAN AS COD_BARRA_PRINCIPAL,');
    SQL.Add('     PRODUTOS.DESCRITIVO_PDV AS DES_REDUZIDA,');
    SQL.Add('     PRODUTOS.DESCRITIVO AS DES_PRODUTO,');
    SQL.Add('     PRODUTOS.QTDE_EMBALAGEME AS QTD_EMBALAGEM_COMPRA,');
    SQL.Add('     PRODUTOS.UNIDADE_COMPRA AS DES_UNIDADE_COMPRA,');
    SQL.Add('     PRODUTOS.QTDE_EMBALAGEMS AS QTD_EMBALAGEM_VENDA,');
    SQL.Add('     PRODUTOS.UNIDADE_VENDA AS DES_UNIDADE_VENDA,');
    SQL.Add('');
    SQL.Add('     CASE ');
    SQL.Add('          WHEN PRODUTOS.IPI_TIPO = ''P'' THEN 0');
    SQL.Add('          ELSE 1');
    SQL.Add('     END AS TIPO_IPI,');
    SQL.Add('');
    SQL.Add('     PRODUTOS.IPI AS VAL_IPI,');
    SQL.Add('');
    SQL.Add('     PRODUTOS.DEPTO AS COD_SECAO,');
    SQL.Add('');
    SQL.Add('     CASE ');
    SQL.Add('          WHEN PRODUTOS.SECAO = 0 THEN 999');
    SQL.Add('          ELSE PRODUTOS.SECAO ');
    SQL.Add('     END AS COD_GRUPO,');
    SQL.Add('');
    SQL.Add('     CASE ');
    SQL.Add('          WHEN PRODUTOS.GRUPO = 0 THEN 999');
    SQL.Add('          ELSE PRODUTOS.GRUPO ');
    SQL.Add('     END AS COD_SUB_GRUPO,');
    SQL.Add('');
    SQL.Add('     PRODUTOS.FAMILIA AS COD_PRODUTO_SIMILAR,');
    SQL.Add('');
    SQL.Add('     CASE ');
    SQL.Add('     WHEN PRODUCAO.COD_PRODUTO_PRODUCAO IS NOT NULL THEN ''S''');
    SQL.Add('     ELSE ');
    SQL.Add('          CASE');
    SQL.Add('          WHEN PRODUTOS.IPV = 0 THEN ''S''');
    SQL.Add('          ELSE ''N''');
    SQL.Add('          END            ');
    SQL.Add('     END AS IPV,');
    SQL.Add('');
    SQL.Add('     PRODUTOS.VALIDADE AS DIAS_VALIDADE,');
    SQL.Add('');
    SQL.Add('     0 AS TIPO_PRODUTO,');
    SQL.Add('');
    SQL.Add('     CASE PRODUTOS.MONOFASICO');
    SQL.Add('          WHEN ''I'' THEN ''S''');
    SQL.Add('          WHEN ''M'' THEN ''S''');
    SQL.Add('          WHEN ''T'' THEN ''N''');
    SQL.Add('          WHEN ''B'' THEN ''S''');
    SQL.Add('          WHEN ''O'' THEN ''S''');
    SQL.Add('          WHEN ''N'' THEN ''S''');
    SQL.Add('          WHEN ''S'' THEN ''S''');
    SQL.Add('     END AS FLG_NAO_PIS_COFINS,');
    SQL.Add('');
    SQL.Add('     CASE ');
    SQL.Add('          WHEN PRODUTOS.IPV = 1 THEN ''N''');
    SQL.Add('          ELSE ''S''');
    SQL.Add('     END  AS FLG_ENVIA_BALANCA,');
    SQL.Add('');
    SQL.Add('     CASE PRODUTOS.MONOFASICO');
    SQL.Add('          WHEN ''I'' THEN 0 -- ISENTO');
    SQL.Add('          WHEN ''M'' THEN 1 -- MONOFASICO');
    SQL.Add('          WHEN ''T'' THEN -1 -- INCIDENTE');
    SQL.Add('          WHEN ''B'' THEN 2 -- SUBSTITUICAO');
    SQL.Add('          WHEN ''O'' THEN 0 -- ALIQUOTA ZERO');
    SQL.Add('          WHEN ''N'' THEN 0 -- NÃO INCIDENTE');
    SQL.Add('          WHEN ''S'' THEN 4 -- SUSPENSO');
    SQL.Add('     END AS TIPO_NAO_PIS_COFINS,');
    SQL.Add('');
    SQL.Add('     CASE PRODUTOS.COMPOSTO');
    SQL.Add('          WHEN 1 THEN 2');
    SQL.Add('          WHEN 4 THEN 1');
    SQL.Add('          WHEN 2 THEN 3');
    SQL.Add('          ELSE 0 ');
    SQL.Add('     END AS TIPO_EVENTO,');
    SQL.Add('');
    SQL.Add('     COALESCE(ASSOCIADO.PRODUTO, 0) AS COD_ASSOCIADO,');
    SQL.Add('     '''' AS DES_OBSERVACAO,');
    SQL.Add('     COALESCE(NUTRICIONAL.ID, 0) AS COD_INFO_NUTRICIONAL,');
//    SQL.Add('     COALESCE(PRODUTOS.ID, 0) AS COD_INFO_RECEITA,');
    SQL.Add('');
//    SQL.Add('     CASE PRODUTOS.MONOFASICO');
//    SQL.Add('         WHEN ''T'' THEN ''000''');
//    SQL.Add('         ELSE LPAD(COALESCE(TRIBUTACAO_ENTRADA.COD_TAB_SPED, ''000''), 3, ''0'')');
//    SQL.Add('     END AS COD_TAB_SPED,');
    SQL.Add('');
    SQL.Add('     ''999'' AS COD_TAB_SPED,');
    SQL.Add('');
    SQL.Add('     ''N'' AS FLG_ALCOOLICO,');
    SQL.Add('');
    SQL.Add('     CASE ');
    SQL.Add('          WHEN PRODUCAO.COD_PRODUTO_PRODUCAO IS NOT NULL THEN 4');
    SQL.Add('          ELSE 0');
    SQL.Add('     END AS TIPO_ESPECIE,');
    SQL.Add('');
    SQL.Add('     0 AS COD_CLASSIF,');
    SQL.Add('');
    SQL.Add('     CASE ');
    SQL.Add('          WHEN PRODUTOS.PESOB = 0 THEN 1');
    SQL.Add('          ELSE PRODUTOS.PESOB');
    SQL.Add('     END AS VAL_VDA_PESO_BRUTO,');
    SQL.Add('');
    SQL.Add('     CASE');
    SQL.Add('          WHEN PRODUTOS.PESOL = 0 THEN 1');
    SQL.Add('          ELSE PRODUTOS.PESOL');
    SQL.Add('     END AS VAL_PESO_EMB,');
    SQL.Add('');
    SQL.Add('     CASE PRODUTOS.COMPOSTO');
    SQL.Add('          WHEN 4 THEN 1');
    SQL.Add('          ELSE 0 ');
    SQL.Add('     END AS TIPO_EXPLOSAO_COMPRA,');
    SQL.Add('');
    SQL.Add('     NULL AS DTA_INI_OPER,');
    SQL.Add('     '''' AS DES_PLAQUETA,');
    SQL.Add('     '''' AS MES_ANO_INI_DEPREC,');
    SQL.Add('     0 AS TIPO_BEM,');
    SQL.Add('     ''0'' AS COD_FORNECEDOR,');
    SQL.Add('     0 AS NUM_NF,');
    SQL.Add('     NULL AS DTA_ENTRADA,');
    SQL.Add('     0 AS COD_NAT_BEM,');
    SQL.Add('     0 AS VAL_ORIG_BEM,');
    SQL.Add('     PRODUTOS.DESCRITIVO AS DES_PRODUTO_ANT,  ');
//    SQL.Add('     ''P'' AS TIPO_CADASTRO,');
    SQL.Add('     COALESCE(PRODUTOS.ID_MARCA_PRODUTO, 0) AS COD_MARCA');
    SQL.Add('');
    SQL.Add('FROM PRODUTOS');
    SQL.Add('LEFT JOIN NUTRICIONAL');
    SQL.Add('ON PRODUTOS.ID = NUTRICIONAL.ID   ');
    SQL.Add('LEFT JOIN PRODUTOS_COMPOSICAO ASSOCIADO');
    SQL.Add('ON PRODUTOS.ID = ASSOCIADO.PRODUTO_BASE');
    SQL.Add('AND PRODUTOS.COMPOSTO = 3   ');
    SQL.Add('');
    SQL.Add('LEFT JOIN');
    SQL.Add('     (');
    SQL.Add('     SELECT');
    SQL.Add('     MAX(PRODUTOS_EAN.EAN) AS EAN,');
    SQL.Add('     PRODUTOS_EAN.PRODUTO');
    SQL.Add('     FROM');
    SQL.Add('     PRODUTOS_EAN');
    SQL.Add('     WHERE');
    SQL.Add('     PRODUTOS_EAN.QTDEE = 1               ');
    SQL.Add('     GROUP BY');
    SQL.Add('     PRODUTOS_EAN.PRODUTO');
    SQL.Add('     ) EAN');
    SQL.Add('ON PRODUTOS.ID = EAN.PRODUTO');
    SQL.Add('');
    SQL.Add('LEFT JOIN');
    SQL.Add('     (');
    SQL.Add('     SELECT DISTINCT');
    SQL.Add('     PRODUTOS_COMPOSICAO.PRODUTO AS COD_PRODUTO_PRODUCAO');
    SQL.Add('     FROM PRODUTOS');
    SQL.Add('     LEFT JOIN PRODUTOS_COMPOSICAO');
    SQL.Add('     ON PRODUTOS.ID = PRODUTOS_COMPOSICAO.PRODUTO_BASE ');
    SQL.Add('');
    SQL.Add('     LEFT JOIN');
    SQL.Add('     COMPOSICAO');
    SQL.Add('     ON PRODUTOS_COMPOSICAO.PRODUTO_BASE = COMPOSICAO.PRODUTO_BASE');
    SQL.Add('     WHERE PRODUTOS.COMPOSTO = 2');
    SQL.Add('     AND PRODUTOS_COMPOSICAO.PRODUTO_BASE IS NOT NULL');
    SQL.Add('     ) PRODUCAO');
    SQL.Add('ON PRODUTOS.ID = PRODUCAO.COD_PRODUTO_PRODUCAO');

    Open;
    First;
    NumLinha := 0;

    AssignFile(Arquivo, EdtCamArquivo.Text + '\SCRIPT_ATUALIZA_MARCAS.SQL');
    Rewrite(Arquivo);

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);
      Layout.SetValues(QryPrincipal2, NumLinha, RecordCount);

      if not PLUValido(Layout.FieldByName('COD_PRODUTO').AsString) then
      begin
        Layout.FieldByName('COD_PRODUTO').AsString := GerarPlu(Copy(Layout.FieldByName('COD_PRODUTO').AsString, 1, Length(Layout.FieldByName('COD_PRODUTO').AsString) - 1));
      end;
//      if Layout.FieldByName('COD_INFO_RECEITA').AsInteger <> 0 then
//      begin
//        if not PLUValido(Layout.FieldByName('COD_PRODUTO').AsString) then
//          Layout.FieldByName('COD_INFO_RECEITA').AsString := GerarPlu(Copy(Layout.FieldByName('COD_INFO_RECEITA').AsString, 1, Length(Layout.FieldByName('COD_INFO_RECEITA').AsString) - 1));
//      end;

      if( QryPrincipal2.FieldByName('COD_MARCA').AsInteger <> 0 ) then
         Writeln(Arquivo, 'UPDATE TAB_PRODUTO SET COD_MARCA = '+ QryPrincipal2.FieldByName('COD_MARCA').ASString +' WHERE COD_PRODUTO = '+ Layout.FieldByName('COD_PRODUTO').AsString +'; ');

//        count := count + 1;


//      if QryPrincipal2.FieldByName('TIPO_CADASTRO').AsString = 'E' then
//         Layout.FieldByName('COD_PRODUTO').AsString := GerarPLU( Layout.FieldByName('COD_PRODUTO').AsString );


      Layout.FieldByName('DES_OBSERVACAO').AsString := StrReplace(StrLBReplace(FieldByName('DES_OBSERVACAO').AsString), '\n', '');
      Layout.FieldByName('DES_REDUZIDA').AsString := StrReplace(StrLBReplace(FieldByName('DES_REDUZIDA').AsString), '\n', '');
      Layout.FieldByName('DES_PRODUTO').AsString := StrReplace(StrLBReplace(FieldByName('DES_PRODUTO').AsString), '\n', '');

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
    Writeln(Arquivo, 'COMMIT WORK;');
    Close
  end;
  CloseFile(Arquivo);
end;

procedure TFrmSmKuzzi.GerarReceitas;
var
  texto : string;
begin
  inherited;
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
    SQL.Add('  COALESCE(PRODUTOS.ID, 999) AS COD_INFO_RECEITA,');
    SQL.Add('  PRODUTOS.DESCRITIVO AS DES_INFO_RECEITA,');
    SQL.Add('  replace(translate(receita, chr(10) || chr(13) || chr(09), '' ''), '','', '' '') AS DETALHAMENTO');
    SQL.Add('FROM PRODUTOS');
    SQL.Add('WHERE PRODUTOS.RECEITA IS NOT NULL');
    SQL.Add('ORDER BY PRODUTOS.ID');


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

//      if not PLUValido(Layout.FieldByName('COD_INFO_RECEITA').AsString) then
//      begin
//        Layout.FieldByName('COD_INFO_RECEITA').AsString := GerarPlu(Copy(Layout.FieldByName('COD_INFO_RECEITA').AsString, 1, Length(Layout.FieldByName('COD_INFO_RECEITA').AsString) - 1));
//      end;

//      Layout.FieldByName('DETALHAMENTO').AsString := StrReplace(StrLBReplace( StringReplace(FieldByName('DETALHAMENTO').AsString,#$A, '', [rfReplaceAll]) ), '\n', '') ;      Layout.FieldByName('DETALHAMENTO').AsString := StrReplace(StrLBReplace( StringReplace(FieldByName('DETALHAMENTO').AsString,#$A, '', [rfReplaceAll]) ), '\n', '') ;

//      texto := StringReplace(StringReplace(StringReplace(Layout.FieldByName('DETALHAMENTO').AsString, #$D#$A, '', [rfReplaceAll]), #$A, '', [rfReplaceAll]), '#$A', '', [rfReplaceAll]);
//      Layout.FieldByName('DETALHAMENTO').AsString := texto;
      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;
end;

procedure TFrmSmKuzzi.GerarScriptAmarrarCEST;
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

procedure TFrmSmKuzzi.GerarScriptCEST;
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

procedure TFrmSmKuzzi.GerarSecao;
begin
  inherited;

  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT DISTINCT');
    SQL.Add('    DEPTOS.DEPTO AS COD_SECAO,');
    SQL.Add('    DEPTOS.DESCRITIVO AS DES_SECAO,');
    SQL.Add('    DEPTOS.MARGEM AS VAL_META');
    SQL.Add('FROM DEPTOS');
    SQL.Add('WHERE DEPTOS.SECAO = 0');
    SQL.Add('AND DEPTOS.GRUPO = 0');
    SQL.Add('AND DEPTOS.SUBGRUPO = 0');


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

procedure TFrmSmKuzzi.GerarSubGrupo;
begin
  inherited;

  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT DISTINCT');
    SQL.Add('    DEPTOS.DEPTO AS COD_SECAO,');
    SQL.Add('    CASE');
    SQL.Add('        WHEN DEPTOS.SECAO = 0 THEN 999');
    SQL.Add('        ELSE DEPTOS.SECAO');
    SQL.Add('    END AS COD_GRUPO,');
    SQL.Add('');
    SQL.Add('    CASE ');
    SQL.Add('        WHEN DEPTOS.GRUPO = 0 THEN 999');
    SQL.Add('        ELSE DEPTOS.GRUPO');
    SQL.Add('    END AS COD_SUB_GRUPO,');
    SQL.Add('');
    SQL.Add('    CASE ');
    SQL.Add('        WHEN DEPTOS.GRUPO = 0 THEN ''A DEFINIR''');
    SQL.Add('        ELSE DEPTOS.DESCRITIVO ');
    SQL.Add('    END AS DES_SUB_GRUPO,');
    SQL.Add('');
    SQL.Add('    DEPTOS.MARGEM AS VAL_META,');
    SQL.Add('    0 AS VAL_MARGEM_REF,');
    SQL.Add('    0 AS QTD_DIA_SEGURANCA,');
    SQL.Add('    ''N'' AS FLG_ALCOOLICO');
    SQL.Add('FROM DEPTOS ');


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

procedure TFrmSmKuzzi.GerarTransportadora;
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

procedure TFrmSmKuzzi.GerarVenda;
var
    QryCusto : TADOQuery;
    custo : currency;
begin
  inherited;

  QryCusto := TADOQuery.Create(FrmProgresso);

  with QryCusto do
  begin
    Connection := ADOOracle;
    SQL.Clear;
    SQL.Add('SELECT ');
    SQL.Add('  CUSTO_ATUAL AS CUSTO_ATUAL ');
    SQL.Add('FROM');
    SQL.Add('  (');
    SQL.Add('    SELECT');
    SQL.Add('       CUSTO_ATUAL AS CUSTO_ATUAL');
    SQL.Add('    FROM');
    SQL.Add('      ALTERACAO_CUSTOS ');
    SQL.Add('    WHERE ');
    SQL.Add('      PRODUTO = :COD_PRODUTO ');
    SQL.Add('    AND ');
    SQL.Add('      POLITICA = :COD_LOJA ');
    SQL.Add('    AND ');
    SQL.Add('      TO_DATE(:DTA_SAIDA, ''DD/MM/YYYY hh24:mi:ss'') BETWEEN DATA_HORA AND DATA_HORA_FIM ');
//    SQL.Add('    AND ');
//    SQL.Add('      TIPO IN (1, 2, 3, 4, 5) ');
//    SQL.Add('    AND ');
//    SQL.Add('      DATA_HORA_FIM <  SYSDATE ');
    SQL.Add('    ORDER BY ');
    SQL.Add('      DATA_HORA');
    SQL.Add('  )');
    SQL.Add('WHERE');
    SQL.Add('  ROWNUM = 1  ');

  end;

  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
    SQL.Add('    ITENS_VENDA.PRODUTO AS COD_PRODUTO,');
    SQL.Add('    VENDAS.EMPRESA AS COD_LOJA,');
    SQL.Add('    0 AS IND_TIPO, --');
    SQL.Add('    VENDAS.PDV AS NUM_PDV,');
    SQL.Add('    ITENS_VENDA.QTDE AS QTD_TOTAL_PRODUTO,');
    SQL.Add('    ITENS_VENDA.VALOR - ITENS_VENDA.DESCONTO + ITENS_VENDA.ACRESCIMO AS VAL_TOTAL_PRODUTO,');
    SQL.Add('    ITENS_VENDA.VENDATAB AS VAL_PRECO_VENDA,');
    SQL.Add('    ITENS_VENDA.VALOR_MARGEM_LUCRO,');
    SQL.Add('    CASE WHEN COALESCE(ITENS_VENDA.CUSTOICMS, 0) <> 0 THEN ITENS_VENDA.CUSTOICMS / ITENS_VENDA.QTDE ELSE 0 END AS VAL_CUSTO_REP,');
    SQL.Add('    VENDAS.DATA_HORA AS DTA_SAIDA,');
    SQL.Add('');
    SQL.Add('    CASE');
    SQL.Add('        WHEN EXTRACT(MONTH FROM VENDAS.DATA_HORA) < 10 THEN ''0'' || EXTRACT(MONTH FROM VENDAS.DATA_HORA) || EXTRACT(YEAR FROM VENDAS.DATA_HORA)');
    SQL.Add('        ELSE EXTRACT(MONTH FROM VENDAS.DATA_HORA) || EXTRACT(YEAR FROM VENDAS.DATA_HORA) ');
    SQL.Add('    END AS DTA_MENSAL,');
    SQL.Add('    ');
    SQL.Add('    VENDAS.DOCUMENTO AS NUM_IDENT,');
    SQL.Add('    '''' AS COD_EAN, ');
    SQL.Add('');
    SQL.Add('    TO_CHAR(VENDAS.DATA_HORA,''HH24'') || TO_CHAR(VENDAS.DATA_HORA,''MI'') AS DES_HORA,');
    SQL.Add('    COALESCE(CLIENTES.ID, 99999) AS COD_CLIENTE, --');
    SQL.Add('    1 AS COD_ENTIDADE, --');
    SQL.Add('    ITENS_VENDA.VALOR AS VAL_BASE_ICMS, --');
    SQL.Add('');
    SQL.Add('        CASE');
    SQL.Add('        WHEN ITENS_VENDA.CST = ''000'' THEN ''T''');
    SQL.Add('        WHEN ITENS_VENDA.CST = ''060'' THEN ''F''');
    SQL.Add('        WHEN ITENS_VENDA.CST = ''040'' THEN ''I''');
    SQL.Add('    END AS DES_SITUACAO_TRIB, --');
    SQL.Add('');
    SQL.Add('    ITENS_VENDA.VALOR_ICMS_VENDA AS VAL_ICMS, --');
    SQL.Add('');
    SQL.Add('    VENDAS.DOCUMENTO AS NUM_CUPOM_FISCAL,');
    SQL.Add('    ITENS_VENDA.VENDATAB AS VAL_VENDA_PDV,');
    SQL.Add('    ');
    SQL.Add('    CASE');
         SQL.Add('   WHEN ITENS_VENDA.CST = ''000'' AND ITENS_VENDA.ALIQUOTA = 7 THEN 2   ');
         SQL.Add('   WHEN ITENS_VENDA.CST = ''000'' AND ITENS_VENDA.ALIQUOTA = 25 THEN 5   ');
         SQL.Add('   WHEN ITENS_VENDA.CST = ''000'' AND ITENS_VENDA.ALIQUOTA = 4.5 THEN 38   ');
         SQL.Add('   WHEN ITENS_VENDA.CST = ''000'' AND ITENS_VENDA.ALIQUOTA = 12 THEN 3   ');
         SQL.Add('   WHEN ITENS_VENDA.CST = ''000'' AND ITENS_VENDA.ALIQUOTA = 11 THEN 35   ');
         SQL.Add('   WHEN ITENS_VENDA.CST = ''000'' AND ITENS_VENDA.ALIQUOTA = 8.8 THEN 24   ');
         SQL.Add('   WHEN ITENS_VENDA.CST = ''000'' AND ITENS_VENDA.ALIQUOTA = 5.5 THEN 38   ');
         SQL.Add('   WHEN ITENS_VENDA.CST = ''000'' AND ITENS_VENDA.ALIQUOTA = 5.6 THEN 38   ');
         SQL.Add('   WHEN ITENS_VENDA.CST = ''000'' AND ITENS_VENDA.ALIQUOTA = 4.7 THEN 53   ');
         SQL.Add('   WHEN ITENS_VENDA.CST = ''000'' AND ITENS_VENDA.ALIQUOTA = 18 THEN 4   ');
         SQL.Add('   WHEN ITENS_VENDA.CST = ''020'' AND ITENS_VENDA.ALIQUOTA = 18 THEN 8   ');
         SQL.Add('   WHEN ITENS_VENDA.CST = ''020'' AND ITENS_VENDA.ALIQUOTA = 7 THEN 6   ');
         SQL.Add('   WHEN ITENS_VENDA.CST = ''020'' AND ITENS_VENDA.ALIQUOTA = 12 THEN 6   ');
         SQL.Add('   WHEN ITENS_VENDA.CST = ''040'' AND ITENS_VENDA.ALIQUOTA = 0 THEN 1   ');
         SQL.Add('   WHEN ITENS_VENDA.CST = ''041'' AND ITENS_VENDA.ALIQUOTA = 0 THEN 23   ');
         SQL.Add('   WHEN ITENS_VENDA.CST = ''060'' AND ITENS_VENDA.ALIQUOTA = 25 THEN 25   ');
         SQL.Add('   WHEN ITENS_VENDA.CST = ''060'' AND ITENS_VENDA.ALIQUOTA = 22 THEN 25   ');
         SQL.Add('   WHEN ITENS_VENDA.CST = ''060'' AND ITENS_VENDA.ALIQUOTA = 30 THEN 25   ');
         SQL.Add('   WHEN ITENS_VENDA.CST = ''060'' AND ITENS_VENDA.ALIQUOTA = 18 THEN 25   ');
         SQL.Add('   WHEN ITENS_VENDA.CST = ''060'' AND ITENS_VENDA.ALIQUOTA = 20 THEN 25   ');
         SQL.Add('   WHEN ITENS_VENDA.CST = ''060'' AND ITENS_VENDA.ALIQUOTA = 0 THEN 25   ');
         SQL.Add('   WHEN ITENS_VENDA.CST = ''060'' AND ITENS_VENDA.ALIQUOTA = 12 THEN 25   ');
    SQL.Add('    END AS COD_TRIBUTACAO,');
    SQL.Add('');
    SQL.Add('    CASE VENDAS.CANCELADO');
    SQL.Add('        WHEN ''F'' THEN ''N''');
    SQL.Add('        ELSE ''N''');
    SQL.Add('    END AS FLG_CUPOM_CANCELADO,');
    SQL.Add('');
    SQL.Add('    CASE ');
    SQL.Add('        WHEN CAST(COALESCE(PRODUTOS.CLASSIFICACAO_FISCAL, ''99999999'') AS INTEGER ) = 0 THEN ''99999999''');
    SQL.Add('        ELSE COALESCE(PRODUTOS.CLASSIFICACAO_FISCAL, ''99999999'')');
    SQL.Add('    END AS NUM_NCM,');
    SQL.Add('');
    SQL.Add('    ''999'' AS COD_TAB_SPED, --');
    SQL.Add('');
    SQL.Add('    ITENS_VENDA.CST_PIS,');
    SQL.Add('');
    SQL.Add('    CASE ITENS_VENDA.CST_PIS');
    SQL.Add('        WHEN ''01'' THEN ''N''');
    SQL.Add('        ELSE ''S''');
    SQL.Add('    END AS FLG_NAO_PIS_COFINS,');
    SQL.Add('');
    SQL.Add('    CASE ITENS_VENDA.CST_PIS');
    SQL.Add('        WHEN ''01'' THEN -1');
    SQL.Add('        WHEN ''04'' THEN 1');
    SQL.Add('        WHEN ''05'' THEN 2');
    SQL.Add('        WHEN ''06'' THEN 0');
    SQL.Add('        WHEN ''07'' THEN 1');
    SQL.Add('        WHEN ''08'' THEN 1');
    SQL.Add('    END AS TIPO_NAO_PIS_COFINS,');
    SQL.Add('');
    SQL.Add('    ''N'' AS FLG_ONLINE,    ');
    SQL.Add('    ''N'' AS FLG_OFERTA,');
    SQL.Add('    0 AS COD_ASSOCIADO  --');
    SQL.Add('FROM VENDAS');
    SQL.Add('');
    SQL.Add('INNER JOIN ITENS_VENDA');
    SQL.Add('ON VENDAS.ID = ITENS_VENDA.VENDA -- L = 17789897 I = 17790711 ');
    SQL.Add('');
    SQL.Add('INNER JOIN PRODUTOS');
    SQL.Add('ON ITENS_VENDA.PRODUTO = PRODUTOS.ID -- 17790581    ');
    SQL.Add('');
    SQL.Add('LEFT JOIN PRODUTOS_LOJA');
    SQL.Add('ON ITENS_VENDA.PRODUTO = PRODUTOS_LOJA.ID');
    SQL.Add('AND VENDAS.EMPRESA = PRODUTOS_LOJA.POLITICA');
    SQL.Add('');
    SQL.Add('LEFT JOIN CLIENTES    ');
    SQL.Add('ON VENDAS.CNPJ_CPF = CLIENTES.CNPJ_CPF');
    SQL.Add('');
    SQL.Add('WHERE ITENS_VENDA.QTDE > 0 ');
    SQL.Add('AND VENDAS.CANCELADO = ''F''  ');
    SQL.Add('AND VENDAS.EMPRESA = 1');

    SQL.Add('AND');
    SQL.Add('    VENDAS.DATA_HORA >= TO_DATE(''' +FormatDateTime('dd/mm/yyyy',DtpInicial.Date)+ ' 00:00:00'', ''DD/MM/YYYY hh24:mi:ss'')');
    SQL.Add('AND ');
    SQL.Add('    VENDAS.DATA_HORA <= TO_DATE(''' +FormatDateTime('dd/mm/yyyy',DtpFinal.Date)+ ' 23:59:59'', ''DD/MM/YYYY hh24:mi:ss'')');

//    showmessage(sql.text);

    Open;

    First;

    NumLinha := 0;

    custo := 0;

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);

      Layout.SetValues(QryPrincipal2, NumLinha, RecordCount);

//      Layout.FieldByName('COD_PRODUTO').AsString := GerarPLU(QryPrincipal2.FieldByName('COD_PRODUTO').AsString);

      if not PLUValido(Layout.FieldByName('COD_PRODUTO').AsString) then
      begin
        Layout.FieldByName('COD_PRODUTO').AsString := GerarPlu(Copy(Layout.FieldByName('COD_PRODUTO').AsString, 1, Length(Layout.FieldByName('COD_PRODUTO').AsString) - 1));
      end;

        if( QryPrincipal2.FieldByName('VAL_CUSTO_REP').AsFloat = 0 ) then
        begin

          QryCusto.Close;
          QryCusto.Parameters.ParamByName('COD_PRODUTO').Value := QryPrincipal2.FieldByName('COD_PRODUTO').AsString;
          QryCusto.Parameters.ParamByName('COD_LOJA').Value := QryPrincipal2.FieldByName('COD_LOJA').AsInteger;

          QryCusto.Parameters.ParamByName('DTA_SAIDA').Value := FormatDateTime('dd/mm/yyyy',QryPrincipal2.FieldByName('DTA_SAIDA').AsDateTime) + '23:59:59';
  //        QryCusto.Parameters.ParamByName('DTA_SAIDA').Value := FormatDateTime('dd/mm/yyyy hh:mm:ss',QryPrincipal2.FieldByName('DTA_SAIDA').AsDateTime);
  //         QryCusto.Parameters.ParamByName('DTA_SAIDA').Value := FormatDateTime('dd/mm/yyyy',DtpFinal.Date)+ ' 23:59:59';

          QryCusto.Open;

  //        Layout.FieldByName('VAL_CUSTO_REP').AsFloat := QryCusto.FieldByName('CUSTO_ATUAL').AsFloat;

             Layout.FieldByName('VAL_CUSTO_REP').AsFloat := RoundTo(QryCusto.FieldByName('CUSTO_ATUAL').AsFloat, -3);
        end;


//        if QryPrincipal2.FieldByName('COD_PRODUTO').AsInteger = 357074 then
//        begin
////         showmessage('atual: ' + QryCusto.FieldByName('CUSTO_ATUAL').AsString);
////         showmessage('data: ' + QryPrincipal2.FieldByName('DTA_SAIDA').AsString);
//
//            custo := custo + QryPrincipal2.FieldByName('VAL_CUSTO_REP').AsCurrency;
//        end;


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

procedure TFrmSmKuzzi.BtnAmarrarCestClick(Sender: TObject);
begin
  inherited;
    inherited;
  FlgGeraAmarrarCest := True;
  BtnGerar.Click;
  FlgGeraAmarrarCest := False;
end;

procedure TFrmSmKuzzi.btnGeraCestClick(Sender: TObject);
begin
  inherited;
  FlgGeraCest := True;
  BtnGerar.Click;
  FlgGeraCest := False;
end;

procedure TFrmSmKuzzi.BtnGerarClick(Sender: TObject);
begin
    ADOOracle.Connected := false;
    ADOOracle.ConnectionString := 'Provider=MSDAORA.1;Password='+ edtSenhaOracle.Text +';User ID='+ edtSchema.Text +';Data Source='+ edtInst.Text +'';
    ADOOracle.Connected := true;
  inherited;
    ADOOracle.Connected := false;
end;

procedure TFrmSmKuzzi.GerarCest;
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

procedure TFrmSmKuzzi.GerarCliente;
begin
  inherited;

  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
    SQL.Add('    18000 + CONVENIADAS.ID AS COD_CLIENTE,');
    SQL.Add('    CONVENIADAS.DESCRITIVO AS DES_CLIENTE,');
    SQL.Add('    '''' AS NUM_CGC,');
    SQL.Add('    CONVENIADAS.INSCRICAO_RG AS NUM_INSC_EST,');
    SQL.Add('    CONVENIADAS.LOGRADOURO || '' '' || CONVENIADAS.ENDERECO AS DES_ENDERECO,');
    SQL.Add('    CONVENIADAS.BAIRRO AS DES_BAIRRO,');
    SQL.Add('    CONVENIADAS.CIDADE AS DES_CIDADE,');
    SQL.Add('    CONVENIADAS.ESTADO AS DES_SIGLA,');
    SQL.Add('    CONVENIADAS.CEP AS NUM_CEP,');
    SQL.Add('    CONVENIADAS.TELEFONE1 AS NUM_FONE,');
    SQL.Add('    CONVENIADAS.FAX AS NUM_FAX,');
    SQL.Add('    '''' AS DES_CONTATO,');
    SQL.Add('    0 AS FLG_SEXO,');
    SQL.Add('    0 AS VAL_LIMITE_CRETID,');
    SQL.Add('    0 AS VAL_LIMITE_CONV,');
    SQL.Add('    0 AS VAL_DEBITO,');
    SQL.Add('    0 AS VAL_RENDA,');
    SQL.Add('    0 AS COD_CONVENIO,');
    SQL.Add('    0 AS COD_STATUS_PDV,');
    SQL.Add('    ''S'' AS FLG_EMPRESA,');
    SQL.Add('    ''S'' AS FLG_CONVENIO,');
    SQL.Add('    ''N'' AS MICRO_EMPRESA,');
    SQL.Add('    CONVENIADAS.DATAHORA_CADASTRO AS DTA_CADASTRO,');
    SQL.Add('    CONVENIADAS.NUMERO AS NUM_ENDERECO,');
    SQL.Add('    '''' AS NUM_RG,');
    SQL.Add('    0 AS FLG_EST_CIVIL,');
    SQL.Add('    CONVENIADAS.TELEFONE2 AS NUM_CELULAR,');
    SQL.Add('    CONVENIADAS.DATAHORA_ALTERACAO AS DTA_ALTERACAO,');
    SQL.Add('    CONVENIADAS.OBSERVACAO AS DES_OBSERVACAO,');
    SQL.Add('    CONVENIADAS.COMPLEMENTO AS DES_COMPLEMENTO,');
    SQL.Add('    '''' AS DES_EMAIL,');
    SQL.Add('    CONVENIADAS.FANTASIA AS DES_FANTASIA,');
    SQL.Add('    null AS DTA_NASCIMENTO,');
    SQL.Add('    '''' AS DES_PAI,');
    SQL.Add('    '''' AS DES_MAE,');
    SQL.Add('    '''' AS DES_CONJUGE,');
    SQL.Add('    '''' AS NUM_CPF_CONJUGE,');
    SQL.Add('    0 AS VAL_DEB_CONV,');
    SQL.Add('    ''N'' AS INATIVO,');
    SQL.Add('    '''' AS DES_MATRICULA,');
    SQL.Add('    ''N'' AS NUM_CGC_ASSOCIADO,');
    SQL.Add('    ''N'' AS FLG_PROD_RURAL,');
    SQL.Add('    0 AS COD_STATUS_PDV_CONV,');
    SQL.Add('    ''N'' AS FLG_ENVIA_CODIGO,');
    SQL.Add('    '''' AS DTA_NASC_CONJUGE,');
    SQL.Add('    0 AS COD_CLASSIF');
    SQL.Add('FROM CONVENIADAS');
    SQL.Add('');
    SQL.Add('UNION ALL');
    SQL.Add('');
    SQL.Add('SELECT');
    SQL.Add('    CLIENTES.ID AS COD_CLIENTE,');
    SQL.Add('    CLIENTES.DESCRITIVO AS DES_CLIENTE,');
    SQL.Add('    CLIENTES.CNPJ_CPF AS NUM_CGC,');
    SQL.Add('    CLIENTES.INSCRICAO_RG AS NUM_INSC_EST,');
    SQL.Add('    CLIENTES.LOGRADOURO || '' '' || CLIENTES.ENDERECO AS DES_ENDERECO,');
    SQL.Add('    CLIENTES.BAIRRO AS DES_BAIRRO,');
    SQL.Add('    CLIENTES.CIDADE AS DES_CIDADE,');
    SQL.Add('    CLIENTES.ESTADO AS DES_SIGLA,');
    SQL.Add('    CLIENTES.CEP AS NUM_CEP,');
    SQL.Add('    CLIENTES.TELEFONE1 AS NUM_FONE,');
    SQL.Add('    CLIENTES.FAX AS NUM_FAX,');
    SQL.Add('    CLIENTES.FANTASIA AS DES_CONTATO,');
    SQL.Add('');
    SQL.Add('    CASE');
    SQL.Add('        WHEN CLIENTES.SEXO = 1 THEN 1');
    SQL.Add('        ELSE 0 ');
    SQL.Add('    END AS FLG_SEXO, -- SEXO');
    SQL.Add('');
    SQL.Add('    0 AS VAL_LIMITE_CRETID, ');
    SQL.Add('    CLIENTES.LIMITE AS VAL_LIMITE_CONV, --VERIFICAR SITUACAO, LIMITE');
    SQL.Add('    0 AS VAL_DEBITO,');
    SQL.Add('    CLIENTES.SALARIO AS VAL_RENDA,');
    SQL.Add('    18000 + CLIENTES.EMPRESA_CONVENIO AS COD_CONVENIO,');
    SQL.Add('    CASE CLIENTES.SITUACAO');
    SQL.Add('        WHEN 0 THEN 0');
    SQL.Add('        ELSE 1');
    SQL.Add('    END AS COD_STATUS_PDV,');
    SQL.Add('');
    SQL.Add('    CASE CLIENTES.PESSOA');
    SQL.Add('        WHEN ''F'' THEN ''N''');
    SQL.Add('        ELSE ''S''');
    SQL.Add('    END AS FLG_EMPRESA,');
    SQL.Add('');
    SQL.Add('    ''N'' AS FLG_CONVENIO, --');
    SQL.Add('    ''N'' AS MICRO_EMPRESA,');
    SQL.Add('    CLIENTES.DATAHORA_CADASTRO AS DTA_CADASTRO,');
    SQL.Add('    CLIENTES.NUMERO AS NUM_ENDERECO,');
    SQL.Add('    CLIENTES.INSCRICAO_RG AS NUM_RG, --OBSERVACAO');
    SQL.Add('');
    SQL.Add('    CASE CLIENTES.ESTADO_CIVIL');
    SQL.Add('        WHEN 0 THEN 0');
    SQL.Add('        WHEN 1 THEN 1');
    SQL.Add('        WHEN 2 THEN 3');
    SQL.Add('        WHEN 3 THEN 2');
    SQL.Add('        WHEN 4 THEN 4');
    SQL.Add('        ELSE 0');
    SQL.Add('    END AS FLG_EST_CIVIL, --ESTADO_CIVIL');
    SQL.Add('');
    SQL.Add('    CLIENTES.TELEFONE2 AS NUM_CELULAR, -- TELEFONE2');
    SQL.Add('    CLIENTES.DATAHORA_ALTERACAO AS DTA_ALTERACAO,');
    SQL.Add('    CLIENTES.OBSERVACAO || '' PONTO DE REFERENCIA: '' || CLIENTES.REFERENCIA AS DES_OBSERVACAO,');
    SQL.Add('    CLIENTES.COMPLEMENTO AS DES_COMPLEMENTO,');
    SQL.Add('    CLIENTES.EMAIL AS DES_EMAIL,');
    SQL.Add('    CLIENTES.FANTASIA AS DES_FANTASIA,');
    SQL.Add('    CLIENTES.DATA_NASCIMENTO AS DTA_NASCIMENTO,');
    SQL.Add('    CLIENTES.PAI AS DES_PAI,');
    SQL.Add('    CLIENTES.MAE AS DES_MAE,');
    SQL.Add('    CLIENTES.CONJUGUE AS DES_CONJUGE,');
    SQL.Add('    CLIENTES.CPF_CONJUGE AS NUM_CPF_CONJUGE,');
    SQL.Add('    0 AS VAL_DEB_CONV,');
    SQL.Add('    ''N'' AS INATIVO,');
    SQL.Add('    '''' AS DES_MATRICULA,');
    SQL.Add('    ''N'' AS NUM_CGC_ASSOCIADO,');
    SQL.Add('    ''N'' AS FLG_PROD_RURAL,');
    SQL.Add('    CASE CLIENTES.SITUACAO');
    SQL.Add('        WHEN 0 THEN 0');
    SQL.Add('        ELSE 1');
    SQL.Add('    END AS COD_STATUS_PDV_CONV,');
    SQL.Add('    ''S'' AS FLG_ENVIA_CODIGO,');
    SQL.Add('    '''' AS DTA_NASC_CONJUGE,');
    SQL.Add('    0 AS COD_CLASSIF');
    SQL.Add('FROM CLIENTES');



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

      Layout.FieldByName('NUM_CEP').AsString := StrRetNums(Layout.FieldByName('NUM_CEP').AsString);

      Layout.FieldByName('DTA_NASCIMENTO').AsDateTime := FieldByName('DTA_NASCIMENTO').AsDateTime;
      Layout.FieldByName('DTA_CADASTRO').AsDateTime := FieldByName('DTA_CADASTRO').AsDateTime;

      Layout.FieldByName('NUM_FONE').AsString := StrRetNums( FieldByName('NUM_FONE').AsString );

      if Layout.FieldByName('FLG_EMPRESA').AsString = 'S' then
      begin
        if Layout.FieldByName('NUM_INSC_EST').AsString = '' then
        begin
          Layout.FieldByName('NUM_INSC_EST').AsString := 'ISENTO';
          Layout.FieldByName('NUM_RG').AsString := '';
        end;
        Layout.FieldByName('NUM_RG').AsString := '';
        if not ValidaCGC(Layout.FieldByName('NUM_CGC').AsString) then
          Layout.FieldByName('NUM_CGC').AsString := '';
      end
      else
      begin
        if Layout.FieldByName('NUM_RG').AsString = '' then
        begin
          Layout.FieldByName('NUM_RG').AsString := 'ISENTO';
          Layout.FieldByName('NUM_INSC_EST').AsString := '';
        end;
        Layout.FieldByName('NUM_INSC_EST').AsString := '';
        if not ValidaCpf(Layout.FieldByName('NUM_CGC').AsString) then
          Layout.FieldByName('NUM_CGC').AsString := '';
      end;

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

procedure TFrmSmKuzzi.GerarCodigoBarras;
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
    SQL.Add('    PRODUTOS.ID AS COD_PRODUTO,');
    SQL.Add('    COALESCE(PRODUTOS_EAN.EAN, '''') AS COD_EAN,');
    SQL.Add('    ''P'' AS TIPO_CADASTRO');
    SQL.Add('FROM');
    SQL.Add('    PRODUTOS');
    SQL.Add('LEFT JOIN');
    SQL.Add('    PRODUTOS_EAN');
    SQL.Add('ON');
    SQL.Add('    PRODUTOS.ID = PRODUTOS_EAN.PRODUTO');
//    SQL.Add('AND');
//    SQL.Add('    PRODUTOS_EAN.QTDEE = 1    ');
//    SQL.Add('');
//    SQL.Add('UNION ALL');
//    SQL.Add('');
//    SQL.Add('SELECT');
//    SQL.Add('    CAST(SUBSTR(1110000 + EAN.PRODUTO, 1, 6) AS INTEGER) AS COD_PRODUTO,');
//    SQL.Add('    EAN.EAN AS COD_BARRA_PRINCIPAL,');
//    SQL.Add('    ''E'' AS TIPO_CADASTRO');
//    SQL.Add('FROM');
//    SQL.Add('    PRODUTOS_EAN EAN');
//    SQL.Add('WHERE');
//    SQL.Add('    EAN.QTDEE > 1    ');


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

      if not PLUValido(Layout.FieldByName('COD_PRODUTO').AsString) then
      begin
        Layout.FieldByName('COD_PRODUTO').AsString := GerarPlu(Copy(Layout.FieldByName('COD_PRODUTO').AsString, 1, Length(Layout.FieldByName('COD_PRODUTO').AsString) - 1));
      end;

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

procedure TFrmSmKuzzi.GerarComposicao;
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

      if not PLUValido(Layout.FieldByName('COD_PRODUTO').AsString) then
      begin
        Layout.FieldByName('COD_PRODUTO').AsString := GerarPlu(Copy(Layout.FieldByName('COD_PRODUTO').AsString, 1, Length(Layout.FieldByName('COD_PRODUTO').AsString) - 1));
      end;

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

procedure TFrmSmKuzzi.GerarCondPagCli;
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

procedure TFrmSmKuzzi.GerarCondPagForn;
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
    SQL.Add('    FORNECEDORES.CNPJ_CPF AS NUM_CGC');
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

procedure TFrmSmKuzzi.GerarDecomposicao;
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

procedure TFrmSmKuzzi.GerarDivisaoForn;
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

procedure TFrmSmKuzzi.GerarFinanceiro(Tipo, Situacao: Integer);
begin
  inherited;
  if Tipo = 1 then
    GerarFinanceiroPagar(IntToStr(Situacao));

  if Tipo = 2 then
    GerarFinanceiroReceber(IntToStr(Situacao));

  if Tipo = 3 then
    GerarFinanceiroReceberCartao;

end;

procedure TFrmSmKuzzi.GerarFinanceiroPagar(Aberto: String);
begin
  inherited;
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
    SQL.Add('    CASE PAGAR.TIPO_CADASTRO');
    SQL.Add('        WHEN 0 THEN 0');
    SQL.Add('        WHEN 3 THEN 2');
    SQL.Add('        WHEN 2 THEN 1');
    SQL.Add('    END AS TIPO_PARCEIRO, -- TIPO_CADASTRO');
    SQL.Add('    PAGAR.ID_CADASTRO AS COD_PARCEIRO,');
    SQL.Add('    0 AS TIPO_CONTA,');
    SQL.Add('');
    SQL.Add('    CASE PAGAR.FORMA_PAGTO');
           SQL.Add('   WHEN 1	THEN 8   ');
           SQL.Add('   WHEN 2	THEN 1   ');
           SQL.Add('   WHEN 3	THEN 3   ');
           SQL.Add('   WHEN 4	THEN 28   ');
           SQL.Add('   WHEN 5	THEN 27   ');
           SQL.Add('   WHEN 6	THEN 26   ');
           SQL.Add('   WHEN 7	THEN 6   ');
    SQL.Add('  ');
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
    SQL.Add('    ''PAGTO'' AS DES_CC,');
    SQL.Add('    0 AS COD_BANDEIRA,');
    SQL.Add('    '''' AS DTA_PRORROGACAO,');
    SQL.Add('    1 AS NUM_SEQ_FIN');
    SQL.Add('FROM CONTAS PAGAR');
    SQL.Add('LEFT JOIN');
    SQL.Add('    (');
    SQL.Add('        SELECT ');
    SQL.Add('            NF,');
    SQL.Add('            TIPO_CADASTRO,');
    SQL.Add('            ID_CADASTRO,');
    SQL.Add('            SUM(VALOR - DESCONTO + ACRESCIMO + CARTORIO + COALESCE(CREDITO, 0)) AS VAL_TOTAL_NF');
    SQL.Add('        FROM CONTAS  ');
    SQL.Add('        WHERE CONTAS.TIPO_CONTA = 0');
    SQL.Add('        AND CONTAS.EMPRESA = 1 ');
    SQL.Add('        AND CONTAS.PARCELA > 0');
    SQL.Add('        AND CONTAS.TIPO_CADASTRO IN (0, 2)');
    SQL.Add('        GROUP BY');
    SQL.Add('            NF,');
    SQL.Add('            TIPO_CADASTRO,');
    SQL.Add('            ID_CADASTRO');
    SQL.Add('    ) NF');
    SQL.Add('ON PAGAR.NF = NF.NF');
    SQL.Add('AND PAGAR.TIPO_CADASTRO = NF.TIPO_CADASTRO');
    SQL.Add('AND PAGAR.ID_CADASTRO = NF.ID_CADASTRO        ');
    SQL.Add('WHERE PAGAR.TIPO_CONTA = 0');
    SQL.Add('AND PAGAR.TIPO_CADASTRO IN (0, 2)');
    SQL.Add('AND PAGAR.PARCELA > 0 ');
    SQL.Add('AND PAGAR.EMPRESA = 1 ');
    SQL.Add('AND PAGAR.ID_CADASTRO IS NOT NULL');


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

procedure TFrmSmKuzzi.GerarFinanceiroReceber(Aberto: String);
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
    SQL.Add('        WHEN 0 THEN 0 -- cliente');
    SQL.Add('        WHEN 2 THEN 1 -- fornecedores');
    SQL.Add('        WHEN 5 THEN 0 -- convenio');
    SQL.Add('    END AS TIPO_PARCEIRO, -- TIPO_CADASTRO');
    SQL.Add('');
    SQL.Add('    CASE');
    SQL.Add('        WHEN RECEBER.TIPO_CADASTRO = 5 THEN 18000 + RECEBER.ID_CADASTRO ');
    SQL.Add('        ELSE RECEBER.ID_CADASTRO');
    SQL.Add('    END AS COD_PARCEIRO,  ');
    SQL.Add('');
    SQL.Add('    1 AS TIPO_CONTA,');
    SQL.Add('');
//    SQL.Add('    CASE RECEBER.FORMA_PAGTO');
//    SQL.Add('        WHEN 15 THEN 4');
//    SQL.Add('        ELSE 1');
    SQL.Add('    4 AS COD_ENTIDADE,');
    SQL.Add('');
//    SQL.Add('    CASE RECEBER.FORMA_PAGTO');
//                   SQL.Add('   WHEN 1	THEN 1   ');
//                   SQL.Add('   WHEN 2	THEN 2   ');
//                   SQL.Add('   WHEN 3	THEN 3   ');
//                   SQL.Add('   WHEN 4	THEN 6   ');
//                   SQL.Add('   WHEN 5	THEN 7   ');
//                   SQL.Add('   WHEN 6	THEN 4   ');
//                   SQL.Add('   WHEN 7	THEN 5   ');
//                   SQL.Add('   WHEN 8	THEN 9   ');
//                   SQL.Add('   WHEN 9	THEN 10   ');
//                   SQL.Add('   WHEN 10	THEN 11   ');
//                   SQL.Add('   WHEN 11	THEN 8   ');
//                   SQL.Add('   WHEN 13	THEN 12   ');
//                   SQL.Add('   WHEN 14	THEN 13   ');
//                   SQL.Add('   WHEN 15	THEN 14   ');
//                   SQL.Add('   WHEN 16	THEN 15   ');
//                   SQL.Add('   WHEN 17	THEN 16   ');
//                   SQL.Add('   WHEN 18	THEN 17   ');
//                   SQL.Add('   WHEN 19	THEN 18   ');
//                   SQL.Add('   WHEN 20	THEN 19   ');
//                   SQL.Add('   WHEN 21	THEN 20   ');
//                   SQL.Add('   WHEN 23	THEN 21   ');
//                   SQL.Add('   WHEN 24	THEN 22   ');
//                   SQL.Add('   WHEN 25	THEN 23   ');
//                   SQL.Add('   WHEN 26	THEN 24   ');
//                   SQL.Add('   WHEN 27	THEN 25   ');
//                   SQL.Add('   WHEN 28	THEN 26   ');
//    SQL.Add('    END AS COD_ENTIDADE,');
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
    SQL.Add('    ''997'' AS COD_CATEGORIA,');
    SQL.Add('');
    SQL.Add('    ''997'' AS COD_SUBCATEGORIA,');
    SQL.Add('');
    SQL.Add('    RECEBER.PARCELA AS NUM_PARCELA,');
    SQL.Add('    RECEBER.TOTAL_PARCELA AS QTD_PARCELA,');
    SQL.Add('    RECEBER.EMPRESA AS COD_LOJA,');
    SQL.Add('    RECEBER.CPF_CNPJ AS NUM_CGC,');
    SQL.Add('    COALESCE(RECEBER.BORDERO, 0) AS NUM_BORDERO,');
    SQL.Add('    RECEBER.NF AS NUM_NF,');
    SQL.Add('    '''' AS NUM_SERIE_NF,');
    SQL.Add('    CASE ');
    SQL.Add('        WHEN NF.VAL_TOTAL_NF = 0 THEN RECEBER.VALOR ');
    SQL.Add('        ELSE NF.VAL_TOTAL_NF ');
    SQL.Add('    END AS VAL_TOTAL_NF, -- EFETUAR A SOMA');
    SQL.Add('    ''COBRANÇA: '' || RECEBER.DATACOB || '' | 1 DEVOL: '' || RECEBER.DEVOLUCAOA || '' | 2 DEVOL : '' || RECEBER.DEVOLUCAOB || '' | ''  || RECEBER.OBSERVACAO AS DES_OBSERVACAO,');
    SQL.Add('    COALESCE(RECEBER.PDV, 0) AS NUM_PDV,');
    SQL.Add('    RECEBER.NF AS NUM_CUPOM_FISCAL,');
    SQL.Add('    0 AS COD_MOTIVO,');
    SQL.Add('');
    SQL.Add('    CASE RECEBER.FORMA_PAGTO');
    SQL.Add('        WHEN 15 THEN (SELECT COALESCE(18000 + CLIENTES.EMPRESA_CONVENIO, 0) FROM CLIENTES WHERE CLIENTES.ID = RECEBER.ID_CADASTRO)');
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
    SQL.Add('    ''RECEBTO'' AS DES_CC,');
    SQL.Add('');
    SQL.Add('    0 AS COD_BANDEIRA,');
    SQL.Add('');
    SQL.Add('');
    SQL.Add('    '''' AS DTA_PRORROGACAO,');
    SQL.Add('    1 AS NUM_SEQ_FIN,');
    SQL.Add('    0 AS COD_COBRANCA,');
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
    SQL.Add('');
    SQL.Add('FROM CONTAS RECEBER');
    SQL.Add('LEFT JOIN');
    SQL.Add('    (');
    SQL.Add('        SELECT ');
    SQL.Add('            NF,');
    SQL.Add('            TIPO_CADASTRO,');
    SQL.Add('            ID_CADASTRO,');
    SQL.Add('            SUM(VALOR - DESCONTO + ACRESCIMO + CARTORIO + COALESCE(CREDITO, 0)) AS VAL_TOTAL_NF');
    SQL.Add('        FROM CONTAS  ');
    SQL.Add('        WHERE CONTAS.TIPO_CONTA = 1');
    SQL.Add('        AND CONTAS.EMPRESA = 1');
    SQL.Add('        AND CONTAS.TIPO_CADASTRO IN (0, 2, 5) -- Adicionar o filtro de cartoes');
    SQL.Add('        AND CONTAS.PARCELA > 0');
    SQL.Add('        AND CONTAS.VALOR > 0');
    SQL.Add('        GROUP BY');
    SQL.Add('            NF,');
    SQL.Add('            TIPO_CADASTRO,');
    SQL.Add('            ID_CADASTRO');
    SQL.Add('    ) NF');
    SQL.Add('ON RECEBER.NF = NF.NF');
    SQL.Add('AND RECEBER.TIPO_CADASTRO = NF.TIPO_CADASTRO');
    SQL.Add('AND RECEBER.ID_CADASTRO = NF.ID_CADASTRO');
    SQL.Add('WHERE RECEBER.TIPO_CONTA = 1');
    SQL.Add('AND RECEBER.TIPO_CADASTRO IN (0, 2, 5) -- Adicionar o filtro de cartoes');
    SQL.Add('AND RECEBER.PARCELA > 0');
    SQL.Add('AND RECEBER.VALOR > 0');
    SQL.Add('AND RECEBER.EMPRESA = 1');

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

procedure TFrmSmKuzzi.GerarFinanceiroReceberCartao;
begin
  inherited;
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
    SQL.Add('');
    SQL.Add('CASE RECEBER.TIPO_CADASTRO');
    SQL.Add('    WHEN 0 THEN 0');
    SQL.Add('    WHEN 1 THEN 3');
    SQL.Add('    WHEN 4 THEN 4');
    SQL.Add('    WHEN 5 THEN 0');
    SQL.Add('END AS TIPO_PARCEIRO, -- TIPO_CADASTRO');
    SQL.Add('');
    SQL.Add('CASE');
    SQL.Add('    WHEN RECEBER.TIPO_CADASTRO = 5 THEN 2400 + RECEBER.ID_CADASTRO ');
    SQL.Add('    WHEN RECEBER.TIPO_CADASTRO = 5 AND COALESCE(RECEBER.ID_CADASTRO, 0) = 0 THEN 6');
    SQL.Add('    WHEN RECEBER.TIPO_CADASTRO = 4 THEN 99');
    SQL.Add('    ELSE CASE WHEN COALESCE(RECEBER.ID_CADASTRO, 0) = 0 THEN 99999 ELSE RECEBER.ID_CADASTRO END');
    SQL.Add('END AS COD_PARCEIRO,  ');
    SQL.Add('');
    SQL.Add('1 AS TIPO_CONTA,');
    SQL.Add('');
    SQL.Add('-- CASE RECEBER.FORMA_PAGTO');
    SQL.Add('--     WHEN 1 THEN 1');
    SQL.Add('--     WHEN 2 THEN 2');
    SQL.Add('--     WHEN 3 THEN 4');
    SQL.Add('--     WHEN 4 THEN 10');
    SQL.Add('--     WHEN 5 THEN 11');
    SQL.Add('--     WHEN 6 THEN 6');
    SQL.Add('--     WHEN 7 THEN 12');
    SQL.Add('--     WHEN 8 THEN 3');
    SQL.Add('--     WHEN 9 THEN 13');
    SQL.Add('--     WHEN 10 THEN 5');
    SQL.Add('--     WHEN 11 THEN 7');
    SQL.Add('--     WHEN 12 THEN 14');
    SQL.Add('--     WHEN 13 THEN 15');
    SQL.Add('--     WHEN 14 THEN 16');
    SQL.Add('--     WHEN 15 THEN 17');
    SQL.Add('--     WHEN 16 THEN 18');
    SQL.Add('--     WHEN 17 THEN 19');
    SQL.Add('--     WHEN 18 THEN 20');
    SQL.Add('--     WHEN 19 THEN 21');
    SQL.Add('--     WHEN 20 THEN 22');
    SQL.Add('--     WHEN 21 THEN 23');
    SQL.Add('--     WHEN 22 THEN 24');
    SQL.Add('--     WHEN 23 THEN 25');
    SQL.Add('--     WHEN 24 THEN 26');
    SQL.Add('--     WHEN 25 THEN 27');
    SQL.Add('--     ELSE 1');
    SQL.Add('-- END AS COD_ENTIDADE,');
    SQL.Add('1 AS COD_ENTIDADE,');
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
    SQL.Add('-- //    ADM_CARTOES.DESCRITIVO AS DES_BANDEIRA,');
    SQL.Add(''''' AS DES_BANDEIRA,');
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
    SQL.Add('10000 + RECEBER.ID_CADASTRO AS COD_BANDEIRA,');
    SQL.Add('  ');
    SQL.Add(''''' AS DTA_PRORROGACAO,');
    SQL.Add('1 AS NUM_SEQ_FIN,');
    SQL.Add('0 AS COD_COBRANCA,');
    SQL.Add('RECEBER.DATACOB AS DTA_COBRANCA,');
    SQL.Add('''N'' AS FLG_ACEITE,');
    SQL.Add('0 AS TIPO_ACEITE');
    SQL.Add('');
    SQL.Add('FROM CONTAS RECEBER');
    SQL.Add('LEFT JOIN');
    SQL.Add('  (');
    SQL.Add('      SELECT ');
    SQL.Add('          NF,');
    SQL.Add('          TIPO_CADASTRO,');
    SQL.Add('          ID_CADASTRO,');
    SQL.Add('          SUM(VALOR - DESCONTO + ACRESCIMO + CARTORIO + COALESCE(CREDITO, 0)) AS VAL_TOTAL_NF');
    SQL.Add('      FROM CONTAS  ');
    SQL.Add('      WHERE CONTAS.TIPO_CONTA = 1');
    SQL.Add('      AND CONTAS.EMPRESA = 1');
    SQL.Add('      AND CONTAS.TIPO_CADASTRO IN (4) -- Adicionar o filtro de cartoes');
    SQL.Add('      AND CONTAS.PARCELA > 0');
    SQL.Add('      AND CONTAS.VALOR > 0');
    SQL.Add('      GROUP BY');
    SQL.Add('          NF,');
    SQL.Add('          TIPO_CADASTRO,');
    SQL.Add('          ID_CADASTRO');
    SQL.Add('  ) NF');
    SQL.Add('ON RECEBER.NF = NF.NF');
    SQL.Add('AND RECEBER.TIPO_CADASTRO = NF.TIPO_CADASTRO');
    SQL.Add('AND RECEBER.ID_CADASTRO = NF.ID_CADASTRO        ');
    SQL.Add('WHERE RECEBER.TIPO_CONTA = 1');
    SQL.Add('AND RECEBER.TIPO_CADASTRO IN (4) -- Adicionar o filtro de cartoes');
    SQL.Add('AND RECEBER.PARCELA > 0');
    SQL.Add('AND RECEBER.VALOR > 0');
    SQL.Add('AND RECEBER.EMPRESA = 1 ');


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

procedure TFrmSmKuzzi.GerarFornecedor;
var
   observacao, email : string;
begin
  inherited;

  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
    SQL.Add('    FORNECEDORES.ID AS COD_FORNECEDOR,');
    SQL.Add('    FORNECEDORES.DESCRITIVO AS DES_FORNECEDOR,');
    SQL.Add('    FORNECEDORES.FANTASIA AS DES_FANTASIA,');
    SQL.Add('    FORNECEDORES.CNPJ_CPF AS NUM_CGC,');
    SQL.Add('    FORNECEDORES.INSCRICAO_RG AS NUM_INSC_EST,');
    SQL.Add('    FORNECEDORES.LOGRADOURO || '' '' || FORNECEDORES.ENDERECO AS DES_ENDERECO,');
    SQL.Add('    FORNECEDORES.BAIRRO AS DES_BAIRRO,');
    SQL.Add('    FORNECEDORES.CIDADE AS DES_CIDADE,');
    SQL.Add('    FORNECEDORES.ESTADO AS DES_SIGLA,');
    SQL.Add('    FORNECEDORES.CEP AS NUM_CEP,');
    SQL.Add('    FORNECEDORES.TELEFONE1 AS NUM_FONE,');
    SQL.Add('    FORNECEDORES.FAX AS NUM_FAX,');
    SQL.Add('    '''' AS DES_CONTATO,');
    SQL.Add('    FORNECEDORES.CARENCIA AS QTD_DIA_CARENCIA,');
    SQL.Add('    FORNECEDORES.FREQUENCIA AS NUM_FREQ_VISITA,');
    SQL.Add('    0 AS VAL_DESCONTO,');
    SQL.Add('    FORNECEDORES.ENTREGA AS NUM_PRAZO,');
    SQL.Add('');
    SQL.Add('    CASE FORNECEDORES.DEVOLUCAO');
    SQL.Add('        WHEN ''T'' THEN ''S''');
    SQL.Add('        ELSE ''N''');
    SQL.Add('    END AS ACEITA_DEVOL_MER, -- DEVOLUCAO');
    SQL.Add('');
    SQL.Add('    ''N'' AS CAL_IPI_VAL_BRUTO,');
    SQL.Add('    ''N'' AS CAL_ICMS_ENC_FIN,');
    SQL.Add('    ''N'' AS CAL_ICMS_VAL_IPI,');
    SQL.Add('');
    SQL.Add('    CASE FORNECEDORES.MICRO');
    SQL.Add('        WHEN ''T'' THEN ''S''');
    SQL.Add('        ELSE ''N''');
    SQL.Add('    END AS MICRO_EMPRESA,');
    SQL.Add('');
    SQL.Add('    0 AS COD_FORNECEDOR_ANT,');
    SQL.Add('    FORNECEDORES.NUMERO AS NUM_ENDERECO,');
    SQL.Add('    FORNECEDORES.OBSERVACAO || '' OBSERVACAO PEDIDO: '' || REPLACE(FORNECEDORES.OBSERVACAO_PEDIDO, ''&'', '''') AS DES_OBSERVACAO,');
    SQL.Add('    FORNECEDORES.EMAIL AS DES_EMAIL,');
    SQL.Add('    FORNECEDORES.SITE AS DES_WEB_SITE,');
    SQL.Add('    ''N'' AS FABRICANTE,');
    SQL.Add('');
    SQL.Add('    CASE FORNECEDORES.PRODUTOR');
    SQL.Add('        WHEN ''T'' THEN ''S''');
    SQL.Add('        ELSE ''N''');
    SQL.Add('    END AS FLG_PRODUTOR_RURAL, -- PRODUTOR');
    SQL.Add('');
    SQL.Add('    FORNECEDORES.FRETE AS TIPO_FRETE, --FRETE');
    SQL.Add('    ''N'' AS FLG_SIMPLES,');
    SQL.Add('    ''N'' AS FLG_SUBSTITUTO_TRIB,');
    SQL.Add('    FORNECEDORES.CONTABIL AS COD_CONTACCFORN,');
    SQL.Add('    ''N'' AS INATIVO,');
    SQL.Add('    0 AS COD_CLASSIF,');
    SQL.Add('    FORNECEDORES.DATAHORA_CADASTRO AS DTA_CADASTRO,');
    SQL.Add('    0 AS VAL_CREDITO,');
    SQL.Add('    0 AS VAL_DEBITO,');
    SQL.Add('    FORNECEDORES.PEDMINIMO AS PED_MIN_VAL,');
    SQL.Add('    '''' AS DES_EMAIL_VEND,');
    SQL.Add('    '''' AS SENHA_COTACAO,');
    SQL.Add('    0 AS TIPO_PRODUTOR, -- VERIFICAR');
    SQL.Add('    FORNECEDORES.TELEFONE2 AS NUM_CELULAR');
    SQL.Add('FROM');
    SQL.Add('    FORNECEDORES');
    SQL.Add('ORDER BY');
    SQL.Add('    FORNECEDORES.ID');

    Open;

    First;
    NumLinha := 0;

//    AssignFile(Arquivo, EdtCamArquivo.Text + '\SCRIPT_ATUALIZA_OBSERVACAO_EMAIL.SQL');
//    Rewrite(Arquivo);

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
        begin
          Layout.FieldByName('NUM_CGC').AsString := '';
//          Layout.FieldByName('NUM_INSC_EST').AsString := '';
        end;
      end
      else
        if not ValidaCPF(Layout.FieldByName('NUM_CGC').AsString) then begin
          Layout.FieldByName('NUM_CGC').AsString := '';
//          Layout.FieldByName('NUM_RG').AsString := '';
        end;


      Layout.FieldByName('NUM_FONE').AsString := StrRetNums( FieldByName('NUM_FONE').AsString );

      observacao := StrReplace(StrLBReplace(FieldByName('DES_OBSERVACAO').AsString), '\n', '');
      email := StrReplace(StrLBReplace(FieldByName('DES_EMAIL').AsString), '\n', '');

//      Writeln(Arquivo, 'UPDATE TAB_FORNECEDOR SET DES_OBSERVACAO = '''+ copy(UTF8Encode(StringReplace(observacao, '&', '', [rfReplaceAll])), 1, 500) +''', DES_EMAIL = '''+ email +''' WHERE COD_FORNECEDOR = '+ QryPrincipal2.FieldByName('COD_FORNECEDOR').AsString +'; ');

      Layout.FieldByName('DES_OBSERVACAO').AsString := observacao;
      Layout.FieldByName('DES_EMAIL').AsString := email;



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
//    Writeln(Arquivo, 'COMMIT WORK;');
//    Close;
  end;
//  CloseFile(Arquivo);
end;

procedure TFrmSmKuzzi.GerarGrupo;
begin
  inherited;
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT DISTINCT');
    SQL.Add('    DEPTOS.DEPTO AS COD_SECAO,');
    SQL.Add('    CASE');
    SQL.Add('        WHEN DEPTOS.SECAO = 0 THEN 999');
    SQL.Add('        ELSE DEPTOS.SECAO');
    SQL.Add('    END AS COD_GRUPO,');
    SQL.Add('    CASE');
    SQL.Add('        WHEN DEPTOS.SECAO = 0 THEN ''A DEFINIR''');
    SQL.Add('        ELSE DEPTOS.DESCRITIVO');
    SQL.Add('    END AS DES_GRUPO,');
    SQL.Add('    DEPTOS.MARGEM AS VAL_META');
    SQL.Add('FROM DEPTOS');
    SQL.Add('WHERE DEPTOS.GRUPO = 0');
    SQL.Add('AND DEPTOS.SUBGRUPO = 0  ');

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

procedure TFrmSmKuzzi.GerarInfoNutricionais;
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
    SQL.Add('    (NUTRICIONAL.FIBRA_ALIMENTAR * 100) / 20 AS VD_FIBRA_ALIMENTAR,');
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

procedure TFrmSmKuzzi.GerarNCM;
var
 count : Integer;
begin
  inherited;
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT DISTINCT');
    SQL.Add('    0 AS COD_NCM,    ');
    SQL.Add('    COALESCE(NCM.DESCRICAO, ''A DEFINIR'') AS DES_NCM,');
    SQL.Add('');
    SQL.Add('    CASE ');
    SQL.Add('        WHEN CAST(COALESCE(PRODUTOS.CLASSIFICACAO_FISCAL, ''99999999'') AS INTEGER ) = 0 THEN ''99999999''');
    SQL.Add('        ELSE COALESCE(PRODUTOS.CLASSIFICACAO_FISCAL, ''99999999'')');
    SQL.Add('    END AS NUM_NCM,');
    SQL.Add('');
    SQL.Add('    CASE ');
    SQL.Add('        WHEN COALESCE(PRODUTOS.CEST, ''0000000'') = ''0000000'' THEN ''9999999'' ');
    SQL.Add('        ELSE PRODUTOS.CEST ');
    SQL.Add('    END AS NUM_CEST,');
    SQL.Add('');
    SQL.Add('    CASE PRODUTOS.MONOFASICO');
    SQL.Add('        WHEN ''I'' THEN ''S''');
    SQL.Add('        WHEN ''M'' THEN ''S''');
    SQL.Add('        WHEN ''T'' THEN ''N''');
    SQL.Add('        WHEN ''B'' THEN ''S''');
    SQL.Add('        WHEN ''O'' THEN ''S''');
    SQL.Add('        WHEN ''N'' THEN ''S''');
    SQL.Add('        WHEN ''S'' THEN ''S''');
    SQL.Add('    END AS FLG_NAO_PIS_COFINS,');
    SQL.Add('');
    SQL.Add('    CASE PRODUTOS.MONOFASICO');
    SQL.Add('        WHEN ''I'' THEN 0 -- ISENTO');
    SQL.Add('        WHEN ''M'' THEN 1 -- MONOFASICO');
    SQL.Add('        WHEN ''T'' THEN -1 -- INCIDENTE');
    SQL.Add('        WHEN ''B'' THEN 2 -- SUBSTITUICAO');
    SQL.Add('        WHEN ''O'' THEN 0 -- ALIQUOTA ZERO');
    SQL.Add('        WHEN ''N'' THEN 0 -- NÃO INCIDENTE');
    SQL.Add('        WHEN ''S'' THEN 4 -- SUSPENSO');
    SQL.Add('    END AS TIPO_NAO_PIS_COFINS,');
    SQL.Add('');
    SQL.Add('    ''999'' AS COD_TAB_SPED,');
    SQL.Add('');
    SQL.Add('           CASE   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 20 AND TRIB_SAIDA.ICMS_VENDA = 18 AND TRIB_SAIDA.REDUCAO_VENDA = ''61,11'' THEN 8   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 60 AND TRIB_SAIDA.ICMS_VENDA = 0 AND TRIB_SAIDA.REDUCAO_VENDA = ''41,67'' THEN 15   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 0 AND TRIB_SAIDA.ICMS_VENDA = 7 AND TRIB_SAIDA.REDUCAO_VENDA = 0 THEN 2   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 90 AND TRIB_SAIDA.ICMS_VENDA = 18 AND TRIB_SAIDA.REDUCAO_VENDA = 0 THEN 22   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 60 AND TRIB_SAIDA.ICMS_VENDA = 18 AND TRIB_SAIDA.REDUCAO_VENDA = ''61,11'' THEN 17   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 60 AND TRIB_SAIDA.ICMS_VENDA = 20 AND TRIB_SAIDA.REDUCAO_VENDA = 0 THEN 29   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 0 AND TRIB_SAIDA.ICMS_VENDA = 12 AND TRIB_SAIDA.REDUCAO_VENDA = 0 THEN 3   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 41 AND TRIB_SAIDA.ICMS_VENDA = 0 AND TRIB_SAIDA.REDUCAO_VENDA = 0 THEN 23   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 60 AND TRIB_SAIDA.ICMS_VENDA = 27 AND TRIB_SAIDA.REDUCAO_VENDA = 0 THEN 13   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 0 AND TRIB_SAIDA.ICMS_VENDA = 18 AND TRIB_SAIDA.REDUCAO_VENDA = 0 THEN 4   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 20 AND TRIB_SAIDA.ICMS_VENDA = 18 AND TRIB_SAIDA.REDUCAO_VENDA = ''51,11'' THEN 39   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 60 AND TRIB_SAIDA.ICMS_VENDA = 18 AND TRIB_SAIDA.REDUCAO_VENDA = 0 THEN 13   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 60 AND TRIB_SAIDA.ICMS_VENDA = 30 AND TRIB_SAIDA.REDUCAO_VENDA = 0 THEN 31   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 60 AND TRIB_SAIDA.ICMS_VENDA = 22 AND TRIB_SAIDA.REDUCAO_VENDA = 0 THEN 30   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 60 AND TRIB_SAIDA.ICMS_VENDA = 25 AND TRIB_SAIDA.REDUCAO_VENDA = 0 THEN 14   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 60 AND TRIB_SAIDA.ICMS_VENDA = 32 AND TRIB_SAIDA.REDUCAO_VENDA = 0 THEN 32   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 20 AND TRIB_SAIDA.ICMS_VENDA = 12 AND TRIB_SAIDA.REDUCAO_VENDA = ''41,67'' THEN 6   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 40 AND TRIB_SAIDA.ICMS_VENDA = 0 AND TRIB_SAIDA.REDUCAO_VENDA = 0 THEN 1   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 20 AND TRIB_SAIDA.ICMS_VENDA = 18 AND TRIB_SAIDA.REDUCAO_VENDA = ''33,33'' THEN 7   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 0 AND TRIB_SAIDA.ICMS_VENDA = 25 AND TRIB_SAIDA.REDUCAO_VENDA = 0 THEN 5   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 60 AND TRIB_SAIDA.ICMS_VENDA = 12 AND TRIB_SAIDA.REDUCAO_VENDA = 0 THEN 12   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 20 AND TRIB_SAIDA.ICMS_VENDA = 12 AND TRIB_SAIDA.REDUCAO_VENDA = ''53,33'' THEN 40   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 90 AND TRIB_SAIDA.ICMS_VENDA = 0 AND TRIB_SAIDA.REDUCAO_VENDA = 0 THEN 22   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 10 AND TRIB_SAIDA.ICMS_VENDA = 12 AND TRIB_SAIDA.REDUCAO_VENDA = 0 THEN 42   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 60 AND TRIB_SAIDA.ICMS_VENDA = 0 AND TRIB_SAIDA.REDUCAO_VENDA = 0 THEN 41   ');
    SQL.Add('               ELSE 1   ');
    SQL.Add('           END AS COD_TRIB_SAIDA,   ');
    SQL.Add('');
    SQL.Add('           CASE   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,95'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''11'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,88'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,51'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''2,85'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''7'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''4,5'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''12'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''2,56'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,85'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,93'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,82'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''2,83'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''1,36'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''2,86'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,38'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,07'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,3'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,67'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''1,54'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,48'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''1,25'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''1,86'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,87'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''2,84'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''2,38'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''2,82'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''2,48'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,9'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,92'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''2,87'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,94'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''4'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 27   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,62'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,41'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''1,44'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''4,7'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''13,3'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''5,5'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''2,58'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,1'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,91'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''2,33'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''1,83'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''2,76'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''25'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 5   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,63'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''2,49'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''0'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,45'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,65'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,15'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,97'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,58'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''1,49'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 10 AND TRIBUTACAO_ENTRADA.ICMS =''20'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 29   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 10 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 10 AND TRIBUTACAO_ENTRADA.ICMS =''25'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 14   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 10 AND TRIBUTACAO_ENTRADA.ICMS =''30'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 31   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 10 AND TRIBUTACAO_ENTRADA.ICMS =''4'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 28   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 10 AND TRIBUTACAO_ENTRADA.ICMS =''1,25'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 10 AND TRIBUTACAO_ENTRADA.ICMS =''3,3'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 10 AND TRIBUTACAO_ENTRADA.ICMS =''13,3'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 50   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 10 AND TRIBUTACAO_ENTRADA.ICMS =''0'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 41   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 10 AND TRIBUTACAO_ENTRADA.ICMS =''12'' AND TRIBUTACAO_ENTRADA.REDUCAO =''41,67'' THEN 15   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 10 AND TRIBUTACAO_ENTRADA.ICMS =''3,41'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 10 AND TRIBUTACAO_ENTRADA.ICMS =''3,07'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 10 AND TRIBUTACAO_ENTRADA.ICMS =''2,87'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 10 AND TRIBUTACAO_ENTRADA.ICMS =''3,1'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 10 AND TRIBUTACAO_ENTRADA.ICMS =''12'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 12   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 10 AND TRIBUTACAO_ENTRADA.ICMS =''1,86'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''61,11'' THEN 8   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''13,3'' AND TRIBUTACAO_ENTRADA.REDUCAO =''47,37'' THEN 57   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''25'' AND TRIBUTACAO_ENTRADA.REDUCAO =''52'' THEN 9   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 8   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''33,333'' THEN 7   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''77'' THEN 47   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''12'' AND TRIBUTACAO_ENTRADA.REDUCAO =''62,09'' THEN 36   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''51,56'' THEN 39   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''12'' AND TRIBUTACAO_ENTRADA.REDUCAO =''60'' THEN 36   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''52'' THEN 39   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''12'' AND TRIBUTACAO_ENTRADA.REDUCAO =''33,33'' THEN 7   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''61,12'' THEN 8   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''79,02'' THEN 47   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''62,09'' THEN 8   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''12'' AND TRIBUTACAO_ENTRADA.REDUCAO =''73,34'' THEN 6   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''12'' AND TRIBUTACAO_ENTRADA.REDUCAO =''82,55'' THEN 6   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''51,1'' THEN 39   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''33,3333'' THEN 7   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''48,89'' THEN 7   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''31,12'' THEN 7   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''52,21'' THEN 39   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''51,11'' THEN 39   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''78,4'' THEN 47   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''67,35'' THEN 8   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''62,94'' THEN 8   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''12'' AND TRIBUTACAO_ENTRADA.REDUCAO =''41,67'' THEN 6   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''78,39'' THEN 47   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''82,66'' THEN 47   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''33,33'' THEN 7   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''75,79'' THEN 47   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''73,97'' THEN 47   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''33,36'' THEN 7   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''13,3'' AND TRIBUTACAO_ENTRADA.REDUCAO =''47,368'' THEN 57   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''12'' AND TRIBUTACAO_ENTRADA.REDUCAO =''46,67'' THEN 6   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''56,97'' THEN 39   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''74,78'' THEN 47   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''69,14'' THEN 8   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''33,334'' THEN 7   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''12'' AND TRIBUTACAO_ENTRADA.REDUCAO =''41,66'' THEN 6   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''66,22'' THEN 8   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''66,67'' THEN 8   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''12'' AND TRIBUTACAO_ENTRADA.REDUCAO =''26,66'' THEN 37   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''12'' AND TRIBUTACAO_ENTRADA.REDUCAO =''40'' THEN 6   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''41,67'' THEN 6   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''83,88'' THEN 8   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 30 AND TRIBUTACAO_ENTRADA.ICMS =''0'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 1   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 40 AND TRIBUTACAO_ENTRADA.ICMS =''0'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 1   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 41 AND TRIBUTACAO_ENTRADA.ICMS =''0'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 23   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 51 AND TRIBUTACAO_ENTRADA.ICMS =''0'' AND TRIBUTACAO_ENTRADA.REDUCAO =''100'' THEN 20   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 51 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 20   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 51 AND TRIBUTACAO_ENTRADA.ICMS =''0'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 20   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''3,07'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''61,11'' THEN 17   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''2,56'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''3,82'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''3,3'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''33,33'' THEN 16   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''1,86'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''1,25'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''3,85'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''3,93'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''2,41'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''7'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 11   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''0'' AND TRIBUTACAO_ENTRADA.REDUCAO =''33,33'' THEN 16   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''25'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 14   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''3,41'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''3,1'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''3,51'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''30'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 31   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''32'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 32   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''22'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 30   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''2,82'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''3,15'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''3,91'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''1,82'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''12'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 12   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''3,95'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''2,58'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''0'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 25   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 70 AND TRIBUTACAO_ENTRADA.ICMS =''12'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 12   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 70 AND TRIBUTACAO_ENTRADA.ICMS =''12'' AND TRIBUTACAO_ENTRADA.REDUCAO =''41,67'' THEN 15   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 70 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''33,3333'' THEN 16   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 70 AND TRIBUTACAO_ENTRADA.ICMS =''12'' AND TRIBUTACAO_ENTRADA.REDUCAO =''33,33'' THEN 16   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 70 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 70 AND TRIBUTACAO_ENTRADA.ICMS =''12'' AND TRIBUTACAO_ENTRADA.REDUCAO =''41,66'' THEN 15   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 70 AND TRIBUTACAO_ENTRADA.ICMS =''13,3'' AND TRIBUTACAO_ENTRADA.REDUCAO =''47,37'' THEN 56   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 70 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''33,32'' THEN 16   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 70 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''51,11'' THEN 17   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 70 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''33,34'' THEN 16   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 70 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''33,33'' THEN 16   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 70 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''61,11'' THEN 17   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 70 AND TRIBUTACAO_ENTRADA.ICMS =''25'' AND TRIBUTACAO_ENTRADA.REDUCAO =''52'' THEN 18   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 70 AND TRIBUTACAO_ENTRADA.ICMS =''13,3'' AND TRIBUTACAO_ENTRADA.REDUCAO =''9,77'' THEN 58   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 90 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''77'' THEN 22   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 90 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 22   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 90 AND TRIBUTACAO_ENTRADA.ICMS =''0'' AND TRIBUTACAO_ENTRADA.REDUCAO =''33,33'' THEN 22   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 90 AND TRIBUTACAO_ENTRADA.ICMS =''0'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 22   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 90 AND TRIBUTACAO_ENTRADA.ICMS =''3,1'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 22   ');
    SQL.Add('               ELSE 1   ');
    SQL.Add('           END AS COD_TRIB_ENTRADA,   ');
    SQL.Add('');
    SQL.Add('    COALESCE(TRIB_SAIDA.IVA, 0) AS PER_IVA,');
    SQL.Add('    ''SP'' AS DES_SIGLA');
    SQL.Add('    ');
    SQL.Add('FROM PRODUTOS');
    SQL.Add('');
    SQL.Add('LEFT JOIN PRODUTOS_ESTADO TRIB_SAIDA');
    SQL.Add('ON PRODUTOS.ID = TRIB_SAIDA.ID');
    SQL.Add('');
    SQL.Add('LEFT JOIN FIS_T_NCM NCM');
    SQL.Add('ON PRODUTOS.CLASSIFICACAO_FISCAL = NCM.ID_NCM');
    SQL.Add('AND ID_NCM > 0');
    SQL.Add('');
    SQL.Add('LEFT JOIN');
    SQL.Add('  (');
    SQL.Add('    SELECT DISTINCT');
    SQL.Add('        ITEM.COD_ITEM,');
    SQL.Add('        SUBSTR(ITEM.CST_ICMS,2,2) AS CST,');
    SQL.Add('        ITEM.ALIQ_ICMS AS ICMS,');
    SQL.Add('        COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) AS REDUCAO');
    SQL.Add('    FROM');
    SQL.Add('        FIS_T_C170 ITEM');
    SQL.Add('    INNER JOIN');
    SQL.Add('        (');
    SQL.Add('            SELECT');
    SQL.Add('                COD_ITEM,');
    SQL.Add('                MAX(ID_C170) ID_C170');
    SQL.Add('            FROM FIS_T_C170 ITEM');
    SQL.Add('            INNER JOIN FIS_T_C100 CAPA');
    SQL.Add('            ON ITEM.ID_C100 = CAPA.ID_C100');
    //SQL.Add('            AND CAPA.ID_COI_A IN (42, 44, 43, 52)');
    SQL.Add('            GROUP BY');
    SQL.Add('                COD_ITEM');
    SQL.Add('        ) AUX');
    SQL.Add('    ON ITEM.COD_ITEM = AUX.COD_ITEM');
    SQL.Add('    AND ITEM.ID_C170 = AUX.ID_C170');
    SQL.Add('  ) TRIBUTACAO_ENTRADA');
    SQL.Add('ON');
    SQL.Add('  PRODUTOS.ID = TRIBUTACAO_ENTRADA.COD_ITEM    ');
    SQL.Add('');
    SQL.Add('ORDER BY');
    SQL.Add('    NUM_NCM,');
    SQL.Add('    DES_NCM,');
    SQL.Add('    FLG_NAO_PIS_COFINS,');
    SQL.Add('    TIPO_NAO_PIS_COFINS,');
    SQL.Add('    COD_TAB_SPED,');
    SQL.Add('    COD_TRIB_SAIDA,');
    SQL.Add('    COD_TRIB_ENTRADA');


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

procedure TFrmSmKuzzi.GerarNCMUF;
var
 count : Integer;
begin
  inherited;
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT DISTINCT');
    SQL.Add('    0 AS COD_NCM,    ');
    SQL.Add('    COALESCE(NCM.DESCRICAO, ''A DEFINIR'') AS DES_NCM,');
    SQL.Add('');
    SQL.Add('    CASE ');
    SQL.Add('        WHEN CAST(COALESCE(PRODUTOS.CLASSIFICACAO_FISCAL, ''99999999'') AS INTEGER ) = 0 THEN ''99999999''');
    SQL.Add('        ELSE COALESCE(PRODUTOS.CLASSIFICACAO_FISCAL, ''99999999'')');
    SQL.Add('    END AS NUM_NCM,');
    SQL.Add('');
    SQL.Add('    CASE ');
    SQL.Add('        WHEN COALESCE(PRODUTOS.CEST, ''0000000'') = ''0000000'' THEN ''9999999'' ');
    SQL.Add('        ELSE PRODUTOS.CEST ');
    SQL.Add('    END AS NUM_CEST,');
    SQL.Add('');
    SQL.Add('    CASE PRODUTOS.MONOFASICO');
    SQL.Add('        WHEN ''I'' THEN ''S''');
    SQL.Add('        WHEN ''M'' THEN ''S''');
    SQL.Add('        WHEN ''T'' THEN ''N''');
    SQL.Add('        WHEN ''B'' THEN ''S''');
    SQL.Add('        WHEN ''O'' THEN ''S''');
    SQL.Add('        WHEN ''N'' THEN ''S''');
    SQL.Add('        WHEN ''S'' THEN ''S''');
    SQL.Add('    END AS FLG_NAO_PIS_COFINS,');
    SQL.Add('');
    SQL.Add('    CASE PRODUTOS.MONOFASICO');
    SQL.Add('        WHEN ''I'' THEN 0 -- ISENTO');
    SQL.Add('        WHEN ''M'' THEN 1 -- MONOFASICO');
    SQL.Add('        WHEN ''T'' THEN -1 -- INCIDENTE');
    SQL.Add('        WHEN ''B'' THEN 2 -- SUBSTITUICAO');
    SQL.Add('        WHEN ''O'' THEN 0 -- ALIQUOTA ZERO');
    SQL.Add('        WHEN ''N'' THEN 0 -- NÃO INCIDENTE');
    SQL.Add('        WHEN ''S'' THEN 4 -- SUSPENSO');
    SQL.Add('    END AS TIPO_NAO_PIS_COFINS,');
    SQL.Add('');
    SQL.Add('    ''999'' AS COD_TAB_SPED,');
    SQL.Add('');
    SQL.Add('           CASE   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 20 AND TRIB_SAIDA.ICMS_VENDA = 18 AND TRIB_SAIDA.REDUCAO_VENDA = ''61,11'' THEN 8   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 60 AND TRIB_SAIDA.ICMS_VENDA = 0 AND TRIB_SAIDA.REDUCAO_VENDA = ''41,67'' THEN 15   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 0 AND TRIB_SAIDA.ICMS_VENDA = 7 AND TRIB_SAIDA.REDUCAO_VENDA = 0 THEN 2   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 90 AND TRIB_SAIDA.ICMS_VENDA = 18 AND TRIB_SAIDA.REDUCAO_VENDA = 0 THEN 22   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 60 AND TRIB_SAIDA.ICMS_VENDA = 18 AND TRIB_SAIDA.REDUCAO_VENDA = ''61,11'' THEN 17   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 60 AND TRIB_SAIDA.ICMS_VENDA = 20 AND TRIB_SAIDA.REDUCAO_VENDA = 0 THEN 29   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 0 AND TRIB_SAIDA.ICMS_VENDA = 12 AND TRIB_SAIDA.REDUCAO_VENDA = 0 THEN 3   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 41 AND TRIB_SAIDA.ICMS_VENDA = 0 AND TRIB_SAIDA.REDUCAO_VENDA = 0 THEN 23   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 60 AND TRIB_SAIDA.ICMS_VENDA = 27 AND TRIB_SAIDA.REDUCAO_VENDA = 0 THEN 13   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 0 AND TRIB_SAIDA.ICMS_VENDA = 18 AND TRIB_SAIDA.REDUCAO_VENDA = 0 THEN 4   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 20 AND TRIB_SAIDA.ICMS_VENDA = 18 AND TRIB_SAIDA.REDUCAO_VENDA = ''51,11'' THEN 39   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 60 AND TRIB_SAIDA.ICMS_VENDA = 18 AND TRIB_SAIDA.REDUCAO_VENDA = 0 THEN 13   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 60 AND TRIB_SAIDA.ICMS_VENDA = 30 AND TRIB_SAIDA.REDUCAO_VENDA = 0 THEN 31   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 60 AND TRIB_SAIDA.ICMS_VENDA = 22 AND TRIB_SAIDA.REDUCAO_VENDA = 0 THEN 30   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 60 AND TRIB_SAIDA.ICMS_VENDA = 25 AND TRIB_SAIDA.REDUCAO_VENDA = 0 THEN 14   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 60 AND TRIB_SAIDA.ICMS_VENDA = 32 AND TRIB_SAIDA.REDUCAO_VENDA = 0 THEN 32   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 20 AND TRIB_SAIDA.ICMS_VENDA = 12 AND TRIB_SAIDA.REDUCAO_VENDA = ''41,67'' THEN 6   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 40 AND TRIB_SAIDA.ICMS_VENDA = 0 AND TRIB_SAIDA.REDUCAO_VENDA = 0 THEN 1   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 20 AND TRIB_SAIDA.ICMS_VENDA = 18 AND TRIB_SAIDA.REDUCAO_VENDA = ''33,33'' THEN 7   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 0 AND TRIB_SAIDA.ICMS_VENDA = 25 AND TRIB_SAIDA.REDUCAO_VENDA = 0 THEN 5   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 60 AND TRIB_SAIDA.ICMS_VENDA = 12 AND TRIB_SAIDA.REDUCAO_VENDA = 0 THEN 12   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 20 AND TRIB_SAIDA.ICMS_VENDA = 12 AND TRIB_SAIDA.REDUCAO_VENDA = ''53,33'' THEN 40   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 90 AND TRIB_SAIDA.ICMS_VENDA = 0 AND TRIB_SAIDA.REDUCAO_VENDA = 0 THEN 22   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 10 AND TRIB_SAIDA.ICMS_VENDA = 12 AND TRIB_SAIDA.REDUCAO_VENDA = 0 THEN 42   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 60 AND TRIB_SAIDA.ICMS_VENDA = 0 AND TRIB_SAIDA.REDUCAO_VENDA = 0 THEN 41   ');
    SQL.Add('               ELSE 1   ');
    SQL.Add('           END AS COD_TRIB_SAIDA,   ');
    SQL.Add('');
    SQL.Add('           CASE   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,95'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''11'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,88'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,51'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''2,85'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''7'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''4,5'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''12'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''2,56'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,85'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,93'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,82'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''2,83'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''1,36'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''2,86'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,38'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,07'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,3'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,67'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''1,54'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,48'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''1,25'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''1,86'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,87'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''2,84'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''2,38'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''2,82'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''2,48'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,9'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,92'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''2,87'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,94'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''4'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 27   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,62'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,41'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''1,44'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''4,7'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''13,3'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''5,5'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''2,58'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,1'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,91'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''2,33'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''1,83'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''2,76'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''25'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 5   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,63'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''2,49'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''0'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,45'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,65'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,15'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,97'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,58'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''1,49'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 10 AND TRIBUTACAO_ENTRADA.ICMS =''20'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 29   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 10 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 10 AND TRIBUTACAO_ENTRADA.ICMS =''25'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 14   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 10 AND TRIBUTACAO_ENTRADA.ICMS =''30'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 31   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 10 AND TRIBUTACAO_ENTRADA.ICMS =''4'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 28   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 10 AND TRIBUTACAO_ENTRADA.ICMS =''1,25'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 10 AND TRIBUTACAO_ENTRADA.ICMS =''3,3'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 10 AND TRIBUTACAO_ENTRADA.ICMS =''13,3'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 50   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 10 AND TRIBUTACAO_ENTRADA.ICMS =''0'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 41   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 10 AND TRIBUTACAO_ENTRADA.ICMS =''12'' AND TRIBUTACAO_ENTRADA.REDUCAO =''41,67'' THEN 15   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 10 AND TRIBUTACAO_ENTRADA.ICMS =''3,41'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 10 AND TRIBUTACAO_ENTRADA.ICMS =''3,07'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 10 AND TRIBUTACAO_ENTRADA.ICMS =''2,87'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 10 AND TRIBUTACAO_ENTRADA.ICMS =''3,1'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 10 AND TRIBUTACAO_ENTRADA.ICMS =''12'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 12   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 10 AND TRIBUTACAO_ENTRADA.ICMS =''1,86'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''61,11'' THEN 8   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''13,3'' AND TRIBUTACAO_ENTRADA.REDUCAO =''47,37'' THEN 57   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''25'' AND TRIBUTACAO_ENTRADA.REDUCAO =''52'' THEN 9   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 8   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''33,333'' THEN 7   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''77'' THEN 47   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''12'' AND TRIBUTACAO_ENTRADA.REDUCAO =''62,09'' THEN 36   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''51,56'' THEN 39   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''12'' AND TRIBUTACAO_ENTRADA.REDUCAO =''60'' THEN 36   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''52'' THEN 39   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''12'' AND TRIBUTACAO_ENTRADA.REDUCAO =''33,33'' THEN 7   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''61,12'' THEN 8   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''79,02'' THEN 47   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''62,09'' THEN 8   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''12'' AND TRIBUTACAO_ENTRADA.REDUCAO =''73,34'' THEN 6   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''12'' AND TRIBUTACAO_ENTRADA.REDUCAO =''82,55'' THEN 6   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''51,1'' THEN 39   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''33,3333'' THEN 7   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''48,89'' THEN 7   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''31,12'' THEN 7   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''52,21'' THEN 39   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''51,11'' THEN 39   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''78,4'' THEN 47   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''67,35'' THEN 8   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''62,94'' THEN 8   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''12'' AND TRIBUTACAO_ENTRADA.REDUCAO =''41,67'' THEN 6   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''78,39'' THEN 47   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''82,66'' THEN 47   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''33,33'' THEN 7   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''75,79'' THEN 47   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''73,97'' THEN 47   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''33,36'' THEN 7   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''13,3'' AND TRIBUTACAO_ENTRADA.REDUCAO =''47,368'' THEN 57   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''12'' AND TRIBUTACAO_ENTRADA.REDUCAO =''46,67'' THEN 6   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''56,97'' THEN 39   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''74,78'' THEN 47   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''69,14'' THEN 8   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''33,334'' THEN 7   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''12'' AND TRIBUTACAO_ENTRADA.REDUCAO =''41,66'' THEN 6   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''66,22'' THEN 8   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''66,67'' THEN 8   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''12'' AND TRIBUTACAO_ENTRADA.REDUCAO =''26,66'' THEN 37   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''12'' AND TRIBUTACAO_ENTRADA.REDUCAO =''40'' THEN 6   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''41,67'' THEN 6   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''83,88'' THEN 8   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 30 AND TRIBUTACAO_ENTRADA.ICMS =''0'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 1   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 40 AND TRIBUTACAO_ENTRADA.ICMS =''0'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 1   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 41 AND TRIBUTACAO_ENTRADA.ICMS =''0'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 23   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 51 AND TRIBUTACAO_ENTRADA.ICMS =''0'' AND TRIBUTACAO_ENTRADA.REDUCAO =''100'' THEN 20   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 51 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 20   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 51 AND TRIBUTACAO_ENTRADA.ICMS =''0'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 20   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''3,07'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''61,11'' THEN 17   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''2,56'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''3,82'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''3,3'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''33,33'' THEN 16   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''1,86'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''1,25'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''3,85'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''3,93'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''2,41'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''7'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 11   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''0'' AND TRIBUTACAO_ENTRADA.REDUCAO =''33,33'' THEN 16   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''25'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 14   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''3,41'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''3,1'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''3,51'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''30'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 31   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''32'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 32   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''22'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 30   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''2,82'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''3,15'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''3,91'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''1,82'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''12'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 12   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''3,95'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''2,58'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''0'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 25   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 70 AND TRIBUTACAO_ENTRADA.ICMS =''12'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 12   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 70 AND TRIBUTACAO_ENTRADA.ICMS =''12'' AND TRIBUTACAO_ENTRADA.REDUCAO =''41,67'' THEN 15   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 70 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''33,3333'' THEN 16   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 70 AND TRIBUTACAO_ENTRADA.ICMS =''12'' AND TRIBUTACAO_ENTRADA.REDUCAO =''33,33'' THEN 16   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 70 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 70 AND TRIBUTACAO_ENTRADA.ICMS =''12'' AND TRIBUTACAO_ENTRADA.REDUCAO =''41,66'' THEN 15   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 70 AND TRIBUTACAO_ENTRADA.ICMS =''13,3'' AND TRIBUTACAO_ENTRADA.REDUCAO =''47,37'' THEN 56   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 70 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''33,32'' THEN 16   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 70 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''51,11'' THEN 17   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 70 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''33,34'' THEN 16   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 70 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''33,33'' THEN 16   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 70 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''61,11'' THEN 17   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 70 AND TRIBUTACAO_ENTRADA.ICMS =''25'' AND TRIBUTACAO_ENTRADA.REDUCAO =''52'' THEN 18   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 70 AND TRIBUTACAO_ENTRADA.ICMS =''13,3'' AND TRIBUTACAO_ENTRADA.REDUCAO =''9,77'' THEN 58   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 90 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''77'' THEN 22   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 90 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 22   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 90 AND TRIBUTACAO_ENTRADA.ICMS =''0'' AND TRIBUTACAO_ENTRADA.REDUCAO =''33,33'' THEN 22   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 90 AND TRIBUTACAO_ENTRADA.ICMS =''0'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 22   ');
               SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 90 AND TRIBUTACAO_ENTRADA.ICMS =''3,1'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 22   ');
    SQL.Add('               ELSE 1   ');
    SQL.Add('           END AS COD_TRIB_ENTRADA,   ');
    SQL.Add('');
    SQL.Add('    COALESCE(TRIB_SAIDA.IVA, 0) AS PER_IVA,');
    SQL.Add('    ''SP'' AS DES_SIGLA');
    SQL.Add('    ');
    SQL.Add('FROM PRODUTOS');
    SQL.Add('');
    SQL.Add('LEFT JOIN PRODUTOS_ESTADO TRIB_SAIDA');
    SQL.Add('ON PRODUTOS.ID = TRIB_SAIDA.ID');
    SQL.Add('');
    SQL.Add('LEFT JOIN FIS_T_NCM NCM');
    SQL.Add('ON PRODUTOS.CLASSIFICACAO_FISCAL = NCM.ID_NCM');
    SQL.Add('AND ID_NCM > 0');
    SQL.Add('');
    SQL.Add('LEFT JOIN');
    SQL.Add('  (');
    SQL.Add('    SELECT DISTINCT');
    SQL.Add('        ITEM.COD_ITEM,');
    SQL.Add('        SUBSTR(ITEM.CST_ICMS,2,2) AS CST,');
    SQL.Add('        ITEM.ALIQ_ICMS AS ICMS,');
    SQL.Add('        COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) AS REDUCAO');
    SQL.Add('    FROM');
    SQL.Add('        FIS_T_C170 ITEM');
    SQL.Add('    INNER JOIN');
    SQL.Add('        (');
    SQL.Add('            SELECT');
    SQL.Add('                COD_ITEM,');
    SQL.Add('                MAX(ID_C170) ID_C170');
    SQL.Add('            FROM FIS_T_C170 ITEM');
    SQL.Add('            INNER JOIN FIS_T_C100 CAPA');
    SQL.Add('            ON ITEM.ID_C100 = CAPA.ID_C100');
    //SQL.Add('            AND CAPA.ID_COI_A IN (42, 44, 43, 52)');
    SQL.Add('            GROUP BY');
    SQL.Add('                COD_ITEM');
    SQL.Add('        ) AUX');
    SQL.Add('    ON ITEM.COD_ITEM = AUX.COD_ITEM');
    SQL.Add('    AND ITEM.ID_C170 = AUX.ID_C170');
    SQL.Add('  ) TRIBUTACAO_ENTRADA');
    SQL.Add('ON');
    SQL.Add('  PRODUTOS.ID = TRIBUTACAO_ENTRADA.COD_ITEM    ');
    SQL.Add('');
    SQL.Add('ORDER BY');
    SQL.Add('    NUM_NCM,');
    SQL.Add('    DES_NCM,');
    SQL.Add('    FLG_NAO_PIS_COFINS,');
    SQL.Add('    TIPO_NAO_PIS_COFINS,');
    SQL.Add('    COD_TAB_SPED,');
    SQL.Add('    COD_TRIB_SAIDA,');
    SQL.Add('    COD_TRIB_ENTRADA');


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

procedure TFrmSmKuzzi.GerarNFClientes;
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
    SQL.Add('    CAPA.ID_EMPRESA_A = 1');

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

procedure TFrmSmKuzzi.GerarNFFornec;
begin
  inherited;
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
    SQL.Add('    SUBSTR(CAPA.COD_PART,2, 60) AS COD_FORNECEDOR,');
    SQL.Add('    CAPA.NUM_DOC AS NUM_NF_FORN,');
    SQL.Add('    COALESCE(CAPA.SER, ''0'') AS NUM_SERIE_NF,');
    SQL.Add('    '''' AS NUM_SUBSERIE_NF,');
    SQL.Add('    CFOP.ID_CFOP AS CFOP,');
    SQL.Add('');
    SQL.Add('    CASE CAPA.ID_COI_A');
    SQL.Add('        WHEN 43 THEN 3');
    SQL.Add('        WHEN 52 THEN 2');
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
    SQL.Add('        WHEN 1 THEN 1');
    SQL.Add('        WHEN 46 THEN 1');
    SQL.Add('        WHEN 48 THEN 2');
    SQL.Add('        WHEN 64 THEN 6');
    SQL.Add('        WHEN 66 THEN 5');
    SQL.Add('    END AS COD_PERFIL,');
    SQL.Add('');
    SQL.Add('    0 AS VAL_DESP_ACESS,');
    SQL.Add('    ''N'' AS FLG_CANCELADO,');
    SQL.Add('    CAPA.OBSERVACAO_A AS DES_OBSERVACAO,');
    SQL.Add('    CAPA.CHV_NFE AS NUM_CHAVE_ACESSO');
    SQL.Add('FROM FIS_T_C100 CAPA');
    SQL.Add('LEFT JOIN BAS_T_COI CFOP ');
    SQL.Add('ON CAPA.ID_COI_A = CFOP.ID_COI     ');
    //SQL.Add('WHERE CAPA.ID_COI_A IN (42, 44, 43, 52) -- Transferencia 1926, 32 / ');
    SQL.Add('WHERE CAPA.STATUS_A = 2  ');
//    SQL.Add('AND CAPA.NUM_DOC =  409677 ');

    SQL.Add('AND    ');
    SQL.Add('    CAPA.DT_DOC >= :INI ');
    SQL.Add('AND    ');
    SQL.Add('    CAPA.DT_DOC <= :FIM ');
    SQL.Add('AND    ');
    SQL.Add('    CAPA.ID_EMPRESA_A = 1');
    SQL.Add('AND CFOP.ID_CFOP NOT IN (5929, 5102, 5405) ');

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

procedure TFrmSmKuzzi.GerarNFitensClientes;
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
    SQL.Add('    CAPA.ID_EMPRESA_A = 1');

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

procedure TFrmSmKuzzi.GerarNFitensFornec;
VAR
  fornecedor, nota, serie : string;
  count : integer;
begin
  inherited;

  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
    SQL.Add('    SUBSTR(CAPA.COD_PART,2, 60) AS COD_FORNECEDOR,');
    SQL.Add('    CAPA.NUM_DOC AS NUM_NF_FORN,');
    SQL.Add('    COALESCE(CAPA.SER, ''0'') AS NUM_SERIE_NF,');
    SQL.Add('    ITEM.COD_ITEM AS COD_PRODUTO,');
    SQL.Add('');
    SQL.Add('    CASE');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 40 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 0 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 1   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 7 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 2   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 90 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 12 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 3   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 12 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 3   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 4   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 90 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 4   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 4   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 30 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 4   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 25 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 5   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 12 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 41.68 THEN 6   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 12 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 33.33 THEN 6   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 12 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 41.667 THEN 6   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 12 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 8.33 THEN 6   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 12 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 41.67 THEN 6   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 10 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 12 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 41.67 THEN 6   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 70 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 12 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 41.66 THEN 6   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 12 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 6   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 12 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 46.67 THEN 6   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 12 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 41.66 THEN 6   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 12 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 40 THEN 6   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 41.67 THEN 6   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 12 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 26.66 THEN 6   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 12 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 41.47 THEN 6   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 33.333 THEN 7   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 38.89 THEN 7   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 33.3 THEN 7   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 31.12 THEN 7   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 33.3333 THEN 7   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 33.34 THEN 7   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 26.11 THEN 7   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 50.54 THEN 7   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 90 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 33.33 THEN 7   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 33.33 THEN 7   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 33.36 THEN 7   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 35.07 THEN 7   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 33.334 THEN 7   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 32.45 THEN 7   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 33.32 THEN 7   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 61.11 THEN 8   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 89.59 THEN 8   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 78.88 THEN 8   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 79.12 THEN 8   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 81.12 THEN 8   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 51.56 THEN 8   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 66.04 THEN 8   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 61.12 THEN 8   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 61.1 THEN 8   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 51.11 THEN 8   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 43.5 THEN 8   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 90 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 61.11 THEN 8   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 83.93 THEN 8   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 82.55 THEN 8   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 78.4 THEN 8   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 67.35 THEN 8   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 62.94 THEN 8   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 100 THEN 8   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 75.75 THEN 8   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 82.66 THEN 8   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 78.39 THEN 8   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 29.98 THEN 8   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 40.81 THEN 8   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 70.66 THEN 8   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 73.97 THEN 8   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 75.79 THEN 8   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 12 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 61.11 THEN 8   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 80.31 THEN 8   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 78.7 THEN 8   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 56.97 THEN 8   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 74.78 THEN 8   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 69.14 THEN 8   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 51.3 THEN 8   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 39.29 THEN 8   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 31.21 THEN 8   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 72.16 THEN 8   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 66.21 THEN 8   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 80.7 THEN 8   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 66.22 THEN 8   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 66.67 THEN 8   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 70.87 THEN 8   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 43.2 THEN 8   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 83.88 THEN 8   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 25 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 52 THEN 9   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 52 THEN 9   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 70 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 25 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 52 THEN 9   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 90 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 25 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 9   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 7 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 11   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 12.11 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 12   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 70 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 12 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 100 THEN 12   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 10 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 12 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 12   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 12 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 12   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 13   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 70 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 13   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 10 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 25 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 14   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 25 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 14   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 70 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 12 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 41.67 THEN 15   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 70 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 12 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 33.33 THEN 15   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 70 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 41.67 THEN 15   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 33.33 THEN 16   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 70 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 12 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 33 THEN 16   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 70 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 33.3333 THEN 16   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 70 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 33.34 THEN 16   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 70 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 33.32 THEN 16   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 70 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 33.33 THEN 16   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 12 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 33.33 THEN 16   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 61.11 THEN 17   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 70 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 51.11 THEN 17   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 70 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 61.11 THEN 17   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 70 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 12 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 61.11 THEN 17   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 51 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 0 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 1 THEN 20   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 51 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 0 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 100 THEN 20   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 51 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 20   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 51 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 0 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 20   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 50 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 0 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 21   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 90 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 0 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 33.33 THEN 22   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 90 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 1.58 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 22   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 90 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 0 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 22   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 13.3 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 47.37 THEN 23   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 90 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 77 THEN 23   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 77 THEN 23   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 41 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 0 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 23   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 30 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 0 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 23   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 0 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 23   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 8.4 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 24   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 10 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 8.4 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 24   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 85.21 THEN 26   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 79.02 THEN 26   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 62.09 THEN 26   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 88.29 THEN 26   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 71.86 THEN 26   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 4 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 27   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 10 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 4 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 28   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 10 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 20 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 29   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 20 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 29   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 20 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 33.33 THEN 29   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 90 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 20 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 29   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 10 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 30 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 31   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 48 THEN 33   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 48.89 THEN 33   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 72.8 THEN 33   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 64.52 THEN 33   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 13.08 THEN 33   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 11 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 35   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 12 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 62.09 THEN 36   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 12 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 82.55 THEN 36   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 12 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 60 THEN 36   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 12 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 73.34 THEN 36   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 4.5 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 38   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 59.83 THEN 39   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 58.36 THEN 39   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 52.21 THEN 39   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 51.1 THEN 39   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 0 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 61.11 THEN 41   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 70 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 0 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 41   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 0 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 33.33 THEN 41   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 10 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 0 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 41   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 10 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 32 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 41   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 60 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 41   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 0 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 41   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 70 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 62 THEN 41   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 10 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 43   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 10 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 18 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 40.6 THEN 43   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 3 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 70 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 1.86 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 3.51 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 3.88 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 3.07 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 1.86 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 2.56 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 3.95 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 70 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 12 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 3.82 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 3.38 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 3.3 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 3.9 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 10 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 2.85 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 2.56 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 3.82 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 1.25 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 3.85 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 3 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 3.85 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 1.36 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 90 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 2.38 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 3.93 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 2.83 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 2.86 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 10 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 1.25 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 10 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 2.56 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 10 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 2.84 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 10 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 3.95 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 3.07 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 3.38 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 3.48 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 3.93 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 1.54 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 3.25 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 3.67 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 2.41 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 1.86 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 2.82 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 1.25 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 3.41 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 3.51 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 2.84 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 3.1 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 32 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 30 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 22 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 10 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 3.3 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 2.38 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 3.9 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 3.92 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 90 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 1.75 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 3.87 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 3.14 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 3.6 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 2.84 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 10 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 2.07 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 2.87 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 2.82 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 2.82 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 10 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 3.41 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 2.33 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 3.41 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 10 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 1.83 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 2.58 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 10 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 1.07 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 1.44 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 3.15 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 3.62 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 3.94 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 2.33 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 3.91 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 3.1 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 10 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 2.87 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 3.91 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 10 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 3.07 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 10 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 3.1 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 2.76 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 0.36 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 1.83 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 1.82 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 2.59 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 3.95 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 3 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 2.33 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 2.58 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 3.45 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 3.63 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 3.61 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 3.65 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 2.49 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 10 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 3.25 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 10 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 1.86 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 1.25 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 3.1 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 90 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 3.1 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 10 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 3.38 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 10 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 1.33 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 3.15 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 3.58 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 1.49 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 1.58 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 1.28 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 3.97 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 60 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 2.48 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 44   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 13.3 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 49   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 10 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 13.3 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 50   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 0 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 4.7 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 0 THEN 53   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 70 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 13.3 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 47.37 THEN 56   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 20 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 13.3 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 47.368 THEN 57   ');
         SQL.Add('   WHEN SUBSTR(ITEM.CST_ICMS,2,2) = 70 AND COALESCE(ITEM.ALIQ_ICMS, 0) = 13.3 AND COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) = 9.77 THEN 58   ');
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
    SQL.Add('FROM FIS_T_C170 ITEM');
    SQL.Add('');
    SQL.Add('INNER JOIN FIS_T_C100 CAPA');
    SQL.Add('ON ITEM.ID_C100 = CAPA.ID_C100');
    SQL.Add('');
    SQL.Add('LEFT JOIN BAS_T_COI CFOP ');
    SQL.Add('ON CAPA.ID_COI_A = CFOP.ID_COI  ');
    SQL.Add('');
    SQL.Add('INNER JOIN PRODUTOS');
    SQL.Add('ON ITEM.COD_ITEM = PRODUTOS.ID');
    SQL.Add('');
    SQL.Add('LEFT JOIN PRODUTOS_FORNECEDOR');
    SQL.Add('ON SUBSTR(CAPA.COD_PART,2, 60) = PRODUTOS_FORNECEDOR.FORNECEDOR');
    SQL.Add('AND PRODUTOS_FORNECEDOR.PRODUTO = ITEM.COD_ITEM    ');
    SQL.Add('');
    //SQL.Add('WHERE CAPA.ID_COI_A IN (42, 44, 43, 52) -- Transferencia 1926, 32');
    SQL.Add('WHERE CAPA.STATUS_A = 2');
//    SQL.Add('AND CAPA.NUM_DOC =  409677 ');


    SQL.Add('AND    ');
    SQL.Add('    CAPA.DT_DOC >= :INI ');
    SQL.Add('AND    ');
    SQL.Add('    CAPA.DT_DOC <= :FIM ');
    SQL.Add('AND    ');
    SQL.Add('    CAPA.ID_EMPRESA_A = 1 ');
    SQL.Add('AND CFOP.ID_CFOP NOT IN (5929, 5102, 5405) ');

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

      if not PLUValido(Layout.FieldByName('COD_PRODUTO').AsString) then
      begin
        Layout.FieldByName('COD_PRODUTO').AsString := GerarPlu(Copy(Layout.FieldByName('COD_PRODUTO').AsString, 1, Length(Layout.FieldByName('COD_PRODUTO').AsString) - 1));
      end;

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
////
//      Layout.FieldByName('NUM_ITEM').AsInteger := count;

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
    end;
    Next;
    end;
  end;

end;

procedure TFrmSmKuzzi.GerarProdForn;
begin
  inherited;

  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT ');
    SQL.Add('    TABELA_FORNECEDOR.PRODUTO AS COD_PRODUTO,');
    SQL.Add('    TABELA_FORNECEDOR.FORNECEDOR AS COD_FORNECEDOR,');
    SQL.Add('    PRODUTOS_FORNECEDOR.REFERENCIA AS DES_REFERENCIA,');
    SQL.Add('    '''' AS NUM_CGC,');
    SQL.Add('    COALESCE(PRODUTOS_FORNECEDOR.LINHA, 0) AS COD_DIVISAO,');
    SQL.Add('    TABELA_FORNECEDOR.UNIDADE_COMPRA AS DES_UNIDADE_COMPRA,');
    SQL.Add('    TABELA_FORNECEDOR.QTDE_EMBALAGEME AS QTD_EMBALAGEM_COMPRA');
    SQL.Add('FROM TABELA_FORNECEDOR');
    SQL.Add('LEFT JOIN PRODUTOS');
    SQL.Add('ON PRODUTOS.ID = TABELA_FORNECEDOR.PRODUTO');
    SQL.Add('LEFT JOIN PRODUTOS_FORNECEDOR');
    SQL.Add('ON TABELA_FORNECEDOR.PRODUTO = PRODUTOS_FORNECEDOR.PRODUTO');
    SQL.Add('AND TABELA_FORNECEDOR.FORNECEDOR = PRODUTOS_FORNECEDOR.FORNECEDOR');




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

      if not PLUValido(Layout.FieldByName('COD_PRODUTO').AsString) then
      begin
        Layout.FieldByName('COD_PRODUTO').AsString := GerarPlu(Copy(Layout.FieldByName('COD_PRODUTO').AsString, 1, Length(Layout.FieldByName('COD_PRODUTO').AsString) - 1));
      end;

//      if QryPrincipal2.FieldByName('TIPO_CADASTRO').AsString = 'E' then
//         Layout.FieldByName('COD_PRODUTO').AsString := GerarPLU( Layout.FieldByName('COD_PRODUTO').AsString );

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

procedure TFrmSmKuzzi.GerarProdLoja;
begin
  inherited;
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
    SQL.Add('    PRODUTOS.ID AS COD_PRODUTO,');
    SQL.Add('    PRODUTOS_LOJA.CUSTO AS VAL_CUSTO_REP,');
    SQL.Add('    COALESCE(VAL_VENDA.VENDA, 0) AS VAL_VENDA,');
    SQL.Add('    COALESCE(OFERTA.PRECO, 0) AS VAL_OFERTA,');
    SQL.Add('    COALESCE(ESTOQUE.QTD_EST_VDA, 0) AS QTD_EST_VDA,');
    SQL.Add('    '''' AS TECLA_BALANCA,');
    SQL.Add('    ');
    SQL.Add('           CASE   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 20 AND TRIB_SAIDA.ICMS_VENDA = 18 AND TRIB_SAIDA.REDUCAO_VENDA = ''61,11'' THEN 8   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 60 AND TRIB_SAIDA.ICMS_VENDA = 0 AND TRIB_SAIDA.REDUCAO_VENDA = ''41,67'' THEN 15   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 0 AND TRIB_SAIDA.ICMS_VENDA = 7 AND TRIB_SAIDA.REDUCAO_VENDA = 0 THEN 2   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 90 AND TRIB_SAIDA.ICMS_VENDA = 18 AND TRIB_SAIDA.REDUCAO_VENDA = 0 THEN 22   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 60 AND TRIB_SAIDA.ICMS_VENDA = 18 AND TRIB_SAIDA.REDUCAO_VENDA = ''61,11'' THEN 17   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 60 AND TRIB_SAIDA.ICMS_VENDA = 20 AND TRIB_SAIDA.REDUCAO_VENDA = 0 THEN 29   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 0 AND TRIB_SAIDA.ICMS_VENDA = 12 AND TRIB_SAIDA.REDUCAO_VENDA = 0 THEN 3   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 41 AND TRIB_SAIDA.ICMS_VENDA = 0 AND TRIB_SAIDA.REDUCAO_VENDA = 0 THEN 23   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 60 AND TRIB_SAIDA.ICMS_VENDA = 27 AND TRIB_SAIDA.REDUCAO_VENDA = 0 THEN 13   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 0 AND TRIB_SAIDA.ICMS_VENDA = 18 AND TRIB_SAIDA.REDUCAO_VENDA = 0 THEN 4   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 20 AND TRIB_SAIDA.ICMS_VENDA = 18 AND TRIB_SAIDA.REDUCAO_VENDA = ''51,11'' THEN 39   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 60 AND TRIB_SAIDA.ICMS_VENDA = 18 AND TRIB_SAIDA.REDUCAO_VENDA = 0 THEN 13   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 60 AND TRIB_SAIDA.ICMS_VENDA = 30 AND TRIB_SAIDA.REDUCAO_VENDA = 0 THEN 31   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 60 AND TRIB_SAIDA.ICMS_VENDA = 22 AND TRIB_SAIDA.REDUCAO_VENDA = 0 THEN 30   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 60 AND TRIB_SAIDA.ICMS_VENDA = 25 AND TRIB_SAIDA.REDUCAO_VENDA = 0 THEN 14   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 60 AND TRIB_SAIDA.ICMS_VENDA = 32 AND TRIB_SAIDA.REDUCAO_VENDA = 0 THEN 32   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 20 AND TRIB_SAIDA.ICMS_VENDA = 12 AND TRIB_SAIDA.REDUCAO_VENDA = ''41,67'' THEN 6   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 40 AND TRIB_SAIDA.ICMS_VENDA = 0 AND TRIB_SAIDA.REDUCAO_VENDA = 0 THEN 1   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 20 AND TRIB_SAIDA.ICMS_VENDA = 18 AND TRIB_SAIDA.REDUCAO_VENDA = ''33,33'' THEN 7   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 0 AND TRIB_SAIDA.ICMS_VENDA = 25 AND TRIB_SAIDA.REDUCAO_VENDA = 0 THEN 5   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 60 AND TRIB_SAIDA.ICMS_VENDA = 12 AND TRIB_SAIDA.REDUCAO_VENDA = 0 THEN 12   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 20 AND TRIB_SAIDA.ICMS_VENDA = 12 AND TRIB_SAIDA.REDUCAO_VENDA = ''53,33'' THEN 40   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 90 AND TRIB_SAIDA.ICMS_VENDA = 0 AND TRIB_SAIDA.REDUCAO_VENDA = 0 THEN 22   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 10 AND TRIB_SAIDA.ICMS_VENDA = 12 AND TRIB_SAIDA.REDUCAO_VENDA = 0 THEN 42   ');
    SQL.Add('               WHEN SUBSTR(TRIB_SAIDA.ST_VENDA,2,2) = 60 AND TRIB_SAIDA.ICMS_VENDA = 0 AND TRIB_SAIDA.REDUCAO_VENDA = 0 THEN 41   ');
    SQL.Add('               ELSE 1   ');
    SQL.Add('           END AS COD_TRIBUTACAO,   ');
    SQL.Add('');
    SQL.Add('    PRODUTOS_LOJA.MARGEM_LUCRO AS VAL_MARGEM,');
    SQL.Add('    1 AS QTD_ETIQUETA,');
    SQL.Add('');
    SQL.Add('           CASE   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,95'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''11'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,88'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,51'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''2,85'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''7'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''4,5'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''12'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''2,56'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,85'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,93'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,82'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''2,83'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''1,36'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''2,86'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,38'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,07'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,3'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,67'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''1,54'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,48'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''1,25'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''1,86'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,87'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''2,84'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''2,38'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''2,82'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''2,48'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,9'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,92'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''2,87'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,94'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''4'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 27   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,62'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,41'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''1,44'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''4,7'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''13,3'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''5,5'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''2,58'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,1'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,91'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''2,33'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''1,83'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''2,76'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''25'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 5   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,63'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''2,49'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''0'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,45'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,65'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,15'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,97'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''3,58'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 0 AND TRIBUTACAO_ENTRADA.ICMS =''1,49'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 4   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 10 AND TRIBUTACAO_ENTRADA.ICMS =''20'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 29   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 10 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 10 AND TRIBUTACAO_ENTRADA.ICMS =''25'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 14   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 10 AND TRIBUTACAO_ENTRADA.ICMS =''30'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 31   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 10 AND TRIBUTACAO_ENTRADA.ICMS =''4'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 28   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 10 AND TRIBUTACAO_ENTRADA.ICMS =''1,25'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 10 AND TRIBUTACAO_ENTRADA.ICMS =''3,3'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 10 AND TRIBUTACAO_ENTRADA.ICMS =''13,3'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 50   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 10 AND TRIBUTACAO_ENTRADA.ICMS =''0'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 41   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 10 AND TRIBUTACAO_ENTRADA.ICMS =''12'' AND TRIBUTACAO_ENTRADA.REDUCAO =''41,67'' THEN 15   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 10 AND TRIBUTACAO_ENTRADA.ICMS =''3,41'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 10 AND TRIBUTACAO_ENTRADA.ICMS =''3,07'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 10 AND TRIBUTACAO_ENTRADA.ICMS =''2,87'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 10 AND TRIBUTACAO_ENTRADA.ICMS =''3,1'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 10 AND TRIBUTACAO_ENTRADA.ICMS =''12'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 12   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 10 AND TRIBUTACAO_ENTRADA.ICMS =''1,86'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''61,11'' THEN 8   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''13,3'' AND TRIBUTACAO_ENTRADA.REDUCAO =''47,37'' THEN 57   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''25'' AND TRIBUTACAO_ENTRADA.REDUCAO =''52'' THEN 9   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 8   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''33,333'' THEN 7   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''77'' THEN 47   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''12'' AND TRIBUTACAO_ENTRADA.REDUCAO =''62,09'' THEN 36   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''51,56'' THEN 39   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''12'' AND TRIBUTACAO_ENTRADA.REDUCAO =''60'' THEN 36   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''52'' THEN 39   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''12'' AND TRIBUTACAO_ENTRADA.REDUCAO =''33,33'' THEN 7   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''61,12'' THEN 8   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''79,02'' THEN 47   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''62,09'' THEN 8   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''12'' AND TRIBUTACAO_ENTRADA.REDUCAO =''73,34'' THEN 6   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''12'' AND TRIBUTACAO_ENTRADA.REDUCAO =''82,55'' THEN 6   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''51,1'' THEN 39   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''33,3333'' THEN 7   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''48,89'' THEN 7   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''31,12'' THEN 7   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''52,21'' THEN 39   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''51,11'' THEN 39   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''78,4'' THEN 47   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''67,35'' THEN 8   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''62,94'' THEN 8   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''12'' AND TRIBUTACAO_ENTRADA.REDUCAO =''41,67'' THEN 6   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''78,39'' THEN 47   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''82,66'' THEN 47   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''33,33'' THEN 7   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''75,79'' THEN 47   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''73,97'' THEN 47   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''33,36'' THEN 7   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''13,3'' AND TRIBUTACAO_ENTRADA.REDUCAO =''47,368'' THEN 57   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''12'' AND TRIBUTACAO_ENTRADA.REDUCAO =''46,67'' THEN 6   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''56,97'' THEN 39   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''74,78'' THEN 47   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''69,14'' THEN 8   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''33,334'' THEN 7   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''12'' AND TRIBUTACAO_ENTRADA.REDUCAO =''41,66'' THEN 6   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''66,22'' THEN 8   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''66,67'' THEN 8   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''12'' AND TRIBUTACAO_ENTRADA.REDUCAO =''26,66'' THEN 37   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''12'' AND TRIBUTACAO_ENTRADA.REDUCAO =''40'' THEN 6   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''41,67'' THEN 6   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 20 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''83,88'' THEN 8   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 30 AND TRIBUTACAO_ENTRADA.ICMS =''0'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 1   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 40 AND TRIBUTACAO_ENTRADA.ICMS =''0'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 1   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 41 AND TRIBUTACAO_ENTRADA.ICMS =''0'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 23   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 51 AND TRIBUTACAO_ENTRADA.ICMS =''0'' AND TRIBUTACAO_ENTRADA.REDUCAO =''100'' THEN 20   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 51 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 20   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 51 AND TRIBUTACAO_ENTRADA.ICMS =''0'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 20   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''3,07'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''61,11'' THEN 17   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''2,56'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''3,82'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''3,3'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''33,33'' THEN 16   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''1,86'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''1,25'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''3,85'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''3,93'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''2,41'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''7'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 11   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''0'' AND TRIBUTACAO_ENTRADA.REDUCAO =''33,33'' THEN 16   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''25'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 14   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''3,41'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''3,1'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''3,51'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''30'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 31   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''32'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 32   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''22'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 30   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''2,82'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''3,15'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''3,91'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''1,82'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''12'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 12   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''3,95'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''2,58'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 60 AND TRIBUTACAO_ENTRADA.ICMS =''0'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 25   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 70 AND TRIBUTACAO_ENTRADA.ICMS =''12'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 12   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 70 AND TRIBUTACAO_ENTRADA.ICMS =''12'' AND TRIBUTACAO_ENTRADA.REDUCAO =''41,67'' THEN 15   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 70 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''33,3333'' THEN 16   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 70 AND TRIBUTACAO_ENTRADA.ICMS =''12'' AND TRIBUTACAO_ENTRADA.REDUCAO =''33,33'' THEN 16   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 70 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 13   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 70 AND TRIBUTACAO_ENTRADA.ICMS =''12'' AND TRIBUTACAO_ENTRADA.REDUCAO =''41,66'' THEN 15   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 70 AND TRIBUTACAO_ENTRADA.ICMS =''13,3'' AND TRIBUTACAO_ENTRADA.REDUCAO =''47,37'' THEN 56   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 70 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''33,32'' THEN 16   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 70 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''51,11'' THEN 17   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 70 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''33,34'' THEN 16   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 70 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''33,33'' THEN 16   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 70 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''61,11'' THEN 17   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 70 AND TRIBUTACAO_ENTRADA.ICMS =''25'' AND TRIBUTACAO_ENTRADA.REDUCAO =''52'' THEN 18   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 70 AND TRIBUTACAO_ENTRADA.ICMS =''13,3'' AND TRIBUTACAO_ENTRADA.REDUCAO =''9,77'' THEN 58   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 90 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''77'' THEN 22   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 90 AND TRIBUTACAO_ENTRADA.ICMS =''18'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 22   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 90 AND TRIBUTACAO_ENTRADA.ICMS =''0'' AND TRIBUTACAO_ENTRADA.REDUCAO =''33,33'' THEN 22   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 90 AND TRIBUTACAO_ENTRADA.ICMS =''0'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 22   ');
                           SQL.Add('   WHEN TRIBUTACAO_ENTRADA.CST = 90 AND TRIBUTACAO_ENTRADA.ICMS =''3,1'' AND TRIBUTACAO_ENTRADA.REDUCAO =''0'' THEN 22   ');
    SQL.Add('               ELSE 1   ');
    SQL.Add('           END AS COD_TRIB_ENTRADA,   ');
    SQL.Add('');
    SQL.Add('    CASE');
    SQL.Add('        WHEN PRODUTOS.STATUS = 0 THEN ''N''');
    SQL.Add('        WHEN PRODUTOS.STATUS = 1 THEN ''N''');
    SQL.Add('        ELSE ''S''');
    SQL.Add('    END AS FLG_INATIVO,');
    SQL.Add('');
    SQL.Add('    PRODUTOS.ID AS COD_PRODUTO_ANT,');
    SQL.Add('');
    SQL.Add('    CASE ');
    SQL.Add('        WHEN CAST(COALESCE(PRODUTOS.CLASSIFICACAO_FISCAL, ''99999999'') AS INTEGER ) = 0 THEN ''99999999''');
    SQL.Add('        ELSE COALESCE(PRODUTOS.CLASSIFICACAO_FISCAL, ''99999999'')');
    SQL.Add('    END AS NUM_NCM,');
    SQL.Add('');
    SQL.Add('    0 AS TIPO_NCM,');
    SQL.Add('    0 AS VAL_VENDA_2,');
    SQL.Add('    OFERTA.DTA_VALIDA_OFERTA AS DTA_VALIDA_OFERTA,');
    SQL.Add('    COALESCE(ESTOQUE.QTD_EST_MINIMO, 0) AS QTD_EST_MINIMO,');
    SQL.Add('    NULL AS COD_VASILHAME,');
    SQL.Add('    CASE PRODUTOS.STATUS');
    SQL.Add('        WHEN 1 THEN ''S'' ');
    SQL.Add('        ELSE ''N''');
    SQL.Add('    END AS FORA_LINHA,');
    SQL.Add('    0 AS QTD_PRECO_DIF,');
    SQL.Add('    0 AS VAL_FORCA_VDA,');
    SQL.Add('    CASE WHEN COALESCE(PRODUTOS.CEST, ''0000000'') = ''0000000'' THEN ''9999999'' ELSE PRODUTOS.CEST END AS NUM_CEST,');
//    SQL.Add('    ''P'' AS TIPO_CADASTRO,');
    SQL.Add('    COALESCE(TRIB_SAIDA.IVA, 0) AS PER_IVA,');
    SQL.Add('    0 AS PER_FCP_ST,');
//    SQL.Add('    0 AS COD_BENEF_FISCAL');
    SQL.Add('    0 AS PER_FIDELIDADE,  ');
    SQL.Add('           CASE   ');
    SQL.Add('               WHEN RECEITA_COD.COD_RECEITA IS NULL THEN 999   ');
    SQL.Add('               ELSE RECEITA_COD.COD_RECEITA   ');
    SQL.Add('           END AS COD_INFO_RECEITA   ');
//    SQL.Add('    -- COALESCE(VAL_VENDA_3.VENDA, 0) AS VAL_VENDA_3,');
//    SQL.Add('    -- COALESCE(TRIBUTACAO_SAIDA.IVA, 0) AS PER_IVA,');
//    SQL.Add('    -- COALESCE(TRIBUTACAO_SAIDA.FCP, 0) AS PER_FCP_ST,');
//    SQL.Add('    -- CASE WHEN TRIBUTACAO_SAIDA.CREDITO_ICMS = ''T'' THEN ''S'' ELSE ''N'' END AS CREDITO_VEDADO');
    SQL.Add('FROM PRODUTOS');
    SQL.Add('');
    SQL.Add('LEFT JOIN PRODUTOS_PRECOS VAL_VENDA -- VALOR DE VENDA ');
    SQL.Add('ON PRODUTOS.ID = VAL_VENDA.PRODUTO');
    SQL.Add('AND VAL_VENDA.POLITICA = 1 --CODIGO DA LOJA');
    SQL.Add('AND VAL_VENDA.ID = 1');
    SQL.Add('');
    SQL.Add('-- LEFT JOIN PRODUTOS_PRECOS VAL_VENDA_3 -- VALOR DE VENDA');
    SQL.Add('-- ON PRODUTOS.ID = VAL_VENDA_3.PRODUTO');
    SQL.Add('-- AND VAL_VENDA_3.POLITICA = 1 --CODIGO DA LOJA');
    SQL.Add('-- AND VAL_VENDA_3.ID = 3');
    SQL.Add('');
    SQL.Add('');
    SQL.Add('LEFT JOIN PRODUTOS_LOJA');
    SQL.Add('ON PRODUTOS.ID = PRODUTOS_LOJA.ID');
    SQL.Add('AND PRODUTOS_LOJA.POLITICA = 1 --CODIGO DA LOJA   ,');
    SQL.Add('');
    SQL.Add('LEFT JOIN PRODUTOS_ESTADO TRIB_SAIDA');
    SQL.Add('ON PRODUTOS.ID = TRIB_SAIDA.ID');
    SQL.Add('');
    SQL.Add('LEFT JOIN');
    SQL.Add('    (        ');
    SQL.Add('        SELECT');
    SQL.Add('            PRODUTO,');
    SQL.Add('            SUM(ESTOQUE_ATUAL) AS QTD_EST_VDA,');
    SQL.Add('            SUM(ESTOQUE_MINIMO) AS QTD_EST_MINIMO');
    SQL.Add('        FROM');
    SQL.Add('            PRODUTOS_ESTOQUES     ');
    SQL.Add('        WHERE');
    SQL.Add('            ESTOQUE = 1');
    SQL.Add('        GROUP BY');
    SQL.Add('            PRODUTO ');
    SQL.Add('    ) ESTOQUE');
    SQL.Add('ON');
    SQL.Add('    PRODUTOS.ID = ESTOQUE.PRODUTO');
    SQL.Add('');
    SQL.Add('LEFT JOIN');
    SQL.Add('    (');
    SQL.Add('        SELECT');
    SQL.Add('            PRODUTO,');
    SQL.Add('            FIM AS DTA_VALIDA_OFERTA,');
    SQL.Add('            PRECO');
    SQL.Add('        FROM PROGRAMACAO_PRECO');
    SQL.Add('        WHERE FIM >= SYSDATE');
    SQL.Add('    ) OFERTA    ');
    SQL.Add('ON');
    SQL.Add('    PRODUTOS.ID = OFERTA.PRODUTO');
    SQL.Add('');
    SQL.Add('LEFT JOIN');
    SQL.Add('  (');
    SQL.Add('    SELECT DISTINCT');
    SQL.Add('        ITEM.COD_ITEM,');
    SQL.Add('        SUBSTR(ITEM.CST_ICMS,2,2) AS CST,');
    SQL.Add('        ITEM.ALIQ_ICMS AS ICMS,');
    SQL.Add('        COALESCE(ITEM.VALOR_REDUCAO_BASE_A, 0) AS REDUCAO');
    SQL.Add('    FROM');
    SQL.Add('        FIS_T_C170 ITEM');
    SQL.Add('    INNER JOIN');
    SQL.Add('        (');
    SQL.Add('            SELECT');
    SQL.Add('                COD_ITEM,');
    SQL.Add('                MAX(ID_C170) ID_C170');
    SQL.Add('            FROM FIS_T_C170 ITEM');
    SQL.Add('            INNER JOIN FIS_T_C100 CAPA');
    SQL.Add('            ON ITEM.ID_C100 = CAPA.ID_C100');
    //SQL.Add('            AND CAPA.ID_COI_A IN (42, 44, 43, 52)');
    SQL.Add('            GROUP BY');
    SQL.Add('                COD_ITEM');
    SQL.Add('        ) AUX');
    SQL.Add('    ON ITEM.COD_ITEM = AUX.COD_ITEM');
    SQL.Add('    AND ITEM.ID_C170 = AUX.ID_C170');
    SQL.Add('  ) TRIBUTACAO_ENTRADA');
    SQL.Add('ON');
    SQL.Add('  PRODUTOS.ID = TRIBUTACAO_ENTRADA.COD_ITEM    ');
    SQL.Add('       LEFT JOIN (   ');
    SQL.Add('            SELECT   ');
    SQL.Add('                 COALESCE(PRODUTOS.ID, 999) AS COD_RECEITA   ');
    SQL.Add('            FROM   ');
    SQL.Add('                 PRODUTOS   ');
    SQL.Add('            WHERE PRODUTOS.RECEITA IS NOT NULL   ');
    SQL.Add('       ) RECEITA_COD   ');
    SQL.Add('       ON PRODUTOS.ID = RECEITA_COD.COD_RECEITA   ');
//    SQL.Add('AND PRODUTOS.STATUS <> 3');

    Open;
    First;
    NumLinha := 0;

//    AssignFile(Arquivo, EdtCamArquivo.Text + '\SCRIPT_ATUALIZA_VEDADO_PRECO3.SQL');
//    Rewrite(Arquivo);

    while not Eof do
    begin
    try
      if Cancelar then
      Break;
      Inc(NumLinha);
      Layout.SetValues(QryPrincipal2, NumLinha, RecordCount);

      if not PLUValido(Layout.FieldByName('COD_PRODUTO').AsString) then
      begin
        Layout.FieldByName('COD_PRODUTO').AsString := GerarPlu(Copy(Layout.FieldByName('COD_PRODUTO').AsString, 1, Length(Layout.FieldByName('COD_PRODUTO').AsString) - 1));
      end;

//      if Layout.FieldByName('COD_INFO_RECEITA').AsInteger <> 0 then
//      begin
//        if not PLUValido(Layout.FieldByName('COD_PRODUTO').AsString) then
//          Layout.FieldByName('COD_INFO_RECEITA').AsString := GerarPlu(Copy(Layout.FieldByName('COD_INFO_RECEITA').AsString, 1, Length(Layout.FieldByName('COD_INFO_RECEITA').AsString) - 1));
//      end;

//      if QryPrincipal2.FieldByName('TIPO_CADASTRO').AsString = 'E' then begin
//         Layout.FieldByName('COD_PRODUTO').AsString := GerarPLU( Layout.FieldByName('COD_PRODUTO').AsString );
//         Layout.FieldByName('COD_PRODUTO_ANT').AsString := GerarPLU(QryPrincipal2.FieldByName('COD_PRODUTO_ANT').AsString);
//      end;

//      if( QryPrincipal2.FieldByName('VAL_VENDA_3').AsFloat > 0 ) or ( QryPrincipal2.FieldByName('CREDITO_VEDADO').AsString = 'S' ) then
//         Writeln(Arquivo, 'UPDATE TAB_PRODUTO_LOJA SET VAL_VENDA_3 = '+ StringReplace(QryPrincipal2.FieldByName('VAL_VENDA_3').AsString, ',', '.', [rfReplaceAll]) +', FLG_VEDA_CRED = '''+ QryPrincipal2.FieldByName('CREDITO_VEDADO').AsString +''' WHERE COD_PRODUTO = '+ Layout.FieldByName('COD_PRODUTO').AsString +' AND COD_LOJA = '+ CbxLoja.Text +'; ');

//      Layout.FieldByName('COD_PRODUTO').AsString := GerarPLU(QryPrincipal2.FieldByName('COD_PRODUTO').AsString);
//      Layout.FieldByName('COD_PRODUTO_ANT').AsString := GerarPLU(QryPrincipal2.FieldByName('COD_PRODUTO_ANT').AsString);

      Layout.FieldByName('DTA_VALIDA_OFERTA').AsDateTime := FieldByName('DTA_VALIDA_OFERTA').AsDateTime;

      Layout.WriteLine;
    except
      On E: Exception do
      FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
    end;
    Next;
    end;
//    Writeln(Arquivo, 'COMMIT WORK;');
//    Close;
  end;
//  CloseFile(Arquivo);
end;

procedure TFrmSmKuzzi.GerarProdSimilar;
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
