unit UFrmSmBomPrecoOurinhosGestora;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes,
  Vcl.Graphics, ComObj,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, UFrmModelo, Data.DBXOracle, Data.DB,
  Data.SqlExpr, Vcl.Menus, Vcl.StdCtrls, Vcl.Buttons, Vcl.ComCtrls,
  Vcl.Imaging.pngimage, Vcl.ExtCtrls, Data.DBXFirebird, Data.Win.ADODB,
  Vcl.Grids, Vcl.DBGrids, Datasnap.Provider, Datasnap.DBClient,
  // dxGDIPlusClasses,
  Math, ShellAPI;

type
  TFrmSmBomPrecoOurinhosGestora = class(TFrmModeloSis)
    CbxLoja: TComboBox;
    lblLoja: TLabel;
    ADOSQLServer: TADOConnection;
    QryPrincipal2: TADOQuery;
    QryAux: TADOQuery;
    Memo1: TMemo;
    CkbScript: TCheckBox;
    BtnCopiar: TButton;
    AbaUnificacao: TTabSheet;
    CkbUnificacao: TCheckBox;
    btnGeraValVenda: TButton;
    lblMatriz: TLabel;
    EdtMatriz: TEdit;
    ShpMatriz: TShape;
    ShpLoja: TShape;
    btnGeraCustoRep: TButton;
    btnGeraEstoque: TButton;
    btnGeraAssociado: TButton;
    btnGeraPromocao: TButton;
    btnGeraPlanilha: TButton;
    btnImportaPlanilha: TButton;
    Button1: TButton;
    BntAbrirPlanilha: TSpeedButton;
    BntAbrirPlanilhaImp: TSpeedButton;
    Label11: TLabel;
    LblBuild: TLabel;
    btnGeraObsForn: TButton;
    GroupBox1: TGroupBox;
    btnGeraClassif: TButton;
    procedure BtnGerarClick(Sender: TObject);
    // procedure CkbProdLojaClick(Sender: TObject);
    procedure QryPrincipal2AfterOpen(DataSet: TDataSet);
    procedure CkbScriptClick(Sender: TObject);
    procedure BtnCopiarClick(Sender: TObject);
    procedure CkbUnificacaoClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure EdtMatrizChange(Sender: TObject);
    procedure EdtMatrizMouseMove(Sender: TObject; Shift: TShiftState;
      X, Y: Integer);
    procedure AbaUnificacaoMouseEnter(Sender: TObject);
    procedure btnGeraPlanilhaClick(Sender: TObject);
    procedure btnImportaPlanilhaClick(Sender: TObject);
    procedure btnGeraValVendaClick(Sender: TObject);
    procedure btnGeraCustoRepClick(Sender: TObject);
    procedure btnGeraEstoqueClick(Sender: TObject);
    procedure BntAbrirPlanilhaClick(Sender: TObject);
    procedure BntAbrirPlanilhaImpClick(Sender: TObject);
    procedure FormMouseEnter(Sender: TObject);
    procedure CbxLojaChange(Sender: TObject);
    procedure btnGeraObsFornClick(Sender: TObject);
    procedure btnGeraClassifClick(Sender: TObject);

  private

    { Private declarations }
  public
    { Public declarations }
    procedure GerarCliente; Override; (* OK *)
    procedure GerarCondPagCli; Override; (* OK *)

    procedure GerarFornecedor; Override; (* OK *)
    procedure GerarCondPagForn; Override; (* OK *)

    procedure GerarSecao; Override; (* OK *) (* UNIFICADO *)
    procedure GerarGrupo; Override; (* OK *) (* UNIFICADO *)
    procedure GerarSubGrupo; Override; (* OK *) (* UNIFICADO *)

    procedure GerarProdSimilar; Override; (* OK *)

    procedure GerarProduto; Override; (* OK *) (* UNIFICADO *)
    procedure GerarReceitas; Override;
    procedure GerarInfoNutricionais; Override;

    procedure GerarCodigoBarras; Override; (* OK *) (* UNIFICADO *)
    procedure GerarProdLoja; Override; (* OK *) (* UNIFICADO *)

    procedure GerarNCM; Override; (* OK *)
    procedure GerarNCMUF; Override; (* OK *)
    procedure GerarCest; Override; (* OK *)

    procedure GerarProdForn; Override; (* OK *)

    procedure GerarNFFornec; Override; (* OK *)
    procedure GerarNFitensFornec; Override; (* OK *)

    procedure GerarVenda; Override; (* OK *)

    procedure GerarFinanceiro(Tipo, Situacao: Integer); Override; (* OK *)
    procedure GerarFinanceiroReceber(Aberto: String); Override; (* OK *)
    procedure GerarFinanceiroPagar(Aberto: String); Override; (* OK *)

    procedure GerarAssociado;
    procedure GerarValorVenda;
    procedure GeraCustoRep;
    procedure GeraEstoqueVenda;

    procedure GeraObsForn;
    procedure GeraClassif;

    // Novas Implementações - BETA
    procedure validaMatriz;
    procedure validaFiliais;

    procedure geraFiliais;

    procedure criaDiretorioUnic;
    procedure criaDiretorioScriptUpdate;
    procedure criaDiretorioPlanilha;
    procedure dataUltimaBuild;

    procedure criaCsv;
    procedure expCsv;
    procedure impCsv;

  end;

var
  FrmSmBomPrecoOurinhosGestora: TFrmSmBomPrecoOurinhosGestora;
  NumLinha: Integer;
  Arquivo: TextFile;

  FlgAtualizaValVenda: Boolean = False;
  FlgAtualizaCustoRep: Boolean = False;
  FlgAtualizaEstoque: Boolean = False;

  FlgAtualizaObsForn: Boolean = False;
  FlgAtualizaClassif: Boolean = False;
  FlgValidaMatriz: Boolean = False;
  FlgValidaFiliais: Boolean = False;

implementation

{$R *.dfm}

uses xProc, UUtilidades, UProgresso;



procedure TFrmSmBomPrecoOurinhosGestora.GerarCliente;
var
  // Obs    : TStringList;
  QryTel: TADOQuery;
begin
  inherited;
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

     SQL.Add('      SELECT      ');
     SQL.Add('          	CLI_CODIGO AS COD_CLIENTE,      ');
     SQL.Add('          	CASE      ');
     SQL.Add('          		WHEN CLI_NOME IS NULL THEN ''Á DEFINIR''      ');
     SQL.Add('          		ELSE CLI_NOME      ');
     SQL.Add('          	END AS DES_CLIENTE,      ');
     SQL.Add('          	CLI_CPFCGC AS NUM_CGC,      ');
     SQL.Add('          	CASE       ');
     SQL.Add('          		WHEN CLI_TIPO = ''J'' THEN        ');
     SQL.Add('          			CASE       ');
     SQL.Add('          				WHEN CLI_RGINS IS NULL THEN ''ISENTO''      ');
     SQL.Add('          				WHEN CLI_RGINS = ''ISENTA'' THEN ''ISENTO''      ');
     SQL.Add('          				WHEN CLI_RGINS = ''000000'' THEN ''ISENTO''      ');
     SQL.Add('          				WHEN CLI_RGINS = ''0000000000'' THEN ''ISENTO''      ');
     SQL.Add('          				ELSE CLI_RGINS      ');
     SQL.Add('          			END      ');
     SQL.Add('          		ELSE ''''      ');
     SQL.Add('          	END AS NUM_INSC_EST,      ');
     SQL.Add('          	CASE       ');
     SQL.Add('          		WHEN CLI_ENDERECO IS NULL THEN  ''A DEFINIR''      ');
     SQL.Add('          		WHEN CLI_ENDERECO = ''''  THEN ''A DEFINIR''      ');
     SQL.Add('          		ELSE CLI_ENDERECO       ');
     SQL.Add('          	END AS DES_ENDERECO,      ');
     SQL.Add('          	CASE       ');
     SQL.Add('          		WHEN CLI_BAIRRO IS NULL THEN ''A DEFINIR''      ');
     SQL.Add('          		WHEN CLI_BAIRRO = '''' THEN ''A DEFINIR''      ');
     SQL.Add('          		ELSE CLI_BAIRRO      ');
     SQL.Add('          	END AS DES_BAIRRO,      ');
     SQL.Add('          	CLI_CIDADE AS DES_CIDADE,      ');
     SQL.Add('          	CLI_ESTADO AS DES_SIGLA,      ');
     SQL.Add('          	CASE       ');
     SQL.Add('          		WHEN CLI_CEP IS NULL THEN ''17400000''      ');
     SQL.Add('          		WHEN CLI_CEP = '''' THEN ''17400000''      ');
     SQL.Add('          		ELSE CLI_CEP       ');
     SQL.Add('          	END AS NUM_CEP,      ');
     SQL.Add('          	REPLACE ((CONCAT( CLI_DDD1, CLI_TELEFONE1)),''-'', '''')  AS NUM_FONE,      ');
     SQL.Add('          	'''' AS NUM_FAX,      ');
     SQL.Add('          	CASE      ');
     SQL.Add('          		WHEN CLI_CONTATO = '''' THEN CLI_NOME      ');
     SQL.Add('          		WHEN CLI_CONTATO IS NULL THEN COALESCE (CLI_NOME, ''ADEFINIR'')      ');
     SQL.Add('          		ELSE CLI_CONTATO      ');
     SQL.Add('          	END AS DES_CONTATO,      ');
     SQL.Add('          	0 AS FLG_SEXO,      ');
     SQL.Add('          	CLI_LIMITE AS VAL_LIMITE_CRETID,      ');
     SQL.Add('          	COALESCE (CLI_LIMITE, 0) AS VAL_LIMITE_CONV,      ');
     SQL.Add('          	0 AS VAL_DEBITO,      ');
     SQL.Add('          	CLI_RENDA AS VAL_RENDA,      ');
     SQL.Add('          	0 AS COD_CONVENIO,      ');
     SQL.Add('          	0 AS COD_STATUS_PDV,      ');
     SQL.Add('          	CASE         ');
     SQL.Add('                  WHEN CLI_TIPO = ''J'' THEN ''S''         ');
     SQL.Add('                  ELSE ''N''         ');
     SQL.Add('          	END AS FLG_EMPRESA,      ');
     SQL.Add('          	''N'' AS FLG_CONVENIO,         ');
     SQL.Add('              ''N'' AS MICRO_EMPRESA,      ');
     SQL.Add('          	COALESCE (CONVERT(CHAR, CLI_DATA_CADASTRO, 103), ''01/01/1899'') AS DTA_CADASTRO,      ');
     SQL.Add('          	CASE       ');
     SQL.Add('          		WHEN CLI_ENDNRO IS NULL THEN ''S/N''      ');
     SQL.Add('          		WHEN CLI_ENDNRO = '''' THEN ''S/N''      ');
     SQL.Add('          		ELSE REPLACE (CLI_ENDNRO, ''-'', '''')       ');
     SQL.Add('          	END AS NUM_ENDERECO,      ');
     SQL.Add('          	CASE       ');
     SQL.Add('          		WHEN CLI_TIPO = ''F'' THEN      ');
     SQL.Add('          			CASE      ');
     SQL.Add('          				WHEN CLI_RGINS IS NULL THEN ''''      ');
     SQL.Add('          				ELSE CLI_RGINS      ');
     SQL.Add('          			END      ');
     SQL.Add('          		ELSE ''''	      ');
     SQL.Add('          	END NUM_RG,      ');
     SQL.Add('          	CASE	      ');
     SQL.Add('          		WHEN CLI_EST_CIVIL IS NULL THEN 0      ');
     SQL.Add('          		WHEN CLI_EST_CIVIL = '''' THEN 0      ');
     SQL.Add('          		WHEN CLI_EST_CIVIL = ''SOLTEIRO'' THEN 0      ');
     SQL.Add('          		WHEN CLI_EST_CIVIL = ''CASADO'' THEN 1      ');
     SQL.Add('          		WHEN CLI_EST_CIVIL = ''CASADA'' THEN 1      ');
     SQL.Add('          		ELSE 0       ');
     SQL.Add('          	END FLG_EST_CIVIL,      ');
     SQL.Add('            COALESCE (CONCAT (CLI_DDDCELULAR, CLI_CELULAR), '''') AS NUM_CELULAR,       ');
     SQL.Add('          	''16/08/2022'' AS DTA_ALTERACAO,      ');
     SQL.Add('          	COALESCE (CLI_OBS1, '''') AS DES_OBSERVACAO,      ');
     SQL.Add('          	''A DEFINIR'' AS DES_COMPLEMENTO,      ');
     SQL.Add('          	CASE       ');
     SQL.Add('          		WHEN ((CLI_EMAIL2 = '''') OR (CLI_EMAIL2 IS NULL)) THEN       ');
     SQL.Add('          			CASE       ');
     SQL.Add('          				WHEN CLI_EMAIL IS NULL THEN ''''      ');
     SQL.Add('          				ELSE UPPER (CLI_EMAIL)      ');
     SQL.Add('          			END      ');
     SQL.Add('          		WHEN ((CLI_EMAIL2 IS NOT NULL) OR (CLI_EMAIL2 <> '' '' )) THEN UPPER ((CLI_EMAIL + ''; '' + CLI_EMAIL2))      ');
     SQL.Add('          	END DES_EMAIL,      ');
     SQL.Add('          	CASE	      ');
     SQL.Add('          		WHEN ((CLI_TIPO = ''J'') AND (CLI_FANTASIA IS NOT NULL)) THEN CLI_FANTASIA      ');
     SQL.Add('          		ELSE ''''      ');
     SQL.Add('          	END AS DES_FANTASIA,      ');
     SQL.Add('          	COALESCE (CONVERT(VARCHAR, CLI_NASCIMENTO , 103), ''01/01/1899'') AS DTA_NASCIMENTO,      ');
     SQL.Add('          	COALESCE (CLI_PAI, '''') AS DES_PAI,      ');
     SQL.Add('          	COALESCE (CLI_MAE, '''') AS DES_MAE,      ');
     SQL.Add('          	COALESCE (CLI_CONJUGUE , '''') AS DES_CONJUGE,      ');
     SQL.Add('          	COALESCE (CLI_CON_CPF , '''') AS NUM_CPF_CONJUGE,      ');
     SQL.Add('          	0 AS VAL_DEB_CONV,      ');
     SQL.Add('          	CASE       ');
     SQL.Add('          		WHEN CLI_STATUS = ''I'' THEN ''S''      ');
     SQL.Add('          		ELSE ''N''      ');
     SQL.Add('          	END AS INATIVO,      ');
     SQL.Add('          	'' '' AS DES_MATRICULA,         ');
     SQL.Add('              ''N'' AS NUM_CGC_ASSOCIADO,         ');
     SQL.Add('              ''N'' AS FLG_PROD_RURAL,      ');
     SQL.Add('          	CASE       ');
     SQL.Add('          		WHEN CLI_BLOQUEADO = ''S'' THEN 1      ');
     SQL.Add('          		ELSE 0      ');
     SQL.Add('          	END AS COD_STATUS_PDV_CONV,      ');
     SQL.Add('          	''S'' AS FLG_ENVIA_CODIGO,         ');
     SQL.Add('              COALESCE (CONVERT (CHAR, CLI_CON_NASCIMENTO , 103), ''01/01/1899'') AS DTA_NASC_CONJUGE,         ');
     SQL.Add('           CASE    ');
     SQL.Add('   		   WHEN GRU_CODIGO IS NULL OR GRU_CODIGO = 0 THEN 0   ');
     SQL.Add('   		   ELSE GRU_CODIGO + 1000   ');
     SQL.Add('   		   END AS COD_CLASSIF      ');
     SQL.Add('                ');
     SQL.Add('          FROM dbo.CLIENTES      ');
     SQL.Add('          ORDER BY      ');
     SQL.Add('          	CLI_CODIGO;   ');



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

        Layout.FieldByName('DES_CLIENTE').AsString :=
          StrRemPont(Layout.FieldByName('DES_CLIENTE').AsString);

        Layout.FieldByName('NUM_CGC').AsString :=
          StrRetNums(Layout.FieldByName('NUM_CGC').AsString);

        if (StrRetNums(Layout.FieldByName('NUM_ENDERECO').AsString) = '') then
          Layout.FieldByName('NUM_ENDERECO').AsString := 'S/N'
        else if strtoint(StrRetNums(Layout.FieldByName('NUM_ENDERECO')
          .AsString)) = 0 then
          Layout.FieldByName('NUM_ENDERECO').AsString := 'S/N'
        else
          Layout.FieldByName('NUM_ENDERECO').AsString :=
            StrRetNums(Layout.FieldByName('NUM_ENDERECO').AsString);

        if StrRetNums(Layout.FieldByName('NUM_RG').AsString) = '' then
          Layout.FieldByName('NUM_RG').AsString := ''
        else
          Layout.FieldByName('NUM_RG').AsString :=
            StrRetNums(Layout.FieldByName('NUM_RG').AsString);

        Layout.FieldByName('NUM_CEP').AsString :=
          StrRetNums(Layout.FieldByName('NUM_CEP').AsString);

        //if QryPrincipal2.FieldByName('DTA_CADASTRO').AsString <> '' then
        //  Layout.FieldByName('DTA_CADASTRO').AsString :=
         //   FORMATDATETIME('dd/mm/yyyy',
         //   QryPrincipal2.FieldByName('DTA_CADASTRO').AsDateTime);

        //if QryPrincipal2.FieldByName('DTA_NASCIMENTO').AsString <> '' then
         /// Layout.FieldByName('DTA_NASCIMENTO').AsString :=
           /// FORMATDATETIME('dd/mm/yyyy', QryPrincipal2.FieldByName('DTA_NASCIMENTO').AsDateTime);

        Layout.FieldByName('NUM_FONE').AsString :=
          StrRetNums(FieldByName('NUM_FONE').AsString);

        Layout.FieldByName('NUM_FAX').AsString :=
          StrRetNums(FieldByName('NUM_FAX').AsString);

        Layout.FieldByName('NUM_CELULAR').AsString := StrRetNums(FieldByName('NUM_CELULAR').AsString);

        if Length(Layout.FieldByName('NUM_CGC').AsString) > 11 then
        begin
          if not ValidaCGC(Layout.FieldByName('NUM_CGC').AsString) then
            Layout.FieldByName('NUM_CGC').AsString := '';
        end
        else if not ValidaCpf(Layout.FieldByName('NUM_CGC').AsString) then
          Layout.FieldByName('NUM_CGC').AsString := '';

        Layout.FieldByName('DES_EMAIL').AsString :=
          StrReplace(StrLBReplace(FieldByName('DES_EMAIL').AsString), '\n', '');
        Layout.FieldByName('DES_ENDERECO').AsString :=
          StrReplace(StrLBReplace(FieldByName('DES_ENDERECO').AsString),
          '\n', '');
        Layout.FieldByName('DES_OBSERVACAO').AsString :=
          StrReplace(StrLBReplace(FieldByName('DES_OBSERVACAO').AsString),
          '\n', '');

        if Layout.FieldByName('DES_EMAIL').AsString = ';' then
          Layout.FieldByName('DES_EMAIL').AsString := '';

        if Layout.FieldByName('FLG_EMPRESA').AsString = 'N' then
          Layout.FieldByName('NUM_INSC_EST').AsString := '';

        Layout.WriteLine;
      except
        On E: Exception do
          FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
      end;
      Next;
    end;
  end;
end;

procedure TFrmSmBomPrecoOurinhosGestora.GerarCondPagCli;
begin
  inherited;

  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('   SELECT    ');
    SQL.Add('   		CLI_CODIGO AS COD_CLIENTE,   ');
    SQL.Add('   		30 AS NUM_CONDICAO,   ');
    SQL.Add('   		2 AS COD_CONDICAO,   ');
    SQL.Add('   		1 AS COD_ENTIDADE   ');
    SQL.Add('   FROM dbo.CLIENTES    ');
    SQL.Add('   WHERE CLI_NOME IS NOT NULL   ');
    SQL.Add('   ORDER BY COD_CLIENTE ASC;   ');

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


procedure TFrmSmBomPrecoOurinhosGestora.GerarFornecedor;
var
  inscEst: string;
begin
  inherited;

  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;
       SQL.Add('    SELECT            ');
       SQL.Add('       FORNECEDOR.FOR_CODIGO AS COD_FORNECEDOR,            ');
       SQL.Add('       FORNECEDOR.FOR_RAZAO AS DES_FORNECEDOR,            ');
       SQL.Add('       CASE             ');
       SQL.Add('           WHEN FOR_FANTASIA IS NULL THEN FORNECEDOR.FOR_RAZAO            ');
       SQL.Add('           WHEN FOR_FANTASIA = '''' THEN FORNECEDOR.FOR_RAZAO            ');
       SQL.Add('           ELSE FOR_FANTASIA            ');
       SQL.Add('       END AS DES_FANTASIA,            ');
       SQL.Add('       FOR_CGC AS NUM_CGC,            ');
       SQL.Add('       CASE             ');
       SQL.Add('           WHEN FOR_INS IS NULL THEN ''ISENTO''            ');
       SQL.Add('           WHEN FOR_INS = '''' THEN ''ISENTO''            ');
       SQL.Add('           WHEN FOR_INS = ''ISENTA'' THEN ''ISENTO''            ');
       SQL.Add('           ELSE FOR_INS            ');
       SQL.Add('       END AS NUM_INSC_EST,            ');
       SQL.Add('       COALESCE (FOR_ENDERECO, ''A DEFINIR'') AS DES_ENDERECO,            ');
       SQL.Add('       COALESCE (FOR_BAIRRO, ''A DEFINIR'') AS DES_BAIRRO,            ');
       SQL.Add('       COALESCE (FOR_CIDADE, ''GARCA'') AS DES_CIDADE,            ');
       SQL.Add('       FOR_ESTADO AS DES_SIGLA,            ');
       SQL.Add('       REPLACE (REPLACE (FOR_CEP, ''.'', ''''), ''-'', '''') AS NUM_CEP,            ');
       SQL.Add('       REPLACE ((CONCAT( FOR_DDD1, FOR_TELEFONE1)),''-'', '''') AS NUM_FONE,            ');
       SQL.Add('       REPLACE ((CONCAT( FOR_FAXDDD, FOR_FAX)),''-'', '''') AS NUM_FAX,            ');
       SQL.Add('       CASE            ');
       SQL.Add('           WHEN FOR_CONTATO IS NULL THEN            ');
       SQL.Add('               CASE             ');
       SQL.Add('                   WHEN FOR_VENDEDOR IS NOT NULL THEN FOR_VENDEDOR            ');
       SQL.Add('                   ELSE ''''            ');
       SQL.Add('               END             ');
       SQL.Add('           ELSE FOR_CONTATO            ');
       SQL.Add('       END AS DES_CONTATO,            ');
       SQL.Add('       0 AS QTD_DIA_CARENCIA,            ');
       SQL.Add('       0 AS NUM_FREQ_VISITA,            ');
       SQL.Add('       0 AS VAL_DESCONTO,            ');
       SQL.Add('       0 AS NUM_PRAZO,            ');
       SQL.Add('       CASE             ');
       SQL.Add('           WHEN (SELECT COUNT(*) FROM DEVOLUCAO_FOR WHERE dbo.DEVOLUCAO_FOR.FOR_CODIGO = dbo.FORNECEDOR.FOR_CODIGO) > 0 THEN ''S''             ');
       SQL.Add('           ELSE ''N''             ');
       SQL.Add('       END AS ACEITA_DEVOL_MER,            ');
       SQL.Add('       ''N'' AS CAL_IPI_VAL_BRUTO,                  ');
       SQL.Add('          ''N'' AS CAL_ICMS_ENC_FIN,                  ');
       SQL.Add('          ''N'' AS CAL_ICMS_VAL_IPI,                  ');
       SQL.Add('          ''N'' AS MICRO_EMPRESA,            ');
       SQL.Add('       FORNECEDOR.FOR_CODIGO AS COD_FORNECEDOR_ANT,            ');
       SQL.Add('       CASE            ');
       SQL.Add('           WHEN (FOR_ENDNRO IS NULL) OR (FOR_ENDNRO = ''.'') OR (FOR_ENDNRO = '' '') OR (FOR_ENDNRO = ''-'') OR (FOR_ENDNRO = ''SN'') OR (FOR_ENDNRO = ''0'') THEN ''S/N''            ');
       SQL.Add('           ELSE REPLACE (REPLACE ( UPPER (FOR_ENDNRO), ''-'', ''''), ''.'', '''')            ');
       SQL.Add('       END AS NUM_ENDERECO,            ');
       SQL.Add('       CONCAT (FOR_OBS, FOR_OBS_PEDIDO) AS DES_OBSERVACAO,            ');
       SQL.Add('       CASE                ');
       SQL.Add('           WHEN FOR_EMAIL IS NULL THEN ''''            ');
       SQL.Add('           ELSE UPPER (FOR_EMAIL)            ');
       SQL.Add('       END AS DES_EMAIL,               ');
       SQL.Add('       COALESCE (FOR_PAGINA, '''') AS DES_WEB_SITE,            ');
       SQL.Add('       ''N'' AS FABRICANTE,                  ');
       SQL.Add('       CASE             ');
       SQL.Add('           WHEN TPF_CODIGO = ''PR'' THEN ''S''            ');
       SQL.Add('           ELSE ''N''             ');
       SQL.Add('          END AS FLG_PRODUTOR_RURAL,            ');
       SQL.Add('       CASE             ');
       SQL.Add('           WHEN FOR_TIPOFRETE = ''CIF'' THEN 0            ');
       SQL.Add('           WHEN FOR_TIPOFRETE = ''FOB'' THEN 1            ');
       SQL.Add('           ELSE 0            ');
       SQL.Add('       END AS TIPO_FRETE,            ');
       SQL.Add('       CASE       ');
       SQL.Add('   		         WHEN TPF_CODIGO = ''SN'' THEN ''S''      ');
       SQL.Add('   		         ELSE ''N''       ');
       SQL.Add('   	       END AS FLG_SIMPLES,            ');
       SQL.Add('       ''N'' AS FLG_SUBSTITUTO_TRIB,            ');
       SQL.Add('       0 AS COD_CONTA,            ');
       SQL.Add('       CASE             ');
       SQL.Add('           WHEN FOR_STATUS = ''A'' THEN ''N''            ');
       SQL.Add('           ELSE ''S''            ');
       SQL.Add('       END AS INATIVO,            ');
       SQL.Add('       CASE    ');
       SQL.Add('   		WHEN GRU_CODIGO IS NULL THEN 21   ');
       SQL.Add('   		ELSE GRU_CODIGO + 2000   ');
       SQL.Add('   	END AS COD_CLASSIF,            ');
       SQL.Add('       CONVERT(CHAR, FOR_ALTERACAO, 103) AS DTA_CADASTRO,            ');
       SQL.Add('       0 AS VAL_CREDITO,                  ');
       SQL.Add('       0 AS VAL_DEBITO,            ');
       SQL.Add('       1 AS PED_MIN_VAL,            ');
       SQL.Add('       CASE                ');
       SQL.Add('           WHEN FOR_EMAIL IS NOT NULL THEN           ');
       SQL.Add('               CASE          ');
       SQL.Add('                   WHEN UPPER(FOR_EMAIL) <> UPPER(FOR_EMAIL_VENDEDOR1) THEN UPPER (FOR_EMAIL+''; ''+ FOR_EMAIL_VENDEDOR1)          ');
       SQL.Add('                   ELSE FOR_EMAIL         ');
       SQL.Add('               END         ');
       SQL.Add('           ELSE ''''         ');
       SQL.Add('       END AS DES_EMAIL_VEND,                  ');
       SQL.Add('       '''' AS SENHA_COTACAO,                  ');
       SQL.Add('       1 AS TIPO_PRODUTOR,            ');
       SQL.Add('       CASE             ');
       SQL.Add('           WHEN FOR_VENCELULAR IS NULL THEN ''''            ');
       SQL.Add('           ELSE REPLACE ((CONCAT( FOR_VENDDDCELULAR, FOR_VENCELULAR)),''-'', '''')             ');
       SQL.Add('       END AS NUM_CELULAR,           ');
       SQL.Add('       FOR_ENT_NAOVALIDAR_PEDIDO AS FLG_RECEB_SEM_PEDIDO,      ');
       SQL.Add('   	   FORNECEDOR_BANCOS.BAN_CODIGO AS COD_BANCOFORN,      ');
       SQL.Add('   	   FORNECEDOR_BANCOS.BAN_AGENCIA AS COD_AGENCIAFORN,      ');
       SQL.Add('   	   FORNECEDOR_BANCOS.BAN_DVAGENCIA AS COD_DIGAGFORN,         ');
       SQL.Add('       FORNECEDOR_BANCOS.BAN_CONTA AS COD_CONTACCFORN,            ');
       SQL.Add('   	   FORNECEDOR_BANCOS.BAN_DVCONTA AS COD_DIGCCFORN      ');
       SQL.Add('   FROM dbo.FORNECEDOR      ');
       SQL.Add('   LEFT JOIN dbo.FORNECEDOR_BANCOS ON      ');
       SQL.Add('   	(FORNECEDOR.FOR_CODIGO = FORNECEDOR_BANCOS.FOR_CODIGO)   ');
       SQL.Add('   ORDER BY   ');
       SQL.Add('       FORNECEDOR.FOR_CODIGO ASC   ');


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

        Layout.FieldByName('DES_FORNECEDOR').AsString :=
          StrSubstLtsAct(Layout.FieldByName('DES_FORNECEDOR').AsString);
        Layout.FieldByName('DES_FANTASIA').AsString :=
          StrSubstLtsAct(Layout.FieldByName('DES_FANTASIA').AsString);
        Layout.FieldByName('DES_BAIRRO').AsString :=
          StrSubstLtsAct(Layout.FieldByName('DES_BAIRRO').AsString);
        Layout.FieldByName('DES_ENDERECO').AsString :=
          StrSubstLtsAct(Layout.FieldByName('DES_ENDERECO').AsString);

        Layout.FieldByName('NUM_CGC').AsString :=
          StrRetNums(Layout.FieldByName('NUM_CGC').AsString);
        Layout.FieldByName('NUM_CEP').AsString :=
          StrRetNums(Layout.FieldByName('NUM_CEP').AsString);
        Layout.FieldByName('NUM_ENDERECO').AsString :=
          StrRetNums(Layout.FieldByName('NUM_ENDERECO').AsString);

        if (Layout.FieldByName('NUM_ENDERECO').AsString = '') then
          Layout.FieldByName('NUM_ENDERECO').AsString := 'S/N';

        if Length(Layout.FieldByName('NUM_CGC').AsString) > 11 then
        begin
          if not ValidaCGC(Layout.FieldByName('NUM_CGC').AsString) then
            Layout.FieldByName('NUM_CGC').AsString := '';
        end
        else if not ValidaCpf(Layout.FieldByName('NUM_CGC').AsString) then
          Layout.FieldByName('NUM_CGC').AsString := '';

        Layout.FieldByName('NUM_FONE').AsString :=
          StrRetNums(FieldByName('NUM_FONE').AsString);
        Layout.FieldByName('NUM_FAX').AsString :=
          StrRetNums(FieldByName('NUM_FAX').AsString);

        inscEst := StrRetNums(Layout.FieldByName('NUM_INSC_EST').AsString);

        if (inscEst = '') then
          Layout.FieldByName('NUM_INSC_EST').AsString := 'ISENTO'
        else
        begin
          if StrToFloat(inscEst) = 0 then
            Layout.FieldByName('NUM_INSC_EST').AsString := ''
          else
            Layout.FieldByName('NUM_INSC_EST').AsString := inscEst;
        end;

        //Layout.FieldByName('DTA_CADASTRO').AsDateTime := DATE;

        if Layout.FieldByName('NUM_CEP').AsString = '' then
          Layout.FieldByName('NUM_CEP').AsString := '17400000';

        if Layout.FieldByName('DES_ENDERECO').AsString = '' then
          Layout.FieldByName('DES_ENDERECO').AsString := 'A DEFINIR';

        if Layout.FieldByName('DES_BAIRRO').AsString = '' then
          Layout.FieldByName('DES_BAIRRO').AsString := 'A DEFINIR';

        //if Layout.FieldByName('DES_CIDADE').AsString = '' then
        //  Layout.FieldByName('DES_CIDADE').AsString := 'GARCA';

        if Layout.FieldByName('DES_SIGLA').AsString = '' then
          Layout.FieldByName('DES_SIGLA').AsString := 'SP';

        Layout.FieldByName('DES_OBSERVACAO').AsString :=
          StrReplace(StrLBReplace(FieldByName('DES_OBSERVACAO').AsString),
          '\n', '');
        Layout.FieldByName('DES_EMAIL').AsString :=
          StrReplace(StrLBReplace(FieldByName('DES_EMAIL').AsString), '\n', '');

        Layout.FieldByName('DES_EMAIL_VEND').AsString :=
          UpperCase(StrDelMultSpaces(StrReplace(StrLBReplace(FieldByName('DES_EMAIL_VEND').AsString),
          '\n', '')));


        //if Layout.FieldByName('COD_BANCOFORN').AsString = '' then
        //   Layout.FieldByName('COD_BANCOFORN').AsString := '0';



        // Layout.FieldByName('DES_OBSERVACAO').AsString := observacao;
        // Layout.FieldByName('DES_EMAIL').AsString := email;

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

procedure TFrmSmBomPrecoOurinhosGestora.GerarCondPagForn;
begin
  inherited;

  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

     SQL.Add('          SELECT DISTINCT       ');
     SQL.Add('          	FORNECEDOR.FOR_CODIGO AS COD_FORNECEDOR,      ');
     SQL.Add('          	COALESCE (PARC_DIAS, ''30'') AS NUM_CONDICAO,      ');
     SQL.Add('          	2 AS COD_CONDICAO,      ');
     SQL.Add('          	CASE   ');
     SQL.Add('   			WHEN FOR_FORMAPAGTO = ''BO'' THEN 8   ');
     SQL.Add('   			WHEN FOR_FORMAPAGTO = ''CH'' THEN 3   ');
     SQL.Add('   			WHEN FOR_FORMAPAGTO = ''CA'' THEN 9   ');
     SQL.Add('   			ELSE 8   ');
     SQL.Add('   		END AS COD_ENTIDADE,      ');
     SQL.Add('          	FOR_CGC AS NUM_CGC      ');
     SQL.Add('          FROM dbo.FORNECEDOR    ');
     SQL.Add('   	   LEFT JOIN dbo.FORNECEDOR_PARCELAS ON   ');
     SQL.Add('   		(FORNECEDOR.FOR_CODIGO = FORNECEDOR_PARCELAS.FOR_CODIGO)   ');
     SQL.Add('   	         ');
     SQL.Add('          ORDER BY    ');
     SQL.Add('   		FORNECEDOR.FOR_CODIGO ASC;   ');

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

        Layout.FieldByName('NUM_CGC').AsString :=
          StrRetNums(Layout.FieldByName('NUM_CGC').AsString);

        Layout.WriteLine;
      except
        On E: Exception do
          FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
      end;
      Next;
    end;
  end;
end;




(*
  procedure TFrmSmBomPrecoOurinhosGestora.CkbProdLojaClick(Sender: TObject);
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
  end; *)


procedure TFrmSmBomPrecoOurinhosGestora.GerarSecao;
begin
  inherited;

  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('   --SECAO   ');
    SQL.Add('   SELECT DISTINCT    ');
    SQL.Add('   	dbo.DEPARTAMENTO.DEP_CODIGO AS COD_SECAO,   ');
    SQL.Add('   	REPLACE (dbo.DEPARTAMENTO.DEP_DESCRICAO, ''.'', '''') AS DES_SECAO,   ');
    SQL.Add('   	0 AS VAL_META   ');
    SQL.Add('   FROM dbo.PRODUTOS    ');
    SQL.Add('   INNER JOIN dbo.DEPARTAMENTO ON   ');
    SQL.Add('   	(dbo.PRODUTOS.DEP_CODIGO = dbo.DEPARTAMENTO.DEP_CODIGO);   ');
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

procedure TFrmSmBomPrecoOurinhosGestora.GerarGrupo;
begin
  inherited;
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

     SQL.Add('   SELECT DISTINCT          ');
     SQL.Add('       CASE    ');
     SQL.Add('   		WHEN dbo.PRODUTOS.DEP_CODIGO = 75 THEN 999   ');
     SQL.Add('   		ELSE COALESCE (dbo.PRODUTOS.DEP_CODIGO, 999)    ');
     SQL.Add('   	END AS COD_SECAO,         ');
     SQL.Add('       COALESCE (dbo.PRODUTOS.GRU_CODIGO, 999) AS COD_GRUPO,         ');
     SQL.Add('       COALESCE (dbo.GRU_PRODUTOS.GRU_DESCRICAO, ''Á DEFINIR'') AS DES_GRUPO,         ');
     SQL.Add('       0 AS VAL_META         ');
     SQL.Add('   FROM dbo.PRODUTOS          ');
     SQL.Add('   LEFT JOIN dbo.GRU_PRODUTOS ON         ');
     SQL.Add('       (dbo.PRODUTOS.GRU_CODIGO = dbo.GRU_PRODUTOS.GRU_CODIGO)      ');
     SQL.Add('   ORDER BY   ');
     SQL.Add('   	COD_SECAO ASC   ');

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

procedure TFrmSmBomPrecoOurinhosGestora.GerarSubGrupo;
begin
  inherited;

  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

     SQL.Add('   --SUBGRUPO      ');
     SQL.Add('   SELECT DISTINCT      ');
     SQL.Add('    CASE    ');
     SQL.Add('   		WHEN dbo.PRODUTOS.DEP_CODIGO = 75 THEN 999   ');
     SQL.Add('   		ELSE COALESCE (dbo.PRODUTOS.DEP_CODIGO, 999)    ');
     SQL.Add('   	END AS COD_SECAO, ');
     SQL.Add('   	COALESCE (dbo.PRODUTOS.GRU_CODIGO, 999) AS COD_GRUPO,      ');
     SQL.Add('   	CASE    ');
     SQL.Add('       	WHEN dbo.PRODUTOS.SUB_CODIGO = 0 OR dbo.PRODUTOS.SUB_CODIGO IS NULL THEN 999   ');
     SQL.Add('       	ELSE COALESCE(dbo.PRODUTOS.SUB_CODIGO, 999)   ');
     SQL.Add('       END AS COD_SUB_GRUPO,      ');
     SQL.Add('   	COALESCE (dbo.SUBGRUPO_PRODUTOS.SUB_DESCRICAO, ''Á DEFINIR'') AS DES_SUB_GRUPO,       ');
     SQL.Add('   	0 AS VAL_META,      ');
     SQL.Add('   	0 AS VAL_MARGEM_REF,      ');
     SQL.Add('   	0 AS QTD_DIA_SEGURANCA,      ');
     SQL.Add('   	''N'' AS FLG_ALCOOLICO      ');
     SQL.Add('   FROM dbo.PRODUTOS       ');
     SQL.Add('   LEFT JOIN dbo.SUBGRUPO_PRODUTOS ON      ');
     SQL.Add('   	(dbo.PRODUTOS.SUB_CODIGO = dbo.SUBGRUPO_PRODUTOS.SUB_CODIGO)   ');
     SQL.Add('   ORDER BY   ');
     SQL.Add('   	COD_SECAO ASC   ');

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

procedure TFrmSmBomPrecoOurinhosGestora.GerarReceitas;
// var
// texto : string;
begin
  inherited;

  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('   SELECT DISTINCT    ');
    SQL.Add('   	dbo.INFORMACAOEXTRA.INF_EXT_CODIGO AS COD_INFO_RECEITA,   ');
    SQL.Add('   	dbo.INFORMACAOEXTRA.INF_EXT_OBS AS DES_INFO_RECEITA,   ');
    SQL.Add('   	CAST (dbo.INFORMACAOEXTRA.INF_EXTRA AS VARCHAR(700)) AS DETALHAMENTO  ');
    SQL.Add('   FROM    ');
    SQL.Add('   	dbo.INFORMACAOEXTRA    ');
    SQL.Add('   INNER JOIN    ');
    SQL.Add('   	dbo.PRODUTOS ON    ');
    SQL.Add('   	(dbo.INFORMACAOEXTRA.INF_EXT_CODIGO = dbo.PRODUTOS.INF_EXT_CODIGO)   ');
    SQL.Add('      ');
    SQL.Add('   ORDER BY    ');
    SQL.Add('   	dbo.INFORMACAOEXTRA.INF_EXT_CODIGO ASC;   ');

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

        (* if not PLUValido(Layout.FieldByName('COD_INFO_RECEITA').AsString) then
          begin
          Layout.FieldByName('COD_INFO_RECEITA').AsString := GerarPlu(Copy(Layout.FieldByName('COD_INFO_RECEITA').AsString, 1, Length(Layout.FieldByName('COD_INFO_RECEITA').AsString) - 1));
          end; *)

        Layout.FieldByName('DETALHAMENTO').AsString :=
          StrReplace(StrLBReplace(StringReplace(FieldByName('DETALHAMENTO')
          .AsString, #$A, '', [rfReplaceAll])), '\n', '');

        // Layout.FieldByName('DETALHAMENTO').AsString := StrReplace(StrLBReplace( StringReplace(FieldByName('DETALHAMENTO').AsString,#$A, '', [rfReplaceAll]) ), '\n', '') ;
        // texto := StringReplace(StringReplace(StringReplace(Layout.FieldByName('DETALHAMENTO').AsString, #$D#$A, '', [rfReplaceAll]), #$A, '', [rfReplaceAll]), '#$A', '', [rfReplaceAll]);
        // Layout.FieldByName('DETALHAMENTO').AsString := texto;

        Layout.WriteLine;
      except
        On E: Exception do
          FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
      end;
      Next;
    end;
  end;
end;

procedure TFrmSmBomPrecoOurinhosGestora.GerarInfoNutricionais;
begin
  inherited;

  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('   SELECT    ');
    SQL.Add('       NUTRI.*,   ');
    SQL.Add('       ROUND( CONVERT( DECIMAL(10,2), NUTRI.VALOR_CALORICO * 100 / 2000), 0) AS VD_VALOR_CALORICO,   ');
    SQL.Add('       ROUND( CONVERT( DECIMAL(10,2), NUTRI.CARBOIDRATO * 100 / 300), 0) AS VD_CARBOIDRATO,   ');
    SQL.Add('       ROUND( CONVERT( DECIMAL(10,2), NUTRI.PROTEINA * 100 / 75), 0) AS VD_PROTEINA,   ');
    SQL.Add('       ROUND( CONVERT( DECIMAL(10,2), NUTRI.GORDURA_TOTAL * 100 / 55), 0) AS VD_GORDURA_TOTAL,   ');
    SQL.Add('       ROUND( CONVERT( DECIMAL(10,2), NUTRI.GORDURA_SATURADA * 100 / 22), 0) AS VD_GORDURA_SATURADA,   ');
    SQL.Add('       ROUND( CONVERT( DECIMAL(10,2), NUTRI.GORDURA_TRANS * 100 / 33), 0) AS VD_GORDURA_TRANS,   ');
    SQL.Add('       ROUND( CONVERT( DECIMAL(10,2), NUTRI.COLESTEROL * 100 / 300), 0) AS VD_COLESTEROL,   ');
    SQL.Add('       ROUND( CONVERT( DECIMAL(10,2), NUTRI.FIBRA_ALIMENTAR * 100 / 25), 0) AS VD_FIBRA_ALIMENTAR,   ');
    SQL.Add('       ROUND( CONVERT( DECIMAL(10,2), NUTRI.CALCIO * 100 / 1000), 0) AS VD_CALCIO,   ');
    SQL.Add('       ROUND( CONVERT( DECIMAL(10,2), NUTRI.FERRO * 100 / 14), 0) AS VD_FERRO,   ');
    SQL.Add('       ROUND( CONVERT( DECIMAL(10,2), NUTRI.SODIO * 100 /2400), 0) AS VD_SODIO   ');
    SQL.Add('   FROM    ');
    SQL.Add('       (   ');
    SQL.Add('       SELECT   ');
    SQL.Add('           NUTRICIONAL.NUT_CODIGO AS COD_INFO_NUTRICIONAL,   ');
    SQL.Add('           UPPER (NUTRICIONAL.NUT_OBS) AS DES_INFO_NUTRICIONAL,   ');
    SQL.Add('           NUTRICIONAL.NUT_PORCAO AS PORCAO,   ');
    SQL.Add('           CASE    ');
    SQL.Add('               WHEN  (INFNUTITENS.INF_CODIGO = 1 ) AND (INFNUTITENS.INF_STATUS = ''S'') THEN 0   ');
    SQL.Add('               ELSE (SELECT INFNUTITENS.INF_QTDE FROM INFNUTITENS LEFT JOIN NUTRICIONAL_INFNUTITENS ON (INFNUTITENS.INFNUT_CODIGO = NUTRICIONAL_INFNUTITENS.INFNUT_CODIGO)    ');
    SQL.Add('                       LEFT JOIN NUTRICIONAL NUT ON (NUTRICIONAL_INFNUTITENS.NUT_CODIGO = NUT.NUT_CODIGO) WHERE INFNUTITENS.INF_CODIGO = 1 AND NUT.NUT_CODIGO = NUTRICIONAL.NUT_CODIGO)    ');
    SQL.Add('           END AS VALOR_CALORICO,   ');
    SQL.Add('           CASE    ');
    SQL.Add('               WHEN (INFNUTITENS.INF_CODIGO = 2 ) AND (INFNUTITENS.INF_STATUS = ''S'') THEN 0   ');
    SQL.Add('               ELSE (SELECT INFNUTITENS.INF_QTDE FROM INFNUTITENS LEFT JOIN NUTRICIONAL_INFNUTITENS ON (INFNUTITENS.INFNUT_CODIGO = NUTRICIONAL_INFNUTITENS.INFNUT_CODIGO)    ');
    SQL.Add('                       LEFT JOIN NUTRICIONAL NUT ON (NUTRICIONAL_INFNUTITENS.NUT_CODIGO = NUT.NUT_CODIGO) WHERE INFNUTITENS.INF_CODIGO = 2 AND NUT.NUT_CODIGO = NUTRICIONAL.NUT_CODIGO)    ');
    SQL.Add('           END AS CARBOIDRATO,   ');
    SQL.Add('           CASE    ');
    SQL.Add('               WHEN (INFNUTITENS.INF_CODIGO = 3 ) AND (INFNUTITENS.INF_STATUS = ''S'') THEN 0    ');
    SQL.Add('               ELSE (SELECT INFNUTITENS.INF_QTDE FROM INFNUTITENS LEFT JOIN NUTRICIONAL_INFNUTITENS ON (INFNUTITENS.INFNUT_CODIGO = NUTRICIONAL_INFNUTITENS.INFNUT_CODIGO)    ');
    SQL.Add('                       LEFT JOIN NUTRICIONAL NUT ON (NUTRICIONAL_INFNUTITENS.NUT_CODIGO = NUT.NUT_CODIGO) WHERE INFNUTITENS.INF_CODIGO = 3 AND NUT.NUT_CODIGO = NUTRICIONAL.NUT_CODIGO)    ');
    SQL.Add('           END AS PROTEINA,   ');
    SQL.Add('           CASE    ');
    SQL.Add('               WHEN(INFNUTITENS.INF_CODIGO = 4 ) AND (INFNUTITENS.INF_STATUS = ''S'') THEN 0   ');
    SQL.Add('               ELSE (SELECT INFNUTITENS.INF_QTDE FROM INFNUTITENS LEFT JOIN NUTRICIONAL_INFNUTITENS ON (INFNUTITENS.INFNUT_CODIGO = NUTRICIONAL_INFNUTITENS.INFNUT_CODIGO)    ');
    SQL.Add('                       LEFT JOIN NUTRICIONAL NUT ON (NUTRICIONAL_INFNUTITENS.NUT_CODIGO = NUT.NUT_CODIGO) WHERE INFNUTITENS.INF_CODIGO = 4 AND NUT.NUT_CODIGO = NUTRICIONAL.NUT_CODIGO)    ');
    SQL.Add('           END AS GORDURA_TOTAL,   ');
    SQL.Add('           CASE    ');
    SQL.Add('               WHEN (INFNUTITENS.INF_CODIGO = 5 ) AND (INFNUTITENS.INF_STATUS = ''S'') THEN 0    ');
    SQL.Add('               ELSE (SELECT INFNUTITENS.INF_QTDE FROM INFNUTITENS LEFT JOIN NUTRICIONAL_INFNUTITENS ON (INFNUTITENS.INFNUT_CODIGO = NUTRICIONAL_INFNUTITENS.INFNUT_CODIGO)    ');
    SQL.Add('                       LEFT JOIN NUTRICIONAL NUT ON (NUTRICIONAL_INFNUTITENS.NUT_CODIGO = NUT.NUT_CODIGO) WHERE INFNUTITENS.INF_CODIGO = 5 AND NUT.NUT_CODIGO = NUTRICIONAL.NUT_CODIGO)    ');
    SQL.Add('           END AS GORDURA_SATURADA,   ');
    SQL.Add('           CASE    ');
    SQL.Add('               WHEN (INFNUTITENS.INF_CODIGO = 6 ) AND (INFNUTITENS.INF_STATUS = ''S'') THEN 0    ');
    SQL.Add('               ELSE (SELECT INFNUTITENS.INF_QTDE FROM INFNUTITENS LEFT JOIN NUTRICIONAL_INFNUTITENS ON (INFNUTITENS.INFNUT_CODIGO = NUTRICIONAL_INFNUTITENS.INFNUT_CODIGO)    ');
    SQL.Add('                       LEFT JOIN NUTRICIONAL NUT ON (NUTRICIONAL_INFNUTITENS.NUT_CODIGO = NUT.NUT_CODIGO) WHERE INFNUTITENS.INF_CODIGO = 6 AND NUT.NUT_CODIGO = NUTRICIONAL.NUT_CODIGO)    ');
    SQL.Add('           END AS GORDURA_TRANS,   ');
    SQL.Add('           CASE    ');
    SQL.Add('               WHEN (INFNUTITENS.INF_CODIGO = 9 ) AND (INFNUTITENS.INF_STATUS = ''S'') THEN 0   ');
    SQL.Add('               ELSE (SELECT INFNUTITENS.INF_QTDE FROM INFNUTITENS LEFT JOIN NUTRICIONAL_INFNUTITENS ON (INFNUTITENS.INFNUT_CODIGO = NUTRICIONAL_INFNUTITENS.INFNUT_CODIGO)    ');
    SQL.Add('                       LEFT JOIN NUTRICIONAL NUT ON (NUTRICIONAL_INFNUTITENS.NUT_CODIGO = NUT.NUT_CODIGO) WHERE INFNUTITENS.INF_CODIGO = 9 AND NUT.NUT_CODIGO = NUTRICIONAL.NUT_CODIGO)    ');
    SQL.Add('           END AS COLESTEROL,   ');
    SQL.Add('           CASE   ');
    SQL.Add('               WHEN (INFNUTITENS.INF_CODIGO = 7 ) AND (INFNUTITENS.INF_STATUS = ''S'') THEN 0   ');
    SQL.Add('               ELSE (SELECT INFNUTITENS.INF_QTDE FROM INFNUTITENS LEFT JOIN NUTRICIONAL_INFNUTITENS ON (INFNUTITENS.INFNUT_CODIGO = NUTRICIONAL_INFNUTITENS.INFNUT_CODIGO)    ');
    SQL.Add('                       LEFT JOIN NUTRICIONAL NUT ON (NUTRICIONAL_INFNUTITENS.NUT_CODIGO = NUT.NUT_CODIGO) WHERE INFNUTITENS.INF_CODIGO = 7 AND NUT.NUT_CODIGO = NUTRICIONAL.NUT_CODIGO)    ');
    SQL.Add('           END AS FIBRA_ALIMENTAR,   ');
    SQL.Add('           CASE   ');
    SQL.Add('               WHEN (INFNUTITENS.INF_CODIGO = 10 ) AND (INFNUTITENS.INF_STATUS = ''S'') THEN 0   ');
    SQL.Add('               ELSE (SELECT INFNUTITENS.INF_QTDE FROM INFNUTITENS LEFT JOIN NUTRICIONAL_INFNUTITENS ON (INFNUTITENS.INFNUT_CODIGO = NUTRICIONAL_INFNUTITENS.INFNUT_CODIGO)    ');
    SQL.Add('                       LEFT JOIN NUTRICIONAL NUT ON (NUTRICIONAL_INFNUTITENS.NUT_CODIGO = NUT.NUT_CODIGO) WHERE INFNUTITENS.INF_CODIGO = 10 AND NUT.NUT_CODIGO = NUTRICIONAL.NUT_CODIGO)    ');
    SQL.Add('           END AS CALCIO,   ');
    SQL.Add('           CASE   ');
    SQL.Add('               WHEN (INFNUTITENS.INF_CODIGO = 11 ) AND (INFNUTITENS.INF_STATUS = ''S'') THEN 0   ');
    SQL.Add('               ELSE (SELECT INFNUTITENS.INF_QTDE FROM INFNUTITENS LEFT JOIN NUTRICIONAL_INFNUTITENS ON (INFNUTITENS.INFNUT_CODIGO = NUTRICIONAL_INFNUTITENS.INFNUT_CODIGO)    ');
    SQL.Add('                       LEFT JOIN NUTRICIONAL NUT ON (NUTRICIONAL_INFNUTITENS.NUT_CODIGO = NUT.NUT_CODIGO) WHERE INFNUTITENS.INF_CODIGO = 11 AND NUT.NUT_CODIGO = NUTRICIONAL.NUT_CODIGO)   ');
    SQL.Add('           END AS FERRO,   ');
    SQL.Add('           CASE    ');
    SQL.Add('               WHEN (INFNUTITENS.INF_CODIGO = 8 ) AND (INFNUTITENS.INF_STATUS = ''S'') THEN 0    ');
    SQL.Add('               ELSE (SELECT INFNUTITENS.INF_QTDE FROM INFNUTITENS LEFT JOIN NUTRICIONAL_INFNUTITENS ON (INFNUTITENS.INFNUT_CODIGO = NUTRICIONAL_INFNUTITENS.INFNUT_CODIGO)    ');
    SQL.Add('                       LEFT JOIN NUTRICIONAL NUT ON (NUTRICIONAL_INFNUTITENS.NUT_CODIGO = NUT.NUT_CODIGO) WHERE INFNUTITENS.INF_CODIGO = 8 AND NUT.NUT_CODIGO = NUTRICIONAL.NUT_CODIGO)   ');
    SQL.Add('           END AS SODIO,   ');
    SQL.Add('           CASE NUTRICIONAL.UNI_CODIGO   ');
    SQL.Add('               WHEN 1 THEN ''G''   ');
    SQL.Add('               WHEN 2 THEN ''ML''   ');
    SQL.Add('               WHEN 3 THEN ''UN''   ');
    SQL.Add('               WHEN 4 THEN ''MG''   ');
    SQL.Add('           END AS UNIDADE_PORCAO,   ');
    SQL.Add('           ''ADEFINIR'' AS DES_PORCAO,   ');
    SQL.Add('           NUTRICIONAL.NUT_QTDE AS PARTE_INTEIRA_MED_CASEIRA,   ');
    SQL.Add('           NUTMEDIDA.MED_MGV AS MED_CASEIRA_UTILIZADA   ');
    SQL.Add('      ');
    SQL.Add('       FROM INFNUTITENS   ');
    SQL.Add('       LEFT JOIN NUTRICIONAL_INFNUTITENS ON   ');
    SQL.Add('           (INFNUTITENS.INFNUT_CODIGO = NUTRICIONAL_INFNUTITENS.INFNUT_CODIGO)   ');
    SQL.Add('       LEFT JOIN INFNUTRICIONAL ON    ');
    SQL.Add('           (INFNUTITENS.INF_CODIGO = INFNUTRICIONAL.INF_CODIGO)   ');
    SQL.Add('       LEFT JOIN NUTRICIONAL ON   ');
    SQL.Add('           (INFNUTITENS.INFNUT_CODIGO = NUTRICIONAL.NUT_CODIGO)   ');
    SQL.Add('       LEFT JOIN NUTMEDIDA ON   ');
    SQL.Add('           (NUTRICIONAL.MED_CODIGO = NUTMEDIDA.MED_CODIGO)   ');
    SQL.Add('       WHERE NUTRICIONAL.NUT_CODIGO IS NOT NULL   ');
    SQL.Add('       ) AS NUTRI;   ');

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

procedure TFrmSmBomPrecoOurinhosGestora.GerarCest;
var
  Count: Integer;
begin
  inherited;
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

       SQL.Add('   SELECT DISTINCT   ');
       SQL.Add('   	0 AS COD_CEST,   ');
       SQL.Add('   	''A DEFINIR'' AS DES_CEST,   ');
       SQL.Add('   	CASE    ');
       SQL.Add('   	WHEN PRO_CEST = '''' THEN ''9999999''   ');
       SQL.Add('   	ELSE PRO_CEST   ');
       SQL.Add('   	END AS NUM_CEST   ');
       SQL.Add('   FROM PRODUTOS   ');
       SQL.Add('   WHERE PRO_CEST IS NOT NULL   ');
       SQL.Add('   ORDER BY    ');
       SQL.Add('   	NUM_CEST ASC   ');

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

        Layout.FieldByName('COD_CEST').AsInteger := Count;

        Layout.FieldByName('NUM_CEST').AsString :=
          StrRetNums(Layout.FieldByName('NUM_CEST').AsString);

        Layout.WriteLine;
      except
        On E: Exception do
          FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
      end;
      Next;
    end;
  end;
end;

procedure TFrmSmBomPrecoOurinhosGestora.GerarNCM;
var
  Count: Integer;
begin
  inherited;
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;
(*

 *)
   SQL.Add('   SELECT  DISTINCT         ');
   SQL.Add('       0 AS COD_NCM,               ');
   SQL.Add('       CASE          ');
   SQL.Add('         WHEN CAST(NBM_DESCRICAO AS VARCHAR(50)) = '''' THEN ''Á DEFINIR''         ');
   SQL.Add('         ELSE REPLACE (CAST(COALESCE(NBM_DESCRICAO, ''Á DEFINIR'') AS VARCHAR(50)), CHAR ( 13 ) +  Char ( 10 ), '''')        ');
   SQL.Add('       END AS DES_NCM,          ');
   SQL.Add('       COALESCE (NBM, PRODUTOS.PRO_CLASFISCAL) AS NUM_NCM,               ');
   SQL.Add('       CASE                ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = '''' ) AND (PRO_CST_COFINS = '''' ) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''N''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 50) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65)        THEN ''N''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 50) AND (PRO_CST_COFINS = 3) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''N''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 50) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA = 1.65) THEN ''N''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 60) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''N''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''N''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 4) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''S''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 4) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL) THEN ''S''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 06) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''S'' --S        ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 06) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''S'' --S         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''S''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''S''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 72) AND (PRO_CST_COFINS = 9) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''S''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 4) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''S''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 06) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''S''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 06) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''S''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''S''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 98) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''N''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 98) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''S''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 98) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''S''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''N''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''S''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''S''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''N''                ');
   SQL.Add('       END  AS FLG_NAO_PIS_COFINS,            ');
   SQL.Add('       CASE                ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = '''' ) AND (PRO_CST_COFINS = '''' ) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''-1''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 50) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65)        THEN ''-1''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 50) AND (PRO_CST_COFINS = 3) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''-1''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 50) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA = 1.65) THEN ''-1''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 60) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''-1''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''-1''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 4) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''1''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 4) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL) THEN ''1''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 6) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''0'' --0         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 6) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''-0'' --0         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''0''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''0''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 72) AND (PRO_CST_COFINS = 9) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''4''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 4) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''0''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 06) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''0''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 06) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''0''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''0''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 98) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''-1''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 98) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''3''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 98) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''3''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''-1''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''3''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''3''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''-1''          ');
   SQL.Add('       END AS TIPO_NAO_PIS_COFINS,               ');
   SQL.Add('       CASE                ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = '''' ) AND (PRO_CST_COFINS = '''' ) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 50) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65)        THEN ''''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 50) AND (PRO_CST_COFINS = 3) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 50) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA = 1.65) THEN ''''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 60) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 4) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN natr_codigo         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 4) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL) THEN natr_codigo         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 6) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN COALESCE (natr_codigo, ''999'')         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 6) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN COALESCE (natr_codigo, ''999'')         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN COALESCE (natr_codigo, ''999'')         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN COALESCE (natr_codigo, ''999'')         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 72) AND (PRO_CST_COFINS = 9) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN COALESCE (natr_codigo, ''999'')         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 4) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN COALESCE (natr_codigo, ''999'')         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 6) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN COALESCE (natr_codigo, ''999'')         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 6) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN COALESCE (natr_codigo, ''999'')         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN COALESCE (natr_codigo, ''999'')         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 98) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 98) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN COALESCE (natr_codigo, ''999'')         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 98) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN COALESCE (natr_codigo, ''999'')         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN COALESCE (natr_codigo, ''999'')         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN COALESCE (natr_codigo, ''999'')         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''''                    ');
   SQL.Add('       END AS COD_TAB_SPED,               ');
   SQL.Add('       CASE         ');
   SQL.Add('         WHEN PRO_CEST  IS NULL OR PRO_CEST = '''' THEN ''99999999''         ');
   SQL.Add('         ELSE PRO_CEST         ');
   SQL.Add('       END AS NUM_CEST,            ');
   SQL.Add('       ''SP'' DES_SIGLA,               ');
   SQL.Add('       CASE                 ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 4.14 ) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''42''               ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 4.50) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''38''          ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 4.70) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''43''          ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 5.50) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''44''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 7.00) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''2''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 9.40) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''45''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 12.00) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''3''            ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 13.30) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''47''          ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 17.00) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''46''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''4''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 25.00) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''5''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 7.00) AND (TRI_REDUCAO = 41.67) AND (CST = 020) THEN ''6''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 12.00) AND (TRI_REDUCAO = 30.00) AND (CST = 020) THEN ''40''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 12.00) AND (TRI_REDUCAO = 41.66) AND (CST = 020) THEN ''39''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 12.00) AND (TRI_REDUCAO = 41.67) AND (CST = 020) THEN ''6''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 26.11) AND (CST = 020) THEN ''41''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 33.33) AND (CST = 020) THEN ''7''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 38.88) AND (CST = 020) THEN ''33''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 61.11) AND (CST = 020) THEN ''8''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 0.00) AND (TRI_REDUCAO = 0.00) AND (CST = 040) THEN ''1''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 0.00) AND (TRI_REDUCAO = 0.00) AND (CST = 041) THEN ''23''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 0.00) AND (TRI_REDUCAO = 0.00) AND (CST = 060) THEN ''25''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 0.00) AND (TRI_REDUCAO = 0.00) AND (CST = 090) THEN ''22''         ');
   SQL.Add('       END AS COD_TRIB_ENTRADA,               ');
   SQL.Add('       CASE                ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 4.14 ) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''42''               ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 4.50) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''38''          ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 4.70) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''43''          ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 5.50) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''44''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 7.00) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''2''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 9.40) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''45''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 12.00) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''3''            ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 13.30) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''47''          ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 17.00) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''46''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''4''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 25.00) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''5''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 7.00) AND (TRI_REDUCAO = 41.67) AND (CST = 020) THEN ''6''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 12.00) AND (TRI_REDUCAO = 30.00) AND (CST = 020) THEN ''40''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 12.00) AND (TRI_REDUCAO = 41.66) AND (CST = 020) THEN ''39''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 12.00) AND (TRI_REDUCAO = 41.67) AND (CST = 020) THEN ''6''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 26.11) AND (CST = 020) THEN ''41''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 33.33) AND (CST = 020) THEN ''7''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 38.88) AND (CST = 020) THEN ''33''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 61.11) AND (CST = 020) THEN ''8''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 0.00) AND (TRI_REDUCAO = 0.00) AND (CST = 040) THEN ''1''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 0.00) AND (TRI_REDUCAO = 0.00) AND (CST = 041) THEN ''23''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 0.00) AND (TRI_REDUCAO = 0.00) AND (CST = 060) THEN ''25''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 0.00) AND (TRI_REDUCAO = 0.00) AND (CST = 090) THEN ''22''         ');
   SQL.Add('       END AS COD_TRIB_SAIDA,               ');
   SQL.Add('       0 AS PER_IVA,               ');
   SQL.Add('       0 AS PER_FCP_ST               ');
   SQL.Add('   FROM dbo.NBM_PRODUTOS                ');
   SQL.Add('   LEFT JOIN dbo.PRODUTOS ON                ');
   SQL.Add('       dbo.NBM_PRODUTOS.NBM = dbo.PRODUTOS.PRO_CLASFISCAL                ');
   SQL.Add('   LEFT JOIN dbo.TRIBUTACAO ON                ');
   SQL.Add('       dbo.PRODUTOS.TRI_CODIGO = dbo.TRIBUTACAO.TRI_CODIGO               ');
   SQL.Add('   WHERE NBM IN (SELECT DISTINCT PRO_CLASFISCAL FROM dbo.PRODUTOS)         ');
   SQL.Add('   AND PRO_CODIGO = 3   ');
   SQL.Add('      ');
   (*SQL.Add('   UNION ALL   ');
   SQL.Add('   	   ');
   SQL.Add('   SELECT DISTINCT   ');
   SQL.Add('   0 AS COD_NCM,               ');
   SQL.Add('       ''ADENIFIR'' AS DES_NCM,          ');
   SQL.Add('       PRODUTOS.PRO_CLASFISCAL AS NUM_NCM,               ');
   SQL.Add('       CASE                ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = '''' ) AND (PRO_CST_COFINS = '''' ) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''N''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 50) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65)        THEN ''N''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 50) AND (PRO_CST_COFINS = 3) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''N''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 50) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA = 1.65) THEN ''N''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 60) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''N''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''N''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 4) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''S''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 4) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL) THEN ''S''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 06) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''S'' --S        ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 06) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''S'' --S         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''S''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''S''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 72) AND (PRO_CST_COFINS = 9) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''S''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 4) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''S''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 06) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''S''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 06) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''S''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''S''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 98) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''N''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 98) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''S''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 98) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''S''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''N''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''S''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''S''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''N''                ');
   SQL.Add('       END  AS FLG_NAO_PIS_COFINS,            ');
   SQL.Add('       CASE                ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = '''' ) AND (PRO_CST_COFINS = '''' ) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''-1''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 50) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65)        THEN ''-1''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 50) AND (PRO_CST_COFINS = 3) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''-1''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 50) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA = 1.65) THEN ''-1''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 60) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''-1''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''-1''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 4) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''1''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 4) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL) THEN ''1''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 6) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''0'' --0         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 6) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''-0'' --0         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''0''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''0''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 72) AND (PRO_CST_COFINS = 9) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''4''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 4) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''0''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 06) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''0''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 06) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''0''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''0''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 98) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''-1''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 98) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''3''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 98) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''3''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''-1''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''3''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''3''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''-1''          ');
   SQL.Add('       END AS TIPO_NAO_PIS_COFINS,               ');
   SQL.Add('       CASE                ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = '''' ) AND (PRO_CST_COFINS = '''' ) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 50) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65)        THEN ''''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 50) AND (PRO_CST_COFINS = 3) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 50) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA = 1.65) THEN ''''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 60) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 4) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN natr_codigo         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 4) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL) THEN natr_codigo         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 6) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN COALESCE (natr_codigo, ''999'')         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 6) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN COALESCE (natr_codigo, ''999'')         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN COALESCE (natr_codigo, ''999'')         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN COALESCE (natr_codigo, ''999'')         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 72) AND (PRO_CST_COFINS = 9) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN COALESCE (natr_codigo, ''999'')         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 4) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN COALESCE (natr_codigo, ''999'')         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 6) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN COALESCE (natr_codigo, ''999'')         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 6) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN COALESCE (natr_codigo, ''999'')         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN COALESCE (natr_codigo, ''999'')         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 98) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 98) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN COALESCE (natr_codigo, ''999'')         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 98) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN COALESCE (natr_codigo, ''999'')         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''''         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN COALESCE (natr_codigo, ''999'')         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN COALESCE (natr_codigo, ''999'')         ');
   SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''''                    ');
   SQL.Add('       END AS COD_TAB_SPED,               ');
   SQL.Add('       CASE         ');
   SQL.Add('         WHEN PRO_CEST  IS NULL OR PRO_CEST = '''' THEN ''99999999''         ');
   SQL.Add('         ELSE PRO_CEST         ');
   SQL.Add('       END AS NUM_CEST,            ');
   SQL.Add('       ''SP'' DES_SIGLA,               ');
   SQL.Add('       CASE                 ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 4.14 ) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''42''               ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 4.50) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''38''          ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 4.70) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''43''          ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 5.50) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''44''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 7.00) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''2''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 9.40) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''45''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 12.00) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''3''            ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 13.30) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''47''          ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 17.00) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''46''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''4''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 25.00) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''5''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 7.00) AND (TRI_REDUCAO = 41.67) AND (CST = 020) THEN ''6''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 12.00) AND (TRI_REDUCAO = 30.00) AND (CST = 020) THEN ''40''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 12.00) AND (TRI_REDUCAO = 41.66) AND (CST = 020) THEN ''39''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 12.00) AND (TRI_REDUCAO = 41.67) AND (CST = 020) THEN ''6''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 26.11) AND (CST = 020) THEN ''41''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 33.33) AND (CST = 020) THEN ''7''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 38.88) AND (CST = 020) THEN ''33''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 61.11) AND (CST = 020) THEN ''8''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 0.00) AND (TRI_REDUCAO = 0.00) AND (CST = 040) THEN ''1''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 0.00) AND (TRI_REDUCAO = 0.00) AND (CST = 041) THEN ''23''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 0.00) AND (TRI_REDUCAO = 0.00) AND (CST = 060) THEN ''25''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 0.00) AND (TRI_REDUCAO = 0.00) AND (CST = 090) THEN ''22''         ');
   SQL.Add('       END AS COD_TRIB_ENTRADA,               ');
   SQL.Add('       CASE                ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 4.14 ) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''42''               ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 4.50) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''38''          ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 4.70) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''43''          ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 5.50) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''44''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 7.00) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''2''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 9.40) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''45''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 12.00) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''3''            ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 13.30) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''47''          ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 17.00) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''46''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''4''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 25.00) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''5''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 7.00) AND (TRI_REDUCAO = 41.67) AND (CST = 020) THEN ''6''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 12.00) AND (TRI_REDUCAO = 30.00) AND (CST = 020) THEN ''40''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 12.00) AND (TRI_REDUCAO = 41.66) AND (CST = 020) THEN ''39''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 12.00) AND (TRI_REDUCAO = 41.67) AND (CST = 020) THEN ''6''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 26.11) AND (CST = 020) THEN ''41''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 33.33) AND (CST = 020) THEN ''7''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 38.88) AND (CST = 020) THEN ''33''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 61.11) AND (CST = 020) THEN ''8''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 0.00) AND (TRI_REDUCAO = 0.00) AND (CST = 040) THEN ''1''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 0.00) AND (TRI_REDUCAO = 0.00) AND (CST = 041) THEN ''23''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 0.00) AND (TRI_REDUCAO = 0.00) AND (CST = 060) THEN ''25''         ');
   SQL.Add('         WHEN (TRI_ALIQUOTA = 0.00) AND (TRI_REDUCAO = 0.00) AND (CST = 090) THEN ''22''         ');
   SQL.Add('       END AS COD_TRIB_SAIDA,               ');
   SQL.Add('       0 AS PER_IVA,               ');
   SQL.Add('       0 AS PER_FCP_ST               ');
   SQL.Add('   FROM dbo.PRODUTOS                             ');
   SQL.Add('   LEFT JOIN dbo.TRIBUTACAO ON                ');
   SQL.Add('     (dbo.PRODUTOS.TRI_CODIGO = dbo.TRIBUTACAO.TRI_CODIGO)   ');
   SQL.Add('   WHERE PRO_CLASFISCAL IN    ');
   SQL.Add('     (''06031000'',''20070000'',''34381010'',''38089090'',''41012010'',''44013000'',''48236000'',''54719012'',''63025900'',''70169990'',''74198090'',''79070000'',''79149100'',   ');
   SQL.Add('     	''81829119'',''82071990'',''84073000'',''84137000'',''84716021'',''84717040'',''84821000'',''85176250'',''85285200'',''85393139'',''85393990'',''85449000'',''87033300'',   ');
   SQL.Add('     			''94032010'',''94032090'',''97019100'')   ');
    *)



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


         if ZeradaOuVazia(Layout.FieldByName('DES_NCM').AsString) then
        begin
          Layout.FieldByName('DES_NCM').AsString := 'Á DEFINIR';
        end;

        //if (Layout.FieldByName('DES_NCM').AsString = '') then
         // Layout.FieldByName('DES_NCM').AsString := 'A DEFINIR'
       // else
         /// Layout.FieldByName('DES_NCM').AsString :=
          ///  Layout.FieldByName('DES_NCM').AsString;

        Layout.FieldByName('NUM_NCM').AsString :=
          StrRetNums(Layout.FieldByName('NUM_NCM').AsString);
        Layout.FieldByName('NUM_CEST').AsString :=
          StrRetNums(Layout.FieldByName('NUM_CEST').AsString);

        Layout.WriteLine;
      except
        On E: Exception do
          FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
      end;
      Next;
    end;
  end;
end;

procedure TFrmSmBomPrecoOurinhosGestora.GerarNCMUF;
var
  Count: Integer;
begin
  inherited;
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

 SQL.Add('   SELECT  DISTINCT         ');
 SQL.Add('       0 AS COD_NCM,               ');
 SQL.Add('       CASE          ');
 SQL.Add('         WHEN CAST(NBM_DESCRICAO AS VARCHAR(50)) = '''' THEN ''Á DEFINIR''         ');
 SQL.Add('         ELSE REPLACE (CAST(COALESCE(NBM_DESCRICAO, ''Á DEFINIR'') AS VARCHAR(50)), CHAR ( 13 ) +  Char ( 10 ), '''')        ');
 SQL.Add('       END AS DES_NCM,          ');
 SQL.Add('       COALESCE (NBM, PRODUTOS.PRO_CLASFISCAL) AS NUM_NCM,               ');
 SQL.Add('       CASE                ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = '''' ) AND (PRO_CST_COFINS = '''' ) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''N''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 50) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65)        THEN ''N''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 50) AND (PRO_CST_COFINS = 3) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''N''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 50) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA = 1.65) THEN ''N''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 60) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''N''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''N''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 4) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''S''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 4) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL) THEN ''S''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 06) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''S'' --S        ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 06) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''S'' --S         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''S''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''S''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 72) AND (PRO_CST_COFINS = 9) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''S''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 4) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''S''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 06) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''S''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 06) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''S''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''S''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 98) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''N''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 98) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''S''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 98) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''S''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''N''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''S''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''S''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''N''                ');
 SQL.Add('       END  AS FLG_NAO_PIS_COFINS,            ');
 SQL.Add('       CASE                ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = '''' ) AND (PRO_CST_COFINS = '''' ) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''-1''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 50) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65)        THEN ''-1''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 50) AND (PRO_CST_COFINS = 3) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''-1''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 50) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA = 1.65) THEN ''-1''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 60) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''-1''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''-1''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 4) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''1''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 4) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL) THEN ''1''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 6) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''0'' --0         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 6) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''-0'' --0         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''0''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''0''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 72) AND (PRO_CST_COFINS = 9) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''4''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 4) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''0''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 06) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''0''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 06) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''0''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''0''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 98) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''-1''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 98) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''3''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 98) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''3''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''-1''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''3''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''3''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''-1''          ');
 SQL.Add('       END AS TIPO_NAO_PIS_COFINS,               ');
 SQL.Add('       CASE                ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = '''' ) AND (PRO_CST_COFINS = '''' ) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 50) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65)        THEN ''''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 50) AND (PRO_CST_COFINS = 3) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 50) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA = 1.65) THEN ''''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 60) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 4) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN natr_codigo         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 4) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL) THEN natr_codigo         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 6) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN COALESCE (natr_codigo, ''999'')         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 6) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN COALESCE (natr_codigo, ''999'')         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN COALESCE (natr_codigo, ''999'')         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN COALESCE (natr_codigo, ''999'')         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 72) AND (PRO_CST_COFINS = 9) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN COALESCE (natr_codigo, ''999'')         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 4) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN COALESCE (natr_codigo, ''999'')         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 6) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN COALESCE (natr_codigo, ''999'')         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 6) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN COALESCE (natr_codigo, ''999'')         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN COALESCE (natr_codigo, ''999'')         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 98) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 98) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN COALESCE (natr_codigo, ''999'')         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 98) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN COALESCE (natr_codigo, ''999'')         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN COALESCE (natr_codigo, ''999'')         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN COALESCE (natr_codigo, ''999'')         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''''                    ');
 SQL.Add('       END AS COD_TAB_SPED,               ');
 SQL.Add('       CASE         ');
 SQL.Add('         WHEN PRO_CEST  IS NULL OR PRO_CEST = '''' THEN ''99999999''         ');
 SQL.Add('         ELSE PRO_CEST         ');
 SQL.Add('       END AS NUM_CEST,            ');
 SQL.Add('       ''SP'' DES_SIGLA,               ');
 SQL.Add('       CASE                 ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 4.14 ) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''42''               ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 4.50) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''38''          ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 4.70) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''43''          ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 5.50) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''44''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 7.00) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''2''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 9.40) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''45''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 12.00) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''3''            ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 13.30) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''47''          ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 17.00) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''46''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''4''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 25.00) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''5''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 7.00) AND (TRI_REDUCAO = 41.67) AND (CST = 020) THEN ''6''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 12.00) AND (TRI_REDUCAO = 30.00) AND (CST = 020) THEN ''40''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 12.00) AND (TRI_REDUCAO = 41.66) AND (CST = 020) THEN ''39''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 12.00) AND (TRI_REDUCAO = 41.67) AND (CST = 020) THEN ''6''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 26.11) AND (CST = 020) THEN ''41''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 33.33) AND (CST = 020) THEN ''7''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 38.88) AND (CST = 020) THEN ''33''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 61.11) AND (CST = 020) THEN ''8''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 0.00) AND (TRI_REDUCAO = 0.00) AND (CST = 040) THEN ''1''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 0.00) AND (TRI_REDUCAO = 0.00) AND (CST = 041) THEN ''23''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 0.00) AND (TRI_REDUCAO = 0.00) AND (CST = 060) THEN ''25''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 0.00) AND (TRI_REDUCAO = 0.00) AND (CST = 090) THEN ''22''         ');
 SQL.Add('       END AS COD_TRIB_ENTRADA,               ');
 SQL.Add('       CASE                ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 4.14 ) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''42''               ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 4.50) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''38''          ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 4.70) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''43''          ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 5.50) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''44''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 7.00) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''2''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 9.40) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''45''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 12.00) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''3''            ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 13.30) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''47''          ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 17.00) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''46''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''4''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 25.00) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''5''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 7.00) AND (TRI_REDUCAO = 41.67) AND (CST = 020) THEN ''6''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 12.00) AND (TRI_REDUCAO = 30.00) AND (CST = 020) THEN ''40''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 12.00) AND (TRI_REDUCAO = 41.66) AND (CST = 020) THEN ''39''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 12.00) AND (TRI_REDUCAO = 41.67) AND (CST = 020) THEN ''6''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 26.11) AND (CST = 020) THEN ''41''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 33.33) AND (CST = 020) THEN ''7''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 38.88) AND (CST = 020) THEN ''33''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 61.11) AND (CST = 020) THEN ''8''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 0.00) AND (TRI_REDUCAO = 0.00) AND (CST = 040) THEN ''1''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 0.00) AND (TRI_REDUCAO = 0.00) AND (CST = 041) THEN ''23''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 0.00) AND (TRI_REDUCAO = 0.00) AND (CST = 060) THEN ''25''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 0.00) AND (TRI_REDUCAO = 0.00) AND (CST = 090) THEN ''22''         ');
 SQL.Add('       END AS COD_TRIB_SAIDA,               ');
 SQL.Add('       0 AS PER_IVA,               ');
 SQL.Add('       0 AS PER_FCP_ST               ');
 SQL.Add('   FROM dbo.NBM_PRODUTOS                ');
 SQL.Add('   INNER JOIN dbo.PRODUTOS ON                ');
 SQL.Add('       dbo.NBM_PRODUTOS.NBM = dbo.PRODUTOS.PRO_CLASFISCAL                ');
 SQL.Add('   INNER JOIN dbo.TRIBUTACAO ON                ');
 SQL.Add('       dbo.PRODUTOS.TRI_CODIGO = dbo.TRIBUTACAO.TRI_CODIGO               ');
 SQL.Add('   WHERE NBM IN (SELECT DISTINCT PRO_CLASFISCAL FROM dbo.PRODUTOS)         ');
 SQL.Add('   AND PRO_CODIGO = 3    ');
 SQL.Add('      ');
 (*SQL.Add('   UNION ALL   ');
 SQL.Add('   	   ');
 SQL.Add('   SELECT DISTINCT   ');
 SQL.Add('   0 AS COD_NCM,               ');
 SQL.Add('       ''ADENIFIR'' AS DES_NCM,          ');
 SQL.Add('       PRODUTOS.PRO_CLASFISCAL AS NUM_NCM,               ');
 SQL.Add('       CASE                ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = '''' ) AND (PRO_CST_COFINS = '''' ) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''N''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 50) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65)        THEN ''N''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 50) AND (PRO_CST_COFINS = 3) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''N''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 50) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA = 1.65) THEN ''N''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 60) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''N''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''N''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 4) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''S''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 4) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL) THEN ''S''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 06) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''S'' --S        ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 06) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''S'' --S         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''S''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''S''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 72) AND (PRO_CST_COFINS = 9) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''S''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 4) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''S''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 06) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''S''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 06) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''S''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''S''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 98) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''N''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 98) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''S''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 98) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''S''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''N''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''S''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''S''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''N''                ');
 SQL.Add('       END  AS FLG_NAO_PIS_COFINS,            ');
 SQL.Add('       CASE                ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = '''' ) AND (PRO_CST_COFINS = '''' ) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''-1''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 50) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65)        THEN ''-1''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 50) AND (PRO_CST_COFINS = 3) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''-1''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 50) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA = 1.65) THEN ''-1''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 60) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''-1''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''-1''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 4) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''1''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 4) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL) THEN ''1''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 6) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''0'' --0         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 6) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''-0'' --0         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''0''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''0''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 72) AND (PRO_CST_COFINS = 9) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''4''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 4) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''0''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 06) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''0''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 06) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''0''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''0''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 98) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''-1''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 98) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''3''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 98) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''3''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''-1''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''3''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''3''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''-1''          ');
 SQL.Add('       END AS TIPO_NAO_PIS_COFINS,               ');
 SQL.Add('       CASE                ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = '''' ) AND (PRO_CST_COFINS = '''' ) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 50) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65)        THEN ''''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 50) AND (PRO_CST_COFINS = 3) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 50) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA = 1.65) THEN ''''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 60) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 4) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN natr_codigo         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 4) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL) THEN natr_codigo         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 6) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN COALESCE (natr_codigo, ''999'')         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 6) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN COALESCE (natr_codigo, ''999'')         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN COALESCE (natr_codigo, ''999'')         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN COALESCE (natr_codigo, ''999'')         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 72) AND (PRO_CST_COFINS = 9) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN COALESCE (natr_codigo, ''999'')         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 4) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN COALESCE (natr_codigo, ''999'')         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 6) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN COALESCE (natr_codigo, ''999'')         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 6) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN COALESCE (natr_codigo, ''999'')         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN COALESCE (natr_codigo, ''999'')         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 98) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 98) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN COALESCE (natr_codigo, ''999'')         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 98) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN COALESCE (natr_codigo, ''999'')         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''''         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN COALESCE (natr_codigo, ''999'')         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN COALESCE (natr_codigo, ''999'')         ');
 SQL.Add('         WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''''                    ');
 SQL.Add('       END AS COD_TAB_SPED,               ');
 SQL.Add('       CASE         ');
 SQL.Add('         WHEN PRO_CEST  IS NULL OR PRO_CEST = '''' THEN ''99999999''         ');
 SQL.Add('         ELSE PRO_CEST         ');
 SQL.Add('       END AS NUM_CEST,            ');
 SQL.Add('       ''SP'' DES_SIGLA,               ');
 SQL.Add('       CASE                 ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 4.14 ) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''42''               ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 4.50) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''38''          ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 4.70) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''43''          ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 5.50) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''44''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 7.00) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''2''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 9.40) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''45''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 12.00) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''3''            ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 13.30) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''47''          ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 17.00) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''46''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''4''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 25.00) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''5''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 7.00) AND (TRI_REDUCAO = 41.67) AND (CST = 020) THEN ''6''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 12.00) AND (TRI_REDUCAO = 30.00) AND (CST = 020) THEN ''40''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 12.00) AND (TRI_REDUCAO = 41.66) AND (CST = 020) THEN ''39''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 12.00) AND (TRI_REDUCAO = 41.67) AND (CST = 020) THEN ''6''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 26.11) AND (CST = 020) THEN ''41''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 33.33) AND (CST = 020) THEN ''7''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 38.88) AND (CST = 020) THEN ''33''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 61.11) AND (CST = 020) THEN ''8''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 0.00) AND (TRI_REDUCAO = 0.00) AND (CST = 040) THEN ''1''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 0.00) AND (TRI_REDUCAO = 0.00) AND (CST = 041) THEN ''23''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 0.00) AND (TRI_REDUCAO = 0.00) AND (CST = 060) THEN ''25''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 0.00) AND (TRI_REDUCAO = 0.00) AND (CST = 090) THEN ''22''         ');
 SQL.Add('       END AS COD_TRIB_ENTRADA,               ');
 SQL.Add('       CASE                ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 4.14 ) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''42''               ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 4.50) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''38''          ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 4.70) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''43''          ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 5.50) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''44''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 7.00) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''2''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 9.40) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''45''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 12.00) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''3''            ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 13.30) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''47''          ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 17.00) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''46''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''4''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 25.00) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''5''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 7.00) AND (TRI_REDUCAO = 41.67) AND (CST = 020) THEN ''6''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 12.00) AND (TRI_REDUCAO = 30.00) AND (CST = 020) THEN ''40''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 12.00) AND (TRI_REDUCAO = 41.66) AND (CST = 020) THEN ''39''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 12.00) AND (TRI_REDUCAO = 41.67) AND (CST = 020) THEN ''6''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 26.11) AND (CST = 020) THEN ''41''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 33.33) AND (CST = 020) THEN ''7''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 38.88) AND (CST = 020) THEN ''33''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 61.11) AND (CST = 020) THEN ''8''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 0.00) AND (TRI_REDUCAO = 0.00) AND (CST = 040) THEN ''1''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 0.00) AND (TRI_REDUCAO = 0.00) AND (CST = 041) THEN ''23''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 0.00) AND (TRI_REDUCAO = 0.00) AND (CST = 060) THEN ''25''         ');
 SQL.Add('         WHEN (TRI_ALIQUOTA = 0.00) AND (TRI_REDUCAO = 0.00) AND (CST = 090) THEN ''22''         ');
 SQL.Add('       END AS COD_TRIB_SAIDA,               ');
 SQL.Add('       0 AS PER_IVA,               ');
 SQL.Add('       0 AS PER_FCP_ST               ');
 SQL.Add('   FROM dbo.PRODUTOS                             ');
 SQL.Add('   INNER JOIN dbo.TRIBUTACAO ON                ');
 SQL.Add('     (dbo.PRODUTOS.TRI_CODIGO = dbo.TRIBUTACAO.TRI_CODIGO)   ');
 SQL.Add('   WHERE PRO_CLASFISCAL IN    ');
 SQL.Add('     (''06031000'',''20070000'',''34381010'',''38089090'',''41012010'',''44013000'',''48236000'',''54719012'',''63025900'',''70169990'',''74198090'',''79070000'',''79149100'',   ');
 SQL.Add('     	''81829119'',''82071990'',''84073000'',''84137000'',''84716021'',''84717040'',''84821000'',''85176250'',''85285200'',''85393139'',''85393990'',''85449000'',''87033300'',   ');
 SQL.Add('     			''94032010'',''94032090'',''97019100'')   ');


    *)


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

        //if ZeradaOuVazia(Layout.FieldByName('DES_NCM').AsString) then
        //begin
        //  Layout.FieldByName('DES_NCM').AsString := 'Á DEFINIR';
        //end;

        Layout.FieldByName('COD_NCM').AsInteger := Count;



        Layout.WriteLine;
      except
        On E: Exception do
          FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
      end;
      Next;
    end;
  end;
end;

procedure TFrmSmBomPrecoOurinhosGestora.GerarProduto;
begin
  inherited;

  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;
(*
   
*)

 SQL.Add('          SELECT                  ');
 SQL.Add('            	PRO_CODIGO AS COD_PRODUTO,                  ');
 SQL.Add('            	CASE                  ');
 SQL.Add('           	 	WHEN LEN(PRO_BARRA) < 8 THEN PRO_CODIGO                 ');
 SQL.Add('           	 	ELSE PRO_BARRA                  ');
 SQL.Add('            	END  AS COD_BARRA_PRINCIPAL,                  ');
 SQL.Add('            	REPLACE (REPLACE (REPLACE (PRO_DESCRICAO2, ''Ã'', ''A''), ''É'', ''E''), ''Ç'', ''C'') AS DES_REDUZIDA,                  ');
 SQL.Add('            	REPLACE (REPLACE (REPLACE (PRO_DESCRICAO, ''Ã'', ''A''), ''É'', ''E''), ''Ç'', ''C'') AS DES_PRODUTO,                  ');
 SQL.Add('            	PRO_EMBALAGEM AS QTD_EMBALAGEM_COMPRA,                  ');
 SQL.Add('            	CASE                ');
 SQL.Add('           	 	WHEN PRO_UNIDADE = ''UNID'' THEN ''UN''            ');
 SQL.Add('           	 	WHEN PRO_UNIDADE = ''UND'' THEN ''UN''            ');
 SQL.Add('           	 	ELSE PRO_UNIDADE               ');
 SQL.Add('            	END AS DES_UNIDADE_COMPRA,                  ');
 SQL.Add('            	PRO_QTDE_BAIXAR AS QTD_EMBALAGEM_VENDA,                  ');
 SQL.Add('            	CASE                ');
 SQL.Add('           	 	WHEN PRO_UNID_REF = ''UNID'' THEN ''UN''            ');
 SQL.Add('           	 	WHEN PRO_UNID_REF = ''UND'' THEN ''UN''            ');
 SQL.Add('           	 	WHEN PRO_UNID_REF = ''KG'' THEN ''KG''            ');
 SQL.Add('           	 	WHEN PRO_UNID_REF = '''' THEN ''UN''      ');
 SQL.Add('           	 	ELSE ''UN''            ');
 SQL.Add('            	END AS DES_UNIDADE_VENDA,                  ');
 SQL.Add('            	0 AS TIPO_IPI ,                  ');
 SQL.Add('            	0 AS VAL_IPI,                  ');
 SQL.Add('              CASE    ');
 SQL.Add('   		        WHEN (dbo.PRODUTOS.DEP_CODIGO = 75) OR (dbo.PRODUTOS.DEP_CODIGO = 0) THEN 999   ');
 SQL.Add('   		        ELSE COALESCE (dbo.PRODUTOS.DEP_CODIGO, 999)    ');
 SQL.Add('   	          END AS COD_SECAO, ');
 SQL.Add('           	  CASE             ');
 SQL.Add('           	 	WHEN GRU_CODIGO = 0 THEN 999             ');
 SQL.Add('           	 	ELSE COALESCE (GRU_CODIGO, 999)             ');
 SQL.Add('            	END AS COD_GRUPO,                  ');
 SQL.Add('            	CASE             ');
 SQL.Add('           	 	WHEN SUB_CODIGO = 0 THEN 999            ');
 SQL.Add('           	 	ELSE COALESCE (SUB_CODIGO, 999)            ');
 SQL.Add('            	END AS COD_SUB_GRUPO,                  ');
 SQL.Add('            	0 AS COD_PRODUTO_SIMILAR,                  ');
 SQL.Add('            	CASE	                  ');
 SQL.Add('           	 	WHEN PRO_UNIDADE = ''KG'' THEN ''S''                  ');
 SQL.Add('           	 	ELSE ''N''                  ');
 SQL.Add('            	END AS IPV,                  ');
 SQL.Add('            	COALESCE (PRO_VALIDADE, 0) AS DIAS_VALIDADE,                  ');
 SQL.Add('            	0 AS TIPO_PRODUTO,                  ');
 SQL.Add('            	CASE             ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = '''' ) AND (PRO_CST_COFINS = '''' ) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''N''      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 50) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65)        THEN ''N''      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 50) AND (PRO_CST_COFINS = 3) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''N''      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 50) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA = 1.65) THEN ''N''      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 60) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''N''      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''N''      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 4) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''S''      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 4) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL) THEN ''S''      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 6) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''S''      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 6) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''S''      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''S''      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''S''      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 72) AND (PRO_CST_COFINS = 9) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''S''      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 4) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''S''      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 6) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''S''      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 6) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''S''      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''S''      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 98) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''N''      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 98) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''S''      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 98) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''S''      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''N''      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''S''      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''S''      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''N''            ');
 SQL.Add('            	END  AS FLG_NAO_PIS_COFINS,                   ');
 SQL.Add('            	PRO_BALANCA AS FLG_ENVIA_BALANCA,                  ');
 SQL.Add('            	CASE                      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = '''' ) AND (PRO_CST_COFINS = '''' ) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''-1''      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 50) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65)        THEN ''-1''      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 50) AND (PRO_CST_COFINS = 3) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''-1''      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 50) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA = 1.65) THEN ''-1''      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 60) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''-1''      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''-1''      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 4) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''1''      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 4) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL) THEN ''1''      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 6) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''0''      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 6) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''0''      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''0''      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''0''      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 72) AND (PRO_CST_COFINS = 9) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''4''      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 4) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''0''      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 6) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''0''      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 6) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''0''      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''0''      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 98) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''-1''      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 98) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''3''      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 98) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''3''      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''-1''      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN ''3''      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''3''      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''-1''                    ');
 SQL.Add('            	END AS TIPO_NAO_PIS_COFINS,                  ');
 SQL.Add('            	0 AS TIPO_EVENTO,                  ');
 SQL.Add('            	COALESCE (PRO_PAI, 0) AS COD_ASSOCIADO,                  ');
 SQL.Add('            	'''' AS DES_OBSERVACAO,                  ');
 SQL.Add('            	COALESCE (NUT_CODIGO, 0) AS COD_INFO_NUTRICIONAL,                  ');
 SQL.Add('            	CASE                      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = '''' ) AND (PRO_CST_COFINS = '''' ) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN ''''      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 50) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65)        THEN ''''      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 50) AND (PRO_CST_COFINS = 3) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''''      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 50) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA = 1.65) THEN ''''      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 60) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''''      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''''      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 4) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN natr_codigo      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 4) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL) THEN natr_codigo      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 6) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN COALESCE (natr_codigo, ''999'')      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 6) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN COALESCE (natr_codigo, ''999'')      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN COALESCE (natr_codigo, ''999'')      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 70) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN COALESCE (natr_codigo, ''999'')      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 72) AND (PRO_CST_COFINS = 9) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN COALESCE (natr_codigo, ''999'')      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 4) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN COALESCE (natr_codigo, ''999'')      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 6) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN COALESCE (natr_codigo, ''999'')      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 6) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN COALESCE (natr_codigo, ''999'')      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 73) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN COALESCE (natr_codigo, ''999'')      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 98) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''''      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 98) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN COALESCE (natr_codigo, ''999'')      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 98) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN COALESCE (natr_codigo, ''999'')      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 1) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''''      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 0.00) AND (PRO_PIS_ENTRADA = 0.00) THEN COALESCE (natr_codigo, ''999'')      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS IS NULL ) AND (PRO_PIS_ENTRADA IS NULL ) THEN COALESCE (natr_codigo, ''999'')      ');
 SQL.Add('           	 	WHEN (PRO_CST_COFINS_ENTRADA = 99) AND (PRO_CST_COFINS = 49) AND (PRO_COFINS = 7.60) AND (PRO_PIS_ENTRADA = 1.65) THEN ''''                     ');
 SQL.Add('            	END AS COD_TAB_SPED,                  ');
 SQL.Add('            	''N'' AS FLG_ALCOOLICO,                      ');
 SQL.Add('            	0 AS TIPO_ESPECIE,                  ');
 SQL.Add('            	0 AS  COD_CLASSIF,                  ');
 SQL.Add('            	1 AS VAL_VDA_PESO_BRUTO,                  ');
 SQL.Add('            	1 AS VAL_PESO_EMB,                  ');
 SQL.Add('            	0 AS TIPO_EXPLOSAO_COMPRA,                  ');
 SQL.Add('            	'''' AS DTA_INI_OPER,                        ');
 SQL.Add('            	'''' AS DES_PLAQUETA,                        ');
 SQL.Add('            	'''' AS MES_ANO_INI_DEPREC,                        ');
 SQL.Add('            	0 AS TIPO_BEM,                  ');
 SQL.Add('            	CASE             ');
 SQL.Add('           	 	WHEN (FOR_CODIGO IS NULL) OR (FOR_CODIGO = 0) THEN 0             ');
 SQL.Add('           	 	ELSE FOR_CODIGO             ');
 SQL.Add('            	END AS COD_FORNECEDOR,                        ');
 SQL.Add('            	0 AS NUM_NF,                        ');
 SQL.Add('            	CONVERT (CHAR, PRO_DATA_CADASTRO, 103) AS DTA_ENTRADA,                        ');
 SQL.Add('            	0 AS COD_NAT_BEM,                        ');
 SQL.Add('            	0 AS VAL_ORIG_BEM,                        ');
 SQL.Add('            	COALESCE (REPLACE (REPLACE (REPLACE (PRO_DESCRICAO, ''Ã'', ''A''), ''É'', ''E''), ''Ç'', ''C''), ''A DEFINIR'') AS DES_PRODUTO_ANT	                    ');
 SQL.Add('           FROM dbo.PRODUTOS                  ');
 SQL.Add('           WHERE PRO_DESCRICAO IS NOT NULL       ');
 SQL.Add('   AND PRO_CODIGO = 3    ');
 SQL.Add('   		     ORDER BY PRO_CODIGO ASC   ');


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

        Layout.FieldByName('COD_PRODUTO').AsString :=
          GerarPLU(Layout.FieldByName('COD_PRODUTO').AsString);

        Layout.FieldByName('COD_ASSOCIADO').AsString :=
          GerarPLU(Layout.FieldByName('COD_ASSOCIADO').AsString);

        // Alterar a palavra 'KG' para 'kg' no campo DES_REDUZIDA
        Layout.FieldByName('DES_REDUZIDA').AsString :=
          StrReplace(UpperCase(Layout.FieldByName('DES_REDUZIDA').AsString),
          'KG', 'kg');

        // Substituir Letras Acentuadas
        Layout.FieldByName('DES_REDUZIDA').AsString :=
          StrSubstLtsAct(Layout.FieldByName('DES_REDUZIDA').AsString);
        Layout.FieldByName('DES_PRODUTO').AsString :=
          StrSubstLtsAct(Layout.FieldByName('DES_PRODUTO').AsString);

        if Length(StrLBReplace
          (Trim(StrRetNums(FieldByName('COD_BARRA_PRINCIPAL').AsString)))) < 8
        then
          Layout.FieldByName('COD_BARRA_PRINCIPAL').AsString :=
            GerarPLU(FieldByName('COD_BARRA_PRINCIPAL').AsString);

        if not CodBarrasValido(Layout.FieldByName('COD_BARRA_PRINCIPAL')
          .AsString) then
          Layout.FieldByName('COD_BARRA_PRINCIPAL').AsString := '';

        (*
          if QryPrincipal.FieldByName('DTA_ENTRADA').AsString <> '' then
          Layout.FieldByName('DTA_ENTRADA').AsDateTime := FieldByName('DTA_ENTRADA').AsDateTime;

          Layout.FieldByName('COD_PRODUTO').AsString := GerarPLU(QryPrincipal2.FieldByName('COD_PRODUTO').AsString);

          Layout.FieldByName('DES_OBSERVACAO').AsString  := StrReplace(StrLBReplace(FieldByName('DES_OBSERVACAO').AsString), '\n', '');
          Layout.FieldByName('DES_REDUZIDA').AsString    := StrRemPont(StrReplace(StrLBReplace(FieldByName('DES_REDUZIDA').AsString), '\n', ''));
          Layout.FieldByName('DES_PRODUTO').AsString     := StrRemPont(StrReplace(StrLBReplace(FieldByName('DES_PRODUTO').AsString), '\n', ''));
          Layout.FieldByName('DES_PRODUTO_ANT').AsString := StrRemPont(StrReplace(StrLBReplace(FieldByName('DES_PRODUTO').AsString), '\n', ''));

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
        *)

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

procedure TFrmSmBomPrecoOurinhosGestora.GerarCodigoBarras;
var
  Count, count1: Integer;
  CodPrincipal: string;
begin
  inherited;
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('   SELECT    ');
    SQL.Add('   	TCONSULTA.PRO_CODIGO AS COD_PRODUTO,   ');
    SQL.Add('   	TCONSULTA.PRO_BARRA AS COD_EAN   ');
    SQL.Add('   FROM TCONSULTA INNER JOIN PRODUTOS ON   ');
    SQL.Add('   	(TCONSULTA.PRO_CODIGO = PRODUTOS.PRO_CODIGO)   ');
    SQL.Add('   WHERE TCONSULTA.PRO_BARRA <> 0   ');
    SQL.Add('   AND LEN(TCONSULTA.PRO_BARRA) >= 8   ');
    SQL.Add('   AND PRODUTOS.PRO_DESCRICAO IS NOT NULL   ');
   // SQL.Add('   AND TCONSULTA.PRO_BARRA <> PRODUTOS.PRO_BARRA   ');
    SQL.Add('      ');
    SQL.Add('   ORDER BY    ');
    SQL.Add('   	TCONSULTA.PRO_CODIGO ASC;   ');

    (*
      SQL.Add('   SELECT    ');
      SQL.Add('   	PRO_CODIGO AS COD_PRODUTO,   ');
      SQL.Add('   	CASE            ');
      SQL.Add('   		WHEN LEN(PRO_BARRA) < 8 THEN PRO_CODIGO    ');
      SQL.Add('           ELSE PRO_BARRA            ');
      SQL.Add('       END  AS COD_EAN       ');
      SQL.Add('   		   ');
      SQL.Add('   FROM dbo.PRODUTOS   ');
      SQL.Add('   WHERE PRO_DESCRICAO IS NOT NULL    ');
      SQL.Add('   ORDER BY          ');
      SQL.Add('   		PRO_CODIGO ASC          ');
    *)

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

        Layout.FieldByName('COD_PRODUTO').AsString :=
          GerarPLU(QryPrincipal2.FieldByName('COD_PRODUTO').AsString);

        // CodPrincipal := FieldByName('COD_EAN').AsString;

        { if ((StrPad(FieldByName('COD_EAN').AsString,13,'0','L')) =
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
          end; }

        if Length(StrLBReplace(Trim(StrRetNums(FieldByName('COD_EAN').AsString))
          )) < 8 then
          Layout.FieldByName('COD_EAN').AsString :=
            GerarPLU(FieldByName('COD_EAN').AsString);

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

procedure TFrmSmBomPrecoOurinhosGestora.GerarProdLoja;
begin
  inherited;
  (*
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
    end; *)

  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;
 (*
 
  *)
   SQL.Add('      SELECT            ');
 SQL.Add('              PRO_CODIGO AS COD_PRODUTO,            ');
  SQL.Add('                  CASE    ');
 SQL.Add('   				WHEN PRO_CUSTOREAL >= 100000 THEN 0   ');
 SQL.Add('   				ELSE PRO_CUSTOREAL   ');
 SQL.Add('   				END AS VAL_CUSTO_REP,                           ');
 SQL.Add('                  CASE    ');
 SQL.Add('   				WHEN PRO_VENDA >= 100000 THEN 0   ');
 SQL.Add('   				ELSE PRO_VENDA   ');
 SQL.Add('   				END AS VAL_VENDA,   ');

 SQL.Add('              PRO_VENDAPRO AS VAL_OFERTA,            ');
 SQL.Add('              PRO_ESTOQUE AS QTD_EST_VDA,            ');
 SQL.Add('              '''' AS TECLA_BALANCA,            ');
 SQL.Add('              CASE                   ');
 SQL.Add('           	   WHEN (TRI_ALIQUOTA = 4.14 ) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''42''            ');
 SQL.Add('           	   WHEN (TRI_ALIQUOTA = 4.50) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''38''     ');
 SQL.Add('           	   WHEN (TRI_ALIQUOTA = 4.70) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''43''       ');
 SQL.Add('           	   WHEN (TRI_ALIQUOTA = 5.50) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''44''      ');
 SQL.Add('           	   WHEN (TRI_ALIQUOTA = 7.00) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''2''      ');
 SQL.Add('           	   WHEN (TRI_ALIQUOTA = 9.40) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''45''      ');
 SQL.Add('           	   WHEN (TRI_ALIQUOTA = 12.00) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''3''         ');
 SQL.Add('           	   WHEN (TRI_ALIQUOTA = 13.30) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''47''       ');
 SQL.Add('           	   WHEN (TRI_ALIQUOTA = 17.00) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''46''      ');
 SQL.Add('           	   WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''4''      ');
 SQL.Add('           	   WHEN (TRI_ALIQUOTA = 25.00) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''5''      ');
 SQL.Add('           	   WHEN (TRI_ALIQUOTA = 7.00) AND (TRI_REDUCAO = 41.67) AND (CST = 020) THEN ''6''      ');
 SQL.Add('           	   WHEN (TRI_ALIQUOTA = 12.00) AND (TRI_REDUCAO = 30.00) AND (CST = 020) THEN ''40''      ');
 SQL.Add('           	   WHEN (TRI_ALIQUOTA = 12.00) AND (TRI_REDUCAO = 41.66) AND (CST = 020) THEN ''39''      ');
 SQL.Add('           	   WHEN (TRI_ALIQUOTA = 12.00) AND (TRI_REDUCAO = 41.67) AND (CST = 020) THEN ''6''      ');
 SQL.Add('           	   WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 26.11) AND (CST = 020) THEN ''41''      ');
 SQL.Add('           	   WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 33.33) AND (CST = 020) THEN ''7''      ');
 SQL.Add('           	   WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 38.88) AND (CST = 020) THEN ''33''      ');
 SQL.Add('           	   WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 61.11) AND (CST = 020) THEN ''8''      ');
 SQL.Add('           	   WHEN (TRI_ALIQUOTA = 0.00) AND (TRI_REDUCAO = 0.00) AND (CST = 040) THEN ''1''      ');
 SQL.Add('           	   WHEN (TRI_ALIQUOTA = 0.00) AND (TRI_REDUCAO = 0.00) AND (CST = 041) THEN ''23''      ');
 SQL.Add('           	   WHEN (TRI_ALIQUOTA = 0.00) AND (TRI_REDUCAO = 0.00) AND (CST = 060) THEN ''25''      ');
 SQL.Add('           	   WHEN (TRI_ALIQUOTA = 0.00) AND (TRI_REDUCAO = 0.00) AND (CST = 090) THEN ''22''              ');
 SQL.Add('              END AS COD_TRIBUTACAO,            ');
 SQL.Add('              PRO_MARGEM AS VAL_MARGEM,            ');
 SQL.Add('              1 AS QTD_ETIQUETA,            ');
 SQL.Add('              CASE                   ');
 SQL.Add('           	   WHEN (TRI_ALIQUOTA = 4.14 ) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''42''            ');
 SQL.Add('           	   WHEN (TRI_ALIQUOTA = 4.50) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''38''       ');
 SQL.Add('           	   WHEN (TRI_ALIQUOTA = 4.70) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''43''       ');
 SQL.Add('           	   WHEN (TRI_ALIQUOTA = 5.50) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''44''      ');
 SQL.Add('           	   WHEN (TRI_ALIQUOTA = 7.00) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''2''      ');
 SQL.Add('           	   WHEN (TRI_ALIQUOTA = 9.40) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''45''      ');
 SQL.Add('           	   WHEN (TRI_ALIQUOTA = 12.00) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''3''         ');
 SQL.Add('           	   WHEN (TRI_ALIQUOTA = 13.30) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''47''       ');
 SQL.Add('           	   WHEN (TRI_ALIQUOTA = 17.00) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''46''      ');
 SQL.Add('           	   WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''4''      ');
 SQL.Add('           	   WHEN (TRI_ALIQUOTA = 25.00) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''5''      ');
 SQL.Add('           	   WHEN (TRI_ALIQUOTA = 7.00) AND (TRI_REDUCAO = 41.67) AND (CST = 020) THEN ''6''      ');
 SQL.Add('           	   WHEN (TRI_ALIQUOTA = 12.00) AND (TRI_REDUCAO = 30.00) AND (CST = 020) THEN ''40''      ');
 SQL.Add('           	   WHEN (TRI_ALIQUOTA = 12.00) AND (TRI_REDUCAO = 41.66) AND (CST = 020) THEN ''39''      ');
 SQL.Add('           	   WHEN (TRI_ALIQUOTA = 12.00) AND (TRI_REDUCAO = 41.67) AND (CST = 020) THEN ''6''      ');
 SQL.Add('           	   WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 26.11) AND (CST = 020) THEN ''41''      ');
 SQL.Add('           	   WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 33.33) AND (CST = 020) THEN ''7''      ');
 SQL.Add('           	   WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 38.88) AND (CST = 020) THEN ''33''      ');
 SQL.Add('           	   WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 61.11) AND (CST = 020) THEN ''8''      ');
 SQL.Add('           	   WHEN (TRI_ALIQUOTA = 0.00) AND (TRI_REDUCAO = 0.00) AND (CST = 040) THEN ''1''      ');
 SQL.Add('           	   WHEN (TRI_ALIQUOTA = 0.00) AND (TRI_REDUCAO = 0.00) AND (CST = 041) THEN ''23''      ');
 SQL.Add('           	   WHEN (TRI_ALIQUOTA = 0.00) AND (TRI_REDUCAO = 0.00) AND (CST = 060) THEN ''25''      ');
 SQL.Add('           	   WHEN (TRI_ALIQUOTA = 0.00) AND (TRI_REDUCAO = 0.00) AND (CST = 090) THEN ''22''                           ');
 SQL.Add('              END AS COD_TRIB_ENTRADA,            ');
 SQL.Add('              CASE            ');
 SQL.Add('           	   WHEN PRO_MIX = ''S'' THEN ''N''            ');
 SQL.Add('           	   WHEN PRO_MIX = ''N'' THEN ''S''      ');
 SQL.Add('           	   ELSE ''N''           ');
 SQL.Add('              END FLG_INATIVO,            ');
 SQL.Add('              PRO_CODIGO AS COD_PRODUTO_ANT,            ');
 SQL.Add('              COALESCE (PRO_CLASFISCAL, ''99999999'') AS NUM_NCM,            ');
 SQL.Add('              0 AS TIPO_NCM,            ');
 SQL.Add('              0 AS VAL_VENDA_2,            ');
 SQL.Add('              '''' AS DTA_VALIDA_OFERTA,            ');
 SQL.Add('              1 AS QTD_EST_MINIMO,            ');
 SQL.Add('              NULL AS COD_VASILHAME,            ');
 SQL.Add('              ''N'' AS FORA_LINHA,            ');
 SQL.Add('              0 AS QTD_PRECO_DIF,            ');
 SQL.Add('              0 AS VAL_FORCA_VDA,            ');
 SQL.Add('              CASE      ');
 SQL.Add('           	   WHEN PRO_CEST  IS NULL OR PRO_CEST = '''' THEN ''99999999''      ');
 SQL.Add('           	   ELSE PRO_CEST      ');
 SQL.Add('              END AS NUM_CEST,            ');
 SQL.Add('              0 AS PER_IVA,            ');
 SQL.Add('              0 AS PER_FCP_ST,            ');
 SQL.Add('              0 AS PER_FIDELIDADE,            ');
 SQL.Add('              COALESCE (INF_EXT_CODIGO, 0) AS COD_INFO_RECEITA,    ');
 SQL.Add('              COALESCE (PRO_PAI , 0) AS COD_ASSOCIADO      ');
 SQL.Add('                 ');
 SQL.Add('           FROM dbo.PRODUTOS                ');
 SQL.Add('           INNER JOIN dbo.TRIBUTACAO ON                ');
 SQL.Add('              (dbo.PRODUTOS.TRI_CODIGO = dbo.TRIBUTACAO.TRI_CODIGO)            ');
 SQL.Add('           WHERE PRO_DESCRICAO IS NOT NULL       ');
 SQL.Add('   AND PRO_CODIGO = 3    ');
 SQL.Add('   		     ORDER BY PRO_CODIGO ASC   ');



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

        Layout.FieldByName('COD_PRODUTO').AsString :=
          GerarPLU(QryPrincipal2.FieldByName('COD_PRODUTO').AsString);
        Layout.FieldByName('COD_PRODUTO').AsString :=
          Layout.FieldByName('COD_PRODUTO').AsString;
        Layout.FieldByName('NUM_NCM').AsString :=
          StrRetNums(Layout.FieldByName('NUM_NCM').AsString);
        Layout.FieldByName('NUM_CEST').AsString :=
          StrRetNums(Layout.FieldByName('NUM_CEST').AsString);
        // StrToInt(Layout.FieldByName('COD_TRIB_ENTRADA').AsString); //:= (Layout.FieldByName('COD_TRIB_ENTRADA').AsInteger);s



        // if QryPrincipal2.FieldByName('COD_INFO_RECEITA').AsInteger = 0 then
        // Layout.FieldByName('COD_INFO_RECEITA').AsInteger  := -1;

        Layout.FieldByName('COD_ASSOCIADO').AsString :=
          GerarPLU(Layout.FieldByName('COD_ASSOCIADO').AsString);

        if QryPrincipal2.FieldByName('DTA_VALIDA_OFERTA').AsString <> '' then
          Layout.FieldByName('DTA_VALIDA_OFERTA').AsString :=
            FORMATDATETIME('dd/mm/yyyy',
            QryPrincipal2.FieldByName('DTA_VALIDA_OFERTA').AsDateTime);

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

procedure TFrmSmBomPrecoOurinhosGestora.GerarProdForn;
begin
  inherited;

  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('   SELECT DISTINCT   ');
    SQL.Add('   	dbo.PRODUTO_FOR.PRO_CODIGO AS COD_PRODUTO,   ');
    SQL.Add('   	dbo.PRODUTO_FOR.FOR_CODIGO AS COD_FORNECEDOR,   ');
    SQL.Add('   	CASE   ');
    SQL.Add('   		WHEN dbo.PRODUTO_FOR.PRO_FORCODIGO IS NULL THEN ''0''   ');
    SQL.Add('   		WHEN dbo.PRODUTO_FOR.PRO_FORCODIGO = '''' THEN ''0''   ');
    SQL.Add('   		ELSE dbo.PRODUTO_FOR.PRO_FORCODIGO   ');
    SQL.Add('   	END AS DES_REFERENCIA,   ');
    SQL.Add('   	dbo.FORNECEDOR.FOR_CGC AS NUM_CGC,   ');
    SQL.Add('   	NULL AS COD_DIVISAO,   ');
    SQL.Add('   	CASE       ');
    SQL.Add('   			WHEN PRO_UNIDADE = ''UNID'' THEN ''UN''   ');
    SQL.Add('   			WHEN PRO_UNIDADE = ''UND'' THEN ''UN''   ');
    SQL.Add('   			ELSE PRO_UNIDADE      ');
    SQL.Add('   	END AS DES_UNIDADE_COMPRA,   ');
    SQL.Add('   	dbo.PRODUTO_FOR.PRO_FOR_EMBALAGEM AS QTD_EMBALAGEM_COMPRA,   ');
    SQL.Add('   	1 AS QTD_TROCA,   ');
    SQL.Add('   	''N'' AS FLG_PREFERENCIAL   ');
    SQL.Add('      ');
    SQL.Add('   FROM dbo.PRODUTO_FOR   ');
    SQL.Add('   INNER JOIN dbo.FORNECEDOR ON   ');
    SQL.Add('   	(dbo.PRODUTO_FOR.FOR_CODIGO = dbo.FORNECEDOR.FOR_CODIGO)   ');
    SQL.Add('   INNER JOIN dbo.PRODUTOS ON   ');
    SQL.Add('   	(dbo.PRODUTO_FOR.PRO_CODIGO = dbo.PRODUTOS.PRO_CODIGO)   ');
    SQL.Add('   WHERE PRO_DESCRICAO IS NOT NULL   ');

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

        Layout.FieldByName('COD_PRODUTO').AsString :=
          GerarPLU(QryPrincipal2.FieldByName('COD_PRODUTO').AsString);
        Layout.FieldByName('NUM_CGC').AsString :=
          StrRetNums(Layout.FieldByName('NUM_CGC').AsString);

        Layout.WriteLine;
      except
        On E: Exception do
          FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
      end;
      Next;
    end;
  end;
end;



procedure TFrmSmBomPrecoOurinhosGestora.GerarProdSimilar;
var
  NumLinha: Integer;
begin
  inherited;
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

     SQL.Add('   SELECT DISTINCT    ');
     SQL.Add('   	PF.ID_PRODUTO_FAMILIA AS COD_PRODUTO_SIMILAR,   ');
     SQL.Add('   	''ADEFNIR'' AS DES_PRODUTO_SIMILAR   ');
     SQL.Add('   FROM PRODUTO_FAMILIA_ITENS PF   ');
     SQL.Add('   ORDER BY COD_PRODUTO_SIMILAR   ');

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





procedure TFrmSmBomPrecoOurinhosGestora.GerarVenda;
begin
  inherited;
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT');
    SQL.Add(' CUPOMPRODUTOS.PRO_CODIGO AS COD_PRODUTO,');
    SQL.Add(' ' + CbxLoja.Text + ' AS COD_LOJA,');
    SQL.Add(' 0 AS IND_TIPO,');
    SQL.Add(' 1 AS NUM_PDV,');
    SQL.Add(' CUPOMPRODUTOS.SAI_QTDE AS QTD_TOTAL_PRODUTO,');
    SQL.Add(' CUPOMPRODUTOS.SAI_TOTAL - CUPOMPRODUTOS.PRO_DESCONTO AS VAL_TOTAL_PRODUTO,');
    SQL.Add(' CUPOMPRODUTOS.PRO_VENDA AS VAL_PRECO_VENDA,');
    SQL.Add(' CUPOMPRODUTOS.PRO_CUSTO AS VAL_CUSTO_REP,');
    SQL.Add(' CUPOMFISCAL.COM_DATA AS DTA_SAIDA,');
    SQL.Add('');
    SQL.Add(' CASE ');
    SQL.Add('  WHEN DATEPART(MONTH,CUPOMFISCAL.COM_DATA)<10 THEN ''0''+CAST(DATEPART(MONTH,CUPOMFISCAL.COM_DATA) AS VARCHAR)+CAST(DATEPART(YEAR,CUPOMFISCAL.COM_DATA) AS VARCHAR)');
    SQL.Add('  ELSE CAST(DATEPART(MONTH,CUPOMFISCAL.COM_DATA) AS VARCHAR)+CAST(DATEPART(YEAR,CUPOMFISCAL.COM_DATA) AS VARCHAR)      ');
    SQL.Add(' END AS DTA_MENSAL,');
    SQL.Add('');
    SQL.Add(' CUPOMPRODUTOS.SAI_REGISTRO AS NUM_IDENT,');
    SQL.Add(' '''' AS COD_EAN,');
    SQL.Add('');
    SQL.Add(' CASE ');
    SQL.Add('  WHEN DATEPART(HH,CUPOMFISCAL.COM_HORA)<10 AND DATEPART(MI,CUPOMFISCAL.COM_HORA)<10 THEN ''0''+CAST(DATEPART(HH,CUPOMFISCAL.COM_HORA) AS VARCHAR)+''0''+CAST(DATEPART(MI,CUPOMFISCAL.COM_HORA) AS VARCHAR)  ');
    SQL.Add('  WHEN DATEPART(HH,CUPOMFISCAL.COM_HORA)>10 AND DATEPART(MI,CUPOMFISCAL.COM_HORA)>10 THEN CAST(DATEPART(HH,CUPOMFISCAL.COM_HORA) AS VARCHAR)+CAST(DATEPART(MI,CUPOMFISCAL.COM_HORA) AS VARCHAR)');
    SQL.Add('  WHEN DATEPART(HH,CUPOMFISCAL.COM_HORA)<10 AND DATEPART(MI,CUPOMFISCAL.COM_HORA)>10 THEN ''0''+CAST(DATEPART(HH,CUPOMFISCAL.COM_HORA) AS VARCHAR)+CAST(DATEPART(MI,CUPOMFISCAL.COM_HORA) AS VARCHAR)');
    SQL.Add('  WHEN DATEPART(HH,CUPOMFISCAL.COM_HORA)>10 AND DATEPART(MI,CUPOMFISCAL.COM_HORA)<10 THEN CAST(DATEPART(HH,CUPOMFISCAL.COM_HORA) AS VARCHAR)+''0''+CAST(DATEPART(MI,CUPOMFISCAL.COM_HORA) AS VARCHAR)  ');
    SQL.Add(' END AS DES_HORA,');
    SQL.Add('');
    SQL.Add(' CUPOMFISCAL.CLI_CODIGO AS COD_CLIENTE,');
    SQL.Add(' 1 AS COD_ENTIDADE,');
    SQL.Add(' 0 AS VAL_BASE_ICMS,');
    SQL.Add(' CUPOMPRODUTOS.PRO_SIT_TRIBUTARIA AS DES_SITUACAO_TRIB,');
    SQL.Add(' 0 AS VAL_ICMS,');
    SQL.Add(' CUPOMFISCAL.COM_NCUPOM AS NUM_CUPOM_FISCAL,');
    SQL.Add(' CUPOMPRODUTOS.SAI_TOTAL - CUPOMPRODUTOS.PRO_DESCONTO AS VAL_VENDA_PDV,');
    SQL.Add(' CUPOMPRODUTOS.TRI_CODIGO AS COD_TRIBUTACAO,');
    SQL.Add('');
    SQL.Add(' CASE ');
    SQL.Add('  WHEN CUPOMPRODUTOS.SAI_STATUS <> ''A''  THEN ''S'' ');
    SQL.Add('  ELSE ''N'' ');
    SQL.Add(' END AS FLG_CUPOM_CANCELADO,');
    SQL.Add('');
    SQL.Add(' CUPOMPRODUTOS.PRO_CLASFISCAL AS NUM_NCM,');
    SQL.Add(' 0 AS COD_TAB_SPED,');
    SQL.Add('');
    SQL.Add(' CASE ');
    SQL.Add('  WHEN COALESCE(PRODUTOS.PRO_CST_COFINS_ENTRADA,50) = 50 THEN ''S'' ');
    SQL.Add('  ELSE ''N'' ');
    SQL.Add(' END AS FLG_NAO_PIS_COFINS,');
    SQL.Add('');
    SQL.Add(' CASE ');
    SQL.Add('  WHEN PRODUTOS.PRO_BALANCA = ''S'' THEN ''S'' ');
    SQL.Add('  ELSE ''N'' ');
    SQL.Add(' END FLG_ENVIA_BALANCA,');
    SQL.Add('');
    SQL.Add(' CASE ');
    SQL.Add('  WHEN PRODUTOS.PRO_CST_COFINS_ENTRADA = 73 AND PRODUTOS.PRO_CST_COFINS IN (01,04,05,06) THEN 0 ');
    SQL.Add('  WHEN PRODUTOS.PRO_CST_COFINS_ENTRADA = 70 AND PRODUTOS.PRO_CST_COFINS IN (01,04,06) THEN 1');
    SQL.Add('  WHEN PRODUTOS.PRO_CST_COFINS_ENTRADA = 75 AND PRODUTOS.PRO_CST_COFINS = 05 THEN 2');
    SQL.Add('  WHEN PRODUTOS.PRO_CST_COFINS_ENTRADA = 74 AND PRODUTOS.PRO_CST_COFINS IN (01,06) THEN 3');
    SQL.Add('  ELSE 4');
    SQL.Add(' END AS TIPO_NAO_PIS_COFINS,');
    SQL.Add('');
    SQL.Add(' ''S'' AS FLG_ONLINE,');
    SQL.Add(' ''N'' AS FLG_OFERTA,');
    SQL.Add(' 0 AS COD_ASSOCIADO');
    SQL.Add('FROM    ');
    SQL.Add(' DBO.SP_' + FORMATDATETIME('MM_YYYY', DTPINICIAL.DATE) +
      ' AS CUPOMPRODUTOS');
    SQL.Add('LEFT JOIN ');
    SQL.Add(' DBO.CP_' + FORMATDATETIME('MM_YYYY', DTPINICIAL.DATE) +
      ' AS CUPOMFISCAL ');
    SQL.Add('ON');
    SQL.Add(' CUPOMPRODUTOS.COM_REGISTRO = CUPOMFISCAL.COM_REGISTRO');
    SQL.Add('LEFT JOIN ');
    SQL.Add(' DBO.PRODUTOS AS PRODUTOS ');
    SQL.Add('ON');
    SQL.Add(' CUPOMPRODUTOS.PRO_CODIGO = PRODUTOS.PRO_CODIGO');

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

        Layout.FieldByName('NUM_NCM').AsString :=
          StrRetNums(Layout.FieldByName('NUM_NCM').AsString);

        Layout.FieldByName('COD_PRODUTO').AsString :=
          GerarPLU(QryPrincipal2.FieldByName('COD_PRODUTO').AsString);
        Layout.FieldByName('COD_ASSOCIADO').AsString :=
          GerarPLU(Layout.FieldByName('COD_ASSOCIADO').AsString);
        if QryPrincipal2.FieldByName('DTA_SAIDA').AsString <> '' then
          Layout.FieldByName('DTA_SAIDA').AsString :=
            FORMATDATETIME('dd/mm/yyyy', QryPrincipal2.FieldByName('DTA_SAIDA')
            .AsDateTime);

        Layout.WriteLine;
      except
        On E: Exception do
          FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
      end;
      Next;
    end;
  end;
end;

procedure TFrmSmBomPrecoOurinhosGestora.QryPrincipal2AfterOpen(DataSet: TDataSet);
begin
  inherited;
  Memo1.Lines.Add(QryPrincipal2.SQL.Text);
end;

(*
  procedure TFrmSmBomPrecoOurinhosGestora.btnGeraAssociadoClick(Sender: TObject);
  begin
  inherited;
  GerarAssociado;
  end;
*)

procedure TFrmSmBomPrecoOurinhosGestora.BtnGerarClick(Sender: TObject);
begin

  ADOSQLServer.Connected := False;
  ADOSQLServer.ConnectionString := 'Provider=MSDASQL.1;Password="' +
    edtSenhaOracle.Text + '";User ID=' + edtInst.Text + ';Data Source=' +
    edtSchema.Text + ';Persist Security Info=False';
  ADOSQLServer.Connected := True;
  Memo1.Lines.Clear;
(*
  ADOSQLServer.Connected := False;
//  ADOSQLServer.ConnectionString := 'Provider=MSDASQL.1;Password="'+edtSenhaOracle.Text+'";ID='+edtInst.Text+';Data Source='+edtSchema.Text+';Persist Security Info=False';
   ADOSQLServer.ConnectionString := 'Provider=MSDASQL.1;Persist Security Info=False;Data Source='+edtSchema.Text+';User ID='+edtInst.Text+';Password='+edtSenhaOracle.Text+'';
  ADOSQLServer.Connected := True; *)



  if FlgAtualizaValVenda then
  begin
    if CbxLoja.Text >= '1' then
    begin
      AssignFile(Arquivo, ExtractFilePath(EdtCamArquivo.Text) +
        '\SCRIPT-FILIAIS\SCRIPT_ATUALIZA_VALOR_VENDA_LJ' + CbxLoja.Text
        + '.TXT');
    end
    else
    begin
      AssignFile(Arquivo, ExtractFilePath(EdtCamArquivo.Text) +
        '\SCRIPT-FILIAIS\SCRIPT_ATUALIZA_VALOR_MARGEM_LJ' + CbxLoja.Text
        + '.TXT');
    end;
    Rewrite(Arquivo);
    geraFiliais;
  end;

  if FlgAtualizaCustoRep then
  begin
    AssignFile(Arquivo, ExtractFilePath(EdtCamArquivo.Text) +
      '\SCRIPT-FILIAIS\SCRIPT_ATUALIZA_CUSTO_REP_LJ' + CbxLoja.Text + '.TXT');
    Rewrite(Arquivo);
    geraFiliais;
  end;

  if FlgAtualizaEstoque then
  begin
    AssignFile(Arquivo, ExtractFilePath(EdtCamArquivo.Text) +
      '\SCRIPT-FILIAIS\SCRIPT_ATUALIZA_QNT_ESTOQUE_LJ' + CbxLoja.Text + '.TXT');
    Rewrite(Arquivo);
    geraFiliais;
  end;

  if FlgAtualizaObsForn then
  begin
    AssignFile(Arquivo, ExtractFilePath(EdtCamArquivo.Text) +
      '\SCRIPT-UPDATE\SCRIPT_INSERE_OBSERV_FORN_LJ' + CbxLoja.Text + '.TXT');
    Rewrite(Arquivo);
    geraFiliais;
  end;

  if FlgAtualizaClassif then
  begin
    AssignFile(Arquivo, ExtractFilePath(EdtCamArquivo.Text) +
      '\SCRIPT-UPDATE\SCRIPT_INSERE_CLASSIF_LJ' + CbxLoja.Text + '.TXT');
    Rewrite(Arquivo);
    geraFiliais;
  end;


  inherited;

  if FlgAtualizaValVenda then
    CloseFile(Arquivo);

  if FlgAtualizaCustoRep then
    CloseFile(Arquivo);

  if FlgAtualizaEstoque then
    CloseFile(Arquivo);

  if FlgAtualizaObsForn then
    CloseFile(Arquivo);

end;




procedure TFrmSmBomPrecoOurinhosGestora.GerarFinanceiro(Tipo, Situacao: Integer);
begin
  inherited;
  if Tipo = 1 then
    GerarFinanceiroPagar(IntToStr(Situacao));

  if Tipo = 2 then
    GerarFinanceiroReceber(IntToStr(Situacao));
end;

procedure TFrmSmBomPrecoOurinhosGestora.GerarFinanceiroPagar(Aberto: String);
var
  num_nf_antigo: string;
  num_parcela, cod_parceiro: Integer;
begin
  inherited;
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;
    if Aberto = '1' then
    begin

      SQL.Add('   SELECT DISTINCT     ');
      SQL.Add('       1 AS TIPO_PARCEIRO,      ');
      SQL.Add('       FIN.FOR_CODIGO AS COD_PARCEIRO,      ');
      SQL.Add('       0 AS TIPO_CONTA,      ');
      SQL.Add('       8 AS COD_ENTIDADE,      ');
      SQL.Add('       FIN.FIN_REGISTRO AS NUM_DOCTO,      ');
      SQL.Add('       999 AS COD_BANCO,      ');
      SQL.Add('       '''' AS DES_BANCO,      ');
      SQL.Add('       CONVERT(CHAR, FIN.FIN_DTEMISSAO, 103) AS DTA_EMISSAO,      ');
      SQL.Add('       CONVERT(CHAR, F_FAT.FINFAT_VENCIMENTO, 103) AS DTA_VENCIMENTO,      ');
      SQL.Add('       F_FAT.FINFAT_VALOR AS VAL_PARCELA,      ');
      SQL.Add('       F_FAT.FINFAT_JURO AS VAL_JUROS,      ');
      SQL.Add('       F_FAT.FINFAT_DESCONTO AS VAL_DESCONTO,      ');
      SQL.Add('       ''N'' AS FLG_QUITADO,      ');
      SQL.Add('       '''' AS DTA_QUITADA,      ');
      SQL.Add('       998 AS COD_CATEGORIA,      ');
      SQL.Add('       998 AS COD_SUBCATEGORIA,      ');
      SQL.Add('       1 AS NUM_PARCELA,      ');
      SQL.Add('       F_FAT.FINFAT_PARCELA AS QTD_PARCELA,      ');
      SQL.Add('       1 AS COD_LOJA,      ');
      SQL.Add('       FORN.FOR_CGC AS NUM_CGC,      ');
      SQL.Add('       0 AS NUM_BORDERO,      ');
      SQL.Add('       FIN.FIN_NUMERONOTA AS NUM_NF,      ');
      SQL.Add('       1 AS NUM_SERIE_NF,      ');
      SQL.Add('       F_FAT.FINFAT_VALORTOTAL AS VAL_TOTAL_NF,      ');
      SQL.Add('       '' '' AS DES_OBSERVACAO,      ');
      SQL.Add('       0 AS NUM_PDV,      ');
      SQL.Add('       0 AS NUM_CUPOM_FISCAL,      ');
      SQL.Add('       0 AS COD_MOTIVO,      ');
      SQL.Add('       0 AS COD_CONVENIO,      ');
      SQL.Add('       0 AS COD_BIN,      ');
      SQL.Add('       '''' AS DES_BANDEIRA,      ');
      SQL.Add('       '''' AS DES_REDE_TEF,      ');
      SQL.Add('       0 AS VAL_RETENCAO,      ');
      SQL.Add('       0 AS COD_CONDICAO,      ');
      SQL.Add('       '''' AS DTA_PAGTO,      ');
      SQL.Add('       CONVERT(CHAR, FIN.FIN_DTEMISSAO, 103) AS DTA_ENTRADA,      ');
      SQL.Add('       '''' AS NUM_NOSSO_NUMERO,      ');
      SQL.Add('       '''' AS COD_BARRA,      ');
      SQL.Add('       ''N'' AS FLG_BOLETO_EMIT,      ');
      SQL.Add('       '''' AS NUM_CGC_CPF_TITULAR,      ');
      SQL.Add('       FORN.FOR_RAZAO AS DES_TITULAR,      ');
      SQL.Add('       30 AS NUM_CONDICAO,      ');
      SQL.Add('       0 AS VAL_CREDITO,      ');
      SQL.Add('       999 AS COD_BANCO_PGTO,      ');
      SQL.Add('       ''PAGTO'' AS DES_CC,      ');
      SQL.Add('       0 AS COD_BANDEIRA,      ');
      SQL.Add('       '''' AS DTA_PRORROGACAO,      ');
      SQL.Add('       1 AS NUM_SEQ_FIN,      ');
      SQL.Add('       0 AS COD_COBRANCA,      ');
      SQL.Add('       '''' AS DTA_COBRANCA,      ');
      SQL.Add('       ''N'' AS FLG_ACEITE,      ');
      SQL.Add('       0 AS TIPO_ACEITE   ');
      SQL.Add('   	   ');
      SQL.Add('   FROM dbo.FIN_FINANCEIRO FIN   ');
      SQL.Add('   LEFT JOIN dbo.FORNECEDOR FORN ON   ');
      SQL.Add('   	(FIN.FOR_CODIGO = FORN.FOR_CODIGO)   ');
      SQL.Add('   LEFT JOIN dbo.FIN_FATURA F_FAT ON    ');
      SQL.Add('   	(FIN.FIN_REGISTRO = F_FAT.FIN_REGISTRO)   ');
      SQL.Add('   WHERE FIN.FIN_STATUS = ''A''   ');
      SQL.Add('   AND F_FAT.FINFAT_VALOR IS NOT NULL   ');
      SQL.Add('   AND F_FAT.FINFAT_VALORPAGO IS NULL   ');
      SQL.Add('	 AND FIN.EMP_CODIGO = ' + CbxLoja.Text + ' ');

    end
    else
    begin
      (*
        SQL.Add('SELECT');
        SQL.Add(' 1 AS TIPO_PARCEIRO,');
        SQL.Add(' CASE ');
        SQL.Add('  WHEN CP.FOR_CODIGO = 0 THEN 99999 ');
        SQL.Add('  ELSE CP.FOR_CODIGO ');
        SQL.Add(' END AS COD_PARCEIRO,');
        SQL.Add(' 0 AS TIPO_CONTA,');
        SQL.Add(' CASE ');
        SQL.Add('  WHEN CP.CON_TIPODOC = ''NF'' THEN 9');
        SQL.Add('  ELSE 8 ');
        SQL.Add(' END AS COD_ENTIDADE,');
        SQL.Add(' CP.CON_NLCTO AS NUM_DOCTO,');
        SQL.Add(' 999 AS COD_BANCO,');
        SQL.Add(' '''' AS DES_BANCO,');
        SQL.Add(' CP.CON_EMISSAO AS DTA_EMISSAO,');
        SQL.Add(' CP.CON_VECTO AS DTA_VENCIMENTO,');
        SQL.Add(' CP.CON_VALOR AS VAL_PARCELA,');
        SQL.Add(' CP.CON_JUROS AS VAL_JUROS,');
        SQL.Add(' 0 AS VAL_DESCONTO,');
        SQL.Add('');
        SQL.Add(' CASE ');
        SQL.Add('  WHEN CP.CON_STATUS = ''X'' THEN ''N''');
        SQL.Add('  ELSE ''S''');
        SQL.Add(' END AS FLG_QUITADO,');
        SQL.Add('');
        SQL.Add(' CASE ');
        SQL.Add('  WHEN CP.CON_STATUS = ''X'' THEN ''''');
        SQL.Add('  ELSE CP.CON_DPAGO');
        SQL.Add(' END AS DTA_QUITADA,');
        SQL.Add('');
        SQL.Add(' 998 AS COD_CATEGORIA,');
        SQL.Add(' 998 AS COD_SUBCATEGORIA,');
        SQL.Add(' 1 AS NUM_PARCELA,');
        SQL.Add(' COALESCE(PARCELAS.QTD_PARCELA, 1) AS QTD_PARCELA,');
        SQL.Add(' CP.EMP_CODIGO AS COD_LOJA,');
        SQL.Add(' FORNECEDOR.FOR_CGC AS NUM_CGC,');
        SQL.Add(' 0 AS NUM_BORDERO,');
        SQL.Add(' CP.ENT_NNOTA AS NUM_NF,');
        SQL.Add(' '''' AS NUM_SERIE_NF,');
        SQL.Add(' PARCELAS.VAL_TOTAL_NF AS VAL_TOTAL_NF,');
        SQL.Add(' '''' AS DES_OBSERVACAO,');
        SQL.Add(' 1 AS NUM_PDV,');
        SQL.Add(' '''' AS NUM_CUPOM_FISCAL,');
        SQL.Add(' 0 AS COD_MOTIVO,');
        SQL.Add(' 0 AS COD_CONVENIO,');
        SQL.Add(' 0 AS COD_BIN,');
        SQL.Add(' '''' AS DES_BANDEIRA,');
        SQL.Add(' '''' AS DES_REDE_TEF,');
        SQL.Add(' 0 AS VAL_RETENCAO,');
        SQL.Add(' 0 AS COD_CONDICAO,');
        SQL.Add(' '''' AS DTA_PAGTO,');
        SQL.Add(' CP.CON_DLCTO AS DTA_ENTRADA,');
        SQL.Add(' '''' AS NUM_NOSSO_NUMERO,');
        SQL.Add(' CP.CON_BARRA AS COD_BARRA,');
        SQL.Add(' ''N'' AS FLG_BOLETO_EMIT,');
        SQL.Add(' '''' AS NUM_CGC_CPF_TITULAR,');
        SQL.Add(' '''' AS DES_TITULAR,');
        SQL.Add(' 0 AS NUM_CONDICAO');
        SQL.Add('FROM');
        SQL.Add(' CONTABIL AS CP');
        SQL.Add('INNER JOIN');
        SQL.Add(' FORNECEDOR');
        SQL.Add('ON');
        SQL.Add(' CP.FOR_CODIGO = FORNECEDOR.FOR_CODIGO');
        SQL.Add('LEFT JOIN');
        SQL.Add(' (');
        SQL.Add('	SELECT');
        SQL.Add(' 	 CONTABIL.ENT_NNOTA,');
        SQL.Add('	 CONTABIL.FOR_CODIGO,');
        SQL.Add('	 COUNT AS QTD_PARCELA,');
        SQL.Add('	 SUM(CON_VALOR_DOC) AS VAL_TOTAL_NF');
        SQL.Add('	FROM');
        SQL.Add('	 CONTABIL');
        SQL.Add('	WHERE');
        SQL.Add('	 COALESCE(CONTABIL.ENT_NNOTA,'''') <> ''''');
        SQL.Add('	AND');
        SQL.Add('	 CON_ACAO IS NULL');
        SQL.Add('	AND');
        SQL.Add('	 CON_DATA_EXCLUSAO IS NULL');
        SQL.Add('	AND');
        SQL.Add('	 CON_CREDITO > CON_DEBITO');
        SQL.Add('	AND');

        //      SQL.Add('	 EMP_CODIGO = 1');
        SQL.Add('	 EMP_CODIGO = '+CbxLoja.Text+' ');

        SQL.Add('	GROUP BY');
        SQL.Add('	 CONTABIL.ENT_NNOTA,');
        SQL.Add('	 CONTABIL.FOR_CODIGO');
        SQL.Add(' ) AS PARCELAS');
        SQL.Add('ON');
        SQL.Add(' CP.FOR_CODIGO = PARCELAS.FOR_CODIGO');
        SQL.Add('AND');
        SQL.Add(' CP.ENT_NNOTA = PARCELAS.ENT_NNOTA');



        SQL.Add('WHERE');
        SQL.Add(' CP.CON_DATA_EXCLUSAO IS NULL');
        SQL.Add('AND');
        SQL.Add(' CP.CON_CREDITO > CP.CON_DEBITO');
        SQL.Add('AND');
        SQL.Add(' CP.CON_ACAO IS NULL');
        SQL.Add('AND');

        //      SQL.Add(' CP.EMP_CODIGO = 1');
        SQL.Add('	 EMP_CODIGO = '+CbxLoja.Text+' ');

        SQL.Add('AND');
        SQL.Add(' CP.CON_EMISSAO >= '''+FormatDateTime('yyyy-mm-dd',DtpInicial.Date)+''' ');
        //      SQL.Add(' CP.CON_EMISSAO >= '''+FormatDateTime('dd-mm-yyyy',DtpInicial.Date)+''' ');
        SQL.Add('AND');
        SQL.Add(' CP.CON_EMISSAO <= '''+FormatDAteTime('yyyy-mm-dd',DtpFinal.Date)+''' ');
        //      SQL.Add(' CP.CON_EMISSAO <= '''+FormatDateTime('dd-mm-yyyy',DtpFinal.Date)+''' ');
        SQL.Add('ORDER BY');
        SQL.Add(' CP.ENT_NNOTA,');
        SQL.Add(' CP.FOR_CODIGO,');
        SQL.Add(' CP.CON_DLCTO');

        //      Parameters.ParamByName('INI').Value := FormatDateTime('yyyy-mm-dd',DtpInicial.Date);
        //      Parameters.ParamByName('FIM').Value := FormatDAteTime('yyyy-mm-dd',DtpFinal.Date);
      *)
    end;

    Open;
    First;
    NumLinha := 0;
    num_nf_antigo := 'inicio';
    cod_parceiro := 0;
    num_parcela := 0;

    while not Eof do
    begin
      try
        if Cancelar then
          Break;
        Inc(NumLinha);
        Layout.SetValues(QryPrincipal2, NumLinha, RecordCount);

        if ((Layout.FieldByName('NUM_NF').AsString <> num_nf_antigo) and
          (Layout.FieldByName('COD_PARCEIRO').AsInteger <> cod_parceiro)) or
          (Layout.FieldByName('QTD_PARCELA').AsInteger = 1) then
        begin
          cod_parceiro := Layout.FieldByName('COD_PARCEIRO').AsInteger;
          num_nf_antigo := Layout.FieldByName('NUM_NF').AsString;
          num_parcela := 0;
        end
        else
        begin
          Inc(num_parcela);
          Layout.FieldByName('NUM_PARCELA').AsInteger := num_parcela;
        end;

        Layout.FieldByName('COD_LOJA').AsInteger := strtoint(CbxLoja.Text);

        Layout.FieldByName('DTA_EMISSAO').AsString :=
          FORMATDATETIME('dd/mm/yyyy', QryPrincipal2.FieldByName('DTA_EMISSAO')
          .AsDateTime);

        Layout.FieldByName('DTA_VENCIMENTO').AsString :=
          FORMATDATETIME('dd/mm/yyyy',
          QryPrincipal2.FieldByName('DTA_VENCIMENTO').AsDateTime);

        if QryPrincipal2.FieldByName('DTA_QUITADA').AsString <> '' then
          Layout.FieldByName('DTA_QUITADA').AsString :=
            FORMATDATETIME('dd/mm/yyyy',
            QryPrincipal2.FieldByName('DTA_QUITADA').AsDateTime);

        if QryPrincipal2.FieldByName('DTA_PAGTO').AsString <> '' then
          Layout.FieldByName('DTA_PAGTO').AsString :=
            FORMATDATETIME('dd/mm/yyyy', QryPrincipal2.FieldByName('DTA_PAGTO')
            .AsDateTime);

        Layout.FieldByName('DTA_ENTRADA').AsString :=
          FORMATDATETIME('dd/mm/yyyy', QryPrincipal2.FieldByName('DTA_ENTRADA')
          .AsDateTime);

        Layout.WriteLine;
      except
        On E: Exception do
          FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
      end;
      Next;
    end;

  end;
end;

procedure TFrmSmBomPrecoOurinhosGestora.GerarFinanceiroReceber(Aberto: String);
begin
  inherited;
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    if Aberto = '1' then
    begin

      SQL.Add('   SELECT      ');
      SQL.Add('                    0 AS TIPO_PARCEIRO,      ');
      SQL.Add('                    VENDAS.CLI_CODIGO AS COD_PARCEIRO,      ');
      SQL.Add('                    1 AS TIPO_CONTA,      ');
      SQL.Add('                    4 AS COD_ENTIDADE,      ');
      SQL.Add('                    COALESCE(VENDAS.COM_REGISTRO,0) AS NUM_DOCTO,      ');
      SQL.Add('                    999 AS COD_BANCO,      ');
      SQL.Add('                    ''RECEBTO'' AS DES_BANCO,      ');
      SQL.Add('                    VENDAS.VEN_DATA AS DTA_EMISSAO,      ');
      SQL.Add('                    VENDAS.VEN_VENCIMENTO AS DTA_VENCIMENTO,      ');
      SQL.Add('                    VENDAS.PRO_VENDA AS VAL_PARCELA,      ');
      SQL.Add('                    VENDAS.VEN_TAXA AS VAL_JUROS,      ');
      SQL.Add('                    0 AS VAL_DESCONTO,      ');
      SQL.Add('                    ''N'' AS FLG_QUITADO,      ');
      SQL.Add('                    '''' AS DTA_QUITADA,      ');
      SQL.Add('                    997 AS COD_CATEGORIA,      ');
      SQL.Add('                    997 AS COD_SUBCATEGORIA,      ');
      SQL.Add('                    1 AS NUM_PARCELA,      ');
      SQL.Add('                    1 AS QTD_PARCELA,      ');
      SQL.Add('                    1 AS COD_LOJA,      ');
      SQL.Add('                    CLIENTES.CLI_CPFCGC AS NUM_CGC,      ');
      SQL.Add('                    0 AS NUM_BORDERO,      ');
      SQL.Add('                    '''' AS NUM_NF,      ');
      SQL.Add('                    0 AS NUM_SERIE_NF,      ');
      SQL.Add('                    VENDAS.PRO_VENDA + VENDAS.VEN_TAXA AS VAL_TOTAL_NF,      ');
      SQL.Add('                    '''' AS DES_OBSERVACAO,      ');
      SQL.Add('                         ');
      SQL.Add('                    CASE UPPER(VENDAS.MAQ_NOME) WHEN ''CAIXA-01'' THEN 1      ');
      SQL.Add('                     WHEN ''CAIXA-02'' THEN 2      ');
      SQL.Add('                     WHEN ''CAIXA-03'' THEN 3       ');
      SQL.Add('                     ELSE 1       ');
      SQL.Add('                    END AS NUM_PDV,      ');
      SQL.Add('                         ');
      SQL.Add('                    COALESCE(VENDAS.COM_NCUPOM, 0) AS NUM_CUPOM_FISCAL,      ');
      SQL.Add('                    0 AS COD_MOTIVO,      ');
      SQL.Add('                    6 AS COD_CONVENIO,      ');
      SQL.Add('                    0 AS COD_BIN,      ');
      SQL.Add('                    '''' AS DES_BANDEIRA,      ');
      SQL.Add('                    '''' AS DES_REDE_TEF,      ');
      SQL.Add('                    0 AS VAL_RETENCAO,      ');
      SQL.Add('                    0 AS COD_CONDICAO,      ');
      SQL.Add('                    '''' AS DTA_PAGTO,      ');
      SQL.Add('                    VENDAS.VEN_DATA AS DTA_ENTRADA,      ');
      SQL.Add('                    VENDAS.VEN_NOSSONUMERO AS NUM_NOSSO_NUMERO,      ');
      SQL.Add('                    VENDAS.VEN_CODIGOBARRA AS COD_BARRA,      ');
      SQL.Add('                    ''N'' AS FLG_BOLETO_EMIT,      ');
      SQL.Add('                    CLIENTES.CLI_CPFCGC AS NUM_CGC_CPF_TITULAR,      ');
      SQL.Add('                    CLIENTES.CLI_NOME AS DES_TITULAR,      ');
      SQL.Add('                    0 AS NUM_CONDICAO      ');
      SQL.Add('                   FROM       ');
      SQL.Add('                    DBO.VENDAS_PRAZO AS VENDAS      ');
      SQL.Add('                   LEFT JOIN       ');
      SQL.Add('                    DBO.CLIENTES AS CLIENTES       ');
      SQL.Add('                   ON      ');
      SQL.Add('                    VENDAS.CLI_CODIGO = CLIENTES.CLI_CODIGO   ');
      SQL.Add('   				WHERE VENDAS.CLI_CODIGO  = 90   ');
      SQL.Add('      ');
      SQL.Add('         UNION ALL   ');
      SQL.Add('      ');
      SQL.Add('         SELECT   ');
      SQL.Add('          0 AS TIPO_PARCEIRO,   ');
      SQL.Add('          CHEQUES.CLI_CODIGO AS COD_PARCEIRO,   ');
      SQL.Add('          1 AS TIPO_CONTA,   ');
      SQL.Add('            ');
      SQL.Add('          CASE    ');
      SQL.Add('           WHEN CHEQUES.CHE_DATA = CHEQUES.CHE_VECTO THEN 2    ');
      SQL.Add('           ELSE 3    ');
      SQL.Add('          END AS COD_ENTIDADE,   ');
      SQL.Add('            ');
      SQL.Add('          COALESCE(CHEQUES.CHE_NUMERO, 0) AS NUM_DOCTO,   ');
      SQL.Add('          999 AS COD_BANCO,   ');
      SQL.Add('          ''RECEBTO'' AS DES_BANCO,   ');
      SQL.Add('          CHEQUES.CHE_DATA AS DTA_EMISSAO,   ');
      SQL.Add('          CHEQUES.CHE_VECTO AS DTA_VENCIMENTO,   ');
      SQL.Add('          CHEQUES.CHE_VALOR AS VAL_PARCELA,   ');
      SQL.Add('          CHEQUES.CHE_JUROS AS VAL_JUROS,   ');
      SQL.Add('          0 AS VAL_DESCONTO,   ');
      SQL.Add('          ''N'' AS FLG_QUITADO,   ');
      SQL.Add('          '''' AS DTA_QUITADA,   ');
      SQL.Add('          997 AS COD_CATEGORIA,   ');
      SQL.Add('          997 AS COD_SUBCATEGORIA,   ');
      SQL.Add('          1 AS NUM_PARCELA,   ');
      SQL.Add('          1 AS QTD_PARCELA,   ');
      SQL.Add('          1 AS COD_LOJA,   ');
      SQL.Add('          CLIENTES.CLI_CPFCGC AS NUM_CGC,   ');
      SQL.Add('          0 AS NUM_BORDERO,   ');
      SQL.Add('          '''' AS NUM_NF,   ');
      SQL.Add('          0 AS NUM_SERIE_NF,   ');
      SQL.Add('          CHEQUES.CHE_JUROS AS VAL_TOTAL_NF,   ');
      SQL.Add('          '''' AS DES_OBSERVACAO,   ');
      SQL.Add('          1 AS NUM_PDV,   ');
      SQL.Add('          0 AS NUM_CUPOM_FISCAL,   ');
      SQL.Add('          0 AS COD_MOTIVO,   ');
      SQL.Add('          6 AS COD_CONVENIO,   ');
      SQL.Add('          0 AS COD_BIN,   ');
      SQL.Add('          '''' AS DES_BANDEIRA,   ');
      SQL.Add('          '''' AS DES_REDE_TEF,   ');
      SQL.Add('          0 AS VAL_RETENCAO,   ');
      SQL.Add('          0 AS COD_CONDICAO,   ');
      SQL.Add('          '''' AS DTA_PAGTO,   ');
      SQL.Add('          CHEQUES.CHE_DATA AS DTA_ENTRADA,   ');
      SQL.Add('          0 AS NUM_NOSSO_NUMERO,   ');
      SQL.Add('          CHEQUES.CHE_CMC7 AS COD_BARRA,   ');
      SQL.Add('          ''N'' AS FLG_BOLETO_EMIT,   ');
      SQL.Add('          CLIENTES.CLI_CPFCGC AS NUM_CGC_CPF_TITULAR,   ');
      SQL.Add('          CLIENTES.CLI_NOME AS DES_TITULAR,   ');
      SQL.Add('          0 AS NUM_CONDICAO   ');
      SQL.Add('         FROM   ');
      SQL.Add('          CHEQUE_REC AS CHEQUES   ');
      SQL.Add('         LEFT JOIN   ');
      SQL.Add('          CLIENTES   ');
      SQL.Add('         ON   ');
      SQL.Add('          CHEQUES.CLI_CODIGO = CLIENTES.CLI_CODIGO   ');
      SQL.Add('         WHERE   ');
      SQL.Add('          CHEQUES.CHE_STATUS IN (''A'', ''D'')   ');

    end
    else
    begin

      SQL.Add('SELECT ');
      SQL.Add(' 0 AS TIPO_PARCEIRO,');
      SQL.Add(' BAIXA.CLI_CODIGO AS COD_PARCEIRO,');
      SQL.Add(' 1 AS TIPO_CONTA,');
      SQL.Add(' 4 AS COD_ENTIDADE,');
      SQL.Add(' COALESCE(BAIXA.COM_NCUPOM, 0) AS  NUM_DOCTO,');
      SQL.Add(' 0 AS COD_BANCO,');
      SQL.Add(' '''' AS DES_BANCO,');
      SQL.Add(' BAIXA.VEN_DATA AS DTA_EMISSAO,');
      SQL.Add(' BAIXA.VEN_VENCIMENTO AS DTA_VENCIMENTO,');
      SQL.Add(' BAIXA.PRO_VENDA + BAIXA.VEN_TAXA AS VAL_PARCELA,');
      SQL.Add(' BAIXA.VEN_TAXA AS VAL_JUROS,');
      SQL.Add(' 0 AS VAL_DESCONTO,');
      SQL.Add(' ''S'' AS FLG_QUITADO,');
      SQL.Add(' BAIXA.DATA_PROCESSO AS DTA_QUITADA,');
      SQL.Add(' 997 AS COD_CATEGORIA,');
      SQL.Add(' 997 AS COD_SUBCATEGORIA,');
      SQL.Add(' 1 AS NUM_PARCELA,');
      SQL.Add(' 1 AS QTD_PARCELA,');
      SQL.Add(' 1 AS COD_LOJA,');
      SQL.Add(' CLIENTES.CLI_CPFCGC AS NUM_CGC,');
      SQL.Add(' 0 AS NUM_BORDERO,');
      SQL.Add(' '''' AS NUM_NF,');
      SQL.Add(' 0 AS NUM_SERIE_NF,');
      SQL.Add(' BAIXA.PRO_VENDA + BAIXA.VEN_TAXA AS VAL_TOTAL_NF,');
      SQL.Add(' '''' AS DES_OBSERVACAO,');
      SQL.Add('');
      SQL.Add(' CASE UPPER(BAIXA.MAQ_NOME) WHEN ''CAIXA-01'' THEN 1');
      SQL.Add('  WHEN ''CAIXA-02'' THEN 2');
      SQL.Add('  WHEN ''CAIXA-03'' THEN 3 ');
      SQL.Add('  ELSE 1 ');
      SQL.Add(' END AS NUM_PDV,');
      SQL.Add('');
      SQL.Add(' COALESCE(BAIXA.COM_NCUPOM, 0) AS NUM_CUPOM_FISCAL,');
      SQL.Add(' 0 AS COD_MOTIVO,');
      SQL.Add('');
      SQL.Add(' CASE COALESCE(CLIENTES.CON_CODIGO,0)');
      SQL.Add('  WHEN 0 THEN 999999');
      SQL.Add('  ELSE ''99999'' + CAST(CLIENTES.CON_CODIGO AS VARCHAR(1))');
      SQL.Add(' END AS COD_CONVENIO,');
      SQL.Add('');
      SQL.Add(' 0 AS COD_BIN,');
      SQL.Add(' '''' AS DES_BANDEIRA,');
      SQL.Add(' '''' AS DES_REDE_TEF,');
      SQL.Add(' 0 AS VAL_RETENCAO,');
      SQL.Add(' 0 AS COD_CONDICAO,');
      SQL.Add(' BAIXA.DATA_PROCESSO AS DTA_PAGTO,');
      SQL.Add(' BAIXA.VEN_DATA AS DTA_ENTRADA,');
      SQL.Add(' 0 AS NUM_NOSSO_NUMERO,');
      SQL.Add(' '''' AS COD_BARRA,');
      SQL.Add(' ''N'' AS FLG_BOLETO_EMIT,');
      SQL.Add(' CLIENTES.CLI_CPFCGC AS NUM_CGC_CPF_TITULAR,');
      SQL.Add(' CLIENTES.CLI_NOME AS DES_TITULAR,');
      SQL.Add(' 0 AS NUM_CONDICAO,');
      SQL.Add(' 1 AS NUM_SEQ_FIN');
      SQL.Add('FROM ');
      SQL.Add(' DBO.BAIXAS_PRAZO AS BAIXA ');
      SQL.Add('LEFT JOIN ');
      SQL.Add(' DBO.CLIENTES AS CLIENTES ');
      SQL.Add('ON');
      SQL.Add(' BAIXA.CLI_CODIGO = CLIENTES.CLI_CODIGO    ');
      SQL.Add('WHERE');
      SQL.Add(' CAST(BAIXA.VEN_DATA AS DATE) BETWEEN :INI AND :FIM ');

      Parameters.ParamByName('INI').Value := FORMATDATETIME('yyyy-mm-dd',
        DTPINICIAL.DATE);
      Parameters.ParamByName('FIM').Value := FORMATDATETIME('yyyy-mm-dd',
        DtpFinal.DATE);
    end;

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

        Layout.FieldByName('DES_OBSERVACAO').AsString :=
          StrLBReplace(QryPrincipal2.FieldByName('DES_OBSERVACAO').AsString);
        Layout.FieldByName('NUM_NF').AsString :=
          StrRetNums(QryPrincipal2.FieldByName('NUM_NF').AsString);
        Layout.FieldByName('NUM_CUPOM_FISCAL').AsString :=
          StrRetNums(QryPrincipal2.FieldByName('NUM_CUPOM_FISCAL').AsString);
        Layout.FieldByName('NUM_CGC').AsString :=
          StrRetNums(QryPrincipal2.FieldByName('NUM_CGC').AsString);

        if QryPrincipal2.FieldByName('DTA_QUITADA').AsString <> '' then
          Layout.FieldByName('DTA_QUITADA').AsString :=
            FORMATDATETIME('dd/mm/yyyy',
            QryPrincipal2.FieldByName('DTA_QUITADA').AsDateTime);

        if QryPrincipal2.FieldByName('DTA_PAGTO').AsString <> '' then
          Layout.FieldByName('DTA_PAGTO').AsString :=
            FORMATDATETIME('dd/mm/yyyy', QryPrincipal2.FieldByName('DTA_PAGTO')
            .AsDateTime);

        if QryPrincipal2.FieldByName('DTA_EMISSAO').AsString <> '' then
          Layout.FieldByName('DTA_EMISSAO').AsString :=
            FORMATDATETIME('dd/mm/yyyy',
            QryPrincipal2.FieldByName('DTA_EMISSAO').AsDateTime);

        if QryPrincipal2.FieldByName('DTA_VENCIMENTO').AsString <> '' then
          Layout.FieldByName('DTA_VENCIMENTO').AsString :=
            FORMATDATETIME('dd/mm/yyyy',
            QryPrincipal2.FieldByName('DTA_VENCIMENTO').AsDateTime);

        if QryPrincipal2.FieldByName('DTA_ENTRADA').AsString <> '' then
          Layout.FieldByName('DTA_ENTRADA').AsString :=
            FORMATDATETIME('dd/mm/yyyy',
            QryPrincipal2.FieldByName('DTA_ENTRADA').AsDateTime);

        Layout.WriteLine;
      except
        On E: Exception do
          FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
      end;
      Next;
    end;
  end;
end;



procedure TFrmSmBomPrecoOurinhosGestora.GerarNFFornec;
begin
  inherited;
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('SELECT ');
    SQL.Add(' FORNECEDOR.FOR_CODIGO AS COD_FORNECEDOR,');
    SQL.Add(' RTRIM(NOTAS.ENT_NNOTA) AS NUM_NF_FORN,');
    SQL.Add(' CASE WHEN COALESCE(NOTAS.ENT_SERIE,'''') = '''' THEN ''0'' ELSE NOTAS.ENT_SERIE END AS NUM_SERIE_NF,');
    SQL.Add(' NOTAS.ENT_SUBSERIE AS NUM_SUBSERIE_NF,');
    SQL.Add(' 5102 AS CFOP,');
    SQL.Add(' 0 AS TIPO_NF,');
    SQL.Add(' NOTAS.TP_NOTA AS DES_ESPECIE,');
    SQL.Add(' NOTAS.ENT_VALOR AS VAL_TOTAL_NF,');
    SQL.Add(' NOTAS.ENT_DATA_EMISSAO AS DTA_EMISSAO,');
    SQL.Add(' NOTAS.ENT_DATA_ENTRADA AS DTA_ENTRADA,');
    SQL.Add(' 0 AS VAL_TOTAL_IPI,');
    SQL.Add(' NOTAS.ENT_VALOR AS VAL_VENDA_VAREJO,');
    SQL.Add(' NOTAS.ENT_FRETE VAL_FRETE,');
    SQL.Add(' NOTAS.ENT_ACRESCIMO AS VAL_ACRESCIMO,');
    SQL.Add(' NOTAS.ENT_DESCONTO AS VAL_DESCONTO,');
    SQL.Add(' FORNECEDOR.FOR_CGC AS NUM_CGC,');
    SQL.Add(' NOTAS.ENT_BASECALCULO AS VAL_TOTAL_BC,');
    SQL.Add(' NOTAS.ENT_ICMS AS VAL_TOTAL_ICMS,');
    SQL.Add(' NOTAS.ENT_REFSUB AS VAL_BC_SUBST,');
    SQL.Add(' NOTAS.ENT_SUB AS VAL_ICMS_SUBST,');
    SQL.Add(' 0 AS VAL_FUNRURAL,');
    SQL.Add(' 1 AS COD_PERFIL,');
    SQL.Add(' 0 AS  VAL_DESP_ACESS,');
    SQL.Add(' ''N'' AS FLG_CANCELADO,');
    SQL.Add(' '''' AS DES_OBSERVACAO,');
    SQL.Add(' NOTAS.ENT_CHAVE_NFE AS NUM_CHAVE_ACESSO');
    SQL.Add('FROM');
    SQL.Add(' DBO.ENTRADA_ESTQ AS NOTAS');
    SQL.Add('LEFT JOIN ');
    SQL.Add(' DBO.FORNECEDOR AS FORNECEDOR ');
    SQL.Add('ON');
    SQL.Add(' NOTAS.FOR_CODIGO = FORNECEDOR.FOR_CODIGO ');
    SQL.Add('WHERE');
    SQL.Add(' CAST(NOTAS.ENT_DATA_EMISSAO AS DATE)  BETWEEN :INI AND :FIM');

    Parameters.ParamByName('INI').Value := FORMATDATETIME('yyyy-mm-dd',
      DTPINICIAL.DATE);
    Parameters.ParamByName('FIM').Value := FORMATDATETIME('yyyy-mm-dd',
      DtpFinal.DATE);

    (*
      SQL.Add('WHERE ');
      //SQL.Add('  CAST(ENTRADANF.DTEMISSAO AS DATE) >= '''+FormatDateTime('yyyy-mm-dd',DtpInicial.Date)+''' ');
      //SQL.Add('  AND  ');
      //SQL.Add('  CAST(ENTRADANF.DTEMISSAO AS DATE) <= '''+FormatDateTime('yyyy-mm-dd',DtpFinal.Date)+''' ');
    *)

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
          Layout.FieldByName('DTA_EMISSAO').AsString :=
            FORMATDATETIME('dd/mm/yyyy',
            QryPrincipal2.FieldByName('DTA_EMISSAO').AsDateTime);

        if QryPrincipal2.FieldByName('DTA_ENTRADA').AsString <> '' then
          Layout.FieldByName('DTA_ENTRADA').AsString :=
            FORMATDATETIME('dd/mm/yyyy',
            QryPrincipal2.FieldByName('DTA_ENTRADA').AsDateTime);

        Layout.FieldByName('DES_OBSERVACAO').AsString :=
          StrLBReplace(FieldByName('DES_OBSERVACAO').AsString);

        Layout.WriteLine;
      except
        On E: Exception do
          FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
      end;
      Next;
    end;
  end;
end;

procedure TFrmSmBomPrecoOurinhosGestora.GerarNFitensFornec;
var
  NumLinha, TotalReg, NumItem: Integer;
  nota, serie, fornecedor, CodNf: string;
  Count, RecordCount: Integer;
begin
  inherited;

  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('   SELECT   ');
    SQL.Add('        CAPA.FOR_CODIGO AS COD_FORNECEDOR,   ');
    SQL.Add('        RTRIM(CAPA.ENT_NNOTA) AS NUM_NF_FORN,   ');
    SQL.Add('        CASE WHEN COALESCE(CAPA.ENT_SERIE,'''') = '''' THEN ''0'' ELSE CAPA.ENT_SERIE END AS NUM_SERIE_NF,   ');
    SQL.Add('        ITENS.PRO_CODIGO AS COD_PRODUTO,   ');
    SQL.Add('        CASE          ');
    SQL.Add('           WHEN (TRI_ALIQUOTA = 0.00) AND (TRI_REDUCAO = 0.00) AND (CST = 060) THEN ''25''      ');
    SQL.Add('           WHEN (TRI_ALIQUOTA = 0.00) AND (TRI_REDUCAO = 0.00) AND (CST = 040) THEN ''1''      ');
    SQL.Add('           WHEN (TRI_ALIQUOTA = 0.00) AND (TRI_REDUCAO = 0.00) AND (CST = 041) THEN ''1''      ');
    SQL.Add('           WHEN (TRI_ALIQUOTA = 7.00) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''2''      ');
    SQL.Add('           WHEN (TRI_ALIQUOTA = 12.00) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''3''      ');
    SQL.Add('           WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''4''      ');
    SQL.Add('           WHEN (TRI_ALIQUOTA = 25.00) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''5''      ');
    SQL.Add('           WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 61.11) AND (CST = 020) THEN ''8''      ');
    SQL.Add('           WHEN (TRI_ALIQUOTA = 12.00) AND (TRI_REDUCAO = 41.67) AND (CST = 020) THEN ''6''      ');
    SQL.Add('           WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 33.33) AND (CST = 020) THEN ''7''      ');
    SQL.Add('           WHEN (TRI_ALIQUOTA = 1.25) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''43''      ');
    SQL.Add('           WHEN (TRI_ALIQUOTA = 2.56) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''44''      ');
    SQL.Add('           WHEN (TRI_ALIQUOTA = 4.00) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''27''      ');
    SQL.Add('           WHEN (TRI_ALIQUOTA = 2.82) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''45''      ');
    SQL.Add('           WHEN (TRI_ALIQUOTA = 2.33) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''46''      ');
    SQL.Add('           WHEN (TRI_ALIQUOTA = 2.84) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''47''      ');
    SQL.Add('           WHEN (TRI_ALIQUOTA = 1.86) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''48''      ');
    SQL.Add('           WHEN (TRI_ALIQUOTA = 2.87) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''49''      ');
    SQL.Add('           WHEN (TRI_ALIQUOTA = 3.10) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''50''      ');
    SQL.Add('           WHEN (TRI_ALIQUOTA = 2.58) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''51''      ');
    SQL.Add('           WHEN (TRI_ALIQUOTA = 3.41) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''52''      ');
    SQL.Add('           WHEN (TRI_ALIQUOTA = 3.45) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''53''      ');
    SQL.Add('           WHEN (TRI_ALIQUOTA = 11.00) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''35''      ');
    SQL.Add('           WHEN (TRI_ALIQUOTA = 12.00) AND (TRI_REDUCAO = 8.33) AND (CST = 020) THEN ''34''      ');
    SQL.Add('           WHEN (TRI_ALIQUOTA = 4.50) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''38''      ');
    SQL.Add('           WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 51.11) AND (CST = 020) THEN ''54''      ');
    SQL.Add('           WHEN (TRI_ALIQUOTA = 2.60) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''55''      ');
    SQL.Add('           WHEN (TRI_ALIQUOTA = 1.89) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''56''      ');
    SQL.Add('           WHEN (TRI_ALIQUOTA = 4.79) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''57''      ');
    SQL.Add('           WHEN (TRI_ALIQUOTA = 0.00) AND (TRI_REDUCAO = 0.00) AND (CST = 090) THEN ''22''      ');
    SQL.Add('           WHEN (TRI_ALIQUOTA = 4.70) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''59''      ');
    SQL.Add('           WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 73.89) AND (CST = 020) THEN ''60''      ');
    SQL.Add('           WHEN (TRI_ALIQUOTA = 12.00 ) AND (TRI_REDUCAO = 60.84) AND (CST = 020) THEN ''61''      ');
    SQL.Add('           WHEN (TRI_ALIQUOTA = 9.40) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''62''      ');
    SQL.Add('           WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 47.78) AND (CST = 020) THEN ''63''      ');
    SQL.Add('           WHEN (TRI_ALIQUOTA = 12.00) AND (TRI_REDUCAO = 21.67) AND (CST = 020) THEN ''64''      ');
    SQL.Add('           WHEN (TRI_ALIQUOTA = 13.30) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''65''      ');
    SQL.Add('           WHEN (TRI_ALIQUOTA = 13.30) AND (TRI_REDUCAO = 47.36) AND (CST = 020) THEN ''66''      ');
    SQL.Add('           WHEN (TRI_ALIQUOTA = 4.14) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''67''      ');
    SQL.Add('           WHEN (TRI_ALIQUOTA = 13.30) AND (TRI_REDUCAO = 47.37) AND (CST = 020) THEN ''68''      ');
    SQL.Add('           WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 77.00) AND (CST = 020) THEN ''69''      ');
    SQL.Add('           WHEN (TRI_ALIQUOTA = 12.00) AND (TRI_REDUCAO = 65.50) AND (CST = 020) THEN ''70''      ');
    SQL.Add('           WHEN (TRI_ALIQUOTA = 9.50) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''71''      ');
    SQL.Add('           WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 47.23) AND (CST = 020) THEN ''72''      ');
    SQL.Add('           WHEN (TRI_ALIQUOTA = 12.00) AND (TRI_REDUCAO = 20.84) AND (CST = 020) THEN ''73''      ');
    SQL.Add('           WHEN (TRI_ALIQUOTA = 9.79) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''74''      ');
    SQL.Add('           WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 45.62) AND (CST = 020) THEN ''75''      ');
    SQL.Add('           WHEN (TRI_ALIQUOTA = 12.00) AND (TRI_REDUCAO = 18.42) AND (CST = 020) THEN ''76''      ');
    SQL.Add('           WHEN (TRI_ALIQUOTA = 11.20) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''77''      ');
    SQL.Add('           WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 37.78) AND (CST = 020) THEN ''78''      ');
    SQL.Add('           WHEN (TRI_ALIQUOTA = 12.00) AND (TRI_REDUCAO = 6.67) AND (CST = 020) THEN ''79''      ');
    SQL.Add('           WHEN (TRI_ALIQUOTA = 18.00) AND (TRI_REDUCAO = 26.11) AND (CST = 020) THEN ''80''      ');
    SQL.Add('           WHEN (TRI_ALIQUOTA = 5.50) AND (TRI_REDUCAO = 0.00) AND (CST = 000) THEN ''81''      ');
    SQL.Add('           WHEN (TRI_ALIQUOTA = 13.30) AND (TRI_REDUCAO = 9.77) AND (CST = 020) THEN ''82''      ');
    SQL.Add('           WHEN (TRI_ALIQUOTA = 25.00) AND (TRI_REDUCAO = 52.00) AND (CST = 020) THEN ''9''      ');
    SQL.Add('        END AS COD_TRIBUTACAO,   ');
    SQL.Add('          ');
    SQL.Add('        ITENS.ENT_EMBALAGEM AS QTD_EMBALAGEM,   ');
    SQL.Add('        COALESCE(ITENS.ENT_QTDE_VOL, 1) AS QTD_ENTRADA,   ');
    SQL.Add('          ');
    SQL.Add('        CASE    ');
    SQL.Add('          WHEN PRODUTOS.PRO_UNIDADE IN ( ''UNID '', '' '', ''. '', ''7896 '', ''7897 '', ''TP '', '', '', ''UND '', ''KL '', ''POTE '', ''LATA '') THEN  ''UN''    ');
    SQL.Add('          WHEN PRODUTOS.PRO_UNIDADE =  ''PCTE '' THEN  ''PCT ''    ');
    SQL.Add('          WHEN PRODUTOS.PRO_UNIDADE IN ( ''BAND '', ''BD '') THEN  ''BJ ''    ');
    SQL.Add('          WHEN PRODUTOS.PRO_UNIDADE =  '' '' THEN  ''UN''    ');
    SQL.Add('          WHEN PRODUTOS.PRO_UNIDADE =  ''KG. '' THEN  ''KG ''    ');
    SQL.Add('          WHEN PRODUTOS.PRO_UNIDADE =  ''LTRO '' THEN  ''LT ''    ');
    SQL.Add('          WHEN PRODUTOS.PRO_UNIDADE =  ''PET '' THEN  ''PT ''    ');
    SQL.Add('          ELSE PRODUTOS.PRO_UNIDADE      ');
    SQL.Add('        END AS DES_UNIDADE,   ');
    SQL.Add('          ');
    SQL.Add('        ITENS.ENT_TOTAL / CASE WHEN COALESCE(ITENS.ENT_QTDE_VOL, 0) = 0 THEN 1 ELSE ITENS.ENT_QTDE_VOL END AS VAL_TABELA,   ');
    SQL.Add('        ITENS.ENT_DESCONTO / CASE WHEN COALESCE(ITENS.ENT_QTDE_VOL, 0) = 0 THEN 1 ELSE ITENS.ENT_QTDE_VOL END AS VAL_DESCONTO_ITEM,   ');
    SQL.Add('        ITENS.ENT_ACRESCIMO / CASE WHEN COALESCE(ITENS.ENT_QTDE_VOL, 0) = 0 THEN 1 ELSE ITENS.ENT_QTDE_VOL END AS VAL_ACRESCIMO_ITEM,   ');
    SQL.Add('        ITENS.ENT_IPI / CASE WHEN COALESCE(ITENS.ENT_QTDE_VOL, 0) = 0 THEN 1 ELSE ITENS.ENT_QTDE_VOL END AS VAL_IPI_ITEM,   ');
    SQL.Add('        ITENS.ENT_IPIPORC AS VAL_IPI_PER,   ');
    SQL.Add('        ITENS.ENT_SUB AS VAL_SUBST_ITEM,   ');
    SQL.Add('        ITENS.ENT_FRETE AS VAL_FRETE_ITEM,   ');
    SQL.Add('        ITENS.ENT_ICMS AS VAL_CREDITO_ICMS,   ');
    SQL.Add('        0 AS VAL_VENDA_VAREJO,   ');
    SQL.Add('        (ITENS.ENT_TOTAL + ITENS.ENT_ACRESCIMO) - ITENS.ENT_DESCONTO AS VAL_TABELA_LIQ,   ');
    SQL.Add('        '''' AS NUM_CGC,   ');
    SQL.Add('        ITENS.ENT_BASEICMS AS VAL_TOT_BC_ICMS,   ');
    SQL.Add('        0 AS VAL_TOT_OUTROS_ICMS,   ');
    SQL.Add('        ITENS.NAT_CODIGO AS CFOP,   ');
    SQL.Add('        0 AS VAL_TOT_ISENTO,   ');
    SQL.Add('        ITENS.ENT_BASEST AS VAL_TOT_BC_ST,   ');
    SQL.Add('        ITENS.ENT_SUB AS VAL_TOT_ST,   ');
    SQL.Add('        ITENS.ENT_ITEM AS NUM_ITEM,   ');
    SQL.Add('        0 AS TIPO_IPI,   ');
    SQL.Add('          ');
    SQL.Add('        CASE     ');
    SQL.Add('         WHEN COALESCE(ITENS.ENT_CLASFISCAL,'''') = '''' THEN ''99999999''   ');
    SQL.Add('         WHEN SUBSTRING(COALESCE(ITENS.ENT_CLASFISCAL,''''),1,8) = ''00000000'' THEN ''99999999''    ');
    SQL.Add('         ELSE SUBSTRING(ITENS.ENT_CLASFISCAL,1,8)    ');
    SQL.Add('        END AS NUM_NCM,   ');
    SQL.Add('          ');
    SQL.Add('        '''' AS DES_REFERENCIA   ');
    SQL.Add('       FROM    ');
    SQL.Add('        dbo.ENTRADA_ITEM AS ITENS   ');
    SQL.Add('       LEFT JOIN dbo.ENTRADA_ESTQ AS CAPA ON   ');
    SQL.Add('        (ITENS.ENT_REGISTRO = CAPA.ENT_REGISTRO)    ');
    SQL.Add('       LEFT JOIN dbo.PRODUTOS ON   ');
    SQL.Add('        (ITENS.PRO_CODIGO = PRODUTOS.PRO_CODIGO)   ');
    SQL.Add('   	LEFT JOIN dbo.TRIBUTACAO ON       ');
    SQL.Add('       (ITENS.TRI_ENTRADA = dbo.TRIBUTACAO.TRI_CODIGO)   ');
    SQL.Add('WHERE');
    SQL.Add(' CAST(CAPA.ENT_DATA_EMISSAO AS DATE) BETWEEN :INI AND :FIM');
    SQL.Add(' ORDER BY CAPA.ENT_NNOTA ASC ');

    Parameters.ParamByName('INI').Value := FORMATDATETIME('yyyy-mm-dd',
      DTPINICIAL.DATE);
    Parameters.ParamByName('FIM').Value := FORMATDATETIME('yyyy-mm-dd',
      DtpFinal.DATE);

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

        // Layout.SetValues(QryPrincipal2, NumLinha, RecordCount);

        (* if CodNf <> QryPrincipal2.FieldByName('CODENTRADANF').AsString then
          begin
          NumItem := 0;
          CodNf   := QryPrincipal2.FieldByName('CODENTRADANF').AsString;
          end;
        *)


        // Layout.SetValues(QryPrincipal, NumLinha, TotalCount);

        // Layout.FieldByName('NUM_ITEM').AsInteger := NumItem;

        Layout.FieldByName('COD_PRODUTO').AsString :=
          GerarPLU(QryPrincipal2.FieldByName('COD_PRODUTO').AsString);
        Layout.FieldByName('NUM_NCM').AsString :=
          StrRetNums(Layout.FieldByName('NUM_NCM').AsString);

        Layout.WriteLine;
      except
        On E: Exception do
          FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
      end;
      Next;
    end;
  end;
end;



procedure TFrmSmBomPrecoOurinhosGestora.GerarValorVenda;
var
  COD_EAN, VAL_VENDA: string;
begin
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    if CbxLoja.Text >= '1' then
    begin
      SQL.Add('   SELECT   ');
      SQL.Add('   	P.PRO_VENDA AS VAL_VENDA,   ');
      SQL.Add('   	P.PRO_BARRA AS COD_EAN   ');
      SQL.Add('   FROM dbo.PRODUTOS P   ');
      SQL.Add('   WHERE    ');
      SQL.Add('   	LEN(P.PRO_BARRA) >= 8   ');
      SQL.Add('   ORDER BY   ');
      SQL.Add('   	P.PRO_CODIGO ASC;   ');
    end;

    Open;
    First;

    NumLinha := 0;

    while not Eof do
    begin
      try
        if Cancelar then
          Break;

        Inc(NumLinha);

        COD_EAN := QryPrincipal2.FieldByName('COD_EAN').AsString;
        // GerarPLU(QryPrincipal2.FieldByName('COD_PRODUTO').AsString);
        VAL_VENDA := QryPrincipal2.FieldByName('VAL_VENDA').AsString;
        // GerarPLU(QryPrincipal2.FieldByName('COD_ASSOCIADO').AsString);

        Writeln(Arquivo, 'UPDATE TAB_PRODUTO_LOJA SET VAL_VENDA = ' + VAL_VENDA
          + ' WHERE COD_LOJA = ' + CbxLoja.Text +
          ' AND COD_PRODUTO IN (SELECT P.COD_PRODUTO FROM TAB_PRODUTO P WHERE P.COD_BARRA_PRINCIPAL = '''
          + COD_EAN + ''';');

        (* if NumLinha = 500 then
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
          Writeln(Arquivo, 'COMMIT WORK;'); *)

      except
        on E: Exception do
          FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
      end;
      Next;
    end;
    Writeln(Arquivo, 'COMMIT WORK;');
    Close;
  end;
end;

procedure TFrmSmBomPrecoOurinhosGestora.GeraCustoRep;
var
  COD_EAN, VAL_CUSTO_REP: string;
begin
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('   SELECT   ');
    SQL.Add('   	CONVERT (DEC(10,2), P.PRO_CUSTOREAL) AS VAL_CUSTO_REP,');
    SQL.Add('   	P.PRO_BARRA AS COD_EAN   ');
    SQL.Add('   FROM dbo.PRODUTOS P   ');
    SQL.Add('   WHERE    ');
    SQL.Add('   	LEN(P.PRO_BARRA) >= 8   ');
    SQL.Add('   ORDER BY   ');
    SQL.Add('   	P.PRO_CODIGO ASC;   ');

    Open;
    First;

    NumLinha := 0;

    while not Eof do
    begin
      try
        if Cancelar then
          Break;

        Inc(NumLinha);

        COD_EAN := QryPrincipal2.FieldByName('COD_EAN').AsString;
        VAL_CUSTO_REP := QryPrincipal2.FieldByName('VAL_CUSTO_REP').AsString;

        Writeln(Arquivo, 'UPDATE TAB_PRODUTO_LOJA SET VAL_CUSTO_REP = ' +
          VAL_CUSTO_REP + ' WHERE COD_LOJA = ' + CbxLoja.Text +
          ' AND COD_PRODUTO IN (SELECT P.COD_PRODUTO FROM TAB_PRODUTO P WHERE P.COD_BARRA_PRINCIPAL = '''
          + COD_EAN + ''');');

        (* if NumLinha = 500 then
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
          Writeln(Arquivo, 'COMMIT WORK;'); *)

      except
        on E: Exception do
          FrmProgresso.AdicionarLog(QryPrincipal.RecNo, 'E', E.Message);
      end;
      Next;
    end;
    Writeln(Arquivo, 'COMMIT WORK;');
    Close;
  end;
end;

procedure TFrmSmBomPrecoOurinhosGestora.GeraEstoqueVenda;
var
  COD_EAN, QTD_EST_VDA: string;
  // codTest
  // COD_CLI_AUTORIZ: integer;
begin
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('   		 SELECT      ');
    SQL.Add('             	P.PRO_ESTOQUE AS QTD_EST_VDA,      ');
    SQL.Add('             	P.PRO_BARRA AS COD_EAN      ');
    SQL.Add('             FROM dbo.PRODUTOS P      ');
    SQL.Add('             WHERE       ');
    SQL.Add('             	LEN(P.PRO_BARRA) >= 8      ');
    SQL.Add('             ORDER BY      ');
    SQL.Add('             	P.PRO_CODIGO ASC;     ');

    Open;
    First;

    NumLinha := 0;

    while not Eof do
    begin
      try
        if Cancelar then
          Break;

        Inc(NumLinha);

        // NUM_CGC := (QryPrincipal.FieldByName('NUM_CGC').AsString);
        // COD_CLI_AUTORIZ := StrToInt((QryPrincipal.FieldByName('COD_CLIENTE_AUTORIZ').AsString));
        // Inc(COD_CLI_AUTORIZ);
        // codTest := IntToStr(NumLinha);

        COD_EAN := QryPrincipal2.FieldByName('COD_EAN').AsString;
        QTD_EST_VDA := QryPrincipal2.FieldByName('QTD_EST_VDA').AsString;

        Writeln(Arquivo, 'UPDATE TAB_PRODUTO_LOJA SET QTD_EST_VDA = ' +
          QTD_EST_VDA + ' WHERE COD_LOJA = ' + CbxLoja.Text +
          ' AND COD_PRODUTO IN (SELECT P.COD_PRODUTO FROM TAB_PRODUTO P WHERE P.COD_BARRA_PRINCIPAL = '''
          + COD_EAN + ''');');
        // Writeln(Arquivo,'INSERT INTO TAB_CLIENTE_AUTORIZ (COD_CLIENTE, NUM_CGC_CPF, DES_AUTORIZ, FLG_REQUISICAO, DTA_INCLUSAO, USUARIO, COD_STATUS_CHEQUE, COD_STATUS_CONV, COD_CLIENTE_AUTORIZ) VALUES ('+QryPrincipal.FieldByName('COD_CLIENTE').AsString+','''+NUM_CGC+''','''+QryPrincipal.FieldByName('DES_AUTORIZ').AsString+''',''N'',''01/02/2023'', NULL, NULL, NULL, '+codTest+');' );

        (* if NumLinha = 500 then
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
          Writeln(Arquivo, 'COMMIT WORK;'); *)

      except
        on E: Exception do
          FrmProgresso.AdicionarLog(QryPrincipal.RecNo, 'E', E.Message);
      end;
      Next;
    end;
    Writeln(Arquivo, 'COMMIT WORK;');
    Close;
  end;
end;

procedure TFrmSmBomPrecoOurinhosGestora.GerarAssociado;
var
  COD_PRODUTO, COD_ASSOCIADO: string;
begin

  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('   SELECT   ');
    SQL.Add('   	PRO_CODIGO AS COD_PRODUTO,   ');
    SQL.Add('   	PRO_PAI AS COD_ASSOCIADO   ');
    SQL.Add('   FROM   ');
    SQL.Add('   	dbo.PRODUTOS   ');
    SQL.Add('   WHERE PRO_PAI <> 0   ');
    SQL.Add('   AND PRO_PAI IS NOT NULL   ');
    SQL.Add('   AND PRO_DESCRICAO IS NOT NULL    ');
    SQL.Add('   ORDER BY    ');
    SQL.Add('   	PRO_CODIGO ASC ;     ');

    Open;
    First;

    NumLinha := 0;

    while not Eof do
    begin
      try
        if Cancelar then
          Break;

        Inc(NumLinha);

        COD_PRODUTO := GerarPLU(QryPrincipal2.FieldByName('COD_PRODUTO')
          .AsString);
        // COD_ASSOCIADO  := GerarPLU(QryPrincipal2.FieldByName('COD_ASSOCIADO').AsString);

        Writeln(Arquivo, 'UPDATE TAB_PRODUTO_LOJA SET COD_ASSOCIADO = ' +
          QryPrincipal2.FieldByName('COD_ASSOCIADO').AsString +
          ' WHERE COD_PRODUTO ' + COD_PRODUTO + ';');

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

      except
        on E: Exception do
          FrmProgresso.AdicionarLog(QryPrincipal2.RecNo, 'E', E.Message);
      end;
      Next;
    end;
    Writeln(Arquivo, 'COMMIT WORK;');
    Close;

  end;

end;

procedure TFrmSmBomPrecoOurinhosGestora.GeraObsForn;
var
  COD_FORN, DES_OBS1, DES_OBS2, DES_CABEC1, DES_CABEC2: string;
  // codTest
  // COD_CLI_AUTORIZ: integer;
begin
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

     SQL.Add('           SELECT       ');
     SQL.Add('           	F.FOR_CODIGO AS COD_FORNECEDOR,      ');
     SQL.Add('           	''OBSERVACOES FORNECEDOR'' AS DES_CABEC1,      ');
     SQL.Add('           	 CHAR(13)+Char(10) + COALESCE (REPLACE(UPPER(CAST(F.FOR_OBS AS VARCHAR(500))),CHAR ( 13 ) +  Char ( 10 ), ''''), '''') AS DES_OBS1,      ');
     SQL.Add('           	 CHAR(13)+Char(10) +''OBSERVACOES PEDIDO'' AS DES_CABEC2,      ');
     SQL.Add('           	 CHAR(13)+Char(10) + COALESCE (REPLACE(UPPER(CAST(F.FOR_OBS_PEDIDO AS VARCHAR(500))),CHAR ( 13 ) +  Char ( 10 ), ''''), '''') AS DES_OBS2      ');
     SQL.Add('           FROM FORNECEDOR F      ');

    Open;
    First;

    NumLinha := 0;

    while not Eof do
    begin
      try
        if Cancelar then
          Break;

        Inc(NumLinha);

        // NUM_CGC := (QryPrincipal.FieldByName('NUM_CGC').AsString);
        // COD_CLI_AUTORIZ := StrToInt((QryPrincipal.FieldByName('COD_CLIENTE_AUTORIZ').AsString));
        // Inc(COD_CLI_AUTORIZ);
        // codTest := IntToStr(NumLinha);

        COD_FORN := QryPrincipal2.FieldByName('COD_FORNECEDOR').AsString;
        DES_CABEC1 := QryPrincipal2.FieldByName('DES_CABEC1').AsString;
        DES_OBS1 := StrSubstLtsAct(QryPrincipal2.FieldByName('DES_OBS1').AsString);
        DES_CABEC2 := QryPrincipal2.FieldByName('DES_CABEC2').AsString;
        DES_OBS2 := StrSubstLtsAct(QryPrincipal2.FieldByName('DES_OBS2').AsString);

        Writeln(Arquivo, 'UPDATE TAB_FORNECEDOR SET DES_OBSERVACAO = ''' + DES_CABEC1 + '''||chr(13)||''' + DES_OBS1 + '''||chr(13)||'''+DES_CABEC2+ '''||chr(13)||'''+ DES_OBS2 +''' WHERE COD_FORNECEDOR = ' + COD_FORN +';');
        // Writeln(Arquivo,'INSERT INTO TAB_CLIENTE_AUTORIZ (COD_CLIENTE, NUM_CGC_CPF, DES_AUTORIZ, FLG_REQUISICAO, DTA_INCLUSAO, USUARIO, COD_STATUS_CHEQUE, COD_STATUS_CONV, COD_CLIENTE_AUTORIZ) VALUES ('+QryPrincipal.FieldByName('COD_CLIENTE').AsString+','''+NUM_CGC+''','''+QryPrincipal.FieldByName('DES_AUTORIZ').AsString+''',''N'',''01/02/2023'', NULL, NULL, NULL, '+codTest+');' );

        (* if NumLinha = 500 then
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
          Writeln(Arquivo, 'COMMIT WORK;'); *)

      except
        on E: Exception do
          FrmProgresso.AdicionarLog(QryPrincipal.RecNo, 'E', E.Message);
      end;
      Next;
    end;
    Writeln(Arquivo, 'COMMIT WORK;');
    Close;
  end;
end;


procedure TFrmSmBomPrecoOurinhosGestora.GeraClassif;
var
  COD_CLASSIF, DES_CLASSIF, TIPO_CLASSIF : string;
begin
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

     SQL.Add('   SELECT    ');
     SQL.Add('   	GC.GRU_CODIGO + 1000 AS COD_CLASSIF,   ');
     SQL.Add('   	GC.GRU_DESCRICAO AS DES_CLASSIF,   ');
     SQL.Add('   	1 AS TIPO_CLASSIF   ');
     SQL.Add('   FROM GRU_CLIENTES GC   ');
     SQL.Add('      ');
     SQL.Add('   UNION ALL   ');
     SQL.Add('      ');
     SQL.Add('   SELECT    ');
     SQL.Add('   	GF.GRU_CODIGO + 2000 AS COD_CLASSIF,   ');
     SQL.Add('   	GF.GRU_DESCRICAO AS DES_CLASSIF,   ');
     SQL.Add('   	2 AS TIPO_CLASSIF   ');
     SQL.Add('   FROM GRU_FORNECEDOR GF   ');


    Open;
    First;

    NumLinha := 0;

    while not Eof do
    begin
      try
        if Cancelar then
          Break;

        Inc(NumLinha);

        // NUM_CGC := (QryPrincipal.FieldByName('NUM_CGC').AsString);
        // COD_CLI_AUTORIZ := StrToInt((QryPrincipal.FieldByName('COD_CLIENTE_AUTORIZ').AsString));
        // Inc(COD_CLI_AUTORIZ);
        // codTest := IntToStr(NumLinha);

        COD_CLASSIF  := QryPrincipal2.FieldByName('COD_CLASSIF').AsString;
        DES_CLASSIF  := Trim(UpperCase(QryPrincipal2.FieldByName('DES_CLASSIF').AsString));
        TIPO_CLASSIF := QryPrincipal2.FieldByName('TIPO_CLASSIF').AsString;

        Writeln(Arquivo, 'INSERT INTO TAB_CLASSIFICACAO C (C.COD_CLASSIF, C.DES_CLASSIF, TIPO_CLASSIF) VALUES ('+COD_CLASSIF+', '''+DES_CLASSIF+''', '+TIPO_CLASSIF+');');

        (* if NumLinha = 500 then
          Writeln(Arquivo, 'COMMIT WORK;');                        S
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
          Writeln(Arquivo, 'COMMIT WORK;'); *)

      except
        on E: Exception do
          FrmProgresso.AdicionarLog(QryPrincipal.RecNo, 'E', E.Message);
      end;
      Next;
    end;
    Writeln(Arquivo, 'COMMIT WORK;');
    Close;
  end;
end;

procedure TFrmSmBomPrecoOurinhosGestora.criaCsv;
begin
  with QryPrincipal2 do
  begin
    Close;
    SQL.Clear;

    SQL.Add('   SELECT      ');
    SQL.Add('   	P.PRO_CODIGO AS COD_PRODUTO,      ');
    SQL.Add('     	P.PRO_DESCRICAO AS DES_PRODUTO,   ');
    SQL.Add('   	''0'' AS VAL_VENDA,   ');
    SQL.Add('   	''0'' AS VAL_CUSTO_REP,   ');
    SQL.Add('   	''0'' AS VAL_OFERTA,   ');
    SQL.Add('   	''0'' AS QTD_EST_VDA,   ');
    SQL.Add('   	''0'' AS VAL_VENDA_2   ');
    SQL.Add('   FROM dbo.PRODUTOS P      ');
    SQL.Add('   WHERE       ');
    SQL.Add('    	LEN(P.PRO_BARRA) <= 8      ');
    SQL.Add('   AND       ');
    SQL.Add('   	P.PRO_DESCRICAO IS NOT NULL      ');
    SQL.Add('   ORDER BY      ');
    SQL.Add('      	TRIM(P.PRO_DESCRICAO) ASC;    ');

    Open;
    First;
  end;
end;

procedure TFrmSmBomPrecoOurinhosGestora.expCsv;
VAR
  Slista: TSTRINGLIST;
  Ncampo: Integer;
  CLinha: string;
  Planilha: Variant;
  SQL: TADOQuery;
begin
  SQL := QryPrincipal2;
  Slista := TSTRINGLIST.Create;
  CLinha := '';

  for Ncampo := 0 to SQL.Fields.Count - 1 do
    CLinha := CLinha + SQL.Fields[Ncampo].DisplayLabel + ';';
  Slista.Add(CLinha);
  SQL.First;
  while not SQL.Eof do
  begin
    CLinha := '';
    for Ncampo := 0 to SQL.Fields.Count - 1 do
      CLinha := CLinha + SQL.Fields[Ncampo].DisplayText + ';';
    Slista.Add(CLinha);
    SQL.Next
  end;
  Slista.SaveToFile(ExtractFilePath(EdtCamArquivo.Text) +
    '/PLANILHAS/EXPORTA/PRODUTOS.CSV');
end;

procedure TFrmSmBomPrecoOurinhosGestora.impCsv;
var
  csv: TextFile;
  quebra: TSTRINGLIST;
  arq: string;
  contador: Integer;
  COD_PRODUTO, DES_PRODUTO, VAL_VENDA, VAL_CUSTO_REP, VAL_OFERTA, QTD_EST_VDA, VAL_VENDA_2: String;
begin
try
  quebra := TSTRINGLIST.Create;
  AssignFile(csv, ExtractFilePath(EdtCamArquivo.Text)+'/PLANILHAS/IMPORTA/PRODUTOS.CSV'); // ExtractFilePath(EdtCamArquivo.Text)+'/PLANILHAS/EXPORTA/PRODUTOS.TXT');
  Reset(csv);
  contador := 1;
  AssignFile(Arquivo, ExtractFilePath(EdtCamArquivo.Text) + '\SCRIPT-FILIAIS\SCRIPT_ATUALIZA_PRODUTOS_EAN_MENOR_Q8_LJ'+CbxLoja.Text+'.TXT' );
  Rewrite(Arquivo);


  while not Eoln(csv) do
  begin
    quebra.Clear;
    Readln(csv, arq);
    ExtractStrings([';'], [], pchar(arq), quebra);
    if contador = 1 then
    begin
      contador := contador + 1;
    end
    else
    begin
      COD_PRODUTO   := GerarPLU(quebra.Strings[0]);
      VAL_VENDA     := StrReplace(quebra.Strings[2], ',', '.');
      VAL_CUSTO_REP := StrReplace(quebra.Strings[3], ',', '.');
      VAL_OFERTA    := StrReplace(quebra.Strings[4], ',', '.');
      QTD_EST_VDA   := StrReplace(quebra.Strings[5], ',', '.');
      VAL_VENDA_2   := StrReplace(quebra.Strings[6], ',', '.');

      Writeln(Arquivo,'UPDATE TAB_PRODUTO_LOJA SET VAL_VENDA = '+VAL_VENDA+ ', VAL_CUSTO_REP = ' +VAL_CUSTO_REP+ ', VAL_OFERTA = ' +VAL_OFERTA+ ', QTD_EST_VDA = '+QTD_EST_VDA+', VAL_VENDA_2 = '+VAL_VENDA_2 + ' WHERE COD_LOJA = '+CbxLoja.Text+' AND COD_PRODUTO = '''+COD_PRODUTO+''';');
    end;

   Next;
  end;
  CloseFile(Arquivo);
  CloseFile(csv);
  Application.MessageBox('Arquivo Gerado com Sucesso!' ,'Gerador - Sucesso',  MB_OK);
except
  on EInOutError do
   Application.MessageBox('O arquivo Produtos.csv está aberto, feche-o e tente novamente!','Gerador - Atenção', MB_ICONWARNING + MB_OK);
  on EStringListError do
   Application.MessageBox('Existe alguma(as) coluna(as) em branco/vazia. As colunas referente há valores não utilizadas devem estar preenchidas com 0' ,'Gerador - Atenção', MB_ICONWARNING + MB_OK);
end;

end;

procedure TFrmSmBomPrecoOurinhosGestora.criaDiretorioPlanilha;
VAR
  Diretorio, Exporta, Importa: String;
begin
  Diretorio := (ExtractFilePath(EdtCamArquivo.Text) + '/PLANILHAS');
  if not DirectoryExists(Diretorio) then
  begin
    ForceDirectories(Diretorio);
  end;

  Exporta := (Diretorio + '/EXPORTA');
  if not DirectoryExists(Exporta) then
  begin
    ForceDirectories(Exporta);
  end;

  Importa := (Diretorio + '/IMPORTA');
  if not DirectoryExists(Importa) then
  begin
    ForceDirectories(Importa);
  end;
end;

procedure TFrmSmBomPrecoOurinhosGestora.dataUltimaBuild;
var
dataArquivo, dataExtraida : TDateTime;
 horaExtraida : String;
begin
 DataArquivo:= FileDateToDateTime(FileAge(Application.ExeName));
 LblBuild.Caption := (FormatDateTime('dd/MM', DataArquivo) + ' ás ' + FormatDateTime('hh:mm', DataArquivo));

 //LblBuild.Caption := ((DateToStr(DataArquivo))) + '' + (TimeToStr(DataArquivo)));
 //Showmessage(TimeToStr(DataArquivo));
end;


procedure TFrmSmBomPrecoOurinhosGestora.validaMatriz;
var
  EdtMat: String;
begin

  EdtMat := ArqConf.ReadString('EdtMatr', 'EdtMatriz', '');

  if EdtMat = '' then
  begin
    EdtMatriz.Visible := True;
    ShpMatriz.Visible := True;
  end
  else
  begin
    ShpMatriz.Visible := False;
    FlgValidaMatriz := True;
  end;
end;

procedure TFrmSmBomPrecoOurinhosGestora.validaFiliais;
begin
  if CbxLoja.Text = EdtMatriz.Text then
  begin
    Application.MessageBox
      ('Escolha uma loja diferente da Matriz para gerar os Script',
      'Gerador - Atenção', MB_ICONWARNING + MB_OK);
    CbxLoja.SetFocus;
    ShpLoja.Visible := True;
  end
  else
  begin
    ShpLoja.Visible := False;
    FlgValidaFiliais := True;
  end;
end;

procedure TFrmSmBomPrecoOurinhosGestora.criaDiretorioUnic;
var
  Diretorio: string;
begin
  Diretorio := (ExtractFilePath(EdtCamArquivo.Text) + '/SCRIPT-FILIAIS');

  if not DirectoryExists(Diretorio) then
  begin
    ForceDirectories(Diretorio);
  end;
end;

procedure TFrmSmBomPrecoOurinhosGestora.criaDiretorioScriptUpdate;
var
  Diretorio: string;
begin
  Diretorio := (ExtractFilePath(EdtCamArquivo.Text) + '/SCRIPT-UPDATE');

  if not DirectoryExists(Diretorio) then
  begin
    ForceDirectories(Diretorio);
  end;
end;

procedure TFrmSmBomPrecoOurinhosGestora.geraFiliais;
begin

  if FlgAtualizaValVenda then
  begin
    GerarValorVenda;
    // Exit;
  end;

  if FlgAtualizaCustoRep then
  begin
    GeraCustoRep;
    // Exit;
  end;

  if FlgAtualizaEstoque then
  begin
    GeraEstoqueVenda;
    // Exit;
  end;

  if FlgAtualizaObsForn then
  begin
    GeraObsForn;
  end;

  if FlgAtualizaClassif then
  begin
    GeraClassif;
  end;
end;

procedure TFrmSmBomPrecoOurinhosGestora.CbxLojaChange(Sender: TObject);
begin
  inherited;
    ArqConf.WriteString('EdtLoja', 'EdtLoja', CbxLoja.Text);
end;

procedure TFrmSmBomPrecoOurinhosGestora.CkbScriptClick(Sender: TObject);
begin
  inherited;
  Memo1.Visible := True;
  BtnCopiar.Visible := True;

  if CkbScript.Checked = False then
  begin
    Memo1.Visible := False;
    BtnCopiar.Visible := False;
  end;
end;

procedure TFrmSmBomPrecoOurinhosGestora.CkbUnificacaoClick(Sender: TObject);
begin
    ArqConf.WriteBool('CkBox', 'ckboxUn', CkbUnificacao.Checked);
    if CkbUnificacao.Checked = True then
    begin
      if EdtCamArquivo.Text <> '' then
      begin
      CkbUnificacao.ShowHint := False;
      AbaUnificacao.Enabled := True;
      EdtMatriz.Visible := True;
      lblMatriz.Visible := True;
      criaDiretorioUnic;
      validaMatriz;
      end
      else
      begin
      CkbUnificacao.ShowHint := True;
      CkbUnificacao.Checked := False;
      end;
    end;


    if CkbUnificacao.Checked = False then
    begin
      AbaUnificacao.Enabled := False;
      EdtMatriz.Visible := False;
      ShpMatriz.Visible := False;
      lblMatriz.Visible := False;
    end;
end;

procedure TFrmSmBomPrecoOurinhosGestora.EdtMatrizChange(Sender: TObject);
begin
  inherited;
  ArqConf.WriteString('EdtMatr', 'EdtMatriz', EdtMatriz.Text);

end;

procedure TFrmSmBomPrecoOurinhosGestora.EdtMatrizMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
  inherited;
  if EdtMatriz.Text <> '' then
  begin
    ShpMatriz.Visible := False;
  end;
  validaFiliais;
end;

procedure TFrmSmBomPrecoOurinhosGestora.FormCreate(Sender: TObject);
VAR
EdtMatri : String;
begin
  inherited;
  CkbUnificacao.Checked := ArqConf.ReadBool('CkBox', 'ckboxUn', False);
  CbxLoja.Text          := ArqConf.ReadString('EdtLoja', 'EdtLoja', '');
  EdtMatriz.Text        := ArqConf.ReadString('EdtMatr', 'EdtMatriz', '');
  dataUltimaBuild;
end;

procedure TFrmSmBomPrecoOurinhosGestora.FormMouseEnter(Sender: TObject);
begin
  inherited;
  if (CkbUnificacao.Checked) And (EdtMatriz.Text = '') then
  begin
    EdtMatriz.SetFocus;
    EdtMatriz.Hint := 'Informe a Loja';
    EdtMatriz.ShowHint := True;
  end;
end;

procedure TFrmSmBomPrecoOurinhosGestora.btnGeraValVendaClick(Sender: TObject);
begin
  inherited;
  FlgAtualizaValVenda := True;
  BtnGerar.Click;
  FlgAtualizaValVenda := False;
end;

procedure TFrmSmBomPrecoOurinhosGestora.btnImportaPlanilhaClick(Sender: TObject);
begin
  inherited;
  impCsv;
end;

procedure TFrmSmBomPrecoOurinhosGestora.btnGeraClassifClick(Sender: TObject);
begin
  inherited;
  FlgAtualizaClassif := True;
  BtnGerar.Click;
  FlgAtualizaClassif := False;
end;

procedure TFrmSmBomPrecoOurinhosGestora.btnGeraCustoRepClick(Sender: TObject);
begin
  inherited;
  FlgAtualizaCustoRep := True;
  BtnGerar.Click;
  FlgAtualizaCustoRep := False;
end;

procedure TFrmSmBomPrecoOurinhosGestora.btnGeraEstoqueClick(Sender: TObject);
begin
  inherited;
  FlgAtualizaEstoque := True;
  BtnGerar.Click;
  FlgAtualizaEstoque := False;
end;

procedure TFrmSmBomPrecoOurinhosGestora.btnGeraObsFornClick(Sender: TObject);
begin
  inherited;
  criaDiretorioScriptUpdate;
  FlgAtualizaObsForn := True;
  BtnGerar.Click;
  FlgAtualizaObsForn := False;
end;

procedure TFrmSmBomPrecoOurinhosGestora.btnGeraPlanilhaClick(Sender: TObject);
begin
  inherited;

  try
  if EdtCamArquivo.Text = '' then
  begin
    MessageDlg(
      'Você não selecionou o caminho em que será salvo o arquivo de texto.',
      mtError, [mbOK], 0);
    Exit;
  end;
    Memo1.Clear;
    criaDiretorioPlanilha;
    criaCsv;
    expCsv;
    Application.MessageBox('Arquivo gerado com sucesso!','Gerador - Sucesso' , MB_OK);
  except
    on EOleException do
    Application.MessageBox('Não foi possivel montar o arquivo.csv! Verifique o Script utilizado...','Gerador - Atenção' , MB_ICONERROR + MB_OK);
    on EFCreateError do
    Application.MessageBox('Não foi possivel salvar o arquivo.csv! Verifique se o mesmo está aberto em outro software...','Gerador - Atenção' , MB_ICONERROR + MB_OK);
  end;

end;



procedure TFrmSmBomPrecoOurinhosGestora.AbaUnificacaoMouseEnter(Sender: TObject);
begin
  inherited;
  validaFiliais;
end;

procedure TFrmSmBomPrecoOurinhosGestora.BntAbrirPlanilhaClick(Sender: TObject);
begin
  //ShellExecute (handle, 'explore', 'E:\Nossos Documentos\Projeto Fisioterapia PEP - PROMED\POPs', 'NULL', 'NULL', SW_SHOWNORMAL);
  ShellExecute(Handle, 'OPEN', PChar(ExtractFilePath (EdtCamArquivo.Text)+'/PLANILHAS/EXPORTA'), nil, nil, SW_NORMAL);
end;

procedure TFrmSmBomPrecoOurinhosGestora.BntAbrirPlanilhaImpClick(Sender: TObject);
begin
  inherited;
  ShellExecute(Handle, 'OPEN', PChar(ExtractFilePath (EdtCamArquivo.Text)+'/PLANILHAS/IMPORTA'), nil, nil, SW_NORMAL);
end;

procedure TFrmSmBomPrecoOurinhosGestora.BtnCopiarClick(Sender: TObject);
begin
  inherited;
  Memo1.SelectAll;
  Memo1.CopyToClipboard;
end;

end.
