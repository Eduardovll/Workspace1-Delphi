unit UFrmSetFiscal;

interface

uses UClasses;

  type
    TFiscal = class
  public
    Function GetFieldTributacoes        : TLayout;

    Function GetFieldNFornecedores      : TLayout;
    Function GetFieldItensNFornecedores : TLayout;

    Function GetFieldINFClientes     : TLayout;
    Function GetFieldItensNFClientes : TLayout;

    Function GetFieldTransferencia      : TLayout;
    Function GetFieldItensTransferencia : TLayout;

    Function GetFieldOutrasNFs      : TLayout;
    Function GetFieldItensOutrasNFs : TLayout;

    Function GetFieldDadoFiscais: TLayout;
  end;

  var Layout :TLayout;

implementation

  { TFiscal }
function TFiscal.GetFieldDadoFiscais: TLayout;
begin
  with Layout do
  begin
   Layout := TLayout.Create('Dados Fiscais de NFs','OUTRAS_NF_ITENS');
  end;
end;

function TFiscal.GetFieldINFClientes: TLayout;
begin
  with Layout do
  begin
  Layout := TLayout.Create('Notas Fiscais Cliente','EMISSAO_NF');
    AddField('COD_CLIENTE',10 ,'0',tpLeft, ftpInteger );
    AddField('NUM_NF_CLI',10 ,'0',tpLeft, ftpInteger );
    AddField('NUM_SERIE_NF',8 ,' ',tpRight, ftpString );
    AddField('CFOP',10 ,'0',tpLeft, ftpInteger );
    AddField('TIPO_NF',1 ,'0',tpLeft, ftpString );
    AddField('DES_ESPECIE',3 ,' ',tpRight, ftpString );
    AddField('VAL_TOTAL_NF',12 ,'0',tpLeft, ftpFloat,'0.00' );
    AddField('DTA_EMISSAO',10 ,' ',tpRight, ftpDateTime,'dd/mm/yyyy' );
    AddField('DTA_ENTRADA',10 ,' ',tpRight, ftpDateTime,'dd/mm/yyyy' );
    AddField('VAL_TOTAL_IPI',12 ,'0',tpLeft, ftpFloat,'0.00' );
    AddField('VAL_FRETE',12 ,'0',tpLeft, ftpFloat,'0.00' );
    AddField('VAL_ENC_FINANC',12 ,'0',tpLeft, ftpFloat,'0.00' );
    AddField('VAL_DESC_FINANC',12 ,'0',tpLeft, ftpFloat,'0.00' );
    AddField('NUM_CGC',19 ,' ',tpRight, ftpString );
    AddField('DES_NATUREZA',35 ,' ',tpRight, ftpString );
    AddField('DES_OBSERVACAO',100 ,' ',tpRight, ftpString );
    AddField('FLG_CANCELADA',1 ,' ',tpRight, ftpString );
    AddField('NUM_CHAVE_ACESSO',100 ,' ',tpRight, ftpString );
  end;
 Result := Layout;
end;

function TFiscal.GetFieldItensNFClientes: TLayout;
begin
  with Layout do
  begin
    Layout := TLayout.Create('Itens Notas Fiscais Cliente','EMISSAO_NF_ITENS');
    AddField('COD_CLIENTE',10 ,'0',tpLeft, ftpInteger );
    AddField('NUM_NF_CLI',10 ,'0',tpLeft, ftpInteger );
    AddField('NUM_SERIE_NF',8 ,' ',tpRight, ftpString );
    AddField('COD_PRODUTO',8 ,'0',tpLeft, ftpString );
    AddField('COD_TRIBUTACAO',3 ,'0',tpLeft, ftpInteger );
    AddField('QTD_EMBALAGEM',12 ,'0',tpLeft, ftpFloat,'0.00' );
    AddField('QTD_ENTRADA',12 ,'0',tpLeft, ftpFloat,'0.000' );
    AddField('DES_UNIDADE',2 ,' ',tpLeft, ftpString );
    AddField('VAL_TABELA',12 ,'0',tpLeft, ftpFloat,'0.000' );
    AddField('VAL_DESCONTO_ITEM',12 ,'0',tpLeft, ftpFloat,'0.000' );
    AddField('VAL_ACRESCIMO_ITEM',12 ,'0',tpLeft, ftpFloat,'0.000' );
    AddField('VAL_IPI_ITEM',12 ,'0',tpLeft, ftpFloat,'0.000' );
    AddField('VAL_CREDITO_ICMS',12 ,'0',tpLeft, ftpFloat,'0.000' );
    AddField('VAL_TABELA_LIQ',12 ,'0',tpLeft, ftpFloat,'0.000' );
    AddField('VAL_CUSTO_REP',12 ,'0',tpLeft, ftpFloat,'0.000' );
    AddField('NUM_CGC',19 ,'0',tpLeft, ftpString );
    AddField('VAL_TOT_BC_ICMS',12 ,'0',tpLeft, ftpFloat,'0.000' );
    AddField('VAL_TOT_OUTROS_ICMS',12 ,'0',tpLeft, ftpFloat,'0.000' );
    AddField('COD_FISCAL',10 ,'0',tpLeft, ftpInteger );
    AddField('NUM_ITEM',10 ,'0',tpLeft, ftpInteger );
    AddField('TIPO_IPI',1 ,'0',tpLeft, ftpString );
  end;
 Result := Layout;
end;

function TFiscal.GetFieldItensNFornecedores: TLayout;
begin
  with Layout do
  begin
    Layout := TLayout.Create('Itens Notas Fiscais Fornecedores','ENTRADA_NF_ITENS');
    AddField('COD_FORNECEDOR',10 ,'0',tpLeft, ftpInteger );
    AddField('NUM_NF_FORN',10 ,'0',tpLeft, ftpInteger );
    AddField('NUM_SERIE_NF',8 ,' ',tpRight, ftpString );
    AddField('COD_PRODUTO',8 ,' ',tpLeft, ftpString );
    AddField('COD_TRIBUTACAO',3 ,'0',tpLeft, ftpInteger );
    AddField('QTD_EMBALAGEM',12 ,'0',tpLeft, ftpFloat,'0.00' );
    AddField('QTD_ENTRADA',12 ,'0',tpLeft, ftpFloat,'0.000' );
    AddField('DES_UNIDADE',2 ,' ',tpRight, ftpString );
    AddField('VAL_TABELA',12 ,'0',tpLeft, ftpFloat,'0.000' );
    AddField('VAL_DESCONTO_ITEM',12 ,'0',tpLeft, ftpFloat,'0.000' );
    AddField('VAL_ACRESCIMO_ITEM',12 ,'0',tpLeft, ftpFloat,'0.000');
    AddField('VAL_IPI_ITEM',12 ,'0',tpLeft, ftpFloat,'0.000' );
    AddField('VAL_SUBST_ITEM',12 ,'0',tpLeft, ftpFloat,'0.000' );
    AddField('VAL_FRETE_ITEM',12 ,'0',tpLeft, ftpFloat,'0.000' );
    AddField('VAL_CREDITO_ICMS',12 ,'0',tpLeft, ftpFloat,'0.000' );
    AddField('VAL_VENDA_VAREJO',12 ,'0',tpLeft, ftpFloat,'0.000' );
    AddField('VAL_TABELA_LIQ',12 ,'0',tpLeft, ftpFloat,'0.000' );
    AddField('NUM_CGC',19 ,' ',tpRight, ftpString );
    AddField('VAL_TOT_BC_ICMS',12 ,'0',tpLeft, ftpFloat,'0.000' );
    AddField('VAL_TOT_OUTROS_ICMS',12 ,'0',tpLeft, ftpFloat,'0.000' );
    AddField('CFOP',10 ,'0',tpLeft, ftpString );
    AddField('VAL_TOT_ISENTO',12 ,'0',tpLeft, ftpFloat,'0.000' );
    AddField('VAL_TOT_BC_ST',12 ,'0',tpLeft, ftpFloat,'0.000' );
    AddField('VAL_TOT_ST',12 ,'0',tpLeft, ftpFloat,'0.000' );
    AddField('NUM_ITEM',10 ,'0',tpLeft, ftpInteger );
    AddField('TIPO_IPI',1 ,' ',tpRight, ftpString );
    AddField('NUM_NCM',8 ,'0',tpLeft, ftpString );
    AddField('DES_REFERENCIA',35 ,' ',tpRight, ftpString );
    AddField('VAL_TOT_ST_FCP', 12, '0', tpLeft, ftpFloat, '0.000' );
    AddField('VAL_DESP_ACESS_ITEM', 12, '0', tpLeft, ftpFloat, '0.000' );
    //AddField('VAL_IPI_PER', 12, '0', tpLeft, ftpFloat, '0.000' );
    AddField('DTA_VALIDADE',10 ,' ',tpRight, ftpDateTime, 'dd/mm/yyyy' );

  end;
 Result := Layout;
end;

function TFiscal.GetFieldItensOutrasNFs: TLayout;
begin
  with Layout do
  begin
   Layout := TLayout.Create('Itens de Outras Notas-Ficais','OUTRAS_NF_ITENS');
   AddField('COD_FORNECEDOR',10 ,'0',tpLeft, ftpInteger );
   AddField('NUM_NF_FORN',10 ,'0',tpLeft, ftpInteger );
   AddField('NUM_SERIE_NF',8 ,' ',tpRight, ftpString );
   AddField('NUM_ITEM',8 ,'0',tpLeft, ftpString );
   AddField('DES_ITEM',50 ,' ',tpRight, ftpString );
   AddField('COD_TRIBUTACAO',3 ,'0',tpLeft, ftpInteger );
   AddField('QTD_EMBALAGEM',12 ,'0',tpLeft, ftpFloat,'0.000' );
   AddField('QTD_ENTRADA',12 ,'0',tpLeft, ftpFloat,'0.000' );
   AddField('DES_UNIDADE',2 ,' ',tpRight, ftpString );
   AddField('VAL_TABELA',12 ,'0',tpLeft, ftpFloat,'0.00' );
   AddField('VAL_DESCONTO_ITEM',12 ,'0',tpLeft, ftpFloat,'0.00' );
   AddField('VAL_ACRESCIMO_ITEM',12 ,'0',tpLeft, ftpFloat,'0.00' );
   AddField('VAL_IPI_ITEM',12 ,'0',tpLeft, ftpFloat,'0.00' );
   AddField('VAL_CREDITO_ICMS',12 ,'0',tpLeft, ftpFloat,'0.00' );
   AddField('VAL_TABELA_LIQ',12 ,'0',tpLeft, ftpFloat,'0.00' );
   AddField('NUM_CGC',19 ,'0',tpLeft, ftpString );
   AddField('VAL_TOT_BC_ICMS',12 ,'0',tpLeft, ftpFloat,'0.00' );
   AddField('VAL_TOT_OUTROS_ICMS',12 ,'0',tpLeft, ftpFloat,'0.00' );
   AddField('COD_FISCAL',10 ,'0',tpLeft, ftpInteger );
  end;
 Result := Layout;
end;

function TFiscal.GetFieldItensTransferencia: TLayout;
begin
  with Layout do
  begin
    Layout := TLayout.Create('Itens Notas Ficais Transfência','TRANSF_NF_ITENS');
    AddField('COD_LOJA_TRANSF',10 ,'0',tpLeft, ftpInteger );
    AddField('NUM_NF_TRANSF',10 ,'0',tpLeft, ftpInteger );
    AddField('NUM_SERIE_NF',8 ,' ',tpRight, ftpString );
    AddField('COD_PRODUTO',8 ,'0',tpLeft, ftpString );
    AddField('COD_TRIBUTACAO',3 ,'0',tpLeft, ftpInteger );
    AddField('QTD_EMBALAGEM',12 ,'0',tpLeft, ftpFloat,'0.00' );
    AddField('QTD_ENTRADA',12 ,'0',tpLeft, ftpFloat,'0.000' );
    AddField('DES_UNIDADE',2 ,' ',tpRight, ftpString );
    AddField('VAL_TABELA',12 ,'0',tpLeft, ftpFloat,'0.000' );
    AddField('VAL_DESCONTO_ITEM',12 ,'0',tpLeft, ftpFloat,'0.000' );
    AddField('VAL_ACRESCIMO_ITEM',12 ,'0',tpLeft, ftpFloat,'0.000' );
    AddField('VAL_IPI_ITEM',12 ,'0',tpLeft, ftpFloat,'0.000' );
    AddField('VAL_CREDITO_ICMS',12 ,'0',tpLeft, ftpFloat,'0.000' );
    AddField('VAL_TABELA_LIQ',12 ,'0',tpLeft, ftpFloat,'0.000' );
    AddField('VAL_CUSTO_REP',12 ,'0',tpLeft, ftpFloat,'0.000' );
    AddField('NUM_CGC',19 ,'0',tpleft, ftpString );
    AddField('NUM_ITEM',4 ,'0',tpLeft, ftpInteger );
    AddField('VAL_TOT_BC_ICMS',12 ,'0',tpLeft, ftpFloat,'0.000' );
    AddField('VAL_TOT_OUTROS_ICMS',12 ,'0',tpLeft, ftpFloat,'0.000' );
    AddField('COD_FISCAL',10 ,'0',tpLeft, FtpInteger );
    AddField('TIPO_IPI',1 ,'0',tpRight, ftpString );
  end;
 Result := Layout;
end;

function TFiscal.GetFieldNFornecedores: TLayout;
begin
  with Layout do
  begin
    Layout := TLayout.Create('Notas Fiscais Fornecedores','ENTRADA_NF');
    AddField('COD_FORNECEDOR',10 ,'0',tpLeft, ftpInteger );
    AddField('NUM_NF_FORN',10 ,'0',tpLeft, ftpInteger );
    AddField('NUM_SERIE_NF',8 ,' ',tpRight, ftpString );
    AddField('NUM_SUBSERIE_NF',8 ,' ',tpRight, ftpString );
    AddField('CFOP',10 ,'0',tpLeft, ftpString );
    AddField('TIPO_NF',1 ,' ',tpRight, ftpString );
    AddField('DES_ESPECIE',3 ,' ',tpRight, ftpString );
    AddField('VAL_TOTAL_NF',12 ,'0',tpLeft, ftpFloat,'0.00' );
    AddField('DTA_EMISSAO',10 ,' ',tpRight, ftpDateTime,'dd/mm/yyyy' );
    AddField('DTA_ENTRADA',10 ,' ',tpRight, ftpDateTime,'dd/mm/yyyy' );
    AddField('VAL_TOTAL_IPI',12 ,'0',tpLeft, ftpFloat,'0.00' );
    AddField('VAL_VENDA_VAREJO',12 ,'0',tpLeft, ftpFloat,'0.00' );
    AddField('VAL_FRETE',12 ,'0',tpLeft, ftpFloat,'0.00' );
    AddField('VAL_ACRESCIMO',12 ,'0',tpLeft, ftpFloat,'0.00' );
    AddField('VAL_DESCONTO',12 ,'0',tpLeft, ftpFloat,'0.00' );
    AddField('NUM_CGC',19 ,' ',tpLeft, ftpString );
    AddField('VAL_TOTAL_BC',12 ,'0',tpLeft, ftpFloat,'0.00' );
    AddField('VAL_TOTAL_ICMS',12 ,'0',tpLeft, ftpFloat,'0.00' );
    AddField('VAL_BC_SUBST',12 ,'0',tpLeft, ftpFloat,'0.00' );
    AddField('VAL_ICMS_SUBST',12 ,'0',tpLeft, ftpFloat,'0.00' );
    AddField('VAL_FUNRURAL',12 ,'0',tpLeft, ftpFloat,'0.00' );
    AddField('COD_PERFIL',10 ,'0',tpLeft, ftpInteger );
    AddField('VAL_DESP_ACESS',12 ,'0',tpLeft, ftpFloat,'0.00' );
    AddField('FLG_CANCELADO',1 ,' ',tpRight, ftpString );
    AddField('DES_OBSERVACAO',100 ,' ',tpLeft, ftpString );
    AddField('NUM_CHAVE_ACESSO',100 ,' ',tpRight, ftpString );
    //AddField('VAL_TOT_ST_FCP',12 ,'0',tpLeft, ftpFloat,'0.00' );

  end;
 Result := Layout;
end;

function TFiscal.GetFieldOutrasNFs: TLayout;
begin
  with Layout do
  begin
    Layout := TLayout.Create('Outras Notas Fiscais','OUTRAS_NF');
    AddField('COD_FORNECEDOR',10 ,'0',tpLeft, ftpInteger );
    AddField('NUM_NF_FORN',10 ,'0',tpLeft, ftpInteger );
    AddField('NUM_SERIE_NF',8 ,' ',tpRight, ftpString );
    AddField('NUM_SUBSERIE_NF',8 ,' ',tpRight, ftpString );
    AddField('COD_FISCAL',10 ,'0',tpLeft, ftpInteger );
    AddField('TIPO_NF',1 ,'0',tpLeft, ftpInteger );
    AddField('DES_ESPECIE',3 ,' ',tpRight, ftpString );
    AddField('VAL_TOTAL_NF',12 ,'0',tpLeft, ftpFloat,'0.00' );
    AddField('DTA_EMISSAO',10 ,' ',tpRight, ftpDateTime,'dd/mm/yyyy' );
    AddField('DTA_ENTRADA',10 ,' ',tpRight, ftpDateTime,'dd/mm/yyyy' );
    AddField('VAL_TOTAL_IPI',12 ,'0',tpLeft, ftpFloat,'0.00' );
    AddField('VAL_VENDA_VAREJO',12 ,'0',tpLeft, ftpFloat,'0.00' );
    AddField('VAL_FRETE',12 ,'0',tpLeft, ftpFloat,'0.00' );
    AddField('VAL_ACRESCIMO',12 ,'0',tpLeft, ftpFloat,'0.00' );
    AddField('VAL_DESCONTO',12 ,'0',tpLeft, ftpFloat,'0.00' );
    AddField('NUM_CGC',19 ,'0',tpLeft, ftpString );
  end;
 Result := Layout;
end;

function TFiscal.GetFieldTransferencia: TLayout;
begin
  with Layout do
  begin
    Layout := TLayout.Create('Notas Ficais Transfência','TRANSF_NF');
    AddField('COD_LOJA_TRANSF',10 ,'0',tpLeft, ftpInteger );
    AddField('NUM_NF_TRANSF',10 ,'0',tpLeft, ftpInteger );
    AddField('NUM_SERIE_NF',8 ,' ',tpRight, ftpString );
    AddField('CFOP',10 ,'0',tpLeft, ftpInteger );
    AddField('TIPO_NF',1 ,'0',tpLeft, ftpInteger );
    AddField('DES_ESPECIE',3 ,' ',tpRight, ftpString );
    AddField('VAL_TOTAL_NF',12 ,'0',tpLeft, ftpFloat,'0.00' );
    AddField('DTA_EMISSAO',10 ,' ',tpRight, ftpDateTime, 'dd/mm/yyyy' );
    AddField('DTA_ENTRADA',10 ,' ',tpRight, ftpDateTime, 'dd/mm/yyyy' );
    AddField('VAL_FRETE',12 ,'0',tpLeft, ftpFloat, '0.00' );
    AddField('NUM_CGC',19 ,'0',tpLeft, ftpString );
    AddField('COD_ORIGEM',2 ,'0',tpLeft, ftpInteger );
    AddField('COD_DESTINO',2 ,'0',tpLeft, ftpInteger );
    AddField('NUM_SERIE_NF',4 ,' ',tpRight, ftpString );
    AddField('DES_NATUREZA',35 ,' ',tpRight, ftpString );
    AddField('FLG_CANCELADO',1 ,' ',tpRight, ftpString );
    AddField('NUM_CHAVE_ACESSO',100 ,'0',tpLeft, ftpString );
    AddField('COD_PERFIL',10 ,'0',tpLeft, ftpInteger);
  end;
 Result := Layout;
end;

function TFiscal.GetFieldTributacoes: TLayout;
begin
  with Layout do
  begin
    Layout := TLayout.Create('Tributações','TRIBUTACOES');
    AddField('COD_TRIBUTACAO',5 ,'0',tpLeft, ftpInteger );
    AddField('DES_TRIBUTACAO',50 ,' ',tpRight, ftpString );
    AddField('VAL_ICMS',10 ,'0',tpLeft, ftpFloat, '0.00' );
    AddField('VAL_REDUCAO_BASE_CALCULO',10 ,'0',tpLeft, ftpFloat, '0.00' );
    AddField('TIPO_TRIBUTACAO',2 ,'0',tpLeft, ftpInteger );
    AddField('DES_OBSERVACAO',200 ,' ',tpRight, ftpString );
  end;
 Result := Layout;
end;

end.
