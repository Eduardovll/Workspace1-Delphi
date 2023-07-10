unit UFrmSetOutros;

interface

  uses UClasses;

  type TFinanceiro = class
  public
    Function GetFieldFinanceiro : TLayout;
  end;

  type TOutros = class
  public
    Function GetFieldMapaResumo  : TLayout;
    Function GetFieldAjuste      : TLayout;
    Function GetFieldVendas      : TLayout;
    Function GetFieldPlanoContas : TLayout;
  end;

implementation

var Layout :TLayout;

{ TFiscal }
function TFinanceiro.GetFieldFinanceiro: TLayout;
begin
 with Layout do
 begin
   Layout := TLayout.Create('Títulos à Pagar/Receber','FINANCEIRO');
    AddField('TIPO_PARCEIRO', 1,'0',tpLeft, ftpInteger );
    AddField('COD_PARCEIRO', 10,'0',tpLeft, ftpInteger );
    AddField('TIPO_CONTA', 1,'0',tpLeft, ftpInteger );
    AddField('COD_ENTIDADE',10 ,'0',tpLeft, ftpString );
    AddField('NUM_DOCTO',25 ,' ',tpRight, ftpString );
    AddField('COD_BANCO',10 ,'0',tpLeft, ftpString );
    AddField('DES_BANCO',30 ,' ',tpRight, ftpString );
    AddField('DTA_EMISSAO',10 ,' ',tpRight, ftpDateTime, 'dd/mm/yyyy' );
    AddField('DTA_VENCIMENTO',10 ,' ',tpRight, ftpDateTime, 'dd/mm/yyyy' );
    AddField('VAL_PARCELA',12 ,'0',tpLeft, ftpFloat,'0.00' );
    AddField('VAL_JUROS',12 ,'0',tpLeft, ftpFloat,'0.00' );
    AddField('VAL_DESCONTO',12 ,'0',tpLeft, ftpFloat,'0.00' );
    AddField('FLG_QUITADO',1 ,' ',tpRight, ftpString );
    AddField('DTA_QUITADA',10 ,' ',tpRight, ftpDateTime, 'dd/mm/yyyy' );
    AddField('COD_CATEGORIA',10 ,'0',tpLeft, ftpInteger );
    AddField('COD_SUBCATEGORIA',10 ,'0',tpLeft, ftpInteger );
    AddField('NUM_PARCELA',4 ,'0',tpLeft, ftpString );
    AddField('QTD_PARCELA',4 ,'0',tpLeft, ftpString );
    AddField('COD_LOJA',10 ,'0',tpLeft, ftpInteger );
    AddField('NUM_CGC',19 ,' ',tpRight, ftpString );
    AddField('NUM_BORDERO',10 ,'0',tpLeft, ftpInteger );
    AddField('NUM_NF',10 ,'0',tpLeft, ftpString );
    AddField('NUM_SERIE_NF',8 ,' ',tpRight, ftpString );
    AddField('VAL_TOTAL_NF',12 ,'0',tpLeft, ftpFloat,'0.00' );
    AddField('DES_OBSERVACAO',200 ,' ',tpRight, ftpString );
    AddField('NUM_PDV',3 ,'0',tpLeft, ftpInteger );
    AddField('NUM_CUPOM_FISCAL',10 ,'0',tpLeft, ftpString );
    AddField('COD_MOTIVO',10 ,'0',tpLeft, ftpInteger );
    AddField('COD_CONVENIO',10 ,'0',tpLeft, ftpInteger );
    AddField('COD_BIN',10 ,'0',tpLeft, ftpInteger );
    AddField('DES_BANDEIRA',25 ,' ',tpRight, ftpString );
    AddField('DES_REDE_TEF',25 ,' ',tpRight, ftpString );
    AddField('VAL_RETENCAO',12 ,'0',tpLeft, ftpFloat );
    AddField('COD_CONDICAO',3 ,'0',tpLeft, ftpInteger );
    AddField('DTA_PAGTO',10 ,' ',tpRight, ftpDateTime, 'dd/mm/yyyy' );
    AddField('DTA_ENTRADA',10 ,' ',tpRight, ftpDateTime, 'dd/mm/yyyy' );
    AddField('NUM_NOSSO_NUMERO',11 ,' ',tpRight, ftpString );
    AddField('COD_BARRA',50 ,' ',tpLeft, ftpString );
    AddField('FLG_BOLETO_EMIT',1 ,' ',tpRight, ftpString );
    AddField('NUM_CGC_CPF_TITULAR',19 ,' ',tpRight, ftpString );
    AddField('DES_TITULAR',50 ,' ',tpRight, ftpString );
    AddField('NUM_CONDICAO',3 ,'0',tpLeft, ftpInteger );
    AddField('VAL_CREDITO',12 ,'0',tpLeft, ftpFloat, '0.00' ); //
    AddField('COD_BANCO_PGTO',10 ,'0',tpLeft, ftpInteger );
    AddField('DES_CC',50 ,' ',tpRight, ftpString );
    AddField('COD_BANDEIRA',10 ,'0',tpLeft, ftpInteger );
    AddField('DTA_PRORROGACAO',10 ,' ',tpRight, ftpDateTime, 'dd/mm/yyyy' );
    AddField('NUM_SEQ_FIN', 2, '0', tpLeft, ftpInteger );
    AddField('COD_COBRANCA', 10,'0',tpLeft, ftpInteger );
    AddField('DTA_COBRANCA',10 ,' ',tpRight, ftpDateTime, 'dd/mm/yyyy' );
    AddField('FLG_ACEITE',1 ,' ',tpRight, ftpString );
    AddField('TIPO_ACEITE', 1,'0',tpLeft, ftpInteger );

 end;
 Result := Layout;
end;

{ TOutros }
function TOutros.GetFieldAjuste: TLayout;
begin
  with Layout do
  begin
    Layout := TLayout.Create('Ajuste','AJUSTE');
    AddField('COD_PRODUTO',8 ,'0',tpLeft, ftpInteger );
    AddField('COD_AJUSTE',10 ,'0',tpLeft, ftpInteger );
    AddField('QTD_AJUSTE',12 ,'0',tpLeft, ftpFloat, '0.000' );
    AddField('QTD_ESTOQUE_ANT',12 ,'0',tpLeft, ftpFloat, '0.000' );
    AddField('QTD_ESTOQUE_POST',12 ,'0',tpLeft, ftpFloat, '0.000' );
    AddField('VAL_CUSTO_SICMS',12 ,'0',tpLeft, ftpFloat, '0.00' );
    AddField('VAL_CUSTO_MED',12 ,'0',tpLeft, ftpFloat, '0.00' );
    AddField('VAL_CUSTO_REP',12 ,'0',tpLeft, ftpFloat, '0.00' );
    AddField('DTA_AJUSTE',10 ,' ',tpRight, ftpDateTime, 'dd/mm/yyyy' );
  end;
 Result := Layout;
end;

function TOutros.GetFieldMapaResumo: TLayout;
begin
 with Layout do
 begin
  Layout := TLayout.Create('Mapa Resumo','MAPA_RESUMO');
  AddField('COD_LOJA',10 ,'0',tpLeft, ftpInteger );
  AddField('DTA_LANCAMENTO',10 ,'0',tpLeft, ftpDatetime );
  AddField('NUM_PDV',8 ,'0',tpLeft, ftpInteger );
  AddField('COD_MR',10 ,'0',tpLeft, ftpInteger );
  AddField('NUM_CUPOM_INICIAL',10 ,'0',tpLeft, ftpInteger );
  AddField('NUM_CUPOM_FINAL',10 ,'0',tpLeft, ftpInteger );
  AddField('VAL_GT_INICIAL',12 ,'0',tpLeft, ftpFloat, '0.00' );
  AddField('VAL_GT_FINAL',12 ,'0',tpLeft, ftpFloat, '0.00' );
  AddField('VAL_VENDA_BRUTA',12 ,'0',tpLeft, ftpFloat, '0.00' );
  AddField('VAL_CANCELAMENTO',12 ,'0',tpLeft, ftpFloat, '0.00' );
  AddField('VAL_DESCONTOS',12 ,'0',tpLeft, ftpFloat, '0.00' );
  AddField('VAL_SUBSTITUICAO',12 ,'0',tpLeft, ftpFloat, '0.00' );
  AddField('VAL_ISENTO',12 ,'0',tpLeft, ftpFloat, '0.00' );
  AddField('VAL_BC_ICMS_0',12 ,'0',tpLeft, ftpFloat, '0.00' );
  AddField('VAL_BC_ICMS_1',12 ,'0',tpLeft, ftpFloat, '0.00' );
  AddField('VAL_BC_ICMS_2',12 ,'0',tpLeft, ftpFloat, '0.00' );
  AddField('VAL_BC_ICMS_3',12 ,'0',tpLeft, ftpFloat, '0.00' );
  AddField('VAL_BC_ICMS_4',12 ,'0',tpLeft, ftpFloat, '0.00' );
  AddField('VAL_BC_ICMS_5',12 ,'0',tpLeft, ftpFloat, '0.00' );
  AddField('VAL_BC_ICMS_6',12 ,'0',tpLeft, ftpFloat, '0.00' );
  AddField('VAL_BC_ICMS_7',12 ,'0',tpLeft, ftpFloat, '0.00' );
  AddField('VAL_BC_ICMS_8',12 ,'0',tpLeft, ftpFloat, '0.00' );
  AddField('VAL_BC_ICMS_9',12 ,'0',tpLeft, ftpFloat, '0.00' );
  AddField('NUM_REDUCAO_Z',8 ,'0',tpLeft, ftpInteger );
  AddField('VAL_NAO_TRIBUT',12 ,'0',tpLeft, ftpFloat, '0.00' );
  AddField('VAL_IMPOSTO',12 ,'0',tpLeft, ftpFloat, '0.00' );
  AddField('NUM_CRO',8 ,'0',tpLeft, ftpInteger );
 end;
 Result := Layout;
end;

function TOutros.GetFieldPlanoContas: TLayout;
begin
 with Layout do
 begin
  Layout := TLayout.Create('Plano de Contas','PLCONTAS');
  AddField('CONTA_CONTABIL',25 ,' ',tpRight, ftpString );
  AddField('DES_PLANO_CONTA',40 ,' ',tpRight, ftpString );
  AddField('CONTA_REDUZIDA',5 ,'0',tpLeft, ftpInteger );
  AddField('CONTA_CONTABIL_SPED',25 ,' ',tpLeft, ftpString );
  AddField('COD_GRUPO_CTB',10 ,'0',tpLeft, ftpInteger );
  AddField('TIPO_CONTA',1 ,' ',tpRight, ftpString );
//AddField('CONTA_CONTABIL_REF', 25, ' ', tpRight, ftpString);
//AddField('DES_CONTA_CONTABIL', 50, ' ', tpRight, ftpString);
 end;
 Result := Layout;
end;

function TOutros.GetFieldVendas: TLayout;
begin
  with Layout do
  begin
    Layout := TLayout.Create('Vendas','VENDAS');
    AddField('COD_PRODUTO',8 ,'0',tpLeft, ftpString );
    AddField('COD_LOJA',10 ,'0',tpLeft, ftpInteger );
    AddField('IND_TIPO',1 ,'0',tpLeft, ftpInteger );
    AddField('NUM_PDV',10 ,'0',tpLeft, ftpString );
    AddField('QTD_TOTAL_PRODUTO',12 ,'0',tpLeft, ftpFloat, '0.000' );
    AddField('VAL_TOTAL_PRODUTO',12 ,'0',tpLeft, ftpFloat, '0.000' );
    AddField('VAL_PRECO_VENDA',12 ,'0',tpLeft, ftpFloat, '0.000' );
    AddField('VAL_CUSTO_REP',12 ,'0',tpLeft, ftpFloat, '0.000' );
    AddField('DTA_SAIDA',10 ,' ',tpRight, ftpDateTime, 'dd/mm/yyyy' );
    AddField('DTA_MENSAL',6 ,' ',tpRight, ftpString  );
    AddField('NUM_IDENT',10 ,'0',tpLeft, ftpInteger );
    AddField('COD_EAN',14 ,'0',tpLeft, ftpString );
    AddField('DES_HORA',4 ,' ',tpRight, ftpString );
    AddField('COD_CLIENTE',10 ,'0',tpLeft, ftpString );
    AddField('COD_ENTIDADE',10 ,'0',tpLeft, ftpInteger );
    AddField('VAL_BASE_ICMS',12 ,'0',tpLeft, ftpFloat, '0.000' );
    AddField('DES_SITUACAO_TRIB',3 ,' ',tpRight, ftpString );
    AddField('VAL_ICMS',12 ,'0',tpLeft, ftpFloat, '0.000' );
    AddField('NUM_CUPOM_FISCAL',10 ,'0',tpLeft, ftpInteger );
    AddField('VAL_VENDA_PDV',12 ,'0',tpLeft, ftpFloat, '0.000' );
    AddField('COD_TRIBUTACAO',10 ,'0',tpLeft, ftpInteger );
    AddField('FLG_CUPOM_CANCELADO',1 ,' ',tpRight, ftpString );
    AddField('NUM_NCM',10 ,' ',tpRight, ftpString );
    AddField('COD_TAB_SPED',3 ,'0',tpLeft, ftpString );
    AddField('FLG_NAO_PIS_COFINS',1 ,' ',tpRight, ftpString );
    AddField('TIPO_NAO_PIS_COFINS',1 ,'0',tpLeft, ftpInteger );
    AddField('FLG_ONLINE',100 ,' ',tpRight, ftpString );
    AddField('FLG_OFERTA',1 ,' ',tpRight, ftpString );
    AddField('COD_ASSOCIADO',8 ,'0',tpLeft, ftpInteger );
  end;
 Result := Layout;
end;

end.
