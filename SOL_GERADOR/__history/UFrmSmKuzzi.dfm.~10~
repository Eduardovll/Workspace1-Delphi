inherited FrmSmKuzzi: TFrmSmKuzzi
  Caption = 'SUPERMERCADO KUZZI'
  PixelsPerInch = 96
  TextHeight = 13
  inherited ImgLogo: TImage
    Left = 624
    Height = 63
    Anchors = [akTop, akRight]
    ExplicitLeft = 624
    ExplicitHeight = 63
  end
  inherited PctArquivos: TPageControl
    Top = 15
    ActivePage = AbaParceiros
    ExplicitTop = 15
    inherited AbaParceiros: TTabSheet
      ExplicitLeft = 4
      ExplicitTop = 24
      ExplicitWidth = 257
      ExplicitHeight = 297
      inherited CkbCondPagForn: TCheckBox
        Top = 28
        ExplicitTop = 28
      end
      inherited CkbDivisaoForn: TCheckBox
        Top = 56
        ParentDoubleBuffered = False
        ExplicitTop = 56
      end
      inherited CkbTransportadora: TCheckBox
        Top = 131
        Enabled = False
        ExplicitTop = 131
      end
      inherited CkbStatusPdv: TCheckBox
        Top = 262
        Enabled = False
        Visible = False
        ExplicitTop = 262
      end
      inherited CkbCliente: TCheckBox
        Top = 82
        ExplicitTop = 82
      end
      inherited CkbCondPagCli: TCheckBox
        Top = 105
        ExplicitTop = 105
      end
      inherited CkbEnderecoCliente: TCheckBox
        Top = 239
        Enabled = False
        Visible = False
        ExplicitTop = 239
      end
    end
    inherited AbaProdutos: TTabSheet
      ExplicitLeft = 4
      ExplicitTop = 24
      ExplicitWidth = 257
      ExplicitHeight = 297
      inherited CkbProdComprador: TCheckBox
        Top = 281
        Enabled = False
        ExplicitTop = 281
      end
      inherited CkbDecomposicao: TCheckBox
        Enabled = False
      end
      inherited CkbProdLocalizacao: TCheckBox
        Enabled = False
      end
      inherited CkbProdProducao: TCheckBox
        Top = 240
        Caption = 'Produ'#231#227'o'
        ExplicitTop = 240
      end
    end
    inherited AbaFiscal: TTabSheet
      ExplicitLeft = 4
      ExplicitTop = 24
      ExplicitWidth = 257
      ExplicitHeight = 297
      inherited CkbOutrasNFs: TCheckBox
        Enabled = False
      end
      inherited CkbNFTransf: TCheckBox
        Enabled = False
      end
      inherited CkbNFClientes: TCheckBox
        Enabled = False
      end
      inherited CkbTributacao: TCheckBox
        Enabled = False
      end
      inherited CkbNf: TCheckBox
        Enabled = False
      end
    end
    inherited Financeiro: TTabSheet
      ExplicitLeft = 4
      ExplicitTop = 24
      ExplicitWidth = 257
      ExplicitHeight = 297
      inherited CkbFinanceiro: TCheckBox
        Enabled = False
      end
      inherited CkbFinanceiroReceberBoleto: TCheckBox
        Enabled = False
      end
      inherited CkbFinanceiroReceberCheque: TCheckBox
        Enabled = False
      end
    end
    inherited AbaOutros: TTabSheet
      ExplicitLeft = 4
      ExplicitTop = 24
      ExplicitWidth = 257
      ExplicitHeight = 297
      inherited CkbMapaResumo: TCheckBox
        Enabled = False
      end
      inherited CkbAjuste: TCheckBox
        Enabled = False
      end
      inherited CkbPlContas: TCheckBox
        Enabled = False
      end
      object btnGeraCest: TButton
        Left = 3
        Top = 248
        Width = 98
        Height = 41
        Caption = 'Cest'
        TabOrder = 4
        OnClick = btnGeraCestClick
      end
      object BtnAmarrarCest: TButton
        Left = 136
        Top = 248
        Width = 89
        Height = 41
        Caption = 'Amarrar Cest'
        TabOrder = 5
        OnClick = BtnAmarrarCestClick
      end
    end
  end
  inherited GroupBox: TGroupBox
    inherited GrpCamBanco: TGroupBox
      Visible = False
      inherited EdtCamBanco: TEdit
        Visible = False
      end
    end
    inherited GrpCamBancoSoliduss: TGroupBox
      Visible = False
    end
    inherited PCBancoDados: TPageControl
      Left = 8
      Top = 176
      Width = 465
      Height = 81
      ExplicitLeft = 8
      ExplicitTop = 176
      ExplicitWidth = 465
      ExplicitHeight = 81
      inherited TabOracle: TTabSheet
        ExplicitLeft = 4
        ExplicitTop = 6
        ExplicitWidth = 457
        ExplicitHeight = 71
        inherited Label5: TLabel
          Left = 9
          Top = 47
          Width = 36
          Caption = 'Usuario'
          ExplicitLeft = 9
          ExplicitTop = 47
          ExplicitWidth = 36
        end
        inherited Label4: TLabel
          Left = 209
          Top = 47
          ExplicitLeft = 209
          ExplicitTop = 47
        end
        inherited Label1: TLabel
          Left = 199
          Width = 37
          Caption = 'Schema'
          ExplicitLeft = 199
          ExplicitWidth = 37
        end
        inherited Label6: TLabel
          Left = 36
          Width = 10
          Caption = 'IP'
          Visible = False
          ExplicitLeft = 36
          ExplicitWidth = 10
        end
        inherited edtSchema: TEdit
          Left = 48
          Top = 43
          ExplicitLeft = 48
          ExplicitTop = 43
        end
        inherited edtSenhaOracle: TEdit
          Left = 249
          Top = 45
          Width = 190
          PasswordChar = '*'
          ExplicitLeft = 249
          ExplicitTop = 45
          ExplicitWidth = 190
        end
        inherited edtInst: TEdit
          Left = 249
          Width = 189
          ExplicitLeft = 249
          ExplicitWidth = 189
        end
        inherited edtIpOra: TEdit
          Left = 48
          Top = 3
          Visible = False
          ExplicitLeft = 48
          ExplicitTop = 3
        end
      end
      inherited TabMySql: TTabSheet
        ExplicitLeft = 4
        ExplicitTop = 6
        ExplicitWidth = 457
        ExplicitHeight = 71
      end
      inherited TabSqlServer: TTabSheet
        ExplicitLeft = 4
        ExplicitTop = 6
        ExplicitWidth = 457
        ExplicitHeight = 71
      end
    end
  end
  inherited GbxData: TGroupBox
    Left = 288
    Top = 34
    ExplicitLeft = 288
    ExplicitTop = 34
  end
  object ADOOracle: TADOConnection
    LoginPrompt = False
    Provider = 'MSDAORA.1'
    Left = 560
    Top = 8
  end
  object QryPrincipal2: TADOQuery
    Connection = ADOOracle
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      'SELECT'
      '    CLIENTES.ID AS COD_CLIENTE,'
      '    CLIENTES.DESCRITIVO AS DES_CLIENTE,'
      '    CLIENTES.CNPJ_CPF AS NUM_CGC,'
      '    CLIENTES.INSCRICAO_RG AS NUM_INSC_EST,'
      
        '    -- CLIENTES.LOGRADOURO || '#39' '#39' || CLIENTES.ENDERECO AS DES_EN' +
        'DERECO,'
      '    CLIENTES.BAIRRO AS DES_BAIRRO,'
      '    CLIENTES.CIDADE AS DES_CIDADE,'
      '    CLIENTES.ESTADO AS DES_SIGLA,'
      '    CLIENTES.CEP AS NUM_CEP,'
      '    CLIENTES.TELEFONE1 AS NUM_FONE,'
      '    CLIENTES.FAX AS NUM_FAX,'
      '    CLIENTES.FANTASIA AS DES_CONTATO,'
      ''
      '    CASE'
      '        WHEN CLIENTES.SEXO = 1 THEN 1'
      '        ELSE 0 '
      '    END AS FLG_SEXO, -- SEXO'
      ''
      '    0 AS VAL_LIMITE_CRETID, '
      '    0 AS VAL_LIMITE_CONV, --VERIFICAR SITUACAO, LIMITE'
      '    0 AS VAL_DEBITO,'
      '    CLIENTES.SALARIO AS VAL_RENDA,'
      '    0 AS COD_CONVENIO,'
      '    0 AS COD_STATUS_PDV,'
      ''
      '    CASE CLIENTES.PESSOA'
      '        WHEN '#39'F'#39' THEN '#39'N'#39
      '        ELSE '#39'S'#39
      '    END AS FLG_EMPRESA,'
      ''
      '    '#39'N'#39' AS FLG_CONVENIO, --'
      '    '#39'N'#39' AS MICRO_EMPRESA,'
      '    CLIENTES.DATAHORA_CADASTRO AS DTA_CADASTRO,'
      '    CLIENTES.NUMERO AS NUM_ENDERECO,'
      '    '#39#39' AS NUM_RG, --OBSERVACAO'
      ''
      '    CASE CLIENTES.ESTADO_CIVIL'
      '        WHEN 0 THEN 0'
      '        WHEN 1 THEN 1'
      '        WHEN 2 THEN 3'
      '        WHEN 3 THEN 2'
      '        WHEN 4 THEN 4'
      '        ELSE 0'
      '    END AS FLG_EST_CIVIL, --ESTADO_CIVIL'
      ''
      '    '#39#39' AS NUM_CELULAR, -- TELEFONE2'
      '    CLIENTES.DATAHORA_ALTERACAO AS DTA_ALTERACAO,'
      '    CLIENTES.OBSERVACAO AS DES_OBSERVACAO,'
      '    CLIENTES.COMPLEMENTO AS DES_COMPLEMENTO,'
      '    CLIENTES.EMAIL AS DES_EMAIL,'
      '    CLIENTES.FANTASIA AS DES_FANTASIA,'
      '    CLIENTES.DATA_NASCIMENTO AS DTA_NASCIMENTO,'
      '    CLIENTES.PAI AS DES_PAI,'
      '    CLIENTES.MAE AS DES_MAE,'
      '    CLIENTES.CONJUGUE AS DES_CONJUGE,'
      '    CLIENTES.CPF_CONJUGE AS NUM_CPF_CONJUGE,'
      '    0 AS VAL_DEB_CONV,'
      '    '#39'N'#39' AS INATIVO,'
      '    0 AS DES_MATRICULA,'
      '    '#39'N'#39' AS NUM_CGC_ASSOCIADO,'
      '    '#39'N'#39' AS FLG_PROD_RURAL,'
      '    0 AS COD_STATUS_PDV_CONV,'
      '    '#39'S'#39' AS FLG_ENVIA_CODIGO,'
      '    '#39#39' AS DTA_NASC_CONJUGE,'
      '    0 AS COD_CLASSIF'
      'FROM'
      '    CLIENTES'
      'ORDER BY'
      '    CLIENTES.ID')
    Left = 480
  end
end
