inherited FrmSmBomDeCarne: TFrmSmBomDeCarne
  Caption = 'SM BOM DE CARNE'
  ClientHeight = 536
  ExplicitHeight = 575
  TextHeight = 13
  inherited ImgLogo: TImage
    Left = 624
    Height = 63
    Anchors = [akTop, akRight]
    ExplicitLeft = 624
    ExplicitHeight = 63
  end
  object lblLoja: TLabel [2]
    Left = 527
    Top = 35
    Width = 20
    Height = 13
    Caption = 'Loja'
  end
  inherited PctArquivos: TPageControl
    Top = 15
    ExplicitTop = 15
    inherited AbaParceiros: TTabSheet
      inherited CkbCondPagForn: TCheckBox
        Top = 28
        ExplicitTop = 28
      end
      inherited CkbDivisaoForn: TCheckBox
        Top = 56
        Enabled = False
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
      object Label11: TLabel [0]
        Left = 114
        Top = 162
        Width = 28
        Height = 13
        Caption = '<<---'
      end
      inherited CkbProdSimilar: TCheckBox
        Enabled = False
      end
      inherited CkbCodigoBarras: TCheckBox
        Top = 104
        ExplicitTop = 104
      end
      inherited CkbProdLoja: TCheckBox
        OnClick = CkbProdLojaClick
      end
      inherited CkbProdForn: TCheckBox
        Enabled = False
      end
      inherited CkbComposicao: TCheckBox
        Enabled = False
      end
      inherited CkbReceitas: TCheckBox
        Enabled = False
      end
      inherited CkbInfoNutricionais: TCheckBox
        Enabled = False
      end
      inherited CkbProdComprador: TCheckBox
        Top = 281
        Enabled = False
        ExplicitTop = 281
      end
      inherited CkbNcm: TCheckBox
        Top = 142
        ExplicitTop = 142
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
        Enabled = False
        ExplicitTop = 240
      end
      object btnGeraValorVenda: TButton
        Left = 159
        Top = 134
        Width = 97
        Height = 25
        Caption = 'Gera val. venda'
        Enabled = False
        TabOrder = 15
        OnClick = btnGeraValorVendaClick
      end
      object btnGeraCustoRep: TButton
        Left = 159
        Top = 165
        Width = 97
        Height = 25
        Caption = 'Gera custo rep.'
        Enabled = False
        TabOrder = 16
        OnClick = btnGeraCustoRepClick
      end
      object btnGerarEstoqueAtual: TButton
        Left = 159
        Top = 196
        Width = 97
        Height = 25
        Caption = 'Gera estoque'
        Enabled = False
        TabOrder = 17
        OnClick = btnGerarEstoqueAtualClick
      end
    end
    inherited AbaFiscal: TTabSheet
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
      inherited CkbFinanceiro: TCheckBox
        Top = 10
        Enabled = False
        ExplicitTop = 10
      end
      inherited CkbFinanceiroPagar: TCheckBox
        Top = 30
        ExplicitTop = 30
      end
      inherited CkbFinanceiroReceber: TCheckBox
        Top = 70
        ExplicitTop = 70
      end
      inherited CkbFinanceiroPagarEmAbertos: TCheckBox
        Top = 50
        ExplicitTop = 50
      end
      inherited CkbFinanceiroReceberEmAberto: TCheckBox
        Top = 90
        ExplicitTop = 90
      end
      inherited CkbFinanceiroReceberCartoes: TCheckBox
        Top = 110
        Enabled = False
        ExplicitTop = 110
      end
      inherited CkbFinanceiroReceberBoleto: TCheckBox
        Top = 130
        Enabled = False
        ExplicitTop = 130
      end
      inherited CkbFinanceiroReceberCheque: TCheckBox
        Top = 150
        Enabled = False
        ExplicitTop = 150
      end
    end
    inherited AbaOutros: TTabSheet
      inherited CkbMapaResumo: TCheckBox
        Enabled = False
      end
      inherited CkbAjuste: TCheckBox
        Enabled = False
      end
      inherited CkbPlContas: TCheckBox
        Enabled = False
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
        ExplicitWidth = 457
        ExplicitHeight = 71
        inherited Label5: TLabel
          Left = 1
          Top = 15
          Width = 44
          Caption = 'Instancia'
          ExplicitLeft = 1
          ExplicitTop = 15
          ExplicitWidth = 44
        end
        inherited Label4: TLabel
          Left = 217
          Top = 47
          ExplicitLeft = 217
          ExplicitTop = 47
        end
        inherited Label1: TLabel
          Left = 4
          Top = 48
          Width = 36
          Caption = 'Usuario'
          ExplicitLeft = 4
          ExplicitTop = 48
          ExplicitWidth = 36
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
          Top = 11
          Width = 377
          ExplicitTop = 11
          ExplicitWidth = 377
        end
        inherited edtSenhaOracle: TEdit
          Left = 255
          Top = 43
          Width = 176
          PasswordChar = '*'
          ExplicitLeft = 255
          ExplicitTop = 43
          ExplicitWidth = 176
        end
        inherited edtInst: TEdit
          Left = 54
          Top = 43
          Width = 139
          ExplicitLeft = 54
          ExplicitTop = 43
          ExplicitWidth = 139
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
        ExplicitWidth = 457
        ExplicitHeight = 71
      end
      inherited TabSqlServer: TTabSheet
        ExplicitWidth = 457
        ExplicitHeight = 71
      end
    end
  end
  inherited GbxData: TGroupBox
    Top = 15
    ExplicitTop = 15
  end
  object CbxLoja: TComboBox [11]
    Left = 553
    Top = 32
    Width = 65
    Height = 21
    TabOrder = 8
    Text = '1'
    Items.Strings = (
      '1')
  end
  object Memo1: TMemo [12]
    Left = 0
    Top = 414
    Width = 840
    Height = 122
    Align = alBottom
    TabOrder = 9
    ExplicitLeft = -6
    ExplicitTop = 409
  end
  inherited MainMenu1: TMainMenu
    Top = 56
  end
  object ADOSQLServer: TADOConnection
    LoginPrompt = False
    Provider = 'MSDASQL.1'
    Left = 576
    Top = 8
  end
  object QryPrincipal2: TADOQuery
    Connection = ADOSQLServer
    AfterOpen = QryPrincipal2AfterOpen
    Parameters = <>
    Left = 464
    Top = 8
  end
  object QryAux: TADOQuery
    Connection = ADOSQLServer
    Parameters = <>
    SQL.Strings = (
      '  CASE PROD.ALIQUOTA '
      ''
      '  WHEN 0 THEN'
      '    CASE COALESCE(PROD.ICMSREDUCAO, 0) '
      #9'  WHEN 61.11 THEN   '
      #9'    CASE PROD.CST WHEN 20 THEN 8 END'#9#9
      #9'  WHEN 0 THEN'
      #9'    CASE PROD.CST '
      #9'      WHEN 0 THEN 25'
      #9#9'  WHEN 50 THEN 21'
      #9#9'  WHEN 51 THEN 23'
      #9#9'  WHEN 90 THEN 22'
      #9#9'  WHEN 60 THEN 25'
      #9#9'  WHEN 40 THEN 1'
      #9#9'  WHEN 41 THEN 1'
      #9#9'  WHEN 20 THEN 4'
      #9'  END'
      '    END'
      '  '
      '  WHEN 4.5 THEN 38'
      ''
      '  WHEN 5.5 THEN '
      '    CASE PROD.CST '
      #9'  WHEN 0 THEN 42'
      #9'  WHEN 70 THEN 44'
      '    END'
      '  '
      '  WHEN 12 THEN'
      '    CASE PROD.ICMSREDUCAO'
      #9'  WHEN 61.11 THEN'
      #9'    CASE WHEN PROD.CST = 20 THEN 40 END'
      #9'  WHEN 0 THEN 3'
      '    END'
      ''
      '  WHEN 13.30 THEN'
      '    CASE PROD.ICMSREDUCAO'
      #9'  WHEN 9.77 THEN'
      #9'    CASE WHEN PROD.CST = 20 THEN 41 END'
      #9'  WHEN 47.37 THEN '
      #9'    CASE WHEN PROD.CST = 20 THEN 43 END'
      #9'  WHEN 0 THEN'
      #9'    CASE WHEN PROD.CST = 20 THEN 41 ELSE 45 END'
      '    END'
      '  '
      '  WHEN 13.33 THEN'
      '    CASE PROD.ICMSREDUCAO'
      #9'  WHEN 9.77 THEN'
      #9'    CASE WHEN PROD.CST = 20 THEN 41 END'
      #9'END'
      ''
      '  WHEN 18 THEN'
      '    CASE PROD.ICMSREDUCAO'
      #9'  WHEN 61.11 THEN'
      #9'    CASE WHEN PROD.CST = 20 THEN 8 END'
      #9'  WHEN 0 THEN'
      #9'    CASE PROD.CST'
      #9#9'  WHEN 20 THEN 4'
      #9#9'  WHEN 0 THEN 4'
      #9#9'  WHEN 10 THEN 13'
      #9#9'  WHEN 70 THEN 13'
      #9#9'END'
      '    END'#9' '
      ''
      #9'WHEN 25 THEN 5'
      '  '
      '  END')
    Left = 568
    Top = 80
  end
  object QryAuxNF: TADOQuery
    Connection = ADOSQLServer
    Parameters = <>
    SQL.Strings = (
      '  CASE NFITEM.ICMS_ALIQUOTA'
      '    '
      #9'WHEN 0 THEN'
      #9'  CASE NFITEM.ICMS_REDUCAO'
      #9'    WHEN 61.11 THEN'#9#9'  '
      #9#9'  CASE WHEN NFITEM.ICMS_CST = 20 THEN 8 END'
      #9#9'WHEN 9.77 THEN'
      #9#9'  CASE WHEN NFITEM.ICMS_CST = 20 THEN 46 END'
      #9#9'WHEN 47.37 THEN'
      #9#9'  CASE WHEN NFITEM.ICMS_CST = 20 THEN 43 END'
      #9#9'WHEN 0 THEN'
      #9#9'  CASE NFITEM.ICMS_CST'
      #9#9'    WHEN 0 THEN 25'
      #9#9#9'WHEN 10 THEN 25'
      #9#9#9'WHEN 41 THEN 23'
      #9#9#9'WHEN 90 THEN 22'
      #9#9#9'WHEN 70 THEN 25'
      #9#9#9'WHEN 60 THEN 25'
      #9#9#9'WHEN 40 THEN 1'
      #9#9#9'WHEN 51 THEN 20'
      #9#9'  END'
      #9'  END'
      #9'  '
      '    WHEN 4 THEN'
      #9'  CASE NFITEM.ICMS_REDUCAO '
      #9'    WHEN 0 THEN'
      #9#9'  CASE NFITEM.ICMS_CST'
      #9#9'    WHEN 0 THEN 28'
      #9#9#9'WHEN 10 THEN 27'
      #9#9'  END'
      #9#9'END'
      '    '
      #9'WHEN 5.5 THEN'
      #9'  CASE NFITEM.ICMS_REDUCAO'
      #9'    WHEN 0 THEN'
      #9#9'  CASE NFITEM.ICMS_CST'
      #9#9'    WHEN 41 THEN 23'
      #9#9#9'WHEN 20 THEN 42'
      #9#9#9'WHEN 0 THEN 42'
      #9#9#9'WHEN 40 THEN 1'
      #9#9'  END'
      '      END'
      ''
      '    WHEN 7 THEN'
      #9'  CASE NFITEM.ICMS_REDUCAO'
      #9'    WHEN 0 THEN'
      #9#9'  CASE NFITEM.ICMS_CST '
      #9#9'    WHEN 0 THEN 2'
      #9#9#9'WHEN 90 THEN 22'
      #9#9'  END'
      #9'  END'
      ''
      #9'WHEN 11.2 THEN'
      #9'  CASE NFITEM.ICMS_REDUCAO'
      #9'    WHEN 0 THEN'
      #9#9'  CASE NFITEM.ICMS_CST'
      #9#9'    WHEN 41 THEN 23'
      #9#9#9'WHEN 40 THEN 1'
      #9#9#9'WHEN 0 THEN 47'
      #9#9'  END'
      #9'  END'
      ''
      #9'WHEN 12 THEN'
      #9'  CASE NFITEM.ICMS_REDUCAO'
      #9'    WHEN 61.11 THEN'
      #9#9'  CASE WHEN NFITEM.ICMS_CST = 70 THEN 15 END'
      #9#9'WHEN 41.67 THEN'
      #9#9'  CASE WHEN NFITEM.ICMS_CST = 10 THEN 48 END'
      #9#9'WHEN 0 THEN'
      #9#9'  CASE NFITEM.ICMS_CST'
      #9#9'    WHEN 10 THEN 42'
      #9#9#9'WHEN 20 THEN 3'
      #9#9#9'WHEN 40 THEN 1 '
      #9#9#9'WHEN 0 THEN 3'
      #9#9'  END'
      #9'  END'
      ''
      #9'WHEN 13.3 THEN'
      #9'  CASE NFITEM.ICMS_REDUCAO'
      #9'    WHEN 47.37 THEN'
      #9#9'  CASE NFITEM.ICMS_CST '
      #9#9'    WHEN 20 THEN 43'
      #9#9#9'WHEN 70 THEN 43'
      #9#9'  END'
      #9#9'WHEN 9.77 THEN'
      #9#9'  CASE WHEN NFITEM.ICMS_CST = 20 THEN 41 END'
      '        WHEN 0 THEN'
      #9#9'  CASE NFITEM.ICMS_CST'
      #9#9'    WHEN 10 THEN 49'
      #9#9#9'WHEN 70 THEN 49'
      #9#9#9'WHEN 0 THEN 0'
      #9#9#9'WHEN 20 THEN 50'
      #9#9'  END'
      #9'  END'
      ''
      #9'WHEN 13.33 THEN'
      #9'  CASE NFITEM.ICMS_REDUCAO'
      #9'    WHEN 9.77 THEN'
      #9#9'  CASE WHEN NFITEM.ICMS_CST = 20 THEN 41 END'
      #9'  END'
      ''
      #9'WHEN 18 THEN'
      #9'  CASE NFITEM.ICMS_REDUCAO'
      #9'    WHEN 61.11 THEN'
      #9#9'  CASE NFITEM.ICMS_CST'
      #9#9'    WHEN 70 THEN 17'
      #9#9#9'WHEN 20 THEN 8'
      #9#9'  END'
      #9#9'WHEN 33.33 THEN'
      #9#9'  CASE NFITEM.ICMS_CST'
      #9#9'    WHEN 70 THEN 16'
      '          END'
      #9#9'WHEN 26.11 THEN'
      #9#9'  CASE NFITEM.ICMS_CST '
      #9#9'    WHEN 70 THEN 17'
      #9#9'  END'
      #9#9'WHEN 0 THEN'
      #9#9'  CASE NFITEM.ICMS_CST'
      #9#9'    WHEN 70 THEN 13'
      #9#9#9'WHEN 90 THEN 22'
      #9#9#9'WHEN 10 THEN 13'
      #9#9#9'WHEN 20 THEN 4'
      #9#9#9'WHEN 0 THEN 4'
      #9#9#9'WHEN 51 THEN 20'
      #9#9'  END'
      #9'  END'
      ''
      #9'WHEN 20 THEN '
      #9'  CASE NFITEM.ICMS_REDUCAO'
      #9'    WHEN 0 THEN '
      #9#9'  CASE WHEN NFITEM.ICMS_CST = 10 THEN 29 END'
      #9'  END'
      ''
      #9'WHEN 25 THEN'
      #9'  CASE NFITEM.ICMS_REDUCAO'
      #9'    WHEN 0 THEN'
      #9#9'  CASE NFITEM.ICMS_CST'
      #9#9'    WHEN 0 THEN 5'
      #9#9#9'WHEN 70 THEN 14'
      #9#9#9'WHEN 10 THEN 14'
      #9#9#9'WHEN 20 THEN 9'
      #9#9'  END'
      #9'  END'
      #9
      #9'WHEN 30 THEN'
      #9'  CASE NFITEM.ICMS_REDUCAO'
      #9'    WHEN 0 THEN'
      #9#9'  CASE WHEN NFITEM.ICMS_CST = 10 THEN 31 END'
      #9'  END'#9'   '#9'       '
      ''
      ''
      '  END ')
    Left = 624
    Top = 80
  end
end
