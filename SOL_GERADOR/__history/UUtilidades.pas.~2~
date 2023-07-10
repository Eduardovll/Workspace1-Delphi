unit UUtilidades;
{Esta unit contém objetos e funções genéricas para uso em qualquer aplicação}

interface

uses
  SysUtils, Classes, Windows, Dialogs, Registry, SqlExpr, DB, DBClient, UClasses;

type
  { Permite cronometrar seções de código, entre outras coisas }
  TCronometro = class(TObject)
  private
    ContagemInicial :Int64;
    ContagemFinal :Int64;
  public
    procedure Iniciar;
    procedure Parar;
    function Contagem :Int64;
    procedure Exibir;
  end;

  { Permite manipular uma coleção de Objetos }
  TList = class(TObject)
  protected
    Items : Array of TObject;
  public
    procedure Add(Item: TObject); //adiciona um item ao final da lista
    procedure Push(Item: TObject); //o mesmo que Add
    function Pop:TObject; //remove um item do fim da lista e retorna o mesmo
    procedure UnShift(Item: TObject); //adiciona um item ao começo da lista
    function Shift:TObject; //remove um item do começo da lista e retorna o mesmo
    function Item(ItemIndex: Integer):TObject; //retorna um item especifico
    function IsEmpty:Boolean; //retorna um booleano informando se a lista está vazia
    function Length:Integer; //retorna o tamanho da lista
  end;

    function GerarPLU(CodProd: String): String;
    function GeraDigitoEAN(Cod: String; Num: word): String;
    function PLUValido(CodProd: String): Boolean; overload;
    function PLUValido(CodProd: Int64): Boolean; overload;
    function CodBarrasValido(CodBarras: String): Boolean; overload;
    function CodBarrasValido(CodBarras: Int64): Boolean; overload;
    function FlagValido(Flag: String): Boolean;
    function TiraLetrasIE(Cadeia: string): String;
    function ValidaCGC(xCNPJ: string): Boolean; deprecated;
    function DigitoCPF(Numero: String): String; deprecated;
    function ValidaCPF(Numero :String) :Boolean; deprecated;
    function CPFEValido(Numero :String) :Boolean;
    function CNPJEValido(Numero : String): Boolean;
    function GerarDigitosCPF(Numero: String): String;
    function GerarDigitosCNPJ(Numero : String): String;
    function GerarCNPJFicticio(Numero: String): String;
    function GerarCPFFicticio(Numero: String): String;
    function ValidarFlag(Flag, NomeFlag: String): String;
    function ZeradaOuVazia(Valor: String): Boolean;
    function ENumero(Caracter: Char): Boolean; overload;
    function ENumero(Texto: String): Boolean; overload;
    function iif(aTest: Boolean; TrueValue, FalseValue: Variant): Variant;
    function DateToFStr(Date: TDateTime; const Format: String):String;
    function ObtValorRegistro(Raiz :Integer; Caminho, Nome: string): string;
    function ObterVariavelDeSistema(const Nome: string): string;

    //Manipulação de Strings
    function StrPad(Text: String; Size: Integer; Filler: String; Side: Char): String; overload;
    function StrPad(Text: Int64; Size: Integer; Filler: String; Side: Char): String; overload;
    function StrPad(Text: Real; Size: Integer; Filler: String; Side: Char): String; overload;
    function StrPad(Text: String; Size: Integer; Filler: String; Side: TSides): String; overload;
    function StrPad(Text: Int64; Size: Integer; Filler: String; Side: TSides): String; overload;
    function StrPad(Text: Real; Size: Integer; Filler: String; Side: TSides): String; overload;
    function StrLBReplace(Text: String): String;
    function StrLBReturn(Text: String): String;
    function StrLBDelete(Text: String): String;
    function StrSubstLtsAct(Text: String): String;
    function StrRemPont(Text: String): String;
    function StrRetNums(Text: String): String;
    function StrSplit(Text, Separator: String): TStringList;
    function StrToIntErr(Numero, Nome: String): Int64;
    function StrToDateErr(Data, Nome: String): TDateTime;
    function StrDelChar(const Text: String; Character: Char): String; overload;
    function StrDelChar(const Text: String; CharList: String): String; overload;
    function StrReplace(Text, Target, Replace: String): String;
    function StrDelMultSpaces(Text :String): String;
    function TiraZerosEsquerda(const Value: string): string;
    function SplitString(const Str: string; Delimiter : Char) : TStringList;

//Debug                              \
procedure DumpQuery(Query :TSQLQuery; FileName :String);
procedure DumpDataSet(DataSet :TDataSet; FileName :String);

function MemoryStreamToString(M: TMemoryStream): AnsiString;

implementation

{ TList }



//adiciona um item ao final da lista
procedure TList.Add(Item: TObject);
begin
  SetLength(Self.Items, System.Length(Self.Items) + 1);
  Self.Items[System.Length(Self.Items) - 1] := Item;
end;

//retorna um booleano informando se a lista está vazia
function TList.IsEmpty: Boolean;
begin
  if System.Length(Self.Items) > 0 then
    Result := false
  else
    Result := true;
end;

//retorna um item especifico da lista, mas não remove ele
function TList.Item(ItemIndex: Integer): TObject;
begin
  Result := Self.Items[ItemIndex];
end;

//remove um item do fim da lista e retorna o mesmo
function TList.Pop: TObject;
begin
  Result := Self.Items[System.Length(Self.Items) - 1];
  SetLength(Self.Items, System.Length(Self.Items) - 1);
end;

//o mesmo que Add
procedure TList.Push(Item: TObject);
begin
  Self.Add(Item);
end;

//remove um item do começo da lista e retorna o mesmo
function TList.Shift: TObject;
var
  ItemIndex :Integer;
begin
  Result := Self.Items[0];
  for ItemIndex := 1 to System.Length(Self.Items) - 1 do
    Self.Items[ItemIndex - 1] := Self.Items[ItemIndex];
  SetLength(Self.Items, System.Length(Self.Items) - 1);
end;

//retorna o tamanho da lista
function TList.Length: Integer;
begin
  Result := System.Length(Self.Items);
end;

//adiciona um item ao começo da lista
procedure TList.UnShift(Item: TObject);
var
  ItemIndex :Integer;
begin
  SetLength(Self.Items, System.Length(Self.Items) + 1);
  for ItemIndex := System.Length(Self.Items) - 1 downto 1 do
    Self.Items[ItemIndex] := Self.Items[ItemIndex - 1];
  Self.Items[0] := Item;
end;

function GerarPLU(CodProd: String): String;
begin
  Result := CodProd + GeraDigitoEAN(CodProd, 13);
end;

function GeraDigitoEAN(Cod: String; Num: word): String;
var
  Tot1, Tot2, Tot3: LongInt;
  I: Integer;
  CodRef: String;
  Digito: Integer;
begin
  Result := '';
  Tot1 := 0;
  Tot2 := 0;
  CodRef := StrPad(Cod, Num - 1, '0', 'L');
  for I := 1 to (Num - 1) do
  begin
    if (I mod 2) = 0 then
      Tot1 := Tot1 + StrToInt(CodRef[i])
    else
      Tot2 := Tot2 + StrToInt(CodRef[i]);
  end;
  Tot1 := Trunc(Tot1 * 3);
  Tot3 := Tot1 + Tot2;
  Digito := Trunc(10 * (int(Tot3 / 10) + 1) - Tot3);
  if Digito = 10 then
    Digito := 0;
  Result := IntToStr(Digito);
end;

function FlagValido(Flag: String): Boolean;
begin
  Result := False;
  Flag := UpperCase(Flag);
  if (Flag = 'S') or (Flag = 'N') then
    Result := True;
end;

{Verifica se o CodProd é um PLU válido, retorna false caso contrario}
function PLUValido(CodProd: String): Boolean; overload;
var
  Digito, CodSemDig: String;
begin
  Result := False;
  CodProd := StrPad(Trim(CodProd), 8, '0', 'L');
  CodSemDig := Copy(CodProd, 1, 7);
  Digito := GeraDigitoEAN(CodSemDig, 13);
  if Copy(CodProd, 8, 1) <> Digito then
    Exit;

  Result := True;
end;

function PLUValido(CodProd: Int64): Boolean; overload;
begin
  Result := PLUValido(IntToStr(CodProd));
end;

{Verifica se o CodBarras é valido, retorna false caso contrario}
function CodBarrasValido(CodBarras: String): Boolean; overload;
var
  CodSemDig, Digito: String;
begin
  CodBarras := Trim(CodBarras);
  Digito := '';
  Result := False;
  if (Length(CodBarras) > 13) then
    Exit;

  CodBarras := StrPad(CodBarras, 13, '0', 'L');
  CodSemDig := Copy(CodBarras, 1, 12);
  Digito := GeraDigitoEAN(CodSemDig, 13);
  if Copy(CodBarras, 13, 1) <> Digito then
    Exit;

  Result := True;
end;

function CodBarrasValido(CodBarras: Int64): Boolean; overload;
begin
  Result := CodBarrasValido(IntToStr(CodBarras));
end;

function TiraLetrasIE(Cadeia: string): String;
var
  I: Integer;
begin
  Result := Cadeia;
  for I := 1 to Length(Result) do
    if (Result[i] in ['A'..'O', 'a'..'o', 'Q'..'Z', 'q'..'z', '/', '.', '-','"']) then
      Delete(Result, I, 1);
end;

function ValidaCGC(xCNPJ: string): Boolean;
begin
  Result := CNPJEValido(xCNPJ);
end;

function DigitoCPF(Numero: String): String; deprecated;
begin
  Result := GerarDigitosCPF(Numero);
end;

function ValidaCPF(Numero :String) :Boolean; deprecated;
begin
  Result := CPFEValido(Numero);
end;

{Retorna true caso Numero seja um CPF válido}
function CPFEValido(Numero :String) :Boolean;
var
  NumeroInt : Int64;
begin
  Result := False;
  NumeroInt := StrToInt64Def(Numero, 0);
  if ((NumeroInt = 0) or (NumeroInt > 100000000000)) then
    Exit;
  if
  (
    (Numero = '00000000000') or
    (Numero = '11111111111') or
    (Numero = '22222222222') or
    (Numero = '33333333333') or
    (Numero = '44444444444') or
    (Numero = '55555555555') or
    (Numero = '66666666666') or
    (Numero = '77777777777') or
    (Numero = '88888888888') or
    (Numero = '99999999999')
  ) then
    Exit;

  Numero := StrPad(Numero, 11, '0', 'L');
  if Copy(Numero, 10, 2) = GerarDigitosCPF(Copy(Numero, 1, 9)) then
    Result := True;
end;

{Retorna true caso Numero seja um CNPJ válido}
function CNPJEValido(Numero : String): Boolean;
var
  NumeroInt : Int64;
begin
  if Numero = '48740351000408' then
    Trim(' ');
  Result := False;
  NumeroInt := StrToInt64Def(Numero, 0);
  if ((NumeroInt = 0) or (NumeroInt > 100000000000000)) then
    Exit;
  if
  (
    (Numero = '00000000000000') or
    (Numero = '11111111111111') or
    (Numero = '22222222222222') or
    (Numero = '33333333333333') or
    (Numero = '44444444444444') or
    (Numero = '55555555555555') or
    (Numero = '66666666666666') or
    (Numero = '77777777777777') or
    (Numero = '88888888888888') or
    (Numero = '99999999999999')
  ) then
    Exit;

  Numero := StrPad(Numero, 14, '0', 'L');
  if Copy(Numero, 13, 2) = GerarDigitosCNPJ(Copy(Numero, 1, 12)) then
    Result := True;
end;

{Calcula e retorna os digitos do CPF para Numero}
function GerarDigitosCPF(Numero: String): String;
var
  i, code : Integer;
  d2 : array[1..12] of Integer;
  DF4, DF5, DF6, RESTO1, Pridig, Segdig : Integer;
  Pridig2, Segdig2 : string;
  Texto : string;
begin
  Numero := StrPad(Numero, 9, '0', 'L');
  Texto :='';
  Texto := Copy(Numero, 1, 3);
  Texto := Texto + Copy(Numero, 4, 3);
  Texto := Texto + Copy(Numero, 7, 3);
  Texto := Texto + Copy(Numero, 10, 2);
  for i := 1 to 9 do
    Val(texto[i],D2[i],Code);
  DF4:=0;
  for i := 1 to 9 do
    DF4 := DF4 + (D2[i] * (11-I));
  DF5 := DF4 div 11;
  DF6 := DF5 * 11;
  Resto1 := DF4 - DF6;
  if (Resto1=0) or (Resto1=1) then
    Pridig:=0
  else
    Pridig:=11 - Resto1;
  for i := 1 to 9 do
    Val(Texto[i],D2[i],Code);
  DF4 := 11 * D2[1] +
         10 * D2[2] +
          9 * D2[3] +
          8 * D2[4] +
          7 * D2[5] +
          6 * D2[6] +
          5 * D2[7] +
          4 * D2[8] +
          3 * D2[9] +
          2 * Pridig;
  DF5 := DF4 div 11;
  DF6 := DF5 * 11;
  Resto1 := DF4 - DF6;
  if (Resto1=0) or (Resto1=1) then
    Segdig:=0
  else
    Segdig:=11 - Resto1;
  Str(Pridig, Pridig2);
  Str(Segdig, Segdig2);
  Result := Pridig2 + SegDig2;
end;

function GerarCPFFicticio(Numero: String): String;
begin
  Result := StrPad(Numero, 9, '0', 'L') + GerarDigitosCPF(Numero);
end;

function GerarCNPJFicticio(Numero: String): String;
begin
  Numero := Numero + '0001';
  Result := StrPad(Numero, 12, '0', 'L') + GerarDigitosCNPJ(Numero);
end;

{Calcula e retorna os digitos do CNPJ para Numero}
function GerarDigitosCNPJ(Numero : String): String;
var
  Dig1, Dig2, Cont: Integer;
begin
  StrToIntErr(Numero, 'número para gerar CNPJ');
  Numero := StrPad(Numero, 12, '0', 'L');

  Dig1 := 0;
  for Cont := 1 to 12 do
  begin
    if Cont < 5 then
      Dig1 := Dig1 + ((6 - Cont) * StrToInt(Numero[Cont]))
    else
      Dig1 := Dig1 + ((14 - Cont) * StrToInt(Numero[Cont]));
  end; //for
  Dig1 := 11 - (Dig1 mod 11);
  if Dig1 > 9 then
    Dig1 := 0;

  Dig2 := 0;
  for Cont := 1 to 12 do
  begin
    if Cont < 6 then
      Dig2 := Dig2 + ((7 - Cont) * StrToInt(Numero[Cont]))
    else
      Dig2 := Dig2 + ((15 - Cont) * StrToInt(Numero[Cont]));
  end; //for
  Dig2 := Dig2 + (2 * Dig1);
  Dig2 := 11 - (Dig2 mod 11);
  if Dig2 > 9 then
    Dig2 := 0;

  Result := IntToStr(Dig1) + IntToStr(Dig2);
end;

{cria uma exceção caso Flag não seja 'S' ou 'N'}
function ValidarFlag(Flag, NomeFlag: String): String;
begin
  Flag := Trim(Flag);
  if (Flag = 'S') or (Flag = 'N') then
    Result := Flag
  else
    Raise Exception.Create('"' + Flag + '" não é um valor válido para o flag "'
     + NomeFlag + '".');
end;

{Retorna true caso Valor seja uma string composta por zeros e/ou espaços}
function ZeradaOuVazia(Valor: String): Boolean;
begin
  Result := True;
  if (Trim(Valor) <> '') and (StrToIntDef(Trim(Valor), 1) <> 0) then
    Result := False;
end;

{Retorna positivo caso Texto seja um número}
function ENumero(Caracter: Char): Boolean; overload;
begin
  if Caracter in ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9'] then
    Result := True
  else
    Result := False;
end;

function ENumero(Texto: String): Boolean; overload;
var
  Cont :Integer;
begin
  Result := False;
  for Cont := 1 to Length(Texto) do
  begin
    if not (StrToInt(Texto[Cont]) in [0..9]) then
      Exit;
  end;
  Result := True;
end;

function iif(aTest: Boolean; TrueValue, FalseValue: Variant): Variant;
begin
  if aTest then
    Result := TrueValue
  else
    Result := FalseValue;
end;

{Transforma Date em uma string de acordo com as regras de Format}
function DateToFStr(Date: TDateTime; const Format: String):String;
begin
  DateTimeToString(Result, Format, Date);
end;

function ObterVariavelDeSistema(const Nome: string): string;
var
  BufSize: Integer;  // buffer size required for value
begin
  // Get required buffer size (inc. terminal #0)
  BufSize := GetEnvironmentVariable(PChar(Nome), nil, 0);
  if BufSize > 0 then
  begin
    // Read env var value into result string
    SetLength(Result, BufSize - 1);
    GetEnvironmentVariable(PChar(Nome), PChar(Result), BufSize);
  end
  else
    // No such environment variable
    Result := '';
end;

{Recupera o valor do registro no Caminho da Raiz}
function ObtValorRegistro(Raiz :Integer; Caminho, Nome: string): string;
var
  Registro: TRegistry;
begin
  Registro := TRegistry.Create(KEY_READ);
  case Raiz of
    1: Registro.RootKey := HKEY_LOCAL_MACHINE;
    2: Registro.RootKey := HKEY_CURRENT_USER;
    3: Registro.RootKey := HKEY_CLASSES_ROOT;
    else
      Raise Exception.Create('Raiz para localização de registro inválida');
  end;
  try
    Registro.OpenKey(Caminho, False);
    Result := Registro.ReadString(Nome);
  finally
    Registro.Free;
  end;
end;

{Preenche Text com Sample do lado Side até ficar do tamanho de text depois
caso Text seja maior que size deleta o excesso do lado Side, Side pode ser
L ou E para esquerda e R ou D para direita}
function StrPad(Text: String; Size: Integer; Filler: String; Side: Char): String; overload;
begin
  Result := Trim(Text);
  while Length(Result) < Size do
  begin
    if (Side = 'L') or (Side = 'E') then
      Result := Filler + Result
    else
      Result := Result + Filler;
  end;

  if Length(Result) > Size then
  begin
    if (Side = 'L') or (Side = 'E') then
      Delete(Result, 1, Length(Result) - Size)
    else
      Delete(Result, Size + 1, Length(Result) - Size);
  end;
end;

function StrPad(Text: Int64; Size: Integer; Filler: String; Side: Char): String; overload;
begin
  Result := StrPad(IntToStr(Text), Size, Filler, Side);
end;

function StrPad(Text: Real; Size: Integer; Filler: String; Side: Char): String; overload;
begin
  Result := StrPad(IntToStr(Round(Text)), Size, Filler, Side);
end;

function StrPad(Text: String; Size: Integer; Filler: String; Side: TSides): String; overload;
var C : Char;
begin
  if Side = tpLeft then
    C := 'L'
  else
    C := 'R';
  Result := StrPad(Text, Size, Filler, C);
end;

function StrPad(Text: Int64; Size: Integer; Filler: String; Side: TSides): String; overload;
var C : Char;
begin
  if Side = tpLeft then
    C := 'L'
  else
    C := 'R';
  Result := StrPad(IntToStr(Text), Size, Filler, C);
end;

function StrPad(Text: Real; Size: Integer; Filler: String; Side: TSides): String; overload;
var C : Char;
begin
  if Side = tpLeft then
    C := 'L'
  else
    C := 'R';
  Result := StrPad(IntToStr(Round(Text)), Size, Filler, C);
end;

{Suibstitui todas as quebras de linha do windows (CR+LF) de Text por \n}
function StrLBReplace(Text: String): String;
begin
  Result := Text;
  Result := StrReplace(Result, #13#10, '\n');
  Result := StrReplace(Result, #10#13, '\n');
  Result := StrReplace(Result, #13, '\n');
  Result := StrReplace(Result, #10, '\n');
end;

{Suibstitui todas \n de Text por quebras de linha}
function StrLBReturn(Text: String): String;
begin
  Result := StrReplace(Text, '\n', ' ' + #13#10);
end;

{Remove todas as quebras de linha de Text}
function StrLBDelete(Text: String): String;
var
  Position: Integer;
begin
  Position := 1;
  Result := Text;
  while Position < Length(Result) do
  begin
    if Result[Position] in [#10, #13] then
      Delete(Result, Position, 1) //remove quebra
    else
      Inc(Position); //vai para proximo caracter
  end;
end;

{Substitui letras acentuadas pelas letras não acentuadas equivalentes}
function StrSubstLtsAct(Text: String): String;
var
  Cont: Integer;
begin
  Result := '';
  for Cont := 1 to Length(Text) do
  begin
    if Text[Cont] = 'Ç' then
      Result := Result + 'C'
    else if Text[Cont] in ['Á', 'À', 'Ã', 'Â', 'Ä'] then
      Result := Result + 'A'
    else if Text[Cont] in ['É', 'È', 'Ê', 'Ë'] then
      Result := Result + 'E'
    else if Text[Cont] in ['Í', 'Ì', 'Î', 'Ï'] then
      Result := Result + 'I'
    else if Text[Cont] in ['Ó', 'Ò', 'Õ', 'Ô', 'Ö'] then
      Result := Result + 'O'
    else if Text[Cont] in ['Ú', 'Ù', 'Û', 'Ü'] then
      Result := Result + 'U'
    else if Text[Cont] = 'ç' then
      Result := Result + 'c'
    else if Text[Cont] in ['á', 'à', 'ã', 'â', 'ä'] then
      Result := Result + 'a'
    else if Text[Cont] in ['é', 'è', 'ê', 'ë'] then
      Result := Result + 'e'
    else if Text[Cont] in ['í', 'ì', 'î', 'ï'] then
      Result := Result + 'i'
    else if Text[Cont] in ['ó', 'ò', 'õ', 'ô', 'ö'] then
      Result := Result + 'o'
    else if Text[Cont] in ['ú', 'ù', 'û', 'ü'] then
      Result := Result + 'u'
    else
      Result := Result + Text[Cont];
  end;
end;

{Substitui letras acentuadas por letras sem acentos e remove tudo o que não for
 letra, numero ou espaço em branco de Text}
function StrRemPont(Text: String): String;
var
  Cont: Integer;
begin
  Result := '';
  Text := UpperCase(StrSubstLtsAct(Trim(Text)));
  for Cont := 1 to Length(Text) do
  begin
    if Text[Cont] in ['A'..'Z', '0'..'9', ' ', '.', '/'] then
      Result := Result + Text[Cont]
    else
      Result := Result + ' ';
  end;
end;

{Retorna somente os caracteres que forem números da função}
function StrRetNums(Text: String): String;
var
  Cont: Integer;
begin
  Result := '';
  for Cont := 1 to Length(Text) do
  begin
    if Text[Cont] in (['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']) then
      Result := Result + Text[Cont];
  end;
end;

{Divide Text nos lugares aonde há um Separator e retorna uma lista com as partes,
retorna uma lista com um único elemento = Text caso não encontre um Separator}
function StrSplit(Text, Separator: String): TStringList;
begin
  Result := TStringList.Create;
  Result.Clear;
  while Pos(Separator, Text) > 0 do
  begin
    Result.Add(copy(Text, 1, Pos(Separator, Text) - 1));
   	Delete(Text, 1, Pos(Separator, Text) + Length(Separator) - 1);
  end;
  Result.Add(Text);
end;

{Converte a string Numero para um integer ou cria uma excessão dizendo que
Nome não é um número válido.}
function StrToIntErr(Numero, Nome: String): Int64;
begin
  try
    Result := StrToInt64(Numero);
  except
    Raise Exception.Create('O ' + Nome + ' "' + Numero + '" não é um número válido.')
  end;
end;

{Tenta converter a string Data para o tipo datetime, cria uma excessão com Nome
caso não consiga.}
function StrToDateErr(Data, Nome: String): TDateTime;
begin
  try
    Result := StrToDate(Data);
  except
    Raise Exception.Create('"' + Data + '" não é uma ' + Nome + '" válida.')
  end;
end;

{Deleta todos os caracteres Character da frase Text}
function StrDelChar(const Text: String; Character: Char): String; overload;
var
  Counter: Integer;
begin
  Result := Text;
  for Counter := Length(Result) downto 1 do
  begin
    if Result[Counter] = Character then
      Delete(Result, Counter, 1);
  end;
end;

{Deleta de Text todos caracteres de CharList, Charlist é uma lista de caracteres
sparada por virgulas, ex: 'a,b,c,d'}
function StrDelChar(const Text: String; CharList: String): String; overload;
var
  CurChar, CurPos: Integer;
  Deleted: Boolean;
  CharArray: TStringList;
  Teste1, Teste2: String;
begin
  CharArray := StrSplit(CharList, ',');
  Deleted := False;
  Result := Text;
  CurPos := 1;
  while (CurPos <= Length(Result)) do
  begin
    CurChar := 0;
    Teste1 := Result[CurPos];
    while (CurChar <= CharArray.Count - 1) do
    begin
      Teste2 := CharArray[CurChar];
      if Result[CurPos] = CharArray[CurChar] then
      begin
        Delete(Result, CurPos, 1);
        Deleted := True;
        Break;
      end
      else
      CurChar := CurChar + 1;
    end;
{Se o caracter da posição atual não foi deletado o loop externo avança para a
 proxima posição de Result, senão o loop externo é executado novamente na mesma
 posição pois um novo caracter está ocupando esta posição e portanto é necessário
 verificar se ele deve ser deletado}
    if Deleted then
      Deleted := False
    else
      CurPos := CurPos + 1;
  end;
end;

{ TCronometro }

function TCronometro.Contagem :Int64;
begin
  Result := Self.ContagemFinal - Self.ContagemInicial;
end;

procedure TCronometro.Exibir;
begin
  ShowMessage(IntToStr(Self.ContagemFinal - Self.ContagemInicial) + ' ticks');
end;

procedure TCronometro.Iniciar;
begin
  Self.ContagemInicial := GetTickCount;
end;

procedure TCronometro.Parar;
begin
  Self.ContagemFinal := GetTickCount;
end;

function StrReplace(Text, Target, Replace: String): String;
begin
  Result := '';
  while Pos(Target, Text) > 0 do
  begin
    Result := Result + Copy(Text, 1, Pos(Target, Text) - 1) + Replace;
    Delete(Text, 1, (Pos(Target, Text) + Length(Target) - 1));
  end;
  Result := Result + Text;
end;

{Substituí vários espaços em branco seguindos por um único espaço}
function StrDelMultSpaces(Text :String): String;
begin
  Result := Trim(Text);
  while Pos('  ', Result) > 0 do
    Result := StringReplace(Result, '  ', ' ', [rfReplaceAll]);
end;

procedure DumpQuery(Query :TSQLQuery; FileName :String);
var
  c :Integer;
  Arq :TextFile;
  QueryText, ParamName, ParamValue, TypeName :String;
  DataType :TFieldType;
begin
  QueryText := Query.SQL.Text;
  for c := 0 to Query.Params.Count - 1 do
  begin
    ParamName := Query.Params[c].Name;
    DataType := Query.Params[c].DataType;
    if Query.Params[c].IsNull then
      ParamValue := 'NULL'
    else if DataType in [ftInteger, ftSmallint, ftWord] then
      ParamValue := Query.Params[c].AsString
    else if DataType in [ftFloat, ftCurrency, ftBCD] then
      ParamValue := FormatFloat('##########0.000', Query.Params[c].AsFloat)
    else if DataType in [ftDate, ftString] then
      ParamValue := '''' + Query.Params[c].AsString + ''''
    else if DataType = ftBoolean then
    begin
      if Query.Params[c].AsBoolean then
        ParamValue := 'true'
      else
        ParamValue := 'false';
    end
    else
    begin
      case DataType of
        ftUnknown : TypeName := 'ftUnknown';
        ftDate : TypeName := 'ftDate';
        ftTime : TypeName := 'ftTime';
        ftDateTime : TypeName := 'ftDateTime';
        ftBytes : TypeName := 'ftBytes';
        ftVarBytes : TypeName := 'ftVarBytes';
        ftAutoInc : TypeName := 'ftAutoInc';
        ftBlob : TypeName := 'ftBlob';
        ftMemo : TypeName := 'ftMemo';
        ftGraphic : TypeName := 'ftGraphic';
        ftFmtMemo : TypeName := 'ftFmtMemo';
        ftParadoxOle : TypeName := 'ftParadoxOle';
        ftDBaseOle : TypeName := 'ftDBaseOle';
        ftTypedBinary : TypeName := 'ftTypedBinary';
        ftCursor : TypeName := 'ftCursor';
        ftFixedChar : TypeName := 'ftFixedChar';
        ftWideString : TypeName := 'ftWideString';
        ftLargeint : TypeName := 'ftLargeint';
        ftADT : TypeName := 'ftADT';
        ftArray : TypeName := 'ftArray';
        ftReference : TypeName := 'ftReference';
        ftDataSet : TypeName := 'ftDataSet';
        ftOraBlob : TypeName := 'ftOraBlob';
        ftOraClob : TypeName := 'ftOraClob';
        ftVariant : TypeName := 'ftVariant';
        ftInterface : TypeName := 'ftInterface';
        ftIDispatch : TypeName := 'ftIDispatch';
        ftGuid : TypeName := 'ftGuid';
        ftTimeStamp : TypeName := 'ftTimeStamp';
        ftFMTBcd : TypeName := 'ftFMTBcd';
        else TypeName := 'Descronhecido';
      end;

      Raise Exception.Create('Parametro não tratado ' + ParamName + ': ' + TypeName);
    end;

    QueryText := StrReplace(QueryText, ':' + ParamName, ParamValue);
  end;

  AssignFile(Arq, FileName);
  Rewrite(Arq);
  WriteLn(Arq, QueryText);
  Close(Arq);
end;

procedure DumpDataSet(DataSet :TDataSet; FileName :String);
var
  c :Integer;
  Arq :TextFile;
  NomeCampo, ValorCampo, TipoCampo :String;
  DataType :TFieldType;
begin
  AssignFile(Arq, FileName);
  Rewrite(Arq);
  WriteLn(Arq, 'Nome do Campo : Tipo do Campo = Valor do Campo');

  for c := 0 to DataSet.FieldCount - 1 do
  begin
    NomeCampo := DataSet.Fields[c].DisplayName;
    DataType := DataSet.Fields[c].DataType;
    case DataType of
      ftInteger : TipoCampo := 'ftInteger';
      ftSmallint : TipoCampo := 'ftSmallint';
      ftWord : TipoCampo := 'ftWord';
      ftFloat : TipoCampo := 'ftFloat';
      ftCurrency : TipoCampo := 'ftCurrency';
      ftBCD : TipoCampo := 'ftBCD';
      ftDate : TipoCampo := 'ftDate';
      ftString : TipoCampo := 'ftString';
      ftBoolean : TipoCampo := 'ftBoolean';
      ftUnknown : TipoCampo := 'ftUnknown';
      ftTime : TipoCampo := 'ftTime';
      ftDateTime : TipoCampo := 'ftDateTime';
      ftBytes : TipoCampo := 'ftBytes';
      ftVarBytes : TipoCampo := 'ftVarBytes';
      ftAutoInc : TipoCampo := 'ftAutoInc';
      ftBlob : TipoCampo := 'ftBlob';
      ftMemo : TipoCampo := 'ftMemo';
      ftGraphic : TipoCampo := 'ftGraphic';
      ftFmtMemo : TipoCampo := 'ftFmtMemo';
      ftParadoxOle : TipoCampo := 'ftParadoxOle';
      ftDBaseOle : TipoCampo := 'ftDBaseOle';
      ftTypedBinary : TipoCampo := 'ftTypedBinary';
      ftCursor : TipoCampo := 'ftCursor';
      ftFixedChar : TipoCampo := 'ftFixedChar';
      ftWideString : TipoCampo := 'ftWideString';
      ftLargeint : TipoCampo := 'ftLargeint';
      ftADT : TipoCampo := 'ftADT';
      ftArray : TipoCampo := 'ftArray';
      ftReference : TipoCampo := 'ftReference';
      ftDataSet : TipoCampo := 'ftDataSet';
      ftOraBlob : TipoCampo := 'ftOraBlob';
      ftOraClob : TipoCampo := 'ftOraClob';
      ftVariant : TipoCampo := 'ftVariant';
      ftInterface : TipoCampo := 'ftInterface';
      ftIDispatch : TipoCampo := 'ftIDispatch';
      ftGuid : TipoCampo := 'ftGuid';
      ftTimeStamp : TipoCampo := 'ftTimeStamp';
      ftFMTBcd : TipoCampo := 'ftFMTBcd';
      else TipoCampo := 'Descronhecido';
    end; //case

    if DataSet.Fields[c].IsNull then
      ValorCampo := 'NULL'
    else if DataType in [ftInteger, ftSmallint, ftWord] then
      ValorCampo := DataSet.Fields[c].AsString
    else if DataType in [ftFloat, ftCurrency, ftBCD] then
      ValorCampo := FormatFloat('##########0.000', DataSet.Fields[c].AsFloat)
    else if DataType in [ftDate, ftString] then
      ValorCampo := '''' + DataSet.Fields[c].AsString + ''''
    else if DataType = ftBoolean then
    begin
      if DataSet.Fields[c].AsBoolean then
        ValorCampo := 'true'
      else
        ValorCampo := 'false';
    end
    else
      Raise Exception.Create('Parametro não tratado ' + NomeCampo + ': ' + TipoCampo);

    WriteLn(Arq, NomeCampo + ' : ' + TipoCampo + ' = ' + ValorCampo);
  end; //while

  Close(Arq);
end;

function TiraZerosEsquerda(const Value: string): string;
begin
  Result := value;
  if Length(Result) > 0 then
    while Result[1] = '0' do
      Delete(Result, 1, 1);
end;

function MemoryStreamToString(M: TMemoryStream): AnsiString;
begin
  SetString(Result, PAnsiChar(M.Memory), M.Size);
end;

function SplitString(const Str: string; Delimiter : Char) : TStringList;
begin
   Result := TStringList.Create;
   Result.LineBreak := Delimiter;
   Result.Text := Str;
end;

end.
