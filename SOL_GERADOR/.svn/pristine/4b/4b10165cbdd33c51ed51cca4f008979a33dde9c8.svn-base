unit UClasses;

interface

uses
  System.SysUtils, System.Classes, System.Variants, DateUtils, Data.DB, Datasnap.DBClient,
  Data.SqlExpr, Data.Win.ADODB, Vcl.Dialogs;

type

  TTypes = (tpVariant, tpString, tpFloat, tpInteger, tpBoolean, tpDateTime);
  TSides = (tpLeft, tpRight);
  TFieldTypes = (ftpString, ftpInteger, ftpFloat, ftpBoolean, ftpDateTime);
  TFileExtensions = (fxtTxt, fxtCsv);

  TLayoutField = class(TObject)
  private
    FName: String;
    FValue: Variant;
    FFormatted, FFiller: String;
    FSide: TSides;
    FSize: Integer;
    FSettedType: TTypes;
    FInitialized: Boolean;
    FFileExtension: TFileExtensions;
  protected
    procedure SetName(const Value: String);
    procedure SetAsVariant(const Value: Variant); virtual;
    procedure SetAsString(const Value: String);
    procedure SetAsFloat(const Value: Extended); virtual;
    procedure SetAsInteger(const Value: Integer); virtual;
    procedure SetAsBoolean(const Value: Boolean);
    procedure SetAsDateTime(const Value: TDateTime); virtual;
    procedure SetFormatted(const Value: String); virtual;
    procedure SetSize(const Value: Integer);
    procedure SetFiller(const Value: String);
    procedure SetSide(const Value: TSides);
    function GetName: String;
    function GetAsVariant: Variant;
    function GetAsString: String;
    function GetAsFloat: Extended;
    function GetAsInteger: Integer;
    function GetAsBoolean: Boolean;
    function GetAsDateTime: TDateTime;
    function GetFormatted: String;
    function GetFiller: String;
    function GetSide: TSides;
    function GetSize: Integer;
  public
    constructor Create(Name: String = ''; Size: Integer = 10; Filler: String = ' '; Side: TSides = tpRight );
    property Name: String read GetName write SetName;
    property Value: Variant read GetAsVariant write SetAsVariant;
    property AsString: String read GetAsString write SetAsString;
    property AsFloat: Extended read GetAsFloat write SetAsFloat;
    property AsCurrency: Extended read GetAsFloat write SetAsFloat;
    property AsInteger: Integer read GetAsInteger write SetAsInteger;
    property AsBoolean: Boolean read GetAsBoolean write SetAsBoolean;
    property AsDateTime: TDateTime read GetAsDateTime write SetAsDateTime;
    property Formatted: String read GetFormatted write SetFormatted;
    property Filler: String read GetFiller write SetFiller;
    property Side: TSides read GetSide write SetSide;
    property Size: Integer read GetSize write SetSize;
    procedure SetFileExtension(FileExtension : TFileExtensions);//!CSV
  end;

  TLayoutFieldInt = class(TLayoutField)
  protected
    procedure SetAsVariant(const Value: Variant); override;
    procedure SetAsInteger(const Value: Integer); override;
    procedure SetFormatted(const Value: String); overload;
    procedure SetFormatted(const Value: Integer); overload;
  public
    constructor Create(Name: String = ''; Size: Integer = 10; Filler: String = '0'; Side: TSides = tpLeft); overload;
  end;

  TLayoutFieldFloat = class(TLayoutField)
  private
    FFormat: String;
    FNumDecimals: Integer;
  protected
    procedure SetAsVariant(const Value: Variant); override;
    procedure SetAsFloat(const Value: Extended); override;
    procedure SetFormat(const Value: String);
    procedure SetFormatted(const Value: String); overload;
    procedure SetFormatted(const Value: Extended); overload;
    function GetFormat: String;
  public
    property Format: String read GetFormat write SetFormat;
    constructor Create(Name: String = ''; Size: Integer = 10; Filler: String = '0'; Side: TSides = tpLeft; Format: String = '0.000'); overload;
  end;

  TLayoutFieldBoolean = class(TLayoutField)
  protected
    procedure SetAsVariant(const Value : Variant); override;
    procedure SetFormatted(const Value: String); overload;
    procedure SetFormatted(const Value: Boolean); overload;
  public
    constructor Create(Name: String = ''); overload;
  end;

  TLayoutFieldDate = class(TLayoutField)
  private
    FFormat: String;
  protected
    procedure SetAsVariant(const Value: Variant); override;
    procedure SetAsDateTime(const Value: TDateTime); override;
    procedure SetFormat(const Value: String);
    procedure SetFormatted(const Value: String); overload;
    procedure SetFormatted(const Value: TDateTime); overload;
    function GetFormat: String;
  public
    property Format: String read GetFormat write SetFormat;
    constructor Create(Name: String = ''; Size: Integer = 10; Filler: String = ' '; Side: TSides = tpRight; Format: String = 'dd/mm/yyyy'); overload;
  end;

  { Permite manipular uma coleção de Objetos }
  TListFields = class(TObject)
  protected
    FItems: Array of TLayoutField;
    Fftpypes: Array of TFieldTypes;
  public
    procedure Add(Item: TLayoutField; ftpype: TFieldTypes); //adiciona um item ao final da lista
    procedure Push(Item: TLayoutField; ftpype: TFieldTypes); //o mesmo que Add
    function Pop:TObject; //remove um item do fim da lista e retorna o mesmo
    procedure UnShiftp(Item: TLayoutField; ftpype: TFieldTypes); //adiciona um item ao começo da lista
    function Shiftp:TObject; //remove um item do começo da lista e retorna o mesmo
    function Item(ItemIndex: Integer):TLayoutField; //retorna um item especifico
    function IsEmpty:Boolean; //retorna um booleano informando se a lista está vazia
    function Length:Integer; //retorna o tamanho da lista
    function IndexOf(const FieldName: String): Integer;
    function FieldByName(const FieldName: String): TLayoutField;
  end;

  TLayout = class(TObject)
  private
    FName, FFileName: String;
    FFields: TListFields;
    FClient: TClientDataSet;
    FFileExtension: TFileExtensions;//!CSV
    procedure WriteHeaders;
  protected
    FPathDir :String;
    procedure SetPathDir( Values :String );
  public
    constructor Create(Name, FileName : String; FileExtension: TFileExtensions = fxtTxt);//!CSV
    procedure AddField(Name: String; Size: Integer = 10; Filler: String = ' '; Side: TSides = tpRight; FieldType: TFieldTypes = ftpString; Format: String = '');
    procedure AddFieldString(Name: String; Size: Integer = 10; Filler: String = ' '; Side: TSides = tpRight);
    procedure AddFieldInteger(Name: String; Size: Integer = 10; Filler: String = '0'; Side: TSides = tpLeft);
    procedure AddFieldFloat(Name: String; Size: Integer = 10; Filler: String = '0'; Side: TSides = tpLeft; Format: String = '0.000');
    procedure AddFieldBoolean(Name: String);
    procedure AddFieldDateTime(Name: String; Size: Integer = 10; Filler: String = ' '; Side: TSides = tpRight; Format: String = 'dd/mm/yyyy');
    procedure SetValues(Qry: TSQLQuery; NumLinha: Integer; Count: Integer); overload;
    procedure SetValues(Qry: TADOQuery; NumLinha: Integer; Count: Integer); overload;
    //procedure SetValues(Qry: TSmartQuery; NumLinha: Integer; Count: Integer); overload;
    procedure Start( Name :String = ''; FileName : String = ''; FileExtension: TFileExtensions = fxtTxt );//!CSV
    procedure WriteLine;
    procedure Finish;

    procedure CreateDataSet(Client: TClientDataSet);
    procedure Append;
    procedure SetFileExtension(FileExtesion : TFileExtensions);//!CSV

    function GetLine: String;
    function FieldCount: Integer;
    function Field(const ItemIndex: Integer): TLayoutField;
    function FieldByName(const FieldName: String): TLayoutField;

    property SetDirectory: String read FPathDir write SetPathDir;
  end;

var
  TxtFile: TextFile;
  PathDir :String;

implementation



uses UUtilidades, UProgresso;

{ TField }

constructor TLayoutField.Create(Name: String; Size: Integer; Filler: String; Side: TSides );
begin
  FInitialized := False;
  SetName(Name);
  SetSize(Size);
  SetFiller(Filler);
  SetSide(Side);
  SetAsString('');
  FInitialized := True;
end;

function TLayoutField.GetAsBoolean: Boolean;
begin
  Result := FValue;
end;

function TLayoutField.GetAsDateTime: TDateTime;
begin
  Result := FValue;
end;

function TLayoutField.GetAsFloat: Extended;
begin
  Result := FValue;
end;

function TLayoutField.GetAsInteger: Integer;
begin
  Result := FValue;
end;

function TLayoutField.GetAsString: String;
begin
  Result := FValue;
end;

function TLayoutField.GetAsVariant: Variant;
begin
  Result := FValue;
end;

function TLayoutField.GetFiller: String;
begin
  Result := FFiller;
end;

function TLayoutField.GetFormatted: String;
begin
  Result := FFormatted;
end;

function TLayoutField.GetName: String;
begin
  Result := FName;
end;

function TLayoutField.GetSide: TSides;
begin
  Result := FSide;
end;

function TLayoutField.GetSize: Integer;
begin
  Result := FSize;
end;

procedure TLayoutField.SetAsBoolean(const Value: Boolean);
begin
  FValue := Value;
  FSettedType := tpBoolean;
  SetFormatted(BoolToStr(Value));
end;

procedure TLayoutField.SetAsDateTime(const Value: TDateTime);
begin
  try
    FValue := Value;
    FSettedType := tpDateTime;
    SetFormatted(DateTimeToStr(Value));
  except
    FValue := 0;
    FSettedType := tpDateTime;
    SetFormatted(DateTimeToStr(0));
  end;
end;

procedure TLayoutField.SetAsFloat(const Value: Extended);
begin
  FValue := Value;
  FSettedType := tpFloat;
  SetFormatted(FloatToStr(Value));
end;

procedure TLayoutField.SetAsInteger(const Value: Integer);
begin
  FValue := Value;
  FSettedType := tpInteger;
  SetFormatted(IntToStr(Value));
end;

procedure TLayoutField.SetAsString(const Value: String);
begin
  FValue := Value;
  FSettedType := tpString;
  SetFormatted(Value);
end;

procedure TLayoutField.SetAsVariant(const Value: Variant);
begin
  FValue := Value;
  FSettedType := tpVariant;
  SetFormatted(Value);
end;

procedure TLayoutField.SetFileExtension(FileExtension: TFileExtensions);//!CSV
begin
  FFileExtension := FileExtension;
end;

procedure TLayoutField.SetFiller(const Value: String);
begin
  FFiller := Value;
end;

procedure TLayoutField.SetFormatted(const Value: String);
begin
  if FInitialized then
  begin
    if FFileExtension = fxtTxt then//!CSV
      FFormatted := StrPad(Value, FSize, FFiller, FSide)
    else
      FFormatted := '"' + Value + '"';
  end;
end;

procedure TLayoutField.SetName(const Value: String);
begin
  FName := Value;
end;

procedure TLayoutField.SetSide(const Value: TSides);
begin
  FSide := Value;
  SetFormatted(GetAsString);
end;

procedure TLayoutField.SetSize(const Value: Integer);
begin
  FSize := Value;
  SetFormatted(GetAsString);
end;

{ TLayoutFieldInt }

constructor TLayoutFieldInt.Create(Name: String; Size: Integer; Filler: String;
  Side: TSides);
begin
  FInitialized := False;
  SetName(Name);
  SetSize(Size);
  SetFiller(Filler);
  SetSide(Side);
  SetAsInteger(0);
  FInitialized := True;
end;

procedure TLayoutFieldInt.SetAsInteger(const Value: Integer);
begin
  inherited;
  SetFormatted(Value);
end;

procedure TLayoutFieldInt.SetAsVariant(const Value: Variant);
begin
  try
    SetAsInteger(Value);
  except
    SetAsInteger(0);
  end;
  SetFormatted(GetAsInteger);
end;

procedure TLayoutFieldInt.SetFormatted(const Value: String);
begin
  SetFormatted(StrToIntDef(Value, 0));
end;

procedure TLayoutFieldInt.SetFormatted(const Value: Integer);
begin
  if FFileExtension = fxtTxt then//!CSV
  begin
    if Value >= 0  then
      FFormatted := StrPad(Value, FSize, FFiller, FSide)
    else
      FFormatted := '-' + StrPad(Value * -1, FSize - 1, FFiller, FSide);
  end
  else
    FFormatted := IntToStr(Value);
end;

{ TLayoutFieldFloat }

constructor TLayoutFieldFloat.Create(Name: String; Size: Integer;
  Filler: String; Side: TSides; Format: String);
begin
  FInitialized := False;
  SetName(Name);
  SetSize(Size);
  SetFiller(Filler);
  SetSide(Side);
  SetFormat(Format);
  SetAsFloat(0);
  FInitialized := True;
end;

function TLayoutFieldFloat.GetFormat: String;
begin
  Result := Format;
end;

procedure TLayoutFieldFloat.SetAsFloat(const Value: Extended);
begin
  inherited;
  SetFormatted(Value);
end;

procedure TLayoutFieldFloat.SetAsVariant(const Value: Variant);
begin
  try
    SetAsFloat(Value);
  except
    SetAsFloat(0);
  end;
end;

procedure TLayoutFieldFloat.SetFormat(const Value: String);
var Stl : TStringList;
begin
  try
    Stl := StrSplit(Value, '.');

    FormatFloat(Value, 1.55444);//teste

    if Stl.Count > 1 then
      FNumDecimals := Length(Stl[1])
    else
      FNumDecimals := 0;

    FFormat := Value;
  except
    FFormat := '';
  end;

  SetFormatted(GetAsFloat);

  Stl.Free;
end;

procedure TLayoutFieldFloat.SetFormatted(const Value: Extended);//!CSV
var
  VStr: String;
begin

  if FFormat <> '' then
   VStr := FormatFloat(FFormat, Value)
  else
   VStr := FloatToStr(Value);

  if FFileExtension = fxtTxt then//!CSV
  begin

    VStr := StringReplace(VStr, ',', '',[rfReplaceAll, rfIgnoreCase]);
    VStr := StringReplace(VStr, '.', '',[rfReplaceAll, rfIgnoreCase]);
    VStr := StringReplace(VStr, '-', '',[rfReplaceAll, rfIgnoreCase]);

    if Value >= 0  then
      FFormatted := StrPad(VStr, FSize, FFiller, FSide)
    else
      FFormatted := '-' + StrPad(VStr, FSize - 1, FFiller, FSide);

  end
  else
    FFormatted := VStr;

end;

procedure TLayoutFieldFloat.SetFormatted(const Value: String);
begin
  SetFormatted(StrToFloatDef(Value, 0));
end;

{ TLayoutFieldBoolean }

constructor TLayoutFieldBoolean.Create(Name: String);
begin
  FInitialized := False;
  SetName(Name);
  SetSize(1);
  SetFiller(' ');
  SetSide(tpRight);
  SetAsBoolean(False);
  FInitialized := True;
end;

procedure TLayoutFieldBoolean.SetAsVariant(const Value: Variant);
begin
  try
    SetAsBoolean(Value);
  except
    try
      SetAsBoolean(Value = 'S');
    except
      SetAsBoolean(False);
    end;
  end;
  SetFormatted(GetAsBoolean);
end;


procedure TLayoutFieldBoolean.SetFormatted(const Value: String);
begin
  SetFormatted(StrToBoolDef(Value, False));
end;

procedure TLayoutFieldBoolean.SetFormatted(const Value: Boolean);
begin
  //FFormatted := StrPad(iif(Value, 'S', 'N'), FSize, FFiller, FSide);
  FFormatted := iif(Value, 'S', 'N');
end;

{ TListFields }

//adiciona um item ao final da lista
procedure TListFields.Add(Item: TLayoutField; ftpype: TFieldTypes);
begin
  SetLength(Self.FItems, System.Length(Self.FItems) + 1);
  Self.FItems[System.Length(Self.FItems) - 1] := Item;
  SetLength(Self.Fftpypes, System.Length(Self.Fftpypes) + 1);
  Self.Fftpypes[System.Length(Self.Fftpypes) - 1] := ftpype;
end;

function TListFields.FieldByName(const FieldName: String): TLayoutField;
var
  ItemIndex: Integer;
begin
  try
    ItemIndex := IndexOf(FieldName);
    Result := FItems[ItemIndex];

    {if Fftpypes[ItemIndex] = ftpInteger then
      Result := (Result as TLayoutFieldInt)
    else if Fftpypes[ItemIndex] = ftpFloat then
      Result := (Result as TLayoutFieldFloat)
    else if Fftpypes[ItemIndex] = ftpBoolean then
      Result := (Result as TLayoutFieldBoolean);}

  except
    raise;
  end;
end;

function TListFields.IndexOf(const FieldName: String): Integer;
var
  ItemIndex: Integer;
begin
  Result := -1;
  for ItemIndex := 0 to System.Length(Self.FItems) - 1 do
    if Self.FItems[ItemIndex].Name = FieldName then
    begin
      Result := ItemIndex;
      Break;
    end;
end;

function TListFields.IsEmpty: Boolean;
begin
  if System.Length(Self.FItems) > 0 then
    Result := false
  else
    Result := true;
end;

//retorna um item especifico da lista, mas não remove ele
function TListFields.Item(ItemIndex: Integer): TLayoutField;
begin
  Result := Self.FItems[ItemIndex];
end;

//remove um item do fim da lista e retorna o mesmo
function TListFields.Pop: TObject;
begin
  Result := Self.FItems[System.Length(Self.FItems) - 1];
  SetLength(Self.FItems, System.Length(Self.FItems) - 1);
  SetLength(Self.Fftpypes, System.Length(Self.Fftpypes) - 1);
end;

//o mesmo que Add
procedure TListFields.Push(Item: TLayoutField; ftpype: TFieldTypes);
begin
  Self.Add(Item, ftpype);
end;

//remove um item do começo da lista e retorna o mesmo
function TListFields.Shiftp: TObject;
var
  ItemIndex: Integer;
begin
  Result := Self.FItems[0];
  for ItemIndex := 1 to System.Length(Self.FItems) - 1 do
  begin
    Self.FItems[ItemIndex - 1]  := Self.FItems[ItemIndex];
    Self.Fftpypes[ItemIndex - 1] := Self.Fftpypes[ItemIndex];
  end;
  SetLength(Self.FItems, System.Length(Self.FItems) - 1);
  SetLength(Self.Fftpypes, System.Length(Self.Fftpypes) - 1);
end;

//retorna o tamanho da lista
function TListFields.Length: Integer;
begin
  Result := System.Length(Self.FItems);
end;

//adiciona um item ao começo da lista
procedure TListFields.UnShiftp(Item: TLayoutField; ftpype: TFieldTypes);
var
  ItemIndex :Integer;
begin
  SetLength(Self.FItems, System.Length(Self.FItems) + 1);
  SetLength(Self.Fftpypes, System.Length(Self.Fftpypes) + 1);
  for ItemIndex := System.Length(Self.FItems) - 1 downto 1 do
  begin
    Self.FItems[ItemIndex]  := Self.FItems[ItemIndex - 1];
    Self.Fftpypes[ItemIndex] := Self.Fftpypes[ItemIndex - 1];
  end;
  Self.FItems[0]  := Item;
  Self.Fftpypes[0] := ftpype;
end;

{ TLayout }


procedure TLayout.AddField(Name: String; Size: Integer; Filler: String;
  Side: TSides; FieldType: TFieldTypes; Format: String);
begin
  if FieldType = ftpString then
    AddFieldString(Name, Size, Filler, Side)
  else if FieldType = ftpInteger then
    AddFieldInteger(Name, Size, Filler, Side)
  else if FieldType = ftpFloat then
    AddFieldFloat(Name, Size, Filler, Side, Format)
  else if FieldType = ftpBoolean then
    AddFieldBoolean(Name)
  else if FieldType = ftpDateTime then
    AddFieldDateTime(Name, Size, Filler, Side, Format);
end;

procedure TLayout.AddFieldBoolean(Name: String);
begin
  FFields.Add(TLayoutFieldBoolean.Create(Name), ftpBoolean);
end;

procedure TLayout.AddFieldDateTime(Name: String; Size: Integer; Filler: String;
  Side: TSides; Format: String);
begin
  FFields.Add(TLayoutFieldDate.Create(Name, Size, Filler, Side, Format), ftpDateTime);
end;

procedure TLayout.AddFieldFloat(Name: String; Size: Integer; Filler: String;
  Side: TSides; Format: String);
begin
  FFields.Add(TLayoutFieldFloat.Create(Name, Size, Filler, Side, Format), ftpFloat);
end;

procedure TLayout.AddFieldInteger(Name: String; Size: Integer; Filler: String;
  Side: TSides);
begin
  FFields.Add(TLayoutFieldInt.Create(Name, Size, Filler, Side), ftpInteger)
end;

procedure TLayout.AddFieldString(Name: String; Size: Integer; Filler: String;
  Side: TSides);
begin
  FFields.Add(TLayoutField.Create(Name, Size, Filler, Side), ftpString);
end;

procedure TLayout.Append;
var I : Integer;
begin
  FClient.Append;
  for I := 0 to Self.FFields.Length - 1 do
    if FrmProgresso.ExibirFormatado then
      FClient.FieldByName( Self.FFields.Item(I).FName ).Value := Self.FFields.Item(I).Value
    else
      FClient.FieldByName( Self.FFields.Item(I).FName ).Value := Self.FFields.Item(I).GetFormatted;

  FClient.Post;
end;

constructor TLayout.Create(Name, FileName : String; FileExtension: TFileExtensions = fxtTxt);
begin
  FName     := Name;
  FFileName := FileName;
  FFields   := TListFields.Create;
  FFileExtension := FileExtension;//!CSV
end;

procedure TLayout.CreateDataSet(Client: TClientDataSet);
var I : Integer;
begin
  FClient := Client;

  for I := 0 to Self.FFields.Length - 1 do
  begin
    if not FrmProgresso.ExibirFormatado then
      with TStringField.Create(Client) do
      begin
        DisplayLabel   := Self.FFields.Item(I).Name;
        FieldName      := Self.FFields.Item(I).Name;
        FieldKind      := fkData;
        Size           := Self.FFields.Item(I).Size;
        DataSet        := Client;
      end
    else if Self.FFields.Fftpypes[I] = ftpString then
      with TStringField.Create(Client) do
      begin
        DisplayLabel   := Self.FFields.Item(I).Name;
        FieldName      := Self.FFields.Item(I).Name;
        FieldKind      := fkData;
        Size           := Self.FFields.Item(I).Size;
        DataSet        := Client;
      end
    else if Self.FFields.Fftpypes[I] = ftpInteger then
      with TIntegerField.Create(Client) do
      begin
        DisplayLabel   := Self.FFields.Item(I).Name;
        FieldName      := Self.FFields.Item(I).Name;
        FieldKind      := fkData;
        //Size           := Self.FFields.Item(I).Size;
        DataSet        := Client;
      end
    else if Self.FFields.Fftpypes[I] = ftpFloat then
      with TFloatField.Create(Client) do
      begin
        DisplayLabel   := Self.FFields.Item(I).Name;
        FieldName      := Self.FFields.Item(I).Name;
        FieldKind      := fkData;
        //Size           := Self.FFields.Item(I).Size;
        DataSet        := Client;
      end
    else if Self.FFields.Fftpypes[I] = ftpBoolean then
      with TBooleanField.Create(Client) do
      begin
        DisplayLabel   := Self.FFields.Item(I).Name;
        FieldName      := Self.FFields.Item(I).Name;
        FieldKind      := fkData;
        //Size           := Self.FFields.Item(I).Size;
        DataSet        := Client;
      end
    else if Self.FFields.Fftpypes[I] = ftpDateTime then
      with TDateTimeField.Create(Client) do
      begin
        DisplayLabel   := Self.FFields.Item(I).Name;
        FieldName      := Self.FFields.Item(I).Name;
        FieldKind      := fkData;
        //Size           := Self.FFields.Item(I).Size;
        DataSet        := Client;
      end;

  end;

  Client.CreateDataSet;
  Client.Open;
end;

function TLayout.Field(const ItemIndex: Integer): TLayoutField;
begin
  Result := FFields.Item(ItemIndex);
end;

function TLayout.FieldByName(const FieldName: String): TLayoutField;
begin
  Result := FFields.FieldByName(FieldName);
end;

procedure TLayout.WriteHeaders;
var HStr : String;
    ItemIndex : Integer;
begin
  if FrmProgresso.ExibirGrade then
    Self.Append;

  HStr := '';

  for ItemIndex := 0 to System.Length(FFields.FItems) - 1 do
    HStr := HStr + iif(HStr <> '', ';', '') + FFields.FItems[ItemIndex].Name;

  Writeln(TxtFile, HStr);
end;

function TLayout.FieldCount: Integer;
begin
  Result := FFields.Length;
end;

procedure TLayout.SetFileExtension(FileExtesion: TFileExtensions);//!CSV
begin
  FFileExtension := FileExtesion;
end;

procedure TLayout.SetPathDir(Values: String);
begin
  if Values <> '' then
  PathDir := Values;
end;

{procedure TLayout.SetValues(Qry: TSmartQuery; NumLinha: Integer; Count: Integer);
var
  ItemIndex :Integer;
begin
    FrmProgresso.ExibirProgresso(NumLinha, Count);
  for ItemIndex := 0 to System.Length(FFields.FItems) - 1 do
  begin
    FFields.FItems[ItemIndex].AsString := '';
    try
      if FFields.Fftpypes[ItemIndex] = ftpString then
        FFields.FItems[ItemIndex].AsString := Qry.FieldByName(FFields.FItems[ItemIndex].Name).AsString
      else if FFields.Fftpypes[ItemIndex] = ftpInteger then
        FFields.FItems[ItemIndex].AsInteger := Qry.FieldByName(FFields.FItems[ItemIndex].Name).AsInteger
      else if FFields.Fftpypes[ItemIndex] = ftpFloat then
        FFields.FItems[ItemIndex].AsFloat := Qry.FieldByName(FFields.FItems[ItemIndex].Name).AsFloat
      else if FFields.Fftpypes[ItemIndex] = ftpBoolean then
        FFields.FItems[ItemIndex].Value := Qry.FieldByName(FFields .FItems[ItemIndex].Name).Value
      else if FFields.Fftpypes[ItemIndex] = ftpDateTime then
        FFields.FItems[ItemIndex].Value := Qry.FieldByName(FFields.FItems[ItemIndex].Name).Value;
    except
      on e : Exception do
      begin
        FrmProgresso.AdicionarLog(NumLinha ,'E', 'Field: ' + FFields.FItems[ItemIndex].Name + ', Error: ' + E.Message );
      end;
    end;
  end;
end;}

procedure TLayout.SetValues(Qry: TADOQuery; NumLinha: Integer; Count: Integer);
var
  ItemIndex :Integer;
begin
    FrmProgresso.ExibirProgresso(NumLinha, Count);
  for ItemIndex := 0 to System.Length(FFields.FItems) - 1 do
  begin
    FFields.FItems[ItemIndex].AsString := '';
    try
      if FFields.Fftpypes[ItemIndex] = ftpString then
        FFields.FItems[ItemIndex].AsString := Qry.FieldByName(FFields.FItems[ItemIndex].Name).AsString
      else if FFields.Fftpypes[ItemIndex] = ftpInteger then
        FFields.FItems[ItemIndex].AsInteger := Qry.FieldByName(FFields.FItems[ItemIndex].Name).AsInteger
      else if FFields.Fftpypes[ItemIndex] = ftpFloat then
        FFields.FItems[ItemIndex].AsFloat := Qry.FieldByName(FFields.FItems[ItemIndex].Name).AsFloat
      else if FFields.Fftpypes[ItemIndex] = ftpBoolean then
        FFields.FItems[ItemIndex].Value := Qry.FieldByName(FFields .FItems[ItemIndex].Name).Value
      else if FFields.Fftpypes[ItemIndex] = ftpDateTime then begin
         if FFields.FItems[ItemIndex].Value <> '' then
            FFields.FItems[ItemIndex].Value := Qry.FieldByName(FFields.FItems[ItemIndex].Name).AsDateTime;
      end;
    except
      on e : Exception do
      begin
        if not (Pos('not found', E.Message) > 0) then
          FrmProgresso.AdicionarLog(NumLinha ,'E', 'Field: ' + FFields.FItems[ItemIndex].Name + ', Error: ' + E.Message );
      end;
    end;
  end;
end;

procedure TLayout.SetValues(Qry: TSQLQuery; NumLinha: Integer; Count: Integer);
var
  ItemIndex :Integer;
  teste : string;
begin
  FrmProgresso.ExibirProgresso(NumLinha, Count);
  for ItemIndex := 0 to Pred(System.Length(FFields.FItems)) do

  begin
    FFields.FItems[ItemIndex].AsString := '';
    try
      FFields.FItems[ItemIndex].SetFileExtension(FFileExtension);
      if FFields.Fftpypes[ItemIndex] = ftpString then
        FFields.FItems[ItemIndex].AsString := Qry.FieldByName(FFields.FItems[ItemIndex].Name).AsString
      else if FFields.Fftpypes[ItemIndex] = ftpInteger then
        FFields.FItems[ItemIndex].AsInteger := Qry.FieldByName(FFields.FItems[ItemIndex].Name).AsInteger
      else if FFields.Fftpypes[ItemIndex] = ftpFloat then
        FFields.FItems[ItemIndex].AsFloat := Qry.FieldByName(FFields.FItems[ItemIndex].Name).AsFloat
      else if FFields.Fftpypes[ItemIndex] = ftpBoolean then
        FFields.FItems[ItemIndex].Value := Qry.FieldByName(FFields.FItems[ItemIndex].Name).Value
      else if FFields.Fftpypes[ItemIndex] = ftpDateTime then begin
         if FFields.FItems[ItemIndex].Value <> '' then
            FFields.FItems[ItemIndex].Value := Qry.FieldByName(FFields.FItems[ItemIndex].Name).AsDateTime;
      end;

     except
     on e : Exception do
      begin
        if not (Pos('not found', E.Message) > 0) then
          FrmProgresso.AdicionarLog(NumLinha ,'E', 'Field: ' + FFields.FItems[ItemIndex].Name + ', Error: ' + E.Message );
      end;
    end;
  end;
end;

function TLayout.GetLine: String;
var
  ItemIndex :Integer;
begin
  Result := '';
  if FFileExtension = fxtTxt then //!CSV
    for ItemIndex := 0 to System.Length(FFields.FItems) - 1 do
      Result := Result + FFields.FItems[ItemIndex].GetFormatted
  else
    for ItemIndex := 0 to System.Length(FFields.FItems) - 1 do
      Result := Result + iif(Result <> '', ';', '') + FFields.FItems[ItemIndex].GetFormatted;
end;

procedure TLayout.Start( Name :string = ''; FileName : String = ''; FileExtension: TFileExtensions = fxtTxt );
 var vPath :String;
begin
  if FileName <> '' then
  begin
     FName     := Name;
     FFileName := FileName;
     FFileExtension := FileExtension;//!CSV
  end;

  if FFileExtension = fxtTxt then//!CSV
    vPath := PathDir+'\'+FFileName+'.txt'
  else
    vPath := PathDir+'\'+FFileName+'.csv';

  AssignFile(TxtFile,vPath);
  Rewrite(TxtFile);
  FrmProgresso.ExibirAcaoAtual('Gerando ' + FName);

  if FFileExtension = fxtCsv then//!CSV
    WriteHeaders;//!CSV

  //
  if FrmProgresso.ExibirGrade then
    FrmProgresso.CriarDS(Self);
end;

procedure TLayout.WriteLine;
begin
  if FrmProgresso.ExibirGrade then
    Self.Append;
  Writeln(TxtFile,GetLine);
end;

procedure TLayout.Finish;
begin
  CloseFile(TxtFile);
end;

{ TLayoutFieldDate }
constructor TLayoutFieldDate.Create(Name: String; Size: Integer; Filler: String;
  Side: TSides; Format: String);
begin
  FInitialized := False;
  SetName(Name);
  SetSize(Size);
  SetFiller(Filler);
  SetFormat(Format);
  SetAsDateTime(0);
  FInitialized := True;
end;

function TLayoutFieldDate.GetFormat: String;
begin
  Result := FFormat;
end;

procedure TLayoutFieldDate.SetAsDateTime(const Value: TDateTime);
begin
  inherited;
  SetFormatted(Value);
end;

procedure TLayoutFieldDate.SetAsVariant(const Value: Variant);
begin
  try
    SetAsDateTime(Value);
  except
    SetAsDateTime(0);
  end;
  SetFormatted(GetAsDateTime);
end;

procedure TLayoutFieldDate.SetFormat(const Value: String);
begin
  FFormat := Value;
  SetFormatted(GetAsDateTime);
end;

procedure TLayoutFieldDate.SetFormatted(const Value: String);
begin
  SetFormatted(StrToDate(Value));
end;

procedure TLayoutFieldDate.SetFormatted(const Value: TDateTime);
var
  VStr: String;
begin

  if Value > 0 then
    VStr := DateToFStr(Value, GetFormat)
  else
    VStr := '';

  if FFileExtension = fxtTxt then//!CSV
    FFormatted := StrPad(VStr, FSize, FFiller, FSide)
  else
    FFormatted := '"' + VStr + '"';

end;

end.

