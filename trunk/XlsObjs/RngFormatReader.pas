unit RngFormatReader;

interface

uses
  Classes, Ranger;

const
  ELEMENT_COLOR                = 'Color';
  ATRIBUT_ITEM_ID              = 'ItemID';
  ATRIBUT_VALUE                = 'Value';

  ELEMENT_BORDER               = 'Border';
  ATRIBUTE_STYLE               = 'Style';
  ATRIBUTE_WEIGHT              = 'Weight';

  ELEMENT_FONT                 = 'Font';
  ATRIBUTE_NAME                = 'Name';
  ATRIBUTE_SIZE                = 'Size';
  ATRIBUTE_BOLD                = 'Bold';
  ATRIBUTE_ITALIC              = 'Italic';

  ELEMENT_RANGE                = 'Range';

  ELEMENT_RANGENAME            = 'RangeName';
  ELEMENT_INTERIORCOLOR        = 'InteriorColor';
  ELEMENT_RANGEFONT            = 'RangeFont';
  ELEMENT_RANGEBORDERS         = 'RangeBorders';
  ELEMENT_RANGEBORDER          = 'RangeBorder';
  ELEMENT_RANGEROWH            = 'RangeRowHeight';
  ELEMENT_RANGECOLW            = 'RangeColWidth';
  ATRIBUTE_ITEM                = 'Item';

const
  COLOR_ATTRIB_NAMES: array[ 0..1 ] of string =
  (
    ATRIBUT_ITEM_ID, ATRIBUT_VALUE
  );

  BORDER_ATTRIB_NAMES: array[ 0..3 ] of string =
  (
    ATRIBUT_ITEM_ID, ELEMENT_COLOR, ATRIBUTE_STYLE, ATRIBUTE_WEIGHT
  );

  FONT_ATTRIB_NAMES: array[ 0..5 ] of string =
  (
    ATRIBUT_ITEM_ID, ATRIBUTE_NAME, ELEMENT_COLOR,
    ATRIBUTE_SIZE, ATRIBUTE_BOLD, ATRIBUTE_ITALIC
  );

type
  TAttributes = array of string;

  PElement = ^TElement;
  TElement = record
    AttributeNames      : TAttributes;
    AttributValues      : TAttributes;
    ElementValue        : string;
  end;
  TElementArr = array of PElement;

  TrfRangeReader = class(TObject)
  private
    FFormatFile: string;
    procedure DoCreate;
  protected
    FRangeFormat: TrfRangeFormat;

    procedure ReadColors(); virtual; abstract;
    procedure ReadFonts(); virtual; abstract;
    procedure ReadBorders(); virtual; abstract;
    procedure ReadRanges(); virtual; abstract;

    class function CreateElement( ANamesOfAttr: TAttributes ): PElement;
    class procedure AddElemtoArr( AElem: PElement; var ElemArr: TElementArr );
    class procedure FreeElementArr( var AElemts: TElementArr );
    class function GetAttribsFromStrings( AStrings: array of string ): TAttributes;

    class function GetRangeRelation(
      ABaseArr: TrfBaseFormatItems; AItemID: string ): TrfBaseFormat;
    class function GetRowColDimension( AValue: string ): integer;

    procedure DoReadRangeFormats();
  public
    constructor Create(AFileName: string); virtual;

    property FormatFile: string read FFormatFile write FFormatFile;

    //Just return reference and free it in client code
    function ReadRangeFormats: TrfRangeFormat; virtual; abstract;
  end;

  function StrToBool( Value: string ): boolean;

implementation

uses
  SysUtils;

const
  TRUE_VAL = 'TRUE';


function StrToBool( Value: string ): boolean;
begin
  Result := UpperCase( Value ) = TRUE_VAL;
end;

{ TrfRangeReader }

class procedure TrfRangeReader.AddElemtoArr(AElem: PElement;
  var ElemArr: TElementArr);
begin
  SetLength( ElemArr, Length( ElemArr ) + 1 );
  ElemArr[ High( ElemArr ) ] := AElem;
end;

constructor TrfRangeReader.Create(AFileName: string);
begin
  DoCreate;
  FFormatFile := AFileName;
end;

class function TrfRangeReader.CreateElement(
  ANamesOfAttr: TAttributes): PElement;
begin
  New( Result );

  Result.AttributeNames := ANamesOfAttr;
  Setlength( Result.AttributValues, Length( ANamesOfAttr ) );
end;

procedure TrfRangeReader.DoCreate;
begin
  FRangeFormat := TrfRangeFormat.Create;
end;

procedure TrfRangeReader.DoReadRangeFormats;
begin
  //Step 1
  ReadColors();
  //Step 2
  ReadFonts();
  //Step 3
  ReadBorders();
  //Step 4
  ReadRanges();
end;

class procedure TrfRangeReader.FreeElementArr(var AElemts: TElementArr);
var
  i: integer;
begin
  for i := Low( AElemts ) to High( AElemts ) do
  begin
    Assert( AElemts[ i ] <> nil );
    Dispose( AElemts[ i ] );
  end;
  SetLength( AElemts, 0 );
end;

class function TrfRangeReader.GetAttribsFromStrings(
  AStrings: array of string): TAttributes;
var
  i: integer;
begin
  SetLength( Result, Length( AStrings ) );

  for i := Low( AStrings ) to High( AStrings ) do
  begin
    Result[ i ] := AStrings[ i ];
  end;
end;

class function TrfRangeReader.GetRangeRelation(
  ABaseArr: TrfBaseFormatItems; AItemID: string): TrfBaseFormat;
begin
  Result := nil;
  Result := TrfRangeFormat.GetItemFromArr( ABaseArr, AItemID );
end;

class function TrfRangeReader.GetRowColDimension(AValue: string): integer;
begin
  Result := WIDTH_DEFAULT;
  try
    Result := StrToInt( AValue );
  except

  end;

  if Result < WIDTH_AUTOFIT then
    Result := WIDTH_DEFAULT
end;

end.
