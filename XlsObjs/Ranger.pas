unit Ranger;

interface

uses
  Windows, Sysutils, Classes, Graphics, IniFiles, Excel97;

const
  STR_COLOR_SEC = 'COLOR_';
  STR_BORDER_SEC = 'BORDER_';
  STR_FONT_SEC = 'FONT_';
  STR_RANGE_SEC = 'RANGE_';

const
  WIDTH_HIDE    = -1;
  WIDTH_AUTOFIT = -2;
  WIDTH_DEFAULT = 0;

type

  ERangeFormatException = class(Exception);

  TrfBaseFormat = class
  private
    FItemID: string;
  public
    property ItemID: string read FItemID write FItemID;
  end;

  TrfBaseFormatItems = array of TrfBaseFormat;

  TrfColor = class(TrfBaseFormat)
  private
    FColor: TColor;
    function GetColorAsString: String;
    procedure SetColorAsString(const Value: String);
  public
    property Color: TColor read FColor write FColor;
    property ColorAsString: String read GetColorAsString write SetColorAsString;
    function ColorAsLong: longint;
  end;

const
//  xlBorderAround        = $0000000D;
  BORDERT_FIRST         = xlDiagonalDown;
  BORDERT_LAST          = xlInsideHorizontal;

type

  TrfBorder = class(TrfBaseFormat)
  private
    FLineColor: TrfColor;
    FWeight: XlBorderWeight;
    FLineStyle: XlLineStyle;
  public
    property LineColor: TrfColor read FLineColor write FLineColor;
    property LineStyle: XlLineStyle read FLineStyle write FLineStyle;
    property Weight: XlBorderWeight read FWeight write FWeight;
  end;

  TrfBorderItems = array [ BORDERT_FIRST..BORDERT_LAST ] of TrfBorder;


  TrfFont = class(TrfBaseFormat)
  private
    FItalic: boolean;
    FBold: boolean;
    FSize: integer;
    FName: String;
    FColor: TrfColor;
  public
    property Name: String read FName write FName;
    property Color: TrfColor read FColor write FColor;
    property Size: integer read FSize write FSize;
    property Bold: boolean read FBold write FBold;
    property Italic: boolean read FItalic write FItalic;

  end;


  TrfRange = class(TrfBaseFormat)
  private
    FRangeName: String;
    FTextFont: TrfFont;
    FInteriorColor: TrfColor;
    FBorders: TrfBorderItems;
    FRowHeight: integer;
    FColWidth: integer;

    procedure ClearBorders;
    function GetBorderItem(Index: integer): TrfBorder;
    procedure SetBorderItem(Index: integer; const Value: TrfBorder);
  public
    constructor Create;
    property InteriorColor: TrfColor read FInteriorColor write FInteriorColor;
    property TextFont: TrfFont read FTextFont write FTextFont;
    property Borders: TrfBorderItems read FBorders;
    property BorderItem[ Index: integer ]: TrfBorder read GetBorderItem write SetBorderItem;
    property RangeName: String read FRangeName write FRangeName;
    property RowHeight: integer read FRowHeight write FRowHeight;
    property ColWidth: integer read FColWidth write FColWidth;

  end;

  TrfRangeFormat = class
  private
    FColorItems: TrfBaseFormatItems;
    FFontItems: TrfBaseFormatItems;
    FBorderItems: TrfBaseFormatItems;
    FRangeFormats: TrfBaseFormatItems;

    procedure ClearArrays;
    procedure ClearBaseItems( AItems: TrfBaseFormatItems );

    function GetBaseFormatItem( AItems: TrfBaseFormatItems;
      AItemID: string ): TrfBaseFormat;

    function GetRangeFormatItem(Index: string): TrfBaseFormat;
    procedure SetRangeFormats(const Value: TrfBaseFormatItems);
  public
    destructor Destroy; override;
    property ColorItems: TrfBaseFormatItems read FColorItems write FColorItems;
    property FontItems: TrfBaseFormatItems read FFontItems write FFontItems;
    property BorderItems: TrfBaseFormatItems read FBorderItems write FBorderItems;

    property RangeFormats: TrfBaseFormatItems read FRangeFormats write FRangeFormats;
    property RangeFormatItem[ Index: string ]: TrfBaseFormat read GetRangeFormatItem;

    class procedure AddBaseItemToArr(
      AItem: TrfBaseFormat; var BaseArr: TrfBaseFormatItems );
    class function GetItemFromArr(
      ABaseArr: TrfBaseFormatItems; AItemID: string ): TrfBaseFormat;
  end;

implementation

{ TrfRangeFormat }

class procedure TrfRangeFormat.AddBaseItemToArr(AItem: TrfBaseFormat;
  var BaseArr: TrfBaseFormatItems);
begin
  Assert( AItem <> nil );
  SetLength( BaseArr, Length( BaseArr ) + 1 );

  BaseArr[ High( BaseArr ) ] := AItem;
end;

procedure TrfRangeFormat.ClearArrays;
begin
  ClearBaseItems( FColorItems );
  ClearBaseItems( FFontItems );
  ClearBaseItems( FBorderItems );
  ClearBaseItems( FRangeFormats );
end;

procedure TrfRangeFormat.ClearBaseItems(AItems: TrfBaseFormatItems);
var
  i: Integer;
begin
  for i := Low( AItems ) to High( AItems ) do
  begin
    FreeAndNil( AItems[ i ] )
  end;
  SetLength( AItems, 0 );
end;

destructor TrfRangeFormat.Destroy;
begin
  ClearArrays;

  inherited;
end;

function TrfRangeFormat.GetBaseFormatItem(
  AItems: TrfBaseFormatItems; AItemID: string): TrfBaseFormat;
var
  i: integer;
begin
  Result := nil;

  i := Low( AItems );
  while i <= High( AItems ) do
  begin
    if AItems[ i ].ItemID = AItemID then
    begin
      Result := AItems[ i ];
      i := High( AItems )
    end;
    Inc( i );
  end;
end;

class function TrfRangeFormat.GetItemFromArr(ABaseArr: TrfBaseFormatItems;
  AItemID: string): TrfBaseFormat;
var
  i: integer;
  Last: integer;
begin
  Result := nil;

  Last  := High( ABaseArr );
  i     := Low( ABaseArr );
  while i <= Last do
  begin
    if ABaseArr[ i ].ItemID = AItemID then
    begin
      Result := ABaseArr[ i ];
      i := Last;
    end;
    i := i + 1;
  end;
end;

function TrfRangeFormat.GetRangeFormatItem(
  Index: string): TrfBaseFormat;
begin
  Result := GetBaseFormatItem( FRangeFormats, Index );

  Assert( Result <> nil );
end;

procedure TrfRangeFormat.SetRangeFormats(const Value: TrfBaseFormatItems);
begin
  FRangeFormats := Value;
end;


{ TrfColor }

function TrfColor.ColorAsLong: longint;
begin
  Result := ColorToRGB( FColor );
end;

function TrfColor.GetColorAsString: String;
begin
  Result := ColorToString( FColor );
end;

procedure TrfColor.SetColorAsString(const Value: String);
begin
  FColor := StringToColor( Value );
end;

{ TrfRange }

procedure TrfRange.ClearBorders;
var
  i: integer;
begin
  for i := Low( TrfBorderItems ) to High( TrfBorderItems ) do
  begin
    FBorders[ i ] := nil;
  end;
end;

constructor TrfRange.Create;
begin
  FTextFont             := nil;
  FInteriorColor        := nil;

  RowHeight     := WIDTH_DEFAULT;
  ColWidth      := WIDTH_DEFAULT;

  ClearBorders;
end;

function TrfRange.GetBorderItem(Index: integer): TrfBorder;
begin
  Assert( ( Index >= BORDERT_FIRST ) and ( Index <= BORDERT_LAST ) );

  Result := FBorders[ Index ];
end;

procedure TrfRange.SetBorderItem(Index: integer; const Value: TrfBorder);
begin
  Assert( ( Index >= BORDERT_FIRST ) and ( Index <= BORDERT_LAST ) );

  FBorders[ Index ] := Value;
end;

end.
