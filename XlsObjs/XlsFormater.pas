unit XlsFormater;

interface

uses
  Classes, XlsObjs, Ranger;

type
  TXlsFormater = class
  private
    FBook: TXlsWorkbook;
    FFormats: TrfRangeFormat;
    FSheetName: string;
  protected
    function GetSheetRange: TXlsRanges;
    procedure InitColors( AColors: TrfBaseFormatItems );
    procedure ApplyFont( ARange: IXlsRange; AFont: TrfFont );
    procedure ApplyBorders( ARange: IXlsRange; ABorders: TrfBorderItems );
    procedure ApplyWidth( ARange: IXlsRange; AWidth: integer );
    procedure ApplyHeight( ARange: IXlsRange; AHeight: integer );
  public
    constructor Create;
    destructor Destroy; override;

    property Book: TXlsWorkbook read FBook write FBook;
    property SheetName: string read FSheetName write FSheetName;
    property RangeFormats: TrfRangeFormat read FFormats write FFormats;

    procedure ApplyFormat(); overload;
//    procedure ApplyFormat( ABook: TXlsWorkbook; AFormat: TrfRangeFormat ); overload;
  end;

implementation

uses
  SysUtils;
{ TXlsFormater }

procedure TXlsFormater.ApplyBorders(ARange: IXlsRange;
  ABorders: TrfBorderItems);
var
  i: integer;
  Border: TrfBorder;
begin
//xlDiagonalDown..xlInsideHorizontal
  for i := BORDERT_FIRST to BORDERT_LAST do
  begin
    Border := nil;
    Border := ABorders[ i ];
    if Border <> nil then
    begin
      with ARange.Borders.Item[ i ] do
      begin
        LineStyle       := Border.LineStyle;
        Weight          := Border.Weight;
        Color           := Border.LineColor.ColorAsLong;
      end;
    end;
  end;
end;

procedure TXlsFormater.ApplyFont(ARange: IXlsRange; AFont: TrfFont);
begin
  Assert( ARange <> nil );
  Assert( AFont <> nil );

  with ARange.Font do
  begin
    Size        := AFont.Size;
    Bold        := AFont.Bold;
    Italic      := AFont.Italic;
    Name        := AFont.Name;
    Color       := AFont.Color.ColorAsLong;
  end;
end;

procedure TXlsFormater.ApplyFormat;
var
  Rngs: TXlsRanges;
  IFormatRng: IXlsRange;
  FmtItem: TrfRange;
  i: integer;
begin
  Assert( FBook <> nil );
  Assert( FFormats <> nil );

  Rngs := nil;
  Rngs := GetSheetRange;
  {
  if FBook.XlsWorksheets.Count > 0 then
    Rngs := TXlsWorksheet(FBook.XlsWorksheets.Items[ 0 ]).XlsRanges;
  }
  //Nothing to format
  if Rngs = nil then
    exit; //Get out here -->>

  InitColors( FFormats.ColorItems );

  for i := Low( FFormats.RangeFormats ) to High( FFormats.RangeFormats ) do
  begin
    FmtItem := TrfRange(FFormats.RangeFormats[ i ]);

    Assert( FmtItem <> nil );

    IFormatRng := nil;
    try
      IFormatRng := Rngs.FindAsIRange( FmtItem.RangeName );
      if IFormatRng <> nil then
      begin
        if FmtItem.InteriorColor <> nil then
          IFormatRng.Interior.Color := FmtItem.InteriorColor.ColorAsLong;

        if FmtItem.TextFont <> nil then
          ApplyFont( IFormatRng, FmtItem.TextFont );

        ApplyWidth( IFormatRng, FmtItem.ColWidth );

        ApplyHeight( IFormatRng, FmtItem.RowHeight );

        ApplyBorders( IFormatRng, FmtItem.Borders );
      end;
    except

    end;
  end;
end;
{
procedure TXlsFormater.ApplyFormat(ABook: TXlsWorkbook;
  AFormat: TrfRangeFormat);
begin
  FBook    := ABook;
  FFormats := AFormat;

  ApplyFormat();
end;
}

procedure TXlsFormater.ApplyHeight(ARange: IXlsRange; AHeight: integer);
begin
  case AHeight of
    WIDTH_AUTOFIT       : ARange.Rows.AutoFit;
    WIDTH_HIDE          : ARange.Rows.Hidden := True;
    WIDTH_DEFAULT       : //Do nothing
  else
    ARange.RowHeight := AHeight;
  end;
end;

procedure TXlsFormater.ApplyWidth(ARange: IXlsRange; AWidth: integer);
begin
  case AWidth of
    WIDTH_AUTOFIT       : ARange.Columns.AutoFit;
    WIDTH_HIDE          : ARange.Columns.Hidden := True;
    WIDTH_DEFAULT       : //Do nothing
  else
    ARange.ColumnWidth := AWidth;
  end;
end;

constructor TXlsFormater.Create;
begin
  SheetName  := '';
  FBook      := nil;
  FFormats   := nil;
end;

destructor TXlsFormater.Destroy;
begin
  SheetName  := '';
  FBook      := nil;
  FFormats   := nil;

  inherited;
end;

function TXlsFormater.GetSheetRange: TXlsRanges;
var
  Sh: TXlsWorksheet;
begin
  Result := nil;
  Sh     := TXlsWorksheet(FBook.XlsWorksheets.GetByName( FSheetName ));
  if Sh <> nil then
    Result := Sh.XlsRanges;
end;

procedure TXlsFormater.InitColors(AColors: TrfBaseFormatItems);
var
  i: integer;
  ColorItm: TrfColor;
begin
  for i := Low( AColors ) to High( AColors ) do
  begin
    ColorItm := TrfColor(AColors[ i ]);
    FBook.IWorkbook.Colors[ MY_COLORS + i, DEF_LCID ] := ColorItm.ColorAsLong;
  end;
end;

end.
