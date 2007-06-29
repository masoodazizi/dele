unit XMLRngFormatReader;

interface

uses
  Classes, Ranger, RngFormatReader, IcXMLParser;

type
  TXMLRangeReader = class(TrfRangeReader)
  private
    FXMLDoc: TIcXMLDocument;
    FRoot: TIcXMLElement;
    FXMLParser: TIcXMLParser;

    procedure ParseDoc;

    function GetElementValues( AElem: TIcXMLElement; AElementName: string;
      ANamesOfAtrrib: TAttributes ): TElementArr;
    procedure GetElemValue( AElem: TIcXMLElement; ElemValue: PElement );
    procedure GetElemAttributes( AElem: TIcXMLElement; ElemValue: PElement );
    function GetElemText( AElem: TIcXMLElement ): string;

    procedure GetRangeProperties( ARng: TrfRange; AElem: TIcXMLElement );
    procedure GetRangeBorders( ARng: TrfRange; AElem: TIcXMLElement );
  protected
    procedure ReadColors(); override;
    procedure ReadFonts(); override;
    procedure ReadBorders(); override;
    procedure ReadRanges(); override;
  public
    constructor Create(AFileName: string); override;
    destructor Destroy(); override;

    function ReadRangeFormats: TrfRangeFormat; override;
  end;

implementation

uses
  SysUtils, Graphics;

{ TXMLRangeReader }

constructor TXMLRangeReader.Create(AFileName: string);
begin
  inherited;
  FXMLDoc       := nil;
  FRoot         := nil;
  FXMLParser    := TIcXMLParser.Create( nil );
end;

destructor TXMLRangeReader.Destroy;
begin
  //Created from parser and destroied.
  if Assigned( FXMLDoc ) then
    FreeAndNil( FXMLDoc );

  FreeAndNil( FXMLParser );
  inherited;
end;

procedure TXMLRangeReader.GetElemAttributes(AElem: TIcXMLElement;
  ElemValue: PElement);
var
  i: integer;
  AttrVal: string;
begin
  Assert( ElemValue <> nil );
  Assert( AElem <> nil );

  for i := Low( ElemValue.AttributeNames ) to High( ElemValue.AttributeNames ) do
  begin
    AttrVal := AElem.getAttribute( ElemValue.AttributeNames[ i ] );
    ElemValue.AttributValues[ i ] := AttrVal;
  end;
end;

function TXMLRangeReader.GetElementValues(AElem: TIcXMLElement;
  AElementName: string; ANamesOfAtrrib: TAttributes): TElementArr;
var
  xmlNodes: TIcNodeList;
  xmlElem: TIcXMLElement;
  i: integer;

  pElemVals: PElement;
begin
  Assert( AElem <> nil );

  //Case semsetive to element names
  xmlNodes := AElem.GetElementsByTagName( AElementName );
  try
    for i := 1 to xmlNodes.Length do
    begin
      xmlElem := TIcXMLElement( xmlNodes.Item( i, ntElement ) );
      if xmlElem <> nil then
      begin
        pElemVals := CreateElement( ANamesOfAtrrib );

        GetElemValue( xmlElem, pElemVals );
        GetElemAttributes( xmlElem, pElemVals );

        AddElemtoArr( pElemVals, Result );
      end;
    end;
  finally
    FreeAndNil( xmlNodes );
  end;
end;

function TXMLRangeReader.GetElemText(AElem: TIcXMLElement): string;
var
  xmlTxt: TIcXMLText;
begin
  Assert( AElem <> nil );

  xmlTxt        := AElem.GetFirstCharData;
  Result        := '';
  if xmlTxt <> nil then
    Result := xmlTxt.GetValue;
end;

procedure TXMLRangeReader.GetElemValue(AElem: TIcXMLElement;
  ElemValue: PElement);
begin
  Assert( ElemValue <> nil );
  Assert( AElem <> nil );

  ElemValue.ElementValue := GetElemText( AElem );
end;

procedure TXMLRangeReader.GetRangeBorders(ARng: TrfRange;
  AElem: TIcXMLElement);
var
  Elems: TElementArr;
  Attr: TAttributes;
  i: integer;
  Ind: integer;
begin
  Assert( ARng <> nil );
  Assert( AElem <> nil );

  Attr := GetAttribsFromStrings( [ ATRIBUTE_ITEM ] );
  Elems := GetElementValues( AElem, ELEMENT_RANGEBORDER, Attr );
  try
    for i := Low( Elems ) to High( Elems ) do
    begin
      Ind := StrToInt( Elems[ i ].AttributValues[ 0 ] );
      if ( Ind < Low( TrfBorderItems ) ) and ( Ind > High( TrfBorderItems ) ) then
        raise ERangeFormatException.Create('Border index not valid ' +
          Elems[ i ].AttributValues[ 0 ] );

      ARng.BorderItem[ Ind ] := TrfBorder(GetRangeRelation( FRangeFormat.BorderItems, Elems[ i ].ElementValue ));
    end;
  finally
    FreeElementArr( Elems );
  end;
end;

procedure TXMLRangeReader.GetRangeProperties(ARng: TrfRange;
  AElem: TIcXMLElement);
var
  ElemName: string;
  ElemVal: string;
  xmlElem: TIcXMLElement;
begin
  Assert( ARng <> nil );
  Assert( AElem <> nil );

  xmlElem := AElem.GetFirstChild;

  repeat
    ElemName := xmlElem.GetName;
    ElemVal := GetElemText( xmlElem );
    if ElemName = ELEMENT_RANGENAME then
      ARng.RangeName := ElemVal
    else if ElemName = ELEMENT_INTERIORCOLOR then
      ARng.InteriorColor := TrfColor(GetRangeRelation(
        FRangeFormat.ColorItems, ElemVal ))
    else if ElemName = ELEMENT_RANGEFONT then
      ARng.TextFont := TrfFont(GetRangeRelation(
        FRangeFormat.FontItems, ElemVal ))
    else if ElemName = ELEMENT_RANGEBORDERS then
      GetRangeBorders( ARng, AElem )
    else if ElemName = ELEMENT_RANGEROWH then
      ARng.RowHeight := GetRowColDimension( ElemVal )
    else if ElemName = ELEMENT_RANGECOLW then
      ARng.ColWidth := GetRowColDimension( ElemVal )
    else
      raise ERangeFormatException.Create('Not Expected element - ' + ElemName );
    xmlElem := xmlElem.NextSibling;
  until ( xmlElem = nil );

end;

procedure TXMLRangeReader.ParseDoc;
begin
  try
    FXMLParser.ValidateDocument := False;
    FXMLParser.StandardXML := False;
    FXMLParser.Parse( FormatFile, FXMLDoc );

    Assert( FXMLDoc <> nil );

    FRoot := FXMLDoc.GetDocumentElement;
  except
    on e: EIcXMLParserError do
    begin
      raise Exception.Create( Format( 'Parse error. File - %s, Line - %d',
        [ FormatFile, e.LineNumber ] ) );
    end;
  end;
end;

procedure TXMLRangeReader.ReadBorders;
var
  Elems: TElementArr;
  Attr: TAttributes;

  ColorItem: TrfColor;
  BorderItem: TrfBorder;
  i: integer;
  Temp: TrfBaseFormatItems;
begin
  Attr := GetAttribsFromStrings( BORDER_ATTRIB_NAMES );

  Temp := FRangeFormat.BorderItems;

  Elems := GetElementValues( FRoot, ELEMENT_BORDER, Attr );
  try
    for i := Low( Elems ) to High( Elems ) do
    begin
      BorderItem := TrfBorder.Create;
      try
        ColorItem               := TrfColor( TrfRangeFormat.GetItemFromArr(
          FRangeFormat.ColorItems, Elems[ i ].AttributValues[ 1 ] ) );

        if ColorItem = nil then
          raise ERangeFormatException.Create( 'Border color not exist.' );

        BorderItem.ItemID       := Elems[ i ].AttributValues[ 0 ];
        BorderItem.LineColor    := ColorItem ;
        BorderItem.LineStyle    := StrToInt( Elems[ i ].AttributValues[ 2 ] );
        BorderItem.Weight       := StrToInt( Elems[ i ].AttributValues[ 3 ] );
      except
        BorderItem.Free;
        raise;
      end;

      TrfRangeFormat.AddBaseItemToArr( BorderItem, Temp );
    end;
  finally
    FreeElementArr( Elems );
  end;

  FRangeFormat.BorderItems := Temp;
end;

procedure TXMLRangeReader.ReadColors;
var
  Elems: TElementArr;
  Attr: TAttributes;

  ColorItem: TrfColor;
  i: integer;
  Temp: TrfBaseFormatItems;
begin

  Attr := GetAttribsFromStrings( COLOR_ATTRIB_NAMES );

  Temp := FRangeFormat.ColorItems;

  Elems := GetElementValues( FRoot, ELEMENT_COLOR, Attr );
  try
    for i := Low( Elems ) to High( Elems ) do
    begin
      ColorItem := TrfColor.Create;
      try
        ColorItem.ItemID        := Elems[ i ].AttributValues[ 0 ];
        ColorItem.ColorAsString := Elems[ i ].AttributValues[ 1 ];
      except
        ColorItem.Free;
        raise;
      end;
      FRangeFormat.AddBaseItemToArr( ColorItem, Temp );
    end;
  finally
    FreeElementArr( Elems );
  end;

  FRangeFormat.ColorItems := Temp;
end;

procedure TXMLRangeReader.ReadFonts;
var
  Elems: TElementArr;
  Attr: TAttributes;

  FontColor: TrfColor;
  FontItem: TrfFont;
  i: integer;
  Temp: TrfBaseFormatItems;
begin
  Attr := GetAttribsFromStrings( FONT_ATTRIB_NAMES );

  Temp := FRangeFormat.FontItems;

  Elems := GetElementValues( FRoot, ELEMENT_FONT, Attr );
  try
    for i := Low( Elems ) to High( Elems ) do
    begin
      FontItem := TrfFont.Create;
      try
        FontColor               := TrfColor( TrfRangeFormat.GetItemFromArr(
          FRangeFormat.ColorItems, Elems[ i ].AttributValues[ 2 ] ) );

        if FontColor = nil then
          raise ERangeFormatException.Create( 'Font color not exist.' );

        FontItem.ItemID    := Elems[ i ].AttributValues[ 0 ];
        FontItem.Name      := Elems[ i ].AttributValues[ 1 ];

        FontItem.Color     := FontColor;
        FontItem.Size      := StrToInt( Elems[ i ].AttributValues[ 3 ] );
        FontItem.Bold      := StrToBool( Elems[ i ].AttributValues[ 4 ] );
        FontItem.Italic    := StrToBool( Elems[ i ].AttributValues[ 5 ] );
      except
        FontItem.Free;
        raise;
      end;

      TrfRangeFormat.AddBaseItemToArr( FontItem, Temp );
    end;
  finally
    FreeElementArr( Elems );
  end;

  FRangeFormat.FontItems := Temp;
end;

function TXMLRangeReader.ReadRangeFormats: TrfRangeFormat;
begin
  Result := nil;

  ParseDoc;
  
  DoReadRangeFormats();

  Result := FRangeFormat;
end;

procedure TXMLRangeReader.ReadRanges;
var
  Elems: TElementArr;
  Attr: TAttributes;

  ReletadItem: TrfBaseFormat;
  RngItem: TrfRange;
  i: integer;

  Temp: TrfBaseFormatItems;

  xmlNodes: TIcNodeList;
  xmlElem: TIcXMLElement;

  pElemVals: PElement;
begin
  Assert( FRoot <> nil );

  Temp          := FRangeFormat.RangeFormats;
  Attr          := GetAttribsFromStrings( [ ATRIBUT_ITEM_ID ] );
  pElemVals     := CreateElement( Attr );
  try
    //Case semsetive to element names
    xmlNodes := FRoot.GetElementsByTagName( ELEMENT_RANGE );
    try
      for i := 1 to xmlNodes.Length do
      begin
        xmlElem := TIcXMLElement( xmlNodes.Item( i, ntElement ) );
        if xmlElem <> nil then
        begin
          RngItem := TrfRange.Create;
          try
            GetElemAttributes( xmlElem, pElemVals );

            RngItem.ItemID := pElemVals.AttributValues[ 0 ];

            GetRangeProperties( RngItem, xmlElem );

          except
            RngItem.Free;
            raise;
          end;

          TrfRangeFormat.AddBaseItemToArr( RngItem, Temp );
        end;
      end;
    finally
      FreeAndNil( xmlNodes );
    end;
  finally
    Dispose( pElemVals );
  end;

  FRangeFormat.RangeFormats := Temp;
end;

end.
