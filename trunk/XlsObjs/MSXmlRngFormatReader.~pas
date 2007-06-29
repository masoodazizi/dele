unit MSXmlRngFormatReader;

interface

uses
  Classes, Ranger, RngFormatReader, Xml;

type

  TMSXMLRangeReader = class(TrfRangeReader)
  private
    FXMLDoc: TXmlDocument;
    FRoot: IDOMElement;
    FFileName: string;

    function GetElementValues( AElem: IDOMNode; AElementName: string;
      ANamesOfAtrrib: TAttributes; AInWholeDoc: boolean ): TElementArr;
    procedure GetElemValue( AElem: IDOMNode; ElemValue: PElement );
    procedure GetElemAttributes( AElem: IDOMNode; ElemValue: PElement );
    function GetElemText( AElem: IDOMNode ): string;


    procedure GetRangeProperties( ARng: TrfRange; AElem: IDOMNode );
    procedure GetRangeBorders( ARng: TrfRange; AElem: IDOMNode );

    procedure LoadXMLFile( AFileName: string );
    procedure ParseDoc;
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

{ TMSXMLRangeReader }


constructor TMSXMLRangeReader.Create(AFileName: string);
begin
  inherited;
  FXMLDoc       := TXmlDocument.Create( nil );
  FRoot         := nil;
  FFileName     := AFileName;
end;


destructor TMSXMLRangeReader.Destroy;
begin
  //Created from parser and destroied.
  if Assigned( FXMLDoc ) then
    FreeAndNil( FXMLDoc );

  inherited;
end;


procedure TMSXMLRangeReader.GetElemAttributes(AElem: IDOMNode;
  ElemValue: PElement);
var
  i: integer;
  AttrVal: string;
  nodeAtrib: IDOMNode;
begin
  Assert( ElemValue <> nil );
  Assert( AElem <> nil );

  for i := Low( ElemValue.AttributeNames ) to High( ElemValue.AttributeNames ) do
  begin
    nodeAtrib := AElem.attributes.getNamedItem( ElemValue.AttributeNames[ i ] );

    Assert( nodeAtrib <> nil );

    ElemValue.AttributValues[ i ] := nodeAtrib.nodeValue;
  end;
end;

function TMSXMLRangeReader.GetElementValues(AElem: IDOMNode;
  AElementName: string; ANamesOfAtrrib: TAttributes; AInWholeDoc: boolean): TElementArr;
var
  xmlNodes: IDOMNodeList;
  xmlElem: IDOMNode;
  i: integer;

  pElemVals: PElement;
  FindPatern: string;
begin
  Assert( AElem <> nil );

  //Case semsetive to element names
  FindPatern := './';
  if AInWholeDoc then
    FindPatern := '//';
  FindPatern := FindPatern + AElementName;

  xmlNodes := AElem.selectNodes( FindPatern );
  try
    for i := 0 to xmlNodes.length - 1 do
    begin
      xmlElem := xmlNodes.Item[ i ];
      if xmlElem <> nil then
      begin
        pElemVals := CreateElement( ANamesOfAtrrib );

        GetElemValue( xmlElem, pElemVals );
        GetElemAttributes( xmlElem, pElemVals );

        AddElemtoArr( pElemVals, Result );
      end;
    end;
  finally
    xmlNodes := nil;
  end;
end;

function TMSXMLRangeReader.GetElemText(AElem: IDOMNode): string;
begin
  Assert( AElem <> nil );
  Result := AElem.text;
end;

procedure TMSXMLRangeReader.GetElemValue(AElem: IDOMNode;
  ElemValue: PElement);
begin
  Assert( ElemValue <> nil );
  Assert( AElem <> nil );

  ElemValue.ElementValue := GetElemText( AElem );
end;

procedure TMSXMLRangeReader.GetRangeBorders(ARng: TrfRange;
  AElem: IDOMNode);
var
  Elems: TElementArr;
  Attr: TAttributes;
  i: integer;
  Ind: integer;
begin
  Assert( ARng <> nil );
  Assert( AElem <> nil );

  Attr := GetAttribsFromStrings( [ ATRIBUTE_ITEM ] );
  Elems := GetElementValues( AElem, ELEMENT_RANGEBORDER, Attr, false );
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

procedure TMSXMLRangeReader.GetRangeProperties(ARng: TrfRange;
  AElem: IDOMNode);
var
  ElemName: string;
  ElemVal: string;
  xmlElem: IDOMNode;
begin
  Assert( ARng <> nil );
  Assert( AElem <> nil );

  xmlElem := AElem.firstChild;
  repeat
    ElemName := xmlElem.nodeName;
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
      GetRangeBorders( ARng, xmlElem )
    else if ElemName = ELEMENT_RANGEROWH then
      ARng.RowHeight := GetRowColDimension( ElemVal )
    else if ElemName = ELEMENT_RANGECOLW then
      ARng.ColWidth := GetRowColDimension( ElemVal );

      //just excape and go to next -->> //    raise ERangeFormatException.Create('Not Expected element - ' + ElemName );
    xmlElem := xmlElem.nextSibling;
  until ( xmlElem = nil );

end;

procedure TMSXMLRangeReader.ParseDoc;
begin
  try
    LoadXMLFile( FFileName );

    Assert( FXMLDoc <> nil );

    FRoot := FXMLDoc.DOM.documentElement;

    Assert( FRoot <> nil );
  except
    raise;
  end;
end;

procedure TMSXMLRangeReader.ReadBorders;
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

  Elems := GetElementValues( FRoot, ELEMENT_BORDER, Attr, true );
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

procedure TMSXMLRangeReader.ReadColors;
var
  Elems: TElementArr;
  Attr: TAttributes;

  ColorItem: TrfColor;
  i: integer;
  Temp: TrfBaseFormatItems;
begin

  Attr := GetAttribsFromStrings( COLOR_ATTRIB_NAMES );

  Temp := FRangeFormat.ColorItems;

  Elems := GetElementValues( FRoot, ELEMENT_COLOR, Attr, true );
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

procedure TMSXMLRangeReader.ReadFonts;
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

  Elems := GetElementValues( FRoot, ELEMENT_FONT, Attr, true );
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

function TMSXMLRangeReader.ReadRangeFormats: TrfRangeFormat;
begin
  Result := nil;

  ParseDoc;

  DoReadRangeFormats();

  Result := FRangeFormat;
end;

procedure TMSXMLRangeReader.ReadRanges;
var
  Elems: TElementArr;
  Attr: TAttributes;

  ReletadItem: TrfBaseFormat;
  RngItem: TrfRange;
  i: integer;

  Temp: TrfBaseFormatItems;

  xmlNodes: IDOMNodeList;
  xmlElem: IDOMNode;

  pElemVals: PElement;
begin
  Assert( FRoot <> nil );

  Temp          := FRangeFormat.RangeFormats;
  Attr          := GetAttribsFromStrings( [ ATRIBUT_ITEM_ID ] );
  pElemVals     := CreateElement( Attr );
  try
    //Case semsetive to element names
    xmlNodes := FRoot.selectNodes( '//' + ELEMENT_RANGE );
    try
      for i := 0 to xmlNodes.length - 1 do
      begin
        xmlElem := xmlNodes.Item[ i ];
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
      xmlNodes := nil;
    end;
  finally
    Dispose( pElemVals );
  end;

  FRangeFormat.RangeFormats := Temp;
end;



procedure TMSXMLRangeReader.LoadXMLFile(AFileName: string);
var
  err: IDOMParseError;
begin
  try
    with FXMLDoc.DOM do
    begin
      async             := false;
      validateOnParse   := false;
      resolveExternals  := false;

      load( AFileName );

      err := parseError;
      if err.errorCode <> 0 then
      begin
        raise Exception.Create( Format(
        'Parse error. Reason - %s, Line - %d, URL - %s, Src - %s, LinePos - %d, FilePos - %d',
          [
          err.reason,
          err.line,
          err.url,
          err.srcText,
          err.linepos,
          err.filepos
          ] ) );
      end;
    end;
  except
    raise;
  end;
end;

end.
