unit RangerTest;

interface

uses
  TestFrameWork, Windows, SysUtils, Forms, Classes,
  Ranger, RngFormatReader;

type
  TMSXMLFormatReaderTest = class(TTestCase)
  private
    FReader: TrfRangeReader;
    FRngFormats: TrfRangeFormat;

    procedure ReadingFormat;
    procedure CheckObject;
  protected
    FXMLFileName: string;
    function CreateXMLReader(): TrfRangeReader; virtual;
  public
    procedure PrepareData;
    procedure UnPrepareData;

    procedure Setup; override;
    procedure Teardown; override;
    function GetRngFormats: TrfRangeFormat;
  published
    procedure CheckColors;
    procedure CheckBorders;
    procedure CheckFonts;
    procedure CheckRanges;
  end;

  function Suite() :ITestSuite;

implementation

uses
  Excel97,
  RangeReaderFactory,
  ActiveX,
  TestHlp;

function Suite() :ITestSuite;
begin
  Result := TTestSuite.Create( TMSXMLFormatReaderTest );
end;


{ TMSXMLFormatReaderTest }

procedure TMSXMLFormatReaderTest.CheckBorders;
begin
  CheckObject;

  CheckEquals( Length( FRngFormats.BorderItems ), 3 );

  Check( FRngFormats.BorderItems[ 0 ] is TrfBorder, '0 Is TrfBorder' );

  CheckEquals( FRngFormats.BorderItems[ 0 ].ItemID, 'Border1' );
  CheckEquals( TrfBorder(FRngFormats.BorderItems[ 0 ]).LineColor.Itemid, 'Color3' );
  CheckEquals( TrfBorder(FRngFormats.BorderItems[ 0 ]).LineStyle, 1 );
  CheckEquals( TrfBorder(FRngFormats.BorderItems[ 0 ]).Weight, 1 );

  CheckEquals( FRngFormats.BorderItems[ 1 ].ItemID, 'Border2' );
  CheckEquals( TrfBorder(FRngFormats.BorderItems[ 1 ]).LineColor.Itemid, 'Color3' );
  CheckEquals( TrfBorder(FRngFormats.BorderItems[ 1 ]).LineStyle, TOleEnum($FFFFEFED) );
  CheckEquals( TrfBorder(FRngFormats.BorderItems[ 1 ]).Weight, TOleEnum($FFFFEFD6) );

  CheckEquals( FRngFormats.BorderItems[ 2 ].ItemID, 'Border3' );
  CheckEquals( TrfBorder(FRngFormats.BorderItems[ 2 ]).LineColor.Itemid, 'Color4' );
  CheckEquals( TrfBorder(FRngFormats.BorderItems[ 2 ]).LineStyle, TOleEnum($FFFFEFE9) );
  CheckEquals( TrfBorder(FRngFormats.BorderItems[ 2 ]).Weight, TOleEnum($00000004) );
end;

procedure TMSXMLFormatReaderTest.CheckColors;
begin
  CheckObject;

  CheckEquals( Length( FRngFormats.ColorItems ), 4 );

  Check( FRngFormats.ColorItems[ 0 ] is TrfColor, '0 Is TrfColor' );

  CheckEquals( FRngFormats.ColorItems[ 0 ].ItemID, 'Color1' );
  CheckEquals( TrfColor(FRngFormats.ColorItems[ 0 ]).ColorAsString, '$00158860' );

  CheckEquals( FRngFormats.ColorItems[ 1 ].ItemID, 'Color2' );
  CheckEquals( TrfColor(FRngFormats.ColorItems[ 1 ]).ColorAsString, '$00F18800' );

  CheckEquals( FRngFormats.ColorItems[ 2 ].ItemID, 'Color3' );
  CheckEquals( TrfColor(FRngFormats.ColorItems[ 2 ]).ColorAsString, 'clRed' );

  CheckEquals( FRngFormats.ColorItems[ 3 ].ItemID, 'Color4' );
  CheckEquals( TrfColor(FRngFormats.ColorItems[ 3 ]).ColorAsString, 'clLime' );
end;

procedure TMSXMLFormatReaderTest.CheckFonts;
begin
  CheckObject;

  CheckEquals( Length( FRngFormats.FontItems ), 1 );

  Check( FRngFormats.FontItems[ 0 ] is TrfFont, '0 Is TrfBorder' );

  with TrfFont(FRngFormats.FontItems[ 0 ]) do
  begin
    CheckEquals( ItemID, 'Font1' );
    CheckEquals( Name, 'Tahoma' );
    CheckEquals( Size, 12 );
    CheckEquals( Color.ItemID, 'Color1' );
    CheckEquals( Bold, True );
    CheckEquals( Italic, False );
  end;
end;

procedure TMSXMLFormatReaderTest.CheckObject;
begin
  CheckNotNull( FRngFormats );
end;

procedure TMSXMLFormatReaderTest.CheckRanges;
begin
  CheckObject;

  CheckEquals( Length( FRngFormats.RangeFormats ), 2 );

  Check( FRngFormats.RangeFormats[ 0 ] is TrfRange, '0 Is TrfRange' );
  Check( FRngFormats.RangeFormatItem[ 'Range1' ] is TrfRange, '0 Is TrfRange' );

  with TrfRange(FRngFormats.RangeFormats[ 0 ]) do
  begin

    CheckEquals( RangeName, 'E3:F4' );

    CheckEquals( InteriorColor.ItemId , 'Color1' );
    CheckEquals( TextFont.ItemId , 'Font1' );

    CheckEquals( BorderItem[7].ItemId , 'Border1' );
    CheckEquals( BorderItem[9].ItemId , 'Border2' );

    CheckEquals( RowHeight , -1 );
    CheckEquals( ColWidth , 0 );
  end;

  Check( FRngFormats.RangeFormats[ 1 ] is TrfRange, '0 Is TrfRange' );
  Check( FRngFormats.RangeFormatItem[ 'Range2' ] is TrfRange, '0 Is TrfRange' );

  with TrfRange(FRngFormats.RangeFormats[ 1 ]) do
  begin

    CheckEquals( RangeName, 'A1:A2' );

    CheckEquals( InteriorColor.ItemId , 'Color2' );
    CheckEquals( TextFont.ItemId , 'Font1' );

    CheckNotNull( BorderItem[xlInsideHorizontal] );
    CheckNull( BorderItem[xlEdgeRight] );

    CheckEquals( BorderItem[5].ItemId , 'Border1' );
    CheckEquals( BorderItem[7].ItemId , 'Border2' );
    CheckEquals( BorderItem[9].ItemId , 'Border3' );
    CheckEquals( BorderItem[12].ItemId , 'Border2' );

    CheckEquals( RowHeight , 10 );
    CheckEquals( ColWidth , 20 );
  end;
end;

function TMSXMLFormatReaderTest.CreateXMLReader: TrfRangeReader;
begin
  Result := RangeReaderFactory.CreateReader( RangeReaderFactory.READER_MSXML, FXMLFileName );
end;

function TMSXMLFormatReaderTest.GetRngFormats: TrfRangeFormat;
begin
  Result := FRngFormats;
end;

procedure TMSXMLFormatReaderTest.PrepareData;
begin
  FRngFormats   := nil;

  THlpTest.CreateXMLFile();

  FXMLFileName  := THlpTest.GetXMLTestFileName();
  FReader       := CreateXMLReader();

  ReadingFormat;
end;

procedure TMSXMLFormatReaderTest.ReadingFormat;
begin
  FRngFormats := FReader.ReadRangeFormats;
end;

procedure TMSXMLFormatReaderTest.Setup;
begin
  inherited;

  PrepareData;
end;

procedure TMSXMLFormatReaderTest.Teardown;
begin
  UnPrepareData;
  inherited;
end;

procedure TMSXMLFormatReaderTest.UnPrepareData;
begin
  FreeAndNil( FRngFormats );
  FreeAndNil( FReader );
end;

end.
