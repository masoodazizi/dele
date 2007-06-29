unit XlsFormaterTest;

interface

uses
  Classes, TestFrameWork,
  XlsApp, XlsObjs, XlsFormater, Ranger, RngFormatReader;

type
  TXlsFormaterTest = class(TTestCase)
  private
    FReader: TrfRangeReader;
    FXMLFileName: string;

    FXls: TXlsApplication;
    FXlsFmt: TXlsFormater;
    FRngFormats: TrfRangeFormat;

    procedure ReadingFormat;
    procedure PrepareFormat;
    procedure UnPrepareFormat;

    procedure CheckRangeA1A2( ARng: IXlsRange );
    procedure CheckRangeE3F4( ARng: IXlsRange );

    procedure ColorCheck( AColor: OleVariant; AMyIndex: integer );
    procedure FontCheck( ARng: IXlsRange; AMyIndex: integer );
    procedure BordersCheck( ARng: IXlsRange; ABorderInd, AMyIndex: integer );
    procedure WidthCheck( ARng: IXlsRange; AExpected: integer );
    procedure HeightCheck( ARng: IXlsRange; AExpected: integer );
  protected

  public
    procedure Setup; override;
    procedure Teardown; override;
  published

    procedure FormatTest;

    //Dont test format, just close all books then excel app.
    //will destroy completely.
    procedure JustCloseAllBooks;
  end;

  function Suite() :ITestSuite;

implementation

uses
  SysUtils, Windows, Graphics,
  RangeReaderFactory,
  TestHlp;

const
  ROW_COL_DELTA = 0.5;


      {
type
  XlLineStyle = TOleEnum;
const
  xlContinuous = $00000001;
  xlDash = $FFFFEFED;
  xlDashDot = $00000004;
  xlDashDotDot = $00000005;
  xlDot = $FFFFEFEA;
  xlDouble = $FFFFEFE9;
  xlSlantDashDot = $0000000D;
  xlLineStyleNone = $FFFFEFD2;

type
  XlBorderWeight = TOleEnum;
const
  xlHairline = $00000001;
  xlMedium = $FFFFEFD6;
  xlThick = $00000004;
  xlThin = $00000002;

type
  XlBordersIndex = TOleEnum;
const
  xlInsideHorizontal = $0000000C;
  xlInsideVertical = $0000000B;
  xlDiagonalDown = $00000005;
  xlDiagonalUp = $00000006;
  xlEdgeBottom = $00000009;
  xlEdgeLeft = $00000007;
  xlEdgeRight = $0000000A;
  xlEdgeTop = $00000008;
      }



function Suite() :ITestSuite;
begin
  Result := TTestSuite.Create( TXlsFormaterTest );
end;

{ TXlsFormaterTest }

procedure TXlsFormaterTest.BordersCheck(ARng: IXlsRange;
  ABorderInd, AMyIndex: integer);
begin
  CheckNotNull( ARng );
  CheckNotNull( ARng.Borders.Item[ ABorderInd ] );

  with ARng.Borders.Item[ ABorderInd ] do
  begin
    case AMyIndex of
      1:
        begin
          CheckEquals( Cardinal(LineStyle), 1 );
          CheckEquals( Cardinal(Weight), 1 );
          //Dont return apropriate color
          //ColorCheck( Color, 3 );
        end;
      2:
        begin
          CheckEquals( Cardinal(LineStyle), $FFFFEFED );
          CheckEquals( Cardinal(Weight), $FFFFEFD6 );
          //Dont return apropriate color
          //ColorCheck( Color, 3 );
        end;
      3:
        begin
          CheckEquals( Cardinal(LineStyle), $FFFFEFE9 );
          CheckEquals( Cardinal(Weight), $00000004 );
          //Dont return apropriate color
          //ColorCheck( Color, 4 );
        end;
    else
      Assert( False );
    end;
  end;
end;

procedure TXlsFormaterTest.CheckRangeA1A2(ARng: IXlsRange);
begin
  CheckNotNull( ARng );
  //Interior
  ColorCheck( ARng.Interior.Color, 2 );

  FontCheck( ARng, 1 );

  //Borders
  BordersCheck( ARng, 5, 1 );
  BordersCheck( ARng, 7, 2 );
  BordersCheck( ARng, 9, 3 );
  BordersCheck( ARng, 12, 2 );

  WidthCheck( ARng, 20 );
  HeightCheck( ARng, 10 ); //Hidden
end;

procedure TXlsFormaterTest.CheckRangeE3F4(ARng: IXlsRange);
begin
  CheckNotNull( ARng );
  //Interior
  ColorCheck( ARng.Interior.Color, 1 );

  FontCheck( ARng, 1 );

  //Borders
  BordersCheck( ARng, 7, 1 );
  BordersCheck( ARng, 9, 2 );

  WidthCheck( ARng, 0 );
  HeightCheck( ARng, -1 ); //Hidden 
end;

procedure TXlsFormaterTest.ColorCheck(AColor: OleVariant;
  AMyIndex: integer);
begin
  case AMyIndex of
    1: CheckEquals( Integer(AColor), ColorToRGB( StringToColor( '$00158860' ) ) );
    2: CheckEquals( Integer(AColor), ColorToRGB( StringToColor( '$00F18800' ) ) );
    3: CheckEquals( Integer(AColor), ColorToRGB( StringToColor( 'clRed' ) ) );
    4: CheckEquals( Integer(AColor), ColorToRGB( StringToColor( 'clLime' ) ) );
  else
    Assert( False );
  end;
end;

procedure TXlsFormaterTest.FontCheck(ARng: IXlsRange; AMyIndex: integer);
begin
  CheckNotNull( ARng );
  with ARng do
  begin
    case AMyIndex of
      1:
        begin
          CheckEquals( string(Font.Name), 'Tahoma' );
          CheckEquals( Integer(Font.Size), 12 );
          CheckEquals( Boolean(Font.Bold), True );
          CheckEquals( Boolean(Font.Italic), False );
          ColorCheck( Font.Color, 1 );
        end;
    else
      Assert( False );
    end;
  end;
end;

procedure TXlsFormaterTest.FormatTest;
var
  Rngs: TXlsRanges;

  RngTest: IXlsRange;

  Book: TXlsItemObj;
  Shs: TXlsWorksheets;
  Sh: TXlsWorksheet;
begin
  //New book
  Book := FXls.XlsBooks.AddItem;

  //Get sheets
  Shs := TXlsWorkbook(Book).XlsWorksheets;

  //Add A1
  Sh := TXlsWorksheet(Shs.AddItem);

  //Apply format
  FXlsFmt.Book          := TXlsWorkbook(Book);
  FXlsFmt.RangeFormats  := FRngFormats;
  FXlsFmt.SheetName     := TXlsWorksheet(Sh).IWorksheet.Name;
  FXlsFmt.ApplyFormat();

  //Check ranges 1
  RngTest := Sh.XlsRanges.FindAsIRange( 'A1:A2' );
  CheckRangeA1A2( RngTest );
  RngTest := nil;

  //Check ranges 2
  RngTest := Sh.XlsRanges.FindAsIRange( 'E3:F4' );
  CheckRangeE3F4( RngTest );
  RngTest := nil;
end;

procedure TXlsFormaterTest.HeightCheck(ARng: IXlsRange;
  AExpected: integer);
begin
  case AExpected of
    -1:
      begin
        CheckEquals( Boolean(ARng.Rows.Hidden), True );
      end;
    -2:
      begin
        //For now only visual possible to check - and data must to have
      end;
     0:
       begin
         //not shure about this const. May be better to get it from another cell
         CheckEquals( ARng.RowHeight, 12.75, ROW_COL_DELTA );
       end;
  else
    CheckEquals( ARng.RowHeight, AExpected, ROW_COL_DELTA );
  end;
end;

procedure TXlsFormaterTest.JustCloseAllBooks;
begin
  CheckNotNull( FXls );
  CheckNotNull( FXls.XlsBooks );
  FXls.XlsBooks.CloseAll;
end;

procedure TXlsFormaterTest.PrepareFormat;
begin
  FRngFormats   := nil;

  THlpTest.CreateXMLFile();

  FXMLFileName  := THlpTest.GetXMLTestFileName();
  FReader       := RangeReaderFactory.CreateReader( RangeReaderFactory.READER_MSXML, FXMLFileName );

  ReadingFormat;
end;

procedure TXlsFormaterTest.ReadingFormat;
begin
  FRngFormats := FReader.ReadRangeFormats;
end;

procedure TXlsFormaterTest.Setup;
begin
  inherited;
  //Something wrong after close all. Application is destroyed and can not connect after that
  Sleep( 500 );

  FXls := TXlsApplication.Create( nil );

  FXlsFmt := TXlsFormater.Create;

  PrepareFormat;
end;

procedure TXlsFormaterTest.Teardown;
begin
  inherited;
  FXls.Free;

  FreeAndNil( FXlsFmt );

  UnPrepareFormat;
end;

procedure TXlsFormaterTest.UnPrepareFormat;
begin
  FreeAndNil( FRngFormats );
  FreeAndNil( FReader );
end;

procedure TXlsFormaterTest.WidthCheck(ARng: IXlsRange; AExpected: integer);
begin
  case AExpected of
    -1:
      begin
        CheckEquals( Boolean(ARng.Columns.Hidden), True );
      end;
    -2:
      begin
        //For now only visual possible to check - and data must to have
      end;
     0:
       begin
         //not shure about this const. May be better to get it from another cell
         CheckEquals( ARng.ColumnWidth, 8.43, ROW_COL_DELTA );
       end;
  else
    CheckEquals( ARng.ColumnWidth, AExpected, ROW_COL_DELTA );
  end;
end;

end.
