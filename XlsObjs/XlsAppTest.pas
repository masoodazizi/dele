unit XlsAppTest;

interface

uses
  TestFrameWork, XlsApp, XlsObjs, Forms, ActiveX;

type
  TXlsAppTest = class(TTestCase)
  private
    FXls: TXlsApplication;
  protected

    procedure Pointers;
    procedure CombineRanges( var Book: TXlsWorkbook; var Ranges: TXlsRanges;
      var RngA1, RngA1C5, RngK5C20, Rng75: TXlsRange; var IRngDir: IXlsRange );
  public
    procedure Setup; override;
    procedure Teardown; override;
  published
    procedure Books;
    procedure Sheets;
    procedure Ranges;
    procedure ConvertFunction;

    //Close all workbooks, and excel because Workbooks.Count = 0;
    procedure CloseAll;
  end;


  function Suite() :ITestSuite;

implementation

uses
  Windows;

const
  SH_NAME          = 'TestSheet';

  RANGE_1          = 'TestRange1';
  RANGE_2          = 'TestRange2';
  RANGE_3          = 'TestRange3';
  RANGE_4          = 'TestRange4';

  RNG_ADRRESS_1    = 'A1';
  RNG_ADRRESS_2    = 'C5';
  RNG_ADRRESS_3    = 'K5';
  RNG_ADRRESS_4    = 'O20';
  RNG_ADRRESS_COL  = 5;
  RNG_ADRRESS_ROW  = 7;

  RANGE_DIR        = 'TestRangeDirect';

  RNG_ADRRESS_DIR1 = 'J1';
  RNG_ADRRESS_DIR2 = 'M5';


function Suite() :ITestSuite;
begin
  Result := TTestSuite.Create( TXlsAppTest );
end;

{ TXlsAppTest }

procedure TXlsAppTest.Books;
var
  Book1: TXlsItemObj;
  Book2: TXlsItemObj;

  procedure AddItems;
  begin
    Book1 := FXls.XlsBooks.AddItem;
    CheckEquals( FXls.XlsBooks.Count, 1 + FXls.XlsBooks.GetOpenedCount );

    Book2 := FXls.XlsBooks.AddItem;
    CheckEquals( FXls.XlsBooks.Count, 2 + FXls.XlsBooks.GetOpenedCount);
  end;
begin
  Pointers;

  AddItems;

  FXls.XlsBooks.Remove( Book2 );
  CheckEquals( FXls.XlsBooks.Count, 1 + FXls.XlsBooks.GetOpenedCount);

  TXlsWorkbook(Book1).Close( False );
  FXls.XlsBooks.Remove( Book1 );
  CheckEquals( FXls.XlsBooks.Count, 0 + FXls.XlsBooks.GetOpenedCount );

  Pointers;
end;

procedure TXlsAppTest.CloseAll;
var
  i: integer;
begin
  Pointers;

  for i := 1 to 10 do
  begin
    FXls.XlsBooks.AddItem;
  end;

  CheckEquals( FXls.XlsBooks.Count, 10 + FXls.XlsBooks.GetOpenedCount );

  //Close all test
  FXls.XlsBooks.CloseAll;
  CheckEquals( FXls.XlsBooks.Count, 0 );
  CheckEquals( FXls.XlsBooks.GetOpenedCount, 0 );

  Pointers;
end;

procedure TXlsAppTest.CombineRanges( var Book: TXlsWorkbook; var Ranges: TXlsRanges; var RngA1, RngA1C5, RngK5C20, Rng75: TXlsRange; var IRngDir: IXlsRange );
var
  Book1: TXlsItemObj;
  Shs: TXlsWorksheets;
  Sh: TXlsWorksheet;
  Count: integer;
  RngNameTemp: string;
begin
  Pointers;

  //New book
  Book1 := FXls.XlsBooks.AddItem;
  CheckEquals( FXls.XlsBooks.Count, 1 + FXls.XlsBooks.GetOpenedCount );

  Book := TXlsWorkbook(Book1);
  //Get sheets
  Shs   := TXlsWorkbook(Book1).XlsWorksheets;
  Count := Shs.Count;
  CheckEquals( Shs.Count, Shs.IWorksheets.Count );

  //Add new sheet
  Sh := TXlsWorksheet(Shs.AddItem);
  CheckEquals( Shs.Count, Count + 1 );
  CheckEquals( Shs.IWorksheets.Count, Count + 1 );
  Sh.IWorksheet.Name := SH_NAME;

  CheckNotNull( Sh.XlsRanges );
  CheckNotNull( Sh.XlsRanges.ContnrDisp );
  Ranges := Sh.XlsRanges;
  CheckNotNull( Ranges );

  RngA1 := Ranges.Add( RANGE_1, RNG_ADRRESS_1, RNG_ADRRESS_1 );
  CheckNotNull( RngA1 );
  CheckNotNull( RngA1.IRange );
  CheckNotNull( RngA1.ItemDisp );
  CheckEquals( string(RngA1.IRange.Name.Name), RANGE_1 );


  RngA1C5 := Ranges.Add( RANGE_2, RNG_ADRRESS_1, RNG_ADRRESS_2 );
  CheckNotNull( RngA1C5 );
  CheckNotNull( RngA1C5.IRange );
  CheckNotNull( RngA1C5.ItemDisp );
  CheckEquals( string(RngA1C5.IRange.Name.Name), RANGE_2 );


  RngK5C20 := Ranges.Add( RANGE_3, RNG_ADRRESS_3, RNG_ADRRESS_4 );
  CheckNotNull( RngK5C20 );
  CheckNotNull( RngK5C20.IRange );
  CheckNotNull( RngK5C20.ItemDisp );
  CheckEquals( string(RngK5C20.IRange.Name.Name), RANGE_3 );


  Rng75 := Ranges.Add( RANGE_4, RNG_ADRRESS_ROW, RNG_ADRRESS_COL );
  CheckNotNull( Rng75 );
  CheckNotNull( Rng75.IRange );
  CheckNotNull( Rng75.ItemDisp );
  CheckEquals( string(Rng75.IRange.Name.Name), RANGE_4 );

  //Direct adding to excel without to the RangeContainer
  IRngDir := Sh.GetSheetRange( RNG_ADRRESS_DIR1, RNG_ADRRESS_DIR2 );

  CheckNotNull( IRngDir );
  IRngDir.Name := RANGE_DIR;
end;

procedure TXlsAppTest.ConvertFunction;
var
  str: string;
begin
  CheckEquals( 'AI1', AbsoluteToRef( 35, 1 ) );
  CheckEquals( 'A10', AbsoluteToRef( 1, 10 ) );
  CheckEquals( 'IV50', AbsoluteToRef( 256, 50 ) );
  CheckEquals( 'EF50', AbsoluteToRef( 136, 50 ) );
end;

procedure TXlsAppTest.Pointers;
begin
  CheckNotNull( FXls, 'IExcelApp' );
  CheckNotNull( IUnknown(FXls.IExcelApp), 'IExcelApp' );

  CheckNotNull( FXls.XlsBooks );
end;

procedure TXlsAppTest.Ranges;
var
  Rngs: TXlsRanges;

  Rng1: TXlsRange;
  Rng2: TXlsRange;
  Rng3: TXlsRange;
  Rng4: TXlsRange;
  RngDir: TXlsRange;

  FndRng: TXlsRange;

  IRngT: IXlsRange;
  Count: integer;

  DummiBook: TXlsWorkbook;

  procedure CompareRngs( ARngOrig, ARngFnd: TXlsRange );
  begin
    CheckNotNull( ARngOrig );
    CheckNotNull( ARngOrig.IRange );
    CheckNotNull( ARngOrig.ItemDisp );

    CheckNotNull( ARngFnd );
    CheckNotNull( ARngFnd.IRange );
    CheckNotNull( ARngFnd.ItemDisp );

    CheckEquals( String(ARngOrig.IRange.Name.Name), String(ARngFnd.IRange.Name.Name) );
    CheckSame( ARngOrig, ARngFnd );
    CheckSame( ARngOrig.IRange, ARngFnd.IRange );
  end;


begin
  Pointers;

  CombineRanges( DummiBook, Rngs, Rng1, Rng2, Rng3, Rng4, IRngT );

  FndRng := Rngs.FindRange( RANGE_1 );
  CompareRngs( Rng1, FndRng );
  CheckEquals( String(FndRng.IRange.Name.Name), RANGE_1 ) ;

  FndRng := Rngs.FindRange( RANGE_2 );
  CompareRngs( Rng2, FndRng );
  CheckEquals( String(FndRng.IRange.Name.Name), RANGE_2 );

  FndRng := Rngs.FindRange( RANGE_3 );
  CompareRngs( Rng3, FndRng );
  CheckEquals( String(FndRng.IRange.Name.Name), RANGE_3 );

  FndRng := Rngs.FindRange( RANGE_4 );
  CompareRngs( Rng4, FndRng );
  CheckEquals( String(FndRng.IRange.Name.Name), RANGE_4 );

  //Direct check
  //first time will add
  FndRng := Rngs.FindRange( RANGE_DIR );
  CheckNotNull( FndRng );
  CheckNotNull( FndRng.IRange );
  CheckNotNull( FndRng.ItemDisp );

  CheckEquals( String(FndRng.IRange.Name.Name), String(IRngT.Name.Name) );

  //return from container
  RngDir := Rngs.FindRange( RANGE_DIR );
  CompareRngs( RngDir, FndRng );

//  rng.Address(False, False, xlA1) - A12:B23
//  rng.Address(True, True, xlA1) - $A$12:$B$23

  Pointers;
end;

procedure TXlsAppTest.Setup;
begin
  inherited;
  FXls := TXlsApplication.Create( nil );
end;

procedure TXlsAppTest.Sheets;
const
  SH_NAME = 'TestRemove';
var
  Book1: TXlsItemObj;
  Sh: TXlsWorksheet;
  Shs: TXlsWorksheets;
  Count: integer;
begin
  Pointers;

  //New book
  Book1 := FXls.XlsBooks.AddItem;
  CheckEquals( FXls.XlsBooks.Count, 1 + FXls.XlsBooks.GetOpenedCount );

  //Get sheets
  Shs   := TXlsWorkbook(Book1).XlsWorksheets;
  Count := Shs.Count;
  CheckEquals( Shs.Count, Shs.IWorksheets.Count);

  //Add new sheet
  Sh := TXlsWorksheet(Shs.AddItem);
  CheckEquals( Shs.Count, Count + 1);
  CheckEquals( Shs.IWorksheets.Count, Count + 1 );

  //Remove sheet
//  Sh.Delete; //remove from excel
  Shs.Remove( Sh );
  CheckEquals( Shs.Count, Count);

  //Add new sheet
  Sh := TXlsWorksheet(Shs.AddItem);
  CheckEquals( Shs.Count, Count + 1);
  Sh.IWorksheet.Name := SH_NAME;

  //Remove sheet by name
//  Sh.Delete; //remove from excel
  Shs.RemoveItem( SH_NAME );
  CheckEquals( Shs.Count, Count);

  Pointers;
end;

procedure TXlsAppTest.Teardown;
begin
  FXls.Free;
  inherited;
end;

initialization
  CoInitialize( nil );

finalization
  CoUnInitialize;

end.
