unit ExcelTests;

interface

uses
  TestFrameWork;

  function GetExcelTestSuite: ITestSuite;

implementation

uses
  RangerTest,
  RangerTest2,
  XlsAppTest,
  XlsFormaterTest,
  RangeReaderFactorytest,
  XmlTest;

function GetExcelTestSuite: ITestSuite;
var
  test: TTestSuite;
begin
  Result := TTestSuite.Create('Excel.Range.XML.Tests');

  Result.AddSuite( XlsAppTest.Suite );
  Result.AddSuite( XmlTest.Suite );

  test := TTestSuite.Create( 'Range' );
  test.AddSuite( RangeReaderFactorytest.Suite );
  test.AddSuite( RangerTest.Suite );
  {$IFDEF VER130}
  test.AddSuite( RangerTest2.Suite );
  {$ENDIF}
  test.AddSuite( XlsFormaterTest.Suite );
  Result.AddSuite( test );
end;


end.
