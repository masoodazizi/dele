unit RangeReaderFactorytest;

interface
uses
  TestFrameWork,
  RngFormatReader;

type
  TRangeReaderFactorytest = class(TTestCase)
  private

  public
  published
    {$IFDEF VER130}
    procedure CreateIcXMLTest;
    {$ENDIF}
    procedure CreateMSXMLTest;
    procedure CreateExceptionTest;
  end;

  function Suite() :ITestSuite;

implementation

uses
  SysUtils,
  RangeReaderFactory;

function Suite() :ITestSuite;
begin
  Result := TTestSuite.Create( TRangeReaderFactorytest );
end;


{ TXmlTest }


{ TRangeReaderFactorytest }

procedure TRangeReaderFactorytest.CreateExceptionTest;
var
  FObj: TrfRangeReader;
begin
  FObj := nil;
  try
    FObj := CreateReader( 'blabla', 'some' );

    Fail( 'Expected Exception Fail' );
  except
    //thats OK;
  end;
  CheckNull( FObj );
end;

{$IFDEF VER130}
procedure TRangeReaderFactorytest.CreateIcXMLTest;
var
  FObj: TrfRangeReader;
begin
  FObj := CreateReader( READER_ICXML, 'some' );
  try
    CheckNotNull( FObj );

  finally
    FreeAndNil( FObj );
  end;

end;
{$ENDIF}

procedure TRangeReaderFactorytest.CreateMSXMLTest;
var
  FObj: TrfRangeReader;
begin
  FObj := CreateReader( READER_MSXML, 'some' );
  try
    CheckNotNull( FObj );
  finally
    FreeAndNil( FObj );
  end;
end;

end.
