unit RangerTest2;

interface

uses
  TestFrameWork,
  RangerTest, RngFormatReader;

type
  TIcXMLFormatReaderTest = class(TMSXMLFormatReaderTest)
  private

  protected
    function CreateXMLReader(): TrfRangeReader; override;
  public

  end;

  function Suite(): ITestSuite;

implementation

uses
  RangeReaderFactory;

function Suite() :ITestSuite;
begin
  Result := TTestSuite.Create( TIcXMLFormatReaderTest );
end;

{ TIcXMLFormatReaderTest }

function TIcXMLFormatReaderTest.CreateXMLReader: TrfRangeReader;
begin
  result := RangeReaderFactory.CreateReader( RangeReaderFactory.READER_ICXML, FXMLFileName );
end;

end.
