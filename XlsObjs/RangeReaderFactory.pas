unit RangeReaderFactory;

interface

uses
  Classes,
  RngFormatReader,
  MSXmlRngFormatReader {$IFDEF VER130}, XMLRngFormatReader {$ENDIF};

const
  READER_ICXML = 'ICXML';
  READER_MSXML = 'MSXML';

//type
  function CreateReader( AType: string; AFileNanme: string ): TrfRangeReader;


implementation

uses
  SysUtils;

function CreateReader( AType: string; AFileNanme: string ): TrfRangeReader;
var
  MyRez: TrfRangeReader;
begin
  MyRez := nil;

  if ( AType = READER_MSXML ) then
  begin
    //MSXML - DOM
    MyRez := TMSXMLRangeReader.Create( AFileNanme );
  end
  {$IFDEF VER130}
  else if ( AType = READER_ICXML ) then
  begin
    //IcXML component use
    MyRez := TXMLRangeReader.Create( AFileNanme );
  end
  {$ENDIF}
  else
    raise Exception.Create( AType + ' Not supported reader type' );

  Result := MyRez;
end;

end.
