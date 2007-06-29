unit Xml;

interface

uses
  Classes, MSXML_TLB;

type
  TMyDOMDocument                = TDOMDocument;
  IDOMElement                   = IXMLDOMElement;
  IDOMNodeList                  = IXMLDOMNodeList;
  IDOMNode                      = IXMLDOMNode;
  IDOMParseError                = IXMLDOMParseError;
  IDOMProcessingInstruction     = IXMLDOMProcessingInstruction;
  IDOMAttribute                 = IXMLDOMAttribute;

  TXmlDocument = class
  private
    FOwner: TComponent;
    FIntf: TMyDOMDocument;
    procedure CreateDOM();

  public
    constructor Create( AOwner: TComponent );
    destructor Destroy; override;

    property Owner: TComponent read FOwner write FOwner;
    property DOM: TMyDOMDocument read FIntf;
  end;


implementation

uses
  SysUtils;

{ TXmlDocument }

constructor TXmlDocument.Create(AOwner: TComponent);
begin
  FOwner        := AOwner;
  FIntf         := nil;

  CreateDOM;
end;

procedure TXmlDocument.CreateDOM;
begin
  FIntf := MSXML_TLB.TDOMDocument.Create( FOwner );

  Assert( FIntf <> nil );

  FIntf.Connect;
end;

destructor TXmlDocument.Destroy;
begin
  FIntf.Disconnect;

  FreeAndNil(FIntf);

  inherited;
end;

end.
