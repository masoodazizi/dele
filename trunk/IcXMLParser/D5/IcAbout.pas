unit IcAbout;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ExtCtrls;

type
  TIcAboutForm = class(TForm)
    Image1: TImage;
    Label1: TLabel;
    Label2: TLabel;
    Button1: TButton;
    Bevel2: TBevel;
    Memo1: TMemo;
    Bevel1: TBevel;
    Label3: TLabel;
    Label4: TLabel;
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  IcAboutForm: TIcAboutForm;

implementation

{$R *.DFM}

procedure TIcAboutForm.Button1Click(Sender: TObject);
begin
  ModalResult := mrOk;
end;

end.
