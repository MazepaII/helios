//���������� ��������� ���������
//����������� �����: Ltlvfpfqbpfzws@yandex.ru

unit Unit3;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, jpeg, ExtCtrls, StdCtrls;

type
  TForm3 = class(TForm)
    Image1: TImage;
    Button1: TButton;
    Label1: TLabel;
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form3: TForm3;

implementation

uses Unit4;

{$R *.dfm}



procedure TForm3.Button1Click(Sender: TObject);
begin
Form3.close;
end;


end.
