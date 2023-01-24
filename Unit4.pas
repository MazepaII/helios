//КООРДИНАТЫ СОЗДАТЕЛЯ ПРОГРАММЫ
//Электронная почта: Ltlvfpfqbpfzws@yandex.ru

unit Unit4;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, ExtCtrls, ComCtrls,XPMan, jpeg;

type
  TForm4 = class(TForm)
    ProgressBar1: TProgressBar;
    ProgressBar2: TProgressBar;
    Image1: TImage;
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form4: TForm4;

const
PBS_MARQUEE = $08;
PBM_SETMARQUEE = WM_USER + 10;

implementation

{$R *.dfm}


/// ОТКРЫТИЕ загруски
procedure TForm4.FormCreate(Sender: TObject);
var FSpeed: Integer;

begin

FSpeed := 50;
SetWindowLong(ProgressBar2.Handle, GWL_STYLE,
   GetWindowLong(ProgressBar2.Handle, GWL_STYLE) Or PBS_MARQUEE);
{ Включить }
SendMessage(ProgressBar2.Handle, PBM_SETMARQUEE, 1, FSpeed);
///{ Выключить }
///SendMessage(ProgressBar1.Handle, PBM_SETMARQUEE, 0, 0);

end;


end.
