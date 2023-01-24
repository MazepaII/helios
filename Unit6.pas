//КООРДИНАТЫ СОЗДАТЕЛЯ ПРОГРАММЫ
//Электронная почта: Ltlvfpfqbpfzws@yandex.ru

unit Unit6;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, jpeg, ExtCtrls;

type
  TForm6 = class(TForm)
    Image1: TImage;
    Vkl: TImage;
    VbIkl: TImage;
    F12_VbIkl: TImage;
    F12_Vkl: TImage;
    F8_VbIkl: TImage;
    F8_Vkl: TImage;
    F6_Vkl: TImage;
    F6_VbIkl: TImage;
    procedure VbIklClick(Sender: TObject);
    procedure VklMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure Image1MouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FormKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form6: TForm6;

implementation

{$R *.dfm}

//Нажетие на кнопку ВЫКЛ        OK
procedure TForm6.VbIklClick(Sender: TObject);
begin
Form6.Close ;
end;

//Появление кнопки ВКЛ
procedure TForm6.Image1MouseMove(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
begin
if VbIkl.Visible=true then
   begin
     VbIkl.Visible:=false ;
     Vkl.Visible:=true ;
   end;
end;

//НАжатие кнопок на форме
procedure TForm6.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
if  Key=Word( VK_F12) then    //F12
  begin
    F12_Vkl.Visible:=false ;
    F12_VbIkl.Visible:=true ;
  end;
if  Key=Word( VK_F8) then    //F8
  begin
    F8_Vkl.Visible:=false ;
    F8_VbIkl.Visible:=true ;
  end;
if  Key=Word( VK_F6) then    //F6
  begin
    F6_Vkl.Visible:=false ;
    F6_VbIkl.Visible:=true ;
  end;
end;

//ОТжатие кнопок на форме
procedure TForm6.FormKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
if  Key=Word( VK_F12) then   //F12
  begin
    F12_VbIkl.Visible:=false ;
    F12_Vkl.Visible:=true ;
  end;
if  Key=Word( VK_F8) then   //F8
  begin
    F8_VbIkl.Visible:=false ;
    F8_Vkl.Visible:=true ;
  end;
if  Key=Word( VK_F6) then   //F6
  begin
    F6_VbIkl.Visible:=false ;
    F6_Vkl.Visible:=true ;
  end;
end;

//Появление кнопки ВЫКЛ      OK
procedure TForm6.VklMouseMove(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
begin
Vkl.Visible:=false ;
VbIkl.Visible:=true ;
end;


end.
 