unit Unit3;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls;

type
  TForm3 = class(TForm)
    Button1: TButton;
    Button2: TButton;
    Label1: TLabel;
    Button3: TButton;
    Label2: TLabel;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
    Yes_All: integer;
  public

    variant_vibora : integer;
    { Public declarations }
  end;

var
  Form3: TForm3;

implementation

uses ComObj, ExcelXP, Unit1, Unit2;

{$R *.dfm}

procedure TForm3.Button1Click(Sender: TObject);
begin
  variant_vibora:=1;
  Yes_all:=1;
  Form3.Close;
end;

procedure TForm3.Button2Click(Sender: TObject);
begin
  variant_vibora:=1;
  Form3.Close;
  Yes_all:=2;
end;

procedure TForm3.Button3Click(Sender: TObject);
begin
  variant_vibora:=3;
  Yes_all:=3;
  Form3.Close;
end;

procedure TForm3.FormCreate(Sender: TObject);
begin
  If Yes_all=2
    then
      begin
        variant_vibora:=1;
        Form3.Close;
      end;
end;

end.
