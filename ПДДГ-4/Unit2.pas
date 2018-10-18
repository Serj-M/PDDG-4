unit Unit2;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls;

type
  TForm2 = class(TForm)
    ComboBox1: TComboBox;
    Button1: TButton;
    procedure ComboBox1Change(Sender: TObject);
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    choice_combo: string;
    vrem: integer;
    { Public declarations }
  end;

var
  Form2: TForm2;
  EndL: integer;
  LastLength: integer;

implementation

uses Unit1;

{$R *.dfm}

procedure TForm2.ComboBox1Change(Sender: TObject);
begin
  if Form2.ComboBox1.ItemIndex <> -1
    then
      begin
        choice_combo := ComboBox1.Items.Strings[ComboBox1.ItemIndex];
        vrem:=Form2.ComboBox1.ItemIndex+1;
      end;
end;

procedure TForm2.Button1Click(Sender: TObject);
begin
  if Form2.combobox1.ItemIndex = -1
    then
      begin
        ShowMessage('Ниче не выбрано');
        exit;
      end;
  //showmessage('финиш массив = '+ choice_combo+' vrem = '+IntToStr(vrem));
  Form2.Combobox1.Items.Clear;
  Form2.Combobox1.Text:='Выберите из списка или введите вещество';
  Form2.Close;
end;

end.
