unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, StrUtils;

type
  TForm1 = class(TForm)
    Button1: TButton;
    Memo1: TMemo;
    Button2: TButton;
    Button3: TButton;
    Memo3: TMemo;
    Memo2: TMemo;
    Button4: TButton;
    Edit1: TEdit;
    Label1: TLabel;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Edit1Change(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;


implementation

{$R *.dfm}

uses ComObj, ExcelXP, Unit2, Unit3;
 var
 E, W :variant;
 StFile : AnsiString;
 Num_Pribor, vrem_edit: string;
 //------------���������� ���� 1 ------------
 Range:Variant;

 ind_arr_gpf_tabl_1 , ind_arr_gpf_tabl_2 :integer;  //������� ������� ������� ������� ������� ��� ������� 1 � 2

 Arr_gpf_tabl_1: array of string;  //������ ������� ������ ������� ��� ���������� �� .gpf
 Arr_gpf_tabl_2: array of string;  //������ ������� ������ ��� �������� �� .gpf, ����� �� �������� ������� ������ ������� � ��������� � �������� �� ��.
 Arr_gpf: array of string;  //������ ������� ������ ��� �������� �� .gpf

 P3orAB: array of string;  // ������ ��� ������������ ���� 2. ������� ������� ����� �.�. ���(�) �.�.

 indx_finish_arr , kol_vesh_gpf_for_tabl_1 , kol_vesh_gpf_for_tabl_2, clik_button : integer;

 LIPA: boolean;


function Parameters: string;    //����� ���������� �� ��������� �������
var                             // ParamStr(0) - ���� �� project.exe; ParamStr(1) - ����� �������;  ParamStr(2) - ���� �� .gpf �����
  cmd : string;
  i : Integer;

begin
  cmd := CmdLine;
  //ShowMessage(cmd);
  //ShowMessage(IntToStr(ParamCount+1)+' ����������');
  for i := 0 to ParamCount-1 do
    if i = 0
      then
        Form1.Memo1.Lines.Add ('����� ������� = '+ParamStr(i+1)+';  ( ����� ������� ������ ������ � ����� ���� �� �����  �������)')
      else
        Form1.Memo1.Lines.Add ('���� �� �����  ������� = '+ParamStr(i+1)+'; ');


  If ParamStr(1)<>''
    then
      begin
        Form1.Edit1.Text:=ParamStr(1);
        vrem_edit:= ParamStr(1);
      end

end;

Function CreateWord:boolean;
begin
CreateWord:=true;
try
W:=CreateOleObject('Word.Application');
except
CreateWord:=false;
end;
End;


Function VisibleWord(visible:boolean):boolean;
begin
VisibleWord:=true;
try
W.visible:= visible;
except
VisibleWord:=false;
end;
End;


Function SaveDocAs(file_:string):boolean;
begin
SaveDocAs:=true;
try
W.ActiveDocument.SaveAs(file_);
except
SaveDocAs:=false;
end;
End;


Function CloseDoc:boolean;
begin
CloseDoc:=true;
try
W.ActiveDocument.Close;
except
CloseDoc:=false;
end;
End;

Function CloseWord:boolean;
begin
CloseWord:=true;
try
W.Quit;
except
CloseWord:=false;
end;
End;


Function OpenDoc(file_:string):boolean;
 Var Doc_:variant;
begin
OpenDoc:=true;
try
Doc_:=W.Documents;
Doc_.Open(file_);
except
OpenDoc:=false;
end;
End;

Function FindTextDoc(text_:string):boolean;
begin
FindTextDoc:=true;
Try
W.Selection.Find.Forward:=true;
W.Selection.Find.Text:=text_;
FindTextDoc := W.Selection.Find.Execute;
except
FindTextDoc:=false;
end;
End;

Function FindAndPasteTextDoc(findtext_,pastetext_:string):boolean;
begin
FindAndPasteTextDoc:=true;
try
W.Selection.Find.Forward:=true;
W.Selection.Find.Text:= findtext_;
if W.Selection.Find.Execute then begin
W.Selection.Delete;
W.Selection.InsertAfter (pastetext_);
end else FindAndPasteTextDoc:=false;
except
FindAndPasteTextDoc:=false;
end;
End;

Function StartOfDoc:boolean;
begin
StartOfDoc:=true;
try
W.Selection.End:=0;
W.Selection.Start:=0;
except
StartOfDoc:=false;
end;
End;

Function GetColumnRowTable(table_:integer; var Column,Row:integer):boolean;
 const
   wdStartOfRangeColumnNumber=16;
   wdStartOfRangeRowNumber=13;
begin
GetColumnRowTable:=true;
try
Column:=W.Selection.Information[wdStartOfRangeColumnNumber];
Row:=W.Selection.Information[wdStartOfRangeRowNumber];
except
GetColumnRowTable:=false;
end;
End;


Function GetSelectionTable:boolean;
  const wdWithInTable=12;
begin
try
GetSelectionTable :=W.Selection.Information[wdWithInTable];
except
GetSelectionTable :=false;
end;
End;

Function SetTextToTable(Table:integer; Row, Column:integer; text:string):boolean;
begin
 SetTextToTable:=true;
 try
  W.ActiveDocument.Tables.Item(Table).Columns.Item(Column).Cells.Item(Row).Range.Text:=text;
 except
  SetTextToTable:=false;
 end;
End;

Function InsertRowTableDoc(table_,position_:integer):boolean;
 var row_:variant;
begin
 InsertRowTableDoc:=true;
 try
 row_:=W.ActiveDocument.Tables.Item(table_).Rows.Item(position_);
 W.ActiveDocument.Tables.Item(table_).Rows.Add(row_);
 except
 InsertRowTableDoc:=false;
 end;
End;



//-------------Excel-------------------

Function CreateExcel:boolean;
begin
CreateExcel:=true;
try
E:=CreateOleObject('Excel.Application');
except
CreateExcel:=false;
end;
End;

Function VisibleExcel(visible:boolean):boolean;
begin
VisibleExcel:=true;
try
E.visible:=visible;
except
VisibleExcel:=false;
end;
End;

Function AddWorkBook:boolean;
begin
 AddWorkBook:=true;
 try
  E.Workbooks.Add;
 except
  AddWorkBook:=false;
 end;
End;

Function OpenWorkBook(file_: string):boolean;
begin
 OpenWorkBook:=true;
 try
  E.Workbooks.Open(file_);
 except
  OpenWorkBook:=false;
 end;
End;

Function SaveWorkBookAs(file_:string): boolean;
begin
SaveWorkBookAs:=true;
try
E.DisplayAlerts:=False;
E.ActiveWorkbook.SaveAs(file_);
E.DisplayAlerts:=True;
except
SaveWorkBookAs:=false;
end;
End;

Function CloseWorkBook:boolean;
begin
 CloseWorkBook:=true;
 try
  E.ActiveWorkbook.Close;
 except
  CloseWorkBook:=false;
 end;
End;

Function CloseExcel:boolean;
begin
 CloseExcel:=true;
 try
  E.Quit;
 except
  CloseExcel:=false;
 end;
End;

Function FindText (text_:string):boolean;
begin
 FindText:=true;
 try
  E.Cells.Find(what:=text_, matchcase:=True).Select;
 except
  FindText:=False;
 end;
End;


procedure TForm1.Button1Click(Sender: TObject);
 Var
 //---------���������� ���� 1------------

  first_row_found, following_row_found, kol_naidenogo, indx_arr, vrem, i, j: integer;
  str1, FirstAddress, Addr, Addr2: string;

  Finish_arr_tabl_1: array of string;   //��������� ������ ������� ������ ������ � ��������� ���������
  Finish_arr_tabl_1_H: array of string; //��������� ������ ������� ������ ������ ������� � ��� ���������� ��������

  str_of_tabl_1: array of string;
  str_of_tabl_1_H: array of string;

  data_tabl1: array of string;

  //---------���������� ���� 2------------
  str_of_tabl_2: array of string;
  str_of_tabl_2_G: array of string;

  Finish_arr_tabl_2: array of string;   //��������� ������ ������� ������ ������ � ��������� ���������
  Finish_arr_tabl_2_G: array of string; //��������� ������ ������� ������ ������ ������� G ��� ���������� ��������

  For_Tabl_2: array of string;

  i2 , j2 , ch_vesh, key_message_tabl_2 : integer;

 begin
  //If Edit1.Text='������� �������� '
  // then
  //   begin
  //     Showmessage('������� � �������');
  //     exit;
  //   end;
  If not(clik_button=1)
   then
     begin
       Showmessage('������� �������� � ������');
       exit;
     end;
 //ComboBox1.Visible := False;  // �������� ComboBox1
 if not CreateExcel
   then
     exit;
 //messagebox(handle,'','��������� Excel.',0);
 VisibleExcel(true);
 //messagebox(handle,'','���������� Excel �� ������.',0);
 if OpenWorkBook('c:\��������� ������������ ��� ����-4 (����-4)\�������� (����).xls')
   then
     begin
       //messagebox(handle,'','������� �����.',0);
     end;

 //----------------------------------����� ������� � ��������� �� ������� 1-----------------------------------------
 key_message_tabl_2:=0;
 SetLength(For_Tabl_2,400);
 SetLength(data_tabl1,200);
 LIPA:= False;
 //All_tabl1:=E.Range['A26:A171'].Value;   //� All_tabl1 ������������ ���������� ����� �� ��������� 26:171 (������� 1), ��� ���������� ������ � ���������, ��� ������� ���� �������� �� �������
 For i:=36 to 171 do                   //���� ������� ������ �� ����� ����1 ��������� A26:A171
   begin
     data_tabl1[i]:=E.Range['A'+IntTostr(i)].Value;
     //Showmessage (data_tabl1[i]);
   end;

 FOR indx_finish_arr:=1 to kol_vesh_gpf_for_tabl_1 do   //���� �� ���������� �������� �������� � ����1 ���������
    BEGIN
      SetLength(Finish_arr_tabl_1,150);
      SetLength(Finish_arr_tabl_1_H,150);
      VarClear(Range);
      //ShowMessage(Arr_veshestv_gpf[indx_finish_arr]);
      Range := E.Range['A25:G178'].Find(What:=Arr_gpf_tabl_1[indx_finish_arr], LookIn:=xlValues,  SearchDirection:=xlNext, MatchCase:=True);
      if not VarIsClear(Range)
        then
          begin
            kol_naidenogo:=0;
            indx_arr:=0;
            FirstAddress := Range.Address;
            SetLength(str_of_tabl_1,150);
            SetLength(str_of_tabl_1_H,150);
            //ShowMessage(Range.Value);
            //ShowMessage(FirstAddress);

            //addr:=Range.Address;
            //addr[2]:='H';
            //ShowMessage(E.Range[addr].value);    //���������� ������ 'H', �������� � ���������

            //kol_naidenogo:=kol_naidenogo+1;
            repeat
              indx_arr:=indx_arr+1;
              //Range.Interior.ColorIndex := 37;
              Range := E.Range['A25:G178'].FindNext(After := Range);
              //ShowMessage(Range.Value);
              //ShowMessage(Range.Address);

              addr:=Range.Address;
              addr[2]:='H';
              //ShowMessage(E.Range[addr].value);  //���������� ������ 'H', �������� � ���������

              kol_naidenogo:=kol_naidenogo+1;
              str_of_tabl_1[indx_arr] := Range.Value;
              str_of_tabl_1_H[indx_arr] := E.Range[addr].value;
            until FirstAddress = Range.Address;                   // ������� ������ ���������� �� ���� ���������
            If kol_naidenogo>1
                then
                  for i:=1 to indx_arr do
                    Form2.combobox1.Items.Add(str_of_tabl_1[i]);
            //ShowMessage('���������� �������� ����� = '+IntToStr(kol_naidenogo));

            If kol_naidenogo=1
              then
                begin
                  Finish_arr_tabl_1[indx_finish_arr]:= str_of_tabl_1[indx_arr]; //���������� str_of_tabl_1[indx_arr] � Finish_arr_tabl_1;
                  Finish_arr_tabl_1_H[indx_finish_arr]:= str_of_tabl_1_H[indx_arr]; //���������� str_of_tabl_1_H[indx_arr] � Finish_arr_tabl_1_H;
                  //ShowMessage('� ������� 1 ����� ��������� '+Finish_arr_tabl_1[indx_finish_arr]+'  '+Finish_arr_tabl_1_H[indx_finish_arr]);
                end
              else
                begin
                  MessageBox(Handle,PChar('���������� �������� ��������� = '+IntToStr(kol_naidenogo)+ '. �������� ������'), '�������', MB_OK or MB_TOPMOST);
                  //ShowMessage('���������� �������� ��������� = '+IntToStr(kol_naidenogo)+ '. �������� ������');
                  Form2.ShowModal;
                  //vrem:=Form2.ComboBox1.ItemIndex;
                  Finish_arr_tabl_1[indx_finish_arr]:=Form2.choice_combo;
                  Finish_arr_tabl_1_H[indx_finish_arr] := str_of_tabl_1_H[Form2.vrem];
                  //ShowMessage('� ������� 1 ����� ��������� '+Finish_arr_tabl_1[indx_finish_arr]+'  '+Finish_arr_tabl_1_H[indx_finish_arr]);
                end;

          end
        else
          begin
            MessageBox(Handle,PChar('� ������� 1 '+Arr_gpf_tabl_1[indx_finish_arr]+' �� �������. �������� ������'), '�������', MB_OK or MB_TOPMOST);
            //Showmessage('� ������� 1 '+Arr_gpf_tabl_1[indx_finish_arr]+' �� �������. �������� ������');
            Memo1.Lines.Add('� ������� 1 '+Arr_gpf_tabl_1[indx_finish_arr]+' �� �������');
            LIPA:= True;
            //Showmessage('�������� ������ ��������');
            For i:=36 to 171 do
              begin
                //data_tabl1:=All_tabl1[i,1];
                Form2.combobox1.Items.Add(data_tabl1[i]);
              end;
            Form2.ShowModal;
            Finish_arr_tabl_1[indx_finish_arr]:=Form2.choice_combo;
            Finish_arr_tabl_1_H[indx_finish_arr] := E.Range['H'+IntToStr(Form2.vrem+25)].Value;
            //ShowMessage('� ������� 1 ����� ��������� '+Finish_arr_tabl_1[indx_finish_arr]+'  '+Finish_arr_tabl_1_H[indx_finish_arr]);
            Memo1.Lines.Add('� ������� 1 ����� ��������� '+Finish_arr_tabl_1[indx_finish_arr]);

            For i:=indx_finish_arr to (indx_finish_arr+indx_finish_arr) do   // ���� ���������� �������� ��� ����1 � ����2
              If (Arr_gpf_tabl_1[indx_finish_arr]=Arr_gpf_tabl_2[i])
                then
                  begin
                    j:=POS(' ', Finish_arr_tabl_1[indx_finish_arr]);
                    Arr_gpf_tabl_2[i]:=copy(Finish_arr_tabl_1[indx_finish_arr], 1, j);
                    If POS (Arr_gpf_tabl_2[i], ' ����� ������� ������������ �������� ')>0
                      then
                        begin
                          //Showmessage(Finish_arr_tabl_1[ch_vesh]);
                          //i2:=POS(' ', Finish_arr_tabl_1[indx_finish_arr]);
                          //Showmessage('i2 = '+IntToStr(i2));
                          j2:=PosEx(' ', Finish_arr_tabl_1[indx_finish_arr], j+1);
                          //Showmessage('j2 = '+IntToStr(j2));
                          Arr_gpf_tabl_2[i]:=copy(Finish_arr_tabl_1[indx_finish_arr], 1, j2);  // ����� ������ ������� ��� ������ �� ������� 2
                          //Showmessage('��������� ��� ������ � ������� 2 '+Arr_gpf_tabl_2[i]);
                        end;
                    If (Arr_gpf_tabl_1[indx_finish_arr]=Arr_gpf_tabl_2[i+1])
                      then
                        begin
                          Arr_gpf_tabl_2[i+1]:=Arr_gpf_tabl_2[i];
                          break;
                        end;
                  end;

          end;
      {i:=POS(' ', Finish_arr_tabl_1[indx_finish_arr]);
      For_Tabl_2[indx_finish_arr]:=copy(Finish_arr_tabl_1[indx_finish_arr], 1, i);  // ����� ������ ������� ��� ������ �� ������� 2
      showmessage('��������� ��� ������ � ������� 2 '+For_Tabl_2[indx_finish_arr]);}
      //str_of_tabl_1:=nil;
      //str_of_tabl_1_H:=nil;
      SetLength(str_of_tabl_1,0);
      SetLength(str_of_tabl_1_H,0);
    END;
  data_tabl1:=nil;

  //-----------------------------------------����� ������ ������� � ��������� �� ������� 1------------------------

  //----------------------------------����� ������� � ��������� �� ������� 2-----------------------------------------

  //Form3.ShowModal;
  MessageBox(0,'����������� ������� 2', '�������', MB_OK or MB_TOPMOST);
  //showmessage('����������� ������� 2');
  If 1 = 1
    then
      begin
        ch_vesh:=1;
        indx_finish_arr:=1;
        //FOR indx_finish_arr:=1 to Form1.kol_vesh_gpf_for_tabl_2 do   //���� �� ���������� �������� �������� � ����2 ���������
        WHILE (kol_vesh_gpf_for_tabl_2>=ch_vesh) do    //���� �� ���������� �������� ����2
          BEGIN
            SetLength(Finish_arr_tabl_2,150);
            SetLength(Finish_arr_tabl_2_G,150);
            VarClear(Range);
            //ShowMessage(For_Tabl_2[ch_vesh]);
            Range := E.Range['A180:A400'].Find(What:=Arr_gpf_tabl_2[ch_vesh], LookIn:=xlValues,  SearchDirection:=xlNext, MatchCase:=True);
            if not VarIsClear(Range)
              then
                begin
                  kol_naidenogo:=0;
                  indx_arr:=0;
                  FirstAddress := Range.Address;
                  SetLength(str_of_tabl_2,150);
                  SetLength(str_of_tabl_2_G,150);
                  //ShowMessage(Range.Value);
                  //ShowMessage(FirstAddress);
                  repeat
                    indx_arr:=indx_arr+1;
                    Range := E.Range['A180:A400'].FindNext(After := Range);
                    //ShowMessage(Range.Value);
                    //ShowMessage(Range.Address);

                    addr:=Range.Address;
                    addr[2]:='G';
                    //ShowMessage(E.Range[addr].value);  //���������� ������ 'G', �������� � ���������

                    kol_naidenogo:=kol_naidenogo+1;
                    str_of_tabl_2[indx_arr] := Range.Value;
                    str_of_tabl_2_G[indx_arr] := E.Range[addr].value;
                  until FirstAddress = Range.Address;                   // ������� ������ ���������� �� ���� ���������
                  If kol_naidenogo>2
                    then
                      for i:=1 to indx_arr do
                        Form2.combobox1.Items.Add(str_of_tabl_2[i]);
                  //ShowMessage('���������� �������� ����� = '+IntToStr(kol_naidenogo));

                  If (kol_naidenogo=1) or (kol_naidenogo=2)
                    then
                      begin
                        For i:=1 to indx_arr do
                          begin
                            If (Pos(P3orAb[ch_vesh], str_of_tabl_2[i])>0)
                              then
                                begin
                                  Finish_arr_tabl_2[indx_finish_arr]:= str_of_tabl_2[i]; //���������� str_of_tabl_2[indx_arr] � Finish_arr_tabl_2;
                                  Finish_arr_tabl_2_G[indx_finish_arr]:= str_of_tabl_2_G[i]; //���������� str_of_tabl_2_G[indx_arr] � Finish_arr_tabl_2_G;
                                  //ShowMessage('� ������� 2 ����� ��������� '+Finish_arr_tabl_2[indx_finish_arr]+'  '+Finish_arr_tabl_2_G[indx_finish_arr]);
                                  indx_finish_arr:=indx_finish_arr+1;
                                end;
                          end;
                     end
                    else
                      If kol_naidenogo>2
                        then
                          begin
                            MessageBox(Handle,PChar('���������� �������� ��������� = '+IntToStr(kol_naidenogo)+ '. �������� ������'), '�������', MB_OK or MB_TOPMOST);
                            //ShowMessage('���������� �������� ��������� = '+IntToStr(kol_naidenogo)+ '. �������� ������');
                            Form2.ShowModal;
                            //vrem:=Form2.ComboBox1.ItemIndex;
                            Finish_arr_tabl_2[indx_finish_arr]:=Form2.choice_combo;
                            Finish_arr_tabl_2_G[indx_finish_arr] := str_of_tabl_2_G[Form2.vrem];
                            //ShowMessage('� ������� 1 ����� ��������� '+Finish_arr_tabl_2[indx_finish_arr]+'  '+Finish_arr_tabl_2_G[indx_finish_arr]);
                            indx_finish_arr:=indx_finish_arr+1;
                          end;
                end
              else    // ���� �� �������
                begin
                  MessageBox(Handle,PChar('� ����2 �������� '+Arr_gpf_tabl_2[ch_vesh]+' �� �������. �������� �������������� ������ � ����1. �������������� ���� � �������� �������.'), '�������', MB_OK or MB_TOPMOST);
                  //Showmessage('� ����2 �������� '+Arr_gpf_tabl_2[ch_vesh]+' �� �������. �������� �������������� ������ � ����1. �������������� ���� � �������� �������.');
                  Memo1.Lines.Add('� ����2 �������� '+Arr_gpf_tabl_2[ch_vesh]+' �� �������');
                  Finish_arr_tabl_2[indx_finish_arr]:='�������� �� �������, �������������� ����';
                  //showmessage('! @ # � ������� 2 ����� ���������� : �������� �� �������, �������������� ����. ! " %');
                  indx_finish_arr:=indx_finish_arr+1;
                  LIPA:= True;
                  key_message_tabl_2:=1;
                end;
            ch_vesh:=ch_vesh+1;
            //str_of_tabl_2:=nil;
            //str_of_tabl_2_G:=nil;
            SetLength(str_of_tabl_2,0);
            SetLength(str_of_tabl_2_G,0);
          END;
      end ;

  //-----------------------------------------����� ������ ������� � ��������� �� ������� 2------------------------

  //SaveWorkBookAs('c:\1.�������� -09(����)SAVE.xls');
  //messagebox(handle,'','��������� ����� ��� "c:\1.�������� -09(����)SAVE.xls".',0);
  CloseWorkBook;
  //messagebox(handle,'','������� ����� "c:\1.xls".',0);
  CloseExcel;

  //��������� � ����������� ��������
  VisibleExcel(true);
  if OpenWorkBook('c:\��������� ������������ ��� ����-4 (����-4)\����������� 1.�������� -09(����).xls')
   then
     begin
       //messagebox(handle,'','������� �����.',0);
     end;
  //----���������� ����1-----------
  For i:=1 to ind_arr_gpf_tabl_1+1 do
    begin
    // ����������, ����������� ����� � ��� ����������� � ������� ����������� ������
      E.Rows[IntToStr(37)].Select;
      E.Selection.Copy;
      E.Rows[IntToStr(36+i)].Select;
      E.Selection.EntireRow.Insert;
      E.Selection.PasteSpecial(Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False);
      E.CutCopyMode := False;
    //
      {E.Rows[IntToStr(25+i)].Select;
      E.Selection.copy;
      E.Selection.EntireRow.Insert;
      E.ActiveCell.Paste; }
      E.Range['A'+IntToStr(36+i)]:= Finish_arr_tabl_1[i];
      E.Range['H'+IntToStr(36+i)]:= Finish_arr_tabl_1_H[i];

    end;
  E.Rows[IntToStr(36+i+1)].Select;
  E.Selection.Delete;
  E.Rows[IntToStr(36+i)].Select;
  E.Selection.Delete;
  E.Cells[36+i-1,1]:='���������� ������� ������������ � �������� 1 � 2.';

  //showmessage('���-�� � .gpf = '+IntToStr(kol_vesh_gpf_for_tabl_1)+ ', ���-�� � ������� 1 = '+IntToStr(i-3));


  //-----------�������� ����2-------------------
  For i:=1 to indx_finish_arr do
    begin
    // ����������, ����������� ����� � ��� ����������� � ������� ����������� ������
      E.Rows[IntToStr(48-1+kol_vesh_gpf_for_tabl_1)].Select;
      E.Selection.Copy;
      E.Rows[IntToStr(47-1+kol_vesh_gpf_for_tabl_1+i)].Select;
      E.Selection.EntireRow.Insert;
      E.Selection.PasteSpecial(Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False);
      E.CutCopyMode := False;
    //
      {E.Rows[IntToStr(25+i)].Select;
      E.Selection.copy;
      E.Selection.EntireRow.Insert;
      E.ActiveCell.Paste; }
      E.Range['A'+IntToStr(47-1+kol_vesh_gpf_for_tabl_1+i)]:= Finish_arr_tabl_2[i];
      E.Range['G'+IntToStr(47-1+kol_vesh_gpf_for_tabl_1+i)]:= Finish_arr_tabl_2_G[i];

    end;
  E.Rows[IntToStr(46-1+kol_vesh_gpf_for_tabl_1+i+1)].Select;
  E.Selection.Delete;
  E.Rows[IntToStr(46-1+kol_vesh_gpf_for_tabl_1+i)].Select;
  E.Selection.Delete;

  //showmessage('���-�� � .gpf = '+IntToStr(kol_vesh_gpf_for_tabl_2)+ ', ���-�� � ������� 2 = '+IntToStr(i-2));


  //----------------------------------------------
  SaveWorkBookAs('c:\��������� ������������ ��� ����-4 (����-4)\SAVE\���������\��������_'+Num_Pribor+'.xls');
  //messagebox(handle,'','��������� ����� ��� "c:\��������_'+Num_Pribor+'.xls".');
  MessageBox(Handle,PChar('��������� ��� "c:\��������� ������������ ��� ����-4 (����-4)\SAVE\���������\��������_'+Num_Pribor+'.xls".'), '�������', MB_OK or MB_TOPMOST);
  //showmessage('��������� ��� "c:\��������� ������������ ��� ����-4 (����-4)\SAVE\��������_'+Num_Pribor+'.xls".');
  CloseWorkBook;
  CloseExcel;

  If (kol_vesh_gpf_for_tabl_2<>(i-2)) or (key_message_tabl_2=1)
    then
      MessageBox(Handle,PChar('������� 2 ��������� �� ���������. �������������� ����������� �������� �������'), '�������', MB_OK or MB_TOPMOST);
      //showmessage('������� 2 ��������� �� ���������. �������������� ����������� �������� �������');
  //Finish_arr_tabl_1:=nil;
  //Finish_arr_tabl_1_H:=nil;

  //Finish_arr_tabl_2:=nil;
  //Finish_arr_tabl_2_G:=nil;

  SetLength(Finish_arr_tabl_1,0);
  SetLength(Finish_arr_tabl_1_H,0);
  SetLength(Finish_arr_tabl_2,0);
  SetLength(Finish_arr_tabl_2_G,0);

  //For_Tabl_2:=nil;
  SetLength(For_Tabl_2,0);

  Form1.Edit1.Text:='������� ��������';
  clik_button:=2;

 end;

procedure TForm1.Button2Click(Sender: TObject);
Var
position:integer;
ToFind: string; // ������, ��������� ������� ����
FindIn: string; // ��� ����
Found: integer; // ��������� ������
FoundLen: integer; //����� ���������� ������
  begin
    If Edit1.Text='������� ��������'
      then
        begin
          Showmessage('������� � �������');
          exit;
        end;
    //memo1.Lines.LoadFromFile('c:\��������� ������������ ��� ����-4 (����-4)\GPF\'+Num_Pribor+'\'+Num_Pribor+'.gpf');
    memo1.Lines.LoadFromFile(StFile);
    ToFind := '[������� �����]'; //��������� ������
    FindIn := Memo1.Lines.Text;//�����, ��� ����� ������
    FoundLen := Length('[������� �����]');
    Found := Pos(AnsiUpperCase(ToFind), AnsiUpperCase(FindIn));
    IF Found > 0 then
      begin
        Memo1.SelStart:= Found-1;
        Memo1.SelLength := FoundLen;
      end;
  end;


procedure TForm1.Button3Click(Sender: TObject);

var List : TStringList;
    //StFile : AnsiString;
    i, j, j1, j2, Nach_poisk, position_PorA, dlina_formuli , ap, pos_ru:Integer;
    St, Naidenie, Formula, Zaglavn_bukva, Eng:string;
    Alfavit, selectedFile : String ;

Begin
     //Parameters;
     If Edit1.Text='������� ��������'
       then
         begin
           Showmessage('������� � �������');
           exit;
         end;
     //Parameters;
     clik_button:=1;
     //Form1.Memo1.Clear;
     Form1.Memo2.Clear;
     Form1.Memo3.Clear;
     Alfavit := 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
     List := TStringList.Create;
     Num_Pribor:=Edit1.Text;

     // ����� .gpf ����� �� �����
     if (ParamStr(2)= '') or (vrem_edit <> Edit1.Text)
       then
         if PromptForFileName(selectedFile,        // ����� ������������� �����
                       'GPF files (*.gpf)|*.gpf',
                       '',
                       '�������� ������ ����',
                       'C:\',
                       False)  // ��������, ��� ������ ��� ����������
           then
             // ����������� ����� ������� �������� �����/����
             //ShowMessage('��������� ���� = '+selectedFile)
             StFile := selectedFile
           else
             begin
               ShowMessage('������ �� �������');
               exit;
             end
       else
         StFile := ParamStr(2);

     Form1.Memo1.Clear;
     //StFile := 'c:\��������� ������������ ��� ����-4 (����-4)\GPF\'+Num_Pribor+'\'+Num_Pribor+'.gpf'
     List.LoadFromFile(StFile);
     //Form1.Edit1.Text:='������� ��������';

     // ���� ������
     Eng:='ABEKMHOPCTaeopc';
     SetLength(Arr_gpf_tabl_1,0);
     SetLength(Arr_gpf_tabl_2,0);
     SetLength(Arr_gpf,0);

     SetLength(P3orAB,0);

     SetLength(Arr_gpf_tabl_1,150);
     SetLength(Arr_gpf_tabl_2,150);
     SetLength(Arr_gpf,150);

     SetLength(P3orAB,150);

     ind_arr_gpf_tabl_1:=1;
     ind_arr_gpf_tabl_2:=1;
     Nach_poisk:=List.IndexOf('[������� �����]');
     For i := Nach_poisk+1 to List.Count-1 do    //���� ������ ( i - �� ����� 1)
       begin
         //ShowMessage(List[i]);
         St:=List[i];
         position_PorA:=POS(' �=', ST);
         ap:=1;
         If position_PorA=0
           then
             begin
               position_PorA:=POS(' �=', ST);
               ap:=2;
               If position_PorA=0
                 then
                   begin
                     ShowMessage('������ � ��������� ��������, �������� �� ������� � ��� �');
                     ap:=0;
                   end;
             end;
         Naidenie:=copy(St,1,position_PorA);

         // 1-�� ����� ������ ���� ���������
         Zaglavn_bukva:= AnsiUpperCase(copy(naidenie,1,1));
         delete(naidenie,1,1);
         naidenie:=Zaglavn_bukva+naidenie;

         // �����������, ������� �������� ��� �������
         For j:=1 to Length(Naidenie) do
          begin
           If Pos('���� ', Naidenie)>0   // ����������. ���� ���� ������ �� �� �������.
             then
               Begin
                 j1:=Pos(' ', Naidenie);
                 Naidenie:=copy(Naidenie,1,j1);
                 break;
               End;
           If Pos('������������ ', Naidenie)>0   // ����������. ���� ���� ������ �� �� �������.
             then
               Begin
                 j1:=Pos(' ', Naidenie);
                 Naidenie:=copy(Naidenie,1,j1);
                 break;
               End;
           Formula:='';
             If (Pos(Naidenie[j], '-�����Ũ���������������������������������������������������������� ')=0)
                then
                  begin
                    //ShowMessage('���� �� ������� = ' +Naidenie);
                    j1:=j;
                    While (Naidenie[j1-1]<> ' ') do   //����� � ������ �������
                      begin
                        j1:=j1-1;
                      end;
                    while (Naidenie[j1]<> ' ') do   // ����� ����� �������
                      begin
                        Formula:=Formula+Naidenie[j1];
                        j1:=j1+1;
                      end;
                    Naidenie:= ' '+Formula+' ';
                    // ������� �������� �������� � ����������� (��� ��������)
                     For j2:=1 to Length(Naidenie) do
                       begin
                        If Pos(Naidenie[j2], '���������������')>0
                          then
                            Begin
                              pos_ru:=Pos(Naidenie[j2], '���������������');
                              //ShowMessage(Naidenie[j2]+ ' - ���');
                              Naidenie[j2]:= Eng[pos_ru];
                              //ShowMessage(Naidenie[j2]+ ' - ENG');
                            End;
                       end;

                    //ShowMessage(Naidenie);
                    break;
                  end;

          end;
         //ShowMessage(Naidenie);

         If not (Arr_gpf_tabl_1[ind_arr_gpf_tabl_1-1]=Naidenie)  // ����� ����������� ������� ��� ����1 � ����2. � ������������ �.�. ��� �.�.
           then
             begin
               Arr_gpf_tabl_1[ind_arr_gpf_tabl_1]:=Naidenie;       // ������ ��� ������� 1.
               ind_arr_gpf_tabl_1:=ind_arr_gpf_tabl_1+1;
               Arr_gpf_tabl_2[ind_arr_gpf_tabl_2]:=Naidenie;       // ������ ��� ������� 2 � ���� ��������. �������� ���������� ������� �� � ��
               Arr_gpf[ind_arr_gpf_tabl_2]:=Naidenie;              // ������ ��� ��������� ��������. �������� ���������� ������� �� � ��
               if ap=1
                 then
                   P3orAB[ind_arr_gpf_tabl_2]:='�.�.'
                 else
                   if ap=2
                     then
                       P3orAB[ind_arr_gpf_tabl_2]:='�.�.'
                     else
                       P3orAB[ind_arr_gpf_tabl_2]:=' ';
               ind_arr_gpf_tabl_2:=ind_arr_gpf_tabl_2+1;
               Memo2.Lines.Add(Arr_gpf_tabl_1[ind_arr_gpf_tabl_1-1]);
             end
           else
             begin
               Arr_gpf_tabl_2[ind_arr_gpf_tabl_2]:=Naidenie;
               Arr_gpf[ind_arr_gpf_tabl_2]:=Naidenie;
               if ap=1
                 then
                   P3orAB[ind_arr_gpf_tabl_2]:='�.�.'
                 else
                   if ap=2
                     then
                       P3orAB[ind_arr_gpf_tabl_2]:='�.�.'
                     else
                       P3orAB[ind_arr_gpf_tabl_2]:=' ';
               ind_arr_gpf_tabl_2:=ind_arr_gpf_tabl_2+1;
             end;

         Memo3.Lines.Add(Arr_gpf_tabl_2[ind_arr_gpf_tabl_2-1]);

         kol_vesh_gpf_for_tabl_2:=ind_arr_gpf_tabl_2-1;
         kol_vesh_gpf_for_tabl_1:=ind_arr_gpf_tabl_1-1;




         //position_probel:=POS(' ', st);
         //ShowMessage('position_probel = ' +IntToStr(position_probel));
         //For j := 1 to Length(St) do
           //ShowMessage(St[j]);
     end;
    MessageBox(Handle,PChar('����1 = '+IntToStr(kol_vesh_gpf_for_tabl_1)+' � ���� 2 = '+IntToStr(kol_vesh_gpf_for_tabl_2)+' �������'), '�������', MB_OK or MB_TOPMOST);
    //ShowMessage('����1 = '+IntToStr(kol_vesh_gpf_for_tabl_1)+' � ���� 2 = '+IntToStr(kol_vesh_gpf_for_tabl_2)+' �������');
    //Arr_veshestv_gpf:=nil;
    List.Free;

end;


procedure TForm1.Button4Click(Sender: TObject);
var
  Finish_arr_pass_orig: array of string;
  Finish_arr_pass_orig_formula: array of string;
  Finish_arr_pass_orig_diap_izm: array of string;
  Finish_arr_pass_orig_tip_datch: array of string;

  Finish_arr_pass_orig_LIPA: array of string;
  Finish_arr_pass_orig_formula_LIPA: array of string;
  Finish_arr_pass_orig_diap_izm_LIPA: array of string;
  Finish_arr_pass_orig_tip_datch_LIPA: array of string;

  str_of_263: array of string;
  str_of_263_formula: array of string;
  str_of_263_diap_izm: array of string;
  str_of_263_tip_datch : array of string;

  data_263: array of string;

  kol_naidenogo, indx_arr, i: integer;

  FirstAddress, addr_vesh, addr_formula, addr_diap_izm_1, addr_diap_izm_2, addr_tip_datch   : string;

  //----------���������� ��� ��������
  Range_tabl1 : Variant;
  tablica_:integer;
   col_,row_:integer;
   a_:integer;
   metki_:array[1..12]of record
   col:integer;
   row:integer;
   metka:string;
   end;

begin
  If clik_button=0
   then
     begin
       Showmessage('������� �������� � ������');
       exit;
     end;
  If not(clik_button=2)
   then
     begin
       Showmessage('������� ����� ������������ ��������');
       exit;
     end;
  if not CreateExcel
   then
     exit;
  //VisibleExcel(true);
 //messagebox(handle,'','���������� Excel �� ������.',0);
  if OpenWorkBook('c:\��������� ������������ ��� ����-4 (����-4)\���� 263 ����2xls.xls')
    then
      begin
        //messagebox(handle,'','������� �����.',0);
      end;
  //---------------------�������-�������� --------------------

  SetLength(Finish_arr_pass_orig,150);
  SetLength(Finish_arr_pass_orig_formula,150);
  SetLength(Finish_arr_pass_orig_diap_izm,150);
  SetLength(Finish_arr_pass_orig_tip_datch,150);
  SetLength(Finish_arr_pass_orig_LIPA,150);
  SetLength(Finish_arr_pass_orig_formula_LIPA,150);
  SetLength(Finish_arr_pass_orig_diap_izm_LIPA,150);
  SetLength(Finish_arr_pass_orig_tip_datch_LIPA,150);
  SetLength(str_of_263,150);
  SetLength(str_of_263_formula,150);
  SetLength(str_of_263_diap_izm,150);
  SetLength(str_of_263_tip_datch,150);
  SetLength(data_263,400);
  For i:=7 to 333 do                   //���� ������� ������ �� ����� ����1 ��������� A26:A171
    begin
      data_263[i]:=E.Range['W'+IntTostr(i)].Value;
      //Showmessage (data_tabl1[i]);
    end;
  FOR indx_finish_arr:=1 to kol_vesh_gpf_for_tabl_2 do   //���� �� ���������� �������� �������� � ����1 ���������
    BEGIN
      VarClear(Range);
      //ShowMessage(Arr_veshestv_gpf[indx_finish_arr]);
      Range := E.Range['C7:E333'].Find(What:=Arr_gpf[indx_finish_arr], LookIn:=xlValues,  SearchDirection:=xlNext, MatchCase:=True);
      if not VarIsClear(Range)
        then
          begin
            kol_naidenogo:=0;
            indx_arr:=0;
            FirstAddress := Range.Address;

            //ShowMessage(Range.Value);
            //ShowMessage(FirstAddress);

            //addr:=Range.Address;
            //addr[2]:='H';
            //ShowMessage(E.Range[addr].value);    //���������� ������ 'H', �������� � ���������

            //kol_naidenogo:=kol_naidenogo+1;
            repeat
              indx_arr:=indx_arr+1;
              //Range.Interior.ColorIndex := 37;
              Range := E.Range['C7:C333'].FindNext(After := Range);
              //ShowMessage(Range.Value);
              //ShowMessage(Range.Address);

              addr_vesh:=Range.Address;
              addr_vesh[2]:='W';

              addr_formula:=Range.Address;
              addr_formula[2]:='E';
              //ShowMessage(E.Range[addr].value);  //���������� ������ 'H', �������� � ���������

              If P3orAB[indx_finish_arr]='�.�.'
                then
                  begin
                    addr_diap_izm_1:=Range.Address;
                    addr_diap_izm_1[2]:='P';
                    addr_diap_izm_2:=Range.Address;
                    addr_diap_izm_2[2]:='R';
                    //ShowMessage(E.Range[addr].value);  //���������� ������ 'H', �������� � ���������
                  end
                else
                  begin
                    addr_diap_izm_1:=Range.Address;
                    addr_diap_izm_1[2]:='M';
                    addr_diap_izm_2:=Range.Address;
                    addr_diap_izm_2[2]:='O';
                    //ShowMessage(E.Range[addr].value);  //���������� ������ 'H', �������� � ���������
                  end;

              addr_tip_datch:=Range.Address;
              addr_tip_datch[2]:='V';
              //ShowMessage(E.Range[addr].value);  //���������� ������ 'H', �������� � ���������

              kol_naidenogo:=kol_naidenogo+1;
              str_of_263[indx_arr] := E.Range[addr_vesh].value;

              str_of_263_formula[indx_arr] := E.Range[addr_formula].value;
              str_of_263_diap_izm[indx_arr]:= E.Range[addr_diap_izm_1].text;
              str_of_263_diap_izm[indx_arr]:=str_of_263_diap_izm[indx_arr]+'-'+E.Range[addr_diap_izm_2].text;
              str_of_263_tip_datch[indx_arr]:= E.Range[addr_tip_datch].value;
            until FirstAddress = Range.Address;                   // ������� ������ ���������� �� ���� ���������
            If kol_naidenogo>1
                then
                  for i:=1 to indx_arr do
                    Form2.combobox1.Items.Add(str_of_263[i]);
            //ShowMessage('���������� �������� ����� = '+IntToStr(kol_naidenogo));

            If kol_naidenogo=1
              then
                begin
                  Finish_arr_pass_orig[indx_finish_arr]:= str_of_263[indx_arr]+' '+P3orAB[indx_finish_arr];
                  Finish_arr_pass_orig_formula[indx_finish_arr]:= str_of_263_formula[indx_arr];
                  Finish_arr_pass_orig_diap_izm[indx_finish_arr]:= str_of_263_diap_izm[indx_arr];
                  Finish_arr_pass_orig_tip_datch[indx_finish_arr]:= str_of_263_tip_datch[indx_arr];
                  MessageBox(Handle,PChar('� �������-�������� ����� ��������� '+Finish_arr_pass_orig[indx_finish_arr]+'  '+Finish_arr_pass_orig_formula[indx_finish_arr]+' '+Finish_arr_pass_orig_diap_izm[indx_finish_arr]+' '+Finish_arr_pass_orig_tip_datch[indx_finish_arr]), '�������', MB_OK or MB_TOPMOST);
                  //ShowMessage('� �������-�������� ����� ��������� '+Finish_arr_pass_orig[indx_finish_arr]+'  '+Finish_arr_pass_orig_formula[indx_finish_arr]+' '+Finish_arr_pass_orig_diap_izm[indx_finish_arr]+' '+Finish_arr_pass_orig_tip_datch[indx_finish_arr]);
                end
              else
                begin
                  MessageBox(Handle,PChar('���������� �������� ��������� = '+IntToStr(kol_naidenogo)+ '. �������� ������'), '�������', MB_OK or MB_TOPMOST);
                  //ShowMessage('���������� �������� ��������� = '+IntToStr(kol_naidenogo)+ '. �������� ������');
                  Form2.ShowModal;
                  Finish_arr_pass_orig[indx_finish_arr]:= Form2.choice_combo+' '+P3orAB[indx_finish_arr];
                  Finish_arr_pass_orig_formula[indx_finish_arr]:= str_of_263_formula[Form2.vrem];
                  Finish_arr_pass_orig_diap_izm[indx_finish_arr]:= str_of_263_diap_izm[Form2.vrem];
                  Finish_arr_pass_orig_tip_datch[indx_finish_arr]:= str_of_263_tip_datch[Form2.vrem];
                  MessageBox(Handle,PChar('� �������-�������� ����� ��������� '+Finish_arr_pass_orig[indx_finish_arr]+'  '+Finish_arr_pass_orig_formula[indx_finish_arr]+' '+Finish_arr_pass_orig_diap_izm[indx_finish_arr]+' '+Finish_arr_pass_orig_tip_datch[indx_finish_arr]), '�������', MB_OK or MB_TOPMOST);
                  //ShowMessage('� �������-�������� ����� ��������� '+Finish_arr_pass_orig[indx_finish_arr]+'  '+Finish_arr_pass_orig_formula[indx_finish_arr]+' '+Finish_arr_pass_orig_diap_izm[indx_finish_arr]+' '+Finish_arr_pass_orig_tip_datch[indx_finish_arr]);
                end;
          end
        else
          begin
            MessageBox(Handle,PChar('� ����� 263 '+Arr_gpf[indx_finish_arr]+' �� �������. �������� ������. �������� ����� ������ �� ������������� ������'), '�������', MB_OK or MB_TOPMOST);
            //Showmessage('� ����� 263 '+Arr_gpf[indx_finish_arr]+' �� �������. �������� ������. �������� ����� ������ �� ������������� ������');
            Memo1.Lines.Add('� ����� 263 '+Arr_gpf[indx_finish_arr]+' �� �������');
            For i:=7 to 333 do
              Form2.combobox1.Items.Add(data_263[i]);
            Form2.ShowModal;
            Finish_arr_pass_orig[indx_finish_arr]:= Form2.choice_combo+' '+P3orAB[indx_finish_arr];
            Finish_arr_pass_orig_formula[indx_finish_arr]:= E.Range['E'+IntToStr(Form2.vrem+6)].Value;
            If P3orAB[indx_finish_arr]='�.�.'
                then
                  begin
                    Finish_arr_pass_orig_diap_izm[indx_finish_arr]:= E.Range['P'+IntToStr(Form2.vrem+6)].Text+'-'+E.Range['R'+IntToStr(Form2.vrem+6)].Text;
                  end
                else
                  begin
                    Finish_arr_pass_orig_diap_izm[indx_finish_arr]:= E.Range['M'+IntToStr(Form2.vrem+6)].Text+'-'+E.Range['O'+IntToStr(Form2.vrem+6)].Text;
                  end;
            Finish_arr_pass_orig_tip_datch[indx_finish_arr]:= E.Range['V'+IntToStr(Form2.vrem+6)].Value;
            MessageBox(Handle,PChar('� �������-�������� ����� ��������� '+Finish_arr_pass_orig[indx_finish_arr]+'  '+Finish_arr_pass_orig_formula[indx_finish_arr]+' '+Finish_arr_pass_orig_diap_izm[indx_finish_arr]+' '+Finish_arr_pass_orig_tip_datch[indx_finish_arr]), '�������', MB_OK or MB_TOPMOST);
            //ShowMessage('� �������-�������� ����� ��������� '+Finish_arr_pass_orig[indx_finish_arr]+'  '+Finish_arr_pass_orig_formula[indx_finish_arr]+' '+Finish_arr_pass_orig_diap_izm[indx_finish_arr]+' '+Finish_arr_pass_orig_tip_datch[indx_finish_arr]);
            Memo1.Lines.Add('� �������-�������� ����� ��������� '+Finish_arr_pass_orig[indx_finish_arr]);
          end;

    END;


  //---------------------��������_4  ------------------------
  If LIPA=True
    then
      begin
        MessageBox(Handle,PChar('!!!������ ������������ �������� 4!!!'), '�������', MB_OK or MB_TOPMOST);
        //Showmessage('!!!������ ������������ �������� 4!!!');
        FOR indx_finish_arr:=1 to kol_vesh_gpf_for_tabl_2 do   //���� �� ���������� �������� �������� � ����1 ���������
          BEGIN
            VarClear(Range);
            //ShowMessage(Arr_veshestv_gpf[indx_finish_arr]);
            Range := E.Range['C7:E333'].Find(What:=Arr_gpf_tabl_2[indx_finish_arr], LookIn:=xlValues,  SearchDirection:=xlNext, MatchCase:=True);
            if not VarIsClear(Range)
              then
                begin
                  kol_naidenogo:=0;
                  indx_arr:=0;
                  FirstAddress := Range.Address;
                  repeat
                    indx_arr:=indx_arr+1;
                    Range := E.Range['C7:C333'].FindNext(After := Range);
                    addr_vesh:=Range.Address;
                    addr_vesh[2]:='W';
                    addr_formula:=Range.Address;
                    addr_formula[2]:='E';
                    If P3orAB[indx_finish_arr]='�.�.'
                      then
                        begin
                          addr_diap_izm_1:=Range.Address;
                          addr_diap_izm_1[2]:='P';
                          addr_diap_izm_2:=Range.Address;
                          addr_diap_izm_2[2]:='R';
                        end
                      else
                        begin
                          addr_diap_izm_1:=Range.Address;
                          addr_diap_izm_1[2]:='M';
                          addr_diap_izm_2:=Range.Address;
                          addr_diap_izm_2[2]:='O';
                        end;
                    addr_tip_datch:=Range.Address;
                    addr_tip_datch[2]:='V';

                    kol_naidenogo:=kol_naidenogo+1;
                    str_of_263[indx_arr] := E.Range[addr_vesh].value;

                    str_of_263_formula[indx_arr] := E.Range[addr_formula].value;
                    str_of_263_diap_izm[indx_arr]:= E.Range[addr_diap_izm_1].text;
                    str_of_263_diap_izm[indx_arr]:=str_of_263_diap_izm[indx_arr]+'-'+E.Range[addr_diap_izm_2].text;
                    str_of_263_tip_datch[indx_arr]:= E.Range[addr_tip_datch].value;
                  until FirstAddress = Range.Address;                   // ������� ������ ���������� �� ���� ���������
                  If kol_naidenogo>1
                    then
                      for i:=1 to indx_arr do
                        Form2.combobox1.Items.Add(str_of_263[i]);
                  If kol_naidenogo=1
                    then
                      begin
                        Finish_arr_pass_orig_LIPA[indx_finish_arr]:= str_of_263[indx_arr]+' '+P3orAB[indx_finish_arr];
                        Finish_arr_pass_orig_formula_LIPA[indx_finish_arr]:= str_of_263_formula[indx_arr];
                        Finish_arr_pass_orig_diap_izm_LIPA[indx_finish_arr]:= str_of_263_diap_izm[indx_arr];
                        Finish_arr_pass_orig_tip_datch_LIPA[indx_finish_arr]:= str_of_263_tip_datch[indx_arr];
                        MessageBox(Handle,PChar('� ���-4 ����� ��������� '+Finish_arr_pass_orig_LIPA[indx_finish_arr]+'  '+Finish_arr_pass_orig_formula_LIPA[indx_finish_arr]+' '+Finish_arr_pass_orig_diap_izm_LIPA[indx_finish_arr]+' '+Finish_arr_pass_orig_tip_datch_LIPA[indx_finish_arr]), '�������', MB_OK or MB_TOPMOST);
                        //ShowMessage('� �������-�������� ����� ��������� '+Finish_arr_pass_orig_LIPA[indx_finish_arr]+'  '+Finish_arr_pass_orig_formula_LIPA[indx_finish_arr]+' '+Finish_arr_pass_orig_diap_izm_LIPA[indx_finish_arr]+' '+Finish_arr_pass_orig_tip_datch_LIPA[indx_finish_arr]);
                      end
                    else
                      begin
                        if Arr_gpf_tabl_2[indx_finish_arr]=Arr_gpf[indx_finish_arr]     //���� ��������������� � ��� �� ��������(� ��� �� �����), ��� � � ���������. �� ����������� �������� �� ���������
                          then
                            begin
                              Finish_arr_pass_orig_LIPA[indx_finish_arr]:= Finish_arr_pass_orig[indx_finish_arr];
                              Finish_arr_pass_orig_formula_LIPA[indx_finish_arr]:= Finish_arr_pass_orig_formula[indx_finish_arr];
                              Finish_arr_pass_orig_diap_izm_LIPA[indx_finish_arr]:= Finish_arr_pass_orig_diap_izm[indx_finish_arr];
                              Finish_arr_pass_orig_tip_datch_LIPA[indx_finish_arr]:= Finish_arr_pass_orig_tip_datch[indx_finish_arr];
                              MessageBox(Handle,PChar('� ���-4 ����� ��������� '+Finish_arr_pass_orig_LIPA[indx_finish_arr]+'  '+Finish_arr_pass_orig_formula_LIPA[indx_finish_arr]+' '+Finish_arr_pass_orig_diap_izm_LIPA[indx_finish_arr]+' '+Finish_arr_pass_orig_tip_datch_LIPA[indx_finish_arr]), '�������', MB_OK or MB_TOPMOST);
                              //ShowMessage('� �������-�������� ����� ��������� '+Finish_arr_pass_orig_LIPA[indx_finish_arr]+'  '+Finish_arr_pass_orig_formula_LIPA[indx_finish_arr]+' '+Finish_arr_pass_orig_diap_izm_LIPA[indx_finish_arr]+' '+Finish_arr_pass_orig_tip_datch_LIPA[indx_finish_arr]);
                            end
                          else      //...����� ���������� ����� �� ����� 2
                            begin
                              MessageBox(Handle,PChar('���������� �������� ��������� = '+IntToStr(kol_naidenogo)+ '. �������� ������'), '�������', MB_OK or MB_TOPMOST);
                              //ShowMessage('���������� �������� ��������� = '+IntToStr(kol_naidenogo)+ '. �������� ������');
                              Form2.ShowModal;
                              Finish_arr_pass_orig_LIPA[indx_finish_arr]:= Form2.choice_combo+' '+P3orAB[indx_finish_arr];
                              Finish_arr_pass_orig_formula_LIPA[indx_finish_arr]:= str_of_263_formula[Form2.vrem];
                              Finish_arr_pass_orig_diap_izm_LIPA[indx_finish_arr]:= str_of_263_diap_izm[Form2.vrem];
                              Finish_arr_pass_orig_tip_datch_LIPA[indx_finish_arr]:= str_of_263_tip_datch[Form2.vrem];
                              MessageBox(Handle,PChar('� �������-�������� ����� ��������� '+Finish_arr_pass_orig_LIPA[indx_finish_arr]+'  '+Finish_arr_pass_orig_formula_LIPA[indx_finish_arr]+' '+Finish_arr_pass_orig_diap_izm_LIPA[indx_finish_arr]+' '+Finish_arr_pass_orig_tip_datch_LIPA[indx_finish_arr]), '�������', MB_OK or MB_TOPMOST);
                              //ShowMessage('� �������-�������� ����� ��������� '+Finish_arr_pass_orig_LIPA[indx_finish_arr]+'  '+Finish_arr_pass_orig_formula_LIPA[indx_finish_arr]+' '+Finish_arr_pass_orig_diap_izm_LIPA[indx_finish_arr]+' '+Finish_arr_pass_orig_tip_datch_LIPA[indx_finish_arr]);
                            end;
                      end;
                end
              else
                begin
                  MessageBox(Handle,PChar('� ����� 263 '+Arr_gpf[indx_finish_arr]+' �� �������. �������� ������. �������� ����� ������ �� ������������� ������'), '�������', MB_OK or MB_TOPMOST);
                  //Showmessage('� ����� 263 '+Arr_gpf[indx_finish_arr]+' �� �������. �������� ������. �������� ����� ������ �� ������������� ������');
                  Memo1.Lines.Add('� ����� 263 '+Arr_gpf[indx_finish_arr]+' �� �������');
                  {Form3.ShowModal;
                  If Form3.variant_vibora = 1
                    then
                      begin
                        //For i:=7 to 333 do
                        //   Form2.combobox1.Items.Add(data_263[i]);
                        //Form2.ShowModal;
                        Finish_arr_pass_orig_LIPA[indx_finish_arr]:= Finish_arr_pass_orig[indx_finish_arr];
                        Finish_arr_pass_orig_formula_LIPA[indx_finish_arr]:= Finish_arr_pass_orig_formula[indx_finish_arr];
                        If P3orAB[indx_finish_arr]='�.�.'
                          then
                            begin
                              Finish_arr_pass_orig_diap_izm_LIPA[indx_finish_arr]:= Finish_arr_pass_orig_diap_izm[indx_finish_arr]
                            end
                          else
                            begin
                              Finish_arr_pass_orig_diap_izm_LIPA[indx_finish_arr]:= Finish_arr_pass_orig_diap_izm[indx_finish_arr]
                            end;
                        Finish_arr_pass_orig_tip_datch_LIPA[indx_finish_arr]:= Finish_arr_pass_orig_tip_datch[indx_finish_arr];
                        ShowMessage('� �������-�������� ����� ��������� '+Finish_arr_pass_orig_LIPA[indx_finish_arr]+'  '+Finish_arr_pass_orig_formula_LIPA[indx_finish_arr]+' '+Finish_arr_pass_orig_diap_izm_LIPA[indx_finish_arr]+' '+Finish_arr_pass_orig_tip_datch_LIPA[indx_finish_arr]);
                      end
                  else
                    begin  }
                      Form2.Combobox1.Items.Clear;
                      For i:=7 to 333 do
                        Form2.combobox1.Items.Add(data_263[i]);
                      Form2.ShowModal;
                      Finish_arr_pass_orig_LIPA[indx_finish_arr]:= Form2.choice_combo+' '+P3orAB[indx_finish_arr];
                      Finish_arr_pass_orig_formula_LIPA[indx_finish_arr]:= E.Range['E'+IntToStr(Form2.vrem+6)].Value;
                      If P3orAB[indx_finish_arr]='�.�.'
                        then
                          begin
                            Finish_arr_pass_orig_diap_izm_LIPA[indx_finish_arr]:= E.Range['P'+IntToStr(Form2.vrem+6)].Text+'-'+E.Range['R'+IntToStr(Form2.vrem+6)].Text;
                          end
                        else
                          begin
                            Finish_arr_pass_orig_diap_izm_LIPA[indx_finish_arr]:= E.Range['M'+IntToStr(Form2.vrem+6)].Text+'-'+E.Range['O'+IntToStr(Form2.vrem+6)].Text;
                          end;
                      Finish_arr_pass_orig_tip_datch_LIPA[indx_finish_arr]:= E.Range['V'+IntToStr(Form2.vrem+6)].Value;
                      MessageBox(Handle,PChar('� �������-�������� ����� ��������� '+Finish_arr_pass_orig_LIPA[indx_finish_arr]+'  '+Finish_arr_pass_orig_formula_LIPA[indx_finish_arr]+' '+Finish_arr_pass_orig_diap_izm_LIPA[indx_finish_arr]+' '+Finish_arr_pass_orig_tip_datch_LIPA[indx_finish_arr]), '�������', MB_OK or MB_TOPMOST);
                      //ShowMessage('� �������-�������� ����� ��������� '+Finish_arr_pass_orig_LIPA[indx_finish_arr]+'  '+Finish_arr_pass_orig_formula_LIPA[indx_finish_arr]+' '+Finish_arr_pass_orig_diap_izm_LIPA[indx_finish_arr]+' '+Finish_arr_pass_orig_tip_datch_LIPA[indx_finish_arr]);
                      Memo1.Lines.Add('� �������-�������� ����� ��������� '+Finish_arr_pass_orig_LIPA[indx_finish_arr]);
                    //end;
                end;

          END;
      end;

  CloseWorkBook;
  CloseExcel;

  //---------------������� (WORD)----------------

  if CreateWord
    then
      begin
        VisibleWord(true);
        If OpenDoc('c:\��������� ������������ ��� ����-4 (����-4)\�������_������.doc')
          then
            begin
              StartOfDoc;
              while FindAndPasteTextDoc('XXX', Num_Pribor) do  // Zavod_Nom - ��� ���������� ������� ������ ������������ ��� ������� �������
                StartOfDoc;
            end
          else
            MessageBox(Handle,PChar('������ �������� ����������� '), '������', MB_OK or MB_TOPMOST);
            //Showmessage('������ �������� ����������� ');
      end
    else
      MessageBox(Handle,PChar('MS Office �� ���������� '), '������', MB_OK or MB_TOPMOST);
      //Showmessage('MS Office �� ���������� ');
  //---------������� 1-----
  tablica_:=1;
  FOR i:=1 to kol_vesh_gpf_for_tabl_2 do
    Begin
      W.ActiveDocument.Tables.Item(1).Cell(i+1,1).Range.text:=IntToStr(i);
      W.ActiveDocument.Tables.Item(1).Cell(i+1,2).Range.text:=Finish_arr_pass_orig[i];
      W.ActiveDocument.Tables.Item(1).Cell(i+1,3).Range.text:=Finish_arr_pass_orig_formula[i];
      W.ActiveDocument.Tables.Item(1).Cell(i+1,4).Range.text:=Finish_arr_pass_orig_diap_izm[i];
      W.ActiveDocument.Tables.Item(1).Cell(i+1,5).Range.text:=Finish_arr_pass_orig_tip_datch[i];
      InsertRowTableDoc(tablica_,i+2);
    End;
  W.ActiveDocument.Tables.Item(tablica_).Rows.Item(i+2).Delete;
  W.ActiveDocument.Tables.Item(tablica_).Rows.Item(i+1).Delete;

  //����� 5-�� ��������
  //W.Selection.End:=2774;
  //W.Selection.Start:=2774;
  If kol_vesh_gpf_for_tabl_2<=33
    then
      if kol_vesh_gpf_for_tabl_2<14
        then
          begin
            FindTextDoc('&');
            For i:=1 to kol_vesh_gpf_for_tabl_2-1 do
              begin
                W.Selection.TypeBackspace;
              end;
          end
        else
          if kol_vesh_gpf_for_tabl_2<25
            then
              begin
                FindTextDoc('&');
                For i:=1 to kol_vesh_gpf_for_tabl_2-2 do
                  begin
                    W.Selection.TypeBackspace;
                  end;
              end
            else
              begin
                FindTextDoc('&');
                For i:=1 to kol_vesh_gpf_for_tabl_2-3 do
                  begin
                    W.Selection.TypeBackspace;
                  end;
              end
    else
      begin
        MessageBox(Handle,PChar('�������� ����� 4 "�������������" �� ������ �������� '), '�������', MB_OK or MB_TOPMOST);
        FindTextDoc('&');
        W.Selection.TypeBackspace;
      end;


  MessageBox(Handle,PChar('��������� ���... '), '�������', MB_OK or MB_TOPMOST);
  //Showmessage('��������� ���... ');
  SaveDocAs('c:\��������� ������������ ��� ����-4 (����-4)\SAVE\�������_'+Num_Pribor+'.doc');
  CloseDoc;
  CloseWord;

  //-------������� 2-----------
  if CreateWord
    then
      begin
        //VisibleWord(true);
        If OpenDoc('c:\��������� ������������ ��� ����-4 (����-4)\��������_������.doc')
          then
            begin
              StartOfDoc;
            end
          else
            MessageBox(Handle,PChar('������ �������� ����������� '), '������', MB_OK or MB_TOPMOST);
            //Showmessage('������ �������� ����������� ');
      end
    else
      MessageBox(Handle,PChar('MS Office �� ���������� '), '������', MB_OK or MB_TOPMOST);
      //Showmessage('MS Office �� ���������� ');
  FOR i:=1 to kol_vesh_gpf_for_tabl_2 do
    Begin
      W.ActiveDocument.Tables.Item(1).Cell(i+1,1).Range.text:=IntToStr(i);
      W.ActiveDocument.Tables.Item(1).Cell(i+1,2).Range.text:=Finish_arr_pass_orig[i];
      W.ActiveDocument.Tables.Item(1).Cell(i+1,3).Range.text:=Finish_arr_pass_orig_formula[i];
      W.ActiveDocument.Tables.Item(1).Cell(i+1,4).Range.text:=Finish_arr_pass_orig_diap_izm[i];
      W.ActiveDocument.Tables.Item(1).Cell(i+1,5).Range.text:=Finish_arr_pass_orig_tip_datch[i];
      InsertRowTableDoc(1,i+2);
    End;
  W.ActiveDocument.Tables.Item(1).Rows.Item(i+2).Delete;
  W.ActiveDocument.Tables.Item(1).Rows.Item(i+1).Delete;

  //MessageBox(Handle,PChar('��������� ���... '), '�������', MB_OK or MB_TOPMOST);
  //Showmessage('��������� ���... ');
  SaveDocAs('c:\��������� ������������ ��� ����-4 (����-4)\SAVE\��������\��������_'+Num_Pribor+'.doc');
  CloseDoc;
  CloseWord;
  //-----------------���_4 -------------------
  If LIPA=True
    then
      begin
        if CreateWord
          then
            begin
              VisibleWord(true);
              If OpenDoc('c:\��������� ������������ ��� ����-4 (����-4)\���_4_������.doc')
                then
                  begin
                    StartOfDoc;
                    while FindAndPasteTextDoc('XXX', Num_Pribor) do
                      StartOfDoc;
                  end
                else
                  MessageBox(Handle,PChar('������ �������� ����������� '), '������', MB_OK or MB_TOPMOST);
                  //Showmessage('������ �������� ����������� ');
            end
          else
            MessageBox(Handle,PChar('MS Office �� ���������� '), '������', MB_OK or MB_TOPMOST);
            //Showmessage('MS Office �� ���������� ');
        //---------������� 1 �-----
        tablica_:=1;
        FOR i:=1 to kol_vesh_gpf_for_tabl_2 do
          Begin
            W.ActiveDocument.Tables.Item(1).Cell(i+1,1).Range.text:=IntToStr(i);
            W.ActiveDocument.Tables.Item(1).Cell(i+1,2).Range.text:=Finish_arr_pass_orig_LIPA[i];
            W.ActiveDocument.Tables.Item(1).Cell(i+1,3).Range.text:=Finish_arr_pass_orig_formula_LIPA[i];
            W.ActiveDocument.Tables.Item(1).Cell(i+1,4).Range.text:=Finish_arr_pass_orig_diap_izm_LIPA[i];
            W.ActiveDocument.Tables.Item(1).Cell(i+1,5).Range.text:=Finish_arr_pass_orig_tip_datch_LIPA[i];
            InsertRowTableDoc(tablica_,i+2);
          End;
        W.ActiveDocument.Tables.Item(tablica_).Rows.Item(i+2).Delete;
        W.ActiveDocument.Tables.Item(tablica_).Rows.Item(i+1).Delete;

        MessageBox(Handle,PChar('��������� ���... '), '�������', MB_OK or MB_TOPMOST);
        //Showmessage('��������� ���... ');
        SaveDocAs('c:\��������� ������������ ��� ����-4 (����-4)\SAVE\���_4-'+Num_Pribor+'.doc');
        CloseDoc;
        CloseWord;
      end;
  //Finish_arr_pass_orig:=nil;
  //Finish_arr_pass_orig_formula:=nil;
  //Finish_arr_pass_orig_diap_izm:=nil;
  //Finish_arr_pass_orig_tip_datch:=nil;

  //Finish_arr_pass_orig_LIPA:=nil;
  //Finish_arr_pass_orig_formula_LIPA:=nil;
  //Finish_arr_pass_orig_diap_izm_LIPA:=nil;
  //Finish_arr_pass_orig_tip_datch_LIPA:=nil;

  //str_of_263:=nil;
  //str_of_263_formula:=nil;
  //str_of_263_diap_izm:=nil;
  //str_of_263_tip_datch :=nil;

  //data_263:=nil;

  //Arr_gpf_tabl_1:=nil;
  //Arr_gpf_tabl_2:=nil;
  //Arr_gpf:=nil;
  //P3orAB:=nil;

  SetLength(Finish_arr_pass_orig,0);
  SetLength(Finish_arr_pass_orig_formula,0);
  SetLength(Finish_arr_pass_orig_diap_izm,0);
  SetLength(Finish_arr_pass_orig_tip_datch,0);
  SetLength(Finish_arr_pass_orig_LIPA,0);
  SetLength(Finish_arr_pass_orig_formula_LIPA,0);
  SetLength(Finish_arr_pass_orig_diap_izm_LIPA,0);
  SetLength(Finish_arr_pass_orig_tip_datch_LIPA,0);
  SetLength(str_of_263,0);
  SetLength(str_of_263_formula,0);
  SetLength(str_of_263_diap_izm,0);
  SetLength(str_of_263_tip_datch,0);
  SetLength(data_263,0);
  SetLength(Arr_gpf_tabl_1,0);
  SetLength(Arr_gpf_tabl_2,0);
  SetLength(Arr_gpf,0);
  SetLength(P3orAB,0);

  Form1.Edit1.Text:='������� ��������';
  clik_button:=0;
  Form1.Memo1.Clear;

end;

procedure TForm1.Edit1Change(Sender: TObject);
begin
  //Num_Pribor:=Edit1.Text;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  Form1.Memo1.Clear;
  Parameters;
  Form1.Memo2.Clear;
  Form1.Memo3.Clear;
end;

end.

