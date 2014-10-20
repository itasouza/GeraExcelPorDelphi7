unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, DB, StdCtrls, ADODB, ComObj;

type
  TForm1 = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    btnGeraExcel: TButton;
    procedure btnGeraExcelClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

uses Unit2;

{$R *.dfm}

procedure TForm1.btnGeraExcelClick(Sender: TObject);
var
  Excel, Sheet: Variant;
  cTitulo: string;
  i: integer;
  sTitCelula: String;
  sArqXLS: String;
begin
   //OBS: Voce deve usar a Clausula ComObj no USES para usar o EXCEL
  try
    cTitulo:= 'Relatório de Faturamento';   //titulo do relatório
    Excel:= CreateOleObject('Excel.Application');
    Excel.Visible := True;
    Excel.WorkBooks.Add;
    Excel.Workbooks[1].Sheets.Add;

    //deleta as planilhas default
//    Excel.Workbooks[1].WorkSheets['Plan1'].delete;
//    Excel.Workbooks[1].WorkSheets['Plan2'].delete;
//    Excel.Workbooks[1].WorkSheets['Plan3'].delete;

    //Cabeçalho
    Excel.Workbooks[1].WorkSheets[1].Name := cTitulo;

    Sheet:=Excel.Workbooks[1].WorkSheets[cTitulo];
    Sheet.Range['A1','F1'].font.name := 'Arial';            // Fonte
    Sheet.Range['A1','F1'].font.size := 10;                 // Tamanho da Fonte
    Sheet.Range['A1','F1'].font.bold := true;               // Negrito
    Sheet.Range['A1','F1'].font.italic := False;            // Italico
    Sheet.Range['A1','F1'].font.color := clNavy;            // Cor da Fonte
    Sheet.Range['A1','F1'].Interior.Color := $16776961;     // Cor da Célula

    // Alinhando as Células
    Sheet.Range['A1','F1'].VerticalAlignment   := 2;        // 1=Top - 2=Center - 3=Bottom
    Sheet.Range['A1','F1'].HorizontalAlignment := 3;        // 3=Center - 4=Right

    Sheet.Range['A2','F400'].font.name := 'Arial';          // Fonte
    Sheet.Range['A2','F400'].font.size := 10;               // Tamanho da Fonte
    Sheet.Range['A2','F400'].font.bold := false;            // Negrito
    Sheet.Range['A2','F400'].font.italic := False;          // Italico
    Sheet.Range['A2','F400'].font.color := clBlack;         // Cor da Fonte

    Sheet.Range['B2','D400'].HorizontalAlignment := 3;      // 3=Center - 4=Right
    Sheet.Range['E2','E400'].HorizontalAlignment := 3;      // 3=Center - 4=Right
    Sheet.Range['F2','F400'].HorizontalAlignment := 3;      // 3=Center - 4=Right
    Sheet.Range['F2','F400'].NumberFormat := 'dd/mm/aaaa';

    // Define o tamanho das Colunas (basta fazer em uma delas e as demais serão alteradas)
    Sheet.Range['A1'].ColumnWidth:= 15;
    Sheet.Range['B1'].ColumnWidth:= 50;
    Sheet.Range['C1'].ColumnWidth:= 15;
    Sheet.Range['D1'].ColumnWidth:= 15;
    Sheet.Range['E1'].ColumnWidth:= 35;
    Sheet.Range['F1'].ColumnWidth:= 15;

    with DM.qryConsulta do
    begin
      Close;
      Open;
      if Not IsEmpty then
      begin
        for i := 1 to FieldCount do
        begin
          sTitCelula := Fields[i-1].FullName;
          Excel.Workbooks[1].WorkSheets[1].Cells[1,i]:= sTitCelula;
        end;

        i:=2;
        while not Eof do
        begin
          Excel.Workbooks[1].Sheets[1].Cells[i,1] := Fields[0].AsString;
          Excel.Workbooks[1].Sheets[1].Cells[i,2] := Fields[1].AsString;
          Excel.Workbooks[1].Sheets[1].Cells[i,3] := Fields[2].AsString;
          Excel.Workbooks[1].Sheets[1].Cells[i,4] := Fields[3].AsString;
          Excel.Workbooks[1].Sheets[1].Cells[i,5] := Fields[4].AsString;
          Excel.Workbooks[1].Sheets[1].Cells[i,6] := Fields[5].AsString;

          Sheet.Range['A1','F'+IntToStr(i-1)].Borders.LineStyle := 1;
          Sheet.Range['A1','F'+IntToStr(i-1)].Borders.Weight := 2;
          Sheet.Range['A1','F'+IntToStr(i-1)].Borders.ColorIndex := 1;

          Inc( i );
          Next;
        end;
        Sheet.Range['A1','F'+IntToStr(i-1)].Borders.LineStyle := 1;
        Sheet.Range['A1','F'+IntToStr(i-1)].Borders.Weight := 2;
        Sheet.Range['A1','F'+IntToStr(i-1)].Borders.ColorIndex := 1;

       // Sheet.Rows.Rows[1].Insert;         // --> Insere uma nova linha na planilha.
       // Sheet.Columns.Columns[1].Insert;   // --> Insere uma nova coluna na planilha.
       // Sheet.Range['A1'].ColumnWidth := 6;

        sArqXLS := 'NomeRelatório_' + FormatDateTime('MMYYYY', Now()) + '.xls';

        if FileExists('C:\Componentes - Sistemas Fontes\Gera Excel\ ' + sArqXLS) then
           DeleteFile('C:\Componentes - Sistemas Fontes\Gera Excel\ ' + sArqXLS);
           Excel.Workbooks[1].SaveAs( 'C:\Componentes - Sistemas Fontes\Gera Excel\ ' + sArqXLS );
           Excel.Quit;
         end;

       end;

    // messagedlg( 'Arquivo gerado!', mtInformation, [mbok], 0 );
  except
    messagedlg( 'Erro ao exportar arquivo !', mterror, [mbok], 0 );
    Excel.Quit;
    Excel.ActiveDocument.Close(SaveChanges := 0);
  end;
  // Fonte para consulta de geração de Excel:
  // http://www.planetadelphi.com.br/dica/5329/exportando-para-o-excel


end;

end.
