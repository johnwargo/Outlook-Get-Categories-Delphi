{ *****************************************************************************
  Get Categories

  Delphi application that opens an OLE connection to Microsoft Outlook and
  lists the names (and ID) for each Category defined within the application.

  John M. Wargo
  September 27, 2015
  ***************************************************************************** }
unit main;

interface

uses
  ComObj, Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs,
  Vcl.StdCtrls, Vcl.ComCtrls;

type
  TForm1 = class(TForm)
    output: TMemo;
    StatusBar1: TStatusBar;
    procedure FormActivate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.FormActivate(Sender: TObject);
var
  category, outlook, ns: OLEVariant;
  i, numItems: Integer;

begin
  // initialize a connection to Outlook
  outlook := CreateOLEObject('Outlook.Application');
  // get the MAPI namespace
  ns := outlook.GetNamespace('MAPI');
  numItems := ns.Categories.Count;
  output.Lines.add(Format('Found %d items', [numItems]));
  if numItems > 0 then
  begin
    for i := 1 to numItems do
    begin
      category := ns.Categories.Item[i];
      // output is a TMemo control
      // category.Name is the name of the category
      // category.CategoryID is an internal, unique ID for the category
      output.Lines.add(Format('%d: %s: (%s)', [i, category.Name,
        category.CategoryID]));
    end;
  end;
end;

end.
