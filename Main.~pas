unit Main;
//////////////////////////////--Расчет извещений на утверждении--///////////////////////
interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ExtCtrls, Menus, ComCtrls, OleCtrls, Buttons, ToolWin, Isp3,
  ActnList, ImgList, SHDocVw, Shellapi,ActiveX, MSHTML, OleServer, AccessXP,
  WebAdapt, WebComp, StrUtils, Registry;

const
  CM_HOMEPAGEREQUEST = WM_USER + $1000;

type
  TMainForm = class(TForm)
    StatusBar1: TStatusBar;
    MainMenu1: TMainMenu;
    File1: TMenuItem;
    Exit1: TMenuItem;
    View1: TMenuItem;
    NavigatorImages: TImageList;
    NavigatorHotImages: TImageList;
    LinksImages: TImageList;
    LinksHotImages: TImageList;
    CoolBar1: TCoolBar;
    ToolBar1: TToolBar;
    BackBtn: TToolButton;
    ForwardBtn: TToolButton;
    StopBtn: TToolButton;
    RefreshBtn: TToolButton;
    ToolBar2: TToolBar;
    ToolButton6: TToolButton;
    ToolButton5: TToolButton;
    ToolButton7: TToolButton;
    ToolButton8: TToolButton;
    URLs: TComboBox;
    Help1: TMenuItem;
    About1: TMenuItem;
    Toolbar3: TMenuItem;
    Statusbar2: TMenuItem;
    Go1: TMenuItem;
    Back1: TMenuItem;
    Forward1: TMenuItem;
    Stop1: TMenuItem;
    Refresh1: TMenuItem;
    N2: TMenuItem;
    ActionList1: TActionList;
    BackAction: TAction;
    ForwardAction: TAction;
    StopAction: TAction;
    RefreshAction: TAction;
    WebBrowser1: TWebBrowser;
    N1: TMenuItem;
    Raschet: TMenuItem;
    Rasch: TMenuItem;
    AccessApplication1: TAccessApplication;
    RaschMod: TMenuItem;
    RaschModItog: TMenuItem;
    RaschAZKRS: TMenuItem;
    Panel1: TPanel;
    ProgressBar1: TProgressBar;
    procedure Exit1Click(Sender: TObject);
    procedure About1Click(Sender: TObject);
    procedure StopClick(Sender: TObject);
    procedure URLsKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FormCreate(Sender: TObject);
    procedure LinksClick(Sender: TObject);
    procedure RefreshClick(Sender: TObject);
    procedure BackClick(Sender: TObject);
    procedure ForwardClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure Toolbar3Click(Sender: TObject);
    procedure Statusbar2Click(Sender: TObject);
    procedure BackActionUpdate(Sender: TObject);
    procedure ForwardActionUpdate(Sender: TObject);
    procedure WebBrowser1BeforeNavigate2(Sender: TObject;
      const pDisp: IDispatch; var URL, Flags, TargetFrameName, PostData,
      Headers: OleVariant; var Cancel: WordBool);
    procedure WebBrowser1DownloadBegin(Sender: TObject);
    procedure WebBrowser1DownloadComplete(Sender: TObject);
    procedure N1Click(Sender: TObject);
    procedure WebBrowser1NewWindow2(Sender: TObject; var ppDisp: IDispatch;
      var Cancel: WordBool);
    procedure RaschModClick(Sender: TObject);
    procedure RaschModItogClick(Sender: TObject);
    procedure RaschAZKRSClick(Sender: TObject);
    procedure WebBrowser1ProgressChange(Sender: TObject; Progress,
      ProgressMax: Integer);
    procedure WebBrowser1DocumentComplete(Sender: TObject;
      const pDisp: IDispatch; var URL: OleVariant);

  private
    HistoryIndex: Integer;
    HistoryList: TStringList;
    UpdateCombo: Boolean;
    FOleInPlaceActiveObject: IOleInPlaceActiveObject;  //для клавы ентер
    procedure FindAddress;
    procedure HomePageRequest(var message: tmessage); message CM_HOMEPAGEREQUEST;
    function GetFrame(FrameNo: Integer): IWebbrowser2;
    procedure MsgHandler(var Msg: TMsg; var Handled: Boolean); //для клавы ентер
    procedure Podgotovka; // подготовка
   end;

var
  MainForm: TMainForm;
  PriznakNew: integer;

 implementation

uses About,MyConsts;

{$R *.dfm}
///////////////////////Функция - оконо Access поверх других окон/////////////////////////////////////
function EnumAccessProc (Wd: HWnd; Param: LongInt): Boolean; stdcall; // Обязательно stdcall !!!
Begin
 If  Wd=FindWindow('MSAccess',nil) then // если это Access
        If not IsWindowVisible(WD) then       // если окно не видимо
            If  IsIconic(WD) then      // если окно свернуто
              If  isWindow(WD) then        // и вообще это - окно.
                 ShowWindow(WD, SW_SHOWMAXIMIZED); // развернем его.
  EnumAccessProc := TRUE; // продолжаем перебирать все окна системы.
end;

procedure TMainForm.Exit1Click(Sender: TObject);
begin
  Close; //выход из программы
end;

procedure TMainForm.FindAddress;// Называем процедуру поиска адреса
var
  Flags: OLEVariant;
begin
  Flags := 0;
  UpdateCombo := True;
  WebBrowser1.Navigate(WideString(Urls.Text), Flags, Flags, Flags, Flags);
end;

procedure TMainForm.About1Click(Sender: TObject); //Вызов справки о системе Windows
begin
  ShowAboutBox;
end;

//Открываю мою справку

procedure TMainForm.N1Click(Sender: TObject);
begin
shellexecute (00000180,'open','СправкаРИАПНРМ.chm',nil,nil,sw_shownormal);
end;

procedure TMainForm.StopClick(Sender: TObject);  //Остановить загрузку
begin
  WebBrowser1.Stop;
end;

procedure TMainForm.URLsKeyDown(Sender: TObject; var Key: Word; //Развернуть
  Shift: TShiftState);   //URL лист вниз
begin
  if Key = VK_Return then   // если нажата клава Enter
  begin
    FindAddress;  // перейти на выбранный адрес
  end;  
end;

procedure TMainForm.LinksClick(Sender: TObject); // Переход на ссылку КХП
begin
  if (Sender as TToolButton).Hint = '' then Exit;
  URLs.Text := (Sender as TToolButton).Hint;
  FindAddress;
end;

procedure TMainForm.RefreshClick(Sender: TObject);  //Обновить
begin
  FindAddress;
end;

procedure TMainForm.BackClick(Sender: TObject);   //Назад
   begin
If LeftStr(URLs.Text,46) = 'http://osn.vaz.ru/matnor/manager/work_ctpi.asp' Then
             begin
                       URLs.Text := HistoryList[2];
             end
else 
       URLs.Text := HistoryList[HistoryIndex - 1];

If LeftStr(URLs.Text,10) = 'javascript' Then
             begin
                       HistoryList.Delete(HistoryIndex);
                       URLs.Text := HistoryList[HistoryList.Count - 2];
             end;


       FindAddress;
   end;

procedure TMainForm.ForwardClick(Sender: TObject);   //Вперёд
begin
  URLs.Text := HistoryList[HistoryIndex + 1];
  FindAddress;
end;


procedure TMainForm.Toolbar3Click(Sender: TObject); // Клик по толбару
begin
  with Sender as TMenuItem do
  begin
    Checked := not Checked;
    Coolbar1.Visible := Checked;
  end;
end;

procedure TMainForm.Statusbar2Click(Sender: TObject);   // Клик по статус бару
begin
  with Sender as TMenuItem do
  begin
    Checked := not Checked;
    StatusBar1.Visible := Checked;
  end;
end;

procedure TMainForm.HomePageRequest(var Message: TMessage); /////////Домашняя страница
begin
//URLs.Text := 'http://osn.vaz.ru/matnor/share/auth_new.asp?log_a=m3225000&passw_a=3225000';
    //Это пароль Новикова
  //URLs.Text := 'http://osn.vaz.ru/matnor/share/auth_new.asp?log_a=m3225106&passw_a=tktyf';
    //Это пароль Полякиной
   //URLs.Text := 'http://osn.vaz.ru/matnor/share/auth_new.asp?log_a=m3225205&passw_a=ssa2509';
   //это пароль Скидановой
  //URLs.Text := 'http://osn.vaz.ru/matnor/share/auth_new.asp?log_a=m3225204&passw_a=test2';
  //Это пароль Стуконоженко
  URLs.Text := 'http://osn.vaz.ru/matnor/DEFAULT.ASP';
  FindAddress;
end;

procedure TMainForm.FormCreate(Sender: TObject); //событие - создание формы
begin
          Application.OnMessage := MsgHandler; //для клавы ентер
  with ProgressBar1 do   //Помещаю прогрессбар в статус бар
  begin
    Parent := StatusBar1;
    Position := 100;
    Top := 4;
    Left := 0;
    Height := StatusBar1.Height - Top;
    Width := StatusBar1.Panels[0].Width + Left +100;
  end;

  with Panel1 do   //Помещаю панель в статус бар
  begin
    Parent := StatusBar1;
    Top := 4;
    Left := 0;
    Height := StatusBar1.Height - Top;
    Width := StatusBar1.Panels[0].Width + Left +100;
  end;
         // Создание исторического индекса и листа
  HistoryIndex := -1;
  HistoryList := TStringList.Create;

  if PriznakNew =1 then  //если новое окно было создано, то выход
   Exit
  Else
   PostMessage(Handle, CM_HOMEPAGEREQUEST, 0, 0); // иначе вернуть домашнюю страницу
end;

procedure TMainForm.FormDestroy(Sender: TObject);  //событие разрушение формы
begin
  HistoryList.Free;   // очистить исторический лист
   begin
     FOleInPlaceActiveObject := nil;  //для клавы ентер
   end;
end;

//////////////////////////////////////////////////////////////////////////////////////
/////////////для клавы ентер\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
procedure TMainForm.MsgHandler(var Msg: TMsg; var Handled: Boolean);
 const
   DialogKeys: set of Byte = [VK_LEFT, VK_RIGHT, VK_BACK,
     $30..$39, $41..42, $44..$55];
     {DialogKeys: set of Byte = [VK_LEFT, VK_RIGHT, VK_BACK, VK_UP, VK_DOWN,
     $30..$39, $41..42, $44..$55, $57, $59..$5A];}
 var
   iOIPAO: IOleInPlaceActiveObject;
   Dispatch: IDispatch;
 begin
   { exit if we don't get back a webbrowser object }
   if (WebBrowser1 = nil) then
   begin
     Handled := System.False;
     Exit;
   end;
   Handled := (IsDialogMessage(WebBrowser1.Handle, Msg) = System.True);
   if (Handled) and (not WebBrowser1.Busy) then
   begin
     if FOleInPlaceActiveObject = nil then
     begin
       Dispatch := WebBrowser1.Application;
       if Dispatch <> nil then
       begin
         Dispatch.QueryInterface(IOleInPlaceActiveObject, iOIPAO);
         if iOIPAO <> nil then
           FOleInPlaceActiveObject := iOIPAO;
       end;
     end;
     if FOleInPlaceActiveObject <> nil then
       if ((Msg.message = WM_KEYDOWN) or (Msg.message = WM_KEYUP)) and
         (Msg.wParam in DialogKeys) then
         // ничего не происходит и не передает DialogKeys
      else
         FOleInPlaceActiveObject.TranslateAccelerator(Msg);
   end;
 end;
/////////////////////////////////////////////////////////////////////////////////////////////


procedure TMainForm.BackActionUpdate(Sender: TObject);
begin
  if HistoryList.Count > 0 then
    BackAction.Enabled := HistoryIndex > 0
    //BackAction.Enabled := True
 else
    BackAction.Enabled := False;
end;

procedure TMainForm.ForwardActionUpdate(Sender: TObject);
begin
  if HistoryList.Count > 0 then
    ForwardAction.Enabled := HistoryIndex < HistoryList.Count - 1
  else
    ForwardAction.Enabled := False;
end;
//////////////////////////Самое главное свойство - навигация по URL\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
procedure TMainForm.WebBrowser1BeforeNavigate2(Sender: TObject;
  const pDisp: IDispatch; var URL, Flags, TargetFrameName, PostData,
  Headers: OleVariant; var Cancel: WordBool);
  var
  NewIndex: Integer;
begin
  NewIndex := HistoryList.IndexOf(URL);
  if NewIndex = -1 then
      begin
           { Удалите входы в HistoryList между последним(прошлым) адресом и текущим адресом }
           if (HistoryIndex >= 0) and (HistoryIndex < HistoryList.Count - 1) then
      while HistoryList.Count > HistoryIndex do
            HistoryList.Delete(HistoryIndex);
            HistoryIndex := HistoryList.Add(URL);
            //HistoryList.Sort;
      end
           else
            HistoryIndex := NewIndex;
                  if UpdateCombo then
                     begin
                          UpdateCombo := False;
                          NewIndex := URLs.Items.IndexOf(URL);
                                   if NewIndex = -1 then
                                      URLs.Items.Insert(0, URL)
                                   else
                                   URLs.Items.Move(NewIndex, 0);
                    end;
              URLs.Text := URL;
   If LeftStr(URLs.Text,54) = 'http://osn.vaz.ru/matnor/manager/nbmdoc_ctpi.asp?ndoc=' Then
  //////////если УРЛ соответсвует раскрытому извещению на утверждении
      Raschet.Visible:=True // доступна кнопка расчет
   else
  //Raschet.Visible:=False; // иначе не доступна кнопка расчет
   If LeftStr(URLs.Text,59) ='http://osn.vaz.ru/matnor/manager/nbmdoc_back_ctpi.asp?ndoc=' Then
  //////////если УРЛ соответсвует раскрытому извещению на возврате
      Raschet.Visible:=True // доступна кнопка расчет
   else
      Raschet.Visible:=False; // иначе не доступна кнопка расчет
end;


procedure TMainForm.WebBrowser1DownloadBegin(Sender: TObject);//событие перед загрузкой документа

 var
  iDoc: IHtmlDocument2;
  i: integer;
  ov: OleVariant;
  iDisp: IDispatch;
  iColl: IHTMLElementCollection;
  iInputElement: IHTMLInputElement;
  Y:string;

////////////////////////Записываю в регистр логин и пароль///////////
 begin
              StopBtn.ImageIndex := 6; ////Пока продолжается загрузка броузера, можно остановить
      if (HistoryIndex >= 1) then
          begin
              WebBrowser1.ControlInterface.Document.QueryInterface(IHtmlDocument2, iDoc);
                 if not assigned(iDoc) then // Если HtmlDocument не существует, то сообщение об ошибке
                     begin
                         ShowMessage(' !!!??? Nothing dowloaded ... ');
                         Exit;
                     end;
                 ov := 'INPUT';
                 IDisp := iDoc.all.tags(ov);//ищу тэги, содержащие INPUT как OleVariant
                 if assigned(IDisp) then //если такой тэг существует, то ...
                     begin
                         IDisp.QueryInterface(IHTMLElementCollection, iColl);//присвоить тэгу итерфейс коллекцион
                          if assigned(iColl) then // если такая структура интерфейса существует, то ...
                              begin
                                    for i := 1 to iColl.Get_length do // цикл от 1 до последнего структурного интерфейса
                                          begin
                                                iDisp := iColl.item(pred(i), 0); //обращаемся к предидущему тэгу начиная с нулевого
                                                iDisp.QueryInterface(IHTMLInputElement, iInputElement); //присваеваем элементу структуру вводимого элемента
                                                    if assigned(iInputElement) then // если такой элемент существует, то...
                                                         begin
                                                              Y:=iInputElement.Get_name;//присвоить переменной имя элемента
                                                                   if (Y ='log_a')  then //если элемент с именем "log_a"
                                                                   R.WriteString('Edit1',iInputElement.Get_value);
                                                                   if (Y ='passw_a')  then //если элемент с именем "passw_a"
                                                                   R.WriteString('Edit2',iInputElement.Get_value);
                                                          end;
                                           end;
                               end;
                       end;
          end;
 end;

procedure TMainForm.WebBrowser1DownloadComplete(Sender: TObject); //Пока продолжается загрузка
                              // документа, можно остановить
  begin
  StopBtn.ImageIndex := 2;
  end;
///////////////////////////////////////////////////////////////////////////////////////////////////
//Функция для кнопки сохранения
function TMainForm.GetFrame(FrameNo: Integer): IWebbrowser2;
 var
   OleContainer: IOleContainer;
   enum: IEnumUnknown;
   unk: IUnknown;
   Fetched: PLongint;
 begin
   while Webbrowser1.ReadyState <> READYSTATE_COMPLETE do
     Application.ProcessMessages;
   if Assigned(Webbrowser1.document) then
   begin
     Fetched := nil;
     OleContainer := Webbrowser1.Document as IOleContainer;
     OleContainer.EnumObjects(OLECONTF_EMBEDDINGS, Enum);
     Enum.Skip(FrameNo);
     Enum.Next(1, Unk, Fetched);
     Result := Unk as IWebbrowser2;
   end
   else
     Result := nil;
 end;
///////////////////////////////////////////////////////////////////////////////////////////////////

procedure TMainForm.WebBrowser1NewWindow2(Sender: TObject;//открыть новое окно в моём броузере
  var ppDisp: IDispatch; var Cancel: WordBool);
var
  NewWindow: TMainForm;
  begin
  PriznakNew:=1;    //признак нового окна
  NewWindow := TMainForm.Create(self);
  NewWindow.Show;
  ppDisp := NewWindow.Webbrowser1.DefaultDispatch;
  Application.OnMessage := MsgHandler; //для клавы ентер
end;

//Отображать движение TProgressBar при помощи фонового потока для ACCESS
//////////////////////////////////////////Создание анимации\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
type
   TAnimationThread = class(TThread)
   private
     { Private declarations }
     FWnd: HWND;
     FPaintRect: TRect;
     FbkColor, FfgColor: TColor;
     FInterval: integer;
   protected
     procedure Execute; override;
   public
     constructor Create(paintsurface : TWinControl; {Control to paint on }
       paintrect : TRect;          {area for animation bar }
       bkColor, barcolor : TColor; {colors to use }
       interval : integer);       {wait in msecs between
paints}
   end;


 constructor TAnimationThread.Create(paintsurface : TWinControl;
   paintrect : TRect; bkColor, barcolor : TColor; interval : integer);
 begin
   inherited Create(True);
   FWnd := paintsurface.Handle;
   FPaintRect := paintrect;
   FbkColor := bkColor;
   FfgColor := barColor;
   FInterval := interval;
   FreeOnterminate := True;
   Resume;
 end; { TAnimationThread.Create }

 procedure TAnimationThread.Execute;
 var
   image : TBitmap;
   DC : HDC;
   left, right : integer;
   increment : integer;
   imagerect : TRect;
   state : (incRight, incLeft, decLeft, decRight);
 begin
   Image := TBitmap.Create;
   try
     with Image do
      begin
       Width := FPaintRect.Right - FPaintRect.Left;
       Height := FPaintRect.Bottom - FPaintRect.Top;
       imagerect := Rect(0, 0, Width, Height);
     end; { with }
     left := 0;
     right := 0;
     increment := imagerect.right div 50;
     state := Low(State);
     while not Terminated do
      begin
       with Image.Canvas do
        begin
         Brush.Color := FbkColor;
         FillRect(imagerect);
         case state of
           incRight:
            begin
             Inc(right, increment);
             if right > imagerect.right then
              begin
               right := imagerect.right;
               Inc(state);
             end; { if }
           end; { Case incRight }
           incLeft:
            begin
             Inc(left, increment);
             if left >= right then
              begin
               left := right;
               Inc(state);
             end; { if }
           end; { Case incLeft }
           decLeft:
            begin
             Dec(left, increment);
             if left <= 0 then
              begin
               left := 0;
               Inc(state);
             end; { if }
           end; { Case decLeft }
           decRight:
            begin
             Dec(right, increment);
             if right <= 0 then
              begin
               right := 0;
               state := incRight;
             end; { if }
           end; { Case decLeft }
         end; { Case }
         Brush.Color := FfgColor;
         FillRect(Rect(left, imagerect.top, right, imagerect.bottom));
       end; { with }
       DC := GetDC(FWnd);
       if DC <> 0 then
         try
           BitBlt(DC,
             FPaintRect.Left,
             FPaintRect.Top,
             imagerect.right,
             imagerect.bottom,
             Image.Canvas.handle,
             0, 0,
             SRCCOPY);
         finally
           ReleaseDC(FWnd, DC);
         end;
       Sleep(FInterval);
     end; { While }
   finally
     Image.Free;
   end;
   InvalidateRect(FWnd, nil, True);
 end; { TAnimationThread.Execute }
//////////////////////////////////////////Конец Создания анимации\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

///////////////Процедуры вызова Access для расчета извещений, находящихся на утверждении или возврате\\\\\\\\\\\

//////////////////Расчет по моделям \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
procedure TMainForm.RaschModClick(Sender: TObject);
 var
   ani : TAnimationThread;
   r : TRect;
 begin
     Podgotovka ;//вызов процедуры подготовка для сохранения страницы на диск С:\
     panel1.Visible :=True;
     r := panel1.clientrect;
     InflateRect(r, - panel1.bevelwidth, - panel1.bevelwidth);
     ani := TanimationThread.Create(panel1, r, panel1.color, clRed, 150);
     RaschMod.Enabled := False;
     Application.ProcessMessages;
     Sleep(3000);  // replace with query.Open or such
     RaschMod.Enabled := True;
//Запускаю макрос для расчета
     AccessApplication1.OpenCurrentDatabaseOld('C:\АрхивПНРМ\ВедениеАрхиваПНРМ.mdb',True);
     EnumWindows (@EnumAccessProc, 0); // отрабатываем оконо Access поверх окна программы.
     AccessApplication1.DoCmd.RunMacro('Импорт_расчет',1,1);
        If AccessApplication1.DCount('*','Усл_год итоги все')=0 then
// Если в запросе Усл_год итоги все нет ни одной строки то
      begin
          ShowMessage('Нет деталей, соответсвующих запросу извещения!');
          ani.Terminate;
          AccessApplication1.CloseCurrentDatabase;
      end
       Else
      begin
          AccessApplication1.DoCmd.Maximize;
          AccessApplication1.Visible:= True;//Отображаю Access на экран
          ani.Terminate;
      end;
 end;
//////////////////Расчет по моделям Итоги\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
procedure TMainForm.RaschModItogClick(Sender: TObject);
 var
   ani : TAnimationThread;
   r : TRect;
 begin
     Podgotovka ;//вызов процедуры подготовка для сохранения страницы на диск С:\
     panel1.Visible :=True;
     r := panel1.clientrect;
     InflateRect(r, - panel1.bevelwidth, - panel1.bevelwidth);
     ani := TanimationThread.Create(panel1, r, panel1.color, clRed, 150);
     RaschModItog.Enabled := False;
     Application.ProcessMessages;
     Sleep(3000);  // replace with query.Open or such
     RaschModItog.Enabled := True;
//Запускаю макрос для расчета
     AccessApplication1.OpenCurrentDatabaseOld('C:\АрхивПНРМ\ВедениеАрхиваПНРМ.mdb',True);
     EnumWindows (@EnumAccessProc, 0); // отрабатываем оконо Access поверх окна программы.
     AccessApplication1.DoCmd.RunMacro('Импорт_расчет_итог',1,1);
       If AccessApplication1.DCount('*','Усл_год итоги все')=0 then
// Если в запросе Усл_год итоги все нет ни одной строки то
          begin
              ShowMessage('Нет деталей, соответсвующих запросу извещения!');
              ani.Terminate;
              AccessApplication1.CloseCurrentDatabase;
          end
        Else
          begin
              AccessApplication1.DoCmd.Maximize;
              AccessApplication1.Visible:= True;//Отображаю Access на экран
              ani.Terminate;
          end;
 end;
//////////////////Расчет АЗКРС\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
procedure TMainForm.RaschAZKRSClick(Sender: TObject);
 var
   ani : TAnimationThread;
   r : TRect;
 begin
      Podgotovka ;//вызов процедуры подготовка для сохранения страницы на диск С:\
      panel1.Visible :=True;
      r := panel1.clientrect;
      InflateRect(r, - panel1.bevelwidth, - panel1.bevelwidth);
      ani := TanimationThread.Create(panel1, r, panel1.color, clRed, 150);
      Application.ProcessMessages;
      Sleep(10000);  // replace with query.Open or such
//Запускаю макрос для расчета
      AccessApplication1.OpenCurrentDatabaseOld('C:\АрхивПНРМ\ВедениеАрхиваПНРМ.mdb',True);
      EnumWindows (@EnumAccessProc, 0); // отрабатываем оконо Access поверх окна программы.
      AccessApplication1.DoCmd.RunMacro('Импорт_расчет_АЗКРС',1,1);
        If AccessApplication1.DCount('*','Анализ')=0 then
// Если в запросе Анализ нет ни одной строки то
          begin
              ani.Terminate;
              ShowMessage('Ошибка!Нет деталей, соответсвующих запросу извещения!');

              AccessApplication1.Quit(1);
          end
         Else
          begin
              AccessApplication1.DoCmd.Maximize;
              AccessApplication1.Visible:= True;//Отображаю Access на экран
              ani.Terminate;
          end;
end;
////////////////////////////Обработка прогресса\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
procedure TMainForm.WebBrowser1ProgressChange(Sender: TObject; Progress,ProgressMax: Integer);
              //прогресс броузера записываю в прогресс на статус баре
            begin
                If Progress =-1 then Exit
              Else
                If ProgressMax <> 0 then
                   Progressbar1.Position:= Progress*100 Div ProgressMax; // прогресс=прогресс броузера*100/макс прогресс
                If ProgressMax = 0 then
                   Progressbar1.Position:=0;
                   Exit
            end;
////////////////////////Процедура Подготовка\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
procedure TMainForm.Podgotovka;//сохраняю страницу из
//веброузер на диск
 var
   IpStream: IPersistStreamInit;
   AStream: TMemoryStream;
   iw: IWebbrowser2;
   i: Integer;
   sl: TStringList;
   fs: TFileStream;
   S: string;
   HTMLDocument: IHTMLDocument2;
   PersistFile: IPersistFile;

  begin
     Raschet.Visible:=False;
     for i := 0 to Webbrowser1.OleObject.Document.frames.Length - 1 do
       begin
           iw := GetFrame(i);
           AStream := TMemoryStream.Create;
        try
        IpStream := iw.document as IPersistStreamInit;
             if Succeeded(IpStream.save(TStreamadapter.Create(AStream), True)) then
                 begin
                     AStream.SaveToFile('c:\frame' + IntToStr(i) + '.htm');
                     AStream.Seek(0, 0);
                     sl := TStringList.Create;
                     sl.LoadFromStream(AStream);
                     sl.SaveToFile('c:\frame' + IntToStr(i) + '.htm');
                     sl.Free;
                 end;
           except
           end;
        AStream.Free;
        end;
           begin
               fs := TFileStream.Create('c:\frame1.htm', fmOpenread or fmShareDenyNone);
               try
                   SetLength(S, fs.Size);
                   fs.ReadBuffer(S[1], fs.Size);
               finally
               fs.Free;
             end;
  //Заменяю '<html>' только при первом вхождении в тексте на '<!--'
                        S  := StringReplace(s, '<html>', '<!--  <!--',[rfIgnoreCase]);
 //Удаляю подстроку '<input name="kprmat_z' во всем тексте чтобы получить код причины
                        S  := StringReplace(s, '<input name="kprmat_z', '  ',[rfReplaceAll, rfIgnoreCase]);
                        fs := TFileStream.Create('c:\frame2.htm', fmCreate);
                   try
                     fs.WriteBuffer(S[1], Length(S));
                    finally
                    fs.Free;
                    end;
                 //fs.Free;
               end;
 //загружаю в броузер страницу с номером извещения
         if FileExists('c:\frame1.htm') = true then
         webBrowser1.Navigate('c:\frame1.htm');
 //сохраняю страницу в виде текста чтобы потом получить номер извещения
           while WebBrowser1.ReadyState < READYSTATE_INTERACTIVE do
           Application.ProcessMessages;
 //когда страница загружена полностью, то...
           try
                   HTMLDocument := WebBrowser1.Document as IHTMLDocument2;
                   PersistFile  := HTMLDocument as IPersistFile;
                   PersistFile.Save(StringToOleStr('c:\test2.txt'), System.True);
                   Sleep(5000);  // replace with query.Open or such
                   Webbrowser1.GoBack; //возврат в КМН
           Except
                    ShowMessage('Ахтунг! Вызвать разработчика! Ошибка test2.txt');
           end;

  //Sleep(5000);  // replace with query.Open or such
  Webbrowser1.Refresh;  //обновить броузер

 end;
///////////////////////////////////////////////////////////////////////////////////////////////////
procedure TMainForm.WebBrowser1DocumentComplete(Sender: TObject;
  const pDisp: IDispatch; var URL: OleVariant);
var
  iDoc: IHtmlDocument2;
  i: integer;
  ov: OleVariant;
  iDisp: IDispatch;
  iColl: IHTMLElementCollection;
  iInputElement: IHTMLInputElement;
  Y:string;

////////////////////////Вставляю из регистра логин и пароль///////////

  begin
      WebBrowser1.ControlInterface.Document.QueryInterface(IHtmlDocument2, iDoc);
      if not assigned(iDoc) then // Если HtmlDocument не существует, то сообщение об ошибке
          begin
               ShowMessage(' !!!??? Nothing dowloaded ... ');
               Exit;
           end;
      ov := 'INPUT';
      IDisp := iDoc.all.tags(ov);//ищу тэги, содержащие INPUT как OleVariant
      if assigned(IDisp) then //если такой тэг существует, то ...
          begin
              IDisp.QueryInterface(IHTMLElementCollection, iColl);//присвоить тэгу итерфейс коллекцион
               if assigned(iColl) then // если такая структура интерфейса существует, то ...
                 begin
                     for i := 1 to iColl.Get_length do // цикл от 1 до последнего структурного интерфейса
                         begin
                                iDisp := iColl.item(pred(i), 0); //обращаемся к предидущему тэгу начиная с нулевого
                                iDisp.QueryInterface(IHTMLInputElement, iInputElement); //присваеваем элементу структуру вводимого элемента
                              if assigned(iInputElement) then // если такой элемент существует, то...
                                          begin
                                                Y:=iInputElement.Get_name;//присвоить переменной имя элемента
                                                if (Y ='log_a')  then //если элемент с именем "log_a"
                                                           iInputElement.value:=R.ReadString('Edit1');
                                                if (Y ='passw_a')  then //если элемент с именем "passw_a"
                                                           iInputElement.value:=R.ReadString('Edit2');
                                          end;
                         end;
                 end;
          end;
 end;
/////////////////////////////////////////////////////////////////////////////////

initialization
   OleInitialize(nil);  //для клавы ентер

 finalization
   OleUninitialize;

   end.
