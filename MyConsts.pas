unit MyConsts;

interface
Uses Registry, Windows,Forms;
Var
R: TRegistry;
implementation
initialization
R:=TRegistry.Create;
R.RootKey:=HKey_Current_User; //кажется это имя ключа
//R.OpenKey('SoftWare'+Application.ExeName,True);
R.OpenKey('SoftWare\K615',True); //а это его расположение
Finalization
R.CloseKey;
R.Free;
end.
