{******************************************************************}
{		CATO Technology - Brazil
{
{Class Name: uEmail
{Class Purpose: Complete and encapsulated send email function with Indy Classes
{Programmer: Carlos Alberto
{Data: 01/07/2014
{Versão: 1.2
{Additional comments:
{******************************************************************}

unit Email;

interface
uses
  IdComponent, IdTCPConnection, IdTCPClient, IdMessageClient, IdSMTP, IdIOHandler,
  IdBaseComponent, IdMessage, IdExplicitTLSClientServerBase, IdSMTPBase,IdAttachmentFile,
  IdIOHandlerSocket, IdIOHandlerStack, IdSSL, IdSSLOpenSSL, IdIMAP4, SysUtils, Classes;


type TMessage = Class
   private
    IdMessage : TIdMessage;

    FFromEmail: string;
    FBody: string;
    FFromName: string;
    FReceiptNotice: boolean;
    FSubject: string;
    FBCCList: string;
    FDestEmail: string;
    FPriority: Integer;
    FAttachmentsList: TStringList;
    FCCList: string;
    procedure SetAttachmentsList(const Value: TStringList);
    procedure SetBody(const Value: string);
    procedure SetBCCList(const Value: string);
    procedure SetDestEmail(const Value: string);
    procedure SetFromEmail(const Value: string);
    procedure SetFromName(const Value: string);
    procedure SetPriority(const Value: Integer);
    procedure SetReceiptNotice(const Value: boolean);
    procedure SetSubject(const Value: string);
    procedure SetCCList(const Value: string);

   protected

   public
      constructor Create;
      destructor Destroy;

      property FromEmail       : string read FFromEmail write SetFromEmail;
      property FromName        : string read FFromName write SetFromName;
      property DestEmail       : string read FDestEmail write SetDestEmail;
      property CCList          : string read FCCList write SetCCList;
      property BCCList         : string read FBCCList write SetBCCList;
      property Body            : string read FBody write SetBody;
      property Subject         : string read FSubject write SetSubject;
      property AttachmentsList : TStringList read FAttachmentsList write SetAttachmentsList;
      property Priority        : Integer read FPriority write SetPriority default 2;
      property ReceiptNotice   : boolean read FReceiptNotice write SetReceiptNotice default False;

      function Configure(AFromEmail : string;
                         AFromName : string;
                         ADestEmail : string;
                         ABCCList : string;
                         ASubject : string;
                         ABody : string;
                         AAttachmentsList : TStringList;
                         AReceiptNotice : boolean ) : boolean;

end;

type TIMAP = Class
   private
    IdImap4 : TIdImap4;
    AuthSSL : TIdSSLIOHandlerSocketOpenSSL;
    Log : TStringList;

    FReadTimeout: Integer;
    FPort: String;
    FSaveToInboxSent: Boolean;
    FConnectTimeout: Integer;
    FPassword: String;
    FHost: String;
    FUser: String;
    FStoreFolder: String;
    FLastMessage: String;
    FUseTLS: Integer;
    procedure SetConnectTimeout(const Value: Integer);
    procedure SetHost(const Value: String);
    procedure SetPassword(const Value: String);
    procedure SetPort(const Value: String);
    procedure SetReadTimeout(const Value: Integer);
    procedure SetSaveToInboxSent(const Value: Boolean);
    procedure SetUser(const Value: String);
    procedure SetStoreFolder(const Value: String);
    procedure SetLastMessage(const Value: String);
    procedure SetUseTLS(const Value: Integer);

    procedure IdImap4Connected(Sender: TObject);
    procedure IdImap4Disconnected(Sender: TObject);
    procedure IdImap4Status(ASender: TObject; const AStatus: TIdStatus; const AStatusText: string);

    procedure AuthSSLStatus(ASender: TObject; const AStatus: TIdStatus; const AStatusText: string);
    procedure AuthSSLStatusInfo(const AMsg: string);

    procedure AddMsg(AMsg : String);
    procedure ClearMsg;

   protected

   public
      constructor Create;
      destructor Destroy;

      property Host : String read FHost write SetHost;
      property Port : String read FPort write SetPort;
      property User : String read FUser write SetUser;
      property Password : String read FPassword write SetPassword;
      property UseTLS : Integer read FUseTLS write SetUseTLS default 0;
      property ConnectTimeout : Integer read FConnectTimeout write SetConnectTimeout {default 30000};
      property ReadTimeout : Integer read FReadTimeout write SetReadTimeout {default 30000};
      property StoreFolder : String read FStoreFolder write SetStoreFolder;
      property SaveToInboxSent : Boolean read FSaveToInboxSent write SetSaveToInboxSent default False;
      property LastMessage : String read FLastMessage write SetLastMessage;

      function AppendMsg(AMsg: TIdMessage; const AFlags: TIdMessageFlagsSet = []) : Boolean;
      function ListMailBoxes(out AList : TStringList) : Boolean;
end;

type TSMTP = Class
   private
    IdSMTP : TIdSMTP;
    AuthSSL : TIdSSLIOHandlerSocketOpenSSL;

    FUseTLS: Integer;
    FPort: String;
    FPassword: String;
    FAuthorizationReq: Boolean;
    FHost: String;
    FUser: String;
    FLastMessage: String;
    procedure SetUseTLS(const Value: Integer);
    procedure SetHost(const Value: String);
    procedure SetAuthorizationReq(const Value: Boolean);
    procedure SetPassword(const Value: String);
    procedure SetPort(const Value: String);
    procedure SetUser(const Value: String);

    procedure AddMsg(AMsg : String);
    procedure ClearMsg;
    procedure SetLastMessage(const Value: String);

   protected

   public
      constructor Create;
      destructor Destroy;

      property Host : String read FHost write SetHost;
      property Port : String read FPort write SetPort;
      property User : String read FUser write SetUser;
      property Password : String read FPassword write SetPassword;
      property AuthorizationReq : Boolean read FAuthorizationReq write SetAuthorizationReq default True;
      property UseTLS : Integer read FUseTLS write SetUseTLS default 0;
      property LastMessage : String read FLastMessage write SetLastMessage;

      function SendMail(AMsg : TMessage; AIMAP : TIMAP) : Boolean;
end;


type TEmail = Class
   private
    FUser: string;
    FPassword: string;
    FIMAP: TIMAP;
    FSMTP: TSMTP;
    FMsg: TMessage;
    FLastMessage: String;

    procedure SetUser(const Value: string);
    procedure SetPassword(const Value: string);
    procedure SetIMAP(const Value: TIMAP);
    procedure SetSMTP(const Value: TSMTP);
    procedure SetMsg(const Value: TMessage);
    procedure SetLastMessage(const Value: String);

    procedure AddMsg(AMsg : String);
    procedure ClearMsg;

   protected

   public
      property User            : string      read FUser            write SetUser;
      property Password        : string      read FPassword        write SetPassword;

      property IMAP            : TIMAP       read FIMAP            write SetIMAP;
      property SMTP            : TSMTP       read FSMTP            write SetSMTP;
      property Msg             : TMessage    read FMsg             write SetMsg;

      property LastMessage     : String      read FLastMessage     write SetLastMessage;

      constructor Create;
      destructor Destroy;
      function SendMail : Boolean;
      procedure Configure;
      procedure New;

   published

end;

implementation

{ IMAP }

procedure TIMAP.AddMsg(AMsg: String);
begin
   if FLastMessage.IsEmpty then
      FLastMessage := AMsg
   else
      FLastMessage := FLastMessage + #13 + AMsg;
end;

function TIMAP.AppendMsg(AMsg: TIdMessage; const AFlags: TIdMessageFlagsSet): Boolean;
var
   Erro : Boolean;
begin
   try
      try
         if IdImap4.Host.IsEmpty then
            raise Exception.Create('IMAP host not specified.');

         if IdImap4.Port = 0 then
            raise Exception.Create('IMAP port not specified.');

         if IdImap4.Username.IsEmpty then
            raise Exception.Create('IMAP username not specified.');

         if IdImap4.Password.IsEmpty then
            raise Exception.Create('IMAP password not specified.');

         if Self.StoreFolder.IsEmpty then
            raise Exception.Create('IMAP folder not specified.');

      except on E: Exception do
         begin
            Erro := True;
            AddMsg(E.Message);
         end;
      end;

      IdImap4.Connect(True);
      if IdImap4.Connected then
         if IdImap4.AppendMsg(self.StoreFolder, AMsg, AFlags) then
            Result := True
         else begin
            Result := False;
            AddMsg('Failed to send copy of message to folder ' + Self.StoreFolder);
            //AddMsg(IdImap4.LastCmdResult.Code + ' - '+ IdImap4.LastCmdResult.Text.Text);
         end;

      IdImap4.Disconnect(True);
   finally
      Result := not Erro;
      //ShowMessage(Log.Text);
   end;
end;

procedure TIMAP.AuthSSLStatus(ASender: TObject; const AStatus: TIdStatus;
  const AStatusText: string);
begin
   Log.Add('SSL Status : ' + AStatusText);
end;

procedure TIMAP.AuthSSLStatusInfo(const AMsg: string);
begin
   Log.Add('SSL Info : ' + AMsg);
end;

procedure TIMAP.ClearMsg;
begin
   FLastMessage := EmptyStr;
end;

constructor TIMAP.Create;
begin
  IdImap4 := TIdIMAP4.Create(nil);
  IdImap4.OnConnected    := IdImap4Connected;
  IdImap4.OnDisconnected := IdImap4Disconnected;
  IdImap4.OnStatus       := IdImap4Status;

  AuthSSL := TIdSSLIOHandlerSocketOpenSSL.Create(nil);
  AuthSSL.OnStatus := AuthSSLStatus;
  AuthSSL.OnStatusInfo := AuthSSLStatusInfo;

  Log := TStringList.Create;
end;

destructor TIMAP.Destroy;
begin
   if IdImap4 <> nil then
      FreeAndNil(IdImap4);

   if AuthSSL <> nil then
      FreeAndNil(AuthSSL);

   if Log <> nil then
      FreeAndNil(Log);
end;

procedure TIMAP.IdImap4Connected(Sender: TObject);
begin
   //Log.Add('IMAP : Connected SUCCESS');
end;

procedure TIMAP.IdImap4Disconnected(Sender: TObject);
begin
   //Log.Add('IMAP : DISCONNECTED');
end;

procedure TIMAP.IdImap4Status(ASender: TObject; const AStatus: TIdStatus;
  const AStatusText: string);
begin
   Log.Add('IMAP Status : ' + AStatusText);
end;

function TIMAP.ListMailBoxes(out AList: TStringList): Boolean;
var
   Erro : Boolean;
begin
   try
      Erro := False;

      if AList = nil then
         AList := TStringList.Create;

      IdImap4.Connect(True);
      if IdImap4.Connected then
         if not IdImap4.ListMailBoxes(AList) then begin
            Erro := True;
            AddMsg('Failed to list folders in the specified email account.');
            //AddMsg(IdImap4.LastCmdResult.Code + ' - '+ IdImap4.LastCmdResult.Text.Text);
         end;

      IdImap4.Disconnect(True);
   finally
      Result := not Erro;
      //ShowMessage(Log.Text);
   end;
end;

procedure TIMAP.SetConnectTimeout(const Value: Integer);
begin
  FConnectTimeout := Value;

  IdImap4.ConnectTimeout := FConnectTimeout;
end;

procedure TIMAP.SetHost(const Value: String);
begin
  FHost := Value;

  IdImap4.Host := FHost;
end;

procedure TIMAP.SetLastMessage(const Value: String);
begin
  FLastMessage := Value;
end;

procedure TIMAP.SetPassword(const Value: String);
begin
  FPassword := Value;

  IdImap4.Password := FPassword;
end;

procedure TIMAP.SetPort(const Value: String);
begin
  FPort := Value;

  IdImap4.Port := StrToInt(FPort);
end;

procedure TIMAP.SetReadTimeout(const Value: Integer);
begin
  FReadTimeout := Value;

  IdImap4.ReadTimeout := FReadTimeout;
end;

procedure TIMAP.SetSaveToInboxSent(const Value: Boolean);
begin
  FSaveToInboxSent := Value;
end;

procedure TIMAP.SetStoreFolder(const Value: String);
begin
  FStoreFolder := Value;
end;

procedure TIMAP.SetUser(const Value: String);
begin
  FUser := Value;

  IdImap4.Username := FUser;
end;

procedure TIMAP.SetUseTLS(const Value: Integer);
begin
   FUseTLS := Value;

   if FUseTLS <> 0 then begin
      IdImap4.IOHandler          := AuthSSL;
      case FUseTLS of
         0 : IdImap4.UseTLS := utNoTLSSupport;
         1 : IdImap4.UseTLS := utUseImplicitTLS;
         2 : IdImap4.UseTLS := utUseRequireTLS;
         3 : IdImap4.UseTLS := utUseExplicitTLS;
      end;

      AuthSSL.DefaultPort       := StrToInt(FPort);
      AuthSSL.SSLOptions.Method := sslvSSLv23; //sslvSSLv3;
      AuthSSL.SSLOptions.Mode   := sslmClient;
   end;
end;

{ Email }

procedure TEmail.AddMsg(AMsg: String);
begin
   if FLastMessage.IsEmpty then
      FLastMessage := AMsg
   else
      FLastMessage := FLastMessage + #13 + AMsg;
end;

procedure TEmail.New;
begin
  Msg.BCCList := EmptyStr;
  Msg.Subject := EmptyStr;
  Msg.Body := EmptyStr;
  Msg.AttachmentsList.Clear;
end;

procedure TEmail.ClearMsg;
begin
   self.FLastMessage := EmptyStr;
   SMTP.LastMessage := EmptyStr;
   IMAP.LastMessage := EmptyStr;
end;

constructor TEmail.Create;
begin
   SMTP := TSMTP.Create;
   IMAP := TIMAP.Create;
   Msg  := TMessage.Create;
end;

destructor TEmail.Destroy;
begin
   //if IMAP <> nil then
   //   FreeAndNil(IMAP);

   //if AttachmentsList <> nil then
   //   FreeAndNil(AttachmentsList);
end;

function TEmail.SendMail : Boolean;
var
   Erro          : Boolean;
   I             : Integer;
begin
   try
      Self.ClearMsg;
      Result := SMTP.SendMail(Self.Msg, Self.IMAP);
   finally
      Self.FLastMessage := SMTP.LastMessage;
   end;
end;

procedure TEmail.SetIMAP(const Value: TIMAP);
begin
  FIMAP := Value;
end;

procedure TEmail.SetLastMessage(const Value: String);
begin
  FLastMessage := Value;
end;

procedure TEmail.SetMsg(const Value: TMessage);
begin
  FMsg := Value;
end;

procedure TEmail.SetPassword(const Value: string);
begin
  FPassword := Value;
  SMTP.Password := FPassword;
  IMAP.Password := FPassword;
end;

procedure TEmail.SetSMTP(const Value: TSMTP);
begin
  FSMTP := Value;
end;

procedure TEmail.SetUser(const Value: string);
begin
   FUser := Value;
   SMTP.User := FUser;
   IMAP.User := FUser;
   Msg.FromEmail := FUser;
end;

procedure TEmail.Configure;
var
   Sufix : String;
begin
   if not FUser.IsEmpty then begin
      Msg.FromEmail := FUser;

      Sufix := Copy(FUser, Pos('@', FUser) + 1, (Length(FUser) - Pos('@', FUser)));
      SMTP.Host := 'smtp.' + Sufix;
      SMTP.Port := '587';

      IMAP.Host := '';
      //IMAP.Port := '';

      SMTP.AuthorizationReq := True;
      SMTP.UseTLS := 0;
      IMAP.UseTLS := 0;

      {
      0 - utNoTLSSupport
      1 - utUseImplicitTLS, // ssl iohandler req, allways tls
      2 - utUseRequireTLS, // ssl iohandler req, user command only accepted when in tls
      3 - utUseExplicitTLS // < user can choose to use tls
      }

      if (Pos('@locaweb', FUser) > 0) then begin
         SMTP.AuthorizationReq := True;
         SMTP.UseTLS := 0;

         IMAP.Host := 'email-ssl.com.br';
         IMAP.Port := '143';
         IMAP.UseTLS := 0;
      end;

      if (Pos('@hotmail', FUser) > 0) or (Pos('@outlook', FUser) > 0) or (Pos('@live', FUser) > 0) then begin
         SMTP.Host := 'smtp-mail.outlook.com';
         SMTP.AuthorizationReq := True;
         SMTP.UseTLS := 3;

         IMAP.Host := 'imap-mail.outlook.com';
         IMAP.Port := '993';
         IMAP.UseTLS := 1;
      end;

      if (Pos('@yahoo', FUser) > 0) then begin
         SMTP.Host := 'smtp.mail.yahoo.com';
         SMTP.AuthorizationReq := True;
         SMTP.UseTLS := 3;

         IMAP.Host := 'imap.mail.yahoo.com';
         IMAP.Port := '993';
         IMAP.UseTLS := 1;
      end;

      if (Pos('@gmail', FUser) > 0) or (Pos('@brturbo', FUser) > 0) then begin
         SMTP.Host := 'smtp.gmail.com';
         SMTP.AuthorizationReq := True;
         SMTP.UseTLS := 3;

         IMAP.Host := 'imap.gmail.com';
         IMAP.Port := '993';
         IMAP.UseTLS := 1;
      end;
   end;
end;

{ SMTP }

constructor TSMTP.Create;
begin
   IdSMTP := TIdSMTP.Create(nil);
   IdSMTP.UseEhlo := True;
   IdSMTP.UseVerp := False;

   AuthSSL := TIdSSLIOHandlerSocketOpenSSL.Create(nil);
end;

destructor TSMTP.Destroy;
begin
  //if IdSMTP <> nil then
  // FreeAndNil(IdSMTP);

  //if AuthSSL <> nil then
  // FreeAndNil(AuthSSL);
end;

procedure TSMTP.AddMsg(AMsg: String);
begin
   if FLastMessage.IsEmpty then
      FLastMessage := AMsg
   else
      FLastMessage := FLastMessage + #13 + AMsg;
end;

procedure TSMTP.ClearMsg;
begin
   FLastMessage := EmptyStr;
end;

function TSMTP.SendMail(AMsg : TMessage; AIMAP : TIMAP) : Boolean;
var
   Erro   : boolean;
   msgRet : string;
   i : integer;
begin
   try
      Erro := False;

      try
         if AMsg.DestEmail.IsEmpty then
            raise Exception.Create('Uninformed Email Recipient.');

         if IdSMTP.Host.IsEmpty then
            raise Exception.Create('SMTP host not specified.');

         if IdSMTP.Username.IsEmpty then
            raise Exception.Create('SMTP username not specified.');

         if IdSMTP.Password.IsEmpty then
            raise Exception.Create('SMTP password not specified.');

      except on E: Exception do
         begin
            Erro := True;
            msgRet := E.Message;
         end;
      end;

      AMsg.IdMessage.MessageParts.Clear;
      for I := 0 to AMsg.AttachmentsList.Count - 1 do
         TIdAttachmentFile.Create(AMsg.IdMessage.MessageParts, AMsg.AttachmentsList[I]) ;

      try
         idSMTP.ReadTimeout := 10000;
         idSMTP.Connect;
         idSMTP.Authenticate;
      except on E: Exception do
         begin
            Erro := True;
            msgRet := E.Message;
         end;
      end;

      if (idSMTP.Connected) and (idSMTP.DidAuthenticate) then begin
         try
            idSMTP.Send(AMsg.IdMessage);
            idSMTP.Disconnect;

            msgRet := 'Email successfully sent.';
         except
            on e: exception do begin
               Erro := True;
               msgRet := 'Email was not sent...' + #13 + #13 +
                         'Reason:'                  + #13 +
                         e.Message                              ;
            end;
         end;

      end
      else begin
         Erro := True;

         if not idSMTP.Connected then
            msgRet := 'Email not sent. Failed to connect.' + #13 + #13 +
                      'Reason:'                               + #13 +
                      msgRet
         else if not idSMTP.DidAuthenticate then
            msgRet := 'Email not sent. Failed to Authenticate.' + #13 + #13 +
                      'Reason:'                                 + #13 +
                      msgRet
      end;

      if not Erro then begin
         if AIMAP.SaveToInboxSent then
            if not AIMAP.AppendMsg(AMsg.IdMessage) then
               msgRet := msgRet + #13 + AIMAP.LastMessage;
      end;
   finally
      idSMTP.Disconnect;
      //AuthSSL.Free;

      AddMsg(msgRet);
      Result := not Erro;;
   end;
end;

procedure TSMTP.SetUseTLS(const Value: Integer);
begin
  FUseTLS := Value;

   if FUseTLS <> 0 then begin
      idSMTP.IOHandler     := AuthSSL;

      case FUseTLS of
         0 : idSMTP.UseTLS := utNoTLSSupport;
         1 : idSMTP.UseTLS := utUseImplicitTLS;
         2 : idSMTP.UseTLS := utUseRequireTLS;
         3 : idSMTP.UseTLS := utUseExplicitTLS;
      end;

      AuthSSL.DefaultPort        := StrToInt(FPort);
      AuthSSL.SSLOptions.Method  := sslvSSLv23; //sslvSSLv3;
      AuthSSL.SSLOptions.Mode    := sslmClient;
   end;
end;

procedure TSMTP.SetHost(const Value: String);
begin
  FHost := Value;

  IdSMTP.Host := FHost;
end;

procedure TSMTP.SetLastMessage(const Value: String);
begin
  FLastMessage := Value;
end;

procedure TSMTP.SetAuthorizationReq(const Value: Boolean);
begin
  FAuthorizationReq := Value;

   if FAuthorizationReq then
      idSMTP.AuthType := satDefault
   else
      idSMTP.AuthType := satNone;
end;

procedure TSMTP.SetPassword(const Value: String);
begin
  FPassword := Value;

  IdSMTP.Password := FPassword;
end;

procedure TSMTP.SetPort(const Value: String);
begin
  FPort := Value;

  IdSMTP.Port := StrToInt(FPort);
end;

procedure TSMTP.SetUser(const Value: String);
begin
  FUser := Value;

  IdSMTP.Username := FUser;
end;

{ TMensagem }

constructor TMessage.Create;
begin
   IdMessage := TIdMessage.Create(nil);
   //IdMessage.ContentType := 'text/html';
   //IdMessage.CharSet := 'ISO-8859-1';

   AttachmentsList := TStringList.Create;
   Priority := 2; {Prioridade Padrão}
end;

destructor TMessage.Destroy;
begin
   // if IdMessage <> nil then
   //    FreeAndNil(IdMessage);

   //if AttachmentsList <> nil then
   //    FreeAndNil(AttachmentsList);
end;

function TMessage.Configure(AFromEmail, AFromName, ADestEmail, ABCCList, ASubject, ABody : string;
  AAttachmentsList : TStringList; AReceiptNotice: boolean): boolean;
begin
   Result := True;
   try
      IdMessage.MessageParts.Clear;

      FFromEmail := AFromEmail;
      FFromName := AFromName;
      FDestEmail := ADestEmail;
      FBCCList := ABCCList;
      FSubject := ASubject;
      FBody := ABody;
      FAttachmentsList := AAttachmentsList;
      FReceiptNotice := AReceiptNotice;
   except
      Result := False;
   end;
end;

procedure TMessage.SetAttachmentsList(const Value: TStringList);
begin
  FAttachmentsList := Value;
end;

procedure TMessage.SetBody(const Value: string);
begin
  FBody := Value;

  IdMessage.Body.Text := FBody;
end;

procedure TMessage.SetCCList(const Value: string);
begin
  FCCList := Value;

  IdMessage.CCList.EMailAddresses := FCCList;
end;

procedure TMessage.SetBCCList(const Value: string);
begin
  FBCCList := Value;

  IdMessage.BCCList.EMailAddresses := FBCCList;
end;

procedure TMessage.SetDestEmail(const Value: string);
begin
  FDestEmail := Value;

  IdMessage.Recipients.EMailAddresses := FDestEmail;
end;

procedure TMessage.SetFromEmail(const Value: string);
begin
  FFromEmail := Value;

  IdMessage.From.Text := FFromEmail;
end;

procedure TMessage.SetFromName(const Value: string);
begin
  FFromName := Value;

  IdMessage.From.Name := FFromName;
end;

procedure TMessage.SetPriority(const Value: Integer);
begin
  FPriority := Value;

  IdMessage.Priority := TIdMessagePriority(FPriority);
end;

procedure TMessage.SetReceiptNotice(const Value: boolean);
begin
  FReceiptNotice := Value;

  if FReceiptNotice then
    IdMessage.ReceiptRecipient.Text := IdMessage.From.Text
  else
    IdMessage.ReceiptRecipient.Text := '';
end;

procedure TMessage.SetSubject(const Value: string);
begin
  FSubject := Value;

  IdMessage.Subject := FSubject;
end;

end.
