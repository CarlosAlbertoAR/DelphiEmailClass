# DelphiSimpleMail
Only file Pascal/Delphi unit form send mail, easy!

Created using Indy non-visual components

# 4 single Steps:

1 - Add and uses uMail.pas to your project
2 - Create var of type TEmail
3 - Configure SMTP and IMAP parameters
4 - Send Emails!


# Example:

var Email : TEmail
begin

  Email               :=  TEmail.Create;
  Email.User          := 'yourmail@yourprovider.com';
  Email.Password      := 'yourpassword';

  // SMTP parameters form sent Emails
  Email.SMTP.Host     := 'smtp.yourprovider.com';
  Email.SMTP.Port     := '587';
  Email.SMTP.AuthorizationReq := True or False;
  Email.SMTP.UseTLS   := 0;
  
  // Imap parameters, for store copie on sent folder (optional)
  Email.IMAP.Host := imap.yourprovider.com;
  Email.IMAP.Port := 143;
  Email.IMAP.StoreFolder := 'sent_folder_name';
  Email.IMAP.SaveToInboxSent := True;

  // Fill email 
  Email.Msg.FromEmail := 'yourmail@yourprovider.com';
  Email.Msg.FromName  := 'You Name';
  Email.Msg.DestEmail := 'emaildest@provider.com';

  Email.Msg.Subject := 'Delphi email';
  Email.Msg.Body    := 'Hello, this is as easy email on Delphi Language.';
                      
  Email.Msg.AttachmentsList.Add('c:\myfolder\log.txt');
  Email.Msg.AttachmentsList.Add('c:\myfolder\image.jpg');  
  
  // Send, return True or False
  Email.SendMail;
  
  // if Error or sucess return message text
  Showmessage(Email.LastMessage);
  
  Email.Free;
end;

# For multiple emails use "Email.New", this clear theses parameters:

  Email.DestEmail
  Email.CCList
  Email.BCCList
  Email.AttachmentsList



