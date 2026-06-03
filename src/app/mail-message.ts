export interface MailMessage {
  id: string;
  fromAddress: string;
  fromName: string;
  subject: string;
  receivedDateTime: string;
  bodyPreview: string;
}
