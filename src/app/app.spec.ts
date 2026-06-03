import { TestBed } from '@angular/core/testing';
import { By } from '@angular/platform-browser';

import { App } from './app';
import { GraphMailService } from './graph-mail.service';
import { MailMessage } from './mail-message';

class GraphMailServiceStub {
  configured = true;
  signedInUser = 'admin@example.com';
  messages: MailMessage[] = [
    {
      id: 'message-1',
      fromAddress: 'sender@example.com',
      fromName: 'Sender Example',
      subject: 'Quarterly update',
      receivedDateTime: '2026-05-01T08:00:00Z',
      bodyPreview: 'Preview text',
    },
  ];
  deletedIds: string[] = [];

  isConfigured(): boolean {
    return this.configured;
  }

  getConfigurationMessage(): string {
    return 'Configure Azure first.';
  }

  async signIn(): Promise<string> {
    return this.signedInUser;
  }

  async signOut(): Promise<void> {
    return Promise.resolve();
  }

  async loadMessagesBySender(): Promise<MailMessage[]> {
    return this.messages;
  }

  async deleteMessages(messageIds: string[]): Promise<void> {
    this.deletedIds = messageIds;
  }
}

describe('App', () => {
  let service: GraphMailServiceStub;

  beforeEach(async () => {
    service = new GraphMailServiceStub();

    await TestBed.configureTestingModule({
      imports: [App],
      providers: [{ provide: GraphMailService, useValue: service }],
    }).compileComponents();
  });

  it('should create the app', () => {
    const fixture = TestBed.createComponent(App);
    const app = fixture.componentInstance;
    expect(app).toBeTruthy();
  });

  it('should render the mail management workflow', () => {
    const fixture = TestBed.createComponent(App);
    fixture.detectChanges();

    const compiled = fixture.nativeElement as HTMLElement;

    expect(compiled.querySelector('h1')?.textContent).toContain('Review and delete messages from a sender');
    expect(compiled.textContent).toContain('Find messages by sender');
    expect(compiled.textContent).toContain('Review and confirm deletion');
  });

  it('loads and deletes selected messages', async () => {
    const fixture = TestBed.createComponent(App);
    fixture.detectChanges();

    const app = fixture.componentInstance;
    await app.signIn();
    app.senderEmail = 'sender@example.com';

    await app.findMessages();
    fixture.detectChanges();

    expect(fixture.nativeElement.textContent).toContain('Quarterly update');

    const checkbox = fixture.debugElement.query(By.css('.message-card input[type="checkbox"]'));
    checkbox.nativeElement.click();
    fixture.detectChanges();

    spyOn(window, 'confirm').and.returnValue(true);

    await app.deleteSelectedMessages();
    fixture.detectChanges();

    expect(service.deletedIds).toEqual(['message-1']);
    expect(fixture.nativeElement.textContent).not.toContain('Quarterly update');
  });
});
