import { DatePipe } from '@angular/common';
import { Component, inject } from '@angular/core';
import { FormsModule } from '@angular/forms';

import { GraphMailService } from './graph-mail.service';
import { MailMessage } from './mail-message';

interface SelectableMailMessage extends MailMessage {
  selected: boolean;
}

@Component({
  selector: 'app-root',
  imports: [DatePipe, FormsModule],
  templateUrl: './app.html',
  styleUrl: './app.css',
})
export class App {
  private readonly graphMailService = inject(GraphMailService);

  readonly senderPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  readonly configurationMessage = this.graphMailService.getConfigurationMessage();

  signedInUser = '';
  senderEmail = '';
  messages: SelectableMailMessage[] = [];
  statusMessage = '';
  errorMessage = '';
  isSignedIn = false;
  isLoading = false;
  isDeleting = false;

  get hasConfiguration(): boolean {
    return this.graphMailService.isConfigured();
  }

  get hasMessages(): boolean {
    return this.messages.length > 0;
  }

  get selectedCount(): number {
    return this.messages.filter((message) => message.selected).length;
  }

  get allSelected(): boolean {
    return this.hasMessages && this.selectedCount === this.messages.length;
  }

  async signIn(): Promise<void> {
    this.errorMessage = '';
    this.statusMessage = '';

    try {
      this.signedInUser = await this.graphMailService.signIn();
      this.isSignedIn = true;
      this.statusMessage = `Signed in as ${this.signedInUser}.`;
    } catch (error) {
      this.errorMessage = this.toErrorMessage(error);
    }
  }

  async signOut(): Promise<void> {
    this.errorMessage = '';

    try {
      await this.graphMailService.signOut();
      this.isSignedIn = false;
      this.signedInUser = '';
      this.messages = [];
      this.statusMessage = 'Signed out.';
    } catch (error) {
      this.errorMessage = this.toErrorMessage(error);
    }
  }

  async findMessages(): Promise<void> {
    this.errorMessage = '';
    this.statusMessage = '';

    if (!this.isSignedIn) {
      this.errorMessage = 'Sign in with Microsoft before loading messages.';
      return;
    }

    if (!this.senderPattern.test(this.senderEmail.trim())) {
      this.errorMessage = 'Enter a valid sender email address before searching.';
      return;
    }

    this.isLoading = true;

    try {
      const messages = await this.graphMailService.loadMessagesBySender(this.senderEmail);
      this.messages = messages.map((message) => ({ ...message, selected: false }));
      this.statusMessage =
        this.messages.length > 0
          ? `Loaded ${this.messages.length} message(s) from ${this.senderEmail.trim()}.`
          : `No messages found from ${this.senderEmail.trim()}.`;
    } catch (error) {
      this.errorMessage = this.toErrorMessage(error);
      this.messages = [];
    } finally {
      this.isLoading = false;
    }
  }

  toggleAllSelections(checked: boolean): void {
    this.messages = this.messages.map((message) => ({ ...message, selected: checked }));
  }

  async deleteSelectedMessages(): Promise<void> {
    const selectedMessages = this.messages.filter((message) => message.selected);

    if (selectedMessages.length === 0) {
      return;
    }

    const confirmed = window.confirm(
      `Delete ${selectedMessages.length} selected message(s) from ${this.senderEmail.trim()}?`,
    );

    if (!confirmed) {
      return;
    }

    this.errorMessage = '';
    this.statusMessage = '';
    this.isDeleting = true;

    try {
      await this.graphMailService.deleteMessages(selectedMessages.map((message) => message.id));
      this.messages = this.messages.filter((message) => !message.selected);
      this.statusMessage = `Deleted ${selectedMessages.length} message(s).`;
    } catch (error) {
      this.errorMessage = this.toErrorMessage(error);
    } finally {
      this.isDeleting = false;
    }
  }

  trackByMessageId(_: number, message: SelectableMailMessage): string {
    return message.id;
  }

  private toErrorMessage(error: unknown): string {
    return error instanceof Error ? error.message : 'An unexpected error occurred.';
  }
}
