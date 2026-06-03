import { Injectable } from '@angular/core';
import type {
  AuthenticationResult,
  PopupRequest,
  PublicClientApplication,
} from '@azure/msal-browser';

import { environment } from '../environments/environment';
import { MailMessage } from './mail-message';

interface GraphMessageResponse {
  value: GraphMessage[];
  '@odata.nextLink'?: string;
}

interface GraphMessage {
  id: string;
  subject?: string;
  receivedDateTime?: string;
  bodyPreview?: string;
  from?: {
    emailAddress?: {
      address?: string;
      name?: string;
    };
  };
}

@Injectable({ providedIn: 'root' })
export class GraphMailService {
  private readonly graphBaseUrl = 'https://graph.microsoft.com/v1.0';
  private readonly authScopes = ['User.Read', 'Mail.ReadWrite'];
  private readonly configuration = environment.azure;
  private msalApp?: PublicClientApplication;

  isConfigured(): boolean {
    return [
      this.configuration.clientId,
      this.configuration.tenantId,
      this.configuration.redirectUri,
    ].every((value) => value.trim().length > 0 && !value.startsWith('YOUR_'));
  }

  getConfigurationMessage(): string {
    return 'Update src/environments/environment.ts with your Azure app registration values before signing in.';
  }

  async signIn(): Promise<string> {
    const result = await this.loginWithPopup();
    return result.account?.username ?? 'Microsoft account';
  }

  async signOut(): Promise<void> {
    const app = await this.getMsalApp();
    const account = app.getActiveAccount() ?? app.getAllAccounts()[0];

    if (account) {
      await app.logoutPopup({ account });
    }
  }

  async loadMessagesBySender(senderEmail: string): Promise<MailMessage[]> {
    const normalizedEmail = senderEmail.trim();

    if (!normalizedEmail) {
      return [];
    }

    const encodedFilter = this.escapeODataLiteral(normalizedEmail);
    let nextUrl =
      `${this.graphBaseUrl}/me/messages` +
      `?$top=100&$select=id,subject,receivedDateTime,bodyPreview,from` +
      `&$filter=from/emailAddress/address eq '${encodedFilter}'`;

    const messages: MailMessage[] = [];

    while (nextUrl) {
      const response = await this.requestGraph<GraphMessageResponse>(nextUrl);
      messages.push(
        ...response.value.map((message) => ({
          id: message.id,
          fromAddress: message.from?.emailAddress?.address ?? normalizedEmail,
          fromName: message.from?.emailAddress?.name ?? normalizedEmail,
          subject: message.subject?.trim() || '(no subject)',
          receivedDateTime: message.receivedDateTime ?? '',
          bodyPreview: message.bodyPreview?.trim() || 'No preview available.',
        })),
      );
      nextUrl = response['@odata.nextLink'] ?? '';
    }

    return messages.sort((left, right) => right.receivedDateTime.localeCompare(left.receivedDateTime));
  }

  async deleteMessages(messageIds: string[]): Promise<void> {
    for (const messageId of messageIds) {
      await this.requestGraph<void>(`${this.graphBaseUrl}/me/messages/${encodeURIComponent(messageId)}`, 'DELETE');
    }
  }

  private async loginWithPopup(): Promise<AuthenticationResult> {
    const app = await this.getMsalApp();
    const loginRequest: PopupRequest = { scopes: this.authScopes };
    const result = await app.loginPopup(loginRequest);
    app.setActiveAccount(result.account);
    return result;
  }

  private async getAccessToken(): Promise<string> {
    const app = await this.getMsalApp();
    const account = app.getActiveAccount() ?? app.getAllAccounts()[0];

    if (!account) {
      const loginResult = await this.loginWithPopup();
      return loginResult.accessToken;
    }

    app.setActiveAccount(account);

    try {
      const result = await app.acquireTokenSilent({
        account,
        scopes: this.authScopes,
      });
      return result.accessToken;
    } catch {
      const result = await app.acquireTokenPopup({
        account,
        scopes: this.authScopes,
      });
      app.setActiveAccount(result.account);
      return result.accessToken;
    }
  }

  private async getMsalApp(): Promise<PublicClientApplication> {
    if (!this.isConfigured()) {
      throw new Error(this.getConfigurationMessage());
    }

    if (!this.msalApp) {
      const { PublicClientApplication } = await import('@azure/msal-browser');

      this.msalApp = new PublicClientApplication({
        auth: {
          clientId: this.configuration.clientId,
          authority: `https://login.microsoftonline.com/${this.configuration.tenantId}`,
          redirectUri: this.configuration.redirectUri,
        },
        cache: {
          cacheLocation: 'sessionStorage',
        },
      });
      await this.msalApp.initialize();

      const account = this.msalApp.getAllAccounts()[0];
      if (account) {
        this.msalApp.setActiveAccount(account);
      }
    }

    return this.msalApp;
  }

  private async requestGraph<T>(url: string, method = 'GET'): Promise<T> {
    const accessToken = await this.getAccessToken();
    const response = await fetch(url, {
      method,
      headers: {
        Authorization: 'Bearer ' + accessToken,
      },
    });

    if (!response.ok) {
      throw new Error(await this.buildGraphError(response));
    }

    if (response.status === 204) {
      return undefined as T;
    }

    return (await response.json()) as T;
  }

  private async buildGraphError(response: Response): Promise<string> {
    try {
      const payload = (await response.json()) as { error?: { message?: string } };
      if (payload.error?.message) {
        return payload.error.message;
      }
    } catch {
      // Ignore JSON parsing failures and fall back to the status text.
    }

    return `Microsoft Graph request failed with status ${response.status}: ${response.statusText || 'Unknown error'}`;
  }

  private escapeODataLiteral(value: string): string {
    return value.replace(/'/g, "''");
  }
}
