import { Injectable } from '@angular/core';
import type {
  EndSessionRequest,
  PublicClientApplication,
  RedirectRequest,
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
  private msalAppInitialization?: Promise<PublicClientApplication>;
  private interactiveAuthRequest?: Promise<unknown>;

  constructor() {
    // Warm up MSAL early so redirect results are processed as soon as the app boots.
    if (this.isConfigured()) {
      void this.getMsalApp();
    }
  }

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
    const app = await this.getMsalApp();
    const account = app.getActiveAccount() ?? app.getAllAccounts()[0];

    if (account) {
      app.setActiveAccount(account);
      return account.username;
    }

    const redirectRequest: RedirectRequest = {
      scopes: this.authScopes,
      redirectUri: this.configuration.redirectUri,
    };

    await this.runInteractiveAuthRequest(() => app.loginRedirect(redirectRequest));
    return 'Redirecting to Microsoft sign-in...';
  }

  async signOut(): Promise<void> {
    const app = await this.getMsalApp();
    const account = app.getActiveAccount() ?? app.getAllAccounts()[0];

    if (account) {
      const logoutRequest: EndSessionRequest = {
        account,
        postLogoutRedirectUri: this.configuration.redirectUri,
      };
      await this.runInteractiveAuthRequest(() => app.logoutRedirect(logoutRequest));
    }
  }

  async getSignedInUsername(): Promise<string> {
    const app = await this.getMsalApp();
    const account = app.getActiveAccount() ?? app.getAllAccounts()[0];

    if (!account) {
      return '';
    }

    app.setActiveAccount(account);
    return account.username;
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

  async deleteMessages(messageIds: string[], onProgress?: (processedCount: number) => void): Promise<void> {
    let processedCount = 0;

    for (const messageId of messageIds) {
      await this.requestGraph<void>(`${this.graphBaseUrl}/me/messages/${encodeURIComponent(messageId)}`, 'DELETE');
      processedCount += 1;
      onProgress?.(processedCount);
    }
  }

  private async getAccessToken(): Promise<string> {
    const app = await this.getMsalApp();
    const account = app.getActiveAccount() ?? app.getAllAccounts()[0];

    if (!account) {
      await this.runInteractiveAuthRequest(() =>
        app.loginRedirect({
          scopes: this.authScopes,
          redirectUri: this.configuration.redirectUri,
        }),
      );
      throw new Error('Redirecting to Microsoft sign-in...');
    }

    app.setActiveAccount(account);

    try {
      const result = await app.acquireTokenSilent({
        account,
        scopes: this.authScopes,
      });
      return result.accessToken;
    } catch {
      await this.runInteractiveAuthRequest(() =>
        app.acquireTokenRedirect({
          account,
          scopes: this.authScopes,
          redirectUri: this.configuration.redirectUri,
        }),
      );
      throw new Error('Redirecting to Microsoft consent page...');
    }
  }

  private async runInteractiveAuthRequest<T>(request: () => Promise<T>): Promise<T> {
    while (this.interactiveAuthRequest) {
      try {
        await this.interactiveAuthRequest;
      } catch {
        // Ignore failures from previous attempts and continue with the next request.
      }
    }

    const activeRequest = request();
    this.interactiveAuthRequest = activeRequest;

    try {
      return await activeRequest;
    } finally {
      if (this.interactiveAuthRequest === activeRequest) {
        this.interactiveAuthRequest = undefined;
      }
    }
  }

  private async getMsalApp(): Promise<PublicClientApplication> {
    if (!this.isConfigured()) {
      throw new Error(this.getConfigurationMessage());
    }

    if (this.msalApp) {
      return this.msalApp;
    }

    if (!this.msalAppInitialization) {
      this.msalAppInitialization = this.initializeMsalApp();
    }

    try {
      return await this.msalAppInitialization;
    } catch (error) {
      this.msalAppInitialization = undefined;
      throw error;
    }
  }

  private async initializeMsalApp(): Promise<PublicClientApplication> {
    const { PublicClientApplication } = await import('@azure/msal-browser');

    const app = new PublicClientApplication({
      auth: {
        clientId: this.configuration.clientId,
        authority: `https://login.microsoftonline.com/${this.configuration.tenantId}`,
        redirectUri: this.configuration.redirectUri,
      },
      cache: {
        cacheLocation: 'sessionStorage',
      },
    });

    await app.initialize();

    try {
      const redirectResult = await app.handleRedirectPromise();
      if (redirectResult?.account) {
        app.setActiveAccount(redirectResult.account);
      }
    } catch (error) {
      if (!(error instanceof Error) || !error.message.includes('no_token_request_cache_error')) {
        throw error;
      }
    }

    const account = app.getAllAccounts()[0];
    if (account) {
      app.setActiveAccount(account);
    }

    this.msalApp = app;
    return app;
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
