import { Injectable } from '@angular/core';
import { MsalService } from '@azure/msal-angular';

import { AlertsService } from './alerts.service';
import { OAuthSettings } from '../oauth';
import { User } from './user';
import { faSignOutAlt } from '@fortawesome/free-solid-svg-icons';

import { HttpClient } from '@angular/common/http';
import { DomSanitizer, SafeUrl } from '@angular/platform-browser';

import { Client } from '@microsoft/microsoft-graph-client';

@Injectable({
  providedIn: 'root'
})
export class AuthService {
  public authenticated: boolean;
  public user: User;
  //imageData: SafeUrl;

  constructor(
    private msalService: MsalService,
    private alertsService: AlertsService){
      this.authenticated = this.msalService.getUser() != null;
      this.getUser().then((user) => {this.user = user});
   }

   // Prompt the user to sign in and
   // grant consent to the requested permission scopes
   async signIn(): Promise<void> {
     let result = await this.msalService.loginPopup(OAuthSettings.scopes)
      .catch((reason) => {
        this.alertsService.add('Login failed', JSON.stringify(reason, null, 2));
      });
     if (result) {
       this.authenticated = true;
       
       this.user = await this.getUser();
     }
   }


  // Sign out
  signOut(): void {
  this.msalService.logout();
  this.user = null;
  this.authenticated = false;
  }

  // Silently request as access token
  async getAccessToken(): Promise<string> {
    let result = await this.msalService.acquireTokenSilent(OAuthSettings.scopes)
      .catch((reason) => {
        this.alertsService.add('Get token failed', JSON.stringify(reason, null, 2));
      });
  
  // Temporary to display token in an error box
  if (result) this.alertsService.add('Token acquired', result);
  return result;
    }

  private async getUser(): Promise<User> {
    if (!this.authenticated) return null;

    let graphClient = Client.init({
      // Initialize the Graph client with an auth
      // provider that requests the token from the
      // auth service
      authProvider: async(done) => {
        let token = await this.getAccessToken()
        .catch((reason) => {
          done(reason, null);
        });

        if (token)
        {
          done(null, token);
        } else {
          done("Could not get an access token", null);
        }
      }
    });

    // Get the user from Graph (GET /me)
    let graphUser = await graphClient.api('/me').get();

    let user = new User();
    user.displayName = graphUser.displayName || "Joe Public";
    // Prefer the mail property, but fall back to userPrincipalName
    user.email = graphUser.mail || graphUser.userPrincipalName;

    return user;
  }

  // private async getUserImage(): Promise<void> {
  //   const imageUrl = 'https://graph.microsoft.com/beta/me/photo/$value';

  //   this.http.get(imageUrl, {
  //     responseType: 'blob'
  //   })
  //   .toPromise()
  //   .then((res: any) => {
  //     let blob = new Blob([res._body], {
  //       type: res.headers.get("Content-Type")
  //     });

  //     let urlCreator = window.URL;
  //     this.imageData = this.sanitizer.bypassSecurityTrustUrl(
  //       urlCreator.createObjectURL(blob));
  //   });
  // }
}