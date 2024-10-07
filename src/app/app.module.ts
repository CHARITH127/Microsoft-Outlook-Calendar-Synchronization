import {APP_INITIALIZER, NgModule} from '@angular/core';
import { BrowserModule } from '@angular/platform-browser';

import { FullCalendarModule } from '@fullcalendar/angular';

import { MsalModule, MsalService, MSAL_INSTANCE, MsalGuard, MsalGuardConfiguration, MsalInterceptor, MsalInterceptorConfiguration } from '@azure/msal-angular';
import {PublicClientApplication, IPublicClientApplication, InteractionType} from '@azure/msal-browser';

import { AppComponent } from './app.component';
import { CalenderComponent } from './calender/calender.component';
import {HttpClientModule} from "@angular/common/http";
import {FormsModule} from "@angular/forms";


// MSAL configuration without async in forRoot, handling the initialization separately
const MSALInstanceFactory = (): IPublicClientApplication => {
  return new PublicClientApplication({
    auth: {
      clientId: '958f7d9c-e3fd-4434-88ca-f01f0516aca8', // Replace with your client ID
      authority: 'https://login.microsoftonline.com/4fc2f3aa-31c4-4dcb-b719-c6c16393e9d3', // Replace with your tenant ID
      redirectUri: 'http://localhost:4200', // Your redirect URI
    },
    cache: {
      cacheLocation: 'localStorage',
      storeAuthStateInCookie: true, // Set to true for IE 11
    },
  });
};

// Function to initialize the MSAL application
const initializeMSAL = (msalInstance: IPublicClientApplication) => {
  return () => msalInstance.initialize(); // Return a function that initializes MSAL
};

// Msal Guard Configuration
const msalGuardConfig: MsalGuardConfiguration = {
  interactionType: InteractionType.Redirect,
  authRequest: {
    scopes: ['User.Read', 'Calendars.ReadWrite'] // Required permissions
  }
};

// Msal Interceptor Configuration
const msalInterceptorConfig: MsalInterceptorConfiguration = {
  interactionType: InteractionType.Redirect,
  protectedResourceMap: new Map([
    ['https://graph.microsoft.com/v1.0/me/events', ['Calendars.ReadWrite']],
    ['https://graph.microsoft.com/v1.0/me/', ['User.Read']],
  ]),
};

@NgModule({
  declarations: [
    AppComponent,
    CalenderComponent
  ],
  imports: [
    BrowserModule,
    FullCalendarModule,
    MsalModule.forRoot(
      MSALInstanceFactory(), // Use synchronous MSAL instance
      msalGuardConfig,
      msalInterceptorConfig
    ),
    HttpClientModule,
    FormsModule
  ],
  providers: [MsalService,
    {
      provide: MSAL_INSTANCE,
      useFactory: MSALInstanceFactory
    },
    {
      provide: APP_INITIALIZER,
      useFactory: initializeMSAL,
      deps: [MSAL_INSTANCE],
      multi: true // Ensure the function runs during app initialization
    }],
  bootstrap: [AppComponent]
})
export class AppModule { }
