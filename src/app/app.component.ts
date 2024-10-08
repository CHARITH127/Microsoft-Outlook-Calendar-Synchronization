import {Component, OnInit, ViewChild} from '@angular/core';
import {MsalBroadcastService, MsalService} from "@azure/msal-angular";
import {AuthenticationResult, EventType, PublicClientApplication} from "@azure/msal-browser";
import { CalenderComponent } from './calender/calender.component';
import {GraphService} from "./services/graph.service";

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})
export class AppComponent implements OnInit{
  title = 'calendar-demo-fe';
  email: string = ''; // Ensure the email property is defined as a string

  @ViewChild(CalenderComponent) calendarComponent!: CalenderComponent;

  constructor(private authService: MsalService, private msalBroadcastService: MsalBroadcastService, private graphService : GraphService) {}

  ngOnInit(): void {
    this.authService.instance.handleRedirectPromise().then((result: AuthenticationResult | null) => {
      if (result !== null && result.account !== null) {
        this.authService.instance.setActiveAccount(result.account);
      }
    });

    this.msalBroadcastService.msalSubject$.subscribe((event) => {
      if (event.eventType === EventType.LOGIN_SUCCESS) {
        console.log('Login successful');
      } else if (event.eventType === EventType.LOGIN_FAILURE) {
        console.error('Login failed', event);
      }
    });
  }

  login() {
    this.authService.loginPopup({
      scopes: ['User.Read', 'MailboxSettings.Read', 'Calendars.ReadWrite']
    }).subscribe({
      next: (result) => {
        this.authService.instance.setActiveAccount(result.account);
        this.loadEventsOnLogin();

        // Call the method to get user timezone after login
        this.getUserTimeZone();
      },
      error: (error) => {
        console.error('Login failed', error);
      }
    });
  }

  private getUserTimeZone() {
    this.graphService.getUserMailboxSettings().subscribe({
      next: (response) => {
        if (response && response.timeZone) {
          this.graphService.setTimeZone(response.timeZone)
        } else {
          console.warn('Time zone information is missing from the response');
        }
      },
      error: (error) => {
        console.error('Error fetching user mailbox settings:', error);
      }
    });
  }

  logout(){
    this.authService.logoutPopup({
      mainWindowRedirectUri: 'http://localhost:4200', // The redirect URI after logout
    });
  }

  private loadEventsOnLogin() {
    if (this.calendarComponent) {
      this.calendarComponent.fetchEvents(); // Call fetchEvents on CalendarComponent
    }
  }

}
