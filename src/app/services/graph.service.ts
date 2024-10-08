import { Injectable } from '@angular/core';
import { HttpClient, HttpHeaders } from '@angular/common/http';
import {BehaviorSubject, Observable, switchMap} from 'rxjs';
import {MsalService} from "@azure/msal-angular";

@Injectable({
  providedIn: 'root'
})
export class GraphService {
  private graphUrl = 'https://graph.microsoft.com/v1.0/me/events'; // Microsoft Graph API URL
  private timeZoneSubject = new BehaviorSubject<string>('UTC'); // Default time zone

  constructor(private http: HttpClient, private authService: MsalService) {}

  // Function to fetch events from Microsoft Graph
  getEvents(): Observable<any> {
    return this.authService.acquireTokenSilent({
      scopes: ['Calendars.ReadWrite']
    }).pipe(
      switchMap(result => {
        const headers = this.createHeaders(result.accessToken);
        return this.http.get(this.graphUrl, { headers });
      })
    );
  }

  // Function to create a new event in Microsoft Graph
  createEvent(event: any): Observable<any> {
    return this.authService.acquireTokenSilent({
      scopes: ['Calendars.ReadWrite']
    }).pipe(
      switchMap(result => {
        const headers = this.createHeaders(result.accessToken);
        return this.http.post(this.graphUrl, event, { headers });
      })
    );
  }

  // Function to delete an event from Microsoft Graph
  deleteEvent(eventId: string): Observable<any> {
    return this.authService.acquireTokenSilent({
      scopes: ['Calendars.ReadWrite']
    }).pipe(
      switchMap(result => {
        const headers = this.createHeaders(result.accessToken);
        return this.http.delete(`${this.graphUrl}/${eventId}`, { headers });
      })
    );
  }

  // Function to create headers for the requests
  private createHeaders(token: string) {
    return new HttpHeaders({
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json'
    });
  }

  getUserMailboxSettings(): Observable<any> {
    return this.authService.acquireTokenSilent({
      scopes: ['User.Read', 'MailboxSettings.Read']
    }).pipe(
      switchMap(result => {
        console.log('Access Token Acquired:', result.accessToken);
        const headers = this.createHeaders(result.accessToken);
        return this.http.get('https://graph.microsoft.com/v1.0/me/mailboxSettings', { headers });
      })
    );
  }

  // Function to set the time zone
  setTimeZone(timeZone: string) {
    this.timeZoneSubject.next(timeZone); // Emit the new time zone
  }

  // Function to get the current time zone
  getTimeZone(): string {
    return this.timeZoneSubject.getValue(); // Return the current time zone
  }

}
