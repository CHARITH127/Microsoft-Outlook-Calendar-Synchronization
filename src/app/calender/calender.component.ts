import {ChangeDetectorRef, Component, OnInit, signal, ViewChild} from '@angular/core';
import {CalendarOptions, DateSelectArg, EventApi, EventClickArg} from '@fullcalendar/core';
import interactionPlugin from '@fullcalendar/interaction';
import dayGridPlugin from '@fullcalendar/daygrid';
import timeGridPlugin from '@fullcalendar/timegrid';
import listPlugin from '@fullcalendar/list';
import {createEventId, INITIAL_EVENTS} from '../services/event-utils';
import { GraphService } from '../services/graph.service';
import { FullCalendarComponent } from '@fullcalendar/angular';
import * as moment from 'moment-timezone';

// Define the structure of an event from Microsoft Graph
interface GraphEvent {
  id: string;
  subject: string;
  start: { dateTime: string };
  end: { dateTime: string };
  isAllDay: boolean;
}

@Component({
  selector: 'app-calender',
  templateUrl: './calender.component.html',
  styleUrls: ['./calender.component.scss']
})
export class CalenderComponent implements OnInit{

  @ViewChild('fullCalendar') calendarComponent!: FullCalendarComponent;

  userTimeZone: string | null = 'UTC';

  calendarVisible = signal(true);
  calendarOptions = signal<CalendarOptions>({
    plugins: [
      interactionPlugin,
      dayGridPlugin,
      timeGridPlugin,
      listPlugin,
    ],
    headerToolbar: {
      left: 'prev,next today',
      center: 'title',
      right: 'dayGridMonth,timeGridWeek,timeGridDay,listWeek'
    },
    initialView: 'dayGridMonth',
    initialEvents: INITIAL_EVENTS, // alternatively, use the `events` setting to fetch from a feed
    weekends: true,
    editable: true,
    selectable: true,
    selectMirror: true,
    dayMaxEvents: true,
    select: this.handleDateSelect.bind(this),
    eventClick: this.handleEventClick.bind(this),
    eventsSet: this.handleEvents.bind(this)
    /* you can update a remote database when these fire:
    eventAdd:
    eventChange:
    eventRemove:
    */
  });
  currentEvents = signal<EventApi[]>([]);

  constructor(private changeDetector: ChangeDetectorRef, private graphService: GraphService) {}


  ngOnInit() {
    this.fetchEvents();
  }

  fetchEvents() {

    this.graphService.getEvents().subscribe((response: { value: GraphEvent[] }) => {
      // Map the response from GraphService to an event object that FullCalendar understands
      const events = response.value.map((event: GraphEvent) => {
        const startInUserTimeZone = moment.tz(event.start.dateTime, this.graphService.getTimeZone()).tz(this.graphService.getTimeZone()).format();
        const endInUserTimeZone = moment.tz(event.end.dateTime, this.graphService.getTimeZone()).tz(this.graphService.getTimeZone()).format();

        return {
          id: event.id,
          title: event.subject,
          start: startInUserTimeZone, // Use the converted start time
          end: endInUserTimeZone, // Use the converted end time
          allDay: event.isAllDay
        };
      });

      // Use the FullCalendar method to add events to the calendar
      const calendarApi = this.calendarComponent.getApi(); // Access FullCalendar's API
      calendarApi.removeAllEvents(); // Clear any existing events
      events.forEach(event => calendarApi.addEvent(event)); // Add each event from the Graph API

      this.changeDetector.detectChanges(); // Ensure the change detection runs after updating events
    }, error => {
      console.error('Error fetching events from Microsoft Graph:', error);
    });

  }

  handleCalendarToggle() {
    this.calendarVisible.update((bool) => !bool);
  }

  handleWeekendsToggle() {
    this.calendarOptions.mutate((options) => {
      options.weekends = !options.weekends;
    });
  }

  handleDateSelect(selectInfo: DateSelectArg) {
    const title = prompt('Please enter a new title for your event');
    const calendarApi = selectInfo.view.calendar;

    calendarApi.unselect(); // Clear date selection

    if (title) {
      const newEvent = {
        subject: title,
        start: {
          dateTime: selectInfo.startStr,
          timeZone: this.graphService.getTimeZone()
        },
        end: {
          dateTime: selectInfo.endStr,
          timeZone: this.graphService.getTimeZone()
        },
        isAllDay: selectInfo.allDay
      };

      // Call the service to save the event to Microsoft Graph
      this.graphService.createEvent(newEvent).subscribe({
        next: (response) => {
          console.log('Event created in Outlook calendar with ID:', response.id);
          // Add the event to FullCalendar using the ID from Microsoft Graph
          calendarApi.addEvent({
            id: response.id, // Use the ID from Microsoft Graph response
            title,
            start: selectInfo.startStr,
            end: selectInfo.endStr,
            allDay: selectInfo.allDay
          });
        },
        error: (error) => {
          console.error('Error creating event in Outlook calendar:', error);
        }
      });
    }
  }

  handleEventClick(clickInfo: EventClickArg) {
    if (confirm(`Are you sure you want to delete the event '${clickInfo.event.title}'`)) {

      // Call the GraphService to delete the event from Microsoft Graph
      this.graphService.deleteEvent(clickInfo.event.id).subscribe({
        next: () => {
          // Remove the event from FullCalendar if successfully deleted from Microsoft Graph
          clickInfo.event.remove();
          console.log(`Event '${clickInfo.event.title}' deleted successfully from Outlook calendar.`);
        },
        error: (error) => {
          console.error('Error deleting event from Outlook calendar:', error);
        }
      });
    }
  }

  handleEvents(events: EventApi[]) {
    this.currentEvents.set(events);
    this.changeDetector.detectChanges(); // workaround for pressionChangedAfterItHasBeenCheckedError
  }


}
