﻿using System.Collections.ObjectModel;
using System.Windows.Input;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using System.Net.Http.Headers;
using Syncfusion.Maui.Scheduler;
using MicrosoftDayOfWeek = Microsoft.Graph.DayOfWeek;
using SchedulerRecurrenceRange = Syncfusion.Maui.Scheduler.SchedulerRecurrenceRange;
using SchedulerRecurrenceType = Syncfusion.Maui.Scheduler.SchedulerRecurrenceType;

namespace MauiSyncOutlookCalendar
{
    public class SchedulerViewModel
    {
        private GraphServiceClient Client;
        private static string[] scopes = { "User.Read", "Calendars.Read", "Calendars.ReadWrite" };

        public ObservableCollection<Meeting> Meetings { get; set; }

        public ICommand ImportButtonCommand { get; set; }

        public ICommand ExportButtonCommand { get; set; }

        public SchedulerViewModel()
        {
            Meetings = new ObservableCollection<Meeting>();
            this.ImportButtonCommand = new Command(ExecuteImportCommand);
            this.ExportButtonCommand = new Command(ExecuteExportCommand);
            this.AddSchedulerEvents();
        }

        /// <summary>
        /// Method to add events to scheduler.
        /// </summary>
        private void AddSchedulerEvents()
        {
            var colors = new List<Brush>
            {
                new SolidColorBrush(Color.FromArgb("#FF8B1FA9")),
                new SolidColorBrush(Color.FromArgb("#FFD20100")),
                new SolidColorBrush(Color.FromArgb("#FFFC571D")),
                new SolidColorBrush(Color.FromArgb("#FF36B37B")),
                new SolidColorBrush(Color.FromArgb("#FF3D4FB5")),
                new SolidColorBrush(Color.FromArgb("#FFE47C73")),
                new SolidColorBrush(Color.FromArgb("#FF636363")),
                new SolidColorBrush(Color.FromArgb("#FF85461E")),
                new SolidColorBrush(Color.FromArgb("#FF0F8644")),
            };

            var subjects = new List<string>
            {
                "Scrum Meeting",
                "Review Meeting",
                "Rating Discussion",
                "Development Meeting",
                "Sprint Review",
                "Sprint Planning",
                "Sprint Retrospective",
                "General Meeting",
                "Yoga Therapy",
            };

            Random ran = new Random();
            for (int startdate = -5; startdate < 5; startdate++)
            {
                var meeting = new Meeting();
                meeting.EventName = subjects[ran.Next(0, subjects.Count)];
                meeting.From = DateTime.Now.Date.AddDays(startdate).AddHours(9);
                meeting.To = meeting.From.AddHours(10);
                meeting.Background = colors[ran.Next(0, colors.Count)];
                this.Meetings.Add(meeting);
            }
        }

        /// <summary>
        /// Method to import the Outlook Calendar to Syncfusion Scheduler.
        /// </summary>
        /// <param name="parameter"></param>
        public void ExecuteImportCommand(object parameter)
        {
            this.Authenticate(true);
        }

        /// <summary>
        /// Method to export the Syncfusion Scheduler events to Outlook Calendar.
        /// </summary>
        /// <param name="parameter"></param>
        public void ExecuteExportCommand(object parameter)
        {
            this.Authenticate(false);
        }

        /// <summary>
        /// Method to connect application authentication with Microsoft Azure. 
        /// </summary>
        /// <param name="import">import or export events</param>
        private async void Authenticate(bool import)
        {
            AuthenticationResult tokenRequest;
            var accounts = await App.ClientApplication.GetAccountsAsync();
            if (accounts.Count() > 0)
            {
                tokenRequest = await App.ClientApplication.AcquireTokenSilent(scopes, accounts.FirstOrDefault())
                    .ExecuteAsync();
            }
            else
            {

                tokenRequest = await App.ClientApplication.AcquireTokenInteractive(scopes).ExecuteAsync();
            }

            Client = new GraphServiceClient("https://graph.microsoft.com/v1.0/",
                                new DelegateAuthenticationProvider(async (requestMessage) =>
                                {
                                    requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", tokenRequest.AccessToken);
                                }));
            if (import)
            {
                this.GetOutlookCalendarEvents();
            }
            else
            {
                this.AddEventToOutlookCalendar();
            }
        }

        /// <summary>
        /// Method to add event to Outlook Calendar.
        /// </summary>
        private void AddEventToOutlookCalendar()
        {
            foreach (Meeting meeting in this.Meetings)
            {
                Event calendarEvent = new Event
                {
                    Subject = meeting.EventName,

                    Start = new DateTimeTimeZone
                    {
                        DateTime = meeting.From.ToString(),
                        TimeZone = "GMT Standard Time"
                    },
                    End = new DateTimeTimeZone()
                    {
                        DateTime = meeting.To.ToString(),
                        TimeZone = "GMT Standard Time"
                    },
                };
                //// Request to add Syncfusion Scheduler event to the Outlook Calendar events.
                Client.Me.Events.Request().AddAsync(calendarEvent);
            }
        }

        /// <summary>
        /// Method to get Outlook Calendar events.
        /// </summary>
        private void GetOutlookCalendarEvents()
        {
            //// Request to get the outlook calendar events.
            var events = Client.Me.Events.Request().GetAsync().Result.ToList();
            if (events != null && events.Count > 0)
            {
                foreach (Event appointment in events)
                {
                    Meeting meeting = new Meeting()
                    {
                        EventName = appointment.Subject,
                        From = Convert.ToDateTime(appointment.Start.DateTime),
                        To = Convert.ToDateTime(appointment.End.DateTime),
                        IsAllDay = (bool)appointment.IsAllDay,
                    };

                    if (appointment.Recurrence != null)
                    {
                        AddRecurrenceRule(appointment, meeting);
                    }
                    this.Meetings.Add(meeting);
                }
            }
        }

        /// <summary>
        /// Method to update recurrence rule to appointments.
        /// </summary>
        /// <param name="appointment"></param>
        /// <param name="meeting"></param>
        private static void AddRecurrenceRule(Event appointment, Meeting meeting)
        {
            // Creating recurrence rule
            SchedulerRecurrenceInfo recurrenceProperties = new SchedulerRecurrenceInfo();
            if (appointment.Recurrence.Pattern.Type == RecurrencePatternType.Daily)
            {
                recurrenceProperties.RecurrenceType = SchedulerRecurrenceType.Daily;
            }
            else if (appointment.Recurrence.Pattern.Type == RecurrencePatternType.Weekly)
            {
                recurrenceProperties.RecurrenceType = SchedulerRecurrenceType.Weekly;
                foreach (var weekDay in appointment.Recurrence.Pattern.DaysOfWeek)
                {
                    if (weekDay == MicrosoftDayOfWeek.Sunday)
                    {
                        recurrenceProperties.WeekDays = SchedulerWeekDays.Sunday;
                    }
                    if (weekDay == MicrosoftDayOfWeek.Monday)
                    {
                        recurrenceProperties.WeekDays = recurrenceProperties.WeekDays | SchedulerWeekDays.Monday;
                    }
                    if (weekDay == MicrosoftDayOfWeek.Tuesday)
                    {
                        recurrenceProperties.WeekDays = recurrenceProperties.WeekDays | SchedulerWeekDays.Tuesday;
                    }
                    if (weekDay == MicrosoftDayOfWeek.Wednesday)
                    {
                        recurrenceProperties.WeekDays = recurrenceProperties.WeekDays | SchedulerWeekDays.Wednesday;
                    }
                    if (weekDay == MicrosoftDayOfWeek.Thursday)
                    {
                        recurrenceProperties.WeekDays = recurrenceProperties.WeekDays | SchedulerWeekDays.Thursday;
                    }
                    if (weekDay == MicrosoftDayOfWeek.Friday)
                    {
                        recurrenceProperties.WeekDays = recurrenceProperties.WeekDays | SchedulerWeekDays.Friday;
                    }
                    if (weekDay == MicrosoftDayOfWeek.Saturday)
                    {
                        recurrenceProperties.WeekDays = recurrenceProperties.WeekDays | SchedulerWeekDays.Saturday;
                    }
                }
            }
            recurrenceProperties.Interval = (int)appointment.Recurrence.Pattern.Interval;
            recurrenceProperties.RecurrenceRange = SchedulerRecurrenceRange.Count;
            recurrenceProperties.RecurrenceCount = 10;
            meeting.RRule = SchedulerRecurrenceManager.GenerateRRule(recurrenceProperties, meeting.From, meeting.To);
        }
    }
}
