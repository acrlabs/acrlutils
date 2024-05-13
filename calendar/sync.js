// Sync events between two calendars
//
// You need to subscribe to the remote calendar; this script will then copy events from
// that calendar to the local calendar.  It will also invite a user of your choice to
// events on the local calendar (the goal is to sync local events with the remote calendar
// if you don't have edit access to the remote calendar).

const BUSY_TITLE = "XXXX";                          // Title to set on events
const IGNORE_TITLES = [BUSY_TITLE, "busy"];         // Ignore these titles when copying to the remote calendar
const LOCAL_EMAIL = "YYYY";                         // Email address for the local calendar
const LOCAL_CALENDAR = "ZZZZ";                      // Name of the local calendar
const REMOTE_CALENDAR = "AAAA";                     // Name of the remote calendar
const REMOTE_EMAIL = "BBBB";                        // Email address for the remote calendar


// Call this function to sync events
function syncCalendarWithRemote() {

  var numDays = 5;

  var start = new Date();
  var end = new Date(start.getTime() + (60 * 60 * 24 * numDays * 1000));
  var remoteCalendar = CalendarApp.getCalendarsByName(REMOTE_CALENDAR)[0];
  var localCalendar = CalendarApp.getCalendarsByName(LOCAL_CALENDAR)[0];

  var remoteEvents = remoteCalendar.getEvents(start, end);
  var localEvents = localCalendar.getEvents(start, end);

  var searchIndex = 0;
  var remoteEventStartTimes = [];
  // Copy events from the remote calendar to the local calendar
  for (var i in remoteEvents) {
    var evt = remoteEvents[i];
    var evtStart = evt.getStartTime();
    var [evts, searchIndex] = findEventsWithStart(evtStart.getTime(), localEvents, searchIndex);
    remoteEventStartTimes.push(evtStart.getTime());
    if (shouldCreateEvent(evts)) {
      console.log("syncing event from remote calendar at " + evtStart);
      localCalendar.createEvent(BUSY_TITLE, evtStart, evt.getEndTime());
    }
  }

  for (var i in localEvents) {
    var evt = localEvents[i];
    var evtStart = evt.getStartTime();
    // Delete events from the local calendar that no longer exist on the remote calendar
    if (evt.getTitle() == BUSY_TITLE && !(remoteEventStartTimes.includes(evtStart.getTime()))) {
      console.log("deleting event from local calendar at " + evtStart);
      evt.deleteEvent();
    // invite the remote user to select local events
    } else if (!IGNORE_TITLES.includes(evt.getTitle()) && shouldInviteRemote(evt)) {
      console.log("inviting remote email for event at " + evtStart);
      evt.addGuest(REMOTE_EMAIL);
    }
  }
}

function findEventsWithStart(start, events, index) {
  var ret = []
  for (var i = index; i < events.length; i++) {
    var evtStart = events[i].getStartTime().getTime();
    if (evtStart == start) {
      ret.push(events[i])
    } else if (evtStart > start) {
      break; // Assume the input list is sorted
    }
  }

  return [ret, i]
}

// Create an event on the local calendar if it doesn't already exist, and the remote user hasn't been invited
function shouldCreateEvent(evts) {
  for (var i in evts) {
    if (evts[i].getTitle() == BUSY_TITLE) {
      return false;
    }

    var guests = evts[i].getGuestList();
    for (var j in guests) {
      if (guests[j].getEmail() == REMOTE_EMAIL) {
        return false;
      }
    }
  }

  return true;
}

// Invite the remote user to a local event if the local user has accepted (or owns) an event on the local calendar
function shouldInviteRemote(evt) {
  var localGuest = evt.getGuestByEmail(LOCAL_EMAIL);
  var remoteGuest = evt.getGuestByEmail(REMOTE_EMAIL);

  if (
    (evt.isOwnedByMe() || (localGuest != null && localGuest.getGuestStatus() == CalendarApp.GuestStatus.YES)) &&
    (remoteGuest == null)
  ) {
        return true;
  }

  return false;
}
