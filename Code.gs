// 複数のカレンダー間でイベントを共有するためのGAS
// SYNC_SETTINGSで記述した複数のカレンダー間で，互いのイベントを自動的にコピーする
// GASの時間ベーストリガーで定期的（1時間〜1日）に実行することを推奨

// 1. How many days do you want to sync your calendars?
const DAYS_TO_SYNC = 14;

// 2. Settings
// calendar_id: string // GoogleカレンダーのID
// private_import: bool or string // trueの場合，このカレンダーには他カレンダーのイベントを「予定あり (imported)」としてインポート．文字列を指定した場合には，その文字列をタイトルとする
// import_from: list string // リストで指定したGoogleカレンダーIDからのみインポートする．
// read_only: bool // trueの場合，このカレンダーからエクスポートはするが，インポートはしない．
const SYNC_SETTINGS = [
  { // Example: Private gmail
    calendar_id: 'foo@gmail.com',
  },
  { // Example: Work gmail
    calendar_id: 'foo@your_company.com',
    private_import: 'z_予定あり (imported)',
  },
  { // Example: foo@your_company.com のオリジナルの予定（インポートされたものではない予定）のみを抜き出したカレンダー
    calendar_id: '000000@group.calendar.google.com',
    import_from: ['foo@your_company.com'],
  },
  { // Example: Apple Calendar等の外部カレンダーをgoogleカレンダーに追加したもの，インポート専用
    calendar_id: '000000@import.calendar.google.com', 
    read_only: true,
  },
];

function main(){
  var startDate = new Date();
  var endDate = new Date(startDate.getTime() + (DAYS_TO_SYNC * 24 * 60 * 60* 1000));

  syncCalendars(SYNC_SETTINGS, startDate, endDate);
}

function unique_id_of(event) {
  const id = event.getId();
  const time = event.getStartTime();
  return id + '_' + time.getTime();
}

function syncCalendars(calendar_settings, startDate, endDate) {
  
  let calendar_events = {};

  for (let calendar_setting of calendar_settings) {
    const calendar_id = calendar_setting.calendar_id;

    const calendar = CalendarApp.getCalendarById(calendar_id);

    if (!calendar) {
      Logger.log('CALENDAR ' + calendar_id + ' IS NOT FOUND!!');
      continue;
    }

    const events = calendar.getEvents(startDate, endDate);
    let copied_events = [];
    let original_events = {};

    for (let event of events) {
      description = event.getDescription();
      const match = description.match(/^Imported from (\S+) (\S+)/);
      if (match) { // Copied event
        copied_events.push({event: event, original_calendar_id: match[1], original_event_id: match[2]})
      }
      else { // Original event
        if (event.getTitle().startsWith("(") || event.getTitle().startsWith("（")) {
          continue;
        }
        if (event.getMyStatus() == CalendarApp.GuestStatus.NO || 
            event.getMyStatus() == CalendarApp.GuestStatus.INVITED) {
          continue;
        }
        original_events[unique_id_of(event)] = event;
      }
    }

    calendar_events[calendar_id] = { 
      calendar: calendar,
      setting: calendar_setting,
      copied_events: copied_events,
      original_events: original_events,
    };
  }

  Logger.log('Pulling');

  for (let [calendar_id, calendar_event] of Object.entries(calendar_events)) {
    if (calendar_event.setting.read_only) continue;

    // Pull copied events
    for (let copied of calendar_event.copied_events) {
      if (!calendar_events[copied.original_calendar_id]) {
        Logger.log('ORIGINAL CALENDAR ' + calendar_id + ' IS NOT FOUND!!');
        continue;
      }
      const original_calendar = calendar_events[copied.original_calendar_id].calendar;
      // Logger.log('hoge');
      Logger.log('Pulling from ' + original_calendar.getName() + ' to ' + calendar_event.calendar.getName());
      const original_event = calendar_events[copied.original_calendar_id].original_events[copied.original_event_id];
      if (original_event) {
        Logger.log('  updated: ' + original_event.getTitle() 
                   + ' (' + original_event.getStartTime() + ', ' + original_event.getEndTime() + ') '
                   + copied.original_event_id);
        if (copied.event.getEndTime() != original_event.getEndTime())
          copied.event.setTime(original_event.getStartTime(), original_event.getEndTime());
        if (copied.event.getLocation() != original_event.getLocation())
          copied.event.setLocation(original_event.getLocation());
      }
      else {
        Logger.log('  deleted: ' + copied.event.getTitle() 
                   + ' (' + copied.event.getStartTime() + ', ' + copied.event.getEndTime() + ') '
                   + copied.original_event_id);
        copied.event.deleteEvent();
      }
    }
  }

  Logger.log('Pushing');
  
  for (let [calendar_id, calendar_event] of Object.entries(calendar_events)) {
    // Push original events
    Logger.log('Pushing from ' + calendar_event.calendar.getName());
    for (let [original_event_id, original_event] of Object.entries(calendar_event.original_events)) {
      for (let calendar_id_to in calendar_events) {
        if (calendar_id_to == calendar_id) continue;

        const calendar_event_to = calendar_events[calendar_id_to];
        if (calendar_event_to.setting.read_only) continue;
        if (calendar_event_to.setting.import_from != undefined && 
            !calendar_event_to.setting.import_from.includes(calendar_id)) continue;

        let already_copied = false;
        for (const copied of calendar_event_to.copied_events) {
          if (copied.original_calendar_id == calendar_id &&
              copied.original_event_id == original_event_id) {
            already_copied = true;
            break;
          }
        }

        if (!already_copied) {
          let title = original_event.getTitle();
          if (calendar_event_to.setting.private_import) {
            if (calendar_event_to.setting.private_import === true) title = '予定あり (imported)';
            else title = calendar_event_to.setting.private_import;
          }

          Logger.log('  created: ' + original_event.getTitle() + ' as ' + title  
                     + ' (' + original_event.getStartTime() + ', ' + original_event.getEndTime() + ') '
                     + original_event_id);
          event = calendar_event_to.calendar.createEvent(title, original_event.getStartTime(), original_event.getEndTime());
          event.setDescription('Imported from ' + calendar_id + ' ' + original_event_id);
          event.setLocation(original_event.getLocation());
        }
      };
    }
  }
}
