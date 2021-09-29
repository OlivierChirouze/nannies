function getPreviousMonth() {
    const previousMonth = new Date();
    previousMonth.setMonth(previousMonth.getMonth() - 1);
    return previousMonth;
}

function onOpen() {
    const currentMonth = new Date();
    const previousMonth = getPreviousMonth();

    const ui = SpreadsheetApp.getUi();
    // Or DocumentApp or FormApp.
    ui.createMenu('Calendar import')
        .addItem('Import ' + formatDate(previousMonth), 'importPreviousMonth')
        .addItem('Import ' + formatDate(currentMonth), 'importCurrentMonth')
        .addToUi();
}

function formatDate(date) {
    return date.toLocaleDateString(undefined, {year: 'numeric', month: 'long'});
}

function importPreviousMonth() {
    importMonth(getPreviousMonth());
}

function importCurrentMonth() {
    importMonth(new Date());
}

// https://developers.google.com/apps-script/guides/sheets
const CalendarImport = function () {
    return this;
};

// Shouldn't be more than 1 event per day!
CalendarImport.maxMonthEvents = 300;

CalendarImport.prototype = {
    getMonthEvents: function (currentDate) {
        const calendarId = 'ie7d48frttnbcj8678ln7lv860@group.calendar.google.com';

        // First day of month
        const monthBegin = new Date(currentDate.getFullYear(), currentDate.getMonth(), 1);
        // Last day of month, end of day
        const monthEnd = new Date(currentDate.getFullYear(), currentDate.getMonth() + 1, 0, 23, 59, 59);

        const optionalArgs = {
            timeMin: monthBegin.toISOString(),
            timeMax: monthEnd.toISOString(),
            showDeleted: false,
            singleEvents: true,
            maxResults: CalendarImport.maxMonthEvents,
            orderBy: 'startTime'
        };

        return Calendar.Events.list(calendarId, optionalArgs).items;
    },

    getTab: function (tabName) {
        const sheets = this.getDoc().getSheets();

        for (let iSheet in sheets) {
            if (tabName === sheets[iSheet].getName()) {
                return sheets[iSheet];
            }
        }

        return undefined;
    },

    getDoc: function () {
        return SpreadsheetApp.getActiveSpreadsheet();
    },

    getOrCreateTab: function (tabName) {
        const existingTab = this.getTab(tabName);
        if (existingTab !== undefined) {
            return existingTab;
        }

        const sheet = getDoc().insertSheet();
        sheet.setName(tabName);

        return sheet;
    },

    copyTab: function (fromName, toName, position) {
        const from = this.getTab(fromName);
        const to = from.copyTo(this.getDoc());
        to.setName(toName);

        to.activate();

        this.getDoc().moveActiveSheet(position);

        return to;
    },

    dump: function (sheet, events) {
        const gridStartRow = 40; // TODO Make it parameter
        const gridStartCol = 2; // TODO Make it parameter

        const durationFormat = "[h]:mm";
        const costFormat = "0.00 €";

        let iCol = gridStartCol;

        // Header
        const header = {
            day: {
                COL: iCol++,
                NAME: 'Jour',
                FORMAT: "dddd d"
            },
            periods: {
                COL: iCol++,
                NAME: 'Périodes',
                FORMAT: ""
            },
            deprecated: {
                COL: iCol++,
                NAME: '',
                FORMAT: "HH:mm"
            },
            duration: {
                COL: iCol++,
                NAME: 'Durée',
                FORMAT: durationFormat
            },
            overtime: {
                COL: iCol++,
                NAME: 'Heures sup',
                FORMAT: durationFormat
            },
            maintenance: {
                COL: iCol++,
                NAME: 'Indemnité d\'entretien',
                FORMAT: costFormat
            },
            food: {
                COL: iCol++,
                NAME: 'Repas',
                FORMAT: costFormat
            },
            notes: {
                COL: iCol++,
                NAME: 'Notes',
                FORMAT: ""
            }
        };

        const columns = Object.keys(header);

        // Keywords to use in calendar
        const bank_holidays = 'férié';
        const vacations = 'congé';
        const no_nanny = 'absence Nounou';
        const no_baby = 'absence';
        const no_lunch = 'sans repas';

        if (events.length > 0) {
            let row = gridStartRow;
            const getDay = (date) => date.toLocaleString().split(',')[0]
            const getTime = (date) => date.toLocaleString('en-US', {hour12: false})
                .replace(/.* /, '')
                .replace(/:\d\d$/, '')

            let duration = 0
            let periods = []
            let currentDay = undefined;

            for (i = 0; i < events.length; i++) {
                const event = events[i];
                let startTime = event.start.dateTime;
                if (!startTime) {
                    startTime = event.start.date;
                }
                let endTime = event.end.dateTime;
                if (!endTime) {
                    endTime = event.end.date;
                }
                // Use date objects!
                const end = new Date(endTime);
                const start = new Date(startTime);

                const summary = (!event.summary || event.summary === 'Nouvel événement')
                    ? undefined
                    : event.summary;

                if (currentDay !== undefined && currentDay !== getDay(start)) {
                    // Different day => go to next row
                    row++;
                    duration = 0
                    periods = []
                }

                currentDay = getDay(start)

                periods.push(getTime(start) + '-' + getTime(end))

                const sameRowCol = function (col) {
                    return sheet.getRange(row, col).getA1Notation();
                };

                duration += (end.getTime() - start.getTime()) / 1000 / 3600 / 24;

                const values = {
                    day: start, // Will be formatted
                    periods: periods.join(' '), // Stores a concatenation of periods
                    duration: duration,
                    overtime:
                        '=if(value(' + sameRowCol(header.duration.COL)
                        + '-vlookup(weekday(' + sameRowCol(header.day.COL) + ';2);$F$23:$H$27;3))>0;'
                        + sameRowCol(header.duration.COL) + '-vlookup(weekday(' + sameRowCol(header.day.COL) + ';2);$F$23:$H$27;3);"")',
                    maintenance: '=if(OR('
                        + sameRowCol(header.notes.COL) + '="' + bank_holidays + '";'
                        + sameRowCol(header.notes.COL) + '="' + vacations + '";'
                        + sameRowCol(header.notes.COL) + '="' + no_nanny + '";'
                        + sameRowCol(header.notes.COL) + '="' + no_baby + '");"";MAX($D$20;' + sameRowCol(header.duration.COL) + '*24*$D$21))',
                    food: '=if(OR('
                        + sameRowCol(header.notes.COL) + '="' + bank_holidays + '";'
                        + sameRowCol(header.notes.COL) + '="' + vacations + '";'
                        + sameRowCol(header.notes.COL) + '="' + no_lunch + '";'
                        + sameRowCol(header.notes.COL) + '="' + no_nanny + '";'
                        + sameRowCol(header.notes.COL) + '="' + no_baby + '");"";$G$32)',
                    notes: (!summary)
                        ? '=if(' + sameRowCol(header.overtime.COL) + '<>"";"Dépassement : "&text(' + sameRowCol(header.day.COL) + ';"hh:mm")&" - "&text(' + sameRowCol(header.deprecated.COL) + ';"hh:mm");"")'  // TODO
                        : summary
                };

                columns.forEach((key, index) => {
                    const cell = sheet.getRange(row, header[key].COL);
                    // Set format right now to ensure next row will have previous values in the right format
                    cell.setNumberFormat(header[key].FORMAT)
                    cell.setValue(values[key]);
                });
            }

            return true;
        } else {
            sheet.getRange(gridStartRow, 1).setValue('Aucun évènement');
            Logger.log('No event found');

            return false;
        }
    }
};

function importMonth(date) {
    const calendarImport = new CalendarImport();

    const monthName = date.getFullYear() + "-" + ("0" + (date.getMonth() + 1)).slice(-2);

    // 1) copy template to current month
    const sheet = calendarImport.copyTab("template", monthName, 3);

    // 2) get events from calendar
    const events = calendarImport.getMonthEvents(date);

    Logger.log(events);

    // 3) update tab
    calendarImport.dump(sheet, events);

    sheet.setTabColor(null);
}
