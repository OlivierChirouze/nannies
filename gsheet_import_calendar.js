function getPreviousMonth() {
    var previousMonth = new Date();
    previousMonth.setMonth(previousMonth.getMonth() - 1);
    return previousMonth;
}

function onOpen() {
    var currentMonth = new Date();
    var previousMonth = getPreviousMonth();

    var ui = SpreadsheetApp.getUi();
    // Or DocumentApp or FormApp.
    ui.createMenu('Calendar import')
        .addItem('Import ' + formatDate(previousMonth), 'importPreviousMonth')
        .addItem('Import ' + formatDate(currentMonth), 'importCurrentMonth')
        .addToUi();
}

function formatDate(date) {
    return date.toLocaleDateString(undefined, { year: 'numeric', month: 'long' });
}

function importPreviousMonth() {
    importMonth(getPreviousMonth());
}

function importCurrentMonth() {
    importMonth(new Date());
}

// https://developers.google.com/apps-script/guides/sheets
var CalendarImport = function () {
    return this;
};

// Shouldn't be more than 1 event per day!
CalendarImport.maxMonthEvents = 300;

CalendarImport.prototype = {
    getMonthEvents: function (currentDate) {
        var calendarId = 'ie7d48frttnbcj8678ln7lv860@group.calendar.google.com';

        // First day of month
        var monthBegin = new Date(currentDate.getFullYear(), currentDate.getMonth(), 1);
        // Last day of month, end of day
        var monthEnd = new Date(currentDate.getFullYear(), currentDate.getMonth() + 1, 0, 23, 59, 59);

        var optionalArgs = {
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
        var sheets = this.getDoc().getSheets();

        for (var iSheet in sheets) {
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
        var existingTab = this.getTab(tabName);
        if (existingTab != undefined) {
            return existingTab;
        }

        var sheet = getDoc().insertSheet();
        sheet.setName(tabName);

        return sheet;
    },

    copyTab: function (fromName, toName, position) {
        var from = this.getTab(fromName);
        var to = from.copyTo(this.getDoc());
        to.setName(toName);

        to.activate();

        this.getDoc().moveActiveSheet(position);

        return to;
    },

    dump: function (sheet, events) {
        var row = 0;

        const gridStartRow = 40; // TODO Make it parameter
        const gridStartCol = 2; // TODO Make it parameter

        const durationFormat = "[h]:mm";
        const costFormat = "0.00 €";

        var iCol = gridStartCol;

        // Header
        const header = {
            day: {
                COL: iCol++,
                NAME: 'Jour',
                FORMAT: "dddd d"
            },
            start: {
                COL: iCol++,
                NAME: 'Arrivée',
                FORMAT: "HH:mm"
            },
            end: {
                COL: iCol++,
                NAME: 'Départ',
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

        if (events.length > 0) {
            for (i = 0; i < events.length; i++) {
                row = i + gridStartRow;

                var sameRowCol = function (col) {
                    return sheet.getRange(row, col).getA1Notation();
                }

                var event = events[i];
                var startTime = event.start.dateTime;
                if (!startTime) {
                    startTime = event.start.date;
                }
                var endTime = event.end.dateTime;
                if (!endTime) {
                    endTime = event.end.date;
                }
                // Use date objects!
                var end = new Date(endTime);
                var start = new Date(startTime);

                var summary = (!event.summary || event.summary == 'Nouvel événement')
                    ? undefined
                    : event.summary;

                // Keywords to use in calendar
                const bank_holidays = 'férié';
                const vacations = 'congé';
                const no_nanny = 'absence Nounou';
                const no_baby = 'absence';
                const no_lunch = 'sans repas';

                var values = {
                    day: start, // Will be formatted
                    start: start,
                    end: end,
                    duration: "="
                        + sameRowCol(header.end.COL) + "-"
                        + sameRowCol(header.start.COL)
                    ,
                    overtime:
                        '=if(value(' + sameRowCol(header.duration.COL)
                        + '-vlookup(weekday(' + sameRowCol(header.start.COL) + ';2);$F$23:$H$27;3))>0;'
                        + sameRowCol(header.duration.COL) + '-vlookup(weekday(' + sameRowCol(header.start.COL) + ';2);$F$23:$H$27;3);"")',
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
                        ? '=if(' + sameRowCol(header.overtime.COL) + '<>"";"Dépassement : "&text(' + sameRowCol(header.start.COL) + ';"hh:mm")&" - "&text(' + sameRowCol(header.end.COL) + ';"hh:mm");"")'  // TODO
                        : summary
                };

                columns.forEach(function (key, index) {
                    sheet.getRange(row, header[key].COL).setValue(values[key]);
                });
            }

            // Now format all columns
            columns.forEach(function (key, index) {
                sheet.getRange(gridStartRow, header[key].COL, events.length).setNumberFormat(header[key].FORMAT);
            });

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

    var monthName = date.getFullYear() + "-" + ("0" + (date.getMonth() + 1)).slice(-2);

    // 1) copy template to current month
    var sheet = calendarImport.copyTab("template", monthName, 3);

    // 2) get events from calendar
    var events = calendarImport.getMonthEvents(date);

    // 3) update tab
    calendarImport.dump(sheet, events);

    sheet.setTabColor(null);
}
