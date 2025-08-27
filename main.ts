// @deno-types="@types/xlsx"
import * as xlsx from "xlsx";
import moment from "moment";
import { parseArgs } from "@std/cli/parse-args";

import { findMerge, timeCellToDate } from "./helpers.ts";

const { help: hasHelp, file, date: dateInputs } = parseArgs(Deno.args, {
  boolean: ["help"],
  string: ["file", "date"],
  collect: ["date"],
});
if (hasHelp) {
  console.error(
    "Usage: %s --date <sheet name>:<2025-09-06> --date <sheet name>:<2025-09-07> --file <file.xlsx>",
    import.meta.filename,
  );
  console.error("Generated json outputted to stdout");
  Deno.exit(0);
}
if (!file) {
  throw new Error("No file.xlxs specified.");
}

const data = await Deno.readFile(file);

// Read the .xlsx file in "dense" mode, otherwise sheets don't have the "!data" property on them
const workbook = xlsx.read(data, {
  dense: true,
});

// Specify the dates corresponding to the sheet names
const dates = new Map(dateInputs.map((dateInput): [string, moment.Moment] => {
  const [sheetName, dateStr] = dateInput.split(":");

  const date = moment(dateStr).startOf("day");
  if (!date.isValid()) {
    throw new Error(`Specified date for "${sheetName}" invalid.`);
  }

  return [sheetName, date];
}));

const roomRow = 0;
const startRow = 1;
const timeCol = 0;
const startCol = 1;

// Helper variable to generate incrementing event ids
let eventIdInc = 0;

interface TimetableEvent {
  eventId: number;
  room: string;
  eventname: string;
  begin: string;
  end: string;
}
const events: TimetableEvent[] = [];

// The magic
for (const [sheetName, startDate] of dates) {
  const sheet = workbook.Sheets[sheetName] as xlsx.DenseWorkSheet; // forced type, bc proper typing is much to be desired in this library
  if (!sheet) {
    throw new Error(
      `The sheet "${sheetName}" doesn't exist, but was specified as a --date flag.`,
    );
  }

  const data = sheet["!data"];
  const merges = sheet["!merges"] as xlsx.Range[]; // forced type, bc proper typing is much to be desired in this library

  if (!data[roomRow]) {
    throw new Error(
      `The room row given for "${sheetName}" doesn't have any data`,
    );
  }

  const lastEventPerCol = new Array<TimetableEvent | undefined>(
    data[roomRow].length,
  ).fill(undefined);

  for (let rowIndex = startRow; rowIndex <= data.length; rowIndex++) {
    const row = data[rowIndex];
    if (!row || !Array.isArray(row)) {
      break;
    }

    for (let colIndex = startCol; colIndex <= row.length; colIndex++) {
      // Lookup the name of the event
      const eventCell = row[colIndex];
      if (!eventCell || !eventCell.w) {
        continue;
      }
      const eventName = eventCell.w.replace(/\s+/g, " ");

      // Lookup the name of the room
      const roomCell = data[roomRow][colIndex];
      if (!roomCell || !roomCell.w) {
        continue;
      }
      const roomName = roomCell.w.replace(/\s+/g, " ");

      // Lookup whether this cell spans multiple cols/rows
      const cellMerge: xlsx.Range = findMerge(merges, colIndex, rowIndex) ??
        // define a default range (1 col/row wide)
        { s: { c: colIndex, r: rowIndex }, e: { c: colIndex, r: rowIndex } };

      // If the cell is more than one column wide, it's not an event, but some other indicator
      if (cellMerge.s.c < cellMerge.e.c) {
        continue;
      }

      // Lookup the start and end times, if they exist
      const startTimeCell = data[cellMerge.s.r]?.[timeCol];
      const endTimeCell = data[cellMerge.e.r + 1]?.[timeCol]; // Add +1 to the end time cell, because actually the row after the last row of the event marks the end

      const startTime = startTimeCell &&
        timeCellToDate(startDate, startTimeCell);
      const endTime = endTimeCell &&
        timeCellToDate(startDate, endTimeCell);

      if (!startTime || !endTime) {
        continue;
      }

      // Timetables where events go on after midnight might have end times (eg. 00:30) before start times (eg. 23:30)
      // so add another day to these times
      if (endTime.isBefore(startTime)) {
        endTime.add(1, "day");
      }

      // Check out the time of the last event in this room, to make sure that when an event started after midnight, time is
      // properly parsed to be on the next day
      // TODO: if the only event in this room is after midnight, this will not work
      //       maybe map all times in the timeCol to actual dates first and then use that list instead
      const lastEvent = lastEventPerCol[colIndex];
      if (lastEvent && moment(lastEvent.begin).isAfter(startTime)) {
        startTime.add(1, "day");
        endTime.add(1, "day");
      }

      const event: TimetableEvent = {
        eventId: eventIdInc++,
        begin: startTime.toISOString(true),
        end: endTime.toISOString(true),
        eventname: eventName,
        room: roomName,
      };

      lastEventPerCol[colIndex] = event;

      events.push(event);
    }
  }
}

// Finally output to stdout
console.log(JSON.stringify(events, null, 2));
