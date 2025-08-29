import moment from "moment";
// @deno-types="@types/xlsx"
import * as xlsx from "xlsx";

export function timeCellToDate(
  startDate: moment.Moment,
  cell: xlsx.CellObject,
) {
  const timeRaw = cell.w;
  if (!timeRaw) {
    return null;
  }

  const match = timeRaw.match(/(\d+)\:(\d+)/);
  if (!match) {
    return null;
  }

  const [_, hourStr, minuteStr] = match;
  const [hour, minute] = [parseInt(hourStr), parseInt(minuteStr)];

  // Prefer setting instead of adding hours and minutes. Remember daylight savings dates!
  return startDate.clone().set({
    hour,
    minute,
  });
}

export function findMerge(
  merges: xlsx.Range[],
  colIndex: number,
  rowIndex: number,
) {
  return merges.find((merge) =>
    merge.s.c === colIndex && merge.s.r === rowIndex
  );
}
