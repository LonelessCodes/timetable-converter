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

  const timeToday = moment(timeRaw, "HH:mm");

  const timeSinceStartToday = moment.duration({
    from: timeToday.clone().startOf("day"),
    to: timeToday,
  });

  return startDate.clone().add(timeSinceStartToday);
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
