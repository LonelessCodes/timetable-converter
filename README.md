# timetable-converter

This is a convenience script to convert a spreadsheet with an event timetable format like this:

<table style="text-align: center; width: 100%">
    <thead>
        <tr>
            <th>Time</th>
            <th>Room 1</th>
            <th>Room 2</th>
        </tr>
    </thead>
    <tbody>
        <tr>
            <td>12:00</td>
            <td rowspan=4>Opening Ceremony</td>
            <td></td>
        </tr>
        <tr>
            <td>12:15</td>
            <td></td>
        </tr>
        <tr>
            <td>12:30</td>
            <td rowspan=3>Crazy Panel</td>
        </tr>
        <tr>
            <td>12:45</td>
        </tr>
        <tr>
            <td>13:00</td>
            <td></td>
        </tr>
        <tr>
            <td>13:15</td>
            <td></td>
            <td></td>
        </tr>
        <tr>
            <td>13:30</td>
            <td colspan=2>Closing</td>
        </tr>
    </tbody>
</table>

where one event can span multiple rows, into a JSON readable array of events.

## Usage

```sh
deno run main --date Friday:2025-05-20 --date Saturday:2025-05-21 --file file.xlsx > events.json

    --date <sheet name>:<date> # Maps a date to a sheet name in the specified workfile. Multiple allowed. Not mentioned sheets will be ignored.
    --file <file.xlsx> # Path to the spreadsheet file

The script outputs the json to stdout.
```
