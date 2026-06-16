# Custom Power BI Gantt Visual

This project contains a custom Power BI visual that renders a Gantt chart with planned vs actual timelines.

## Included Data Fields

- Sequence
- Task Name
- Start Date
- End Date
- Planned Start Date
- Planned Finish Date
- Group 1
- Group 2
- Group 3
- Progress %
- Planned Progress %
- Status
- URL

## Build and Package

1. Install Node.js LTS (if not already installed).
2. Open a terminal in this folder.
3. Install dependencies:

   ```powershell
   npm install
   ```

4. Run the visual in developer mode:

   ```powershell
   npm run start
   ```

5. Package the visual:

   ```powershell
   npm run package
   ```

The packaged visual file is generated in the `dist` folder.

## Notes

- Use `Sequence` to control row order.
- `Progress %` and `Planned Progress %` accept values in either `0-1` or `0-100` format.
- If a `URL` is supplied, clicking the task label launches that link.
