# Attainment Report Generator

React + TypeScript + Vite Power Apps Code App for generating manager-specific attainment reports and Outlook draft emails.

## Workflow

1. Upload a Global Attainment Report (`.xlsx` or `.csv`).
2. Upload a Sales Compensation Report (`.xlsx`).
3. Select regions and generate manager reports.
4. Review managers and create Outlook drafts with the generated Excel attachments.

## Report Logic

- Attainment rows are matched to SCR rows by `LI_EMP_ID` / `Employee ID`.
- `Level_1_Manager` is refreshed from SCR `Supervisory Manager`.
- `Level_2_Manager` is refreshed from SCR `Superior Organization - Level 01 Away`.
- Each manager report starts with that manager's own attainment rows before the team hierarchy section.
- Managers whose generated report rows are all `Employee Status = Terminated` are excluded.
- Report rows are sorted by `LI_EMP_ID` ascending, `Quota Start Date` ascending, and `Measure Weight` descending.

## Commands

```bash
npm run dev
npm run build
npm run lint
```

Power Apps deployment uses the project `power.config.json` and `pac code push`.
