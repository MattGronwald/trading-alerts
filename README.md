# Google Apps Script for Trading Signals

A Google Apps Script project that monitors stock price movements and sends email alerts based on:
- Golden Cross signals (50 SMA crosses above 200 SMA)
- Death Cross signals (50 SMA crosses below 200 SMA)

## Google Sheet Structure

The script works with a Google Sheet containing the following structure:
| Column         | Description                                  | Data Type     |
|:---------------|:---------------------------------------------|:--------------|
| Symbol         | Stock symbol (e.g., "AMZN")                  | Text          |
| Name           | Company name (e.g., "Amazon.com Inc")        | Text          |
| Close          | Latest stock price                           | Number        |
| 50 SMA         | 50-day Simple Moving Average                 | Number        |
| 200 SMA        | 200-day Simple Moving Average                | Number        |
| Cross Signal   | Cross signal status ("GOLDEN", "DEATH")      | Text          |
| Previous Signal| Previous cross signal status                 | Text          |
| Last Change    | Last update date                             | Date/Time     |
| Currency       | Currency of the stock (e.g., "USD", "EUR")   | Text          |

## SMA Calculation in the Google Sheet

The 50-day and 200-day Simple Moving Averages (SMAs) are calculated using Google Sheet formulas that retrieve historical stock data:

### 50-day SMA Formula
```
=IF(A2=""; ""; ROUND(AVERAGE(QUERY(GOOGLEFINANCE(A2; "close"; TODAY()-100; TODAY()); "select Col2 where Col2 is not null order by Col1 desc limit 50"; 0));2))
```

### 200-day SMA Formula
```
=IF(A2=""; ""; ROUND(AVERAGE(QUERY(GOOGLEFINANCE(A2; "close"; TODAY()-300; TODAY()); "select Col2 where Col2 is not null order by Col1 desc limit 200"; 0));2))
```

These formulas:
1. Use GOOGLEFINANCE to retrieve historical closing prices
2. Query the last 50 or 200 non-null values
3. Calculate the average and round to 2 decimal places
4. Include error handling for empty cells

## Features

- Automated detection of Golden Cross and Death Cross events
- Email notifications with detailed stock information
- Current price and percentage calculations relative to moving averages
- Currency information displayed for each stock in email notifications
- Organized tables of buy and sell signals with relevant metrics

This repository contains a Google Apps Script project managed locally using `clasp` (Command Line Apps Script Projects) and Node.js.

## Project Setup

### Prerequisites

- [Node.js](https://nodejs.org/) installed (recommended version: 20.x)
- [nvm](https://github.com/nvm-sh/nvm) for Node.js version management (optional but recommended)
- [clasp](https://github.com/google/clasp) installed globally

Install clasp globally:
```bash
npm install -g @google/clasp
```

Login to your Google account:
```bash
npx clasp login
```

### Cloning the Existing Script

Clone the Google Apps Script project using the Script ID:
```bash
npx clasp clone <SCRIPT-ID>
```

Replace `<SCRIPT-ID>` with your actual Google Apps Script ID.

---

## Local Development Workflow

After cloning the project:

- Pull the latest changes:
  ```bash
  npm run pull
  ```

- Make your changes locally in `.gs` or `.html` files.

- Push your changes back to Apps Script:
  ```bash
  npm run push
  ```

- Open the project in the online editor (optional):
  ```bash
  npm run open
  ```

---

## Scripts

Defined in `package.json` under `"scripts"`:

| Script       | Command             | Description                                |
|:-------------|:---------------------|:-------------------------------------------|
| **pull**     | `npx clasp pull`       | Pull latest remote changes |
| **push**     | `npx clasp push`       | Push local changes to Google Apps Script |
| **open**     | `npx clasp open`       | Open the project in the Google Script Editor |
| **deploy**   | `npx clasp deploy`     | Create a new version (if needed) |

---

## Notes

- `node_modules/` and `.clasp.json` are excluded via `.gitignore`.
- Always `pull` before starting to avoid conflicts.
- Always `push` after making local changes.

---

## License

This project is licensed under the [MIT License](LICENSE) unless otherwise specified.
