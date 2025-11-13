# CricAuction Scraper

This project scrapes upcoming cricket auction data from [CricAuction](https://cricauction.live/upcoming-auction) and stores it locally in a CSV file. It also syncs the data to a Microsoft SharePoint Excel worksheet.

## Features

- **Web Scraping**: Uses Selenium to scrape auction details (tournament name, location, total players, auction date)
- **Local Storage**: Saves data to `cricauction_upcoming.csv`
- **SharePoint Integration**: Syncs data to a SharePoint Excel worksheet using Microsoft Graph API
- **Automated Daily Runs**: GitHub Actions workflow runs daily at 9:00 AM IST

## Requirements

- Python 3.11+
- Chrome/Chromium browser
- ChromeDriver (for local execution)

## Installation

1. Clone the repository:
```bash
git clone https://github.com/Darshilshahcricbattle/CricAuction.git
cd CricAuction
```

2. Install dependencies:
```bash
pip install selenium requests
```

## Configuration

### SharePoint Credentials (Optional)

If you want to sync data to SharePoint, create a `credential.json` file:

```json
{
  "tenant_id": "your-tenant-id",
  "client_id": "your-client-id",
  "client_secret": "your-client-secret"
}
```

**Note**: This file is excluded from Git for security. The script will work without it (local CSV only).

## Usage

Run the scraper manually:
```bash
python scrape_auctions_updated.py
```

## GitHub Actions

The scraper runs automatically every day at 9:00 AM IST via GitHub Actions.

### Setting up GitHub Secrets

To enable SharePoint sync in GitHub Actions:

1. Go to your repository on GitHub
2. Navigate to **Settings** → **Secrets and variables** → **Actions**
3. Click **New repository secret**
4. Add a secret named `CREDENTIAL_JSON` with your credentials JSON content

### Manual Trigger

You can manually trigger the workflow from the **Actions** tab in your GitHub repository.

## Output

- **Local**: `cricauction_upcoming.csv` - Contains all scraped auction data
- **SharePoint**: Data is synced to the configured SharePoint Excel worksheet (if credentials are provided)

## License

MIT
