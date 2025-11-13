# Google Timeline to CRA Vehicle Logbook

A Python tool that converts Google Timeline JSON exports into CRA-compliant vehicle logbooks with automated address resolution, odometer tracking, and intelligent filtering.

## Features

- 📍 **Automated Geocoding** - Converts Google Place IDs to full addresses using Google Maps API
- 🚗 **Odometer Calculations** - Automatically calculates start/end odometer readings for each trip
- 📊 **CRA Compliance** - Generates logbooks meeting Canadian Revenue Agency requirements
- 🔍 **Smart Filtering** - Removes duplicate trips, impossible sequences, and applies custom filters
- 📝 **Purpose Assignment** - Automatically categorizes trips based on location
- 📈 **Distance Tracking** - Calculates total kilometers driven with configurable thresholds

## Requirements

- Python 3.7+
- Google Maps API key (for geocoding)
- Google Service Account credentials (for Sheets API access)
- Google Timeline JSON export

## Installation

1. Clone this repository:
```bash
git clone https://github.com/BananaRama123/Timeline2Sheets
cd timeline2sheets
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

3. Set up your credentials:
   - Place your Google service account JSON in the project directory as `credentials.json`
   - Get a Google Maps API key from [Google Cloud Console](https://console.cloud.google.com/)
   - Add your IP Address to Maps Platform API Key in Application restrictions or temporary remove all restrictions

## Configuration

Edit the configuration section in `timeline_complete_pipeline.py`:

```python
# Google API Credentials
GOOGLE_MAPS_API_KEY = "your-maps-api-key"
SHEETS_CREDENTIALS_PATH = "credentials.json"
SPREADSHEET_ID = "your-google-sheet-id"

# JSON File Settings
TIMELINE_JSON_PATH = "Timelinechunk.json"
SKIP_JSON_IMPORT = False  # Set to True to skip JSON import

# Processing Settings
START_DATE = "2024-09-30"  # Fiscal year start
END_DATE = "2025-10-01"    # Fiscal year end

# Odometer Settings
ODOMETER_END_DATE = "2025-10-01"  # Date of known odometer reading
ODOMETER_END_READING = 172000      # Odometer reading on that date
```

## Usage

### Step 1: Export Your Google Timeline Data

1. Go to [Google Takeout](https://takeout.google.com/)
2. Deselect all, then select only "Location History"
3. Choose JSON format
4. Download and extract the Timeline data

### Step 2: Share Your Google Sheet

Share your Google Sheet with the service account email:
```
your-service-account@project-id.iam.gserviceaccount.com
```
Grant **Editor** permissions.

### Step 3: Run the Script

```bash
python timeline_complete_pipeline.py
```

## Output Format

The script generates a CRA-compliant logbook with the following columns:

| Column | Description |
|--------|-------------|
| Item | Sequential trip number |
| Date | Date of trip (YYYY-MM-DD) |
| Start Time | Trip start time |
| End Time | Trip end time |
| Starting Point | Origin address |
| Destination | Destination address |
| Purpose of Trip | Business purpose (auto-assigned or blank) |
| Km Driven | Distance traveled |
| Start Odometer | Odometer reading at trip start |
| End Odometer | Odometer reading at trip end |
| Duration (min) | Trip duration in minutes |
| Activity_type | Type of transportation |

## How It Works

### 1. JSON Import & Address Resolution
- Reads Google Timeline semantic segments (visits and activities)
- Resolves Place IDs to full addresses using Google Maps Geocoding API
- Caches results to minimize API calls

### 2. Trip Processing
- Identifies vehicle trips (`IN_PASSENGER_VEHICLE` activities)
- Matches trips with origin (previous visit) and destination (next visit)
- Extracts dates, times, and distances

### 3. Odometer Calculation
- Works backwards from a known odometer reading
- Calculates start/end readings for each trip
- Accounts for ALL vehicle trips (including filtered ones)

### 4. Intelligent Filtering
- Removes trips outside specified regions
- Filters impossible consecutive trips (same origin twice)
- Excludes specific locations or date ranges
- Applies minimum distance thresholds

### 5. Purpose Assignment
- Automatically categorizes trips based on destination
- Supports custom city-to-purpose mappings
- Leaves unmatched trips blank for manual entry

### 6. Sheet Generation
- Writes formatted data to Google Sheets
- Applies professional styling (bold headers, centered columns)
- Includes total distance calculation

## Customization

### Filtering Trips

Customize filters in the `write_final_report()` method:

```python
# Example: Filter by province/state
start_in_region = 'desired_region' in start_addr.lower()
dest_in_region = 'desired_region' in dest_addr.lower()

# Example: Exclude specific cities
excluded_cities = ['city1', 'city2', 'city3']
```

### Purpose Assignment

Define custom purpose mappings:

```python
# Cities for specific purposes
customer_site_cities = ['city1', 'city2', 'city3']
meeting_cities = ['city4', 'city5', 'city6']

# Assign purpose based on location
for city in customer_site_cities:
    if city in dest_addr or city in start_addr:
        purpose = 'Your Custom Purpose'
```

### Distance Thresholds

Adjust minimum trip distance:

```python
# Only include trips >= 15km
if distance >= 15:
    filtered_data.append(entry)
```

## Configuration Options

### Skip JSON Import

If you've already imported your Timeline data, skip the import step:

```python
SKIP_JSON_IMPORT = True
```

### Date Range Filtering

Process only specific fiscal years or date ranges:

```python
START_DATE = "2024-01-01"
END_DATE = "2024-12-31"
```

### Odometer Settings

Set your known odometer reading as a reference point:

```python
ODOMETER_END_DATE = "2025-01-01"  # Any date with known reading
ODOMETER_END_READING = 150000      # Actual odometer value
```

## API Costs

- **Google Maps Geocoding API**: ~$5 per 1,000 Place ID lookups
- **Google Sheets API**: Free (within quota limits)

The script uses caching to minimize API calls. Typical usage:
- First run: Resolves all unique Place IDs
- Subsequent runs: Uses cached addresses

## Troubleshooting

### "403 Forbidden" Error
- Ensure you've shared the Google Sheet with the service account email
- Verify the service account has "Editor" permissions

### Missing Addresses
- Check that your Google Maps API key is valid
- Ensure the Geocoding API is enabled in Google Cloud Console

### Incorrect Odometer Readings
- Verify your reference date and reading are correct
- Check that trips are sorted chronologically

### No Data in Timeline JSON
- Ensure you exported "Location History" not "Maps"
- Verify the JSON contains `semanticSegments`

## Contributing

Contributions are welcome! Please feel free to submit pull requests or open issues.

## License

MIT License - see LICENSE file for details

## Disclaimer

This tool is provided as-is for generating vehicle logbooks. Always verify the accuracy of your logbook data before submitting it for tax purposes. The authors are not responsible for any errors or compliance issues.

## Support

For issues or questions:
- Open an issue on GitHub
- Check existing issues for solutions
- Review the configuration section carefully

---

**Note**: This tool processes location data. Ensure you comply with all applicable privacy laws and regulations when handling location information.
