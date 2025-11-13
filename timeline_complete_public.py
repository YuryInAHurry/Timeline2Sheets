#!/usr/bin/env python3
"""
Complete Google Timeline Pipeline
1. Converts Google Timeline JSON to Google Sheets with resolved addresses
2. Filters and processes data into Final Report with proper distance associations
"""

import json
import time
from datetime import datetime
from typing import List, Dict, Any, Optional
import requests
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


class CompleteTimelinePipeline:
    def __init__(self, google_maps_api_key: str, sheets_credentials_path: str, spreadsheet_id: str):
        """
        Initialize the complete pipeline
        
        Args:
            google_maps_api_key: Your Google Maps API key
            sheets_credentials_path: Path to Google Sheets service account JSON
            spreadsheet_id: Google Sheets spreadsheet ID
        """
        self.maps_api_key = google_maps_api_key
        self.spreadsheet_id = spreadsheet_id
        self.geocode_cache = {}
        self.reverse_geocode_cache = {}
        
        # Initialize Google Sheets API
        self.sheets_service = self._init_sheets_service(sheets_credentials_path)
    
    def _init_sheets_service(self, credentials_path: str):
        """Initialize Google Sheets API service"""
        SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
        creds = service_account.Credentials.from_service_account_file(
            credentials_path, scopes=SCOPES)
        return build('sheets', 'v4', credentials=creds)
    
    # ========================================================================
    # PART 1: JSON TO SHEETS (Original timeline_to_sheets_v2.py functionality)
    # ========================================================================
    
    def geocode_place_id(self, place_id: str) -> Dict[str, str]:
        """Convert Google Place ID to formatted address"""
        if place_id in self.geocode_cache:
            return self.geocode_cache[place_id]
        
        url = "https://maps.googleapis.com/maps/api/place/details/json"
        params = {
            'place_id': place_id,
            'key': self.maps_api_key
        }
        
        try:
            response = requests.get(url, params=params, timeout=10)
            response.raise_for_status()
            data = response.json()
            
            # Changed: data['result'] instead of data['results'][0]
            if data['status'] == 'OK' and 'result' in data:
                result = data['result']  # Single object, not array
                address_info = {
                    'formatted_address': result.get('formatted_address', ''),
                    'name': result.get('name', ''),
                    'place_id': place_id
                }
                
                for component in result.get('address_components', []):
                    types = component['types']
                    if 'street_number' in types:
                        address_info['street_number'] = component['long_name']
                    elif 'route' in types:
                        address_info['street'] = component['long_name']
                    elif 'locality' in types:
                        address_info['city'] = component['long_name']
                    elif 'administrative_area_level_1' in types:
                        address_info['state'] = component['short_name']
                    elif 'country' in types:
                        address_info['country'] = component['long_name']
                    elif 'postal_code' in types:
                        address_info['postal_code'] = component['long_name']
                
                self.geocode_cache[place_id] = address_info
                time.sleep(0.1)
                return address_info
            else:
                print(f"  ⚠️ API Status: {data.get('status')}, Error: {data.get('error_message', 'N/A')}")
                return {'formatted_address': f'Unknown Place ID', 'place_id': place_id}
                
        except Exception as e:
            print(f"  ❌ Error geocoding {place_id}: {e}")
            return {'formatted_address': f'Error resolving address', 'place_id': place_id}
        
    def reverse_geocode_coords(self, lat: float, lng: float) -> str:
        """Reverse geocode coordinates to an address"""
        cache_key = f"{lat:.6f},{lng:.6f}"
        
        if cache_key in self.reverse_geocode_cache:
            return self.reverse_geocode_cache[cache_key]
        
        url = "https://maps.googleapis.com/maps/api/geocode/json"
        params = {
            'latlng': f"{lat},{lng}",
            'key': self.maps_api_key
        }
        
        try:
            response = requests.get(url, params=params, timeout=10)
            response.raise_for_status()
            data = response.json()
            
            if data['status'] == 'OK' and data['results']:
                address = data['results'][0].get('formatted_address', 'Unknown')
                self.reverse_geocode_cache[cache_key] = address
                time.sleep(0.1)
                return address
            else:
                return f"({lat:.6f}, {lng:.6f})"
                
        except Exception as e:
            print(f"  ❌ Error reverse geocoding {lat},{lng}: {e}")
            return f"({lat:.6f}, {lng:.6f})"
    
    def parse_latlng_string(self, latlng_str: str) -> tuple[Optional[float], Optional[float]]:
        """Parse a lat/lng string like '43.5707239°, -79.5797226°'"""
        try:
            parts = latlng_str.replace('°', '').split(',')
            lat = float(parts[0].strip())
            lng = float(parts[1].strip())
            return lat, lng
        except:
            return None, None
    
    def parse_timeline_json(self, json_path: str) -> List[Dict[str, Any]]:
        """Parse Google Timeline JSON file with semanticSegments"""
        print(f"📖 Reading JSON file: {json_path}")
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        records = []
        
        if 'semanticSegments' not in data:
            print("❌ Error: JSON does not contain 'semanticSegments'. Please check the file format.")
            return records
        
        segments = data['semanticSegments']
        print(f"✅ Found {len(segments)} semantic segments")
        
        for i, segment in enumerate(segments):
            if i % 50 == 0:
                print(f"  Processing segment {i+1}/{len(segments)}...")
            
            if 'visit' in segment:
                record = self._parse_visit(segment)
                if record:
                    records.append(record)
            elif 'activity' in segment:
                record = self._parse_activity(segment)
                if record:
                    records.append(record)
            elif 'timelinePath' in segment:
                record = self._parse_timeline_path(segment)
                if record:
                    records.append(record)
        
        print(f"✅ Parsed {len(records)} total records")
        return records
    
    def _parse_visit(self, segment: Dict) -> Dict[str, Any]:
        """Parse a visit (place) segment"""
        visit = segment.get('visit', {})
        top_candidate = visit.get('topCandidate', {})
        
        start_time = self._parse_timestamp(segment.get('startTime'))
        end_time = self._parse_timestamp(segment.get('endTime'))
        duration_minutes = self._calculate_duration(segment.get('startTime'), segment.get('endTime'))
        
        place_location = top_candidate.get('placeLocation', {})
        latlng_str = place_location.get('latLng', '')
        lat, lng = self.parse_latlng_string(latlng_str)
        
        record = {
            'type': 'Visit',
            'start_time': start_time,
            'end_time': end_time,
            'duration_minutes': duration_minutes,
            'place_id': top_candidate.get('placeId', ''),
            'semantic_type': top_candidate.get('semanticType', 'UNKNOWN'),
            'probability': top_candidate.get('probability', 0),
            'lat': lat,
            'lng': lng,
            'hierarchy_level': visit.get('hierarchyLevel', 0),
            'visit_confidence': visit.get('probability', 0)
        }
        
        return record
    
    def _parse_activity(self, segment: Dict) -> Dict[str, Any]:
        """Parse an activity (movement) segment"""
        activity = segment.get('activity', {})
        start_loc = activity.get('start', {})
        end_loc = activity.get('end', {})
        top_candidate = activity.get('topCandidate', {})
        
        start_time = self._parse_timestamp(segment.get('startTime'))
        end_time = self._parse_timestamp(segment.get('endTime'))
        duration_minutes = self._calculate_duration(segment.get('startTime'), segment.get('endTime'))
        
        start_lat, start_lng = self.parse_latlng_string(start_loc.get('latLng', ''))
        end_lat, end_lng = self.parse_latlng_string(end_loc.get('latLng', ''))
        
        distance_meters = activity.get('distanceMeters', 0)
        distance_km = distance_meters / 1000 if distance_meters else 0
        
        record = {
            'type': 'Activity',
            'start_time': start_time,
            'end_time': end_time,
            'duration_minutes': duration_minutes,
            'activity_type': top_candidate.get('type', 'UNKNOWN'),
            'activity_confidence': top_candidate.get('probability', 0),
            'distance_km': round(distance_km, 2),
            'distance_meters': distance_meters,
            'start_lat': start_lat,
            'start_lng': start_lng,
            'end_lat': end_lat,
            'end_lng': end_lng,
        }
        
        return record
    
    def _parse_timeline_path(self, segment: Dict) -> Optional[Dict[str, Any]]:
        """Parse a timeline path segment (raw location points)"""
        timeline_path = segment.get('timelinePath', [])
        
        if not timeline_path:
            return None
        
        start_time = self._parse_timestamp(segment.get('startTime'))
        end_time = self._parse_timestamp(segment.get('endTime'))
        
        first_point = timeline_path[0]
        last_point = timeline_path[-1]
        
        first_lat, first_lng = self.parse_latlng_string(first_point.get('point', ''))
        last_lat, last_lng = self.parse_latlng_string(last_point.get('point', ''))
        
        record = {
            'type': 'Location Path',
            'start_time': start_time,
            'end_time': end_time,
            'num_points': len(timeline_path),
            'first_point_lat': first_lat,
            'first_point_lng': first_lng,
            'last_point_lat': last_lat,
            'last_point_lng': last_lng,
        }
        
        return record
    
    def _parse_timestamp(self, timestamp_str: str) -> str:
        """Parse ISO timestamp to readable format"""
        if not timestamp_str:
            return ''
        try:
            dt = datetime.fromisoformat(timestamp_str)
            return dt.strftime('%Y-%m-%d %H:%M:%S')
        except:
            return timestamp_str
    
    def _calculate_duration(self, start_str: str, end_str: str) -> int:
        """Calculate duration in minutes between two timestamps"""
        if not start_str or not end_str:
            return 0
        try:
            start = datetime.fromisoformat(start_str)
            end = datetime.fromisoformat(end_str)
            duration = (end - start).total_seconds() / 60
            return int(duration)
        except:
            return 0
    
    def resolve_addresses(self, records: List[Dict[str, Any]], resolve_activities: bool = True) -> List[Dict[str, Any]]:
        """Resolve Place IDs and coordinates to addresses for all records"""
        print(f"\n🔍 Resolving addresses for {len(records)} records...")
        
        visit_count = 0
        activity_count = 0
        
        for i, record in enumerate(records):
            if (i + 1) % 20 == 0:
                print(f"  Progress: {i+1}/{len(records)} records processed...")
            
            if record.get('type') == 'Visit' and record.get('place_id'):
                place_id = record['place_id']
                address_info = self.geocode_place_id(place_id)
                record['address'] = address_info.get('formatted_address', '')
                record['place_name'] = address_info.get('name', '')
                visit_count += 1
            
            if record.get('type') == 'Activity' and resolve_activities:
                if record.get('start_lat') and record.get('start_lng'):
                    record['start_address'] = self.reverse_geocode_coords(
                        record['start_lat'], record['start_lng']
                    )
                
                if record.get('end_lat') and record.get('end_lng'):
                    record['end_address'] = self.reverse_geocode_coords(
                        record['end_lat'], record['end_lng']
                    )
                activity_count += 1
        
        print(f"✅ Resolved {visit_count} visit addresses and {activity_count} activity locations")
        return records
    
    def write_to_sheet(self, records: List[Dict[str, Any]], sheet_name: str):
        """Write records to Google Sheet"""
        if not records:
            print("❌ No records to write!")
            return
        
        print(f"\n📊 Preparing data for Google Sheets...")
        
        headers = self._get_headers(records)
        
        rows = [headers]
        for record in records:
            row = []
            for header in headers:
                value = record.get(header, '')
                if isinstance(value, (int, float)):
                    row.append(str(value))
                else:
                    row.append(str(value) if value else '')
            rows.append(row)
        
        body = {'values': rows}
        
        try:
            print(f"📤 Writing {len(rows)} rows to '{sheet_name}' tab...")
            
            # Clear existing data first
            self.sheets_service.spreadsheets().values().clear(
                spreadsheetId=self.spreadsheet_id,
                range=f"'{sheet_name}'!A:Z"
            ).execute()
            
            # Write new data
            result = self.sheets_service.spreadsheets().values().update(
                spreadsheetId=self.spreadsheet_id,
                range=f"'{sheet_name}'!A1",
                valueInputOption='RAW',
                body=body
            ).execute()
            
            updated_rows = result.get('updatedRows', 0)
            print(f"✅ Successfully wrote {updated_rows} rows to '{sheet_name}' tab!")
            
        except HttpError as error:
            print(f"❌ An error occurred: {error}")
    
    def _get_headers(self, records: List[Dict[str, Any]]) -> List[str]:
        """Determine appropriate headers based on records"""
        all_keys = set()
        for record in records:
            all_keys.update(record.keys())
        
        priority_fields = [
            'type', 'start_time', 'end_time', 'duration_minutes',
            'place_name', 'address', 'semantic_type', 'place_id', 'probability',
            'visit_confidence', 'hierarchy_level', 'activity_type', 'activity_confidence',
            'distance_km', 'distance_meters', 'start_address', 'end_address',
            'lat', 'lng', 'start_lat', 'start_lng', 'end_lat', 'end_lng',
            'num_points', 'first_point_lat', 'first_point_lng', 'last_point_lat', 'last_point_lng',
        ]
        
        headers = []
        for field in priority_fields:
            if field in all_keys:
                headers.append(field)
                all_keys.remove(field)
        
        headers.extend(sorted(all_keys))
        return headers
    
    # ========================================================================
    # PART 2: PROCESS TO FINAL REPORT (Original process_timeline_report.py)
    # ========================================================================
    
    def read_timeline_data(self, sheet_name: str = "Timeline Data") -> List[Dict[str, Any]]:
        """Read data from Timeline Data tab"""
        print(f"\n📖 Reading data from '{sheet_name}' tab...")
        
        try:
            result = self.sheets_service.spreadsheets().values().get(
                spreadsheetId=self.spreadsheet_id,
                range=f"'{sheet_name}'!A:Z"
            ).execute()
            
            values = result.get('values', [])
            
            if not values:
                print("❌ No data found!")
                return []
            
            headers = values[0]
            records = []
            
            for row in values[1:]:
                while len(row) < len(headers):
                    row.append('')
                
                record = {}
                for i, header in enumerate(headers):
                    record[header] = row[i] if i < len(row) else ''
                records.append(record)
            
            print(f"✅ Read {len(records)} records with {len(headers)} columns")
            return records
            
        except HttpError as error:
            print(f"❌ Error reading sheet: {error}")
            return []
    
    def parse_datetime(self, date_str: str) -> datetime:
        """Parse datetime string in format YYYY-MM-DD HH:MM:SS"""
        if not date_str or date_str.strip() == '':
            return None
        try:
            return datetime.strptime(date_str.strip(), '%Y-%m-%d %H:%M:%S')
        except:
            try:
                return datetime.strptime(date_str.strip(), '%Y-%m-%d')
            except:
                return None
    
    def filter_by_date_range(self, records: List[Dict], start_date_str: str, end_date_str: str) -> List[Dict]:
        """Filter records by date range (fiscal year)"""
        print(f"\n🗓️  Filtering by date range: {start_date_str} to {end_date_str}")
        
        start_date = datetime.strptime(start_date_str, '%Y-%m-%d')
        end_date = datetime.strptime(end_date_str, '%Y-%m-%d')
        
        # Debug: Check first few records
        print(f"\n🔍 DEBUG: Checking first 3 records...")
        for i, record in enumerate(records[:3]):
            print(f"  Record {i+1}:")
            print(f"    Available keys: {list(record.keys())}")
            print(f"    start_time value: '{record.get('start_time', 'NOT FOUND')}'")
            parsed = self.parse_datetime(record.get('start_time', ''))
            print(f"    Parsed datetime: {parsed}")
            if parsed:
                print(f"    In range? {start_date <= parsed <= end_date}")
        
        filtered = []
        parse_errors = 0
        for record in records:
            start_time = self.parse_datetime(record.get('start_time', ''))
            if start_time is None:
                parse_errors += 1
            elif start_date <= start_time <= end_date:
                filtered.append(record)
        
        print(f"✅ Filtered to {len(filtered)} records in date range")
        if parse_errors > 0:
            print(f"⚠️  Warning: {parse_errors} records had unparseable dates")
        return filtered
    
    def associate_distances_with_addresses(self, records: List[Dict]) -> List[Dict]:
        """Associate distances with addresses for IN_PASSENGER_VEHICLE activities"""
        print("\n🚗 Processing vehicle trips for CRA logbook...")
        
        sorted_records = sorted(records, key=lambda x: x.get('start_time', ''))
        all_vehicle_trips = []
        
        for i, record in enumerate(sorted_records):
            record_type = record.get('type', '')
            
            if record_type == 'Activity' and record.get('activity_type', '') == 'IN_PASSENGER_VEHICLE':
                from_address = ''
                to_address = ''
                
                # Look backward for the previous Visit (origin)
                for j in range(i - 1, -1, -1):
                    if sorted_records[j].get('type') == 'Visit':
                        from_address = sorted_records[j].get('address', '')
                        break
                
                # Look forward for the next Visit (destination)
                for j in range(i + 1, len(sorted_records)):
                    if sorted_records[j].get('type') == 'Visit':
                        to_address = sorted_records[j].get('address', '')
                        break
                
                # Extract date from start_time (format: YYYY-MM-DD HH:MM:SS)
                start_time_str = record.get('start_time', '')
                date = start_time_str.split(' ')[0] if ' ' in start_time_str else start_time_str
                
                # Get distance as float
                distance_km = 0
                try:
                    distance_km = float(record.get('distance_km', 0))
                except:
                    pass
                
                entry = {
                    'Date': date,
                    'Start Time': record.get('start_time', ''),
                    'End Time': record.get('end_time', ''),
                    'Duration': self._format_duration(record.get('duration_minutes', '')),
                    'Starting Point': from_address,
                    'Destination': to_address,
                    'Purpose of Trip': '',  # Blank for now
                    'Distance': distance_km,  # Store as number
                    'Activity_type': record.get('activity_type', '')
                }
                
                all_vehicle_trips.append(entry)
        
        print(f"✅ Found {len(all_vehicle_trips)} total vehicle trips")
        return all_vehicle_trips
    
    def calculate_odometer_readings(self, trips: List[Dict], end_date: str, end_odometer: float) -> List[Dict]:
        """
        Calculate odometer readings by working backwards from a known end point
        
        Args:
            trips: List of trip entries (must be sorted by date/time, oldest first)
            end_date: Date of the known odometer reading (YYYY-MM-DD format)
            end_odometer: Odometer reading on end_date
            
        Returns:
            List of trips with Start Odometer and End Odometer added
        """
        print(f"\n📊 Calculating odometer readings...")
        print(f"   Reference point: {end_date} = {end_odometer:,.0f} km")
        
        # Parse the end date
        from datetime import datetime
        end_dt = datetime.strptime(end_date, '%Y-%m-%d')
        
        # Separate trips into before and after the reference date
        trips_before = []
        trips_after = []
        
        for trip in trips:
            trip_date_str = trip.get('Date', '')
            if trip_date_str:
                try:
                    trip_dt = datetime.strptime(trip_date_str, '%Y-%m-%d')
                    if trip_dt <= end_dt:
                        trips_before.append(trip)
                    else:
                        trips_after.append(trip)
                except:
                    pass
        
        # Work backwards from the reference date for trips before/on that date
        current_odometer = end_odometer
        for trip in reversed(trips_before):
            distance = trip.get('Distance', 0)
            trip['End Odometer'] = current_odometer
            trip['Start Odometer'] = current_odometer - distance
            current_odometer = trip['Start Odometer']
        
        # Work forwards from the reference date for trips after that date
        current_odometer = end_odometer
        for trip in trips_after:
            distance = trip.get('Distance', 0)
            trip['Start Odometer'] = current_odometer
            trip['End Odometer'] = current_odometer + distance
            current_odometer = trip['End Odometer']
        
        print(f"✅ Calculated odometer readings for {len(trips)} trips")
        if trips_before:
            earliest_trip = trips_before[0]
            print(f"   Earliest trip: {earliest_trip.get('Date')} - Odometer: {earliest_trip.get('Start Odometer', 0):,.0f} km")
        
        return trips
    
    def write_final_report(self, report_data: List[Dict], sheet_name: str = "Final Report", 
                          end_date: str = "2025-10-01", end_odometer: float = 150000):
        """Write the final CRA-compliant logbook report to Google Sheets"""
        if not report_data:
            print("❌ No data to write!")
            return
        
        print(f"\n📤 Preparing CRA logbook report...")
        
        # Calculate odometer readings for ALL trips first
        report_data = self.calculate_odometer_readings(report_data, end_date, end_odometer)
        
        # Filter 1: Remove trips outside Ontario
        print(f"\n🔍 Filtering trips by location...")
        ontario_trips = []
        removed_non_ontario = []
        
        for trip in report_data:
            start_addr = trip.get('Starting Point', '').lower()
            dest_addr = trip.get('Destination', '').lower()
            
            # Check if both addresses contain "ON" or "Ontario"
            start_in_ontario = ', on ' in start_addr or ', ontario' in start_addr
            dest_in_ontario = ', on ' in dest_addr or ', ontario' in dest_addr
            
            if start_in_ontario and dest_in_ontario:
                ontario_trips.append(trip)
            else:
                removed_non_ontario.append(trip)
                if not start_in_ontario:
                    print(f"   ⚠️  Removed trip outside Ontario - Starting Point: {trip.get('Starting Point', '')[:60]}...")
                if not dest_in_ontario:
                    print(f"   ⚠️  Removed trip outside Ontario - Destination: {trip.get('Destination', '')[:60]}...")
        
        print(f"   ✅ {len(ontario_trips)} trips in Ontario")
        if removed_non_ontario:
            print(f"   ❌ Removed {len(removed_non_ontario)} trips outside Ontario")
        
        report_data = ontario_trips
        
        # Filter 2: Remove trips to/from excluded cities (example: City A)
        print(f"\n🔍 Filtering out excluded city trips...")
        non_excluded_trips = []
        removed_excluded = []
        
        # Example: Replace 'city_a' with your own cities to exclude
        excluded_cities = ['city_a']
        
        for trip in report_data:
            start_addr = trip.get('Starting Point', '').lower()
            dest_addr = trip.get('Destination', '').lower()
            
            # Check if either address contains excluded cities
            is_excluded = False
            for city in excluded_cities:
                if city in start_addr or city in dest_addr:
                    is_excluded = True
                    break
            
            if is_excluded:
                removed_excluded.append(trip)
                print(f"   ⚠️  Removed excluded city trip: {trip.get('Date')} {trip.get('Start Time', '').split()[1] if ' ' in trip.get('Start Time', '') else ''}")
                print(f"       {trip.get('Starting Point', '')[:50]}... → {trip.get('Destination', '')[:50]}...")
            else:
                non_excluded_trips.append(trip)
        
        if removed_excluded:
            print(f"   ❌ Removed {len(removed_excluded)} trips to/from excluded cities")
        else:
            print(f"   ✅ No excluded city trips found")
        
        report_data = non_excluded_trips
        
        # Filter 3: Remove trips to/from additional excluded cities (example: City B)
        print(f"\n🔍 Filtering out additional excluded city trips...")
        non_excluded2_trips = []
        removed_excluded2 = []
        
        # Example: Replace 'city_b' with your own cities to exclude
        excluded_cities_2 = ['city_b']
        
        for trip in report_data:
            start_addr = trip.get('Starting Point', '').lower()
            dest_addr = trip.get('Destination', '').lower()
            
            # Check if either address contains excluded cities
            is_excluded = False
            for city in excluded_cities_2:
                if city in start_addr or city in dest_addr:
                    is_excluded = True
                    break
            
            if is_excluded:
                removed_excluded2.append(trip)
                print(f"   ⚠️  Removed excluded city trip: {trip.get('Date')} {trip.get('Start Time', '').split()[1] if ' ' in trip.get('Start Time', '') else ''}")
                print(f"       {trip.get('Starting Point', '')[:50]}... → {trip.get('Destination', '')[:50]}...")
            else:
                non_excluded2_trips.append(trip)
        
        if removed_excluded2:
            print(f"   ❌ Removed {len(removed_excluded2)} trips to/from additional excluded cities")
        else:
            print(f"   ✅ No additional excluded city trips found")
        
        report_data = non_excluded2_trips
        
        # Filter 4: Remove trips from specific date range (optional)
        print(f"\n🔍 Filtering out specific date range trips...")
        date_filtered_trips = []
        removed_dates = []
        
        from datetime import datetime
        # Example: Replace with your own date range to exclude
        exclude_start = datetime.strptime('2025-01-01', '%Y-%m-%d')
        exclude_end = datetime.strptime('2025-01-05', '%Y-%m-%d')
        
        for trip in report_data:
            trip_date_str = trip.get('Date', '')
            if trip_date_str:
                try:
                    trip_date = datetime.strptime(trip_date_str, '%Y-%m-%d')
                    if exclude_start <= trip_date <= exclude_end:
                        removed_dates.append(trip)
                        print(f"   ⚠️  Removed date range trip: {trip.get('Date')} {trip.get('Start Time', '').split()[1] if ' ' in trip.get('Start Time', '') else ''}")
                        print(f"       {trip.get('Starting Point', '')[:50]}... → {trip.get('Destination', '')[:50]}...")
                    else:
                        date_filtered_trips.append(trip)
                except:
                    date_filtered_trips.append(trip)
            else:
                date_filtered_trips.append(trip)
        
        if removed_dates:
            print(f"   ❌ Removed {len(removed_dates)} trips from excluded date range")
        else:
            print(f"   ✅ No trips found in excluded date range")
        
        report_data = date_filtered_trips
        
        # Filter 5: Remove impossible consecutive trips (same starting point twice in a row)
        print(f"\n🔍 Checking for impossible consecutive trips...")
        filtered_impossible = []
        removed_impossible = []
        
        for i, trip in enumerate(report_data):
            if i == 0:
                filtered_impossible.append(trip)
                continue
            
            prev_trip = filtered_impossible[-1]
            prev_start = prev_trip.get('Starting Point', '').strip().lower()
            current_start = trip.get('Starting Point', '').strip().lower()
            
            if prev_start == current_start and prev_start != '':
                # Two consecutive trips from the same location - impossible!
                removed_impossible.append(trip)
                print(f"   ⚠️  Removed impossible trip: {trip.get('Date')} {trip.get('Start Time', '').split()[1] if ' ' in trip.get('Start Time', '') else ''}")
                print(f"       {trip.get('Starting Point', '')[:50]}... → {trip.get('Destination', '')[:50]}...")
            else:
                filtered_impossible.append(trip)
        
        if removed_impossible:
            print(f"   ❌ Removed {len(removed_impossible)} impossible consecutive trips")
        else:
            print(f"   ✅ No impossible consecutive trips found")
        
        report_data = date_filtered_trips
        
        # Filter 5: Remove impossible consecutive trips (same starting point twice in a row)
        print(f"\n🔍 Checking for impossible consecutive trips...")
        filtered_impossible = []
        removed_impossible = []
        
        for i, trip in enumerate(report_data):
            if i == 0:
                filtered_impossible.append(trip)
                continue
            
            prev_trip = filtered_impossible[-1]
            prev_start = prev_trip.get('Starting Point', '').strip().lower()
            current_start = trip.get('Starting Point', '').strip().lower()
            
            if prev_start == current_start and prev_start != '':
                # Two consecutive trips from the same location - impossible!
                removed_impossible.append(trip)
                print(f"   ⚠️  Removed impossible trip: {trip.get('Date')} {trip.get('Start Time', '').split()[1] if ' ' in trip.get('Start Time', '') else ''}")
                print(f"       {trip.get('Starting Point', '')[:50]}... → {trip.get('Destination', '')[:50]}...")
            else:
                filtered_impossible.append(trip)
        
        if removed_impossible:
            print(f"   ❌ Removed {len(removed_impossible)} impossible consecutive trips")
        else:
            print(f"   ✅ No impossible consecutive trips found")
        
        report_data = filtered_impossible
        
        # Add Purpose of Trip based on destination/starting point
        print(f"\n📝 Assigning trip purposes...")
        
        # Example cities for "Travel to Customer Site"
        # Replace these with your actual customer site cities
        customer_site_cities = [
            'Tiverton', 'Port Elgin', 'Kincardine', 'city4', 'city5',
            'city6', 'city7', 'city8', 'city9', 'city10'
        ]
        
        # Example cities for "Meeting with Customers"
        # Replace these with your actual meeting cities
        meeting_cities = [
            'city11', 'city12', 'city13',
            'city14', 'city15', 'city16', 'city17'
        ]
        
        # Example cities for "Delivery"
        # Replace these with your actual delivery cities
        delivery_cities = ['city18']
        
        for trip in report_data:
            start_addr = trip.get('Starting Point', '').lower()
            dest_addr = trip.get('Destination', '').lower()
            
            purpose = ''
            
            # Check for customer site cities
            for city in customer_site_cities:
                if city in dest_addr or city in start_addr:
                    purpose = 'Travel to Customer Site'
                    break
            
            # Check for meeting cities
            if not purpose:
                for city in meeting_cities:
                    if city in dest_addr or city in start_addr:
                        purpose = 'Meeting with Customers'
                        break
            
            # Check for delivery cities
            if not purpose:
                for city in delivery_cities:
                    if city in dest_addr or city in start_addr:
                        purpose = 'Delivery'
                        break
            
            trip['Purpose of Trip'] = purpose
        
        print(f"   ✅ Trip purposes assigned")
        
        # Filter 6: Remove trips under 15km for the final report
        filtered_data = []
        for entry in report_data:
            distance = entry.get('Distance', 0)
            if distance >= 15:
                filtered_data.append(entry)
        
        print(f"\n📊 Final filtering results:")
        print(f"   • {len(filtered_data)} trips >= 15km will be shown")
        print(f"   • {len(report_data) - len(filtered_data)} trips < 15km excluded")
        
        # CRA-compliant headers
        headers = [
            'Item', 
            'Date', 
            'Start Time', 
            'End Time', 
            'Starting Point', 
            'Destination', 
            'Purpose of Trip',
            'Km Driven',
            'Start Odometer',
            'End Odometer',
            'Duration (min)',
            'Activity_type',
            ''  # Empty column for total
        ]
        
        # Calculate total distance (only filtered trips >= 15km)
        total_distance = sum(entry.get('Distance', 0) for entry in filtered_data)
        
        # First row with total distance in M1 (column 13)
        rows = [['', '', '', '', '', '', '', '', '', '', '', '', f'{total_distance:.2f}']]
        
        # Header row
        rows.append(headers)
        
        # Data rows with item numbers
        for i, entry in enumerate(filtered_data, start=1):
            # Extract time from datetime
            start_time = entry.get('Start Time', '')
            end_time = entry.get('End Time', '')
            start_time_only = start_time.split(' ')[1] if ' ' in start_time else start_time
            end_time_only = end_time.split(' ')[1] if ' ' in end_time else end_time
            
            # Convert duration to minutes only
            duration_str = entry.get('Duration', '')
            duration_minutes = self._extract_total_minutes(duration_str)
            
            row = [
                str(i),  # Item number
                entry.get('Date', ''),
                start_time_only,
                end_time_only,
                entry.get('Starting Point', ''),
                entry.get('Destination', ''),
                entry.get('Purpose of Trip', ''),
                f"{entry.get('Distance', 0):.2f}",  # Km Driven
                f"{entry.get('Start Odometer', 0):.0f}",  # Start Odometer (no decimals)
                f"{entry.get('End Odometer', 0):.0f}",  # End Odometer (no decimals)
                str(duration_minutes),
                entry.get('Activity_type', ''),
                ''  # Empty column M
            ]
            rows.append(row)
        
        # Add empty row
        rows.append(['', '', '', '', '', '', '', '', '', '', '', '', ''])
        
        # Add total row (in column M)
        rows.append(['', '', '', '', '', '', '', 'Total Distance Traveled in Fiscal Year', '', '', '', '', f'{total_distance:.2f}'])
        
        body = {'values': rows}
        
        try:
            # Clear existing data
            self.sheets_service.spreadsheets().values().clear(
                spreadsheetId=self.spreadsheet_id,
                range=f"'{sheet_name}'!A:Z"
            ).execute()
            
            # Write new data
            result = self.sheets_service.spreadsheets().values().update(
                spreadsheetId=self.spreadsheet_id,
                range=f"'{sheet_name}'!A1",
                valueInputOption='USER_ENTERED',
                body=body
            ).execute()
            
            # Apply formatting
            requests = []
            
            # Make header row (row 2) bold and underlined
            requests.append({
                'repeatCell': {
                    'range': {
                        'sheetId': self._get_sheet_id(sheet_name),
                        'startRowIndex': 1,
                        'endRowIndex': 2,
                        'startColumnIndex': 0,
                        'endColumnIndex': 13
                    },
                    'cell': {
                        'userEnteredFormat': {
                            'textFormat': {
                                'bold': True,
                                'underline': True
                            }
                        }
                    },
                    'fields': 'userEnteredFormat.textFormat'
                }
            })
            
            # Center align Item column (column A, index 0)
            requests.append({
                'repeatCell': {
                    'range': {
                        'sheetId': self._get_sheet_id(sheet_name),
                        'startColumnIndex': 0,
                        'endColumnIndex': 1,
                        'startRowIndex': 2,
                        'endRowIndex': len(rows)
                    },
                    'cell': {
                        'userEnteredFormat': {
                            'horizontalAlignment': 'CENTER'
                        }
                    },
                    'fields': 'userEnteredFormat.horizontalAlignment'
                }
            })
            
            # Center align Duration column (column K, index 10)
            requests.append({
                'repeatCell': {
                    'range': {
                        'sheetId': self._get_sheet_id(sheet_name),
                        'startColumnIndex': 10,
                        'endColumnIndex': 11,
                        'startRowIndex': 2,
                        'endRowIndex': len(rows)
                    },
                    'cell': {
                        'userEnteredFormat': {
                            'horizontalAlignment': 'CENTER'
                        }
                    },
                    'fields': 'userEnteredFormat.horizontalAlignment'
                }
            })
            
            # Center align and format Km Driven column (column H, index 7)
            requests.append({
                'repeatCell': {
                    'range': {
                        'sheetId': self._get_sheet_id(sheet_name),
                        'startColumnIndex': 7,
                        'endColumnIndex': 8,
                        'startRowIndex': 2,
                        'endRowIndex': len(rows)
                    },
                    'cell': {
                        'userEnteredFormat': {
                            'horizontalAlignment': 'CENTER',
                            'numberFormat': {
                                'type': 'NUMBER',
                                'pattern': '#,##0.00'
                            }
                        }
                    },
                    'fields': 'userEnteredFormat.horizontalAlignment,userEnteredFormat.numberFormat'
                }
            })
            
            # Center align Start Odometer (column I, index 8)
            requests.append({
                'repeatCell': {
                    'range': {
                        'sheetId': self._get_sheet_id(sheet_name),
                        'startColumnIndex': 8,
                        'endColumnIndex': 9,
                        'startRowIndex': 2,
                        'endRowIndex': len(rows)
                    },
                    'cell': {
                        'userEnteredFormat': {
                            'horizontalAlignment': 'CENTER',
                            'numberFormat': {
                                'type': 'NUMBER',
                                'pattern': '#,##0'
                            }
                        }
                    },
                    'fields': 'userEnteredFormat.horizontalAlignment,userEnteredFormat.numberFormat'
                }
            })
            
            # Center align End Odometer (column J, index 9)
            requests.append({
                'repeatCell': {
                    'range': {
                        'sheetId': self._get_sheet_id(sheet_name),
                        'startColumnIndex': 9,
                        'endColumnIndex': 10,
                        'startRowIndex': 2,
                        'endRowIndex': len(rows)
                    },
                    'cell': {
                        'userEnteredFormat': {
                            'horizontalAlignment': 'CENTER',
                            'numberFormat': {
                                'type': 'NUMBER',
                                'pattern': '#,##0'
                            }
                        }
                    },
                    'fields': 'userEnteredFormat.horizontalAlignment,userEnteredFormat.numberFormat'
                }
            })
            
            # Format M1 (total distance) as number
            requests.append({
                'repeatCell': {
                    'range': {
                        'sheetId': self._get_sheet_id(sheet_name),
                        'startColumnIndex': 12,
                        'endColumnIndex': 13,
                        'startRowIndex': 0,
                        'endRowIndex': 1
                    },
                    'cell': {
                        'userEnteredFormat': {
                            'numberFormat': {
                                'type': 'NUMBER',
                                'pattern': '#,##0.00'
                            }
                        }
                    },
                    'fields': 'userEnteredFormat.numberFormat'
                }
            })
            
            # Format last row total distance (column M) as number
            requests.append({
                'repeatCell': {
                    'range': {
                        'sheetId': self._get_sheet_id(sheet_name),
                        'startColumnIndex': 12,
                        'endColumnIndex': 13,
                        'startRowIndex': len(rows) - 1,
                        'endRowIndex': len(rows)
                    },
                    'cell': {
                        'userEnteredFormat': {
                            'numberFormat': {
                                'type': 'NUMBER',
                                'pattern': '#,##0.00'
                            }
                        }
                    },
                    'fields': 'userEnteredFormat.numberFormat'
                }
            })
            
            # Apply formatting
            self.sheets_service.spreadsheets().batchUpdate(
                spreadsheetId=self.spreadsheet_id,
                body={'requests': requests}
            ).execute()
            
            updated_rows = result.get('updatedRows', 0)
            print(f"\n✅ Successfully wrote {updated_rows} rows to '{sheet_name}' tab!")
            print(f"📊 Total Distance (trips >= 15km): {total_distance:.2f} km")
            print(f"🚗 CRA-compliant logbook ready!")
            
        except HttpError as error:
            print(f"❌ Error writing to sheet: {error}")
    
    def _format_duration(self, duration_str: str) -> str:
        """Format duration from minutes to readable format"""
        if not duration_str or duration_str == '':
            return ''
        
        try:
            minutes = float(duration_str)
            
            if minutes < 60:
                return f"{int(minutes)} min"
            else:
                hours = int(minutes // 60)
                mins = int(minutes % 60)
                return f"{hours}h {mins}min"
        except:
            return duration_str
    
    def _format_distance(self, distance_str: str) -> str:
        """Format distance with km unit"""
        if not distance_str or distance_str == '':
            return '0 km'
        
        try:
            distance = float(distance_str)
            return f"{distance:.2f} km"
        except:
            return distance_str
    
    
    def _extract_total_minutes(self, duration_str: str) -> int:
        """Extract total minutes from duration string like '1h 30min' or '45min'"""
        if not duration_str:
            return 0
        
        try:
            # Handle formats like "1h 30min" or "45min" or "1h 30 min"
            duration_str = duration_str.lower().strip()
            total_minutes = 0
            
            if 'h' in duration_str:
                parts = duration_str.split('h')
                hours = int(parts[0].strip())
                total_minutes = hours * 60
                
                # Check if there are remaining minutes
                if len(parts) > 1:
                    min_part = parts[1].replace('min', '').strip()
                    if min_part:
                        total_minutes += int(min_part)
            else:
                # Just minutes
                min_part = duration_str.replace('min', '').strip()
                if min_part:
                    total_minutes = int(min_part)
            
            return total_minutes
        except:
            return 0
    
    def _get_sheet_id(self, sheet_name: str) -> int:
        """Get the sheet ID for a given sheet name"""
        try:
            spreadsheet = self.sheets_service.spreadsheets().get(
                spreadsheetId=self.spreadsheet_id
            ).execute()
            
            for sheet in spreadsheet.get('sheets', []):
                if sheet.get('properties', {}).get('title') == sheet_name:
                    return sheet.get('properties', {}).get('sheetId')
            
            return 0  # Default to first sheet
        except:
            return 0
    
    # ========================================================================
    # MAIN PIPELINE ORCHESTRATION
    # ========================================================================
    
    def run_complete_pipeline(self, 
                             json_path: str = None,
                             resolve_activities: bool = False,
                             start_date: str = "2024-09-30",
                             end_date: str = "2025-10-01",
                             skip_json_import: bool = False):
        """
        Run the complete pipeline
        
        Args:
            json_path: Path to JSON file (required if skip_json_import=False)
            resolve_activities: Whether to reverse geocode activity locations
            start_date: Start date for fiscal year filter
            end_date: End date for fiscal year filter
            skip_json_import: If True, skip JSON import and go straight to processing
        """
        print("=" * 70)
        print("🗺️  COMPLETE GOOGLE TIMELINE PIPELINE")
        print("=" * 70)
        
        # STEP 1: Import JSON to Timeline Data tab (if not skipping)
        if not skip_json_import:
            if not json_path:
                print("❌ Error: json_path is required when skip_json_import=False")
                return
            
            print("\n" + "=" * 70)
            print("STEP 1: Import JSON to Timeline Data")
            print("=" * 70)
            
            records = self.parse_timeline_json(json_path)
            if not records:
                return
            
            records = self.resolve_addresses(records, resolve_activities=resolve_activities)
            self.write_to_sheet(records, "Timeline Data")
        else:
            print("\n⏭️  Skipping JSON import (using existing Timeline Data)")
        
        # STEP 2: Process Timeline Data to Final Report
        print("\n" + "=" * 70)
        print("STEP 2: Process to Final Report")
        print("=" * 70)
        
        timeline_records = self.read_timeline_data("Timeline Data")
        
        if not timeline_records:
            return
        
        filtered_records = self.filter_by_date_range(timeline_records, start_date, end_date)
        final_report = self.associate_distances_with_addresses(filtered_records)
        
        # Odometer settings - change these if needed
        ODOMETER_END_DATE = "2025-10-01"  # Date of known odometer reading
        ODOMETER_END_READING = 60000      # Odometer reading on that date (example: 150,000 km)
        
        self.write_final_report(final_report, "Final Report", ODOMETER_END_DATE, ODOMETER_END_READING)
        
        # Summary
        print("\n" + "=" * 70)
        print("✅ PIPELINE COMPLETE!")
        print("=" * 70)
        print(f"📊 Total vehicle trips in Final Report: {len(final_report)}")
        print(f"📅 Date range: {start_date} to {end_date}")
        print("=" * 70)


def main():
    """Run the complete pipeline"""
    
    # ========== CONFIGURATION ==========
    
    # Google API Credentials
    GOOGLE_MAPS_API_KEY = "YOUR-MAPS-API-KEY"
    SHEETS_CREDENTIALS_PATH = "credentials.json"
    SPREADSHEET_ID = "YOUR-SPREADSHEET-ID"
    SERVICE_ACCOUNT_EMAIL = "your-service-account@email"
    
    # JSON File Settings
    TIMELINE_JSON_PATH = "Timeline2024-2025.json"  # Your JSON file
    SKIP_JSON_IMPORT = False  # Set to True if you already imported JSON and just want to process
    
    # Processing Settings
    RESOLVE_ACTIVITIES = False  # Set to True to reverse geocode activity locations (slower, more API calls)
    START_DATE = "2024-10-01"  # Fiscal year start
    END_DATE = "2025-10-01"    # Fiscal year end
    
    # ===================================
    
    print("\n⚠️  IMPORTANT: Make sure you've shared the Google Sheet with:")
    print(f"   📧 {SERVICE_ACCOUNT_EMAIL}")
    print("   (Give it 'Editor' permissions)\n")
    
    try:
        # Create pipeline
        pipeline = CompleteTimelinePipeline(
            GOOGLE_MAPS_API_KEY, 
            SHEETS_CREDENTIALS_PATH,
            SPREADSHEET_ID
        )
        
        # Run complete pipeline
        pipeline.run_complete_pipeline(
            json_path=TIMELINE_JSON_PATH,
            resolve_activities=RESOLVE_ACTIVITIES,
            start_date=START_DATE,
            end_date=END_DATE,
            skip_json_import=SKIP_JSON_IMPORT
        )
        
    except Exception as e:
        print(f"\n❌ Error: {e}")
        if "403" in str(e) or "Forbidden" in str(e):
            print("\n⚠️  Permission Error! Please share your Google Sheet with:")
            print(f"   📧 {SERVICE_ACCOUNT_EMAIL}")
            print(f"   1. Open: https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit")
            print(f"   2. Click 'Share' button")
            print(f"   3. Add the service account email above")
            print(f"   4. Grant 'Editor' access")
            print(f"   5. Run this script again")


if __name__ == "__main__":
    main()
