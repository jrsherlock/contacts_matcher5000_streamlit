import pandas as pd
import re
import json
import os
from fuzzywuzzy import process, fuzz
from typing import Dict, List, Optional, Tuple
from tqdm import tqdm
import csv

SETTINGS_FILE = "matcher_settings.json"

def try_read_csv(file_path):
    """Try to read a CSV file with different encodings and delimiters"""
    encodings = ['utf-8', 'latin1', 'iso-8859-1', 'cp1252', 'macroman']
    delimiters = [',', ';', '\t']
    
    for encoding in encodings:
        for delimiter in delimiters:
            try:
                df = pd.read_csv(file_path, encoding=encoding, delimiter=delimiter, dtype=str)
                
                # Clean up column names
                df.columns = [col.strip() for col in df.columns]
                
                # Drop unnamed columns
                df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
                
                return df
            except Exception as e:
                continue
    
    print(f"ERROR: Failed to read {file_path} with all encodings and delimiters")
    return None

def print_box(title: str, content: List[str], width: int = 60):
    """Print content in a nice box."""
    # Clean and pad content lines
    clean_content = []
    for line in content:
        # Split long lines
        while len(line) > width - 4:
            split_point = line[:width-4].rfind(' ')
            if split_point == -1:
                split_point = width - 4
            clean_content.append(line[:split_point])
            line = line[split_point:].strip()
        clean_content.append(line)

    print("╔" + "═" * (width - 2) + "╗")
    print(f"║ {title.center(width-4)} ║")
    print("╠" + "═" * (width - 2) + "╣")
    for line in clean_content:
        print(f"║ {line.ljust(width-4)} ║")
    print("╚" + "═" * (width - 2) + "╝")
    print()  # Add blank line after box

def save_settings(settings: Dict, filename: str = SETTINGS_FILE):
    """Save all settings to a JSON file."""
    with open(filename, 'w') as f:
        json.dump(settings, f, indent=2)

def load_settings(filename: str = SETTINGS_FILE) -> Optional[Dict]:
    """Load settings from JSON file if it exists."""
    try:
        with open(filename) as f:
            return json.load(f)
    except FileNotFoundError:
        return None

def display_current_settings(settings):
    """Display current settings in a formatted box"""
    print("\n╔" + "═" * 52 + "╗")
    print("║" + " " * 21 + "Current Settings" + " " * 18 + "║")
    print("╠" + "═" * 52 + "╣")
    
    # Files
    print("║ Input File:", settings.get('input_file', 'Not set').ljust(41) + "║")
    print("║ Target File:", settings.get('target_file', 'Not set').ljust(40) + "║")
    print("║" + " " * 52 + "║")
    
    # Thresholds
    print("║ Matching Thresholds:" + " " * 34 + "║")
    thresholds = settings.get('thresholds', {
        'company_name': 85,
        'person_name': 85,
        'email': 100,
        'title': 70,
        'department': 70
    })
    for key, value in thresholds.items():
        name = key.replace('_', ' ').title()
        print(f"║   {name}: {value}%".ljust(52) + "║")
    print("║" + " " * 52 + "║")
    
    # Column Mappings
    print("║ Column Mappings:" + " " * 37 + "║")
    mappings = settings.get('column_mapping', {})
    for source, target in mappings.items():
        mapping = f"  {source} → {target}"
        print("║" + mapping.ljust(52) + "║")
    
    print("╚" + "═" * 52 + "╝")

def get_csv_files():
    """Get list of CSV files in current directory"""
    files = []
    for file in os.listdir('.'):
        if file.lower().endswith('.csv'):
            files.append(file)
    return sorted(files)

def select_file(prompt="Select a file"):
    """Let user select a single file from available CSVs"""
    csv_files = get_csv_files()
    if not csv_files:
        print("No CSV files found in current directory")
        return None
    
    print(f"\n{prompt}:")
    for i, file in enumerate(csv_files, 1):
        print(f"{i}. {file}")
    
    while True:
        try:
            choice = input("\nEnter file number: ").strip()
            idx = int(choice) - 1
            if 0 <= idx < len(csv_files):
                return csv_files[idx]
            print("Invalid selection")
        except ValueError:
            print("Please enter a valid number")

def select_multiple_files(prompt="Select files"):
    """Let user select multiple files from available CSVs"""
    csv_files = get_csv_files()
    if not csv_files:
        print("No CSV files found in current directory")
        return []
    
    print(f"\n{prompt} (enter multiple numbers separated by spaces):")
    for i, file in enumerate(csv_files, 1):
        print(f"{i}. {file}")
    
    while True:
        try:
            choices = input("\nEnter file numbers: ").strip().split()
            selections = []
            for choice in choices:
                idx = int(choice) - 1
                if 0 <= idx < len(csv_files):
                    selections.append(csv_files[idx])
                else:
                    print(f"Invalid selection: {choice}")
                    selections = []
                    break
            if selections:
                return selections
        except ValueError:
            print("Please enter valid numbers")

def detect_company_column(columns: List[str]) -> str:
    """Try to automatically detect the company column name."""
    company_keywords = ['company name', 'company', 'organization', 'organisation', 'employer', 'business', 'firm']
    # First try exact matches
    for col in columns:
        if str(col).lower() in ['company name', 'company']:
            return col
    # Then try partial matches
    for col in columns:
        if any(keyword in str(col).lower() for keyword in company_keywords):
            return col
    return None

def get_column_mapping(input_cols: List[str], target_cols: List[str], batch_size: int = 10) -> Dict[str, str]:
    """Let user map input columns to target columns with improved interface."""
    print("\nColumn Mapping")
    print("For each input column, select the corresponding target column.")
    print("Enter '0' to skip a column or 'q' to finish mapping")
    
    mapping = {}
    for i in range(0, len(input_cols), batch_size):
        batch = input_cols[i:i+batch_size]
        print(f"\nMapping columns {i+1}-{min(i+batch_size, len(input_cols))} of {len(input_cols)}")
        
        # Show target columns once for the batch
        print("\nAvailable target columns:")
        for j, col in enumerate(target_cols, 1):
            print(f"{j}. {col}")
        print("0. Skip column")
        print("q. Finish mapping")
        
        for input_col in batch:
            while True:
                try:
                    choice = input(f"\nSelect target column for '{input_col}' (0-{len(target_cols)} or q): ").strip()
                    if choice.lower() == 'q':
                        if mapping:  # Only return if we have at least one mapping
                            return mapping
                        else:
                            print("Error: Must map at least one column")
                            continue
                    choice = int(choice)
                    if choice == 0:
                        break
                    if 1 <= choice <= len(target_cols):
                        mapping[input_col] = target_cols[choice - 1]
                        break
                except ValueError:
                    pass
                print("Invalid choice. Please try again.")
    return mapping

def save_column_mapping(mapping: Dict[str, str], filename: str = "column_mapping.json"):
    """Save column mapping to JSON file."""
    with open(filename, 'w') as f:
        json.dump(mapping, f, indent=2)

def load_column_mapping(filename: str = "column_mapping.json") -> Optional[Dict[str, str]]:
    """Load column mapping from JSON file if it exists."""
    try:
        with open(filename) as f:
            return json.load(f)
    except FileNotFoundError:
        return None

def safe_get_column(row, column_name, column_mapping=None, default=''):
    """Safely get a column value from a row, using column mapping if provided."""
    try:
        if column_mapping and column_name in column_mapping:
            mapped_name = column_mapping[column_name]
            value = row.get(mapped_name, default)
        else:
            value = row.get(column_name, default)
        return str(value).lower() if value is not None else default
    except Exception:
        return default

def print_overlap_stats(name, overlaps, target_companies):
    """Print statistics about overlaps for debugging."""
    print(f"\n{name} Stats:")
    print(f"Total overlapping companies: {len(overlaps)}")
    print(f"Total target companies: {len(target_companies)}")
    if len(target_companies) > 0:
        print(f"Overlap percentage: {(len(overlaps) / len(target_companies) * 100):.1f}%")
    else:
        print("Overlap percentage: 0% (no target companies)")
    if len(overlaps) <= 10:
        print(f"Overlapping companies: {sorted(overlaps)}")
    else:
        print(f"First 10 overlapping companies: {sorted(overlaps)[:10]}")

def normalize_company_name(name):
    """Normalize company names for better matching"""
    if pd.isna(name):
        return ''
    
    # Convert to lowercase and remove punctuation except & and .
    name = str(name).lower().strip()
    name = re.sub(r'[^\w\s\&\.]', '', name)
    
    # Keep & symbol but standardize spacing around it
    name = re.sub(r'\s*\&\s*', ' & ', name)
    
    # Remove leading "the"
    name = re.sub(r'^the\s+', '', name)
    
    # Handle special cases
    if 'qcr' in name.lower():
        # Special handling for QCR Holdings variations
        name = re.sub(r'qcr\s*holdings?\s*(?:inc|incorporated)?', 'qcr holdings', name)
        return name.strip()
    
    # Handle educational institutions
    edu_keywords = ['university', 'college', 'institute', 'school']
    is_edu = any(keyword in name for keyword in edu_keywords)
    
    # Remove common company suffixes and prefixes
    suffixes = [
        'inc', 'corp', 'corporation', 'llc', 'ltd', 'limited', 'co',
        'company', 'group', 'holdings', 'international', 'intl',
        'worldwide', 'global', 'solutions', 'services', 'technologies',
        'technology', 'tech', 'plc', 'lp', 'llp', 'gmbh', 'sa', 'ag',
        'nv', 'bv', 'pty', 'proprietary'
    ]
    
    # Handle department/division indicators
    divisions = ['division of', 'subsidiary of', 'part of', 'a division of', 'a subsidiary of']
    for div in divisions:
        name = re.sub(rf'\s*{div}\s+', ' ', name)
    
    # Standardize common abbreviations
    abbrev_map = {
        'corp': 'corporation',
        'inc': 'incorporated',
        'intl': 'international',
        'tech': 'technology',
        'mfg': 'manufacturing',
        'svcs': 'services',
        'sys': 'systems',
        'grp': 'group',
        'hldg': 'holding',
        'univ': 'university',
        'hosp': 'hospital',
        'med': 'medical',
        'ctr': 'center'
    }
    
    # First expand abbreviations
    words = name.split()
    for i, word in enumerate(words):
        if word in abbrev_map:
            words[i] = abbrev_map[word]
    name = ' '.join(words)
    
    # Only remove suffixes if not an educational institution
    if not is_edu:
        for suffix in sorted(suffixes, key=len, reverse=True):
            pattern = rf'\s+{suffix}(?:\s+|$)'
            name = re.sub(pattern, ' ', name)
    
    # Standardize whitespace and dots
    name = re.sub(r'\s+', ' ', name)
    name = re.sub(r'\.+', '.', name)
    name = name.strip(' .')
    
    return name

def normalize_person_name(name):
    """Normalize person names for better matching"""
    if pd.isna(name):
        return ''
        
    # Convert to lowercase and remove punctuation (except . for initials)
    name = re.sub(r'[^\w\s\.]', '', str(name).lower()).strip()
    
    # Remove common titles and suffixes
    titles = [
        'dr', 'mr', 'mrs', 'ms', 'miss', 'prof', 'professor',
        'phd', 'md', 'mba', 'esq', 'cpa', 'jr', 'sr', 'ii', 'iii', 'iv'
    ]
    for title in titles:
        name = re.sub(rf'^{title}\s+', '', name)
        name = re.sub(rf'\s+{title}$', '', name)
    
    # Handle initials (ensure consistent spacing)
    name = re.sub(r'([A-Za-z])\.([A-Za-z])', r'\1. \2', name)
    
    # Remove multiple spaces
    name = re.sub(r'\s+', ' ', name)
    return name.strip()

def normalize_job_title(title):
    """Normalize job titles for better matching"""
    if pd.isna(title):
        return ''
    
    title = str(title).lower().strip()
    
    # Common role level mappings
    level_map = {
        r'\bsr\b': 'senior',
        r'\bjr\b': 'junior',
        r'\bsvp\b': 'senior vice president',
        r'\bvp\b': 'vice president',
        r'\bceo\b': 'chief executive officer',
        r'\bcfo\b': 'chief financial officer',
        r'\bcto\b': 'chief technology officer',
        r'\bcoo\b': 'chief operating officer',
        r'\bcio\b': 'chief information officer',
        r'\bcmo\b': 'chief marketing officer',
        r'\bpres\b': 'president',
        r'\bexec\b': 'executive',
        r'\bmgr\b': 'manager',
        r'\bdir\b': 'director',
        r'\beng\b': 'engineer',
        r'\bdev\b': 'developer'
    }
    
    # Apply mappings
    for abbr, full in level_map.items():
        title = re.sub(abbr, full, title)
    
    # Remove common filler words
    fillers = ['of', 'the', 'and', '&', 'for', 'to', 'in', 'at']
    for filler in fillers:
        title = re.sub(rf'\b{filler}\b', ' ', title)
    
    # Remove multiple spaces
    title = re.sub(r'\s+', ' ', title)
    return title.strip()

def get_person_key(row, column_mapping=None):
    """Generate a key for person matching using multiple fields"""
    # Get raw values first
    last_name = safe_get_column(row, 'Last Name', column_mapping)
    first_name = safe_get_column(row, 'First Name', column_mapping)
    email = safe_get_column(row, 'Email Address', column_mapping)
    
    # Handle "LastName, FirstName" format
    if first_name and ',' in first_name:
        parts = first_name.split(',')
        if len(parts) == 2:
            first_name = parts[1].strip()
            last_name = parts[0].strip()
            
    # Handle "First Middle Last" format in first name field
    elif first_name and ' ' in first_name:
        parts = first_name.split()
        if len(parts) >= 2:
            # If what we think is the last name contains the second part,
            # then the second part is likely a middle name
            if parts[1].lower() in last_name.lower():
                first_name = parts[0]
            else:
                # The second part might be part of the last name
                first_name = parts[0]
                last_name = ' '.join(parts[1:]) + ' ' + last_name
    
    # Extract email components if available
    email_name = ''
    if email and '@' in email:
        email_name = email.split('@')[0].lower()
        # Remove common email prefixes
        email_name = re.sub(r'^(info|contact|admin|support|sales)', '', email_name)
        # Remove numbers from email
        email_name = re.sub(r'\d+', '', email_name)
        # Remove special characters
        email_name = re.sub(r'[^\w\s]', '', email_name)
    
    # Normalize everything
    last_name = normalize_person_name(last_name)
    first_name = normalize_person_name(first_name)
    job_title = normalize_job_title(safe_get_column(row, 'Position', column_mapping))
    linkedin = safe_get_column(row, 'URL', column_mapping).lower().strip()
    
    # Create composite key for matching
    return {
        'last_name': last_name,
        'first_name': first_name,
        'job_title': job_title,
        'linkedin': linkedin,
        'email_name': email_name
    }

def find_person_matches(input_contacts, target_contacts, thresholds):
    """Find matches between people using multiple criteria"""
    matches = []
    
    for input_contact in tqdm(input_contacts, desc="Processing input contacts"):
        input_key = get_person_key(input_contact)
        input_company = normalize_company_name(input_contact.get('company', ''))
        
        best_match = None
        best_score = 0
        
        for target_contact in target_contacts:
            target_key = get_person_key(target_contact)
            target_company = normalize_company_name(target_contact.get('company', ''))
            
            # Skip if normalized companies don't match with fuzzy matching
            if not input_company or not target_company:
                continue
                
            company_score = max(
                fuzz.token_sort_ratio(input_company, target_company),
                fuzz.token_set_ratio(input_company, target_company),
                fuzz.partial_ratio(input_company, target_company)
            )
            
            if company_score < thresholds['company_name']:
                continue
                
            # Compare names using multiple methods
            name_score = max(
                fuzz.token_sort_ratio(input_key, target_key),
                fuzz.token_set_ratio(input_key, target_key),
                fuzz.partial_ratio(input_key, target_key)  
            )
            
            # Check nicknames if score is below threshold
            if name_score < thresholds['person_name']:
                input_first = input_contact.get('first_name', '').lower()
                target_first = target_contact.get('first_name', '').lower()
                
                # Check if either name has nicknames
                input_nicknames = ['will', 'bill', 'billy'] if input_first == 'william' else [input_first]
                target_nicknames = ['rob', 'bob', 'bobby'] if target_first == 'robert' else [target_first]
                
                # Compare all nickname combinations
                for i_nick in input_nicknames:
                    for t_nick in target_nicknames:
                        nick_key = f"{i_nick} {input_contact.get('last_name', '')}"
                        target_nick_key = f"{t_nick} {target_contact.get('last_name', '')}"
                        nick_score = max(
                            fuzz.token_sort_ratio(nick_key, target_nick_key),
                            fuzz.token_set_ratio(nick_key, target_nick_key),
                            fuzz.partial_ratio(nick_key, target_nick_key)  
                        )
                        name_score = max(name_score, nick_score)
            
            # Check email if available
            email_score = 0
            input_email = input_contact.get('email', '')
            target_email = target_contact.get('email', '')
            if input_email and target_email:
                email_score = fuzz.ratio(input_email.lower(), target_email.lower())
                
                # If emails match exactly, boost the overall score
                if email_score == 100:
                    name_score = max(name_score, thresholds['person_name'] + 10)
            
            # Update best match if this is better
            if name_score >= thresholds['person_name'] and name_score > best_score:
                best_score = name_score
                best_match = {
                    'input_contact': input_contact,
                    'target_contact': target_contact,
                    'score': name_score
                }
        
        if best_match:
            matches.append(best_match)
    
    print(f"\nFound {len(matches)} person matches.")
    return matches

def find_matches(source_df: pd.DataFrame, target_df: pd.DataFrame, thresholds: Dict[str, int]) -> List[Tuple[str, str, float]]:
    """Find matches between two dataframes using fuzzy string matching."""
    matches = []
    seen_matches = set()
    
    # Print column names for debugging
    print("\nSource columns:", source_df.columns.tolist())
    print("Target columns:", target_df.columns.tolist())
    
    # Detect company column names
    source_company_col = None
    target_company_col = None
    
    # Try common company column names
    company_cols = ['Company', 'Company Name', 'Organization', 'Employer', 'Business']
    for col in company_cols:
        if col in source_df.columns:
            source_company_col = col
        if col in target_df.columns:
            target_company_col = col
    
    if not source_company_col or not target_company_col:
        print("Error: Could not find company columns. Available columns:")
        print("Source:", source_df.columns.tolist())
        print("Target:", target_df.columns.tolist())
        return matches
    
    print(f"\nUsing columns: {source_company_col} (source) and {target_company_col} (target)")
    
    # Get unique company names from both dataframes
    source_companies = source_df[source_company_col].dropna().unique()
    target_companies = target_df[target_company_col].dropna().unique()
    
    print(f"\nFound {len(source_companies)} source companies and {len(target_companies)} target companies")
    
    print("\nProcessing company matches...")
    
    # Use tqdm for progress tracking
    with tqdm(total=len(source_companies), ncols=80, 
              bar_format='{l_bar}{bar}| {n_fmt}/{total_fmt}') as pbar:
        for source_company in source_companies:
            # Normalize source company name
            norm_source = normalize_company_name(source_company)
            if not norm_source:
                pbar.update(1)
                continue
            
            # Find best match in target companies
            best_match = process.extractOne(
                norm_source,
                [normalize_company_name(tc) for tc in target_companies],
                scorer=fuzz.token_sort_ratio
            )
            
            if best_match and best_match[1] >= thresholds['company_name']:
                match_idx = best_match[2]
                target_company = target_companies[match_idx]
                
                # Create a unique key for this match
                match_key = f"{norm_source}|||{normalize_company_name(target_company)}"
                
                if match_key not in seen_matches:
                    matches.append((source_company, target_company, best_match[1]))
                    seen_matches.add(match_key)
                    
                    # Print match details on a new line
                    print(f"\nMatch found: '{source_company}' -> '{target_company}' (score: {best_match[1]})")
                    print(f"Normalized: '{norm_source}' -> '{normalize_company_name(target_company)}'")
            
            pbar.update(1)
    
    return matches

def detect_company_column(columns):
    """Detect which column contains company names."""
    company_keywords = ['company', 'organization', 'employer', 'business', 'firm']
    
    # Convert column names to lowercase for case-insensitive matching
    columns = [str(col).lower() for col in columns]
    
    # First try exact matches
    for keyword in company_keywords:
        for col in columns:
            if keyword in col:
                return col
    
    # If no match found, return None
    return None

def write_contact_info(f, contact_data, prefix=""):
    """Helper function to write contact information consistently"""
    # Write name
    name = f"{contact_data.get('First Name', '')} {contact_data.get('Last Name', '')}".strip()
    f.write(f"{prefix}Name: {name}\n")
    
    # Write title if available
    title = contact_data.get('Job Title', '').strip()
    f.write(f"{prefix}Title: {title if title else 'Not Available'}\n")
    
    # Write email if available
    email = contact_data.get('Email Address', '').strip()
    f.write(f"{prefix}Email: {email if email else 'Not Available'}\n")
    
    # Write LinkedIn URL if available
    linkedin = contact_data.get('LinkedIn', '').strip()
    f.write(f"{prefix}LinkedIn: {linkedin if linkedin else 'Not Available'}\n")
    
    # Write company if available
    company = contact_data.get('Company', '').strip()
    f.write(f"{prefix}Company: {company if company else 'Not Available'}\n")
    
    # Write connected on date if available
    connected_on = contact_data.get('Connected On', '').strip()
    if connected_on:
        f.write(f"{prefix}Connected On: {connected_on}\n")
    
    f.write(f"{prefix}" + "-" * 50 + "\n")

def write_overlap_report(company_matches, input_file, target_file):
    """Write a focused overlap report for outreach purposes"""
    output_file = 'company_overlaps.txt'
    print(f"Writing results to {output_file}...")
    
    with open(output_file, 'w', encoding='utf-8') as f:
        # Write header
        f.write("Contact Matching Results\n")
        f.write("=" * 50 + "\n\n")
        f.write("These are the contacts that the Source person knows at the Target list of companies and people.\n\n")
        f.write(f"Source File: {input_file}\n")
        f.write(f"Target File: {target_file}\n\n")
        
        # Write statistics
        f.write("Summary\n")
        f.write("-" * 30 + "\n")
        f.write(f"Total contacts matched: {sum(len(m[2]) for m in company_matches)}\n")
        f.write(f"Companies with matches: {len(company_matches)}\n")
        f.write(f"Source: {input_file}\n")
        f.write(f"Target companies: {target_file}\n\n")
        
        # Write detailed matches
        f.write("Contacts by Target Company\n")
        f.write("=" * 50 + "\n\n")
        
        for match in company_matches:
            if match[2]:  # If there are contacts for this company
                f.write(f"\nCompany: {match[1]}\n")
                f.write(f"Number of contacts: {len(match[2])}\n")
                f.write("-" * 50 + "\n\n")
                
                for person in match[2]:
                    if isinstance(person, dict) and 'score' in person:
                        # This is a matched contact
                        f.write(f"Match Score: {person['score']:.1f}%\n")
                        f.write("\nYour Connection:\n")
                        write_contact_info(f, person['input_contact'], "  ")
                        f.write("\nProspect:\n")
                        write_contact_info(f, person['target_contact'], "  ")
                    else:
                        # This is an unmatched contact
                        write_contact_info(f, person, "  ")
                    f.write("-" * 50 + "\n\n")

    print(f"Matching complete! Results written to {output_file}")
    print(f"Total contacts matched: {sum(len(m[2]) for m in company_matches)}")
    print(f"Companies with matches: {len(company_matches)}")

def test_company_match(name1, name2):
    """Test if two company names would match using our normalization and scoring."""
    norm1 = normalize_company_name(name1)
    norm2 = normalize_company_name(name2)
    
    print(f"\nTesting company name matching:")
    print(f"Original 1: {name1}")
    print(f"Normalized 1: {norm1}")
    print(f"Original 2: {name2}")
    print(f"Normalized 2: {norm2}")
    
    score = fuzz.token_sort_ratio(norm1, norm2)
    print(f"Match score: {score}")
    
    if score >= 80:
        print("MATCH!")
    else:
        print("NO MATCH")
    
    return score

def find_matches(input_file, target_file, thresholds):
    """Find matches between input and target contacts using fuzzy string matching"""
    print("\nContact Matcher")
    print("=" * 50 + "\n")

    # Load files
    print("Reading files...")
    target_contacts = try_read_csv(target_file)
    input_contacts = try_read_csv(input_file)
    if target_contacts is None or input_contacts is None:
        print(f"Error: Could not read input or target files")
        return []

    # First find all company matches
    print("Finding company matches...")
    company_matches = {}  # Use dict to track unique normalized company names
    
    # Process each target company
    for _, target_row in tqdm(target_contacts.iterrows(), desc="Processing companies"):
        target_company = None
        for col in ['Company', 'Company Name', 'Company Division Name']:
            if col in target_contacts.columns and not pd.isna(target_row[col]):
                target_company = str(target_row[col])
                break
        
        if not target_company:
            continue
            
        target_norm = normalize_company_name(target_company)
        if not target_norm:
            continue
            
        # Find all contacts at companies that match this target company
        contact_dict = {}  # Use dict to track unique contacts
        for _, source_row in input_contacts.iterrows():
            source_company = None
            for col in ['Company', 'Company Name', 'Company Division Name']:
                if col in input_contacts.columns and not pd.isna(source_row[col]):
                    source_company = str(source_row[col])
                    break
                    
            if not source_company:
                continue
                
            source_norm = normalize_company_name(source_company)
            if not source_norm:
                continue
                
            # Compare normalized company names
            company_score = fuzz.token_sort_ratio(target_norm, source_norm)
            
            # If companies match, add all contacts and check for person matches
            if company_score >= thresholds['company_name']:
                contact_data = {}
                # Map source fields to standardized fields
                field_mapping = {
                    'First Name': 'First Name',
                    'Last Name': 'Last Name',
                    'Email Address': 'Email Address',
                    'Position': 'Job Title',
                    'Company': 'Company',
                    'URL': 'LinkedIn',
                    'Connected On': 'Connected On'
                }
                
                # Copy all available fields with proper handling of NaN values
                for source_field, target_field in field_mapping.items():
                    if source_field in source_row:
                        value = source_row[source_field]
                        if pd.isna(value):
                            contact_data[target_field] = ''
                        else:
                            contact_data[target_field] = str(value).strip()
                
                # Check for person match
                person_match = False
                
                # Check name match
                source_name = normalize_person_name(f"{source_row.get('First Name', '')} {source_row.get('Last Name', '')}")
                target_name = normalize_person_name(f"{target_row.get('First Name', '')} {target_row.get('Last Name', '')}")
                if source_name and target_name:
                    name_score = fuzz.token_sort_ratio(source_name, target_name)
                    if name_score >= thresholds['person_name']:
                        person_match = True
                
                # Check email match
                if not person_match and 'Email Address' in source_row and 'Email Address' in target_row:
                    source_email = str(source_row['Email Address']).lower().strip() if not pd.isna(source_row['Email Address']) else ''
                    target_email = str(target_row['Email Address']).lower().strip() if not pd.isna(target_row['Email Address']) else ''
                    if source_email and target_email and source_email == target_email:
                        person_match = True
                
                # Check title match
                if not person_match and 'Position' in source_row and 'Job Title' in target_row:
                    source_title = normalize_job_title(str(source_row['Position'])) if not pd.isna(source_row['Position']) else ''
                    target_title = normalize_job_title(str(target_row['Job Title'])) if not pd.isna(target_row['Job Title']) else ''
                    if source_title and target_title:
                        title_score = fuzz.token_sort_ratio(source_title, target_title)
                        if title_score >= thresholds['title']:
                            person_match = True
                
                # Add a flag to indicate if this contact has a personal match
                contact_data['has_person_match'] = person_match
                
                # Create a unique key for this contact
                contact_key = f"{contact_data.get('First Name', '')}-{contact_data.get('Last Name', '')}-{contact_data.get('Email Address', '')}"
                contact_dict[contact_key] = contact_data
        
        if contact_dict:
            # Store unique contacts for this normalized company name
            if target_norm in company_matches:
                # Get existing contacts and merge with new ones
                _, best_company_name, existing_contacts = company_matches[target_norm]
                # Convert existing contacts to dictionary for deduplication
                existing_dict = {}
                for contact in existing_contacts:
                    key = f"{contact.get('First Name', '')}-{contact.get('Last Name', '')}-{contact.get('Email Address', '')}"
                    existing_dict[key] = contact
                # Update with new contacts
                existing_dict.update(contact_dict)
                # Use the longer/more complete company name
                final_company_name = target_company if len(target_company) > len(best_company_name) else best_company_name
                company_matches[target_norm] = (target_norm, final_company_name, list(existing_dict.values()))
            else:
                company_matches[target_norm] = (target_norm, target_company, list(contact_dict.values()))
    
    # Convert the dictionary values back to a list and sort by company name
    return sorted(company_matches.values(), key=lambda x: x[1].lower())

def process_company_names(df):
    """Extract and process company names from DataFrame"""
    company_cols = ['Company', 'Company Name', 'Company Division Name']
    for col in company_cols:
        if col in df.columns:
            return df[col].dropna().unique()
    return []

def find_company_matches(normalized_name, input_contacts, thresholds):
    """Find contacts that match a given company name"""
    matches = []
    company_cols = ['Company', 'Company Name', 'Company Division Name']
    
    for _, contact in tqdm(input_contacts.iterrows(), desc="Matching contacts", leave=False):
        # Find company column
        contact_company = None
        for col in company_cols:
            if col in contact and not pd.isna(contact[col]):
                contact_company = contact[col]
                break
                
        if not contact_company:
            continue
            
        # Compare normalized names
        contact_norm = normalize_company_name(contact_company)
        if not contact_norm:
            continue
            
        score = fuzz.token_sort_ratio(normalized_name, contact_norm)
        if score >= thresholds['company_name']:
            matches.append(contact)
    
    return matches

def configure_column_mapping(settings):
    """Configure column mapping between input and target files"""
    if not settings.get('input_file') or not settings.get('target_file'):
        print("\nError: Please select input and target files first")
        input("Press Enter to continue...")
        return settings
        
    input_df = try_read_csv(settings['input_file'])
    target_df = try_read_csv(settings['target_file'])
    
    if input_df is None or target_df is None:
        print("\nError: Could not read input or target files")
        input("Press Enter to continue...")
        return settings
    
    column_mapping = settings.get('column_mapping', {})
    
    while True:
        print("\nCurrent Column Mappings:")
        print("-" * 50)
        for i, (source, target) in enumerate(column_mapping.items(), 1):
            print(f"{i}. {source} → {target}")
        print(f"{len(column_mapping) + 1}. Add New Mapping")
        print(f"{len(column_mapping) + 2}. Done")
        
        try:
            choice = int(input("\nSelect mapping to modify (or select Done to finish): "))
            if choice == len(column_mapping) + 2:
                break
            elif choice == len(column_mapping) + 1:
                # Add new mapping
                print("\nAvailable source columns:")
                for i, col in enumerate(input_df.columns, 1):
                    print(f"{i}. {col}")
                source_idx = int(input("\nSelect source column number: ")) - 1
                source_col = input_df.columns[source_idx]
                
                print("\nAvailable target columns:")
                for i, col in enumerate(target_df.columns, 1):
                    print(f"{i}. {col}")
                target_idx = int(input("\nSelect target column number: ")) - 1
                target_col = target_df.columns[target_idx]
                
                column_mapping[source_col] = target_col
            elif 1 <= choice <= len(column_mapping):
                # Modify existing mapping
                source_col = list(column_mapping.keys())[choice - 1]
                
                print("\nAvailable source columns:")
                for i, col in enumerate(input_df.columns, 1):
                    print(f"{i}. {col}")
                source_idx = int(input("\nSelect new source column number (or 0 to delete mapping): ")) - 1
                
                if source_idx == -1:
                    # Delete mapping
                    del column_mapping[source_col]
                else:
                    # Update mapping
                    new_source = input_df.columns[source_idx]
                    print("\nAvailable target columns:")
                    for i, col in enumerate(target_df.columns, 1):
                        print(f"{i}. {col}")
                    target_idx = int(input("\nSelect target column number: ")) - 1
                    target_col = target_df.columns[target_idx]
                    
                    if new_source != source_col:
                        del column_mapping[source_col]
                    column_mapping[new_source] = target_col
            
        except (ValueError, IndexError):
            print("\nInvalid selection. Please try again.")
            input("Press Enter to continue...")
            continue
            
    settings['column_mapping'] = column_mapping
    return settings

def modify_thresholds(settings):
    """Modify matching thresholds for different entity types"""
    thresholds = settings.get('thresholds', {
        'company_name': 85,  # For company name matching
        'person_name': 85,   # For first/last name matching
        'email': 100,        # For email exact matching
        'title': 70,         # For job title fuzzy matching
        'department': 70     # For department fuzzy matching
    })
    
    print("\nCurrent Matching Thresholds:")
    print("-" * 50)
    for entity, value in thresholds.items():
        print(f"{entity.replace('_', ' ').title()}: {value}%")
    print("-" * 50)
    
    print("\nWhich threshold would you like to modify?")
    print("1. Company Name Matching (used for initial company matching)")
    print("2. Person Name Matching (for first/last name comparison)")
    print("3. Email Matching (100% recommended for exact matches)")
    print("4. Title Matching (for job title comparison)")
    print("5. Department Matching (for department/function matching)")
    print("6. Done")
    
    mapping = {
        '1': ('company_name', 'company name'),
        '2': ('person_name', 'person name'),
        '3': ('email', 'email'),
        '4': ('title', 'job title'),
        '5': ('department', 'department')
    }
    
    while True:
        choice = input("\nEnter your choice (1-6): ").strip()
        if choice == '6':
            break
            
        if choice in mapping:
            key, desc = mapping[choice]
            while True:
                try:
                    print(f"\nCurrent {desc} threshold: {thresholds[key]}%")
                    print("Recommended ranges:")
                    print("- Company/Person Name: 80-90%")
                    print("- Email: 100% (exact match)")
                    print("- Title/Department: 70-80%")
                    value = int(input(f"\nEnter new threshold for {desc} (0-100): "))
                    if 0 <= value <= 100:
                        thresholds[key] = value
                        break
                    print("Threshold must be between 0 and 100")
                except ValueError:
                    print("Please enter a valid number")
    
    settings['thresholds'] = thresholds
    return settings

def modify_settings(settings):
    """Allow user to modify settings interactively"""
    while True:
        print("\nWhat would you like to modify?")
        print("1. Input Files")
        print("2. Target File")
        print("3. Matching Thresholds")
        print("4. Column Mapping")
        print("5. Done")
        
        choice = input("Enter your choice (1-5): ").strip()
        
        if choice == '1':
            files = select_multiple_files("Select input files")
            if files:
                settings['input_files'] = files
                print(f"\nSelected input files: {', '.join(files)}")
        elif choice == '2':
            file = select_file("Select target file")
            if file:
                settings['target_file'] = file
                print(f"\nSelected target file: {file}")
        elif choice == '3':
            settings = modify_thresholds(settings)
        elif choice == '4':
            settings = configure_column_mapping(settings)
        else:
            break
    
    # Save modified settings
    with open(SETTINGS_FILE, 'w') as f:
        json.dump(settings, f, indent=2)
    
    return settings

def validate_settings(settings):
    """Validate that all required settings are configured"""
    required = ['input_file', 'target_file', 'column_mapping']
    missing = [field for field in required if not settings.get(field)]
    
    if missing:
        print("\nMissing required settings:", ", ".join(missing))
        return False
        
    if not os.path.exists(settings['input_file']):
        print(f"\nInput file {settings['input_file']} not found")
        return False
        
    if not os.path.exists(settings['target_file']):
        print(f"\nTarget file {settings['target_file']} not found")
        return False
    
    return True

def main():
    """Main function to run the contact matcher"""
    settings = load_settings()
    if not settings:
        settings = {}
    
    while True:
        # Always show current settings first
        display_current_settings(settings)
        
        print("\nMain Menu:")
        print("1. Select Input File")
        print("2. Select Target File")
        print("3. Modify Thresholds")
        print("4. Configure Column Mapping")
        print("5. Run Program")
        print("6. Exit")
        
        choice = input("\nEnter your choice (1-6): ").strip()
        
        if choice == '1':
            settings['input_file'] = select_file("Select input file")
        elif choice == '2':
            settings['target_file'] = select_file("Select target file")
        elif choice == '3':
            settings = modify_thresholds(settings)
        elif choice == '4':
            settings = configure_column_mapping(settings)
        elif choice == '5':
            if validate_settings(settings):
                input_contacts = try_read_csv(settings['input_file'])
                target_contacts = try_read_csv(settings['target_file'])
                company_matches = find_matches(settings['input_file'], settings['target_file'], settings['thresholds'])
                write_overlap_report(company_matches, settings['input_file'], settings['target_file'])
                input("\nPress Enter to return to main menu...")
            else:
                print("\nPlease configure all required settings before running.")
                input("Press Enter to continue...")
        elif choice == '6':
            print("\nExiting program...")
            break
        else:
            print("\nInvalid choice. Please try again.")
            input("Press Enter to continue...")
        
        # Save settings after each modification
        save_settings(settings)

if __name__ == "__main__":
    main()