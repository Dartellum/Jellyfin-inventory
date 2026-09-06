# ==========================================
# 1. STANDARD LIBRARIES (Always available)
# ==========================================
import sys
import os
import subprocess
import glob
import re
import urllib.parse
from datetime import datetime
from difflib import SequenceMatcher

# ==========================================
# 2. STANDARDIZED FRONT-END: DEPENDENCY CHECK
# ==========================================
def in_virtual_environment():
    """Returns True if running inside a venv, virtualenv, or conda environment."""
    in_venv = sys.prefix != sys.base_prefix
    in_conda = "CONDA_PREFIX" in os.environ
    return in_venv or in_conda

def install_requirements():
    """Checks for required libraries and installs them gracefully."""
    requirements = ["pandas", "openpyxl", "requests"]
    missing = []

    for req in requirements:
        try:
            __import__(req)
        except ImportError:
            missing.append(req)

    if missing:
        print(f"[*] Missing required libraries: {', '.join(missing)}")

        if not in_virtual_environment():
            print("\n[WARNING] You are not currently in a Python virtual environment (venv/conda).")
            print("Installing packages globally can cause conflicts with system-level Python tools.")
            choice = input("Do you want to proceed with a global installation anyway? (y/N): ").strip().lower()
            if choice != 'y':
                print("\nAborting. Please activate a virtual environment and try again.")
                sys.exit(1)

        print("\n[*] Attempting to install dependencies automatically...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", *missing])
            print("[*] Dependencies installed successfully!\n")
        except subprocess.CalledProcessError:
            print("\n[ERROR] Automatic installation was blocked by your OS (PEP 668).")
            print(f"Even though an environment may be active, '{sys.executable}' is locked.")
            print("\nPlease run this command manually in your terminal to fix it:")
            print(f"    pip install {' '.join(missing)}\n")
            print("If your terminal still throws an externally-managed error, recreate your environment:")
            print("    python3 -m venv .venv")
            print("    source .venv/bin/activate")
            print(f"    pip install {' '.join(missing)}")
            sys.exit(1)

# Run the check before importing the rest of the script
install_requirements()

# ==========================================
# 3. MAIN SCRIPT LOGIC
# ==========================================
import pandas as pd
import requests
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font

def load_env_file(filepath=".env"):
    """Loads key-value pairs from .env into os.environ if not already set."""
    if os.path.exists(filepath):
        try:
            with open(filepath, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if line and not line.startswith("#") and "=" in line:
                        k, v = line.split("=", 1)
                        k = k.strip()
                        v = v.strip().strip("'\"")
                        if k and k not in os.environ:
                            os.environ[k] = v
        except Exception:
            pass

# Load environment configuration from .env if present
load_env_file()

# 1. TMDb V4 Read Access Token (from https://www.themoviedb.org/settings/api)
TMDB_V4_TOKEN = os.environ.get("TMDB_V4_TOKEN", "YOUR_TMDB_V4_TOKEN_HERE").strip().strip("'\"")

# 2. PATH TO YOUR ENTERTAINMENT ROOT
MEDIA_LIBRARY_DIR = os.environ.get("MEDIA_LIBRARY_DIR", "/path/to/your/Entertainment").strip().strip("'\"")
MOVIES_DIR = os.path.join(MEDIA_LIBRARY_DIR, "Movies")
TV_DIR = os.path.join(MEDIA_LIBRARY_DIR, "TV-Shows")

# 3. DIRECTORIES TO IGNORE
EXCLUDED_DIRECTORIES = [
    "lost+found",
    "timeshift",
    ".Trash-1000"
]

# Configure secure Bearer Auth headers for TMDB v4 Token globally
HEADERS = {
    "Authorization": f"Bearer {TMDB_V4_TOKEN}",
    "accept": "application/json"
}

def calculate_folder_size_gb(folder_path):
    """Calculates folder payload size in Gigabytes."""
    total_size = 0
    if not os.path.exists(folder_path):
        return 0
    for dirpath, _, filenames in os.walk(folder_path):
        for f in filenames:
            fp = os.path.join(dirpath, f)
            if not os.path.islink(fp):
                total_size += os.path.getsize(fp)
    return total_size / (1024 ** 3)

def analyze_format_and_upgrade(size_gb, is_tv=False, expected_seasons=1):
    """Deduces source quality from raw folder storage size and suggests upgrades."""
    if size_gb == 0:
        return "[ EMPTY FOLDER ]", "Action Required: Check/populate folder."

    # Adjust size expectations if it's a massive multi-season TV show folder
    divisor = expected_seasons if is_tv else 1
    adjusted_size = size_gb / divisor

    if adjusted_size < 4.5:
        current_fmt = "DVD / SD"
        rec = "Recommendation: Upgrade to 1080p Blu-ray or 4K UHD"
    elif 4.5 <= adjusted_size <= 12.0:
        current_fmt = "1080p Encode / Web-DL"
        rec = "Recommendation: Upgrade to uncompressed 1080p Blu-ray (Remux) or 4K UHD"
    elif 12.0 < adjusted_size <= 45.0:
        current_fmt = "1080p Blu-ray (Remux)"
        if is_tv:
            rec = "Optimal for TV (Most series cap at 1080p physical master)"
        else:
            rec = "Recommendation: Upgrade to 4K Ultra HD if a UHD disc exists"
    else:
        current_fmt = "4K Ultra HD"
        rec = "Optimal (Max 4K quality achieved)"

    return f"[{current_fmt}]", rec

def analyze_tv_season_format(size_gb, ep_count=1):
    """
    Deduces quality of an individual TV season based on episode count and folder size.
    """
    if size_gb == 0 or ep_count == 0:
        return "[ EMPTY FOLDER ]", "Action Required: Check/populate season folder."

    avg_ep_gb = size_gb / ep_count if ep_count > 0 else size_gb

    # If single episode (e.g., special / miniseries)
    if ep_count == 1:
        if size_gb < 1.0:
            return "[DVD / SD]", "Recommendation: Upgrade to 1080p Blu-ray or 4K UHD"
        elif 1.0 <= size_gb <= 4.0:
            return "[1080p Encode / Web-DL]", "Recommendation: Upgrade to uncompressed 1080p Blu-ray (Remux) or 4K UHD"
        elif 4.0 < size_gb <= 20.0:
            return "[1080p Blu-ray (Remux)]", "Optimal for TV (Most series cap at 1080p physical master)"
        else:
            return "[4K Ultra HD]", "Optimal (Max 4K quality achieved)"

    # Multi-episode season: evaluate based on average episode size
    if avg_ep_gb < 0.6:
        current_fmt = "DVD / SD"
        rec = "Recommendation: Upgrade to 1080p Blu-ray or 4K UHD"
    elif 0.6 <= avg_ep_gb <= 3.0:
        current_fmt = "1080p Encode / Web-DL"
        rec = "Recommendation: Upgrade to uncompressed 1080p Blu-ray (Remux) or 4K UHD"
    elif 3.0 < avg_ep_gb <= 12.0:
        current_fmt = "1080p Blu-ray (Remux)"
        rec = "Optimal for TV (Most series cap at 1080p physical master)"
    else:
        current_fmt = "4K Ultra HD"
        rec = "Optimal (Max 4K quality achieved)"

    return f"[{current_fmt}]", rec

def get_season_number(name):
    """
    Extracts season number from directory or file name.
    Recognizes 'Season 01', 'Season 1', 'S01', 'S1', 'Specials', 'Season 00', etc.
    """
    nl = name.lower()
    if "special" in nl:
        return 0
    m = re.search(r'\bseason\s*(\d+)\b', nl)
    if m:
        return int(m.group(1))
    m = re.search(r'\bs(\d{1,2})(?:e\d+)?\b', nl)
    if m:
        return int(m.group(1))
    return None

def get_movie_collection_data(movie_name, release_year=None):
    # Format the search query
    safe_name = urllib.parse.quote(movie_name)
    search_url = f"https://api.themoviedb.org/3/search/movie?query={safe_name}"

    # If we have a year, append it to force an exact match
    if release_year:
         search_url += f"&primary_release_year={release_year}"

    try:
        response = requests.get(search_url, headers=HEADERS)
        if response.status_code == 401: return "AUTH_ERROR", []
        res_json = response.json()
        if not res_json.get('results'): return None, []

        movie_id = res_json['results'][0]['id']
        movie_details = requests.get(f"https://api.themoviedb.org/3/movie/{movie_id}", headers=HEADERS).json()

        collection = movie_details.get('belongs_to_collection')
        if not collection: return None, []

        coll_details = requests.get(f"https://api.themoviedb.org/3/collection/{collection['id']}", headers=HEADERS).json()

        parts_list = []
        for part in coll_details.get('parts', []):
            title = part['title']
            if title == "Harry Potter and the Philosopher's Stone":
                title = "Harry Potter and the Sorcerer's Stone"
            parts_list.append(title)

        return collection['name'], parts_list
    except:
        return None, []

def get_tv_series_data(series_name, release_year=None):
    # Format the search query
    safe_name = urllib.parse.quote(series_name)
    search_url = f"https://api.themoviedb.org/3/search/tv?query={safe_name}"

    # For TV shows, TMDb uses first_air_date_year
    if release_year:
         search_url += f"&first_air_date_year={release_year}"

    try:
        response = requests.get(search_url, headers=HEADERS)
        if response.status_code == 401: return "AUTH_ERROR", 0, []
        res_json = response.json()
        if not res_json.get('results'): return None, 0, []

        tv_id = res_json['results'][0]['id']
        tv_details = requests.get(f"https://api.themoviedb.org/3/tv/{tv_id}", headers=HEADERS).json()

        seasons = tv_details.get('seasons', [])
        season_list = []
        for s in seasons:
            season_list.append({
                'season_number': s.get('season_number', 0),
                'name': s.get('name', f"Season {s.get('season_number', 0)}"),
                'episode_count': s.get('episode_count', 0)
            })
        return tv_details.get('name'), tv_details.get('number_of_seasons', 1), season_list
    except:
        return None, 0, []

def load_existing_upcs():
    """Finds the most recent Excel audit and extracts previously scanned UPC codes."""
    excel_files = glob.glob("Jellyfin_Audit_List*.xlsx")
    # Exclude dummy/sample files from being considered historical audits
    excel_files = [f for f in excel_files if not os.path.basename(f).startswith("sample_")]
    if not excel_files:
        return {}, {}

    # Find the newest spreadsheet file in the current directory
    latest_file = max(excel_files, key=os.path.getmtime)
    print(f"[*] Found previous audit file: {os.path.basename(latest_file)}")
    print("[*] Preserving existing UPC codes...")

    movie_upcs = {}
    tv_upcs = {}

    try:
        xls = pd.ExcelFile(latest_file)

        # Pull UPCs from Movies Audit
        if 'Movies Audit' in xls.sheet_names:
            df_m = pd.read_excel(xls, 'Movies Audit', dtype={'UPC Code': str})
            for _, row in df_m.dropna(subset=['UPC Code']).iterrows():
                title = str(row['Title']).strip()
                upc = str(row['UPC Code']).strip()
                if upc and upc.lower() != 'nan':
                    if upc.endswith('.0'): upc = upc[:-2]
                    movie_upcs[title] = upc

        # Pull UPCs from TV Shows Audit
        if 'TV Shows Audit' in xls.sheet_names:
            df_tv = pd.read_excel(xls, 'TV Shows Audit', dtype={'UPC Code': str})
            for _, row in df_tv.dropna(subset=['UPC Code']).iterrows():
                series = str(row['Series']).strip()
                item = str(row['Item/Season']).strip()
                upc = str(row['UPC Code']).strip()
                if upc and upc.lower() != 'nan':
                    if upc.endswith('.0'): upc = upc[:-2]
                    tv_upcs[(series, item)] = upc

    except Exception as e:
        print(f"[!] Warning: Could not read previous UPCs: {e}")

    return movie_upcs, tv_upcs

def audit_movies(existing_upcs={}):
    """Compiles movie data into a list of dictionaries for Excel export."""
    if not os.path.exists(MOVIES_DIR):
        print(f"[INFO] Movies folder '{MOVIES_DIR}' not found. Skipping.")
        return []

    actual_folders = os.listdir(MOVIES_DIR)
    audited_collections = set()
    movie_data = []

    print(f"Scanning {len(actual_folders)} Movie Folders...")

    for folder in actual_folders:
        if folder in EXCLUDED_DIRECTORIES:
            continue

        clean_name = folder.split('(')[0].strip()
        full_path = os.path.join(MOVIES_DIR, folder)
        if not os.path.isdir(full_path): continue

        # Extract the year from the folder name (e.g., grabs 1984 from "Dune (1984)")
        year_match = re.search(r'\((\d{4})\)', folder)
        folder_year = year_match.group(1) if year_match else None

        local_size = calculate_folder_size_gb(full_path)

        # Pass both title AND year to TMDb! No more hardcoded overrides needed.
        collection_name, franchise_checklist = get_movie_collection_data(clean_name, folder_year)

        if collection_name == "AUTH_ERROR":
            print("[ERROR] TMDb v4 Token Unauthorized.")
            return []

        if collection_name and collection_name not in audited_collections:
            audited_collections.add(collection_name)
            used_folders_in_collection = set()  # Tracks folders claimed in this collection

            for expected_movie in franchise_checklist:
                if expected_movie == "Harry Potter and the Philosopher's Stone":
                    expected_movie = "Harry Potter and the Sorcerer's Stone"

                matched_folder = None
                api_clean = "".join(c for c in expected_movie.lower() if c.isalnum())
                api_numbers = re.findall(r'\d+', api_clean)

                for actual in actual_folders:
                    # Skip excluded directories AND folders already matched to another movie
                    if actual in EXCLUDED_DIRECTORIES or actual in used_folders_in_collection:
                        continue

                    folder_title = re.sub(r'\s*\(\d{4}\).*', '', actual).strip()
                    local_clean = "".join(c for c in folder_title.lower() if c.isalnum())
                    local_numbers = re.findall(r'\d+', local_clean)

                    # Stage A: Exact Match
                    if api_clean == local_clean:
                        matched_folder = actual
                        break

                    # Stage B: Fuzzy Match (Forgives typos, enforces numbers)
                    similarity = SequenceMatcher(None, api_clean, local_clean).ratio()
                    if similarity >= 0.85 and api_numbers == local_numbers:
                        matched_folder = actual
                        break

                saved_upc = existing_upcs.get(expected_movie, "")

                if matched_folder:
                    used_folders_in_collection.add(matched_folder)  # Mark folder as claimed
                    m_path = os.path.join(MOVIES_DIR, matched_folder)
                    size = calculate_folder_size_gb(m_path)
                    fmt_string, rec_string = analyze_format_and_upgrade(size)

                    movie_data.append({
                        "Collection/Group": collection_name,
                        "Title": expected_movie,
                        "Status": "✓ Found",
                        "Format": fmt_string,
                        "Size (GB)": round(size, 2),
                        "Recommendation": rec_string,
                        "UPC Code": saved_upc
                    })
                else:
                    movie_data.append({
                        "Collection/Group": collection_name,
                        "Title": expected_movie,
                        "Status": "✕ Missing",
                        "Format": "-",
                        "Size (GB)": 0.0,
                        "Recommendation": "Missing Film",
                        "UPC Code": saved_upc
                    })
        elif not collection_name:
            fmt_string, rec_string = analyze_format_and_upgrade(local_size)
            saved_upc = existing_upcs.get(clean_name, "")
            movie_data.append({
                "Collection/Group": "Standalone",
                "Title": clean_name,
                "Status": "✓ Found",
                "Format": fmt_string,
                "Size (GB)": round(local_size, 2),
                "Recommendation": rec_string,
                "UPC Code": saved_upc
            })

    return movie_data

def audit_tv_shows(existing_upcs={}):
    """Compiles TV show data into a list of dictionaries for Excel export."""
    if not os.path.exists(TV_DIR):
        print(f"[INFO] TV Shows folder '{TV_DIR}' not found. Skipping.")
        return []

    actual_folders = sorted(os.listdir(TV_DIR))
    tv_data = []

    print(f"\nScanning {len(actual_folders)} TV Show Folders...")

    VIDEO_EXTENSIONS = ('.mkv', '.mp4', '.avi', '.m4v', '.ts', '.wmv')

    for folder in actual_folders:
        if folder in EXCLUDED_DIRECTORIES:
            continue

        full_path = os.path.join(TV_DIR, folder)
        if not os.path.isdir(full_path):
            continue

        # Extract the year and clean title, stripping any [imdbid-xxxx] tags
        clean_name = re.sub(r'\[.*?\]', '', folder)
        year_match = re.search(r'\((\d{4})\)', clean_name)
        folder_year = year_match.group(1) if year_match else None
        clean_name = clean_name.split('(')[0].strip()

        local_size = calculate_folder_size_gb(full_path)

        # Pass both title AND year to the API function
        official_name, total_seasons, season_list = get_tv_series_data(clean_name, folder_year)

        if official_name == "AUTH_ERROR":
            print("[ERROR] TMDb v4 Token Unauthorized.")
            return []

        if official_name:
            display_series = f"{official_name} ({folder_year})" if folder_year else official_name
            print(f"  ✓ Auditing TV Series: {display_series}")

            # 1. Complete Series overview row
            fmt_string, rec_string = analyze_format_and_upgrade(local_size, is_tv=True, expected_seasons=total_seasons)

            series_upc = (
                existing_upcs.get((display_series, "Complete Series"))
                or existing_upcs.get((official_name, "Complete Series"))
                or ""
            )

            tv_data.append({
                "Series": display_series,
                "Item/Season": "Complete Series",
                "Status": "✓ Found (Folder exists)",
                "Format Overall": fmt_string,
                "Size (GB)": round(local_size, 2),
                "Recommendation": rec_string,
                "UPC Code": series_upc
            })

            # 2. Map local season folders & standalone files to season numbers
            subfiles = os.listdir(full_path)
            local_seasons = {}

            for sf in subfiles:
                sf_path = os.path.join(full_path, sf)
                s_num = get_season_number(sf)
                if s_num is not None:
                    if os.path.isdir(sf_path):
                        v_files = [f for f in os.listdir(sf_path) if f.lower().endswith(VIDEO_EXTENSIONS) and not f.startswith('.')]
                        local_seasons[s_num] = {
                            'path': sf_path,
                            'is_dir': True,
                            'files': v_files,
                            'name': sf
                        }
                    elif sf.lower().endswith(VIDEO_EXTENSIONS):
                        if s_num not in local_seasons:
                            local_seasons[s_num] = {
                                'path': sf_path,
                                'is_dir': False,
                                'files': [sf],
                                'name': sf
                            }
                        else:
                            local_seasons[s_num]['files'].append(sf)

            # Fallback for single-season shows with loose files in root folder without season tags
            if not local_seasons and total_seasons == 1:
                loose_videos = [sf for sf in subfiles if sf.lower().endswith(VIDEO_EXTENSIONS) and not sf.startswith('.')]
                if loose_videos:
                    local_seasons[1] = {
                        'path': full_path,
                        'is_dir': False,
                        'files': loose_videos,
                        'name': 'Root'
                    }

            # 3. Process Season 0 (Specials) if present locally on disk
            if 0 in local_seasons:
                s_info = local_seasons[0]
                ep_count = len(s_info['files'])
                if s_info['is_dir']:
                    s_size = calculate_folder_size_gb(s_info['path'])
                else:
                    s_size = sum(os.path.getsize(os.path.join(full_path, f)) for f in s_info['files']) / (1024 ** 3)

                fmt, rec = analyze_tv_season_format(s_size, ep_count)
                status_str = f"✓ Found ({ep_count} eps)" if ep_count > 0 else "✕ Empty (0 eps)"
                item_label = "Season 0 (Specials)"
                s_upc = (
                    existing_upcs.get((display_series, item_label))
                    or existing_upcs.get((official_name, item_label))
                    or ""
                )
                tv_data.append({
                    "Series": display_series,
                    "Item/Season": item_label,
                    "Status": status_str,
                    "Format Overall": fmt,
                    "Size (GB)": round(s_size, 2) if ep_count > 0 else 0.0,
                    "Recommendation": rec,
                    "UPC Code": s_upc
                })

            # 4. Process broadcast seasons 1 through total_seasons
            tmdb_seasons = {s['season_number']: s for s in season_list if s['season_number'] > 0}

            for idx in range(1, total_seasons + 1):
                s_meta = tmdb_seasons.get(idx, {})
                season_name = s_meta.get('name', f"Season {idx}")
                expected_eps = s_meta.get('episode_count', 0)
                item_label = f"Season {idx} ({season_name})"

                s_upc = (
                    existing_upcs.get((display_series, item_label))
                    or existing_upcs.get((official_name, item_label))
                    or ""
                )

                if idx in local_seasons:
                    s_info = local_seasons[idx]
                    ep_count = len(s_info['files'])

                    if s_info['is_dir']:
                        s_size = calculate_folder_size_gb(s_info['path'])
                    else:
                        s_size = sum(os.path.getsize(os.path.join(full_path, f)) for f in s_info['files']) / (1024 ** 3)

                    if ep_count > 0:
                        fmt, rec = analyze_tv_season_format(s_size, ep_count)
                        if expected_eps > 0 and ep_count < expected_eps:
                            status_str = f"✓ Found (Partial: {ep_count}/{expected_eps} eps)"
                        else:
                            status_str = f"✓ Found ({ep_count} eps)"

                        tv_data.append({
                            "Series": display_series,
                            "Item/Season": item_label,
                            "Status": status_str,
                            "Format Overall": fmt,
                            "Size (GB)": round(s_size, 2),
                            "Recommendation": rec,
                            "UPC Code": s_upc
                        })
                    else:
                        tv_data.append({
                            "Series": display_series,
                            "Item/Season": item_label,
                            "Status": "✕ Empty (0 eps)",
                            "Format Overall": "[ EMPTY FOLDER ]",
                            "Size (GB)": 0.0,
                            "Recommendation": "Action Required: Check/populate season folder.",
                            "UPC Code": s_upc
                        })
                else:
                    tv_data.append({
                        "Series": display_series,
                        "Item/Season": item_label,
                        "Status": "✕ Missing",
                        "Format Overall": "-",
                        "Size (GB)": 0.0,
                        "Recommendation": "Missing Season",
                        "UPC Code": s_upc
                    })
        else:
            unknown_upc = existing_upcs.get((clean_name, "Unrecognized Directory"), "")
            tv_data.append({
                "Series": clean_name,
                "Item/Season": "Unrecognized Directory",
                "Status": "? Unknown",
                "Format Overall": "-",
                "Size (GB)": round(local_size, 2),
                "Recommendation": "-",
                "UPC Code": unknown_upc
            })

    return tv_data

def export_to_excel(movies, tv_shows, output_file="Jellyfin_Audit_List.xlsx"):
    """Writes the collected dictionary lists to an Excel workbook with formatting."""
    print("\nWriting data to Excel with formatting...")

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:

        # --- Define our Red Highlight Style ---
        red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        red_font = Font(color="9C0006")

        # --- Helper function for formatting sheets ---
        def format_sheet(sheet_name, df):
            worksheet = writer.sheets[sheet_name]

            # 1. Freeze the top header row
            worksheet.freeze_panes = 'A2'

            # 2. Enable Sorting & Filtering Dropdowns
            worksheet.auto_filter.ref = worksheet.dimensions

            # 3. Auto-expand Column Widths
            for idx, col_name in enumerate(df.columns, start=1):
                series_lengths = df[col_name].astype(str).map(len)
                max_len = max(series_lengths.max() if not series_lengths.empty else 0, len(str(col_name)))
                adjusted_width = min(max_len + 3, 60)
                col_letter = get_column_letter(idx)
                worksheet.column_dimensions[col_letter].width = adjusted_width

            # 4. Highlight "Missing" and "Empty" Rows in Red
            if "Status" in df.columns:
                status_col_idx = df.columns.get_loc("Status") + 1

                for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row):
                    status_cell = row[status_col_idx - 1]

                    if status_cell.value and any(k in str(status_cell.value) for k in ["Missing", "Empty"]):
                        for cell in row:
                            cell.fill = red_fill
                            cell.font = red_font

            # 5. Format UPC column as Text to prevent dropping leading zeros
            if "UPC Code" in df.columns:
                upc_col_idx = df.columns.get_loc("UPC Code") + 1
                upc_col_letter = get_column_letter(upc_col_idx)

                for cell in worksheet[upc_col_letter]:
                    cell.number_format = '@'

        # --- Process Movies ---
        if movies:
            df_movies = pd.DataFrame(movies)
            df_movies.to_excel(writer, sheet_name='Movies Audit', index=False)
            format_sheet('Movies Audit', df_movies)

        # --- Process TV Shows ---
        if tv_shows:
            df_tv = pd.DataFrame(tv_shows)
            df_tv.to_excel(writer, sheet_name='TV Shows Audit', index=False)
            format_sheet('TV Shows Audit', df_tv)

    print(f"Success! Audit saved to: {os.path.abspath(output_file)}")

if __name__ == "__main__":
    if not TMDB_V4_TOKEN or TMDB_V4_TOKEN == "YOUR_TMDB_V4_TOKEN_HERE":
        print("\n" + "=" * 65)
        print("[!] TMDb V4 Read Access Token Not Configured!")
        print("=" * 65)
        print("Please configure your token by creating a .env file or setting the environment variable:")
        print("  1. Copy .env.example to .env:  cp .env.example .env")
        print("  2. Add your token to .env:     TMDB_V4_TOKEN=\"your_token_here\"")
        print("You can get a free token from: https://www.themoviedb.org/settings/api\n")
        sys.exit(1)
    else:
        print("Starting Media Library Audit...\n")

        # 1. Load any UPCs from previous exports
        prev_movie_upcs, prev_tv_upcs = load_existing_upcs()

        # 2. Run audits using the saved UPC maps
        movies_data = audit_movies(prev_movie_upcs)
        tv_data = audit_tv_shows(prev_tv_upcs)

        if movies_data or tv_data:
            timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            filename = f"Jellyfin_Audit_List_{timestamp}.xlsx"

            export_to_excel(movies_data, tv_data, output_file=filename)
        else:
            print("No data was found to export.")