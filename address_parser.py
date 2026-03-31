import usaddress
import re

PROVINCES = {
    'ON': 'Ontario', 'QC': 'Quebec', 'BC': 'British Columbia', 'AB': 'Alberta',
    'NS': 'Nova Scotia', 'MB': 'Manitoba', 'SK': 'Saskatchewan', 'NB': 'New Brunswick',
    'NL': 'Newfoundland and Labrador', 'PE': 'Prince Edward Island', 'YT': 'Yukon',
    'NT': 'Northwest Territories', 'NU': 'Nunavut'
}

STATES = {
    'AL': 'Alabama', 'AK': 'Alaska', 'AZ': 'Arizona', 'AR': 'Arkansas', 'CA': 'California',
    'CO': 'Colorado', 'CT': 'Connecticut', 'DE': 'Delaware', 'FL': 'Florida', 'GA': 'Georgia',
    'HI': 'Hawaii', 'ID': 'Idaho', 'IL': 'Illinois', 'IN': 'Indiana', 'IA': 'Iowa',
    'KS': 'Kansas', 'KY': 'Kentucky', 'LA': 'Louisiana', 'ME': 'Maine', 'MD': 'Maryland',
    'MA': 'Massachusetts', 'MI': 'Michigan', 'MN': 'Minnesota', 'MS': 'Mississippi', 'MO': 'Missouri',
    'MT': 'Montana', 'NE': 'Nebraska', 'NV': 'Nevada', 'NH': 'New Hampshire', 'NJ': 'New Jersey',
    'NM': 'New Mexico', 'NY': 'New York', 'NC': 'North Carolina', 'ND': 'North Dakota', 'OH': 'Ohio',
    'OK': 'Oklahoma', 'OR': 'Oregon', 'PA': 'Pennsylvania', 'RI': 'Rhode Island', 'SC': 'South Carolina',
    'SD': 'South Dakota', 'TN': 'Tennessee', 'TX': 'Texas', 'UT': 'Utah', 'VT': 'Vermont',
    'VA': 'Virginia', 'WA': 'Washington', 'WV': 'West Virginia', 'WI': 'Wisconsin', 'WY': 'Wyoming'
}

COUNTRIES = {
    'AFGHANISTAN': 'AF', 'ALBANIA': 'AL', 'ALGERIA': 'DZ', 'ARGENTINA': 'AR', 'AUSTRALIA': 'AU',
    'AUSTRIA': 'AT', 'BAHRAIN': 'BH', 'BANGLADESH': 'BD', 'BELGIUM': 'BE', 'BRAZIL': 'BR',
    'BULGARIA': 'BG', 'CANADA': 'CA', 'CHILE': 'CL', 'CHINA': 'CN', 'COLOMBIA': 'CO',
    'CROATIA': 'HR', 'CZECH REPUBLIC': 'CZ', 'DENMARK': 'DK', 'ECUADOR': 'EC', 'EGYPT': 'EG',
    'ESTONIA': 'EE', 'FINLAND': 'FI', 'FRANCE': 'FR', 'GERMANY': 'DE', 'GHANA': 'GH',
    'GREECE': 'GR', 'HONG KONG': 'HK', 'HUNGARY': 'HU', 'INDIA': 'IN', 'INDONESIA': 'ID',
    'IRAN': 'IR', 'IRAQ': 'IQ', 'IRELAND': 'IE', 'ISRAEL': 'IL', 'ITALY': 'IT',
    'JAPAN': 'JP', 'JORDAN': 'JO', 'KENYA': 'KE', 'KOREA': 'KR', 'SOUTH KOREA': 'KR',
    'KUWAIT': 'KW', 'LATVIA': 'LV', 'LEBANON': 'LB', 'LITHUANIA': 'LT', 'LUXEMBOURG': 'LU',
    'MALAYSIA': 'MY', 'MEXICO': 'MX', 'MOROCCO': 'MA', 'NETHERLANDS': 'NL', 'NEW ZEALAND': 'NZ',
    'NIGERIA': 'NG', 'NORWAY': 'NO', 'OMAN': 'OM', 'PAKISTAN': 'PK', 'PERU': 'PE',
    'PHILIPPINES': 'PH', 'POLAND': 'PL', 'PORTUGAL': 'PT', 'QATAR': 'QA', 'ROMANIA': 'RO',
    'RUSSIA': 'RU', 'SAUDI ARABIA': 'SA', 'SINGAPORE': 'SG', 'SLOVAKIA': 'SK', 'SLOVENIA': 'SI',
    'SOUTH AFRICA': 'ZA', 'SPAIN': 'ES', 'SRI LANKA': 'LK', 'SWEDEN': 'SE', 'SWITZERLAND': 'CH',
    'TAIWAN': 'TW', 'THAILAND': 'TH', 'TUNISIA': 'TN', 'TURKEY': 'TR', 'UKRAINE': 'UA',
    'UNITED ARAB EMIRATES': 'AE', 'UAE': 'AE', 'UNITED KINGDOM': 'GB', 'UK': 'GB',
    'UNITED STATES': 'US', 'USA': 'US', 'VENEZUELA': 'VE', 'VIETNAM': 'VN'
}

def parse_address_string(address_str):
    """
    Parses a single address block into a dictionary of fields.
    Detects leading 1Z tracking numbers and extracts them first.
    """
    if not address_str:
        return None
        
    parsed = {
        'Street': '',
        'Suite': '',
        'City': '',
        'State': '',
        'Zip': '',
        'Country': 'CA',
        'Phone': '',
        'CompanyName': '',
        'ContactName': ''
    }
    
    # 1. Detect and extract leading 1Z tracking number
    lines = [l.strip() for l in address_str.strip().splitlines() if l.strip()]
    
    # Check for tracking numbers in any line, but focus on the first
    tracking_found = None
    for i, line in enumerate(lines):
        tracking_match = re.search(r'\b(1Z[A-Z0-9]{16})\b', line, re.IGNORECASE)
        if tracking_match:
            tracking_found = tracking_match.group(1).upper()
            parsed["TrackingNumber"] = tracking_found
            # Remove the tracking number from the line
            lines[i] = line.replace(tracking_match.group(0), "").strip()
            break # Only take the first one found

    # Remove empty lines that might have been created
    lines = [l for l in lines if l]
            
    try:
        # 1. Extract Phone Number (Improved Regex)
        # Matches: (514) 288-6664, 514-288-6664, 514.288.6664, 5142886664, etc.
        phone_pattern = r'(\+?1?[\s.-]?)?\(?\d{3}\)?[\s.-]?\d{3}[\s.-]?\d{4}'
        for i, line in enumerate(lines):
            phone_match = re.search(phone_pattern, line)
            if phone_match:
                parsed['Phone'] = phone_match.group(0).strip()
                lines[i] = line.replace(phone_match.group(0), '').strip()
                break
        
        lines = [l for l in lines if l]

        # 2. Extract Canadian Postal Code if present
        postal_pattern = r'\b([A-Za-z]\d[A-Za-z][\s-]?\d[A-Za-z]\d)\b'
        for i, line in enumerate(lines):
            postal_match = re.search(postal_pattern, line)
            if postal_match:
                parsed['Zip'] = postal_match.group(1).upper().replace(' ', '').replace('-', '')
                parsed['Country'] = 'CA'
                lines[i] = line.replace(postal_match.group(0), '').strip()
                break

        lines = [l for l in lines if l]

        # Identify Country/State indicators
        for i, line in enumerate(lines):
            # Check for country name (full text → ISO code)
            line_upper_stripped = line.strip().upper()
            if line_upper_stripped in COUNTRIES:
                parsed['Country'] = COUNTRIES[line_upper_stripped]
                lines[i] = ''  # Remove the country line from remaining processing
                continue
            # Also check if country name appears anywhere in the line (e.g., "Barcelona, Spain")
            for country_name, country_code in COUNTRIES.items():
                if re.search(r'\b' + re.escape(country_name) + r'\b', line_upper_stripped):
                    parsed['Country'] = country_code
                    lines[i] = re.sub(r'\b' + re.escape(country_name) + r'\b', '', line, flags=re.IGNORECASE).strip().strip(',')
                    break
            
            # Proactive CA Province detection
            ca_indicators = ['ONTARIO', 'QUEBEC', 'ALBERTA', 'NOVA SCOTIA']
            if any(ind in line_upper_stripped for ind in ca_indicators):
                parsed['Country'] = 'CA'

            # Check CA Provinces
            for code, name in PROVINCES.items():
                if re.search(r'\b' + code + r'\b', line.upper()) or re.search(r'\b' + name.upper() + r'\b', line.upper()):
                    parsed['State'] = code
                    parsed['Country'] = 'CA'
                    break
            
            if parsed['State']: continue

            # Check US States
            for code, name in STATES.items():
                if re.search(r'\b' + code + r'\b', line.upper()) or re.search(r'\b' + name.upper() + r'\b', line.upper()):
                    parsed['State'] = code
                    parsed['Country'] = 'US'
                    break

        lines = [l for l in lines if l]
        # 3. Identify Company and Contact (Heuristic)
        # Find the street address line. It usually starts with a digit or looks like a Po Box.
        STREET_KEYWORDS = [' ave', ' ave.', ' st', ' st.', ' rd', ' rd.', ' blvd', ' street',
                           ' avenue', ' road', ' p.o', ' box ', ' calle', ' carrer', ' rue',
                           ' via ', ' corso', ' viale', ' way', ' dr', ' drive', ' ln', ' lane']
        street_idx = -1
        for i, line in enumerate(lines):
            if re.match(r'^\d+', line) or any(keyword in line.lower() for keyword in STREET_KEYWORDS):
                street_idx = i
                break
        
        if street_idx > 0:
            # Lines before the street index are likely Company/Contact
            header_lines = lines[:street_idx]
            if len(header_lines) == 1:
                parsed['CompanyName'] = header_lines[0]
            elif len(header_lines) >= 2:
                parsed['CompanyName'] = header_lines[1] # Often many labels put Contact then Company
                parsed['ContactName'] = header_lines[0]
                # If the first line is very short or looks like a person's name, maybe swap
                # Actually, many systems use LINE 1 for NAME and LINE 2 for COMPANY
                # But let's check for "Inc", "Corp", "Ltd" in header_lines[0]
                company_keywords = ['inc', 'corp', 'ltd', 'company', 'shop', 'services', 'logistics']
                if any(kw in header_lines[0].lower() for kw in company_keywords):
                    parsed['CompanyName'] = header_lines[0]
                    parsed['ContactName'] = header_lines[1]
                elif any(kw in header_lines[1].lower() for kw in company_keywords):
                    parsed['CompanyName'] = header_lines[1]
                    parsed['ContactName'] = header_lines[0]
                else:
                    # Default: Name then Company is common for person-specific mailings
                    # but Company then Name is common for business.
                    # We'll stick to 1=Company, 2=Contact if we can't tell?
                    # Actually, the user's manual UI has Company then Contact.
                    parsed['CompanyName'] = header_lines[0]
                    parsed['ContactName'] = header_lines[1]

        # Remove header lines and phone/postal modifications for usaddress fallback
        # Normalize the string: remove multiple spaces and newlines
        remaining_str = ' '.join(lines[street_idx:] if street_idx != -1 else lines).strip()
        clean_str = re.sub(r'\s+', ' ', remaining_str).strip()
        
        # 4. Parse remaining using usaddress
        if clean_str:
            try:
                tagged_address, address_type = usaddress.tag(clean_str)
            except usaddress.RepeatedLabelError as e:
                tagged_address = {}
                for val, label in e.parsed_string:
                    if label not in tagged_address:
                        tagged_address[label] = val
                    else:
                        tagged_address[label] += f" {val}"
            except Exception:
                tagged_address = {}
                
            street_parts = []
            potential_street_parts = [
                'AddressNumber', 'StreetNamePreDirectional', 'StreetName', 
                'StreetNamePostType', 'AddressNumberSuffix', 'StreetNamePostDirectional',
                'BuildingName', 'SubaddressIdentifier', 'SubaddressType',
                'LandmarkName', 'CornerOf', 'IntersectionSeparator'
            ]
            for part in potential_street_parts:
                if part in tagged_address:
                    street_parts.append(tagged_address[part])
            
            parsed['Street'] = ' '.join(street_parts).strip()
            if not parsed['Street'] and 'Recipient' in tagged_address:
                 parsed['Street'] = tagged_address['Recipient']

            if not parsed['CompanyName'] and 'Recipient' in tagged_address:
                 # If we didn't find company by heuristic, and usaddress found Recipient
                 parsed['CompanyName'] = tagged_address['Recipient']

            parsed['Suite'] = tagged_address.get('OccupancyIdentifier', '')
            parsed['City'] = tagged_address.get('PlaceName', '').strip().rstrip(',')
            
            if not parsed['State']:
                parsed['State'] = tagged_address.get('StateName', '').strip().rstrip(',')

            if not parsed['Zip']:
                parsed['Zip'] = tagged_address.get('ZipCode', '')
            
            if tagged_address.get('OccupancyType'):
                parsed['Suite'] = f"{tagged_address.get('OccupancyType')} {parsed['Suite']}".strip()

        # 5. Normalize State/Province to 2-letter codes
        if parsed['State']:
            state_upper = parsed['State'].strip().upper()
            # Map Full Name -> Code for Provinces
            name_to_code_prov = {v.upper(): k for k, v in PROVINCES.items()}
            if state_upper in name_to_code_prov:
                parsed['State'] = name_to_code_prov[state_upper]
            else:
                # Map Full Name -> Code for States
                name_to_code_state = {v.upper(): k for k, v in STATES.items()}
                if state_upper in name_to_code_state:
                    parsed['State'] = name_to_code_state[state_upper]
                # If it's already a code or not found, leave as is (but uppercase)
                elif len(state_upper) == 2:
                    parsed['State'] = state_upper

        return parsed
    except Exception as e:
        print(f"Error parsing address: {e}")
        return None

def split_addresses(raw_text):
    """
    Splits a raw block of text into individual address strings.
    Handles double newlines primarily, but falls back to single newlines
     if it detects a list of single-line addresses.
    """
    raw_text = raw_text.strip()
    if not raw_text:
        return []

    # If double newlines exist, use them as the primary delimiter
    if '\n\n' in raw_text or '\n \n' in raw_text:
        blocks = re.split(r'\n\s*\n', raw_text)
        # Check if any block looks like it contains multiple single-line addresses
        # (e.g. if a block has multiple lines and each line looks like an address)
        final_blocks = []
        for b in blocks:
            lines = [l.strip() for l in b.split('\n') if l.strip()]
            if len(lines) >= 2: # Heuristic: if 2 or more lines, check if it's a list
                # If most lines start with a number, it's likely a list
                if sum(1 for l in lines if re.match(r'^\d+', l)) >= len(lines) / 2:
                    final_blocks.extend(lines)
                else:
                    final_blocks.append(b)
            else:
                final_blocks.append(b)
        return final_blocks
    
    # Fallback to single newlines if no double newlines found
    lines = [line.strip() for line in raw_text.split('\n') if line.strip()]
    return lines

if __name__ == "__main__":
    test_batch = """
730-9600 rue Meilleur
Montréal, PQ  H2N 2E3
514.385.7909

123 Main St
Springfield, IL 62704
217-555-0199

456 Oak Ave Suite 10
Los Angeles, CA 90001
    """
    print("Testing Split Addresses:")
    split = split_addresses(test_batch)
    for i, addr in enumerate(split):
        print(f"Block {i+1}:\n{addr}")
        print(f"Parsed: {parse_address_string(addr)}\n")
