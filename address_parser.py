import usaddress
import re
import logging

# Configure logging to match main app or standalone
if not logging.getLogger().hasHandlers():
    logging.basicConfig(level=logging.INFO, filename='ups_api.log', format='%(asctime)s - %(levelname)s - %(message)s')

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

# US ZIP prefix → State code lookup (covers all standard ranges)
ZIP_PREFIX_TO_STATE = {
    '006': 'PR', '007': 'PR', '008': 'PR', '009': 'PR',
    '010': 'MA', '011': 'MA', '012': 'MA', '013': 'MA', '014': 'MA', '015': 'MA', '016': 'MA', '017': 'MA', '018': 'MA', '019': 'MA',
    '020': 'MA', '021': 'MA', '022': 'MA', '023': 'MA', '024': 'MA', '025': 'MA', '026': 'MA', '027': 'MA',
    '028': 'RI', '029': 'RI',
    '030': 'NH', '031': 'NH', '032': 'NH', '033': 'NH', '034': 'NH', '035': 'NH', '036': 'NH', '037': 'NH', '038': 'NH',
    '039': 'ME', '040': 'ME', '041': 'ME', '042': 'ME', '043': 'ME', '044': 'ME', '045': 'ME', '046': 'ME', '047': 'ME', '048': 'ME', '049': 'ME',
    '050': 'VT', '051': 'VT', '052': 'VT', '053': 'VT', '054': 'VT', '055': 'VT', '056': 'VT', '057': 'VT', '058': 'VT', '059': 'VT',
    '060': 'CT', '061': 'CT', '062': 'CT', '063': 'CT', '064': 'CT', '065': 'CT', '066': 'CT', '067': 'CT', '068': 'CT', '069': 'CT',
    '070': 'NJ', '071': 'NJ', '072': 'NJ', '073': 'NJ', '074': 'NJ', '075': 'NJ', '076': 'NJ', '077': 'NJ', '078': 'NJ', '079': 'NJ',
    '080': 'NJ', '081': 'NJ', '082': 'NJ', '083': 'NJ', '084': 'NJ', '085': 'NJ', '086': 'NJ', '087': 'NJ', '088': 'NJ', '089': 'NJ',
    '100': 'NY', '101': 'NY', '102': 'NY', '103': 'NY', '104': 'NY', '105': 'NY', '106': 'NY', '107': 'NY', '108': 'NY', '109': 'NY',
    '110': 'NY', '111': 'NY', '112': 'NY', '113': 'NY', '114': 'NY', '115': 'NY', '116': 'NY', '117': 'NY', '118': 'NY', '119': 'NY',
    '120': 'NY', '121': 'NY', '122': 'NY', '123': 'NY', '124': 'NY', '125': 'NY', '126': 'NY', '127': 'NY', '128': 'NY', '129': 'NY',
    '130': 'NY', '131': 'NY', '132': 'NY', '133': 'NY', '134': 'NY', '135': 'NY', '136': 'NY', '137': 'NY', '138': 'NY', '139': 'NY',
    '140': 'NY', '141': 'NY', '142': 'NY', '143': 'NY', '144': 'NY', '145': 'NY', '146': 'NY', '147': 'NY', '148': 'NY', '149': 'NY',
    '150': 'PA', '151': 'PA', '152': 'PA', '153': 'PA', '154': 'PA', '155': 'PA', '156': 'PA', '157': 'PA', '158': 'PA', '159': 'PA',
    '160': 'PA', '161': 'PA', '162': 'PA', '163': 'PA', '164': 'PA', '165': 'PA', '166': 'PA', '167': 'PA', '168': 'PA', '169': 'PA',
    '170': 'PA', '171': 'PA', '172': 'PA', '173': 'PA', '174': 'PA', '175': 'PA', '176': 'PA', '177': 'PA', '178': 'PA', '179': 'PA',
    '180': 'PA', '181': 'PA', '182': 'PA', '183': 'PA', '184': 'PA', '185': 'PA', '186': 'PA', '187': 'PA', '188': 'PA', '189': 'PA',
    '190': 'PA', '191': 'PA', '192': 'PA', '193': 'PA', '194': 'PA', '195': 'PA', '196': 'PA',
    '197': 'DE', '198': 'DE', '199': 'DE',
    '200': 'DC', '201': 'VA', '202': 'DC', '203': 'DC', '204': 'DC', '205': 'DC',
    '206': 'MD', '207': 'MD', '208': 'MD', '209': 'MD', '210': 'MD', '211': 'MD', '212': 'MD', '214': 'MD', '215': 'MD', '216': 'MD', '217': 'MD', '218': 'MD', '219': 'MD',
    '220': 'VA', '221': 'VA', '222': 'VA', '223': 'VA', '224': 'VA', '225': 'VA', '226': 'VA', '227': 'VA', '228': 'VA', '229': 'VA',
    '230': 'VA', '231': 'VA', '232': 'VA', '233': 'VA', '234': 'VA', '235': 'VA', '236': 'VA', '237': 'VA', '238': 'VA', '239': 'VA',
    '240': 'VA', '241': 'VA', '242': 'VA', '243': 'VA', '244': 'VA', '245': 'VA', '246': 'VA',
    '247': 'WV', '248': 'WV', '249': 'WV', '250': 'WV', '251': 'WV', '252': 'WV', '253': 'WV', '254': 'WV', '255': 'WV', '256': 'WV', '257': 'WV', '258': 'WV', '259': 'WV',
    '260': 'WV', '261': 'WV', '262': 'WV', '263': 'WV', '264': 'WV', '265': 'WV', '266': 'WV', '267': 'WV', '268': 'WV',
    '270': 'NC', '271': 'NC', '272': 'NC', '273': 'NC', '274': 'NC', '275': 'NC', '276': 'NC', '277': 'NC', '278': 'NC', '279': 'NC',
    '280': 'NC', '281': 'NC', '282': 'NC', '283': 'NC', '284': 'NC', '285': 'NC', '286': 'NC', '287': 'NC', '288': 'NC', '289': 'NC',
    '290': 'SC', '291': 'SC', '292': 'SC', '293': 'SC', '294': 'SC', '295': 'SC', '296': 'SC', '297': 'SC', '298': 'SC', '299': 'SC',
    '300': 'GA', '301': 'GA', '302': 'GA', '303': 'GA', '304': 'GA', '305': 'GA', '306': 'GA', '307': 'GA', '308': 'GA', '309': 'GA',
    '310': 'GA', '311': 'GA', '312': 'GA', '313': 'GA', '314': 'GA', '315': 'GA', '316': 'GA', '317': 'GA', '318': 'GA', '319': 'GA',
    '320': 'FL', '321': 'FL', '322': 'FL', '323': 'FL', '324': 'FL', '325': 'FL', '326': 'FL', '327': 'FL', '328': 'FL', '329': 'FL',
    '330': 'FL', '331': 'FL', '332': 'FL', '333': 'FL', '334': 'FL', '335': 'FL', '336': 'FL', '337': 'FL', '338': 'FL', '339': 'FL',
    '340': 'FL', '341': 'FL', '342': 'FL', '344': 'FL', '346': 'FL', '347': 'FL', '349': 'FL',
    '350': 'AL', '351': 'AL', '352': 'AL', '353': 'AL', '354': 'AL', '355': 'AL', '356': 'AL', '357': 'AL', '358': 'AL', '359': 'AL',
    '360': 'AL', '361': 'AL', '362': 'AL', '363': 'AL', '364': 'AL', '365': 'AL', '366': 'AL', '367': 'AL', '368': 'AL', '369': 'AL',
    '370': 'TN', '371': 'TN', '372': 'TN', '373': 'TN', '374': 'TN', '375': 'TN', '376': 'TN', '377': 'TN', '378': 'TN', '379': 'TN',
    '380': 'TN', '381': 'TN', '382': 'TN', '383': 'TN', '384': 'TN', '385': 'TN',
    '386': 'MS', '387': 'MS', '388': 'MS', '389': 'MS', '390': 'MS', '391': 'MS', '392': 'MS', '393': 'MS', '394': 'MS', '395': 'MS', '396': 'MS', '397': 'MS',
    '398': 'GA', '399': 'GA',
    '400': 'KY', '401': 'KY', '402': 'KY', '403': 'KY', '404': 'KY', '405': 'KY', '406': 'KY', '407': 'KY', '408': 'KY', '409': 'KY',
    '410': 'KY', '411': 'KY', '412': 'KY', '413': 'KY', '414': 'KY', '415': 'KY', '416': 'KY', '417': 'KY', '418': 'KY',
    '420': 'KY', '421': 'KY', '422': 'KY', '423': 'KY', '424': 'KY', '425': 'KY', '426': 'KY', '427': 'KY',
    '430': 'OH', '431': 'OH', '432': 'OH', '433': 'OH', '434': 'OH', '435': 'OH', '436': 'OH', '437': 'OH', '438': 'OH', '439': 'OH',
    '440': 'OH', '441': 'OH', '442': 'OH', '443': 'OH', '444': 'OH', '445': 'OH', '446': 'OH', '447': 'OH', '448': 'OH', '449': 'OH',
    '450': 'OH', '451': 'OH', '452': 'OH', '453': 'OH', '454': 'OH', '455': 'OH', '456': 'OH', '457': 'OH', '458': 'OH',
    '460': 'IN', '461': 'IN', '462': 'IN', '463': 'IN', '464': 'IN', '465': 'IN', '466': 'IN', '467': 'IN', '468': 'IN', '469': 'IN',
    '470': 'IN', '471': 'IN', '472': 'IN', '473': 'IN', '474': 'IN', '475': 'IN', '476': 'IN', '477': 'IN', '478': 'IN', '479': 'IN',
    '480': 'MI', '481': 'MI', '482': 'MI', '483': 'MI', '484': 'MI', '485': 'MI', '486': 'MI', '487': 'MI', '488': 'MI', '489': 'MI',
    '490': 'MI', '491': 'MI', '492': 'MI', '493': 'MI', '494': 'MI', '495': 'MI', '496': 'MI', '497': 'MI', '498': 'MI', '499': 'MI',
    '500': 'IA', '501': 'IA', '502': 'IA', '503': 'IA', '504': 'IA', '505': 'IA', '506': 'IA', '507': 'IA', '508': 'IA', '509': 'IA',
    '510': 'IA', '511': 'IA', '512': 'IA', '513': 'IA', '514': 'IA', '515': 'IA', '516': 'IA', '520': 'IA', '521': 'IA', '522': 'IA', '523': 'IA', '524': 'IA', '525': 'IA', '526': 'IA', '527': 'IA', '528': 'IA',
    '530': 'WI', '531': 'WI', '532': 'WI', '534': 'WI', '535': 'WI', '537': 'WI', '538': 'WI', '539': 'WI',
    '540': 'WI', '541': 'WI', '542': 'WI', '543': 'WI', '544': 'WI', '545': 'WI', '546': 'WI', '547': 'WI', '548': 'WI', '549': 'WI',
    '550': 'MN', '551': 'MN', '553': 'MN', '554': 'MN', '555': 'MN', '556': 'MN', '557': 'MN', '558': 'MN', '559': 'MN',
    '560': 'MN', '561': 'MN', '562': 'MN', '563': 'MN', '564': 'MN', '565': 'MN', '566': 'MN', '567': 'MN',
    '570': 'SD', '571': 'SD', '572': 'SD', '573': 'SD', '574': 'SD', '575': 'SD', '576': 'SD', '577': 'SD',
    '580': 'ND', '581': 'ND', '582': 'ND', '583': 'ND', '584': 'ND', '585': 'ND', '586': 'ND', '587': 'ND', '588': 'ND',
    '590': 'MT', '591': 'MT', '592': 'MT', '593': 'MT', '594': 'MT', '595': 'MT', '596': 'MT', '597': 'MT', '598': 'MT', '599': 'MT',
    '600': 'IL', '601': 'IL', '602': 'IL', '603': 'IL', '604': 'IL', '605': 'IL', '606': 'IL', '607': 'IL', '608': 'IL', '609': 'IL',
    '610': 'IL', '611': 'IL', '612': 'IL', '613': 'IL', '614': 'IL', '615': 'IL', '616': 'IL', '617': 'IL', '618': 'IL', '619': 'IL',
    '620': 'IL', '622': 'IL', '623': 'IL', '624': 'IL', '625': 'IL', '626': 'IL', '627': 'IL', '628': 'IL', '629': 'IL',
    '630': 'MO', '631': 'MO', '633': 'MO', '634': 'MO', '635': 'MO', '636': 'MO', '637': 'MO', '638': 'MO', '639': 'MO',
    '640': 'MO', '641': 'MO', '644': 'MO', '645': 'MO', '646': 'MO', '647': 'MO', '648': 'MO', '649': 'MO',
    '650': 'MO', '651': 'MO', '652': 'MO', '653': 'MO', '654': 'MO', '655': 'MO', '656': 'MO', '657': 'MO', '658': 'MO',
    '660': 'KS', '661': 'KS', '662': 'KS', '664': 'KS', '665': 'KS', '666': 'KS', '667': 'KS', '668': 'KS', '669': 'KS',
    '670': 'KS', '671': 'KS', '672': 'KS', '673': 'KS', '674': 'KS', '675': 'KS', '676': 'KS', '677': 'KS', '678': 'KS', '679': 'KS',
    '680': 'NE', '681': 'NE', '683': 'NE', '684': 'NE', '685': 'NE', '686': 'NE', '687': 'NE', '688': 'NE', '689': 'NE',
    '690': 'NE', '691': 'NE', '692': 'NE', '693': 'NE',
    '700': 'LA', '701': 'LA', '703': 'LA', '704': 'LA', '705': 'LA', '706': 'LA', '707': 'LA', '708': 'LA',
    '710': 'LA', '711': 'LA', '712': 'LA', '713': 'LA', '714': 'LA',
    '716': 'AR', '717': 'AR', '718': 'AR', '719': 'AR', '720': 'AR', '721': 'AR', '722': 'AR', '723': 'AR', '724': 'AR', '725': 'AR', '726': 'AR', '727': 'AR', '728': 'AR', '729': 'AR',
    '730': 'OK', '731': 'OK', '733': 'OK', '734': 'OK', '735': 'OK', '736': 'OK', '737': 'OK', '738': 'OK', '739': 'OK',
    '740': 'OK', '741': 'OK', '743': 'OK', '744': 'OK', '745': 'OK', '746': 'OK', '747': 'OK', '748': 'OK', '749': 'OK',
    '750': 'TX', '751': 'TX', '752': 'TX', '753': 'TX', '754': 'TX', '755': 'TX', '756': 'TX', '757': 'TX', '758': 'TX', '759': 'TX',
    '760': 'TX', '761': 'TX', '762': 'TX', '763': 'TX', '764': 'TX', '765': 'TX', '766': 'TX', '767': 'TX', '768': 'TX', '769': 'TX',
    '770': 'TX', '771': 'TX', '772': 'TX', '773': 'TX', '774': 'TX', '775': 'TX', '776': 'TX', '777': 'TX', '778': 'TX', '779': 'TX',
    '780': 'TX', '781': 'TX', '782': 'TX', '783': 'TX', '784': 'TX', '785': 'TX', '786': 'TX', '787': 'TX', '788': 'TX', '789': 'TX',
    '790': 'TX', '791': 'TX', '792': 'TX', '793': 'TX', '794': 'TX', '795': 'TX', '796': 'TX', '797': 'TX', '798': 'TX', '799': 'TX',
    '800': 'CO', '801': 'CO', '802': 'CO', '803': 'CO', '804': 'CO', '805': 'CO', '806': 'CO', '807': 'CO', '808': 'CO', '809': 'CO',
    '810': 'CO', '811': 'CO', '812': 'CO', '813': 'CO', '814': 'CO', '815': 'CO', '816': 'CO',
    '820': 'WY', '821': 'WY', '822': 'WY', '823': 'WY', '824': 'WY', '825': 'WY', '826': 'WY', '827': 'WY', '828': 'WY', '829': 'WY', '830': 'WY', '831': 'WY',
    '832': 'ID', '833': 'ID', '834': 'ID', '835': 'ID', '836': 'ID', '837': 'ID', '838': 'ID',
    '840': 'UT', '841': 'UT', '842': 'UT', '843': 'UT', '844': 'UT', '845': 'UT', '846': 'UT', '847': 'UT',
    '850': 'AZ', '851': 'AZ', '852': 'AZ', '853': 'AZ', '855': 'AZ', '856': 'AZ', '857': 'AZ', '859': 'AZ', '860': 'AZ', '863': 'AZ', '864': 'AZ', '865': 'AZ',
    '870': 'NM', '871': 'NM', '872': 'NM', '873': 'NM', '874': 'NM', '875': 'NM', '876': 'NM', '877': 'NM', '878': 'NM', '879': 'NM', '880': 'NM', '881': 'NM', '882': 'NM', '883': 'NM', '884': 'NM',
    '885': 'TX',
    '890': 'NV', '891': 'NV', '893': 'NV', '894': 'NV', '895': 'NV', '897': 'NV', '898': 'NV',
    '900': 'CA', '901': 'CA', '902': 'CA', '903': 'CA', '904': 'CA', '905': 'CA', '906': 'CA', '907': 'CA', '908': 'CA', '909': 'CA',
    '910': 'CA', '911': 'CA', '912': 'CA', '913': 'CA', '914': 'CA', '915': 'CA', '916': 'CA', '917': 'CA', '918': 'CA', '919': 'CA',
    '920': 'CA', '921': 'CA', '922': 'CA', '923': 'CA', '924': 'CA', '925': 'CA', '926': 'CA', '927': 'CA', '928': 'CA',
    '930': 'CA', '931': 'CA', '932': 'CA', '933': 'CA', '934': 'CA', '935': 'CA', '936': 'CA', '937': 'CA', '938': 'CA', '939': 'CA',
    '940': 'CA', '941': 'CA', '942': 'CA', '943': 'CA', '944': 'CA', '945': 'CA', '946': 'CA', '947': 'CA', '948': 'CA', '949': 'CA',
    '950': 'CA', '951': 'CA', '952': 'CA', '953': 'CA', '954': 'CA', '955': 'CA', '956': 'CA', '957': 'CA', '958': 'CA', '959': 'CA',
    '960': 'CA', '961': 'CA',
    '970': 'OR', '971': 'OR', '972': 'OR', '973': 'OR', '974': 'OR', '975': 'OR', '976': 'OR', '977': 'OR', '978': 'OR', '979': 'OR',
    '980': 'WA', '981': 'WA', '982': 'WA', '983': 'WA', '984': 'WA', '985': 'WA', '986': 'WA', '988': 'WA', '989': 'WA', '990': 'WA', '991': 'WA', '992': 'WA', '993': 'WA', '994': 'WA',
    '995': 'AK', '996': 'AK', '997': 'AK', '998': 'AK', '999': 'AK',
    '967': 'HI', '968': 'HI'
}

def infer_state_from_zip(zip_code, country):
    """For US addresses, infer state code from zip prefix if state is missing."""
    if country != 'US':
        return ''
    zip_clean = re.sub(r'\D', '', zip_code)[:5]
    if len(zip_clean) >= 3:
        return ZIP_PREFIX_TO_STATE.get(zip_clean[:3], '')
    return ''

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

        # 6. If state still blank, try to infer from zip code (US/ambiguous only)
        if not parsed['State'] and parsed.get('Zip'):
            zip_clean = re.sub(r'\D', '', parsed['Zip'])[:5]
            # 5-digit zip with no explicit Canadian postal: try US lookup
            if len(zip_clean) == 5:
                inferred_state = ZIP_PREFIX_TO_STATE.get(zip_clean[:3], '')
                if inferred_state:
                    logging.info(f"Inferred state '{inferred_state}' from zip '{parsed['Zip']}' - setting Country to US")
                    parsed['State'] = inferred_state
                    parsed['Country'] = 'US'  # Override default CA since zip is clearly American

        return parsed
    except Exception as e:
        logging.error(f"Error parsing address block: {e}")
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
