import pandas as pd
import os
import re

# ==========================================
# 1. تنظیمات فایل و مسیر
# ==========================================
input_dir = r'G:\Paper\nema-Nanopore-Sequencing\zoology new zealand and australia\data\Suspect'
output_path = os.path.join(input_dir, 'Location_Counts_Summary_With_Details.xlsx')

files_info = [
    '5.8S.xlsx',
    'ITS2.xlsx',
    'ITS1.xlsx',
    '18S.xlsx',
    '28S.xlsx',
    'COX1.xlsx',
]

# ==========================================
# 2. دیکشنری‌ها (برای تشخیص دقیق باید پر باشند)
# ==========================================
  
Australia = {
    "New South Wales": [
        "Aarons Pass", "Abbeyvale", "Abbotsbury", "Abbotsford", "Aberdare", "Aberdeen", "Aberfeldy", "Aberfoyle",
        "Adaminaby", "Adelong", "Adjungbilly", "Afterlee", "Albert Parish", "Albury", "Albury-Wodonga",
        "Albury–Wodonga (Albury part)", "Alstonville", "Armidale", "Ballina", "Balranald", "Batemans Bay", "Bathurst",
        "Bega", "Binnaway", "Blackheath", "Blayney", "Blue Mountains", "Bomaderry", "Booligal", "Boorowa", "Bourke",
        "Bowral", "Bowral – Mittagong", "Bredbo", "Broken Hill", "Byron Bay", "Camden", "Camden Haven", "Campbelltown",
        "Canberra – Queanbeyan (Queanbeyan part)", "Captains Flat", "Casino", "Central Coast", "Cessnock", "Cobar",
        "Cobargo", "Coffs Harbour", "Coleambally", "Collarenebri", "Cooma", "Coonabarabran", "Cooranbong",
        "Cootamundra", "Corowa", "Cowra", "Deniliquin", "Dubbo", "Eden", "Estella", "Eurobodalla", "Forbes", "Forster",
        "Forster – Tuncurry", "Gilgandra", "Glen Innes", "Gold Coast – Tweed Heads (part in NSW)", "Gosford",
        "Goulburn", "Grafton", "Griffith", "Gundagai", "Gunnedah", "Harden", "Hay", "Helensburgh", "Inverell",
        "Jerilderie", "Junee", "Katoomba", "Kempsey", "Kiama", "Kurri Kurri", "Kyogle", "Lake Cargelligo", "Leeton",
        "Lennox Head", "Lismore", "Lithgow", "Liverpool", "Lockhart", "Lord Howe Island", "Maitland", "Medowie",
        "Merimbula", "Merriwa", "Mittagong", "Moama", "Molong", "Moree", "Morisset", "Morisset – Cooranbong",
        "Moruya", "Moss Vale", "Mudgee", "Mulwala", "Murwillumbah", "Muswellbrook", "Nambucca Heads", "Narrabri",
        "Narrandera", "Nelson Bay", "Newcastle", "Newcastle – Maitland", "Nowra", "Nowra – Bomaderry",
        "Nowra-Bomaderry", "Nyngan", "Oberon", "Orange", "Pambula", "Parkes", "Parramatta", "Penrith", "Picton",
        "Port Macquarie", "Pottsville", "Queanbeyan", "Quirindi", "Raymond Terrace", "Richmond", "Richmond North",
        "Salamander Bay", "Sanctuary Point", "Sawtell", "Scone", "Silverdale", "Singleton", "Sofala",
        "Soldiers Point", "South West Rocks", "St Georges Basin", "St Georges Basin – Sanctuary Point", "Sydney",
        "Tahmoor", "Talbingo", "Tamworth", "Taree", "Temora", "Tenterfield", "The Rock", "Tocumwal", "Tumut",
        "Tuncurry", "Tweed Heads", "Ulladulla", "Uralla", "Urunga", "Wagga Wagga", "Walcha", "Walgett",
        "Warragamba", "Warren", "Wauchope", "Wee Waa", "Wellington", "Wentworth", "West Wyalong", "Wilcannia",
        "Windsor", "Wingham", "Wollongong", "Woolgoolga", "Wyong", "Yamba", "Yass", "Young",
        "Helensburgh", "Casino", "Alstonville", "Lennox Head", "Pottsville", "Tweed Heads", "Woolgoolga", "Sawtell",
        "South West Rocks", "Salamander Bay", "Soldiers Point", "Tahmoor", "Silverdale", "Sydney", "Newcastle", "Central Coast", "Wollongong", "Maitland", "Tweed Heads", "Albury", "Coffs Harbour", "Wagga Wagga", "Port Macquarie", "Orange", "Dubbo", "Queanbeyan", "Bathurst", "Tamworth", "Nowra", "Bomaderry", "Blue Mountains", "Lismore", "Goulburn", "Bowral", "Mittagong", "Cessnock", "Morisset", "Cooranbong", "Armidale", "Griffith", "Forster", "Tuncurry", "Kurri Kurri", "Ballina", "Taree", "Broken Hill", "Grafton", "Kiama", "Nelson Bay", "Ulladulla", "Singleton", "Raymond Terrace", "Batemans Bay", "Mudgee", "Lithgow", "Kempsey", "St Georges Basin", "Sanctuary Point", "Muswellbrook", "Byron Bay", "Medowie", "Casino", "Parkes", "Murwillumbah", "Inverell", "Moss Vale", "Gunnedah", "Cowra", "Merimbula", "Camden Haven", "Wauchope", "Young", "Lennox Head", "Leeton", "Pottsville", "Moree", "Forbes", "Estella", "Nambucca Heads", "Moama", "Tumut", "Cooma", "Deniliquin", "Yamba", "Helensburgh", "Yass", "Woolgoolga", "Cootamundra", "Salamander Bay", "Soldiers Point", "Narrabri", "Richmond North", "South West Rocks", "Silverdale", "Warragamba", "Glen Innes", "Alstonville", "Tahmoor", "Scone","Canberra – Queanbeyan (ACT part)", "Canberra", "Hall", "Oaks Estate", "Pierces Creek", "Barton", "Deakin", "Deakin West", "Harman", "Harrison",
        "Hmas Creswell", "Jervis Bay", "Dalgety", "Yass-Canberra", "Tooma", "Lyndhurst", "Armidale", "Farrer", "Downer", "Florey", "Oaks Estate", "Canberra", "Franklin", "Kambah", "Deakin", "Dalgety",
        "Barton", "Aranda", "Deakin West", "Bonner", "Dickson", "Braddon", "Casey", "Chifley", "Ngunnawal",
        "Belconnen", "Crace", "Calwell", "Campbell", "Hall", "Chapman", "Coombs", "Conder", "Lyndhurst", "Forde",
        "Jervis Bay", "Fadden", "Curtin", "Charnwood", "City", "Evatt", "Flynn", "Hmas Creswell", "Banks",
        "Canberra – Queanbeyan (ACT part)", "Cook", "Amaroo", "Pierces Creek", "Harrison", "Duffy", "Fisher",
        "Armidale", "Dunlop", "Harman", "Chisholm", "Tooma", "Yass-Canberra", "Bruce", "Forrest", "Mount Ainslie"
    ],
    "Victoria": [
        "A1 Mine Settlement", "Abbeyard", "Aberfeldy", "Acheron", "Adams Estate", "Addington", "Adelaide Lead",
        "Agnes", "Aireys Inlet", "Albury-Wodonga", "Albury–Wodonga (Wodonga part)", "Allambee", "Alphington",
        "Ancona", "Appin", "Ararat", "Archies Creek", "Ashbourne", "Avonsleigh", "Axedale", "Bacchus Marsh",
        "Bagshot", "Bairnsdale", "Ballan", "Ballarat", "Balmattum", "Balmoral", "Balnarring", "Bandiana",
        "Bangholme", "Bannockburn", "Barfold", "Barkers Creek", "Barnawartha", "Barrabool", "Barrys Reef",
        "Barwon Heads", "Basalt", "Batesford", "Bayindeen", "Bayles", "Baynton", "Bayswater", "Beechworth",
        "Benalla", "Bendigo", "Bright", "Castlemaine", "Churchill", "Clematis", "Clifton Springs", "Cobram", "Colac",
        "Cowes", "Dandenong", "Drouin", "Drysdale", "Echuca", "Echuca – Moama", "Emerald", "Forrest", "Frankston",
        "Geelong", "Gisborne", "Glengarry", "Hamilton", "Healesville", "Horsham", "Inverloch", "Jan Juc", "Kerang",
        "Kilmore", "Kyabram", "Kyneton", "Lakes Entrance", "Lara", "Latrobe City", "Leongatha", "Leopold",
        "Macedon", "Mahaikah", "Maiden Gully", "Maiden Town", "Mansfield", "Maryborough", "Melbourne", "Melton",
        "Mildura", "Mildura – Buronga (part in Victoria)", "Moe", "Moe – Newborough", "Mooroopna", "Morwell",
        "Newborough", "Newhaven", "Ocean Grove", "Pakenham", "Port Fairy", "Portarlington", "Portland", "Sale",
        "Sea Lake", "Seddon", "Sedgwick", "Selby", "Seldom Seen", "Seymour", "Shepparton", "Shepparton – Mooroopna",
        "St Leonards", "Stawell", "Sunbury", "Swan Hill", "Torquay", "Traralgon", "Traralgon – Morwell", "Viewbank",
        "Vinifera", "Violet Town", "Waaia", "Waanyarra", "Wahgunyah", "Waitchie", "Walhalla", "Walkerville",
        "Wallace", "Wallan", "Wangaratta", "Warragul", "Warragul – Drouin", "Warrnambool", "Werribee", "Whittlesea",
        "Wodonga", "Wonthaggi", "Yarrawonga", "Abbotsford", "Aberfeldie", "Aberfeldy", "Adelaide Lead", "Aireys Inlet", "Airport West", "Alamein", "Albanvale", "Albert Park", "Alberton", "Alberton Shire", "Albion", "Alexandra and Alexandra Shire", "Alfredton", "Allambee", "Allans Flat", "Allansford", "Allendale", "Alma", "Alphington", "Alpine Shire", "Altona and Altona City", "Altona Meadows", "Altona North", "Alvie", "Amherst", "Amphitheatre", "Anakie", "Anglesea", "Annuello", "Antwerp", "Apollo Bay", "Apsley", "Arapiles Shire", "Ararat", "Ararat Rural City", "Arcadia", "Archies Creek", "Ardeer", "Ardmona", "Armadale", "Armstrong", "Arthurs Creek", "Arthurs Seat", "Ascot (near Bendigo)", "Ascot Vale", "Ashburton", "Ashwood", "Aspendale", "Aspendale Gardens", "Attwood", "Auburn", "Avalon", "Avenel", "Avoca and Avoca Shire", "Avon Plains", "Avon Shire", "Avondale Heights", "Avonsleigh", "Axe Creek, Emu Creek and Eppalock", "Axedale", "Bacchus Marsh and Shire", "Baddaginnie", "Badger Creek", "Baillieston", "Bairnsdale", "Bairnsdale Shire", "Balaclava", "Ballan", "Ballan Shire", "Ballarat City", "Ballarat East", "Ballarat North", "Ballarat Shire", "Ballendella", "Balliang", "Balmattum", "Balmoral", "Balnarring", "Balwyn", "Balwyn North", "Bamawm and Bamawm Extension", "Bandiana", "Bangholme", "Bannockburn", "Bannockburn Shire", "Banyena", "Banyule City", "Baranduda", "Barengi Gadjin", "Baringhup", "Barkly", "Barmah", "Barnawartha", "Barongarook", "Barrabool and Barrabool Shire", "Barramunga", "Barraport", "Barry's Reef", "Barwon Downs", "Barwon Heads", "Bass", "Bass Coast Shire", "Bass Shire", "Batesford", "Baw Baw Shire", "Baxter", "Bayles", "Baynton", "Bayside City", "Bayswater", "Bayswater North", "Beaconsfield", "Beaconsfield Upper", "Bealiba", "Beaufort", "Beaumaris", "Beeac", "Beech Forest", "Beechworth", "Beechworth Shire", "Belfast", "Belfast Shire", "Belgrave", "Belgrave Heights and Belgrave South", "Bell Park", "Bell Post Hill", "Bellarine", "Bellarine Rural City", "Bellbridge", "Bellfield (near Heidelberg)", "Belmont", "Bena", "Benalla", "Benambra", "Bendigo", "Bendoc", "Bennettswood", "Bentleigh and Bentleigh East", "Berringa", "Berriwillock", "Berwick", "Berwick Shire and City", "Bessiebelle", "Bet Bet and Shire", "Bethanga", "Betley", "Beulah", "Beverford", "Beveridge", "Big River", "Binginwarri", "Birchip and Birchip Shire", "Birdwoodton and Cabarita", "Birregurra", "Bittern", "Black Hill", "Black Lead and Scotchmans Lead", "Black Rock", "Blackburn", "Blackburn and Mitcham Shire", "Blackburn North", "Blackburn South", "Blackwood", "Blairgowrie", "Blakeville and Korweinguboora", "Blind Bight", "Bogong", "Boho and Boho South", "Boinka", "Boisdale", "Bolwarra and Allestree", "Bolwarrah", "Bonbeach", "Bonegilla", "Boneo, Cape Schanck, Fingal", "Bonnie Doon", "Boolarra", "Boorcan", "Boorhaman", "Boort", "Boronia", "Boroondara", "Borung", "Box Hill and Box Hill City", "Box Hill North", "Box Hill South", "Braeside and Waterways", "Branxholme", "Braybrook and Braybrook Shire", "Breakwater", "Breamlea", "Briagolong", "Briar Hill", "Bridgewater", "Bright", "Bright Shire", "Brighton", "Brighton East", "Brim", "Brimbank City", "Britannia Creek", "Broadford and Broadford Shire", "Broadmeadows and Broadmeadows City", "Bromley", "Brookfield", "Brooklyn", "Broomfield", "Broughton", "Brown Hill", "Browns and Scarsdale", "Brunswick and Brunswick City", "Brunswick East", "Brunswick West", "Bruthen", "Buangor", "Buchan", "Buckrabanyule", "Buckley", "Budgeree", "Buffalo River", "Bulla and Bulla Shire", "Bullarook", "Bullarto", "Bulleen", "Bullengarook", "Bullumwaal", "Buln Buln", "Buln Buln Shire", "Buloke Shire", "Bundalaguah", "Bundalong", "Bundoora", "Bungaree and Bungaree Shire", "Buninyong", "Buninyong Shire", "Bunyip", "Burnley", "Burnside and Burnside Heights", "Burramine", "Burrumbeet", "Burwood and Burwood East", "Bushfield", "Buxton", "Byaduk", "Bylands", "Cabbage Tree Creek", "Cairnlea", "Caldermeade", "California Gully", "Calivil", "Callignee", "Camberwell", "Cambrian Hill", "Campaspe Shire", "Campbellfield", "Campbells Creek", "Camperdown", "Canadian", "Caniambo", "Cann River", "Cannons Creek", "Canterbury", "Cape Clear", "Cape Paterson", "Cape Woolamai", "Caramut", "Carapooee", "Cardigan and Cardigan Village", "Cardinia", "Cardinia Shire", "Cardross", "Carisbrook", "Carlisle River", "Carlsruhe", "Carlton", "Carlton North", "Carnegie", "Carngham", "Caroline Springs", "Carrajung", "Carrum", "Carrum Downs", "Carrum Swamp", "Carwarp", "Casey City", "Cassilis and Tongio West", "Casterton", "Castle Donnington", "Castlemaine", "Catani", "Cathcart", "Caulfield", "Cavendish", "Central Goldfields Shire", "Ceres", "Chadstone", "Charlton and Charlton Shire", "Chelsea", "Chelsea Heights", "Cheltenham", "Cheshunt", "Chewton", "Chiltern and Chiltern Shire", "Chiltern Valley", "Chinkapook", "Chirnside Park", "Christmas Hills", "Christmastown", "Chum Creek", "Churchill", "Clarendon", "Clarinda", "Clarkefield", "Clayton and Clayton North", "Clayton South", "Clematis", "Clementson", "Clifton Hill", "Clifton Springs", "Clunes", "Clyde and Clyde North", "Coalville", "Cobden", "Cobram and Cobram Shire", "Coburg and Coburg City", "Coburg North", "Cockatoo", "Cocoroc", "Cohuna and Cohuna Shire", "Coimadai", "Colac", "Colac Otway Shire", "Colac Shire", "Colbinabbin", "Coldstream", "Coleraine", "Collingwood", "Condah", "Congupna", "Connewarre", "Coode Island", "Coolaroo", "Coongulla", "Cora Lynn", "Corack", "Coragulac", "Corangamite Shire", "Corindhap", "Corinella", "Corio", "Corio Shire", "Corop", "Cororooke", "Corryong", "Costerfield", "Cowangie", "Cowwarr", "Craigie", "Craigieburn", "Cranbourne", "Cranbourne North, South, East, West", "Cranbourne Shire", "Cremorne", "Cressy", "Creswick", "Creswick Shire", "Crib Point", "Crossley", "Croxton", "Croydon and Croydon City", "Croydon Hills", "Croydon North", "Croydon South", "Cudgee", "Cudgewa", "Culgoa", "Dallas", "Dalmore", "Dalyston", "Dandenong", "Dandenong North", "Dandenong Ranges", "Dandenong Shire and City", "Dandenong South", "Darebin and Darebin City", "Dargo", "Darley", "Darlington", "Darnum", "Darraweit Guim", "Dartmoor", "Dartmouth", "Daylesford", "Daylesford and Glenlyon Shire", "Deakin Shire", "Dean", "Deans Marsh", "Dederang", "Deep Lead", "Deepdene", "Deer Park", "Delacombe", "Delahey", "Delatite Shire", "Derrimut", "Derrinallum", "Devenish", "Devon Meadows", "Devon North", "Dhurringile", "Diamond Creek", "Diamond Hill", "Diamond Valley Shire", "Digby", "Diggers Rest and Plumpton", "Diggora and Diggora West", "Dimboola", "Dimboola Shire", "Dingee", "Dingley Village", "Dinner Plain", "Dixons Creek", "Dja Dja Wurrung", "Docklands", "Donald and Donald Shire", "Doncaster", "Doncaster and Templestowe", "Doncaster East", "Donnybrook and Kalkallo", "Donvale", "Dooen", "Dookie", "Doreen", "Doveton", "Dreeite", "Dromana", "Drouin", "Drouin West and Drouin South", "Drumborg", "Drummond", "Drysdale", "Dudley and Dudley South", "Dumbalk", "Dundas Shire", "Dunkeld", "Dunmunkle Shire", "Dunnstown", "Dunolly", "Durham Lead", "Eagle Point", "Eaglehawk", "Eaglemont", "East Bendigo", "East Geelong", "East Gippsland", "East Loddon Shire", "East Melbourne", "Eastern Maar", "Echuca", "Echuca Shire", "Echuca Village", "Ecklin", "Eddington", "Eden Park", "Edenhope", "Edi and Edi Upper", "Edithvale", "Eganstown", "Eildon", "Elaine", "Eldorado", "Elingamite", "Ellerslie", "Elliminyt", "Ellinbank", "Elmhurst", "Elmore", "Elphinstone", "Elsternwick", "Eltham and Eltham Shire", "Eltham North", "Elwood", "Emerald", "Emerald Hill", "Endeavour Hills", "Enfield", "Ensay", "Epping", "Epsom", "Erica", "Eskdale", "Essendon and Essendon City", "Essendon North", "Essendon West", "Eumemmerring", "Eureka", "Euroa and Euroa Shire", "Eurobin", "Evansford", "Everton", "Eynesbury", "Fairfield", "Falls Creek", "Fawkner", "Fernshaw", "Ferntree Gully", "Ferny Creek", "Fish Creek", "Fishermans Bend", "Fitzroy", "Fitzroy North", "Flemington", "Flinders", "Flinders Shire", "Flora Hill", "Flowerdale and Hazeldene", "Footscray and Footscray City", "Footscray West", "Forest Hill", "Forrest", "Foster", "Fosterville", "Framlingham", "Frankston", "Frankston City", "Frankston North", "Frankston South", "Freeburgh", "French Island", "Freshwater Creek", "Fryerstown", "Fyansford", "Gaffneys Creek", "Gannawarra", "Gannawarra Shire", "Garden City", "Gardenvale", "Gardiner", "Garfield", "Garvoc", "Geelong", "Geelong North", "Geelong West", "Gellibrand", "Gembrook", "Gerang Gerung", "Gheringhap", "Gippsland", "Gippsland Lakes", "Girgarre", "Gisborne", "Gladstone Park", "Gladysdale", "Glen Alvie", "Glen Eira City", "Glen Huntly", "Glen Iris", "Glen Waverley", "Glenelg Shire", "Glenferrie", "Glengarry", "Glenlyon and Glenlyon Shire", "Glenmaggie District", "Glenorchy", "Glenormiston", "Glenrowan", "Glenroy", "Glenthompson", "Gobur", "Golden Beach", "Golden Lake", "Golden Plains Shire", "Golden Point", "Golden Square", "Goldsborough", "Goonawarra", "Goorambat", "Goornong", "Gorae and Gorae West", "Gordon", "Gordon Shire", "Gormandale", "Goroke", "Goulburn Shire", "Goulburn Valley", "Gowanbrae", "Gowar", "Gowerville", "Grahamvale", "Grampians", "Granite Flat", "Grantville", "Granya", "Grassmere", "Graytown", "Great Northern", "Great Southern", "Great Western", "Greater Bendigo", "Greater Dandenong City", "Greater Geelong City", "Greater Shepparton", "Greendale", "Greensborough", "Greenvale", "Grenville Shire", "Greta", "Grovedale", "Gruyere", "Guildford", "Gunaikurnai", "Gunbower", "Gunditj Mirring", "Gundowring and Gundowring Upper", "Haddon", "Hadfield", "Hallam", "Hallora and Ripplebrook", "Halls Gap", "Hamilton", "Hamlyn Heights", "Hampden Shire", "Hampton", "Hampton Park", "Happy Valley", "Harcourt", "Harkaway", "Harrietville", "Harrisfield", "Harrow", "Hartwell", "Hastings", "Havelock", "Haven", "Hawkesdale", "Hawksburn", "Hawthorn", "Hawthorn East", "Hazelwood", "Healesville and Healesville Shire", "Heathcote", "Heatherdale", "Heatherton", "Hedley", "Heidelberg", "Heidelberg West and Heidelberg Heights", "Henty", "Hepburn and Hepburn Springs", "Hepburn Shire", "Herne Hill", "Hernes Oak", "Hexham", "Heyfield", "Heytesbury Shire", "Heywood", "Heywood Shire", "Highett", "Highton", "Hill End", "Hillside", "Hillside (near Bairnsdale)", "HMAS Cerberus", "Hobsons Bay City", "Hoddles Creek", "Hollands Landing", "Hollybush", "Holmesglen", "Homebush", "Hopetoun", "Hoppers Crossing", "Horsham", "Horsham Rural City", "Hotham", "Howqua Shire", "Hughesdale", "Hume City", "Huntingdale", "Huntly", "Huntly Shire", "Hurstbridge", "Illabarook", "Illowa", "Indented Head", "Indigo", "Indigo Shire", "Inglewood", "Invergordon", "Inverleigh", "Inverloch", "Invermay and Invermay Park", "Iona", "Ironbark", "Irrewillipe", "Irymple", "Italian Gully", "Ivanhoe", "Ivanhoe East", "Jacana", "Jamieson", "Jan Juc and Bellbrae", "Jancourt", "Janefield", "Jeeralang", "Jeparit", "Jericho", "Jindivick", "Johnsonville", "Jolimont", "Jordanville", "Jumbunna", "Jung", "Junortoun", "Kalimna", "Kalkee", "Kallista", "Kalorama", "Kamarooka", "Kangaroo Flat", "Kangaroo Ground", "Kaniva", "Kara Kara Shire", "Kardella", "Karingal", "Karkarooc Shire", "Katamatite", "Katandra", "Katandra West", "Katunga", "Katyil", "Kealba", "Keilor", "Keilor Downs", "Keilor East", "Keilor Lodge", "Keilor Park", "Kennington", "Kensington", "Keon Park", "Kerang", "Kerang Shire", "Kergunyah and Kergunyah South", "Kerrimuir", "Kew", "Kew East", "Kewell", "Keysborough", "Kialla and Kialla West", "Kiata", "Kiewa", "Kilcunda", "Killarney", "Kilmore", "Kilsyth and Kilsyth South", "Kinglake District", "Kingower", "Kings Park", "Kingsbury", "Kingston City", "Kingston Township", "Melbourne", "Geelong", "Ballarat", "Bendigo", "Melton (Melbourne)", "Shepparton", "Mooroopna", "Pakenham (Melbourne)", "Sunbury (Melbourne)", "Wodonga", "Mildura", "Warrnambool", "Traralgon", "Torquay", "Jan Juc", "Ocean Grove", "Barwon Heads", "Bacchus Marsh", "Wangaratta", "Warragul", "Horsham", "Drysdale", "Clifton Springs", "Lara", "Moe", "Newborough", "Drouin", "Wallan", "Sale", "Morwell", "Echuca", "Bairnsdale", "Colac", "Leopold", "Gisborne", "Swan Hill", "Castlemaine", "Portland", "Benalla", "Hamilton", "Portarlington", "St Leonards", "Kilmore", "Healesville", "Yarrawonga", "Wonthaggi", "Maryborough", "Ararat", "Cowes", "Lakes Entrance", "Bannockburn", "Inverloch", "Seymour", "Kyabram", "Stawell", "Leongatha", "Whittlesea", "Cobram", "Kyneton", "Churchill"
    ],
    "Queensland": [
        "Abbeywood", "Abbotsford", "Abercorn", "Abergowrie", "Abingdon Downs", "Abington", "Acacia Ridge", "Acland",
        "Adavale", "Agnes Water", "Airlie Beach", "Airlie Beach – Cannonvale", "Aitkenvale", "Albany Creek",
        "Albert Shire", "Albion", "Aldershot", "Alice River", "Allingham", "Allora", "Alpha", "Amamoor",
        "Apple Tree Creek", "Aramac", "Armstrong Beach", "Atherton", "Augathella", "Aurukun", "Ayr", "Babinda",
        "Badu Island", "Bakers Creek", "Balgal Beach", "Ball Bay", "Balmoral", "Balonne Shire", "Bamaga", "Banana Shire", "Banyo", "Baralaba", "Barcaldine", "Barcaldine Regional Council", "Barcoo Shire", "Bardon", "Bargara", "Barkly Tableland", "Barolin Shire", "Barron Shire", "Basin Pocket", "Battery Hill", "Bauhinia Shire", "Beachmere", "Beaconsfield", "Beaudesert", "Beaudesert Shire", "Beechmont and Lower Beechmont", "Beenleigh", "Beerburrum", "Beerwah", "Belgian Gardens", "Bell", "Bellbird Park", "Bellbowrie", "Bellmere", "Belmont and Belmont Shire", "Belyando Shire", "Bendemere Shire", "Benowa", "Bentley Park", "Bethania", "Biggenden and Biggenden Shire", "Biggera Waters", "Bilinga", "Biloela", "Bingera and South Bingera", "Bingil Bay", "Birkdale", "Birtinya", "Blackall and Blackall Shire", "Blackall-Tambo Regional Council", "Blackbutt", "Blacks Beach", "Blackstone and Bundamba", "Blackwater", "Blair Athol", "Blenheim", "Bli Bli", "Bluff", "Bohle Plains", "Bonogin", "Boooie", "Boonah", "Boonah Shire", "Boondall", "Booral", "Booringa Shire", "Booroobin", "Booval", "Boronia Heights", "Bouldercombe", "Boulia and Boulia Shire", "Bowen", "Bowen Basin", "Bowen Hills and Mayne", "Bowen Shire", "Boyne Island", "Bracken Ridge", "Bridgeman Downs", "Brigalow Belt", "Brighton", "Brisbane and Greater Brisbane", "Brisbane Central", "Brisbane River System", "Broadbeach", "Broadbeach Waters", "Broadsound Shire", "Brookfield and Upper Brookfield", "Brookwater", "Browns Plains", "Bucasia", "Bucca", "Buccan", "Buddina", "Buderim", "Buderim Hinterland Localities", "Bulimba", "Bulloo Shire", "Bundaberg", "Bundaberg Eastern Localities", "Bundaberg Regional Council", "Bundaberg Suburbs", "Bundall", "Bungalow and Portsmith", "Bungil Shire", "Bunya", "Buranda", "Burbank", "Burdekin Shire", "Burdell", "Burke Shire", "Burleigh Heads", "Burleigh Waters", "Burnett Heads", "Burnett Shire", "Burnside", "Burpengary", "Burrum Heads", "Burrum Shire", "Bushland Beach", "Cabarlah", "Caboolture", "Caboolture Shire", "Cairns", "Cairns Regional Council", "Cairns Suburbs", "Calamvale", "Calliope", "Calliope Shire", "Calliungal Shire", "Caloundra", "Caloundra Suburbs", "Cambooya", "Cambooya Shire", "Camira", "Camp Hill", "Camp Mountain", "Cannon Hill", "Cannonvale", "Canoona", "Canungra", "Capalaba", "Cape York Peninsula", "Capella", "Caravonica", "Carbrook", "Cardwell", "Cardwell Shire", "Carina", "Carina Heights", "Carindale", "Carole Park", "Carpentaria Shire", "Carrara and Merrimac", "Carseldine", "Cashmere", "Cassowary Coast Regional Council", "Cawarral", "Cecil Plains", "Cedar Creek (near Tamborine)", "Cedar Grove and Cedar Vale", "Centenary Heights", "Central Highlands Regional Council", "Chambers Flat", "Chandler", "Channel Country, Queensland", "Chapel Hill", "Charleville", "Charters Towers", "Charters Towers Regional Council", "Charters Towers Suburbs", "Chatsworth, Glastonbury and The Palms", "Chelmer", "Cherbourg Aboriginal Shire Council", "Chermside and Chermside West", "Childers", "Chillagoe and Chillagoe Shire", "Chinchilla", "Chinchilla Shire", "Churchill", "Chuwar", "Clayfield", "Clear Island Waters", "Clear Mountain", "Clermont and Copperfield", "Cleveland and Cleveland Shire", "Clifton", "Clifton Beach", "Clifton Shire", "Cloncurry", "Cloncurry Shire", "Clontarf", "Coalfalls", "Coes Creek", "Collingwood Park", "Collinsville and Scottville", "Condon", "Conondale", "Cook Shire", "Cooktown", "Coolabunia", "Coolangatta", "Cooloola Cove", "Cooloola Shire", "Coolum Beach", "Coombabah", "Coomera and Coomera Shire", "Coominya", "Coopers Plains", "Cooroy", "Coorparoo and Coorparoo Shire", "Cooya Beach", "Coral Cove", "Corinda", "Cornubia", "Cotswold Hills and Torrington", "Cracow", "Craiglie", "Cranbrook", "Cranley", "Crestmead", "Cribb Island", "Crows Nest", "Crows Nest Shire", "Croydon and Croydon Shire", "Cunnamulla", "Curra", "Currajong", "Currimundi", "Currumbin", "Currumbin Valley", "Currumbin Waters", "Daintree Shire", "Daisy Hill", "Dakabin", "Dalby", "Dalrymple Shire", "Darling Downs", "Darling Heights", "Darra", "Dayboro", "Deagon", "Deception Bay", "Deebing Heights", "Deeral", "Depot Hill", "Diamantina Shire", "Dicky Beach", "Diddillibah", "Dimbulah", "Dinmore", "Dirranbandi", "Donnybrook", "Doolandella", "Doomadgee Aboriginal Shire Council", "Doonan", "Douglas", "Douglas Shire", "Drayton and Drayton Shire", "Drewvale", "Dry Tropics", "Duaringa and Duaringa Shire", "Duchess", "Dugandan", "Dundowran, Craignish and Dundowran Beach", "Durack", "Dutton Park", "Dysart", "D’Aguilar", "Eacham Shire", "Eagle Farm", "Eagle Heights", "Eagleby", "Earlville", "East Brisbane", "East Ipswich", "East Mackay", "East Toowoomba", "Eastern Heights", "Eatons Hill", "Ebbw Vale", "Edens Landing", "Edge Hill", "Edmonton", "Eerwah Vale", "Eidsvold and Eidsvold Shire", "Eight Mile Plains", "Eimeo", "Einasleigh", "Ekibin", "El Arish", "Elanora", "Elimbah", "Elliott Heads", "Emerald", "Emerald Shire", "Emu Park", "Enoggera", "Esk", "Esk Shire", "Etheridge", "Eton", "Eudlo", "Eumundi", "Everton Hills", "Everton Park", "Fairfield", "Fairney View and Glamorgan Vale", "Farleigh", "Fernvale", "Ferny Grove", "Ferny Hills", "Fig Tree Pocket", "Finch Hatton", "Fitzgibbon", "Fitzroy River Basin", "Fitzroy Shire", "Flaxton", "Flinders Shire", "Flinders View", "Flying Fish Point", "Forest Hill", "Forest Lake", "Forestdale", "Forsayth", "Fortitude Valley", "Fraser Coast Regional Council", "Fraser Island", "Frenchville", "Freshwater", "Gailes", "Garbutt", "Gatton", "Gatton Shire", "Gaven", "Gayndah", "Gayndah Shire", "Gaythorne", "Geebung", "Georgetown", "Gin Gin", "Giru", "Gladstone", "Gladstone Localities", "Gladstone Regional Council", "Gladstone Suburbs", "Glass House Mountains", "Glencoe", "Glenden", "Gleneagle", "Glenella", "Glengallan Shire", "Glenore Grove", "Glenvale", "Glenview", "Gold Coast", "Gold Coast Inner Hinterland", "Golden Beach", "Golden Gate", "Gooburrum Shire", "Goodna", "Goomboorian", "Goombungee", "Goondi", "Goondiwindi", "Goondiwindi Regional Council", "Gordon Park", "Gordonvale", "Gowrie Junction", "Gowrie Shire", "Gracemere", "Graceville", "Grandchester", "Grange", "Granville", "Grasstree Beach", "Great Barrier Reef", "Greenbank", "Greenmount", "Greenslopes", "Griffin", "Grovely", "Gulf Country, Queensland", "Gulliver", "Gumdale", "Gunalda and Glenwood", "Gunpowder", "Gympie", "Gympie Regional Council", "Habana", "Half Tide Beach", "Halifax", "Hamilton", "Hamilton Island", "Hann Shire", "Harlaxton", "Harristown", "Harrisville and Normanby Shire", "Hatton Vale", "Hawthorne", "Hay Point", "Hayman Island", "Heatley", "Helensvale", "Helidon", "Hemmant", "Hendon", "Hendra", "Herberton", "Herberton Minerals Area", "Herberton Shire", "Heritage Park", "Hermit Park", "Herston", "Hervey Bay", "Hervey Bay Suburbs", "Highfields and Highfields Shire", "Highgate Hill", "Highland Park", "Highvale", "Hillcrest", "Hinchinbrook Shire", "Hodgkinson Minerals Area", "Holland Park and Holland Park West", "Holloways Beach", "Hollywell", "Holmview", "Home Hill", "Homebush", "Hope Island", "Hope Vale Aboriginal Shire Council", "Horn Island", "Horton", "Howard", "Hughenden", "Hyde Park", "Ilfracombe Shire", "Ilkley", "Imbil", "Inala", "Indooroopilly", "Ingham", "Inglewood", "Inglewood Shire", "Inkerman", "Innes Park", "Innisfail", "Innisfail Suburbs", "Ipswich", "Ipswich Southern Localities", "Irvinebank", "Isaac Regional Council", "Isis Shire", "Isisford and Isisford Shire", "Ithaca and Ithaca Shire", "Jacobs Well", "Jamboree Heights", "Jandowae", "Japoonvale", "Jarvisfield", "Jensen", "Jericho Shire", "Jimboomba, Flagstone and Stockleigh", "Jindalee", "Johnstone Shire", "Jondaryan", "Jondaryan Shire", "Joyner", "Julia Creek", "Kalbar", "Kalinga", "Kallangur", "Kandanga", "Kangaroo Point", "Karalee and Barellan Point", "Karana Downs", "Karumba", "Kawana (Rockhampton)", "Kawana Waters", "Kearneys Spring", "Kedron", "Kelso", "Kelvin Grove", "Kenilworth", "Kenmore", "Kenmore Hills", "Keperra", "Keppel Bay Area", "Kerry", "Kewarra Beach", "Kholo", "Kilcoy and Kilcoy Shire", "Kilkivan and Kilkivan Shire", "Killarney", "Kin Kin", "Kingaroy", "Kingaroy Shire", "Kings Beach", "Kingsthorpe", "Kingston", "Kinka Beach", "Kippa-Ring", "Kirwan", "Kolan Shire", "Kooralbyn", "Koumala", "Kowanyama Aboriginal Shire Council", "Kuluin", "Kuraby", "Kuranda", "Kureelpa", "Kuridala", "Kurrimine Beach", "Kurwongbah", "Kuttabul", "Labrador", "Laidley", "Laidley Localities", "Laidley Shire", "Lake Macdonald", "Lakes Creek and Koongal", "Landsborough", "Lawnton", "Leichhardt", "Lindum", "Livingstone Shire", "Lockhart River Aboriginal Shire Council", "Lockyer Valley and Regional Council", "Logan Central", "Logan City", "Logan Reserve", "Logan Village", "Loganholme", "Loganlea", "Longreach", "Longreach Regional Council", "Longreach Shire", "Lota", "Lower Burnett and Kolan Localities", "Lowood", "Lucinda", "Lutwyche", "Lytton", "MacGregor", "Machans Beach", "Mackay", "Mackay Regional Council", "Mackenzie", "Macknade", "Magnetic Island", "Main Beach", "Malanda", "Maleny", "Mango Hill", "Manly", "Manly West", "Mansfield", "Manunda", "Many Peaks", "Mapleton", "Mapoon Aboriginal Shire Council", "Maranoa Regional Council", "Marburg, Haigslea, Ironbark", "Marcoola", "Mareeba", "Mareeba Shire", "Margate", "Marian", "Maroochy and Maroochy Shire", "Maroochy River", "Maroochydore", "Marsden", "Mary Kathleen", "Mary Valley", "Maryborough", "Maryborough Suburbs", "McDowall", "McKinlay Shire", "Meadowbrook", "Mena Creek", "Meridan Plains", "Meringandan", "Mermaid Beach", "Mermaid Waters", "Merthyr", "Miami", "Middle Park", "Middle Ridge", "Middlemount", "Midge Point", "Miles", "Millaa Millaa", "Millmerran", "Millmerran Shire", "Milton", "Minden", "Minyama", "Mirani", "Mirani Shire", "Miriam Vale and Miriam Vale Shire", "Miriwinni", "Mission Beach", "Mitchell", "Mitchelton", "Moffat Beach", "Moggill", "Molendinar", "Monto", "Monto Shire", "Montville", "Mooloolaba", "Mooloolah", "Moore Park Beach", "Moorooka", "Moranbah", "Morayfield", "Moreton Bay", "Moreton Bay Regional Council", "Moreton District", "Moreton Shire", "Morningside", "Mornington Shire", "Mossman", "Mount Britton", "Mount Chalmers", "Mount Coolum", "Mount Coot-tha", "Mount Cotton", "Mount Crosby", "Mount Cuthbert", "Mount Garnet", "Mount Gravatt and Mount Gravatt East", "Mount Isa", "Mount Isa Shire and City", "Mount Isa Suburbs", "Mount Larcom", "Mount Lofty and Prince Henry Heights", "Mount Louisa", "Mount Low", "Mount Molloy", "Mount Morgan", "Mount Morgan Shire", "Mount Ommaney", "Mount Perry and Perry Shire", "Mount Pleasant", "Mount Samson", "Mount Tyson", "Mount Warren Park", "Mountain Creek", "Moura", "Mourilyan", "Mudgeeraba", "Mudjimba", "Mulgildie", "Mulgrave Shire", "Mundingburra", "Mundoolun", "Mundubbera and Mundubbera Shire", "Mungana", "Mungindi", "Munruben", "Murarrie", "Murgon and Murgon Shire", "Murilla Shire", "Murrumba Downs", "Murweh Shire", "Mutdapilly and Mutdapilly Shire", "Nambour", "Nanango", "Nanango Shire", "Napranum Aboriginal Shire Council", "Narangba", "Nathan", "Nebo Shire", "Nerang and Nerang Shire", "New Beith", "New Farm", "Newmarket", "Newport", "Newstead", "Newtown (near Ipswich)", "Newtown, Toowoomba", "Ninderry, North Arm and Yandina Creek", "Ningi", "Brisbane", "Gold Coast", "Sunshine Coast", "Townsville", "Cairns", "Toowoomba", "Mackay", "Rockhampton", "Hervey Bay", "Bundaberg", "Gladstone", "Maryborough", "Mount Isa", "Gympie", "Nambour", "Bongaree–Woorim", "Yeppoon", "Warwick", "Emerald", "Dalby", "Bargara–Innes Park", "Gracemere", "Kingaroy", "Tannum Sands–Boyne Island", "Highfields", "Airlie Beach", "Sandstone Point–Ningi", "Bowen", "Moranbah", "Ayr", "Charters Towers", "Mareeba", "Tamborine Mountain", "Innisfail", "Atherton", "Roma", "Gatton", "Gordonvale", "Chinchilla", "Beaudesert", "Biloela", "Mount Cotton", "Jimboomba – West", "Goondiwindi", "Beerwah", "Stanthorpe", "Kensington Grove–Regency Downs", "Blackwater", "Emu Park", "Ingham", "Oakey", "Port Douglas–Craiglie", "Jimboomba", "Palmwoods", "Yarrabilba", "Lowood", "Weipa", "Doonan–Tinbeerwah", "Landsborough", "Glass House Mountains", "Laidley", "Westbrook", "Beachmere", "Calliope", "Proserpine", "Nanango", "Sarina", "Walkerston", "Charleville", "Mooloolah", "Pittsworth", "Home Hill", "Thursday Island", "Woodford", "Cooroy", "Russell Island", "Maleny", "Fernvale", "Longreach", "Boonah", "Macleay Island", "Burnett Heads", "Cooloola Village", "Yarrabah", "Cedar Vale", "Logan Village", "Palm Island", "Dysart", "Mount Morgan", "St George", "Marian", "Rosewood", "Meringandan West", "Cloncurry", "Kuranda", "Tin Can Bay", "Tully", "Samford Valley–Highvale", "Alice River", "Moore Park", "Toogoom", "Murgon", "Clermont", "Oakhurst", "Gowrie Junction", "Cedar Grove", "Dayboro", "Kilcoy", "Yandina", "Hamilton Island", "Cooroibah", "Middlemount", "Wondai", "Glenview", "Jacobs Well", "Mossman", "Curra", "Cooktown", "Pomona", "Gayndah", "Crows Nest", "Kooralbyn", "Malanda", "Kingsthorpe", "Withcott", "River Heads", "Moura", "Glenwood", "Wyreema", "Doomadgee", "Meridan Plains", "Millmerran", "Hay Point", "Gleneagle", "Childers", "Barcaldine", "Willowbank", "Tolga – West", "Aurukun", "Cherbourg", "Esk", "Burrum Heads", "Plainland", "Cardwell", "Clifton", "Wongaling Beach", "Walloon", "Rainbow Beach", "Normanton", "Laidley Heights", "Jandowae", "Allingham", "Nelly Bay", "Bamaga", "Mundubbera", "Bakers Creek", "Cambooya", "Gununa", "Miles", "Tieri", "Collinsville", "Woodgate", "Kiels Mountain", "Blackall", "Howard", "Hughenden", "Yungaburra", "Mulambin", "Monto", "Toorbul", "Forest Acres", "Rubyvale", "Boyland", "Yabulu", "Blackbutt", "Kinka Beach", "Kurrimine Beach", "Coominya", "Marburg", "Wangan", "Glenden", "Aldershot", "Horseshoe Bay", "Sloping Hummock", "Gooburrum", "Prince Henry Heights", "Sapphire", "Glenore Grove", "Quilpie", "Nome", "Glendale", "Injinoo", "Pie Creek", "Beechmont", "Cooya Beach", "Taroom", "Horn Island", "Mareeba – South", "Karumba", "Apple Tree Creek", "Campwin Beach", "Nebo", "Goomeri", "Richmond", "Greenmount", "Eromanga", "Brisbane", "Gold Coast", "Moreton Bay", "Logan", "Sunshine Coast", "Ipswich", "Townsville", "Toowoomba", "Cairns", "Redland", "Mackay", "Fraser Coast", "Bundaberg", "Rockhampton", "Gladstone", "Noosa", "Gympie", "Scenic Rim", "Lockyer Valley", "Livingstone", "Southern Downs", "Whitsunday", "Western Downs", "South Burnett", "Cassowary Coast", "Central Highlands", "Tablelands", "Somerset", "Mareeba", "Isaac", "Mount Isa", "Burdekin", "Banana", "Maranoa", "Charters Towers", "Douglas", "Hinchinbrook", "Goondiwindi", "North Burnett", "Torres Strait Island", "Balonne", "Murweh", "Cook", "Weipa", "Longreach", "Torres", "Cloncurry", "Barcaldine", "Northern Peninsula Area", "Yarrabah", "Palm Island", "Carpentaria", "Blackall-Tambo", "Paroo", "Flinders", "Doomadgee", "Aurukun", "Cherbourg", "Mornington", "Winton", "Woorabinda", "Napranum", "Kowanyama", "Hope Vale", "Etheridge", "McKinlay", "Boulia", "Bulloo", "Burke", "Mapoon", "Croydon", "Diamantina", "Wujal Wujal", "Barcoo", "Torres Strait"
    ],
    "Western Australia": [
        "APPADENE", "APPERTARRA", "APPLECROSS", "ARAGOON", "ARALUEN", "Abbotts", "Acton Park", "Adamsvale", "Agnew",
        "Ajana", "Albany", "Aldersyde", "Alexandra Bridge", "Alkimos", "Allanooka", "Allanson", "Alma", "Ambania",
        "Amelup", "Amery", "Angelo River", "Anketell", "Antonymyre", "Ardath", "Arrino", "Arrowsmith", "Arthur River",
        "Augusta", "Austin", "Australind", "Baandee", "Babakin", "Badgebup", "Badgingarra", "Badjaling", "Bailup", "Bakers Hill", "Balgo", "Balingup", "Balkuling", "Balladonia", "Ballidu", "Banksiadale", "Bardi", "Beacon", "Bedfordale", "Beermullah", "Bejoording", "Belka", "Bencubbin", "Bendering", "Benger", "Benjaberring", "Beverley", "Big Bell", "Bilbarin", "Bindi Bindi", "Bindoon", "Binningup", "Binnu", "Bodallin", "Boddington", "Bolgart", "Bonnie Rock", "Bonnie Vale", "Boranup", "Borden", "Bornholm", "Boscabel", "Bow Bridge", "Boxwood Hill", "Boyanup", "Boyup Brook", "Bremer Bay", "Bridgetown", "Broad Arrow", "Brookton", "Broome", "Broomehill", "Bruce Rock", "Brunswick Junction", "Bullabulling", "Bullaring", "Bullfinch", "Bullsbrook", "Bulong", "Bunbury", "Bungulla", "Bunjil", "Buntine", "Burakin", "Burekup", "Burracoppin", "Busselton", "Byford", "Eagle Bay", "Ejanding", "Elgin", "Elleker", "Emu Hill", "Eneabba", "Eradu", "Erikin", "Esperance", "Eucla", "Exmouth", "Gabbin", "Gabbadah", "Gairdner", "Gascoyne Junction", "Geraldton", "Gibb River", "Gibson", "Gidgegannup", "Gingin", "Gleneagle", "Gnarabup", "Gnowangerup", "Goldsworthy", "Goomalling", "Gracetown", "Grass Patch", "Grass Valley", "Green Head", "Greenbushes", "Greenhills", "Greenough", "Guilderton", "Gutha", "Gwalia", "Kalannie", "Kalbarri", "Kalgan", "Kalgoorlie", "Kambalda", "Kanowna", "Karakin", "Karlgarin", "Karratha", "Karridale", "Katanning", "Kellerberrin", "Kendenup", "Keysbrook", "King River", "Kirup", "Kiwirrkurra", "Kojarena", "Kojonup", "Kondinin", "Kondut", "Koojan", "Kookynie", "Koolyanobbing", "Koorda", "Korrelocking", "Kukerin", "Kulin", "Kulja", "Kumarina", "Kunjin", "Kununoppin", "Kununurra", "Kweda", "Kwelkan", "Kwolyin", "Nabawa", "Nanga Brook", "Nangeenan", "Nangetty", "Nannine", "Nannup", "Nanson", "Nanutarra", "Narembeen", "Narrikup", "Narrogin", "New Norcia", "Newdegate", "Newman", "Nilgen", "Nornalup", "Norseman", "North Bannister", "North Dandalup", "Northam", "Northampton", "Northcliffe", "Nullagine", "Nungarin", "Nyabing", "Palgarup", "Pannawonica", "Papulankutja", "Pantapin", "Paraburdoo", "Patjarr", "Paynes Find", "Paynesville", "Peak Hill", "Pemberton", "Peppermint Grove Beach", "Perenjori", "Perth", "Piawaning", "Piesseville", "Pindar", "Pingaring", "Pingelly", "Pingrup", "Pinjarra", "Pintharuka", "Pithara", "Point Samson", "Popanyinning", "Porlell", "Porongurup", "Port Denison", "Port Gregory", "Port Hedland", "Preston Beach", "Prevelly", "Princess Royal", "Ranford", "Ravensthorpe", "Rawlinna", "Redmond", "Reedy", "Regans Ford", "Rocky Gully", "Roebourne", "Roelands", "Roleystone", "Rosa Brook", "Rothsay", "Rottnest Island", "Salmon Gums", "Sandstone", "Scaddan", "Seabird", "Serpentine", "Shackleton", "Shay Gap", "Schotts", "Sir Samuel", "South Hedland", "South Kumminin", "Southern Cross", "Stratham", "Tambellup", "Tammin", "Tampa", "Tardun", "Telfer", "Tenindewa", "Tenterden", "The Lakes", "Three Springs", "Tincurrin", "Tjirrkarli", "Tjukurla", "Tom Price", "Toodyay", "Torbay", "Trayning", "Tuckanarra", "Tunney", "Unicup", "Useless Loop", "Xantippe", "Yalgoo", "Yallingup", "Yandanooka", "Yarloop", "Yarri", "Yealering", "Yelbeni", "Yellowdine", "Yerecoin", "Yerilla", "Yilliminning", "Yoongarillup", "York", "Yorkrakine", "Yornaning", "Yornup", "Yoting", "Youanmi", "Youndegin", "Yoweragabbie", "Yuna", "Yundamindera", "Yunndaga", "Zanthus", "Munster"
    ],
    "South Australia": [
        "Adelaide", "Agery", "Alawoona", "Aldgate", "Aldinga", "Alford", "Allendale East", "Allendale North", "Alma",
        "Amata", "American River", "Andamooka", "Andrews", "Angas Plains", "Angas Valley", "Angaston", "Angle Vale",
        "Anna Creek", "Appila", "Ardrossan", "Arkaroola", "Armagh", "Arno Bay", "Arthurton", "Ashville", "Auburn",
        "Avenue", "Avoca Dell", "Babbage", "Baker Gully", "Baker Sandhill", "Balaklava", "Baldina Creek", "Baldon",
        "Balhannah", "Ballast Head", "Bangham", "Bangor", "Barabba", "Baratta", "Barker", "Barmera", "Barna",
        "Baroota", "Barossa", "Basket Range", "Baudin Beach", "Beachport", "Belair", "Beltana", "Belton",
        "Belvidere", "Benbournie", "Berri", "Birdwood", "Blanchetown", "Blyth", "Booleroo Centre", "Bordertown",
        "Boston", "Burra", "Bute", "Callington", "Ceduna", "Charleston", "Clare", "Clarendon", "Clayton Bay", "Cleve",
        "Cobdogla", "Cockatoo Valley", "Cockburn", "Coffin Bay", "Coober Pedy", "Coobowie", "Coonalpyn", "Cowell",
        "Crafers-Bridgewater", "Crystal Brook", "Cummins", "Echunga", "Edithburgh", "Elliston", "Eudunda", "Frances",
        "Freeling", "Gawler", "Gladstone", "Goolwa", "Greenock", "Gumeracha", "Hahndorf", "Hamley Bridge", "Hanson",
        "Hawker", "Hayborough", "Houghton", "Indulkana", "Inglewood", "Iron Knob", "Jamestown", "Kadina",
        "Kalangadoo", "Kaltjiti", "Kanmantoo", "Kapunda", "Karoonda", "Keith", "Kersbrook", "Kimba", "Kingscote",
        "Kingston SE", "Lameroo", "Laura", "Leigh Creek", "Lewiston", "Littlehampton", "Lobethal", "Lock", "Loxton",
        "Lucindale", "Lyndoch", "Macclesfield", "Maitland", "Mallala", "Mannum", "McCracken", "McLaren Flat",
        "McLaren Vale", "Meadows", "Meningie", "Middleton", "Milang", "Millicent", "Mimili", "Minlaton", "Moonta",
        "Morgan", "Mount Barker", "Mount Burr", "Mount Compass", "Mount Gambier", "Mount Pleasant", "Mount Torrens",
        "Mullaquana", "Mundulla", "Murray Bridge", "Myponga", "Nairne", "Nangwarry", "Napperby", "Naracoorte",
        "Normanville", "Nuriootpa", "One Tree Hill", "Oodnadatta", "Orroroo", "Owen", "Parham", "Paringa",
        "Penneshaw", "Penola", "Peterborough", "Pinnaroo", "Point Turton", "Port Adelaide Enfield", "Port Augusta",
        "Port Broughton", "Port Germein", "Port Lincoln", "Port MacDonnell", "Port Neill", "Port Pirie",
        "Port Victoria", "Port Vincent", "Port Wakefield", "Pukatja", "Quorn", "Qualco", "Renmark", "Riverton", "Robe",
        "Roseworthy", "Roxby Downs", "Saddleworth", "Smokey Bay", "Snowtown", "Southend", "Spalding", "Springton",
        "Stansbury", "Strathalbyn", "Streaky Bay", "Summertown", "Swan Reach", "Tailem Bend", "Tantanoola",
        "Tanunda", "Tarpeena", "Tintinara", "Truro", "Tumby Bay", "Two Wells", "Uraidla", "Victor Harbor",
        "Victor Harbor – Goolwa", "Virginia", "Waikerie", "Wallaroo", "Warooka", "Wasleys", "Whyalla",
        "Williamstown", "Willunga", "Willyaroo", "Wilmington", "Wirrabara", "Woodside", "Woomera", "Wudinna",
        "Yahl", "Yalata", "Yankalilla", "Yorketown", "Lochiel", "Geranium"
    ],
    "Tasmania": [
        "Abbotsham", "Abels Bay", "Abercrombie", "Aberdeen", "Acacia Hills", "Acton", "Adventure Bay", "Akaroa",
        "Alberton", "Allens Rivulet", "Alonnah", "Ambleside", "Andover", "Ansons Bay", "Antill Ponds", "Apollo Bay",
        "Arthur River", "Austins Ferry", "Avoca", "Badger Head", "Bagdad", "Bakers Beach", "Bangor", "Barnes Bay",
        "Barrington", "Battery Point", "Beaconsfield", "Beaumaris", "Beauty Point", "Beechford", "Bell Bay",
        "Bellerive", "Ben Lomond", "Bicheno", "Binalong Bay", "Birchs Bay", "Birralee", "Bishopsbourne",
        "Black Hills", "Black River", "Blackmans Bay", "Blessington", "Blue Rocks", "Boat Harbour", "Burnie", "Burnie – Somerset", "Kingston", "Huonville", "Antill Ponds", "Black River", "Alberton",
        "Bicheno", "Smithton", "Kempton", "Avoca", "Bellerive", "Launceston", "Plenty", "Aberdeen", "Birchs Bay",
        "Legana", "Acacia Hills", "Howden", "Blessington", "Queenstown", "Beaconsfield", "Rosebery", "Abercrombie", "Adventure Bay", "Cygnet", "Hagley", "Richmond", "Hobart", "Scottsdale", "Pierces Creek", "Clarence",
        "Abbotsham", "Devonport", "Barnes Bay", "Bakers Beach", "Boat Harbour", "Burnie", "Glenorchy", "Andover",
        "Akaroa", "Allens Rivulet", "Baden", "Badger Head", "Bagdad", "Bakers Beach", "Banca", "Bangor", "Barnes Bay", "Barretta", "Barrington", "Beaconsfield", "Beaumaris", "Beauty Point", "Beechford", "Bell Bay", "Bellingham", "Ben Lomond", "Beulah", "Bicheno", "Binalong Bay", "Birchs Bay", "Birralee", "Bishopsbourne", "Black Hills", "Black River", "Blackwall", "Blackwood Creek", "Blessington", "Blue Rocks", "Blumont", "Boat Harbour Beach", "Boat Harbour", "Boobyalla", "Boomer Bay", "Bothwell", "Boyer", "Bracknell", "Bradys Lake", "Brandum", "Branxholm", "Breadalbane", "Bream Creek", "Breona", "Bridgenorth", "Bridport", "Brighton", "Brittons Swamp", "Broadmarsh", "Broadmeadows", "Bronte Park", "Brooks Bay", "Buckland", "Bungaree", "Burns Creek", "Bushy Park", "Butlers Gorge", "Hobart", "Launceston", "Devonport", "Burnie – Somerset", "Ulverstone", "New Norfolk", "Wynyard", "Dodges Ferry – Lewisham", "Latrobe", "George Town", "Legana", "Port Sorell", "Longford", "Midway Point", "Penguin", "Smithton", "Perth", "Sorell", "Deloraine", "Margate", "Hadspen", "Huonville", "Scottsdale", "Snug", "Queenstown", "Westbury", "St Helens", "Bridport", "Beauty Point", "Primrose Sands", "Sheffield", "Beaconsfield", "Exeter", "Evandale", "Ranelagh", "Cygnet", "Grindelwald", "Richmond", "Railton", "Campbell Town", "Bicheno", "Fern Tree", "Rosebery", "Stieglitz", "Triabunna", "Howden", "Swansea", "Zeehan", "South Arm", "Cressy", "Currie", "Geeveston", "Gravelly Beach", "Scamander", "Strahan", "Sulphur Creek", "Orford", "Clifton Beach", "Campania", "Electrona", "Cremorne", "Oatlands", "Dover", "Low Head", "Stanley", "Carrick", "Bagdad", "St Marys", "Sisters Beach", "Hillwood", "Franklin", "Kettering", "Ridgley", "Forth", "Bracknell", "Eaglehawk Neck", "Opossum Bay", "Bothwell", "Kempton", "Nubeena", "Fingal", "Heybridge", "Dilston", "Lilydale", "White Beach", "Dunalley", "Swan Point", "Ross", "Gawler", "Binalong Bay", "Waratah", "Mole Creek", "Tullah", "Collinsvale", "Macquarie Island"
    ],
    
    "Northern Territory": [
        "Darwin", "Alice Springs", "Alice Springs", "Anthony Lagoon", "Darwin", "Katherine", "Tennant Creek",
        "Palmerston", "Wavell Heights", "Humpty Doo", "Gunyangara", "Hermannsburg", "Howard Springs",
        "Nhulunbuy", "Adelaide River", "Alawa", "Ali Curung", "Alpurrurulam", "Alyangula", "Acacia Hills",
        "Aherrenge", "Bagot", "Baines", "Bakewell", "Barunga", "Napperby", "Narwietooma", "Neutral Junction",
        "Amanbidji", "Amangal Indigenous Village", "Amelia Creek", "Daguragu", "Barunga", "Amanbidji", "Anthony Lagoon", "Aherrenge", "Namerinni", "Alice Springs", "Elliott",
        "Daly River", "Nhulunbuy", "Narwietooma", "Charlotte Waters", "Howard Springs", "Alawa", "Delamere",
        "Nadderns Yard", "Gapuwiyak", "Connellan", "Amelia Creek", "Bakewell", "Desert Springs", "Wavell Heights",
        "Cossack", "Daly Waters", "East Arnhem", "Borroloola", "Darwin", "Baines", "Adelaide River", "Bagot",
        "Dundee Beach", "Ciccone", "Amangal Indigenous Village", "Beswick", "Katherine", "Bulman", "Bynoe",
        "Galiwinku", "Alpurrurulam", "Belyuen", "Darwin River", "Docker River", "Alyangula", "Douglas Daly",
        "Dunmarra", "Ali Curung", "Acacia Hills", "Darwin River Dam", "Atitjere", "Davenport", "Angurugu",
        "Palmerston", "Batchelor", "Gunyangara", "Namaidpa District", "Tennant Creek", "Hermannsburg", "Humpty Doo",
        "Napperby", "Neutral Junction", "Driver", "Areyonga", "Daly River Mango Farm", "Daly River", "Edith River", "Mulga Park", "Maningrida", "Darwin Harbour", "Darwin", "Darwin, Northern Territory"
    ]
}

New_Zealand = {
    "Auckland": [
        "Auckland","Botanical Garden", "Auckland Botanic Gardens", "Beachlands", "Beachlands-Pine Harbour", "Clarks Beach", "Clevedon", "Devonport", "Drury", "Great Barrier Island", "Helensville", "Hibiscus Coast", "Howick", "Kumeū-Huapai", "Maraetai", "Muriwai", "Orewa", "Parakai", "Patumāhoe", "Pukekohe", "Riverhead", "Smith's Bush", "Snells Beach", "St Johns Bush", "Tuakau", "Waiatarua Reserve", "Waiheke West", "Waimauku", "Waiuku", "Warkworth", "Wellsford", "Western Springs Park", "Pōkeno"
    ],
    "Canterbury": [
        "Akaroa", "Amberley", "Arthurs Pass National Park", "Ashburton", "Banks Peninsula", "Birdlings Flat", "Christchurch", "Culverden", "Darfield", "Diamond Harbour", "Geraldine", "Hanmer Springs", "Kaiapoi", "Kaikōura", "Leeston", "Lincoln", "Lyttelton", "Methven", "Oxford", "Pegasus", "Pleasant Point", "Prebbleton", "Rakaia", "Rangiora", "Rolleston", "Taylors Mistake", "Temuka", "Timaru", "Twizel", "Waikari", "Waikuku Beach", "West Melton", "Woodend"
    ],
    "Wellington": [
        "Carterton", "Featherston", "Greytown", "Kaitoke", "Kaitoke Waterworks", "Kapiti Coast", "Lower Hutt", "Martinborough", "Masterton", "Paekākāriki", "Porirua", "Upper Hutt", "Wellington", "Ōtaki", "Ōtaki Beach", "Paraparaumu", "Waikanae"
    ],
    "Waikato": [
        "Cambridge", "Coromandel", "Desert Road", "Hamilton", "Huntly", "Kihikihi", "Kinloch", "Matamata", "Morrinsville", "Ngatea", "Ngāruawāhia", "Opepe Historic Reserve", "Paeroa", "Pauanui", "Pirongia", "Putāruru", "Raglan", "Rauroa Bush Reserve", "Tairua", "Taupō", "Te Aroha", "Te Awamutu", "Te Kauwhata", "Te Kūiti", "Thames", "Tokoroa", "Tuakau", "Tūrangi", "Waihi", "Waiau Falls", "Waitomo", "Whangamatā", "Whitianga", "Ōtorohanga", "Pōkeno"
    ],
    "Bay of Plenty": [
        "Edgecumbe", "Katikati", "Kawerau", "Maketu", "Murupara", "Ngongotahā", "Ōhope", "Ōmokoroa", "Ōpōtiki", "Rotorua", "Tauranga", "Te Puke", "Wai-O-Tapu", "Waihi Beach-Bowentown", "Whakatāne"
    ],
    "Otago": [
        "Alexandra", "Arrowtown", "Balclutha", "Brighton", "Clyde", "Cromwell", "Dunedin", "Kaitangata", "Kaka Point", "Lake Hāwea", "Lake Wanaka", "Macandrew Bay", "Macraes Flat", "Milton", "Mosgiel", "Nugget Point", "Oamaru", "Port Chalmers", "Queenstown", "Waikouaiti", "Wānaka"
    ],
    "Manawatu-Wanganui": [
        "Ashhurst", "Bruce Park", "Bulls", "Dannevirke", "Feilding", "Foxton", "Foxton Beach", "Himatangi Beach", "Hunterville", "Levin", "Manawatu", "Marton", "Ohakune", "Pahiatua", "Palmerston North", "Raetihi", "Shannon", "Taihape", "Taumarunui", "Wanganui", "Whanganui", "Woodville"
    ],
    "Northland": [
        "Ahipara", "Dargaville", "Haruru", "Hikurangi", "Kaikohe", "Kaitaia", "Kawakawa", "Kerikeri", "Mangawhai", "Mangawhai Heads", "Moerewa", "Ngunguru", "One Tree Point", "Opua", "Paihia", "Ruakākā", "Russell", "Waipu", "Waiotama", "Whangārei"
    ],
    "Nelson": [
        "Brightwater", "Hope", "Māpua", "Nelson", "Richmond", "Wakefield"
    ],
    "Taranaki": [
        "Eltham", "Hawera", "Inglewood", "Kapuni", "New Plymouth", "Normanby", "Ōakura", "Ōpunake", "Patea", "Stratford", "Waitara"
    ],
    "Hawke’s Bay": [
        "Clive","Hawke’s Bay", "Hastings", "Haumoana", "Havelock North", "Hawke’s Bay", "Napier", "Waipawa", "Waipukurau", "Wairoa"
    ],
    "Gisborne": [
        "Gisborne"
    ],
    "Southland": [
        "Bluff", "Gore", "Invercargill", "Lumsden", "Mataura", "Niagara Falls", "Riverton", "Te Anau", "Winton"
    ],
    "Marlborough": [
        "Blenheim", "Havelock", "Maud Island", "Motuara Island", "Picton", "Renwick"
    ],
    "Tasman": [
        "Brightwater", "Hope", "Māpua", "Motueka", "Murchison", "Richmond", "Tākaka", "Wakefield"
    ],
    "West Coast": [
        "Greymouth", "Hokitika", "Okarito", "Reefton", "Ross", "Runanga", "Saltwater Forest", "Westport"
    ],
    "Chatham Islands": [
        "Chatham Islands", "Chatham Rise"
    ],
 "Pacific Ocean": [
        "New Zealand coast" , "Pacific Ocean"
    ]

}

Australia_states = {
    "New South Wales": "NSW",
    "Queensland": "QLD",
    "South Australia": "SA",
    "Tasmania": "TAS",
    "Victoria": "VIC",
    "Western Australia": "WA"
}

New_Zealand_regions = {
    "Auckland": "AUK",
    "Bay of Plenty": "BOP",
    "Canterbury": "CAN",
    "Gisborne": "GIS",
    "Hawke’s Bay": "HKB",
    "Manawatu-Wanganui": "MWT",
    "Marlborough": "MBH",
    "Nelson": "NSN",
    "Northland": "NTL",
    "Otago": "OTA",
    "Southland": "STL",
    "Taranaki": "TKI",
    "Tasman": "TAS",
    "Waikato": "WKO",
    "Wellington": "WGN",
    "West Coast": "WTC"
}


# ==========================================
# 3. پیش‌پردازش
# ==========================================
au_city_to_state = {city.lower(): state for state, cities in Australia.items() for city in cities}
nz_city_to_region = {city.lower(): region for region, cities in New_Zealand.items() for city in cities}

au_abbr_to_state = {v: k for k, v in Australia_states.items()}
nz_abbr_to_region = {v: k for k, v in New_Zealand_regions.items()}

# ==========================================
# 4. تابع منطق تشخیص (Classification)
# ==========================================
def classify_location(text):
    if not isinstance(text, str):
        return "Other", "Other"
    
    text_lower = text.lower()
    
    # --- Check Australia ---
    if "australia" in text_lower:
        country = "Australia"
        
        # 1. State Name
        for state in Australia.keys():
            if state.lower() in text_lower:
                return country, state
        
        # 2. Abbreviation
        for abbr, state in au_abbr_to_state.items():
            if re.search(r'\b' + re.escape(abbr.lower()) + r'\b', text_lower):
                return country, state
                
        # 3. City Name
        for city, state in au_city_to_state.items():
            if city in text_lower:
                return country, state
        
        # 4. Not Found -> Unknown State
        return country, "Unknown State"

    # --- Check New Zealand ---
    elif "new zealand" in text_lower:
        country = "New Zealand"
        
        # 1. Region Name
        for region in New_Zealand.keys():
            if region.lower() in text_lower:
                return country, region
                
        # 2. Abbreviation
        for abbr, region in nz_abbr_to_region.items():
            if re.search(r'\b' + re.escape(abbr.lower()) + r'\b', text_lower):
                return country, region
                
        # 3. City Name
        for city, region in nz_city_to_region.items():
            if city in text_lower:
                return country, region
                
        # 4. Not Found -> Unknown Region
        return country, "Unknown Region"

    else:
        return "Other", "Other"

# ==========================================
# 5. اجرای حلقه اصلی
# ==========================================
print("Starting process...")

all_summaries = [] # برای شیت اول
all_unknowns = []  # برای شیت دوم (هر دو کشور)

for file_name in files_info:
    file_path = os.path.join(input_dir, file_name)
    
    if not os.path.exists(file_path):
        print(f"Warning: File not found: {file_name} -> Skipped")
        continue
        
    print(f"Processing: {file_name}")
    
    try:
        df = pd.read_excel(file_path)
        
        if 'Geo Loc Name' not in df.columns:
            print(f"   -> Column 'Geo Loc Name' missing in {file_name}")
            continue

        # اعمال تابع تشخیص
        df[['Derived_Country', 'Derived_State']] = df['Geo Loc Name'].apply(
            lambda x: pd.Series(classify_location(x))
        )

        # فیلتر کردن: فقط ردیف‌هایی که یا استرالیا هستند یا نیوزلند
        result_df = df[df['Derived_Country'].isin(['Australia', 'New Zealand'])]
        
        if result_df.empty:
            continue

        # --- بخش ۱: محاسبه خلاصه آماری (شیت اول) ---
        summary = result_df.groupby(['Derived_Country', 'Derived_State']).size().reset_index(name='Count')
        summary['files_info'] = file_name 
        all_summaries.append(summary)
        
        # --- بخش ۲: جدا کردن ناشناخته‌ها (شیت دوم) ---
        # در اینجا هم 'Unknown State' (مال استرالیا) و هم 'Unknown Region' (مال نیوزلند) انتخاب می‌شوند
        unknowns_df = result_df[result_df['Derived_State'].isin(['Unknown State', 'Unknown Region'])].copy()
        
        if not unknowns_df.empty:
            unknowns_df['files_info'] = file_name
            # انتخاب ستون‌های مورد نیاز برای نمایش
            unknowns_df = unknowns_df[['Geo Loc Name', 'Derived_Country', 'Derived_State', 'files_info']]
            all_unknowns.append(unknowns_df)
            
    except Exception as e:
        print(f"   -> Error processing {file_name}: {e}")

# ==========================================
# 6. ذخیره فایل نهایی
# ==========================================

print("\nSaving files...")

with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    
    # Sheet 1: Summary Counts
    if all_summaries:
        final_summary = pd.concat(all_summaries, ignore_index=True)
        final_summary = final_summary[['Derived_Country', 'Derived_State', 'Count', 'files_info']]
        final_summary.to_excel(writer, sheet_name='Summary', index=False)
        print(" -> 'Summary' sheet created.")
    else:
        pd.DataFrame({'Message': ['No Data']}).to_excel(writer, sheet_name='Summary', index=False)

    # Sheet 2: Unknown Details (Both Countries)
    if all_unknowns:
        final_unknowns = pd.concat(all_unknowns, ignore_index=True)
        final_unknowns.to_excel(writer, sheet_name='Unknown_Details', index=False)
        print(" -> 'Unknown_Details' sheet created (Includes both Australia & NZ unknowns).")
    else:
        pd.DataFrame({'Message': ['No Unknown States/Regions found for Australia or NZ']}).to_excel(writer, sheet_name='Unknown_Details', index=False)
        print(" -> No Unknowns found (Empty sheet created).")

print(f"\nDone! File saved at:\n{output_path}")