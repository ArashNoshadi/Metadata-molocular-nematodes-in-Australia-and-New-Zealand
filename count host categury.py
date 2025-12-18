import pandas as pd
import os
import re

# ==========================================
# 1. تنظیمات مسیر و فایل‌ها
# ==========================================
input_dir = r'G:\Paper\nema-Nanopore-Sequencing\zoology new zealand and australia\data\Suspect'
output_path = os.path.join(input_dir, 'Host_Analysis_Deep_Level.xlsx')

files_info = [
    '5.8S.xlsx',
    'ITS2.xlsx',
    'ITS1.xlsx',
    '18S.xlsx',
    '28S.xlsx',
    'COX1.xlsx',
]

# ==========================================
# 2. دیکشنری جدید میزبان (Nested Host Dictionary)
# ==========================================
# توجه: لطفاً لیست‌های خالی [] را با کلمات کلیدی خودتان پر کنید.
Host = {
    "Animal": {
        "Aquatic Invertebrates": {
            "Crustaceans": ["diplopod Spirostreptida",
                            "Oncocladosoma castaneum",
                            "Balanus sp."],
            "Mollusks": ["Codakia paytenorum (mollusc)",
                         "Octopus sp. d MA-2020 (octopus)",
                         "Athoracophorus bitentaculatus",
                         "Nototodarus sloanii (Arrow squid)",
                         "Moroteuthopsis ingens (Warty squid)",
                         "Nototodarus solanii",
                         "Moroteuthopsis ingens",
                         "freshwater snail",
                         "Arrow squid, Nototodarus sloanii",
                         "Nototodarus sloanii",
                         "Warty squid, Moroteuthopsis ingens"]
        },
        "Aquatic Vertebrates": {
            "Aquatic Reptiles": ["Crocodylus porosus",
                                 "Chelodina rugosa",
                                 "Chelodina burrungandjii",
                                 "Emydura tanybaraga",
                                 "Chelodina canni",
                                 "Hydrophis peronii (sea snake)",
                                 "Elseya dentata",
                                 "Chelodina expansa",
                                 "Emydura macquarii",
                                 "Emydura australis"],
            "Fish": ["Trygonorrhina fasciata",
                     "Upeneichthys lineatus",
                     "Hyporhamphus regularis",
                     "Heterodontus portusjacksoni",
                     "Anguilla reinhardtii",
                     "Anguilla australis",
                     "Leiopotherapon unicolor",
                     "Maccullochella peelii",
                     "Galaxias olidus",
                     "Pastinachus ater (stingray)",
                     "Hypseleotris sp. 5",
                     "Hypseleotris klunzingeri",
                     "Acanthopagrus australis",
                     "Rhabdosargus sarba",
                     "Nanoperca australis",
                     "Neochanna apoda",
                     "Seriolella brama (Blue warehou)",
                     "Chelidonichthys cuculus (Red gurnard)",
                     "Paratrachichthys trailli (Common roughy)",
                     "Tripterygiidae gen. sp. (Triplefin)",
                     "Bovichtus variegatus (Thornfish)",
                     "Notolabrus celidotus (Scarlett wrasse)",
                     "Thyrsites atun (Barracouta)",
                     "Congiopodus leucopaecilus (Pigfish)",
                     "Aldrichetta forsteri (Mullet)",
                     "Pseduolabrus fucicola (Banded wrasse)",
                     "Sprattus muelleri (Sprat)",
                     "Acanthoclinus fuscus (Olive rockfish)",
                     "Nemadactylus macropterus (Tarahiki)",
                     "Parapercis colias (Blue cod)",
                     "Seriolella punctata (Silver warehou)",
                     "Sprattus antipodum (Sprat)",
                     "Sprat, Sprattus antipodum",
                     "Triplefin, Tripterygiidae gen. sp.",
                     "Giant stargazer, Kathetostoma giganteum",
                     "NZ sole, Peltorhamphus novaezeelandiae",
                     "Spiny dogfish, Squalus acanthias",
                     "Sixgill shark, Hexanchus griseus",
                     "Mullet, Aldrichetta forsteri",
                     "Triplefin, Forsterygion capito",
                     "Sprat, Sprattus muelleri",
                     "Anchovy, Engraulis australis",
                     "Pigfish, Congiopodus leucopaecilus",
                     "Triplefin, Tripterygiidae gen. spp.",
                     "Scaly gurnard, Lepidotrigla brachyoptera",
                     "Crested bellowsfish, Notopogon lilliei",
                     "Witch, Arnoglossus sp.",
                     "Pelotretis flavilatus (Lemon sole)",
                     "Clingfish, Gastroscyphus hectoris",
                     "Seahorse, Hippocampus abdominalis",
                     "Opalfish, Hemerocoetes monopterygius",
                     "Tarahiki, Nemadactylus macropterus",
                     "School shark, Galeorhinus galeus",
                     "Red gurnard, Chelidonichthys cuculus",
                     "Silver warehou, Seriolella punctata",
                     "Red cod, Pseudophycis bachus",
                     "Blue warehou, Seriolella brama",
                     "Peltorhamphus novaezeelandiae",
                     "Scomber australasicus",
                     "Platycephalus richardsoni",
                     "Platycephalus bassensis",
                     "Platycephalus fuscus",
                     "Anguilla sp.",
                     "Banded wrasse, Pseduolabrus fucicola",
                     "Barracouta, Thyrsites atun",
                     "Blue cod, Parapercis colias",
                     "Blue warehou, Seriolella brama",
                     "Clingfish, Gastroscyphus hectoris",
                     "Crested bellowsfish, Notopogon lilliei",
                     "Olive rockfish, Acanthoclinus fuscus",
                     "Parapercis colias (blue cod)",
                     "Red cod, Pseudophycis bachus",
                     "Red gurnard, Chelidonichthys cuculus",
                     "Scaly gurnard, Lepidotrigla brachyoptera",
                     "Scarlett wrasse, Notolabrus celidotus",
                     "School shark, Galeorhinus galeus",
                     "Sillago flindersi",
                     "Sprat, Sprattus antipodum",
                     "Sprat, Sprattus muelleri",
                     "Sprattus antipodum (sprat)",
                     "Tarakihi, Nemadactylus macropterus",
                     "Thornfish, Bovichtus variegatus",
                     "Arripis georgianus",
                     "Trachurus novaezelandiae",
                     "Arripis georgianus, Trachurus novaezelandiae",
                     "Forsterygion capito (triplefin)",
                     "Kathetostoma giganteum (giant stargazer)",
                     "Engraulis australis (anchovy)",
                     "Tripterygiidae sp. (triplefin)",
                     "Lepidotrigla brachyoptera (scaly gurnard)",
                     "Notopogon lilliei (crested bellowsfish)",
                     "Arnoglossus sp. (witch)",
                     "Gastroscyphus hectoris (clingfish)",
                     "Hippocampus abdominalis (seahorse)",
                     "Nemadactylus macropterus (tarakihi)",
                     "Pseudophycis bachus (red cod)",
                     "Johnius sp.",
                     "Polydactylus macrochir",
                     "Scomberoides commersonnianus"]
        },
        "Terrestrial Vertebrates": {
            "Amphibians": ["Litoria lesueuri",
                           "Salamandrella keyserlingii"],
            "Birds": ["Gymnorhina tibicen (magpie)",
                      "Eudyptula novaehollandiae",
                      "Gallus gallus",
                      "Phalacrocorax sulcirostris",
                      "Falco berigora",
                      "Falco longipennis",
                      "Apteryx rowi",
                      "Circus approximans (Australasian harrier)",
                      "Pelecanoides urinatrix (Common diving petrel)",
                      "Procellaria parkinsoni (Black petrel)",
                      "Procellaria cinerea (Grey petrel)",
                      "Ardenna carneipes (Flesh-footed shearwater)",
                      "Leucocarbo carunculatus (King shag)",
                      "Ardenna grisea (Sooty shearwater)",
                      "Procellaria westlandica (Westland petrel)",
                      "Haematopus unicolor (Variable oystercatcher)",
                      "Procellaria aequinoctialis (White-chinned petrel)",
                      "Thalassarche cauta (White-capped mollymawk)",
                      "Macronectes halli (Northern giant petrel)",
                      "Thalassarche salvini (Salvin's mollymawk)",
                      "Larus dominicanus (Black-backed gull)",
                      "Chroicocephalus scopulinus (Red-billed gull)",
                      "Haematopus finschi (South Island pied oystercatcher)",
                      "Diomedea sanfordi (Northern royal albatross)",
                      "Chroicocephalus scopulinus (Red-billed gull)",
                      "Phalacrocorax punctatus (Spotted shag)",
                      "Leucocarbo chalconotus (Otago shag)",
                      "Hydroprogne caspia (Caspian tern)",
                      "Eudyptes pachyrhynchus (Fiordland crested penguin)",
                      "Eudyptula novaehollandiae (Little blue penguin)",
                      "Megadyptes antipodes (Yellow-eyed penguin)",
                      "Royal spoonbill, Platalea flavipes",
                      "Black-backed gull, Larus dominicanus",
                      "Salvin's mollymawk, Thalassarche salvini",
                      "Grey-headed mollymawk, Thalassarche chrysostoma",
                      "Flesh-footed shearwater, Ardenna carneipes",
                      "White-chinned petrel, Procellaria aequinoctialis",
                      "Variable oystercatcher, Haematopus unicolor",
                      "Spotted shag, Phalacrocorax punctatus",
                      "Common diving petrel, Pelecanoides urinatrix",
                      "Little blue penguin, Eudyptula novaehollandiae",
                      "Northern royal albatross, Diomedea sanfordi",
                      "Little pied shag, Microcarbo melanoleucos",
                      "Red-billed gull, Chroicocephalus scopulinus",
                      "Kingfisher, Todiramphus sanctus",
                      "Spotted shag, Phalacrocorax punctatus",
                      "Snares crested penguin, Eudyptes robustus",
                      "White-faced heron, Egretta novaehollandiae",
                      "Australasian crested grebe, Podiceps cristatus australis",
                      "Microcarbo melanoleucos brevirostris (Little shag)",
                      "Otago shag, Leucocarbo chalconotus",
                      "Chroicocephalus scopulinus",
                      "Circus approximans",
                      "Apteryx mantelli mantelli",
                      "Pelecanus conspicillatus",
                      "Chroicocephalus scopulinus (Red-billed gull)",
                      "Diomedea sanfordi (northern royal albatross)",
                      "Eudyptes pachyrhynchus (fiordland crested penguin)",
                      "Eudyptula novaehollandiae (little blue penguin)",
                      "Fiordland crested penguin, Eudyptes pachyrhynchus",
                      "Haematopus finschi (South Island pied oystercatcher)",
                      "King shag, Leucocarbo carunculatus",
                      "Larus dominicanus (Black-backed gull)",
                      "Leucocarbo carunculatus (king shag)",
                      "Leucocarbo colensoi (Auckland island shag)",
                      "Little blue penguin, Eudyptula novaehollandiae",
                      "Macropus giganteus (eastern grey kangaroo)",
                      "Megadyptes antipodes (yellow-eyed penguin)",
                      "Northern giant petrel, Macronectes halli",
                      "Northern royal albatross, Diomedea sanfordi",
                      "Phalacrocorax punctatus (spotted shag)",
                      "Spotted shag, Phalacrocorax punctatus",
                      "Western scrub wallaby, Notamacropus irma",
                      "White-capped mollymawk, Thalassarche cauta",
                      "White-chinned petrel, Procellaria aequinoctialis",
                      "Yellow-eyed penguin, Megadyptes antipodes",
                      "Eudyptes robustus (Snares crested penguin)",
                      "Elseya latisternum",
                      "Thalassarche salvini (Salvin's albatross)",
                      "Microcarbo melanoleucos (Little pied shag)",
                      "Chroicocephalus novaehollandiae scopulinus"],
            "Mammals": ["Wallabia bicolor",
                         "Rattus fuscipes",
                         "Perameles gunnii",
                         "Macropus giganteus",
                         "Isoodon obesulus",
                         "Osphranter robustus woodwardi",
                         "Thylogale stigmatica",
                         "Setonix brachyurus",
                         "Macropus dorsalis",
                         "Petrogale assimilis",
                         "Petrogale purpureicollis",
                         "Macropus antilopinus",
                         "Neophoca cinerea (Australian sea lion); voucher: SBDP12-076",
                         "Macropus rufus",
                         "Petrogale herberti",
                         "Petrogale persephone",
                         "Petrogale inornata",
                         "Neophoca cinerea (Australian sea lion); voucher: SBDP12-062",
                         "Neophoca cinerea (Australian sea lion); voucher: SBDP12-073",
                         "Neophoca cinerea (Australian sea lion); voucher: SBDP12-061",
                         "Neophoca cinerea (Australian sea lion); voucher: SBDP12-095",
                         "Neophoca cinerea (Australian sea lion); voucher: SBDP12-071",
                         "Neophoca cinerea (Australian sea lion); voucher: SBDP12-088",
                         "Neophoca cinerea (Australian sea lion); voucher: SBDP12-082",
                         "Neophoca cinerea (Australian sea lion); voucher: SBDP12-092",
                         "Neophoca cinerea (Australian sea lion); voucher: SBDP12-022; collected by C. Kennedy",
                         "Neophoca cinerea (Australian sea lion); voucher: PSDP09-01; collected by C. Kennedy",
                         "Neophoca cinerea (Australian sea lion); voucher: DRDP11-13",
                         "Neophoca cinerea (Australian sea lion); voucher: DRDP13-06",
                         "Neophoca cinerea (Australian sea lion); voucher: DRDP11-26",
                         "Neophoca cinerea (Australian sea lion); voucher: DRDP11-21",
                         "Neophoca cinerea (Australian sea lion); voucher: DRDP11-06",
                         "Neophoca cinerea (Australian sea lion); voucher: DRDP11-03",
                         "Neophoca cinerea (Australian sea lion); voucher: DRDP11-15",
                         "Neophoca cinerea (Australian sea lion); voucher: DRDP11-08",
                         "Arctocephalus pusillus doriferus",
                         "Mirounga leonina",
                         "Macropus parryi",
                         "Macropus agilis",
                         "Canis lupus familiaris",
                         "Homo sapiens",
                         "Petrogale mareeba",
                         "Notamacropus rufogriseus",
                         "Thylogale billardierii",
                         "Pseudocheirus peregrinus",
                         "Osphranter robustus",
                         "Notamacropus agilis",
                         "Thoroughbred horse",
                         "Felis catus (domestic cat)",
                         "Vulpes vulpes (red fox)",
                         "Isoodon macrourus",
                         "Gymnobelideus leadbeateri",
                         "dog",
                         "Kogia breviceps (Pygmy sperm whale)",
                         "Notamacropus irma",
                         "sheep",
                         "cattle",
                         "Cervus elaphus (red deer)",
                         "Phocarctos hookeri",
                         "Hydrurga leptonyx (Leopard seal)",
                         "Eudyptula minor",
                         "Leopard seal, Hydrurga leptonyx",
                         "Mesoplodon grayi",
                         "Globicephala melas",
                         "Canis familiaris (domestic dog)",
                         "Sarcophilus harrisii",
                         "Macropus fuliginosus",
                         "Vombatus ursinus",
                         "Lasiorhinus latifrons",
                         "Notamacropus dorsalis",
                         "Dendrolagus bennettianus",
                         "Lagorchestes conspicillatus",
                         "Macropus (Notamacropus) dorsalis",
                         "Dendrolagus lumholtzi",
                         "Notamacropus parryi",
                         "Osphranter rufus",
                         "dog; breed: greyhound",
                         "Thylogale thetis",
                         "Macropus fuliginosus ocydromus",
                         "Macropus fuliginosus fuliginosus",
                         "Mirounga leonine",
                         "Neophoca cinerea",
                         "Arctocephalus forsteri",
                         "Canis familiaris (dog)",
                         "Canis familiaris (domestic dog)",
                         "cow",
                         "Macropus giganteus (eastern grey kangaroo)",
                         "Macropus robustus",
                         "Macropus robustus erubescens",
                         "Macropus robustus robustus",
                         "Neophoca cinerea (Australian sea lion); voucher: DRDP11-03",
                         "Neophoca cinerea (Australian sea lion); voucher: DRDP11-06",
                         "Neophoca cinerea (Australian sea lion); voucher: DRDP11-08",
                         "Neophoca cinerea (Australian sea lion); voucher: DRDP11-13",
                         "Neophoca cinerea (Australian sea lion); voucher: DRDP11-15",
                         "Neophoca cinerea (Australian sea lion); voucher: DRDP11-21",
                         "Neophoca cinerea (Australian sea lion); voucher: DRDP11-26",
                         "Neophoca cinerea (Australian sea lion); voucher: DRDP13-06",
                         "Neophoca cinerea (Australian sea lion); voucher: SBDP12-061",
                         "Neophoca cinerea (Australian sea lion); voucher: SBDP12-062",
                         "Neophoca cinerea (Australian sea lion); voucher: SBDP12-071",
                         "Neophoca cinerea (Australian sea lion); voucher: SBDP12-073",
                         "Neophoca cinerea (Australian sea lion); voucher: SBDP12-076",
                         "Neophoca cinerea (Australian sea lion); voucher: SBDP12-082",
                         "Neophoca cinerea (Australian sea lion); voucher: SBDP12-088",
                         "Neophoca cinerea (Australian sea lion); voucher: SBDP12-092",
                         "Notomacropus irma (Black-glove wallaby)",
                         "Osphranter antilopinus (antilopine wallaroo)",
                         "Osphranter bernardus (black wallaroo)",
                         "Osphranter robustus (commom wallaroo)",
                         "Petauroides volans",
                         "Petrogale penicillata",
                         "Pseudocherius peregrinus",
                         "rabbit",
                         "Rattus rattus",
                         "Trichosurus vulpecula",
                         "Cervus elaphus",
                         "Felis catus",
                         "Notamacropus eugenii",
                         "swamp wallaby",
                         "Ceratotherium simum",
                         "Ovis aries",
                         "red deer",
                         "Sus scrofa (wild pig)",
                         "Dasyurus geoffroii",
                         "Tachyglossus aculeatus",
                         "ox",
                         "Dasyurus hallucatus",
                         "Dasyurus viverrinus",
                         "Rattus lutreolus"],
            "Reptiles": ["Varanus indicus",
                         "Dactylocnemis pacificus",
                         "Woodworthia sp. type Otago large",
                         "Woodworthia maculata",
                         "Woodworthia brunneus",
                         "Naultinus punctatus",
                         "Naultinus gemmeus",
                         "Oligosoma polychroma",
                         "Oligosoma aeneum",
                         "Oligosoma maccanni",
                         "Oligosoma infrapunctatum type crenulate",
                         "Tiliqua scincoides",
                         "Cyclodomorphus gerrardii",
                         "Morelia spilota spilota (diamond python)"]
        }
    },
    "Insect": {
        "Mutualists": {
            "Pollinators": [],
            "Protective Ants": []
        },
        "Primary Consumers (Herbivores)": {
            "Chewing Insects": ["Helicoverpa armigera (Hubner)",
                                "Sirex noctilio",
                                "Fergusonina sp.",
                                "Forficula auricularia",
                                "Austrocidaria sp.",
                                "Pyrgotis plagiatana",
                                "Declana floccosa",
                                "Austrocidaria anguligera",
                                "Pseudocoremia suavis",
                                "Proteodes profunda",
                                "Xyridacma alectoraria",
                                "Gellonia pannularia",
                                "Apoctena orthropis",
                                "Forficula auricularia"],
            "Sucking Insects": []
        },
        "Secondary Consumers": {
            "Predators": [],
            "Parasitoids": []
        },
        "Vectors": {
            "Virus Vectors": [],
            "Bacterial Vectors": [],
            "Nematode-Associated Vectors": []
        }
    },
    "Plant": {
        "Herbaceous Plant": {
            "Agricultural Crops": {
                "Cereals": ["wheat (Triticum sp.)",
                            "Zea mays",
                            "maize",
                            "wheat"],
                "Legumes": ["Vicia faba",
                            "Snake bean"],
                "Oil Crops": []
            },
            "Vegetables": ["potato (Solanum tuberosum)",
                           "Solanum physalifolium",
                           "Butternut Pumpkin",
                           "Capsicum",
                           "cucumber",
                           "Cucumis sativus",
                           "Sweet Potato",
                           "Salicornia quinqueflora",
                           "banana",
                           "Musa acuminata AAA Group",
                           "Hop"],
            "Grasses": ["kikuyu",
                        "bentgrass",
                        "Lachnagrostis filiformis",
                        "Polypogon monspeliensis; annual beardgrass",
                        "Stipa sp.",
                        "Holcus lanatus; velvetgrass",
                        "Astrebla pectinata",
                        "Microlaena stipodes",
                        "Lolium rigidum; annual ryegrass",
                        "Ficinia spiralis",
                        "ryegrass",
                        "Polypogon monspeliensis",
                        "Festuca nigrescens",
                        "Lolium perenne",
                        "Trifolium repens",
                        "Eleocharis gracilis",
                        "grasses",
                        "Lolium rigidum",
                        "Astrebra pectinata",
                        "Saccharum sp."],
            "Ornamental Herbaceous Plants": []
        },
        "Woody Plant": {
            "Fruit Trees": ["Coffea arabica",
                            "Ficus racemosa",
                            "Ficus rubiginosa",
                            "Ficus benjamina",
                            "Ficus hispida",
                            "Ficus obliqua",
                            "apple (Malus domestica Borkh.)",
                            "Syzygium sp.",
                            "Coffea arabica cv. Catuai Rojo",
                            "Coffea arabica cv. Bourbon",
                            "Pecan",
                            "fig",
                            "Ficus variegata",
                            "Syzygium luehmannii"],
            "Forest Trees": ["Eucalyptus macrorhyncha",
                             "Melaleuca quinquenervia",
                             "Eucalyptus camaldulensis",
                             "Eucalyptus bridgesiana",
                             "eucalyptus",
                             "Melaleuca leucadendra",
                             "Eucalyptus tereticornis",
                             "Corymbia tessellaris",
                             "Eucalyptus delegatensis",
                             "Angophora floribunda",
                             "Melaleuca fluviatilis",
                             "Eucalyptus cosmophylla",
                             "Eucalyptus intertexta",
                             "Corymbia sp.",
                             "Melaleuca dealbata",
                             "Melaleuca argentea",
                             "Eucalyptus sp.",
                             "Pinus radiata",
                             "Callitris preissii",
                             "Metrosideros excelsa",
                             "Nothofagus sp.",
                             "Corymbia maculata",
                             "Corymbia ptychocarpa",
                             "Melaleuca nervosa",
                             "Melaleuca viridiflora",
                             "Myrtaceae",
                             "Plagianthus regius",
                             "Sophora microphylla"],
            "Ornamental Trees": [],
            "Shrubs": ["Leptospermum laevigatum",
                       "Melaleuca cajuputi",
                       "Melaleuca decora",
                       "Melaleuca linariifolia",
                       "Melaleuca nodosa",
                       "Melaleuca armillaris",
                       "Melaleuca stenostachya"],
            "Woody Vines": []
        }
    }
}

# ==========================================
# 3. پیش‌پردازش بازگشتی (Recursive Processing)
# ==========================================

host_search_list = []      # لیست کلمات برای جستجو
all_defined_paths = set()  # لیست تمام مسیرهای تعریف شده (برای شیت سوم)

def flatten_host_dict(d, path=()):
    """
    این تابع به صورت بازگشتی دیکشنری تو در تو را می‌خواند
    و کلمات کلیدی را به همراه مسیر کاملشان استخراج می‌کند.
    """
    for k, v in d.items():
        current_path = path + (k,)
        
        if isinstance(v, dict):
            # اگر هنوز دیکشنری است، عمیق‌تر برو
            flatten_host_dict(v, current_path)
        elif isinstance(v, list):
            # اگر به لیست رسیدیم (برگ‌های درخت)
            # مسیر را به 4 سطح استاندارد می‌رسانیم (با None پر می‌کنیم)
            padded_path = list(current_path) + [None] * (4 - len(current_path))
            padded_path_tuple = tuple(padded_path)
            
            # ذخیره مسیر تعریف شده (حتی اگر لیست خالی باشد)
            all_defined_paths.add(padded_path_tuple)
            
            # اضافه کردن کلمات کلیدی به لیست جستجو
            for item in v:
                clean_item = item.strip().lower()
                # فرمت: (کلمه, سطح1, سطح2, سطح3, سطح4)
                host_search_list.append((clean_item, *padded_path))

# اجرای تابع بازگشتی روی دیکشنری اصلی
flatten_host_dict(Host)

# مرتب‌سازی لیست جستجو بر اساس طول کلمات (کلمات طولانی‌تر اولویت دارند)
host_search_list.sort(key=lambda x: len(x[0]), reverse=True)

# ==========================================
# 4. توابع تشخیص (Classification Functions)
# ==========================================

def get_country_only(text):
    if not isinstance(text, str): return None
    t = text.lower()
    if "australia" in t: return "Australia"
    if "new zealand" in t: return "New Zealand"
    return None

def classify_host_deep(text):
    """جستجوی میزبان و بازگرداندن 4 سطح دسته‌بندی"""
    if not isinstance(text, str):
        return "Unknown Host", None, None, None
    
    text_lower = text.lower()
    
    for entry in host_search_list:
        keyword = entry[0]
        # entry[1] تا entry[4] همان سطوح 1 تا 4 هستند
        if keyword in text_lower:
            return entry[1], entry[2], entry[3], entry[4]
            
    return "Unknown Host", "Unknown", "Unknown", "Unknown"

# ==========================================
# 5. اجرای اصلی
# ==========================================
print("Starting process with Deep Level Host Dictionary...")

all_summaries = []
all_unknown_hosts = []

for file_name in files_info:
    file_path = os.path.join(input_dir, file_name)
    
    if not os.path.exists(file_path):
        print(f"Warning: File not found: {file_name} -> Skipped")
        continue
        
    print(f"Processing: {file_name}")
    
    try:
        df = pd.read_excel(file_path)
        
        if 'Geo Loc Name' not in df.columns:
            print(f"   -> 'Geo Loc Name' missing")
            continue
        if 'Host' not in df.columns:
             df['Host'] = ""

        # 1. فیلتر کشور
        df['Derived_Country'] = df['Geo Loc Name'].apply(get_country_only)
        result_df = df.dropna(subset=['Derived_Country']).copy()
        
        if result_df.empty:
            continue

        # 2. تشخیص میزبان (چند سطحی)
        # خروجی تابع classify_host_deep یک تاپل 4 تایی است
        host_classification = result_df['Host'].apply(lambda x: classify_host_deep(x))
        
        # تبدیل تاپل‌ها به ستون‌های مجزا در دیتافریم
        result_df['Level_1'] = [x[0] for x in host_classification]
        result_df['Level_2'] = [x[1] for x in host_classification]
        result_df['Level_3'] = [x[2] for x in host_classification]
        result_df['Level_4'] = [x[3] for x in host_classification]
        
        # پر کردن مقادیر None با خط تیره برای زیبایی گزارش
        result_df.fillna(value={"Level_2": "-", "Level_3": "-", "Level_4": "-"}, inplace=True)

        # --- A: جمع‌آوری داده‌ها برای Summary ---
        summary_cols = ['Derived_Country', 'Level_1', 'Level_2', 'Level_3', 'Level_4']
        summary = result_df.groupby(summary_cols).size().reset_index(name='Count')
        summary['files_info'] = file_name
        all_summaries.append(summary)
        
        # --- B: جمع‌آوری Unknown Host ---
        unknowns_df = result_df[result_df['Level_1'] == 'Unknown Host'].copy()
        if not unknowns_df.empty:
            unknowns_df['files_info'] = file_name
            cols_to_keep = ['Geo Loc Name', 'Host', 'Derived_Country', 'files_info']
            valid_cols = [c for c in cols_to_keep if c in unknowns_df.columns]
            all_unknown_hosts.append(unknowns_df[valid_cols])
            
    except Exception as e:
        print(f"   -> Error processing {file_name}: {e}")

# ==========================================
# 6. ذخیره فایل نهایی
# ==========================================
print("\nPreparing final report...")

with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    
    found_paths_set = set()
    
    # --- Sheet 1: Summary ---
    if all_summaries:
        final_summary = pd.concat(all_summaries, ignore_index=True)
        # ستون‌ها را مرتب می‌کنیم
        final_summary = final_summary[['Derived_Country', 'Level_1', 'Level_2', 'Level_3', 'Level_4', 'Count', 'files_info']]
        final_summary.to_excel(writer, sheet_name='Summary', index=False)
        print(" -> Sheet 'Summary' created.")
        
        # استخراج مسیرهای پیدا شده برای مقایسه با کل
        # باید خط تیره ها را دوباره به None تبدیل کنیم تا با all_defined_paths قابل مقایسه باشند
        for _, row in final_summary.iterrows():
            # ساخت تاپل مسیر
            l1 = row['Level_1']
            l2 = row['Level_2'] if row['Level_2'] != '-' else None
            l3 = row['Level_3'] if row['Level_3'] != '-' else None
            l4 = row['Level_4'] if row['Level_4'] != '-' else None
            
            if l1 != 'Unknown Host':
                found_paths_set.add((l1, l2, l3, l4))
    else:
        pd.DataFrame({'Message': ['No data found']}).to_excel(writer, sheet_name='Summary')

    # --- Sheet 2: Unknown Host Details ---
    if all_unknown_hosts:
        final_unknowns = pd.concat(all_unknown_hosts, ignore_index=True)
        final_unknowns.to_excel(writer, sheet_name='Unknown_Host_Details', index=False)
        print(" -> Sheet 'Unknown_Host_Details' created.")
    else:
        pd.DataFrame({'Message': ['All Hosts identified']}).to_excel(writer, sheet_name='Unknown_Host_Details')

    # --- Sheet 3: Missing Categories ---
    # محاسبه اختلاف مجموعه‌ها
    missing_paths = list(all_defined_paths - found_paths_set)
    
    if missing_paths:
        missing_df = pd.DataFrame(missing_paths, columns=['Level_1', 'Level_2', 'Level_3', 'Level_4'])
        # پر کردن None ها با '-' برای زیبایی
        missing_df.fillna('-', inplace=True)
        # مرتب‌سازی
        missing_df = missing_df.sort_values(by=['Level_1', 'Level_2', 'Level_3', 'Level_4'])
        missing_df.to_excel(writer, sheet_name='Missing_Categories', index=False)
        print(" -> Sheet 'Missing_Categories' created.")
    else:
        pd.DataFrame({'Message': ['All defined categories were found!']}).to_excel(writer, sheet_name='Missing_Categories')

print(f"\nDone! File saved at:\n{output_path}")