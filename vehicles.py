"""
vehicles.py - Λίστα μαρκών και μοντέλων οχημάτων
"""

VEHICLES = {
    # ==================== ΑΥΤΟΚΙΝΗΤΑ ====================
    "Alfa Romeo": ["147", "156", "159", "166", "Giulia", "Giulietta", "MiTo", "Stelvio", "Tonale"],
    "Audi": ["A1", "A2", "A3", "A4", "A5", "A6", "A7", "A8", "Q2", "Q3", "Q4 e-tron", "Q5", "Q7", "Q8", "TT", "R8", "e-tron"],
    "BMW": ["116", "118", "120", "125", "316", "318", "320", "325", "330", "520", "523", "525", "530", "730", "740", "X1", "X2", "X3", "X4", "X5", "X6", "X7", "Z3", "Z4", "M3", "M5"],
    "Chevrolet": ["Aveo", "Captiva", "Cruze", "Epica", "Lacetti", "Matiz", "Spark", "Trax"],
    "Chrysler": ["300C", "Grand Voyager", "Neon", "PT Cruiser", "Sebring", "Voyager"],
    "Citroen": ["Berlingo", "C1", "C2", "C3", "C3 Aircross", "C4", "C4 Cactus", "C4 Picasso", "C5", "C5 Aircross", "C5 X", "C8", "DS3", "DS4", "DS5", "Jumper", "Jumpy", "Nemo", "Saxo", "Xsara", "Xsara Picasso"],
    "Dacia": ["Dokker", "Duster", "Jogger", "Lodgy", "Logan", "Sandero", "Spring"],
    "Daewoo": ["Kalos", "Lacetti", "Lanos", "Leganza", "Matiz", "Nubira"],
    "Fiat": ["500", "500L", "500X", "Barchetta", "Bravo", "Brava", "Croma", "Doblo", "Ducato", "Fiorino", "Grande Punto", "Idea", "Linea", "Marea", "Multipla", "Panda", "Punto", "Qubo", "Scudo", "Sedici", "Stilo", "Tipo", "Ulysse"],
    "Ford": ["B-Max", "C-Max", "EcoSport", "Edge", "Explorer", "Fiesta", "Focus", "Fusion", "Galaxy", "Grand C-Max", "Ka", "Kuga", "Maverick", "Mondeo", "Mustang", "Puma", "Ranger", "S-Max", "Transit", "Transit Connect", "Transit Custom"],
    "Honda": ["Accord", "City", "Civic", "CR-V", "CR-Z", "Fit", "HR-V", "Jazz", "Legend", "Odyssey", "Pilot", "Stream"],
    "Hyundai": ["Accent", "Atos", "Bayon", "Coupé", "Elantra", "Getz", "i10", "i20", "i30", "i40", "i45", "ioniq", "ix20", "ix35", "ix55", "Kona", "Matrix", "Santa Fe", "Sonata", "Terracan", "Trajet", "Tucson"],
    "Isuzu": ["D-Max", "Trooper"],
    "Jeep": ["Cherokee", "Commander", "Compass", "Grand Cherokee", "Patriot", "Renegade", "Wrangler"],
    "Kia": ["Carens", "Carnival", "Ceed", "Cerato", "EV6", "Joice", "Magentis", "Niro", "Optima", "Picanto", "ProCeed", "Rio", "Sorento", "Soul", "Sportage", "Stinger", "Stonic", "Venga", "XCeed"],
    "Lancia": ["Delta", "Lybra", "Musa", "Phedra", "Thesis", "Ypsilon"],
    "Land Rover": ["Defender", "Discovery", "Discovery Sport", "Freelander", "Range Rover", "Range Rover Evoque", "Range Rover Sport", "Range Rover Velar"],
    "Mazda": ["2", "3", "5", "6", "323", "626", "BT-50", "CX-3", "CX-5", "CX-7", "CX-9", "CX-30", "CX-60", "MX-5", "MPV", "Premacy", "RX-8", "Tribute"],
    "Mercedes": ["A 140", "A 150", "A 160", "A 170", "A 180", "A 200", "A 220", "B 150", "B 160", "B 170", "B 180", "B 200", "C 180", "C 200", "C 220", "C 250", "C 300", "CLA", "CLS", "E 200", "E 220", "E 250", "E 300", "E 350", "GLA", "GLB", "GLC", "GLE", "GLS", "ML", "S 300", "S 350", "S 400", "S 500", "SLK", "Sprinter", "Viano", "Vito"],
    "Mini": ["Cabrio", "Clubman", "Cooper", "Cooper S", "Countryman", "One", "Paceman"],
    "Mitsubishi": ["ASX", "Colt", "Eclipse Cross", "Galant", "Grandis", "L200", "Lancer", "Outlander", "Pajero", "Space Star"],
    "Nissan": ["350Z", "370Z", "Almera", "Ariya", "Juke", "Leaf", "Micra", "Murano", "Navara", "Note", "Pathfinder", "Patrol", "Pixo", "Primera", "Pulsar", "Qashqai", "Sentra", "Terrano", "Tiida", "X-Trail"],
    "Opel": ["Adam", "Agila", "Ampera", "Antara", "Astra", "Cascada", "Combo", "Corsa", "Crossland", "Grandland", "Insignia", "Kadett", "Meriva", "Mokka", "Movano", "Omega", "Signum", "Tigra", "Vectra", "Vivaro", "Zafira"],
    "Peugeot": ["107", "108", "206", "207", "208", "301", "306", "307", "308", "406", "407", "408", "3008", "4007", "4008", "5008", "Bipper", "Boxer", "Expert", "Ion", "Partner", "RCZ"],
    "Renault": ["Captur", "Clio", "Espace", "Express", "Fluence", "Grand Scenic", "Kadjar", "Kangoo", "Koleos", "Laguna", "Latitude", "Logan", "Master", "Megane", "Modus", "Scenic", "Symbol", "Talisman", "Trafic", "Twingo", "Zoe"],
    "Seat": ["Altea", "Arona", "Ateca", "Cordoba", "Exeo", "Ibiza", "Inca", "Leon", "Tarraco", "Toledo"],
    "Skoda": ["Citigo", "Fabia", "Kamiq", "Karoq", "Kodiaq", "Octavia", "Rapid", "Roomster", "Scala", "Superb", "Yeti"],
    "Smart": ["Forfour", "Fortwo"],
    "Subaru": ["Forester", "Impreza", "Legacy", "Outback", "Tribeca", "XV"],
    "Suzuki": ["Alto", "Baleno", "Grand Vitara", "Ignis", "Jimny", "Liana", "S-Cross", "Splash", "Swift", "SX4", "Vitara", "Wagon R"],
    "Tesla": ["Model 3", "Model S", "Model X", "Model Y"],
    "Toyota": ["Auris", "Avensis", "Aygo", "C-HR", "Camry", "Corolla", "FJ Cruiser", "GR Yaris", "Hilux", "Land Cruiser", "Picnic", "Prius", "ProAce", "RAV4", "Rush", "Supra", "Urban Cruiser", "Verso", "Yaris"],
    "Volkswagen": ["Amarok", "Arteon", "Beetle", "Caddy", "Caravelle", "CC", "Crafter", "Eos", "Golf", "ID.3", "ID.4", "ID.5", "Jetta", "Multivan", "Passat", "Phaeton", "Polo", "Scirocco", "Sharan", "T-Cross", "T-Roc", "Tiguan", "Touareg", "Touran", "Transporter", "Up"],
    "Volvo": ["C30", "C70", "S40", "S60", "S80", "S90", "V40", "V50", "V60", "V70", "V90", "XC40", "XC60", "XC70", "XC90"],

    # ==================== ΜΗΧΑΝΕΣ ====================
    "Honda Moto": ["CB 500", "CB 600 Hornet", "CB 1000R", "CBR 600RR", "CBR 1000RR", "Forza 300", "Forza 350", "Forza 750", "NC 700", "NC 750", "PCX 125", "SH 125", "SH 150", "SH 300", "VFR 800", "XL 1000V Varadero"],
    "Kawasaki": ["ER-6N", "ER-6F", "J125", "J300", "Ninja 300", "Ninja 400", "Ninja 650", "Ninja 1000", "Versys 650", "Versys 1000", "W800", "Z400", "Z650", "Z750", "Z900", "Z1000"],
    "Kymco": ["Agility 125", "Agility 150", "Downtown 125", "Downtown 300", "Grandvista 250", "Like 125", "People 125", "People 150", "People 200", "People 250", "People S 200", "Xciting 400", "Xciting 500"],
    "Piaggio": ["Beverly 125", "Beverly 300", "Beverly 350", "Beverly 500", "Liberty 125", "Liberty 150", "MP3 300", "MP3 400", "MP3 500", "Medley 125", "Medley 150", "Vespa GTS 125", "Vespa GTS 150", "Vespa GTS 300", "Vespa Primavera 125", "Vespa Sprint 125", "X-Evo 250", "X-Evo 400"],
    "SYM": ["Citycom 300", "Fiddle 125", "Joymax 250", "Joymax 300", "Orbit 125", "Symphony 125", "Symphony 150"],
    "Suzuki Moto": ["Address 110", "AN 400 Burgman", "AN 650 Burgman", "Bandit 650", "Bandit 1250", "DL 650 V-Strom", "DL 1000 V-Strom", "GSF 650", "GSR 600", "GSR 750", "GSX-R 600", "GSX-R 750", "GSX-R 1000", "GSX-S 750", "GSX-S 1000", "Hayabusa", "UH 125 Burgman", "UH 200 Burgman"],
    "Vespa": ["ET4 125", "GTS 125", "GTS 150", "GTS 250", "GTS 300", "LX 125", "LX 150", "Primavera 125", "Primavera 150", "S 125", "Sprint 125", "Sprint 150"],
    "Yamaha": ["FJR 1300", "FZ6", "FZ8", "MT-03", "MT-07", "MT-09", "MT-10", "NMAX 125", "NMAX 155", "R1", "R6", "T-Max 500", "T-Max 530", "T-Max 560", "Tracer 700", "Tracer 900", "X-Max 125", "X-Max 250", "X-Max 300", "X-Max 400", "XJ6", "XJR 1300", "XSR 700", "XSR 900"],
    "Bmw Moto": ["C 400 GT", "C 400 X", "C 600 Sport", "C 650 GT", "F 700 GS", "F 800 GS", "F 850 GS", "F 900 R", "R 1200 GS", "R 1250 GS", "S 1000 RR", "S 1000 XR"],
}

def get_markes():
    return sorted(VEHICLES.keys())

def get_montela(marka):
    return VEHICLES.get(marka, [])
