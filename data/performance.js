// Performance data for all PAC models and sizing calculations
// Sources: Atlantic catalog, ACV/Izea datasheets, NF EN 12831, ADEME/COSTIC

const PERFORMANCE_DATA = {

  // ============================================================
  // TBASE BY DEPARTMENT (NF EN 12831 / NF P 52-612/CN)
  // ============================================================
  tbase: {
    "01":{tbase:-10,zone:"G",climat:"H1a"},"02":{tbase:-7,zone:"D",climat:"H1c"},
    "03":{tbase:-8,zone:"E",climat:"H1c"},"04":{tbase:-8,zone:"E",climat:"H2d"},
    "05":{tbase:-10,zone:"G",climat:"H1c"},"06":{tbase:-6,zone:"A",climat:"H3"},
    "07":{tbase:-6,zone:"D",climat:"H2d"},"08":{tbase:-10,zone:"G",climat:"H1c"},
    "09":{tbase:-5,zone:"C",climat:"H2d"},"10":{tbase:-10,zone:"G",climat:"H1c"},
    "11":{tbase:-5,zone:"C",climat:"H3"},"12":{tbase:-8,zone:"E",climat:"H2d"},
    "13":{tbase:-5,zone:"C",climat:"H3"},"14":{tbase:-7,zone:"D",climat:"H1c"},
    "15":{tbase:-8,zone:"E",climat:"H1c"},"16":{tbase:-5,zone:"C",climat:"H2c"},
    "17":{tbase:-5,zone:"C",climat:"H2b"},"18":{tbase:-7,zone:"D",climat:"H2c"},
    "19":{tbase:-8,zone:"E",climat:"H1c"},"2A":{tbase:-2,zone:"A",climat:"H3"},
    "2B":{tbase:-2,zone:"A",climat:"H3"},"21":{tbase:-10,zone:"G",climat:"H1c"},
    "22":{tbase:-4,zone:"B",climat:"H2a"},"23":{tbase:-8,zone:"E",climat:"H1c"},
    "24":{tbase:-5,zone:"C",climat:"H2c"},"25":{tbase:-12,zone:"H",climat:"H1b"},
    "26":{tbase:-6,zone:"D",climat:"H2d"},"27":{tbase:-7,zone:"D",climat:"H1c"},
    "28":{tbase:-7,zone:"D",climat:"H1c"},"29":{tbase:-4,zone:"B",climat:"H2a"},
    "30":{tbase:-5,zone:"C",climat:"H3"},"31":{tbase:-5,zone:"C",climat:"H2d"},
    "32":{tbase:-5,zone:"C",climat:"H2d"},"33":{tbase:-5,zone:"C",climat:"H2c"},
    "34":{tbase:-5,zone:"C",climat:"H3"},"35":{tbase:-4,zone:"C",climat:"H2a"},
    "36":{tbase:-7,zone:"D",climat:"H2c"},"37":{tbase:-7,zone:"D",climat:"H2b"},
    "38":{tbase:-10,zone:"G",climat:"H1a"},"39":{tbase:-10,zone:"G",climat:"H1b"},
    "40":{tbase:-5,zone:"C",climat:"H2c"},"41":{tbase:-7,zone:"D",climat:"H2b"},
    "42":{tbase:-8,zone:"E",climat:"H1a"},"43":{tbase:-8,zone:"E",climat:"H1c"},
    "44":{tbase:-5,zone:"C",climat:"H2a"},"45":{tbase:-7,zone:"D",climat:"H1c"},
    "46":{tbase:-6,zone:"D",climat:"H2d"},"47":{tbase:-5,zone:"C",climat:"H2c"},
    "48":{tbase:-8,zone:"E",climat:"H2d"},"49":{tbase:-7,zone:"D",climat:"H2a"},
    "50":{tbase:-4,zone:"B",climat:"H2a"},"51":{tbase:-10,zone:"G",climat:"H1c"},
    "52":{tbase:-12,zone:"H",climat:"H1c"},"53":{tbase:-7,zone:"C",climat:"H2a"},
    "54":{tbase:-15,zone:"I",climat:"H1c"},"55":{tbase:-12,zone:"H",climat:"H1c"},
    "56":{tbase:-4,zone:"B",climat:"H2a"},"57":{tbase:-15,zone:"I",climat:"H1c"},
    "58":{tbase:-10,zone:"G",climat:"H1b"},"59":{tbase:-9,zone:"F",climat:"H1c"},
    "60":{tbase:-7,zone:"D",climat:"H1c"},"61":{tbase:-7,zone:"D",climat:"H1c"},
    "62":{tbase:-9,zone:"F",climat:"H1c"},"63":{tbase:-8,zone:"E",climat:"H1c"},
    "64":{tbase:-5,zone:"C",climat:"H2c"},"65":{tbase:-5,zone:"C",climat:"H2d"},
    "66":{tbase:-5,zone:"C",climat:"H3"},"67":{tbase:-15,zone:"I",climat:"H1c"},
    "68":{tbase:-15,zone:"I",climat:"H1c"},"69":{tbase:-10,zone:"G",climat:"H1a"},
    "70":{tbase:-12,zone:"H",climat:"H1b"},"71":{tbase:-10,zone:"G",climat:"H1b"},
    "72":{tbase:-7,zone:"D",climat:"H2b"},"73":{tbase:-10,zone:"G",climat:"H1a"},
    "74":{tbase:-10,zone:"G",climat:"H1a"},"75":{tbase:-5,zone:"D",climat:"H1c"},
    "76":{tbase:-7,zone:"D",climat:"H1c"},"77":{tbase:-7,zone:"D",climat:"H1c"},
    "78":{tbase:-7,zone:"D",climat:"H1c"},"79":{tbase:-7,zone:"C",climat:"H2b"},
    "80":{tbase:-9,zone:"F",climat:"H1c"},"81":{tbase:-5,zone:"C",climat:"H2d"},
    "82":{tbase:-5,zone:"C",climat:"H2d"},"83":{tbase:-5,zone:"A",climat:"H3"},
    "84":{tbase:-6,zone:"D",climat:"H2d"},"85":{tbase:-5,zone:"C",climat:"H2b"},
    "86":{tbase:-7,zone:"D",climat:"H2b"},"87":{tbase:-8,zone:"E",climat:"H1c"},
    "88":{tbase:-15,zone:"I",climat:"H1c"},"89":{tbase:-10,zone:"G",climat:"H1b"},
    "90":{tbase:-15,zone:"I",climat:"H1b"},"91":{tbase:-7,zone:"D",climat:"H1c"},
    "92":{tbase:-7,zone:"D",climat:"H1c"},"93":{tbase:-7,zone:"D",climat:"H1c"},
    "94":{tbase:-7,zone:"D",climat:"H1c"},"95":{tbase:-7,zone:"D",climat:"H1c"}
  },

  // Zone-Altitude correction matrix (NF P 52-612/CN)
  zoneAltitude: {
    "A": [-2,-4,-6,-8,-10,-12,-14,-16,-18,-20],
    "B": [-4,-5,-6,-7,-8,-9,-10,-12,-13,-14],
    "C": [-5,-6,-7,-8,-9,-10,-11,-12,-13,-14],
    "D": [-7,-8,-9,-11,-13,-14,-15,-17,-19,-21],
    "E": [-8,-9,-11,-13,-15,-17,-19,-21,-23,-25],
    "F": [-9,-10,-11,-12,-13,-15,-17,-19,-21,-23],
    "G": [-10,-11,-13,-14,-17,-19,-21,-23,-24,-25],
    "H": [-12,-13,-15,-17,-19,-21,-23,-24,-25,-27],
    "I": [-15,-15,-19,-21,-23,-24,-25,-27,-29,-31]
  },
  altitudeBands: [200,400,600,800,1000,1200,1400,1600,1800,2000],

  // ============================================================
  // PAC PERFORMANCE DATA
  // ============================================================

  // EFFIPAC (R32) - Atlantic - up to 60°C
  effipac: {
    models: [
      {code:"AHP60_14",nom:"Effipac 14",puissance_nom:14.10,refrigerant:"R32",t_max:60,chassis:"S",
       performance:{
         "A7/W35":{pcalo:14.10,pabs:2.91,cop:4.85},
         "A7/W45":{pcalo:14.41,pabs:3.63,cop:3.97},
         "A7/W55":{pcalo:13.44,pabs:4.35,cop:3.09},
         "A-7/W55":{pcalo:10.60,pabs:5.07,cop:2.09}
       }},
      {code:"AHP60_18",nom:"Effipac 18",puissance_nom:17.90,refrigerant:"R32",t_max:60,chassis:"S",
       performance:{
         "A7/W35":{pcalo:17.90,pabs:4.07,cop:4.40},
         "A7/W45":{pcalo:18.31,pabs:5.03,cop:3.64},
         "A7/W55":{pcalo:17.25,pabs:5.99,cop:2.88},
         "A-7/W55":{pcalo:12.30,pabs:6.03,cop:2.04}
       }},
      {code:"AHP60_26",nom:"Effipac 26",puissance_nom:26.00,refrigerant:"R32",t_max:60,chassis:"M",
       performance:{
         "A7/W35":{pcalo:26.00,pabs:6.44,cop:4.04},
         "A7/W45":{pcalo:26.65,pabs:7.98,cop:3.34},
         "A7/W55":{pcalo:25.10,pabs:9.51,cop:2.64},
         "A-7/W55":{pcalo:17.00,pabs:9.44,cop:1.80}
       }},
      {code:"AHP60_32",nom:"Effipac 32",puissance_nom:32.10,refrigerant:"R32",t_max:60,chassis:"M",
       performance:{
         "A7/W35":{pcalo:32.10,pabs:7.85,cop:4.09},
         "A7/W45":{pcalo:33.60,pabs:9.97,cop:3.37},
         "A7/W55":{pcalo:31.80,pabs:12.10,cop:2.63},
         "A-7/W55":{pcalo:21.70,pabs:11.92,cop:1.82}
       }},
      {code:"AHP60_50",nom:"Effipac 50",puissance_nom:50.20,refrigerant:"R32",t_max:60,chassis:"L",
       performance:{
         "A7/W35":{pcalo:50.20,pabs:12.21,cop:4.11},
         "A7/W45":{pcalo:51.34,pabs:15.10,cop:3.40},
         "A7/W55":{pcalo:48.30,pabs:18.02,cop:2.68},
         "A-7/W55":{pcalo:32.90,pabs:21.79,cop:1.51}
       }},
      {code:"AHP60_70",nom:"Effipac 70",puissance_nom:66.80,refrigerant:"R32",t_max:60,chassis:"XL",
       performance:{
         "A7/W35":{pcalo:66.80,pabs:16.29,cop:4.10},
         "A7/W45":{pcalo:67.37,pabs:20.05,cop:3.36},
         "A7/W55":{pcalo:61.90,pabs:23.80,cop:2.60},
         "A-7/W55":{pcalo:46.40,pabs:30.13,cop:1.54}
       }}
    ],
    maxCascade: 420, // kW max (6 units)
    maxUnits: 6
  },

  // APTAE (R290) - Atlantic/ACV Izea - up to 75°C
  aptae: {
    models: [
      {code:"AHP70_15",nom:"Aptae 15",puissance_nom:16.33,refrigerant:"R290",t_max:75,chassis:"S",
       poids:174, dimensions:{h:1380,l:1180,p:460},
       performance:{
         "A7/W35":{pcalo:16.33,pabs:3.31,cop:4.94},
         "A7/W45":{pcalo:15.50,pabs:3.88,cop:4.00},
         "A7/W55":{pcalo:14.80,pabs:4.65,cop:3.18},
         "A-7/W35":{pcalo:12.00,pabs:3.50,cop:3.43},
         "A-7/W55":{pcalo:10.50,pabs:4.90,cop:2.14}
       }},
      {code:"AHP70_18",nom:"Aptae 18",puissance_nom:18.72,refrigerant:"R290",t_max:75,chassis:"S",
       poids:174, dimensions:{h:1380,l:1180,p:460},
       performance:{
         "A7/W35":{pcalo:18.72,pabs:4.05,cop:4.62},
         "A7/W45":{pcalo:17.80,pabs:4.74,cop:3.76},
         "A7/W55":{pcalo:17.00,pabs:5.67,cop:3.00},
         "A-7/W35":{pcalo:13.80,pabs:4.25,cop:3.25},
         "A-7/W55":{pcalo:12.10,pabs:5.95,cop:2.03}
       }},
      {code:"AHP70_23",nom:"Aptae 23",puissance_nom:22.80,refrigerant:"R290",t_max:75,chassis:"M",
       poids:254, dimensions:{h:1380,l:1570,p:460},
       performance:{
         "A7/W35":{pcalo:22.80,pabs:4.78,cop:4.77},
         "A7/W45":{pcalo:21.70,pabs:5.58,cop:3.89},
         "A7/W55":{pcalo:20.70,pabs:6.68,cop:3.10},
         "A-7/W35":{pcalo:16.80,pabs:5.00,cop:3.36},
         "A-7/W55":{pcalo:14.70,pabs:7.00,cop:2.10}
       }},
      {code:"AHP70_27",nom:"Aptae 27",puissance_nom:27.00,refrigerant:"R290",t_max:75,chassis:"M",
       poids:264, dimensions:{h:1380,l:1570,p:460},
       performance:{
         "A7/W35":{pcalo:27.00,pabs:6.21,cop:4.35},
         "A7/W45":{pcalo:25.70,pabs:7.24,cop:3.55},
         "A7/W55":{pcalo:24.50,pabs:8.68,cop:2.82},
         "A-7/W35":{pcalo:19.90,pabs:6.50,cop:3.06},
         "A-7/W55":{pcalo:17.40,pabs:9.10,cop:1.91}
       }},
      {code:"AHP70_40",nom:"Aptae 40",puissance_nom:40.00,refrigerant:"R290",t_max:75,chassis:"L",
       poids:542, dimensions:{h:1680,l:2340,p:780},
       performance:{
         "A7/W35":{pcalo:40.00,pabs:9.76,cop:4.10},
         "A7/W45":{pcalo:38.00,pabs:11.40,cop:3.33},
         "A7/W55":{pcalo:36.20,pabs:13.66,cop:2.65},
         "A-7/W35":{pcalo:29.50,pabs:10.20,cop:2.89},
         "A-7/W55":{pcalo:25.80,pabs:14.30,cop:1.80}
       }},
      {code:"AHP70_50",nom:"Aptae 50",puissance_nom:50.00,refrigerant:"R290",t_max:75,chassis:"L",
       poids:560, dimensions:{h:1680,l:2340,p:780},
       performance:{
         "A7/W35":{pcalo:50.00,pabs:12.20,cop:4.10},
         "A7/W45":{pcalo:47.50,pabs:14.25,cop:3.33},
         "A7/W55":{pcalo:45.30,pabs:17.08,cop:2.65},
         "A-7/W35":{pcalo:36.80,pabs:12.75,cop:2.89},
         "A-7/W55":{pcalo:32.20,pabs:17.90,cop:1.80}
       }},
      {code:"AHP70_65",nom:"Aptae 65",puissance_nom:62.00,refrigerant:"R290",t_max:75,chassis:"XL",
       poids:650, dimensions:{h:1800,l:2500,p:850},
       performance:{
         "A7/W35":{pcalo:62.00,pabs:15.12,cop:4.10},
         "A7/W45":{pcalo:58.90,pabs:17.67,cop:3.33},
         "A7/W55":{pcalo:56.20,pabs:21.21,cop:2.65},
         "A-7/W35":{pcalo:45.70,pabs:15.81,cop:2.89},
         "A-7/W55":{pcalo:39.90,pabs:22.17,cop:1.80}
       }}
    ],
    maxCascade: 450, // kW max
    maxUnits: 6
  },

  // ============================================================
  // COMPETITORS
  // ============================================================
  competitors: {
    daikin: {
      brand: "Daikin",
      gamme: "Altherma 3 H HT",
      refrigerant: "R290",
      t_max: 70,
      models: [
        {nom:"EABH16DA9W",puissance_nom:16,
         performance:{
           "A7/W35":{pcalo:16.00,cop:4.60},
           "A7/W55":{pcalo:16.00,cop:3.10},
           "A-7/W35":{pcalo:16.00,cop:3.20},
           "A-7/W55":{pcalo:14.00,cop:2.00}
         }},
        {nom:"EABX16DA9W",puissance_nom:16,
         performance:{
           "A7/W35":{pcalo:16.00,cop:4.56},
           "A7/W55":{pcalo:16.00,cop:3.08}
         }},
        {nom:"Daikin Altherma 3 R ECH2O 500",puissance_nom:14,
         performance:{
           "A7/W35":{pcalo:14.50,cop:4.30},
           "A7/W55":{pcalo:14.50,cop:2.90}
         }}
      ],
      note: "Daikin propose la gamme Altherma pour le residentiel individuel et petit collectif. Pour le gros collectif, Daikin propose le VRV avec recup sur ECS."
    },
    mitsubishi: {
      brand: "Mitsubishi Electric",
      gamme: "Ecodan CAHV",
      refrigerant: "R744 (CO2)",
      t_max: 90,
      models: [
        {nom:"CAHV-P500YA-HPB",puissance_nom:45,
         performance:{
           "A7/W65":{pcalo:45.00,cop:3.80},
           "A7/W90":{pcalo:40.00,cop:2.80},
           "A-7/W65":{pcalo:36.00,cop:2.90},
           "A-7/W90":{pcalo:32.00,cop:2.20}
         }},
        {nom:"CAHV-P500YB-HPB",puissance_nom:50,
         performance:{
           "A7/W65":{pcalo:50.00,cop:3.90},
           "A7/W90":{pcalo:44.00,cop:2.85},
           "A-7/W65":{pcalo:40.00,cop:2.95},
           "A-7/W90":{pcalo:36.00,cop:2.25}
         }}
      ],
      note: "Mitsubishi Ecodan CAHV utilise le CO2 (R744), permettant des temperatures jusqu'a 90C. Ideal pour ECS collective. Cascade jusqu'a 300kW."
    }
  },

  // ============================================================
  // ECS SIZING DATA (ADEME/COSTIC)
  // ============================================================
  ecs: {
    // Logement standard = T3 social housing, 2.1 persons, 70 L/day at 60°C
    logementStandard: 70, // L/day at 60°C

    // Equivalence coefficients per dwelling type (New method - Parc social)
    equivalenceParc: {
      social: { T1:0.6, T2:0.7, T3:1.0, T4:1.4, T5:1.8, T6:1.9 },
      prive:  { T1:0.6, T2:0.7, T3:0.9, T4:1.1, T5:1.3, T6:1.4 }
    },

    // Peak volume formulas: Vp = a * Ns^b (litres at 60°C, Tef=9°C)
    peakFormulas: {
      "10min": {a:61, b:0.503},
      "1h":    {a:83, b:0.708},
      "2h":    {a:108, b:0.773},
      "3h":    {a:116, b:0.815},
      "4h":    {a:162, b:0.789},
      "5h":    {a:189, b:0.784},
      "6h":    {a:241, b:0.758},
      "7h":    {a:277, b:0.750},
      "8h":    {a:294, b:0.762}
    },

    // Non-residential consumption (L/unit/day at 60°C)
    tertiaire: {
      hotel_affaire:       {valeur:90, unite:"chambre", label:"Hotel affaire"},
      hotel_tourisme:      {valeur:135, unite:"chambre", label:"Hotel tourisme 3-4*"},
      hotel_montagne:      {valeur:185, unite:"chambre", label:"Hotel montagne"},
      restaurant_trad:     {valeur:15, unite:"repas", label:"Restaurant traditionnel"},
      restaurant_rapide:   {valeur:3, unite:"repas", label:"Restauration rapide"},
      restaurant_collectif:{valeur:7, unite:"repas", label:"Restauration collective"},
      ehpad:               {valeur:35, unite:"lit", label:"EHPAD"},
      hopital:             {valeur:55, unite:"lit", label:"Hopital"},
      maison_retraite:     {valeur:40, unite:"lit", label:"Maison de retraite"},
      residence_etudiante: {valeur:35, unite:"lit", label:"Residence etudiante"},
      foyer_jeunes:        {valeur:50, unite:"lit", label:"Foyer jeunes travailleurs"},
      caserne:             {valeur:50, unite:"lit", label:"Caserne"},
      bureau:              {valeur:4, unite:"personne", label:"Bureau"},
      camping:             {valeur:12, unite:"campeur", label:"Camping 3-4*"},
      sport_football:      {valeur:900, unite:"match", label:"Football (par match)"},
      sport_rugby:         {valeur:1250, unite:"match", label:"Rugby (par match)"},
      gymnase:             {valeur:448, unite:"heure", label:"Gymnase (par heure)"}
    },

    // Seasonal coefficients
    seasonal: {
      1:1.10, 2:1.10, 3:1.10, 4:1.10, 5:1.10,
      6:0.85, 7:0.75, 8:0.75,
      9:0.90, 10:1.05, 11:1.10, 12:1.10
    },

    // Legionella prevention
    legionella: {
      stockage_min: 55, // °C minimum
      distribution_min: 50, // °C at all points
      boost_temp: 60, // °C
      boost_duration: 30, // minutes
      robinet_max_sdb: 50, // °C max bathroom
      robinet_max_autre: 60 // °C max other
    }
  },

  // ============================================================
  // HEATING BIN DATA (typical hours per temperature bin)
  // ============================================================
  // Approximate bin hours for 3 climate zones (hours/year at each 1°C bin)
  binHours: {
    H1: { // e.g. Paris/Strasbourg
      "-20":1,"-19":2,"-18":4,"-17":6,"-16":10,"-15":15,"-14":22,"-13":30,
      "-12":40,"-11":52,"-10":65,"-9":80,"-8":100,"-7":120,"-6":145,"-5":170,
      "-4":200,"-3":230,"-2":265,"-1":300,"0":340,"1":380,"2":420,"3":460,
      "4":500,"5":540,"6":560,"7":570,"8":560,"9":530,"10":500,"11":460,
      "12":420,"13":380,"14":340,"15":300,"16":250,"17":200,"18":150,"19":80
    },
    H2: { // e.g. Nantes/Bordeaux
      "-15":2,"-14":4,"-13":6,"-12":10,"-11":15,"-10":22,"-9":30,"-8":42,
      "-7":55,"-6":72,"-5":92,"-4":115,"-3":142,"-2":175,"-1":210,"0":250,
      "1":295,"2":340,"3":385,"4":430,"5":470,"6":500,"7":520,"8":530,
      "9":520,"10":500,"11":470,"12":430,"13":385,"14":340,"15":295,"16":250,
      "17":200,"18":150,"19":100
    },
    H3: { // e.g. Marseille/Nice
      "-10":2,"-9":4,"-8":8,"-7":14,"-6":22,"-5":35,"-4":50,"-3":72,
      "-2":100,"-1":135,"0":175,"1":220,"2":270,"3":325,"4":385,"5":445,
      "6":500,"7":545,"8":580,"9":590,"10":580,"11":560,"12":530,"13":490,
      "14":445,"15":395,"16":340,"17":280,"18":215,"19":130
    }
  },

  // ============================================================
  // BUILDING TYPES (for quick heat loss estimation)
  // ============================================================
  buildingTypes: {
    passif:     {label:"Passif / BBC",      specific: 0.015, description:"< 15 W/m2"},
    re2020:     {label:"RE2020 / RT2012",   specific: 0.040, description:"~40 W/m2"},
    renove:     {label:"Renove (post-1990)",specific: 0.060, description:"~60 W/m2"},
    ancien:     {label:"Ancien (1970-1990)",specific: 0.080, description:"~80 W/m2"},
    tres_ancien:{label:"Tres ancien (<1970)",specific: 0.120, description:"~120 W/m2"}
  },

  // Ventilation R coefficients (W/m3.K for air renewal)
  ventilation: {
    vmc_auto:  {label:"VMC autoreglable",     R: 0.20},
    vmc_hygro_a:{label:"VMC hygroreglable A", R: 0.14},
    vmc_hygro_b:{label:"VMC hygroreglable B", R: 0.12}
  },

  // ============================================================
  // HYDRAULIC SIZING TABLES
  // ============================================================
  pipeDiameters: [
    {int:14, ext:16, maxFlow:150},
    {int:20, ext:22, maxFlow:380},
    {int:26, ext:28, maxFlow:770},
    {int:33, ext:35, maxFlow:1500},
    {int:40, ext:42, maxFlow:2450},
    {int:50, ext:54, maxFlow:4200},
    {int:66, ext:70, maxFlow:8000}
  ]
};

// Make available globally
if (typeof window !== 'undefined') window.PERFORMANCE_DATA = PERFORMANCE_DATA;
if (typeof module !== 'undefined') module.exports = PERFORMANCE_DATA;
