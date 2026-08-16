const backBtn = document.getElementById('backBtn');
const searchBtn = document.getElementById('searchBtn');
const mapBtn = document.getElementById('mapBtn');
const archiveBtn = document.getElementById('archiveBtn');
const addEquipmentBtn = document.getElementById('addEquipmentBtn');
const contentArea = document.getElementById('content-area');
const mapDiv = document.getElementById('map');
const addEquipmentModal = document.getElementById('addEquipmentModal');
const closeAddEquipmentModal = document.getElementById('closeAddEquipmentModal');
const addEquipmentForm = document.getElementById('addEquipmentForm');
const selectLocation = document.getElementById('selectLocation');
const dynamicContainer = document.getElementById('dynamicContainer');
const dynamicLabel = document.getElementById('dynamicLabel');
const selectDynamic = document.getElementById('selectDynamic');
const subDynamicContainer = document.getElementById('subDynamicContainer');
const subDynamicLabel = document.getElementById('subDynamicLabel');
const selectSubDynamic = document.getElementById('selectSubDynamic');
const equipmentNameInput = document.getElementById('equipmentName');
const equipmentQuantityInput = document.getElementById('equipmentQuantity');
const confirmModal = document.getElementById('confirmModal');
const confirmModalTitle = document.getElementById('confirmModalTitle');
const confirmModalMessage = document.getElementById('confirmModalMessage');
const confirmModalConfirmBtn = document.getElementById('confirmModalConfirmBtn');
const confirmModalCancelBtn = document.getElementById('confirmModalCancelBtn');
const syncStatus = document.getElementById('syncStatus');
const signInBtn = document.getElementById('signInBtn');
const passwordGate = document.getElementById('passwordGate');
const gateForm = document.getElementById('gateForm');
const gateUsernameInput = document.getElementById('gateUsername');
const gatePasswordInput = document.getElementById('gatePassword');
const gateError = document.getElementById('gateError');
const userGreeting = document.getElementById('userGreeting');
const switchUserBtn = document.getElementById('switchUserBtn');
const themeMenuBtn = document.getElementById('themeMenuBtn');
const themeMenu = document.getElementById('themeMenu');

// ----------------------------------------------------------------------------
// THEMES
// Every colour/radius/blur value in styles.css is read from a CSS variable,
// so switching themes is just toggling a class on <body> - the variable
// overrides in styles.css do the rest. Applied immediately (not waiting on
// DOMContentLoaded) so a returning visitor doesn't see a flash of the wrong
// theme before it kicks in.
// ----------------------------------------------------------------------------
const THEME_KEY = 'labInventoryTheme';
const THEME_CLASSES = ['theme-dark', 'theme-gecko'];

function applyTheme(theme) {
    document.body.classList.remove(...THEME_CLASSES);
    if (theme === 'dark' || theme === 'gecko') {
        document.body.classList.add(`theme-${theme}`);
    }
    document.querySelectorAll('.theme-option').forEach(btn => {
        btn.classList.toggle('active', btn.dataset.theme === theme);
    });
    localStorage.setItem(THEME_KEY, theme);
}

applyTheme(localStorage.getItem(THEME_KEY) || 'light');

// ----------------------------------------------------------------------------
// ACCESS GATE
// A lightweight username + shared-password screen, not real security (the
// password lives in this file, visible to anyone who views source) - it's a
// soft deterrent to keep casual/accidental visitors out, not a substitute
// for an actual login system. The Microsoft sign-in for shared-data sync is
// the real access control on the data itself.
// ----------------------------------------------------------------------------
const GATE_PASSWORD = 'JGHSScience';
const GATE_USER_KEY = 'labInventoryUser';
const GATE_AUTHED_KEY = 'labInventoryAuthed';

function setGreeting(username) {
    if (userGreeting) {
        userGreeting.textContent = `Hi, ${username}, what are you looking for?`;
    }
}

function startApp() {
    showWelcome();
    initSync();
}

function unlockApp(username) {
    localStorage.setItem(GATE_USER_KEY, username);
    localStorage.setItem(GATE_AUTHED_KEY, 'true');
    setGreeting(username);
    passwordGate.classList.add('hidden');
    startApp();
}

function checkGate() {
    const storedUser = localStorage.getItem(GATE_USER_KEY);
    const authed = localStorage.getItem(GATE_AUTHED_KEY) === 'true';
    if (authed && storedUser) {
        setGreeting(storedUser);
        passwordGate.classList.add('hidden');
        startApp();
    } else {
        passwordGate.classList.remove('hidden');
        gateUsernameInput.focus();
    }
}

// ----------------------------------------------------------------------------
// SHARED DATA CONFIG - Microsoft Graph / Excel Online
// Fill these in from your Azure app registration and SharePoint site (see
// the setup guide). Until you do, the app just runs with the built-in data
// on this device only - nothing breaks, it just won't be shared.
// ----------------------------------------------------------------------------
const MSAL_CLIENT_ID = 'PASTE_YOUR_AZURE_APP_CLIENT_ID_HERE';
const MSAL_TENANT_ID = 'PASTE_YOUR_AZURE_TENANT_ID_HERE';
// Path Graph uses to find the shared workbook, built from your SharePoint site's
// address. If your site is https://contoso.sharepoint.com/sites/ScienceDept,
// that becomes hostname "contoso.sharepoint.com" and site name "ScienceDept" below.
const GRAPH_SITE_HOSTNAME = 'PASTE_YOUR_SHAREPOINT_HOSTNAME_HERE';
const GRAPH_SITE_NAME = 'PASTE_YOUR_SHAREPOINT_SITE_NAME_HERE';
const GRAPH_FILE_PATH = 'LabInventory.xlsx'; // path to the file within that site's document library
const GRAPH_SHEET_NAME = 'InventoryData'; // worksheet name inside that file

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
const GRAPH_SCOPES = ['Files.ReadWrite'];
const POLL_INTERVAL_MS = 20000;

const SYNC_CONFIGURED = ![MSAL_CLIENT_ID, MSAL_TENANT_ID, GRAPH_SITE_HOSTNAME, GRAPH_SITE_NAME]
    .some(v => !v || v.startsWith('PASTE_YOUR'));

const GRAPH_WORKBOOK_URL = SYNC_CONFIGURED
    ? `${GRAPH_BASE}/sites/${GRAPH_SITE_HOSTNAME}:/sites/${GRAPH_SITE_NAME}:/drive/root:/${GRAPH_FILE_PATH}:/workbook/worksheets('${GRAPH_SHEET_NAME}')/range(address='A1')`
    : null;

const msalInstance = SYNC_CONFIGURED ? new msal.PublicClientApplication({
    auth: {
        clientId: MSAL_CLIENT_ID,
        authority: `https://login.microsoftonline.com/${MSAL_TENANT_ID}`,
        redirectUri: window.location.origin + window.location.pathname
    },
    cache: { cacheLocation: 'localStorage' }
}) : null;

let lastSyncedSnapshot = '';
let saveDebounceTimer = null;

let navigationStack = [];
let pendingAction = null;

let equipmentData = {
    'Lab 1': {
        cabinets: {
            'Cupboard A': [{ name: 'Paperwork', qty: 1 }, { name: 'Bread Makers', qty: 2 }]
        }
    },
    'Lab 5': {
        cabinets: {
            'Cupboard A': [{ name: '2 x 5kg weighing scales', qty: 2 }],
            'Cupboard B': [{ name: '7 x joule meters', qty: 7 }, { name: 'Trolleys', qty: 20 }],
            'Cupboard C': [{ name: 'DJB Data Loggers and Light Gates', qty: 20 }]
        }
    },
    'Lab 6': {
        cabinets: {
            'Cupboard A': [{ name: 'Loudspeakers', qty: 5 }, { name: 'Ruben’s tube', qty: 5 }, { name: 'Ramps', qty: 5 }],
            'Cupboard B': [{ name: 'C cores', qty: 20 }, { name: 'Microammeters', qty: 20 }, { name: 'Microwave kit', qty: 1 }, { name: 'Coils', qty: 20 }],
            'Cupboard D': [{ name: 'Ripple tank', qty: 1 }]
        }
    },
    'Lab 7': {
        cabinets: {
            'Cupboard A': [{ name: 'Oscilloscopes', qty: 10 }, { name: 'Steam engine', qty: 5 }, { name: 'Turntables', qty: 5 }, { name: 'Lenz’s law apparatus', qty: 5 }, { name: 'Large rheostats', qty: 20 }, { name: 'Ramps', qty: 5 }],
            'Cupboard B': [{ name: 'Signal generators', qty: 15 }],
            'Cupboard C': [{ name: 'Teletron tubes – Cathode Ray tube, Demonstration Triode, Double Beam tube, Electron Diffraction tube, Maltese Cross, Perrin tube', qty: 6 }, { name: 'stands for teletron tubes', qty: 2 }, { name: 'Helmholtz coils', qty: 2 }]
        }
    },
    'Greenhouse': {
        cabinets: {
            'Cabinet A': [{ name: 'Radioactive cupboard containing sources', qty: 1 }, { name: '“wooden box” Geiger Muller tube', qty: 1 }, { name: 'Laser', qty: 1 }, { name: 'EHT supply', qty: 2 }, { name: 'Scalar/timer', qty: 1 }],
            'Cabinet B': [{ name: 'Sodium Lamp with Transformer', qty: 5 }, { name: 'Travelling microscope', qty: 5 }, { name: 'Rayzer Boards with adapters', qty: 10 }, { name: 'Rectangular and semi-circular blocks', qty: 10 }, { name: 'Convex and concave lenses', qty: 10 }]
        }
    },
    'Bookstore': {
        basic: [{ name: 'Periodic Tables (poster)', qty: 20 }, { name: 'Linear Air Track', qty: 2 }, { name: 'Blowers', qty: 2 }, { name: 'Vehicles', qty: 5 }, { name: 'Pasco track and vehicles', qty: 5 }, { name: 'Conservation of momentum tray', qty: 1 }, { name: 'Oiled and unoiled hinges tray', qty: 1 }]
    },
    'Academy': {
        basic: [{ name: 'Petri dishes (sterile)', qty: 50 }, { name: 'Old manuals', qty: 10 }]
    },
    'Preproom': {
        subcategories: {
            'Biology': {
                cabinets: {
                    'Biology Column 01': [{ name: 'Syringes 5 + 10 ml (new)', qty: 10 }, { name: 'Syringes 20 ml (new + used)', qty: 5 }, { name: 'Syringes 5 + 10 ml (used)', qty: '50 packs' }, { name: 'Syringes 1 + 2.5 ml (new + used)', qty: '20 packs' }, { name: 'Petri Dishes new (Small and Large)', qty: '1 box' }, { name: 'Petri Dishes Small (used)', qty: '1 box' }, { name: 'Petri Dishes Regular (used)', qty: 15 }, { name: 'Petri Dishes (Glass)', qty: 10 }, { name: 'Syringes 0.25 - 12.5 ml', qty: 25 }],
                    'Biology Column 02': [{ name: 'Specimen Bottles', qty: 100 }, { name: 'Sample Bottles', qty: 50 }, { name: 'Plastic Pots', qty: '1 box' }, { name: 'Media Bottles', qty: '10 boxes' }, { name: 'Respiration Chamber Experiment', qty: '10 boxes' }, { name: 'Respiration Kit + Mouth Pieces', qty: 1 }, { name: 'Respirometer Demo Kit', qty: '5 bottles' }, { name: 'Glass Tube Attached to Thermometer', qty: '5 bottles' }, { name: 'Glass Tube Attached to Stoppers', qty: 500 }, { name: 'Spotting Tiles Plastic', qty: 1000 }, { name: 'Spotting Tiles Ceramic', qty: 2 }, { name: '96-well Plates', qty: 2 }, { name: 'Ceramic Tiles', qty: 1 }],
                    'Biology Column 03': [{ name: 'Plastic Sterile Loops', qty: 1 }, { name: 'Pipettes 1, 5, 10 ml & Fillers', qty: 3 }, { name: 'Pasteur Pipettes & Fillers', qty: 2 }, { name: 'Pipettes 10 ml', qty: 1 }, { name: 'Hand Blender', qty: 1 }, { name: 'Cork Borers', qty: 5 }, { name: 'DNA Bracelet Kit', qty: '1L' }, { name: 'DNA Bracelet Kit Materials', qty: '1L' }, { name: 'Lab Pipettes & Disposable Pipettes', qty: 3 }, { name: 'Potometer', qty: 5 }, { name: 'Loch Experiment', qty: 20 }, { name: 'A Simple Fermenter', qty: 15 }],
                    'Biology Column 04': [{ name: 'Sodium Alginate + Calcium Chloride + DCPIP', qty: 50 }, { name: 'Neutrase Bottles', qty: 30 }, { name: 'pH Buffers 2.8-13', qty: 25 }, { name: 'PEA\' Practical Mark + Recapture', qty: 10 }, { name: 'AH Unit 2 Genetic Drift Simulation', qty: 5 }, { name: 'Disinfectant + Discard Jar', qty: 5 }, { name: 'Muslin Cloth', qty: 2 }, { name: 'Garlic Root Experiment', qty: 1 }, { name: 'Flour', qty: 1 }, { name: 'Green Leaf Chromatography', qty: 1 }, { name: 'Bung + Tubing', qty: 3 }, { name: 'Pooters', qty: 10 }],
                    'Biology Column 05': [{ name: 'Taps', qty: 5 }, { name: 'Sodium Phosphate 0 - 0.3 M', qty: 10 }, { name: 'Hydrogen Peroxide Bottles', qty: 2 }, { name: 'Mirrors + Acetate Sheets', qty: 5 }, { name: 'Yeast + Flour + Sugar', qty: 'None' }, { name: 'Sugar Solutions + Resazurin Dye Bottles', qty: 'None' }, { name: 'Glucose + Starch + Amylase Bottles', qty: 'None' }, { name: 'Sand + Silver Sand', qty: 'None' }, { name: 'Starch Powder', qty: 'None' }, { name: 'Starch Solution + Bile Salts Bottles', qty: 'None' }, { name: 'Fabric Detergent + Stained Cloths', qty: 'None' }, { name: 'Mortar + Pestle', qty: 'None' }],
                    'Biology Column 06': [{ name: 'Glucose-1-phosphate bottles (handwritten note)', qty: 'None' }, { name: 'Methylene Blue', qty: 'None' }, { name: 'Liquid Paraffin + Phenolphthalein Indicator', qty: 'None' }, { name: 'Chromatography Solvent + Lipase Bottles', qty: 'None' }, { name: 'Detergent + Nail Polish + Manganese Dioxide', qty: 'None' }, { name: 'Limewater + Sodium Carbonate 0.05 M', qty: 'None' }, { name: 'Sodium Chloride', qty: 'None' }, { name: 'Chloroplast PCR Kit', qty: 'None' }],
                    'Biology Column 07': [{ name: 'Bubble Experiment', qty: 'None' }, { name: 'ELISA Kit', qty: 'None' }, { name: 'Humidity Tanks', qty: 'None' }, { name: 'Botrytis Kit', qty: 'None' }, { name: 'Seed Models', qty: 'None' }, { name: 'Rapid Cycling Brassica Kit', qty: 'None' }, { name: 'Butterflies + Moths', qty: 'None' }, { name: 'Woodlice Choice Chamber', qty: 'None' }, { name: 'Calcium Chloride + Soda Lime + Paint Brush', qty: 'None' }],
                    'Biology Column 08': [{ name: 'Woodlice Choice Chamber', qty: 'None' }, { name: 'Medical Examination Miscellaneous', qty: 'None' }, { name: 'Lung Model in Bell Jar', qty: 'None' }, { name: 'Brain Models', qty: 'None' }, { name: 'Heart Models', qty: 'None' }, { name: 'Skulls', qty: 'None' }, { name: 'Skulls (Teeth + Jaws)', qty: 'None' }, { name: 'Wood + Tree Sections', qty: 'None' }, { name: 'Thermos Flasks', qty: 'None' }],
                    'Biology Column 09': [{ name: 'Drosophila Kit', qty: 'None' }, { name: 'Food Calorimeter', qty: 'None' }, { name: 'Seeds', qty: 'None' }, { name: 'Seeds', qty: 'None' }, { name: 'Pipette Tips', qty: 'None' }, { name: 'Soil Test Kit', qty: 'None' }, { name: 'Meiosis Kit', qty: 'None' }, { name: 'Ear Model', qty: 'None' }],
                    'Biology Column 10': [{ name: 'Photosynthesis Kit', qty: 'None' }, { name: 'DNA Extraction Kit', qty: 'None' }, { name: 'Paperclip PCR', qty: 'None' }, { name: 'Amino Acid Paper Chromatography', qty: 'None' }, { name: 'Electrophoresis Kit', qty: 'None' }, { name: 'Electrophoresis Kit', qty: 'None' }, { name: 'Electrophoresis Kit (old)', qty: 'None' }, { name: 'Higher Bio Unit 1 Protein Synthesis', qty: 'None' }, { name: 'Cauliflower Cloning Experiment', qty: 'None' }],
                    'Biology Column 11': [{ name: 'Chromatography Tanks', qty: 'None' }, { name: 'Pond Nets + Double Thermometer', qty: 'None' }, { name: 'Sieves + Knitting Needles + Juice Extractors', qty: 'None' }, { name: 'Magnifying Glass', qty: 'None' }, { name: 'Peristalsis Kit', qty: 'None' }, { name: 'UV Sensitive Yeast Experiment', qty: 'None' }, { name: 'Sugar in Blood and Urine Experiment', qty: 'None' }],
                    'Biology Column 12': [{ name: 'Flower Models + Dissection Trays', qty: 'None' }, { name: 'Digital Thermometers', qty: 'None' }, { name: 'Light + Moisture Meters', qty: 'None' }, { name: 'Light + Moisture Meters', qty: 'None' }, { name: 'pH + Light + Moisture Meters', qty: 'None' }, { name: 'Trowels + Plastic Bags + Cups', qty: 'None' }, { name: 'Dissection Kits', qty: 'None' }, { name: 'Dissection Kit Spares', qty: 'None' }, { name: 'Dissection Kit Spares', qty: 'None' }, { name: 'Dissection Kit Plastic + New Scalpels', qty: 'None' }],
                    'Biology Column 13': [{ name: 'General Microbiology Equipment', qty: 'None' }, { name: 'Smokey Sue', qty: 'None' }, { name: 'Stethoscopes', qty: 'None' }, { name: 'Thermometers + Medical Thermometers', qty: 'None' }, { name: 'Heart Rate Monitor + Oximeters', qty: 'None' }, { name: 'Dynamometer', qty: 'None' }, { name: 'Body Composition Monitor Cards', qty: 'None' }, { name: 'Skinfold Callipers', qty: 'None' }, { name: 'Peak Flow Mouthpieces + Instructions', qty: 'None' }, { name: 'Peak Flow Mouthpieces (paper)', qty: 'None' }, { name: 'Spirometers + Peak Flow Meter', qty: 'None' }],
                    'Biology Column 14': [{ name: 'Blood Pressure Monitors', qty: 'None' }, { name: 'Stature Meters + Measuring Tape', qty: 'None' }, { name: 'Lung Volume Kits', qty: 'None' }, { name: 'Deflagrating Spoons', qty: 'None' }, { name: 'Random Laminates', qty: 'None' }, { name: 'Shells', qty: 'None' }, { name: 'Shells', qty: 'None' }, { name: 'Cuvettes', qty: 'None' }, { name: 'Finger Mazes', qty: 'None' }, { name: 'Serial Position Effect', qty: 'None' }, { name: 'Centrifuge + Eppendorfs', qty: 'None' }, { name: 'Eppendorfs', qty: 'None' }, { name: 'Universal Adapter', qty: 'None' }],
                    'Biology Column 15': [{ name: 'Centrifuge Tubes', qty: 'None' }, { name: 'Haemocytometers (+ Disposable)', qty: 'None' }, { name: 'Slides + Coverslips', qty: 'None' }, { name: 'Prepared Slides', qty: 'None' }, { name: 'Prepared Slides', qty: 'None' }, { name: 'Prepared Slides', qty: 'None' }, { name: 'Teaching Notes', qty: 'None' }, { name: 'Biobanking', qty: 'None' }, { name: 'Deflagrating Spoons', qty: 'None' }, { name: 'Photosynthesis Notes', qty: 'None' }],
                    'Biology Column 16': [{ name: 'Colorimeter', qty: 'None' }, { name: 'Glass Tubing w/ Rubber Tubing + Stoppers Attached', qty: 'None' }, { name: 'Burning Sugar Kit', qty: 'None' }, { name: 'Burning Food (just food)', qty: 'None' }, { name: 'Candies', qty: 'None' }, { name: 'Crispy Foods', qty: 'None' }, { name: 'Juice + More Food', qty: 'None' }],
                    'Biology Column 17': [{ name: 'Big Glass Jar + Crystallising Dish', qty: 'None' }, { name: 'PTC + Visking Tubing + Clinix', qty: 'None' }, { name: 'Drugs', qty: 'None' }, { name: 'Media + Albustix', qty: 'None' }, { name: 'Spaghetti + Food', qty: 'None' }, { name: 'Custard Bomb', qty: 'None' }],
                    'Biology Column 18': [{ name: 'Cooking Oil + Vinegar', qty: 'None' }, { name: 'Food', qty: 'None' }, { name: 'Milk Powder + Honey + Coffee', qty: 'None' }, { name: 'Flour + Sugar', qty: 'None' }, { name: 'Juice + Flours + Salt', qty: 'None' }, { name: 'Tea + Spices', qty: 'None' }, { name: 'Tea + Spices (small jars)', qty: 'None' }, { name: 'Cooking Oil + Vinegar', qty: 'None' }]
                }
            },
            'Chemistry': {
                cabinets: {
                    'Chemistry Column 01': [{ name: 'Wooden white tiles', qty: 'None' }, { name: 'Wide Neck Conical Flask', qty: 'None' }, { name: 'Narrow Neck Conical Flask', qty: 'None' }, { name: 'Graduated + Bulb Pipette (1 Ml)', qty: 'None' }, { name: 'Graduated + Bulb Pipette (2Ml)', qty: 'None' }, { name: 'Graduated + Bulb Pipette (5 Ml)', qty: 'None' }, { name: 'Graduated + Bulb Pipette (10 Ml)', qty: 'None' }, { name: 'Graduated + Bulb Pipette (20 + 25 Ml)', qty: 'None' }, { name: 'Distillation kit (Odd bits)', qty: 'None' }, { name: 'Distillation Kit (Large bits)', qty: 'None' }, { name: 'Distillation Kit', qty: 'None' }, { name: '3 Neck round bottom flask + Other distillation bits', qty: 'None' }, { name: 'Round bottom flask (50 + 100ml)', qty: 'None' }],
                    'Chemistry Column 02': [{ name: 'Quick Fit Distillation Kit', qty: 'None' }, { name: 'Quick Fit Distillation Kit', qty: 'None' }, { name: 'Quick Fit Distillation Kit', qty: 'None' }, { name: 'Quick Fit Distillation Kit', qty: 'None' }, { name: 'Quick Fit Distillation Kit', qty: 'None' }, { name: 'Quick Fit Distillation Kit', qty: 'None' }, { name: 'Quick Fit Distillation Kit', qty: 'None' }, { name: 'Quick Fit Distillation Kit', qty: 'None' }, { name: 'Quick Fit Distillation Kit', qty: 'None' }, { name: 'Pear Shaped flask with 2 neck + odd bits to fit distillation kit', qty: 'None' }, { name: 'Distillation Sets + Liebig condenser', qty: 'None' }, { name: 'Fractionating column', qty: 'None' }, { name: 'Part of quick for kit + conical flask + round bottom flask', qty: 'None' }, { name: 'Buchner funnel (Small)', qty: 'None' }],
                    'Chemistry Column 03': [{ name: 'Test Tubes + Boiling Tubes with Side arm', qty: 'None' }, { name: 'Buchner Funnel (Large)', qty: 'None' }, { name: 'Glass funnel with filter attached', qty: 'None' }, { name: 'Separating Conical flask (250 Ml)', qty: 'None' }, { name: 'Separating conical flask + Cylindrical flask with stoppers (500 Ml)', qty: 'None' }, { name: 'Conical flask with side arm (500 Ml)', qty: 'None' }, { name: 'Conical flask with side arm + tubing attached (250 + 500 Ml)', qty: 'None' }, { name: 'Conical flask with side arm + tubing attached (250 + 500 ml)', qty: 'None' }, { name: 'Filter paper (Different type)', qty: 'None' }, { name: 'Filter funnel + Buchner funnel filter Paper (Different type)', qty: 'None' }, { name: 'Delivery tube Glass (Right angle bend with stopper attached)', qty: 'None' }],
                    'Chemistry Column 04': [{ name: 'Delivery tube (With tubing + stopper attached)', qty: 'None' }, { name: 'Glass Deliver tube', qty: 'None' }, { name: 'Glass delivery tube', qty: 'None' }, { name: 'Thistle funnel + glass hollow pipe', qty: 'None' }, { name: 'Small Bunsen burner', qty: 'None' }, { name: 'Deflagrating spoon', qty: 'None' }, { name: 'Deflagrating spoon', qty: 'None' }, { name: 'Crucible', qty: 'None' }, { name: 'Porcelain boat', qty: 'None' }, { name: 'Porcelain + metal + Triangle with metal gauze', qty: 'None' }, { name: 'Spouted boiling tube with tubing attached', qty: 'None' }, { name: 'Spirit oil burners', qty: 'None' }, { name: 'Water Pumps', qty: 'None' }, { name: 'Large filter funnel (Plastic)', qty: 'None' }],
                    'Chemistry Column 05': [{ name: 'Copper cans', qty: 'None' }, { name: 'Product of combustion', qty: 'None' }, { name: 'Conical flask with side arm + Buchner funnel (250 Ml)', qty: 'None' }, { name: 'Delivery tubes', qty: 'None' }, { name: 'Thistle funnel with stopper', qty: 'None' }, { name: 'Stoppers (Solids) No.s 23, 25, 30.', qty: 'None' }, { name: 'Stoppers (Solids) No.s 13 (TT), 21 (BT)', qty: 'None' }, { name: 'Stoppers (Solids) No.s 11, 15, 17, 27, 29, 41, 49.', qty: 'None' }, { name: 'Stoppers (Single hole) No.s 10, 11, 15, 17, 19', qty: 'None' }, { name: 'Stoppers (Single hole) No.s 13 (TT) , 21 (BT)', qty: 'None' }, { name: 'Stoppers (Single hole) No.s 23, 25, 27, 29, 30.', qty: 'None' }, { name: 'Stopper (Double hole) No.s 15, 17, 19, 23, 25, 27, 29', qty: 'None' }, { name: 'Cork Borers (Assorted sizes)', qty: 'None' }, { name: 'Cork Borers (Assorted sizes)', qty: 'None' }, { name: 'Dessicator (Used for cobalt paper)', qty: 'None' }],
                    'Chemistry Column 06': [{ name: 'Thermometers (Various temperatures)', qty: 'None' }, { name: 'Dreschel bottle + Plain head + Long neck round bottom flask', qty: 'None' }, { name: 'Gas syringes', qty: 'None' }, { name: 'Boiling tubes with side arm', qty: 'None' }, { name: '‘U’ shaped tubes with side arm', qty: 'None' }, { name: 'Large filter funnel (Glass)', qty: 'None' }, { name: 'Glass filter funnel (Small)', qty: 'None' }, { name: 'Assorted glassware (Small)', qty: 'None' }, { name: 'Fermentation lock', qty: 'None' }, { name: 'Film canisters', qty: 'None' }, { name: 'Capillary Tube', qty: 'None' }],
                    'Chemistry Column 07': [{ name: 'Colorimeter', qty: 'None' }, { name: 'PH Metre', qty: 'None' }, { name: 'Magnetic Stirrers', qty: 'None' }, { name: 'Magnetic Stirrers', qty: 'None' }, { name: 'Magnetic Stirrers', qty: 'None' }, { name: 'PH Metre', qty: 'None' }, { name: 'Flea for Magnetic Stirrer', qty: 'None' }, { name: 'Odd Broken glass bits', qty: 'None' }, { name: 'Magnetic Stirrer', qty: 'None' }, { name: 'Protein amino acid kit', qty: 'None' }, { name: 'Odd Cables with plugs', qty: 'None' }, { name: 'Technicians guide for Chemistry + S2', qty: 'None' }],
                    'Chemistry Column 08': [{ name: 'Kosar’s Tray', qty: 'None' }, { name: 'Roman’s Tray', qty: 'None' }, { name: 'Hazards Labels tray', qty: 'None' }, { name: 'Polypockets + A4 Folders', qty: 'None' }, { name: 'Guoy + Almasa LCD display Balance', qty: 'None' }, { name: 'Tubing clamp + Tubing Adaptor', qty: 'None' }, { name: 'Lightning tapers', qty: 'None' }, { name: 'Lamp Holders 2.5V', qty: 'None' }],
                    'Chemistry Column 09': [{ name: 'Chemistry Handouts + Laminated cards', qty: 'None' }, { name: 'Assorted metals tray (Mg, Cu, Ni, Zn, in form of squares + strips. Iron nails)', qty: 'None' }, { name: 'Assorted metals tray', qty: 'None' }, { name: 'Assorted metals tray', qty: 'None' }, { name: 'Assorted metals tray', qty: 'None' }, { name: 'Citrus candle making kit', qty: 'None' }, { name: 'Metal Lids used for metals cans', qty: 'None' }, { name: 'Beer bottle tops', qty: 'None' }, { name: 'Leads Black colours', qty: 'None' }],
                    'Chemistry Column 10': [{ name: 'Leads Red colours', qty: 'None' }, { name: 'Assorted leads', qty: 'None' }, { name: 'Multimeter testing kit + Carbon rods + More leads', qty: 'None' }, { name: 'Hydrogen + Carbon electrodes', qty: 'None' }, { name: 'Crocodile clips', qty: 'None' }, { name: '‘D’ Plates + Crocodile clips in black holder', qty: 'None' }, { name: 'Plastic beaker electrolysis kit', qty: 'None' }, { name: 'Plastic beaker electrolysis kit', qty: 'None' }, { name: 'Plastic beaker electrolysis kit', qty: 'None' }, { name: 'Microchem kit + current direction meter', qty: 'None' }, { name: 'Alkali & Compound displays', qty: 'None' }],
                    'Chemistry Column 11': [{ name: 'Ionic Solution + Roman Compound Displays', qty: 'None' }, { name: 'Elements Display', qty: 'None' }, { name: 'Compounds Display', qty: 'None' }, { name: 'Minerals of Britain', qty: 'None' }, { name: 'Nichrome wire', qty: 'None' }, { name: 'Washable glue', qty: 'None' }, { name: 'Elastic Wire + Ideas & Suggestions', qty: 'None' }, { name: 'Polymorph', qty: 'None' }, { name: 'Polymer Kit + Memory wire', qty: 'None' }, { name: 'Photographic Paper + Golf Ball + Other Balls', qty: 'None' }, { name: 'Molymods (4 sets in a tub)', qty: 'None' }],
                    'Chemistry Column 12': [{ name: 'Molymods (4 sets in a tub)', qty: 'None' }, { name: 'Molymods (4 sets in a tub)', qty: 'None' }, { name: 'Molymods (4 sets in a tub)', qty: 'None' }, { name: 'Molymods (4 sets in a tub)', qty: 'None' }, { name: 'Molymods (4 sets in a tub)', qty: 'None' }, { name: 'Molymods top up kit', qty: 'None' }, { name: 'Molymods top up kit', qty: 'None' }, { name: 'Molymods top up kit', qty: 'None' }, { name: 'Wooden blocks', qty: 'None' }, { name: 'Molymods molecular model', qty: 'None' }, { name: 'Polystyrene balls', qty: 'None' }, { name: 'Compound Cards', qty: 'None' }, { name: 'Different material blocks + cubes', qty: 'None' }, { name: 'Wooden cones + squares blocks', qty: 'None' }, { name: 'Burner wicks', qty: 'None' }]
                }
            },
            'Physics': {
                cabinets: {
                    'Physics Column 01': [{ name: 'Fuses', qty: 'None' }, { name: 'Resistors', qty: 'None' }, { name: 'Resistors, LEDs', qty: 'None' }, { name: 'Lamps, fuses', qty: 'None' }, { name: 'Capacitors', qty: 'None' }, { name: 'Lamps', qty: 'None' }, { name: 'Lamps', qty: 'None' }, { name: 'Masses', qty: 'None' }, { name: 'Masses', qty: 'None' }, { name: 'Phenaskistoscopes', qty: 'None' }],
                    'Physics Column 02': [{ name: 'Variable voltage boards, thermistors', qty: 'None' }, { name: 'Accessories for variable voltage boards', qty: 'None' }, { name: 'Accessories for variable voltage boards', qty: 'None' }, { name: 'Phototronics explorer', qty: 'None' }, { name: 'Electromagnetic radiation cards – Higher', qty: 'None' }, { name: 'LDRs', qty: 'None' }, { name: 'Centre of gravity experiment', qty: 'None' }, { name: 'Vernier callipers', qty: 'None' }, { name: 'Micrometers', qty: 'None' }, { name: 'Long leads, semaphore flags etc', qty: 'None' }, { name: 'Telephones', qty: 'None' }, { name: 'Phenaskistoscopes', qty: 'None' }, { name: 'Masses', qty: 'None' }, { name: 'Masses', qty: 'None' }, { name: 'Wheatstone bridges', qty: 'None' }],
                    'Physics Column 03': [{ name: 'Wheatstone bridges', qty: 'None' }, { name: 'Potential divider boards', qty: 'None' }, { name: 'Momentum and impulse golf club experiment', qty: 'None' }, { name: 'UV lamps and UV meters, Researching Physics unit', qty: 'None' }, { name: 'UV lamps and UV meters, Researching Physics unit', qty: 'None' }, { name: 'DVD cases, vibration detectors, sun cream, pulleys, Researching Physics unit', qty: 'None' }, { name: 'Pressure laws equipment', qty: 'None' }, { name: 'Charles’ laws equipment', qty: 'None' }, { name: 'Hydrometers', qty: 'None' }, { name: 'Masses', qty: 'None' }, { name: 'Capacitor display boards', qty: 'None' }],
                    'Physics Column 04': [{ name: 'Electric fields apparatus, poppers', qty: 'None' }, { name: 'Microwave apparatus', qty: 'None' }, { name: 'Microwave apparatus', qty: 'None' }, { name: 'Piezo-electric crystals, brass blocks', qty: 'None' }, { name: 'Eureka cans', qty: 'None' }, { name: 'Eureka cans', qty: 'None' }, { name: 'Heating elements, long thermometers', qty: 'None' }, { name: 'Metal blocks', qty: 'None' }, { name: 'Trundle wheel', qty: 'None' }],
                    'Physics Column 05': [{ name: 'Mountain', qty: 'None' }, { name: 'Compasses, rope, chalk, long tape measure', qty: 'None' }, { name: 'Doppler rocket and accessories', qty: 'None' }, { name: 'Mechanical stop clocks, digital tape measure', qty: 'None' }, { name: 'Vibrators', qty: 'None' }, { name: 'Projectile apparatus', qty: 'None' }, { name: 'Tin cans', qty: 'None' }, { name: 'Tin cans', qty: 'None' }, { name: 'Plastic beakers, 1kg masses', qty: 'None' }, { name: 'G-clamps', qty: 'None' }, { name: 'Newton balances', qty: 'None' }],
                    'Physics Column 06': [{ name: 'Newton balances', qty: 'None' }, { name: 'Newton balances', qty: 'None' }, { name: 'Newton balances', qty: 'None' }, { name: 'Pulleys', qty: 'None' }, { name: 'Pulleys', qty: 'None' }, { name: 'Springs', qty: 'None' }, { name: 'Springs', qty: 'None' }, { name: 'Bench pulleys', qty: 'None' }, { name: 'Resistance boxes', qty: 'None' }, { name: 'Boards with fixed points', qty: 'None' }, { name: 'Calorimeter cans', qty: 'None' }, { name: 'Laboratory jacks', qty: 'None' }, { name: 'Young’s modulus', qty: 'None' }, { name: 'Box with polystyrene bits', qty: 'None' }],
                    'Physics Column 07': [{ name: 'Trolleys', qty: 'None' }, { name: 'Trolley accessories', qty: 'None' }, { name: 'Trolleys', qty: 'None' }, { name: 'Trolleys', qty: 'None' }, { name: 'Trolleys', qty: 'None' }, { name: 'Trolleys', qty: 'None' }, { name: 'Trolleys', qty: 'None' }, { name: 'Perspex shapes', qty: 'None' }],
                    'Physics Column 08': [{ name: 'Accessories for grey meters', qty: 'None' }, { name: 'Accessories for grey meters', qty: 'None' }, { name: 'Accessories for grey meters', qty: 'None' }, { name: 'Accessories for grey meters', qty: 'None' }, { name: 'Grey meters', qty: 'None' }, { name: 'Digital meters', qty: 'None' }, { name: 'Accessories for digital meters', qty: 'None' }, { name: 'Digital meters', qty: 'None' }, { name: 'Digital meters', qty: 'None' }, { name: 'Grey meters', qty: 'None' }, { name: 'Optical kits with mirrors', qty: 'None' }],
                    'Physics Column 09': [{ name: 'X-rays', qty: 'None' }, { name: 'Glass blocks', qty: 'None' }, { name: 'Glass blocks', qty: 'None' }, { name: 'Glass blocks', qty: 'None' }, { name: 'Glass blocks, lenses', qty: 'None' }, { name: 'Lenses', qty: 'None' }, { name: 'Ray boxes', qty: 'None' }, { name: 'Ray boxes', qty: 'None' }, { name: 'Ray boxes', qty: 'None' }, { name: 'Bits for ray boxes', qty: 'None' }, { name: 'Attachments for optical benches', qty: 'None' }, { name: 'Digital meters', qty: 'None' }, { name: 'Digital meters', qty: 'None' }, { name: 'Whoosh of air apparatus', qty: 'None' }],
                    'Physics Column 10': [{ name: 'Very large magnifying lenses and holders', qty: 'None' }, { name: 'Glass blocks', qty: 'None' }, { name: 'Mirrors', qty: 'None' }, { name: 'Optical fibres, glass rods, 250ml measuring cylinder, chalk, Perspex block for laser', qty: 'None' }, { name: 'UV lamps', qty: 'None' }, { name: '12V mounted lamps', qty: 'None' }, { name: 'Sensors – light, magnetic flux, infra-red, distance', qty: 'None' }, { name: 'Sodium flame pencils, Newtons ring apparatus, hologram, adjustable slit , slits', qty: 'None' }, { name: 'CDs, diffraction gratings, colour mixing set, glass plates for interference, Moiré’s fringes', qty: 'None' }, { name: 'Mirrors', qty: 'None' }, { name: 'Filters, large camera lenses', qty: 'None' }, { name: 'Large lamps', qty: 'None' }, { name: 'Prisms', qty: 'None' }, { name: 'Accessories for Monkey and Hunter experiment', qty: 'None' }],
                    'Physics Column 11': [{ name: 'Space rocket with bottles and copper wire', qty: 'None' }, { name: 'Strobe light', qty: 'None' }, { name: 'Accessories for strobe light, tuning forks', qty: 'None' }, { name: 'Marbles, ball bearings, polystyrene balls etc', qty: 'None' }, { name: 'Tennis balls, table tennis balls, straws, super balls', qty: 'None' }, { name: 'Radiant heaters, halogen lamp for model eye', qty: 'None' }, { name: 'Radios, tape recorders and microphones', qty: 'None' }, { name: 'Pasco extras, projectile apparatus', qty: 'None' }],
                    'Physics Column 12': [{ name: 'TSA timers', qty: 'None' }, { name: 'TSA timers', qty: 'None' }, { name: 'TSA timers adaptors', qty: 'None' }, { name: 'Lights gates, receivers, nuts and bolts for attaching masks to trolleys', qty: 'None' }, { name: 'Masks, blutak, elastics for trolleys, thread, 10g masses, pulley for Pasco track', qty: 'None' }, { name: 'Laptop and charger', qty: 'None' }, { name: 'USB connector, motion sensor, force sensor, voltage/current sensor', qty: 'None' }, { name: 'Fizz pop rocket equipment', qty: 'None' }, { name: 'Tin cans, large coffee tin “g” apparatus', qty: 'None' }, { name: 'Double pan balances', qty: 'None' }, { name: 'Sodium lamp', qty: 'None' }],
                    'Physics Column 13': [{ name: 'Wooden blocks with sandpaper etc', qty: 'None' }, { name: 'Wind-up animals', qty: 'None' }, { name: 'Stearic acid/Salol experiments', qty: 'None' }, { name: 'Radio kits', qty: 'None' }, { name: 'Water pumps', qty: 'None' }, { name: 'Coils of wire – 20,40,60', qty: 'None' }, { name: 'Coils 60:60, 500:125, clips and c-cores', qty: 'None' }, { name: 'Transmission line demonstration', qty: 'None' }, { name: 'Properties of matter Higher equipment', qty: 'None' }],
                    'Physics Column 14': [{ name: 'Magnets', qty: 'None' }, { name: 'Magnets', qty: 'None' }, { name: 'Electronics trolley – power packs', qty: 'None' }, { name: 'Electronics trolley – power packs', qty: 'None' }, { name: 'Electronics trolley', qty: 'None' }, { name: 'Model eye – flask in Chemistry prep room', qty: 'None' }],
                    'Physics Column 15': [{ name: 'Electronics trolley – guide in drawers', qty: 'None' }, { name: 'Electronics trolley', qty: 'None' }, { name: 'Electronics trolley', qty: 'None' }, { name: 'Electronics trolley', qty: 'None' }, { name: 'Electronics trolley', qty: 'None' }, { name: 'Electronics trolley', qty: 'None' }, { name: 'Electronics trolley', qty: 'None' }, { name: 'Electronics trolley', qty: 'None' }, { name: 'Electronics trolley', qty: 'None' }, { name: 'Electronics trolley', qty: 'None' }, { name: 'Electronics trolley', qty: 'None' }, { name: 'Electronics trolley including large LDRs and thermocouples', qty: 'None' }, { name: 'Instructions A-N', qty: 'None' }],
                    'Physics Column 16': [{ name: 'Instructions O-Z', qty: 'None' }, { name: 'Resistor holders, switches, two-way switches, lamps', qty: 'None' }, { name: 'Tools', qty: 'None' }, { name: 'Cells', qty: 'None' }, { name: 'Leads and crocodile clips', qty: 'None' }, { name: 'Leads', qty: 'None' }, { name: 'Old Standard Grade booklets', qty: 'None' }, { name: 'Cable', qty: 'None' }, { name: 'Resistors', qty: 'None' }],
                    'Physics Column 17': [{ name: 'Resistors', qty: 'None' }, { name: 'Resistors', qty: 'None' }, { name: 'Resistors', qty: 'None' }, { name: 'Resistors', qty: 'None' }, { name: 'Resistors', qty: 'None' }, { name: 'Resistors', qty: 'None' }, { name: 'Resistors', qty: 'None' }, { name: 'Resistors', qty: 'None' }, { name: 'Resistors', qty: 'None' }, { name: 'Resistors', qty: 'None' }, { name: 'Diodes', qty: 'None' }, { name: 'Resistors – unknown, resistor charts, hand lenses', qty: 'None' }, { name: 'Variable resistors', qty: 'None' }, { name: 'Variable resistors', qty: 'None' }, { name: 'Odd/large capacitors', qty: 'None' }],
                    'Physics Column 18': [{ name: 'Capacitors', qty: 'None' }, { name: 'Capacitors', qty: 'None' }, { name: 'Capacitors', qty: 'None' }, { name: 'Capacitors', qty: 'None' }, { name: 'Capacitors', qty: 'None' }, { name: 'Capacitors', qty: 'None' }, { name: 'Capacitors', qty: 'None' }, { name: 'Large yellow holders for cells', qty: 'None' }, { name: 'Short leads, leads with jack plugs', qty: 'None' }, { name: 'Speed of sound equipment', qty: 'None' }],
                    'Physics Column 19': [{ name: 'Morse code equipment', qty: 'None' }, { name: 'Morse code equipment', qty: 'None' }, { name: 'Buzzers, candles, rattling eggs etc', qty: 'None' }, { name: 'Slinkys', qty: 'None' }, { name: 'Sound meters', qty: 'None' }, { name: 'Sound meters, large', qty: 'None' }, { name: 'ear defenders', qty: 'None' }, { name: 'Loudspeaker making equipment', qty: 'None' }, { name: 'Brownian Motion apparatus', qty: 'None' }, { name: 'Wire cutters, 3 core cable', qty: 'None' }],
                    'Physics Column 20': [{ name: 'Series circuit boards', qty: 'None' }, { name: 'Parallel circuit boards', qty: 'None' }, { name: 'Electrical plugs', qty: 'None' }, { name: 'Screwdrivers, shake flashlight, hand generator torch', qty: 'None' }, { name: 'Electric motors', qty: 'None' }, { name: 'Fred model, Megger, conductivity experiment', qty: 'None' }, { name: 'Plug wiring kit', qty: 'None' }, { name: 'Appliances with faults', qty: 'None' }, { name: 'Appliances with faults', qty: 'None' }, { name: 'Brass terminals, panel pins, wire', qty: 'None' }],
                    'Physics Column 21': [{ name: 'Iron rods, leads, paper clips', qty: 'None' }, { name: 'Leads, electric motors', qty: 'None' }, { name: 'Electric motor kit', qty: 'None' }, { name: 'Electric motor kit', qty: 'None' }, { name: 'Electric motor kit', qty: 'None' }, { name: 'Mains ammeter, iron rods', qty: 'None' }, { name: 'Bells', qty: 'None' }, { name: 'Tray of electromagnetic equipment', qty: 'None' }, { name: '12V lamps – different wattages', qty: 'None' }, { name: 'Plotting compasses, reed switches, reed coils, conductor/insulator kits, paper clips, electromagnetic relay', qty: 'None' }, { name: 'Circuit breakers, fuse holders, metre lengths of wire', qty: 'None' }, { name: 'Elements with ratings', qty: 'None' }, { name: 'Electric motor', qty: 'None' }, { name: 'Sonometers', qty: 'None' }],
                    'Next to Column 7': [{ name: 'Metre bridges', qty: 'None' }, { name: 'Bicycle pump', qty: 'None' }],
                    'Big tub next to Column 16': [{ name: 'Cables, large, optical fibre', qty: 'None' }, { name: 'Free fall tube, Guinea and Feather apparatus', qty: 'None' }, { name: 'Optical benches', qty: 'None' }, { name: 'Rocket launcher', qty: 'None' }, { name: 'Stands for Monkey and Hunter experiment', qty: 'None' }, { name: 'Sulfuric Acid (conc)', qty: 'None' }]
                }
            },
            'Chemical Store': {
                'C1': [],
                'C2': [],
                'C3': [],
                'C4': [],
                'Shelves': {
                    'Main Shelf 1': [],
                    'Main Shelf 2': []
                }
            }
        }
    },
    'Lab 2': {},
    'Lab 3': {},
    'Lab 4': {},
    'Lab 8': {},
    'Lab 9': {},
    'Lab 10': {}
};

let standardBasicEquipment = [{ name: 'Goggles', qty: 21 }, { name: 'Bunsen Burners', qty: 10 }, { name: 'Heating Mats', qty: 10 }, { name: 'Tripods', qty: 10 }, { name: 'Clampstands', qty: 10 }];
let standardGeneralLabware = [{ name: '250ml Beakers', qty: 10 }, { name: '100ml Beakers', qty: 10 }, { name: '100ml Measuring Cylinders', qty: 10 }, { name: 'Spatulas', qty: 10 }, { name: 'Stirring Rods', qty: 10 }, { name: 'Funnels', qty: 10 }, { name: 'Wooden Tongs', qty: 10 }, { name: 'Metal Tongs', qty: 10 }, { name: 'Beaker Holders', qty: 10 }, { name: 'Test Tubes', qty: '2 tubs' }, { name: 'Boiling Tubes', qty: '2 tubs' }];

let archivedEquipment = [];

// ----------------------------------------------------------------------------
// SHARED EXCEL ONLINE SYNC (Microsoft Graph)
// The whole inventory travels as one JSON blob in cell A1 of the
// "InventoryData" worksheet in a shared Excel file on your SharePoint site.
// Each person signs in with their own Microsoft account (Files.ReadWrite),
// then reads/writes that cell directly via the Graph API - no middle-man
// server needed. See the setup guide for how to register the Azure app and
// point GRAPH_SITE_HOSTNAME/GRAPH_SITE_NAME at your file.
// ----------------------------------------------------------------------------

function setSyncStatus(text, isError = false) {
    if (!syncStatus) return;
    syncStatus.textContent = text;
    syncStatus.classList.toggle('sync-error', isError);
}

function currentDataSnapshot() {
    return JSON.stringify({ equipmentData, archivedEquipment, standardBasicEquipment, standardGeneralLabware });
}

function applySharedData(data) {
    equipmentData = data.equipmentData;
    archivedEquipment = data.archivedEquipment || [];
    standardBasicEquipment = data.standardBasicEquipment || standardBasicEquipment;
    standardGeneralLabware = data.standardGeneralLabware || standardGeneralLabware;
}

// Only ever called from a real button click (see signInBtn handler below) -
// loginPopup must run in direct response to a user gesture or the browser
// will block the popup.
async function ensureSignedIn() {
    try {
        const result = await msalInstance.loginPopup({ scopes: GRAPH_SCOPES });
        msalInstance.setActiveAccount(result.account);
        return true;
    } catch (err) {
        console.error('Microsoft sign-in failed:', err);
        setSyncStatus('Sign-in failed - try again', true);
        return false;
    }
}

// Silent-only token fetch, safe to call from background code (polling, auto-save).
// Never opens a popup itself - if the cached session is gone, it asks the
// person to click "Sign in" again rather than surprising them with a popup.
async function getGraphToken() {
    if (!SYNC_CONFIGURED) return null;
    const account = msalInstance.getActiveAccount() || msalInstance.getAllAccounts()[0];
    if (!account) return null;
    try {
        const result = await msalInstance.acquireTokenSilent({ scopes: GRAPH_SCOPES, account });
        return result.accessToken;
    } catch (err) {
        console.error('Silent token acquisition failed:', err);
        setSyncStatus('Signed out - click "Sign in to sync"', true);
        signInBtn.classList.remove('hidden');
        return null;
    }
}

async function loadFromSheet() {
    if (!SYNC_CONFIGURED) {
        setSyncStatus('Local only (sharing not set up)');
        return;
    }
    const token = await getGraphToken();
    if (!token) return;
    setSyncStatus('Loading shared inventory…');
    try {
        const res = await fetch(GRAPH_WORKBOOK_URL, { headers: { Authorization: `Bearer ${token}` } });
        if (!res.ok) throw new Error(`Graph GET ${res.status}`);
        const result = await res.json();
        const raw = result.values && result.values[0] && result.values[0][0];
        const data = raw ? JSON.parse(raw) : null;
        if (data && data.equipmentData) {
            applySharedData(data);
            lastSyncedSnapshot = currentDataSnapshot();
            setSyncStatus('Synced ✓');
        } else {
            // Sheet cell is empty - first person to sign in seeds it with the built-in defaults.
            await saveToSheet();
        }
    } catch (err) {
        console.error('Failed to load shared inventory:', err);
        setSyncStatus('Offline - showing local copy', true);
    }
}

async function saveToSheet() {
    if (!SYNC_CONFIGURED) return;
    const token = await getGraphToken();
    if (!token) return;
    setSyncStatus('Saving…');
    const payload = currentDataSnapshot();
    try {
        const res = await fetch(GRAPH_WORKBOOK_URL, {
            method: 'PATCH',
            headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
            body: JSON.stringify({ values: [[payload]] })
        });
        if (!res.ok) throw new Error(`Graph PATCH ${res.status}`);
        lastSyncedSnapshot = payload;
        setSyncStatus('Saved ✓');
    } catch (err) {
        console.error('Failed to save to shared inventory:', err);
        setSyncStatus('Sync failed - saved locally only', true);
    }
}

function scheduleSave() {
    if (!SYNC_CONFIGURED) return;
    clearTimeout(saveDebounceTimer);
    setSyncStatus('Saving…');
    saveDebounceTimer = setTimeout(saveToSheet, 700);
}

function startPolling() {
    if (!SYNC_CONFIGURED) return;
    setInterval(async () => {
        // Don't yank data out from under someone mid-edit.
        if (confirmModal.style.display === 'flex' || addEquipmentModal.style.display === 'flex') return;
        const token = await getGraphToken();
        if (!token) return;
        try {
            const res = await fetch(GRAPH_WORKBOOK_URL, { headers: { Authorization: `Bearer ${token}` } });
            if (!res.ok) return;
            const result = await res.json();
            const raw = result.values && result.values[0] && result.values[0][0];
            if (!raw) return;
            const data = JSON.parse(raw);
            if (!data || !data.equipmentData) return;
            if (raw !== lastSyncedSnapshot) {
                applySharedData(data);
                lastSyncedSnapshot = raw;
                const currentState = navigationStack[navigationStack.length - 1];
                if (currentState) renderState(currentState);
                setSyncStatus('Updated by a teammate');
            }
        } catch (err) {
            // Stay quiet on background poll failures.
        }
    }, POLL_INTERVAL_MS);
}

// Called once on page load. If MSAL already has a cached account (returning
// visitor), sync starts immediately with no popup needed. Otherwise the
// "Sign in to sync" button stays visible until clicked.
function initSync() {
    if (!SYNC_CONFIGURED) {
        setSyncStatus('Local only (sharing not set up)');
        return;
    }
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
        msalInstance.setActiveAccount(accounts[0]);
        signInBtn.classList.add('hidden');
        loadFromSheet().finally(startPolling);
    } else {
        signInBtn.classList.remove('hidden');
        setSyncStatus('Not signed in');
    }
}

function showContent(title, html) {
    contentArea.innerHTML = `<h2>${title}</h2>${html}`;
}

function pushState(state) {
    navigationStack.push(state);
    backBtn.disabled = navigationStack.length <= 1;
    updateButtonVisibility();
}

function goBack() {
    if (navigationStack.length > 1) {
        navigationStack.pop();
        const prevState = navigationStack[navigationStack.length - 1];
        renderState(prevState);
    } else {
        showWelcome();
    }
}

function updateButtonVisibility() {
    const currentState = navigationStack[navigationStack.length - 1];
    const isSearchOrArchive = currentState && (currentState.type === 'search' || currentState.type === 'archive');
    searchBtn.classList.toggle('hidden', isSearchOrArchive);
    addEquipmentBtn.classList.toggle('hidden', isSearchOrArchive);
    mapBtn.classList.toggle('hidden', currentState && currentState.type === 'map');
    archiveBtn.classList.toggle('hidden', currentState && currentState.type === 'archive');
}

function renderState(state) {
    mapDiv.style.display = 'none';
    contentArea.classList.remove('hidden');

    switch (state.type) {
        case 'map':
            showMap();
            break;
        case 'lab':
            displayLab(state.name, state.showBasic, state.showGeneralLabware);
            break;
        case 'main':
            if (state.area === 'Preproom') showPreproom();
            else if (state.area === 'Greenhouse') showGreenhouse();
            else if (state.area === 'Bookstore') showBookstore();
            else if (state.area === 'Academy') showAcademy();
            break;
        case 'subarea':
            showSubArea(state.parent, state.name);
            break;
        case 'cabinet':
            showCabinet(state.parent, state.mainArea, state.cabinetName);
            break;
        case 'shelf':
            showShelf(state.parent, state.subArea, state.shelfName);
            break;
        case 'search':
            showEquipmentSearch(state.resultsHtml);
            break;
        case 'archive':
            showArchiveView();
            break;
        default:
            showWelcome();
            break;
    }
    attachDynamicEventListeners();
}

function showWelcome() {
    showContent('Welcome to the Laboratory Inventory', `<p>Select an area on the map to explore.</p>`);
    mapDiv.style.display = 'grid';
    navigationStack = [{ type: 'map' }];
    backBtn.disabled = true;
    updateButtonVisibility();
}

function showMap() {
    showWelcome();
}

function showArchiveView() {
    pushState({ type: 'archive' });
    const html = archivedEquipment.length === 0
        ? '<p>The archive is currently empty.</p>'
        : generateEquipmentListHTML(archivedEquipment, 'archived', ['archived'], true);
    showContent('Archived Items', html);
    attachDynamicEventListeners();
}

function generateEquipmentListHTML(items, itemType, itemPath, isArchived = false) {
    let html = '<ul>';
    if (items && items.length > 0) {
        items.forEach((item, index) => {
            const uniqueId = `qty-${itemType}-${itemPath.join('-')}-${index}`;
            html += `
                <li class="equipment-item">
                    <span class="equipment-name">${item.name}</span>
                    <div class="quantity-controls">
                        <button data-action="decrease" data-type="${itemType}" data-path="${itemPath.join(',')}" data-index="${index}">-</button>
                        <input type="text" id="${uniqueId}" value="${item.qty}" data-type="${itemType}" data-path="${itemPath.join(',')}" data-index="${index}">
                        <button data-action="increase" data-type="${itemType}" data-path="${itemPath.join(',')}" data-index="${index}">+</button>
                    </div>
                    <div class="item-actions">
                        ${isArchived ? `
                            <button class="restore-item" data-action="restore" data-type="${itemType}" data-path="${itemPath.join(',')}" data-index="${index}">
                                <i class="fa-solid fa-undo"></i> Restore
                            </button>
                            <button class="delete-item" data-action="permanent-delete" data-type="${itemType}" data-path="${itemPath.join(',')}" data-index="${index}">
                                <i class="fa-solid fa-trash-alt"></i> Delete
                            </button>
                        ` : `
                            <button class="archive-item" data-action="archive" data-type="${itemType}" data-path="${itemPath.join(',')}" data-index="${index}">
                                <i class="fa-solid fa-box-archive"></i> Archive
                            </button>
                            <button class="delete-item" data-action="delete" data-type="${itemType}" data-path="${itemPath.join(',')}" data-index="${index}">
                                <i class="fa-solid fa-trash-alt"></i> Delete
                            </button>
                        `}
                    </div>
                </li>
            `;
        });
    } else {
        html += `<p>No equipment found here.</p>`;
    }
    html += '</ul>';
    return html;
}

function showLab(labName) {
    const currentState = navigationStack[navigationStack.length - 1];
    const newState = (currentState && currentState.type === 'lab' && currentState.name === labName) ? currentState : { type: 'lab', name: labName, showBasic: false, showGeneralLabware: false };
    if (newState !== currentState) {
        pushState(newState);
    }
    displayLab(newState.name, newState.showBasic, newState.showGeneralLabware);
}

function displayLab(labName, showBasic = false, showGeneralLabware = false) {
    const data = equipmentData[labName];
    if (!data) {
        showContent(labName, `<p>No data available for ${labName}.</p>`);
        return;
    }

    let html = '';
    if (data.basic && data.basic.length > 0) {
        html += `<h3>General Equipment</h3>${generateEquipmentListHTML(data.basic, 'basic', [labName, 'basic'])}`;
    }

    html += `
        <h3>Basic Equipment</h3>
        <button class="toggle-button" data-toggle="basic" data-lab="${labName}">
            ${showBasic ? 'Hide Basic Equipment' : 'Show Basic Equipment'}
        </button>
        <div id="basicEquipmentContainer" class="${showBasic ? '' : 'hidden'}">
            ${generateEquipmentListHTML(standardBasicEquipment, 'standardBasic', ['standardBasic'])}
        </div>
        <h3>General Labware</h3>
        <button class="toggle-button" data-toggle="generalLabware" data-lab="${labName}">
            ${showGeneralLabware ? 'Hide General Labware' : 'Show General Labware'}
        </button>
        <div id="generalLabwareContainer" class="${showGeneralLabware ? '' : 'hidden'}">
            ${generateEquipmentListHTML(standardGeneralLabware, 'standardLabware', ['standardLabware'])}
        </div>
    `;

    const cabinetNames = Object.keys(data.cabinets || {});
    if (cabinetNames.length > 0) {
        html += '<h3>Cabinets</h3><div class="cabinet-buttons">';
        cabinetNames.forEach(cab => {
            html += `<button class="cabinet-button-js" data-parent="Lab" data-mainarea="${labName}" data-cabinet="${cab}">${cab}</button>`;
        });
        html += '</div>';
    }

    showContent(labName, html);
    attachDynamicEventListeners();
}

function showCabinet(parentType, mainAreaName, cabinetName) {
    pushState({ type: 'cabinet', parent: parentType, mainArea: mainAreaName, cabinetName: cabinetName });

    let items, title;
    let path = [];
    if (parentType === 'Lab') {
        items = equipmentData[mainAreaName]?.cabinets?.[cabinetName];
        title = `${mainAreaName} - ${cabinetName}`;
        path = [mainAreaName, 'cabinets', cabinetName];
    } else if (parentType === 'Preproom' && ['Biology', 'Chemistry', 'Physics'].includes(mainAreaName)) {
        items = equipmentData['Preproom']?.subcategories?.[mainAreaName]?.cabinets?.[cabinetName];
        title = `Preproom - ${mainAreaName} - ${cabinetName}`;
        path = ['Preproom', 'subcategories', mainAreaName, 'cabinets', cabinetName];
    } else if (['Greenhouse', 'Bookstore', 'Academy'].includes(parentType)) {
        items = equipmentData[parentType]?.cabinets?.[cabinetName];
        title = `${parentType} - ${cabinetName}`;
        path = [parentType, 'cabinets', cabinetName];
    }

    const html = generateEquipmentListHTML(items, 'cabinet', path);
    showContent(title, html);
    attachDynamicEventListeners();
}

function showShelf(parent, subArea, shelfName) {
    pushState({ type: 'shelf', parent: parent, subArea: subArea, shelfName: shelfName });

    // "Shelves" is a category, not an item list - render buttons for the nested shelves it contains
    if (shelfName === 'Shelves') {
        const shelvesData = equipmentData[parent]?.subcategories?.[subArea]?.Shelves;
        let html = '<h3>Shelves:</h3><div class="cabinet-buttons">';
        if (shelvesData) {
            Object.keys(shelvesData).sort().forEach(nestedName => {
                html += `<button class="nested-shelf-button-js" data-parent="${parent}" data-subarea="${subArea}" data-nestedshelf="${nestedName}">${nestedName}</button>`;
            });
        }
        html += '</div>';
        showContent(`${parent} - ${subArea} - Shelves`, html);
        attachDynamicEventListeners();
        return;
    }

    let items, title, itemType, path;
    if (['C1', 'C2', 'C3', 'C4'].includes(shelfName)) {
        items = equipmentData[parent]?.subcategories?.[subArea]?.[shelfName];
        title = `${parent} - ${subArea} - ${shelfName}`;
        itemType = 'shelf';
        path = [parent, 'subcategories', subArea, shelfName];
    } else {
        // A nested shelf living inside the "Shelves" category (e.g. "Main Shelf 1")
        items = equipmentData[parent]?.subcategories?.[subArea]?.Shelves?.[shelfName];
        title = `${parent} - ${subArea} - Shelves - ${shelfName}`;
        itemType = 'nestedShelf';
        path = [parent, 'subcategories', subArea, 'Shelves', shelfName];
    }

    const html = generateEquipmentListHTML(items, itemType, path);
    showContent(title, html);
    attachDynamicEventListeners();
}

function showEquipmentSearch(resultsHtml = '') {
    pushState({ type: 'search', resultsHtml: resultsHtml });
    let html = `
        <div class="search-container">
            <input type="text" id="searchInput" placeholder="Search equipment...">
            <button id="searchBtnConfirm">Search</button>
        </div>
        <div id="searchResults">${resultsHtml}</div>
    `;
    showContent('Search Equipment', html);

    const searchInput = document.getElementById('searchInput');
    const searchBtnConfirm = document.getElementById('searchBtnConfirm');
    const searchResultsDiv = document.getElementById('searchResults');
    searchInput.focus();

    const performSearch = () => {
        const query = searchInput.value.toLowerCase().trim();
        if (query.length < 2) {
            searchResultsDiv.innerHTML = '<p>Please enter at least 2 characters to search.</p>';
            return;
        }

        let resultsHtml = '<h3>Search Results:</h3><ul>';
        let found = false;

        const searchList = (list, location, typeDetail = '') => {
            if (!list || !Array.isArray(list)) return;
            list.forEach(item => {
                if (item.name.toLowerCase().includes(query)) {
                    const locationDetail = `<strong>${location}</strong>${typeDetail ? `, ${typeDetail}` : ''}`;
                    resultsHtml += `<li>${item.name} (Qty: ${item.qty}) - Found in: ${locationDetail}</li>`;
                    found = true;
                }
            });
        };

        for (const locationName in equipmentData) {
            const data = equipmentData[locationName];
            if (data.basic) searchList(data.basic, locationName, 'General Items');
            if (locationName.startsWith('Lab')) {
                searchList(standardBasicEquipment, locationName, 'Standard Basic Equipment');
                searchList(standardGeneralLabware, locationName, 'Standard General Labware');
            }
            if (data.cabinets) {
                for (const cabinetName in data.cabinets) {
                    searchList(data.cabinets[cabinetName], locationName, `Cabinet: ${cabinetName}`);
                }
            }
            if (data.subcategories) {
                for (const subcategoryName in data.subcategories) {
                    const subcatData = data.subcategories[subcategoryName];
                    if (subcatData.cabinets) {
                        for (const cabinetName in subcatData.cabinets) {
                            searchList(subcatData.cabinets[cabinetName], locationName, `Section: ${subcategoryName}, Column/Cabinet: ${cabinetName}`);
                        }
                    }
                    if (subcategoryName === 'Chemical Store') {
                        for (const chemShelfName of ['C1', 'C2', 'C3', 'C4']) {
                            if (Array.isArray(subcatData[chemShelfName])) {
                                searchList(subcatData[chemShelfName], locationName, `Section: ${subcategoryName}, ${chemShelfName}`);
                            }
                        }
                        if (subcatData.Shelves) {
                            for (const nestedShelfName in subcatData.Shelves) {
                                searchList(subcatData.Shelves[nestedShelfName], locationName, `Section: ${subcategoryName}, Shelves: ${nestedShelfName}`);
                            }
                        }
                    }
                }
            }
        }

        if (!found) {
            resultsHtml += '<li>No equipment found matching your search.</li>';
        }
        resultsHtml += '</ul>';
        searchResultsDiv.innerHTML = resultsHtml;

        const currentState = navigationStack[navigationStack.length - 1];
        if (currentState.type === 'search') {
            currentState.resultsHtml = resultsHtml;
        }
    };

    searchBtnConfirm.addEventListener('click', performSearch);
    searchInput.addEventListener('keypress', (e) => {
        if (e.key === 'Enter') {
            performSearch();
        }
    });
}

function addEquipment(location, dynamicSelection, subDynamicSelection, name, quantity) {
    let targetList = null;
    let displayPath = [location];

    if (location.startsWith('Lab')) {
        const labData = equipmentData[location];
        if (labData?.cabinets?.[dynamicSelection]) {
            targetList = labData.cabinets[dynamicSelection];
            displayPath.push(dynamicSelection);
        } else if (dynamicSelection === 'Standard Basic Equipment') {
            targetList = standardBasicEquipment;
            displayPath.push('Standard Basic Equipment');
        } else if (dynamicSelection === 'Standard General Labware') {
            targetList = standardGeneralLabware;
            displayPath.push('Standard General Labware');
        } else if (labData.basic && dynamicSelection === 'General Items') {
            targetList = labData.basic;
            displayPath.push('General Items');
        }
    } else if (location === 'Preproom') {
        const subcatData = equipmentData[location]?.subcategories?.[dynamicSelection];
        if (subcatData?.cabinets?.[subDynamicSelection]) {
            targetList = subcatData.cabinets[subDynamicSelection];
            displayPath.push(dynamicSelection, subDynamicSelection);
        } else if (dynamicSelection === 'Chemical Store') {
            if (subcatData?.[subDynamicSelection]) {
                targetList = subcatData[subDynamicSelection];
                displayPath.push(dynamicSelection, subDynamicSelection);
            } else if (subcatData?.Shelves?.[subDynamicSelection]) {
                targetList = subcatData.Shelves[subDynamicSelection];
                displayPath.push(dynamicSelection, 'Shelves', subDynamicSelection);
            }
        }
    } else if (['Greenhouse', 'Bookstore', 'Academy'].includes(location)) {
        const facilityData = equipmentData[location];
        if (dynamicSelection === 'General Items' && facilityData.basic) {
            targetList = facilityData.basic;
            displayPath.push('General Items');
        } else if (facilityData?.cabinets?.[dynamicSelection]) {
            targetList = facilityData.cabinets[dynamicSelection];
            displayPath.push(dynamicSelection);
        }
    }

    if (!targetList) {
        showCustomModal('Error', 'Could not determine where to add equipment. Please check your selections.', 'error');
        return;
    }

    const existingItem = targetList.find(item => item.name.toLowerCase() === name.toLowerCase());
    if (existingItem) {
        if (typeof existingItem.qty === 'number' && typeof quantity === 'number') {
            existingItem.qty += quantity;
        } else if (typeof existingItem.qty === 'string' && typeof quantity === 'string') {
            const existingParts = existingItem.qty.match(/^(\d+)\s*(.*)$/);
            const newParts = quantity.match(/^(\d+)\s*(.*)$/);
            if (existingParts && newParts && existingParts[2] === newParts[2]) {
                existingItem.qty = `${parseInt(existingParts[1], 10) + parseInt(newParts[1], 10)} ${existingParts[2]}`.trim();
            } else {
                targetList.push({ name, qty: quantity });
                showCustomModal('Warning', `Units for "${name}" did not match. Added as a new separate entry. Current: ${existingItem.qty}, New: ${quantity}`);
            }
        } else {
            targetList.push({ name, qty: quantity });
            showCustomModal('Warning', `Quantity type mismatch for "${name}". Added as a new separate entry. Current: ${existingItem.qty}, New: ${quantity}`);
        }
    } else {
        targetList.push({ name, qty: quantity });
    }

    renderState(navigationStack[navigationStack.length - 1]);
    showCustomModal('Success', `Equipment "${name}" (Qty: ${quantity}) added successfully to ${displayPath.join(' -> ')}.`);
    scheduleSave();
}

function updateQuantity(item, valueOrDelta, isDirectValue = false) {
    let currentQty = item.qty;
    if (typeof currentQty === 'string') {
        const parts = currentQty.match(/^(\d+)\s*(.*)$/);
        let numPart = parts ? parseInt(parts[1], 10) : 0;
        const unitPart = parts ? parts[2] : '';
        if (isDirectValue) {
            let newNum = parseInt(valueOrDelta, 10);
            if (isNaN(newNum)) newNum = 0;
            item.qty = `${Math.max(0, newNum)} ${unitPart}`.trim();
        } else {
            numPart = Math.max(0, numPart + valueOrDelta);
            item.qty = `${numPart} ${unitPart}`.trim();
        }
    } else {
        if (isDirectValue) {
            let newNum = parseInt(valueOrDelta, 10);
            if (isNaN(newNum)) newNum = 0;
            item.qty = Math.max(0, newNum);
        } else {
            item.qty = Math.max(0, currentQty + valueOrDelta);
        }
    }
}

// Walks a comma/array path of literal object keys down into equipmentData
// (e.g. ['Lab 1','cabinets','Cupboard A'] or ['Preproom','subcategories','Chemistry','cabinets','Chemistry Column 01'])
// and returns the array found there, or null if the path doesn't resolve to an array.
function getItemListFromPath(path) {
    if (path[0] === 'standardBasic') {
        return standardBasicEquipment;
    }
    if (path[0] === 'standardLabware') {
        return standardGeneralLabware;
    }
    let currentLevel = equipmentData;
    for (let i = 0; i < path.length; i++) {
        const segment = path[i];
        if (currentLevel && currentLevel[segment] !== undefined) {
            currentLevel = currentLevel[segment];
        } else {
            return null;
        }
    }
    return Array.isArray(currentLevel) ? currentLevel : null;
}

// Single source of truth for "give me the list an item lives in", used by both
// read (getItemFromData) and write (findAndManipulateItem) operations so they can't drift apart.
function resolveList(itemType, path) {
    return itemType === 'archived' ? archivedEquipment : getItemListFromPath(path);
}

function getItemFromData(type, path, index) {
    const list = resolveList(type, path);
    return list ? list[index] : null;
}

function findAndManipulateItem(itemType, itemPathString, index, action) {
    const currentList = resolveList(itemType, itemPathString.split(','));

    if (!currentList) {
        console.error('List not found for path:', itemPathString);
        showCustomModal('Error', 'Could not locate that item. Please refresh and try again.', 'error');
        return;
    }

    const item = currentList[index];
    if (!item) {
        console.error('Item not found at index:', index, 'for list:', currentList);
        return;
    }

    if (action === 'delete') {
        currentList.splice(index, 1);
        showCustomModal('Deleted', `"${item.name}" has been deleted.`, 'success');
    } else if (action === 'archive') {
        currentList.splice(index, 1);
        archivedEquipment.push({ ...item, originalPath: itemPathString });
        showCustomModal('Archived', `"${item.name}" has been moved to the archive.`, 'info');
    } else if (action === 'permanent-delete') {
        currentList.splice(index, 1);
        showCustomModal('Deleted', `"${item.name}" has been permanently deleted.`, 'success');
    } else if (action === 'restore') {
        const { originalPath, ...restoredItem } = item;
        const originalList = getItemListFromPath(originalPath.split(','));
        if (originalList) {
            originalList.push(restoredItem);
            currentList.splice(index, 1);
            showCustomModal('Restored', `"${item.name}" has been restored to its original location.`, 'success');
        } else {
            console.error('Failed to find original location for restore:', originalPath);
            showCustomModal('Error', 'Failed to restore item. Original location not found.', 'error');
        }
    }

    const currentState = navigationStack[navigationStack.length - 1];
    if (currentState) {
        renderState(currentState);
    }
    scheduleSave();
}

function showCustomModal(title, message, type) {
    confirmModalTitle.textContent = title;
    confirmModalMessage.textContent = message;
    confirmModal.style.display = 'flex';
    confirmModalConfirmBtn.classList.add('hidden');
    confirmModalCancelBtn.textContent = 'OK';
    confirmModalCancelBtn.classList.remove('danger-btn');
    confirmModalCancelBtn.classList.add('success-btn');
    confirmModalConfirmBtn.onclick = null;
    confirmModalCancelBtn.onclick = () => confirmModal.style.display = 'none';
}

function showConfirmationModal(title, message, confirmCallback) {
    confirmModalTitle.textContent = title;
    confirmModalMessage.textContent = message;
    confirmModal.style.display = 'flex';
    confirmModalConfirmBtn.classList.remove('hidden');
    confirmModalCancelBtn.textContent = 'Cancel';
    confirmModalCancelBtn.classList.remove('success-btn');
    confirmModalCancelBtn.classList.add('danger-btn');

    pendingAction = confirmCallback;

    confirmModalConfirmBtn.onclick = () => {
        pendingAction();
        confirmModal.style.display = 'none';
    };
    confirmModalCancelBtn.onclick = () => {
        confirmModal.style.display = 'none';
    };
}

function attachDynamicEventListeners() {
    contentArea.removeEventListener('click', handleDynamicClick);
    contentArea.addEventListener('click', handleDynamicClick);
    contentArea.removeEventListener('change', handleQuantityChange);
    contentArea.addEventListener('change', handleQuantityChange);
}

function handleDynamicClick(event) {
    const target = event.target;
    const itemButton = target.closest('[data-action]');
    const cabinetButton = target.closest('.cabinet-button-js');
    const subAreaButton = target.closest('.subarea-button-js');
    const toggleButton = target.closest('.toggle-button');
    const nestedShelfButton = target.closest('.nested-shelf-button-js');
    const shelfButton = target.closest('.shelf-button-js');
    const shelfCategoryButton = target.closest('.shelf-category-js');
    const mapCell = target.closest('.map-cell');

    if (mapCell) {
        handleAreaSelection(mapCell.dataset.area);
    } else if (itemButton) {
        const action = itemButton.dataset.action;
        const type = itemButton.dataset.type;
        const path = itemButton.dataset.path;
        const index = parseInt(itemButton.dataset.index, 10);

        if (action === 'increase' || action === 'decrease') {
            const delta = action === 'increase' ? 1 : -1;
            const item = getItemFromData(type, path.split(','), index);
            if (item) {
                updateQuantity(item, delta);
                document.getElementById(`qty-${type}-${path.split(',').join('-')}-${index}`).value = item.qty;
                scheduleSave();
            }
        } else if (action === 'archive' || action === 'delete' || action === 'permanent-delete' || action === 'restore') {
            const item = getItemFromData(type, path.split(','), index);
            const itemName = item?.name || 'this item';
            let title, message;

            if (action === 'archive') {
                title = 'Confirm Archive';
                message = `Are you sure you want to archive "${itemName}"? You can restore it later from the Archive.`;
            } else if (action === 'delete') {
                title = 'Confirm Delete';
                message = `Are you sure you want to delete "${itemName}"? This action cannot be undone.`;
            } else if (action === 'permanent-delete') {
                title = 'Confirm Permanent Delete';
                message = `Are you sure you want to PERMANENTLY delete "${itemName}"? This action cannot be undone.`;
            } else if (action === 'restore') {
                title = 'Confirm Restore';
                message = `Are you sure you want to restore "${itemName}" to its original location?`;
            }

            showConfirmationModal(title, message, () => {
                findAndManipulateItem(type, path, index, action);
            });
        }
    } else if (cabinetButton) {
        showCabinet(cabinetButton.dataset.parent, cabinetButton.dataset.mainarea, cabinetButton.dataset.cabinet);
    } else if (subAreaButton) {
        showSubArea(subAreaButton.dataset.parent, subAreaButton.dataset.subarea);
    } else if (toggleButton) {
        const toggleType = toggleButton.dataset.toggle;
        const labName = toggleButton.dataset.lab;
        const currentState = navigationStack[navigationStack.length - 1];
        if (currentState && currentState.type === 'lab' && currentState.name === labName) {
            if (toggleType === 'basic') {
                currentState.showBasic = !currentState.showBasic;
            } else if (toggleType === 'generalLabware') {
                currentState.showGeneralLabware = !currentState.showGeneralLabware;
            }
            displayLab(labName, currentState.showBasic, currentState.showGeneralLabware);
        }
    } else if (nestedShelfButton) {
        showShelf(nestedShelfButton.dataset.parent, nestedShelfButton.dataset.subarea, nestedShelfButton.dataset.nestedshelf);
    } else if (shelfButton) {
        showShelf(shelfButton.dataset.parent, shelfButton.dataset.subarea, shelfButton.dataset.shelf);
    } else if (shelfCategoryButton) {
        showShelf(shelfCategoryButton.dataset.parent, shelfCategoryButton.dataset.subarea, 'Shelves');
    }
}

function handleQuantityChange(event) {
    const input = event.target.closest('input[type="text"]');
    if (input) {
        const type = input.dataset.type;
        const path = input.dataset.path.split(',');
        const index = parseInt(input.dataset.index, 10);
        const item = getItemFromData(type, path, index);
        if (item) {
            updateQuantity(item, input.value, true);
            scheduleSave();
        }
    }
}

function handleAreaSelection(area) {
    mapDiv.style.display = 'none';
    contentArea.classList.remove('hidden');

    if (area.startsWith("Lab")) {
        showLab(area);
    } else {
        switch (area) {
            case 'Preproom': showPreproom(); break;
            case 'Greenhouse': showGreenhouse(); break;
            case 'Bookstore': showBookstore(); break;
            case 'Academy': showAcademy(); break;
            default: showContent(area, `<p>Details about ${area} coming soon.</p>`); pushState({ type: 'facility', name: area }); break;
        }
    }
}

function showPreproom() {
    pushState({ type: 'main', area: 'Preproom' });
    const preproomData = equipmentData['Preproom'];
    let html = '<h3>Sections:</h3><div class="cabinet-buttons">';
    if (preproomData && preproomData.subcategories) {
        const sortedSubcategories = Object.keys(preproomData.subcategories).sort();
        sortedSubcategories.forEach(subArea => {
            html += `<button class="subarea-button-js" data-parent="Preproom" data-subarea="${subArea}">${subArea}</button>`;
        });
    }
    html += '</div>';
    showContent('Preproom', html);
    attachDynamicEventListeners();
}

function showSubArea(parent, subArea) {
    pushState({ type: 'subarea', parent: parent, name: subArea });
    const subAreaContent = equipmentData[parent]?.subcategories?.[subArea];
    let html = '';
    let title = subArea;

    if (parent === 'Preproom') {
        if ((subArea === 'Biology' || subArea === 'Chemistry' || subArea === 'Physics') && subAreaContent?.cabinets) {
            title = `Preproom - ${subArea}`;
            html += `<h3>${subArea} Columns:</h3><div class="cabinet-buttons">`;
            for (const cabinetName in subAreaContent.cabinets) {
                html += `<button class="cabinet-button-js" data-parent="Preproom" data-mainarea="${subArea}" data-cabinet="${cabinetName}">${cabinetName}</button>`;
            }
            html += '</div>';
        } else if (subArea === 'Chemical Store') {
            title = `Preproom - Chemical Store`;
            html += '<h3>Chemical Store Divisions:</h3><div class="cabinet-buttons">';
            for (const chemDivision of ['C1', 'C2', 'C3', 'C4']) {
                if (Array.isArray(subAreaContent[chemDivision])) {
                    html += `<button class="shelf-button-js" data-parent="Preproom" data-subarea="Chemical Store" data-shelf="${chemDivision}">${chemDivision}</button>`;
                }
            }
            if (subAreaContent.Shelves) {
                html += `<button class="shelf-category-js" data-parent="Preproom" data-subarea="Chemical Store">Shelves</button>`;
            }
            html += '</div>';
        }
    }

    showContent(title, html);
    attachDynamicEventListeners();
}

function showGreenhouse() {
    pushState({ type: 'main', area: 'Greenhouse' });
    const greenhouseData = equipmentData['Greenhouse'];
    let html = '<h3>Long Room with Cabinets</h3><div class="cabinet-buttons">';
    if (greenhouseData && greenhouseData.cabinets) {
        for (const cabinetName in greenhouseData.cabinets) {
            html += `<button class="cabinet-button-js" data-parent="Greenhouse" data-mainarea="Long Room" data-cabinet="${cabinetName}">${cabinetName}</button>`;
        }
    }
    html += '</div>';
    showContent('Greenhouse', html);
    attachDynamicEventListeners();
}

function showBookstore() {
    pushState({ type: 'main', area: 'Bookstore' });
    const bookstoreData = equipmentData['Bookstore'];
    let html = '';
    if (bookstoreData && bookstoreData.basic && bookstoreData.basic.length > 0) {
        html += '<h3>General Items:</h3>' + generateEquipmentListHTML(bookstoreData.basic, 'basic', ['Bookstore', 'basic']);
    }
    if (bookstoreData && bookstoreData.cabinets && Object.keys(bookstoreData.cabinets).length > 0) {
        html += '<h3>Shelves:</h3><div class="cabinet-buttons">';
        for (const cabinetName in bookstoreData.cabinets) {
            html += `<button class="cabinet-button-js" data-parent="Bookstore" data-mainarea="Main Area" data-cabinet="${cabinetName}">${cabinetName}</button>`;
        }
        html += '</div>';
    }
    showContent('Bookstore', html);
    attachDynamicEventListeners();
}

function showAcademy() {
    pushState({ type: 'main', area: 'Academy' });
    const academyData = equipmentData['Academy'];
    let html = '';
    if (academyData && academyData.basic && academyData.basic.length > 0) {
        html += '<h3>General Equipment:</h3>' + generateEquipmentListHTML(academyData.basic, 'basic', ['Academy', 'basic']);
    }
    if (academyData && academyData.cabinets && Object.keys(academyData.cabinets).length > 0) {
        html += '<h3>Cabinets:</h3><div class="cabinet-buttons">';
        for (const cabinetName in academyData.cabinets) {
            html += `<button class="cabinet-button-js" data-parent="Academy" data-mainarea="Main Room" data-cabinet="${cabinetName}">${cabinetName}</button>`;
        }
        html += '</div>';
    }
    showContent('Academy', html);
    attachDynamicEventListeners();
}

document.addEventListener('DOMContentLoaded', () => {
    backBtn.addEventListener('click', goBack);
    searchBtn.addEventListener('click', () => renderState({ type: 'search' }));
    mapBtn.addEventListener('click', showMap);
    archiveBtn.addEventListener('click', showArchiveView);
    addEquipmentBtn.addEventListener('click', openAddEquipmentModal);
    closeAddEquipmentModal.addEventListener('click', () => addEquipmentModal.style.display = 'none');
    selectLocation.addEventListener('change', () => {
        const location = selectLocation.value;
        populateDynamicSelect(location);
        selectDynamic.value = '';
        dynamicContainer.classList.add('hidden');
        selectSubDynamic.innerHTML = '<option value="">Select an option</option>';
        subDynamicContainer.classList.add('hidden');
        if (location) dynamicContainer.classList.remove('hidden');
    });
    selectDynamic.addEventListener('change', () => {
        const location = selectLocation.value;
        const dynamicSelection = selectDynamic.value;
        populateSubDynamicSelect(location, dynamicSelection);
        subDynamicContainer.classList.add('hidden');
        if (location === 'Preproom' && ['Biology', 'Chemistry', 'Physics', 'Chemical Store'].includes(dynamicSelection)) {
            subDynamicContainer.classList.remove('hidden');
        }
    });
    addEquipmentForm.addEventListener('submit', (event) => {
        event.preventDefault();
        const location = selectLocation.value;
        const dynamicSelection = selectDynamic.value;
        const subDynamicSelection = selectSubDynamic.value;
        const name = equipmentNameInput.value.trim();
        let quantity = equipmentQuantityInput.value.trim();
        if (name && quantity) {
            if (!isNaN(quantity) && !isNaN(parseFloat(quantity))) {
                quantity = parseInt(quantity, 10);
            }
            addEquipment(location, dynamicSelection, subDynamicSelection, name, quantity);
            addEquipmentModal.style.display = 'none';
            addEquipmentForm.reset();
        } else {
            showCustomModal('Error', 'Please fill in all required fields.', 'error');
        }
    });
    signInBtn.addEventListener('click', async () => {
        const ok = await ensureSignedIn();
        if (ok) {
            signInBtn.classList.add('hidden');
            await loadFromSheet();
            startPolling();
        }
    });
    mapDiv.addEventListener('click', (event) => {
        const cell = event.target.closest('.map-cell');
        if (cell) {
            handleAreaSelection(cell.dataset.area);
        }
    });

    gateForm.addEventListener('submit', (event) => {
        event.preventDefault();
        const username = gateUsernameInput.value.trim();
        const password = gatePasswordInput.value;
        if (!username) {
            gateError.textContent = 'Please enter your name.';
            return;
        }
        if (password !== GATE_PASSWORD) {
            gateError.textContent = 'Incorrect password.';
            gatePasswordInput.value = '';
            gatePasswordInput.focus();
            return;
        }
        gateError.textContent = '';
        unlockApp(username);
    });

    switchUserBtn.addEventListener('click', () => {
        localStorage.removeItem(GATE_USER_KEY);
        localStorage.removeItem(GATE_AUTHED_KEY);
        location.reload();
    });

    themeMenuBtn.addEventListener('click', (event) => {
        event.stopPropagation();
        const isOpen = !themeMenu.classList.contains('hidden');
        themeMenu.classList.toggle('hidden');
        themeMenuBtn.setAttribute('aria-expanded', String(!isOpen));
    });

    document.querySelectorAll('.theme-option').forEach(btn => {
        btn.addEventListener('click', () => {
            applyTheme(btn.dataset.theme);
            themeMenu.classList.add('hidden');
            themeMenuBtn.setAttribute('aria-expanded', 'false');
        });
    });

    document.addEventListener('click', (event) => {
        if (!themeMenu.classList.contains('hidden') && !event.target.closest('.theme-menu-wrapper')) {
            themeMenu.classList.add('hidden');
            themeMenuBtn.setAttribute('aria-expanded', 'false');
        }
    });

    checkGate();
});

function openAddEquipmentModal() {
    addEquipmentModal.style.display = 'flex';
    populateLocationSelect();
    selectLocation.value = '';
    selectDynamic.innerHTML = '<option value="">Select an option</option>';
    dynamicContainer.classList.add('hidden');
    selectSubDynamic.innerHTML = '<option value="">Select an option</option>';
    subDynamicContainer.classList.add('hidden');
    equipmentNameInput.value = '';
    equipmentQuantityInput.value = '';
}

function populateLocationSelect() {
    selectLocation.innerHTML = '<option value="">Select a location</option>';
    const allLocations = Object.keys(equipmentData).sort();
    allLocations.forEach(location => {
        const option = document.createElement('option');
        option.value = location;
        option.textContent = location;
        selectLocation.appendChild(option);
    });
}

function populateDynamicSelect(location) {
    selectDynamic.innerHTML = '<option value="">Select an option</option>';
    dynamicLabel.textContent = 'Cabinet/Section:';
    const data = equipmentData[location];
    if (!data) return;

    if (location.startsWith('Lab')) {
        const labOptions = [];
        if (data.basic) {
            labOptions.push({ value: 'General Items', text: 'General Items (No specific cabinet)' });
        }
        if (data.cabinets) {
            Object.keys(data.cabinets).sort().forEach(cabinet => {
                labOptions.push({ value: cabinet, text: cabinet });
            });
        }

        labOptions.forEach(opt => {
            const option = document.createElement('option');
            option.value = opt.value;
            option.textContent = opt.text;
            selectDynamic.appendChild(option);
        });

    } else if (location === 'Preproom' && data.subcategories) {
        dynamicLabel.textContent = 'Section:';
        const sortedSubcategories = Object.keys(data.subcategories).sort();
        sortedSubcategories.forEach(subcatName => {
            const option = document.createElement('option');
            option.value = subcatName;
            option.textContent = subcatName;
            selectDynamic.appendChild(option);
        });
    } else if (['Greenhouse', 'Bookstore', 'Academy'].includes(location)) {
        dynamicLabel.textContent = 'Cabinet/Shelf:';
        if (data.basic && Array.isArray(data.basic)) {
            const option = document.createElement('option');
            option.value = 'General Items';
            option.textContent = 'General Items (No specific cabinet)';
            selectDynamic.appendChild(option);
        }
        if (data.cabinets) {
            Object.keys(data.cabinets).sort().forEach(cabinet => {
                const option = document.createElement('option');
                option.value = cabinet;
                option.textContent = cabinet;
                selectDynamic.appendChild(option);
            });
        }
    }
}

function populateSubDynamicSelect(location, dynamicSelection) {
    selectSubDynamic.innerHTML = '<option value="">Select an option</option>';
    subDynamicLabel.textContent = 'Column/Shelf:';
    if (location === 'Preproom') {
        const subcatData = equipmentData[location]?.subcategories?.[dynamicSelection];
        if (subcatData?.cabinets) {
            subDynamicLabel.textContent = 'Column:';
            Object.keys(subcatData.cabinets).sort().forEach(cabinetName => {
                const option = document.createElement('option');
                option.value = cabinetName;
                option.textContent = cabinetName;
                selectSubDynamic.appendChild(option);
            });
        } else if (dynamicSelection === 'Chemical Store' && subcatData) {
            subDynamicLabel.textContent = 'Shelf:';
            const chemicalStoreOptions = ['C1', 'C2', 'C3', 'C4'];
            chemicalStoreOptions.forEach(chemDivision => {
                const option = document.createElement('option');
                option.value = chemDivision;
                option.textContent = chemDivision;
                selectSubDynamic.appendChild(option);
            });
            if (subcatData.Shelves) {
                Object.keys(subcatData.Shelves).sort().forEach(shelfName => {
                    const option = document.createElement('option');
                    option.value = shelfName;
                    option.textContent = shelfName;
                    selectSubDynamic.appendChild(option);
                });
            }
        }
    }
}
