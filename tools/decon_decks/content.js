// Company-specific content for the decontamination sales decks.
// Unit lists are indicative (typical for the facility type) and are confirmed during the joint site survey.

const CHEM = {
  VD: 'Vapour-phase degreaser',
  WB: 'Water-based solvent',
  OB: 'Oil-based solvent',
  OX: 'Sulphide / pyrophoric oxidiser',
  NE: 'Neutraliser',
  AF: 'Antifoam / dispersant',
};

const COMPANIES = [
  {
    key: 'OQ8_Duqm_Refinery', name: 'OQ8', full: 'OQ8 – Duqm Refinery', location: 'Duqm Special Economic Zone, Al Wusta',
    facility: 'Grassroots refinery processing heavy and sour crudes, with hydrocracking, delayed coking and sulphur recovery.',
    contaminants: ['H₂S', 'Benzene', 'LEL (hydrocarbons)', 'Pyrophoric FeS', 'Heavy oil & sludge', 'Coke', 'Ammonia', 'Mercaptans'],
    focus: 'Turnaround-ready decontamination of the crude, vacuum, conversion and sour-service units – fast, gas-free and pyrophoric-safe entry.',
    units: [
      { name: 'Crude Distillation Unit (CDU)', draw: 'column', ch: ['Heavy fouling in bottoms and pump-arounds', 'Steam distribution through large columns', 'Light ends and benzene in overheads'], chem: ['VD', 'OB'], opt: ['Vapour-phase through the steam-out lines', 'Bottoms circuit pre-flush'] },
      { name: 'Vacuum Distillation Unit (VDU)', draw: 'column', ch: ['Pyrophoric FeS in packed beds', 'Heavy vacuum residue in bottoms', 'Pump-around exchanger fouling'], chem: ['OB', 'VD', 'OX'], opt: ['Viscosity flush (circulation)', 'Packing pre-treatment', 'Vapour-phase with oxidiser post-rinse'] },
      { name: 'Hydrocracker (HCU)', draw: 'reactor', ch: ['H₂S and NH₃ in reactor effluent circuit', 'Pyrophoric scale in exchangers and separators', 'Light hydrocarbons trapped in catalyst beds'], chem: ['VD', 'OX'], opt: ['Vapour-phase of fractionation section', 'Circulation of HP/LP separators and air coolers'] },
      { name: 'Delayed Coker (DCU)', draw: 'coker', ch: ['Heaviest bottoms and coke fines', 'Time-consuming mechanical cleaning', 'Feed pre-heat exchanger fouling'], chem: ['OB', 'VD'], opt: ['Viscosity flush of bottoms circuit', 'Vapour-phase / boil-out combination on the main fractionator'] },
      { name: 'Sour Water Strippers & Amine Regeneration', draw: 'column', ch: ['High H₂S and ammonia', 'FeS deposits on trays and reboilers', 'Foaming during decontamination'], chem: ['OX', 'NE', 'AF'], opt: ['Vapour-phase with oxidiser', 'Neutralising rinse of the stripper bottoms'] },
      { name: 'Sulphur Recovery & Tail Gas (SRU / TGT)', draw: 'separator', ch: ['Residual H₂S and SO₂', 'Pyrophoric iron sulphide', 'Sulphur and amine residues'], chem: ['OX', 'WB'], opt: ['Circulation of the TGT absorber / quench', 'Boil-out of sulphur pits and drums'] },
      { name: 'Saturated Gas Plant & LPG Treating', draw: 'column', ch: ['LEL from LPG and light ends', 'Mercaptans and H₂S', 'Caustic and amine carry-over'], chem: ['VD', 'OX'], opt: ['Vapour-phase of absorber / stripper / debutaniser', 'Drum boil-out'] },
    ],
  },
  {
    key: 'OQ_Sohar_Refinery', name: 'OQ Sohar Refinery', full: 'OQ – Sohar Refinery', location: 'Sohar Port, Al Batinah North',
    facility: 'Residue-upgrading refinery with a residue fluid catalytic cracker (RFCC), hydrotreating and sulphur recovery.',
    contaminants: ['H₂S', 'Benzene', 'LEL (hydrocarbons)', 'Pyrophoric FeS', 'Slurry & catalyst fines', 'Heavy oil & sludge', 'Ammonia'],
    focus: 'Decontamination of the crude, RFCC and treating units and tank farm so mechanical work can start safely on day one of the shutdown.',
    units: [
      { name: 'Crude & Vacuum Distillation', draw: 'column', ch: ['Heavy fouling in bottoms', 'Pyrophoric deposits in packed sections', 'Benzene and LEL in overhead system'], chem: ['VD', 'OB'], opt: ['Vapour-phase through steam-out lines', 'Bottoms pre-flush'] },
      { name: 'RFCC Main Fractionator & Gas Concentration', draw: 'column', ch: ['Slurry, catalyst fines and coke in bottoms', 'Pyrophoric packing deposits', 'Light ends in gas concentration unit'], chem: ['OB', 'VD', 'OX'], opt: ['Viscosity flush of slurry circuit', 'Packing pre-treatment', 'Vapour-phase (with boil-out where needed)'] },
      { name: 'Hydrotreaters', draw: 'reactor', ch: ['H₂S in reactor effluent and separators', 'FeS in exchangers', 'Hydrocarbons in catalyst beds'], chem: ['VD', 'OX'], opt: ['Vapour-phase of stripper section', 'Circulation of separators and air coolers'] },
      { name: 'Desalter', draw: 'separator', ch: ['Heavy oil / water emulsions and sludge', 'Contact and distribution in a large vessel'], chem: ['WB', 'VD'], opt: ['Boil-out'] },
      { name: 'Sour Water & Amine Systems', draw: 'column', ch: ['High H₂S and ammonia', 'FeS on trays and reboilers', 'Foaming'], chem: ['OX', 'NE', 'AF'], opt: ['Vapour-phase with oxidiser', 'Neutralising rinse'] },
      { name: 'Sulphur Recovery Unit', draw: 'separator', ch: ['Residual H₂S / SO₂', 'Pyrophoric iron sulphide', 'Sulphur residues'], chem: ['OX', 'WB'], opt: ['Circulation of quench / absorber', 'Boil-out of drums and pits'] },
      { name: 'Crude & Product Tanks', draw: 'tank', ch: ['Sludge and heavy oil on the floor', 'Benzene and LEL in the vapour space', 'Man-entry restrictions'], chem: ['WB', 'VD'], opt: ['Gamma-jet circulation with decon chemistry', 'Heating where required'] },
    ],
  },
  {
    key: 'OQ_Mina_Al_Fahal_Refinery', name: 'OQ Mina Al Fahal Refinery', full: 'OQ – Mina Al Fahal Refinery', location: 'Mina Al Fahal, Muscat',
    facility: 'Coastal refinery and crude export terminal with crude distillation, hydrotreating and reforming.',
    contaminants: ['H₂S', 'Benzene', 'LEL (hydrocarbons)', 'Pyrophoric FeS', 'Heavy oil & sludge', 'Ammonia'],
    focus: 'Short-window decontamination of process units and tankage close to a city – low emissions and minimal waste.',
    units: [
      { name: 'Crude Distillation Unit', draw: 'column', ch: ['Heavy fouling in bottoms', 'Steam distribution', 'Benzene / LEL in overheads'], chem: ['VD', 'OB'], opt: ['Vapour-phase', 'Bottoms pre-flush'] },
      { name: 'Desalter', draw: 'separator', ch: ['Sludge and emulsions', 'Contact and distribution'], chem: ['WB', 'VD'], opt: ['Boil-out'] },
      { name: 'Naphtha Hydrotreater & Reformer', draw: 'reactor', ch: ['Benzene-rich reformate', 'H₂S and FeS in the hydrotreater circuit', 'Chlorides in the reformer section'], chem: ['VD', 'OX', 'NE'], opt: ['Vapour-phase of strippers / stabiliser', 'Circulation of separators'] },
      { name: 'Kerosene & Gas Oil Treating', draw: 'separator', ch: ['Mercaptans and caustic residues', 'Hydrocarbon LEL'], chem: ['VD', 'NE'], opt: ['Vessel boil-out', 'Circulation'] },
      { name: 'Sour Water Stripper', draw: 'column', ch: ['H₂S and ammonia', 'FeS deposits'], chem: ['OX', 'NE'], opt: ['Vapour-phase with oxidiser'] },
      { name: 'Crude & Product Tank Farm', draw: 'tank', ch: ['Bottom sludge', 'Benzene / LEL in vapour space', 'Limited man-entry'], chem: ['WB', 'VD'], opt: ['Gamma-jet circulation', 'Heating where required'] },
    ],
  },
  {
    key: 'OQ_Aromatics_Sohar', name: 'OQ Aromatics', full: 'OQ – Aromatics Complex, Sohar', location: 'Sohar Industrial Port, Al Batinah North',
    facility: 'Aromatics complex producing paraxylene and benzene from reformate.',
    contaminants: ['Benzene', 'Toluene & xylenes', 'LEL (hydrocarbons)', 'Heavy aromatics', 'Pyrophoric FeS'],
    focus: 'Benzene-critical decontamination: bringing aromatics equipment below benzene exposure limits quickly and verifiably.',
    units: [
      { name: 'Reforming & Stabiliser Section', draw: 'reactor', ch: ['Benzene-rich reformate', 'Chlorides and FeS in the stabiliser', 'Light hydrocarbons in reactors'], chem: ['VD', 'OX'], opt: ['Vapour-phase of stabiliser and separators'] },
      { name: 'Aromatics Extraction (Benzene / Toluene)', draw: 'column', ch: ['High benzene in columns and solvent circuit', 'Solvent residues', 'Strict benzene entry limits'], chem: ['VD', 'WB'], opt: ['Vapour-phase', 'Water-based rinse of solvent circuit'] },
      { name: 'Xylene Fractionation & Paraxylene', draw: 'column', ch: ['Heavy aromatics in column bottoms', 'Xylenes and LEL', 'Large columns and reboilers'], chem: ['VD', 'OB'], opt: ['Vapour-phase', 'Bottoms viscosity flush'] },
      { name: 'Exchangers & Air Coolers', draw: 'exchanger', ch: ['Aromatics trapped in bundles', 'Benzene release when opened'], chem: ['VD', 'WB'], opt: ['Circulation before bundle pulling'] },
      { name: 'Aromatics Storage Tanks', draw: 'tank', ch: ['Benzene in vapour space', 'Heavy aromatic residues'], chem: ['WB', 'VD'], opt: ['Gamma-jet circulation'] },
    ],
  },
  {
    key: 'Liwa_Plastics_Sohar', name: 'Liwa Plastics', full: 'OQ – Liwa Plastics Industries Complex, Sohar', location: 'Sohar Industrial Port, Al Batinah North',
    facility: 'Mixed-feed steam cracker with downstream polyethylene and polypropylene plants.',
    contaminants: ['Benzene', 'Tar & coke', 'Polymer fouling', 'LEL (hydrocarbons)', 'Red oil & caustic', 'H₂S'],
    focus: 'Decontaminating cracker quench and compression systems and polymer plants – heavy tar, polymer and benzene.',
    units: [
      { name: 'Quench Oil Tower', draw: 'quench', ch: ['Heavy coke and tar deposits', 'High benzene levels', 'Utility limitations'], chem: ['OB', 'VD'], opt: ['Circulation – viscosity flush', 'Boil-out', 'Vapour-phase'] },
      { name: 'Quench Water Tower', draw: 'column', ch: ['Polymer and emulsion fouling', 'Pyrolysis gasoline and benzene'], chem: ['VD', 'WB', 'AF'], opt: ['Vapour-phase', 'Circulation'] },
      { name: 'Caustic Wash Tower', draw: 'column', ch: ['Red oil / polymer', 'Spent caustic and H₂S'], chem: ['WB', 'NE', 'OX'], opt: ['Circulation with neutralising rinse'] },
      { name: 'Cracked Gas Compression Knock-out Drums', draw: 'separator', ch: ['Polymer and hydrocarbon condensate', 'LEL and benzene'], chem: ['VD', 'OB'], opt: ['Vapour-phase', 'Drum boil-out'] },
      { name: 'Polyethylene & Polypropylene Plants', draw: 'reactor', ch: ['Polymer and wax deposits', 'Monomer / LEL in purge bins and recovery'], chem: ['OB', 'VD'], opt: ['Circulation', 'Vapour-phase of recovery columns'] },
      { name: 'Pyrolysis Gasoline & Storage', draw: 'tank', ch: ['Benzene-rich vapour space', 'Gums and heavy residues'], chem: ['WB', 'VD'], opt: ['Gamma-jet circulation'] },
    ],
  },
  {
    key: 'Oman_Salalah_Methanol', name: 'Oman & Salalah Methanol', full: 'Oman Methanol (Sohar) & Salalah Methanol (Salalah)', location: 'Sohar Industrial Port and Salalah Free Zone',
    facility: 'Natural-gas-based methanol plants: desulphurisation, steam reforming, synthesis and distillation.',
    contaminants: ['LEL (methanol & hydrocarbons)', 'H₂S (feed gas)', 'Pyrophoric FeS', 'Wax / heavy alcohols', 'Iron deposits'],
    focus: 'Gas-freeing and cleaning feed-gas, synthesis and distillation equipment for inspection with minimal downtime.',
    units: [
      { name: 'Feed Gas Desulphurisation', draw: 'reactor', ch: ['H₂S and sulphur in guard beds', 'Pyrophoric iron sulphide', 'Hydrocarbon LEL'], chem: ['OX', 'VD'], opt: ['Vapour-phase', 'Oxidiser circulation before catalyst change'] },
      { name: 'Reformer Heat Recovery & Process Gas Coolers', draw: 'exchanger', ch: ['Deposits on exchanger tubes', 'Trapped process gas'], chem: ['WB', 'VD'], opt: ['Circulation'] },
      { name: 'Synthesis Loop Separators', draw: 'separator', ch: ['Methanol and wax residues', 'Iron and nickel deposits', 'LEL'], chem: ['VD', 'WB'], opt: ['Vapour-phase', 'Boil-out'] },
      { name: 'Methanol Distillation Columns', draw: 'column', ch: ['Methanol LEL', 'Higher alcohols and wax on trays', 'Large column volumes'], chem: ['VD', 'WB'], opt: ['Vapour-phase', 'Water-based rinse'] },
      { name: 'Methanol Storage Tanks', draw: 'tank', ch: ['Methanol vapour (LEL)', 'Sediments'], chem: ['WB'], opt: ['Gamma-jet circulation'] },
    ],
  },
  {
    key: 'PDO_Gas_Plants', name: 'PDO Gas Plants', full: 'Petroleum Development Oman – Gas Plants', location: 'Saih Rawl, Saih Nihayda, Yibal Khuff, Kauther and other sites',
    facility: 'Gas processing plants handling sweet and sour gas: separation, acid gas removal, dehydration, condensate stabilisation and sulphur recovery.',
    contaminants: ['H₂S (sour gas)', 'Mercaptans', 'Pyrophoric FeS', 'LEL (condensate)', 'BTEX / benzene', 'Glycol & amine residues'],
    focus: 'Safe, fast entry to sour-gas and condensate equipment – H₂S and pyrophoric control without long steaming or water washing.',
    units: [
      { name: 'Inlet Separation & Slug Catchers', draw: 'separator', ch: ['Condensate and sludge', 'H₂S and FeS', 'Large horizontal volumes'], chem: ['VD', 'OX', 'WB'], opt: ['Boil-out', 'Vapour-phase of separators'] },
      { name: 'Acid Gas Removal (Amine)', draw: 'column', ch: ['High H₂S', 'FeS on trays and reboilers', 'Amine foaming'], chem: ['OX', 'NE', 'AF'], opt: ['Vapour-phase with oxidiser', 'Neutralising rinse'] },
      { name: 'Glycol (TEG) Dehydration', draw: 'column', ch: ['Glycol and BTEX residues', 'Hydrocarbon carry-over'], chem: ['VD', 'WB'], opt: ['Vapour-phase', 'Water-based rinse'] },
      { name: 'Condensate Stabilisation', draw: 'column', ch: ['LEL from condensate', 'Benzene / BTEX', 'FeS in reboilers'], chem: ['VD', 'OX'], opt: ['Vapour-phase'] },
      { name: 'Sulphur Recovery (Sour Gas Plants)', draw: 'separator', ch: ['H₂S / SO₂', 'Pyrophoric iron sulphide', 'Sulphur residues'], chem: ['OX', 'WB'], opt: ['Circulation', 'Boil-out'] },
      { name: 'Produced Water & Condensate Tanks', draw: 'tank', ch: ['Sludge and oil layers', 'H₂S and LEL in vapour space'], chem: ['WB', 'OX'], opt: ['Gamma-jet circulation'] },
    ],
  },
  {
    key: 'BP_Khazzan', name: 'BP Oman – Khazzan', full: 'BP Oman – Khazzan & Ghazeer (Block 61)', location: 'Block 61, central Oman',
    facility: 'Tight-gas central processing facility: inlet separation, gas dehydration, condensate stabilisation and compression.',
    contaminants: ['LEL (condensate)', 'BTEX / benzene', 'Pyrophoric FeS', 'Glycol residues', 'H₂S (where present)'],
    focus: 'Gas-freeing condensate and dehydration equipment for planned shutdowns with verified entry conditions.',
    units: [
      { name: 'Inlet Separation', draw: 'separator', ch: ['Condensate and sludge', 'LEL', 'Large horizontal vessels'], chem: ['VD', 'WB'], opt: ['Boil-out', 'Vapour-phase'] },
      { name: 'Gas Dehydration', draw: 'column', ch: ['Glycol and BTEX residues', 'Hydrocarbon carry-over'], chem: ['VD', 'WB'], opt: ['Vapour-phase', 'Water-based rinse'] },
      { name: 'Condensate Stabilisation', draw: 'column', ch: ['LEL and benzene', 'FeS in reboilers'], chem: ['VD', 'OX'], opt: ['Vapour-phase'] },
      { name: 'Compression Scrubbers & KO Drums', draw: 'separator', ch: ['Condensate and lube-oil residues', 'LEL'], chem: ['VD', 'OB'], opt: ['Vapour-phase', 'Drum boil-out'] },
      { name: 'Condensate & Produced Water Tanks', draw: 'tank', ch: ['Oil / sludge layers', 'LEL in vapour space'], chem: ['WB'], opt: ['Gamma-jet circulation'] },
    ],
  },
  {
    key: 'Shell_Blocks_10_11', name: 'Shell Development Oman', full: 'Shell Development Oman – Blocks 10 & 11', location: 'Blocks 10 and 11, central Oman',
    facility: 'Gas and condensate production and processing facilities.',
    contaminants: ['LEL (condensate)', 'BTEX / benzene', 'Pyrophoric FeS', 'H₂S (where present)', 'Sludge'],
    focus: 'Decontamination of separation, dehydration and condensate equipment for inspection and tie-in work.',
    units: [
      { name: 'Inlet Separation & Pig Receivers', draw: 'separator', ch: ['Condensate, sludge and pigging debris', 'LEL and FeS'], chem: ['VD', 'WB', 'OX'], opt: ['Boil-out', 'Vapour-phase'] },
      { name: 'Gas Dehydration', draw: 'column', ch: ['Glycol and BTEX', 'Hydrocarbon carry-over'], chem: ['VD', 'WB'], opt: ['Vapour-phase'] },
      { name: 'Condensate Stabilisation', draw: 'column', ch: ['LEL and benzene', 'FeS in reboilers'], chem: ['VD', 'OX'], opt: ['Vapour-phase'] },
      { name: 'Exchangers & Coolers', draw: 'exchanger', ch: ['Hydrocarbons trapped in bundles', 'Deposits'], chem: ['VD', 'WB'], opt: ['Circulation'] },
      { name: 'Condensate & Water Tanks', draw: 'tank', ch: ['Sludge layers', 'LEL'], chem: ['WB'], opt: ['Gamma-jet circulation'] },
    ],
  },
  {
    key: 'OQ_Base_Industries_Salalah', name: 'OQ Base Industries', full: 'OQ Base Industries – Salalah', location: 'Salalah, Dhofar',
    facility: 'Gas processing in Salalah: LPG extraction and fractionation, with treating, dehydration and LPG storage.',
    contaminants: ['LEL (LPG & condensate)', 'Mercaptans', 'H₂S (traces)', 'Pyrophoric FeS', 'Amine & glycol residues'],
    focus: 'Gas-freeing LPG and fractionation equipment quickly, with odour (mercaptan) control near the community.',
    units: [
      { name: 'Inlet Gas Treating (Amine)', draw: 'column', ch: ['H₂S and CO₂', 'FeS and amine residues', 'Foaming'], chem: ['OX', 'NE', 'AF'], opt: ['Vapour-phase with oxidiser'] },
      { name: 'Dehydration & Guard Beds', draw: 'reactor', ch: ['Adsorbed hydrocarbons', 'Pyrophoric material at bed change-out'], chem: ['VD', 'OX'], opt: ['Vapour-phase before unloading'] },
      { name: 'Cryogenic Extraction / Cold Box', draw: 'exchanger', ch: ['Trapped light hydrocarbons', 'LEL'], chem: ['VD'], opt: ['Vapour-phase (warm-up first)'] },
      { name: 'Fractionation (C2 / C3 / C4 Columns)', draw: 'column', ch: ['LPG and LEL', 'Mercaptans', 'FeS in reboilers'], chem: ['VD', 'OX'], opt: ['Vapour-phase'] },
      { name: 'LPG Storage Spheres', draw: 'sphere', ch: ['LPG vapour (LEL)', 'Mercaptan odour', 'Water and sediments'], chem: ['VD', 'WB'], opt: ['Vapour-phase / steam with degreaser', 'Water-based rinse'] },
    ],
  },
];

module.exports = { CHEM, COMPANIES };
