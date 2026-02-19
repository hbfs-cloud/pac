// ============================================================
// SCHEMA_CATALOG - Mapping images extraites de l'Excel original
// ============================================================
const SCHEMA_CATALOG = {

  // ---- Photos produit par gamme_chassis ----
  photos: {
    effipac_S: 'image67.png',   // Effipac petit (14-18 kW)
    effipac_M: 'image65.png',   // Effipac moyen (26-32 kW)
    effipac_L: 'image62.png',   // Effipac grand (50 kW)
    effipac_XL: 'image66.png',  // Effipac XL (70 kW)
    aptae_S: 'image63.png',     // Aptae petit (15-18 kW)
    aptae_M: 'image64.png',     // Aptae moyen (23-27 kW)
    aptae_L: 'image68.png',     // Aptae grand (40-50 kW)
    aptae_XL: 'image66.png',    // Aptae XL (65 kW)
  },

  // ---- Plans dimensionnels (encombrement) ----
  dimensions: {
    effipac_S: 'image3.png',    // 1041mm face, raccords G 1"M
    effipac_M: 'image5.png',    // 1600mm face, 1 ventilateur
    effipac_L: 'image7.png',    // 1850mm, tubes rainures 1"1/2
    effipac_XL: 'image15.png',  // 2250mm, 2" taraude
    aptae_S: 'image9.png',      // 1100mm face, G 1"M
    aptae_M: 'image11.png',     // 1602mm face, G 1"1/4M
    aptae_L: 'image13.png',     // 1895mm, 2 ventilateurs top
    aptae_XL: 'image15.png',    // 2250mm, 2" taraude
  },

  // ---- Degagements techniques minimums ----
  degagements: {
    effipac_S: 'image4.png',
    effipac_M: 'image6.png',
    effipac_L: 'image8.png',
    effipac_XL: 'image16.png',
    aptae_S: 'image10.png',     // Avec zone de danger R290
    aptae_M: 'image12.png',     // Avec zone de danger R290
    aptae_L: 'image14.png',     // Avec zone de danger R290
    aptae_XL: 'image16.png',    // Avec zone de danger R290
  },

  // ---- Schemas hydrauliques: gamme_mode_nombre ----
  hydraulic: {
    // EFFIPAC - Chauffage seul
    effipac_elec_1: 'image17.png',   // 1 PAC, ballon charge + ECS
    effipac_elec_2: 'image22.png',   // 2 PACs cascade, 2 ballons
    effipac_elec_3: 'image24.png',   // 3 PACs cascade

    // EFFIPAC - Double service (chauffage + ECS)
    effipac_double_1: 'image26.png', // 1 PAC, coupleur ERP, circuits
    effipac_double_2: 'image28.png', // 2 PACs, circuits regules + ECS

    // EFFIPAC - ECS seul
    effipac_ecs_1: 'image25.png',    // 1 PAC, coupleur ERP
    effipac_ecs_2: 'image29.png',    // 2 PACs + secours

    // EFFIPAC - Hybride (PAC + chaudiere)
    effipac_hybrid_1: 'image19.png', // 1 PAC, 2 ballons (charge + suppl)
    effipac_hybrid_2: 'image23.png', // 2 PACs, ballon supplementaire

    // APTAE - Chauffage seul
    aptae_elec_1: 'image36.png',     // 1 PAC 40-50, chauffage
    aptae_elec_2: 'image46.png',     // 2 PACs, chauffage + circuits

    // APTAE - Double service (chauffage + ECS)
    aptae_double_1: 'image30.png',   // 2 PACs, double service + secours
    aptae_double_2: 'image35.png',   // 2 PACs, double service + ECS

    // APTAE - ECS seul
    aptae_ecs_1: 'image31.png',      // Double service ECS focus
    aptae_ecs_2: 'image33.png',      // 2 PACs + ECS

    // APTAE - Hybride
    aptae_hybrid_1: 'image32.png',   // 2 PACs double service variant
    aptae_hybrid_2: 'image34.png',   // 2 PACs complex + circuits

    // Schemas complexes multi-PAC
    aptae_double_3: 'image40.png',   // PAC 40-50, 2 PACs + secours elect
    aptae_elec_3: 'image45.png',     // 2 PACs + 3 circuits regules
  },

  // ---- Schemas supplementaires par configuration specifique ----
  variants: {
    // Effipac chauffage seul variantes
    effipac_elec_1_ballon2: 'image18.png',  // 1 PAC, variante compacte
    effipac_elec_1_suppl: 'image19.png',    // 1 PAC, 2 ballons
    effipac_elec_2_ballon1: 'image21.png',  // 2 PACs, 1 ballon charge

    // Aptae double service variantes
    aptae_double_2_secours: 'image40.png',  // Avec secours electrique
    aptae_double_2_circuits: 'image45.png', // Avec circuits regules
  },

  // ---- Mapping code modele → cle chassis ----
  getChassisKey: function(model) {
    if (!model) return null;
    const gamme = model.refrigerant === 'R290' ? 'aptae' : 'effipac';
    const chassis = model.chassis || 'S';
    return gamme + '_' + chassis;
  },

  // ---- Obtenir le schema hydraulique adapte ----
  getHydraulicKey: function(gamme, mode, nombre) {
    const g = gamme === 'aptae' ? 'aptae' : 'effipac';
    const n = nombre >= 3 ? 3 : nombre;
    const key = g + '_' + mode + '_' + n;
    // Fallback: essayer avec moins de PACs
    if (this.hydraulic[key]) return key;
    const fallback = g + '_' + mode + '_' + Math.min(n, 2);
    if (this.hydraulic[fallback]) return fallback;
    return g + '_' + mode + '_1';
  },

  // ---- Obtenir toutes les images pour une solution ----
  getImagesForSolution: function(model, mode, nombre) {
    const chassisKey = this.getChassisKey(model);
    const hydraulicKey = this.getHydraulicKey(
      model.refrigerant === 'R290' ? 'aptae' : 'effipac',
      mode,
      nombre
    );
    const basePath = 'assets/images/schemas/';

    return {
      photo: this.photos[chassisKey] ? basePath + this.photos[chassisKey] : null,
      dimensions: this.dimensions[chassisKey] ? basePath + this.dimensions[chassisKey] : null,
      degagements: this.degagements[chassisKey] ? basePath + this.degagements[chassisKey] : null,
      hydraulic: this.hydraulic[hydraulicKey] ? basePath + this.hydraulic[hydraulicKey] : null,
      chassisKey: chassisKey,
      hydraulicKey: hydraulicKey
    };
  }
};
