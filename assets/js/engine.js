// ============================================================
// PAC SIZING ENGINE
// Implements: NF EN 12831, EN 14825 bin method, monotone curve,
// ECS ADEME/COSTIC, hybrid bivalent point calculation
// ============================================================

const SizingEngine = {

  // ----------------------------------------------------------
  // 1. CLIMATE & LOCATION
  // ----------------------------------------------------------

  getTbase(departement, altitude, bordDeMer) {
    const d = PERFORMANCE_DATA.tbase[departement];
    if (!d) return null;
    const zone = d.zone;
    const bands = PERFORMANCE_DATA.altitudeBands;
    const corrections = PERFORMANCE_DATA.zoneAltitude[zone];
    if (!corrections) return d.tbase;

    let idx = 0;
    for (let i = 0; i < bands.length; i++) {
      if (altitude <= bands[i]) { idx = i; break; }
      if (i === bands.length - 1) idx = i;
    }
    let tbase = corrections[idx];
    if (bordDeMer) tbase += 2;
    return tbase;
  },

  getClimateZone(departement) {
    const d = PERFORMANCE_DATA.tbase[departement];
    return d ? d.climat : null;
  },

  getBinHours(climat) {
    const zone = climat ? climat.substring(0, 2) : "H1";
    return PERFORMANCE_DATA.binHours[zone] || PERFORMANCE_DATA.binHours.H1;
  },

  // ----------------------------------------------------------
  // 2. BUILDING HEAT LOSS (Deperditions)
  // ----------------------------------------------------------

  calcDeperditions(params) {
    // params: {surface, hauteur, ubat, ventilationType, tInt, tBase}
    // OR: {surface, buildingType, tInt, tBase}
    const tInt = params.tInt || 19;
    const tBase = params.tBase;
    let dp; // W/K

    if (params.buildingType) {
      const bt = PERFORMANCE_DATA.buildingTypes[params.buildingType];
      if (!bt) return null;
      const pTotal = params.surface * bt.specific * 1000; // W at deltaT ~25-30K
      dp = pTotal / (tInt - tBase);
    } else {
      const volume = params.surface * (params.hauteur || 2.5);
      const ventR = PERFORMANCE_DATA.ventilation[params.ventilationType || 'vmc_hygro_b'];
      dp = (params.ubat || 0.8) * params.surface + (ventR ? ventR.R : 0.12) * volume;
    }

    const deperditions = dp * (tInt - tBase); // Watts
    return {
      dp: dp, // W/K
      deperditions: deperditions, // W at Tbase
      deperditionsKW: deperditions / 1000 // kW
    };
  },

  // Quick estimate from surface and building type
  quickEstimate(surface, buildingType, tBase, tInt) {
    tInt = tInt || 19;
    const bt = PERFORMANCE_DATA.buildingTypes[buildingType];
    if (!bt) return null;
    const depKW = surface * bt.specific; // simplified kW
    return {
      deperditionsKW: depKW,
      dp: depKW * 1000 / (tInt - tBase)
    };
  },

  // ----------------------------------------------------------
  // 3. PAC PERFORMANCE INTERPOLATION
  // ----------------------------------------------------------

  // Interpolate PAC capacity and COP at any outdoor temperature
  interpolatePerformance(model, tExt, tWater) {
    const perf = model.performance;
    // Find the two closest reference conditions
    const waterKey = `W${tWater}`;

    // Extract all conditions for this water temperature
    const conditions = [];
    for (const key in perf) {
      if (key.includes(waterKey)) {
        const tAir = parseInt(key.replace('A', '').split('/')[0]);
        conditions.push({ tAir, data: perf[key] });
      }
    }

    if (conditions.length === 0) {
      // Try to find the closest water temperature
      return this._interpolateWater(model, tExt, tWater);
    }

    conditions.sort((a, b) => a.tAir - b.tAir);

    if (conditions.length === 1) {
      return { pcalo: conditions[0].data.pcalo, cop: conditions[0].data.cop, pabs: conditions[0].data.pabs || conditions[0].data.pcalo / conditions[0].data.cop };
    }

    // Find bracket
    let lower = conditions[0], upper = conditions[conditions.length - 1];
    for (let i = 0; i < conditions.length - 1; i++) {
      if (tExt >= conditions[i].tAir && tExt <= conditions[i + 1].tAir) {
        lower = conditions[i];
        upper = conditions[i + 1];
        break;
      }
    }

    // Linear interpolation
    const range = upper.tAir - lower.tAir;
    const ratio = range === 0 ? 0 : (tExt - lower.tAir) / range;

    const pcalo = lower.data.pcalo + ratio * (upper.data.pcalo - lower.data.pcalo);
    const cop = lower.data.cop + ratio * (upper.data.cop - lower.data.cop);
    const pabs = pcalo / cop;

    return { pcalo: Math.round(pcalo * 100) / 100, cop: Math.round(cop * 100) / 100, pabs: Math.round(pabs * 100) / 100 };
  },

  _interpolateWater(model, tExt, tWater) {
    // Try to interpolate between water temperatures
    const temps = [35, 45, 55];
    const results = {};

    for (const tw of temps) {
      const r = this.interpolatePerformance(model, tExt, tw);
      if (r && r.pcalo > 0) results[tw] = r;
    }

    const keys = Object.keys(results).map(Number).sort((a, b) => a - b);
    if (keys.length < 2) return keys.length === 1 ? results[keys[0]] : null;

    let lower = keys[0], upper = keys[keys.length - 1];
    for (let i = 0; i < keys.length - 1; i++) {
      if (tWater >= keys[i] && tWater <= keys[i + 1]) {
        lower = keys[i]; upper = keys[i + 1]; break;
      }
    }

    const ratio = (tWater - lower) / (upper - lower);
    const pcalo = results[lower].pcalo + ratio * (results[upper].pcalo - results[lower].pcalo);
    const cop = results[lower].cop + ratio * (results[upper].cop - results[lower].cop);

    return { pcalo: Math.round(pcalo * 100) / 100, cop: Math.round(cop * 100) / 100, pabs: Math.round((pcalo / cop) * 100) / 100 };
  },

  // ----------------------------------------------------------
  // 4. PAC SELECTION ALGORITHM
  // ----------------------------------------------------------

  selectPAC(params) {
    // params: {deperditionsKW, gamme, tWater, mode, preference}
    // mode: 'elec' (100%), 'hybrid', 'ecs', 'double_service'
    // Returns sorted solutions
    const { deperditionsKW, gamme, tWater, mode, preference, tBase, tInt } = params;

    const gammeData = gamme === 'aptae' ? PERFORMANCE_DATA.aptae : PERFORMANCE_DATA.effipac;
    const solutions = [];

    // Target power range depends on mode
    let targetMin, targetMax;
    if (mode === 'hybrid') {
      targetMin = deperditionsKW * 0.20;
      targetMax = deperditionsKW * 0.50;
    } else {
      targetMin = deperditionsKW * 0.80;
      targetMax = deperditionsKW * 1.20;
    }

    // Try each model and number of units
    for (const model of gammeData.models) {
      for (let n = 1; n <= gammeData.maxUnits; n++) {
        const totalPower = model.puissance_nom * n;

        // Check if total is within acceptable range
        if (totalPower < targetMin * 0.7) continue;
        if (totalPower > gammeData.maxCascade) break;

        // Get performance at Tbase
        const perfAtBase = this.interpolatePerformance(model, tBase || -7, tWater || 45);
        const perfAt7 = model.performance["A7/W35"] || model.performance["A7/W45"];

        if (!perfAtBase) continue;

        const totalPowerAtBase = perfAtBase.pcalo * n;
        const partPAC = (totalPowerAtBase / deperditionsKW) * 100;

        // Coverage rate estimation
        const coverageRate = this._estimateCoverageRate(partPAC, mode);

        // Score the solution
        let score = 0;
        if (mode === 'hybrid') {
          // Hybrid: prefer 30-45% part PAC
          if (partPAC >= 25 && partPAC <= 50) score += 50;
          if (coverageRate >= 70) score += 30;
          if (n <= 3) score += 20 - n * 5;
        } else {
          // Electric: prefer 80-110% coverage
          if (partPAC >= 80 && partPAC <= 120) score += 50;
          if (coverageRate >= 90) score += 30;
          if (n <= 4) score += 20 - n * 4;
        }

        // Prefer fewer large units
        if (preference === 'moins_d_appoints') score += (partPAC > 90) ? 10 : 0;

        solutions.push({
          pac: model,
          nombre: n,
          puissance_totale_nom: Math.round(totalPower * 100) / 100,
          puissance_totale_tbase: Math.round(totalPowerAtBase * 100) / 100,
          part_pac: Math.round(partPAC * 10) / 10,
          taux_couverture: Math.round(coverageRate * 10) / 10,
          cop_a7: perfAt7 ? perfAt7.cop : null,
          cop_tbase: perfAtBase.cop,
          pabs_tbase: Math.round(perfAtBase.pabs * n * 100) / 100,
          score: score
        });
      }
    }

    // Sort by score descending
    solutions.sort((a, b) => b.score - a.score);
    return solutions.slice(0, 6); // Return top 6
  },

  _estimateCoverageRate(partPAC, mode) {
    // Approximate coverage rate from PAC share at Tbase
    // Based on monotone integration heuristics
    if (mode === 'hybrid') {
      // Bivalent parallel
      if (partPAC >= 50) return 92;
      if (partPAC >= 40) return 85;
      if (partPAC >= 30) return 78;
      if (partPAC >= 20) return 68;
      return 55;
    } else {
      // Electric monovalent
      if (partPAC >= 100) return 100;
      if (partPAC >= 90) return 99;
      if (partPAC >= 80) return 97;
      if (partPAC >= 70) return 94;
      if (partPAC >= 60) return 90;
      if (partPAC >= 50) return 85;
      if (partPAC >= 40) return 78;
      return 65;
    }
  },

  // ----------------------------------------------------------
  // 5. ANNUAL ENERGY CALCULATION (Bin Method / EN 14825)
  // ----------------------------------------------------------

  calcAnnualEnergy(params) {
    // params: {deperditionsKW, dp, tInt, tBase, model, nombre, tWater, climat, mode}
    const { deperditionsKW, dp, tInt, tBase, model, nombre, tWater, climat, mode } = params;
    const bins = this.getBinHours(climat);
    const n = nombre || 1;

    let ePac = 0, eBackup = 0, eTotal = 0, elecPac = 0, elecBackup = 0;
    const binResults = [];

    for (const [tStr, hours] of Object.entries(bins)) {
      const t = parseInt(tStr);
      if (t >= (tInt || 19)) continue;

      const load = dp ? dp * ((tInt || 19) - t) / 1000 : deperditionsKW * ((tInt || 19) - t) / ((tInt || 19) - tBase);
      if (load <= 0) continue;

      const perf = this.interpolatePerformance(model, t, tWater || 45);
      const pacCapacity = perf ? perf.pcalo * n : 0;
      const cop = perf ? perf.cop : 1;

      let pacEnergy, backupEnergy;

      if (mode === 'hybrid' && t < -5 && pacCapacity < load * 0.3) {
        // Bivalent alternatif: PAC stops at very low temps
        pacEnergy = 0;
        backupEnergy = load * hours;
      } else {
        pacEnergy = Math.min(pacCapacity, load) * hours;
        backupEnergy = Math.max(0, load - pacCapacity) * hours;
      }

      ePac += pacEnergy;
      eBackup += backupEnergy;
      eTotal += load * hours;
      elecPac += cop > 0 ? pacEnergy / cop : 0;
      elecBackup += backupEnergy; // Direct electric backup or gas

      binResults.push({ t, hours, load: Math.round(load * 10) / 10, pacCapacity: Math.round(pacCapacity * 10) / 10, cop: Math.round(cop * 100) / 100 });
    }

    const tauxCouverture = eTotal > 0 ? (ePac / eTotal) * 100 : 0;
    const scopWeighted = elecPac > 0 ? ePac / elecPac : 0;

    return {
      ePac: Math.round(ePac),
      eBackup: Math.round(eBackup),
      eTotal: Math.round(eTotal),
      elecPac: Math.round(elecPac),
      elecBackup: Math.round(elecBackup),
      tauxCouverture: Math.round(tauxCouverture * 10) / 10,
      scopWeighted: Math.round(scopWeighted * 100) / 100,
      bins: binResults
    };
  },

  // ----------------------------------------------------------
  // 6. ECS SIZING
  // ----------------------------------------------------------

  calcECS(params) {
    // params: {logements (array of {type, count}), parc, tEcs, tEf}
    // OR: {buildingType, units, tEcs, tEf}
    const tEcs = params.tEcs || 60;
    const tEf = params.tEf || 10;

    if (params.logements) {
      // Residential collective
      const equivCoeffs = PERFORMANCE_DATA.ecs.equivalenceParc[params.parc || 'social'];
      let ns = 0;
      for (const log of params.logements) {
        const coeff = equivCoeffs[log.type] || 1.0;
        ns += coeff * log.count;
      }

      const formulas = PERFORMANCE_DATA.ecs.peakFormulas;
      const vj = ns * PERFORMANCE_DATA.ecs.logementStandard; // L/day at 60°C
      const v10 = formulas["10min"].a * Math.pow(ns, formulas["10min"].b);
      const v1h = formulas["1h"].a * Math.pow(ns, formulas["1h"].b);

      // Storage sizing
      const vStockage = v10 * 2.4; // Minimum without loop losses

      // Production power (semi-accumulation)
      const pProd = 14 * Math.pow(vStockage, -0.365) * ns; // kW (permanent circ)
      const pInst = 1.163 * (tEcs - tEf) * v10 / (10 * 60); // kW instantaneous

      return {
        ns: Math.round(ns * 10) / 10,
        vj: Math.round(vj),
        v10: Math.round(v10),
        v1h: Math.round(v1h),
        vStockage: Math.round(vStockage),
        pProduction: Math.round(pProd * 10) / 10,
        pInstantanee: Math.round(pInst * 10) / 10,
        tEcs, tEf
      };
    }

    if (params.buildingType) {
      const bt = PERFORMANCE_DATA.ecs.tertiaire[params.buildingType];
      if (!bt) return null;
      const vj = bt.valeur * (params.units || 1);
      const pAccumulation = 1.163 * vj * (tEcs - tEf) / (8 * 1000); // kW, 8h recharge

      return {
        type: bt.label,
        unite: bt.unite,
        nombre: params.units || 1,
        vj: Math.round(vj),
        pAccumulation: Math.round(pAccumulation * 10) / 10,
        tEcs, tEf
      };
    }

    return null;
  },

  // ----------------------------------------------------------
  // 7. HYDRAULIC SIZING
  // ----------------------------------------------------------

  calcHydraulics(puissanceKW, deltaT) {
    deltaT = deltaT || 5; // PAC typical: 5K
    const debit = puissanceKW / (deltaT * 1.163); // m3/h
    const debitLh = debit * 1000;

    // Find pipe diameter
    const pipes = PERFORMANCE_DATA.pipeDiameters;
    let selectedPipe = pipes[pipes.length - 1];
    for (const pipe of pipes) {
      if (pipe.maxFlow >= debitLh) {
        selectedPipe = pipe;
        break;
      }
    }

    // Buffer tank sizing (inverter)
    const bufferTank = Math.round(14 * puissanceKW); // L for inverter

    return {
      debit_m3h: Math.round(debit * 100) / 100,
      debit_Lh: Math.round(debitLh),
      diametre_int: selectedPipe.int,
      diametre_ext: selectedPipe.ext,
      diametre_bouteille: selectedPipe.int * 3,
      ballon_tampon_L: bufferTank,
      deltaT
    };
  },

  // ----------------------------------------------------------
  // 8. HYBRID BIVALENT POINT
  // ----------------------------------------------------------

  calcBivalentPoint(params) {
    // params: {deperditionsKW, dp, tInt, tBase, model, nombre, tWater}
    const { deperditionsKW, dp, tInt, tBase, model, nombre, tWater } = params;
    const n = nombre || 1;
    const ti = tInt || 19;

    // Search for intersection of load line and PAC capacity
    for (let t = tBase; t <= ti; t += 0.5) {
      const load = deperditionsKW * (ti - t) / (ti - tBase);
      const perf = this.interpolatePerformance(model, t, tWater || 45);
      if (!perf) continue;
      const capacity = perf.pcalo * n;

      if (capacity >= load) {
        return {
          tBivalent: t,
          loadAtBivalent: Math.round(load * 10) / 10,
          capacityAtBivalent: Math.round(capacity * 10) / 10,
          copAtBivalent: perf.cop
        };
      }
    }

    return { tBivalent: ti, loadAtBivalent: 0, capacityAtBivalent: 0, copAtBivalent: 0 };
  },

  // ----------------------------------------------------------
  // 9. COMPETITOR COMPARISON
  // ----------------------------------------------------------

  compareWithCompetitors(atlanticSolution, deperditionsKW, tBase, tWater) {
    const comparisons = [];

    // Atlantic solution
    comparisons.push({
      brand: atlanticSolution.pac.refrigerant === 'R290' ? 'Atlantic Aptae' : 'Atlantic Effipac',
      model: atlanticSolution.pac.nom,
      nombre: atlanticSolution.nombre,
      puissance_nom: atlanticSolution.puissance_totale_nom,
      part_pac: atlanticSolution.part_pac,
      taux_couverture: atlanticSolution.taux_couverture,
      cop_a7: atlanticSolution.cop_a7,
      refrigerant: atlanticSolution.pac.refrigerant,
      t_max: atlanticSolution.pac.t_max,
      prix_ht: atlanticSolution.pac.prix_ht ? atlanticSolution.pac.prix_ht * atlanticSolution.nombre : null,
      isAtlantic: true
    });

    // Check competitors
    for (const [key, comp] of Object.entries(PERFORMANCE_DATA.competitors)) {
      for (const cModel of comp.models) {
        for (let n = 1; n <= 6; n++) {
          const totalPower = cModel.puissance_nom * n;
          if (totalPower < deperditionsKW * 0.6 || totalPower > deperditionsKW * 1.5) continue;

          const perfKey = Object.keys(cModel.performance).find(k => k.includes('W35') || k.includes('W45'));
          const perf = perfKey ? cModel.performance[perfKey] : null;
          if (!perf) continue;

          const partPAC = (totalPower / deperditionsKW) * 100;
          const coverage = this._estimateCoverageRate(partPAC, 'elec');

          comparisons.push({
            brand: comp.brand + ' ' + comp.gamme,
            model: cModel.nom,
            nombre: n,
            puissance_nom: totalPower,
            part_pac: Math.round(partPAC * 10) / 10,
            taux_couverture: Math.round(coverage * 10) / 10,
            cop_a7: perf.cop,
            refrigerant: comp.refrigerant,
            t_max: comp.t_max,
            prix_ht: cModel.prix_ht ? cModel.prix_ht * n : null,
            isAtlantic: false
          });
          break; // Only first valid config per model
        }
      }
    }

    return comparisons;
  },

  // ----------------------------------------------------------
  // 10. GENERATE MONOTONE CURVE DATA (for charts)
  // ----------------------------------------------------------

  generateMonotoneData(params) {
    const { deperditionsKW, dp, tInt, tBase, climat, model, nombre, tWater } = params;
    const bins = this.getBinHours(climat);
    const n = nombre || 1;
    const ti = tInt || 19;

    // Build hourly demand array
    const demands = [];
    for (const [tStr, hours] of Object.entries(bins)) {
      const t = parseInt(tStr);
      if (t >= ti) continue;
      const load = dp ? dp * (ti - t) / 1000 : deperditionsKW * (ti - t) / (ti - tBase);
      if (load <= 0) continue;

      const perf = this.interpolatePerformance(model, t, tWater || 45);
      const pacCap = perf ? perf.pcalo * n : 0;

      for (let h = 0; h < hours; h++) {
        demands.push({ load, pacCap, t });
      }
    }

    // Sort descending by load
    demands.sort((a, b) => b.load - a.load);

    // Sample ~100 points for chart
    const step = Math.max(1, Math.floor(demands.length / 100));
    const chartData = [];
    let cumHours = 0;
    let ePacCum = 0, eTotalCum = 0;

    for (let i = 0; i < demands.length; i++) {
      cumHours++;
      eTotalCum += demands[i].load;
      ePacCum += Math.min(demands[i].load, demands[i].pacCap);

      if (i % step === 0 || i === demands.length - 1) {
        chartData.push({
          hours: cumHours,
          load: Math.round(demands[i].load * 10) / 10,
          pacCapacity: Math.round(demands[i].pacCap * 10) / 10,
          cumEnergy: Math.round(eTotalCum),
          cumPacEnergy: Math.round(ePacCum)
        });
      }
    }

    return chartData;
  }
};

if (typeof window !== 'undefined') window.SizingEngine = SizingEngine;
