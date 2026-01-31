/**
 * ═══════════════════════════════════════════════════════════════
 * ARMITEC - Code.gs v5.3.3 PRODUCTION CORRIGÉ
 * Système de Gestion de Maintenance Industrielle
 * ═══════════════════════════════════════════════════════════════
 * 
 * @author ARMITEC Development Team
 * @version 5.3.3
 * @license MIT
 * @changelog v5.3.3: 
 *   - Correction complète syntaxe
 *   - Fix champs disabled dans payload
 *   - Amélioration gestion gammes
 *   - Compatible Apps Script V8
 */

// ═══════════════════════════════════════════════════════════════
// CONFIGURATION GLOBALE
// ═══════════════════════════════════════════════════════════════

var CONFIG = {
  MODELS_FOLDER_ID: "1BWbBtT4S5MhdC-3rnG2jhlwnbDQwrXJB",
  OUTPUT_FOLDER_ID: "1lLDepJLNxk6VyIP0K4yQPqyfln3MoxUB",
  PIECES_PHOTOS_FOLDER_ID: "1lLDepJLNxk6VyIP0K4yQPqyfln3MoxUB",
  SPREADSHEET_ID: "1klKvOhSYGjXTatyidu60ozDCXj7KWlk8_uG4ImMY1Dk",
  SHEET_INTERVENTIONS: "Interventions",
  SHEET_PIECES: "Pièces détachées",
  SHEET_LOGS: "Logs",
  SHEET_PHOTOS: "Photos",
  SHEET_FORM_SCHEMA: "FORM_SCHEMA",
  SHEET_MACHINES: "Machines",
  SHEET_GAMMES: "Gammes",
  SOCLE_MAX_ORDER: 24,
  CACHE_TTL: 300,
  MAX_FILE_SIZE_MB: 10,
  MAX_RETRIES: 3,
  RETRY_DELAY_MS: 1000
};

var SHEET_HEADERS = {
  INTERVENTIONS: ["ID","Horodatage","Date","Site","Machine","Technicien","Type","Temps","Common","Specific","GammeKey","CorrectifDetecte","Statut","Description"],
  PIECES: ["ID","Intervention","Horodatage","PieceID","Machine","Zone","Référence","Désignation","Quantité","ContexteDécouverte","ActionImmédiate","Criticité","Origine","Usage","Décision","Impact","PrixUnitaireHT","Fournisseur","Commentaire","PhotosURLs"],
  PHOTOS: ["Intervention","Horodatage","Filename","Url","Type"],
  LOGS: ["Date","Level","Source","Message","Details","Duration","User"]
};

// ═══════════════════════════════════════════════════════════════
// UTILITAIRES DE BASE
// ═══════════════════════════════════════════════════════════════

function retryOperation(operation, context, maxRetries) {
  maxRetries = maxRetries || CONFIG.MAX_RETRIES;
  var startTime = new Date().getTime();
  
  for (var attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      var result = operation();
      return result;
    } catch (error) {
      if (attempt === maxRetries) {
        Logger.log("[ERROR] " + context + " - Échec après " + maxRetries + " tentatives: " + error.message);
        throw error;
      }
      
      var retryDelay = CONFIG.RETRY_DELAY_MS * attempt;
      Logger.log("[WARN] " + context + " - Tentative " + attempt + "/" + maxRetries + " échouée, retry dans " + retryDelay + "ms");
      Utilities.sleep(retryDelay);
    }
  }
}

function validatePayload(payload) {
  if (!payload) throw new Error("Payload vide");
  if (!payload.common) throw new Error("Données communes manquantes");
  
  var required = ['machine_id', 'type_maintenance'];
  for (var i = 0; i < required.length; i++) {
    if (!payload.common[required[i]]) {
      throw new Error("Champ requis manquant: " + required[i]);
    }
  }
  
  if (payload.pieces && Array.isArray(payload.pieces)) {
    for (var j = 0; j < payload.pieces.length; j++) {
      var piece = payload.pieces[j];
      if (piece.quantite && (isNaN(piece.quantite) || piece.quantite <= 0)) {
        throw new Error("Quantité invalide pour pièce " + j);
      }
    }
  }
}

function sanitizeInput(input) {
  if (!input) return '';
  return String(input)
    .replace(/<[^>]*>/g, '')
    .replace(/[<>"'&]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .substring(0, 1000);
}

function log(level, source, message, details) {
  details = details || null;
  
  try {
    var now = new Date();
    var user = Session.getActiveUser().getEmail() || 'system';
    var duration = (details && details.duration) ? details.duration : '0ms';
    
    var logMessage = "[" + level + "] " + source + ": " + message;
    Logger.log(logMessage);
    
    retryOperation(function() {
      var sheet = getOrCreateSheet(CONFIG.SHEET_LOGS, SHEET_HEADERS.LOGS);
      sheet.appendRow([
        now, 
        level, 
        source, 
        message, 
        details ? JSON.stringify(details) : '', 
        duration,
        user
      ]);
    }, 'log_to_sheet', 2);
    
  } catch (error) {
    Logger.log("[ERROR] Erreur logging: " + error.message);
  }
}

function logError(source, error) {
  var message = error && error.message ? error.message : String(error);
  var stack = error && error.stack ? error.stack : '';
  log("ERROR", source, message, { 
    stack: stack,
    type: error.constructor ? error.constructor.name : 'Error'
  });
}

function getOrCreateSheet(name, headers) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(name);
    
    if (!sheet) {
      sheet = ss.insertSheet(name);
      if (Array.isArray(headers) && headers.length > 0) {
        sheet.appendRow(headers);
        sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      }
    }
    
    return sheet;
  } catch (error) {
    logError('getOrCreateSheet', error);
    throw error;
  }
}

function generateInterventionId(prefix) {
  prefix = prefix || 'INT';
  var now = new Date();
  var dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyyMMdd");
  var timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "HHmmss");
  var random = Math.random().toString(36).substring(2, 6).toUpperCase();
  return prefix + "-" + dateStr + "-" + timeStr + "-" + random;
}

// ═══════════════════════════════════════════════════════════════
// WEBAPP ENTRYPOINTS
// ═══════════════════════════════════════════════════════════════

function doGet(e) {
  var startTime = new Date().getTime();
  log("INFO", "doGet", "Début requête GET");
  
  try {
    var params = e && e.parameter ? e.parameter : {};
    var template = HtmlService.createTemplateFromFile('index');
    
    template.prefill = {
      machine: sanitizeInput(params.machine || ''),
      type: sanitizeInput(params.type || ''),
      gamme: sanitizeInput(params.gamme || ''),
      site: sanitizeInput(params.site || '')
    };
    
    var result = template.evaluate()
      .setTitle('ARMITEC - Déclaration d\'intervention')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
    
    var duration = new Date().getTime() - startTime;
    log("INFO", "doGet", "Requête GET traitée", { duration: duration + "ms" });
    
    return result;
    
  } catch (error) {
    logError('doGet', error);
    return HtmlService.createHtmlOutput('<h1>Erreur</h1><p>' + error.message + '</p>');
  }
}

function doPost(e) {
  var startTime = new Date().getTime();
  log("INFO", "doPost", "Début requête POST");
  
  try {
    var payload = e && e.postData && e.postData.contents ? JSON.parse(e.postData.contents) : {};
    var result = submitIntervention(payload);
    
    var duration = new Date().getTime() - startTime;
    log("INFO", "doPost", "Requête POST traitée", { 
      duration: duration + "ms",
      interventionId: result.id
    });
    
    return ContentService.createTextOutput(JSON.stringify({ 
      status: 'OK', 
      result: result,
      processingTime: duration + "ms"
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    logError('doPost', error);
    return ContentService.createTextOutput(JSON.stringify({ 
      status: 'ERROR', 
      message: error.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function include(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (error) {
    logError('include', error);
    return '<!-- Erreur inclusion ' + filename + ': ' + error.message + ' -->';
  }
}

// ═══════════════════════════════════════════════════════════════
// SOUMISSION INTERVENTION
// ═══════════════════════════════════════════════════════════════

function submitIntervention(payload) {
  var startTime = new Date().getTime();
  var interventionId = generateInterventionId('INT');
  
  log("INFO", "submitIntervention", "Début soumission " + interventionId);
  
  try {
    validatePayload(payload);
    
    var now = new Date();
    var common = payload.common || {};
    var specific = payload.specific || {};
    var pieces = Array.isArray(payload.pieces) ? payload.pieces : [];
    var gammeKey = specific.gamme_key || payload.gamme_key || common.selected_gamme_key || '';
    
    var correctifDetecte = detectCorrectifNeeded(pieces);
    
    var interventionSheet = retryOperation(function() {
      return getOrCreateSheet(CONFIG.SHEET_INTERVENTIONS, SHEET_HEADERS.INTERVENTIONS);
    }, 'get_intervention_sheet');
    
    retryOperation(function() {
      interventionSheet.appendRow([
        interventionId,
        now,
        common.date_intervention || Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd"),
        common.site || "",
        common.machine_id || "",
        common.technicien || Session.getActiveUser().getEmail(),
        common.type_maintenance || "",
        common.temps_passe || "",
        JSON.stringify(common),
        JSON.stringify(specific),
        gammeKey,
        correctifDetecte ? 'OUI' : 'NON',
        'Créée',
        common.description || ""
      ]);
    }, 'insert_intervention');
    
    log("INFO", "submitIntervention", "Intervention " + interventionId + " créée");
    
    if (pieces.length > 0) {
      processPieces(interventionId, pieces, common.machine_id, now);
      if (correctifDetecte) {
        createCorrectiveInterventions(interventionId, pieces, common.machine_id);
      }
    }
    
    var duration = new Date().getTime() - startTime;
    log("INFO", "submitIntervention", "Intervention finalisée", {
      duration: duration + "ms",
      id: interventionId,
      piecesCount: pieces.length
    });
    
    return { 
      success: true, 
      message: "Intervention enregistrée", 
      id: interventionId, 
      piecesCount: pieces.length, 
      clearForm: true,
      processingTime: duration + "ms"
    };
    
  } catch (error) {
    logError('submitIntervention', error);
    throw new Error("Enregistrement échoué : " + error.message);
  }
}

function processPieces(interventionId, pieces, machineId, timestamp) {
  var startTime = new Date().getTime();
  log("INFO", "processPieces", "Traitement " + pieces.length + " pièces");
  
  try {
    var piecesSheet = retryOperation(function() {
      return getOrCreateSheet(CONFIG.SHEET_PIECES, SHEET_HEADERS.PIECES);
    }, 'get_pieces_sheet');
    
    var rows = pieces.map(function(piece) {
      return [
        Utilities.getUuid(),
        interventionId,
        timestamp,
        piece.piece_id || Utilities.getUuid(),
        piece.machine_id || machineId || "",
        piece.zone_composant || "",
        piece.reference || "",
        piece.designation || "",
        piece.quantite || 1,
        piece.contexte_decouverte || "",
        piece.action_immediate || "",
        piece.criticite || "",
        piece.origine || "",
        piece.usage || "",
        piece.decision || "",
        piece.impact_attendu || "",
        piece.prix_unitaire_ht || "",
        piece.fournisseur || "",
        piece.commentaire || "",
        ""
      ];
    });
    
    if (rows.length > 0) {
      retryOperation(function() {
        var startRow = piecesSheet.getLastRow() + 1;
        piecesSheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
      }, 'batch_insert_pieces');
    }
    
    var duration = new Date().getTime() - startTime;
    log("INFO", "processPieces", pieces.length + " pièces enregistrées", {
      duration: duration + "ms"
    });
    
    processDecisionsPieces(pieces);
    
  } catch (error) {
    logError('processPieces', error);
    throw error;
  }
}

function detectCorrectifNeeded(pieces) {
  if (!Array.isArray(pieces)) return false;
  
  return pieces.some(function(piece) {
    return piece.decision === 'commander_pour_intervention' ||
           piece.decision === 'Commander pour intervention' ||
           (piece.action_immediate === 'Non' && (piece.criticite === 'Haute' || piece.criticite === 'Critique'));
  });
}

// ═══════════════════════════════════════════════════════════════
// HELPER : RECHERCHE FICHIER MODÈLE
// ═══════════════════════════════════════════════════════════════

function getModelFileForMachine(folder, machineId) {
  var startTime = new Date().getTime();
  
  machineId = String(machineId || '').trim();
  if (!machineId) return null;
  
  var candidates = [
    machineId + ".json",
    machineId + "_model.json"
  ];

  for (var i = 0; i < candidates.length; i++) {
    var fileName = candidates[i];
    
    try {
      var files = folder.getFilesByName(fileName);
      if (files.hasNext()) {
        var file = files.next();
        var duration = new Date().getTime() - startTime;
        log("INFO", "getModelFileForMachine", "Fichier trouvé: " + fileName, {
          duration: duration + "ms",
          fileSize: (file.getSize() / 1024).toFixed(2) + " KB"
        });
        return file;
      }
    } catch (error) {
      log("WARN", "getModelFileForMachine", "Erreur accès " + fileName + ": " + error.message);
    }
  }
  
  log("WARN", "getModelFileForMachine", "Aucun fichier pour " + machineId);
  return null;
}

// ═══════════════════════════════════════════════════════════════
// GESTION DES GAMMES
// ═══════════════════════════════════════════════════════════════

function getAvailableGammes(machineId) {
  var startTime = new Date().getTime();
  log("INFO", "getAvailableGammes", "Récupération gammes pour " + (machineId || "non défini"));
  
  try {
    var gammes = getGammesForMachine(machineId);
    var duration = new Date().getTime() - startTime;
    
    log("INFO", "getAvailableGammes", gammes.length + " gammes récupérées", {
      duration: duration + "ms",
      machineId: machineId || "non défini",
      gammesList: gammes.length > 0 ? JSON.stringify(gammes.map(function(g) { return g.key; })) : "aucune gamme"
    });
    
    return gammes;
  } catch (error) {
    logError('getAvailableGammes', error);
    log("ERROR", "getAvailableGammes", "Erreur lors de la récupération des gammes: " + error.message);
    return [];
  }
}

// ═══════════════════════════════════════════════════════════════
// CORRECTION MAJEURE : CHARGEMENT GAMMES DEPUIS GOOGLE SHEETS
// ═══════════════════════════════════════════════════════════════

/**
 * Récupère les gammes disponibles pour une machine donnée
 * NOUVELLE VERSION : Lit depuis Google Sheets (feuille "Gammes")
 */
function getGammesForMachine(machineId) {
  var startTime = new Date().getTime();
  log("INFO", "getGammesForMachine", "🔍 Chargement gammes pour machine: " + (machineId || "non défini"));

  if (!machineId) {
    log("WARN", "getGammesForMachine", "❌ ID machine vide");
    return [];
  }

  try {
    // ✅ Étape 1 : Vérifier le cache
    var cache = CacheService.getScriptCache();
    var cacheKey = 'gammes_' + machineId;
    var cached = cache.get(cacheKey);
    
    if (cached) {
      try {
        var result = JSON.parse(cached);
        var duration = new Date().getTime() - startTime;
        log("INFO", "getGammesForMachine", "✅ Cache HIT - " + result.length + " gammes", {
          duration: duration + "ms",
          gammesKeys: result.map(function(g) { return g.key; }).join(", ")
        });
        return result;
      } catch (parseError) {
        log("WARN", "getGammesForMachine", "⚠️ Cache corrompu, suppression");
        cache.remove(cacheKey);
      }
    }

    log("DEBUG", "getGammesForMachine", "💾 Cache MISS, chargement depuis Google Sheets");

    // ✅ Étape 2 : Lire la feuille "Gammes"
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.SHEET_GAMMES);
    
    if (!sheet) {
      log("ERROR", "getGammesForMachine", "❌ Feuille 'Gammes' introuvable");
      return [];
    }

    log("INFO", "getGammesForMachine", "📊 Feuille 'Gammes' trouvée");

    var data = sheet.getDataRange().getValues();
    
    if (data.length < 2) {
      log("WARN", "getGammesForMachine", "⚠️ Feuille 'Gammes' vide (pas de données)");
      return [];
    }

    log("DEBUG", "getGammesForMachine", "📋 " + (data.length - 1) + " lignes de données dans la feuille");

    // ✅ Étape 3 : Parser les en-têtes
    var headers = data[0];
    var colIndex = {
      machine: headers.indexOf("Machine"),
      type: headers.indexOf("Type"),
      codeGamme: headers.indexOf("Code Gamme"),
      libelle: headers.indexOf("Libellé"),
      url: headers.indexOf("Url"),
      periodicite: headers.indexOf("Périodicité"),
      statut: headers.indexOf("Statut")
    };

    log("DEBUG", "getGammesForMachine", "📑 Colonnes identifiées:", {
      machine: colIndex.machine,
      type: colIndex.type,
      codeGamme: colIndex.codeGamme,
      libelle: colIndex.libelle,
      url: colIndex.url,
      statut: colIndex.statut
    });

    // Vérifier que les colonnes essentielles existent
    if (colIndex.machine === -1 || colIndex.codeGamme === -1) {
      log("ERROR", "getGammesForMachine", "❌ Colonnes essentielles manquantes dans la feuille");
      return [];
    }

    // ✅ Étape 4 : Filtrer les gammes pour cette machine
    var gammes = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      
      var machine = colIndex.machine >= 0 ? String(row[colIndex.machine] || "").trim() : "";
      var type = colIndex.type >= 0 ? String(row[colIndex.type] || "").trim().toLowerCase() : "";
      var codeGamme = colIndex.codeGamme >= 0 ? String(row[colIndex.codeGamme] || "").trim() : "";
      var libelle = colIndex.libelle >= 0 ? String(row[colIndex.libelle] || "").trim() : "";
      var url = colIndex.url >= 0 ? String(row[colIndex.url] || "").trim() : "";
      var statut = colIndex.statut >= 0 ? String(row[colIndex.statut] || "").trim() : "";

      log("DEBUG", "getGammesForMachine", "📝 Ligne " + (i + 1) + ":", {
        machine: machine,
        type: type,
        codeGamme: codeGamme,
        statut: statut,
        matches: (machine === machineId && type === 'preventif' && statut === 'Actif')
      });

      // Filtrer : même machine, type préventif, statut actif
      if (machine === machineId && type === 'preventif' && statut === 'Actif') {
        gammes.push({
          key: codeGamme,
          label: libelle || codeGamme,
          url: url,
          periodicite: colIndex.periodicite >= 0 ? String(row[colIndex.periodicite] || "") : "",
          stepsCount: 0 // Sera mis à jour lors du chargement du schéma
        });

        log("INFO", "getGammesForMachine", "✅ Gamme ajoutée:", {
          key: codeGamme,
          label: libelle,
          url: url ? "OUI" : "NON"
        });
      }
    }

    log("INFO", "getGammesForMachine", "📊 Résultat: " + gammes.length + " gammes trouvées pour " + machineId);

    // ✅ Étape 5 : Mettre en cache
    if (gammes.length > 0) {
      try {
        cache.put(cacheKey, JSON.stringify(gammes), CONFIG.CACHE_TTL || 300);
        log("INFO", "getGammesForMachine", "💾 Gammes mises en cache");
      } catch (cacheError) {
        log("WARN", "getGammesForMachine", "⚠️ Échec mise en cache: " + cacheError.message);
      }
    }

    var duration = new Date().getTime() - startTime;
    log("INFO", "getGammesForMachine", "✅ Terminé en " + duration + "ms");

    return gammes;

  } catch (error) {
    logError('getGammesForMachine', error);
    log("ERROR", "getGammesForMachine", "❌ Erreur générale: " + error.message);
    return [];
  }
}

// Fonction de secours (fallback) pour lire les gammes depuis Google Sheets
function getGammesFromSheetsFallback(machineId, startTime, reason) {
  log("INFO", "getGammesFromSheetsFallback", "Tentative de fallback vers Google Sheets", {
    reason: reason,
    machineId: machineId || "non défini"
  });

  try {
    // Lire les données depuis la feuille Google Sheets
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = "Gammes"; // Adaptez ce nom si votre feuille s'appelle différemment
    var sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      log("ERROR", "getGammesFromSheetsFallback", "Feuille introuvable: " + sheetName);
      return [];
    }
    log("INFO", "getGammesFromSheetsFallback", "Feuille trouvée: " + sheetName);

    var data = sheet.getDataRange().getValues();
    log("DEBUG", "getGammesFromSheetsFallback", "Données lues", {
      rowCount: data.length
    });

    var gammes = [];
    
    // Parcourir les lignes de la feuille (en supposant que la première ligne est l'en-tête)
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row.length < 9) continue; // Ignorer les lignes incomplètes
      
      var machine = row[0]; // Colonne A: Machine
      var type = row[1]; // Colonne B: Type
      var codeGamme = row[2]; // Colonne C: Code Gamme
      var libelle = row[3]; // Colonne D: Libellé
      var statut = row[8]; // Colonne I: Statut (ajustez si nécessaire)
      
      // Vérifier si la ligne correspond à la machine et au type preventif avec statut Actif
      if (machine === machineId && type.toLowerCase() === 'preventif' && statut === 'Actif') {
        gammes.push({
          key: codeGamme,
          label: libelle,
          url: row[5] || '', // Colonne F: Url (ajustez si nécessaire)
          stepsCount: 0 // Valeur par défaut, ajustez si vous avez des données pour cela
        });
        log("DEBUG", "getGammesFromSheetsFallback", "Gamme ajoutée depuis Sheets", {
          key: codeGamme,
          label: libelle
        });
      }
    }
    
    var duration = new Date().getTime() - startTime;
    log("INFO", "getGammesFromSheetsFallback", "Gammes chargées depuis Google Sheets", {
      duration: duration + "ms",
      gammesCount: gammes.length,
      gammesList: gammes.length > 0 ? JSON.stringify(gammes.map(function(g) { return g.key; })) : "aucune gamme"
    });

    // Mettre en cache les résultats du fallback si des gammes sont trouvées
    if (gammes.length > 0) {
      try {
        var cache = CacheService.getScriptCache();
        var cacheKey = 'gammes_' + machineId;
        cache.put(cacheKey, JSON.stringify(gammes), CONFIG.CACHE_TTL || 3600);
        log("INFO", "getGammesFromSheetsFallback", "Gammes mises en cache depuis fallback", {
          cacheKey: cacheKey
        });
      } catch (cacheError) {
        log("WARN", "getGammesFromSheetsFallback", "Échec mise en cache fallback: " + cacheError.message);
      }
    }

    return gammes;

  } catch (error) {
    logError('getGammesFromSheetsFallback', error);
    log("ERROR", "getGammesFromSheetsFallback", "Erreur lors du fallback: " + error.message);
    return [];
  }
}

/**
 * Récupère le schéma d'une gamme spécifique
 * NOUVELLE VERSION : Utilise l'URL de la feuille Gammes
 */
function getGammeSchema(machineId, gammeKey) {
  var startTime = new Date().getTime();
  log("INFO", "getGammeSchema", "🔍 Récupération schéma gamme: " + gammeKey + " pour machine: " + machineId);
  
  if (!machineId || !gammeKey) {
    log("WARN", "getGammeSchema", "❌ Paramètres manquants", {
      machineId: machineId || "manquant",
      gammeKey: gammeKey || "manquant"
    });
    return { fields: [], images: [], title: '' };
  }
  
  try {
    // ✅ Étape 1 : Récupérer l'URL depuis la feuille Gammes
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.SHEET_GAMMES);
    
    if (!sheet) {
      log("ERROR", "getGammeSchema", "❌ Feuille 'Gammes' introuvable");
      return { fields: [], images: [], title: gammeKey };
    }

    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    
    var colIndex = {
      machine: headers.indexOf("Machine"),
      codeGamme: headers.indexOf("Code Gamme"),
      libelle: headers.indexOf("Libellé"),
      url: headers.indexOf("Url")
    };

    log("DEBUG", "getGammeSchema", "📑 Recherche de la gamme dans la feuille");

    var gammeUrl = null;
    var gammeLabel = gammeKey;

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var machine = colIndex.machine >= 0 ? String(row[colIndex.machine] || "").trim() : "";
      var code = colIndex.codeGamme >= 0 ? String(row[colIndex.codeGamme] || "").trim() : "";
      
      if (machine === machineId && code === gammeKey) {
        gammeUrl = colIndex.url >= 0 ? String(row[colIndex.url] || "").trim() : "";
        gammeLabel = colIndex.libelle >= 0 ? String(row[colIndex.libelle] || "").trim() : gammeKey;
        
        log("INFO", "getGammeSchema", "✅ Gamme trouvée dans la feuille", {
          machine: machine,
          code: code,
          label: gammeLabel,
          url: gammeUrl || "URL manquante",
          urlLength: gammeUrl ? gammeUrl.length : 0
        });
        break;
      }
    }

    if (!gammeUrl) {
      log("ERROR", "getGammeSchema", "❌ URL de gamme non trouvée pour " + gammeKey);
      return { 
        fields: [
          {
            name: 'gamme_note',
            label: 'Notes de la gamme ' + gammeLabel,
            type: 'textarea',
            required: false,
            gamme: gammeKey
          }
        ], 
        images: [], 
        title: gammeLabel 
      };
    }

    // ✅ Étape 2 : Extraire l'ID du fichier Drive depuis l'URL
    var fileId = extractDriveFileId(gammeUrl);
    
    if (!fileId) {
      log("ERROR", "getGammeSchema", "❌ Impossible d'extraire l'ID du fichier depuis l'URL", {
        url: gammeUrl
      });
      return { fields: [], images: [], title: gammeLabel };
    }

    log("INFO", "getGammeSchema", "📄 ID du fichier extrait: " + fileId);

    // ✅ Étape 3 : Charger le fichier depuis Drive
    var file;
    try {
      file = DriveApp.getFileById(fileId);
      log("INFO", "getGammeSchema", "✅ Fichier Drive trouvé", {
        name: file.getName(),
        size: (file.getSize() / 1024).toFixed(2) + " KB",
        mimeType: file.getMimeType()
      });
    } catch (driveError) {
      log("ERROR", "getGammeSchema", "❌ Erreur accès fichier Drive: " + driveError.message, {
        fileId: fileId
      });
      return { fields: [], images: [], title: gammeLabel };
    }

    // ✅ Étape 4 : Lire et parser le JSON
    var modelText;
    try {
      modelText = file.getBlob().getDataAsString("UTF-8");
      log("DEBUG", "getGammeSchema", "📖 Contenu fichier lu: " + modelText.length + " caractères");
    } catch (readError) {
      log("ERROR", "getGammeSchema", "❌ Erreur lecture fichier: " + readError.message);
      return { fields: [], images: [], title: gammeLabel };
    }

    var model;
    try {
      model = JSON.parse(modelText);
      log("DEBUG", "getGammeSchema", "✅ JSON parsé", {
        hasTimeline: !!model.timeline,
        timelineLength: (model.timeline || []).length
      });
    } catch (parseError) {
      log("ERROR", "getGammeSchema", "❌ JSON invalide: " + parseError.message);
      return { fields: [], images: [], title: gammeLabel };
    }

    // ✅ Étape 5 : Extraire TOUTE la timeline
    if (!model.timeline || !Array.isArray(model.timeline)) {
      log("ERROR", "getGammeSchema", "❌ Timeline invalide ou absente");
      return { fields: [], images: [], title: gammeLabel };
    }

    // ✅ CORRECTION : Utiliser extractGammeFromTimeline au lieu de extractWholeTimeline
    var gammeData = extractGammeFromTimeline(model.timeline, gammeKey);
    
    // ✅ Vérification : gamme valide ?
    if (!gammeData || !gammeData.fields || gammeData.fields.length === 0) {
      log("ERROR", "getGammeSchema", "❌ Extraction gamme échouée (aucun champ)");
      return { 
        fields: [], 
        images: [], 
        title: gammeLabel 
      };
    }

    // ✅ Utiliser le label de la feuille Gammes comme titre
    gammeData.title = gammeLabel;

    var duration = new Date().getTime() - startTime;
    log("INFO", "getGammeSchema", "✅ Schéma récupéré en " + duration + "ms", {
      fieldsCount: gammeData.fields.length,
      imagesCount: gammeData.images.length,
      title: gammeData.title
    });
    
    return gammeData;
    
  } catch (error) {
    logError('getGammeSchema', error);
    return { fields: [], images: [], title: gammeKey };
  }
}

/**
 * Extrait l'ID d'un fichier Google Drive depuis une URL
 * Gère plusieurs formats d'URL Drive
 */
function extractDriveFileId(url) {
  if (!url) return null;
  
  url = String(url).trim();
  
  log("DEBUG", "extractDriveFileId", "🔗 Extraction ID depuis URL", {
    url: url,
    urlLength: url.length
  });
  
  // Format 1 : https://drive.google.com/file/d/FILE_ID/view
  var match1 = url.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (match1) {
    log("INFO", "extractDriveFileId", "✅ ID extrait (format 1): " + match1[1]);
    return match1[1];
  }
  
  // Format 2 : https://drive.google.com/open?id=FILE_ID
  var match2 = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (match2) {
    log("INFO", "extractDriveFileId", "✅ ID extrait (format 2): " + match2[1]);
    return match2[1];
  }
  
  // Format 3 : https://docs.google.com/document/d/FILE_ID/edit
  var match3 = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (match3) {
    log("INFO", "extractDriveFileId", "✅ ID extrait (format 3): " + match3[1]);
    return match3[1];
  }
  
  // Si c'est déjà juste un ID (pas d'URL)
  if (url.match(/^[a-zA-Z0-9_-]+$/)) {
    log("INFO", "extractDriveFileId", "✅ Déjà un ID: " + url);
    return url;
  }
  
  log("ERROR", "extractDriveFileId", "❌ Impossible d'extraire l'ID depuis: " + url);
  return null;
}

/**
 * Extrait les gammes depuis une timeline
 * (Cette fonction reste inchangée mais avec plus de logs)
 */
function extractGammesFromTimeline(timeline) {
  log("DEBUG", "extractGammesFromTimeline", "🔍 Extraction gammes depuis timeline (" + timeline.length + " items)");

  if (!Array.isArray(timeline)) {
    log("ERROR", "extractGammesFromTimeline", "❌ Timeline invalide (pas un tableau)");
    return [];
  }

  var gammes = [];
  var currentGamme = null;

  for (var i = 0; i < timeline.length; i++) {
    var item = timeline[i] || {};

    // Détecter une nouvelle section de gamme
if (item.kind === 'section') {
  var order = parseInt(item.order, 10) || 0;

  // Si order >= seuil, c'est une gamme
  if (order >= (CONFIG.SOCLE_MAX_ORDER + 1)) {

    // 🔒 Finaliser la gamme précédente UNIQUEMENT si elle a des steps
    if (currentGamme) {
      if (currentGamme.stepsCount > 0) {
        gammes.push(currentGamme);
      } else {
        log(
          "WARN",
          "extractGammesFromTimeline",
          "❌ Gamme ignorée (aucun champ)",
          currentGamme
        );
      }
    }

    // Initialiser nouvelle gamme
    currentGamme = {
      key: item.section_id || ('section_' + i),
      label: (item.title && String(item.title).trim())
        ? String(item.title).trim()
        : (item.section_id || 'Section'),
      order: order,
      stepsCount: 0
    };

    log(
      "DEBUG",
      "extractGammesFromTimeline",
      "📌 Nouvelle gamme détectée",
      currentGamme
    );

  } else if (currentGamme && order < (CONFIG.SOCLE_MAX_ORDER + 1)) {

    // 🔒 Fin de gamme : ne garder que si elle contient des steps
    if (currentGamme.stepsCount > 0) {
      gammes.push(currentGamme);
    } else {
      log(
        "WARN",
        "extractGammesFromTimeline",
        "❌ Gamme ignorée (aucun champ)",
        currentGamme
      );
    }

    currentGamme = null;
  }
}

// Compter les steps dans la gamme actuelle
if (currentGamme && item.kind === 'field') {
  currentGamme.stepsCount++;
}

  }

  // Ajouter la dernière gamme
  if (currentGamme) {
    gammes.push(currentGamme);
  }

  log("INFO", "extractGammesFromTimeline", "✅ " + gammes.length + " gammes extraites", {
    gammesKeys: gammes.map(function(g) { return g.key; }).join(", ")
  });

  return gammes;
}

// ═══════════════════════════════════════════════════════════════
// FONCTION POUR TESTER LE SYSTÈME DEPUIS LE MENU
// ═══════════════════════════════════════════════════════════════

function testGammeLoading() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    "🧪 Test chargement gamme", 
    "Entrez l'ID de la machine (ex: M015):", 
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  var machineId = response.getResponseText().trim();
  if (!machineId) {
    ui.alert("❌ ID machine requis");
    return;
  }
  
  try {
    log("INFO", "testGammeLoading", "🧪 TEST DÉMARRÉ pour machine: " + machineId);
    
    // Test 1 : Charger les gammes
    var gammes = getGammesForMachine(machineId);
    
    var msg = "📋 TEST GAMMES - Machine " + machineId + "\n\n";
    msg += "✅ Gammes trouvées : " + gammes.length + "\n\n";
    
    if (gammes.length === 0) {
      msg += "❌ Aucune gamme disponible\n";
      msg += "Vérifiez :\n";
      msg += "1. Feuille 'Gammes' existe\n";
      msg += "2. Machine = " + machineId + "\n";
      msg += "3. Type = 'preventif'\n";
      msg += "4. Statut = 'Actif'\n";
    } else {
      for (var i = 0; i < gammes.length; i++) {
        var g = gammes[i];
        msg += (i + 1) + ". " + g.label + "\n";
        msg += "   Code: " + g.key + "\n";
        msg += "   URL: " + (g.url ? "✅ Présente" : "❌ Manquante") + "\n\n";
      }
      
      // Test 2 : Charger le schéma de la première gamme
      if (gammes.length > 0) {
        var firstGamme = gammes[0];
        msg += "\n🔍 Test chargement schéma: " + firstGamme.key + "\n\n";
        
        var schema = getGammeSchema(machineId, firstGamme.key);
        msg += "   Champs: " + schema.fields.length + "\n";
        msg += "   Images: " + schema.images.length + "\n";
        msg += "   Titre: " + schema.title + "\n";
      }
    }
    
    msg += "\n📊 Consultez les logs (Outils > Historique des exécutions)";
    
    ui.alert(msg);
    
  } catch (error) {
    logError('testGammeLoading', error);
    ui.alert("❌ Erreur: " + error.message + "\n\nConsultez les logs");
  }
}
/**
 * Extrait TOUTE la timeline d'un fichier de gamme
 * (1 fichier = 1 gamme complète)
 */
function extractGammeFromTimeline(timeline, gammeKey) {
  var startTime = new Date().getTime();

  if (!Array.isArray(timeline)) {
    log("ERROR", "extractGammeFromTimeline", "❌ Timeline invalide");
    return null;
  }

  var title = gammeKey; // Par défaut
  var fields = [];
  var images = [];
  var stepOrder = 0;

  log("INFO", "extractGammeFromTimeline", "🔍 Extraction de TOUTE la timeline (" + timeline.length + " items)");

  // ✅ Parcourir TOUS les éléments de la timeline
  for (var i = 0; i < timeline.length; i++) {
    var item = timeline[i] || {};

    // 📋 Section = titre de la gamme (prendre le premier)
    if (item.kind === 'section' && !title && item.title) {
      title = String(item.title).trim();
      log("DEBUG", "extractGammeFromTimeline", "📌 Titre trouvé: " + title);
    }

    // 📝 Champs (field)
    if (item.kind === 'field' && item.field) {
      stepOrder++;
      var field = item.field;

      var fieldData = {
        name: "gamme_" + gammeKey + "_" + (field.key || ("step_" + stepOrder)),
        label: field.label || field.key || ("Étape " + stepOrder),
        type: field.type || 'text',
        required: !!field.required,
        options: parseOptions(field.options),
        gamme: gammeKey,
        stepOrder: stepOrder,
        key: field.key || ''
      };

      fields.push(fieldData);
      log("DEBUG", "extractGammeFromTimeline", "✅ Champ #" + stepOrder + ": " + field.label);
    }

    // 🖼️ Images
    if (item.kind === 'image' && item.media && item.media.drive_id) {
      var imageData = {
        drive_id: item.media.drive_id,
        filename: item.media.filename || '',
        caption: item.media.caption || '',
        order: item.order || 0
      };

      images.push(imageData);
      log("DEBUG", "extractGammeFromTimeline", "✅ Image: " + imageData.caption);
    }
  }

  // ❌ Vérification : au moins 1 champ requis
  if (fields.length === 0) {
    log("ERROR", "extractGammeFromTimeline", "❌ Aucun champ trouvé dans la timeline");
    return null;
  }

  var duration = new Date().getTime() - startTime;
  log("INFO", "extractGammeFromTimeline", "✅ Extraction terminée en " + duration + "ms", {
    title: title,
    fieldsCount: fields.length,
    imagesCount: images.length
  });

  return {
    title: title,
    fields: fields,
    images: images
  };
}




function parseOptions(optionsRaw) {
  if (!optionsRaw) return [];

  if (Array.isArray(optionsRaw)) {
    var out = [];
    for (var i = 0; i < optionsRaw.length; i++) {
      var v = String(optionsRaw[i] || '').trim();
      if (v) out.push(v);
    }
    return out;
  }

  var s = String(optionsRaw || '').trim();
  if (!s) return [];

  s = s.replace(/\r?\n/g, '|')
       .replace(/;/g, '|')
       .replace(/,/g, '|');

  var parts = s.split('|');
  var res = [];
  for (var j = 0; j < parts.length; j++) {
    var p = String(parts[j] || '').trim();
    if (p) res.push(p);
  }
  return res;
}

// ═══════════════════════════════════════════════════════════════
// SCHÉMA FORMULAIRE
// ═══════════════════════════════════════════════════════════════

function getFormSchema() {
  var startTime = new Date().getTime();
  log("INFO", "getFormSchema", "Récupération schéma formulaire");
  
  try {
    var cache = CacheService.getScriptCache();
    var cacheKey = 'form_schema';
    var cached = cache.get(cacheKey);
    
    if (cached) {
      try {
        var result = JSON.parse(cached);
        log("INFO", "getFormSchema", "Schéma depuis cache");
        return result;
      } catch (parseError) {
        log("WARN", "getFormSchema", "Cache corrompu, suppression", {
          error: parseError.message || parseError
        });
        cache.remove(cacheKey);
      }
    }
    
    // Ouvrir le spreadsheet spécifié dans la configuration
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.SHEET_FORM_SCHEMA);
    
    if (!sheet) {
      log("WARN", "getFormSchema", "Feuille " + (CONFIG.SHEET_FORM_SCHEMA || "non défini") + " introuvable");
      return { fields: [] };
    }
    
    var data = sheet.getDataRange().getValues();
    if (!data || data.length < 2) {
      log("WARN", "getFormSchema", "Feuille vide ou sans en-têtes", {
        rowCount: data.length
      });
      return { fields: [] };
    }
    
    var headers = data[0].map(function(h) { return String(h || '').trim().toLowerCase(); });
    var colIndex = function(name) {
      var idx = headers.indexOf(name.toLowerCase());
      return idx >= 0 ? idx : -1;
    };
    
    var fields = [];
    
    for (var row = 1; row < data.length; row++) {
      var rowData = data[row];
      
      var activeIdx = colIndex('active');
      var activeVal = activeIdx >= 0 ? String(rowData[activeIdx] || '').toUpperCase() : 'TRUE';
      if (activeVal !== 'TRUE') continue;
      
      var keyIdx = colIndex('key');
      var key = keyIdx >= 0 ? String(rowData[keyIdx] || '').trim() : '';
      if (!key) continue;
      
      var field = {
        name: key,
        label: getColValue(rowData, colIndex('label'), key),
        type: getColValue(rowData, colIndex('type'), 'text'),
        required: getColValue(rowData, colIndex('required'), '').toUpperCase() === 'TRUE',
        options: parseOptions(getColValue(rowData, colIndex('options'), '')),
        phase: getColValue(rowData, colIndex('phase'), 'pre').toLowerCase(),
        display_if: getColValue(rowData, colIndex('display_if'), ''),
        placeholder: getColValue(rowData, colIndex('placeholder'), ''),
        hint: getColValue(rowData, colIndex('hint'), ''),
        multiple: getColValue(rowData, colIndex('multiple'), '').toUpperCase() === 'TRUE',
        accept: getColValue(rowData, colIndex('accept'), ''),
        bloc: getColValue(rowData, colIndex('bloc'), 'AUTRES'),
        ordre: parseInt(getColValue(rowData, colIndex('ordre'), '0'), 10) || 0
      };
      
      fields.push(field);
    }
    
    // Trier les champs par phase et ordre
    fields.sort(function(a, b) {
      var phaseOrder = { 'pre': 1, 'metier': 2, 'post': 3 };
      var pa = phaseOrder[a.phase] || 99;
      var pb = phaseOrder[b.phase] || 99;
      if (pa !== pb) return pa - pb;
      return (a.ordre || 0) - (b.ordre || 0);
    });
    
    var schema = { fields: fields };
    
    try {
      cache.put(cacheKey, JSON.stringify(schema), CONFIG.CACHE_TTL || 3600); // Ajout d'une valeur par défaut pour CACHE_TTL
      log("INFO", "getFormSchema", "Schéma mis en cache");
    } catch (cacheError) {
      log("WARN", "getFormSchema", "Échec mise en cache: " + (cacheError.message || cacheError));
    }
    
    var duration = new Date().getTime() - startTime;
    log("INFO", "getFormSchema", "Schéma chargé", {
      duration: duration + "ms",
      fieldsCount: fields.length
    });
    
    return schema;
    
  } catch (error) {
    logError('getFormSchema', error);
    log("ERROR", "getFormSchema", "Erreur récupération schéma: " + (error.message || error));
    return { fields: [] };
  }
}

function getColValue(row, index, defaultValue) {
  if (index < 0 || index >= row.length) return defaultValue;
  var val = row[index];
  return val !== null && val !== undefined ? String(val).trim() : defaultValue;
}


// ═══════════════════════════════════════════════════════════════
// GESTION MACHINES ET SITES
// ═══════════════════════════════════════════════════════════════

function getMachinesAndSites() {
  var startTime = new Date().getTime();
  log("INFO", "getMachinesAndSites", "Chargement machines et sites");
  
  try {
    var cache = CacheService.getScriptCache();
    var cacheKey = 'machines_and_sites';
    var cached = cache.get(cacheKey);
    
    if (cached) {
      var result = JSON.parse(cached);
      log("INFO", "getMachinesAndSites", "Données depuis cache");
      return result;
    }
    
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sh = ss.getSheetByName("Machines");
    
    if (!sh) {
      log("ERROR", "getMachinesAndSites", "Onglet Machines introuvable");
      return { machines: [], sites: [], sitesByMachine: {} };
    }

    var values = sh.getDataRange().getValues();
    if (values.length < 2) {
      return { machines: [], sites: [], sitesByMachine: {} };
    }

    var header = values[0].map(function(h) { return String(h).trim(); });
    var colID = header.indexOf("ID");
    var colNom = header.indexOf("Nom");
    var colType = header.indexOf("Type");
    var colZone = header.indexOf("Zone");

    if (colNom === -1 || colZone === -1) {
      log("ERROR", "getMachinesAndSites", "Colonnes manquantes");
      return { machines: [], sites: [], sitesByMachine: {} };
    }

    var machinesMap = {};
    var sitesSet = {};

    for (var i = 1; i < values.length; i++) {
      var id = colID >= 0 ? String(values[i][colID] || "").trim() : "";
      var nom = String(values[i][colNom] || "").trim();
      var type = colType >= 0 ? String(values[i][colType] || "").trim() : "";
      var zone = String(values[i][colZone] || "").trim();

      var machineKey = id || nom;
      if (!machineKey) continue;

      if (!machinesMap[machineKey]) {
        machinesMap[machineKey] = {
          id: id || nom,
          nom: nom || id,
          type: type,
          zones: {}
        };
      }

      if (zone) {
        sitesSet[zone] = true;
        machinesMap[machineKey].zones[zone] = true;
      }
    }

    var machines = [];
    for (var key in machinesMap) {
      if (machinesMap.hasOwnProperty(key)) {
        var m = machinesMap[key];
        var zonesArray = [];
        for (var z in m.zones) {
          if (m.zones.hasOwnProperty(z)) zonesArray.push(z);
        }
        machines.push({
          id: m.id,
          label: m.nom ? m.id + " - " + m.nom : m.id,
          type: m.type,
          zones: zonesArray.sort()
        });
      }
    }

    machines.sort(function(a, b) {
      return a.id.localeCompare(b.id);
    });

    var sites = [];
    for (var s in sitesSet) {
      if (sitesSet.hasOwnProperty(s)) sites.push(s);
    }
    sites.sort();

    var sitesByMachine = {};
    for (var j = 0; j < machines.length; j++) {
      sitesByMachine[machines[j].id] = machines[j].zones;
    }

    var result = {
      machines: machines,
      sites: sites,
      sitesByMachine: sitesByMachine
    };

    try {
      cache.put(cacheKey, JSON.stringify(result), CONFIG.CACHE_TTL);
    } catch (cacheError) {
      log("WARN", "getMachinesAndSites", "Échec cache");
    }

    var duration = new Date().getTime() - startTime;
    log("INFO", "getMachinesAndSites", "Données chargées", {
      duration: duration + "ms",
      machinesCount: machines.length,
      sitesCount: sites.length
    });

    return result;

  } catch (error) {
    logError('getMachinesAndSites', error);
    return { machines: [], sites: [], sitesByMachine: {} };
  }
}

function getMachinesList() {
  try {
    var data = getMachinesAndSites();
    return data.machines || [];
  } catch (error) {
    logError('getMachinesList', error);
    return [];
  }
}

function getSitesList() {
  try {
    var data = getMachinesAndSites();
    return data.sites || [];
  } catch (error) {
    logError('getSitesList', error);
    return [];
  }
}

function getSitesForMachine(machineId) {
  try {
    if (!machineId) return [];
    var data = getMachinesAndSites();
    return data.sitesByMachine[machineId] || [];
  } catch (error) {
    logError('getSitesForMachine', error);
    return [];
  }
}

function getMachinesForSite(siteName) {
  try {
    siteName = String(siteName || "").trim();
    if (!siteName) return [];

    var data = getMachinesAndSites();
    var machines = data.machines || [];

    var result = [];
    for (var i = 0; i < machines.length; i++) {
      var m = machines[i];
      if ((m.zones || []).indexOf(siteName) !== -1) {
        result.push(m);
      }
    }

    return result;

  } catch (error) {
    logError('getMachinesForSite', error);
    return [];
  }
}

// ═══════════════════════════════════════════════════════════════
// GESTION PHOTOS
// ═══════════════════════════════════════════════════════════════

function uploadPhoto(interventionId, filename, base64, mimeType) {
  var startTime = new Date().getTime();
  log("INFO", "uploadPhoto", "Upload photo pour " + interventionId);
  
  try {
    if (!interventionId || !base64) {
      throw new Error("Paramètres manquants");
    }
    
    var dataSize = base64.length * 0.75;
    var maxSize = CONFIG.MAX_FILE_SIZE_MB * 1024 * 1024;
    if (dataSize > maxSize) {
      throw new Error("Fichier trop volumineux");
    }
    
    var blob = createBlobFromBase64(base64, mimeType, filename, interventionId);
    
    var folder = retryOperation(function() {
      return DriveApp.getFolderById(CONFIG.OUTPUT_FOLDER_ID);
    }, 'get_output_folder');
    
    var file = retryOperation(function() {
      return folder.createFile(blob);
    }, 'create_file');
    
    retryOperation(function() {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    }, 'set_file_sharing');
    
    var photosSheet = retryOperation(function() {
      return getOrCreateSheet(CONFIG.SHEET_PHOTOS, SHEET_HEADERS.PHOTOS);
    }, 'get_photos_sheet');
    
    retryOperation(function() {
      photosSheet.appendRow([interventionId, new Date(), file.getName(), file.getUrl(), 'intervention']);
    }, 'log_photo_upload');
    
    var duration = new Date().getTime() - startTime;
    log("INFO", "uploadPhoto", "Photo uploadée", {
      duration: duration + "ms",
      fileName: file.getName()
    });
    
    return { success: true, url: file.getUrl(), name: file.getName() };
    
  } catch (error) {
    logError('uploadPhoto', error);
    throw new Error("Upload photo échoué : " + error.message);
  }
}

function uploadPiecePhoto(interventionId, pieceId, fileName, base64Data, mimeType) {
  var startTime = new Date().getTime();
  log("INFO", "uploadPiecePhoto", "Upload photo pièce " + pieceId);
  
  try {
    if (!interventionId || !base64Data) {
      throw new Error("Paramètres manquants");
    }
    
    var blob = createBlobFromBase64(base64Data, mimeType, fileName, interventionId + "_" + pieceId);
    
    var rootFolder = retryOperation(function() {
      return DriveApp.getFolderById(CONFIG.PIECES_PHOTOS_FOLDER_ID);
    }, 'get_pieces_folder');
    
    var interventionFolder = getOrCreateSubfolder(rootFolder, interventionId);
    var pieceFolder = getOrCreateSubfolder(interventionFolder, "piece_" + pieceId);
    
    var file = retryOperation(function() {
      return pieceFolder.createFile(blob);
    }, 'create_piece_file');
    
    retryOperation(function() {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    }, 'set_piece_file_sharing');
    
    updatePiecePhotoUrl(interventionId, pieceId, file.getUrl());
    
    var photosSheet = retryOperation(function() {
      return getOrCreateSheet(CONFIG.SHEET_PHOTOS, SHEET_HEADERS.PHOTOS);
    }, 'get_photos_sheet_piece');
    
    retryOperation(function() {
      photosSheet.appendRow([interventionId, new Date(), file.getName(), file.getUrl(), "piece_" + pieceId]);
    }, 'log_piece_photo');
    
    var duration = new Date().getTime() - startTime;
    log("INFO", "uploadPiecePhoto", "Photo pièce uploadée", {
      duration: duration + "ms"
    });
    
    return { success: true, url: file.getUrl(), fileName: file.getName() };
    
  } catch (error) {
    logError('uploadPiecePhoto', error);
    return { success: false, error: error.message };
  }
}

function createBlobFromBase64(base64, mimeType, filename, prefix) {
  var comma = base64.indexOf(',');
  var rawBase64 = comma >= 0 ? base64.slice(comma + 1) : base64;
  var bytes = Utilities.base64Decode(rawBase64);
  var safeName = prefix + "_" + (filename || 'photo_' + Date.now());
  return Utilities.newBlob(bytes, mimeType || 'application/octet-stream', safeName);
}

function getOrCreateSubfolder(parentFolder, folderName) {
  try {
    var folders = parentFolder.getFoldersByName(folderName);
    if (folders.hasNext()) {
      return folders.next();
    } else {
      return parentFolder.createFolder(folderName);
    }
  } catch (error) {
    logError('getOrCreateSubfolder', error);
    throw error;
  }
}

function updatePiecePhotoUrl(interventionId, pieceId, photoUrl) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.SHEET_PIECES);
    if (!sheet) return;
    
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === interventionId && data[i][3] === pieceId) {
        var currentUrls = data[i][19] || '';
        var newUrls = currentUrls ? currentUrls + "\n" + photoUrl : photoUrl;
        sheet.getRange(i + 1, 20).setValue(newUrls);
        break;
      }
    }
  } catch (error) {
    logError('updatePiecePhotoUrl', error);
  }
}

// ═══════════════════════════════════════════════════════════════
// GESTION PIÈCES - DÉCISIONS
// ═══════════════════════════════════════════════════════════════

function processDecisionsPieces(pieces) {
  if (!Array.isArray(pieces)) return;
  
  try {
    for (var i = 0; i < pieces.length; i++) {
      var piece = pieces[i];
      var decision = normalizeDecision(piece.decision);
      
      switch (decision) {
        case 'recommander_reassort':
          addToReassortList(piece);
          break;
        case 'commander_piece':
        case 'commander_pour_stock':
        case 'commander_pour_intervention':
          addToCommandesList(piece);
          break;
        case 'chiffrer_devis':
          addToDevisList(piece);
          break;
      }
    }
  } catch (error) {
    logError('processDecisionsPieces', error);
  }
}

function normalizeDecision(decision) {
  if (!decision) return '';
  return String(decision).toLowerCase().replace(/\s+/g, '_').replace(/[^a-z0-9_]/g, '');
}

function addToReassortList(piece) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Réassort');
    
    if (!sheet) {
      sheet = ss.insertSheet('Réassort');
      sheet.appendRow(['Date', 'Machine', 'Référence', 'Désignation', 'Quantité', 'Criticité', 'Statut']);
      sheet.getRange(1, 1, 1, 7).setFontWeight('bold');
    }
    
    sheet.appendRow([
      new Date(), 
      piece.machine_id || '', 
      piece.reference || '', 
      piece.designation || '', 
      piece.quantite || 1, 
      piece.criticite || '', 
      'À commander'
    ]);
    
  } catch (error) {
    logError('addToReassortList', error);
  }
}

function addToCommandesList(piece) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Commandes');
    
    if (!sheet) {
      sheet = ss.insertSheet('Commandes');
      sheet.appendRow(['Date', 'Machine', 'Référence', 'Désignation', 'Quantité', 'Fournisseur', 'Prix HT', 'Type', 'Urgence', 'Statut']);
      sheet.getRange(1, 1, 1, 10).setFontWeight('bold');
    }
    
    var urgence = piece.criticite === 'Critique' || piece.criticite === 'Haute' ? 'Urgent' : 'Normal';
    sheet.appendRow([
      new Date(), 
      piece.machine_id || '', 
      piece.reference || '', 
      piece.designation || '', 
      piece.quantite || 1, 
      piece.fournisseur || '', 
      piece.prix_unitaire_ht || '', 
      piece.decision || '', 
      urgence, 
      'À commander'
    ]);
    
  } catch (error) {
    logError('addToCommandesList', error);
  }
}

function addToDevisList(piece) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Devis');
    
    if (!sheet) {
      sheet = ss.insertSheet('Devis');
      sheet.appendRow(['Date', 'Machine', 'Zone', 'Référence', 'Désignation', 'Quantité', 'Commentaire', 'Demandeur', 'Statut']);
      sheet.getRange(1, 1, 1, 9).setFontWeight('bold');
    }
    
    sheet.appendRow([
      new Date(), 
      piece.machine_id || '', 
      piece.zone_composant || '', 
      piece.reference || '', 
      piece.designation || '', 
      piece.quantite || 1, 
      piece.commentaire || '', 
      Session.getActiveUser().getEmail(), 
      'À chiffrer'
    ]);
    
  } catch (error) {
    logError('addToDevisList', error);
  }
}

// ═══════════════════════════════════════════════════════════════
// INTERVENTIONS CORRECTIVES
// ═══════════════════════════════════════════════════════════════

function createCorrectiveInterventions(interventionIdSource, pieces, machineId) {
  var startTime = new Date().getTime();
  log("INFO", "createCorrectiveInterventions", "Création correctif pour " + interventionIdSource);
  
  try {
    var piecesCorrectifs = pieces.filter(function(p) {
      return normalizeDecision(p.decision) === 'commander_pour_intervention' ||
             (p.action_immediate === 'Non' && (p.criticite === 'Haute' || p.criticite === 'Critique'));
    });
    
    if (piecesCorrectifs.length === 0) {
      log("INFO", "createCorrectiveInterventions", "Aucune pièce nécessitant correctif");
      return;
    }
    
    var sheet = retryOperation(function() {
      return getOrCreateSheet(CONFIG.SHEET_INTERVENTIONS, SHEET_HEADERS.INTERVENTIONS);
    }, 'get_interventions_sheet_correctif');
    
    var correctifId = generateInterventionId('CORR');
    var now = new Date();
    
    var piecesDesc = [];
    for (var i = 0; i < piecesCorrectifs.length; i++) {
      piecesDesc.push(piecesCorrectifs[i].reference || piecesCorrectifs[i].designation);
    }
    
    var description = "Intervention corrective suite à " + interventionIdSource + 
                     "\nPièces: " + piecesDesc.join(', ');
    
    retryOperation(function() {
      sheet.appendRow([
        correctifId, now, '', '', machineId, Session.getActiveUser().getEmail(),
        'correctif', '', JSON.stringify({ origine: interventionIdSource }), '', '', 'NON', 'À planifier', description
      ]);
    }, 'insert_correctif_intervention');
    
    var duration = new Date().getTime() - startTime;
    log("INFO", "createCorrectiveInterventions", "Correctif " + correctifId + " créé", {
      duration: duration + "ms",
      piecesCount: piecesCorrectifs.length
    });
    
    sendCorrectifNotification(correctifId, machineId, piecesCorrectifs);
    
  } catch (error) {
    logError('createCorrectiveInterventions', error);
  }
}

function sendCorrectifNotification(correctifId, machineId, pieces) {
  try {
    var email = Session.getActiveUser().getEmail();
    var subject = "⚠️ [ARMITEC] Intervention corrective " + correctifId + " - Machine " + machineId;
    
    var body = "Une intervention corrective a été créée automatiquement.\n\n";
    body += "ID: " + correctifId + "\n";
    body += "Machine: " + machineId + "\n";
    body += "Date: " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm") + "\n\n";
    body += "Pièces concernées:\n";
    
    for (var i = 0; i < pieces.length; i++) {
      var p = pieces[i];
      body += (i + 1) + ". " + (p.reference || 'N/A') + " - " + (p.designation || 'N/A') + " (Criticité: " + (p.criticite || 'N/A') + ")\n";
    }
    
    MailApp.sendEmail(email, subject, body);
    log("INFO", "sendCorrectifNotification", "Notification envoyée");
  } catch (error) {
    logError('sendCorrectifNotification', error);
  }
}

// ═══════════════════════════════════════════════════════════════
// MENU GOOGLE SHEETS
// ═══════════════════════════════════════════════════════════════

function onOpen(e) {
  log("INFO", "onOpen", "Initialisation menu ARMITEC");
  
  try {
    var ui = SpreadsheetApp.getUi();
    
    ui.createMenu("🔧 ARMITEC")
      .addSubMenu(ui.createMenu("📋 Gammes")
        .addItem("Visualiser une gamme", "visualizeGamme")
        .addItem("Tester extraction gammes", "testExtractGammes")
        .addItem("🧪 TEST: Chargement gamme complète", "testGammeLoading"))  // ✅ AJOUTER CETTE LIGNE

    log("INFO", "onOpen", "Menu ajouté avec succès");
  } catch (error) {
    logError("onOpen", error);
  }
}

// ═══════════════════════════════════════════════════════════════
// ACTIONS MENU - INITIALISATION
// ═══════════════════════════════════════════════════════════════

function initializeAllSheets() {
  try {
    getOrCreateSheet(CONFIG.SHEET_INTERVENTIONS, SHEET_HEADERS.INTERVENTIONS);
    getOrCreateSheet(CONFIG.SHEET_PIECES, SHEET_HEADERS.PIECES);
    getOrCreateSheet(CONFIG.SHEET_PHOTOS, SHEET_HEADERS.PHOTOS);
    getOrCreateSheet(CONFIG.SHEET_LOGS, SHEET_HEADERS.LOGS);
    createMachinesSheetIfNeeded();
    createGammesSheetIfNeeded();
    createFormSchemaSheet();
    SpreadsheetApp.getUi().alert("✅ Toutes les feuilles ont été initialisées.");
  } catch (error) {
    logError("initializeAllSheets", error);
    SpreadsheetApp.getUi().alert("❌ Erreur: " + error.message);
  }
}

function createMachinesSheetIfNeeded() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.SHEET_MACHINES);
  if (sheet) return;
  
  sheet = ss.insertSheet(CONFIG.SHEET_MACHINES);
  var data = [
    ["ID", "Nom", "Type", "Marque", "Modèle", "Année", "Zone"],
    ["M001", "Tour CNC 1", "Tour", "Mazak", "QT-15N", 2018, "Atelier A"],
    ["M002", "Fraiseuse DMU", "Fraiseuse", "DMG MORI", "DMU 50", 2020, "Atelier A"],
    ["M006", "Centre usinage", "Centre", "HAAS", "VF-2", 2019, "Atelier B"],
    ["M064", "Machine exemple", "Exemple", "Generic", "EX-100", 2020, "Atelier A"],
    ["M075", "Imprimante 3D", "3D", "Ultimaker", "S5 Pro", 2021, "Lidec"]
  ];
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  sheet.getRange(1, 1, 1, data[0].length).setFontWeight("bold");
}

function createGammesSheetIfNeeded() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.SHEET_GAMMES);
  if (sheet) return;
  
  sheet = ss.insertSheet(CONFIG.SHEET_GAMMES);
  var data = [
    ["Machine", "Code Gamme", "Libellé", "Périodicité", "Durée estimée", "Statut"],
    ["M075", "sec_7", "Maintenance 06 mois", "6 mois", "3h", "Actif"]
  ];
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  sheet.getRange(1, 1, 1, data[0].length).setFontWeight("bold");
}

function createFormSchemaSheet() {
  try {
    var headers = ["active", "bloc", "ordre", "phase", "key", "label", "type", "required", "options", "display_if", "placeholder", "hint", "multiple", "accept"];
    getOrCreateSheet(CONFIG.SHEET_FORM_SCHEMA, headers);
    SpreadsheetApp.getUi().alert("✅ Feuille FORM_SCHEMA créée.");
  } catch (error) {
    logError("createFormSchemaSheet", error);
    SpreadsheetApp.getUi().alert("❌ Erreur: " + error.message);
  }
}

function populateFormSchemaExamples() {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.SHEET_FORM_SCHEMA);
    if (!sheet) throw new Error("Feuille FORM_SCHEMA introuvable");

    var examples = [
      ["TRUE", "IDENTIFICATION", 1, "pre", "date_intervention", "Date intervention", "date", "TRUE", "", "", "", "", "FALSE", ""],
      ["TRUE", "IDENTIFICATION", 2, "pre", "site", "Site", "select", "TRUE", "Atelier Bélliparc|Commun de site|Lidec", "", "Choisissez le site", "", "FALSE", ""],
      ["TRUE", "IDENTIFICATION", 3, "pre", "machine_id", "Machine", "select", "TRUE", "", "", "Choisissez la machine", "", "FALSE", ""],
      ["TRUE", "IDENTIFICATION", 4, "pre", "technicien", "Technicien", "text", "TRUE", "", "", "Nom du technicien", "", "FALSE", ""],
      ["TRUE", "CLASSIFICATION", 5, "pre", "type_maintenance", "Type maintenance", "select", "TRUE", "preventif|correctif|diagnostique|reglementaire|modification", "", "", "", "FALSE", ""],
      ["TRUE", "TEMPS_RESSOURCES", 40, "post", "temps_passe", "Temps passé (HH:MM)", "time", "TRUE", "", "", "Format HH:MM", "", "FALSE", ""]
    ];

    var startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, examples.length, examples[0].length).setValues(examples);
    SpreadsheetApp.getUi().alert("✅ " + examples.length + " exemples ajoutés.");
  } catch (error) {
    logError("populateFormSchemaExamples", error);
    SpreadsheetApp.getUi().alert("❌ Erreur: " + error.message);
  }
}

function applyFormSchemaValidations() {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.SHEET_FORM_SCHEMA);
    if (!sheet) throw new Error("Feuille FORM_SCHEMA introuvable");

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var colIndex = function(name) {
      for (var i = 0; i < headers.length; i++) {
        if (String(headers[i]).toLowerCase() === name.toLowerCase()) return i + 1;
      }
      return -1;
    };

    var startRow = 2;
    var lastRow = Math.max(sheet.getLastRow(), startRow);
    var numRows = lastRow - startRow + 1;

    var validations = {
      "active": ["TRUE", "FALSE"],
      "phase": ["pre", "metier", "post"],
      "type": ["text", "textarea", "select", "date", "time", "number", "file", "radio", "checkbox"],
      "required": ["TRUE", "FALSE"],
      "multiple": ["TRUE", "FALSE"]
    };

    for (var colName in validations) {
      if (validations.hasOwnProperty(colName)) {
        var col = colIndex(colName);
        if (col > 0 && numRows > 0) {
          var range = sheet.getRange(startRow, col, numRows, 1);
          var rule = SpreadsheetApp.newDataValidation()
            .requireValueInList(validations[colName], true)
            .setAllowInvalid(false)
            .build();
          range.setDataValidation(rule);
        }
      }
    }

    SpreadsheetApp.getUi().alert("✅ Validations appliquées.");
  } catch (error) {
    logError("applyFormSchemaValidations", error);
    SpreadsheetApp.getUi().alert("❌ Erreur: " + error.message);
  }
}

function clearAllCaches() {
  try {
    var cache = CacheService.getScriptCache();
    var keys = ["form_schema", "machines_and_sites", "gammes_M001", "gammes_M002", "gammes_M006", "gammes_M064", "gammes_M075"];
    cache.removeAll(keys);
    log("INFO", "clearAllCaches", "Caches nettoyés");
    SpreadsheetApp.getUi().alert("✅ Cache nettoyé.");
  } catch (error) {
    logError("clearAllCaches", error);
    SpreadsheetApp.getUi().alert("❌ Erreur: " + error.message);
  }
}

function openWebApp() {
  try {
    var url = ScriptApp.getService().getUrl();
    var html = HtmlService.createHtmlOutput('<script>window.open("' + url + '","_blank");google.script.host.close();</script>');
    SpreadsheetApp.getUi().showModalDialog(html, "Ouverture du formulaire...");
  } catch (error) {
    logError("openWebApp", error);
    SpreadsheetApp.getUi().alert("❌ Erreur: " + error.message);
  }
}

function showLastLogs() {
  try {
    var sheet = getOrCreateSheet(CONFIG.SHEET_LOGS, SHEET_HEADERS.LOGS);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      SpreadsheetApp.getUi().alert("Aucun log disponible.");
      return;
    }
    var startRow = Math.max(2, lastRow - 49);
    var data = sheet.getRange(startRow, 1, lastRow - startRow + 1, 4).getValues().reverse();
    var text = data.map(function(r) { return r.join(" | "); }).join("\n");
    SpreadsheetApp.getUi().alert(text || "Aucun log");
  } catch (error) {
    logError("showLastLogs", error);
    SpreadsheetApp.getUi().alert("❌ Erreur: " + error.message);
  }
}

function showInterventionStats() {
  try {
    var sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_INTERVENTIONS);
    if (!sheet) {
      SpreadsheetApp.getUi().alert("Aucune intervention enregistrée.");
      return;
    }
    var data = sheet.getDataRange().getValues();
    var total = Math.max(0, data.length - 1);
    var preventif = 0;
    var correctif = 0;
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][6] === "preventif") preventif++;
      if (data[i][6] === "correctif") correctif++;
    }
    
    var msg = "📊 Statistiques Interventions\n\n";
    msg += "Total: " + total + "\n";
    msg += "Préventif: " + preventif + "\n";
    msg += "Correctif: " + correctif + "\n";
    
    SpreadsheetApp.getUi().alert(msg);
  } catch (error) {
    logError("showInterventionStats", error);
    SpreadsheetApp.getUi().alert("❌ Erreur: " + error.message);
  }
}

function testMachinesAndSites() {
  try {
    var data = getMachinesAndSites();
    var msg = "🔧 Test Machines et Sites\n\n";
    msg += "Machines: " + (data.machines || []).length + "\n";
    msg += "Sites: " + (data.sites || []).length + "\n\n";
    msg += "Sites: " + (data.sites || []).join(", ") + "\n\n";
    msg += "Exemples machines:\n";
    
    var machines = data.machines || [];
    for (var i = 0; i < Math.min(5, machines.length); i++) {
      var m = machines[i];
      msg += "• " + m.id + " (" + (m.type || "N/A") + ") - Zones: " + (m.zones || []).join(", ") + "\n";
    }
    
    SpreadsheetApp.getUi().alert(msg);
  } catch (error) {
    logError("testMachinesAndSites", error);
    SpreadsheetApp.getUi().alert("❌ Erreur: " + error.message);
  }
}

function testExtractGammes() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt("🧪 Test extraction gammes", "Entrez l'ID de la machine (ex: M075):", ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  var machineId = response.getResponseText().trim();
  if (!machineId) {
    ui.alert("ID machine requis");
    return;
  }
  
  try {
    var gammes = getGammesForMachine(machineId);
    var msg = "📋 Gammes extraites pour " + machineId + ":\n\n";
    
    if (!gammes || gammes.length === 0) {
      msg += "Aucune gamme trouvée.";
    } else {
      for (var i = 0; i < gammes.length; i++) {
        var g = gammes[i];
        msg += (i + 1) + ". " + g.label + " (" + g.stepsCount + " étapes)\n";
        msg += "   Key: " + g.key + "\n";
        msg += "   Order: " + g.order + "\n\n";
      }
    }
    
    ui.alert(msg);
  } catch (error) {
    logError("testExtractGammes", error);
    ui.alert("❌ Erreur: " + error.message);
  }
}

function visualizeGamme() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt("📋 Visualiser une gamme", "Entrez l'ID de la machine (ex: M075):", ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  var machineId = response.getResponseText().trim();
  if (!machineId) {
    ui.alert("ID machine requis");
    return;
  }
  
  try {
    var html = renderGammeHtml(machineId);
    var output = HtmlService.createHtmlOutput(html).setWidth(1000).setHeight(700);
    ui.showModalDialog(output, "Gammes - Machine " + machineId);
  } catch (error) {
    logError("visualizeGamme", error);
    ui.alert("❌ Erreur: " + error.message);
  }
}

function renderGammeHtml(machineId) {
  try {
    if (!machineId) {
      return "<h1>Erreur</h1><p>ID machine manquant</p>";
    }
    
    var folder = DriveApp.getFolderById(CONFIG.MODELS_FOLDER_ID);
    var modelFile = getModelFileForMachine(folder, machineId);
    
    if (!modelFile) {
      return "<h1>Erreur</h1><p>Modèle introuvable pour " + machineId + "</p>";
    }
    
    var model = JSON.parse(modelFile.getBlob().getDataAsString("UTF-8"));
    var gammes = extractGammesFromTimeline(model.timeline);
    
    var html = "<h1>Gammes - " + machineId + "</h1>";
    html += "<p>Machine: " + (model.meta && model.meta.materiel ? model.meta.materiel : machineId) + "</p>";
    html += "<p>Site: " + (model.meta && model.meta.site ? model.meta.site : 'N/A') + "</p>";
    html += "<hr>";
    
    for (var i = 0; i < gammes.length; i++) {
      var gamme = gammes[i];
      html += "<h2>" + gamme.label + " (" + gamme.stepsCount + " étapes)</h2>";
    }
    
    return html;
    
  } catch (error) {
    logError('renderGammeHtml', error);
    return "<h1>Erreur</h1><p>" + error.message + "</p>";
  }
}

function listAllModelFiles() {
  try {
    var folder = DriveApp.getFolderById(CONFIG.MODELS_FOLDER_ID);
    var files = folder.getFiles();
    var fileList = [];
    
    while (files.hasNext()) {
      var file = files.next();
      var name = file.getName();
      if (name.indexOf('.json') !== -1) {
        fileList.push({
          name: name,
          size: (file.getSize() / 1024).toFixed(2) + ' KB',
          lastModified: Utilities.formatDate(file.getLastUpdated(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
        });
      }
    }
    
    var msg = "📁 Fichiers JSON dans MODELS_FOLDER:\n\n";
    if (fileList.length === 0) {
      msg += "Aucun fichier JSON trouvé.";
    } else {
      for (var i = 0; i < fileList.length; i++) {
        var f = fileList[i];
        msg += (i + 1) + ". " + f.name + " (" + f.size + ")\n";
        msg += "   Modifié: " + f.lastModified + "\n\n";
      }
    }
    
    SpreadsheetApp.getUi().alert(msg);
  } catch (error) {
    logError("listAllModelFiles", error);
    SpreadsheetApp.getUi().alert("❌ Erreur: " + error.message);
  }
}

function showAbout() {
  var msg = "🔧 ARMITEC v5.3.3\n\n";
  msg += "Système de Gestion de Maintenance Industrielle\n\n";
  msg += "Fonctionnalités:\n";
  msg += "• Gestion gammes JSON timeline (order>=25)\n";
  msg += "• Support M075.json et M075_model.json\n";
  msg += "• Cache + retry + logs détaillés\n";
  msg += "• Gestion pièces + correctifs auto\n";
  msg += "• Upload photos\n\n";
  msg += "v5.3.3: Fix champs disabled + syntaxe complète\n\n";
  msg += "© 2024 ARMITEC";
  
  SpreadsheetApp.getUi().alert(msg);
}
