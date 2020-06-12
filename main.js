/**
 * TODO: ajouter la fonctionnalité de calcul du nombre d'UV par UC
 * TODO: ajouter la fonctionnalité de calcul du prix moyen des UV par UC
 * TODO: ajouter la fonctionnalité de tri par mois plutot que par semaine
 */

// Fonction effectuée au démarrage de la SS
function onOpen() {
    var UI = SpreadsheetApp.getUi();
    UI.createMenu('Test') // on ajoute un nouveau menu
        .addItem('Load data', 'loadData') // une fonction qui permet de charger les données dans le SS
        .addItem('Trier par semaine', 'sortByWeek') // permet de trier par semaine
        .addSeparator() // on ajoute un séparateur purement décoratif
        .addItem('Calculer Prix moyen', 'getAveragePrice') // définition d'une action contextuelle permettant de calculer le prix moyen d'une UV
        .addItem('UV count per UC', 'getUVperUC') // définition d'une action contextuelle permettant d'obtenir le nombre d'UV présent dans chaque UC
        .addItem('Delete Week Sheets', 'delWeekSheets') // définition d'une action contexuelle permettant de supprimer les sheets des détails des semaines (DEBUG)
        .addToUi(); // on ajoute notre nouveau menu et ses actions à l'UI
}

function loadData() {
    // on ouvre la SS active locale
    var destSS = SpreadsheetApp.getActiveSheet();

    // on ouvre la sheet contenant les informations souhaitées
    var srcSS = SpreadsheetApp.openById('1CIGmAglLEjQ8CcHgx2oNJHkcR5b2v5lWfw7aT8-KyTk');

    // on récupère les données
    var srcSheet = srcSS.getSheetByName('Analyse des ventes');
    var srcDataRange = srcSheet.getDataRange();
    var srcDataValues = srcDataRange.getValues();

    // on les colle dans la SS locale
    destSS.getRange(1, 1, srcDataRange.getHeight(), srcDataRange.getWidth()).setValues(srcDataValues);

    // on redimensionne les colones pour faciliter la lecture
    destSS.setName("Analyse Ventes");
    destSS.autoResizeColumns(1, srcDataRange.getWidth());
}

/**
 * permet de supprimer les sheets contenant les détails de chaque semaine
 */
function delWeekSheets () {
    // on récupère toutes les sheets présentes dans le Spreadsheet
    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();

    // on supprime toutes les sheets (sauf la première sheet qui contient les données)
    for (let index = 1; index < sheets.length; index++) {
        SpreadsheetApp.getActiveSpreadsheet().deleteSheet(sheets[index]); // on supprime la sheet
    }
}

function newSheet_ (string) {
    var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
    var newSheet = spreadSheet.getSheetByName(string);

    // si la sheet existe déjà on la supprime
    if (newSheet != null) {
        spreadSheet.deleteSheet(string);
    }

    // sinon on créé une nouvelle sheet avec le nom passé en paramètre
    newSheet = spreadSheet.insertSheet();
    newSheet.setName(string);

    // on redéfini la première sheet en tant que sheet active
    SpreadsheetApp.setActiveSheet(spreadSheet.getSheets()[0]);
    // on retourne le pointeur de la nouvelle sheet
    return newSheet;
}

/**
 * Fonction permettant d'identifier un mois sous la forme "[mois] [année]"
 * @param {String} string la description d'un article 
 */
function isMonthID_ (string) {
    var flag = false;
    var monthRegExp = new RegExp('(\\s*[a-z]+\\s[0-9]{4}\\s*)', 'i'); // permet de reconnaitre les indetifications de mois

    // si on à détecté un mois et que le mot 'ifco' n'est pas détecté alors on est bien sur une cellule identifiant un mois
    if (monthRegExp.exec(string) && string.lastIndexOf('ifco') < 0) {
        flag = true;
    }

    return flag;
}

/**
 * Fonction permettant d'identifier une semaine sous la forme "WXX [année]" où "XX" représente le numéro de la semaine
 * @param {String} string la description d'un article
 */
function isWeekID_(string) {
    var flag = false;
    var weekRegExp = new RegExp('(\\s*(W[0-9]+)\\s([0-9]{4})\\s*)', 'i'); // expression régulière qui permet de définir si un string correspond à la description d'une semaine

    if (weekRegExp.exec(string)) {
        flag = true;
    }

    return flag;
}

function sortByWeek() {

    var dataSheet = SpreadsheetApp.getActiveSheet(); // la sheet active contenant les donénes
    var range = dataSheet.getDataRange(); // la sélection qui contient toutes les données présentes dans la sheet contenant les données
    var values = range.getValues(); // les valeurs de la sélection des données
    var weekRowIndex; // l'indice de la dernière semaine indentifiée

    for (var i = 0; i <= values.length; i++) {
        // si on identifie une semaine ou qu'on est à la fin de la sélection
        if (isWeekID_(values[i]) || i == values.length) {
            // on vérifie qu'une semaine à été détectée précédement (pour éviter de sélectionner des semaines vides ou non correspondantes)
            if (typeof weekRowIndex !== 'undefined') {
                // on nettoie le nom de la semaine
                let weekName = String(values[weekRowIndex]).split(",")[0];

                // on créé une nouvelle sheet nommée avec la semaine actuelle 
                var dstSheet = newSheet_(weekName);

                // on définit la range de source et la range de destination (les trois colonnes sur la semaine actuelle)
                var srcRange = dataSheet.getRange(weekRowIndex + 1, 1, (i - weekRowIndex), range.getWidth());

                // on colle les articles dans la nouvelle sheet
                dstSheet.getRange(1, 1, (i - weekRowIndex), range.getWidth()).setValues(srcRange.getValues());

                // on resize les données pour faciliter la lecture
                dstSheet.autoResizeColumns(1, range.getWidth());
            }
            weekRowIndex = i; // on met à jour le numéro de la ligne de la dernière semaine identifiée
        }
        // si on identifie aucune semaine alors on passe à la ligne suivante (aucune action nécessaire)
    }

    // TODO: ajouter une fonction qui permet de trier les sheets selon l'ordre alphabétique (W53 2019 < W2 2020)
}

/**
 * 
 * @param {String} description une string contenant une description d'article
 */
function getCondFromName_ (description) {
    var cond = 1;

    // conditionnement en "- de XX kg" ou "- XX kg"
    if (description.lastIndexOf(' - ') > 0) {
        var tmpString = String(description.split(' - ')[1]); // on coupe la description en deux afin de récupérer seulement la partie contenant le conditionnement
    
        // dans cette partie de la description on vérifie la présence du séparateur 'de '
        if (tmpString.lastIndexOf('de ') >= 0) {
            cond = parseInt(tmpString.split('de ')[1].split(' ')[0]); // on récupère le conditionnement
        } else if (tmpString.lastIndexOf('par ') >= 0) {
            cond = parseInt(tmpString.split(' ')[1]); // on récupère le conditonnement
        } else if (tmpString.lastIndexOf('kg') > 0) {
            cond = parseInt(tmpString.split(' ')[0]); // on récupère le conditonnement
        }
    }

    return cond;
}

// permet d'obtenir le nombre d'UV par UC
function getUVperUC() {
    // on récupère les données présentes sur la sheet principale
    var dataSheet = SpreadsheetApp.getActiveSheet(); // la sheet active contenant les donénes
    var range = dataSheet.getDataRange(); // la sélection qui contient toutes les données présentes dans la sheet contenant les données
    var values = range.getValues();

    // pour chaque article des données récupérées
    for (var i = 0; i <= values.length; i++) {
        var article = String(values[i]).split(",")[0]; // on définit les données présentes dans chaque ligne

        // on vérifie que la case contient bien une description d'article (pas de mois, semaine, total ou cellule vide)
        if (isWeekID_(article) || isMonthID_(article) || article === '' || article.indexOf('Total') >= 0) {
            continue;
        } 

        // on récupère le conditionnement de chaque article
        var cond = getCondFromName_(article);

        // on définit la nouvelle range et on y place le conditonnement
        dataSheet.getRange(i + 1, range.getWidth() + 1, 1, 1).setValue(cond);
    }
}


function getAveragePrice () {
    // on récupère les données présentes sur la sheet principale
    var dataSheet = SpreadsheetApp.getActiveSheet(); // la sheet active contenant les donénes
    var range = dataSheet.getDataRange(); // la sélection qui contient toutes les données présentes dans la sheet contenant les données
    var values = range.getValues();

    // pour chaque article des données récupérées
    for (var i = 0; i <= values.length; i++) {
        var article = String(values[i]).split(",")[0]; // on définit les données présentes dans chaque ligne
        var nbUV = (String(values[i]).split(",")[1]);
        var TotalHT = (String(values[i]).split(",")[2]);
        var UC = (String(values[i]).split(",")[3]);

        // on vérifie que la case contient bien une description d'article (pas de mois, semaine, total ou cellule vide)
        if (isWeekID_(article) || isMonthID_(article) || article === '' || article.indexOf('Total') >= 0) {
            continue;
        }

        Logger.log(TotalHT + '/' + '(' + nbUV + '*' + UC + ')');
        dataSheet.getRange(i + 1, range.getWidth() + 1, 1, 1).setValue(TotalHT / (nbUV * UC));
    }
}