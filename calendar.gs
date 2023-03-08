/**
 * Ajoute les évènements de l'agenda public ISI 2A à l'agenda personnel.
 * Les évènements de la semaine actuelle sont ajoutés.
 * Permet de filtrer les évènements et d'ajouter des couleurs.
 * 
 * GUIDE D'UTILISATION :
 * 
 * 1) Suivre les étapes de https://developers.google.com/calendar/api/quickstart/apps-script?hl=fr
 *    pour créer un nouveau script.
 * 
 * 2) Remplacer l'exemple de code par ce script.
 * 
 * 3) Remplacer les constantes de la fonction main() par les valeurs voulues.
 * 
 * 4) Sauvegarder puis exécuter le script en cliquant sur les boutons en haut de l'écran.
 */
function main() {

  /* Nom de l'agenda personnel à utiliser */
  const agendaPerso = 'Cours';

  /* Agenda public, ici ISI 2A */
  const calendarId = '10d79c21648406b69db795f20f87f48d5bb330bef3dcbe11ab31d3325e5f9812@group.calendar.google.com'; // ID of the public calendar

  /* Ajouter les mots-clés correspondant au titre des matières à garder */
  const matieres = ['ACSS', 'PIEP', 'ALGOA', 'responsable', 'IHM', 'ACOL', 'Algo avancee', 'Anglais G12', 'CAWEB', 'Management'];

  /* Permet d'exclure les évènements non voulus */
  const excludedKeywords = ['ISI G1', 'ISI G2)', 'ISI G3', 'Management G3', 'TP G2', 'TP G3'];

  /* Permet de renommer les évènements */
  const titres = {
    'ACSS': 'ACSS',
    'PIEP': 'PIEP',
    'ALGOA': 'Algo A',
    'Algo avancee': 'Algo A',
    'responsable': 'Numérique responsable',
    'IHM': 'IHM',
    'ACOL': 'ACOL',
    'Anglais G12': 'Théâtre',
    'CAWEB': 'CAWEB',
    'Management': 'Management'
  };

  /* Permet d'ajouter une couleur, voir https://lukeboyle.com/blog/posts/google-calendar-api-color-id pour la liste des codes couleur */
  const couleurs = {
    'ACSS': '2',
    'PIEP': '9',
    'Algo A': '3',
    'Numérique responsable': '3',
    'IHM': '2',
    'ACOL': '9',
    'Théâtre': '4',
    'CAWEB': '9',
    'Management': '4'
  };

  const events = getEventsFromPublicCalendar(calendarId);
  const filteredEvents = filterEventsByKeywords(events, matieres, excludedKeywords);
  const replacedEvents = replaceEventTitles(filteredEvents, titres);
  const coloredEvents = addColorToEvents(replacedEvents, couleurs);

  // Ajout des événements filtrés et colorés à votre agenda personnel
  const calendar = CalendarApp.getCalendarsByName(agendaPerso)[0];
  for (const event of coloredEvents) {
    addEvent(calendar, event);
  }
}

/**
 * Ajoute un évènement au bon calendrier.
 */
function addEvent(calendar, event) {
  const start = new Date(event.start.dateTime);
  const end = new Date(event.end.dateTime);
  const title = event.summary;
  const colorId = event.colorId;
  const options = { description: event.description, location: event.location };

  console.log(`AJOUT -- Titre : ${event.summary} | Couleur : ${event.colorId} | Date : ${event.start.dateTime}`);
  calendar.createEvent(title, start, end, options).setColor(colorId);
}

/**
 * Récupère les événements de la semaine en cours à partir d'un agenda Google public.
 *
 * @return {Array<Object>} Les événements de la semaine en cours.
 */
function getEventsFromPublicCalendar(calendarId) {
  const now = new Date();
  const startOfWeek = new Date(now.getFullYear(), now.getMonth(), now.getDate() - now.getDay()); // Start of current week
  const endOfWeek = new Date(now.getFullYear(), now.getMonth(), now.getDate() + (6 - now.getDay())); // End of current week
  const optionalArgs = {
    timeMin: startOfWeek.toISOString(),
    timeMax: endOfWeek.toISOString(),
    singleEvents: true,
    orderBy: 'startTime'
  };

  try {
    const events = Calendar.Events.list(calendarId, optionalArgs).items;
    return events;
  } catch (err) {
    console.log('Failed with error %s', err.message);
    return [];
  }
}


/**
 * Filtrer les événements en fonction des mots-clés du titre
 *
 * @param {Array<Object>} events - Tableau d'objets représentant les événements.
 * @param {Array<String>} keywords - Liste de mots-clés à utiliser pour filtrer les événements.
 * @param {Array<String>} excludedKeywords - Liste de mots-clés à exclure.
 * @return {Array<Object>} Les événements dont le titre contient au moins un des mots-clés.
 */
function filterEventsByKeywords(events, keywords, excludedKeywords) {
  const filteredEvents = [];

  for (const event of events) {
    let shouldExclude = false;
    for (const excludedKeyword of excludedKeywords) {
      if (event.summary.toLowerCase().indexOf(excludedKeyword.toLowerCase()) !== -1) {
        shouldExclude = true;
        break;
      }
    }

    if (shouldExclude) {
      continue;
    }

    for (const keyword of keywords) {
      if (event.summary.toLowerCase().indexOf(keyword.toLowerCase()) !== -1) {
        filteredEvents.push(event);
        break;
      }
    }
  }

  return filteredEvents;
}

/**
 * Ajoute la couleur des événements en fonction du titre
 *
 * @param {Object} colorMap - HashMap avec un mot-clé et une couleur associée.
 * @param {Array<Object>} events - Tableau d'objets représentant les événements.
 * @return {Array<Object>} Les événements mis à jour avec la couleur associée au titre.
 */
function addColorToEvents(events, colorMap) {
  for (const event of events) {
    event.colorId = colorMap[event.summary];
  }
  return events;
}

/**
 * Remplace le titre des événements en fonction des mots-clés du dictionnaire
 *
 * @param {Array<Object>} events - Tableau d'objets représentant les événements.
 * @param {Object} dictionary - Dictionnaire avec un mot-clé et une chaîne de caractères associée.
 * @return {Array<Object>} Les événements mis à jour avec les titres remplacés par la chaîne de caractères associée au mot-clé.
 */
function replaceEventTitles(events, dictionary) {
  for (const event of events) {
    for (const keyword in dictionary) {
      if (event.summary.toLowerCase().indexOf(keyword.toLowerCase()) !== -1) {
        if (event.description) {
          event.description += `\n\nAncien titre : ${event.summary}`;
        } else {
          event.description = `Ancien titre : ${event.summary}`;
        }
        event.summary = dictionary[keyword];
        break;
      }
    }
  }

  return events;
}
