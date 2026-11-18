# Sales Opportunity Hub

## Sul tuo PC Windows (3 passi)
1. Estrai tutto lo zip in una cartella (clic destro > Estrai tutto).
2. Doppio clic su **AVVIA.bat** (per provarlo con dati di prova: **AVVIA-DEMO.bat**).
3. Si apre il browser. Accedi con `admin` / `admin` (poi scegli una nuova password).

Le librerie sono già incluse: serve solo Python (se manca, AVVIA prova a installarlo da solo).
Il programma gira in background, senza finestra. Il primo avvio crea sul Desktop il collegamento
«Sicuritalia Opportunity Hub»: dalla volta dopo basta quello.
Per spegnerlo: pulsante «Spegni il programma» nel menu (admin) oppure ARRESTA.bat.
Se non parte: AVVIA-DIAGNOSI.bat mostra il messaggio d'errore.

## Altri dispositivi
- Telefono sulla stessa rete Wi-Fi: AVVIA-RETE.bat e apri l'indirizzo mostrato (http://IP-del-PC:8000).
- Per tutta la rete vendita, con icona installabile sul telefono: mettilo online (Docker/Procfile, vedi sotto).

## Online
    docker build -t soh .   &&   docker run -d -p 8000:8000 -v soh-data:/data soh
Servizio con disco persistente su /data e variabili SOH_DATA=/data SOH_BEHIND_PROXY=1 SOH_SECURE_COOKIE=1 (HTTPS dalla piattaforma).
Usare 1 solo worker (già impostato). Backup: copiare la cartella dei dati (hub.db). Controllo: /healthz

## Funzioni
- Admin: aree e Direttori (Aree e geografia). Direttore: District, Team Leader, Venditori e zone (Utenti), rete in Organigramma.
- Zone di lavoro: Regione > Province > Comuni a tendina. Elenco comuni completo: download automatico da GeoNames (CC-BY 3.0) se c'è internet, oppure da «Aree e geografia».
- Dashboard con mappa (contratti verde, KO rosso, in corso blu); la cartina richiede internet, senza c'è una vista semplificata.

## Carica da email (tutti i profili)
In alto a destra c'è il pulsante **📧 Carica da email**. Scegli il PDF (o .eml) della mail di Disdette Retail o Caring, oppure incolla il testo: il programma legge cliente, telefono, località, appuntamento, motivazione, canoni ecc. Controlla i dati e premi "Crea opportunità".
- Agente / Recuperatore: l'opportunità è assegnata a lui e la gestisce subito.
- Direttore, District, Team Leader: devono scegliere l'agente (viene suggerito in base all'agente gestore della mail o alla zona di lavoro).
- Una stessa mail caricata due volte viene segnalata. L'abilitazione si può togliere dal menu Utenti > permessi ("Caricare opportunità da email").

## Proposte degli agenti (assistente + inserimento manuale)
Nel menu **💬 Proposte** l'agente descrive a parole un'opportunità (cliente, telefono, località, cosa chiede): l'assistente ricava i dati, l'agente li controlla e li invia. Si può anche inserire tutto a mano ("+ Nuova opportunità"). Ogni proposta va al **Direttore di area**, che la approva (scegliendo l'agente a cui assegnarla) o la rifiuta con un motivo; District e Team Leader vedono le proposte della loro squadra. L'assistente funziona con regole sul testo italiano, non usa servizi esterni.

## Liste di prospect dall'assistente
In **💬 Proposte** scrivi ad esempio «cerca mobilifici a Treviso» o «elenco ristoranti a Verona entro 5 km»: l'assistente cerca le attività vicine su OpenStreetMap (serve Internet) e mostra nome, indirizzo, telefono, sito e P.IVA **solo se presenti nella mappa** (spesso mancano). Scegli quelle che ti interessano e le invii al Direttore, oppure scarichi il file CSV (si apre con Excel).

## Calendario degli agenti in Outlook (v3.1)

In **Agenda → Collega calendario** ogni utente trova il proprio indirizzo personale (un link che finisce in `.ics`).
Si incolla in Outlook (*Aggiungi calendario → Iscriviti dal Web* / *Da Internet*) e gli appuntamenti fissati e le scadenze SLA
delle opportunità da gestire compaiono nel calendario, con promemoria. Vale anche per Google Calendar e iPhone.

- Funziona solo con la **versione online** del programma (con quella sul PC l'indirizzo non è raggiungibile da fuori).
- È in sola lettura: dal calendario non si modifica nulla. Outlook lo aggiorna da solo, in pochi minuti o alcune ore.
- Chi conosce l'indirizzo vede nomi e indirizzi dei clienti: «Genera un nuovo indirizzo» annulla quello vecchio.
  Se un utente viene disattivato o perde la funzione Agenda, il suo indirizzo smette di funzionare.
- Nel dettaglio di un appuntamento c'è anche **Aggiungi a calendario**, che scarica un file `.ics` da aprire con Outlook.

## Dashboard: filtri per agenti, campagne e periodo (v3.2)

Nella Dashboard (direttore, district, team leader) si possono combinare:
- **Agenti**: tutti, uno solo o più di uno (anche district e team leader: includono i loro agenti). C'è la ricerca per nome.
- **Categoria** (Customer base, Prospect, Disdette) e **Campagne** specifiche, una o più, come sotto-categoria. L'elenco mostra solo le campagne della categoria scelta.
- **Periodo** in base alla data di arrivo delle opportunità: preset (ultimi 7/30 giorni, questo mese, mese scorso, quest'anno) oppure date a scelta. Il grafico dell'andamento segue il periodo (al massimo gli ultimi 92 giorni).
I filtri valgono per numeri, grafici, tabelle e mappa. «Azzera filtri» li toglie tutti.

Mappa: il flag **Zone di lavoro** ora mostra i comuni di lavoro anche quando agli agenti è stata assegnata un'intera provincia (prima compariva solo per i comuni scelti uno a uno).

## Assistente AI per le liste di prospect (v3.3)

Nella pagina **Proposte** l'assistente capisce richieste scritte a parole («voglio tutti i ristoranti della provincia di Gorizia»,
«parrucchieri a Monfalcone e Grado, solo con telefono») e prepara la lista con nome, indirizzo, telefono, sito e P.IVA (se presenti).
Si può affinare scrivendo di seguito («solo quelli con telefono», «anche Cormons»). L'agente spunta i prospect e li invia al
Direttore di area, che li approva e li assegna come per le altre proposte. Si possono inviare fino a 500 prospect alla volta.

**Come si attiva** (solo amministratore):
1. Creare una chiave API su console.anthropic.com (sezione API keys) e caricare un po' di credito: si paga a consumo.
2. Online (Render o simili): aggiungere la variabile d'ambiente `ANTHROPIC_API_KEY` con la chiave. È il modo più sicuro.
   Sul PC: *Impostazioni → Assistente AI*, incollare la chiave e usare «Prova il collegamento».
3. Facoltativo: variabile `SOH_AI_MODEL` per cambiare modello (predefinito `claude-haiku-5-5`); `SOH_AI_MAX_ORA` e `SOH_AI_MAX_GIORNO`
   per i limiti di richieste per utente (predefiniti 40 all'ora e 150 al giorno).

**Cosa viene inviato ad Anthropic:** solo il testo scritto nella chat dell'assistente (che può contenere nomi e telefoni se l'agente
descrive un'opportunità). Le liste trovate e i dati del programma non vengono inviati. L'AI non cerca direttamente: restituisce solo
tipo di attività e luoghi; la ricerca la fa il programma, che controlla ogni valore.

**Se l'AI non è disponibile** (chiave assente o errata, credito finito, nessuna connessione) la chat usa la ricerca semplice a parole chiave
e avvisa l'agente. La ricerca per provincia intera funziona solo con l'AI attiva. Se la chiave è salvata dalle Impostazioni finisce anche
nella copia di sicurezza: conservarla con cura.

I dati vengono da OpenStreetMap: non sono completi e la P.IVA c'è raramente. Per zone molto grandi si vede solo una parte dei risultati.

**Ricerca lenta o in errore.** La ricerca usa il servizio gratuito di OpenStreetMap, che a volte è lento o sovraccarico, soprattutto per zone grandi.
Il programma prova più server, tiene i risultati 30 minuti (le richieste ripetute o gli affinamenti sono immediati) e, se una provincia
intera non risponde, cerca entro 15 km dal capoluogo avvisando l'agente. Se compare «lento o sovraccarico» basta riprovare dopo qualche minuto
o cercare un comune alla volta. Se invece compare «non riesco a collegarmi», il server non ha accesso a Internet o un firewall blocca overpass-api.de.

## Appuntamenti: limite consigliato, concomitanze e mail con data (v3.5)

- Il numero di appuntamenti al giorno (Impostazioni, predefinito 3) è ora un **limite consigliato**: oltre quel numero il programma avvisa
  e l'agente può inserire comunque l'appuntamento.
- Il programma segnala anche gli appuntamenti **concomitanti** (a meno di un'ora l'uno dall'altro): mostra l'elenco della giornata,
  evidenzia quali sono in conflitto (con il link «Apri per spostarlo»), suggerisce un'alternativa libera e permette di inserire comunque.
  L'inserimento forzato resta annotato nello storico dell'opportunità.
- **Carica da email**: se la mail (lead) contiene la data dell'appuntamento, compare il riquadro «Nella mail c'è un appuntamento» e si sceglie
  tra «Crea e inserisci in agenda» e «Crea senza fissare l'appuntamento». I controlli sull'agenda dell'agente scelto avvengono prima di creare.
- Corretto: la cartina della dashboard restava visibile sopra le finestre (ad esempio «Carica da email»). Ora le finestre stanno sempre sopra.

## Disdette: chiamate, chiusura e statistica trimestrale (v3.6)

- **Registro chiamate**: in ogni opportunità, «Registra una chiamata» (ho parlato / non risponde / da richiamare / numero errato / il cliente rifiuta) con note.
  Ogni chiamata resta nello storico, conta nel numero di chiamate e porta l'opportunità allo stato **In lavorazione**; da lì si può fissare un appuntamento o chiudere.
- **Chiusura delle disdette**: «Disdetta recuperata» (valore in euro + note obbligatorie) oppure «Disdetta non recuperata» con motivo a scelta tra
  *portata a naturale scadenza*, *con applicazione penali*, *senza applicazione penali*, *non lavorabile* (+ note obbligatorie).
- **Statistica** (Dashboard, «Disdette esitate per trimestre», anche in Excel): trimestri solari (gen–mar, apr–giu, lug–set, ott–dic) in base alla data di chiusura,
  con disdette esitate, recuperate (n, %, valore), non recuperate (n, %) e dettaglio per motivo. L'esportazione della rendicontazione include esito, motivo, valore e chiamate.
- Le disdette già chiuse prima di questa versione risultano recuperate (ex «contratto») o non recuperate (ex «KO»); queste ultime con «motivo non indicato».

**Aggiornamento v3.7 – statistica disdette.** Ogni disdetta conta nel trimestre solare in cui è stata assegnata. Dopo la fine del trimestre c'è un mese solare
in più per completare la lavorazione (es. 1° trimestre: fino al 30 aprile): la statistica resta «provvisoria» e diventa «definitiva» dal giorno dopo.
La tabella mostra anche le disdette ricevute e quelle ancora aperte.

## Utente «Ufficio Disdette» (v3.8)

Nuovo ruolo **Ufficio Disdette**: vede tutte le disdette dell'area e tutti i recuperatori (opportunità, dashboard, mappa, statistica trimestrale, esportazione, organigramma),
solo in lettura. Non vede customer base, prospect, proposte e agende. Si crea da *Utenti → + Ufficio Disdette* (o dall'Organigramma); risponde al Direttore di area.
Le funzioni si possono ampliare o ridurre utente per utente nella scheda «Accesso per funzione».

## Canone originario delle disdette (v3.9)

- «Carica da email»: dalla mail di disdetta viene letto il **Canone complessivo** (es. 50,00 €) e proposto nel campo «Canone complessivo attuale», modificabile.
- Alla chiusura («Disdetta recuperata» / «non recuperata») il recuperatore deve **confermare il canone originario** (precompilato).
- Statistica trimestrale: oltre a numero e % delle recuperate sulle esitate, mostra il **valore complessivo recuperato** (somma dei «Valore recuperato»),
  la sua percentuale sul canone originario delle recuperate e il canone originario complessivo delle esitate. Anche nell'Excel.

## Statistica disdette: dettaglio e non lavorabili (v3.10)

- Ogni numero della statistica trimestrale (ricevute, aperte, esitate, recuperate, valore, non recuperate, motivi) è **cliccabile**: si apre l'elenco delle disdette
  che lo compongono (cliente, recuperatore, date, esito e motivo, canone, valore, note, chiamate); cliccando una riga si apre la scheda. Vale per recuperatori e Ufficio Disdette (ciascuno nel proprio ambito).
- Le disdette chiuse come **non lavorabili** sono **escluse** dalle percentuali (recuperate e non recuperate sono calcolate sulle sole disdette lavorate); restano visibili in una colonna a parte.

## Grafici nella dashboard delle disdette (v3.11)

La sezione «Disdette esitate per trimestre» (recuperatori, Ufficio Disdette, direttori e tutti i ruoli con dashboard) ora ha anche i grafici:
esiti per trimestre (recuperate / non recuperate / non lavorabili), valore recuperato contro canone originario, percentuale di recupero per trimestre,
ciambella degli esiti dell'anno e ciambella dei motivi delle non recuperate. La tabella resta sotto, con i numeri cliccabili.

## Area manager e visibilità della dashboard disdette (v3.12)

- La sezione «Disdette esitate per trimestre» (grafici, tabella, dettaglio, Excel) è visibile **solo a Direttore di area, Recuperatore e Ufficio Disdette** (e all'amministratore).
- Nuovo ruolo **Area manager**, tra Direttore di area e District. Si crea da *Utenti → + Area manager* o dall'Organigramma e risponde al Direttore.
  Si assegnano le **zone di lavoro** (regioni, province, comuni) come per gli agenti: vede dashboard, mappa, opportunità, proposte, agenda e rendicontazione
  **come il Direttore ma solo per le sue zone**, più quelle della sua rete (District, Team Leader e agenti che rispondono a lui). Senza zone vede solo la sua rete.
  I District possono rispondere a un Area manager. Le funzioni si regolano utente per utente.

## Google Places (v3.13)

In **Impostazioni → Google Places** l'amministratore incolla la chiave (inizia con «AIza»), con i pulsanti «Salva», «Rimuovi» e «Prova il collegamento».
La chiave si può anche impostare sul server con la variabile `GOOGLE_PLACES_KEY` (consigliato in cloud). Con la chiave attiva la ricerca prospect e l'assistente
usano Google Places (nome, indirizzo, telefono, sito); la partita IVA non è disponibile. Se Google non risponde o rifiuta la chiave, il programma ripiega su OpenStreetMap e lo segnala.
Per contenere i costi: massimo 12 richieste Google per ricerca (3 luoghi × 3 categorie × 3 pagine, entro quel tetto), risultati ripetuti ricordati 30 minuti e blocco a 300 richieste al giorno
(modificabile con `SOH_GOOGLE_MAX_GIORNO`). In Google Cloud conviene impostare anche un budget con avviso e limitare la chiave a «Places API (New)».

## Elenco cliccabile in dashboard e mappa più robusta (v3.14)

- In **Dashboard** i riquadri (Totale, Gestite, Da gestire, SLA scaduto, Contratti, Non assegnate) e i numeri delle tabelle «Per categoria» e «Team» sono cliccabili:
  si apre l'elenco delle opportunità che compongono quel numero (con gli stessi filtri della dashboard) e cliccando una riga si vede il dettaglio. Vale per tutti i ruoli, ciascuno nel proprio ambito.
- **Mappa**: la cartina (libreria Leaflet) viene cercata prima nel programma (`static/leaflet/leaflet.js` e `leaflet.css`, opzionali), poi su tre reti di distribuzione diverse; se una è bloccata dal firewall si prova la successiva
  e un caricamento fallito non resta memorizzato. Se le immagini della cartina non si caricano compare un avviso; la spunta «Zone di lavoro» non spegne mai la cartina.
  In una rete aziendale restrittiva si può copiare Leaflet 1.9.4 (leaflet.js, leaflet.css e la cartella images) in `static/leaflet/` e la cartina funziona senza uscire verso altri siti (restano da raggiungere le immagini di OpenStreetMap).

## Controllo duplicati e customer base (v3.15)
**Customer base** (menu «Customer base», di default Direttore di area e amministratore): si carica l'elenco completo dei clienti dell'area (colonne: Cliente, P.IVA/Codice fiscale, Telefono, Località, Indirizzo). Per aggiornarla si ricarica il file: il programma mostra prima l'anteprima (nuovi, non più presenti, modificati, invariati) e poi si conferma «sostituisci» (il file è l'elenco completo) oppure «aggiungi soltanto». Resta lo storico degli aggiornamenti. Dopo ogni aggiornamento le proposte in attesa e i prospect ancora aperti vengono ricontrollati e il Direttore riceve la segnalazione di quelli che risultano ora già clienti.

**Come si riconosce lo stesso cliente**: stessa P.IVA/CF o stesso telefono = «già presente»; stesso nome (senza punteggiatura, srl/spa/snc, ordine delle parole) = «probabilmente già presente»; nome molto simile = «nome simile».

**Dove viene fatto il controllo**
- Import Excel di una campagna: prima dell'importazione compare l'elenco dei clienti già presenti (customer base, altre campagne, proposte in attesa, doppioni nel file) evidenziati; si sceglie «importa solo i nuovi», «importa tutti e segnala» o si annulla. Per le campagne di tipo customer base e disdette non si confronta con la customer base. Le righe segnalate mostrano il badge ⚠ nell'elenco e il dettaglio nell'opportunità.
- Prospect generati dall'assistente e proposte manuali: ogni nominativo è confrontato con customer base, campagne e proposte di altri venditori; i già presenti sono evidenziati e non preselezionati. Il Direttore vede l'avviso nella proposta da vagliare.
- «Verifica clienti» (tutti i ruoli con la funzione): controllo libero di un file o di un elenco incollato, con scarico del risultato in Excel.
Nuove funzioni nei permessi: «Customer base» e «Verifica clienti».

**Stesso indirizzo (v3.15b)**: il controllo segnala anche i clienti con indirizzo identico (via e numero civico, abbreviazioni come «V.» o «P.zza» comprese; se le località sono indicate in entrambi e sono diverse, non c'è segnalazione). È una segnalazione («Stesso indirizzo») e non fa saltare la riga con «Importa solo i nuovi», perché allo stesso indirizzo possono esserci attività diverse.

## Messaggi con allegati e documenti (v3.16)
**Messaggi** (menu «Messaggi»): Direttore di area, Area manager, District e Team Leader possono scrivere ai venditori che rispondono a loro (tutti, solo agenti, solo recuperatori, solo team leader/district oppure persone scelte) con oggetto, testo e fino a 10 allegati (15 MB ciascuno). I destinatari ricevono una notifica (campanella) e trovano il messaggio in «Messaggi»; chi invia vede chi lo ha letto. Chi invia (o l'amministratore) può eliminare il messaggio, allegati compresi.
**Documenti** (menu «Documenti»): ogni utente ha uno spazio personale (300 MB, visibile solo a lui; gli altri, direttore compreso, non lo vedono) dove salva i propri file e gli allegati ricevuti («Salva nei miei documenti»). L'**area comune** è visibile a tutti gli utenti dell'area (l'amministratore può pubblicare anche per tutte le aree) ed è organizzata in cartelle; pubblicano Direttore, Area manager, District e Team Leader, e chi ha caricato un file (o il Direttore) può eliminarlo.
Permessi nuovi (modificabili per utente): «Ricevere messaggi», «Inviare messaggi», «Documenti», «Pubblicare nell'area comune».
File non ammessi: eseguibili e script (.exe, .bat, .js, .html, .svg, ecc.).
**Copie di sicurezza**: i file sono nella cartella `data/allegati` (nel cloud, nel volume dati): vanno copiati insieme al file .db del backup.
**Nota sulle notifiche push**: le notifiche arrivano nell'app (campanella e pagina Messaggi). Le notifiche sul telefono a schermo bloccato richiedono un indirizzo HTTPS pubblico (versione cloud) e una configurazione dedicata: non sono incluse in questa versione.

## Guide d'uso e video (v3.16)
Insieme al programma trovi le guide illustrate per ruolo (PDF) e il video di presentazione. Correzione: l'Area manager vede ora anche la statistica delle disdette per le zone assegnate.

## Cartina e verifica clienti (v3.17)
- **Cartina:** i prospect inviati dalla ricerca (OpenStreetMap/Google) ora salvano anche le coordinate, e il comune viene riconosciuto dall'indirizzo («Via Roma 5, 31100 Treviso TV») o dalla posizione. Prima, senza città nel dato, non comparivano. Le opportunità già presenti senza località vengono collegate al comune al primo caricamento della dashboard. Se qualcuna resta fuori, sotto la cartina compare l'elenco con i nomi da correggere.
- **Verifica clienti:** controlla anche i prospect e le opportunità aperte di tutti i venditori dell'area. Un cliente già aperto da un altro è segnato «già in lavorazione» con il nome di chi lo segue, anche nel file Excel scaricato.

## Preventivi ai clienti (v3.18)
- Nell'opportunità (campagne customer base e prospect) l'agente trova «Preventivo al cliente» → **Fai un preventivo**. Sceglie i servizi da un menu (ispettivo, pronto intervento allarme, videosorveglianza, videoispezioni, collegamento incendio), risponde alle domande di ciascuno (ispezioni interne/esterne e notturne/diurne, allarme in comodato o solo periferica, interventi in franchigia, una tantum di installazione, telecamere, videoispezioni, periferica) e scrive il canone mensile. Si possono mettere più servizi nello stesso preventivo.
- Il **PDF** ha il formato dell'«Offerta tipo» Sicuritalia: copertina, presentazione, referenze, certificazioni e condizioni generali identiche; «Spett.le», città e data e le tabelle dell'offerta economica sono compilate dal programma. Le pagine fisse sono in `assets/offerta_base.pdf`: se il modello cambia, si rigenerano da lì.
- L'opportunità passa allo stato **Preventivo inviato**. Poi l'agente fissa l'appuntamento per la verifica con il cliente e registra l'esito (contratto, KO, da riprogrammare).
- **Invio:** dal telefono si apre la scelta dell'app (mail, WhatsApp…) con il PDF già allegato; dal PC il PDF viene scaricato e si apre la mail già scritta, da completare allegando il file. Il programma non invia mail da solo (non ha un server di posta).
- Permesso «Fare preventivi ai clienti»: attivo per agenti, team leader, district, area manager e direttore; si cambia per singolo utente.

### Preventivi: aggiornamento (v3.19)
- Il **Recuperatore disdette** può fare preventivi (anche sulle disdette, per proporre una nuova offerta al cliente); l'esito resta «Disdetta recuperata / non recuperata».
- Nuovo campo **Note** dopo l'elenco dei servizi: compare nel PDF in un riquadro «NOTE» sotto la tabella dell'offerta.
- **Elenco materiale (v3.20):** nella maschera del preventivo, sotto i servizi, si scelgono dal listino (assets/listino.json e foto in assets/listino_img, 198 articoli senza prezzi, ricavato dal file Excel «LISTINO TECNOLOGIE») gli articoli scrivendo solo il numero dei pezzi. Nel PDF, in fondo, compare l'«Allegato - Elenco materiale dell'impianto» con i soli articoli a quantità > 0 (per i kit anche i componenti). Per aggiornare il listino: `python3 listino_to_json.py LISTINO.xlsx assets/listino.json assets/listino_img` (rigenera anche le foto, prese dalla colonna «Immagine prodotto»; serve openpyxl e Pillow) (script incluso nello zip).

### Copia e ripristino dei dati dal programma (v3.21)
- **Impostazioni → Copia di sicurezza → «Scarica la copia (.zip)»** (solo amministratore): file con database e allegati.
- **Impostazioni → «Ripristina da una copia»** (solo amministratore): carica un .zip (o un vecchio .db) e sostituisce i dati correnti. Prima salva `prima-ripristino-AAAAMMGG-hhmmss.db` nella cartella dei dati. Dopo il ripristino si rifà l'accesso con un utente della copia.
- Uso tipico: scaricare la copia dal programma sul PC e caricarla sul programma online (Render, ecc.). Limite di caricamento: 70 MB.

### Valori del contratto e riepilogo del mese (v3.22)
- **Diventa contratto** apre il riquadro «Dati del contratto». **Prospect/lead**: si inserisce il valore del nuovo contratto (canone). **Customer base**: si sceglie *Modifica canone* o *Trasformazione* e si inseriscono valore del vecchio e del nuovo contratto. Valori obbligatori.
- **Dashboard**: riquadro «Riepilogo del mese solare» = somma dei nuovi canoni dei prospect + differenza positiva (nuovo − vecchio) sulla customer base, per i contratti chiusi nel mese; dettaglio per tipo, colonna «Valore» in «Per categoria» e «Team», tile «Valore contratti (periodo)» e, dall'elenco «Contratti», vecchio/nuovo canone di ogni opportunità.
- Export Excel: colonne Tipo contratto, Vecchio canone, Nuovo canone, Valore contratto. Le disdette non cambiano. I contratti chiusi prima di v3.22 non hanno valori e contano 0.
- **Stampa resoconto mensile (v3.22):** pulsante «🖨 Stampa resoconto» nel riquadro del mese in dashboard: apre la stampa con totale e dettaglio per tipo (rispetta i filtri scelti).
