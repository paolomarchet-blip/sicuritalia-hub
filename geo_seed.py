"""Archivio geografico di base: 20 regioni, 107 province (con capoluogo e coordinate approssimate)
e alcuni comuni del Nord-Est. L'elenco completo dei comuni italiani si carica da Impostazioni > Archivio geografico
(download automatico da GeoNames, CC-BY 3.0, oppure caricamento di un file)."""
import re
import unicodedata


def key(s):
    s = unicodedata.normalize("NFKD", str(s or "")).encode("ascii", "ignore").decode().lower()
    return re.sub(r"[^a-z0-9]", "", s)


REGIONI = ["Abruzzo", "Basilicata", "Calabria", "Campania", "Emilia-Romagna", "Friuli-Venezia Giulia", "Lazio", "Liguria",
           "Lombardia", "Marche", "Molise", "Piemonte", "Puglia", "Sardegna", "Sicilia", "Toscana", "Trentino-Alto Adige",
           "Umbria", "Valle d'Aosta", "Veneto"]

# (provincia, sigla, regione, capoluogo, lat, lon)
PROVINCE = [
    ("L'Aquila", "AQ", "Abruzzo", "L'Aquila", 42.35, 13.40), ("Chieti", "CH", "Abruzzo", "Chieti", 42.35, 14.17),
    ("Pescara", "PE", "Abruzzo", "Pescara", 42.46, 14.21), ("Teramo", "TE", "Abruzzo", "Teramo", 42.66, 13.70),
    ("Potenza", "PZ", "Basilicata", "Potenza", 40.64, 15.80), ("Matera", "MT", "Basilicata", "Matera", 40.67, 16.60),
    ("Catanzaro", "CZ", "Calabria", "Catanzaro", 38.91, 16.59), ("Cosenza", "CS", "Calabria", "Cosenza", 39.30, 16.25),
    ("Crotone", "KR", "Calabria", "Crotone", 39.08, 17.13), ("Reggio di Calabria", "RC", "Calabria", "Reggio di Calabria", 38.11, 15.65),
    ("Vibo Valentia", "VV", "Calabria", "Vibo Valentia", 38.68, 16.10),
    ("Napoli", "NA", "Campania", "Napoli", 40.85, 14.27), ("Avellino", "AV", "Campania", "Avellino", 40.91, 14.79),
    ("Benevento", "BN", "Campania", "Benevento", 41.13, 14.78), ("Caserta", "CE", "Campania", "Caserta", 41.07, 14.33),
    ("Salerno", "SA", "Campania", "Salerno", 40.68, 14.77),
    ("Bologna", "BO", "Emilia-Romagna", "Bologna", 44.49, 11.34), ("Ferrara", "FE", "Emilia-Romagna", "Ferrara", 44.84, 11.62),
    ("Forlì-Cesena", "FC", "Emilia-Romagna", "Forlì", 44.22, 12.04), ("Modena", "MO", "Emilia-Romagna", "Modena", 44.65, 10.93),
    ("Parma", "PR", "Emilia-Romagna", "Parma", 44.80, 10.33), ("Piacenza", "PC", "Emilia-Romagna", "Piacenza", 45.05, 9.69),
    ("Ravenna", "RA", "Emilia-Romagna", "Ravenna", 44.42, 12.20), ("Reggio nell'Emilia", "RE", "Emilia-Romagna", "Reggio nell'Emilia", 44.70, 10.63),
    ("Rimini", "RN", "Emilia-Romagna", "Rimini", 44.06, 12.57),
    ("Trieste", "TS", "Friuli-Venezia Giulia", "Trieste", 45.65, 13.78), ("Udine", "UD", "Friuli-Venezia Giulia", "Udine", 46.06, 13.24),
    ("Gorizia", "GO", "Friuli-Venezia Giulia", "Gorizia", 45.94, 13.62), ("Pordenone", "PN", "Friuli-Venezia Giulia", "Pordenone", 45.96, 12.66),
    ("Roma", "RM", "Lazio", "Roma", 41.90, 12.50), ("Frosinone", "FR", "Lazio", "Frosinone", 41.64, 13.34),
    ("Latina", "LT", "Lazio", "Latina", 41.47, 12.90), ("Rieti", "RI", "Lazio", "Rieti", 42.40, 12.86), ("Viterbo", "VT", "Lazio", "Viterbo", 42.42, 12.11),
    ("Genova", "GE", "Liguria", "Genova", 44.41, 8.93), ("Imperia", "IM", "Liguria", "Imperia", 43.89, 8.03),
    ("La Spezia", "SP", "Liguria", "La Spezia", 44.10, 9.82), ("Savona", "SV", "Liguria", "Savona", 44.31, 8.48),
    ("Milano", "MI", "Lombardia", "Milano", 45.46, 9.19), ("Bergamo", "BG", "Lombardia", "Bergamo", 45.70, 9.67),
    ("Brescia", "BS", "Lombardia", "Brescia", 45.54, 10.22), ("Como", "CO", "Lombardia", "Como", 45.81, 9.09),
    ("Cremona", "CR", "Lombardia", "Cremona", 45.13, 10.02), ("Lecco", "LC", "Lombardia", "Lecco", 45.86, 9.40),
    ("Lodi", "LO", "Lombardia", "Lodi", 45.31, 9.50), ("Mantova", "MN", "Lombardia", "Mantova", 45.16, 10.79),
    ("Monza e della Brianza", "MB", "Lombardia", "Monza", 45.58, 9.27), ("Pavia", "PV", "Lombardia", "Pavia", 45.19, 9.16),
    ("Sondrio", "SO", "Lombardia", "Sondrio", 46.17, 9.87), ("Varese", "VA", "Lombardia", "Varese", 45.82, 8.83),
    ("Ancona", "AN", "Marche", "Ancona", 43.62, 13.52), ("Ascoli Piceno", "AP", "Marche", "Ascoli Piceno", 42.85, 13.58),
    ("Fermo", "FM", "Marche", "Fermo", 43.16, 13.72), ("Macerata", "MC", "Marche", "Macerata", 43.30, 13.45),
    ("Pesaro e Urbino", "PU", "Marche", "Pesaro", 43.91, 12.91),
    ("Campobasso", "CB", "Molise", "Campobasso", 41.56, 14.66), ("Isernia", "IS", "Molise", "Isernia", 41.60, 14.23),
    ("Torino", "TO", "Piemonte", "Torino", 45.07, 7.69), ("Alessandria", "AL", "Piemonte", "Alessandria", 44.91, 8.61),
    ("Asti", "AT", "Piemonte", "Asti", 44.90, 8.21), ("Biella", "BI", "Piemonte", "Biella", 45.57, 8.05),
    ("Cuneo", "CN", "Piemonte", "Cuneo", 44.38, 7.54), ("Novara", "NO", "Piemonte", "Novara", 45.45, 8.62),
    ("Verbano-Cusio-Ossola", "VB", "Piemonte", "Verbania", 45.92, 8.55), ("Vercelli", "VC", "Piemonte", "Vercelli", 45.32, 8.42),
    ("Bari", "BA", "Puglia", "Bari", 41.12, 16.87), ("Barletta-Andria-Trani", "BT", "Puglia", "Barletta", 41.32, 16.28),
    ("Brindisi", "BR", "Puglia", "Brindisi", 40.63, 17.94), ("Foggia", "FG", "Puglia", "Foggia", 41.46, 15.54),
    ("Lecce", "LE", "Puglia", "Lecce", 40.35, 18.17), ("Taranto", "TA", "Puglia", "Taranto", 40.47, 17.24),
    ("Cagliari", "CA", "Sardegna", "Cagliari", 39.22, 9.12), ("Nuoro", "NU", "Sardegna", "Nuoro", 40.32, 9.33),
    ("Oristano", "OR", "Sardegna", "Oristano", 39.90, 8.59), ("Sassari", "SS", "Sardegna", "Sassari", 40.73, 8.56),
    ("Sud Sardegna", "SU", "Sardegna", "Carbonia", 39.17, 8.52),
    ("Palermo", "PA", "Sicilia", "Palermo", 38.12, 13.36), ("Agrigento", "AG", "Sicilia", "Agrigento", 37.31, 13.58),
    ("Caltanissetta", "CL", "Sicilia", "Caltanissetta", 37.49, 14.06), ("Catania", "CT", "Sicilia", "Catania", 37.50, 15.09),
    ("Enna", "EN", "Sicilia", "Enna", 37.57, 14.28), ("Messina", "ME", "Sicilia", "Messina", 38.19, 15.55),
    ("Ragusa", "RG", "Sicilia", "Ragusa", 36.93, 14.73), ("Siracusa", "SR", "Sicilia", "Siracusa", 37.08, 15.29),
    ("Trapani", "TP", "Sicilia", "Trapani", 38.02, 12.51),
    ("Firenze", "FI", "Toscana", "Firenze", 43.77, 11.25), ("Arezzo", "AR", "Toscana", "Arezzo", 43.46, 11.88),
    ("Grosseto", "GR", "Toscana", "Grosseto", 42.76, 11.11), ("Livorno", "LI", "Toscana", "Livorno", 43.55, 10.31),
    ("Lucca", "LU", "Toscana", "Lucca", 43.84, 10.50), ("Massa-Carrara", "MS", "Toscana", "Massa", 44.04, 10.14),
    ("Pisa", "PI", "Toscana", "Pisa", 43.72, 10.40), ("Pistoia", "PT", "Toscana", "Pistoia", 43.93, 10.92),
    ("Prato", "PO", "Toscana", "Prato", 43.88, 11.10), ("Siena", "SI", "Toscana", "Siena", 43.32, 11.33),
    ("Trento", "TN", "Trentino-Alto Adige", "Trento", 46.07, 11.12), ("Bolzano", "BZ", "Trentino-Alto Adige", "Bolzano", 46.50, 11.35),
    ("Perugia", "PG", "Umbria", "Perugia", 43.11, 12.39), ("Terni", "TR", "Umbria", "Terni", 42.56, 12.64),
    ("Aosta", "AO", "Valle d'Aosta", "Aosta", 45.74, 7.32),
    ("Verona", "VR", "Veneto", "Verona", 45.44, 10.99), ("Belluno", "BL", "Veneto", "Belluno", 46.14, 12.22),
    ("Padova", "PD", "Veneto", "Padova", 45.41, 11.88), ("Rovigo", "RO", "Veneto", "Rovigo", 45.07, 11.79),
    ("Treviso", "TV", "Veneto", "Treviso", 45.67, 12.24), ("Venezia", "VE", "Veneto", "Venezia", 45.44, 12.33),
    ("Vicenza", "VI", "Veneto", "Vicenza", 45.55, 11.55),
]

# Alcuni comuni aggiuntivi del Nord-Est (coordinate approssimate; il caricamento completo le sostituisce con quelle precise)
EXTRA = [
    ("VR", "Villafranca di Verona", 45.35, 10.84), ("VR", "Legnago", 45.19, 11.31), ("VR", "San Bonifacio", 45.40, 11.27),
    ("VR", "Bussolengo", 45.47, 10.85), ("VR", "Peschiera del Garda", 45.44, 10.69), ("VR", "Soave", 45.42, 11.25),
    ("VI", "Bassano del Grappa", 45.77, 11.73), ("VI", "Schio", 45.71, 11.36), ("VI", "Thiene", 45.71, 11.48),
    ("VI", "Arzignano", 45.52, 11.33), ("VI", "Valdagno", 45.65, 11.30),
    ("PD", "Cittadella", 45.65, 11.78), ("PD", "Este", 45.23, 11.66), ("PD", "Monselice", 45.24, 11.75),
    ("PD", "Abano Terme", 45.36, 11.79), ("PD", "Piove di Sacco", 45.30, 12.04),
    ("TV", "Castelfranco Veneto", 45.67, 11.93), ("TV", "Conegliano", 45.89, 12.30), ("TV", "Montebelluna", 45.78, 12.04),
    ("TV", "Vittorio Veneto", 45.99, 12.30), ("TV", "Oderzo", 45.78, 12.49),
    ("VE", "Chioggia", 45.22, 12.28), ("VE", "San Donà di Piave", 45.63, 12.56), ("VE", "Jesolo", 45.51, 12.64),
    ("VE", "Mirano", 45.49, 12.11), ("VE", "Portogruaro", 45.78, 12.84),
    ("BL", "Feltre", 46.02, 11.91), ("BL", "Cortina d'Ampezzo", 46.54, 12.14), ("RO", "Adria", 45.05, 12.06), ("RO", "Badia Polesine", 45.10, 11.49),
    ("BS", "Desenzano del Garda", 45.47, 10.54), ("BS", "Salò", 45.61, 10.52), ("BS", "Chiari", 45.54, 9.93),
    ("BS", "Montichiari", 45.41, 10.39), ("BS", "Lumezzane", 45.65, 10.27),
    ("BG", "Treviglio", 45.52, 9.59), ("BG", "Seriate", 45.68, 9.72), ("MN", "Castiglione delle Stiviere", 45.39, 10.49),
    ("TN", "Rovereto", 45.89, 11.04), ("TN", "Riva del Garda", 45.89, 10.84), ("BZ", "Merano", 46.67, 11.16), ("BZ", "Bressanone", 46.72, 11.66),
    ("UD", "Cividale del Friuli", 46.09, 13.43), ("UD", "Codroipo", 45.96, 12.98), ("PN", "Sacile", 45.95, 12.50), ("GO", "Monfalcone", 45.81, 13.53),
]


def seed(conn):
    """Inserisce regioni, province, capoluoghi ed extra (idempotente: non tocca i dati già presenti)."""
    for r in REGIONI:
        conn.execute("INSERT OR IGNORE INTO geo_regioni(nome) VALUES(?)", (r,))
    reg = {r["nome"]: r["id"] for r in conn.execute("SELECT id,nome FROM geo_regioni")}
    for nome, sigla, regione, cap, lat, lon in PROVINCE:
        conn.execute("INSERT OR IGNORE INTO geo_province(nome,sigla,regione_id) VALUES(?,?,?)", (nome, sigla, reg[regione]))
    prov = {r["sigla"]: r["id"] for r in conn.execute("SELECT id,sigla FROM geo_province")}
    for nome, sigla, regione, cap, lat, lon in PROVINCE:
        conn.execute("INSERT OR IGNORE INTO geo_comuni(nome,provincia_id,lat,lon) VALUES(?,?,?,?)", (cap, prov[sigla], lat, lon))
    for sigla, nome, lat, lon in EXTRA:
        conn.execute("INSERT OR IGNORE INTO geo_comuni(nome,provincia_id,lat,lon) VALUES(?,?,?,?)", (nome, prov[sigla], lat, lon))
