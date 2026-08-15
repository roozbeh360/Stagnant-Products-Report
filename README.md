# Stagnant Products Report mit Python und MySQL

Dieses Projekt erstellt eine Excel-Liste von Produkten, die aktuell Lagerbestand haben, aber seit mindestens 90 Tagen nicht verkauft wurden.

Die Daten werden aus einer MySQL- beziehungsweise MariaDB-Datenbank gelesen. Die SQL-Abfrage ermittelt den letzten Verkaufszeitpunkt jedes Produkts. Anschließend werden die Ergebnisse mit Pandas verarbeitet und als Excel-Datei gespeichert.

## Funktionen

- Verbindung zu MySQL oder MariaDB mit SQLAlchemy
- Abruf von Produkt- und Bestandsdaten
- Ermittlung des letzten Verkaufsdatums
- Ausschluss stornierter und erstatteter Bestellungen
- Erkennung von Produkten ohne Verkauf
- Erkennung von Produkten ohne Verkauf innerhalb der letzten 90 Tage
- Klassifizierung des Lagerstillstands
- Export des Berichts als Excel-Datei

## Voraussetzungen

- Python 3.10 oder höher
- MySQL oder MariaDB
- Zugriff auf die Datenbank
- Berechtigung zum Lesen der Tabellen
- Excel-Unterstützung durch `openpyxl`

## Installation

Erstelle zunächst eine virtuelle Umgebung:
```bash
python -m venv .venv

Aktivierung unter Linux oder macOS:

bash
source .venv/bin/activate

Aktivierung unter Windows:

powershell
.venv\Scripts\activate
```

Installiere anschließend die benötigten Pakete:

bash
pip install pandas sqlalchemy pymysql openpyxl python-dotenv

## Datenbankstruktur

Die Abfrage verwendet drei Tabellen:

1. `product`
2. `order_product`
3. ``order``

Das folgende Diagramm zeigt die erwarteten Beziehungen:

mermaid
erDiagram
PRODUCT ||--o{ ORDER_PRODUCT : "enthält"
ORDER ||--o{ ORDER_PRODUCT : "besteht aus"

PRODUCT {
bigint id PK
varchar barcode
varchar technical_code
varchar name
decimal stock
decimal price_wholesale
}

ORDER_PRODUCT {
bigint id PK
bigint product_id FK
bigint order_id FK
decimal quantity
decimal unit_price
}

ORDER {
bigint id PK
datetime created_date
varchar order_status
}

### Tabelle `product`

| Spalte | Typ | Beschreibung |
|---|---|---|
| `id` | `BIGINT` | Primärschlüssel des Produkts |
| `barcode` | `VARCHAR` | Barcode des Produkts |
| `technical_code` | `VARCHAR` | Technische Artikelnummer |
| `name` | `VARCHAR` | Produktname |
| `stock` | `DECIMAL` oder `INT` | Aktueller Lagerbestand |
| `price_wholesale` | `DECIMAL` | Großhandelspreis |

### Tabelle `order_product`

| Spalte | Typ | Beschreibung |
|---|---|---|
| `id` | `BIGINT` | Primärschlüssel der Position |
| `product_id` | `BIGINT` | Verweis auf `product.id` |
| `order_id` | `BIGINT` | Verweis auf `order.id` |
| `quantity` | `DECIMAL` oder `INT` | Verkaufte Menge |
| `unit_price` | `DECIMAL` | Verkaufspreis pro Einheit |

### Tabelle `order`

| Spalte | Typ | Beschreibung |
|---|---|---|
| `id` | `BIGINT` | Primärschlüssel der Bestellung |
| `created_date` | `DATETIME` | Erstellungsdatum der Bestellung |
| `order_status` | `VARCHAR` | Status der Bestellung |

> Hinweis: `order` ist ein reserviertes Schlüsselwort in MySQL und MariaDB. Deshalb muss der Tabellenname mit Backticks geschrieben werden: `` `order` ``.

## Beziehungen zwischen den Tabellen

Die Beziehungen lauten:

text
product.id
|
| 1:n
v
order_product.product_id

order.id
|
| 1:n
v
order_product.order_id

Ein Produkt kann in mehreren Bestellpositionen vorkommen. Eine Bestellung kann mehrere Produkte enthalten.

## Konfiguration

Lege im Projektverzeichnis eine Datei mit dem Namen `.env` an:

env
DB_USER=root
DB_PASSWORD=dein_passwort
DB_HOST=127.0.0.1
DB_PORT=3306
DB_NAME=amadpart_vip_05_5_19

Die Datei `.env` darf nicht in ein öffentliches GitHub-Repository hochgeladen werden.

Ergänze deshalb deine `.gitignore`:

gitignore
.env
.venv/
__pycache__/
*.xlsx

## Python-Skript

Speichere den folgenden Code beispielsweise als `stagnant_products.py`:

python
import os

import pandas as pd
from dotenv import load_dotenv
from sqlalchemy import create_engine
from sqlalchemy.engine import URL


load_dotenv()


DB_USER = os.getenv("DB_USER", "root")
DB_PASSWORD = os.getenv("DB_PASSWORD", "")
DB_HOST = os.getenv("DB_HOST", "127.0.0.1")
DB_PORT = os.getenv("DB_PORT", "3306")
DB_NAME = os.getenv("DB_NAME", "amadpart_vip_05_5_19")


connection_url = URL.create(
drivername="mysql+pymysql",
username=DB_USER,
password=DB_PASSWORD,
host=DB_HOST,
port=int(DB_PORT),
database=DB_NAME,
)

engine = create_engine(
connection_url,
pool_pre_ping=True,
)


query = """
SELECT
p.barcode,
p.technical_code,
p.id AS product_id,
p.name AS product_name,
COALESCE(p.stock, 0) AS stock_qty,
MAX(o.created_date) AS last_sale_date,
CASE
WHEN MAX(o.created_date) IS NULL THEN NULL
ELSE DATEDIFF(CURDATE(), MAX(o.created_date))
END AS days_since_last_sale
FROM product AS p
LEFT JOIN order_product AS op
ON op.product_id = p.id
LEFT JOIN `order` AS o
ON o.id = op.order_id
AND o.order_status NOT IN ('cancelled', 'refunded')
WHERE p.price_wholesale > 0
GROUP BY
p.barcode,
p.technical_code,
p.id,
p.name,
p.stock
HAVING
COALESCE(p.stock, 0) > 0
AND (
MAX(o.created_date) IS NULL
OR DATEDIFF(CURDATE(), MAX(o.created_date)) > 90
)
ORDER BY
last_sale_date ASC;
"""


def classify_stagnation(days):
"""
Klassifiziert ein Produkt anhand der Tage seit dem letzten Verkauf.
"""
if pd.isna(days):
return "Ohne Verkauf"

days = int(days)

if days <= 180:
return "Leicht stagnierend"

if days <= 365:
return "Mittel stagnierend"

return "Stark stagnierend"


try:
df = pd.read_sql(
sql=query,
con=engine,
parse_dates=["last_sale_date"],
)

df["stagnation_status"] = df["days_since_last_sale"].apply(
classify_stagnation
)

output_file = "stagnant_products.xlsx"
df.to_excel(output_file, index=False)

print("Die Verbindung zur Datenbank wurde erfolgreich hergestellt.")
print(f"Der Bericht wurde gespeichert: {output_file}")
print(df.head(10).to_string(index=False))

except Exception as error:
print("Fehler bei der Datenbankverbindung oder SQL-Abfrage:")
print(error)

finally:
engine.dispose()

## Skript ausführen

bash
python stagnant_products.py

Bei erfolgreicher Ausführung wird folgende Datei erstellt:

text
stagnant_products.xlsx

## Beispielausgabe

| barcode | technical_code | product_id | product_name | stock_qty | last_sale_date | days_since_last_sale | stagnation_status |
|---|---|---:|---|---:|---|---:|---|
| 123456789 | BP-1001 | 15 | Bremsbelag vorne | 12 | 2025-09-10 | 339 | Mittel stagnierend |
| 987654321 | FL-2004 | 21 | Ölfilter | 8 | 2025-01-15 | 577 | Stark stagnierend |
| 555666777 | SP-3002 | 34 | Zündkerze | 4 | `NULL` | `NULL` | Ohne Verkauf |

## Definition der Lagerstagnation

Das Skript verwendet folgende Klassifizierung:

| Bedingung | Status |
|---|---|
| Kein bisheriger Verkauf | Ohne Verkauf |
| 91 bis 180 Tage ohne Verkauf | Leicht stagnierend |
| 181 bis 365 Tage ohne Verkauf | Mittel stagnierend |
| Mehr als 365 Tage ohne Verkauf | Stark stagnierend |

Die SQL-Abfrage berücksichtigt nur Produkte mit:

sql
p.stock > 0

und:

sql
price_wholesale > 0

Zusätzlich werden Bestellungen mit folgenden Status ignoriert:

text
cancelled
refunded

## Gültige Bestellstatus

Die folgende Bedingung schließt stornierte und erstattete Bestellungen aus:

sql
o.order_status NOT IN ('cancelled', 'refunded')

Falls in deinem Shopsystem nur bestimmte Status als gültiger Verkauf gelten, ist eine Positivliste oft genauer:

sql
o.order_status IN ('paid', 'completed', 'shipped', 'delivered')

In diesem Fall sollte die Join-Bedingung wie folgt angepasst werden:

sql
LEFT JOIN `order` AS o
ON o.id = op.order_id
AND o.order_status IN (
'paid',
'completed',
'shipped',
'delivered'
)

Die tatsächlichen Statuswerte müssen an das verwendete Shopsystem angepasst werden.

## Wichtige technische Hinweise

### 1. Spaltenname `order_id`

Im ursprünglichen Code stand:

sql
si.oder_id

Das ist vermutlich ein Schreibfehler. Der korrekte Spaltenname lautet normalerweise:

sql
op.order_id

Falls die Spalte in deiner Datenbank tatsächlich anders heißt, muss die Abfrage entsprechend angepasst werden.

### 2. `GROUP BY`

Bei aktivierter MySQL-Option `ONLY_FULL_GROUP_BY` müssen alle nicht aggregierten Spalten im `GROUP BY` enthalten sein. Deshalb enthält die Abfrage alle ausgewählten Produktspalten im `GROUP BY`.

### 3. `parse_dates`

Die SQL-Abfrage liefert nur eine Datumsspalte:

text
last_sale_date

Daher wird in Pandas nur diese Spalte als Datum geparst:

python
parse_dates=["last_sale_date"]

Die Spalten `created_at`, `shipped_at`, `delivered_at` und `promise_date` waren in der ursprünglichen Abfrage nicht enthalten.

### 4. Passwort mit Sonderzeichen

Durch die Verwendung von `URL.create()` funktionieren auch Passwörter mit Sonderzeichen wie:

text
@ : / # ? %

sicherer als bei einer manuell zusammengesetzten Connection-URL.

## Performance

Bei großen Datenmengen sollten die folgenden Spalten indexiert sein:

sql
CREATE INDEX idx_product_stock
ON product (stock);

CREATE INDEX idx_product_wholesale_price
ON product (price_wholesale);

CREATE INDEX idx_order_product_product_id
ON order_product (product_id);

CREATE INDEX idx_order_product_order_id
ON order_product (order_id);

CREATE INDEX idx_order_status_created_date
ON `order` (order_status, created_date);

Vor dem Anlegen von Indizes sollte geprüft werden, ob entsprechende Indizes bereits existieren.

## Sicherheit

Verwende niemals echte Zugangsdaten direkt im Python-Code:

python
DB_PASSWORD = "mein_passwort"

Verwende stattdessen Umgebungsvariablen oder eine `.env`-Datei:

env
DB_PASSWORD=mein_passwort

Die `.env`-Datei muss in `.gitignore` eingetragen werden:

gitignore
.env

## Fehlerbehebung

### Fehler: `Access denied for user`

Prüfe:

- Benutzername
- Passwort
- Host
- Port
- Berechtigungen des Datenbankbenutzers

Beispiel für eine Leseberechtigung:

sql
GRANT SELECT ON amadpart_vip_05_5_19.*
TO 'report_user'@'localhost';

### Fehler: `Unknown column 'op.order_id'`

Prüfe die tatsächlichen Spaltennamen:

sql
DESCRIBE order_product;

### Fehler: `Unknown column 'o.created_date'`

Prüfe die Struktur der Bestellungstabelle:

sql
DESCRIBE `order`;

### Fehler: `Table 'product' doesn't exist`

Prüfe die vorhandenen Tabellen:

sql
SHOW TABLES;

Falls die Tabelle beispielsweise `products` statt `product` heißt, muss die SQL-Abfrage angepasst werden.

## Erweiterungsmöglichkeiten

Der Bericht kann später um folgende Funktionen erweitert werden:

- Lagerwert der stagnierenden Produkte
- Einkaufspreis und Verkaufspreis
- Produktmarke und Produktgruppe
- durchschnittlicher monatlicher Verkauf
- ABC-Klassifizierung
- farbliche Formatierung der Excel-Datei
- separate Excel-Arbeitsblätter für leichte, mittlere und starke Stagnation
- automatischer Versand des Berichts per E-Mail
- tägliche Ausführung über Cron oder Task Scheduler

## Projektstruktur

text
project/
├── stagnant_products.py
├── .env
├── .gitignore
├── requirements.txt
└── README.md

Beispiel für `requirements.txt`:

text
pandas
SQLAlchemy
PyMySQL
openpyxl
python-dotenv

## Lizenz

Dieses Projekt kann entsprechend der Lizenz des jeweiligen Unternehmens oder Projekts verwendet werden.
`

یک نکته مهم در کد اصلی شما وجود دارد: در `SELECT` ستون‌های `barcode` و `technical_code` انتخاب شده‌اند، اما در `GROUP BY` قرار نگرفته‌اند. در MariaDB با فعال بودن `ONLY_FULL_GROUP_BY` این موضوع خطا ایجاد می‌کند؛ در نسخه آماده‌شده این مشکل اصلاح شده است.# Stagnant Products Report mit Python und MySQL

Dieses Projekt erstellt eine Excel-Liste von Produkten, die aktuell Lagerbestand haben, aber seit mindestens 90 Tagen nicht verkauft wurden.

Die Daten werden aus einer MySQL- beziehungsweise MariaDB-Datenbank gelesen. Die SQL-Abfrage ermittelt den letzten Verkaufszeitpunkt jedes Produkts. Anschließend werden die Ergebnisse mit Pandas verarbeitet und als Excel-Datei gespeichert.

## Funktionen

- Verbindung zu MySQL oder MariaDB mit SQLAlchemy
- Abruf von Produkt- und Bestandsdaten
- Ermittlung des letzten Verkaufsdatums
- Ausschluss stornierter und erstatteter Bestellungen
- Erkennung von Produkten ohne Verkauf
- Erkennung von Produkten ohne Verkauf innerhalb der letzten 90 Tage
- Klassifizierung des Lagerstillstands
- Export des Berichts als Excel-Datei

## Voraussetzungen

- Python 3.10 oder höher
- MySQL oder MariaDB
- Zugriff auf die Datenbank
- Berechtigung zum Lesen der Tabellen
- Excel-Unterstützung durch `openpyxl`

## Installation

Erstelle zunächst eine virtuelle Umgebung:
```bash
python -m venv .venv

Aktivierung unter Linux oder macOS:

bash
source .venv/bin/activate

Aktivierung unter Windows:

powershell
.venv\Scripts\activate

Installiere anschließend die benötigten Pakete:

bash
pip install pandas sqlalchemy pymysql openpyxl python-dotenv

## Datenbankstruktur

Die Abfrage verwendet drei Tabellen:

1. `product`
2. `order_product`
3. ``order``

Das folgende Diagramm zeigt die erwarteten Beziehungen:

mermaid
erDiagram
PRODUCT ||--o{ ORDER_PRODUCT : "enthält"
ORDER ||--o{ ORDER_PRODUCT : "besteht aus"

PRODUCT {
bigint id PK
varchar barcode
varchar technical_code
varchar name
decimal stock
decimal price_wholesale
}

ORDER_PRODUCT {
bigint id PK
bigint product_id FK
bigint order_id FK
decimal quantity
decimal unit_price
}

ORDER {
bigint id PK
datetime created_date
varchar order_status
}

### Tabelle `product`

| Spalte | Typ | Beschreibung |
|---|---|---|
| `id` | `BIGINT` | Primärschlüssel des Produkts |
| `barcode` | `VARCHAR` | Barcode des Produkts |
| `technical_code` | `VARCHAR` | Technische Artikelnummer |
| `name` | `VARCHAR` | Produktname |
| `stock` | `DECIMAL` oder `INT` | Aktueller Lagerbestand |
| `price_wholesale` | `DECIMAL` | Großhandelspreis |

### Tabelle `order_product`

| Spalte | Typ | Beschreibung |
|---|---|---|
| `id` | `BIGINT` | Primärschlüssel der Position |
| `product_id` | `BIGINT` | Verweis auf `product.id` |
| `order_id` | `BIGINT` | Verweis auf `order.id` |
| `quantity` | `DECIMAL` oder `INT` | Verkaufte Menge |
| `unit_price` | `DECIMAL` | Verkaufspreis pro Einheit |

### Tabelle `order`

| Spalte | Typ | Beschreibung |
|---|---|---|
| `id` | `BIGINT` | Primärschlüssel der Bestellung |
| `created_date` | `DATETIME` | Erstellungsdatum der Bestellung |
| `order_status` | `VARCHAR` | Status der Bestellung |

> Hinweis: `order` ist ein reserviertes Schlüsselwort in MySQL und MariaDB. Deshalb muss der Tabellenname mit Backticks geschrieben werden: `` `order` ``.

## Beziehungen zwischen den Tabellen

Die Beziehungen lauten:

text
product.id
|
| 1:n
v
order_product.product_id

order.id
|
| 1:n
v
order_product.order_id

Ein Produkt kann in mehreren Bestellpositionen vorkommen. Eine Bestellung kann mehrere Produkte enthalten.

## Konfiguration

Lege im Projektverzeichnis eine Datei mit dem Namen `.env` an:

env
DB_USER=root
DB_PASSWORD=dein_passwort
DB_HOST=127.0.0.1
DB_PORT=3306
DB_NAME=amadpart_vip_05_5_19

Die Datei `.env` darf nicht in ein öffentliches GitHub-Repository hochgeladen werden.

Ergänze deshalb deine `.gitignore`:

gitignore
.env
.venv/
__pycache__/
*.xlsx

## Python-Skript

Speichere den folgenden Code beispielsweise als `stagnant_products.py`:

python
import os

import pandas as pd
from dotenv import load_dotenv
from sqlalchemy import create_engine
from sqlalchemy.engine import URL


load_dotenv()


DB_USER = os.getenv("DB_USER", "root")
DB_PASSWORD = os.getenv("DB_PASSWORD", "")
DB_HOST = os.getenv("DB_HOST", "127.0.0.1")
DB_PORT = os.getenv("DB_PORT", "3306")
DB_NAME = os.getenv("DB_NAME", "amadpart_vip_05_5_19")


connection_url = URL.create(
drivername="mysql+pymysql",
username=DB_USER,
password=DB_PASSWORD,
host=DB_HOST,
port=int(DB_PORT),
database=DB_NAME,
)

engine = create_engine(
connection_url,
pool_pre_ping=True,
)


query = """
SELECT
p.barcode,
p.technical_code,
p.id AS product_id,
p.name AS product_name,
COALESCE(p.stock, 0) AS stock_qty,
MAX(o.created_date) AS last_sale_date,
CASE
WHEN MAX(o.created_date) IS NULL THEN NULL
ELSE DATEDIFF(CURDATE(), MAX(o.created_date))
END AS days_since_last_sale
FROM product AS p
LEFT JOIN order_product AS op
ON op.product_id = p.id
LEFT JOIN `order` AS o
ON o.id = op.order_id
AND o.order_status NOT IN ('cancelled', 'refunded')
WHERE p.price_wholesale > 0
GROUP BY
p.barcode,
p.technical_code,
p.id,
p.name,
p.stock
HAVING
COALESCE(p.stock, 0) > 0
AND (
MAX(o.created_date) IS NULL
OR DATEDIFF(CURDATE(), MAX(o.created_date)) > 90
)
ORDER BY
last_sale_date ASC;
"""


def classify_stagnation(days):
"""
Klassifiziert ein Produkt anhand der Tage seit dem letzten Verkauf.
"""
if pd.isna(days):
return "Ohne Verkauf"

days = int(days)

if days <= 180:
return "Leicht stagnierend"

if days <= 365:
return "Mittel stagnierend"

return "Stark stagnierend"


try:
df = pd.read_sql(
sql=query,
con=engine,
parse_dates=["last_sale_date"],
)

df["stagnation_status"] = df["days_since_last_sale"].apply(
classify_stagnation
)

output_file = "stagnant_products.xlsx"
df.to_excel(output_file, index=False)

print("Die Verbindung zur Datenbank wurde erfolgreich hergestellt.")
print(f"Der Bericht wurde gespeichert: {output_file}")
print(df.head(10).to_string(index=False))

except Exception as error:
print("Fehler bei der Datenbankverbindung oder SQL-Abfrage:")
print(error)

finally:
engine.dispose()

## Skript ausführen

bash
python stagnant_products.py

Bei erfolgreicher Ausführung wird folgende Datei erstellt:

text
stagnant_products.xlsx

## Beispielausgabe

| barcode | technical_code | product_id | product_name | stock_qty | last_sale_date | days_since_last_sale | stagnation_status |
|---|---|---:|---|---:|---|---:|---|
| 123456789 | BP-1001 | 15 | Bremsbelag vorne | 12 | 2025-09-10 | 339 | Mittel stagnierend |
| 987654321 | FL-2004 | 21 | Ölfilter | 8 | 2025-01-15 | 577 | Stark stagnierend |
| 555666777 | SP-3002 | 34 | Zündkerze | 4 | `NULL` | `NULL` | Ohne Verkauf |

## Definition der Lagerstagnation

Das Skript verwendet folgende Klassifizierung:

| Bedingung | Status |
|---|---|
| Kein bisheriger Verkauf | Ohne Verkauf |
| 91 bis 180 Tage ohne Verkauf | Leicht stagnierend |
| 181 bis 365 Tage ohne Verkauf | Mittel stagnierend |
| Mehr als 365 Tage ohne Verkauf | Stark stagnierend |

Die SQL-Abfrage berücksichtigt nur Produkte mit:

sql
p.stock > 0

und:

sql
price_wholesale > 0

Zusätzlich werden Bestellungen mit folgenden Status ignoriert:

text
cancelled
refunded

## Gültige Bestellstatus

Die folgende Bedingung schließt stornierte und erstattete Bestellungen aus:

sql
o.order_status NOT IN ('cancelled', 'refunded')

Falls in deinem Shopsystem nur bestimmte Status als gültiger Verkauf gelten, ist eine Positivliste oft genauer:

sql
o.order_status IN ('paid', 'completed', 'shipped', 'delivered')

In diesem Fall sollte die Join-Bedingung wie folgt angepasst werden:

sql
LEFT JOIN `order` AS o
ON o.id = op.order_id
AND o.order_status IN (
'paid',
'completed',
'shipped',
'delivered'
)

Die tatsächlichen Statuswerte müssen an das verwendete Shopsystem angepasst werden.

## Wichtige technische Hinweise

### 1. Spaltenname `order_id`

Im ursprünglichen Code stand:

sql
si.oder_id

Das ist vermutlich ein Schreibfehler. Der korrekte Spaltenname lautet normalerweise:

sql
op.order_id

Falls die Spalte in deiner Datenbank tatsächlich anders heißt, muss die Abfrage entsprechend angepasst werden.

### 2. `GROUP BY`

Bei aktivierter MySQL-Option `ONLY_FULL_GROUP_BY` müssen alle nicht aggregierten Spalten im `GROUP BY` enthalten sein. Deshalb enthält die Abfrage alle ausgewählten Produktspalten im `GROUP BY`.

### 3. `parse_dates`

Die SQL-Abfrage liefert nur eine Datumsspalte:

text
last_sale_date

Daher wird in Pandas nur diese Spalte als Datum geparst:

python
parse_dates=["last_sale_date"]

Die Spalten `created_at`, `shipped_at`, `delivered_at` und `promise_date` waren in der ursprünglichen Abfrage nicht enthalten.

### 4. Passwort mit Sonderzeichen

Durch die Verwendung von `URL.create()` funktionieren auch Passwörter mit Sonderzeichen wie:

text
@ : / # ? %

sicherer als bei einer manuell zusammengesetzten Connection-URL.

## Performance

Bei großen Datenmengen sollten die folgenden Spalten indexiert sein:

sql
CREATE INDEX idx_product_stock
ON product (stock);

CREATE INDEX idx_product_wholesale_price
ON product (price_wholesale);

CREATE INDEX idx_order_product_product_id
ON order_product (product_id);

CREATE INDEX idx_order_product_order_id
ON order_product (order_id);

CREATE INDEX idx_order_status_created_date
ON `order` (order_status, created_date);

Vor dem Anlegen von Indizes sollte geprüft werden, ob entsprechende Indizes bereits existieren.

## Sicherheit

Verwende niemals echte Zugangsdaten direkt im Python-Code:

python
DB_PASSWORD = "mein_passwort"

Verwende stattdessen Umgebungsvariablen oder eine `.env`-Datei:

env
DB_PASSWORD=mein_passwort

Die `.env`-Datei muss in `.gitignore` eingetragen werden:

gitignore
.env

## Fehlerbehebung

### Fehler: `Access denied for user`

Prüfe:

- Benutzername
- Passwort
- Host
- Port
- Berechtigungen des Datenbankbenutzers

Beispiel für eine Leseberechtigung:

sql
GRANT SELECT ON amadpart_vip_05_5_19.*
TO 'report_user'@'localhost';

### Fehler: `Unknown column 'op.order_id'`

Prüfe die tatsächlichen Spaltennamen:

sql
DESCRIBE order_product;

### Fehler: `Unknown column 'o.created_date'`

Prüfe die Struktur der Bestellungstabelle:

sql
DESCRIBE `order`;

### Fehler: `Table 'product' doesn't exist`

Prüfe die vorhandenen Tabellen:

sql
SHOW TABLES;

Falls die Tabelle beispielsweise `products` statt `product` heißt, muss die SQL-Abfrage angepasst werden.

## Erweiterungsmöglichkeiten

Der Bericht kann später um folgende Funktionen erweitert werden:

- Lagerwert der stagnierenden Produkte
- Einkaufspreis und Verkaufspreis
- Produktmarke und Produktgruppe
- durchschnittlicher monatlicher Verkauf
- ABC-Klassifizierung
- farbliche Formatierung der Excel-Datei
- separate Excel-Arbeitsblätter für leichte, mittlere und starke Stagnation
- automatischer Versand des Berichts per E-Mail
- tägliche Ausführung über Cron oder Task Scheduler

## Projektstruktur

text
project/
├── stagnant_products.py
├── .env
├── .gitignore
├── requirements.txt
└── README.md

Beispiel für `requirements.txt`:

text
pandas
SQLAlchemy
PyMySQL
openpyxl
python-dotenv

## Lizenz

Dieses Projekt kann entsprechend der Lizenz des jeweiligen Unternehmens oder Projekts verwendet werden.
`
