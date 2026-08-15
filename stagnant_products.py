import os

import pandas as pd
from dotenv import load_dotenv
from sqlalchemy import create_engine
from sqlalchemy.engine import URL


load_dotenv()

DB_USER = os.getenv("DB_USER", "root")
DB_PASSWORD = os.getenv("DB_PASSWORD", "")
DB_HOST = os.getenv("DB_HOST", "127.0.0.1")
DB_PORT = int(os.getenv("DB_PORT", "3306"))
DB_NAME = os.getenv("DB_NAME", "data_base")

connection_url = URL.create(
    drivername="mysql+pymysql",
    username=DB_USER,
    password=DB_PASSWORD,
    host=DB_HOST,
    port=DB_PORT,
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
