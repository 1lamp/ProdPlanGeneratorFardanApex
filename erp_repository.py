"""Read-only access to Fardan ERP sales orders."""
import os
import keyring
import pyodbc


def _extract_delivery_time(conditions):
    """Extract the third numbered proforma condition when present."""
    if not conditions:
        return None
    import re
    match = re.search(
        r"(?ims)^\s*(?:۳|3)[.٫،)\-]\s*زمان\s*تحویل\s*[:：]?\s*(.*?)(?=^\s*(?:[۰-۹0-9]+)[.٫،)\-]|\Z)",
        str(conditions),
    )
    return match.group(1).strip() if match and match.group(1).strip() else None


class ErpRepository:
    """Provides parameterized, read-only queries to the Fardan OLTP database."""

    def __init__(self, settings):
        self.settings = settings

    @staticmethod
    def credential_service(settings):
        """Return a per-server credential key for Windows Credential Manager."""
        server = (settings.get("sql_server") or "").strip()
        database = (settings.get("sql_database") or "FARDAN_OLTP").strip()
        return f"FardanApex.ProdPlanGenerator|{server}|{database}"

    def _connect(self, password_override=None):
        server = os.getenv("FARDAN_SQL_SERVER") or self.settings.get("sql_server", "")
        database = os.getenv("FARDAN_SQL_DATABASE") or self.settings.get("sql_database", "FARDAN_OLTP")
        username = os.getenv("FARDAN_SQL_USERNAME") or self.settings.get("sql_username", "")
        driver = self.settings.get("sql_driver", "ODBC Driver 18 for SQL Server")
        authentication = self.settings.get("sql_authentication", "sql" if username else "windows")
        if not server:
            raise RuntimeError("آدرس SQL Server را در تنظیمات برنامه وارد کنید.")
        if not database:
            raise RuntimeError("نام پایگاه داده را در تنظیمات برنامه وارد کنید.")
        if authentication == "sql":
            if not username:
                raise RuntimeError("برای احراز هویت SQL، نام کاربری را در تنظیمات وارد کنید.")
            password = (
                password_override
                or os.getenv("FARDAN_SQL_PASSWORD")
                or keyring.get_password(self.credential_service(self.settings), username)
            )
            if not password:
                raise RuntimeError(
                    "رمز SQL ذخیره نشده است. رمز را در تنظیمات وارد و ذخیره کنید، "
                    "یا متغیر محیطی FARDAN_SQL_PASSWORD را تنظیم کنید."
                )
            auth = f"UID={username};PWD={password};"
        else:
            auth = "Trusted_Connection=yes;"
        connection_string = (
            f"DRIVER={{{driver}}};SERVER={server};DATABASE={database};"
            f"{auth}Encrypt=yes;TrustServerCertificate=yes;Connection Timeout=8;"
            "ApplicationIntent=ReadOnly;"
        )
        return pyodbc.connect(connection_string, readonly=True, autocommit=True)

    def test_connection(self, password_override=None):
        """Open a read-only connection and run a harmless connectivity check."""
        with self._connect(password_override=password_override) as connection:
            connection.cursor().execute("SELECT 1").fetchone()

    @staticmethod
    def _has_serial_timestamp(cursor):
        cursor.execute("""
            SELECT c.name
            FROM sys.columns c
            INNER JOIN sys.tables t ON t.object_id = c.object_id
            INNER JOIN sys.schemas s ON s.schema_id = t.schema_id
            WHERE s.name = N'sales'
              AND t.name = N'order_details'
              AND c.name IN (N'production_serial_created_at', N'serial_created_at')
        """)
        row = cursor.fetchone()
        return row[0] if row else None

    def list_ready_orders(self):
        """Only orders whose every item has a production serial are eligible."""
        with self._connect() as connection:
            cursor = connection.cursor()
            timestamp_column = self._has_serial_timestamp(cursor)
            serial_date_expression = (
                f"MAX(od.[{timestamp_column}])" if timestamp_column else "CAST(NULL AS DATETIME2)"
            )
            if timestamp_column:
                order_by = (
                    f"CASE WHEN {serial_date_expression} IS NULL THEN 1 ELSE 0 END, "
                    f"{serial_date_expression} DESC, o.order_date DESC, o.order_id DESC"
                )
            else:
                order_by = "o.order_date DESC, o.order_id DESC"

            cursor.execute(f"""
                SELECT
                    o.order_id, o.order_number, o.order_date, c.company_name,
                    COUNT(od.order_detail_id) AS item_count,
                    {serial_date_expression} AS serial_approved_at
                FROM sales.orders o
                INNER JOIN sales.companies c ON c.company_id = o.company_id
                INNER JOIN sales.order_details od ON od.order_id = o.order_id
                WHERE o.is_active = 1
                GROUP BY o.order_id, o.order_number, o.order_date, c.company_name
                HAVING COUNT(od.order_detail_id) > 0
                   AND COUNT(od.production_serial) = COUNT(od.order_detail_id)
                ORDER BY {order_by}
            """)
            columns = [column[0] for column in cursor.description]
            return [dict(zip(columns, row)) for row in cursor.fetchall()], bool(timestamp_column)

    def get_orders_for_processing(self, order_numbers):
        """Fetch order headers and line details for manually entered or listed orders."""
        if not order_numbers:
            return {}
        placeholders = ",".join("?" for _ in order_numbers)
        with self._connect() as connection:
            cursor = connection.cursor()
            cursor.execute(f"""
                SELECT
                    o.order_id, o.order_number, o.order_date, o.notes,
                    c.company_name, pf.delivery_time, pf.sales_conditions,
                    od.order_detail_id, od.ordered_quantity,
                    od.production_serial, od.item_notes,
                    p.technical_code, p.customer_part_name, p.unit_of_measure,
                    pd.delivery_commitment AS item_delivery_time,
                    pd.item_notes AS proforma_item_notes
                FROM sales.orders o
                INNER JOIN sales.companies c ON c.company_id = o.company_id
                INNER JOIN sales.order_details od ON od.order_id = o.order_id
                INNER JOIN sales.products p ON p.product_id = od.product_id
                LEFT JOIN sales.proformas pf ON pf.proforma_id = o.proforma_id
                LEFT JOIN sales.proforma_details pd
                    ON pd.proforma_detail_id = od.proforma_detail_id
                WHERE o.is_active = 1 AND o.order_number IN ({placeholders})
                ORDER BY o.order_date, o.order_id, od.order_detail_id
            """, order_numbers)
            columns = [column[0] for column in cursor.description]
            result = {}
            for row in cursor.fetchall():
                item = dict(zip(columns, row))
                number = str(item["order_number"]).strip()
                order = result.setdefault(number, {
                    "order_id": item["order_id"],
                    "order_number": number,
                    "order_date": item["order_date"],
                    "notes": item["notes"],
                    "company_name": item["company_name"],
                    "delivery_time": _extract_delivery_time(item.get("sales_conditions")) or item["delivery_time"],
                    "items": [],
                })
                order["items"].append({
                    "order_detail_id": item["order_detail_id"],
                    "quantity": item["ordered_quantity"],
                    "serial": item["production_serial"],
                    "technical_code": item["technical_code"],
                    "product_name": item["customer_part_name"],
                    "unit": item["unit_of_measure"],
                    "item_notes": item["proforma_item_notes"] or item["item_notes"],
                    "delivery_time": item["item_delivery_time"],
                })
            return result